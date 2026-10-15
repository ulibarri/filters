try { require('dotenv').config(); } catch (_) { /* dotenv es opcional */ }

const express = require('express');
const multer = require('multer');
const crypto = require('node:crypto');
const { parse } = require('csv-parse/sync');
const XLSX = require('xlsx');
const path = require('path');


const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 } // 25MB por archivo
});

// ---------- Config de la tarjeta de Teams / verificación de participante ----------
// Participante a verificar automáticamente en cada run (tarjeta de Teams + búsqueda
// automática de la Sección 3 de la web). Configurable por entorno (ver .env.example).
const TARGET_PARTICIPANT = process.env.TARGET_PARTICIPANT || 'MARIA CARO';

// Columnas mostradas en las tablas de la tarjeta de Teams (máximo 5). El modelo de
// datos actual (ver getDraftProjection) solo expone estas 4 columnas para identificar
// un viaje — no hay un ID de viaje ni origen/destino separados del Draft.
const TEAMS_CARD_COLUMNS = ['PACE MRN', 'Patient', 'Patient Address', 'Time'];

// Filas máximas por tabla de resultados antes de truncar (se reduce más si la
// tarjeta completa excede MAX_CARD_BYTES).
const MAX_TABLE_ROWS = 15;

// Límite duro de Teams/Power Automate para el tamaño de una Adaptive Card.
const MAX_CARD_BYTES = 28 * 1024;

// JSON solo se usa para la API (/api/process con base64) y no estorba a multipart/form-data.
app.use(express.json({ limit: '30mb' }));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  const lang = (req.query.lang === 'en' ? 'en' : 'es');
  res.render('index', { lang, targetParticipant: TARGET_PARTICIPANT });
});

// ---------- utils ----------
function get(obj, key) {
  if (!obj) return undefined;
  const k = String(key).trim().toLowerCase();
  let found;
  for (const [kk, vv] of Object.entries(obj)) {
    if (String(kk).trim().toLowerCase() === k) { found = vv; break; }
  }
  return typeof found === 'string' ? found.trim() : found;
}

function parseCsv(buffer) {
  const text = buffer.toString('utf8');
  return parse(text, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
    // Specialty SÍ depende de comillas reales para poder incluir comas dentro de
    // un campo (p.ej. "VOGEL, DON D" o "2330 LAKE MORENA DR SPC 32, CAMPO, CA 91906"),
    // así que aquí se mantiene el manejo de comillas, solo relajando comillas de
    // apertura sueltas para que no truene con "Invalid Opening Quote".
    relax_quotes: true
  });
}

// Draft, en cambio, no usa comillas para escapar comas (ningún campo las necesita);
// las comillas que aparecen ahí son basura literal de la fuente (apodos entre
// comillas, direcciones con símbolo de pulgadas, etc.) que rompen el parseo
// estricto con errores como "Invalid Opening/Closing Quote". Por eso Draft-como-CSV
// se parsea con quote:null (comillas tratadas como texto plano, no como delimitador),
// y luego sanitizeRows() limpia el caracter " de "Patient Address".
function parseCsvPermissive(buffer) {
  const text = buffer.toString('utf8');
  return parse(text, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
    quote: null
  });
}

// Limpieza que se aplica ANTES de cualquier otro procesamiento (filtrado/comparación):
// quita el caracter " del contenido de "Patient Address" en cada fila.
function sanitizeRows(rows) {
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (String(key).trim().toLowerCase() === 'patient address' && typeof row[key] === 'string') {
        row[key] = row[key].replace(/"/g, '').replace(/\s+/g, ' ').trim();
      }
    }
  }
  return rows;
}

// V2: el Draft ahora puede venir como .xlsx / .xls / .csv
function parseDraftFile(file) {
  const name = String(file.originalname || '').toLowerCase();
  if (name.endsWith('.csv')) {
    return parseCsvPermissive(file.buffer);
  }
  // Excel: se toma la primera hoja y se formatean las celdas tal como se ven en Excel
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });

  // Trim de strings y se descartan columnas sin encabezado (auto-generadas por SheetJS)
  return rawRows.map(r => {
    const out = {};
    for (const [k, v] of Object.entries(r)) {
      if (!k || /^__EMPTY/.test(k)) continue;
      out[k] = typeof v === 'string' ? v.trim() : v;
    }
    return out;
  });
}

function toCsv(rows) {
  if (!rows || rows.length === 0) return '';
  const headers = Object.keys(rows[0]);
  const esc = (v) => {
    if (v == null) return '';
    let s = String(v);
    if (s.includes('"')) s = s.replace(/"/g, '""');
    if (/[",\n]/.test(s)) s = `"${s}"`;
    return s;
  };
  const head = headers.map(esc).join(',');
  const body = rows.map(r => headers.map(h => esc(r[h])).join(',')).join('\n');
  return `${head}\n${body}`;
}

// Normaliza un nombre de participante para comparar sin importar mayúsculas/
// minúsculas, acentos, comas o el orden de nombre/apellido: quita diacríticos,
// pasa a minúsculas, trata comas como espacios, y ordena los tokens
// alfabéticamente. Así "CARO, MARIA", "Maria Caro" y "maría   caro" normalizan
// al mismo valor y se pueden comparar con ===.
function normalizeNameForMatch(raw) {
  const s = String(raw == null ? '' : raw)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/,/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return '';
  return s.split(' ').filter(Boolean).sort().join(' ');
}

function computePatientFromDraft(row) {
  const first = get(row, 'First Name') || '';
  const last = get(row, 'Last Name') || '';
  const middle = get(row, 'Middle_Name') || '';
  const full = [first, middle, last].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
  return full || get(row, 'Participant') || get(row, 'Patient') || '';
}

function getDraftProjection(draftRow) {
  return {
    'PACE MRN': get(draftRow, 'PACE MRN') || get(draftRow, 'Pace MRN'),
    'Patient': computePatientFromDraft(draftRow),
    'Patient Address': get(draftRow, 'Patient Address'),
    'Time': get(draftRow, 'Time') || get(draftRow, 'Start Time') || get(draftRow, 'Appointment Time')
  };
}

// ---- prefijo YYYYMMDD_ desde "Visit Date" del DRAFT ----
function normalizeDateToYYYYMMDD(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  const datePart = s.split(/[ T]/)[0];

  let m = datePart.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) {
    const [_, y, mo, d] = m;
    return `${y}${mo.padStart(2, '0')}${d.padStart(2, '0')}`;
  }
  m = datePart.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m) {
    let [_, a, b, y] = m;
    a = a.padStart(2, '0'); b = b.padStart(2, '0');
    const A = parseInt(a, 10), B = parseInt(b, 10);
    return (A > 12 && B <= 12) ? `${y}${b}${a}` : `${y}${a}${b}`;
  }
  m = datePart.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return datePart;

  const coerced = datePart.replace(/\//g, '-');
  const dt = new Date(coerced);
  if (!isNaN(dt.getTime())) {
    const y = String(dt.getFullYear());
    const mo = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}${mo}${d}`;
  }
  return null;
}

function computeVisitDatePrefixFromDraft(draftRows) {
  const counts = new Map();
  for (const r of draftRows) {
    const v = get(r, 'Visit Date') || get(r, 'Visit date') || get(r, 'VisitDate');
    if (!v) continue;
    const key = String(v).trim();
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  let chosen = null;
  if (counts.size > 0) {
    chosen = [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
  } else {
    const now = new Date();
    const y = String(now.getFullYear());
    const mo = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}${mo}${d}_`;
  }
  const ymd = normalizeDateToYYYYMMDD(chosen);
  if (ymd) return `${ymd}_`;
  const now = new Date();
  const y = String(now.getFullYear());
  const mo = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}${mo}${d}_`;
}

// ---------- V2: generación de Filters.csv a partir de Draft ----------
// Criterios (no distinguen mayúsculas/minúsculas):
// 1) Remover filas cuyo "Provider/Resource" contenga "Glenner"
// 2) Remover filas cuyo "Type" coincida (exacto, ignorando espacios) con la lista dada
// 3) Remover filas cuyas notas ("Appt Notes"/"Notes") contengan "dialysis"
const TYPE_EXCLUDE_LIST = [
  'alpine acs day [574]',
  'Day Center[458]',
  'hilltop acs day center [575]',
  'Magnolia Day Center [572]',
  'PACE Dial [469]'
];

function normalizeForMatch(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, '');
}

const TYPE_EXCLUDE_NORMALIZED = new Set(TYPE_EXCLUDE_LIST.map(normalizeForMatch));

function shouldExcludeFromFilters(row) {
  const provider = get(row, 'Provider/Resource') || '';
  if (String(provider).toLowerCase().includes('glenner')) return true;

  const type = get(row, 'Type') || '';
  if (TYPE_EXCLUDE_NORMALIZED.has(normalizeForMatch(type))) return true;

  const notes = get(row, 'Appt Notes') || get(row, 'Notes') || '';
  if (String(notes).toLowerCase().includes('dialysis')) return true;

  return false;
}

function generateFiltersFromDraft(draftRows) {
  return draftRows.filter(r => !shouldExcludeFromFilters(r));
}

// ---------- núcleo de procesamiento (compartido por /process y /api/process) ----------
// draftFile / specialtyFile: { originalname, buffer }
function processFiles(draftFile, specialtyFile) {
  let draftRows = parseDraftFile(draftFile);
  let specialtyRows = parseCsv(specialtyFile.buffer);

  // Limpieza previa a cualquier procesamiento: quita el caracter " de "Patient Address"
  draftRows = sanitizeRows(draftRows);
  specialtyRows = sanitizeRows(specialtyRows);

  // V2: Filters.csv ya no se sube, se genera a partir de Draft
  const filtersRows = generateFiltersFromDraft(draftRows);
  const filtersCsv = toCsv(filtersRows);

  const normalizeKey = (v) => (v == null ? '' : String(v).trim());
  const setFrom = (rows) => new Set(rows.map(r => normalizeKey(get(r, 'CSN'))).filter(Boolean));

  const csnFilters = setFrom(filtersRows);
  const csnSpecialty = setFrom(specialtyRows);

  // El lookup de proyección sigue usando el Draft original (sin filtrar)
  const draftByCSN = new Map();
  for (const r of draftRows) {
    const k = normalizeKey(get(r, 'CSN'));
    if (!k) continue;
    if (!draftByCSN.has(k)) draftByCSN.set(k, r);
  }

  const missingInSpecialty = [];
  for (const k of csnFilters) {
    if (!csnSpecialty.has(k)) {
      const draftMatch = draftByCSN.get(k);
      if (draftMatch) missingInSpecialty.push(getDraftProjection(draftMatch));
    }
  }

  const missingInFilters = [];
  for (const k of csnSpecialty) {
    if (!csnFilters.has(k)) {
      const draftMatch = draftByCSN.get(k);
      if (draftMatch) missingInFilters.push(getDraftProjection(draftMatch));
    }
  }

  const csv1 = toCsv(missingInSpecialty);
  const csv2 = toCsv(missingInFilters);

  // Índice para búsqueda (Sección 3 sigue usando el DRAFT original, no el Filters generado)
  const draftIndex = draftRows.map(getDraftProjection);

  // Prefijo por Visit Date (tomado del Draft original)
  const visitDatePrefix = computeVisitDatePrefixFromDraft(draftRows);

  return {
    trips_In_Filters_Not_In_Specialty: missingInSpecialty,
    trips_In_Specialty_Not_In_Filters: missingInFilters,
    csvFiles: {
      trips_In_Filters_Not_In_Specialty: csv1,
      trips_In_Specialty_Not_In_Filters: csv2
    },
    // V2: Filters.csv generado + su conteo de trips
    filtersCsv,
    filtersRowCount: filtersRows.length,
    draftRowCount: draftRows.length,
    specialtyRowCount: specialtyRows.length,
    // Filters.csv generado, ya proyectado — lo usa la verificación de participante
    // de la tarjeta de Teams (busca en Filters además de en las dos excepciones).
    filtersIndex: filtersRows.map(getDraftProjection),
    meta: {
      visitDatePrefix,
      draftIndex  // <-- para la sección de búsqueda (solo lo usa la página web)
    }
  };
}

// ---------- endpoint principal (página web, sin cambios) ----------
app.post(
  '/process',
  upload.fields([
    { name: 'draft', maxCount: 1 },
    { name: 'specialty', maxCount: 1 }
  ]),
  (req, res) => {
    try {
      if (!req.files?.draft?.[0] || !req.files?.specialty?.[0]) {
        return res.status(400).json({ error: 'Faltan archivos. Sube draft (xlsx/xls/csv) y specialty.csv.' });
      }
      // Se crea un run (igual que /api/process) para que la página web también
      // pueda usar /api/publish después de que el usuario revise los resultados.
      const run = createRun({
        draftFilename: req.files.draft[0].originalname,
        specialtyFilename: req.files.specialty[0].originalname
      });
      const result = processFiles(req.files.draft[0], req.files.specialty[0]);
      touchRun(run, { status: 'completed', result });
      res.json({ runId: run.id, ...result });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Error procesando los archivos. Verifica encabezados y formato (xlsx/xls/csv).' });
    }
  }
);

// =====================================================================
// ---------- API REST (para el Chatbot de MS Teams / Copilot) ----------
// =====================================================================
//
// Capa de autenticación temporal: por ahora deja pasar todo. Cuando se
// defina el mecanismo definitivo (API Key o Microsoft Entra ID), la
// validación real se agrega aquí mismo sin tener que tocar las rutas.
const authMiddleware = (req, res, next) => {
  next();
};
app.use('/api', authMiddleware);

// ---- Runs en memoria (igual que el resto de la app: sin base de datos) ----
// Render Free tier suspende el servicio tras ~15 min sin tráfico y lo reinicia
// en el próximo request; cualquier redeploy también reinicia el proceso. En
// ambos casos este Map se vacía, así que un runId previo deja de existir aunque
// no haya pasado RUN_TTL_MS. Por eso cualquier lookup que no encuentre el run
// responde 410 (no 404): el resultado existió pero ya no está disponible.
const RUNS = new Map();
const RUN_TTL_MS = 24 * 60 * 60 * 1000; // 24h
const RUN_EXPIRED_MESSAGE = 'Este resultado expiró; vuelve a procesar los archivos.';

function createRun(input) {
  const id = crypto.randomUUID();
  const now = new Date().toISOString();
  const run = {
    id,
    status: 'queued', // queued -> processing -> completed | failed
    createdAt: now,
    updatedAt: now,
    input, // { draftFilename, specialtyFilename }
    error: null,
    result: null,
    published: false,
    publishedAt: null
  };
  RUNS.set(id, run);
  return run;
}

function touchRun(run, patch) {
  Object.assign(run, patch, { updatedAt: new Date().toISOString() });
  return run;
}

// Limpieza periódica de runs viejos para no crecer memoria indefinidamente
setInterval(() => {
  const cutoff = Date.now() - RUN_TTL_MS;
  for (const [id, run] of RUNS) {
    if (new Date(run.updatedAt).getTime() < cutoff) RUNS.delete(id);
  }
}, 60 * 60 * 1000).unref();

// Procesa un run en segundo plano (no bloquea la respuesta del POST /api/process)
function runProcessingAsync(run, draftFile, specialtyFile) {
  touchRun(run, { status: 'processing' });
  setImmediate(() => {
    try {
      const result = processFiles(draftFile, specialtyFile);
      touchRun(run, { status: 'completed', result });
    } catch (err) {
      console.error(`[run ${run.id}] error:`, err);
      touchRun(run, {
        status: 'failed',
        error: 'Error procesando los archivos. Verifica encabezados y formato (xlsx/xls/csv).'
      });
    }
  });
}

// Acepta el archivo como multipart (igual que la página web) o, si no viene
// por multipart, como JSON: { draft: { filename, contentBase64 }, specialty: {...} }
// Esto da flexibilidad para distintos flujos de Copilot Studio / Power Automate.
function extractUploadedFile(req, fieldName) {
  const fromMulter = req.files?.[fieldName]?.[0];
  if (fromMulter) return { originalname: fromMulter.originalname, buffer: fromMulter.buffer };

  const fromBody = req.body?.[fieldName];
  if (fromBody && typeof fromBody.contentBase64 === 'string') {
    return {
      originalname: fromBody.filename || `${fieldName}.csv`,
      buffer: Buffer.from(fromBody.contentBase64, 'base64')
    };
  }
  return null;
}

function buildBaseUrl(req) {
  return process.env.API_BASE_URL || `${req.protocol}://${req.get('host')}`;
}

// POST /api/process — crea un run y lo procesa sin interacción humana.
// Responde de inmediato con el runId (202) y el procesamiento sigue en segundo plano;
// el llamador consulta el avance con GET /api/status/{runId}.
app.post('/api/process', upload.fields([
  { name: 'draft', maxCount: 1 },
  { name: 'specialty', maxCount: 1 }
]), (req, res) => {
  const draftFile = extractUploadedFile(req, 'draft');
  const specialtyFile = extractUploadedFile(req, 'specialty');

  if (!draftFile || !specialtyFile) {
    return res.status(400).json({
      error: 'Faltan archivos. Sube "draft" y "specialty" como multipart/form-data, o en JSON como { draft: { filename, contentBase64 }, specialty: { filename, contentBase64 } }.'
    });
  }

  const run = createRun({ draftFilename: draftFile.originalname, specialtyFilename: specialtyFile.originalname });
  const baseUrl = buildBaseUrl(req);

  res.status(202).json({
    runId: run.id,
    status: run.status,
    statusUrl: `${baseUrl}/api/status/${run.id}`,
    summaryUrl: `${baseUrl}/api/summary/${run.id}`,
    exceptionsUrl: `${baseUrl}/api/exceptions/${run.id}`
  });

  runProcessingAsync(run, draftFile, specialtyFile);
});

// GET /api/status/{runId}
app.get('/api/status/:runId', (req, res) => {
  const run = RUNS.get(req.params.runId);
  if (!run) return res.status(410).json({ error: RUN_EXPIRED_MESSAGE });
  res.json({
    runId: run.id,
    status: run.status,
    createdAt: run.createdAt,
    updatedAt: run.updatedAt,
    published: run.published,
    publishedAt: run.publishedAt,
    error: run.error
  });
});

// GET /api/summary/{runId}
app.get('/api/summary/:runId', (req, res) => {
  const run = RUNS.get(req.params.runId);
  if (!run) return res.status(410).json({ error: RUN_EXPIRED_MESSAGE });
  if (run.status !== 'completed') {
    return res.status(409).json({ runId: run.id, status: run.status, error: run.error || 'El run todavía no está completado.' });
  }
  const r = run.result;
  const prefix = r.meta.visitDatePrefix;
  res.json({
    runId: run.id,
    status: run.status,
    createdAt: run.createdAt,
    updatedAt: run.updatedAt,
    draftRowCount: r.draftRowCount,
    specialtyRowCount: r.specialtyRowCount,
    filtersRowCount: r.filtersRowCount,
    missingInSpecialtyCount: r.trips_In_Filters_Not_In_Specialty.length,
    missingInFiltersCount: r.trips_In_Specialty_Not_In_Filters.length,
    visitDatePrefix: prefix,
    files: {
      filters: `${prefix}Filters.csv`,
      missingInSpecialty: `${prefix}trips_In_Filters_Not_In_Specialty.csv`,
      missingInFilters: `${prefix}trips_In_Specialty_Not_In_Filters.csv`
    }
  });
});

// Campos permitidos para filtrar excepciones vía ?field=&value= — son las únicas
// columnas que expone la proyección de Draft usada en missingInSpecialty /
// missingInFilters (ver getDraftProjection). El match es "contiene" (substring),
// sin distinguir mayúsculas/minúsculas, igual que la búsqueda de la Sección 3 web.
const EXCEPTION_SEARCH_FIELDS = ['PACE MRN', 'Patient', 'Patient Address', 'Time'];

function matchesFieldValue(row, field, value) {
  const cell = String(row[field] ?? '').toLowerCase();
  return cell.includes(String(value).toLowerCase().trim());
}

// GET /api/exceptions/{runId}?type=missingInSpecialty|missingInFilters&field=<campo>&value=<valor>
// - type (opcional): limita a una de las dos categorías; si se omite, regresa ambas.
// - field/value (opcionales, deben venir juntos): filtran las filas devueltas.
//   field debe ser uno de EXCEPTION_SEARCH_FIELDS: 'PACE MRN', 'Patient',
//   'Patient Address' o 'Time'.
app.get('/api/exceptions/:runId', (req, res) => {
  const run = RUNS.get(req.params.runId);
  if (!run) return res.status(410).json({ error: RUN_EXPIRED_MESSAGE });
  if (run.status !== 'completed') {
    return res.status(409).json({ runId: run.id, status: run.status, error: run.error || 'El run todavía no está completado.' });
  }

  const { type, field, value } = req.query;

  if ((field && !value) || (!field && value)) {
    return res.status(400).json({
      error: 'Para filtrar por campo debes enviar "field" y "value" juntos.',
      allowedFields: EXCEPTION_SEARCH_FIELDS
    });
  }
  if (field && !EXCEPTION_SEARCH_FIELDS.includes(field)) {
    return res.status(400).json({
      error: `Campo "${field}" no permitido para filtrar.`,
      allowedFields: EXCEPTION_SEARCH_FIELDS
    });
  }

  const r = run.result;
  const applyFilter = (rows) => (field ? rows.filter(row => matchesFieldValue(row, field, value)) : rows);

  const payload = { runId: run.id };
  if (!type || type === 'missingInSpecialty') payload.missingInSpecialty = applyFilter(r.trips_In_Filters_Not_In_Specialty);
  if (!type || type === 'missingInFilters') payload.missingInFilters = applyFilter(r.trips_In_Specialty_Not_In_Filters);
  if (field) payload.filter = { field, value };
  res.json(payload);
});

// GET /api/files/{runId}/{which}  (which: filters | missingInSpecialty | missingInFilters)
// Descarga directa del CSV de un run completado — usado por los links del publish a Teams.
app.get('/api/files/:runId/:which', (req, res) => {
  const run = RUNS.get(req.params.runId);
  if (!run) return res.status(410).json({ error: RUN_EXPIRED_MESSAGE });
  if (run.status !== 'completed') return res.status(409).json({ error: 'El run todavía no está completado.' });

  const r = run.result;
  const prefix = r.meta.visitDatePrefix;
  const map = {
    filters: { csv: r.filtersCsv, filename: `${prefix}Filters.csv` },
    missingInSpecialty: { csv: r.csvFiles.trips_In_Filters_Not_In_Specialty, filename: `${prefix}trips_In_Filters_Not_In_Specialty.csv` },
    missingInFilters: { csv: r.csvFiles.trips_In_Specialty_Not_In_Filters, filename: `${prefix}trips_In_Specialty_Not_In_Filters.csv` }
  };
  const entry = map[req.params.which];
  if (!entry) return res.status(404).json({ error: 'Archivo no reconocido. Usa filters, missingInSpecialty o missingInFilters.' });

  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${entry.filename}"`);
  res.send(entry.csv);
});

// Busca TARGET_PARTICIPANT en las 3 categorías de resultados de un run (Filters
// generado, Missing in Specialty, Missing in Filters), usando comparación de
// nombre normalizada (ver normalizeNameForMatch). Cada match se etiqueta con la
// categoría donde se encontró — un mismo viaje puede aparecer en más de una
// (p.ej. un trip "missing in Specialty" por definición también está en Filters).
function searchParticipant(result, targetParticipant) {
  const target = normalizeNameForMatch(targetParticipant);
  if (!target) return [];

  const buckets = [
    { label: 'Filters', rows: result.filtersIndex || [] },
    { label: 'Missing in Specialty', rows: result.trips_In_Filters_Not_In_Specialty || [] },
    { label: 'Missing in Filters', rows: result.trips_In_Specialty_Not_In_Filters || [] }
  ];

  const matches = [];
  for (const bucket of buckets) {
    for (const row of bucket.rows) {
      if (normalizeNameForMatch(row['Patient']) === target) {
        matches.push({ ...row, 'Found in': bucket.label });
      }
    }
  }
  return matches;
}

function cellText(v) {
  return v == null ? '' : String(v);
}

// Construye un elemento "Table" de Adaptive Cards 1.5 a partir de columnas +
// filas. La primera fila son los encabezados (en negritas, firstRowAsHeaders
// además le indica al host que las estilice como tal).
function buildTableElement(headers, rows) {
  const headerCells = headers.map(h => ({
    type: 'TableCell',
    items: [{ type: 'TextBlock', text: cellText(h), weight: 'bolder', wrap: true }]
  }));
  const dataRows = rows.map(row => ({
    type: 'TableRow',
    cells: headers.map(h => ({
      type: 'TableCell',
      items: [{ type: 'TextBlock', text: cellText(row[h]), wrap: true }]
    }))
  }));
  return {
    type: 'Table',
    firstRowAsHeaders: true,
    showGridLines: true,
    columns: headers.map(() => ({ width: 1 })),
    rows: [{ type: 'TableRow', cells: headerCells }, ...dataRows]
  };
}

// Sección de una tabla de resultados (missingInSpecialty / missingInFilters) con
// truncado a rowLimit filas. rowLimit puede bajar hasta 0 (lo controla el ajuste
// de tamaño en buildTeamsMessageFitted) sin tocar las filas de Participant check.
function buildResultsSection(titleText, rows, rowLimit) {
  const elements = [
    { type: 'TextBlock', text: titleText, weight: 'bolder', size: 'medium', wrap: true, spacing: 'medium' }
  ];

  if (!rows || rows.length === 0) {
    elements.push({ type: 'TextBlock', text: 'No trips found in this category.', wrap: true, isSubtle: true });
    return elements;
  }

  const shown = rows.slice(0, Math.max(0, rowLimit));
  if (shown.length > 0) {
    elements.push(buildTableElement(TEAMS_CARD_COLUMNS, shown));
  }

  if (rows.length > shown.length) {
    elements.push({
      type: 'TextBlock',
      text: `Showing ${shown.length} of ${rows.length} trips. Download the full CSV for the complete list.`,
      wrap: true,
      isSubtle: true,
      size: 'small'
    });
  }

  return elements;
}

// Sección "Participant check" al inicio de la tarjeta. Nunca se trunca por
// tamaño — ver buildTeamsMessageFitted.
function buildParticipantCheckSection(result, targetParticipant) {
  const matches = searchParticipant(result, targetParticipant);
  const heading = {
    type: 'TextBlock',
    text: `Participant check: ${targetParticipant}`,
    weight: 'bolder',
    size: 'medium',
    wrap: true,
    spacing: 'medium'
  };

  if (matches.length === 0) {
    return [
      heading,
      {
        type: 'Container',
        style: 'good',
        bleed: true,
        items: [
          { type: 'TextBlock', text: `✅ No appointments found for ${targetParticipant}.`, wrap: true }
        ]
      }
    ];
  }

  const participantColumns = [...TEAMS_CARD_COLUMNS, 'Found in'];
  return [
    heading,
    {
      type: 'Container',
      style: 'attention',
      bleed: true,
      items: [
        { type: 'TextBlock', text: `⚠️ ${matches.length} appointment(s) found for ${targetParticipant}`, wrap: true, weight: 'bolder' },
        buildTableElement(participantColumns, matches)
      ]
    }
  ];
}

// Construye el mensaje para un webhook de "Workflows" de Teams (tarjeta Adaptive Card).
// Si el canal usa en cambio un conector clásico "Incoming Webhook", el formato
// esperado es distinto (MessageCard) y habría que ajustar esta función.
// rows1Limit/rows2Limit controlan cuántas filas se muestran en cada tabla de
// resultados — los ajusta buildTeamsMessageFitted si la tarjeta no cabe en
// MAX_CARD_BYTES.
function buildTeamsMessage(run, baseUrl, rows1Limit, rows2Limit) {
  const r = run.result;
  const prefix = r.meta.visitDatePrefix;

  const body = [
    { type: 'TextBlock', text: 'Dispatch Run Summary', weight: 'bolder', size: 'large', wrap: true },
    { type: 'TextBlock', text: `Run ${run.id}`, isSubtle: true, size: 'small', wrap: true },
    {
      type: 'FactSet',
      facts: [
        { title: 'Date', value: new Date(run.createdAt).toISOString() },
        { title: 'Draft', value: `${r.draftRowCount} rows` },
        { title: 'Specialty', value: `${r.specialtyRowCount} rows` },
        { title: 'Filters generated', value: `${r.filtersRowCount} trips` },
        { title: 'Missing in Specialty', value: String(r.trips_In_Filters_Not_In_Specialty.length) },
        { title: 'Missing in Filters', value: String(r.trips_In_Specialty_Not_In_Filters.length) },
        { title: 'Visit date prefix', value: prefix }
      ]
    },
    ...buildParticipantCheckSection(r, TARGET_PARTICIPANT),
    ...buildResultsSection('TRIPS FOUND IN FILTERS MISSING IN SPECIALTY', r.trips_In_Filters_Not_In_Specialty, rows1Limit),
    ...buildResultsSection('TRIPS FOUND IN SPECIALTY MISSING IN FILTERS', r.trips_In_Specialty_Not_In_Filters, rows2Limit),
    {
      type: 'TextBlock',
      text: '⚠️ Links available temporarily',
      weight: 'bolder',
      isSubtle: true,
      size: 'small',
      wrap: true
    }
  ];

  const card = {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body,
    actions: [
      { type: 'Action.OpenUrl', title: `Download ${prefix}Filters.csv`, url: `${baseUrl}/api/files/${run.id}/filters` },
      { type: 'Action.OpenUrl', title: 'Download exceptions (Filters→Specialty)', url: `${baseUrl}/api/files/${run.id}/missingInSpecialty` },
      { type: 'Action.OpenUrl', title: 'Download exceptions (Specialty→Filters)', url: `${baseUrl}/api/files/${run.id}/missingInFilters` }
    ]
  };
  return {
    type: 'message',
    attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: card }]
  };
}

// Envuelve buildTeamsMessage ajustando rows1Limit/rows2Limit hacia abajo (nunca
// las filas de Participant check) hasta que la tarjeta quepa en MAX_CARD_BYTES,
// el límite real de Teams/Power Automate.
function buildTeamsMessageFitted(run, baseUrl) {
  let rows1 = MAX_TABLE_ROWS;
  let rows2 = MAX_TABLE_ROWS;
  let message = buildTeamsMessage(run, baseUrl, rows1, rows2);
  let size = Buffer.byteLength(JSON.stringify(message), 'utf8');

  while (size > MAX_CARD_BYTES && (rows1 > 0 || rows2 > 0)) {
    if (rows1 >= rows2 && rows1 > 0) rows1--;
    else if (rows2 > 0) rows2--;
    message = buildTeamsMessage(run, baseUrl, rows1, rows2);
    size = Buffer.byteLength(JSON.stringify(message), 'utf8');
  }

  return message;
}

// POST /api/publish — body: { "runId": "..." }
// Entrega el resultado de un run completado a un canal de MS Teams, mediante un
// webhook configurado en el servidor (nunca lo manda el llamador, por seguridad).
app.post('/api/publish', async (req, res) => {
  const { runId } = req.body || {};
  if (!runId) return res.status(400).json({ error: 'Falta "runId" en el body.' });

  const run = RUNS.get(runId);
  if (!run) return res.status(410).json({ error: RUN_EXPIRED_MESSAGE });
  if (run.status !== 'completed') {
    return res.status(409).json({ error: `El run está en estado "${run.status}", no se puede publicar hasta que esté "completed".` });
  }

  const webhookUrl = process.env.TEAMS_WEBHOOK_URL;
  if (!webhookUrl) {
    return res.status(501).json({
      error: 'TEAMS_WEBHOOK_URL no está configurada en el servidor. Crea un webhook de "Workflows" en el canal de Teams y ponlo en la variable de entorno TEAMS_WEBHOOK_URL.'
    });
  }

  try {
    const baseUrl = buildBaseUrl(req);
    const message = buildTeamsMessageFitted(run, baseUrl);
    const resp = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(message)
    });

    if (!resp.ok) {
      const bodyText = await resp.text().catch(() => '');
      return res.status(502).json({ error: `El webhook de Teams respondió ${resp.status}.`, details: bodyText.slice(0, 500) });
    }

    touchRun(run, { published: true, publishedAt: new Date().toISOString() });
    res.json({ runId: run.id, published: true, publishedAt: run.publishedAt });
  } catch (err) {
    console.error(`[run ${run.id}] error publicando a Teams:`, err);
    res.status(502).json({ error: 'No se pudo entregar el mensaje al webhook de Teams.' });
  }
});

// Guard para poder hacer require('./server.js') desde scripts (p.ej.
// scripts/card-preview.js) sin levantar el servidor HTTP.
if (require.main === module) {
  const PORT = process.env.PORT || 3200;
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`CSV Diff app V2 corriendo en http://localhost:${PORT}`);
  });
}

module.exports = {
  app,
  processFiles,
  buildTeamsMessage,
  buildTeamsMessageFitted,
  searchParticipant,
  normalizeNameForMatch,
  TARGET_PARTICIPANT,
  TEAMS_CARD_COLUMNS,
  MAX_TABLE_ROWS,
  MAX_CARD_BYTES
};
