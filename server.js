const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const XLSX = require('xlsx');
const path = require('path');


const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  const lang = (req.query.lang === 'en' ? 'en' : 'es');
  res.render('index', { lang });
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
  return parse(text, { columns: true, skip_empty_lines: true, trim: true });
}

// V2: el Draft ahora puede venir como .xlsx / .xls / .csv
function parseDraftFile(file) {
  const name = String(file.originalname || '').toLowerCase();
  if (name.endsWith('.csv')) {
    return parseCsv(file.buffer);
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
    return `${y}${mo.padStart(2,'0')}${d.padStart(2,'0')}`;
  }
  m = datePart.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m) {
    let [_, a, b, y] = m;
    a = a.padStart(2,'0'); b = b.padStart(2,'0');
    const A = parseInt(a,10), B = parseInt(b,10);
    return (A > 12 && B <= 12) ? `${y}${b}${a}` : `${y}${a}${b}`;
  }
  m = datePart.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return datePart;

  const coerced = datePart.replace(/\//g, '-');
  const dt = new Date(coerced);
  if (!isNaN(dt.getTime())) {
    const y = String(dt.getFullYear());
    const mo = String(dt.getMonth()+1).padStart(2,'0');
    const d = String(dt.getDate()).padStart(2,'0');
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
    chosen = [...counts.entries()].sort((a,b)=>b[1]-a[1])[0][0];
  } else {
    const now = new Date();
    const y = String(now.getFullYear());
    const mo = String(now.getMonth()+1).padStart(2,'0');
    const d = String(now.getDate()).padStart(2,'0');
    return `${y}${mo}${d}_`;
  }
  const ymd = normalizeDateToYYYYMMDD(chosen);
  if (ymd) return `${ymd}_`;
  const now = new Date();
  const y = String(now.getFullYear());
  const mo = String(now.getMonth()+1).padStart(2,'0');
  const d = String(now.getDate()).padStart(2,'0');
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

// ---------- endpoint principal ----------
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

      const draftRows = parseDraftFile(req.files.draft[0]);
      const specialtyRows = parseCsv(req.files.specialty[0].buffer);

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

      res.json({
        trips_In_Filters_Not_In_Specialty: missingInSpecialty,
        trips_In_Specialty_Not_In_Filters: missingInFilters,
        csvFiles: {
          trips_In_Filters_Not_In_Specialty: csv1,
          trips_In_Specialty_Not_In_Filters: csv2
        },
        // V2: Filters.csv generado + su conteo de trips
        filtersCsv,
        filtersRowCount: filtersRows.length,
        meta: {
          visitDatePrefix,
          draftIndex  // <-- para la sección de búsqueda
        }
      });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Error procesando los archivos. Verifica encabezados y formato (xlsx/xls/csv).' });
    }
  }
);

const PORT = process.env.PORT || 3200;
app.listen(PORT, () => {
  console.log(`CSV Diff app V2 corriendo en http://localhost:${PORT}`);
});
