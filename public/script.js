const i18n = {
  es: {
    title: 'Comparar viajes filtrados vs viajes CSV Specialty',
    language: 'Idioma',
    hint: 'Sube los archivos <strong>Draft (.xlsx/.xls/.csv)</strong> y <strong>specialty.csv</strong>. La app genera Filters automáticamente.',
    step1: '1) Cargar archivos',
    step2: '2) Resultados',
    step3: '3) Búsqueda en Draft.csv',
    searchHint: 'Busca por <strong>PACE MRN</strong> o <strong>Patient</strong>.',
    process: 'Procesar datos',
    download: 'Descargar CSV',
    empty: 'Sin datos',
    footer: 'CFS California Fleet Services Inc. 2025',
    processing: 'Procesando...',
    errorPrefix: 'Error: ',
    searchBtn: 'Buscar',
    searchPlaceholder: 'Criterio de búsqueda (ej. Caro, SDP10035365)',
    countSingle: '{n} viaje',
    countPlural: '{n} viajes',
    done: 'Listo',
    draftLabel: 'Draft (.xlsx, .xls o .csv)',
    filtersTitle: 'Filters.csv (generado)',
    downloadFilters: 'Descargar Filters.csv generado'
  },
  en: {
    title: 'Compare Filtered trips vs Specialty CSV trips',
    language: 'Language',
    hint: 'Upload <strong>Draft (.xlsx/.xls/.csv)</strong> and <strong>specialty.csv</strong> files. Filters is generated automatically.',
    step1: '1) Upload files',
    step2: '2) Results',
    step3: '3) Search in Draft.csv',
    searchHint: 'Search by <strong>PACE MRN</strong> or <strong>Patient</strong>.',
    process: 'Process data',
    download: 'Download CSV',
    empty: 'No data',
    footer: 'CFS California Fleet Services Inc. 2025',
    processing: 'Processing...',
    errorPrefix: 'Error: ',
    searchBtn: 'Search',
    searchPlaceholder: 'Search criteria (e.g., Caro, SDP10035365)',
    countSingle: '{n} trip',
    countPlural: '{n} trips',
    done: 'Done',
    draftLabel: 'Draft (.xlsx, .xls or .csv)',
    filtersTitle: 'Filters.csv (generated)',
    downloadFilters: 'Download generated Filters.csv'
  }
};

function formatCount(lang, n) {
  const d = i18n[lang] || i18n.es;
  const key = (n === 1) ? 'countSingle' : 'countPlural';
  return (d[key] || '{n}').replace('{n}', n);
}


function applyLang(lang) {
  const d = i18n[lang] || i18n.es;
  document.getElementById('t_title').textContent = d.title;
  document.getElementById('t_language').textContent = d.language;
  document.getElementById('t_hint').innerHTML = d.hint;
  document.getElementById('t_step1').textContent = d.step1;
  document.getElementById('t_step2').textContent = d.step2;
  document.getElementById('t_process_btn').textContent = d.process;
  document.getElementById('t_download').textContent = d.download;
  document.getElementById('t_download2').textContent = d.download;
  document.getElementById('t_footer').textContent = d.footer;
  document.getElementById('t_draft_label').textContent = d.draftLabel;
  document.getElementById('t_filters_title').textContent = d.filtersTitle;
  document.getElementById('t_download_filters').textContent = d.downloadFilters;

  // Búsqueda
  document.getElementById('t_step3').textContent = d.step3;
  document.getElementById('t_search_hint').innerHTML = d.searchHint;
  document.getElementById('t_search_btn').textContent = d.searchBtn;
  const input = document.getElementById('searchInput');
  input.placeholder = d.searchPlaceholder;

  // Tablas vacías
  const t1 = document.getElementById('table1');
  const t2 = document.getElementById('table2');
  const st = document.getElementById('searchTable');
  if (t1.classList.contains('empty')) t1.textContent = d.empty;
  if (t2.classList.contains('empty')) t2.textContent = d.empty;
  if (st.classList.contains('empty')) st.textContent = d.empty;

    // Mostrar instrucciones según idioma
  const listEn = document.getElementById('t_instructions_list');
  const listEs = document.getElementById('t_instructions_list_es');
  if (listEn && listEs) {
    if (lang === 'en') {
      listEn.style.display = 'block';
      listEs.style.display = 'none';
    } else {
      listEn.style.display = 'none';
      listEs.style.display = 'block';
    }
  }

}

(function initLang(){
  const initial = window.__APP_LANG__ || 'es';
  const select = document.getElementById('langSelect');
  applyLang(initial);
  select.value = initial;
  select.addEventListener('change', () => {
    const val = select.value;
    applyLang(val);
    const url = new URL(window.location.href);
    url.searchParams.set('lang', val);
    window.history.replaceState({}, '', url.toString());
  });
})();

// ---------- Conteo de filas para archivos en Sección 1 ----------
const fileDraft = document.getElementById('fileDraft');
const fileSpecialty = document.getElementById('fileSpecialty');

const fileCountDraft = document.getElementById('fileCountDraft');
const fileCountSpecialty = document.getElementById('fileCountSpecialty');

// Conteo simple de filas: cuenta líneas no vacías menos 1 (header)
// Nota: si hay saltos de línea dentro de campos entre comillas muy complejos,
// este conteo podría variar; para uso rápido de UI funciona bien.
function estimateCsvRowCount(text) {
  if (!text) return 0;
  const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (lines.length === 0) return 0;
  return Math.max(0, lines.length - 1); // restar encabezado
}

// V2: Draft puede ser .xlsx/.xls, así que el conteo rápido en el navegador
// usa SheetJS para esos formatos, y el conteo por líneas para .csv
function estimateRowCountForFile(file, callback) {
  const name = (file.name || '').toLowerCase();
  const isExcel = name.endsWith('.xlsx') || name.endsWith('.xls');

  if (isExcel && window.XLSX) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });
        callback(rows.length);
      } catch (e) {
        callback(0);
      }
    };
    reader.onerror = () => callback(0);
    reader.readAsArrayBuffer(file);
    return;
  }

  const reader = new FileReader();
  reader.onload = () => callback(estimateCsvRowCount(reader.result || ''));
  reader.onerror = () => callback(0);
  reader.readAsText(file);
}

function showFileRowCount(targetEl, n) {
  const lang = (document.getElementById('langSelect').value || 'es');
  targetEl.textContent = formatCount(lang, n);
}

function resetFileRowCount(targetEl) {
  targetEl.textContent = '—';
}

function handleFileChange(inputEl, countEl) {
  const file = inputEl.files && inputEl.files[0];
  if (!file) { resetFileRowCount(countEl); return; }
  estimateRowCountForFile(file, (n) => showFileRowCount(countEl, n));
}

fileDraft.addEventListener('change', () => handleFileChange(fileDraft, fileCountDraft));
fileSpecialty.addEventListener('change', () => handleFileChange(fileSpecialty, fileCountSpecialty));

// Si cambia el idioma, re-renderiza textos de conteo actuales si son números
(function hookLangToFileCounts(){
  const origApplyLang = applyLang;
  window.applyLang = function(lang){
    origApplyLang(lang);
    // Reaplicar si hay números
    [fileCountDraft, fileCountSpecialty, document.getElementById('filtersCount')].forEach(el=>{
      if (!el) return;
      const n = parseInt(el.textContent);
      if (!isNaN(n)) el.textContent = formatCount(lang, n);
    });
  };
})();


// ---------- lógica de app ----------
const form = document.getElementById('csvForm');
const statusEl = document.getElementById('status');

const table1 = document.getElementById('table1');
const table2 = document.getElementById('table2');

const downloadCsv1 = document.getElementById('downloadCsv1');
const downloadCsv2 = document.getElementById('downloadCsv2');

// V2: sección de Filters generado
const filtersGenerated = document.getElementById('filtersGenerated');
const filtersCountEl = document.getElementById('filtersCount');
const downloadFilters = document.getElementById('downloadFilters');

// Búsqueda
const searchInput = document.getElementById('searchInput');
const searchBtn = document.getElementById('searchBtn');
const searchTable = document.getElementById('searchTable');

// Índice de Draft (se llena tras /process)
let DRAFT_INDEX = [];

function renderTable(el, rows) {
  el.innerHTML = '';
  el.classList.remove('empty');

  if (!rows || rows.length === 0) {
    const lang = (document.getElementById('langSelect').value || 'es');
    el.textContent = i18n[lang].empty;
    el.classList.add('empty');
    return;
  }

  const headers = Object.keys(rows[0]);
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const trh = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  rows.forEach(r => {
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const td = document.createElement('td');
      td.textContent = r[h] ?? '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  el.appendChild(table);
}

// Datagrid para búsqueda: fuerza columnas solicitadas
function renderSearchTable(rows) {
  const cols = ['PACE MRN', 'Patient', 'Patient Address', 'Time'];
  searchTable.innerHTML = '';
  searchTable.classList.remove('empty');

  if (!rows || rows.length === 0) {
    const lang = (document.getElementById('langSelect').value || 'es');
    searchTable.textContent = i18n[lang].empty;
    searchTable.classList.add('empty');
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const trh = document.createElement('tr');
  cols.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  rows.forEach(r => {
    const tr = document.createElement('tr');
    cols.forEach(h => {
      const td = document.createElement('td');
      td.textContent = r[h] ?? '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  searchTable.appendChild(table);
}

function runSearchSingle(query) {
  if (!query || !DRAFT_INDEX || DRAFT_INDEX.length === 0) return [];
  const q = String(query).toLowerCase().trim();
  return DRAFT_INDEX.filter(r => {
    const pm = String(r['PACE MRN'] ?? '').toLowerCase();
    const pt = String(r['Patient'] ?? '').toLowerCase();
    return pm.includes(q) || pt.includes(q);
  });
}

// Búsqueda combinada (unión de varios criterios)
function runSearchCombined(queries) {
  const seen = new Set();
  const out = [];
  for (const q of queries) {
    const rows = runSearchSingle(q);
    for (const r of rows) {
      const key = `${r['PACE MRN']}|${r['Patient']}|${r['Patient Address']}|${r['Time']}`;
      if (!seen.has(key)) { seen.add(key); out.push(r); }
    }
  }
  return out;
}

searchBtn.addEventListener('click', () => {
  const val = searchInput.value || '';
  const rows = runSearchSingle(val);
  renderSearchTable(rows);
});

function csvToBlobUrl(csvText) {
  const blob = new Blob([csvText || ''], { type: 'text/csv;charset=utf-8;' });
  return URL.createObjectURL(blob);
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const lang = (document.getElementById('langSelect').value || 'es');
  statusEl.textContent = i18n[lang].processing;

  // Limpieza previa
  downloadCsv1.style.display = 'none';
  downloadCsv2.style.display = 'none';
  filtersGenerated.style.display = 'none';
  resetFileRowCount(filtersCountEl);
  table1.classList.add('empty');
  table2.classList.add('empty');
  searchTable.classList.add('empty');
  table1.textContent = i18n[lang].empty;
  table2.textContent = i18n[lang].empty;
  searchTable.textContent = i18n[lang].empty;

  try {
    const fd = new FormData(form);
    const res = await fetch('/process', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unexpected error');

    // Tablas de resultados
    renderTable(table1, data.trips_In_Filters_Not_In_Specialty);
    renderTable(table2, data.trips_In_Specialty_Not_In_Filters);
    //show row count
    const c1 = data.trips_In_Filters_Not_In_Specialty?.length || 0;
    const c2 = data.trips_In_Specialty_Not_In_Filters?.length || 0;
    document.getElementById('count1').textContent = `${c1} ${c1 === 1 ? 'trip' : 'trips'}`;
    document.getElementById('count2').textContent = `${c2} ${c2 === 1 ? 'trip' : 'trips'}`;

    // Enlaces de descarga con prefijo
    const prefix = data?.meta?.visitDatePrefix || '';
    if (data.csvFiles?.trips_In_Filters_Not_In_Specialty) {
      const url1 = csvToBlobUrl(data.csvFiles.trips_In_Filters_Not_In_Specialty);
      downloadCsv1.href = url1;
      downloadCsv1.download = `${prefix}trips_In_Filters_Not_In_Specialty.csv`;
      downloadCsv1.style.display = 'inline-block';
    }
    if (data.csvFiles?.trips_In_Specialty_Not_In_Filters) {
      const url2 = csvToBlobUrl(data.csvFiles.trips_In_Specialty_Not_In_Filters);
      downloadCsv2.href = url2;
      downloadCsv2.download = `${prefix}trips_In_Specialty_Not_In_Filters.csv`;
      downloadCsv2.style.display = 'inline-block';
    }

    // V2: Filters.csv generado automáticamente
    if (typeof data.filtersCsv === 'string') {
      const urlF = csvToBlobUrl(data.filtersCsv);
      downloadFilters.href = urlF;
      downloadFilters.download = `${prefix}Filters.csv`;
      const fc = data.filtersRowCount || 0;
      showFileRowCount(filtersCountEl, fc);
      filtersGenerated.style.display = 'block';
    }

    // Índice para búsqueda
    DRAFT_INDEX = data?.meta?.draftIndex || [];

    // --- BÚSQUEDA AUTOMÁTICA TRAS "Procesar" ---
    // Criterios: Patient="Caro" y PACE MRN="SDP10035365"
    const autoRows = runSearchCombined(['Caro', 'SDP10035365']);
    renderSearchTable(autoRows);

    statusEl.textContent = i18n[lang].done
  } catch (err) {
    const lang2 = (document.getElementById('langSelect').value || 'es');
    statusEl.textContent = i18n[lang2].errorPrefix + err.message;
    console.error(err);
  }
});
