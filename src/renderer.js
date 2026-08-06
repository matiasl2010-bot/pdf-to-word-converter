let selectedFiles = [];
let outputFolder = null;
let engines = { word: null, soffice: null };

const dropzone = document.getElementById('dropzone');
const fileListEl = document.getElementById('file-list');
const fileCountEl = document.getElementById('file-count');
const btnConvert = document.getElementById('btn-convert');
const outputPathEl = document.getElementById('output-path');
const sofficeWarning = document.getElementById('soffice-warning');
const noEnginesWarning = document.getElementById('no-engines-warning');
const engineInfo = document.getElementById('engine-info');
const progressWrap = document.getElementById('progress-wrap');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const resultsSection = document.getElementById('results');
const resultsList = document.getElementById('results-list');

function isPdfFile(name) {
  return /\.pdf$/i.test(name);
}

function renderFileList() {
  fileListEl.innerHTML = '';
  selectedFiles.forEach((f, idx) => {
    const li = document.createElement('li');
    const name = f.split(/[\\/]/).pop();

    // textContent y no innerHTML: el nombre viene de un archivo del disco y en
    // Linux/macOS puede contener HTML (en Windows el SO ya prohibe < y >).
    const spanNombre = document.createElement('span');
    spanNombre.textContent = name;

    const spanQuitar = document.createElement('span');
    spanQuitar.className = 'remove';
    spanQuitar.dataset.idx = idx;
    spanQuitar.textContent = '✕';

    li.append(spanNombre, spanQuitar);
    fileListEl.appendChild(li);
  });
  fileCountEl.textContent = `${selectedFiles.length} archivo${selectedFiles.length === 1 ? '' : 's'}`;
  updateConvertButtonState();
}

function updateConvertButtonState() {
  const hasEngine = !!(engines.word || engines.soffice);
  btnConvert.disabled = selectedFiles.length === 0 || !hasEngine;
}

fileListEl.addEventListener('click', (e) => {
  if (e.target.classList.contains('remove')) {
    const idx = Number(e.target.dataset.idx);
    selectedFiles.splice(idx, 1);
    renderFileList();
  }
});

document.getElementById('btn-clear').addEventListener('click', () => {
  selectedFiles = [];
  renderFileList();
  resultsSection.classList.add('hidden');
});

// --- Selección de archivos/carpeta ---

document.getElementById('btn-select-files').addEventListener('click', async () => {
  const files = await window.api.selectFiles();
  addFiles(files);
});

document.getElementById('btn-select-folder').addEventListener('click', async () => {
  const files = await window.api.selectFolder();
  addFiles(files);
});

function addFiles(paths) {
  for (const p of paths) {
    if (!selectedFiles.includes(p)) selectedFiles.push(p);
  }
  renderFileList();
}

// --- Drag & drop ---

['dragenter', 'dragover'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });
});

['dragleave', 'drop'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
  });
});

dropzone.addEventListener('drop', (e) => {
  const paths = [];
  for (const file of e.dataTransfer.files) {
    if (isPdfFile(file.name)) paths.push(file.path);
  }
  addFiles(paths);
});

// --- Carpeta de salida ---

document.getElementById('btn-select-output').addEventListener('click', async () => {
  const folder = await window.api.selectOutputFolder();
  if (folder) {
    outputFolder = folder;
    outputPathEl.textContent = folder;
  }
});

// --- Detección de motores (Word / LibreOffice) ---

function renderEngineInfo() {
  if (engines.word) {
    engineInfo.classList.remove('hidden');
    engineInfo.innerHTML = `<span class="dot"></span> Se usará <strong>Microsoft Word</strong> (mejor reconstrucción del texto editable)${engines.soffice ? ' — LibreOffice queda como respaldo.' : '.'}`;
    sofficeWarning.classList.toggle('hidden', !!engines.soffice);
    noEnginesWarning.classList.add('hidden');
  } else if (engines.soffice) {
    engineInfo.classList.remove('hidden');
    engineInfo.innerHTML = `<span class="dot"></span> Se usará <strong>LibreOffice</strong> (no se encontró Word instalado).`;
    sofficeWarning.classList.add('hidden');
    noEnginesWarning.classList.add('hidden');
  } else {
    engineInfo.classList.add('hidden');
    sofficeWarning.classList.add('hidden');
    noEnginesWarning.classList.remove('hidden');
  }
  updateConvertButtonState();
}

async function checkEngines() {
  engines = await window.api.checkEngines();
  renderEngineInfo();
}

document.getElementById('btn-download-lo').addEventListener('click', () => {
  window.open('https://www.libreoffice.org/download/download/', '_blank');
});

document.getElementById('btn-locate-lo').addEventListener('click', async () => {
  const chosen = await window.api.selectSofficeManually();
  if (chosen) {
    engines.soffice = chosen;
    renderEngineInfo();
  }
});

checkEngines();

// --- Conversión ---

window.api.onProgress(({ current, total, fileName }) => {
  const pct = Math.round((current / total) * 100);
  progressFill.style.width = pct + '%';
  progressText.textContent = `Convirtiendo ${current}/${total}: ${fileName}`;
});

document.getElementById('btn-convert').addEventListener('click', async () => {
  if (selectedFiles.length === 0) return;
  if (!engines.word && !engines.soffice) {
    await checkEngines();
    if (!engines.word && !engines.soffice) return;
  }

  btnConvert.disabled = true;
  progressWrap.classList.remove('hidden');
  resultsSection.classList.add('hidden');
  progressFill.style.width = '0%';

  // Las rutas de Word/LibreOffice las resuelve el proceso main: mandarlas desde
  // aca seria decirle que ejecutable lanzar.
  const results = await window.api.convertFiles({
    files: selectedFiles,
    outputDir: outputFolder || null,
    preferWord: !!engines.word
  });

  progressWrap.classList.add('hidden');
  btnConvert.disabled = false;
  showResults(results);
});

function showResults(results) {
  resultsList.innerHTML = '';
  results.forEach((r) => {
    const li = document.createElement('li');
    const name = r.file.split(/[\\/]/).pop();
    const engineLabel = r.engine === 'word' ? 'Word' : r.engine === 'libreoffice' ? 'LibreOffice' : '';

    // Igual que en la lista de archivos: el nombre se inserta como texto, nunca
    // como HTML. El icono y el motor si son literales del codigo.
    const icono = document.createElement('span');
    icono.className = r.success ? 'ok' : 'fail';
    icono.textContent = r.success ? '✓' : '✕';

    const texto = document.createElement('span');
    texto.textContent = r.success ? name : `${name} — error`;

    li.append(icono, texto);

    if (engineLabel) {
      const motor = document.createElement('span');
      motor.style.color = '#6c6f7a';
      motor.textContent = ` (${engineLabel})`;
      texto.appendChild(motor);
    }

    resultsList.appendChild(li);
  });
  resultsSection.classList.remove('hidden');
}

document.getElementById('btn-open-output').addEventListener('click', () => {
  const folder = outputFolder || (selectedFiles[0] ? selectedFiles[0].substring(0, selectedFiles[0].lastIndexOf(selectedFiles[0].includes('\\') ? '\\' : '/')) : null);
  if (folder) window.api.openFolder(folder);
});
