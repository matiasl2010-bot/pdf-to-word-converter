const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');
const Store = require('electron-store');

const store = new Store();

let mainWindow;

// Rutas comunes donde suele instalarse LibreOffice
const COMMON_SOFFICE_PATHS = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/bin/libreoffice',
  '/Applications/LibreOffice.app/Contents/MacOS/soffice'
];

function findSoffice() {
  const saved = store.get('sofficePath');
  if (saved && fs.existsSync(saved)) return saved;
  for (const p of COMMON_SOFFICE_PATHS) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

// Busca Word instalado consultando el registro de Windows (rápido, no lanza Word)
function findWordViaRegistry() {
  return new Promise((resolve) => {
    if (process.platform !== 'win32') return resolve(null);

    const keys = [
      'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE',
      'HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE'
    ];
    let i = 0;

    function tryNext() {
      if (i >= keys.length) return resolve(null);
      const key = keys[i++];
      const proc = spawn('reg', ['query', key, '/ve']);
      let out = '';
      proc.stdout.on('data', (d) => { out += d.toString(); });
      proc.on('close', () => {
        const match = out.match(/REG_SZ\s+(.+)/);
        if (match) {
          const exePath = match[1].trim();
          if (fs.existsSync(exePath)) return resolve(exePath);
        }
        tryNext();
      });
      proc.on('error', () => tryNext());
    }
    tryNext();
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 720,
    height: 700,
    resizable: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.setMenuBarVisibility(false);
  mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ---- IPC: detección de motores disponibles ----

ipcMain.handle('check-engines', async () => {
  const [wordPath, sofficePath] = await Promise.all([
    findWordViaRegistry(),
    Promise.resolve(findSoffice())
  ]);
  return { word: wordPath, soffice: sofficePath };
});

ipcMain.handle('select-soffice-manually', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Ubicá el ejecutable soffice.exe de LibreOffice',
    properties: ['openFile'],
    filters: [{ name: 'Ejecutable', extensions: ['exe', ''] }]
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const chosen = result.filePaths[0];
  store.set('sofficePath', chosen);
  return chosen;
});

// ---- IPC: selección de archivos ----

ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Seleccioná archivos PDF',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Documentos PDF', extensions: ['pdf'] }]
  });
  if (result.canceled) return [];
  return result.filePaths;
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Seleccioná una carpeta (se buscarán .pdf dentro, incluyendo subcarpetas)',
    properties: ['openDirectory']
  });
  if (result.canceled) return [];

  const folder = result.filePaths[0];
  const files = [];

  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(full);
      } else if (/\.pdf$/i.test(entry.name)) {
        files.push(full);
      }
    }
  }
  walk(folder);
  return files;
});

ipcMain.handle('select-output-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Seleccioná dónde guardar los Word',
    properties: ['openDirectory']
  });
  if (result.canceled) return null;
  return result.filePaths[0];
});

ipcMain.handle('open-folder', async (_event, folderPath) => {
  shell.openPath(folderPath);
});

// ---- Conversión ----

function outputDocxPathFor(filePath, outDir) {
  const dir = outDir || path.dirname(filePath);
  const base = path.parse(filePath).name;
  return path.join(dir, base + '.docx');
}

// Convierte un archivo con soffice --headless
function convertOneWithSoffice(sofficePath, filePath, outDir) {
  return new Promise((resolve) => {
    const dir = outDir || path.dirname(filePath);
    const args = ['--headless', '--norestore', '--convert-to', 'docx', '--outdir', dir, filePath];
    const proc = spawn(sofficePath, args);
    let stderr = '';
    proc.stderr.on('data', (d) => { stderr += d.toString(); });
    proc.on('close', (code) => {
      resolve({ file: filePath, success: code === 0, error: code === 0 ? null : stderr, engine: 'libreoffice' });
    });
    proc.on('error', (err) => {
      resolve({ file: filePath, success: false, error: err.message, engine: 'libreoffice' });
    });
  });
}

// Convierte un lote entero con Word (una sola instancia de Word para todo el lote)
function convertBatchWithWord(files, outputDir, onProgress) {
  return new Promise((resolve) => {
    const jobs = files.map((f) => ({ input: f, output: outputDocxPathFor(f, outputDir) }));
    const tmpFile = path.join(os.tmpdir(), `pdf-word-job-${Date.now()}.json`);
    fs.writeFileSync(tmpFile, JSON.stringify(jobs));

    const scriptPath = path.join(__dirname, 'scripts', 'convert-with-word.ps1');
    const proc = spawn('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath, '-JobFile', tmpFile
    ]);

    let buffer = '';
    let finalResult = null;
    let fatal = null;

    proc.stdout.on('data', (d) => {
      buffer += d.toString();
      let lines = buffer.split(/\r?\n/);
      buffer = lines.pop();
      for (const line of lines) {
        if (line.startsWith('PROGRESS:')) {
          const [, current, total, fileName] = line.split(':');
          onProgress && onProgress({ current: Number(current), total: Number(total), fileName: path.basename(fileName || '') });
        } else if (line.startsWith('RESULT:')) {
          try {
            finalResult = JSON.parse(line.slice('RESULT:'.length));
            if (!Array.isArray(finalResult)) finalResult = [finalResult];
          } catch (e) {
            finalResult = null;
          }
        } else if (line.startsWith('FATAL:')) {
          fatal = line.slice('FATAL:'.length);
        }
      }
    });

    proc.on('close', () => {
      try { fs.unlinkSync(tmpFile); } catch (e) {}
      if (fatal || !finalResult) {
        resolve({ ok: false, reason: fatal || 'sin resultado de Word' });
      } else {
        resolve({ ok: true, results: finalResult.map((r) => ({ ...r, engine: 'word' })) });
      }
    });

    proc.on('error', (err) => {
      try { fs.unlinkSync(tmpFile); } catch (e) {}
      resolve({ ok: false, reason: err.message });
    });
  });
}

ipcMain.handle('convert-files', async (event, { files, outputDir, preferWord, wordPath, sofficePath }) => {
  const send = (progress) => event.sender.send('conversion-progress', progress);
  let results = [];

  if (preferWord && wordPath) {
    const batch = await convertBatchWithWord(files, outputDir, send);
    if (batch.ok) {
      results = batch.results;
    } else if (sofficePath) {
      for (let i = 0; i < files.length; i++) {
        send({ current: i + 1, total: files.length, fileName: path.basename(files[i]) });
        const r = await convertOneWithSoffice(sofficePath, files[i], outputDir);
        results.push(r);
      }
    } else {
      results = files.map((f) => ({ file: f, success: false, error: `Word falló (${batch.reason}) y no hay LibreOffice de respaldo`, engine: 'ninguno' }));
    }

    if (sofficePath) {
      for (let i = 0; i < results.length; i++) {
        if (!results[i].success && results[i].engine === 'word') {
          const retry = await convertOneWithSoffice(sofficePath, results[i].file, outputDir);
          if (retry.success) results[i] = retry;
        }
      }
    }
  } else if (sofficePath) {
    for (let i = 0; i < files.length; i++) {
      send({ current: i + 1, total: files.length, fileName: path.basename(files[i]) });
      const r = await convertOneWithSoffice(sofficePath, files[i], outputDir);
      results.push(r);
    }
  } else {
    results = files.map((f) => ({ file: f, success: false, error: 'No hay Word ni LibreOffice disponibles', engine: 'ninguno' }));
  }

  return results;
});
