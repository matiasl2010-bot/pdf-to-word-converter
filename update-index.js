// Regenera index.md con la estructura de carpetas/archivos del proyecto.
// Uso: node update-index.js  (o: npm run update-index)

const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const IGNORE = new Set(['node_modules', 'dist', '.git', 'index.md']);

function walk(dir, prefix = '') {
  let lines = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true })
    .filter((e) => !IGNORE.has(e.name))
    .sort((a, b) => {
      if (a.isDirectory() !== b.isDirectory()) return a.isDirectory() ? -1 : 1;
      return a.name.localeCompare(b.name);
    });

  entries.forEach((entry, i) => {
    const isLast = i === entries.length - 1;
    const connector = isLast ? '└── ' : '├── ';
    lines.push(prefix + connector + entry.name + (entry.isDirectory() ? '/' : ''));

    if (entry.isDirectory()) {
      const nextPrefix = prefix + (isLast ? '    ' : '│   ');
      lines = lines.concat(walk(path.join(dir, entry.name), nextPrefix));
    }
  });

  return lines;
}

const tree = walk(ROOT).join('\n');

const content = `# pdf-to-word-converter

App de Electron (instalador NSIS + portable) para convertir archivos PDF a
Word (.docx) en lote. Usa Microsoft Word (automatización COM, función de
reflow de PDF) como motor preferido si está instalado; si no hay Word, o si
un archivo puntual falla, cae automáticamente a LibreOffice headless.

**No hace OCR**: solo funciona bien con PDFs nativos (texto real), no con
escaneos o fotos.

Generado/actualizado automáticamente por \`update-index.js\` — no editar a mano
la sección de estructura, correr \`npm run update-index\` después de agregar
o quitar archivos.

## Estructura

\`\`\`
${tree}
\`\`\`

## Componentes clave

- **main.js** — proceso principal de Electron: detecta Word (registro de
  Windows) y LibreOffice (soffice), maneja diálogos de selección de
  archivos/carpeta/salida, y orquesta la conversión con fallback automático
  Word → LibreOffice.
- **scripts/convert-with-word.ps1** — abre una única instancia de Word vía
  COM para todo el lote; al abrir cada PDF, Word dispara su conversión
  automática a texto editable, y el script guarda el resultado con
  \`SaveAs2\` en formato .docx.
- **preload.js** — expone la API segura (\`window.api\`) al renderer vía
  contextBridge, sin nodeIntegration.
- **src/index.html / styles.css / renderer.js** — interfaz: drag&drop de
  PDFs, selección de carpeta, indicador de motor en uso, aviso sobre
  limitación de OCR, barra de progreso, resultados con motor por archivo.
- **update-index.js** — este script; regenera la estructura de arriba.

## Requisitos externos

- **Word** (opcional, recomendado): detección automática vía registro de
  Windows.
- **LibreOffice** (gratuito): motor único si no hay Word, o respaldo
  automático si Word falla con algún archivo.

## Build

- \`npm start\` — correr en modo desarrollo
- \`npm run dist\` — generar instalador (NSIS) + versión portable para Windows
- \`npm run dist:portable\` — solo portable
- \`npm run dist:installer\` — solo instalador
`;

fs.writeFileSync(path.join(ROOT, 'index.md'), content);
console.log('index.md actualizado.');
