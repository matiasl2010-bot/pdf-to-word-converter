# pdf-to-word-converter

App de Electron (instalador NSIS + portable) para convertir archivos PDF a
Word (.docx) en lote. Usa Microsoft Word (automatización COM, función de
reflow de PDF) como motor preferido si está instalado; si no hay Word, o si
un archivo puntual falla, cae automáticamente a LibreOffice headless.

**No hace OCR**: solo funciona bien con PDFs nativos (texto real), no con
escaneos o fotos.

Generado/actualizado automáticamente por `update-index.js` — no editar a mano
la sección de estructura, correr `npm run update-index` después de agregar
o quitar archivos.

## Estructura

```
├── scripts/
│   └── convert-with-word.ps1
├── src/
│   ├── index.html
│   ├── renderer.js
│   └── styles.css
├── main.js
├── package.json
├── preload.js
├── README.md
└── update-index.js
```

## Componentes clave

- **main.js** — proceso principal de Electron: detecta Word (registro de
  Windows) y LibreOffice (soffice), maneja diálogos de selección de
  archivos/carpeta/salida, y orquesta la conversión con fallback automático
  Word → LibreOffice.
- **scripts/convert-with-word.ps1** — abre una única instancia de Word vía
  COM para todo el lote; al abrir cada PDF, Word dispara su conversión
  automática a texto editable, y el script guarda el resultado con
  `SaveAs2` en formato .docx.
- **preload.js** — expone la API segura (`window.api`) al renderer vía
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

- `npm start` — correr en modo desarrollo
- `npm run dist` — generar instalador (NSIS) + versión portable para Windows
- `npm run dist:portable` — solo portable
- `npm run dist:installer` — solo instalador
