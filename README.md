# PDF a Word — app de escritorio

App de Electron para convertir archivos PDF a Word (.docx), individual o en
lote, con drag&drop. Genera un **instalador** y una **versión portable**
para Windows.

## ⚠️ Importante: sin OCR

Esta app **no reconoce texto en imágenes**. Funciona bien con PDFs "nativos"
(los que salen de exportar un Word, Google Docs, LaTeX, etc., donde el texto
ya es texto real dentro del PDF). Si el PDF es un **escaneo o una foto**, el
resultado va a salir vacío, o con texto ilegible/basura — ni Word ni
LibreOffice tienen OCR integrado por defecto. Para PDFs escaneados haría
falta una herramienta distinta (con motor OCR tipo Tesseract).

## 1. Requisitos previos

- **Node.js** (18 o superior) instalado en tu PC.
- **Microsoft Word** (opcional pero recomendado): al abrir un PDF, Word
  dispara su función de "PDF Reflow" (reconstruye el documento como texto
  editable, con más fidelidad de estructura que LibreOffice en la mayoría de
  los casos). La app lo usa como motor preferido si está instalado.
- **LibreOffice** (gratis, https://www.libreoffice.org/download/download/):
  se usa si no hay Word, o como respaldo automático si Word falla con algún
  archivo puntual.

## 2. Poner el proyecto en marcha

abrís una
terminal ahí y corrés:

```bash
npm install
npm start
```

## 3. Generar el instalador y la versión portable

```bash
npm run dist
```

Esto crea en la carpeta `dist/`:

- `PDF a Word Setup 1.0.0.exe` → instalador (NSIS), con acceso directo de escritorio.
- `PDF a Word 1.0.0.exe` → versión portable.

```bash
npm run dist:installer   # solo el instalador
npm run dist:portable    # solo el portable
```

## 4. Cómo se usa la app

1. Arrastrás archivos `.pdf` a la zona de drop, o elegís archivos/carpeta
   (busca recursivamente en subcarpetas).
2. Opcional: elegís una carpeta de salida fija. Si no elegís ninguna, cada
   .docx se guarda al lado de su PDF original.
3. Apretás **Convertir a Word**. Barra de progreso archivo por archivo.
4. Resultado final con ✓/✕ y motor usado (Word o LibreOffice) por archivo.

## 5. Mantenimiento del índice del proyecto

```bash
npm run update-index
```

Regenera `index.md` con la estructura actual del proyecto.

## Notas técnicas

- **Motor Word**: automatización COM vía `scripts/convert-with-word.ps1`.
  Abre una única instancia de Word para todo el lote, deja que Word haga su
  conversión automática de PDF a texto editable al abrir el archivo, y
  guarda con `SaveAs2` en formato .docx (`wdFormatXMLDocument`).
- **Motor LibreOffice**: `soffice --headless --convert-to docx`. Suele
  reconstruir el texto en cuadros/frames en vez de párrafos fluidos, así que
  el documento resultante es editable pero con más trabajo de limpieza que
  con Word.
- **Selección automática**: Word como principal si está disponible; fallback
  automático a LibreOffice archivo por archivo si Word falla.
- Documentos con layouts muy complejos (columnas, tablas, texto superpuesto
  a imágenes) van a necesitar una revisión manual después de convertir, sea
  cual sea el motor usado — es una limitación inherente a "PDF a Word", no
  del motor en particular.
- No sube archivos a ningún servidor — todo corre localmente.
