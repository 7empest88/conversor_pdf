from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pythoncom
import win32com.client as win32
from flask import Flask, flash, redirect, render_template_string, request, send_file, url_for
from pypdf import PdfReader, PdfWriter
from werkzeug.utils import secure_filename


# Constantes de Microsoft Office.
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_PORTRAIT = 1
XL_PAPER_LETTER = 1
XL_SHEET_VISIBLE = -1
WORD_EXPORT_FORMAT_PDF = 17
POWERPOINT_FIXED_FORMAT_PDF = 2
POWERPOINT_INTENT_PRINT = 2

EXCEL_EXTENSIONS = {"xlsx", "xls", "xlsm", "xlsb"}
WORD_EXTENSIONS = {"docx", "doc", "docm", "dotx", "dotm"}
POWERPOINT_EXTENSIONS = {"pptx", "ppt", "pptm", "potx", "potm"}
PDF_EXTENSIONS = {"pdf"}
ALLOWED_EXTENSIONS = EXCEL_EXTENSIONS | WORD_EXTENSIONS | POWERPOINT_EXTENSIONS | PDF_EXTENSIONS

MAX_FILE_SIZE = 100 * 1024 * 1024
MAX_FILES = 30

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "cambia-esta-clave-local")
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE * MAX_FILES

OUTPUT_DIR = Path(tempfile.gettempdir()) / "conversor_documentos_pdf_outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


HTML = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Conversor local de documentos a PDF</title>
  <style>
    :root {
      --bg: #f3f6fb;
      --card: #ffffff;
      --text: #172033;
      --muted: #5e6c80;
      --primary: #1467dd;
      --primary-hover: #0e52b6;
      --border: #d9e2f0;
      --soft-blue: #f7fbff;
      --success-bg: #eaf8ef;
      --success-border: #9bd8ad;
      --error-bg: #fff1f1;
      --error-border: #efb0b0;
      --warning-bg: #fff8df;
      --warning-border: #ecd888;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: "Segoe UI", Roboto, Arial, sans-serif;
      line-height: 1.45;
    }

    main {
      width: min(920px, calc(100% - 32px));
      margin: 40px auto;
    }

    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 16px;
      box-shadow: 0 14px 35px rgba(35, 57, 91, .10);
      overflow: hidden;
    }

    header {
      padding: 29px 32px 23px;
      border-bottom: 1px solid var(--border);
    }

    h1 { margin: 0; font-size: clamp(1.45rem, 3vw, 2rem); }
    h2 { margin: 0 0 8px; font-size: 1.04rem; }

    .subtitle { margin: 8px 0 0; color: var(--muted); }
    .content { padding: 27px 32px 32px; }

    .notice {
      padding: 13px 15px;
      border-radius: 10px;
      margin-bottom: 17px;
    }
    .error { background: var(--error-bg); border: 1px solid var(--error-border); }
    .success { background: var(--success-bg); border: 1px solid var(--success-border); }
    .warning { background: var(--warning-bg); border: 1px solid var(--warning-border); }

    .config {
      background: var(--soft-blue);
      border: 1px solid #d7e7fb;
      border-radius: 12px;
      padding: 16px 19px;
      margin-bottom: 23px;
    }
    .config ul { margin: 0; padding-left: 20px; color: #3e4d63; }

    .upload-zone {
      padding: 19px;
      border: 2px dashed #afc1da;
      border-radius: 12px;
      background: #fbfdff;
    }
    .upload-zone.dragging { border-color: var(--primary); background: #edf5ff; }

    label { display: block; font-weight: 700; margin-bottom: 8px; }
    input[type=file] { width: 100%; color: #3b4a60; }
    .hint { margin: 10px 0 0; color: var(--muted); font-size: .92rem; }

    #lista-archivos {
      list-style: none;
      margin: 19px 0 0;
      padding: 0;
      display: grid;
      gap: 9px;
    }
    #lista-archivos:empty { display: none; }

    .archivo-item {
      display: flex;
      align-items: center;
      gap: 11px;
      border: 1px solid var(--border);
      border-radius: 10px;
      background: #fff;
      padding: 10px 12px;
    }
    .numero {
      display: grid;
      place-items: center;
      min-width: 28px;
      height: 28px;
      border-radius: 50%;
      background: #eaf2ff;
      color: #174f9c;
      font-weight: 800;
      font-size: .87rem;
    }
    .archivo-nombre { flex: 1; min-width: 0; overflow-wrap: anywhere; }
    .archivo-info { color: var(--muted); font-size: .86rem; white-space: nowrap; }

    .small-button {
      border: 1px solid #bfcce0;
      background: #fff;
      color: #33435a;
      border-radius: 7px;
      min-width: 34px;
      height: 32px;
      cursor: pointer;
      font: inherit;
      font-weight: 800;
    }
    .small-button:hover:not(:disabled) { background: #edf4ff; }
    .small-button.remove:hover { background: #fff0f0; color: #a62929; }
    .small-button:disabled { opacity: .4; cursor: not-allowed; }

    .action {
      width: 100%;
      border: 0;
      border-radius: 10px;
      padding: 13px 18px;
      margin-top: 21px;
      background: var(--primary);
      color: #fff;
      cursor: pointer;
      font: inherit;
      font-weight: 800;
      font-size: 1rem;
    }
    .action:hover:not(:disabled) { background: var(--primary-hover); }
    .action:disabled { opacity: .62; cursor: wait; }

    .result {
      margin-top: 23px;
      background: var(--success-bg);
      border: 1px solid var(--success-border);
      border-radius: 12px;
      padding: 19px;
    }
    .result p { margin: 7px 0; }
    .download {
      display: block;
      text-align: center;
      text-decoration: none;
      margin-top: 15px;
      border-radius: 10px;
      padding: 12px;
      background: var(--primary);
      color: #fff;
      font-weight: 800;
    }
    .download:hover { background: var(--primary-hover); }
    code { overflow-wrap: anywhere; }

    footer { color: var(--muted); text-align: center; font-size: .88rem; margin-top: 16px; }

    @media (max-width: 650px) {
      main { width: min(100% - 18px, 920px); margin: 18px auto; }
      header, .content { padding-left: 18px; padding-right: 18px; }
      .archivo-item { gap: 7px; padding: 9px; }
      .archivo-info { display: none; }
      .small-button { min-width: 30px; }
    }
  </style>
</head>
<body>
  <main>
    <section class="card">
      <header>
        <h1>📚 Conversor local de documentos a PDF</h1>
        <p class="subtitle">Carga, ordena y concatena documentos en un único PDF, sin separadores entre archivos.</p>
      </header>

      <div class="content">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% for category, message in messages %}
            <div class="notice {{ category }}">{{ message }}</div>
          {% endfor %}
        {% endwith %}

        <section class="config">
          <h2>Reglas de conversión</h2>
          <ul>
            <li>Excel: hojas visibles en su orden original, carta, vertical, márgenes estrechos y columnas ajustadas a una página de ancho.</li>
            <li>Word y PowerPoint: se exportan usando Microsoft Office local.</li>
            <li>PDF: se agrega sin conversión.</li>
            <li>El PDF final concatena los documentos según el orden elegido abajo, sin portadas ni páginas separadoras.</li>
          </ul>
        </section>

        <form id="form-conversion" method="post" action="{{ url_for('convertir') }}" enctype="multipart/form-data">
          <label for="archivos">Documentos a concatenar</label>
          <div class="upload-zone" id="upload-zone">
            <input id="archivos" name="archivos" type="file" multiple required
              accept=".xlsx,.xls,.xlsm,.xlsb,.docx,.doc,.docm,.dotx,.dotm,.pptx,.ppt,.pptm,.potx,.potm,.pdf">
            <p class="hint">Formatos: Excel, Word, PowerPoint y PDF. Máximo {{ max_files }} archivos; hasta {{ max_file_mb }} MB por archivo.</p>
          </div>

          <ol id="lista-archivos" aria-live="polite"></ol>
          <input id="orden-archivos" name="orden_archivos" type="hidden">

          <button class="action" id="boton-convertir" type="submit">Convertir y concatenar a PDF</button>
        </form>

        {% if resultado %}
          <section class="result">
            <h2>✅ Conversión terminada</h2>
            <p>Documentos concatenados: <strong>{{ resultado.archivos }}</strong></p>
            <p>Páginas totales del PDF: <strong>{{ resultado.paginas }}</strong></p>
            <p>Archivo generado: <code>{{ resultado.nombre_pdf }}</code></p>
            <a class="download" href="{{ url_for('descargar', token=resultado.token) }}">Descargar PDF final</a>
          </section>
        {% endif %}
      </div>
    </section>
    <footer>El procesamiento se realiza en la PC donde ejecutas el servidor Flask. Los documentos originales no se modifican.</footer>
  </main>

  <script>
  const input = document.getElementById("archivos");
  const lista = document.getElementById("lista-archivos");
  const form = document.getElementById("form-conversion");
  const boton = document.getElementById("boton-convertir");
  const zona = document.getElementById("upload-zone");

  let archivos = [];

  function formatoTamano(bytes) {
    if (bytes < 1024 * 1024) {
      return (bytes / 1024).toFixed(1) + " KB";
    }

    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
  }

  function sincronizarInputConLista() {
    const transferencia = new DataTransfer();

    archivos.forEach((archivo) => {
      transferencia.items.add(archivo);
    });

    input.files = transferencia.files;
  }

  function actualizarLista() {
    lista.innerHTML = "";

    archivos.forEach((archivo, indice) => {
      const item = document.createElement("li");
      item.className = "archivo-item";

      item.innerHTML = `
        <span class="numero">${indice + 1}</span>
        <span class="archivo-nombre">${archivo.name}</span>
        <span class="archivo-info">${formatoTamano(archivo.size)}</span>

        <button
          class="small-button"
          type="button"
          title="Subir posición"
          aria-label="Subir ${archivo.name}"
          ${indice === 0 ? "disabled" : ""}
        >↑</button>

        <button
          class="small-button"
          type="button"
          title="Bajar posición"
          aria-label="Bajar ${archivo.name}"
          ${indice === archivos.length - 1 ? "disabled" : ""}
        >↓</button>

        <button
          class="small-button remove"
          type="button"
          title="Quitar archivo"
          aria-label="Quitar ${archivo.name}"
        >×</button>
      `;

      const botones = item.querySelectorAll("button");

      botones[0].addEventListener("click", () => {
        mover(indice, indice - 1);
      });

      botones[1].addEventListener("click", () => {
        mover(indice, indice + 1);
      });

      botones[2].addEventListener("click", () => {
        archivos.splice(indice, 1);
        actualizarLista();
      });

      lista.appendChild(item);
    });

    sincronizarInputConLista();
  }

  function mover(origen, destino) {
    if (destino < 0 || destino >= archivos.length) {
      return;
    }

    [archivos[origen], archivos[destino]] = [
      archivos[destino],
      archivos[origen],
    ];

    actualizarLista();
  }

  function agregarArchivos(nuevos) {
    const nuevosArchivos = Array.from(nuevos);

    nuevosArchivos.forEach((archivo) => {
      const existe = archivos.some(
        (actual) =>
          actual.name === archivo.name &&
          actual.size === archivo.size &&
          actual.lastModified === archivo.lastModified
      );

      if (!existe) {
        archivos.push(archivo);
      }
    });

    actualizarLista();
  }

  input.addEventListener("change", (evento) => {
    agregarArchivos(evento.target.files);
  });

  ["dragenter", "dragover"].forEach((tipoEvento) => {
    zona.addEventListener(tipoEvento, (evento) => {
      evento.preventDefault();
      zona.classList.add("dragging");
    });
  });

  ["dragleave", "drop"].forEach((tipoEvento) => {
    zona.addEventListener(tipoEvento, (evento) => {
      evento.preventDefault();
      zona.classList.remove("dragging");
    });
  });

  zona.addEventListener("drop", (evento) => {
    agregarArchivos(evento.dataTransfer.files);
  });

  form.addEventListener("submit", (evento) => {
    sincronizarInputConLista();

    if (archivos.length === 0 || input.files.length === 0) {
      evento.preventDefault();
      alert("Es necesario seleccionar uno o más archivos.");
      return;
    }

    if (archivos.length > {{ max_files }}) {
      evento.preventDefault();
      alert("El máximo permitido es {{ max_files }} archivos.");
      return;
    }

    boton.disabled = true;
    boton.textContent = "Convirtiendo documentos... espera";
  });
</script>
</body>
</html>
"""


def extension_permitida(nombre: str) -> bool:
    return "." in nombre and nombre.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def tipo_documento(ruta: Path) -> str:
    extension = ruta.suffix.lower().lstrip(".")
    if extension in EXCEL_EXTENSIONS:
        return "excel"
    if extension in WORD_EXTENSIONS:
        return "word"
    if extension in POWERPOINT_EXTENSIONS:
        return "powerpoint"
    if extension in PDF_EXTENSIONS:
        return "pdf"
    raise ValueError(f"Formato no admitido: {ruta.name}")


def nombre_pdf_final() -> str:
    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"documentos_concatenados_{marca_tiempo}.pdf"


def configurar_hoja_excel(hoja) -> None:
    page_setup = hoja.PageSetup
    page_setup.Orientation = XL_PORTRAIT
    page_setup.PaperSize = XL_PAPER_LETTER
    page_setup.LeftMargin = 0.25 * 72
    page_setup.RightMargin = 0.25 * 72
    page_setup.TopMargin = 0.25 * 72
    page_setup.BottomMargin = 0.25 * 72
    page_setup.HeaderMargin = 0.10 * 72
    page_setup.FooterMargin = 0.10 * 72
    page_setup.Zoom = False
    page_setup.FitToPagesWide = 1
    page_setup.FitToPagesTall = False
    page_setup.Order = 1
    page_setup.CenterHorizontally = True
    page_setup.LeftHeader = ""
    page_setup.CenterHeader = ""
    page_setup.RightHeader = ""
    page_setup.LeftFooter = ""
    page_setup.CenterFooter = ""
    page_setup.RightFooter = ""


def convertir_excel(ruta_entrada: Path, ruta_salida: Path) -> None:
    excel = None
    libro = None
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        libro = excel.Workbooks.Open(
            str(ruta_entrada),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        hojas_visibles = 0
        for hoja in libro.Worksheets:
            if hoja.Visible == XL_SHEET_VISIBLE:
                configurar_hoja_excel(hoja)
                hojas_visibles += 1
        if hojas_visibles == 0:
            raise ValueError(f'El Excel "{ruta_entrada.name}" no contiene hojas visibles.')
        libro.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=str(ruta_salida),
            Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        if libro is not None:
            libro.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def convertir_word(ruta_entrada: Path, ruta_salida: Path) -> None:
    word = None
    documento = None
    pythoncom.CoInitialize()
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        documento = word.Documents.Open(
            FileName=str(ruta_entrada),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        documento.ExportAsFixedFormat(
            OutputFileName=str(ruta_salida),
            ExportFormat=WORD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
        )
    finally:
        if documento is not None:
            documento.Close(SaveChanges=False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def convertir_powerpoint(ruta_entrada: Path, ruta_salida: Path) -> None:
    powerpoint = None
    presentacion = None
    pythoncom.CoInitialize()
    try:
        powerpoint = win32.DispatchEx("PowerPoint.Application")
        presentacion = powerpoint.Presentations.Open(
            FileName=str(ruta_entrada),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        presentacion.ExportAsFixedFormat(
            Path=str(ruta_salida),
            FixedFormatType=POWERPOINT_FIXED_FORMAT_PDF,
            Intent=POWERPOINT_INTENT_PRINT,
            FrameSlides=False,
            HandoutOrder=1,
            OutputType=1,
            PrintHiddenSlides=False,
            IncludeDocProperties=True,
            KeepIRMSettings=True,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    finally:
        if presentacion is not None:
            presentacion.Close()
        if powerpoint is not None:
            powerpoint.Quit()
        pythoncom.CoUninitialize()


def convertir_a_pdf(ruta_entrada: Path, carpeta_trabajo: Path, indice: int) -> Path:
    tipo = tipo_documento(ruta_entrada)
    if tipo == "pdf":
        return ruta_entrada

    ruta_salida = carpeta_trabajo / f"convertido_{indice:03d}.pdf"
    if tipo == "excel":
        convertir_excel(ruta_entrada, ruta_salida)
    elif tipo == "word":
        convertir_word(ruta_entrada, ruta_salida)
    elif tipo == "powerpoint":
        convertir_powerpoint(ruta_entrada, ruta_salida)

    if not ruta_salida.is_file() or ruta_salida.stat().st_size == 0:
        raise RuntimeError(f'No se pudo generar el PDF de "{ruta_entrada.name}".')
    return ruta_salida


def concatenar_pdfs(rutas_pdf: list[Path], ruta_final: Path) -> int:
    escritor = PdfWriter()
    try:
        for ruta_pdf in rutas_pdf:
            lector = PdfReader(str(ruta_pdf))
            if lector.is_encrypted:
                resultado = lector.decrypt("")
                if resultado == 0:
                    raise ValueError(
                        f'El PDF "{ruta_pdf.name}" está protegido con contraseña y no se puede concatenar.'
                    )
            escritor.append(lector)
        with open(ruta_final, "wb") as salida:
            escritor.write(salida)
        return len(escritor.pages)
    finally:
        escritor.close()


def ordenar_archivos(archivos, orden_json: str):
    archivos_por_nombre: dict[str, list] = {}
    for archivo in archivos:
        archivos_por_nombre.setdefault(archivo.filename, []).append(archivo)

    try:
        import json
        orden = json.loads(orden_json)
    except Exception:
        orden = []

    ordenados = []
    for nombre in orden:
        disponibles = archivos_por_nombre.get(nombre, [])
        if disponibles:
            ordenados.append(disponibles.pop(0))

    for disponibles in archivos_por_nombre.values():
        ordenados.extend(disponibles)

    return ordenados


@app.route("/", methods=["GET"])
def inicio():
    return render_template_string(
        HTML,
        resultado=None,
        max_files=MAX_FILES,
        max_file_mb=MAX_FILE_SIZE // (1024 * 1024),
    )


@app.route("/convertir", methods=["POST"])
def convertir():
    archivos = [archivo for archivo in request.files.getlist("archivos") if archivo.filename]

    if not archivos:
        flash("Selecciona al menos un documento.", "error")
        return redirect(url_for("inicio"))

    if len(archivos) > MAX_FILES:
        flash(f"El máximo permitido es {MAX_FILES} archivos por conversión.", "error")
        return redirect(url_for("inicio"))

    archivos = ordenar_archivos(archivos, request.form.get("orden_archivos", "[]"))

    for archivo in archivos:
        if not extension_permitida(archivo.filename):
            flash(f"Formato no admitido: {archivo.filename}", "error")
            return redirect(url_for("inicio"))

    token = f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{uuid4().hex[:8]}"
    carpeta_trabajo = Path(tempfile.mkdtemp(prefix=f"documentos_{token}_"))
    nombre_salida = nombre_pdf_final()
    ruta_pdf_temporal = carpeta_trabajo / nombre_salida
    ruta_pdf_final = OUTPUT_DIR / f"{token}__{nombre_salida}"

    try:
        pdfs_para_unir: list[Path] = []
        for indice, archivo in enumerate(archivos, start=1):
            nombre_seguro = secure_filename(archivo.filename)
            if not nombre_seguro:
                raise ValueError("Uno de los archivos tiene un nombre no válido.")
            ruta_entrada = carpeta_trabajo / f"{indice:03d}_{nombre_seguro}"
            archivo.save(ruta_entrada)
            pdfs_para_unir.append(convertir_a_pdf(ruta_entrada, carpeta_trabajo, indice))

        paginas = concatenar_pdfs(pdfs_para_unir, ruta_pdf_temporal)
        if not ruta_pdf_temporal.is_file() or ruta_pdf_temporal.stat().st_size == 0:
            raise RuntimeError("No se pudo crear el PDF final.")
        shutil.move(str(ruta_pdf_temporal), str(ruta_pdf_final))

    except Exception as error:
        flash(f"No se pudo completar la conversión: {error}", "error")
        return redirect(url_for("inicio"))

    finally:
        shutil.rmtree(carpeta_trabajo, ignore_errors=True)

    resultado = {
        "archivos": len(archivos),
        "paginas": paginas,
        "nombre_pdf": nombre_salida,
        "token": ruta_pdf_final.name,
    }
    return render_template_string(
        HTML,
        resultado=resultado,
        max_files=MAX_FILES,
        max_file_mb=MAX_FILE_SIZE // (1024 * 1024),
    )


@app.route("/descargar/<path:token>", methods=["GET"])
def descargar(token: str):
    nombre_seguro = Path(token).name
    ruta_pdf = OUTPUT_DIR / nombre_seguro
    if not ruta_pdf.is_file():
        flash("El PDF ya no está disponible. Convierte nuevamente los documentos.", "error")
        return redirect(url_for("inicio"))

    partes = nombre_seguro.split("__", 1)
    nombre_descarga = partes[1] if len(partes) == 2 else "documentos_concatenados.pdf"
    return send_file(
        ruta_pdf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=nombre_descarga,
        max_age=0,
    )


@app.errorhandler(413)
def archivo_demasiado_grande(_error):
    flash(
        f"La carga supera el límite permitido. Cada archivo puede tener hasta {MAX_FILE_SIZE // (1024 * 1024)} MB.",
        "error",
    )
    return redirect(url_for("inicio"))


if __name__ == "__main__":
    from waitress import serve
    print("Aplicación local disponible en: http://127.0.0.1:5000")
    serve(app, host="127.0.0.1", port=5000)