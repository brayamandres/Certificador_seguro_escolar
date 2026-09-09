import io
import os
import sys
import re
import shutil
import subprocess
import traceback
import pandas as pd
from functools import wraps
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, Response
from docxtpl import DocxTemplate

app = Flask(__name__)

# --- SECRET_KEY: obligatoria en producción, sin fallback inseguro ---
SECRET_KEY = os.environ.get("SECRET_KEY")
if not SECRET_KEY:
    if os.environ.get("FLASK_ENV") == "production" or os.environ.get("AWS_EXECUTION_ENV") or os.environ.get("PORT"):
        raise RuntimeError("Debes definir la variable de entorno SECRET_KEY antes de arrancar en producción.")
    SECRET_KEY = "clave-dev-solo-local"  # solo para desarrollo local
app.secret_key = SECRET_KEY

# --- Límite de tamaño de subida (16 MB) ---
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

# --- Autenticación básica (uso interno) ---
BASIC_AUTH_USER = os.environ.get("BASIC_AUTH_USER")
BASIC_AUTH_PASS = os.environ.get("BASIC_AUTH_PASS")

def requiere_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not BASIC_AUTH_USER or not BASIC_AUTH_PASS:
            # Si no se configuraron credenciales, no se exige auth (útil en dev local)
            return f(*args, **kwargs)
        auth = request.authorization
        if not auth or auth.username != BASIC_AUTH_USER or auth.password != BASIC_AUTH_PASS:
            return Response(
                "Acceso restringido.", 401,
                {"WWW-Authenticate": 'Basic realm="Login requerido"'}
            )
        return f(*args, **kwargs)
    return decorated

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
PLANTILLA_DOCX = os.path.join(DIRECTORIO_BASE, "PLANTILLA.docx")
PLANTILLA_EXCEL = os.path.join(DIRECTORIO_BASE, "PLANTILLA_CERTIFICADOS.xlsx")
CARPETA_UPLOADS = os.path.join(DIRECTORIO_BASE, "uploads")
CARPETA_DOC_WORD = os.path.join(DIRECTORIO_BASE, "certificados_word")
CARPETA_PDF_FINAL = os.path.join(DIRECTORIO_BASE, "certificados_pdf")
CARPETA_OUTPUT_ZIP = os.path.join(DIRECTORIO_BASE, "archivos_zip")

for carpeta in [CARPETA_UPLOADS, CARPETA_DOC_WORD, CARPETA_PDF_FINAL, CARPETA_OUTPUT_ZIP]:
    os.makedirs(carpeta, exist_ok=True)


def limpiar_texto(valor):
    if pd.isna(valor) or valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()

def formatear_fecha(valor):
    if pd.isna(valor) or valor is None or str(valor).strip() == "":
        return ""
    try:
        fecha_dt = pd.to_datetime(valor, dayfirst=True)
        return fecha_dt.strftime("%d/%m/%Y")
    except Exception:
        return str(valor).strip()

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", nombre).strip()

def limpiar_carpeta(ruta_carpeta):
    """Elimina todo el contenido dentro de una carpeta sin borrar la carpeta en sí."""
    if os.path.exists(ruta_carpeta):
        for elem in os.listdir(ruta_carpeta):
            ruta_elem = os.path.join(ruta_carpeta, elem)
            try:
                if os.path.isfile(ruta_elem) or os.path.islink(ruta_elem):
                    os.remove(ruta_elem)
                elif os.path.isdir(ruta_elem):
                    shutil.rmtree(ruta_elem)
            except Exception as e:
                app.logger.warning(f"No se pudo eliminar {ruta_elem}: {e}")

def limpiar_carpetas_temporales():
    """Limpia todas las carpetas de trabajo para no acumular archivos en el servidor."""
    for carpeta in [CARPETA_DOC_WORD, CARPETA_PDF_FINAL, CARPETA_OUTPUT_ZIP, CARPETA_UPLOADS]:
        limpiar_carpeta(carpeta)

def convertir_a_pdf(carpeta_docx, carpeta_pdf, docx_generados):
    """
    Convierte archivos .docx a .pdf de forma compatible con multiples entornos:
    1. Si 'soffice' (LibreOffice) esta disponible en el PATH (Linux/Docker o Windows con LibreOffice),
       lo utiliza por lotes con soffice --headless.
    2. Si estamos en Windows y no hay 'soffice', utiliza Microsoft Word mediante docx2pdf.
    """
    soffice_cmd = shutil.which("soffice")
    if soffice_cmd:
        subprocess.run(
            [soffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", carpeta_pdf] + docx_generados,
            check=True,
            timeout=60 + 10 * len(docx_generados),
        )
        return

    if sys.platform == "win32":
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        try:
            convert(carpeta_docx, carpeta_pdf)
        finally:
            pythoncom.CoUninitialize()
        return

    raise RuntimeError("No se encontro LibreOffice ('soffice') ni Microsoft Word en el sistema para convertir a PDF.")


@app.route("/", methods=["GET"])
@requiere_auth
def index():
    zip_pdf_existe = os.path.exists(os.path.join(CARPETA_OUTPUT_ZIP, "certificados_pdf.zip"))
    return render_template("index.html", pdf_listo=zip_pdf_existe)


@app.route("/descargar-plantilla", methods=["GET"])
@requiere_auth
def descargar_plantilla():
    if not os.path.exists(PLANTILLA_EXCEL):
        flash("La plantilla de ejemplo no está disponible en el servidor.", "error")
        return redirect(url_for("index"))
    return send_file(
        PLANTILLA_EXCEL,
        as_attachment=True,
        download_name="PLANTILLA_CERTIFICADOS.xlsx",
    )


@app.route("/procesar", methods=["POST"])
@requiere_auth
def procesar_excel():
    if "archivo_excel" not in request.files:
        flash("No se seleccionó ningún archivo.", "error")
        return redirect(url_for("index"))

    file = request.files["archivo_excel"]
    if file.filename == "":
        flash("Por favor selecciona un archivo Excel válido.", "error")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("El archivo debe ser un Excel (.xlsx o .xls).", "error")
        return redirect(url_for("index"))

    if not os.path.exists(PLANTILLA_DOCX):
        flash("ERROR: No se encontró el archivo PLANTILLA.docx en el servidor.", "error")
        return redirect(url_for("index"))

    # Limpiamos archivos temporales previos
    limpiar_carpetas_temporales()

    try:
        excel_path = os.path.join(CARPETA_UPLOADS, "excel_procesado.xlsx")
        file.save(excel_path)

        # 1. Metadatos
        try:
            df_info = pd.read_excel(excel_path, skiprows=3, nrows=4, usecols="B:C", header=None)
            df_info.columns = ["Campo", "Valor"]
            info_dict = dict(zip(df_info["Campo"].astype(str).str.strip(), df_info["Valor"]))
            año = "2026"
            for k, v in info_dict.items():
                if "a" in k.lower() and "o" in k.lower():
                    año = limpiar_texto(v)
                    break
        except Exception:
            año = "2026"

        # 2. Datos
        df = pd.read_excel(excel_path, header=9)
        df.columns = df.columns.astype(str).str.strip()

        # Una vez leídos los datos en memoria, eliminamos el Excel del servidor
        if os.path.exists(excel_path):
            try:
                os.remove(excel_path)
            except Exception:
                pass

        columnas_clave = [col for col in ["NOMBRES", "RUT"] if col in df.columns]
        if columnas_clave:
            df = df.dropna(subset=columnas_clave, how="all")
        if "NOMBRES" in df.columns:
            df = df[df["NOMBRES"].astype(str).str.strip() != ""]

        if df.empty:
            flash("No se encontraron registros válidos en el Excel. Revisa el formato de la planilla.", "error")
            return redirect(url_for("index"))

        # 3. Generar Word temporalmente (necesarios para convertirlos a PDF)
        filas_con_error = []
        docx_generados = []
        for index, row in df.iterrows():
            try:
                doc = DocxTemplate(PLANTILLA_DOCX)
                nombres = limpiar_texto(row.get("NOMBRES"))
                apellidos = limpiar_texto(row.get("APELLIDOS"))

                contexto = {
                    "Año": año,
                    "NOMBRE_DEL_ENCARGADO_A": limpiar_texto(row.get("NOMBRE DEL ENCARGADO/A")),
                    "CARGO_DEL_ENCARGADO": limpiar_texto(row.get("CARGO DEL ENCARGADO")),
                    "NOMBRE_EMPRESA_O_INSTITUCIÓN": limpiar_texto(row.get("NOMBRE EMPRESA O INSTITUCIÓN")),
                    "NOMBRES": nombres,
                    "APELLIDOS": apellidos,
                    "RUT": limpiar_texto(row.get("RUT")),
                    "CARRERA": limpiar_texto(row.get("CARRERA")),
                    "PRACTICA_PASANTIA": limpiar_texto(row.get("PRÁCTICA / PASANTÍA")),
                    "FECHA_DE_INICIO": formatear_fecha(row.get("FECHA DE INICIO")),
                    "FECHA_DE_TÉRMINO": formatear_fecha(row.get("FECHA DE TÉRMINO")),
                    "HORAS_TOTALES": limpiar_texto(row.get("HORAS TOTALES")),
                }

                doc.render(contexto)
                nombre_persona = f"{nombres} {apellidos}".strip() or f"Registro_{index+1}"
                nombre_docx = limpiar_nombre_archivo(f"NUEVO SEGURO GENERAL {nombre_persona}.docx")
                ruta_docx = os.path.join(CARPETA_DOC_WORD, nombre_docx)
                doc.save(ruta_docx)
                docx_generados.append(ruta_docx)
            except Exception:
                app.logger.exception(f"Error generando certificado para la fila {index + 1}")
                filas_con_error.append(index + 1)

        if not docx_generados:
            flash("No se pudo generar ningún certificado. Revisa el formato de la planilla y la plantilla.", "error")
            return redirect(url_for("index"))

        # 4. Generar PDF y eliminar inmediatamente los archivos Word temporales
        try:
            convertir_a_pdf(CARPETA_DOC_WORD, CARPETA_PDF_FINAL, docx_generados)
        finally:
            # Requisito: Los Word se eliminan inmediatamente tras la creación de los PDF
            limpiar_carpeta(CARPETA_DOC_WORD)

        # Crear ZIP únicamente con los PDF generados
        shutil.make_archive(
            os.path.join(CARPETA_OUTPUT_ZIP, "certificados_pdf"), "zip", CARPETA_PDF_FINAL
        )
        # Limpiar los PDFs individuales en disco para liberar espacio (ya están en el ZIP)
        limpiar_carpeta(CARPETA_PDF_FINAL)

        mensaje = f"Éxito: Se generaron {len(docx_generados)} de {len(df)} certificados en PDF."
        if filas_con_error:
            mensaje += f" Filas con error (revisar planilla): {', '.join(map(str, filas_con_error))}."
        flash(mensaje, "warning" if filas_con_error else "success")

    except Exception:
        app.logger.exception("Error durante el procesamiento del Excel")
        limpiar_carpetas_temporales()
        flash("Ocurrió un error durante el procesamiento. Revisa el formato del archivo e intenta nuevamente.", "error")

    return redirect(url_for("index"))


@app.route("/descargar/pdf", methods=["GET"])
@app.route("/descargar/<tipo>", methods=["GET"])
@requiere_auth
def descargar_zip(tipo="pdf"):
    if tipo != "pdf":
        flash("Solo está habilitada la descarga de certificados en PDF.", "warning")
        return redirect(url_for("index"))

    ruta_zip = os.path.join(CARPETA_OUTPUT_ZIP, "certificados_pdf.zip")
    if not os.path.exists(ruta_zip):
        flash("Los certificados en PDF ya fueron descargados o no están disponibles.", "warning")
        return redirect(url_for("index"))

    try:
        # Leemos el archivo en memoria para poder eliminarlo inmediatamente del servidor
        with open(ruta_zip, "rb") as f:
            archivo_memoria = io.BytesIO(f.read())

        # Requisito: Luego de descargar exitosamente los PDF, se eliminan del servidor
        limpiar_carpetas_temporales()

        archivo_memoria.seek(0)
        return send_file(
            archivo_memoria,
            mimetype="application/zip",
            as_attachment=True,
            download_name="Certificados_PDF.zip",
        )
    except Exception as e:
        app.logger.exception(f"Error sirviendo archivo ZIP: {e}")
        flash("Ocurrió un error al descargar el archivo.", "error")
        return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)