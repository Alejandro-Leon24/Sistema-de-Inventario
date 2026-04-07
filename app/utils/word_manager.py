import os
import re
import json
import logging
import platform
import shutil
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import datetime
import tempfile
import threading
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
# docx2pdf works perfectly on windows if MS Word is installed. 
# We wrap it in a try-except to log an error gracefully if MS Word is missing.
try:
    from docx2pdf import convert
except ImportError:
    convert = None

try:
    import mammoth
except ImportError:
    mammoth = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
except ImportError:
    A4 = None
    canvas = None

logger = logging.getLogger(__name__)
PDF_CONVERT_LOCK = threading.Lock()

_JINJA_TAG_NORMALIZE_PATTERNS = [
    (re.compile(r"\{\%\s*end\s+for\s*\%\}", re.IGNORECASE), "{% endfor %}"),
    (re.compile(r"\{\%\s*end\s+if\s*\%\}", re.IGNORECASE), "{% endif %}"),
    (re.compile(r"\{\%\s*end\s+block\s*\%\}", re.IGNORECASE), "{% endblock %}"),
]


def _find_libreoffice_binary():
    for candidate in ("libreoffice", "soffice", "soffice.exe"):
        path = shutil.which(candidate)
        if path:
            return path
    return None


def _convert_docx_to_pdf_with_libreoffice(docx_path, pdf_path):
    binary = _find_libreoffice_binary()
    if not binary:
        return False, "LibreOffice no está instalado o no está en PATH."

    out_dir = os.path.dirname(pdf_path)
    basename = os.path.splitext(os.path.basename(docx_path))[0]
    generated_pdf = os.path.join(out_dir, f"{basename}.pdf")

    cmd = [
        binary,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        docx_path,
    ]

    try:
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=120, check=False)
    except Exception as exc:
        return False, f"Falló la ejecución de LibreOffice: {exc}"

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"Código de salida {completed.returncode}"
        return False, f"LibreOffice no pudo convertir el DOCX a PDF. {detail}"

    if not os.path.exists(generated_pdf):
        return False, "LibreOffice finalizó, pero no se encontró el PDF generado."

    try:
        if os.path.abspath(generated_pdf) != os.path.abspath(pdf_path):
            os.replace(generated_pdf, pdf_path)
    except Exception as exc:
        return False, f"No se pudo mover el PDF convertido a su destino final: {exc}"

    return True, None


def get_preview_unavailable_reason():
    if mammoth is None:
        return "No se pudo mostrar la vista previa. Instala mammoth con: pip install mammoth"
    return None


def generate_area_summary_pdf_fallback(pdf_path, titulo_acta, table_rows):
    """
    Fallback puro Python para generar PDF cuando no hay Word/LibreOffice.
    table_rows: [{'descripcion': str, 'cantidad': int}, ...]
    """
    if not canvas or not A4:
        return False, "No está instalada la librería reportlab."

    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        page_w, page_h = A4

        margin_x = 42
        y = page_h - 56

        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin_x, y, str(titulo_acta or "ACTA"))
        y -= 24

        # Encabezado de tabla
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_x, y, "DESCRIPCION")
        c.drawRightString(page_w - margin_x, y, "CANTIDAD")
        y -= 8
        c.line(margin_x, y, page_w - margin_x, y)
        y -= 14

        c.setFont("Helvetica", 10)
        for row in table_rows or []:
            descripcion = str((row or {}).get("descripcion") or "-")
            cantidad = str((row or {}).get("cantidad") or 0)

            if y < 60:
                c.showPage()
                y = page_h - 56
                c.setFont("Helvetica-Bold", 10)
                c.drawString(margin_x, y, "DESCRIPCION")
                c.drawRightString(page_w - margin_x, y, "CANTIDAD")
                y -= 8
                c.line(margin_x, y, page_w - margin_x, y)
                y -= 14
                c.setFont("Helvetica", 10)

            # recorte simple de linea larga
            if len(descripcion) > 95:
                descripcion = f"{descripcion[:92]}..."

            is_total = descripcion.strip().upper() == "TOTAL DE BIENES"
            if is_total:
                c.setFont("Helvetica-Bold", 10)

            c.drawString(margin_x, y, descripcion)
            c.drawRightString(page_w - margin_x, y, cantidad)
            y -= 14

            if is_total:
                c.setFont("Helvetica", 10)

        c.save()
        return True, None
    except Exception as exc:
        return False, str(exc)


def _normalize_jinja_in_docx(template_path):
    """
    Crea una copia temporal del DOCX corrigiendo variantes comunes de tags Jinja
    que Word suele fragmentar o modificar (por ejemplo: {% end for %}).
    """
    try:
        temp_dir = tempfile.mkdtemp(prefix="docx_tpl_")
        normalized_path = os.path.join(temp_dir, os.path.basename(template_path))

        with zipfile.ZipFile(template_path, "r") as src_zip, zipfile.ZipFile(normalized_path, "w") as dst_zip:
            for item in src_zip.infolist():
                raw = src_zip.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    xml_text = raw.decode("utf-8", errors="ignore")
                    # Normaliza espacios no separables que rompen el parser de jinja.
                    xml_text = xml_text.replace("\u00a0", " ")
                    # Quita marcas de ortografía/gramática que Word inserta en medio de tags.
                    xml_text = re.sub(r"<w:proofErr[^>]*/>", "", xml_text)
                    for pattern, replacement in _JINJA_TAG_NORMALIZE_PATTERNS:
                        xml_text = pattern.sub(replacement, xml_text)
                    raw = xml_text.encode("utf-8")

                dst_zip.writestr(item, raw)

        return normalized_path
    except Exception as exc:
        logger.warning(f"No se pudo normalizar plantilla DOCX ({template_path}): {exc}")
        return template_path

# Configuración y estilos para las tablas extraidas
TABLE_STYLES_CONFIG = {
    # Aquí puedes personalizar cómo se renderiza la tabla generada dinámicamente
    "font_size": 10,
    "font_name": "Arial",
    "header_bg_color": "D9EAF7",
    "header_font_color": "000000",
    "header_font_bold": True,
    "border_color": "000000",
    "border_size": 4  # tamaño estándar en docx
}


def _set_cell_shading(cell, fill_color_hex):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), str(fill_color_hex or "D9EAF7"))
    tc_pr.append(shd)

def extract_variables_from_template(file_path):
    """
    Lee un documento Word (docx) y extrae todas las variables definidas como {{ variable }}
    Usando el motor de docxtpl.
    """
    pattern = re.compile(r"\{\{\s*([a-zA-Z_][a-zA-Z0-9_\.-]*)\s*\}\}")

    internal_vars = {"tabla_items", "tabla_columnas", "tabla_filas", "celda"}

    def _normalize(found):
        normalized = []
        for item in found:
            name = str(item or "").strip().lower()
            if not name:
                continue
            if name in internal_vars:
                continue
            if name.startswith("item.") or name.startswith("col.") or name.startswith("fila."):
                continue
            if name not in normalized:
                normalized.append(name)
        return normalized

    def _extract_with_docxtpl(path):
        doc = DocxTemplate(path)
        variables = doc.get_undeclared_template_variables() or set()
        return _normalize(list(variables))

    def _extract_with_python_docx(path):
        found = set()
        doc = Document(path)

        for p in doc.paragraphs:
            for match in pattern.findall(p.text or ""):
                found.add(match)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for match in pattern.findall(cell.text or ""):
                        found.add(match)

        for section in doc.sections:
            for p in section.header.paragraphs:
                for match in pattern.findall(p.text or ""):
                    found.add(match)
            for p in section.footer.paragraphs:
                for match in pattern.findall(p.text or ""):
                    found.add(match)

        return _normalize(list(found))

    def _extract_with_xml(path):
        found = set()
        with zipfile.ZipFile(path, "r") as zip_docx:
            targets = [
                name for name in zip_docx.namelist()
                if name.startswith("word/") and name.endswith(".xml")
            ]
            for name in targets:
                try:
                    xml = zip_docx.read(name).decode("utf-8", errors="ignore")
                except Exception:
                    continue
                try:
                    root = ET.fromstring(xml)
                    text_nodes = [node.text or "" for node in root.iter() if node.tag.endswith("}t")]
                    xml_text = "\n".join(text_nodes)
                except Exception:
                    xml_text = re.sub(r"<[^>]+>", "", xml)

                for match in pattern.findall(xml_text):
                    found.add(match)
        return _normalize(list(found))

    try:
        variables = _extract_with_docxtpl(file_path)
        if variables:
            return variables
    except Exception as e:
        logger.warning(f"Extractor docxtpl falló para {file_path}: {e}")

    try:
        variables = _extract_with_python_docx(file_path)
        if variables:
            return variables
    except Exception as e:
        logger.warning(f"Extractor python-docx falló para {file_path}: {e}")

    try:
        return _extract_with_xml(file_path)
    except Exception as e:
        logger.error(f"Error extrayendo variables de {file_path}: {e}")
        return []

def configure_table_data(extraction_data):
    """
    Toma los datos de la selección de inventario y los preprocesa
    para ser inyectados en el Word respetando la configuración de estilos.
    
    extraction_data: lista de diccionarios, e.g., [{"CÓDIGO INV.": "123", "DESCRIPCIÓN": "Silla"}, ...]
    """
    if not extraction_data:
        return []

    if not isinstance(extraction_data, list):
        return []

    normalized_rows = []
    for raw_row in extraction_data:
        if not isinstance(raw_row, dict):
            continue

        normalized_row = {}
        for key, value in raw_row.items():
            normalized_key = str(key).strip()
            if not normalized_key:
                continue

            if value is None:
                normalized_row[normalized_key] = ""
            elif isinstance(value, bool):
                normalized_row[normalized_key] = "Si" if value else "No"
            elif isinstance(value, float):
                normalized_row[normalized_key] = f"{value:.2f}".rstrip("0").rstrip(".")
            else:
                text_value = str(value).strip()
                if text_value.lower() in {"none", "null", "nan", "undefined"}:
                    text_value = ""
                normalized_row[normalized_key] = text_value

        if normalized_row:
            normalized_rows.append(normalized_row)

    return normalized_rows

def build_dynamic_table_context(table_rows, table_columns):
    """
    Construye estructuras para tablas dinámicas (columnas y filas)
    consumibles por docxtpl en Word.

    table_rows: lista de dicts con filas de inventario.
    table_columns: lista de columnas seleccionadas, por ejemplo:
        [{"id": "cod_inventario", "label": "CODIGO INV."}, ...]
    """
    if not table_rows or not table_columns:
        return [], []

    columnas = []
    for col in table_columns:
        col_id = str(col.get("id", "")).strip()
        if not col_id:
            continue
        columnas.append({
            "id": col_id,
            "label": str(col.get("label", col_id)).strip() or col_id,
        })

    if not columnas:
        return [], []

    filas = []
    for row in table_rows:
        celdas = []
        for col in columnas:
            raw_value = row.get(col["id"], "-")
            if raw_value is None:
                celdas.append("-")
            elif isinstance(raw_value, bool):
                celdas.append("Si" if raw_value else "No")
            else:
                text = str(raw_value).strip()
                if text.lower() in {"none", "null", "nan", "undefined"}:
                    text = ""
                celdas.append(text if text else "-")
        filas.append({"celdas": celdas})

    return columnas, filas


def build_dynamic_table_subdoc(doc_tpl, table_rows, table_columns):
    """
    Construye una tabla dinámica dentro de un subdocumento para evitar depender
    de tags complejos {%tc %}/{%tr %} sensibles a cambios de Word.

    Se renderiza en plantilla usando una sola variable: {{ tabla_dinamica }}
    """
    if not table_rows or not table_columns:
        return None

    columnas, filas = build_dynamic_table_context(table_rows, table_columns)
    if not columnas:
        return None

    try:
        subdoc = doc_tpl.new_subdoc()
    except Exception as exc:
        # new_subdoc depends on docxcompose; if it is missing, keep generation alive
        # so fixed-field templates and preview continue working.
        if "docxcompose" in str(exc).lower():
            logger.warning("docxcompose no instalado. tabla_dinamica se omitira temporalmente.")
            return None
        raise
    table = subdoc.add_table(rows=1, cols=len(columnas))
    # Algunas plantillas no incluyen el estilo en ingles; degradamos con fallback seguro.
    for style_name in ("Table Grid", "Tabla con cuadricula", "Tabla con cuadrícula"):
        try:
            table.style = style_name
            break
        except Exception:
            continue

    # Encabezados
    header_cells = table.rows[0].cells
    for idx, col in enumerate(columnas):
        _set_cell_shading(header_cells[idx], TABLE_STYLES_CONFIG.get("header_bg_color", "D9EAF7"))
        run = header_cells[idx].paragraphs[0].add_run(str(col.get("label", "")))
        run.bold = bool(TABLE_STYLES_CONFIG.get("header_font_bold", True))
        run.font.size = Pt(int(TABLE_STYLES_CONFIG.get("font_size", 10)))

    # Filas de datos
    for fila in filas:
        row_cells = table.add_row().cells
        celdas = fila.get("celdas", [])
        for idx in range(len(columnas)):
            val = celdas[idx] if idx < len(celdas) else "-"
            run = row_cells[idx].paragraphs[0].add_run(str(val if val is not None else "-"))
            run.font.size = Pt(10)

    return subdoc

def generate_acta(
    template_path,
    context_data,
    table_data=None,
    table_columns=None,
    output_dir=None,
    doc_name="acta",
    generate_pdf=True,
    use_date_subfolder=True,
    include_time_suffix=True,
):
    """
    Genera el archivo DOCX final y un PDF.
    
    template_path: ruta de la plantilla DOCX base
    context_data: Diccionario con los datos del form (e.g. {"entregado_por": "Juan", "ubicacion": "Bloque A"})
    table_data: Lista de diccionarios con la tabla del inventario
    output_dir: Directorio base de descargas (por ej: C:/Users/x/Downloads/inventario)
    """
    normalized_template = _normalize_jinja_in_docx(template_path)
    doc = DocxTemplate(normalized_template)
    # Evita mutar el contexto original del request (puede romper json.dumps en historial).
    render_context = dict(context_data or {})
    
    # Inyectar tabla si existe
    if table_data:
        render_context['tabla_items'] = configure_table_data(table_data)

    # Modo dinámico completo: encabezados y celdas según columnas seleccionadas
    if table_data and table_columns:
        tabla_columnas, tabla_filas = build_dynamic_table_context(table_data, table_columns)
        render_context['tabla_columnas'] = tabla_columnas
        render_context['tabla_filas'] = tabla_filas
        # Camino robusto recomendado: insertar tabla como subdocumento
        render_context['tabla_dinamica'] = build_dynamic_table_subdoc(doc, table_data, table_columns)
        
    try:
        doc.render(render_context)
    except Exception as e:
        msg = str(e)
        if "unknown tag 'endfor'" in msg.lower() or "unexpected 'endfor'" in msg.lower():
            raise ValueError(
                "Sintaxis de tabla dinámica inválida en Word. "
                "Recomendado: use una sola variable {{ tabla_dinamica }} en la celda donde debe ir la tabla. "
                "Alternativa avanzada: "
                "{%tc for col in tabla_columnas %}{{ col.label }}{%tc endfor %} "
                "y {%tr for fila in tabla_filas %}{%tc for celda in fila.celdas %}{{ celda }}{%tc endfor %}{%tr endfor %}."
            )
        raise
    
    # Preparar el directorio de salida (e.g., ../Downloads/inventario/acta entrega/31-03-2026/ )
    if not output_dir:
        # Default de reserva si no proveen uno
        home = os.path.expanduser("~")
        output_dir = os.path.join(home, "Downloads", "inventario", doc_name.lower())
        
    fecha_str = datetime.datetime.now().strftime("%d-%m-%Y")
    final_output_dir = os.path.join(output_dir, fecha_str) if use_date_subfolder else output_dir
    
    if not os.path.exists(final_output_dir):
        os.makedirs(final_output_dir, exist_ok=True)
    
    # Generar ruta del archivo
    if include_time_suffix:
        time_str = datetime.datetime.now().strftime("%H%M%S")
        docx_filename = f"{doc_name}_{time_str}.docx"
        pdf_filename = f"{doc_name}_{time_str}.pdf"
    else:
        docx_filename = f"{doc_name}.docx"
        pdf_filename = f"{doc_name}.pdf"
    
    docx_path = os.path.join(final_output_dir, docx_filename)
    pdf_path = os.path.join(final_output_dir, pdf_filename)
    
    # Guardar docx
    try:
        doc.save(docx_path)
    except Exception as e:
        logger.error(f"Error guardando DOCX en {docx_path}: {e}")
        return None, None
        
    # Guardar PDF con fallback multiplataforma:
    # 1) Word/docx2pdf en Windows
    # 2) LibreOffice headless en cualquier SO
    gen_pdf = False
    converter = convert
    if platform.system() == "Windows" and not converter:
        # Intento tardío por si la app inició antes de instalar docx2pdf.
        try:
            from docx2pdf import convert as runtime_convert
            converter = runtime_convert
        except Exception:
            converter = None

    if generate_pdf:
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)

        if platform.system() == "Windows" and converter:
            try:
                # Word COM conversion is not reliable in parallel; serialize calls.
                with PDF_CONVERT_LOCK:
                    if pythoncom:
                        pythoncom.CoInitialize()
                    converter(abs_docx, abs_pdf)
                    if pythoncom:
                        pythoncom.CoUninitialize()
                gen_pdf = True
            except Exception as e:
                try:
                    if pythoncom:
                        pythoncom.CoUninitialize()
                except Exception:
                    pass
                logger.error(f"Error convirtiendo DOCX a PDF con Word/docx2pdf. Detalle: {e}")
                gen_pdf = False

        if not gen_pdf:
            with PDF_CONVERT_LOCK:
                lo_ok, lo_error = _convert_docx_to_pdf_with_libreoffice(abs_docx, abs_pdf)
            gen_pdf = lo_ok
            if not gen_pdf:
                logger.warning(f"No fue posible convertir DOCX a PDF. {lo_error}")
            
    # Retornar rutas locales (absolutas preferiblemente o directas)
    return docx_path, pdf_path if gen_pdf else None


def render_docx_preview_html(docx_path):
    """
    Convierte DOCX a HTML simple para fallback de vista previa cuando no se
    dispone de PDF.
    """
    if not docx_path or not os.path.exists(docx_path) or mammoth is None:
        return None

    try:
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value or ""
        if not html.strip():
            return None

        return (
            "<html><head><meta charset='utf-8'></head>"
            "<body style='font-family: Arial, sans-serif; margin: 20px;'>"
            f"{html}"
            "</body></html>"
        )
    except Exception as exc:
        logger.warning(f"No se pudo generar preview HTML para {docx_path}: {exc}")
        return None
