import os
import re
import json
import logging
import platform
import datetime
from docxtpl import DocxTemplate
# docx2pdf works perfectly on windows if MS Word is installed. 
# We wrap it in a try-except to log an error gracefully if MS Word is missing.
try:
    from docx2pdf import convert
except ImportError:
    convert = None

logger = logging.getLogger(__name__)

# Configuración y estilos para las tablas extraidas
TABLE_STYLES_CONFIG = {
    # Aquí puedes personalizar cómo se renderiza la tabla generada dinámicamente
    "font_size": 10,
    "font_name": "Arial",
    "header_bg_color": "EFEFEF",
    "header_font_color": "000000",
    "header_font_bold": True,
    "border_color": "000000",
    "border_size": 4  # tamaño estándar en docx
}

def extract_variables_from_template(file_path):
    """
    Lee un documento Word (docx) y extrae todas las variables definidas como {{ variable }}
    Usando el motor de docxtpl.
    """
    try:
        doc = DocxTemplate(file_path)
        # get_undeclared_template_variables returns a set of variables found in the template
        variables = doc.get_undeclared_template_variables()
        return list(variables)
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
    
    # Podrías agregar más manipulación base según lo que requieras aquí con TABLE_STYLES_CONFIG
    return extraction_data

def generate_acta(template_path, context_data, table_data=None, output_dir=None, doc_name="acta"):
    """
    Genera el archivo DOCX final y un PDF.
    
    template_path: ruta de la plantilla DOCX base
    context_data: Diccionario con los datos del form (e.g. {"entregado_por": "Juan", "ubicacion": "Bloque A"})
    table_data: Lista de diccionarios con la tabla del inventario
    output_dir: Directorio base de descargas (por ej: C:/Users/x/Downloads/inventario)
    """
    doc = DocxTemplate(template_path)
    
    # Inyectar tabla si existe
    if table_data:
        context_data['tabla_items'] = configure_table_data(table_data)
        
    doc.render(context_data)
    
    # Preparar el directorio de salida (e.g., ../Downloads/inventario/acta entrega/31-03-2026/ )
    if not output_dir:
        # Default de reserva si no proveen uno
        home = os.path.expanduser("~")
        output_dir = os.path.join(home, "Downloads", "inventario", doc_name.lower())
        
    fecha_str = datetime.datetime.now().strftime("%d-%m-%Y")
    final_output_dir = os.path.join(output_dir, fecha_str)
    
    if not os.path.exists(final_output_dir):
        os.makedirs(final_output_dir, exist_ok=True)
    
    # Generar ruta del archivo
    time_str = datetime.datetime.now().strftime("%H%M%S")
    docx_filename = f"{doc_name}_{time_str}.docx"
    pdf_filename = f"{doc_name}_{time_str}.pdf"
    
    docx_path = os.path.join(final_output_dir, docx_filename)
    pdf_path = os.path.join(final_output_dir, pdf_filename)
    
    # Guardar docx
    try:
        doc.save(docx_path)
    except Exception as e:
        logger.error(f"Error guardando DOCX en {docx_path}: {e}")
        return None, None
        
    # Guardar pd (Asume Windows con MS Word instalado para docx2pdf)
    gen_pdf = False
    if platform.system() == "Windows" and convert:
        try:
            # requiere ruta absoluta para ambos parms!
            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(pdf_path)
            convert(abs_docx, abs_pdf)
            gen_pdf = True
        except Exception as e:
            logger.error(f"Error convirtiendo de DOCX a PDF (Word podría no estar instalado o abierto). Detalle: {e}")
            gen_pdf = False
            
    # Retornar rutas locales (absolutas preferiblemente o directas)
    return docx_path, pdf_path if gen_pdf else None
