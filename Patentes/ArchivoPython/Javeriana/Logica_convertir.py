import pandas as pd
import spacy
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import logging
import datetime

# Inicializar el log
logging.basicConfig(filename='xml_conversion.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Función para inicializar el detector de idiomas en spaCy
def get_lang_detector(nlp, name):
    return LanguageDetector()

# Cargar el modelo de procesamiento de lenguaje natural en español de spaCy
nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector", func=get_lang_detector)
nlp.add_pipe('language_detector', last=True)

# Función para crear el archivo XML a partir de cada fila de Excel
def crear_archivo_xml(fila_datos, version="VX"):
    logging.info(f"Procesando fila: {fila_datos['No.']}")

    try:
        # Crear el elemento raíz "patentes"
        patentes = ET.Element("v1:patentes", xmlns="v1.patentes.pure.atira.dk", xmlns_v3="v3.commons.pure.atira.dk")

        # Crear el elemento "patente"
        patente = ET.SubElement(patentes, "v1:patente", id=f"patente_{fila_datos['No.']}", type=fila_datos['Tipo de activo de PI'].strip())

        # Detectar el idioma del título
        titulo_patente = fila_datos['Titulo de la Patente'].strip()
        titulo_patente_lang = nlp(titulo_patente)._.language['language']

        # Crear el título de la patente
        title = ET.SubElement(patente, "v1:title")
        title_text = ET.SubElement(title, "v3:text", lang=titulo_patente_lang, country=fila_datos['JURISDICCIÓN'].strip())
        title_text.text = titulo_patente

        # Descripción de la patente
        descripcion_patente = fila_datos['Descripcion'].strip()
        descripcion_patente_lang = 'es' if not descripcion_patente else nlp(descripcion_patente)._.language['language']

        if descripcion_patente:
            description = ET.SubElement(patente, "v1:description")
            description_text = ET.SubElement(description, "v3:text", lang=descripcion_patente_lang, country=fila_datos['JURISDICCIÓN'].strip())
            description_text.text = descripcion_patente

        # Jurisdicción
        jurisdiccion = ET.SubElement(patente, "v1:jurisdiccion")
        jurisdiccion.text = fila_datos['JURISDICCIÓN'].strip()

        # Número de Solicitud
        if pd.notna(fila_datos['Número de Solicitud']):
            numero_solicitud = ET.SubElement(patente, "v1:numeroSolicitud")
            numero_solicitud.text = fila_datos['Número de Solicitud'].strip()

        # Número de Prioridad
        if pd.notna(fila_datos['Número de Prioridad']):
            numero_prioridad = ET.SubElement(patente, "v1:numeroPrioridad")
            numero_prioridad.text = fila_datos['Número de Prioridad'].strip()

        # Fecha de Solicitud
        if pd.notna(fila_datos['Fecha de Solicitud']):
            fecha_solicitud = ET.SubElement(patente, "v1:fechaSolicitud")
            fecha_solicitud.text = fila_datos['Fecha de Solicitud'].strip()

        # Titular
        if pd.notna(fila_datos['Titular']):
            titular = ET.SubElement(patente, "v1:titular")
            titular.text = fila_datos['Titular'].strip()

        # Inventores / Autores
        if pd.notna(fila_datos['INVENTORES/AUTORES']):
            inventores_autores = ET.SubElement(patente, "v1:inventoresAutores")
            for inventor in fila_datos['INVENTORES/AUTORES'].split("\n"):
                inventor_elem = ET.SubElement(inventores_autores, "v1:inventor")
                inventor_elem.text = inventor.strip()

        # Tipo de Vinculación PUJ
        if pd.notna(fila_datos['Tipo Vinculación PUJ (Interno - Externo)']):
            tipo_vinculacion = ET.SubElement(patente, "v1:tipoVinculacionPUJ")
            tipo_vinculacion.text = fila_datos['Tipo Vinculación PUJ (Interno - Externo)'].strip()

        # Estado actual del trámite
        if pd.notna(fila_datos['Estado actual del trámite']):
            estado_tramite = ET.SubElement(patente, "v1:estadoTramite")
            estado_tramite.text = fila_datos['Estado actual del trámite'].strip()

        # Fecha de Concesión
        if pd.notna(fila_datos['Fecha de concesión']):
            fecha_concesion = ET.SubElement(patente, "v1:fechaConcesion")
            fecha_concesion.text = fila_datos['Fecha de concesión'].strip()

        # Facultad
        if pd.notna(fila_datos['Facultad']):
            facultad = ET.SubElement(patente, "v1:facultad")
            facultad.text = fila_datos['Facultad'].strip()

        # Departamento / Instituto
        if pd.notna(fila_datos['Departamento/Instituto']):
            departamento_instituto = ET.SubElement(patente, "v1:departamentoInstituto")
            departamento_instituto.text = fila_datos['Departamento/Instituto'].strip()

        # Grupo de Investigación
        if pd.notna(fila_datos['Grupo De Investigación']):
            grupo_investigacion = ET.SubElement(patente, "v1:grupoInvestigacion")
            grupo_investigacion.text = fila_datos['Grupo De Investigación'].strip()

        # Link de Consulta
        if pd.notna(fila_datos.get('Link De Consulta', None)):
            link_consulta = ET.SubElement(patente, "v1:linkConsulta")
            link_consulta.text = fila_datos['Link De Consulta']

        # Guardar el archivo XML en la carpeta "Resultado"
        hoy = datetime.datetime.now().strftime("%Y_%m_%d")
        nombre_archivo = f"{hoy}_patente_{fila_datos['No.']}.xml"
        tree = ET.ElementTree(patentes)

        with open(f"Resultado/{nombre_archivo}", "wb") as archivo:
            tree.write(archivo, encoding="utf-8", xml_declaration=True)

        logging.info(f"Archivo XML generado: {nombre_archivo}")
        return nombre_archivo

    except Exception as e:
        logging.error(f"Error procesando la fila {fila_datos['No.']}: {str(e)}")
        raise

# Leer el archivo Excel y cargar la plantilla XML
data = pd.read_excel('Productos de PI Javeriana.xlsx')

# Iterar sobre cada fila para generar los archivos XML
for index, fila in data.iterrows():
    try:
        archivo_xml = crear_archivo_xml(fila)
        print(f"Generado: {archivo_xml}")
    except Exception as e:
        logging.error(f"Error procesando la fila {index}: {str(e)}")
