import os
import pandas as pd
import spacy
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import logging
import datetime
import uuid

# Inicializar el log
logging.basicConfig(filename='xml_conversion.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Función para inicializar el detector de idiomas en spaCy
def get_lang_detector(nlp, name):
    return LanguageDetector()

# Cargar el modelo de procesamiento de lenguaje natural en español de spaCy
nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector", func=get_lang_detector)
nlp.add_pipe('language_detector', last=True)

# Cargar el archivo Excel que contiene los datos de las patentes desde la pestaña 'Hoja1'
dataframe_patentes = pd.read_excel(r"pruebas.xlsx", sheet_name='Hoja1', dtype=object)

# Verificar las columnas disponibles
print("Columnas disponibles en la hoja de Excel:", dataframe_patentes.columns)

# Crear la carpeta Resultado si no existe
if not os.path.exists("Resultado"):
    os.makedirs("Resultado")

# Crear el elemento raíz del archivo XML que agrupará todas las patentes
root = ET.Element("patentes", {
    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xsi:schemaLocation": "https://puj-staging.elsevierpure.com/ws/api/524 https://puj-staging.elsevierpure.com/ws/api/524/xsd/schema1.xsd"
})

# Función para aplicar indentación en el XML
def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for subelem in elem:
            indent(subelem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

# Crear un conjunto para almacenar los títulos de las patentes ya procesadas
titulos_procesados = set()

# Iterar sobre todas las filas del DataFrame para procesar cada patente
for index, fila in dataframe_patentes.iterrows():
    print(f"Procesando la fila {index + 1} de {len(dataframe_patentes)}")

    # Extraer la información clave de la patente
    temp_patente_id = str(fila['No.']) if pd.notna(fila['No.']) else "Desconocido"
    temp_patente_titulo = str(fila['Titulo de la Patente']) if pd.notna(fila['Titulo de la Patente']) else "Sin título"

    # Verificar si la patente con el mismo título ya ha sido procesada
    if temp_patente_titulo in titulos_procesados:
        print(f"Patente con título '{temp_patente_titulo}' ya procesada, omitiendo...")
        continue  # Si ya fue procesada, omitimos esta patente

    # Agregar el título de la patente al conjunto para marcarla como procesada
    titulos_procesados.add(temp_patente_titulo)

    # Generar un UUID único para la patente
    temp_uuid = str(uuid.uuid4())
    
    # Crear el elemento de la patente dentro del archivo XML general
    patent = ET.SubElement(root, "patent", {
        "uuid": temp_uuid,
        "pureId": temp_patente_id,
        "externalId": temp_patente_id,
        "externalIdSource": "espacenet"
    })

    # Crear el título de la patente
    title = ET.SubElement(patent, "title", {"formatted": "true"})
    ET.SubElement(title, "text").text = temp_patente_titulo

    # Crear el tipo de la patente
    patent_type = ET.SubElement(patent, "type", {
        "pureId": "13611",
        "uri": "/dk/atira/pure/researchoutput/researchoutputtypes/patent/patent"
    })
    term = ET.SubElement(patent_type, "term", {"formatted": "false"})
    ET.SubElement(term, "text", {"locale": "en_US"}).text = "Invention patent"
    ET.SubElement(term, "text", {"locale": "es_CO"}).text = "Patente de Invención"

    # Categoría de la patente
    category = ET.SubElement(patent, "category", {
        "pureId": "4255",
        "uri": "/dk/atira/pure/researchoutput/category/research"
    })
    category_term = ET.SubElement(category, "term", {"formatted": "false"})
    ET.SubElement(category_term, "text", {"locale": "en_US"}).text = "Research"
    ET.SubElement(category_term, "text", {"locale": "es_CO"}).text = "Investigación"

    # Crear la descripción de la patente
    if pd.notna(fila['Descripcion']):
        abstract = ET.SubElement(patent, "abstract", {"formatted": "true"})
        abstract_text = ET.SubElement(abstract, "text", {"locale": "es_CO"})
        abstract_text.text = f"<![CDATA[{fila['Descripcion']}]]>"

    # Añadir los datos de publicación
    publication_statuses = ET.SubElement(patent, "publicationStatuses")
    publication_status = ET.SubElement(publication_statuses, "publicationStatus", {
        "current": "true",
        "pureId": temp_patente_id,
        "externalId": temp_patente_id,
        "externalIdSource": "espacenet"
    })
    pub_status = ET.SubElement(publication_status, "publicationStatus", {
        "pureId": "1017",
        "uri": "/dk/atira/pure/researchoutput/status/published"
    })
    term = ET.SubElement(pub_status, "term", {"formatted": "false"})
    ET.SubElement(term, "text", {"locale": "en_US"}).text = "Published"
    ET.SubElement(term, "text", {"locale": "es_CO"}).text = "Publicada"

    # Fecha de publicación (si está presente)
    if pd.notna(fila['Fecha de Solicitud']):
        publication_date = ET.SubElement(publication_status, "publicationDate")
        fecha_solicitud = str(fila['Fecha de Solicitud'])
        if fecha_solicitud:
            date_parts = fecha_solicitud.split("-")
            if len(date_parts) == 3:
                ET.SubElement(publication_date, "year").text = date_parts[0]
                ET.SubElement(publication_date, "month").text = date_parts[1]
                ET.SubElement(publication_date, "day").text = date_parts[2]

    # Idioma
    language = ET.SubElement(patent, "language", {"pureId": "240", "uri": "/dk/atira/pure/core/languages/es_CO"})
    language_term = ET.SubElement(language, "term", {"formatted": "false"})
    ET.SubElement(language_term, "text", {"locale": "en_US"}).text = "Spanish"
    ET.SubElement(language_term, "text", {"locale": "es_CO"}).text = "Español"

    # Añadir jurisdicción
    if pd.notna(fila['Jurisdiccion']):
        jurisdiccion = ET.SubElement(patent, "jurisdiccion")
        jurisdiccion.text = fila['Jurisdiccion']

    # Añadir número de patente
    if pd.notna(fila['Numero de Solicitud ']):
        patent_number = ET.SubElement(patent, "patentNumber")
        patent_number.text = str(fila['Numero de Solicitud '])

    # Asociaciones de personas (inventores)
    if pd.notna(fila['Inventores/Autores']):
        person_associations = ET.SubElement(patent, "personAssociations")
        for inventor in str(fila['Inventores/Autores']).split(","):
            person_association = ET.SubElement(person_associations, "personAssociation", {"pureId": temp_patente_id})
            external_person = ET.SubElement(person_association, "externalPerson", {"uuid": str(uuid.uuid4())})
            ET.SubElement(external_person, "name").text = inventor
            person_role = ET.SubElement(person_association, "personRole", {
                "pureId": "13480",
                "uri": "/dk/atira/pure/researchoutput/roles/patent/inventor"
            })
            term = ET.SubElement(person_role, "term", {"formatted": "false"})
            ET.SubElement(term, "text", {"locale": "en_US"}).text = "Inventor"
            ET.SubElement(term, "text", {"locale": "es_CO"}).text = "Inventor"

        # Unidades organizacionales
    organisational_units = ET.SubElement(patent, "organisationalUnits")

    if pd.notna(fila['Facultad']):
        organisational_unit = ET.SubElement(organisational_units, "organisationalUnit", {
            "uuid": str(uuid.uuid4()),
            "externalId": "Facultad",
            "externalIdSource": "synchronisedOrganisation",
            "externallyManaged": "true"
        })
        ET.SubElement(organisational_unit, "name", {"formatted": "false"}).text = fila['Facultad']
        unit_type = ET.SubElement(organisational_unit, "type", {
            "pureId": "1079",
            "uri": "/dk/atira/pure/organisation/organisationtypes/organisation/department"
        })
        term = ET.SubElement(unit_type, "term", {"formatted": "false"})
        ET.SubElement(term, "text", {"locale": "en_US"}).text = "Department"
        ET.SubElement(term, "text", {"locale": "es_CO"}).text = "Departamento"

    if pd.notna(fila['Departamento/Instituto']):
        organisational_unit = ET.SubElement(organisational_units, "organisationalUnit", {
            "uuid": str(uuid.uuid4()),
            "externalId": "Departamento/Instituto",
            "externalIdSource": "synchronisedOrganisation",
            "externallyManaged": "true"
        })
        ET.SubElement(organisational_unit, "name", {"formatted": "false"}).text = fila['Departamento/Instituto']
        unit_type = ET.SubElement(organisational_unit, "type", {
            "pureId": "1079",
            "uri": "/dk/atira/pure/organisation/organisationtypes/organisation/department"
        })
        term = ET.SubElement(unit_type, "term", {"formatted": "false"})
        ET.SubElement(term, "text", {"locale": "en_US"}).text = "Department"
        ET.SubElement(term, "text", {"locale": "es_CO"}).text = "Departamento"

    # Añadir datos adicionales
    if pd.notna(fila['IPC/CIP']):
        ipc = ET.SubElement(patent, "ipc")
        ipc.text = fila['IPC/CIP']

    if pd.notna(fila['Estado actual del tramite']):
        estado_tramite = ET.SubElement(patent, "estadoTramite")
        estado_tramite.text = fila['Estado actual del tramite']

    if pd.notna(fila['Fecha de concesion']):
        fecha_concesion = ET.SubElement(patent, "fechaConcesion")
        fecha_concesion.text = str(fila['Fecha de concesion'])

    # Añadir el titular
    if pd.notna(fila['Titular']):
        titular = ET.SubElement(patent, "titular")
        titular.text = fila['Titular']

    # Añadir el grupo de investigación
    if pd.notna(fila['Grupo De Investigacion']):
        grupo_investigacion = ET.SubElement(patent, "grupoInvestigacion")
        grupo_investigacion.text = fila['Grupo De Investigacion']

    # Añadir link de consulta
    if pd.notna(fila['Link De Consulta']):
        link_consulta = ET.SubElement(patent, "linkConsulta")
        link_consulta.text = fila['Link De Consulta']

# Aplicar indentación al XML
indent(root)

# Guardar todas las patentes en un único archivo XML
hoy = datetime.datetime.now().strftime("%Y_%m_%d")
nombre_archivo = f"Resultado/{hoy}_patentes.xml"

tree = ET.ElementTree(root)
tree.write(nombre_archivo, encoding="utf-8", xml_declaration=True)

logging.info(f"Archivo XML generado: {nombre_archivo}")
print(f"Proceso de generación de XML completado. Archivo generado: {nombre_archivo}")
