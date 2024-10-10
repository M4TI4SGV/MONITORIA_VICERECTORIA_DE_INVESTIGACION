import os
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

# Cargar el archivo Excel que contiene los datos de las patentes
dataframe_patentes = pd.read_excel(r"pruebas.xlsx", sheet_name='Hoja1', dtype=object)

# Crear la carpeta Resultado si no existe
if not os.path.exists("Resultado"):
    os.makedirs("Resultado")

# Crear el elemento raíz del archivo XML que agrupará todas las publicaciones
root = ET.Element("publications", {
    "xmlns": "v1.publication-import.base-uk.pure.atira.dk",
    "xmlns:ns2": "v3.commons.pure.atira.dk"
})

# Función auxiliar para crear etiquetas <ns2:text>
def create_ns2_text(parent, text_value):
    ns2_text = ET.SubElement(parent, "ns2:text")
    ns2_text.text = text_value

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
    temp_patente_titulo = str(fila['Titulo de la Patente']) if pd.notna(fila['Titulo de la Patente']) else "Sin título"
    tipo_activo = str(fila['Tipo de Activo De PI']) if pd.notna(fila['Tipo de Activo De PI']) else ""

    # Filtrar para procesar solo los tipos "Patente - Invención" y "Patente - PCT"
    if tipo_activo not in ["Patente - Invención", "Patente - PCT"]:
        print(f"Omitiendo tipo de activo: {tipo_activo}")
        continue

    # Verificar si la patente con el mismo título ya ha sido procesada
    if temp_patente_titulo in titulos_procesados:
        print(f"Patente con título '{temp_patente_titulo}' ya procesada, omitiendo...")
        continue  # Si ya fue procesada, omitimos esta patente

    # Agregar el título de la patente al conjunto para marcarla como procesada
    titulos_procesados.add(temp_patente_titulo)

    # Crear el elemento de la patente dentro del archivo XML general
    patent = ET.SubElement(root, "patent", {"subType": tipo_activo})

    # Añadir el campo peerReviewed y workflow (fijos)
    peer_reviewed = ET.SubElement(patent, "peerReviewed")
    peer_reviewed.text = "false"
    workflow = ET.SubElement(patent, "workflow")
    workflow.text = "approved"

    # Añadir los datos de publicación
    publication_statuses = ET.SubElement(patent, "publicationStatuses")
    publication_status = ET.SubElement(publication_statuses, "publicationStatus")
    status_type = ET.SubElement(publication_status, "statusType")
    status_type.text = "published"
    date = ET.SubElement(publication_status, "date")

    if pd.notna(fila['Fecha de Solicitud']):
        fecha_solicitud = str(fila['Fecha de Solicitud']).split(" ")[0]
        date_parts = fecha_solicitud.split("-")
        if len(date_parts) == 3:
            year = ET.SubElement(date, "ns2:year")
            year.text = date_parts[0]
            month = ET.SubElement(date, "ns2:month")
            month.text = date_parts[1]
            day = ET.SubElement(date, "ns2:day")
            day.text = date_parts[2]

    # Añadir la etiqueta del idioma
    language = ET.SubElement(patent, "language")
    language.text = "es_CO"

    # Crear el título de la patente
    title = ET.SubElement(patent, "title")
    create_ns2_text(title, temp_patente_titulo)

    # Crear el campo abstract con idioma y país detectado
    if pd.notna(fila['Descripcion']) and fila['Descripcion'].strip() != "":
        doc = nlp(fila['Descripcion'])
        detected_language = doc._.language['language']
        country = "CO" if detected_language == "es" else "GB"
        abstract = ET.SubElement(patent, "abstract")
        abstract_text = ET.SubElement(abstract, "ns2:text", {
            "lang": detected_language,
            "country": country
        })
        abstract_text.text = fila['Descripcion']
    else:
        print(f"Advertencia: No hay descripción para la patente en la fila {index + 1}")
        abstract = ET.SubElement(patent, "abstract")
        abstract_text = ET.SubElement(abstract, "ns2:text", {"lang": "es", "country": "CO"})
        abstract_text.text = "Descripción no disponible"

    # Añadir los inventores
    if pd.notna(fila['Inventores/Autores']):
        persons = ET.SubElement(patent, "persons")
        inventores = str(fila['Inventores/Autores']).split(",")

        for inventor in inventores:
            nombres = inventor.strip().split()
            first_name = " ".join(nombres[:-2])
            last_name = " ".join(nombres[-2:])

            author = ET.SubElement(persons, "author")
            role = ET.SubElement(author, "role")
            role.text = "inventor"

            person = ET.SubElement(author, "person")
            first_name_elem = ET.SubElement(person, "firstName")
            last_name_elem = ET.SubElement(person, "lastName")
            first_name_elem.text = first_name
            last_name_elem.text = last_name

    # Crear el campo de organizaciones y separar por guion ("-")
    if pd.notna(fila['Titular']):
        organisations = ET.SubElement(patent, "organisations")
        titulares = str(fila['Titular']).split(" - ")
        for titular in titulares:
            organisation = ET.SubElement(organisations, "organisation")
            name = ET.SubElement(organisation, "name")
            create_ns2_text(name, titular.strip())

    # Añadir el owner con el ID "PUJAV"
    owner = ET.SubElement(patent, "owner", {"id": "PUJAV"})

    # Crear el campo número de patente
    if pd.notna(fila['Numero de Solicitud ']):
        patent_number = ET.SubElement(patent, "patentNumber")
        numero_solicitud_limpio = str(fila['Numero de Solicitud ']).replace("\n", " ").strip()
        create_ns2_text(patent_number, numero_solicitud_limpio)

    # Añadir la fecha de prioridad
    if pd.notna(fila['Fecha de concesion']):
        priority_date = ET.SubElement(patent, "priorityDate")
        priority_date_text = ET.SubElement(priority_date, "ns2:text")
        priority_date_text.text = str(fila['Fecha de concesion']).split(" ")[0]

# Aplicar indentación al XML
indent(root)

# Guardar todas las patentes en un único archivo XML
hoy = datetime.datetime.now().strftime("%Y_%m_%d")
nombre_archivo = f"Resultado/{hoy}_patentes.xml"

tree = ET.ElementTree(root)
tree.write(nombre_archivo, encoding="utf-8", xml_declaration=True)

logging.info(f"Archivo XML generado: {nombre_archivo}")
print(f"Proceso de generación de XML completado. Archivo generado: {nombre_archivo}")
