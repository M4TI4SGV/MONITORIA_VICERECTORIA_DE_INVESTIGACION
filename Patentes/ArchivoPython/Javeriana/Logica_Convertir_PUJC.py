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

# Cargar el archivo Excel que contiene los datos de las patentes desde la pestaña 'Hoja1'
dataframe_patentes = pd.read_excel(r"Pruebas_Patentes_PUJC.xlsx", sheet_name='Patentes', dtype=object)

# Verificar las columnas disponibles
print("Columnas disponibles en la hoja de Excel:", dataframe_patentes.columns)

# Crear la carpeta Resultado Cali si no existe
if not os.path.exists("Resultado Cali"):
    os.makedirs("Resultado Cali")

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

# Función para separar correctamente los nombres y apellidos
def separar_nombres(nombre_completo):
    nombres = nombre_completo.strip().split()
    if len(nombres) > 2:
        first_name = " ".join(nombres[:-2])  # Todas las palabras menos las dos últimas son nombres
        last_name = " ".join(nombres[-2:])   # Las últimas dos palabras son apellidos
    elif len(nombres) == 2:
        first_name = nombres[0]  # La primera palabra es el nombre
        last_name = nombres[1]   # La segunda palabra es el apellido
    else:
        first_name = nombres[0]  # Si solo hay una palabra, se considera como nombre
        last_name = ""  # No hay apellido
    return first_name, last_name

# Iterar sobre todas las filas del DataFrame para procesar cada patente
for index, fila in dataframe_patentes.iterrows():
    print(f"Procesando la fila {index + 1} de {len(dataframe_patentes)}")

    # Extraer la información clave de la patente
    temp_patente_titulo = str(fila['titulo']) if pd.notna(fila['titulo']) else "Sin título"

    # Verificar si la patente con el mismo título ya ha sido procesada
    if temp_patente_titulo in titulos_procesados:
        print(f"Patente con título '{temp_patente_titulo}' ya procesada, omitiendo...")
        continue  # Si ya fue procesada, omitimos esta patente

    # Agregar el título de la patente al conjunto para marcarla como procesada
    titulos_procesados.add(temp_patente_titulo)
    
    # Crear el elemento de la patente dentro del archivo XML general
    patent = ET.SubElement(root, "patent", {"subType": "Patente - Invención"})

    # Crear el título de la patente
    title = ET.SubElement(patent, "title")
    create_ns2_text(title, temp_patente_titulo)

    # Organizaciones
    if pd.notna(fila['departamento']):
        organisations = ET.SubElement(patent, "organisations")
        organizaciones = str(fila['departamento']).split(" - ")
        for org in organizaciones:
            organisation = ET.SubElement(organisations, "organisation")
            name = ET.SubElement(organisation, "name")
            create_ns2_text(name, org.strip())
    
    # Añadir el propietario
    owner = ET.SubElement(patent, "owner", {"id": "PUJAV"})

    # Añadir el número de patente
    if pd.notna(fila['Número de patente']):
        patent_number = ET.SubElement(patent, "patentNumber")
        create_ns2_text(patent_number, str(fila['Número de patente']).replace("\n", " ").strip())

    # Crear el abstract
    if pd.notna(fila['resumen']):
        abstract = ET.SubElement(patent, "abstract")
        abstract_text = ET.SubElement(abstract, "ns2:text", {
            "lang": "es",
            "country": str(fila['pais']).upper() if pd.notna(fila['pais']) else "CO"
        })
        abstract_text.text = fila['resumen']

    # Personas asociadas (inventores y principal investigador)
    persons = ET.SubElement(patent, "persons")

    # Inventores
    if pd.notna(fila['nombres']) and pd.notna(fila['apellidos']):
        nombres = str(fila['nombres']).split(",")
        apellidos = str(fila['apellidos']).split(",")
        for nombre, apellido in zip(nombres, apellidos):
            first_name, last_name = separar_nombres(nombre + " " + apellido)

            author = ET.SubElement(persons, "author")
            person = ET.SubElement(author, "person")
            ET.SubElement(person, "firstName").text = first_name
            ET.SubElement(person, "lastName").text = last_name

    # Publicación y fechas
    publication_statuses = ET.SubElement(patent, "publicationStatuses")
    publication_status = ET.SubElement(publication_statuses, "publicationStatus")
    status_type = ET.SubElement(publication_status, "statusType")
    status_type.text = "published"
    date = ET.SubElement(publication_status, "date")
    
    # Fecha de publicación (year, month, day) con el formato <ns2:year>, <ns2:month>, <ns2:day>
    if pd.notna(fila['fecha_publicacion']):
        try:
            fecha_publicacion = pd.to_datetime(fila['fecha_publicacion'], dayfirst=True)
            ET.SubElement(date, "ns2:year").text = str(fecha_publicacion.year)
            ET.SubElement(date, "ns2:month").text = str(fecha_publicacion.month).zfill(2)
            ET.SubElement(date, "ns2:day").text = str(fecha_publicacion.day).zfill(2)
        except Exception as e:
            logging.warning(f"Error procesando la fecha de publicación en la fila {index + 1}: {e}")
    else:
        # Si no hay fecha de publicación, eliminar el campo date vacío
        publication_statuses.remove(date)

    # Idioma
    language = ET.SubElement(patent, "language")
    language.text = "es_CO"  # Cambiado de "es" a "es_CO" como en Bogotá

    # Prioridad de la patente con el formato <ns2:text>
    if pd.notna(fila['fecha_prioridad']):
        priority_date = ET.SubElement(patent, "priorityDate")
        create_ns2_text(priority_date, str(fila['fecha_prioridad']).split(" ")[0])

# Aplicar indentación al XML
indent(root)

# Guardar todas las patentes en un único archivo XML
hoy = datetime.datetime.now().strftime("%Y_%m_%d")
nombre_archivo = f"Resultado Cali/{hoy}_patentes.xml"

tree = ET.ElementTree(root)
tree.write(nombre_archivo, encoding="utf-8", xml_declaration=True)

logging.info(f"Archivo XML generado: {nombre_archivo}")
print(f"Proceso de generación de XML completado. Archivo generado: {nombre_archivo}")
