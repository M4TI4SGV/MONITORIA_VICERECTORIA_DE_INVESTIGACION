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
dataframe_patentes = pd.read_excel(r"pruebas.xlsx", sheet_name='Hoja1', dtype=object)

# Verificar las columnas disponibles
print("Columnas disponibles en la hoja de Excel:", dataframe_patentes.columns)

# Crear la carpeta Resultado si no existe
if not os.path.exists("Resultado"):
    os.makedirs("Resultado")

# Crear el elemento raíz del archivo XML que agrupará todas las patentes
root = ET.Element("publications", {
    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xsi:schemaLocation": "https://puj-staging.elsevierpure.com/ws/api/524 https://puj-staging.elsevierpure.com/ws/api/524/xsd/schema1.xsd"
})

# Función para crear el prefijo 'ns2' en los campos de texto
def create_ns2_text(element, text):
    ns2 = ET.SubElement(element, "{ns2}text")
    ns2.text = text

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

    # Verificar si la patente con el mismo título ya ha sido procesada
    if temp_patente_titulo in titulos_procesados:
        print(f"Patente con título '{temp_patente_titulo}' ya procesada, omitiendo...")
        continue  # Si ya fue procesada, omitimos esta patente

    # Agregar el título de la patente al conjunto para marcarla como procesada
    titulos_procesados.add(temp_patente_titulo)
    
    # Crear el elemento de la patente dentro del archivo XML general
    patent = ET.SubElement(root, "patent")

    # Crear el título de la patente
    title = ET.SubElement(patent, "title", {"formatted": "true"})
    create_ns2_text(title, temp_patente_titulo)

    # Crear el tipo de la patente
    patent_type = ET.SubElement(patent, "type")
    term = ET.SubElement(patent_type, "term", {"formatted": "false"})
    create_ns2_text(term, "Invention patent")
    create_ns2_text(term, "Patente de Invención")

    # Categoría de la patente
    category = ET.SubElement(patent, "category")
    category_term = ET.SubElement(category, "term", {"formatted": "false"})
    create_ns2_text(category_term, "Research")
    create_ns2_text(category_term, "Investigación")

    # Crear la descripción de la patente
    if pd.notna(fila['Descripcion']):
        abstract = ET.SubElement(patent, "abstract", {"formatted": "false"})
        abstract_text = ET.SubElement(abstract, "text", {"locale": "es_CO"})
        abstract_text.text = f"<![CDATA[{fila['Descripcion']}]]>"

    # Añadir los datos de publicación
    publication_statuses = ET.SubElement(patent, "publicationStatuses")
    publication_status = ET.SubElement(publication_statuses, "publicationStatus", {"current": "true"})
    pub_status = ET.SubElement(publication_status, "publicationStatus")
    term = ET.SubElement(pub_status, "term", {"formatted": "false"})
    create_ns2_text(term, "Published")
    create_ns2_text(term, "Publicada")

    # Fecha de publicación
    if pd.notna(fila['Fecha de Solicitud']):
        publication_date = ET.SubElement(publication_status, "publicationDate")
        fecha_solicitud = str(fila['Fecha de Solicitud']).split(" ")[0]
        date_parts = fecha_solicitud.split("-")
        if len(date_parts) == 3:
            year = ET.SubElement(publication_date, "year")
            create_ns2_text(year, date_parts[0])
            month = ET.SubElement(publication_date, "month")
            create_ns2_text(month, date_parts[1])
            day = ET.SubElement(publication_date, "day")
            create_ns2_text(day, date_parts[2])

    # Idioma
    language = ET.SubElement(patent, "language")
    language_term = ET.SubElement(language, "term", {"formatted": "false"})
    create_ns2_text(language_term, "Spanish")
    create_ns2_text(language_term, "Español")

    # Cambiar jurisdicción a country
    if pd.notna(fila['Jurisdiccion']):
        country = ET.SubElement(patent, "country")
        term = ET.SubElement(country, "term", {"formatted": "false"})
        create_ns2_text(term, fila['Jurisdiccion'])
        create_ns2_text(term, fila['Jurisdiccion'])

    # Añadir número de patente
    if pd.notna(fila['Numero de Solicitud ']):
        patent_number = ET.SubElement(patent, "patentNumber")
        create_ns2_text(patent_number, str(fila['Numero de Solicitud ']))

    # Asociaciones de personas (inventores)
if pd.notna(fila['Inventores/Autores']):
    persons = ET.SubElement(patent, "persons")  # Cambia personAssociations por persons
    inventores = str(fila['Inventores/Autores']).split(",")

    for inventor in inventores:
        # Asegurarnos de quitar espacios en blanco innecesarios
        inventor = inventor.strip()

        # Separar el nombre completo en una lista de palabras
        nombres = inventor.split()

        # Verificar que haya al menos un nombre y un apellido
        if len(nombres) > 2:
            # Si hay más de 2 palabras, las dos últimas se consideran apellidos
            first_name = " ".join(nombres[:-2])  # Todos los elementos menos los últimos dos son nombres
            last_name = " ".join(nombres[-2:])   # Los últimos dos elementos son apellidos
        elif len(nombres) == 2:
            # Si hay exactamente dos palabras, una es el nombre y otra el apellido
            first_name = nombres[0]
            last_name = nombres[1]
        else:
            # Si solo hay una palabra, la asignamos como nombre y dejamos el apellido vacío
            first_name = nombres[0]
            last_name = ""

        # Crear la estructura de author y person para cada inventor
        author = ET.SubElement(persons, "author")
        role = ET.SubElement(author, "role")
        role.text = "inventor"  # Asignamos el rol de inventor

        person = ET.SubElement(author, "person")
        first_name_elem = ET.SubElement(person, "firstName")
        last_name_elem = ET.SubElement(person, "lastName")

        first_name_elem.text = first_name  # Primer nombre o nombres
        last_name_elem.text = last_name    # Apellido o apellidos

    # Unidades organizacionales
    organisational_units = ET.SubElement(patent, "organisationalUnits")

    if pd.notna(fila['Facultad']):
        organisational_unit = ET.SubElement(organisational_units, "organisationalUnit")
        name = ET.SubElement(organisational_unit, "name", {"formatted": "false"})
        create_ns2_text(name, fila['Facultad'])
        unit_type = ET.SubElement(organisational_unit, "type")
        term = ET.SubElement(unit_type, "term", {"formatted": "false"})
        create_ns2_text(term, "Department")
        create_ns2_text(term, "Departamento")

    if pd.notna(fila['Departamento/Instituto']):
        organisational_unit = ET.SubElement(organisational_units, "organisationalUnit")
        name = ET.SubElement(organisational_unit, "name", {"formatted": "false"})
        create_ns2_text(name, fila['Departamento/Instituto'])
        unit_type = ET.SubElement(organisational_unit, "type")
        term = ET.SubElement(unit_type, "term", {"formatted": "false"})
        create_ns2_text(term, "Department")
        create_ns2_text(term, "Departamento")

    # Manejar las organizaciones (organisationalUnits)
    if pd.notna(fila['Titular']):
        organisations = ET.SubElement(patent, "organisations")  # Crear el elemento organisations
        
        # Dividir las organizaciones por el guion "-" y procesarlas individualmente
        titulares = str(fila['Titular']).split(" - ")
        
        for titular in titulares:
            organisation = ET.SubElement(organisations, "organisation")
            name = ET.SubElement(organisation, "name")
            create_ns2_text(name, titular.strip())  # Usar create_ns2_text para insertar el nombre


        # Añadir datos adicionales como IPC/CIP y Estado del trámite
        if pd.notna(fila['IPC/CIP']):
            ipc = ET.SubElement(patent, "ipc")
            create_ns2_text(ipc, fila['IPC/CIP'])

        if pd.notna(fila['Estado actual del tramite']):
            estado_tramite = ET.SubElement(patent, "estadoTramite")
            create_ns2_text(estado_tramite, fila['Estado actual del tramite'])

        # Añadir la fecha de concesión
        if pd.notna(fila['Fecha de concesion']):
            priority_date = ET.SubElement(patent, "priorityDate")
            create_ns2_text(priority_date, str(fila['Fecha de concesion']).split(" ")[0])

        # Añadir el titular (separado en organizacionales)
        if pd.notna(fila['Titular']):
            organisational_units = ET.SubElement(patent, "organisationalUnits")
            for titular in str(fila['Titular']).split("-"):
                organisational_unit = ET.SubElement(organisational_units, "organisationalUnit")
                name = ET.SubElement(organisational_unit, "name", {"formatted": "false"})
                create_ns2_text(name, titular.strip())
                unit_type = ET.SubElement(organisational_unit, "type")
                term = ET.SubElement(unit_type, "term", {"formatted": "false"})
                create_ns2_text(term, "Department")
                create_ns2_text(term, "Departamento")

    # Añadir el grupo de investigación
    if pd.notna(fila['Grupo De Investigacion']):
        grupo_investigacion = fila['Grupo De Investigacion']
        organisational_unit = ET.SubElement(organisational_units, "organisationalUnit")
        name = ET.SubElement(organisational_unit, "name", {"formatted": "false"})
        create_ns2_text(name, grupo_investigacion)
        unit_type = ET.SubElement(organisational_unit, "type")
        term = ET.SubElement(unit_type, "term", {"formatted": "false"})
        create_ns2_text(term, "Department")
        create_ns2_text(term, "Departamento")

        # Managing organization
        managing_unit = ET.SubElement(patent, "managingOrganisationalUnit")
        name = ET.SubElement(managing_unit, "name", {"formatted": "false"})
        create_ns2_text(name, grupo_investigacion)

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
