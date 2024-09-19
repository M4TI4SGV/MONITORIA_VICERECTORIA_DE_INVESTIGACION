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
root = ET.Element("patentes")

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

    # Corregir el nombre de la columna para el tipo de activo
    temp_patente_tipo = str(fila['Tipo de Activo De PI']) if pd.notna(fila['Tipo de Activo De PI']) else "Desconocido"
    
    temp_patente_jurisdiccion = str(fila['Jurisdiccion']) if pd.notna(fila['Jurisdiccion']) else "Desconocido"
    
    # Detectar el idioma del título de la patente
    temp_patente_titulo_lang = nlp(temp_patente_titulo)._.language['language']
    
    # Extraer la descripción de la patente
    temp_patente_descripcion = str(fila['Descripcion']) if pd.notna(fila['Descripcion']) else ""
    
    # Si la descripción está vacía, establecer el idioma predeterminado a español
    if not temp_patente_descripcion:
        temp_patente_descripcion_lang = "es"
    else:
        temp_patente_descripcion_lang = nlp(temp_patente_descripcion)._.language['language']
        if temp_patente_descripcion_lang not in ["es", "en"]:
            temp_patente_descripcion_lang = "es"

    # Asignar el país correspondiente al idioma del título y la descripción
    temp_patente_titulo_country = "CO" if temp_patente_titulo_lang == "es" else "US"
    temp_patente_descripcion_country = "CO" if temp_patente_descripcion_lang == "es" else "US"

    # Crear el elemento de la patente dentro del archivo XML general
    patente = ET.SubElement(root, "patente", id=temp_patente_id, type=temp_patente_tipo)

    # Crear el título de la patente
    title = ET.SubElement(patente, "title")
    ET.SubElement(title, "text", lang=temp_patente_titulo_lang, country=temp_patente_titulo_country).text = temp_patente_titulo

    # Crear la descripción de la patente
    if temp_patente_descripcion:
        description = ET.SubElement(patente, "description")
        ET.SubElement(description, "text", lang=temp_patente_descripcion_lang, country=temp_patente_descripcion_country).text = temp_patente_descripcion

    # Añadir otros elementos de la patente (ej. Jurisdicción)
    jurisdiccion = ET.SubElement(patente, "jurisdiccion")
    jurisdiccion.text = temp_patente_jurisdiccion

    # Añadir otros datos según sea necesario, convirtiendo a cadena cada valor
    if pd.notna(fila['Numero de Solicitud ']):
        numero_solicitud = ET.SubElement(patente, "numeroSolicitud")
        numero_solicitud.text = str(fila['Numero de Solicitud '])

# Aplicar indentación al XML
indent(root)

# Guardar todas las patentes en un único archivo XML
hoy = datetime.datetime.now().strftime("%Y_%m_%d")
nombre_archivo = f"{hoy}_patentes.xml"

tree = ET.ElementTree(root)
tree.write(f"Resultado/{nombre_archivo}", encoding="utf-8", xml_declaration=True)

logging.info(f"Archivo XML generado: {nombre_archivo}")

print("Proceso de generación de XML completado.")

