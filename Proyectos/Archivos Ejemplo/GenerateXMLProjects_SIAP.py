import pandas as pd
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import spacy
import xml.etree.ElementTree as ET
import datetime

# Función para inicializar el detector de idiomas en spaCy
def get_lang_detector(nlp, name):
    return LanguageDetector()

# Cargar el modelo de procesamiento de lenguaje natural en español de spaCy
nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector", func=get_lang_detector)
# Añadir el detector de idiomas al pipeline de spaCy
nlp.add_pipe('language_detector', last=True)

# Crear el elemento raíz del árbol XML con los espacios de nombres adecuados
upmprojects = ET.Element("upmprojects")
upmprojects.set('xmlns', "v1.upmproject.pure.atira.dk")
upmprojects.set('xmlns:ns2', "v3.commons.pure.atira.dk")

# Cargar el archivo Excel que contiene los datos del proyecto
dataframe1 = pd.read_excel(r"Profesores para Perfiles y Capacidades - SIAP Nuevo.xlsx", sheet_name='default_1', dtype=object)

# Leer las columnas únicas de 'id_unico' para identificar proyectos únicos
column = dataframe1.id_unico.unique()

# Iterar sobre cada proyecto único identificado por 'id_unico'
for uid in column:
    # Filtrar el DataFrame para obtener la información del proyecto específico
    temp_project = dataframe1[dataframe1["id_unico"] == uid]
    
    # Extraer el ID del proyecto
    temp_project_id = str(uid)
    
    # Asignar un tipo de proyecto, en este caso "research" (investigación)
    temp_project_type = "research"
    
    # Extraer el título del proyecto
    temp_project_title = temp_project["Titulo del proyecto"].unique()[0]
    
    # Detectar el idioma del título del proyecto
    temp_project_title_lang = nlp(temp_project_title)._.language['language']
    
    # Extraer la descripción del proyecto
    temp_project_description_t = temp_project["descripcion_final"].unique()[0]

    # Extraer la organización que gestiona el proyecto
    temp_project_managing_org = temp_project["owner id"].unique()[0]
    
    # Asignar visibilidad del proyecto, en este caso, se fija a "public"
    temp_project_visibility = "public"
    
    # Extraer la fecha de inicio y finalización del proyecto
    temp_project_startDate = temp_project['Fecha de inicio'].unique()[0]
    temp_project_endDate = temp_project['Fecha final'].unique()[0]
    
    # Extraer la información de los participantes en el proyecto
    temp_project_participants = temp_project[["ID Empleado", "Rol en el proyecto", "Nombres", "Apellidos"]]

    # Extraer información del patrocinador y tipo de financiador
    temp_project_ext_org = temp_project['Nombre patrocinador'].unique()[0]
    temp_project_int_org = temp_project['Tipo de financiador'].unique()[0]

    # Comprobar si la descripción del proyecto es un número (int o float) y reemplazarla por una cadena vacía si es así
    if (isinstance(temp_project_description_t, int) | isinstance(temp_project_description_t, float)):
        temp_project_description = ""
    else:
        # Limpiar la descripción para remover saltos de línea
        temp_project_description = temp_project_description_t.replace("\n", " ")
    
    # Si la descripción está vacía, establecer el idioma predeterminado a español
    if temp_project_description == "":
        temp_project_description_lang = "es"
    else:
        # Detectar el idioma de la descripción del proyecto
        temp_project_description_lang = nlp(temp_project_description)._.language['language']
        
        # Si el idioma detectado no es "es" o "en", forzar a "es"
        if (not(temp_project_description_lang == "es")) | (not(temp_project_description_lang == "en")):
            temp_project_description_lang = "es"

    # Asignar el país correspondiente al idioma del título
    if temp_project_title_lang == "en":
        temp_project_title_country = "US"
    elif temp_project_title_lang == "es":
        temp_project_title_country = "CO"
    else:
        temp_project_title_country = ""

    # Asignar el país correspondiente al idioma de la descripción
    if temp_project_description_lang == "en":
        temp_project_description_country = "US"
    elif temp_project_title_lang == "es":
        temp_project_description_country = "CO"

    #Create project
    upmproject = ET.SubElement(upmprojects,"upmproject", id=temp_project_id,type=temp_project_type)
    #Create title class
    title = ET.SubElement(upmproject,"title")
    if temp_project_title_country=="":
        title_text = ET.SubElement(title,"ns2:text",lang=temp_project_title_lang).text =temp_project_title
    else:
        title_text = ET.SubElement(title,"ns2:text",lang=temp_project_title_lang,country=temp_project_title_country).text =temp_project_title
    #Create description
    if not(temp_project_description == ""):
        description = ET.SubElement(upmproject,"descriptions")
        description_ns2 = ET.SubElement(description,"ns2:description",type="projectdescription")
        description_text = ET.SubElement(description_ns2,"ns2:text",lang=temp_project_description_lang,country=temp_project_description_country).text = temp_project_description
    #Create id
    ids = ET.SubElement(upmproject,"ids")
    ids_ns2 = ET.SubElement(ids,"ns2:id",type="siap").text=temp_project_id
    
    # Create internal participants
    internalParticipants = ET.SubElement(upmproject,"internalParticipants")
    if temp_project_participants['ID Empleado'].isnull().values.any():
        # Create external participants
        externalParticipants = ET.SubElement(upmproject,"externalParticipants")


    for ind in range(len(temp_project_participants)):
        temp_person_id_t = temp_project_participants.iloc[ind,0]
        if (isinstance(temp_person_id_t,int) | isinstance(temp_person_id_t,float)):
            temp_person_id = ""
        else:
            temp_person_id = temp_person_id_t

        temp_person_role = temp_project_participants.iloc[ind,1]
        # if temp_person_id=="":


        #     temp_person_FN = temp_project_participants.iloc[ind,2]
        #     temp_person_LN = temp_project_participants.iloc[ind,3]
        #     temp_person_ExOrg_t = temp_project_participants.iloc[ind,5]

        #     if (isinstance(temp_person_ExOrg_t,int) | isinstance(temp_person_ExOrg_t,float)):
        #         temp_person_ExOrg = ""
        #     else:
        #         temp_person_ExOrg = temp_person_ExOrg_t

        #     externalParticipant = ET.SubElement(externalParticipants,"externalParticipant")
        #     firstName = ET.SubElement(externalParticipant,"firstName").text=temp_person_FN
        #     lastName = ET.SubElement(externalParticipant,"lastName").text = temp_person_LN
        #     role = ET.SubElement(externalParticipant,"role").text=temp_person_role.replace(" ","")
        #     if temp_person_ExOrg != "":
        #         externalOrgName = ET.SubElement(externalParticipant,"externalOrgName").text=temp_person_ExOrg

        # else:
        # temp_person_Org = temp_project_participants.iloc[ind,4]
        internalParticipant = ET.SubElement(internalParticipants,"internalParticipant")
        personId = ET.SubElement(internalParticipant,"personId").text=str(temp_person_id).replace(".0","")
        # organisationIds = ET.SubElement(internalParticipant,"organisationIds")
        # organisation = ET.SubElement(organisationIds,"organisation", id=str(temp_person_Org).replace(".0",""))
        role = ET.SubElement(internalParticipant,"role").text=temp_person_role.replace(" ","")

    if (isinstance(temp_project_ext_org,int) | isinstance(temp_project_ext_org,float)):
        temp_project_ext_org = ""

    if (isinstance(temp_project_int_org,int) | isinstance(temp_project_int_org,float)):
        temp_project_int_org = ""

    if not(temp_project_ext_org==""):
        externalOrganisations = ET.SubElement(upmproject,"externalOrganisations")
        externalOrganisationsAssociation = ET.SubElement(externalOrganisations,"ns2:externalOrganisationAssociation")
        externalOrgName = ET.SubElement(externalOrganisationsAssociation,"ns2:externalOrgName").text=temp_project_ext_org

    if not(temp_project_int_org==""):
        organisations = ET.SubElement(upmproject,"organisations")
        if temp_project_int_org == "PUJ":
            organisation = ET.SubElement(organisations,"organisation",id="PUJAV")
        elif temp_project_int_org == "HUSI":
            organisation = ET.SubElement(organisations,"organisation",id="HUSI")

    managedByOrganisation = ET.SubElement(upmproject,"managedByOrganisation",id=temp_project_managing_org)
    startDate = ET.SubElement(upmproject,"startDate").text=str(temp_project_startDate).replace(" 00:00:00","")
    endDate = ET.SubElement(upmproject,"endDate").text=str(temp_project_endDate).replace(" 00:00:00","")
    visibility = ET.SubElement(upmproject,"visibility").text=temp_project_visibility.replace("p","P")
    workflow = ET.SubElement(upmproject,"workflow").text="validated"
    


#Create XML
# input("Press Enter to continue...\n")
print("Creating XML...")
ET.indent(upmprojects)
tree  = ET.ElementTree(upmprojects)
with open('2024_08_01_Bogota_SIAP.xml','w')as f:
    tree.write('2024_08_01_Bogota_SIAP.xml',encoding="utf-8",xml_declaration=True)

f.close()
print("Complete")