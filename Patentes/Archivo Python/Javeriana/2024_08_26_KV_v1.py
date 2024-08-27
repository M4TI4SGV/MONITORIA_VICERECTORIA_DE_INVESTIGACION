import pandas as pd
import spacy
from spacy.language import Language
from spacy_langdetect import LanguageDetector
import xml.etree.ElementTree as ET
import logging
import datetime

# Initialize logging
logging.basicConfig(filename='xml_conversion.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Define language detector 
def get_lang_detector(nlp, name):
    return LanguageDetector()

nlp = spacy.load("es_core_news_sm")
Language.factory("language_detector", func=get_lang_detector)
nlp.add_pipe('language_detector', last=True)

#Define root node
piprojects = ET.Element("piprojects")
piprojects.set('xmlns',"v1.piproject.pure.atira.dk")
piprojects.set('xmlns:ns2',"v3.commons.pure.atira.dk")

# Read data from parent folder
data = pd.read_excel('Productos de PI Javeriana.xlsx')

# Iterate over rows
column = data.id_unico.unique()


# print some data
if data is not None:
    print(data.head())
