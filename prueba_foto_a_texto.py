import pytesseract
from PIL import Image
import pandas as pd
import numpy as np
from dotenv import load_dotenv
import os

load_dotenv()

# Establece la ruta al tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Cargar la imagen y extraer texto con pytesseract
def imagen_a_texto(ruta_de_la_imagen):
    imagen = Image.open(ruta_de_la_imagen)
    texto = pytesseract.image_to_string(imagen)
    return texto

# Procesar el texto extraído para transformarlo en lista
def procesar_texto(texto):
    data = [linea for linea in texto.strip().split("\n") if linea != '']
    return data

#Extrae la ruta de la imagen del .env
RUTA_IMAGEN = os.getenv('RUTA_IMAGEN')

texto_extraido = imagen_a_texto(fr'{RUTA_IMAGEN}')

#Trasforma el texto en lista para procesar el Dataframe
lista_texto = procesar_texto(texto_extraido)

#Extre los posiciones necesarias para el dataframe
index_Year = lista_texto.index('Year')+1
index_Category = lista_texto.index('Category')+1
index_Product = lista_texto.index('Product')+1
index_Sales = lista_texto.index('Sales')+1
index_Rating = lista_texto.index('Rating')+1

#Crea el diccionario
info_a_excel = {'Year' : lista_texto[index_Year:index_Category-1],
                'Category' : lista_texto[index_Category:index_Product-1],
                'Product' : lista_texto[index_Product:index_Sales-1],
                'Sales' : lista_texto[index_Sales:index_Rating-1],
                'Rating' : lista_texto[index_Rating: ]}


#Revisa cual es la lista más larga
max_len = max(len(col) for col in info_a_excel.values())

#Rellena las listas que no tienen la cantidad necesaria con Nan
for key in info_a_excel:
    info_a_excel[key].extend([np.nan] * (max_len - len(info_a_excel[key])))

#Transforma el diccionario a Dataframe
info_a_excel = pd.DataFrame(info_a_excel)

#Extrae la ruta donde se creara el excel del .env
RUTA_EXCEL = os.getenv('RUTA_EXCEL')

name = (Fr'{RUTA_EXCEL}')

#Convierte el dataframe a excel
with pd.ExcelWriter(name) as ExcelWriter:
    info_a_excel.to_excel(excel_writer=ExcelWriter, index=False)

