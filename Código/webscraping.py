# Librerias a utilizar
import pandas as pd 
from bs4 import BeautifulSoup
import requests
import re
import word
import registro
import datetime

def datos_handle(hdl,especialista,maestria,referencista, modalidad):

    estudiantes = get_data(hdl , maestria)
    create_certif(estudiantes=estudiantes,especialista=especialista,modalidad=modalidad, referencista=referencista)

def get_data(handle: str, postgrado: str) -> tuple:
    """Busca los datos segun el handle y devuelve una tupla de diccionarios que contienen los datos de cada estudiante"""
    url = handle.strip() + '?mode=full'
    datos_certificado = dict()

    # Lectura de tabla
    table = pd.read_html(handle + '?mode=full')
    df = table[0]
    datos_certificado['nombre'] = tuple(df[df['DC Field'] == 'dc.contributor.author']['Value'])
    datos_certificado['cedula'] = tuple(df[df['DC Field'] == 'dc.ucuenca.idautor']['Value'])

    # Extracción de Html
    pagina = requests.get(handle + '?mode=full')
    soup = BeautifulSoup(pagina.content, 'html.parser')

    # Obtener Facultad y carrera
    colecciones = soup.find_all('ol' , {'class':"breadcrumb btn-success"})
    lista_coleccion = limpiar_html(str(colecciones)).split('\n')
    datos_certificado['facultad'] = lista_coleccion[3]
    datos_certificado["handle"] = "Link: " + handle
    
    if postgrado == "":
        datos_certificado['carrera_maestria'] = "Carrera de " + lista_coleccion[4]
    else:
        datos_certificado['carrera_maestria'] = "de la " + postgrado

    estudiantes = []
    cant_estudiantes = len(datos_certificado['nombre'])
    for num in range(cant_estudiantes):
        estudiantes.append({"nombre":datos_certificado['nombre'][num],
                            "cedula":datos_certificado['cedula'][num],
                            "facultad": datos_certificado['facultad'],
                            "carrera_maestria":datos_certificado['carrera_maestria'],
                            "handle": datos_certificado["handle"]
                            }
                           )
    return tuple(estudiantes)

# Limpiar html
def limpiar_html(texto):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr,'',texto)
    return cleantext

def create_certif(estudiantes: tuple, especialista: str, modalidad: str, referencista: str) -> bool:
    """ Cree los certificados para cada estudiante e inserta en la base de datos a cada registro"""
    fecha_hora = datetime.datetime.now()
    fecha = str(fecha_hora.date())
    hora = str(fecha_hora.time())

    for estudiante in estudiantes:
        registro.insertar_datos(fecha = fecha,
                                    hora = hora,
                                    name = estudiante['nombre'],
                                    ci = estudiante['cedula'],
                                    facultad = estudiante['facultad'],
                                    carr_o_maes = estudiante['carrera_maestria'],
                                    hdl = estudiante['handle'],
                                    firma = especialista,
                                    tipo_cert="Trabajo de titulacion",
                                    referencista = referencista
                                    )
        try:
            word.crear_certificado(estudiante=estudiante['nombre'],
                                   cedula = estudiante['cedula'],
                                   facultad = estudiante['facultad'],
                                   carrera = estudiante['carrera_maestria'],
                                   handle = estudiante['handle'],
                                   especialista = especialista,
                                   modalidad = modalidad
                                   )
        except:
            print("error en crear el certificado")
            return False

    return True

