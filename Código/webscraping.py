# Librerias a utilizar
import pandas as pd 
from bs4 import BeautifulSoup
import requests
import re
import word
import registro
import datetime

def datos_handle(hdl,especialista,grado,maestria,referencista):
    handle = hdl
    fecha_hora = datetime.datetime.now()
    fecha = str(fecha_hora.date())
    hora = str(fecha_hora.time())

    # Lectura de tabla
    table = pd.read_html(handle + '?mode=full')
    df = table[0]
    autores = list(df[df['DC Field'] == 'dc.contributor.author']['Value'])
    cedulas = list(df[df['DC Field'] == 'dc.ucuenca.idautor']['Value'])

    # Extracción de Html
    pagina = requests.get(handle + '?mode=full')
    soup = BeautifulSoup(pagina.content, 'html.parser')

    # Limpiar html
    def limpiar_html(texto):
        cleanr = re.compile('<.*?>')
        cleantext = re.sub(cleanr,'',texto)
        return cleantext\

    # Obtener Facultad y carrera
    colecciones = soup.find_all('ol' , {'class':"breadcrumb btn-success"})
    lista_coleccion = limpiar_html(str(colecciones)).split('\n')
    facultad = lista_coleccion[3]
    carrera = lista_coleccion[4]
    handle = "Link: " + handle

    if grado == 0:
        carrera_o_maestria = "Carrera de " + carrera
    else:
        carrera_o_maestria = 'de la ' + maestria

    if len(autores) == 2:
        word.crear_certificado(estudiante = autores[0],cedula = cedulas[0],facultad = facultad ,carrera = carrera_o_maestria,handle = handle,especialista = especialista)
        word.crear_certificado(estudiante = autores[1],cedula = cedulas[1],facultad = facultad ,carrera = carrera_o_maestria,handle = handle,especialista = especialista)
        registro.insertar_datos(fecha = fecha , hora = hora ,name = autores[0] , ci = cedulas[0] , facultad = facultad, carr_o_maes = carrera_o_maestria, hdl = handle, firma = especialista,tipo_cert="Trabajo de titulacion",referencista = referencista)
        registro.insertar_datos(fecha = fecha , hora = hora ,name = autores[1] , ci = cedulas[1] , facultad = facultad, carr_o_maes = carrera_o_maestria, hdl = handle, firma = especialista,tipo_cert="Trabajo de titulacion", referencista = referencista)

    else:
        word.crear_certificado(estudiante = autores[0],cedula = cedulas[0],facultad = facultad ,carrera = carrera_o_maestria,handle = handle,especialista = especialista)
        registro.insertar_datos(fecha = fecha , hora = hora ,name = autores[0] , ci = cedulas[0] , facultad = facultad, carr_o_maes = carrera_o_maestria, hdl = handle, firma = especialista,tipo_cert="Trabajo de titulacion", referencista = referencista)
    return True

