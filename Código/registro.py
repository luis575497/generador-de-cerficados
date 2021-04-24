# Conectar la base de datos con la aplicación
# Importaciones principales
import sqlite3
import os
import pandas as pd

def insertar_datos(fecha, hora, name , ci, facultad, carr_o_maes, hdl, firma,referencista,tipo_cert):
    #Datos a Insertar
    especialista = firma.split('\n')[1]
    datos = [fecha, hora , name, ci, facultad, carr_o_maes, hdl, especialista, referencista, tipo_cert]

    # Conectar con la base de datos sino existe crear una y crear un cursor
    if os.path.isfile('./recursos/registro de certificados.db') == True:
        conector = sqlite3.connect('./recursos/registro de certificados.db')
        cursor = conector.cursor()
    else:
        conector = sqlite3.connect('./recursos/registro de certificados.db')
        cursor = conector.cursor()
        conector.execute('''CREATE TABLE "certificados" ("id"	INTEGER UNIQUE,"Fecha"	TEXT,"Hora"	TEXT,"Nombres y Apellidos"	TEXT,"Cédula"	TEXT,"Facultad"	TEXT,"Carrera o Maestría"	TEXT,"Handle"	TEXT,"Fimado por"  TEXT,  "Referencista"	TEXT,	"Tipo"	TEXT,PRIMARY KEY("id" AUTOINCREMENT))''')

    # Insertar en la Tabla
    query_insert = 'INSERT INTO certificados VALUES (NULL, ? , ? , ? , ? , ? , ? , ? , ?, ?, ?)'
    cursor.execute(query_insert,datos)

    # Actualizar los cambios
    conector.commit()

    # Cerrar la conexión con la base de datos
    conector.close()

def crear_registro(fecha):
    # Conectar con la base de datos sino existe crear una y crear un cursor
    if os.path.isfile('./recursos/registro de certificados.db') == True:
        fecha_val = all([x.isdigit() for x in fecha.split("-")])
        if fecha_val == True and len(fecha.split("-")) < 4:
            year = fecha+'%'
            conector = sqlite3.connect('./recursos/registro de certificados.db')
            query = f"SELECT * FROM certificados WHERE Fecha LIKE '{year}'"
            data_from_certificados = pd.DataFrame(pd.read_sql_query(query , conector))
            if os.path.isfile("./Reporte de Certificados.xlsx") == True:
                os.remove("./Reporte de Certificados.xlsx")
                data_from_certificados.to_excel("./Reporte de Certificados.xlsx", sheet_name = "Registro" , index = False)
            else:
                data_from_certificados.to_excel("./Reporte de Certificados.xlsx", sheet_name = "Registro" , index = False)
            conector.commit()
            conector.close()
            return "Exito en la operacion"
        else:
            return "Formato incorrecto"

def buscar(consulta):
    if os.path.isfile("./recursos/registro de certificados.db") == True:
        conector = sqlite3.connect("./recursos/registro de certificados.db")
        cursor = conector.cursor()
        cursor.execute(consulta)
        datos = cursor.fetchall()
        conector.commit()
        conector.close
        return datos
