# Crear la interfaz de usuario
from tkinter import *
import word
import webscraping
import registro
from PIL import ImageTk,Image
import sqlite3
from tkinter import ttk
import datetime

# Variables de datos
paraiso_esp = '________________________________________\nLcda. Marghot Elizabeth Maza\nEspecialista de Biblioteca Campus Paraíso \n CDR “Juan Bautista Vázquez”'
central_esp = '________________________________________\nLcda. Ximena Carrasco Aguilar\nEspecialista de Biblioteca Campus Central \n CDR “Juan Bautista Vázquez”'
yanuncay_esp = '________________________________________\nLcda. Ximena Carrasco Aguilar\nEspecialista de Biblioteca Campus Yanuncay \n CDR “Juan Bautista Vázquez”'
coordinacion = '________________________________________\nLcda. Rocío Campoverde Carpio, Mg.\nCoordinadora General \n CDR “Juan Bautista Vázquez”'
centrohist = '________________________________________\nLcda. Diana Fajardo Pasán \nBibliotecaria Campus Centro Histórico\n CDR “Juan Bautista Vázquez”'
# Facultades
facultades = {"Arquitectura y Urbanismo":["Arquitectura"], 
              "Artes":["Diseño Gráfico","Diseño de interiores","Artes Visuales","Artes Músicales","Artes Escénicas"],
              "Ciencias Agropecuarias":["Ingeniería Agronómica","Medicina Veterinaria y Zootecnia"],
              "Ciencias de la Hospitalidad":["Gastronomía","Hospitalidad y Hotelería","Turismo"], 
              "Ciencias Económicas y Administrativas":["Administración de Empresas","Economía","Mercadotecnia","Contabilidad y Auditoría","Sociología"],
              "Ciencias Médicas":["Enfermería","Laboratorio Clínico","Estimulación Temprana","Medicina","Fonoadiología","Nutrición y Dietética","Imagenología y Radiología","Fisioterapia"],
              "Ciencias Químicas":["Bioquímica y Farmacia","Ingeniería Industrial","Ingeniería Ambiental","Ingeniería Química"],
              "Odontología":["Odontología"],
              "Filosofía, Letras y Ciencias de la Educación":["Pedagogía de las Artes y las Humanidades",
                                                              "Pedagogía de las Ciencias Experimentales",
                                                              "Pedagogía de la Historia y las Ciencias Sociales",
                                                              "Pedagogía de la Actividad Física y Deporte",
                                                              "Pedagogía de la Lengua y Literatura",
                                                              "Pedagogía de los idiomas Nacionales y Extranjeros",
                                                              "Educación Inicial",
                                                              "Educación General Básica",
                                                              "Cine",
                                                              "Comunicación",
                                                              "Periodismo",
                                                              "Educación Básica",
                                                              "Cine y Audiovisuales",
                                                              "Lengua y Literatura Inglesa",
                                                              "Matemáticas y Física",
                                                              "Cultura Física",
                                                              "Lengua, Literatura y Lenguajes Audiovisuales",
                                                              "Comunicación Social en Comunicación Organizacional y Relaciones Públicas",
                                                              "Comunicación Social en Periodismo y Comunicación Organizacional",
                                                              "Filosofía, Sociología y Economía",
                                                              "Historia y Geografía",
                                                              ],
              "Ingeniería":["Ingeniería Civil","Ingeniería Eléctrica","Ingeniería de Sistemas","Ingeniería en Electrónica y Telecomnicaciones"],
              "Jurisprudencia y Ciencias Políticas y Sociales":["Derecho","Trabajo Social","Orientación Familiar","Género y Desarrollo"],
              "Psicología":["Psicología Clínica","Psicología Social","Psicología Educativa","Licenciatura en Psicología"]
              }
# Funcion para seleccionar la carrera
def seleccionfac(values):
    global lista_carrera
    i = lista_facultad.get()
    lista_carrera = ttk.Combobox(pest_movil, values = facultades[i])
    lista_carrera.configure(width=36, height=30)
    lista_carrera.grid(row=8, column=0, columnspan=2, pady=5)


# Funcion de generar certificado
def certificado():
    handl = cajahandle.get()
    nombre_maestria = caja_maestria.get()
    
    try:
        print(nombre_maestria)
        webscraping.datos_handle(hdl = handl, especialista = espe.get(), maestria = nombre_maestria,referencista= referencista.get(), modalidad = virtual.get())
        etiqueta_datos = Label(pest_tesis, text = "  Su Certificado se ha generado exitosamente  ")
        etiqueta_datos.grid(row = 3, column = 0, columnspan = 2)
        cajahandle.delete(0, END)
        caja_maestria.delete(0, END)
        referencista.delete(0, END)
        if len(nombre_maestria) != 0:
            caja_maestria.delete(0, END)
    except:
        etiqueta_datos = Label(pest_tesis, text = "     No se ha podido generar el certificado    ")
        etiqueta_datos.grid(row = 3, column = 0, columnspan = 2)
        cajahandle.delete(0, END)

# Función para generar el Registro
def butt_reporte():
    
    def crear_reg():
        year = year_reporte.get()
        registro.crear_registro(year)
        year_reporte.delete(0, END)
    
    ventana_reporte = Tk()
    ventana_reporte.geometry("230x180")
    ventana_reporte.resizable(width = 0, height = 0)
    ventana_reporte.title('Generar certificados')
    
    etiqueta_reporte = Label(ventana_reporte, text = 'Escriba el año a consultar\nde la forma año-mes-dia\n(2020-03-03)')
    etiqueta_reporte.grid(row = 0, column = 0, columnspan = 4, pady = 10)

    year_reporte = Entry(ventana_reporte, font = "Helvetica 15")
    year_reporte.grid(row = 1, column = 0, columnspan = 4, pady = 10, padx=2)

    boton_rep = Button(ventana_reporte, text = 'Crear Reporte', command = crear_reg)
    boton_rep.grid(row = 2, column = 0, pady = 10, padx = 10)

    boton_cancel = Button(ventana_reporte, text = '    Cancelar    ', command = ventana_reporte.destroy)
    boton_cancel.grid(row=2, column=1, pady=10, padx=10)

    ventana_reporte.mainloop()

# Funcion para buscar si ya se realizo el certificado

def buscarcertificado():
    # Funciones de los Botones
    def insertar_busqueda():
        ced_id = cajabusqueda.get()
        if len(ced_id) == 10 and ced_id.isdigit():
            for record in tree_busqueda.get_children():
                tree_busqueda.delete(record)
            query = f"SELECT id, Cédula,Tipo,Fecha FROM certificados WHERE Cédula LIKE '{ced_id}'"
            datos_ci = registro.buscar(query)
            for count in range(len(datos_ci)):
                tree_busqueda.insert(parent='', index='end', iid=count, text="", values= datos_ci[count])
                cajabusqueda.delete(0, END)
            else:
                cajabusqueda.delete(0,END)
    
    def generar_again_cert():
        fecha_hora = datetime.datetime.now()
        fecha = str(fecha_hora.date())
        hora = str(fecha_hora.time())
        try:
            seleccion_tree = tree_busqueda.item(tree_busqueda.selection())["values"][0]
            query = f"SELECT * FROM certificados WHERE id = {seleccion_tree}"
            datos_a_rehacer = list(registro.buscar(query)[0])
            for esp in (yanuncay_esp,paraiso_esp,central_esp,centrohist,coordinacion):
                esp_nombre = datos_a_rehacer[8]
                if esp_nombre in esp:
                    datos_a_rehacer[8] = esp
            if (datos_a_rehacer[10] == "Ex. Complexivo"): 
                datos_a_rehacer[7] = "Certificado válido para Movilidad"
            elif (datos_a_rehacer[10] == "Movilidad"):
                datos_a_rehacer[7] = "Certificado válido para Examen Complexivo"
            word.crear_certificado(estudiante=datos_a_rehacer[3], cedula=datos_a_rehacer[4], facultad=datos_a_rehacer[5], carrera=datos_a_rehacer[6], handle=datos_a_rehacer[7], especialista=datos_a_rehacer[8])
            registro.insertar_datos(fecha = fecha , hora = hora ,name=datos_a_rehacer[3] , ci=datos_a_rehacer[4] , facultad=datos_a_rehacer[5], carr_o_maes=datos_a_rehacer[6], hdl=datos_a_rehacer[7], firma=datos_a_rehacer[8] ,tipo_cert=datos_a_rehacer[10], referencista=datos_a_rehacer[-2])
        except:
            print("Seleciona un valor")
 
    # Creacion de la ventana
    ventana_busqueda = Tk()
    ventana_busqueda.geometry("320x360")
    ventana_busqueda.resizable(width = 0, height = 0)
    ventana_busqueda.title('Buscar certificados')
    indicaciones_busqueda = Label(ventana_busqueda, text="Escriba el numero de cedula del estudiante")
    indicaciones_busqueda.grid(row=0, column=0, padx=5, pady=7,columnspan=3)
    cajabusqueda = Entry(ventana_busqueda, font= "Helvetica 20")
    cajabusqueda.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
    boton_buscar_ci = Button(ventana_busqueda, text='  Buscar  ', command=insertar_busqueda)
    boton_buscar_ci.grid(row=3, column=0,padx=0.2, pady=5)
    boton_salir = Button(ventana_busqueda, text="  Cancelar  ", command=ventana_busqueda.destroy)
    boton_salir.grid(row=3,column=2,padx=0.2, pady=5)
    boton_rehacer = Button(ventana_busqueda, text="Generar Cert.",command=generar_again_cert)
    boton_rehacer.grid(row=3,column=1,padx=0.2, pady=5)

    tree_busqueda = ttk.Treeview(ventana_busqueda)
    # Columnas de la vista por arbol
    tree_busqueda['columns'] = ("Id","Cedula","Tipo de Certificado","Fecha")
    tree_busqueda.column("#0", width=0, stretch=NO)
    tree_busqueda.column("Id",width=40, anchor=W)
    tree_busqueda.column("Cedula",width=100, anchor=W)
    tree_busqueda.column("Tipo de Certificado", width=80, anchor=CENTER)
    tree_busqueda.column("Fecha", width=85, anchor=W)

    tree_busqueda.heading("#0", text="", anchor=W)
    tree_busqueda.heading("Id", text="ID",anchor=W)
    tree_busqueda.heading("Cedula",text="Cedula",anchor=CENTER)
    tree_busqueda.heading("Tipo de Certificado", text="Tipo C.", anchor=CENTER)
    tree_busqueda.heading("Fecha", text="Fecha", anchor=CENTER)

    tree_busqueda.grid(row=2, column=0, columnspan=3, pady=5,padx=5)
    ventana_busqueda.mainloop()

def cert_movilidad_complexivo():
    try:
        fecha_hora = datetime.datetime.now()
        fecha = str(fecha_hora.date())
        hora = str(fecha_hora.time())
        nombre = cajanombre_movilidad.get()
        cedula = caja_cedula.get()
        fac = "Facultad de " + lista_facultad.get()

        if grado_m.get() == 0:
            carr = "Carrera de " + lista_carrera.get()
        else:
            carr = 'de la ' + caja_maestria_mov.get()
        ref = referencista_mov.get()
        firma = espe_mov.get()
        mov_o_compl = certif.get()
        if "Movilidad" in mov_o_compl:
            tipo_certificado = "Movilidad"
        else:
            tipo_certificado = "Ex. Complexivo"
        
        word.crear_certificado(estudiante=nombre,cedula=cedula,facultad=fac,carrera=carr,handle=mov_o_compl,especialista=firma, modalidad = virtual.get())
        registro.insertar_datos(fecha = fecha , hora = hora ,name = nombre , ci = cedula , facultad = fac, carr_o_maes = carr, hdl = "", firma = firma, tipo_cert=tipo_certificado,referencista = ref)
        
        cajanombre_movilidad.delete(0, END)
        caja_cedula.delete(0, END)
        caja_maestria_mov.delete(0, END)
        referencista_mov.delete(0, END)
        lista_facultad.delete(0, END)
        lista_carrera.delete(0, END)
    except:
        print("Faltan datos ")

if __name__ == "__main__":
    # Crear Ventana Principal
    ventana = Tk()
    ventana.geometry("350x680")
    ventana.resizable(width=0, height=0)
    ventana.title('Certificado de no Adeudar')

    # Crear las pestanas para cada tipo de certificado
    pestanas = ttk.Notebook(ventana)
    pestanas.pack(fill='both', expand='yes')
    pest_tesis = ttk.Frame(pestanas)
    pest_movil = ttk.Frame(pestanas)
    pest_complex = ttk.Frame(pestanas) 
    pestanas.add(pest_tesis, text="Tesis")
    pestanas.add(pest_movil, text="Movilidad y Complexivo")

#
#   Pestana de Certificados por Tesis
#
    # Imagen o Logotipo
    logo_img = ImageTk.PhotoImage(Image.open('recursos/logo de certificado.png'))
    logo = Label(pest_tesis,image = logo_img)
    logo.grid(row = 0, column = 0, columnspan = 2, pady = 20)

    etiqueta = Label(pest_tesis, text = 'Escribe el handle de la tesis')
    etiqueta.grid(row = 1, column = 0, columnspan = 2)

    cajahandle = Entry(pest_tesis, font = "Helvetica 20")
    cajahandle.grid(row = 2, column = 0, columnspan = 2)
    cajahandle.focus()
    
    etiqueta_datos = Label(pest_tesis, text = "                     ")
    etiqueta_datos.grid(row = 3, column = 0, columnspan = 2)

    pregrado = LabelFrame(pest_tesis, text = 'Tesis de pregrado o posgrado', padx = 5, pady = 10)
    pregrado.grid(row = 4 , column = 0, padx = 13, pady =5, columnspan = 2)

    # Variable de tipo de grado
    grado = IntVar()
    grado.set(0) # Valor Predeterminado para pregrado

    pregrado_but = Radiobutton(pregrado, text = 'Pregrado', variable = grado, value = 0)
    pregrado_but.grid(row = 0, column = 0)

    posgrado_but = Radiobutton(pregrado, text = 'Maestría', variable = grado, value = 1)
    posgrado_but.grid(row = 0, column = 2)

    caja_maestria = Entry(pregrado, font = "Helvetica 20")
    caja_maestria.grid(row = 1, column = 0, columnspan = 6)

    especialista = LabelFrame(pest_tesis, text = 'Especialista de Biblioteca')
    especialista.grid(row =5, padx = 10, pady =10, column = 0, rowspan = 4)
    # Variable para especialistas
    espe = StringVar()
    espe.set(coordinacion) # Valor predeterminado para la coordinacion

    coordinador_but = Radiobutton(especialista, text = 'Coordinador/a', variable = espe, value = coordinacion)
    coordinador_but.grid(row = 0, column =0, sticky="w")

    yanuncay_but = Radiobutton(especialista, text = 'Campus Yanuncay', variable = espe, value = yanuncay_esp)
    yanuncay_but.grid(row = 1, column =0, sticky="w")

    paraiso_but = Radiobutton(especialista, text = 'Campus Paraíso', variable = espe, value = paraiso_esp)
    paraiso_but.grid(row = 2, column =0, sticky="w")

    central_but = Radiobutton(especialista, text = 'Campus Central', variable = espe, value = central_esp)
    central_but.grid(row = 3, column =0, sticky="w")

    centro_historico_but = Radiobutton(especialista, text = 'Campus Centro H.', variable = espe, value = centrohist)
    centro_historico_but.grid(row = 4, column =0, sticky="w")

    boton_generar = Button(pest_tesis, text = '  Crear Certificado  ', command = certificado, padx = 10, pady = 5)
    boton_generar.grid(row = 6, column = 1)

    boton_reporte = Button(pest_tesis, text = 'Reporte', padx = 40, pady = 5, command = butt_reporte)
    boton_reporte.grid(row = 7, column = 1)

    boton_buscar = Button(pest_tesis, text = ' Buscar ', padx = 40, pady = 5, command=buscarcertificado)
    boton_buscar.grid(row = 8, column = 1)

    frame_referencista = LabelFrame(pest_tesis, text = 'Referencista')
    frame_referencista.grid( row = 9, column=0, columnspan=2, padx=10, pady=5)

    referencista = Entry(frame_referencista, font="Helvetica 20")
    referencista.grid(row=0, column=0, columnspan = 2, pady = 3)

    # 
    # Frame para selecionar si es un certificado virtual o impreso
    #
    caract_certificado = LabelFrame(pest_tesis,text='Características del certificado') 
    caract_certificado.grid(row=10, column=0, padx=10,pady=5)

    virtual = StringVar()
    virtual.set("digital") # Valor predeterminado para el certificado virtual

    virtual_but = Radiobutton(caract_certificado, text="Certif. Virtual", variable = virtual, value = "digital")
    virtual_but.grid(row = 0, column = 0, sticky = "w")

    impreso_but = Radiobutton(caract_certificado, text="Certif. Impreso", variable = virtual, value = "impreso")
    impreso_but.grid(row = 1, column = 0, sticky = "w")

    boton_exit = Button(pest_tesis, text = 'Salir', padx = 40, pady = 5, command=ventana.destroy)
    boton_exit.grid(row = 11, column = 0, sticky='w', pady=5, padx=15)

#
#   Certificados por Movilidad
#
    etiqueta_m = Label(pest_movil, text = "  ")
    etiqueta_m.grid(row = 0, column = 0, columnspan = 2, pady=5)

    etiqueta_movilidad = Label(pest_movil, text = 'Escriba los nombres y apellidos del estudiante')
    etiqueta_movilidad.grid(row = 1, column = 0, columnspan = 2, pady=5)

    cajanombre_movilidad = Entry(pest_movil, font = "Helvetica 20")
    cajanombre_movilidad.grid(row = 2, column = 0, columnspan = 2)
    cajanombre_movilidad.focus()
    
    etiqueta_datos_cedula = Label(pest_movil, text = "Escriba la cédula del estudiante")
    etiqueta_datos_cedula.grid(row = 3, column = 0, columnspan = 2)

    caja_cedula = Entry(pest_movil, font="Helvetica 20")
    caja_cedula.grid(row=4, column=0, columnspan=2)
    
    etiqueta_datos_facultad = Label(pest_movil, text = "Seleccione la Facultad")
    etiqueta_datos_facultad.grid(row = 5, column = 0, columnspan = 2, pady=5)

    lista_facultad = ttk.Combobox(pest_movil, values = list(facultades.keys()))
    lista_facultad.configure(width=36, height=30)
    lista_facultad.bind("<<ComboboxSelected>>",seleccionfac)
    lista_facultad.grid(row=6, column=0, columnspan=2, pady=5)
    
    etiqueta_datos_carrera = Label(pest_movil, text = "Seleccione la Carrera")
    etiqueta_datos_carrera.grid(row = 7, column = 0, columnspan = 2, pady=5)

    lista_carrera = ttk.Combobox(pest_movil)
    lista_carrera.configure(width=36, height=30)
    lista_carrera.grid(row=8, column=0, columnspan=2, pady=5)

    pregrado_movilidad = LabelFrame(pest_movil, text = 'Tesis de pregrado o posgrado', padx = 5, pady = 10)
    pregrado_movilidad.grid(row = 9 , column = 0, padx = 13, pady =5, columnspan = 2)
    
    # Variable de tipo de grado
    grado_m = IntVar()
    grado_m.set(0) # Valor Predeterminado para pregrado

    pregrado_but_mov = Radiobutton(pregrado_movilidad, text = 'Pregrado', variable = grado_m, value = 0)
    pregrado_but_mov.grid(row = 0, column = 0)

    posgrado_but_mov = Radiobutton(pregrado_movilidad, text = 'Maestría', variable = grado_m, value = 1)
    posgrado_but_mov.grid(row = 0, column = 2)

    caja_maestria_mov = Entry(pregrado_movilidad, font = "Helvetica 20")
    caja_maestria_mov.grid(row = 1, column = 0, columnspan = 6)
    
    especialista_mov = LabelFrame(pest_movil, text = 'Especialista de Biblioteca')
    especialista_mov.grid(row =10, padx = 10, pady =10, column = 0, rowspan = 4)
    
    # Variable para especialistas
    espe_mov = StringVar()
    espe_mov.set(coordinacion) # Valor predeterminado para la coordinacion

    coordinador_but = Radiobutton(especialista_mov, text = 'Coordinador/a', variable = espe_mov, value = coordinacion)
    coordinador_but.grid(row = 0, column =0, sticky="w")

    yanuncay_but = Radiobutton(especialista_mov, text = 'Campus Yanuncay', variable = espe_mov, value = yanuncay_esp)
    yanuncay_but.grid(row = 1, column =0, sticky="w")

    paraiso_but = Radiobutton(especialista_mov, text = 'Campus Paraíso', variable = espe_mov, value = paraiso_esp)
    paraiso_but.grid(row = 2, column =0, sticky="w")

    central_but = Radiobutton(especialista_mov, text = 'Campus Central', variable = espe_mov, value = central_esp)
    central_but.grid(row = 3, column =0, sticky="w")

    centro_historico_but = Radiobutton(especialista_mov, text = 'Campus Centro H.', variable = espe_mov, value = centrohist)
    centro_historico_but.grid(row = 4, column =0, sticky="w")
    
    movilidad_tipo = "Certificado válido para Movilidad"
    complexico_tipo = "Certificado válido para Examen Complexivo"
 
    certif = StringVar()
    certif.set(movilidad_tipo)
    
    tipocertif = LabelFrame(pest_movil, text="Tipo de Certif.")
    tipocertif.grid(row=10, column=1, pady=5, padx=5)
   
    movilidad_but = Radiobutton(tipocertif, text = 'Movilidad', variable = certif, value = movilidad_tipo)
    movilidad_but.grid(row = 1, column =0, sticky="w")

    complexivo_but = Radiobutton(tipocertif, text = 'Ex. Complex.', variable = certif, value = complexico_tipo)
    complexivo_but.grid(row = 2, column =0, sticky="w")
    
    cert_mov_compl = Button(pest_movil, text="Generar Certif.", command=cert_movilidad_complexivo)
    cert_mov_compl.grid(row=11, column=1)
    
    frame_referencista_mov = LabelFrame(pest_movil, text = 'Referencista')
    frame_referencista_mov.grid( row = 16, column=0, columnspan=2, padx=10)
    
    referencista_mov = Entry(frame_referencista_mov, font="Helvetica 20")
    referencista_mov.grid(row=0, column=0, columnspan = 2, pady = 3)
    
    # 
    # Frame para selecionar si es un certificado virtual o impreso
    #
    
    caract_certificado = LabelFrame(pest_movil,text='Características del certificado') 
    caract_certificado.grid(row=15, column=0, padx=10,pady=5)

    virtual = StringVar()
    virtual.set("digital") # Valor predeterminado para el certificado virtual

    virtual_but = Radiobutton(caract_certificado, text="Certif. Virtual", variable = virtual, value = "digital")
    virtual_but.grid(row = 0, column = 0, sticky = "w")
    
    impreso_but = Radiobutton(caract_certificado, text="Certif. Impreso", variable = virtual, value = "impreso")
    impreso_but.grid(row = 1, column = 0, sticky = "w")
    
    
    boton_exit = Button(pest_tesis, text = 'Salir', padx = 40, pady = 5, command=ventana.destroy)
    boton_exit.grid(row = 11, column = 0, sticky='w', pady=5, padx=15)


    ventana.mainloop()
