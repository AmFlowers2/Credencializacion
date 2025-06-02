import pandas as pd 
import re, os

def procesarDatosDocentes(ProfesoresNuevos, Todos, rutaFotos):
    #Verificar qué docentes activos no se encuentran en la BD de Todas las credenciales para poder pedir su credencial
    dfDocentesIntranet = pd.read_excel(ProfesoresNuevos, 
                                    usecols=["appaterno", "apmaterno", "nombre", "clave", "Sexo", "fechanacimiento", "rfc", "Nacionalidad", "Plantel"],
                                    dtype=str)
    dfTodos = pd.read_excel(Todos, usecols=["Clave"])

    #Docentes nuevos seran aquellos activos cuya clave no esté en la BD de Todos
    dfDocentesNuevos = dfDocentesIntranet[~dfDocentesIntranet["clave"].isin(dfTodos["Clave"])]

    ruta = rutaFotos #Fotos obtenidas de la oficina de credenciales
    fotos_set = set() #Conjunto que contendrá todas las fotos dentro de la carpeta mencionada arriba
    for archivo in os.listdir(ruta): #Por cada archivo que haya en la carpeta ...
        nombre, ext = os.path.splitext(archivo) #Obten su nombre y su formato ...
        if ext.lower() in [".jpg", ".jpeg", ".png"]: #Si su formato es un formato de imagen ...
            if nombre.startswith("C"):                
                fotos_set.add(nombre) #Añadelo al conjunto de fotos
            else:
                fotos_set.add("C"+nombre) #Añadelo al conjunto de fotos

    #Docentes con fotos seran aquellos cuya clave se encuentre dentro del conjunto de fotos
    dfDocentesConFoto = dfDocentesNuevos[dfDocentesNuevos["clave"].astype(str).isin(fotos_set)]
    #Usaremos dfDocentesConFoto para hacer el pedido

    #Verificar que el docente tenga apellido paterno
    for i, registro in dfDocentesConFoto.iterrows():
        if not isinstance(registro['appaterno'], str):
            registro['appaterno'] = registro['apmaterno']
            registro['apmaterno'] = ""
        if len(f"{registro['appaterno']}{registro['apmaterno']}{registro['nombre']}") > 45:
            print(f"El docente {registro['appaterno']} {registro['apmaterno']} {registro['nombre']} tiene un nombre demasiado largo")

    #Todo a mayusculas
    dfDocentesConFoto = dfDocentesConFoto.apply(lambda x: x.map(lambda val: val.upper() if isinstance(val, str) else val))

    #Quitar caracteres especiales
    def quitar_caracteres(txt):
        txt = re.sub(r'[^A-ZÁÉÍÓÚÑ ]', '', txt)
        txt = txt.replace('Ñ', 'N').replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        return txt

    columnas = ["appaterno", "apmaterno", "nombre"]
    for columna in columnas:
        dfDocentesConFoto[columna] = dfDocentesConFoto[columna].apply(lambda x: quitar_caracteres(str(x)) if isinstance(x, str) else x)

    #Arreglar el sexo
    dfDocentesConFoto["Sexo"] = dfDocentesConFoto["Sexo"].replace({"M" : "H", "F": "M"})

    #Verificar fecha de nacimiento igual a RFC
    for i, registro in dfDocentesConFoto.iterrows():
        fecha_nac = str(registro["fechanacimiento"]).replace("-","")[:8]
        rfc = str(registro["rfc"])

        #Si la fecha de nacimiento tiene 8 digitos AAAAMMDD y el RFC 10 digitos xxxxAAMMDD
        if ((len(fecha_nac) == 8) and (len(rfc) == 13)):
            #Si fecha de nacimiento a partir de segundo digito en adelante (AAMMDD)
            #E S   D I S T I N T O 
            #al RFC en su cuarto dígito en adelante (AAMMDD)
            if (fecha_nac[2:] != rfc[4:10]):
                #Imprime un error
                print(f"* Error con el docente {i}: {registro['nombre']} {registro['appaterno']} {registro['apmaterno']}"
                +f"\n\tSu fecha de nacimiento :{fecha_nac} y su RFC: {rfc} no coinciden \n")
                    
                #La fecha será mandada por el RFC y dependerá de los últimos 2 digitos de la fecha de nacimiento
                if int(rfc[4:6]) <= 25: #Si los ultimos 2 digitos del año de nacimiento son menores o iguales a 25 ...
                    fecha_nac = "20" + rfc[4:10] #El docente nació en los 2000's
                else: #Si no ...
                    fecha_nac = "19" + rfc[4:10] #El docente nación en los 1900's

                #Si dicha fecha se conforma de puros numeros...
                if fecha_nac.isdigit():
                    print(f"LA FECHA CORREGIDA ES {fecha_nac}\n")
                    #Corrige la fecha
                    dfDocentesConFoto.at[i, "fechanacimiento"] = int(fecha_nac)
                else: #Si no, corrigela manunalmente
                    print("Verificar manualmente los datos del docente\n")                
        else:
            print("Error en fecha de nacimiento o RFC")

    borrador_pedido = pd.DataFrame(columns=
            ["APELLIDO P", "APELLIDO M", "NOMBRE", "SEXO",
            "FEC NACIMI", "RFC", "MATRICULA", "CONDICION",
            "CAMPUS", "PROGRAMA", "MOVIMIENTO", "DATO ADICIONAL 1",
            "DATO ADICIONAL 2","MODIFICACION EN NOMBRE",
            "Codigo NACIONALIDAD", "TELEFONO", "E-MAIL",
            "NOMBRE DE VIA (CALLE)", "NUM DE VIA","INTERIOR",
            "COLONIA", "CP", "PAIS", "POBLACION", "ESTADO",
            "COD PROV", "DEL/MUN","Nacionalidad","Pais de residencia"])
        
        #Por caaada registro dentro de dfDocentesPedido ...
    for i, valor in dfDocentesConFoto.iterrows():
        borrador_pedido.at[i, "APELLIDO P"] = valor["appaterno"]
        borrador_pedido.at[i, "APELLIDO M"] = valor["apmaterno"]
        borrador_pedido.at[i, "NOMBRE"] = valor["nombre"]
        borrador_pedido.at[i, "SEXO"] = valor["Sexo"]
        borrador_pedido.at[i, "FEC NACIMI"] = int(str(valor["fechanacimiento"]).replace("-","")[:8])
        borrador_pedido.at[i, "RFC"] = valor["rfc"][:10]
        borrador_pedido.at[i, "MATRICULA"] = valor["clave"]
        borrador_pedido.at[i, "CONDICION"] = "08" #Docente
        borrador_pedido.at[i, "CAMPUS"] = "04" #Campus vacío
        borrador_pedido.at[i, "PROGRAMA"] = "403" #Programa Vacio
        borrador_pedido.at[i, "MOVIMIENTO"] = "A" #Todos los Docentes de "Alta" se harán en automatico
        borrador_pedido.at[i, "DATO ADICIONAL 1"] = ""
        borrador_pedido.at[i, "DATO ADICIONAL 2"] = ""
        borrador_pedido.at[i, "MODIFICACION EN NOMBRE"] = "NO" #Al ser alta, no requiere cambio de nombre

        if (valor["Nacionalidad"]) != "MEXICANA": #Si la nacionalidad no es mexicana ...
            clave = valor["Clave"]
            print(f"ADVERTENCIA \nEl docente {clave} no es de nacionalidad mexicana, ajustar manualmente") #Ajustar manualmente

        borrador_pedido.at[i, "Codigo NACIONALIDAD"] = "052"  #Hay que buscar el código de la nacionalidad del docente en caso de no ser mexicano

        #Todos los demas son datos predeterminados
        borrador_pedido.at[i, "TELEFONO"] = "7222624817"
        borrador_pedido.at[i, "E-MAIL"] = "telecom@universidadiuem.edu.mx"
        borrador_pedido.at[i, "NOMBRE DE VIA (CALLE)"] = "BOULEVARD TOLUCA METEPEC NORTE"
        borrador_pedido.at[i, "NUM DE VIA"] = "814"
        borrador_pedido.at[i, "INTERIOR"] = ""
        borrador_pedido.at[i, "COLONIA"] = "HIPICO"
        borrador_pedido.at[i, "CP"] = "52156"
        borrador_pedido.at[i, "PAIS"] = "052"
        borrador_pedido.at[i, "POBLACION"] = "METEPEC"
        borrador_pedido.at[i, "ESTADO"] = "0000008MC"
        borrador_pedido.at[i, "COD PROV"] = "00054"
        borrador_pedido.at[i, "DEL/MUN"] = "METEPEC"
        borrador_pedido.at[i, "Nacionalidad"] = "MEXICANA" #Hay que buscar el pais del alumno en caso de no ser mexicano
        borrador_pedido.at[i, "Pais de residencia"] = "MEXICO"

    return borrador_pedido

def renombrarFotos(rutaFotos):
    for nombre in os.listdir(rutaFotos):
        origen = os.path.join(rutaFotos, nombre)
        if os.path.isfile(origen) and not nombre.startswith("C"):
            nuevo_nombre = "C"+nombre
            destino = os.path.join(rutaFotos, nuevo_nombre)
            os.rename(origen, destino)