import pandas as pd, re, os, zipfile
from PIL import Image

fotos_set = set()

def quitar_caracteres(txt): #Quitar caracteres especiales de los nombres
        txt = re.sub(r'[^A-ZÁÉÍÓÚÑ ]', '', txt)
        txt = txt.replace('Ñ', 'N').replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        return txt

def procesarDatosDocentes(ProfesoresNuevos, Todos, rutaFotos):
    #Verificar qué docentes activos no se encuentran en la BD de Todas las credenciales para poder pedir su credencial
    dfDocentesIntranet = pd.read_excel(ProfesoresNuevos, 
                                    usecols=["appaterno", "apmaterno", "nombre", "clave", "Sexo", "fechanacimiento", "rfc", "Nacionalidad", "Plantel"],
                                    dtype=str)
    
    dfTodos = pd.read_excel(Todos, usecols=["Clave"])

    #Docentes nuevos seran aquellos activos cuya clave no esté en la BD de Todos
    dfDocentesNuevos = dfDocentesIntranet[~dfDocentesIntranet["clave"].isin(dfTodos["Clave"])]

    #fotos_set = set() #Conjunto que contendrá todas las fotos dentro de la carpeta seleccionada
    global fotos_set
    for foto in os.listdir(rutaFotos): #Por cada foto dentro de la carpeta
        nombre, _ = os.path.splitext(foto)
        if nombre.startswith("C"): #Si el nombre ya empieza con C, no lo cambiamos               
            fotos_set.add(nombre)
        else:
            fotos_set.add("C"+nombre) #Si no, le agregamos una C al inicio del nombre

    #Docentes con fotos seran aquellos cuya clave se encuentre dentro del conjunto de fotos
    dfDocentesConFoto = dfDocentesNuevos[dfDocentesNuevos["clave"].astype(str).isin(fotos_set)]

    #Verificar que el docente tenga apellido paterno
    for _, registro in dfDocentesConFoto.iterrows():
        if not isinstance(registro['appaterno'], str):
            registro['appaterno'] = registro['apmaterno']
            registro['apmaterno'] = ""
        if len(f"{registro['appaterno']}{registro['apmaterno']}{registro['nombre']}") > 45:
            print(f"El docente {registro['appaterno']} {registro['apmaterno']} {registro['nombre']} tiene un nombre demasiado largo")

    #Todo a mayusculas
    dfDocentesConFoto = dfDocentesConFoto.apply(lambda x: x.map(lambda val: val.upper() if isinstance(val, str) else val))
   
    #Quitar caracteres especiales de los nombres
    for columna in ["appaterno", "apmaterno", "nombre"]:
        dfDocentesConFoto[columna] = dfDocentesConFoto[columna].apply(lambda x: quitar_caracteres(str(x)) if isinstance(x, str) else x)

    #Cambiar el sexo, Masculino a Hombre y Femenino a Mujer
    dfDocentesConFoto["Sexo"] = dfDocentesConFoto["Sexo"].replace({"M" : "H", "F": "M"})

    #Verificar fecha de nacimiento igual a RFC
    for i, registro in dfDocentesConFoto.iterrows():
        fecha_nac = str(registro["fechanacimiento"]).replace("-","")[:8]
        rfc = str(registro["rfc"])

        #La fecha de nacimiento debe tener 8 digitos (AAAAMMDD) y el RFC 13 (xxxxAAMMDDxxx)
        if ((len(fecha_nac) != 8) or (len(rfc) != 13)): #Si alguno de los dos no tiene la longitud correcta...
            print("Error en fecha de nacimiento o RFC")               
        else:
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

                #La fecha deben contener unicamente números
                if fecha_nac.isdigit():
                    print(f"LA FECHA CORREGIDA ES {fecha_nac}\n")
                    dfDocentesConFoto.at[i, "fechanacimiento"] = int(fecha_nac) #Para poder corregir la fecha de nacimiento
                else: #Si no, corrigela manunalmente
                    print("Verificar manualmente los datos del docente\n") 

    borrador_pedido = pd.DataFrame(columns=
            ["APELLIDO P", "APELLIDO M", "NOMBRE", "SEXO",
            "FEC NACIMI", "RFC", "MATRICULA", "CONDICION",
            "CAMPUS", "PROGRAMA", "MOVIMIENTO", "DATO ADICIONAL 1",
            "DATO ADICIONAL 2","MODIFICACION EN NOMBRE",
            "Codigo NACIONALIDAD", "TELEFONO", "E-MAIL",
            "NOMBRE DE VIA (CALLE)", "NUM DE VIA","INTERIOR",
            "COLONIA", "CP", "PAIS", "POBLACION", "ESTADO",
            "COD PROV", "DEL/MUN","Nacionalidad","Pais de residencia"])
        
    #Por caaada registro dentro de dfDocentesConFoto...
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
            print(f"ADVERTENCIA \nEl docente {valor['Clave']} no es de nacionalidad mexicana, ajustar manualmente") 
            #Ajustar manualmente los datos de nacionalidad, en el Excel se deja en blanco
            borrador_pedido.at[i, "Codigo NACIONALIDAD"] = ""
            borrador_pedido.at[i, "Nacionalidad"] = ""
            borrador_pedido.at[i, "PAIS"] = ""
        else:
            borrador_pedido.at[i, "Codigo NACIONALIDAD"] = "052"
            borrador_pedido.at[i, "Nacionalidad"] = "MEXICANA"
            borrador_pedido.at[i, "PAIS"] = "052"

        #Todos los demas son datos predeterminados
        borrador_pedido.at[i, "TELEFONO"] = "7222624817"
        borrador_pedido.at[i, "E-MAIL"] = "telecom@universidadiuem.edu.mx"
        borrador_pedido.at[i, "NOMBRE DE VIA (CALLE)"] = "BOULEVARD TOLUCA METEPEC NORTE"
        borrador_pedido.at[i, "NUM DE VIA"] = "814"
        borrador_pedido.at[i, "INTERIOR"] = ""
        borrador_pedido.at[i, "COLONIA"] = "HIPICO"
        borrador_pedido.at[i, "CP"] = "52156"
        borrador_pedido.at[i, "POBLACION"] = "METEPEC"
        borrador_pedido.at[i, "ESTADO"] = "0000008MC"
        borrador_pedido.at[i, "COD PROV"] = "00054"
        borrador_pedido.at[i, "DEL/MUN"] = "METEPEC"     
        borrador_pedido.at[i, "Pais de residencia"] = "MEXICO"

    return borrador_pedido

def genZip(rutaFotos, fecha, borrador_pedido): #C/Users/Abraham/Desktop/Docentes , 2025 06 04

    rutaRaiz = os.path.dirname(rutaFotos) #C/Users/Abraham/Desktop
    zipName = os.path.join(rutaRaiz, f"Pedido DOC {fecha}.zip") #C/Users/Abraham/Desktop/Pedido DOC 2025 06 04.zip
    fotosValidas = fotos_set & set(borrador_pedido["MATRICULA"].astype(str))


    with zipfile.ZipFile(zipName, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
        for foto in os.listdir(rutaFotos):
            if foto.lower().endswith(".jpg"):
                nombreOriginal = os.path.join(rutaFotos, foto)
                nombreSinExt, _ = os.path.splitext(foto)
                if not foto.startswith("C"):
                    nombreNuevo = os.path.join(rutaFotos, "C"+foto)
                    os.rename(nombreOriginal, nombreNuevo)
                    imgRedimensionada = Image.open(nombreNuevo).resize((182,230))
                    imgRedimensionada.save(nombreNuevo, "JPEG")
                    if "C"+nombreSinExt in fotosValidas:
                        zipf.write(nombreNuevo, "C"+foto) 
                    else:
                        print(f"La foto {foto} no se incluyó en el zip.")
                else:
                    imgRedimensionada = Image.open(nombreOriginal).resize((182,230))
                    imgRedimensionada.save(nombreOriginal, "JPEG")
                    if nombreSinExt in fotosValidas:
                        zipf.write(nombreOriginal, foto)
                    else:
                        print(f"La foto {foto} no se incluyó en el zip.")