import re, os, sys, zipfile, pandas as pd, pdfplumber, numpy as np
from PIL import Image

fotos_set = set()
        
def limpiarDatos(dfAlumnos):
    
    # Verificar que el alumno tenga apellido paterno
    # Verificar que la longitud del nombre sea <= 45 caracteres
    for _, registro in dfAlumnos.iterrows():
        if not isinstance(registro["Paterno"], str):
            registro["Paterno"] = registro["Materno"]
            registro["Materno"] = ""
        if (len(f"{registro['Paterno']} {registro['Materno']} {registro['Nombre']}") > 45):
            print(f"El alumno {registro['Paterno']} {registro['Materno']} {registro['Nombre']} tiene un nombre demasiado largo")

    # Todo a mayusculas
    dfAlumnos = dfAlumnos.apply(lambda x: x.map(lambda val: val.upper() if isinstance(val, str) else val))

    # Quitar caracteres especiales
    def quitar_caracteres(txt):
        txt = re.sub(r"[^A-ZÁÉÍÓÚÑ ]", "", txt)
        txt = (txt.replace("Ñ", "N").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U"))
        return txt

    for columna in ["Paterno", "Materno", "Nombre"]:
        dfAlumnos[columna] = dfAlumnos[columna].apply(lambda x: quitar_caracteres(str(x)) if isinstance(x, str) else x)

    # Arreglar el sexo
    dfAlumnos["Sexo"] = dfAlumnos["Sexo"].replace({"M": "H", "F": "M"})

    # Verificar fecha de nacimiento igual a RFC
    for i, registro in dfAlumnos.iterrows():
        fecha_nac = str(registro["Fecha de Nacimiento"])
        rfc = str(registro["RFC"])

        # Si la fecha de nacimiento NO tiene 8 digitos AAAAMMDD o el RFC NO tiene 10 digitos xxxxAAMMDD ...
        if (len(fecha_nac) != 8) or (len(rfc) != 10):
            print("Error en fecha de nacimiento o RFC")  # ... imprime un error
        else: #Si no...
            if fecha_nac[2:] != rfc[4:]: #Si la fecha de nacimiento (AAAAMMDD) no coincide con el RFC (xxxxAAMMDD) ...
                print(f"* Error con el alumno {i}: {registro['Nombre']} {registro['Paterno']} {registro['Materno']}" # Imprime un error
                    + f"\n\tSu fecha de nacimiento :{fecha_nac} y su RFC: {rfc} no coinciden \n\tPertenece a {registro['Carrera']}")
                
                #El RFC tiene prioridad en la fecha de nacimiento
                if int(rfc[4:6]) <= 25: # Si el mes del RFC es menor o igual a 25, entonces el alumno nació en el siglo XXI
                    fecha_nac = "20" + rfc[4:]
                else:
                    fecha_nac = "19" + rfc[4:] #Si no, entonces el alumno nació en el siglo XX

                if fecha_nac.isdigit(): #Si la fecha de nacimiento contiene solo numeros se puede corregir
                    print(f"LA FECHA CORREGIDA ES {fecha_nac}\n")
                    dfAlumnos.at[i, "Fecha de Nacimiento"] = int(fecha_nac) #Y Asignar la fecha corregida al DataFrame
                else:  # Si no son puros numeros, y contiene letras, hay que corregir manualmente los datos del alumno
                    print("EL RFC contiene errores de captura\nVerificar manualmente los datos del alumno\n")

    return dfAlumnos

def procesarArchivos(rutaAlumnosActivos, rutaTodos, rutaFotos, rutaPdf = None):

    dfAlumnosIntranet = pd.read_excel(rutaAlumnosActivos, usecols=["Paterno","Materno","Nombre","Clave","Sexo","Fecha de Nacimiento","RFC","Carrera","Nacionalidad","Plantel"],dtype=str)
    dfTodos = pd.read_excel(rutaTodos, usecols=["Clave"])

    # Alumnos nuevos seran aquellos activos cuya clave no esté en la BD de Todos
    dfAlumnosNuevos = dfAlumnosIntranet[~dfAlumnosIntranet["Clave"].isin(dfTodos["Clave"])]

    dfAlumnosNuevos = dfAlumnosNuevos[
        ~(dfAlumnosNuevos["Carrera"].isin(["BACHILLERATO TECNOLOGICO DE LA UNIVERSIDAD IUEM", "PREPARATORIA UAEM", "PREPARATORIA SE"])) 
        &
        (dfAlumnosNuevos["Plantel"].isin(["IUEM", "ONLINE", "TENANCINGO", "UNIVERSIDAD IUEM"]))  
    ]

    # Conjunto que contendrá todas las fotos dentro de la carpeta mencionada arriba
    #fotos_set = set()
    global fotos_set
    for foto in os.listdir(rutaFotos):
        nombre, _ = os.path.splitext(foto)
        fotos_set.add(nombre)  # Añadelo al conjunto de fotos

    # Alumnos con fotos seran aquellos cuya clave se encuentre dentro del conjunto de fotos
    dfAlumnosConFoto = dfAlumnosNuevos[dfAlumnosNuevos["Clave"].astype(str).isin(fotos_set)]
    
    #Si se proporcionó un PDF con las claves de los alumnos que pagaron reposición ..
    if rutaPdf is not None:
        clavesRepo = []
        with pdfplumber.open(rutaPdf) as pdf:
            for page in pdf.pages:
                text = page.extract_table()
                for i,row in enumerate(text):
                    if i > 0:
                        clavesRepo.append(row[3])
        dfAlumnosRepo = dfAlumnosNuevos[dfAlumnosNuevos['Clave'].isin(clavesRepo)]
        
        dfAlumnosRepo = limpiarDatos(dfAlumnosRepo)
        dfAlumnosConFoto = limpiarDatos(dfAlumnosConFoto)
        
        dfAlumnosRepo = crearBorrador(dfAlumnosRepo, "R")
        dfAlumnosConFoto = crearBorrador(dfAlumnosConFoto, "A")
        
        borrador_pedido = pd.concat([dfAlumnosConFoto, dfAlumnosRepo], ignore_index=True)
    else:
        dfAlumnosConFoto = limpiarDatos(dfAlumnosConFoto)
        dfAlumnosConFoto = crearBorrador(dfAlumnosConFoto, "A")
        borrador_pedido = dfAlumnosConFoto
    
    return borrador_pedido

def crearBorrador(dfAlumnos, movimiento):            
    
    # Obtener la condición, plantel y departamento del alumno, recibe como parámetro el nombre de la carrera del alumno
    def get_condicion(carrera):
        # Primero, quita todos los caracteres especiales y carreras con "A" p.e. "Arqitectura A", para que solo sea "Arquitectura"
        carrera = re.sub(r"\sA$", "", carrera).strip()
        carrera = (
            carrera.replace("Ñ", "N")
            .replace("Á", "A")
            .replace("É", "E")
            .replace("Í", "I")
            .replace("Ó", "O")
            .replace("Ú", "U")
        )

        def getRoute():
            if getattr(sys, "frozen", False):
                baseRoute = sys._MEIPASS
            else:
                baseRoute = os.path.dirname(__file__)
            #Archivo donde se especifican los departamentos segun su nombre COMPLETO, el Mapping original tiene abreviaciones
            return os.path.join(baseRoute, "IEUM MAPPING 2024 10 04 OK - copia.xlsx")
        
        # Usará la hoja llamada "Departamento"
        mapping = pd.read_excel(getRoute(), sheet_name="Departamento", dtype={"Depto": str})

        # Listas con los códigos de cada departamento
        deptos_prepa = ["005", "053", "054"]
        deptos_maestria = [
            "021",
            "022",
            "023",
            "024",
            "025",
            "026",
            "027",
            "036",
            "037",
            "038",
            "039",
            "040",
            "041",
            "042",
            "043",
            "044",
            "045",
            "046",
            "047",
            "048",
        ]
        deptos_doctorado = ["017", "018"]

        # Aqui se asignará el departamento correspondiente al alumno
        depto = ""

        for i, registro in enumerate(mapping["Descripción"]):  # Por cada registro en la columna "descripción" del archivo mapping ...
            if (carrera == registro):  # Si el nombre de la carrera coincide con el nombre del registro ...
                depto = mapping.at[i, "Depto"]  # Asigna el valor del departamento
                break  # Y rompe el bucle, no hace falta seguir buscando

        if (depto in deptos_prepa):  # Si el valor de depto se encuentra dentro de deptos_prepa...
            condicion = "01"  # El alumno es de preparatoria
        elif (depto in deptos_maestria):  # Si no, si se encuentra dentro de deptos_maestria ...
            condicion = "03"  # El alumno es de maestria/posgrado
        elif (depto in deptos_doctorado):  # Si no, si se encuentra dentro de deptos_doctorado
            condicion = "04"  # El alumno es de doctorado
        else:  # En cualquier otro caso ...
            condicion = "02"  # El alumno es de licenciatura

        #Y regresa condicion, campus, departamento
        return condicion, depto
        # Para poder asignarlos a cada registro

    COLUMNAS_BORRADOR = [
    "APELLIDO P", "APELLIDO M", "NOMBRE", "SEXO", "FEC NACIMI", "RFC", "MATRICULA",
    "CONDICION", "CAMPUS", "DEPAR", "MOVIMIENTO", "DATO ADICIONAL 1", "DATO ADICIONAL 2",
    "MODIFICACION EN NOMBRE", "Codigo NACIONALIDAD", "TELEFONO", "E-MAIL",
    "NOMBRE DE VIA (CALLE)", "NUM DE VIA", "INTERIOR", "COLONIA", "CP", "PAIS",
    "POBLACION", "ESTADO", "COD PROV", "DEL/MUN", "Nacionalidad", "Pais de residencia"
    ]
    
    dfTemp = pd.DataFrame(index=dfAlumnos.copy().index, columns=COLUMNAS_BORRADOR)
    
    dfTemp["APELLIDO P"] = dfAlumnos["Paterno"]
    dfTemp["APELLIDO M"] = dfAlumnos["Materno"]
    dfTemp["NOMBRE"] = dfAlumnos["Nombre"]
    dfTemp["SEXO"] = dfAlumnos["Sexo"]
    dfTemp["FEC NACIMI"] = dfAlumnos["Fecha de Nacimiento"]
    dfTemp["RFC"] = dfAlumnos["RFC"]
    dfTemp["MATRICULA"] = dfAlumnos["Clave"]
    
    dfTemp["MOVIMIENTO"] = movimiento
    
    dfTemp[["CONDICION", "DEPAR"]] = dfAlumnos["Carrera"].apply(lambda x: pd.Series(get_condicion(x)))

    esmexicano = dfAlumnos["Nacionalidad"] == "MEXICANA"
    dfTemp["Codigo NACIONALIDAD"] = np.where(esmexicano, "052", "")
    dfTemp["Nacionalidad"] = np.where(esmexicano, "052", "")
    dfTemp["PAIS"] = np.where(esmexicano, "052", "")
    
    dfTemp["CAMPUS"] = "04"
    dfTemp["MODIFICACION EN NOMBRE"] = "NO"
    dfTemp["TELEFONO"] = "7222624817"
    dfTemp["E-MAIL"] = "telecom@universidadiuem.edu.mx"
    dfTemp["NOMBRE DE VIA (CALLE)"] = "BOULEVARD TOLUCA METEPEC NORTE"
    dfTemp["NUM DE VIA"] = "814"
    dfTemp["COLONIA"] = "HIPICO"
    dfTemp["CP"] = "52156"
    dfTemp["POBLACION"] = "METEPEC"
    dfTemp["ESTADO"] = "0000008MC"
    dfTemp["COD PROV"] = "00054"
    dfTemp["DEL/MUN"] = "METEPEC"
    dfTemp["Pais de residencia"] = "MEXICO"
    
    return dfTemp

# Genera un zip con las fotos redimensionadas EN 182x230px y las guarda en la misma ruta que la carpeta de fotos
def genZip(rutaExcel, rutaFotos, fecha, borrador_pedido):

    zipName = os.path.join(os.path.dirname(rutaExcel),f"Pedido A {fecha}.zip")
    fotosValidas = fotos_set & set(borrador_pedido["MATRICULA"].astype(str))

    with zipfile.ZipFile(zipName, 'w', compression= zipfile.ZIP_DEFLATED) as zipf:
        for foto in os.listdir(rutaFotos):
            if foto.lower().endswith(".jpg"):
                rutaImg = os.path.join(rutaFotos, foto)
                Image.open(rutaImg).resize((182, 230)).save(rutaImg, "JPEG")
                if os.path.splitext(foto)[0] in fotosValidas:
                    zipf.write(rutaImg, foto)
                else:
                    print(f"La foto {foto} no se incluyó en el zip.")
                    