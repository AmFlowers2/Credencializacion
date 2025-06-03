import re, os, sys, zipfile, pandas as pd
from PIL import Image

def ProcesarArchivos(AlumnosActivos, Todos, rutaFotos):

    dfAlumnosIntranet = pd.read_excel(
        AlumnosActivos,
        usecols=[
            "Paterno",
            "Materno",
            "Nombre",
            "Clave",
            "Sexo",
            "Fecha de Nacimiento",
            "RFC",
            "Carrera",
            "Nacionalidad",
            "Plantel",
        ],
    )
    dfTodos = pd.read_excel(Todos, usecols=["Clave"])

    # Alumnos nuevos seran aquellos activos cuya clave no esté en la BD de Todos
    dfAlumnosNuevos = dfAlumnosIntranet[~dfAlumnosIntranet["Clave"].isin(dfTodos["Clave"])]

    dfAlumnosNuevos = dfAlumnosNuevos[
        ~(dfAlumnosNuevos["Carrera"].isin(["BACHILLERATO TECNOLOGICO DE LA UNIVERSIDAD IUEM", "PREPARATORIA UAEM"])) 
        &
        (dfAlumnosNuevos["Plantel"].isin(["IUEM", "ONLINE", "TENANCINGO", "UNIVERSIDAD IUEM"]))  
    ]

    # Conjunto que contendrá todas las fotos dentro de la carpeta mencionada arriba
    fotos_set = set()
    for foto in os.listdir(rutaFotos):
        nombre, _ = os.path.splitext(foto)
        fotos_set.add(nombre)  # Añadelo al conjunto de fotos

    # Alumnos con fotos seran aquellos cuya clave se encuentre dentro del conjunto de fotos
    dfAlumnosConFoto = dfAlumnosNuevos[dfAlumnosNuevos["Clave"].astype(str).isin(fotos_set)]

    # Verificar que el alumno tenga apellido paterno
    # Verificar que la longitud del nombre sea <= 45 caracteres
    for _, registro in dfAlumnosConFoto.iterrows():
        if not isinstance(registro["Paterno"], str):
            registro["Paterno"] = registro["Materno"]
            registro["Materno"] = ""
        if (len(f"{registro['Paterno']} {registro['Materno']} {registro['Nombre']}") > 45):
            print(f"El alumno {registro['Paterno']} {registro['Materno']} {registro['Nombre']} tiene un nombre demasiado largo")

    # Todo a mayusculas
    dfAlumnosConFoto = dfAlumnosConFoto.apply(lambda x: x.map(lambda val: val.upper() if isinstance(val, str) else val))

    # Quitar caracteres especiales
    def quitar_caracteres(txt):
        txt = re.sub(r"[^A-ZÁÉÍÓÚÑ ]", "", txt)
        txt = (txt.replace("Ñ", "N").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U"))
        return txt

    for columna in ["Paterno", "Materno", "Nombre"]:
        dfAlumnosConFoto[columna] = dfAlumnosConFoto[columna].apply(lambda x: quitar_caracteres(str(x)) if isinstance(x, str) else x)

    # Arreglar el sexo
    dfAlumnosConFoto["Sexo"] = dfAlumnosConFoto["Sexo"].replace({"M": "H", "F": "M"})

    # Verificar fecha de nacimiento igual a RFC
    for i, registro in dfAlumnosConFoto.iterrows():
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
                    dfAlumnosConFoto.at[i, "Fecha de Nacimiento"] = int(fecha_nac) #Y Asignar la fecha corregida al DataFrame
                else:  # Si no son puros numeros, y contiene letras, hay que corregir manualmente los datos del alumno
                    print("EL RFC contiene errores de captura\nVerificar manualmente los datos del alumno\n")
            
    # Preparar el DataFrame que creará el Excel con el formato solicitado por Santander
    borrador_pedido = pd.DataFrame(
        columns=[
            "APELLIDO P",
            "APELLIDO M",
            "NOMBRE",
            "SEXO",
            "FEC NACIMI",
            "RFC",
            "MATRICULA",
            "CONDICION",
            "CAMPUS",
            "DEPAR",
            "MOVIMIENTO",
            "DATO ADICIONAL 1",
            "DATO ADICIONAL 2",
            "MODIFICACION EN NOMBRE",
            "Codigo NACIONALIDAD",
            "TELEFONO",
            "E-MAIL",
            "NOMBRE DE VIA (CALLE)",
            "NUM DE VIA",
            "INTERIOR",
            "COLONIA",
            "CP",
            "PAIS",
            "POBLACION",
            "ESTADO",
            "COD PROV",
            "DEL/MUN",
            "Nacionalidad",
            "Pais de residencia",
        ]
    )

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

    # Por caaada registro dentro de dfAlumnosConFoto ...
    for i, valor in dfAlumnosConFoto.iterrows():

        # Estos primeros valores son los que arreglamos o que ya teníamos al principio
        borrador_pedido.at[i, "APELLIDO P"] = valor["Paterno"]
        borrador_pedido.at[i, "APELLIDO M"] = valor["Materno"]
        borrador_pedido.at[i, "NOMBRE"] = valor["Nombre"]
        borrador_pedido.at[i, "SEXO"] = valor["Sexo"]
        borrador_pedido.at[i, "FEC NACIMI"] = valor["Fecha de Nacimiento"]
        borrador_pedido.at[i, "RFC"] = valor["RFC"]
        borrador_pedido.at[i, "MATRICULA"] = valor["Clave"]

        # Llamada a la función get_condicion()
        borrador_pedido.at[i, "CONDICION"], borrador_pedido.at[i, "DEPAR"] = get_condicion(valor["Carrera"])

        borrador_pedido.at[i, "CAMPUS"] = "04" # El campus es siempre "04"
        borrador_pedido.at[i, "MOVIMIENTO"] = "A"  # Todos los alumnos de "Alta" se harán en automatico
        borrador_pedido.at[i, "DATO ADICIONAL 1"] = ""
        borrador_pedido.at[i, "DATO ADICIONAL 2"] = ""
        borrador_pedido.at[i, "MODIFICACION EN NOMBRE"] = "NO"  # Al ser alta, no requiere cambio de nombre

        if (valor["Nacionalidad"]) != "MEXICANA":  # Si la nacionalidad no es mexicana ...
            print(f"ADVERTENCIA \nEl alumno {valor['Clave']} no es de nacionalidad mexicana, ajustar manualmente")  # Ajustar manualmente
            borrador_pedido.at[i, "Codigo NACIONALIDAD"] = ""  # Dejar en blanco el código de nacionalidad
            borrador_pedido.at[i, "Nacionalidad"] = ""  # Dejar en blanco la nacionalidad
            borrador_pedido.at[i, "PAIS"] = ""  # Dejar en blanco el país
        else:
            borrador_pedido.at[i, "Codigo NACIONALIDAD"] = "052"
            borrador_pedido.at[i, "Nacionalidad"] = "MEXICO"
            borrador_pedido.at[i, "PAIS"] = "052"

        # Todos los demas son datos predeterminados
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

# Genera un zip con las fotos redimensionadas y las guarda en la ruta especificada
def genZip(rutaFotos, fecha):

    rutaRaiz = os.path.dirname(rutaFotos)
    zipName = os.path.join(rutaRaiz,f"Pedido A {fecha}.zip")

    for foto in os.listdir(rutaFotos):
        if foto.lower().endswith(".jpg"):
            rutaImg = os.path.join(rutaFotos, foto)
            img_redimensionada = Image.open(rutaImg).resize((182, 230))
            img_redimensionada.save(rutaImg, "JPEG") 

    with zipfile.ZipFile(zipName, 'w', compression= zipfile.ZIP_DEFLATED) as zipf:
        for foto in os.listdir(rutaFotos):
            rutaFoto = os.path.join(rutaFotos, foto)
            zipf.write(rutaFoto, foto)