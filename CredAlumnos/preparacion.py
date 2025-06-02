import pandas as pd
import re, os, sys


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
    dfAlumnosNuevos = dfAlumnosIntranet[
        ~dfAlumnosIntranet["Clave"].isin(dfTodos["Clave"])
    ]

    dfAlumnosNuevos = dfAlumnosNuevos[
        ~(
            dfAlumnosNuevos["Carrera"].isin(
                ["BACHILLERATO TECNOLOGICO DE LA UNIVERSIDAD IUEM", "PREPARATORIA UAEM"]# Alumnos que no sean de prepa
            )
        )  
        & (
            dfAlumnosNuevos["Plantel"].isin(
                ["IUEM", "ONLINE", "TENANCINGO", "UNIVERSIDAD IUEM"] # Y que solo sean de estos planteles
            )
        )  
    ]

    # Fotos obtenidas de la oficina de credenciales
    ruta = rutaFotos  # "Fotos/RecibidasRecientes/"

    # Conjunto que contendrá todas las fotos dentro de la carpeta mencionada arriba
    fotos_set = set()
    for archivo in os.listdir(ruta):  # Por cada archivo que haya en la carpeta ...
        nombre, ext = os.path.splitext(archivo)  # Obten su nombre y su formato ...
        if ext.lower() in [".jpg",".jpeg",".png",]:  # Si su formato es un formato de imagen ...
            fotos_set.add(nombre)  # Añadelo al conjunto de fotos

    # Alumnos con fotos seran aquellos cuya clave se encuentre dentro del conjunto de fotos
    dfAlumnosConFoto = dfAlumnosNuevos[
        dfAlumnosNuevos["Clave"].astype(str).isin(fotos_set)
    ]

    # Verificar que el alumno tenga apellido paterno
    # Verificar que la longitud del nombre sea <= 45 caracteres
    for _, registro in dfAlumnosConFoto.iterrows():
        if not isinstance(registro["Paterno"], str):
            registro["Paterno"] = registro["Materno"]
            registro["Materno"] = ""
        if (len(f"{registro['Paterno']} {registro['Materno']} {registro['Nombre']}") > 45):
            print(
                f"El alumno {registro['Paterno']} {registro['Materno']} {registro['Nombre']} tiene un nombre demasiado largo"
            )

    # Todo a mayusculas
    dfAlumnosConFoto = dfAlumnosConFoto.apply(
        lambda x: x.map(lambda val: val.upper() if isinstance(val, str) else val)
    )

    # Quitar caracteres especiales
    def quitar_caracteres(txt):
        # re.sub(Caracteres a eliminar, con qué se van a reemplazar, el texto del que se van a eliminar)
        txt = re.sub(r"[^A-ZÁÉÍÓÚÑ ]", "", txt)  # Pero, al contenter ^, indica que esos caracteres en vez de eliminarse, se van a conservar
        txt = (
            txt.replace("Ñ", "N")
            .replace("Á", "A")
            .replace("É", "E")
            .replace("Í", "I")
            .replace("Ó", "O")
            .replace("Ú", "U")
        )
        return txt

    columnas = ["Paterno", "Materno", "Nombre"]
    for columna in columnas:
        dfAlumnosConFoto[columna] = dfAlumnosConFoto[columna].apply(
            lambda x: quitar_caracteres(str(x)) if isinstance(x, str) else x
        )

    # Arreglar el sexo
    dfAlumnosConFoto["Sexo"] = dfAlumnosConFoto["Sexo"].replace({"M": "H", "F": "M"})

    # Verificar fecha de nacimiento igual a RFC
    for i, registro in dfAlumnosConFoto.iterrows():
        fecha_nac = str(registro["Fecha de Nacimiento"])
        rfc = str(registro["RFC"])

        # Si la fecha de nacimiento tiene 8 digitos aaaammdd Y el RFC 10 digitos xxxxAAMMDD ...
        if (len(fecha_nac) == 8) and (len(rfc) == 10):

            # Si fecha de nacimiento a partir de segundo digito en adelante (AAMMDD)
            # E S   D I S T I N T O
            # al RFC en su cuarto dígito en adelante (AAMMDD)
            if fecha_nac[2:] != rfc[4:]:
                # Imprime un error
                print(
                    f"* Error con el alumno {i}: {registro['Nombre']} {registro['Paterno']} {registro['Materno']}"
                    + f"\n\tSu fecha de nacimiento :{fecha_nac} y su RFC: {rfc} no coinciden \n\tPertenece a {registro['Carrera']}"
                )
                # La fecha será arreglada por el RFC
                fecha_nac = "20" + rfc[4:]
                # PUEDE HABER ERRORES DE CAPTURA EN EL RFC DESDE LA INTRANET
                # Por lo que, si la fecha de nacimiento obtenida por el RFC son puros numeros...
                if fecha_nac.isdigit():
                    print(f"LA FECHA CORREGIDA ES {fecha_nac}\n")
                    dfAlumnosConFoto.at[i, "Fecha de Nacimiento"] = int(fecha_nac)  # ... Corrige la fecha
                else:  # Si no son puros numeros, y contiene letras, hay que corregir manualmente los datos del alumno
                    print("EL RFC contiene errores de captura\nVerificar manualmente los datos del alumno\n")
        else:  # Si la fecha de nacimiento NO tiene 8 digitos Y el rfc no tiene 10 digitos ...
            print("Error en fecha de nacimiento o RFC")  # ... imprime un error

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
            return os.path.join(baseRoute, "IEUM MAPPING 2024 10 04 OK - copia.xlsx")

        # Archivo donde se especifican los departamentos segun su nombre COMPLETO, el Mapping original tiene abreviaciones
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

        # Y regresa condicion, campus/plantel, departamento
        return condicion, "04", depto
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
        borrador_pedido.at[i, "CONDICION"], borrador_pedido.at[i, "CAMPUS"],borrador_pedido.at[i, "DEPAR"] = get_condicion(valor["Carrera"])

        borrador_pedido.at[i, "MOVIMIENTO"] = "A"  # Todos los alumnos de "Alta" se harán en automatico
        borrador_pedido.at[i, "DATO ADICIONAL 1"] = ""
        borrador_pedido.at[i, "DATO ADICIONAL 2"] = ""
        borrador_pedido.at[i, "MODIFICACION EN NOMBRE"] = "NO"  # Al ser alta, no requiere cambio de nombre

        if (valor["Nacionalidad"]) != "MEXICANA":  # Si la nacionalidad no es mexicana ...
            print(f"ADVERTENCIA \nEl alumno {valor['Clave']} no es de nacionalidad mexicana, ajustar manualmente")  # Ajustar manualmente

        borrador_pedido.at[i, "Codigo NACIONALIDAD"] = "052"  # Hay que buscar el código de la nacionalidad del alumno en caso de no ser mexicano

        # Todos los demas son datos predeterminados
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
        borrador_pedido.at[i, "Nacionalidad"] = "MEXICO"  # Hay que buscar el pais del alumno en caso de no ser mexicano
        borrador_pedido.at[i, "Pais de residencia"] = "MEXICO"

    return borrador_pedido