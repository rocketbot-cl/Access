# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Función que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"
    
    pip install <package> -t .

"""
import os
import sys

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'Access' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)
import pyodbc

"""
    Obtengo el modulo que fueron invocados
"""
module = GetParams("module")

"""
    Automate Android devices
"""
global Access

if module == "connect":
    path = GetParams('path')

    try:
        extension = path.split(".")[-1]
        drivers = [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

        if 'Microsoft Access Driver (*.mdb, *.accdb)' in drivers:

            drivers = drivers[0].replace("(*.mdb)", "(*.mdb, *.accdb)")
        else:
            drivers = "{" + drivers[0] + "}"

        conn_str = (
            f'Driver={drivers};'
            f'DBQ={path};'
        )


        conn = pyodbc.connect(conn_str)

        Access = {
            "conn": conn,
            "path": path
        }

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "execute_query":
    try:
        query = GetParams('query')
        result = GetParams('var_')

        cursor = Access["conn"].cursor()
        cursor.execute(query)


        fetch = cursor.fetchall()
        print(result, fetch)
        if result:
            SetVar(result, fetch)

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if (module == "closeConnection"):
    try:
        Access["conn"].close()

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e


