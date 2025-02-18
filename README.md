# Automatización para obtener y guardar la Tasa de Cambio del BCV en Excel (con Python y Programador de tareas de Windows)

Guía paso a paso para crear un script en Python que obtiene la tasa de cambio oficial del Banco Central de Venezuela (BCV) desde una API y la guarda automáticamente en un archivo de Excel local en tu computadora, utilizando el Programador de tareas de Windows para la automatización.

1. Introducción

Esta guía te mostrará cómo automatizar la obtención de la tasa de cambio oficial del Banco Central de Venezuela (BCV) y guardarla en un archivo de Excel en tu computadora.  Utilizaremos un script en Python para obtener los datos de una API pública, y el Programador de tareas de Windows para ejecutar el script automáticamente de forma periódica.

Esta guía es ideal para:

*   Personas que necesitan **monitorizar y registrar la tasa del BCV de forma regular**.
*   Usuarios que quieren **aprender a automatizar tareas básicas con Python y el Programador de tareas de Windows**, incluso sin tener experiencia previa en programación.
*   Aquellos que buscan una **solución práctica, gratuita y sencilla** para obtener datos web y guardarlos en formato Excel.

**Objetivo:** 
Al finalizar esta guía, tendrás un script de Python que se ejecuta automáticamente en tu computadora, obtiene la tasa del BCV, y la guarda en un archivo Excel, ¡sin que tengas que hacer nada manualmente!

2. Prerrequisitos

Antes de empezar, asegúrate de tener instalado lo siguiente en tu computadora:

*   **Python 3.x:** Si no lo tienes instalado, puedes descargarlo desde la [página oficial de Python](https://www.python.org/downloads/).  Durante la instalación, asegúrate de marcar la opción "Add Python to PATH" para poder ejecutar Python desde la línea de comandos.

*   **Biblioteca `openpyxl` de Python:**  Esta biblioteca es necesaria para trabajar con archivos de Excel desde Python.  Para instalarla, abre la línea de comandos (Símbolo del sistema o PowerShell en Windows) y ejecuta el siguiente comando:

    pip install openpyxl

*   **Acceso a Internet:** Necesitas conexión a internet para descargar el script de GitHub, instalar la biblioteca `openpyxl` y para que el script pueda acceder a la API del BCV.

*   **Editor de texto (Opcional):** Aunque puedes usar el Bloc de notas de Windows, un editor de texto más avanzado como [Notepad++](https://notepad-plus-plus.org/) o [Visual Studio Code](https://code.visualstudio.com/) (gratuito) puede facilitar la edición del código Python y el archivo `.bat`.

*   **Microsoft Excel u otro programa compatible con archivos `.xlsx`:** Para poder abrir y visualizar el archivo Excel donde se guardará la tasa del BCV.

*   **Sistema Operativo Windows:** El Programador de tareas es una herramienta nativa de Windows.  Si utilizas otro sistema operativo (Linux o macOS), existen herramientas equivalentes (como `cron` en Linux o `launchd` en macOS), pero esta guía se centra en el Programador de tareas de Windows para mayor simplicidad.

3. Código Python (TASABCV.py)

Crea un nuevo archivo de texto y guárdalo con el nombre `TASABCV.py` (asegúrate de que la extensión sea `.py`).  Copia y pega el siguiente código Python en el archivo:

import http.client
import json
import openpyxl
import os

# Información de la API y archivo Excel
API_URL = "ve.dolarapi.com"
API_ENDPOINT = "/v1/dolares/oficial"
EXCEL_FILENAME = "tasa_dolar_bcv.xlsx"
SHEET_NAME = "TasaBCV"
CELL_FECHA = "A1"
CELL_VALOR = "B1"

def obtener_dolar_oficial():
    """Obtiene la información del dólar oficial desde la API."""
    conn = http.client.HTTPSConnection(API_URL)
    conn.request("GET", API_ENDPOINT)
    res = conn.getresponse()
    data = res.read()
    return json.loads(data.decode("utf-8"))

def escribir_en_excel_local(data, excel_filename, sheet_name, cell_fecha, cell_valor):
    """Escribe los datos en un archivo de Excel local usando openpyxl."""
    fecha = data['fechaActualizacion']
    valor = data['promedio']

    # Comprueba si el archivo Excel ya existe
    if os.path.exists(excel_filename):
        # Si existe, carga el archivo existente
        workbook = openpyxl.load_workbook(excel_filename)
        # Comprueba si la hoja de cálculo ya existe
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
        else:
            # Si la hoja no existe, crea una nueva hoja
            sheet = workbook.create_sheet(sheet_name)
    else:
        # Si el archivo no existe, crea un nuevo libro de trabajo (archivo Excel)
        workbook = openpyxl.Workbook()
        sheet = workbook.active # Hoja activa por defecto
        sheet.title = sheet_name # Cambia el nombre de la hoja

    # Escribe los datos en las celdas especificadas
    sheet[cell_fecha] = fecha
    sheet[cell_valor] = valor

    # Guarda el libro de trabajo (archivo Excel)
    try:
        workbook.save(excel_filename)
        print(f"Datos escritos en el archivo Excel local: '{excel_filename}', hoja '{sheet_name}', celdas '{cell_fecha}' y '{cell_valor}'.")
    except Exception as e:
        print(f"Ocurrió un error al guardar el archivo Excel: {e}")


if __name__ == '__main__':
    datos_dolar = obtener_dolar_oficial()
    if datos_dolar:
        print("Datos de la API obtenidos:")
        print(datos_dolar)

        escribir_en_excel_local(datos_dolar, EXCEL_FILENAME, SHEET_NAME, CELL_FECHA, CELL_VALOR)
        print("Datos escritos exitosamente en archivo Excel local.")
    else:
        print("No se pudieron obtener los datos de la API.")
