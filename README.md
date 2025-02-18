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

```python
import http.client
import json
import openpyxl
import os

# Información de la API y archivo Excel
API_URL = "ve.dolarapi.com"
API_ENDPOINT = "/v1/dolares/oficial"
EXCEL_FILENAME = "TASABCV.xlsx"
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

```
Explicación del código:

Importaciones:
        import http.client: Biblioteca para hacer peticiones HTTP a la API.
        import json: Biblioteca para trabajar con datos en formato JSON que devuelve la API.
        import openpyxl: Biblioteca para leer y escribir archivos de Excel (.xlsx).
        import os: Biblioteca para interactuar con el sistema operativo, en este caso, para verificar si el archivo Excel existe.

- Variables de Configuración:
        API_URL = "ve.dolarapi.com" y API_ENDPOINT = "/v1/dolares/oficial": Definen la URL base y el punto final de la API del BCV que vamos a consultar.
        EXCEL_FILENAME = "tasa_dolar_bcv.xlsx": Nombre del archivo Excel que se creará (o se actualizará) en la misma carpeta que el script Python.
        SHEET_NAME = "TasaBCV": Nombre de la hoja de cálculo dentro del archivo Excel donde se guardarán los datos.
        CELL_FECHA = "A1" y CELL_VALOR = "B1": Celdas específicas dentro de la hoja de cálculo donde se escribirán la fecha y el valor de la tasa del BCV. Puedes cambiar estas variables si deseas guardar los datos en otras celdas.

- Función obtener_dolar_oficial():
        Establece una conexión HTTPS a la API del BCV (API_URL).
        Realiza una petición GET al punto final /v1/dolares/oficial (API_ENDPOINT).
        Recibe la respuesta de la API en formato JSON.
        Decodifica la respuesta JSON y la devuelve como un diccionario de Python.

- Función escribir_en_excel_local(data, excel_filename, sheet_name, cell_fecha, cell_valor):
     * Recibe los datos de la API (data), el nombre del archivo Excel, el nombre de la hoja y las celdas de               destino como argumentos.
     
     * Extrae la fecha de actualización y el valor de la tasa del BCV del diccionario data: fecha =                       data['fechaActualizacion'] y valor = data['promedio']. Importante: Se accede directamente a las claves             'fechaActualizacion' y 'promedio' que están presentes en la respuesta de la API.
     
     * Verifica si el archivo Excel existe: Utiliza os.path.exists() para comprobar si el archivo especificado            en excel_filename ya existe en la misma carpeta que el script.
     
     * Carga o crea el libro de trabajo Excel:
       Si el archivo Excel existe, lo abre (openpyxl.load_workbook()).
       Si el archivo Excel no existe, crea un nuevo libro de trabajo (openpyxl.Workbook()).

     * Crea o selecciona la hoja de cálculo:
       Si la hoja con el nombre sheet_name existe dentro del libro de trabajo, la selecciona.
       Si la hoja no existe, crea una nueva hoja con el nombre sheet_name y la selecciona como hoja activa.

     * Escribe los datos en las celdas: Escribe la fecha en la celda especificada por cell_fecha y el valor en la         celda especificada por cell_valor dentro de la hoja de cálculo seleccionada.

     * Guarda el archivo Excel: Guarda los cambios en el archivo Excel utilizando workbook.save(excel_filename).          Imprime un mensaje de éxito o error en la consola.

- Bloque if __name__ == '__main__':
     * Este bloque de código se ejecuta solo cuando el script se ejecuta directamente (no cuando se importa como          un módulo).

     * Llama a la función obtener_dolar_oficial() para obtener los datos de la API.
        Si los datos se obtienen correctamente:
            Imprime los datos de la API en la consola (para verificar).
            Llama a la función escribir_en_excel_local() para guardar los datos en el archivo Excel local.
            Imprime un mensaje de éxito.
        Si no se pueden obtener los datos de la API, imprime un mensaje de error.

```
```
Creación del Archivo .bat (ejecutar_tasa_bcv.bat)
Para automatizar la ejecución del script Python con el Programador de tareas de Windows, crearemos un archivo .bat (archivo por lotes) que contiene el comando para ejecutar el script.

   * Abre el Bloc de notas de Windows.

   * Copia y pega el siguiente código en el Bloc de notas:
```
@echo off
"C:\Program Files\Python311\python.exe" "C:\Users\AnalistaIT\Desktop\TASA BCV\TASABCV.py"
pause

```
¡Importante! Reemplaza las rutas de ejemplo con las rutas CORRECTAS de tu sistema:

  C:\Program Files\Python311\python.exe: Reemplaza esto con la ruta completa al ejecutable de Python en tu computadora.  Puedes encontrar esta ruta ejecutando el comando where python o where python3 en el Símbolo del sistema o PowerShell.  Asegúrate de que la ruta sea correcta y que el archivo python.exe realmente exista en esa ubicación.

  C:\Users\USUARIO\Desktop\TASA BCV\TASABCV.py: Reemplaza esto con la ruta completa al archivo TASABCV.py que creaste en el paso anterior. Guarda el archivo TASABCV.py en una carpeta de tu elección (por ejemplo, Desktop\TASA BCV).  Luego, copia la ruta completa a ese archivo y reemplaza la ruta de ejemplo en el archivo .bat.

¡Asegúrate de mantener las comillas dobles " alrededor de AMBAS rutas si contienen espacios!

La línea pause al final es opcional. Se utiliza para mantener la ventana de comandos abierta después de ejecutar el script, para que puedas ver si hay algún mensaje de error (útil para depuración).  Puedes quitar la línea pause una vez que estés seguro de que el script funciona correctamente.

Guarda el archivo .bat:

   Haz clic en "Archivo" -> "Guardar como..." en el Bloc de notas.
    Navega hasta la misma carpeta donde guardaste el archivo TASABCV.py (C:\Users\AnalistaIT\Desktop\TASA BCV en este ejemplo).
    En "Nombre de archivo:", escribe un nombre descriptivo para el archivo .bat terminado en .bat (ej: ejecutar_tasa_bcv.bat).
    En "Guardar como tipo:", selecciona "Todos los archivos (.)". ¡Esto es crucial para que el archivo se guarde como .bat y no como .txt!
    Haz clic en "Guardar".

Prueba el archivo .bat manualmente:

Navega hasta la carpeta donde guardaste el archivo .bat con el Explorador de archivos de Windows.
  Haz doble clic en el archivo .bat (ej: ejecutar_tasa_bcv.bat).
    Debería abrirse brevemente una ventana de comandos y luego cerrarse (si no has incluido la línea pause).
     Verifica que se haya creado o actualizado el archivo Excel tasa_dolar_bcv.xlsx en la misma carpeta. Abre el archivo Excel y confirma que la fecha y el valor de la tasa del BCV se hayan guardado correctamente.
```
