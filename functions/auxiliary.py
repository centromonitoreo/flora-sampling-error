# import fiona
import geopandas as gpd
import os
import zipfile
from tkinter import  messagebox
import shutil
import time
from functions.root import ROOT_GDB
import re
import pandas as pd
import gc  # Módulo para manejo de la recolección de basura

def get_list_ncobert(gdb_path):
    """
    Retrieves a list of unique values for 'N_COBERT' from a layer within a geospatial database file (GDB).

    This function loads data from the 'PuntoMuestreoFlora' layer in the GDB file, extracts the unique values 
    from the 'N_COBERT' column, and returns this list. Additionally, it ensures that the memory used by the 
    GeoDataFrame is released after use.

    Args:
        gdb_path (str): Path to the GDB file containing the data layer.

    Returns:
        list: A list of unique values for 'N_COBERT' extracted from the 'PuntoMuestreoFlora' layer.
    """


    lista_unicos_n_cobert = []  # Inicializar la lista

    df_punto_muestreo_flora = gpd.read_file(gdb_path, layer="PuntoMuestreoFlora")
                
    # Extraer la lista de valores únicos de 'N_COBERT'
    lista_unicos_n_cobert = df_punto_muestreo_flora['N_COBERT'].unique().tolist()

    # Usada para test con muchos n_cobert
    lista_unicos_n_cobert = [elemento for elemento in lista_unicos_n_cobert for _ in range(1)]

    # Liberar memoria eliminando la referencia al GeoDataFrame
    del df_punto_muestreo_flora
    
    # Forzar la recolección de basura
    gc.collect()

    return lista_unicos_n_cobert

def get_list_ncobert_excel(file_path):
    """
    Reads an Excel file and extracts the unique values from the 'N_COBERT' column.

    This function opens the Excel file located at the given path and retrieves
    all the unique values from the 'N_COBERT' column. The unique values are 
    returned as a list.

    Args:
        file_path (str): The path to the Excel file to be read.

    Returns:
        list: A list of unique values from the 'N_COBERT' column. If an error 
              occurs while loading the file, the function prints an error message 
              and returns None.

    Raises:
        Exception: If there is an error while loading the Excel file (e.g., 
                   file not found, incorrect format), an error message is printed.

    Example:
        file_path = 'path/to/excel_file.xlsx'
        ncobert_list = get_list_ncobert_excel(file_path)
    """
    
    try:
        df = pd.read_excel(file_path)
        lista_unicos_n_cobert = df['N_COBERT'].unique().tolist()
        return lista_unicos_n_cobert
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")


def extract_zip_gdb(file_path, root_unzip):
    """
    Extracts a ZIP file containing a .gdb folder and returns the path to the extracted folder.

    This function first removes any existing .gdb folder in the destination directory, unzips the specified 
    ZIP file, and checks for .gdb folders in the extraction location. If a .gdb folder is found, its path 
    is returned. In case of errors, warning or error messages are displayed.

    Args:
        file_path (str): Path to the ZIP file containing the .gdb folder.
        root_unzip (str): Path where the ZIP file contents will be extracted.

    Returns:
        str: Path to the first .gdb folder found after extraction, or None if none is found.
    """

    delete_folder(ROOT_GDB)

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(root_unzip)  # Descomprimir el archivo

        # Esperar un poco para asegurarnos de que todo se haya descomprimido correctamente
        time.sleep(1)

        # Verificar si hay carpetas `.gdb` en la carpeta de descompresión
        gdb_dirs = [f for f in os.listdir(root_unzip) if f.endswith('.gdb') and os.path.isdir(os.path.join(root_unzip, f))]
        if gdb_dirs:
            return os.path.join(root_unzip, gdb_dirs[0])  # Retorna la primera carpeta .gdb encontrada
        else:
            messagebox.showwarning("Advertencia", "No se encontró ninguna carpeta .gdb en el ZIP.")
            return None

    except zipfile.BadZipFile:
        messagebox.showerror("Error", "El archivo ZIP está corrupto o no es válido.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al descomprimir: {str(e)}")

    return None

def delete_folder(path_folder):
    """
    Deletes all files and subdirectories in the specified folder.

    This function checks if the provided path exists and is a directory. If so, 
    it iterates over all items within the folder, deleting files and subdirectories. 
    If an error occurs during the process, the exception is caught, and an error 
    message is printed.

    Args:
        path_folder (str): The path to the folder to be cleared.

    Returns:
        None: This function does not return any value.
    """

    if os.path.exists(path_folder) and os.path.isdir(path_folder):
        try:
            # Iterar sobre todos los elementos en el directorio
            for item in os.listdir(path_folder):
                item_path = os.path.join(path_folder, item)  # Obtener la ruta completa del elemento
                
                # Comprobar si es un archivo o un directorio y eliminarlo
                if os.path.isfile(item_path):
                    os.remove(item_path)  # Eliminar archivo
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)  # Eliminar directorio y su contenido
            
            # print(f"Se han eliminado todos los archivos y subdirectorios en {path_folder}.")
        except Exception as e:
            print(f"Error al limpiar el directorio: {e}")
    else:
        print(f"El directorio {path_folder} no existe o no es un directorio.")      

def clean_sheet_name(name):
    """
    Cleans a spreadsheet name by removing invalid characters and limiting its length 
    to 31 characters, which is the maximum allowed by Excel.

    Args:
        name (str): The spreadsheet name to be cleaned.

    Returns:
        str: The cleaned and validated name for the spreadsheet.
    """

    # Reemplazar caracteres no permitidos
    nombre_limpio = re.sub(r'[\\/*\[\]:?"]', '', name)
    # Limitar el nombre a 31 caracteres (máximo permitido por Excel)
    return nombre_limpio[:31]  

def validate_gdb_folder(file_path):
    """
    Validates whether the specified folder is an ArcGIS geodatabase.

    Checks if the folder name ends with '.gdb' and contains certain required 
    files typical of a geodatabase.

    Args:
        file_path (str): The path to the directory to be validated as a geodatabase.

    Returns:
        bool: True if the folder is a valid geodatabase, False otherwise.
    """

    # Comprueba si el nombre de la carpeta termina con '.gdb'
    if not file_path.endswith('.gdb'):
        return False
    
    # Verifica si contiene algunos archivos típicos de una geodatabase
    required_files = ['a00000001.gdbtable', 'gdb', 'spx']  # Lista de archivos comunes en una gdb
    folder_contents = os.listdir(file_path)

    for req_file in required_files:
        if not any(file_name.endswith(req_file) for file_name in folder_contents):
            return False
    
    return True

def is_number(valor):
    """
    Checks if the provided value can be converted to a number.

    Attempts to convert the value to a float and determines if the conversion 
    is successful.

    Args:
        value: The value to be checked (can be of any type).

    Returns:
        bool: True if the value can be converted to a number, False otherwise.
    """

    try:
        float(valor)  # Intenta convertir el valor a un número flotante
        return True
    except ValueError:
        return False
    
def truncar_string(texto: str, longitud: int):
    """
    Truncates a text to a specific length and adds "..." if necessary.

    If the text exceeds the specified length, it is cut and "..." is added at the end.
    If the text is shorter than the specified length, spaces are added at the end 
    to make the text match the desired length.

    Args:
        text (str): The text to be truncated or padded.
        length (int): The maximum allowed length for the text.

    Returns:
        str: The text truncated or padded to the specified length.
    """

    
    if len(texto) > longitud:
        # Retornar los primeros 'longitud' caracteres + "..."
        return texto[:longitud] + "..."
    else:
        # Rellenar con espacios al final para que el texto tenga la longitud deseada
        return texto.ljust(longitud+8)
    
def compressed_files(root_file_: str, name_file: str):
    """
    Creates a ZIP file containing all files in the specified folder, excluding the ZIP file if it already exists, 
    and deletes the original files.

    This function compresses all files in the specified folder into a ZIP file. the original files are deleted, leaving only the resulting ZIP file.

    Parameters:
    root_file_ (str): The path to the folder containing the files to be compressed.
    name_file (str): The name of the ZIP file to be created. Must include the '.zip' extension.

    Returns:
    None: The function does not return any value. Prints messages about the operation to the console.

    Example:
    >>> compressed_files("/home/user/result", "compressed_file.zip")
    (Creates a 'compressed_file.zip' file in the '/home/user/result' folder with all the files in the folder, except the ZIP file.)
"""


    # Nombre de la carpeta y del archivo zip resultante
    carpeta = root_file_
    archivo_zip = os.path.join(carpeta, name_file)

    # Obtener la lista de archivos en la carpeta
    archivos = [f for f in os.listdir(carpeta) if os.path.isfile(os.path.join(carpeta, f))]

    # Verificar el número de archivos en la carpeta
    if len(archivos) <= 0:
        # print(f'Solo hay {len(archivos)} archivo(s) en la carpeta. No se realizará ninguna acción.')
        pass
    else:
        # Crear el archivo zip en la carpeta 'result'
        with zipfile.ZipFile(archivo_zip, 'w') as zipf:
            for file in archivos:
                if file != name_file:  # Excluir el archivo zip si ya existe
                    # Añadir cada archivo sin incluir la ruta de la carpeta 'result'
                    zipf.write(os.path.join(carpeta, file), arcname=file)

        # print(f'Archivos comprimidos en {archivo_zip}')

        # Borrar todos los archivos excepto el archivo zip resultante
        for file in archivos:
            if file != name_file:
                os.remove(os.path.join(carpeta, file))

def list_files(root_files: str):
    """
    Retrieves a list of full file paths from a specific folder.

    This function takes the path to a folder and generates a list of full paths for all files in that folder. 
    The list includes only files, excluding subdirectories and other non-file items.

    Parameters:
    root_files (str): The path to the directory where files will be searched.

    Returns:
    list: A list of strings, where each string is the full path to a file in the specified directory.

    Example:
    >>> root_files = '/home/user/documents'
    >>> list_files(root_files)
    ['/home/user/documents/file1.txt', '/home/user/documents/file2.csv', '/home/user/documents/file3.pdf']
"""


    lista_archivos = [os.path.join(root_files, archivo) for archivo in os.listdir(root_files) if os.path.isfile(os.path.join(root_files, archivo))]
  
    return lista_archivos