import os
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import json
from tkinterdnd2 import DND_FILES, TkinterDnD
import csv
import re
import pandas as pd
import string
from openpyxl import load_workbook
import aux_fun as aux



# --- Function to close the application on Escape key press ---
def close_on_escape(event):
    event.widget.quit()

# --- Function to get the column mapping between two lists of columns ---
def get_column_index(col_str):
    index = 0
    for c in col_str:
        index = index * 26 + (ord(c.upper()) - ord('A') + 1)
    return index - 1  # Restamos 1 porque Excel tiene una numeración de columnas 1-based y Python 0-based


# --- Function to validate a file ---
def validate_file(full_path):
    if not os.path.exists(full_path):
        return f"El archivo '{os.path.basename(full_path)}' no existe."
    if not full_path.lower().endswith(".xlsx"):
        return f"El archivo '{os.path.basename(full_path)}' no es un archivo .xlsx."
    return None  # No errors


# --- Function to transform the EXCEL file to a CSV file ---
def from_excel_to_csv(input_excel, output_csv):
    df = pd.read_excel(input_excel, dtype=str)
    df.fillna('-', inplace=True)
    
    # Replace lines breaks only in text columns
    for col in df.select_dtypes(include='object'):
        df[col] = df[col].str.replace('\n', ' ', regex=False)
    
    df.to_csv(output_csv, sep=';', index=False, encoding='utf-8')
    # print(f"File converted successfully: {output_csv}")


# --- Function to get KPI values and indexes (column) from CSV ---
def get_kpi_values_and_indexes(csv_file):
    with open(csv_file, 'r') as f:
        reader = csv.reader(f, delimiter=';')
        for i, row in enumerate(reader):
            if i == 3:
                index_numerico = [(index, int(value)) for index, value in enumerate(row) if value.isdigit()]
                break
    return index_numerico


# --- Function to print the first row of each CSV file ---
def print_first_row_of_csv(csv_file_1, csv_file_2):

    array_head_colores = []
    array_head_base = []
    # print(f"Primera fila de {file_colores_csv}:")
    with open(csv_file_1, 'r') as f:
        reader = csv.reader(f, delimiter=';')
        for i, row in enumerate(reader):
            if i == 0:
                # print(row)
                array_head_colores = row
                break

    # print(f"Primera fila de {file_dataBase_csv}:")
    with open(csv_file_2, 'r') as f:
        reader = csv.reader(f, delimiter=';')
        for i, row in enumerate(reader):
            if i == 0:
                # print(row)
                array_head_base = row
                break
    return array_head_colores, array_head_base

# --- Function to get the column letter from its index ---
def get_letter_column(n):
    """Transform a column index (0-based) to its letter representation (a, b, ..., z, aa, ab, ...)."""
    result = ''
    while n >= 0:
        remainder = n % 26
        result = chr(ord('a') + remainder) + result
        n = n // 26 - 1
    return result

# --- Function to obtain the column mapping ---
def get_column_mapping(columnas_archivo, array_head_base):
    """
    Search each field in columnas_archivo within array_head_base and returns
    a dictionary mapping the field name to the corresponding column letter.
    
    Args:
        columnas_archivo (list): List of strings with the names of the columns in the file.
        array_head_base (list): List of strings with the names of the base columns.
        
    Returns:
        dict: A dictionary where the keys are the names of the columns in columnas_archivo
              and the values are the corresponding column letters in array_head_base.
              If a field is not found, its value will be None.
            
    """
    mapping = {}
    for column in columnas_archivo:
        try:
            index = array_head_base.index(column)
            # Get the letter of the column
            column_letter = get_letter_column(index)
            # Replace dots with underscores to match the format of the example JSON
            formatted_key = column.replace(".", "_")
            # The correction is here: we simply assign the letter of the column
            mapping[formatted_key] = column_letter
        except ValueError:
            formatted_key = column.replace(".", "_")
            mapping[formatted_key] = None # Indicates that the column was not found
    return mapping


# --- Function to save the column mapping in a JSON file ---
def save_column_mapping_to_json(columns_file1, columns_file2, array_head_base, array_head_colors, file_name="mapeo_columnas.json"):
    """
    Generates the mapping of columns and saves the result in a JSON file.

    Args: 
        columns_file1 (list): List of strings with the names of the columns in file 1.
        columns_file2 (list): List of strings with the names of the columns in file 2.
        array_head_base (list): Base list to compare with columns_file1.
        array_head_colors (list): Base list to compare with columns_file2.
        file_name (str, optional): The name of the JSON file to save.
                                   Defaults to "mapeo_columnas.json".

    """
    mapping_file1 = get_column_mapping(columns_file1, array_head_base)
    mapping_file2 = get_column_mapping(columns_file2, array_head_colors)

    data = {
        "archivo1": mapping_file1,
        "archivo2": mapping_file2
    }

    with open(file_name, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


# --- Function to load the data from JSON file ---
def obtener_valores_json(nombre_fichero):
    """
    Carga el contenido de un fichero JSON y devuelve los valores de los dos objetos principales.

    Args:
        nombre_fichero (str): El nombre del fichero JSON a cargar.

    Returns:
        tuple: Una tupla que contiene dos listas:
               - valores1: Los valores del primer objeto en el JSON.
               - valores2: Los valores del segundo objeto en el JSON.
               Devuelve (None, None) si el fichero no se encuentra o tiene un formato inesperado.
    """
    try:
        with open(nombre_fichero, 'r') as f:
            data = json.load(f)

        if not isinstance(data, dict) or len(data) != 2:
            print("Error: El fichero JSON debe contener exactamente dos objetos.")
            return None, None

        primer_objeto = list(data.values())[0]
        segundo_objeto = list(data.values())[1]

        if not isinstance(primer_objeto, dict) or not isinstance(segundo_objeto, dict):
            print("Error: Los elementos del JSON deben ser diccionarios.")
            return None, None

        valores1 = list(primer_objeto.values())
        valores2 = list(segundo_objeto.values())

        return valores1, valores2

    except FileNotFoundError:
        print(f"Error: El fichero '{nombre_fichero}' no se encontró.")
        return None, None
    except json.JSONDecodeError:
        print(f"Error: No se pudo decodificar el JSON en '{nombre_fichero}'.")
        return None, None
    except IndexError:
        print("Error: El fichero JSON no tiene la estructura esperada.")
        return None, None

# --- Function to format the name of an ingredient and the cooking process ---
#     -> format: main ingredient + . + cooking process (no spaces, first letter of each word in uppercase)
def format_ingredient_and_cooking_method(main_ingredient, cooking_method):
    
    # Remove spaces and capitalize the first letter of each word in the main ingredient
    formatted_main_ingredient = ''.join(word.capitalize() for word in main_ingredient.split())

    # Remove spaces and capitalize the first letter of each word in the cooking method
    formatted_cooking_method = ''.join(word.capitalize() for word in cooking_method.split())
    
    # Combine them with a dot
    formatted_string = f"{formatted_main_ingredient}.{formatted_cooking_method}"

    # detelete spaces
    formatted_string = formatted_string.replace(" ", "")

    # Delete New: only if exist
    formatted_string = formatted_string.replace("New:", "")
    
    return formatted_string


# --- Return the no cooking methods ---
def list_cooking_methods(file):
    # Only load the third sheet (index 2) of the Excel file
    sheet_3 = pd.read_excel(file, sheet_name=2)  # The index starts at 0

    # Convert the content to a vector (list of lists)
    array = sheet_3.values.tolist()

    # Convert the first element of each sublist to minuscules
    array = [[str(row[0]).lower()] + row[1:] for row in array]

    # convert all the elements of the array to minuscules
    array = [[str(element).lower() for element in row] for row in array]

    # print("lista de métodos de cocinado:", array)
    # array = [['fry', None], ['shallow fry', None], ['deep fry', None], ['simmering', None], ['mix boil', None], ['steaming', None], ['reducing', None], ['bain marie', None], ['pressure cooker', None], ['sous vide', None], ['melting', None], ['grill', None], ['pan fry', None], ['melting', None], ['cut', 'No cooking method'], ['peel', 'No cooking method'], ['wash', 'No cooking method'], ['sliced', 'No cooking method'], ['…','No cooking method']]
    return array

# --- Function to check if there are matches ---
def check_if_matches(array_cooking, cooking_process):

    match_found, no_cooking_method = False, False

    # For each cooking process in the list of cooking methods, check if it matches with the cooking process of the row
    for process, is_cooking_method in array_cooking:

        # If there is a match, we will check if it is a cooking method or not, and we will break the loop
        if cooking_process == process:
            match_found = True

            if is_cooking_method == "No cooking method":
                no_cooking_method = True
            break

    return match_found, no_cooking_method



# --- Function to check if there is a cooking method, it's not, or if it is new ---
def check_cooking_method(row, indexK, file):

    value = None

    # Get the cooking method of the row (orange column) and convert it to lowercase
    cooking_process = row.iloc[indexK].lower()
    # print(f"Cooking process: '{cooking_process}'")

    # Get the list of cooking methods from the Excel file (third sheet)
    array_cooking =list_cooking_methods(file)
    # print(f"List of cooking methods: {array_cooking}")

    # Check if the cooking process is NEW
    if cooking_process.startswith("new:"):

        value = 0 # CASE 0 -> NEW cooking method

    else: 

        print("Checking if the cooking process is in the list of cooking methods...")

        # print("DEBUG array_cooking (elementos):")
        # for elem in array_cooking:
        #     print("  ->", elem, "len =", len(elem) if hasattr(elem, "__len__") else "N/A")

        # Check if the cooking process is in the list of cooking methods
        # for process, is_cooking_method in array_cooking:
        for elem in array_cooking:
            process = elem[0]
            is_cooking_method = elem[1]
            print(f"Comparing '{cooking_process}' with '{process}'...")
            if cooking_process == process:
                print(f"Match found for cooking process: '{cooking_process}'")
                print(f"Is it a cooking method? '{is_cooking_method}'")
                if is_cooking_method == "no cooking method":
                    value = 2 # CASE 2 -> No cooking method
                    print("No cooking method found")
                    break
                else:
                    value = 1 # CASE 1 -> Cooking method found
                    print("Cooking method found")
                    break
 
    return value


# --- Function to convert the cooking method into a format that matches the catalogue ---
def convert_cooking_method_format(cooking_method_color_excel):

    cooking_method = ""

    if cooking_method_color_excel == "fry": # Fry
        cooking_method = "FRY"

    elif  cooking_method_color_excel == "shallowfry": # Shallow fry
        cooking_method = "SHALLOW_FRY"

    elif  cooking_method_color_excel == "deepfry": # Deep fry
        cooking_method = "DEEP_FRY"

    elif  cooking_method_color_excel == "simmering": # Simmering
        cooking_method = "SIMMERING"

    elif  cooking_method_color_excel == "mixboil": # Mix Boil
        cooking_method = "MIX_BOIL"

    elif  cooking_method_color_excel == "steaming": # Steaming
        cooking_method = "STEAMING"

    elif  cooking_method_color_excel == "reducing": # Reducing
        cooking_method = "REDUCING"

    elif  cooking_method_color_excel == "bainmarie": # Bain Marie
        cooking_method = "BAIN_MARIE"

    elif  cooking_method_color_excel == "pressurecooker": # Pressure cooker
        cooking_method = "PRESSURE_COOKER"

    elif  cooking_method_color_excel == "panfry": # Pan Fry
        cooking_method = "PAN_FRY"

    elif  cooking_method_color_excel == "melting": # Melting
        cooking_method = "MELTING"

    return cooking_method


def obtain_file_name_1(file):
    name_file1 = os.path.basename(file)
    name_file1 = name_file1.replace('.xlsx', '')
    return name_file1


def has_partial_match(cell_value, search_words, STOPWORDS):
    if pd.isna(cell_value):
        return False

    cell_text = str(cell_value).lower()
    cell_words = re.findall(r"\w+", cell_text)

    # Elimina stopwords y deja las palabras "útiles"
    cell_words = [w for w in cell_words if w not in STOPWORDS]

    # Coincidencia si al menos una palabra del search está incluida (como subcadena) en alguna celda
    for search_word in search_words:
        for cell_word in cell_words:
            if search_word in cell_word:  # <-- permite subpalabras (ej. 'prawns' dentro de 'kingprawns')
                return True
    return False



# --- Get the number of matches ---
def get_number_of_matches(case, df_base, main_ingredient_to_search, STOPWORDS):

    # print("BUSCANDO MATCHES...")
    # Search for exact matches and partial matches
    # print(f"   → Searching for: '{main_ingredient_to_search}'")
    exact_matches = df_base[df_base.iloc[:, 0] == main_ingredient_to_search]
    # partial_matches = df_base[df_base.iloc[:, 0].str.contains(main_ingredient_to_search, na=False, case=False)] # busqueda parcial 

    # --- Partial matches ---
    search_words = [
        w for w in re.findall(r"\w+", str(main_ingredient_to_search).lower())
        if w not in STOPWORDS
    ]

    # print(f"   → Search words for partial match: {search_words}")
    # print (f"   → Total rows in base DataFrame: {len(df_base)}")

    partial_matches = df_base[df_base.iloc[:, 0].apply(
        lambda x: has_partial_match(x, search_words, STOPWORDS)
    )]


    # print(f"   → Exact matches found: {len(exact_matches)}")
    # print(f"   → Partial matches found: {len(partial_matches)}")

    # Case 0: No matches
    # Case 1: One exact match
    # Case 2: One partial match
    # Case 3: More than one exact match
    # Case 4: More than one partial match

    # Check the number of matches
    if exact_matches.empty:

        if partial_matches.empty:
            case = 0
            print(f"\n🔸 No matches found for '{main_ingredient_to_search}'.")

        elif len(partial_matches) == 1: # print('   1 match (parcial)')
            case = 2
            print(f"\n🔸 Partial match found for '{main_ingredient_to_search}':")
            print(f"   → Found: {partial_matches.iloc[0, 0]}")

        else:
            case = 4
            print(f"\n🔸 Multiple partial matches found for '{main_ingredient_to_search}':")
            for val in partial_matches.iloc[:, 0]:
                print(f"   → Found: {val}")

    elif len(exact_matches) == 1:
        case = 1
        print(f"\n🔸 Exact match found for '{main_ingredient_to_search}':")
        print(f"   → Found: {exact_matches.iloc[0, 0]}")

    else:
        case = 3
        print(f"\n🔸 Multiple exact matches found for '{main_ingredient_to_search}':")
        for val in exact_matches.iloc[:, 0]:
            print(f"   → Found: {val}")

    return case, exact_matches, partial_matches


# Formats both words in the same way: first letter of each word in each variable in uppercase and the rest in lowercase, and removes spaces. And concatenates them with a dot in between
# Example: word1 = "Ejemplo Prueba" and word2 = "Air Fried" -> "Ejemploprueba.AirFried"
def format_words(word1, word2):

    # word1: ingredient
    # word2: cooking method

    # Format the cooking method to have the first letter of each word in uppercase and the rest in lowercase ("Air Fried -> Air Fried")
    word2 = string.capwords(word2)

    # Delete all the spaces in the cooking method ("Air Fried -> AirFried")
    word2 = word2.replace(" ", "")

    # Format the ingredient to have the first letter of each word in uppercase and the rest in lowercase ("KingPrawns -> Kingprawns")
    word1 = string.capwords(word1)

    # Delete all the spaces in the ingredient ("Ejemplo Prueba -> EjemploPrueba")
    word1 = word1.replace(" ", "")

    # Concatenate the main ingredient and the cooking method with a dot in between, and convert it to lowercase (MainIngredient.AirFried)
    new_ID_cooking_method_ = word1 + '.' + word2

    return new_ID_cooking_method_