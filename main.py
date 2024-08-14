import os
import pandas as pd

path_principal = "C:\\bin\\calidad\\"

lst_files_xls = [
    "hola", "hola", "hola", "hola", 
    "hola", "hola", "hola", "Codigo de Calendario tcccp0110m000.xlsx", 
    "Codigos impositivos tcmcs0137m000.xlsx", "componentesdecosto_prueba.xlsx", "Conjuntos de unidades tcmcs0106m000.xlsx", "hola", 
    "hola", "hola", "Formatos de direccion tccom4135s000.xlsx", "hola"
    ]

lst_template_destino = [
    "hola", "hola", "hola", "hola", 
    "hola", "hola", "hola", "Código de calendario", 
    "Códigos impositivos", "Componentes de costo", "Conjunto de unidades", "hola",
    "hola", "hola", "Formatos dirección", "hola"
    ]


def comprobar_archivos(path_principal, lst_files_xls):
    archivos_encontrados = []
    archivos_no_encontrados = []

    for archivo in lst_files_xls:
        ruta_completa = os.path.join(path_principal, archivo)
        if os.path.isfile(ruta_completa):
            archivos_encontrados.append(archivo)
        else:
            archivos_no_encontrados.append(archivo)

    return archivos_encontrados, archivos_no_encontrados


def detect_duplicates_and_missing(file_path):
    # Verificar existencia del archivo
    if not os.path.isfile(file_path):
        return "Error: El archivo no existe en la ruta especificada."
    # Cargar el archivo Excel
    try:
        excel_data = pd.ExcelFile(file_path)
    except Exception as e:
        return f"Error al leer el archivo: {e}"
    
    # Diccionario para almacenar los resultados
    results = {}

    # Iterar sobre todas las hojas
    for sheet_name in excel_data.sheet_names:
        if sheet_name != "enum":
            # Leer cada hoja en un DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Inicializar variables para cálculos
            total_values = 0
            duplicated_values_total = 0
            null_values_total = 0
            empty_values_total = 0

            # Diccionario para almacenar los resultados por columna
            column_results = {}

            # Iterar sobre las columnas
            for column in df.columns:
                # Calcular total de valores y valores duplicados
                total = df[column].notna().count()
                duplicated = df[column][df[column].duplicated()].count()
                null_values = df[column].isna().sum()
                empty_values = (df[column] == '').sum()

                # Total de valores en la columna, considerando nulos y vacíos
                total_column_values = total + null_values + empty_values

                # Si hay valores, calcular el porcentaje de duplicidad
                if total_column_values > 0:
                    percentage_duplicates = (duplicated / total_column_values) * 100
                    percentage_nulls = (null_values / total_column_values) * 100
                    percentage_empty = (empty_values / total_column_values) * 100
                    column_results[column] = {
                        'total': total_column_values,
                        'duplicated': duplicated,
                        'nulls': null_values,
                        'empties': empty_values,
                        'percentage_duplicates': percentage_duplicates,
                        'percentage_nulls': percentage_nulls,
                        'percentage_empty': percentage_empty
                    }

                    total_values += total_column_values
                    duplicated_values_total += duplicated
                    null_values_total += null_values
                    empty_values_total += empty_values

            # Calcular el porcentaje de duplicidad promedio, nulos y vacíos para la hoja
            if total_values > 0:
                avg_percentage_duplicates = (duplicated_values_total / total_values) * 100
                avg_percentage_nulls = (null_values_total / total_values) * 100
                avg_percentage_empty = (empty_values_total / total_values) * 100
            else:
                avg_percentage_duplicates = 0
                avg_percentage_nulls = 0
                avg_percentage_empty = 0
        
            results[sheet_name] = {
                'average_percentage_duplicates': avg_percentage_duplicates,
                'average_percentage_nulls': avg_percentage_nulls,
                'average_percentage_empty': avg_percentage_empty,
                'columns': column_results
            }

        # Formatear los resultados
        result_str = []
        for sheet_name, sheet_data in results.items():
            result_str.append(f"Hoja: {sheet_name}")
            result_str.append(f"  Porcentaje promedio de duplicidad: {sheet_data['average_percentage_duplicates']:.2f}%")
            result_str.append(f"  Porcentaje promedio de valores nulos: {sheet_data['average_percentage_nulls']:.2f}%")
            result_str.append(f"  Porcentaje promedio de valores vacíos: {sheet_data['average_percentage_empty']:.2f}%")
            for column, col_data in sheet_data['columns'].items():
                result_str.append(f"  Columna: {column}")
                result_str.append(f"    Total: {col_data['total']}")
                result_str.append(f"    Duplicados: {col_data['duplicated']}")
                result_str.append(f"    Nulos: {col_data['nulls']}")
                result_str.append(f"    Vacíos: {col_data['empties']}")
                result_str.append(f"    Porcentaje de duplicados: {col_data['percentage_duplicates']:.2f}%")
                result_str.append(f"    Porcentaje de nulos: {col_data['percentage_nulls']:.2f}%")
                result_str.append(f"    Porcentaje de vacíos: {col_data['percentage_empty']:.2f}%")

        return "\n".join(result_str)


def calcular_porcentajes_columna(file_path, columna_a_evaluar):
    # Verificar la existencia del archivo
    if not os.path.isfile(file_path):
        return "Error: El archivo no existe en la ruta especificada."
    
    # Cargar el archivo Excel
    try:
        excel_data = pd.ExcelFile(file_path)
    except Exception as e:
        return f"Error al leer el archivo: {e}"
    
    # Inicializar contadores globales
    total_values = 0
    duplicated_values_total = 0
    null_values_total = 0

    # Iterar sobre todas las hojas
    for sheet_name in excel_data.sheet_names:
        if sheet_name != "enum":
            # Leer la hoja en un DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Verificar si la columna especificada existe en la hoja
            if columna_a_evaluar in df.columns:
                # Calcular total de valores y valores duplicados
                total = df[columna_a_evaluar].notna().count()
                duplicated = df[columna_a_evaluar][df[columna_a_evaluar].duplicated()].count()
                null_values = df[columna_a_evaluar].isna().sum()

                # Total de valores en la columna, considerando nulos y vacíos
                total_column_values = total

                # Acumular los valores globales
                total_values += total_column_values
                duplicated_values_total += duplicated
                null_values_total += null_values
    
    # Calcular los porcentajes globales
    if total_values > 0:
        percentage_duplicates = (duplicated_values_total / total_values) * 100
        percentage_nulls = (null_values_total / total_values) * 100
    else:
        percentage_duplicates = 0
        percentage_nulls = 0

    # Formatear los resultados
    """
    result_str = []
    result_str.append(f"Columna: {columna_a_evaluar}")
    result_str.append(f"Total de registros evaluados: {total_values}")
    result_str.append(f"Porcentaje de duplicados: {percentage_duplicates:.2f}%")
    result_str.append(f"Porcentaje de valores nulos/vacios: {percentage_nulls:.2f}%")
    return "\n".join(result_str)
    """
    result_dict = {
        "Columnas": columna_a_evaluar,
        "Total_Registros_Evaluados": total_values,
        "Porcentaje_Duplicados": f"{percentage_duplicates:.2f}%",
        "Porcentaje_Vacios": f"{percentage_nulls:.2f}%"
    }
    return result_dict


def buscar_parametro(result_str, termino_a_buscar):
    if termino_a_buscar in result_str:
        return result_str[termino_a_buscar]
    else:
        return f"La llave '{termino_a_buscar}' no se encontró en el diccionario."


def actualizar_columnas_excel(file_path, resultado_incumple_valor, resultado_cumple_valor, output_path=None):
    pass


if __name__ == "__main__":
    # Uso del servicio
    file_path = r'C:\bin\calidad\Conjuntos de unidades tcmcs0106m000.xlsx'  # Cambia esta ruta según tu archivo
    print(detect_duplicates_and_missing(file_path))

    # Validamos que exista la ruta + archivo
    encontrados, no_encontrados = comprobar_archivos(path_principal, lst_files_xls) 
    #print(encontrados)
    for archivo in encontrados:
        #print(path_principal+archivo)
        # Creamos la variable con la ruta completa
        file_path = path_principal+archivo
        
        if archivo == "componentesdecosto_prueba1.xlsx":
            resultado_compania = calcular_porcentajes_columna(file_path, "Compañía")
            #print(type(resultado_compania))
            Porcentaje_Duplicados = buscar_parametro(resultado_compania, "Porcentaje_Duplicados")
            print(Porcentaje_Duplicados)

            resultado_Componente_de_costo = calcular_porcentajes_columna(file_path, "Componente de costo")
            Porcentaje_Duplicados = buscar_parametro(resultado_Componente_de_costo, "Porcentaje_Duplicados")
            print(Porcentaje_Duplicados)
