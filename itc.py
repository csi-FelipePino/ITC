import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import time
import threading
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

############################################################################################
########### Directorios
############################################################################################

# Crear la ventana principal
root = tk.Tk()
root.title("Selección de Directorios")
root.geometry("700x300")

# Etiqueta para el título
label_title = tk.Label(root, text="Selecciona un directorio principal", font=("Arial", 16))
label_title.pack(pady=20)

# Etiqueta para mostrar el directorio seleccionado
label_dir = tk.Label(root, text="", font=("Arial", 10))
label_dir.pack(pady=10)

# Variables globales para rutas
selected_directory = None
csv_file_path = None
excel_file_path = None

# Función para seleccionar el directorio
def select_directory():
    global selected_directory, csv_file_path, excel_file_path
    selected_directory = filedialog.askdirectory()
    
    if selected_directory:
        # Mostrar el directorio principal seleccionado
        label_dir.config(text=f"Directorio principal seleccionado: {selected_directory}")
        
        # Buscar archivo CSV en el directorio principal
        for file in os.listdir(selected_directory):
            if file.endswith(".csv"):
                csv_file_path = os.path.join(selected_directory, file)
                break
        else:
            messagebox.showerror("Error", "No se encontró ningún archivo CSV en el directorio seleccionado")
            return
        
        # Buscar archivo Excel en la subcarpeta 'data'
        excel_file_path = os.path.join(selected_directory, "data", "itc.xlsx")
        if not os.path.isfile(excel_file_path):
            messagebox.showerror("Error", "No se encontró el archivo itc.xlsx en la carpeta 'data'")
            return
        

# Botón para seleccionar el directorio
btn_select_dir = tk.Button(root, text="Seleccionar Directorio", command=select_directory)
btn_select_dir.pack(pady=10)

# Ejecutar la ventana principal
root.mainloop()

############################################################################################
########### Logica
############################################################################################

def rename_unnamed_columns(df):
    col_counter = 1

    new_columns = []
    for col in df.columns:
        if 'Unnamed' in col:
            new_columns.append(f'{col_counter}')
            col_counter += 1
        else:
            new_columns.append(col)

    # Asignamos los nuevos nombres de columna al DataFrame
    df.columns = new_columns
    return df

df = pd.read_csv(csv_file_path, delimiter=';')

df = rename_unnamed_columns(df)

############################################################################################
########### Tabla 1

work_999_index = df[df.iloc[:, 0].str.contains('Work.999', na=False)].index

if not work_999_index.empty:
    next_row_index = work_999_index[0] + 1
    tabla1 = df.iloc[next_row_index]

Nombre_Inter = tabla1.iloc[0]
IP =  tabla1.iloc[2]
ID =  tabla1.iloc[3]
Num_grupos =  tabla1.iloc[6]
Num_E_Logicos =  tabla1.iloc[8]
Num_S_Logicos =  tabla1.iloc[7]
Num_paneles =  tabla1.iloc[4]
Num_D_Logicos =  tabla1.iloc[12]

data = {
    'Dato': [
        'Nombre de la intersección',
        'IP controlador',
        'ID panel de control',
        'Número de grupos',
        'Cantidad de escenarios lógicos',
        'Cantidad de secuencias lógicas',
        'Número de planes',
        'Número de detectores lógicos'
    ],
    'Valor': [
        Nombre_Inter,
        IP,
        ID,
        Num_grupos,
        Num_E_Logicos,
        Num_S_Logicos,
        Num_paneles,
        Num_D_Logicos
    ]
}

df1 = pd.DataFrame(data)

################################### Completar datos
sheet_name = "Tabla 1-1"
wb = load_workbook(excel_file_path)
ws = wb[sheet_name]

ws["B2"] = df1.iloc[0,1]
ws["B3"] = df1.iloc[1,1]
ws["B4"] = df1.iloc[2,1]
ws["B5"] = df1.iloc[3,1]
ws["B6"] = df1.iloc[4,1]
ws["B7"] = df1.iloc[5,1]
ws["B8"] = df1.iloc[6,1]
ws["B9"] = df1.iloc[7,1]

wb.save(excel_file_path)



############################################################################################
########### Tabla 2

work_998_index = df[df.iloc[:, 0].str.contains('Work.998', na=False)].index

if not work_998_index.empty and not work_999_index.empty:
    tabla2 = df.iloc[work_998_index[0] +1 : work_999_index[0]-1] #quito el work y el Next
else:
    print("No se encontraron ambas entradas de 'Work.998' y 'Work.999'")

df2 = pd.DataFrame({
    'Grupo': pd.Series(dtype='str'),   # Columna vacía 'Grupo'
    'Nombre': pd.Series(dtype='str'),  # Columna vacía 'Nombre'
    'Tiempos1': pd.Series(dtype='object'),  # Columna vacía 'Tiempos1'
    'Tiempos2': pd.Series(dtype='object'),  # Columna vacía 'Tiempos2'
    'Tiempos3': pd.Series(dtype='object')   # Columna vacía 'Tiempos3'
})


resultados = []
num=1

for idx, row in tabla2.iterrows():
    valores = [row[0], row[11], row[14], row[9]]

    # Insertar el número de fila al inicio de la lista
    fila_data = [num] + valores  # +1 porque queremos que la primera fila sea 1
    num = num +1
    resultados.append(fila_data)

# Convertir la lista de resultados en un DataFrame
df_nuevo = pd.DataFrame(resultados, columns=['Fila', 'Dato1', 'Dato2', 'Dato3', 'Dato4'])

df2['Grupo'] = df_nuevo['Fila'].values.tolist()
df2['Nombre'] = df_nuevo[ 'Dato1'].values.tolist()
df2['Tiempos1'] = df_nuevo['Dato2'].values.tolist()
df2['Tiempos2'] = df_nuevo['Dato3'].values.tolist()
df2['Tiempos3'] = df_nuevo['Dato4'].values.tolist()


def procesar_valor(valor):
    # Realizamos el split y tomamos el primer elemento
    primer_valor = valor.split('-')[0]
    # Convertimos a entero
    return int(float(primer_valor))

# Aplicar la función para todas las columnas de tiempos
df2['Tiempos1'] = df2['Tiempos1'].apply(procesar_valor)
df2['Tiempos2'] = df2['Tiempos2'].apply(procesar_valor)
df2['Tiempos3'] = df2['Tiempos3'].apply(procesar_valor)


################################### Completar datos
sheet_name = "Tabla 1-2"
wb = load_workbook(excel_file_path)
ws = wb[sheet_name]

# GRUPO
for i in range(len(df2)):
    ws[f"A{i + 2}"] = df2.iloc[i, 0]  # Escribir en la columna B, fila dinámica (i + 1)

# NOMBRE
for i in range(len(df2)):
    ws[f"B{i + 2}"] = df2.iloc[i, 1]  

# TIEMPOS 1
for i in range(len(df2)):
    ws[f"C{i + 2}"] = df2.iloc[i, 2]  

# TIEMPOS 2
for i in range(len(df2)):
    ws[f"D{i + 2}"] = df2.iloc[i, 3]  

# TIEMPOS 3
for i in range(len(df2)):
    ws[f"D{i + 2}"] = df2.iloc[i, 4]  

wb.save(excel_file_path)



############################################################################################
########### Tabla 3

work_997_index = df[df.iloc[:, 0].str.contains('Work.997', na=False)].index

if not work_997_index.empty and not work_998_index.empty:
    tabla3 = df.iloc[work_997_index[0] +1 : work_998_index[0]-1] #quito el work y el Next
else:
    print("No se encontraron ambas entradas de 'Work.997' y 'Work.998'")

grupo_vals = df2['Nombre'].values
# Crear una tabla simétrica vacía usando estos valores
tabla_simetrica = pd.DataFrame(index=grupo_vals, columns=grupo_vals)

for i in range(len(tabla_simetrica.columns)):
    tabla_simetrica.iloc[:, i] = tabla3.iloc[:, i].values

def convertir_a_int(valor):
    try:
        return int(float(valor))  # Convertir a float y luego a int
    except (ValueError, TypeError):
        return valor  # Si no se puede convertir, devolver el valor original

tabla_simetrica = tabla_simetrica.applymap(convertir_a_int)

def eliminar_espacios(valor):
    if isinstance(valor, str):  # Verificar si el valor es una cadena de texto
        return valor.strip()  # Eliminar espacios en blanco al inicio y al final
    return valor  # Si no es una cadena, devolver el valor tal como está

# Aplicar la función a todo el DataFrame
tabla_simetrica = tabla_simetrica.applymap(eliminar_espacios)
tabla_simetrica = tabla_simetrica.replace('.', '-')
tabla_simetrica = tabla_simetrica.fillna('-')

################################### Completar datos
sheet_name = "Tabla 1-3"
wb = load_workbook(excel_file_path)
ws = wb[sheet_name]

# GRUPO VERTICAL
for i in range(len(df2)):
    ws[f"B{i + 3}"] = df2.iloc[i,1]

# GRUPO HORIZONTAL
for i in range(len(df2)):
    col_letra = get_column_letter(3 + i)  # Empieza en la columna C (columna 3)
    ws[f"{col_letra}2"] = df2.iloc[i, 1]  # Escribir en la columna dinámica y fila fija 3

# Iterar sobre todas las filas y columnas de tabla_simetrica
for col_idx in range(len(tabla_simetrica.columns)):
    col_letter = get_column_letter(3 + col_idx)  # 3 corresponde a la columna "C" y avanzamos
    for row_idx in range(len(tabla_simetrica)):
        ws[f"{col_letter}{row_idx + 3}"] = tabla_simetrica.iloc[row_idx, col_idx]


wb.save(excel_file_path)
############################################################################################
########### Tabla 4

work_006_index = df[df.iloc[:, 0].str.contains('Work.006', na=False)].index
work_007_index = df[df.iloc[:, 0].str.contains('Work.007', na=False)].index

if not work_006_index.empty and not work_007_index.empty:
    tabla4 = df.iloc[work_006_index[0] +1 : work_007_index[0]-1] #quito el work y el Next
else:
    print("No se encontraron ambas entradas de 'Work.997' y 'Work.998'")

grupo_vals = df2['Nombre'].values
# Crear una tabla simétrica vacía usando estos valores
tabla_simetrica2 = pd.DataFrame(index=grupo_vals, columns=grupo_vals)

for i in range(len(tabla_simetrica.columns)):
    tabla_simetrica2.iloc[:, i] = tabla4.iloc[:, i].values

def convertir_a_int(valor):
    try:
        primer_valor = valor.split('-')[1]
        return int(float(primer_valor))  # Convertir a float y luego a int
    except (ValueError, TypeError):
        return valor  # Si no se puede convertir, devolver el valor original

tabla_simetrica2 = tabla_simetrica2.applymap(convertir_a_int)

def eliminar_puntos(valor):
    if isinstance(valor, str):  # Verificar si el valor es una cadena de texto
        return valor.replace(".", "")  # Eliminar todos los espacios en blanco
    return valor  # Si no es una cadena, devolver el valor tal como está

# Aplicar la función a todo el DataFrame
df4 = tabla_simetrica2.applymap(eliminar_puntos)


################################### Completar datos
sheet_name = "Tabla 1-4"
wb = load_workbook(excel_file_path)
ws = wb[sheet_name]

# GRUPO VERTICAL
for i in range(len(df2)):
    ws[f"B{i + 3}"] = df2.iloc[i,1]

# GRUPO HORIZONTAL
for i in range(len(df2)):
    col_letra = get_column_letter(3 + i)  # Empieza en la columna C (columna 3)
    ws[f"{col_letra}2"] = df2.iloc[i, 1]  # Escribir en la columna dinámica y fila fija 3

# Iterar sobre todas las filas y columnas de tabla_simetrica
for col_idx in range(len(df4.columns)):
    col_letter = get_column_letter(3 + col_idx)  # 3 corresponde a la columna "C" y avanzamos
    for row_idx in range(len(df4)):
        ws[f"{col_letter}{row_idx + 3}"] = df4.iloc[row_idx, col_idx]

wb.save(excel_file_path)
############################################################################################
########### Tabla 5

work_009_index = df[df.iloc[:, 0].str.contains('Work.009', na=False)].index
work_010_index = df[df.iloc[:, 0].str.contains('Work.010', na=False)].index

if not work_009_index.empty and not work_010_index.empty:
    tabla5 = df.iloc[work_009_index[0] +1 : work_010_index[0]-1] #quito el work y el Next
else:
    print("No se encontraron ambas entradas de 'Work.997' y 'Work.998'")

segunda_columna_lista = tabla5.iloc[:, 1].tolist()


fase_vals = [1, 2, 3, 4]
grupo_vals = df2['Nombre'].values

df5 = pd.DataFrame({'fase': fase_vals})


for grupo in grupo_vals:
    df5[grupo] = pd.Series([None] * len(fase_vals))  # Inicializamos las columnas con valores vacíos

segunda_columna_lista = list(map(int, segunda_columna_lista))

# Definir las potencias de 2 que vamos a usar
potencias = [16, 8, 4, 2]

# Función para descomponer el valor en potencias de 2 más grandes posibles
def descomponer_en_potencias(valor):
    resultado = []
    for potencia in potencias:
        if valor >= potencia:
            resultado.append(potencia)
            valor -= potencia  # Restamos la potencia del valor
    return resultado

# Función para llenar el DataFrame
def llenar_columna(valor, nombre_columna):
    # Descomponemos el valor en las potencias de 2 correspondientes
    potencias_descompuestas = descomponer_en_potencias(valor)

    # Llenamos el DataFrame según las potencias descompuestas
    for potencia in potencias_descompuestas:
        # Fila 1 corresponde a 2, fila 2 a 4, fila 3 a 8, fila 4 a 16
        if potencia == 2:
            df5.loc[0, nombre_columna] = 'ok'  # Fila 1
        elif potencia == 4:
            df5.loc[1, nombre_columna] = 'ok'  # Fila 2
        elif potencia == 8:
            df5.loc[2, nombre_columna] = 'ok'  # Fila 3
        elif potencia == 16:
            df5.loc[3, nombre_columna] = 'ok'  # Fila 4

# Recorrer la lista de valores binarios y llenar las columnas
for idx, valor in enumerate(segunda_columna_lista):
    nombre_columna = grupo_vals[idx]  # Usamos el nombre de la columna directamente
    llenar_columna(int(valor), nombre_columna) 

df5.dropna(how='all', inplace=True)

################################### Completar datos
sheet_name = "Tabla 1-5"
wb = load_workbook(excel_file_path)
ws = wb[sheet_name]


# GRUPO HORIZONTAL
for i in range(len(df2)):
    col_letra = get_column_letter(2 + i)  
    ws[f"{col_letra}2"] = df2.iloc[i, 1]  

# Iterar sobre todas las filas y columnas de tabla_simetrica
# Obtener el valor y el estilo de la celda J14
valor_j14 = ws["J14"].value
estilo_j14 = ws["J14"]._style  # El estilo de la celda J14

# Recorrer las columnas y filas de df5
for col_idx in range(len(df5.columns)):
    col_letter = get_column_letter(1 + col_idx)  # Comienza en la columna "A"
    for row_idx in range(len(df5)):
        if df5.iloc[row_idx, col_idx] == "ok":
            # Obtener la dirección de la celda de destino
            destino = f"{col_letter}{row_idx + 3}"

            # Asignar el valor de J14 a la celda de destino
            ws[destino].value = valor_j14

            # Copiar el formato desde J14
            ws[destino]._style = estilo_j14  # Aplicar el estilo de J14 a la celda de destino

# fila - columna
wb.save(excel_file_path)