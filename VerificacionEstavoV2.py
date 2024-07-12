import os
import pandas as pd
import xlsxwriter

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/EstadoBeneficio_Credito_08172024.xlsx'

if not os.path.isfile(file_path):
    raise FileNotFoundError(f"{file_path} no encontrado.")
else:
    print(f"Archivo {file_path} encontrado.")

# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")

# Carga los DataFrames de trabajo
try:
    # Lectura de los insumos en un diccionario de dataframes
    dic_insumos = pd.read_excel(file_path, sheet_name=['CAROLINADEUDAS','2021-2','2022-1','2022-2','2023-1','2023-2'], engine='openpyxl')

    # Limpia los nombres de columnas
    for df in dic_insumos.values():
        df.columns = df.columns.str.strip()

    credito, piam20212, piam20221, piam20222, piam20231, piam20232 = dic_insumos['CAROLINADEUDAS'],dic_insumos['2021-2'], dic_insumos['2022-1'], dic_insumos['2022-2'], dic_insumos['2023-1'], dic_insumos['2023-2']
except Exception as e:
    print(f"Error al cargar los DataFrames: {e}")

# Función para actualizar 'EstadoBeneficio'
def actualizar_estado_beneficio(row):
    if pd.isna(row['EstadoBeneficio']) and not pd.isna(row['ESTADO F']):
        return row['ESTADO F']
    return row['EstadoBeneficio']

# Función para actualizar 'CriterioExclusion'
def actualizar_criterio_exclusion(row):
    if pd.isna(row['CriterioExclusion']) and not pd.isna(row['ESTADO']):
        return row['ESTADO']
    return row['CriterioExclusion']

# Función para realizar merge y actualización
def merge_and_update(df_credito, df_piam):

    # Elimina la columna 'BOLETA' si ya existe en df_credito
    if 'BOLETA' in df_credito.columns:
        df_credito = df_credito.drop(columns=['BOLETA'])

    df_merged = pd.merge(
        df_credito,
        df_piam[['BOLETA', 'ESTADO F', 'ESTADO']],
        left_on='Documento',
        right_on='BOLETA',
        how='left')

    df_merged['EstadoBeneficio'] = df_merged.apply(actualizar_estado_beneficio, axis=1)
    df_merged['CriterioExclusion'] = df_merged.apply(actualizar_criterio_exclusion, axis=1)

    df_merged.drop(columns=['ESTADO F', 'ESTADO'], inplace=True)
    
    return df_merged

# Lista de DataFrames piam
piam_list = [piam20212, piam20221, piam20222, piam20231, piam20232]

# Itera sobre la lista de DataFrames piam para realizar los merges y actualizaciones
for df_piam in piam_list:
    credito = merge_and_update(credito, df_piam)

# Asegúrate de eliminar la columna 'BOLETA' si todavía está presente
if 'BOLETA' in credito.columns:
    credito = credito.drop(columns=['BOLETA'])

# Agrupa los resultados para el análisis
filtro_Facturacion = credito.groupby('Periodico Academico')['Documento'].size().reset_index(name='Poblacion')

# Exporta el resultado a un nuevo archivo Excel
output_path = "/content/VerificacionEstadoBeneficioCreditoCartera.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    filtro_Facturacion.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=1, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Generalidades']
    formato = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    worksheet.merge_range('B1:C1', "INSUMO FACTURACION 2024-1", formato)
    
    credito.to_excel(writer, sheet_name='Facturacion20241', index=False)
 
print(f"Archivo guardado en {output_path}")
