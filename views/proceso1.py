import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from io import BytesIO
import streamlit as st
import calendar
import xlsxwriter
import io
import locale


# Subir los archivos
def subir_archivos():
    archivo1 = st.file_uploader("Sube el primer archivo", type=["csv", "xlsx"])
    archivo2 = st.file_uploader("Sube el segundo archivo para el proceso", type=["csv", "xlsx"])
    if archivo1 and archivo2:
        return archivo1, archivo2
    return None, None

# Renombrar las columnas
def renombrar_columnas(df):
    df = df.rename(columns={
        'FECHA DE ENVIO': 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)',
        'DIAS': 'DIAS TRANSCURRIDOS (HABILES)'
    })
    return df

# Ordenar por fecha
def ordenar_por_fecha(df):
    df['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(
        df['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce')
    df = df.sort_values(by='FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)', ascending=True)
    df = df.drop(columns=['Nº'], errors='ignore')  # Elimina la columna anterior si existe
    df.insert(0, 'Nº', range(1, len(df) + 1))       # Inserta una nueva columna "Nº"
    return df


# Reemplazar la columna "GUIA ASEGURADO" por "DIRECCION" en la hoja MENSAJERO
def reemplazar_columna_guias(df_mensajero):
    if 'GUIA ASEGURADO' in df_mensajero.columns:
        df_mensajero['DIRECCION'] = df_mensajero['GUIA ASEGURADO']
        df_mensajero = df_mensajero.drop(columns=['GUIA ASEGURADO'])
    return df_mensajero

# Función para modificar la columna "RAD DE SALIDA"
def modificar_rad_salida(df):
    df['RAD DE SALIDA'] = df['RAD DE SALIDA'].apply(
        lambda x: f"SAL-{x}" if not str(x).startswith("SAL-") else x)
    return df

# ---- GRAFICAS --------
# ----- ---- ---- ---- ----
# Colores
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') 
colores = ['#FFB897', '#B8E6A7', '#809bce', "#64a09d", '#CBE6FF']

# Función para graficar el gráfico de torta con los datos de COURIER y agregarlo al Excel
def autopct_custom(pct):
    return f'{pct:.1f}%' if pct >= 1 else ''

def graficar_torta_courier(df_courier, writer):
    df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(
        df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce'
    )

    df_courier['MES'] = df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.month

    conteo = df_courier.groupby('MES').size().reset_index(name='Conteo')
    total = conteo['Conteo'].sum()
    conteo['Porcentaje'] = conteo['Conteo'] / total
    conteo['MesNombre'] = conteo['MES'].apply(lambda x: calendar.month_name[x].capitalize())

    # Función personalizada para etiquetas
    def autopct_custom(pct):
        return f'{pct:.1f}%' if pct >= 1 else ''

    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(
        conteo['Conteo'],
        labels=None,
        autopct=autopct_custom,
        startangle=90,
        colors=colores
    )

    ax.legend(wedges, conteo['MesNombre'], title="Meses", loc="upper right", bbox_to_anchor=(1.1, 1))
    ax.set_title('Distribución de las Guías por Mes (COURIER)', pad=30)
    ax.axis('equal')

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    worksheet = writer.sheets['COURIER TABLA']
    worksheet.insert_image('D2', 'grafico_torta.png', {'image_data': img_stream})

    plt.close(fig)

# Función para graficar el gráfico de barras con los datos de COURIER y agregarlo al Excel
def graficar_barras_courier_por_mes(df_courier, writer):
    # Convertir fechas
    df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(
        df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce'
    )

    # Eliminar nulos
    df_courier = df_courier.dropna(subset=['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'])

    # Agrupar por mes
    df_courier['MES'] = df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.month
    conteo = df_courier.groupby('MES').size().reset_index(name='Conteo')
    conteo['MesNombre'] = conteo['MES'].apply(lambda x: calendar.month_name[x].capitalize())

    # Ordenar por número de mes
    conteo = conteo.sort_values('MES')

    # Crear gráfico
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(conteo['MesNombre'], conteo['Conteo'], color=colores)

    ax.set_title('Distribución de las Guías por Mes (COURIER)', pad=20)
    ax.set_xlabel('Mes')
    ax.set_ylabel('Cantidad de guías')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    # Guardar imagen
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    # Insertar en Excel
    worksheet = writer.sheets['COURIER TABLA']
    worksheet.insert_image('D2', 'grafico_barras_courier.png', {'image_data': img_stream})

    plt.close(fig)

# Función para graficar el gráfico de torta con los datos de MENSAJERO y agregarlo al Excel
def graficar_torta_mensajero(df_mensajero, writer):
    # Convertir la columna de fechas a datetime
    df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(
        df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce'
    )

    # Extraer el mes
    df_mensajero['MES'] = df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.month

    # Contar las guías por mes
    conteo = df_mensajero.groupby('MES').size().reset_index(name='Conteo')
    total = conteo['Conteo'].sum()
    conteo['Porcentaje'] = conteo['Conteo'] / total
    conteo['MesNombre'] = conteo['MES'].apply(lambda x: calendar.month_name[x].capitalize())

    # Función para ocultar etiquetas menores a 1%
    def autopct_custom(pct):
        return f'{pct:.1f}%' if pct >= 1 else ''

    # Crear gráfico
    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(
        conteo['Conteo'],
        labels=None,
        autopct=autopct_custom,
        startangle=90,
        colors=colores
    )

    ax.legend(wedges, conteo['MesNombre'], title="Meses", loc="upper right", bbox_to_anchor=(1.1, 1))
    ax.set_title('Distribución de las Guías por Mes (MENSAJERO)', pad=30)
    ax.axis('equal')

    # Guardar imagen
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    # Insertar en Excel
    worksheet = writer.sheets['MENSAJERO TABLA']
    worksheet.insert_image('D2', 'grafico_torta_mensajero.png', {'image_data': img_stream})

    plt.close(fig)

# Función para graficar el gráfico de barras con los datos de MENSAJERO y agregarlo al Excel
def graficar_barras_mensajero_por_dia(df_mensajero, writer):

    # Convertir la columna de fechas a datetime
    df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(
        df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce'
    )

    # Eliminar nulos
    df_mensajero = df_mensajero.dropna(subset=['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'])

    # Agrupar por día (fecha completa)
    conteo_por_dia = df_mensajero.groupby(
        df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.date
    ).size().reset_index(name='Conteo')

    # Crear gráfico
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(conteo_por_dia['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], conteo_por_dia['Conteo'], color=colores)

    # Formatear fechas en el eje X
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b'))  # Ej: 24-Jun
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))   # Mostrar todos los días (ajusta el intervalo si hay muchos)

    plt.xticks(rotation=45, ha='right')
    ax.set_title('Distribución diaria de las Guías (MENSAJERO)', pad=20)
    ax.set_xlabel('Fecha')
    ax.set_ylabel('Cantidad de guías')
    plt.tight_layout()

    # Guardar imagen
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    # Insertar en Excel
    worksheet = writer.sheets['MENSAJERO TABLA']
    worksheet.insert_image('D2', 'grafico_barras_mensajero.png', {'image_data': img_stream})

    plt.close(fig)

# ---- TABLAS --------
# ----- ---- ---- ---- ----

# Función para crear la tabla en la hoja COURIER TABLA
def crear_tabla_courier(df_courier, writer):
    # Convertir la columna de fechas a datetime
    df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce')

    # Extraer el mes de la fecha
    df_courier['MES'] = df_courier['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.month

    # Contar las guías por mes
    conteo_guias_courier = df_courier.groupby('MES').size().reset_index(name='Conteo de guías')

    # Crear una nueva hoja COURIER TABLA con la tabla (sin sobrescribir COURIER)
    worksheet = writer.book.add_worksheet('COURIER TABLA')

    # Definir formato para el encabezado y bordes
    formato_encabezado = writer.book.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    formato_bordes = writer.book.add_format({'border': 1})

    # Escribir los encabezados en la tabla
    worksheet.write('A1', 'FECHA DE ENVIO', formato_encabezado)
    worksheet.write('B1', 'CUENTA DE GUIAS', formato_encabezado)

    # Escribir los datos de la tabla (sin espacio entre encabezado y datos)
    for i, row in conteo_guias_courier.iterrows():
        worksheet.write(i + 1, 0, calendar.month_name[row['MES']])  # Mes en formato textual
        worksheet.write(i + 1, 1, row['Conteo de guías'])

        # Aplicar bordes solo a las celdas con datos (meses y conteo de guías)
        worksheet.write(i + 1, 0, calendar.month_name[row['MES']], formato_bordes)
        worksheet.write(i + 1, 1, row['Conteo de guías'], formato_bordes)

    # Agregar total general en la última fila
    total_guias_courier = conteo_guias_courier['Conteo de guías'].sum()
    worksheet.write(len(conteo_guias_courier) + 1, 0, 'Total General', formato_encabezado)
    worksheet.write(len(conteo_guias_courier) + 1, 1, total_guias_courier, formato_bordes)

    # Ajustar el ancho de las columnas
    worksheet.set_column('A:B', 20)

    # Graficar el gráfico de torta y agregarlo a la hoja
    graficar_torta_courier(df_courier, writer)
    graficar_barras_courier_por_mes(df_courier, writer)

# Crear la tabla en la hoja MENSAJERO TABLA
def crear_tabla_mensajero(df_mensajero, writer):
    # Convertir la columna de fechas a datetime
    df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'] = pd.to_datetime(df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'], errors='coerce')

    # Extraer el mes de la fecha
    df_mensajero['MES'] = df_mensajero['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].dt.month

    # Contar las guías por mes
    conteo_guias = df_mensajero.groupby('MES').size().reset_index(name='Conteo de guías')

    # Crear una nueva hoja MENSAJERO TABLA con la tabla (sin sobrescribir MENSAJERO)
    worksheet = writer.book.add_worksheet('MENSAJERO TABLA')

    # Definir formato para el encabezado y bordes
    formato_encabezado = writer.book.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    formato_bordes = writer.book.add_format({'border': 1})

    # Escribir los encabezados en la tabla
    worksheet.write('A1', 'FECHA DE ENVIO', formato_encabezado)
    worksheet.write('B1', 'CUENTA DE GUIAS', formato_encabezado)

    # Escribir los datos de la tabla (sin espacio entre encabezado y datos)
    for i, row in conteo_guias.iterrows():
        worksheet.write(i + 1, 0, calendar.month_name[row['MES']])  # Mes en formato textual
        worksheet.write(i + 1, 1, row['Conteo de guías'])

        # Aplicar bordes solo a las celdas con datos (meses y conteo de guías)
        worksheet.write(i + 1, 0, calendar.month_name[row['MES']], formato_bordes)
        worksheet.write(i + 1, 1, row['Conteo de guías'], formato_bordes)

    # Agregar total general en la última fila
    total_guias = conteo_guias['Conteo de guías'].sum()
    worksheet.write(len(conteo_guias) + 1, 0, 'Total General', formato_encabezado)
    worksheet.write(len(conteo_guias) + 1, 1, total_guias, formato_bordes)

    # Ajustar el ancho de las columnas
    worksheet.set_column('A:B', 20)

    # Graficar el gráfico de torta y agregarlo a la hoja
    graficar_torta_mensajero(df_mensajero, writer)
    graficar_barras_mensajero_por_dia(df_mensajero, writer)

# ---PROCESAMIENTO DE DATOS ----

#  Función de procesamiento de datos
def procesar_datos(archivo1, archivo2):
    try:
        # Procesar el primer archivo
        if archivo1.name.endswith((".xlsx", ".xls")):
            xl1 = pd.ExcelFile(archivo1)
            if 'COURIER' in xl1.sheet_names and 'MENSAJERO' in xl1.sheet_names:
                df1_courier = xl1.parse('COURIER')
                df1_mensajero = xl1.parse('MENSAJERO')
            else:
                return None, None
        else:
            return None, None

        # Procesar el segundo archivo
        if archivo2.name.endswith((".xlsx", ".xls")):
            xl2 = pd.ExcelFile(archivo2)
            if 'COURIER' in xl2.sheet_names and 'MENSAJERO' in xl2.sheet_names:
                df2_courier = xl2.parse('COURIER')
                df2_mensajero = xl2.parse('MENSAJERO')
            else:
                return None, None
        else:
            return None, None

        # Renombrar las columnas para que coincidan en cada hoja
        df1_courier = renombrar_columnas(df1_courier)
        df2_courier = renombrar_columnas(df2_courier)
        df1_mensajero = renombrar_columnas(df1_mensajero)
        df2_mensajero = renombrar_columnas(df2_mensajero)

        # Reemplazar "GUIA ASEGURADO" por "DIRECCION" en la hoja MENSAJERO
        df1_mensajero = reemplazar_columna_guias(df1_mensajero)
        df2_mensajero = reemplazar_columna_guias(df2_mensajero)

        # Combinar los DataFrames de las hojas COURIER y MENSAJERO
        df_courier = pd.concat([df1_courier, df2_courier], ignore_index=True)
        df_mensajero = pd.concat([df1_mensajero, df2_mensajero], ignore_index=True)

        # Ordenar por fecha (de más antiguo a más reciente)
        df_courier = ordenar_por_fecha(df_courier)
        df_mensajero = ordenar_por_fecha(df_mensajero)

        # Modificar la columna "RAD DE SALIDA" usando la nueva función
        df_courier = modificar_rad_salida(df_courier)
        df_mensajero = modificar_rad_salida(df_mensajero)

        # Eliminar duplicados (opcional, según el caso)
        df_courier = df_courier.drop_duplicates()
        df_mensajero = df_mensajero.drop_duplicates()

        return df_courier, df_mensajero

    except Exception as e:
        print(f"Error: {e}")
        return None, None

# Guardar el archivo Excel generado
def guardar_archivo_excel(df_courier, df_mensajero):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Guardar las hojas COURIER y MENSAJERO (sin sobrescribir MENSAJERO existente)
            df_courier.to_excel(writer, sheet_name='COURIER', index=False)
            df_mensajero.to_excel(writer, sheet_name='MENSAJERO', index=False)

            # Tablas
            crear_tabla_mensajero(df_mensajero, writer)
            crear_tabla_courier(df_courier, writer)

        output.seek(0)
        return output
    except Exception as e:
        return None

# Descargar el archivo generado
def descargar_excel(archivo_combinado):
    try:
        st.download_button(
            label="📥 Descargar",
            data=archivo_combinado,
            file_name="logística_472.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error al preparar el archivo para descarga: {e}")