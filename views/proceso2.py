import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import streamlit as st
from datetime import datetime, timedelta
import holidays

# Subir los archivos
def subir_archivos():
    archivo1 = st.file_uploader("Sube el primer archivo", type=["csv", "xlsx"], key="archivo1")
    archivo2 = st.file_uploader("Sube la matriz de trayectos", type=["csv", "xlsx"], key="archivo2")
    
    if archivo1 is not None and archivo2 is not None:
        return archivo1, archivo2
    return None, None

# Función para leer los archivos
def leer_archivos(archivo1, archivo2):
    try:
        # Leer el primer archivo (solo la hoja 'COURIER')
        if archivo1.name.endswith((".xlsx", ".xls")):
            xl1 = pd.ExcelFile(archivo1)
            if 'COURIER' in xl1.sheet_names:
                df1_courier = xl1.parse('COURIER')  # Leer solo la hoja 'COURIER'
                st.write("Primer archivo cargado exitosamente con la hoja 'COURIER'.")
                st.write("Columnas en el archivo 1 ('COURIER'):", df1_courier.columns)  # Depuración: Ver columnas
            else:
                st.error("La hoja 'COURIER' no está presente en el archivo 1.")
                return None, None
        else:
            st.error("El primer archivo no es válido.")
            return None, None
        
        # Leer el segundo archivo (todo el contenido)
        if archivo2.name.endswith((".xlsx", ".xls")):
            df2_courier = pd.read_excel(archivo2)  # Leer el archivo completo
            st.write("Segundo archivo cargado exitosamente.")
            st.write("Columnas en el archivo 2:", df2_courier.columns)  # Depuración: Ver columnas
        else:
            st.error("El segundo archivo no es válido.")
            return None, None
        
        return xl1, df2_courier  # Devolvemos xl1 junto con df2_courier
    except Exception as e:
        st.error(f"Error al leer los archivos: {e}")
        return None, None


def hacer_merge(xl1, df2):
    try:
        # Cargar la hoja 'COURIER' de xl1 como DataFrame
        df1_courier = xl1.parse('COURIER')  # Asegúrate de usar parse para obtener un DataFrame de la hoja 'COURIER'
        
        # Eliminar espacios en blanco en las columnas de ambos DataFrames
        df1_courier.columns = df1_courier.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        # Mostrar las columnas de df2 para depuración
        st.write("Columnas en el archivo 2 (df2):", df2.columns)

        # Verificar si la columna exacta existe en df2
        columna_buscada = 'CORRESPONDENCIA PRIORITARIA, ENCOMIENDAS Y CORREO TELEGRÁFICO'
        if columna_buscada not in df2.columns:
            st.error(f"La columna '{columna_buscada}' no está en el archivo.")
            return None

        # Renombrar la columna
        df2 = df2.rename(columns={columna_buscada: 'CORRESPONDENCIA PRIORITARIA'})

        # Verificar que la columna fue renombrada correctamente
        st.write("Columnas en df2 después del renombrado:", df2.columns)

        # Realizar el merge con 'DIVIPOLA' como columna clave
        df_merged = df1_courier.merge(
            df2[['DIVIPOLA', 'CORRESPONDENCIA PRIORITARIA']], 
            on='DIVIPOLA', how='left'
        )

        # Eliminar las columnas duplicadas
        df_merged = df_merged.loc[:, ~df_merged.columns.duplicated()]

        return df_merged
    except Exception as e:
        st.error(f"Error al hacer el merge: {e}")
        return None


# Función para calcular los días transcurridos
def calcular_dias(df_merged):
    try:
        # Verificar si la columna 'CORRESPONDENCIA PRIORITARIA' está en el DataFrame
        if 'CORRESPONDENCIA PRIORITARIA' not in df_merged.columns:
            st.error("La columna 'CORRESPONDENCIA PRIORITARIA' no está en el DataFrame.")
            return None

        # Asegurarse de que la columna 'CORRESPONDENCIA PRIORITARIA' sea numérica
        df_merged['CORRESPONDENCIA PRIORITARIA'] = pd.to_numeric(df_merged['CORRESPONDENCIA PRIORITARIA'], errors='coerce')

        # Crear una nueva columna 'DIAS' sumando 2 a los valores de la columna 'CORRESPONDENCIA PRIORITARIA'
        # No debe ocurrir la asignación a una columna con múltiples valores
        df_merged['DIAS'] = df_merged['CORRESPONDENCIA PRIORITARIA'] + 2

        st.write("Primeras filas después de calcular los días:", df_merged[['DIVIPOLA', 'CORRESPONDENCIA PRIORITARIA', 'DIAS']].head())  # Depuración

        return df_merged
    except Exception as e:
        st.error(f"Error al calcular los días: {e}")
        return None

# Función para calcular el estado de "TERMINO"
def calcular_termino(df_merged):
    try:
        # Verificar si las columnas necesarias existen
        if 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)' not in df_merged.columns:
            st.error("La columna 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)' no está en el archivo.")
            return None
        
        if 'DIAS' not in df_merged.columns:
            st.error("La columna 'DIAS' no está en el archivo.")
            return None
        
        # Obtener las fechas festivas en Colombia
        festivos_colombia = holidays.Colombia(years=[datetime.today().year])

        # Limpiar cualquier espacio extra o caracteres no visibles
        df_merged['FECHA_DE_RECIBIDO_LIMPIA'] = df_merged['FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'].astype(str).str.strip()

        # Verificar los valores vacíos o mal formateados
        valores_vacios = df_merged[df_merged['FECHA_DE_RECIBIDO_LIMPIA'] == '']
        if not valores_vacios.empty:
            st.warning(f"Existen {len(valores_vacios)} valores vacíos en 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'. Revísalos.")

        # Función para convertir la fecha utilizando datetime.strptime() de manera manual
        def convertir_fecha(fecha_recibido):
            try:
                if pd.isna(fecha_recibido) or fecha_recibido == '':
                    return None
                # Verificar si la fecha tiene hora (longitud de la cadena)
                if len(fecha_recibido.split()) > 1:  # Si la fecha contiene espacio, tiene hora
                    fecha_convertida = datetime.strptime(fecha_recibido, '%Y-%m-%d %H:%M:%S')
                else:
                    # Si no tiene hora, solo tiene la fecha
                    fecha_convertida = datetime.strptime(fecha_recibido, '%Y-%m-%d')
                
                return fecha_convertida
            except Exception as e:
                st.warning(f"Error al convertir la fecha: {fecha_recibido}. Detalles del error: {e}")
                return None  # Si no se puede convertir, devolver None

        # Aplicar la conversión de fechas a la columna limpia
        df_merged['FECHA_DE_RECIBIDO_CONVERTIDA'] = df_merged['FECHA_DE_RECIBIDO_LIMPIA'].apply(convertir_fecha)

        # Verificar los valores que no se pudieron convertir
        fechas_invalidas = df_merged[df_merged['FECHA_DE_RECIBIDO_CONVERTIDA'].isna()]
        if not fechas_invalidas.empty:
            st.warning(f"Hay {len(fechas_invalidas)} filas con fechas no válidas. Revisa los datos de la columna 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)'.")
        
        # Función para calcular los días hábiles entre las fechas (sin numpy)
        def calcular_dias_habiles(fecha_recibido):
            if fecha_recibido is None:
                return None  # Si la fecha es inválida, retornar None
            hoy = datetime.today()  # Obtener la fecha actual
            dias_habiles = 0
            # Empezamos desde el día siguiente de la fecha recibida (solo se incrementa una vez)
            fecha_recibido += timedelta(days=1)  # Comenzar desde el siguiente día

            # Iteramos día a día entre la fecha recibida y hoy
            while fecha_recibido <= hoy:
                if fecha_recibido.weekday() < 5 and fecha_recibido not in festivos_colombia:
                    dias_habiles += 1
                fecha_recibido += timedelta(days=1)
            return dias_habiles

        # Crear la columna 'DIAS TRANSCURRIDOS' con el cálculo de los días hábiles
        df_merged['DIAS TRANSCURRIDOS'] = df_merged['FECHA_DE_RECIBIDO_CONVERTIDA'].apply(calcular_dias_habiles)

        # Verificar si hay alguna columna con valores nulos que no pudieron ser procesados
        if df_merged['DIAS TRANSCURRIDOS'].isna().sum() > 0:
            st.warning(f"Existen valores no válidos en 'FECHA DE RECIBIDO EN CORRESPONDENCIA (GESTOR DOCUMENTAL)' que no se pudieron convertir.")
        
        # Crear la columna 'TERMINO' basada en la comparación de 'DIAS TRANSCURRIDOS' con la columna 'DIAS'
        df_merged['TERMINO'] = df_merged.apply(
            lambda row: 'FUERA DE TÉRMINO' if row['DIAS TRANSCURRIDOS'] > row['DIAS'] else 'EN TÉRMINO', axis=1
        )
        df_merged = df_merged.drop(columns=['FECHA_DE_RECIBIDO_LIMPIA', 'FECHA_DE_RECIBIDO_CONVERTIDA'])
        return df_merged
    
    except Exception as e:
        st.error(f"Error al calcular el término: {e}")
        return None
    
# Función para guardar el archivo Excel generado con todas las hojas
def guardar_archivo_excel2(df_merged, xl1):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Escribir la hoja 'COURIER' procesada con el nombre 'COURIER GESTOR'
            df_merged.to_excel(writer, sheet_name='COURIER GESTOR', index=False)
            
            # Escribir todas las demás hojas del archivo original
            for sheet_name in xl1.sheet_names:
                if sheet_name != 'COURIER':  # No escribir la hoja 'COURIER' ya que ya se escribió
                    df = xl1.parse(sheet_name)  # Leer cada hoja del archivo original
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        output.seek(0)  # Reposicionar el puntero al inicio
        return output
    except Exception as e:
        st.error(f"Error al guardar el archivo: {e}")
        return None


# Descargar el archivo generado
def descargar_excel2(archivo_combinado):
    try:
        st.download_button(
            label="📥 Descargar",
            data=archivo_combinado,
            file_name="estado_gestor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error al preparar el archivo para descarga: {e}")