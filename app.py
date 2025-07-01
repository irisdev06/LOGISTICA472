import streamlit as st
from views.proceso1 import subir_archivos, procesar_datos, guardar_archivo_excel, descargar_excel
from views.proceso2 import leer_archivos, hacer_merge, calcular_dias, guardar_archivo_excel2, calcular_termino, descargar_excel2

# Título y menú de navegación
st.title("📦 Logística 472")
st.sidebar.title("Menú de Navegación")
menu = st.sidebar.selectbox("Selecciona una opción", ["Consolidación Base", "Estado Gestor"])

if menu == "Consolidación Base":
    st.title("Consolidación Base")
    # Subir los archivos
    archivo1, archivo2 = subir_archivos()

    if archivo1 and archivo2:
        # Procesar los datos
        df_courier, df_mensajero = procesar_datos(archivo1, archivo2)

        if df_courier is not None and df_mensajero is not None:
            st.write("✅ Archivos combinados y procesados correctamente.")

            # Guardar el archivo Excel generado
            archivo_combinado = guardar_archivo_excel(df_courier, df_mensajero)

            if archivo_combinado:
                # Llamar a la función para descargar el archivo
                descargar_excel(archivo_combinado)
            else:
                st.write("❌ Hubo un problema al guardar el archivo.")
        else:
            st.write("❌ Hubo un problema al procesar los datos.")
    else:
        st.write("⚠️ Por favor, sube ambos archivos.")


elif menu == "Estado Gestor":
    st.title("Estado Gestor")
    # Flujo para "Estado Gestor"
    archivo1, archivo2 = subir_archivos()

    if archivo1 and archivo2:
        # Leer los archivos
        df1_courier, df2_courier = leer_archivos(archivo1, archivo2)

        if df1_courier is not None and df2_courier is not None:
            # Hacer el merge de los archivos
            df_merged = hacer_merge(df1_courier, df2_courier)

            if df_merged is not None:
                # Calcular los días
                df_merged = calcular_dias(df_merged)

                if df_merged is not None:
                    # Calcular el término basado en los días y otros valores
                    df_merged = calcular_termino(df_merged)

                    if df_merged is not None:
                        # Guardar el archivo Excel generado (solo con df_merged)
                        archivo_combinado = guardar_archivo_excel2(df_merged, xl1=df1_courier)  # Solo pasar df_merged aquí


                        if archivo_combinado:
                            # Descargar el archivo generado
                            descargar_excel2(archivo_combinado)
                        else:
                            st.write("❌ Hubo un problema al guardar el archivo.")
                    else:
                        st.write("❌ Hubo un problema al calcular el término.")
                else:
                    st.write("❌ Hubo un problema al calcular los días.")
            else:
                st.write("❌ Hubo un problema al hacer el merge de los archivos.")
        else:
            st.write("❌ Hubo un problema al leer los archivos.")
    else:
        st.write("⚠️ Por favor, sube ambos archivos.")
