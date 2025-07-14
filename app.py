# app.py
import streamlit as st
from generador_excel import generar_excel_en_memoria # Importamos la función refactorizada

# --- Configuración de la página ---
st.set_page_config(
    page_title="Generador de Lentes",
    page_icon="👁️",
    layout="centered"
)

st.title("Generador de Hojas de Cálculo para Lentes 👁️")
st.write("Introduce los parámetros de la lente para generar el archivo de cálculo de espesores.")

# --- Formulario de Entrada ---
with st.form(key='lens_form'):
    st.header("Parámetros de la Receta")
    
    col1, col2 = st.columns(2)
    with col1:
        lado_ojo = st.selectbox("Lado del Ojo", ("R", "L"), index=0, help="Selecciona el ojo derecho (R) o izquierdo (L).")
        esfera_d = st.number_input("Esfera (D)", value=-4.50, step=0.25, help="Potencia esférica en dioptrías.")
        cilindro_d = st.number_input("Cilindro (D)", value=0.00, step=0.25, help="Potencia cilíndrica en dioptrías.")
        eje_cilindro_grados = st.slider("Eje del Cilindro (°)", 0, 180, 90, help="Eje del astigmatismo.")

    with col2:
        indice_refraccion = st.number_input("Índice de Refracción", value=1.586, step=0.001, format="%.3f", help="Índice de refracción del material.")
        grosor_orilla_mm = st.number_input("Grosor de Borde Mín. (mm)", value=1.70, step=0.1, help="Grosor mínimo deseado en el borde.")
        grosor_centro_mm = st.number_input("Grosor de Centro Mín. (mm)", value=2.10, step=0.1, help="Grosor mínimo deseado en el centro.")
    
    st.header("Prisma y Decentración")
    col3, col4 = st.columns(2)
    with col3:
        decentracion_h = st.number_input("Decentración Horizontal (mm)", value=0.0, step=0.1, help="Desplazamiento del centro óptico en el eje X.")
        decentracion_v = st.number_input("Decentración Vertical (mm)", value=0.0, step=0.1, help="Desplazamiento del centro óptico en el eje Y.")

    with col4:
        prisma_magnitud = st.number_input("Magnitud del Prisma (Δ)", value=5.0, step=0.25, help="Potencia del prisma en dioptrías prismáticas.")
        prisma_eje = st.slider("Base del Prisma (°)", 0, 360, 0, help="Orientación de la base del prisma.")
        
    st.header("Radios de Borde (Obligatorio)")
    radios_str = st.text_area(
        "Introduce los radios en centésimas de mm, separados por punto y coma (;)",
        height=200,
        placeholder="2395;2404;2415;...",
        help="Pega aquí la lista de radios de borde. Puedes usar saltos de línea o puntos y coma como separadores."
    )

    submit_button = st.form_submit_button(label='✨ Generar Archivo Excel')

# --- Lógica de Procesamiento ---
if submit_button:
    if not radios_str.strip():
        st.error("¡Error! Debes introducir los radios de borde para poder generar el archivo.")
    else:
        with st.spinner('Calculando y generando el archivo... Por favor, espera.'):
            try:
                datos_lente = {
                    "lado_ojo": lado_ojo,
                    "esfera_d": esfera_d,
                    "cilindro_d": cilindro_d,
                    "eje_cilindro_grados": eje_cilindro_grados,
                    "prisma_magnitud_dp": prisma_magnitud,
                    "prisma_eje_base_grados": prisma_eje,
                    "grosor_orilla_mm": grosor_orilla_mm,
                    "grosor_centro_mm": grosor_centro_mm,
                    "indice_refraccion": indice_refraccion,
                    "radios_borde_centesimas_mm_str": radios_str.replace('\n', ';').replace(',', ';'),
                    "decentracion_co_horizontal_mm": decentracion_h,
                    "decentracion_co_vertical_mm": decentracion_v,
                }

                excel_en_memoria = generar_excel_en_memoria(datos_lente)

                st.success("¡Archivo generado con éxito!")
                st.download_button(
                    label="📥 Descargar Archivo Excel",
                    data=excel_en_memoria,
                    file_name=f"Calculo_Lente_{lado_ojo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except ValueError as ve:
                st.error(f"Error en los datos de entrada: {ve}")
            except Exception as e:
                st.error(f"Ocurrió un error inesperado: {e}")