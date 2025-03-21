import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import re

# Configuración de la página
st.set_page_config(
    page_title="Resultados encuestas de Satisfacción",
    page_icon="📊",
    layout="wide"
)

# Extraer país del nombre de la hoja
def extract_country_from_sheet_name(sheet_name):
    # Solo mantener los códigos de país que realmente usamos
    country_codes = {
        'BRA': 'Brasil',
        'ARG': 'Argentina',
        'MEX': 'México',
        'CHL': 'Chile',
        'COL': 'Colombia'
    }
    
    for code in country_codes:
        if code in sheet_name:
            return code
    
    return "Otro"

# Extraer mes del nombre de la hoja
def extract_month_from_sheet_name(sheet_name):
    # Lista de meses en español
    months = {
        "enero": "Enero",
        "febrero": "Febrero",
        "marzo": "Marzo",
        "abril": "Abril",
        "mayo": "Mayo",
        "junio": "Junio",
        "julio": "Julio",
        "agosto": "Agosto",
        "septiembre": "Septiembre",
        "octubre": "Octubre",
        "noviembre": "Noviembre",
        "diciembre": "Diciembre"
    }
    
    # Buscar el mes en el nombre de la hoja
    sheet_name_lower = sheet_name.lower()
    for month_name, display_name in months.items():
        if month_name in sheet_name_lower:
            return display_name
    
    # Si no encontramos un mes por nombre, buscamos una fecha en formato DD/MM o similar
    date_patterns = [
        r"(\d{1,2})[\s/\-\.]+(\d{1,2})",  # 21/01, 21-01, 21.01
        r"(\d{1,2})[\s]+de[\s]+(\w+)"     # 21 de enero
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, sheet_name_lower)
        if matches:
            # Intentar extraer mes de la fecha
            try:
                if len(matches[0]) > 1:
                    # Si el patrón capturo grupos, el segundo puede ser el mes
                    month_part = matches[0][1]
                    # Si es numérico, convertir a nombre de mes
                    if month_part.isdigit():
                        month_num = int(month_part)
                        if 1 <= month_num <= 12:
                            month_names = list(months.values())
                            return month_names[month_num - 1]
                    # Si es texto, ver si coincide con un mes
                    elif month_part in months:
                        return months[month_part]
            except:
                pass
    
    return "Sin mes específico"

# Función para verificar si una hoja tiene formato de taller
def is_workshop_sheet(df, sheet_name):
    # Excluir explícitamente la hoja "Plantilla Base"
    if sheet_name == "Plantilla Base":
        return False
    
    # Verificar si contiene al menos uno de los encabezados clave
    has_resultados = any(df.iloc[:, 0].astype(str).str.contains("Resultados encuesta", case=False))
    has_facilitadores = any(df.iloc[:, 0].astype(str).str.contains("Facilitadores", case=False))
    has_fishbowl = any(df.iloc[:, 0].astype(str).str.contains("Fishbowl", case=False))
    has_verbatims = False
    
    # Buscar VERBATIMS en cualquier columna
    for col in range(len(df.columns)):
        if any(df.iloc[:, col].astype(str).str.contains("VERBATIMS", case=False)):
            has_verbatims = True
            break
    
    # Es una hoja de taller si tiene al menos dos de los elementos
    criteria_met = [has_resultados, has_facilitadores, has_fishbowl, has_verbatims]
    return sum(criteria_met) >= 2

# Crear un diccionario con todos los talleres y sus facilitadores
def create_facilitator_index(xls, workshop_sheets):
    facilitator_index = {}
    axialent_facilitators = set()  # Conjunto para almacenar solo facilitadores de Axialent
    
    for sheet_name in workshop_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            taller_data = extract_data_from_sheet(df)
            
            if not taller_data["facilitadores"].empty:
                for _, row in taller_data["facilitadores"].iterrows():
                    facilitator = row["Nombre"]
                    empresa = row.get("Empresa", "").strip()
                    
                    # Registrar al facilitador en el índice general
                    if facilitator not in facilitator_index:
                        facilitator_index[facilitator] = []
                    if sheet_name not in facilitator_index[facilitator]:
                        facilitator_index[facilitator].append(sheet_name)
                    
                    # Si es de Axialent, añadirlo al conjunto de facilitadores de Axialent
                    if empresa.lower() == "axialent":
                        axialent_facilitators.add(facilitator)
        except Exception as e:
            print(f"Error al procesar facilitadores en {sheet_name}: {str(e)}")
    
    return facilitator_index, axialent_facilitators

# Función simplificada para extraer datos con posiciones fijas
def extract_data_from_sheet(df):
    results = {
        "resultados_encuesta": pd.DataFrame(),
        "facilitadores": pd.DataFrame(),
        "fishbowl": pd.DataFrame(),
        "verbatims": []
    }
    
    try:
        # 1. EXTRAER RESULTADOS DE ENCUESTA (posiciones fijas)
        # Verificamos primero si es el formato de encuesta con % o no
        resultados_data = {
            "Métrica": ["Favorabilidad", "Aplicabilidad", "Response Rate"],
            "Valor": ["", "", ""]
        }
        
        # Verificamos la columna A para encontrar las filas correctas
        fav_row = None
        aplic_row = None
        resp_row = None
        
        for i in range(min(10, len(df))):  # Buscamos en las primeras 10 filas
            if isinstance(df.iloc[i, 0], str):
                cell_value = df.iloc[i, 0].strip().lower()
                if "favorabilidad" in cell_value:
                    fav_row = i
                elif "aplicabilidad" in cell_value:
                    aplic_row = i
                elif "response rate" in cell_value:
                    resp_row = i
        
        # Si encontramos las filas, extraemos los valores
        if fav_row is not None:
            # Convertir a formato de porcentaje
            val = df.iloc[fav_row, 1]
            if isinstance(val, (int, float)):
                resultados_data["Valor"][0] = f"{int(val*100) if val*100 == int(val*100) else val*100}%"
            else:
                resultados_data["Valor"][0] = f"{val}%"
        
        if aplic_row is not None:
            # Convertir a formato de porcentaje
            val = df.iloc[aplic_row, 1]
            if isinstance(val, (int, float)):
                resultados_data["Valor"][1] = f"{int(val*100) if val*100 == int(val*100) else val*100}%"
            else:
                resultados_data["Valor"][1] = f"{val}%"
        
        if resp_row is not None:
            # Convertir a formato de porcentaje
            val = df.iloc[resp_row, 1]
            if isinstance(val, (int, float)):
                resultados_data["Valor"][2] = f"{int(val*100) if val*100 == int(val*100) else val*100}%"
            else:
                resultados_data["Valor"][2] = f"{val}%"
            
        # Crear el dataframe
        results["resultados_encuesta"] = pd.DataFrame(resultados_data)
        
    except Exception as e:
        st.error(f"Error al extraer resultados de encuesta: {str(e)}")

    try:
        # 2. EXTRAER FACILITADORES
        facilitadores_row = None
        
        # Buscar la fila que contiene "Facilitadores"
        for i in range(len(df)):
            if isinstance(df.iloc[i, 0], str) and "facilitadores" in df.iloc[i, 0].lower():
                facilitadores_row = i
                break
        
        if facilitadores_row is not None:
            # Buscar el inicio de la siguiente sección o final de datos
            next_section_row = len(df)
            for i in range(facilitadores_row + 1, len(df)):
                if isinstance(df.iloc[i, 0], str) and ("fishbowl" in df.iloc[i, 0].lower() or df.iloc[i, 0].strip() == ""):
                    next_section_row = i
                    break
            
            # Extraer los datos de facilitadores (columnas A y B)
            facilitadores_data = []
            for i in range(facilitadores_row + 1, next_section_row):
                nombre = df.iloc[i, 0]
                empresa = df.iloc[i, 1]  # Cambiamos "Rol" por "Empresa"
                
                # Solo incluir filas con datos válidos
                if pd.notna(nombre) and str(nombre).strip() and nombre != "None":
                    facilitadores_data.append({"Nombre": nombre, "Empresa": empresa})  # Cambiamos "Rol" por "Empresa"
            
            results["facilitadores"] = pd.DataFrame(facilitadores_data)
    except Exception as e:
        st.error(f"Error al extraer facilitadores: {str(e)}")

    try:
        # 3. EXTRAER FISHBOWL (solo nombres)
        fishbowl_row = None
        
        # Buscar la fila que contiene "Fishbowl"
        for i in range(len(df)):
            if isinstance(df.iloc[i, 0], str) and "fishbowl" in df.iloc[i, 0].lower():
                fishbowl_row = i
                break
        
        if fishbowl_row is not None:
            # Buscar el final de la sección o filas vacías
            next_section_row = len(df)
            for i in range(fishbowl_row + 1, len(df)):
                # Verificar si es el final de los datos o el inicio de otra sección
                if (pd.isna(df.iloc[i, 0]) or df.iloc[i, 0] == "" or 
                    (isinstance(df.iloc[i, 0], str) and len(df.iloc[i, 0].strip()) == 0)):
                    # Verificar si hay varias filas vacías consecutivas
                    empty_count = 0
                    for j in range(i, min(i+3, len(df))):
                        if pd.isna(df.iloc[j, 0]) or df.iloc[j, 0] == "":
                            empty_count += 1
                        else:
                            break
                    if empty_count >= 2:  # Si hay 2+ filas vacías consecutivas
                        next_section_row = i
                        break
            
            # Extraer solo los nombres
            fishbowl_data = []
            for i in range(fishbowl_row + 1, next_section_row):
                nombre = df.iloc[i, 0]
                
                # Solo incluir filas con datos válidos
                if pd.notna(nombre) and str(nombre).strip() and nombre != "None":
                    fishbowl_data.append({"Nombre": nombre})
            
            results["fishbowl"] = pd.DataFrame(fishbowl_data)
    except Exception as e:
        st.error(f"Error al extraer fishbowl: {str(e)}")

    try:
        # 4. EXTRAER VERBATIMS - SIEMPRE en columna D (índice 3)
        # No buscamos la palabra VERBATIMS, simplemente extraemos todo de la columna D
        verbatims_col = 3  # Corresponde a la columna D
        
        # Extraer todos los valores no vacíos de la columna D, excluyendo la palabra "VERBATIMS"
        verbatims = []
        for i in range(len(df)):
            if i < len(df) and verbatims_col < len(df.columns):
                cell_value = df.iloc[i, verbatims_col]
                if pd.notna(cell_value):
                    value_str = str(cell_value).strip()
                    if value_str and value_str != "VERBATIMS" and "verbatims" not in value_str.lower():
                        verbatims.append(value_str)
        
        results["verbatims"] = verbatims
    except Exception as e:
        st.error(f"Error al extraer verbatims: {str(e)}")
        
    return results

# Título principal
st.title("📊 Resultados encuestas de Satisfacción")

# Cargar archivo Excel
uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Mostrar mensaje de carga
        with st.spinner("Analizando archivo Excel..."):
            # Leer el archivo Excel completo para obtener los nombres de las hojas
            xls = pd.ExcelFile(uploaded_file)
            all_worksheet_names = xls.sheet_names
            
            # Filtrar solo las hojas que tienen formato de taller
            workshop_sheets = []
            workshop_countries = set()
            workshop_months = set()
            
            for sheet_name in all_worksheet_names:
                try:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                    if is_workshop_sheet(df, sheet_name):  # Pasamos también el nombre de la hoja
                        workshop_sheets.append(sheet_name)
                        # Extraer el país
                        country = extract_country_from_sheet_name(sheet_name)
                        workshop_countries.add(country)
                        # Extraer el mes
                        month = extract_month_from_sheet_name(sheet_name)
                        workshop_months.add(month)
                except Exception as e:
                    st.warning(f"Error al analizar la hoja '{sheet_name}': {e}")
            
            if not workshop_sheets:
                st.warning("No se encontraron hojas con formato de taller en el archivo")
            else:
                st.success(f"Se encontraron {len(workshop_sheets)} hojas con formato de taller")
                
            # Crear índice de facilitadores
            facilitator_index, axialent_facilitators = create_facilitator_index(xls, workshop_sheets)
            
            # Convertir a lista ordenada
            workshop_countries = sorted(list(workshop_countries))
            workshop_months = sorted(list(workshop_months), 
                                  key=lambda x: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", 
                                              "Junio", "Julio", "Agosto", "Septiembre", 
                                              "Octubre", "Noviembre", "Diciembre", 
                                              "Sin mes específico"].index(x) 
                                         if x in ["Enero", "Febrero", "Marzo", "Abril", "Mayo", 
                                               "Junio", "Julio", "Agosto", "Septiembre", 
                                               "Octubre", "Noviembre", "Diciembre", 
                                               "Sin mes específico"] else 999)
            all_facilitators = sorted(list(facilitator_index.keys()))
            axialent_facilitators = sorted(list(axialent_facilitators))
        
        # FILTROS
        with st.expander("Filtros", expanded=True):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Filtro por país - sin valor por defecto
                selected_countries = st.multiselect(
                    "Filtrar por país:",
                    ["Todos"] + workshop_countries
                )
            
            with col2:
                # Filtro por mes - sin valor por defecto
                selected_months = st.multiselect(
                    "Filtrar por mes:",
                    ["Todos"] + workshop_months
                )
            
            with col3:
                # Filtro por facilitador - mostrar solo facilitadores de Axialent
                selected_facilitator = st.selectbox(
                    "Filtrar por facilitador:",
                    ["Todos"] + axialent_facilitators
                )
            
            # Aplicar filtros
            # 1. Filtro por país
            if "Todos" in selected_countries or not selected_countries:
                filtered_by_country = workshop_sheets
            else:
                filtered_by_country = [
                    sheet for sheet in workshop_sheets 
                    if extract_country_from_sheet_name(sheet) in selected_countries
                ]
            
            # 2. Filtro por mes
            if "Todos" in selected_months or not selected_months:
                filtered_by_country_month = filtered_by_country
            else:
                filtered_by_country_month = [
                    sheet for sheet in filtered_by_country
                    if extract_month_from_sheet_name(sheet) in selected_months
                ]
            
            # 3. Filtro por facilitador
            if selected_facilitator == "Todos":
                filtered_worksheets = filtered_by_country_month
            else:
                # Intersección de talleres filtrados por país+mes y los del facilitador seleccionado
                facilitator_workshops = facilitator_index.get(selected_facilitator, [])
                filtered_worksheets = [
                    sheet for sheet in filtered_by_country_month 
                    if sheet in facilitator_workshops
                ]
                
            if not filtered_worksheets:
                st.warning("No hay talleres que coincidan con los filtros seleccionados")
                
            # Mostrar el número de talleres filtrados
            st.info(f"Mostrando {len(filtered_worksheets)} de {len(workshop_sheets)} talleres")
        
        # Mostrar todos los talleres filtrados directamente (sin opción de selección manual)
        selected_worksheets = filtered_worksheets
        
        # Mostrar datos por taller
        if selected_worksheets:
            # Botón para expandir/colapsar todos los talleres
            expander_col1, expander_col2 = st.columns([1, 5])
            with expander_col1:
                if "expand_all" not in st.session_state:
                    st.session_state.expand_all = True
                
                if st.button("Expandir/Colapsar Todos"):
                    st.session_state.expand_all = not st.session_state.expand_all
            
            with expander_col2:
                st.write(f"Estado actual: {'Expandidos' if st.session_state.expand_all else 'Colapsados'}")
            
            st.header("Resumen de Talleres")
            
            # Crear organización de 3 columnas para las tarjetas
            cols = st.columns(3)
            
            # Procesar cada taller seleccionado
            for i, ws_name in enumerate(selected_worksheets):
                with cols[i % 3]:
                    # Crear tarjeta expandible para cada taller
                    with st.expander(f"📘 {ws_name}", expanded=st.session_state.expand_all):
                        # Leer la hoja específica
                        df = pd.read_excel(uploaded_file, sheet_name=ws_name)
                        
                        # Extraer todos los datos de la hoja
                        taller_data = extract_data_from_sheet(df)
                        
                        # 1. MOSTRAR RESULTADOS DE ENCUESTA
                        st.subheader("Resultados de Encuesta")
                        if not taller_data["resultados_encuesta"].empty:
                            # Mostrar como tabla sin crear gráfico
                            st.dataframe(taller_data["resultados_encuesta"], hide_index=True)
                        else:
                            st.info("No se encontraron datos de resultados de encuesta")
                        
                        # 2. MOSTRAR FACILITADORES
                        st.subheader("Facilitadores")
                        if not taller_data["facilitadores"].empty:
                            # Filtrar de nuevo para asegurarnos que no hay filas con None
                            facilitadores_cleaned = taller_data["facilitadores"].copy()
                            facilitadores_cleaned = facilitadores_cleaned[
                                facilitadores_cleaned["Nombre"].astype(str).str.lower() != "none"
                            ]
                            # Mostrar tabla sin índices
                            st.dataframe(facilitadores_cleaned, hide_index=True)
                        else:
                            st.info("No se encontraron datos de facilitadores")
                        
                        # 3. MOSTRAR FISHBOWL
                        st.subheader("Fishbowl")
                        if not taller_data["fishbowl"].empty:
                            # Solo mostrar nombre
                            st.dataframe(taller_data["fishbowl"], hide_index=True)
                        else:
                            st.info("No se encontraron datos de fishbowl")
                        
                        # 4. MOSTRAR VERBATIMS
                        st.subheader("Verbatims")
                        if taller_data["verbatims"]:
                            # Mostrar verbatims en una lista con numeración
                            for j, verbatim in enumerate(taller_data["verbatims"]):
                                st.markdown(f"**{j+1}.** {verbatim}")
                        else:
                            st.info("No se encontraron verbatims")
    
    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        st.info("Detalles técnicos del error para ayudar a diagnosticar el problema:")
        st.code(str(e))

# Sección de configuración
with st.sidebar:
    st.header("Configuración")
    
    st.subheader("Acerca de la App")
    st.write("""
    Esta aplicación visualiza los Resultados encuestas de Satisfacción extraídos de archivos Excel.
    Muestra resultados de encuesta, información de facilitadores, fishbowl y verbatims.
    """)
    
    st.subheader("Instrucciones")
    st.write("""
    1. Carga el archivo Excel que contiene los datos de los talleres
    2. Utiliza los filtros para seleccionar talleres por país o facilitador
    3. Explora los datos en las tarjetas expandibles
    """)
    
    st.info("""
    Nota: Si encuentras problemas con la visualización, asegúrate de que:
    1. El archivo Excel está en el formato esperado
    2. Los nombres de las secciones son correctos (Resultados encuesta, Facilitadores, Fishbowl, VERBATIMS)
    3. La estructura de datos es coherente entre hojas
    """)