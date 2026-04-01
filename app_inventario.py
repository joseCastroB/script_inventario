import streamlit as st
from streamlit_sortables import sort_items
import xmlrpc.client
import base64
import io
import xlsxwriter
from PIL import Image
from datetime import datetime

# Configuración visual
st.set_page_config(page_title="Exportador Odoo Dinámico", page_icon="📦", layout="wide")
st.title("📦 Exportador de Inventario Avanzado")

# --- CONEXIÓN CACHEADA PARA EXTRAER CLIENTES EN VIVO ---
@st.cache_data(ttl=600) # Se actualiza cada 10 minutos para no saturar Odoo
def obtener_clientes_odoo():
    try:
        url = 'https://omr-work-group-sac.odoo.com'
        db = 'omr-work-group-sac'
        username = 'oscar.moscoso@omrworkgroup.com'
        password = st.secrets["ODOO_PASSWORD"]
        
        common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        uid = common.authenticate(db, username, password, {})
        models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
        
        # OBTENER SOLO LOS CLIENTES QUE EXISTEN EN EL CAMPO x_studio_cliente_1 DE LOS PRODUCTOS
        grupos_clientes = models.execute_kw(db, uid, password, 'product.template', 'read_group', 
            [[('x_studio_cliente_1', '!=', False)]], # Filtramos para que no traiga los vacíos
            ['x_studio_cliente_1'], 
            ['x_studio_cliente_1'])
        
        nombres_clientes = []
        for grupo in grupos_clientes:
            campo = grupo.get('x_studio_cliente_1')
            if isinstance(campo, list) and len(campo) == 2:
                nombres_clientes.append(campo[1]) # Si es Many2one saca el nombre
            elif isinstance(campo, str):
                nombres_clientes.append(campo) # Si es un campo de texto simple
                
        return sorted(list(set(nombres_clientes)))
    except Exception as e:
        return []

# 1. DICCIONARIO DE COLUMNAS 
CAMPOS_DISPONIBLES = {
    "Favorito": "is_favorite",
    "Nombre": "name",
    "Marca": "x_studio_marca",
    "Medidas": "x_studio_medidas",
    "Estado del Activo": "x_studio_estado_del_activo",
    "Referencia interna": "default_code",
    "Responsable": "responsible_id",
    "Etiquetas": "product_tag_ids",
    "Cinta": "website_ribbon_id",
    "Código de barras": "barcode",
    "Precio de venta": "list_price",
    "Costo": "standard_price",
    "Categoría de producto de PdV": "pos_categ_ids",
    "Disponible en PdV": "available_in_pos",
    "Categoría del producto": "categ_id",
    "Tipo de producto": "type",
    "Stock": "qty_available",
    "Pronosticado": "virtual_available",
    "Cliente": "x_studio_cliente_1",
    "Evento": "x_studio_evento",
    "Unidad": "uom_id",
    "Decoración de la actividad de excepción": "activity_exception_decoration"
}

# 2. DICCIONARIO DE FILTROS PREDEFINIDOS
FILTROS_DISPONIBLES = {
    "Todos los registros (Sin filtro)": [],
    "ACTIVOS BCP": [("x_studio_cliente_1", "ilike", "BCP")],
    "ACTIVOS DECOPRINT": [("x_studio_cliente_1", "=", "DECOPRINT")],
    "ACTIVOS GSK": [("x_studio_cliente_1", "=", "GSK")],
    "ACTIVOS MOMENTUM": ["&", "&", ("type", "=", "consu"), ("x_studio_cliente_1", "=", "MOMENTUM"), ("qty_available", ">", 0)],
    "ACTIVOS OMR": [("x_studio_cliente_1", "=", "OMR")],
    "ACTIVOS TOTAL OMR": [("type", "=", "consu"), ("purchase_ok", "=", True)],
    "ACTIVOS UNILEVER": ["&", ("x_studio_cliente_1", "=", "UNILEVER"), ("qty_available", ">", 0)],
    "Productos": [("x_studio_cliente_1", "=", "UNILEVER")]
}

# --- INTERFAZ VISUAL ---
st.subheader("1. Filtra los registros (Filas)")

# Selector de tipo de filtro
tipo_filtro = st.radio(
    "¿Qué tipo de filtro deseas usar?", 
    ["Filtros Predefinidos (Casuísticas)", "Buscar por Cliente Específico en Vivo"],
    horizontal=True
)

if tipo_filtro == "Filtros Predefinidos (Casuísticas)":
    filtro_elegido = st.selectbox("Selecciona la casuística a extraer:", list(FILTROS_DISPONIBLES.keys()))
    dominio_odoo = FILTROS_DISPONIBLES[filtro_elegido]
    
    # Inteligencia para extraer el nombre limpio de la empresa desde el filtro predefinido
    nombre_empresa_reporte = filtro_elegido.replace("ACTIVOS ", "").replace("TOTAL ", "").strip()
    if nombre_empresa_reporte == "Todos los registros (Sin filtro)" or nombre_empresa_reporte == "Productos":
        nombre_empresa_reporte = "GENERAL"

else:
    lista_clientes = obtener_clientes_odoo()
    if lista_clientes:
        cliente_elegido = st.selectbox("Selecciona un Cliente de la base de datos de Odoo:", lista_clientes)
    else:
        # Respaldo de seguridad por si falla la conexión inicial
        cliente_elegido = st.text_input("Escribe el nombre del Cliente (Ej: MOMENTUM):")
    
    dominio_odoo = [("x_studio_cliente_1", "ilike", cliente_elegido)]
    nombre_empresa_reporte = cliente_elegido if cliente_elegido else "GENERAL"

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("2. Elige los campos (Columnas)")
    campos_seleccionados = st.multiselect(
        "Agrega o quita campos de la lista:",
        options=list(CAMPOS_DISPONIBLES.keys()),
        # Ajustado para que se parezca a tu imagen por defecto
        default=[
            "Nombre", "Marca", "Medidas", "Estado del Activo", 
            "Referencia interna","Precio de venta", "Costo" ,
            "Categoría del producto", "Pronosticado",
            "Tipo de producto", "Stock", "Cliente", "Unidad", "Evento"
        ] 
    )

with col2:
    st.subheader("3. Ordena las columnas")
    st.write("Arrastra los bloques para cambiar el orden:")
    if campos_seleccionados:
        campos_ordenados = sort_items(campos_seleccionados)
    else:
        campos_ordenados = []

st.divider()

# --- LÓGICA DE EXPORTACIÓN ---
if st.button("Generar Reporte Corporativo", type="primary"):
    
    if not campos_ordenados:
        st.warning("Por favor, selecciona al menos un campo para exportar.")
    else:
        with st.spinner('Extrayendo datos filtrados y aplicando diseño corporativo...'):
            try:
                # Credenciales 
                url = 'https://omr-work-group-sac.odoo.com'
                db = 'omr-work-group-sac'
                username = 'oscar.moscoso@omrworkgroup.com'
                password = st.secrets["ODOO_PASSWORD"]

                # Conexión XML-RPC
                common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
                uid = common.authenticate(db, username, password, {})
                models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

                # Preparar búsqueda
                campos_tecnicos_a_buscar = [CAMPOS_DISPONIBLES[campo] for campo in campos_ordenados]
                campos_a_consultar = campos_tecnicos_a_buscar + ['image_128'] 

                # Buscar en Odoo
                productos = models.execute_kw(db, uid, password, 
                        'product.template', 'search_read',
                        [dominio_odoo], 
                        {'fields': campos_a_consultar, 'limit': 300}) 

                # --- ESCRITURA EN EXCEL (CON DISEÑO AVANZADO) ---
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet('Kardex')

                # 1. Definición de Formatos
                formato_titulo_1 = workbook.add_format({'bold': True, 'font_size': 36, 'font_color': '#000000', 'valign': 'vcenter'})
                formato_titulo_2 = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#000000', 'valign': 'vcenter'})
                formato_cabecera = workbook.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                
                # Formatos de celdas de datos
                formato_normal = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                formato_stock = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00', 'bold': True})
                formato_categoria = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'bold': True})

                # 2. Escribir Títulos Superiores
                worksheet.set_row(0, 45) # Fila gigante para el nombre de la empresa
                worksheet.write(0, 0, nombre_empresa_reporte.lower(), formato_titulo_1)
                
                worksheet.set_row(2, 25) # Fila para el subtítulo
                worksheet.write(2, 0, f"KARDEX {nombre_empresa_reporte.upper()}", formato_titulo_2)

                # 3. Escribir Cabeceras Negras (Fila 4)
                fila_encabezados = 4
                worksheet.set_row(fila_encabezados, 30)
                
                for col_num, campo_humano in enumerate(campos_ordenados):
                    worksheet.write(fila_encabezados, col_num, campo_humano.upper(), formato_cabecera)
                    worksheet.set_column(col_num, col_num, 20) 
                
                col_imagen = len(campos_ordenados)
                worksheet.write(fila_encabezados, col_imagen, 'FOTO', formato_cabecera) # Cambiado a FOTO
                worksheet.set_column(col_imagen, col_imagen, 26)
                
                # Expandir columna de nombre para que se vea mejor
                if "Nombre" in campos_ordenados:
                    idx_nombre = campos_ordenados.index("Nombre")
                    worksheet.set_column(idx_nombre, idx_nombre, 35)

                # 4. Llenar datos con colores por columna
                row = 5
                for prod in productos:
                    worksheet.set_row(row, 105) # Fila alta para la imagen
                    
                    for col_num, campo_tecnico in enumerate(campos_tecnicos_a_buscar):
                        campo_humano = campos_ordenados[col_num]
                        
                        # Decidir qué formato pintar según la columna
                        if campo_humano == "Stock":
                            formato_usar = formato_stock
                        elif campo_humano == "Categoría del producto":
                            formato_usar = formato_categoria
                        else:
                            formato_usar = formato_normal

                        # Extraer y limpiar el valor
                        valor = prod.get(campo_tecnico, '')
                        if isinstance(valor, list) and len(valor) == 2:
                            valor = valor[1]
                        elif isinstance(valor, bool):
                            valor = "Sí" if valor else "No"
                        elif valor is False or valor is None:
                            valor = ''
                            
                        worksheet.write(row, col_num, valor, formato_usar)

                    # Procesar imagen
                    worksheet.write_blank(row, col_imagen, '', formato_normal) # Poner borde a la celda de la imagen
                    imagen_base64 = prod.get('image_128')
                    
                    if imagen_base64:
                        try:
                            image_data = base64.b64decode(imagen_base64)
                            imagen_pil = Image.open(io.BytesIO(image_data))
                            stream_imagen = io.BytesIO()
                            imagen_pil.save(stream_imagen, format="PNG")
                            stream_imagen.seek(0)
                            
                            worksheet.insert_image(row, col_imagen, 'img.png', {
                                'image_data': stream_imagen, 
                                'x_scale': 1.04,  
                                'y_scale': 1.04,
                                'object_position': 1,
                                'x_offset': 5, # Pequeño margen
                                'y_offset': 5
                            })
                        except Exception:
                            worksheet.write(row, col_imagen, 'Error', formato_normal)
                    
                    row += 1

                # 5. Aplicar los Autofiltros en la tabla (A partir de la fila 4)
                worksheet.autofilter(fila_encabezados, 0, row - 1, col_imagen)

                workbook.close()
                st.success(f"¡Se exportaron {len(productos)} registros exitosamente con el nuevo diseño!")
                
                # Descarga
                fecha_actual = datetime.now().strftime("%d-%m-%Y") # Formato: 01-04-2026
                nombre_archivo = f"KARDEX_{nombre_empresa_reporte.replace(' ', '_')}_{fecha_actual}.xlsx"
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo}",
                    data=output.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error técnico: {e}")