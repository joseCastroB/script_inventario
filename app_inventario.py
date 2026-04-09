import streamlit as st
from streamlit_sortables import sort_items
import xmlrpc.client
import base64
import io
import xlsxwriter
from PIL import Image
from datetime import datetime
import ast
import os

# Configuración visual
st.set_page_config(page_title="Exportador Odoo Dinámico", page_icon="📦", layout="wide")
st.title("📦 Exportador de Inventario Avanzado")

# --- 1. CONEXIÓN EN VIVO PARA EXTRAER FILTROS (FAVORITOS) DE ODOO ---
@st.cache_data(ttl=30) 
def obtener_filtros_odoo():
    try:
        url = 'https://omr-work-group-sac.odoo.com'
        db = 'omr-work-group-sac'
        username = 'oscar.moscoso@omrworkgroup.com'
        password = st.secrets["ODOO_PASSWORD"]
        
        common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        uid = common.authenticate(db, username, password, {})
        models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
        
        filtros_crudos = models.execute_kw(db, uid, password, 'ir.filters', 'search_read', 
            [[('model_id', '=', 'product.template')]], 
            {'fields': ['name', 'domain']})
        
        filtros_dict = {"Todos los registros (Sin filtro)": []}
        
        for f in filtros_crudos:
            nombre = f.get('name')
            dominio_str = f.get('domain')
            if nombre and dominio_str and dominio_str != '[]':
                try:
                    dominio_lista = ast.literal_eval(dominio_str)
                    filtros_dict[nombre] = dominio_lista
                except Exception:
                    pass 
                    
        return filtros_dict
    except Exception as e:
        return {"Todos los registros (Sin filtro)": []}

# --- 2. CONEXIÓN EN VIVO PARA EXTRAER CLIENTES DEL INVENTARIO ---
@st.cache_data(ttl=600) 
def obtener_clientes_odoo():
    try:
        url = 'https://omr-work-group-sac.odoo.com'
        db = 'omr-work-group-sac'
        username = 'oscar.moscoso@omrworkgroup.com'
        password = st.secrets["ODOO_PASSWORD"]
        
        common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        uid = common.authenticate(db, username, password, {})
        models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
        
        grupos_clientes = models.execute_kw(db, uid, password, 'product.template', 'read_group', 
            [[('x_studio_cliente_1', '!=', False)]], 
            ['x_studio_cliente_1'], 
            ['x_studio_cliente_1'])
        
        nombres_clientes = []
        for grupo in grupos_clientes:
            campo = grupo.get('x_studio_cliente_1')
            if isinstance(campo, list) and len(campo) == 2:
                nombres_clientes.append(campo[1])
            elif isinstance(campo, str):
                nombres_clientes.append(campo)
                
        return sorted(list(set(nombres_clientes)))
    except Exception as e:
        return []

# --- 3. DICCIONARIO DE COLUMNAS ---
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

# --- INTERFAZ VISUAL ---
st.subheader("1. Filtra los registros (Filas)")

tipo_filtro = st.radio(
    "¿Qué tipo de filtro deseas usar?", 
    ["Filtros Guardados de Odoo (Favoritos)", "Buscar por Cliente Específico en Vivo"],
    horizontal=True
)

if tipo_filtro == "Filtros Guardados de Odoo (Favoritos)":
    filtros_dinamicos = obtener_filtros_odoo()
    
    col_filtro, col_btn = st.columns([4, 1])
    with col_filtro:
        filtro_elegido = st.selectbox("Selecciona un filtro guardado:", list(filtros_dinamicos.keys()))
    with col_btn:
        st.write("") 
        st.write("")
        if st.button("🔄 Refrescar filtros"):
            obtener_filtros_odoo.clear() 
            st.rerun() 
            
    dominio_odoo = filtros_dinamicos[filtro_elegido]
    
    nombre_empresa_reporte = filtro_elegido.replace("ACTIVOS ", "").replace("TOTAL ", "").strip()
    if nombre_empresa_reporte == "Todos los registros (Sin filtro)" or nombre_empresa_reporte == "Productos":
        nombre_empresa_reporte = "GENERAL"

else:
    lista_clientes = obtener_clientes_odoo()
    if lista_clientes:
        cliente_elegido = st.selectbox("Selecciona un Cliente (Extraído del inventario actual):", lista_clientes)
    else:
        cliente_elegido = st.text_input("Escribe el nombre del Cliente (Ej: MOMENTUM):")
    
    dominio_odoo = [("x_studio_cliente_1", "ilike", cliente_elegido)]
    nombre_empresa_reporte = cliente_elegido if cliente_elegido else "GENERAL"

dominio_expandido = []
for regla in dominio_odoo:
    if isinstance(regla, (list, tuple)) and len(regla) == 3:
        campo, operador, valor = regla
        if campo == 'type' and valor == 'consu' and operador == '=':
            dominio_expandido.append(('type', 'in', ['consu', 'service']))
        else:
            dominio_expandido.append(regla)
    else:
        dominio_expandido.append(regla)
        
dominio_odoo = dominio_expandido

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("2. Elige los campos (Columnas)")
    campos_seleccionados = st.multiselect(
        "Agrega o quita campos de la lista:",
        options=list(CAMPOS_DISPONIBLES.keys()),
        default=[
            "Nombre", "Medidas", "Estado del Activo", "Categoría del producto",
            "Tipo de producto", "Stock", "Cliente", "Evento"
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

                common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
                uid = common.authenticate(db, username, password, {})
                models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

                campos_tecnicos_a_buscar = [CAMPOS_DISPONIBLES[campo] for campo in campos_ordenados]
                campos_a_consultar = campos_tecnicos_a_buscar + ['image_128'] 

                # Buscar en Odoo
                productos = models.execute_kw(db, uid, password, 
                        'product.template', 'search_read',
                        [dominio_odoo], 
                        {'fields': campos_a_consultar, 'limit': 1000}) 

                # --- ESCRITURA EN EXCEL ---
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet('Kardex')

                # Formatos
                formato_titulo_1 = workbook.add_format({'bold': True, 'font_size': 36, 'font_color': '#000000', 'valign': 'vcenter'})
                formato_titulo_2 = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#000000', 'valign': 'vcenter'})
                formato_cabecera = workbook.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                formato_normal = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                formato_stock = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00', 'bold': True})
                formato_categoria = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'bold': True})

                # --- AJUSTE MANUAL DE LOGO Y CELDA A1 ---
                escala_logo = 0.38 
                altura_fila_0 = 150 
                worksheet.set_row(0, altura_fila_0)
                
                ancho_columna_A = 35
                worksheet.set_column(0, 0, ancho_columna_A)
                
                if os.path.exists("logo.png"):
                    try:
                        worksheet.insert_image(0, 0, "logo.png", {
                            'x_scale': escala_logo, 
                            'y_scale': escala_logo, 
                            'x_offset': 10, 
                            'y_offset': 10
                        })
                    except Exception:
                        pass 
                
                # Subtítulo KARDEX
                worksheet.set_row(2, 25) 
                worksheet.write(2, 0, f"KARDEX ACTIVOS {nombre_empresa_reporte.upper()}", formato_titulo_2)

                # Cabeceras Negras
                fila_encabezados = 4
                worksheet.set_row(fila_encabezados, 30)
                
                for col_num, campo_humano in enumerate(campos_ordenados):
                    worksheet.write(fila_encabezados, col_num, campo_humano.upper(), formato_cabecera)
                    worksheet.set_column(col_num, col_num, 20) 
                
                col_imagen = len(campos_ordenados)
                worksheet.write(fila_encabezados, col_imagen, 'FOTO', formato_cabecera)
                worksheet.set_column(col_imagen, col_imagen, 26)
                
                if "Nombre" in campos_ordenados:
                    idx_nombre = campos_ordenados.index("Nombre")
                    worksheet.set_column(idx_nombre, idx_nombre, 35)

                # Llenar datos 
                row = 5
                total_stock_acumulado = 0 

                for prod in productos:
                    worksheet.set_row(row, 105) 
                    
                    for col_num, campo_tecnico in enumerate(campos_tecnicos_a_buscar):
                        campo_humano = campos_ordenados[col_num]
                        valor = prod.get(campo_tecnico, '')

                        if campo_tecnico == 'type':
                            val_str = str(valor).lower().strip()
                            if valor == 'consu':
                                valor = 'BIENES'
                            elif valor == 'service':
                                valor = 'SERVICIOS'
                            elif valor == 'product':
                                valor = 'ALMACENABLES'
                                
                        if isinstance(valor, list) and len(valor) == 2:
                            valor = valor[1]
                        elif isinstance(valor, bool):
                            valor = "Sí" if valor else "No"
                        elif valor is False or valor is None:
                            valor = ''
                            
                        if campo_humano == "Stock" and valor != '':
                            try:
                                total_stock_acumulado += float(valor)
                            except:
                                pass

                        if campo_humano == "Stock":
                            formato_usar = formato_stock
                        elif campo_humano == "Categoría del producto":
                            formato_usar = formato_categoria
                        else:
                            formato_usar = formato_normal
                            
                        worksheet.write(row, col_num, valor, formato_usar)

                    # --- LÓGICA DE CENTRADO MATEMÁTICO DE IMAGEN ---
                    worksheet.write_blank(row, col_imagen, '', formato_normal) 
                    imagen_base64 = prod.get('image_128')
                    
                    if imagen_base64:
                        try:
                            image_data = base64.b64decode(imagen_base64)
                            imagen_pil = Image.open(io.BytesIO(image_data))
                            
                            # 1. Definir tamaño de la celda de Excel en píxeles
                            # (Ancho col 26 = aprox 187px, Alto fila 105 = aprox 140px)
                            ancho_celda_px = 187 
                            alto_celda_px = 140  
                            
                            # 2. Definir margen de seguridad para no tocar los bordes
                            margen = 20
                            max_ancho = ancho_celda_px - margen
                            max_alto = alto_celda_px - margen
                            
                            # 3. Redimensionar inteligentemente sin deformar
                            imagen_pil.thumbnail((max_ancho, max_alto))
                            ancho_final, alto_final = imagen_pil.size
                            
                            # 4. Calcular el centro exacto de la celda
                            centro_x = int((ancho_celda_px - ancho_final) / 2)
                            centro_y = int((alto_celda_px - alto_final) / 2)
                            
                            stream_imagen = io.BytesIO()
                            imagen_pil.save(stream_imagen, format="PNG")
                            stream_imagen.seek(0)
                            
                            # Insertamos la imagen con escala normal (1.0) usando los offsets matemáticos
                            worksheet.insert_image(row, col_imagen, 'img.png', {
                                'image_data': stream_imagen, 
                                'x_scale': 1.0,  
                                'y_scale': 1.0,
                                'object_position': 1,
                                'x_offset': centro_x,
                                'y_offset': centro_y
                            })
                        except Exception:
                            worksheet.write(row, col_imagen, 'Error', formato_normal)
                    
                    row += 1

                # Autofiltros 
                worksheet.autofilter(fila_encabezados, 0, row - 1, col_imagen)

                # FILA FINAL DE TOTALES
                row += 1 
                worksheet.set_row(row, 25)
                worksheet.write(row, 0, "TOTALES GENERALES", formato_cabecera)
                worksheet.write(row, 1, f"Cantidad: {len(productos)} Registros", formato_categoria)
                
                if "Stock" in campos_ordenados:
                    idx_stock = campos_ordenados.index("Stock")
                    worksheet.write(row, idx_stock, total_stock_acumulado, formato_stock)

                workbook.close()
                st.success(f"¡Se exportaron {len(productos)} registros exitosamente!")
                
                fecha_actual = datetime.now().strftime("%d-%m-%Y")
                nombre_archivo = f"KARDEX_{nombre_empresa_reporte.replace(' ', '_')}_{fecha_actual}.xlsx"
                
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo}",
                    data=output.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error técnico: {e}")