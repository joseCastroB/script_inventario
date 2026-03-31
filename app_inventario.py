import streamlit as st
from streamlit_sortables import sort_items
import xmlrpc.client
import base64
import io
import xlsxwriter
from PIL import Image

# Configuración visual
st.set_page_config(page_title="Exportador Odoo Dinámico", page_icon="📦", layout="wide")
st.title("📦 Exportador de Inventario Avanzado")

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

# 2. DICCIONARIO DE FILTROS (NUEVO)
# Aquí puedes agregar todos los escenarios (casuísticas) que tu empresa necesite.
# Usamos 'ilike' en lugar de '=' para que Odoo busque el texto "BCP" dentro del campo Cliente.
FILTROS_DISPONIBLES = {
    "Todos los registros (Sin filtro)": [],
    "ACTIVOS BCP": [("x_studio_cliente_1", "ilike", "BCP")],
    "ACTIVOS DECOPRINT": [("x_studio_cliente_1", "=", "DECOPRINT")],
    "ACTIVOS GSK": [("x_studio_cliente_1", "=", "GSK")],
    "ACTIVOS MOMENTUM": ["&", "&", ("type", "=", "consu"), ("x_studio_cliente_1", "=", "MOMENTUM"), ("qty_available", ">", 0)],
    "ACTIVOS OMR": [("x_studio_cliente_1", "=", "OMR")],
    "ACTIVOS TOTAL OMR": ["&", ("type", "=", "consu"), "&", ("purchase_ok", "=", True), "|", ("can_be_expensed", "=", False), ("can_be_expensed", "!=", False)],
    "ACTIVOS UNILEVER": ["&", ("x_studio_cliente_1", "=", "UNILEVER"), ("qty_available", ">", 0)],
    "Productos": [("x_studio_cliente_1", "=", "UNILEVER")]
}

# --- INTERFAZ VISUAL ---

# Sección 1: Filtro de filas
st.subheader("1. Filtra los registros (Filas)")
filtro_elegido = st.selectbox("Selecciona qué registros quieres extraer de Odoo:", list(FILTROS_DISPONIBLES.keys()))
dominio_odoo = FILTROS_DISPONIBLES[filtro_elegido]

st.divider()

# Sección 2 y 3: Selección y orden de columnas
col1, col2 = st.columns(2)

with col1:
    st.subheader("2. Elige los campos (Columnas)")
    campos_seleccionados = st.multiselect(
        "Agrega o quita campos de la lista:",
        options=list(CAMPOS_DISPONIBLES.keys()),
        default=["Nombre", 
                    "Marca", 
                    "Medidas", 
                    "Estado del Activo", 
                    "Referencia interna", 
                    "Precio de venta", 
                    "Costo", 
                    "Stock", 
                    "Pronosticado", 
                    "Cliente", 
                    "Evento", 
                    "Unidad"
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
if st.button("Generar Excel de Inventario", type="primary"):
    
    if not campos_ordenados:
        st.warning("Por favor, selecciona al menos un campo para exportar.")
    else:
        with st.spinner('Extrayendo datos filtrados y procesando imágenes...'):
            try:
                # Credenciales 
                url = 'https://cons240326.odoo.com'
                db = 'cons240326'
                username = 'oscar.moscoso@omrworkgroup.com'
                password = 'Cons290705$$'


                # Conexión XML-RPC
                common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
                uid = common.authenticate(db, username, password, {})
                models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

                # Preparar búsqueda
                campos_tecnicos_a_buscar = [CAMPOS_DISPONIBLES[campo] for campo in campos_ordenados]
                campos_a_consultar = campos_tecnicos_a_buscar + ['image_128'] 

                # Buscar con el filtro aplicado
                productos = models.execute_kw(db, uid, password, 
                        'product.template', 'search_read',
                        [dominio_odoo], 
                        {'fields': campos_a_consultar}) 

                # Iniciar Excel en memoria
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet('Productos')

                # --- NUEVO: Título del filtro en el Excel ---
                formato_titulo = workbook.add_format({'bold': True, 'font_size': 12, 'color': '#333333'})
                worksheet.write(0, 0, f"Filtro aplicado: {filtro_elegido}", formato_titulo)

                # Escribir Encabezados (Ahora empiezan en la fila 2, dejando una en blanco)
                fila_encabezados = 2
                for col_num, campo_humano in enumerate(campos_ordenados):
                    worksheet.write(fila_encabezados, col_num, campo_humano)
                    worksheet.set_column(col_num, col_num, 25) 
                
                # Encabezado Imagen
                col_imagen = len(campos_ordenados)
                worksheet.write(fila_encabezados, col_imagen, 'Imagen')
                worksheet.set_column(col_imagen, col_imagen, 20)

                # Llenar datos (Ahora empieza en la fila 3)
                row = 3
                for prod in productos:
                    for col_num, campo_tecnico in enumerate(campos_tecnicos_a_buscar):
                        valor = prod.get(campo_tecnico, '')
                        
                        if isinstance(valor, list) and len(valor) == 2:
                            valor = valor[1]
                        elif isinstance(valor, bool):
                            valor = "Sí" if valor else "No"
                        elif valor is False or valor is None:
                            valor = ''
                            
                        worksheet.write(row, col_num, valor)

                    # Procesar imagen
                    imagen_base64 = prod.get('image_128')
                    worksheet.set_row(row, 80) 
                    if imagen_base64:
                        try:
                            image_data = base64.b64decode(imagen_base64)
                            imagen_pil = Image.open(io.BytesIO(image_data))
                            stream_imagen = io.BytesIO()
                            imagen_pil.save(stream_imagen, format="PNG")
                            stream_imagen.seek(0)
                            
                            worksheet.insert_image(row, col_imagen, 'img.png', {
                                'image_data': stream_imagen, 
                                'x_scale': 0.8, 
                                'y_scale': 0.8,
                                'object_position': 1 
                            })
                        except Exception:
                            worksheet.write(row, col_imagen, 'Error')
                    else:
                        worksheet.write(row, col_imagen, 'Sin foto')
                    
                    row += 1

                workbook.close()
                st.success(f"¡Se exportaron {len(productos)} registros exitosamente!")
                
                # --- NUEVO: Nombre dinámico del archivo ---
                # Reemplazamos espacios por guiones bajos para que el nombre del archivo sea limpio
                nombre_archivo = f"Inventario_{filtro_elegido.replace(' ', '_')}.xlsx"
                
                # Descarga
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo}",
                    data=output.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error técnico: {e}")