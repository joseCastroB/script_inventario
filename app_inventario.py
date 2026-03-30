import streamlit as st
from streamlit_sortables import sort_items # <-- NUEVA LIBRERÍA
import xmlrpc.client
import base64
import io
import xlsxwriter
from PIL import Image

# 1. Configuración visual y Diccionario de Campos
st.set_page_config(page_title="Exportador Odoo Dinámico", page_icon="📦", layout="wide")
st.title("📦 Exportador de Inventario Avanzado")
st.write("Selecciona los campos y luego arrástralos para definir el orden exacto de las columnas en tu Excel.")

# DICCIONARIO COMPLETO
CAMPOS_DISPONIBLES = {
    "Favorito": "is_favorite",
    "Nombre": "name",
    "Marca": "x_studio_marca",
    "Medidas": "x_studio_medidas",
    "Estado del Activo": "x_studio_estado_del_activo",
    "Referencia interna": "default_code",
    "Precio de venta": "list_price",
    "Categoría del producto": "categ_id",
    "Costo": "standard_price",
    "Cantidad a la mano": "qty_available",
    "Cantidad pronosticada": "virtual_available",
    "Cliente": "x_studio_cliente_1",
    "Evento": "x_studio_evento",
    "Unidad": "uom_id",
    "Decoración de la actividad de excepción": "activity_exception_decoration",
    "Tipo de producto": "type"
}

# 2. Interfaz de Selección y ORDENAMIENTO (DRAG & DROP)
col1, col2 = st.columns(2) # Dividimos la pantalla en dos para que se vea más ordenado

with col1:
    st.subheader("1. Elige los campos")
    campos_seleccionados = st.multiselect(
        "Agrega o quita campos de la lista:",
        options=list(CAMPOS_DISPONIBLES.keys()),
        default=["Nombre", "Medidas", "Unidad", "Estado del Activo", "Categoría del producto", "Tipo de producto", "Cantidad a la mano", "Cliente", "Evento"] 
    )

with col2:
    st.subheader("2. Ordena las columnas")
    st.write("Arrastra los bloques arriba o abajo para ordenar el Excel:")
    
    # Aquí ocurre la magia del Drag & Drop
    if campos_seleccionados:
        # sort_items toma la lista original y devuelve la lista reordenada por el usuario
        campos_ordenados = sort_items(campos_seleccionados)
    else:
        campos_ordenados = []

st.divider()

# 3. Botón de acción principal
if st.button("Generar Excel de Inventario", type="primary"):
    
    # ATENCIÓN: Ahora validamos y usamos 'campos_ordenados' en lugar de 'campos_seleccionados'
    if not campos_ordenados:
        st.warning("Por favor, selecciona al menos un campo para exportar.")
    else:
        with st.spinner('Extrayendo datos y procesando imágenes desde Odoo...'):
            try:
                # Credenciales de Odoo
                url = 'https://cons240326.odoo.com'
                db = 'cons240326'
                username = 'oscar.moscoso@omrworkgroup.com'
                password = st.secrets["ODOO_PASSWORD"]

                # Conexión XML-RPC
                common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
                uid = common.authenticate(db, username, password, {})
                models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

                # Preparar los campos técnicos usando la LISTA ORDENADA POR EL USUARIO
                campos_tecnicos_a_buscar = [CAMPOS_DISPONIBLES[campo] for campo in campos_ordenados]
                campos_a_consultar = campos_tecnicos_a_buscar + ['image_128'] 

                # Buscar productos 
                productos = models.execute_kw(db, uid, password, 
                        'product.template', 'search_read',
                        [[]], 
                        {'fields': campos_a_consultar}) 

                # Iniciar Excel en memoria
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet('Productos')

                # Escribir Encabezados Dinámicos (Respetando el nuevo orden)
                for col_num, campo_humano in enumerate(campos_ordenados):
                    worksheet.write(0, col_num, campo_humano)
                    worksheet.set_column(col_num, col_num, 25) 
                
                # Encabezado de la Imagen (Siempre al final)
                col_imagen = len(campos_ordenados)
                worksheet.write(0, col_imagen, 'Imagen')
                worksheet.set_column(col_imagen, col_imagen, 20)

                # Llenar datos
                row = 1
                for prod in productos:
                    # Usamos la lista de técnicos ordenada
                    for col_num, campo_tecnico in enumerate(campos_tecnicos_a_buscar):
                        valor = prod.get(campo_tecnico, '')
                        
                        if isinstance(valor, list) and len(valor) == 2:
                            valor = valor[1]
                        elif isinstance(valor, bool):
                            valor = "Sí" if valor else "No"
                        elif valor is False or valor is None:
                            valor = ''
                            
                        worksheet.write(row, col_num, valor)

                    # Procesar la imagen
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
                st.success("¡Excel listo para descargar!")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Inventario.xlsx",
                    data=output.getvalue(),
                    file_name="Inventario_Odoo_Ordenado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error técnico: {e}")