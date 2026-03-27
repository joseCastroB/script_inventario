import streamlit as st
import xmlrpc.client
import base64
import io
import xlsxwriter
from PIL import Image

# 1. Configuración visual de la página web
st.set_page_config(page_title="Exportador Odoo", page_icon="📦")
st.title("📦 Exportador de Inventario con Imágenes")
st.write("Haz clic en el botón inferior para conectarte a Odoo, extraer los productos consumibles y generar el archivo Excel.")

# 2. Botón de acción principal
if st.button("Generar Excel de Inventario", type="primary"):
    
    # Mostrar un mensaje de carga mientras el script trabaja
    with st.spinner('Conectando a Odoo y procesando imágenes... Esto puede tomar unos segundos.'):
        
        try:
            # Credenciales de Odoo (Asegúrate de poner las tuyas)
            url = 'https://cons240326.odoo.com'
            db = 'cons240326'
            username = 'oscar.moscoso@omrworkgroup.com'
            password = st.secrets["ODOO_PASSWORD"]

            # Conexión
            common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
            uid = common.authenticate(db, username, password, {})
            models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

            # Buscar productos
            productos = models.execute_kw(db, uid, password, 
                    'product.template', 'search_read',
                    [[('type', '=', 'consu')]], 
                    {'fields': ['name', 'list_price', 'qty_available', 'image_128']})

            # Crear el archivo Excel en la memoria (BytesIO) en vez de en el disco duro
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Productos')

            # Formato de columnas y encabezados
            worksheet.set_column('A:A', 30)
            worksheet.set_column('B:B', 15)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('D:D', 20)
            worksheet.write('A1', 'Productos')
            worksheet.write('B1', 'Precio')
            worksheet.write('C1', 'A la mano')
            worksheet.write('D1', 'Imagen')

            # Llenar datos
            row = 1
            for prod in productos:
                worksheet.write(row, 0, prod.get('name', ''))
                worksheet.write(row, 1, prod.get('list_price', 0))
                worksheet.write(row, 2, prod.get('qty_available', 0))

                imagen_base64 = prod.get('image_128')
                if imagen_base64:
                    try:
                        image_data = base64.b64decode(imagen_base64)
                        imagen_pil = Image.open(io.BytesIO(image_data))
                        stream_imagen = io.BytesIO()
                        imagen_pil.save(stream_imagen, format="PNG")
                        stream_imagen.seek(0)
                        
                        worksheet.set_row(row, 80) 
                        worksheet.insert_image(row, 3, 'imagen_limpia.png', {'image_data': stream_imagen, 'x_scale': 0.8, 'y_scale': 0.8})
                    except Exception as e:
                        worksheet.write(row, 3, 'Sin imagen válida')
                row += 1

            # Cerrar el libro de Excel
            workbook.close()
            
            # Mensaje de éxito
            st.success("¡El archivo se ha generado correctamente!")
            
            # 3. Crear el botón de descarga web
            st.download_button(
                label="📥 Descargar Inventario.xlsx",
                data=output.getvalue(), # Extraer los datos de la memoria
                file_name="Inventario_con_Imagenes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as error_conexion:
            st.error(f"Ocurrió un error al conectar con Odoo: {error_conexion}")