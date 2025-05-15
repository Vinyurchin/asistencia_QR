import qrcode
import random
import string
import pymysql
import os
import tkinter as tk
from PIL import Image, ImageTk
import re

# Crear la carpeta /imagenes_qr si no existe
output_dir = "imagenes_qr"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Función para generar un código QR único
def generar_codigo_qr():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))  # Código de 8 caracteres

# Función para generar un QR para un usuario
def generar_qr_para_usuario(nombre, apellido, segundo_apellido=None):
    # Validar entradas
    patron = re.compile(r'^[A-Za-zÁÉÍÓÚáéíóúÑñüÜ\s]+$')
    if not nombre or not apellido:
        return "Error: El nombre y el apellido no pueden estar vacíos."
    if not patron.match(nombre) or not patron.match(apellido) or (segundo_apellido and not patron.match(segundo_apellido)):
        return "Error: El nombre y los apellidos solo deben contener letras y espacios."

    try:
        # Conexión a MySQL
        conexion = pymysql.connect(
            host="localhost",
            user="root",
            password="1234",
            database="asistencia"
        )
        cursor = conexion.cursor()

        # Verificar si ya existe un código QR para este usuario
        cursor.execute("SELECT qr_code FROM usuarios WHERE nombre = %s AND apellido = %s AND (segundo_apellido = %s OR segundo_apellido IS NULL)", (nombre, apellido, segundo_apellido))
        resultado = cursor.fetchone()

        if resultado:
            codigo_qr = resultado[0]  # Ya tiene un código QR
            if codigo_qr is None:
                # Generar un nuevo código QR único
                codigo_qr = generar_codigo_qr()
                while verificar_codigo_qr_existente(cursor, codigo_qr):
                    codigo_qr = generar_codigo_qr()
                cursor.execute("UPDATE usuarios SET qr_code = %s WHERE nombre = %s AND apellido = %s AND (segundo_apellido = %s OR segundo_apellido IS NULL)", 
                               (codigo_qr, nombre, apellido, segundo_apellido))
                conexion.commit()
            print(f"{nombre} {apellido} {segundo_apellido or ''} ya tiene un código QR: {codigo_qr}")
        else:
            # Generar un nuevo código QR único
            codigo_qr = generar_codigo_qr()
            while verificar_codigo_qr_existente(cursor, codigo_qr):
                codigo_qr = generar_codigo_qr()
            # Insertar usuario en MySQL con su código QR
            cursor.execute("INSERT INTO usuarios (nombre, apellido, segundo_apellido, qr_code) VALUES (%s, %s, %s, %s)",
                           (nombre, apellido, segundo_apellido, codigo_qr))
            conexion.commit()

        # Crear código QR
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4
        )
        qr.add_data(codigo_qr)
        qr.make(fit=True)

        # Generar imagen del QR
        imagen_qr = qr.make_image(fill="black", back_color="white")

        # Guardar la imagen con el nombre del usuario
        nombre_archivo = os.path.join(output_dir, f"{nombre}_{apellido}{'_' + segundo_apellido if segundo_apellido else ''}.png").replace(" ", "_")
        imagen_qr.save(nombre_archivo)

        # Show the QR code to the user
        imagen_qr.show()

        # Display the generated QR code in a window similar to "Mostrar Todos los QRs"
        ventana_qr = tk.Toplevel()
        ventana_qr.title("QR Generado")
        ventana_qr.geometry("400x400")

        label_titulo = tk.Label(ventana_qr, text="QR Generado Correctamente", font=("Arial", 14))
        label_titulo.pack(pady=10)

        imagen = Image.open(nombre_archivo)
        imagen_tk = ImageTk.PhotoImage(imagen)

        label_imagen = tk.Label(ventana_qr, image=imagen_tk)
        label_imagen.image = imagen_tk
        label_imagen.pack(pady=10)

        btn_cerrar = tk.Button(ventana_qr, text="Cerrar", command=ventana_qr.destroy, bg="red", fg="white", font=("Arial", 10))
        btn_cerrar.pack(pady=10)

        print(f"Código QR para {nombre} {apellido} {segundo_apellido or ''}: {codigo_qr} → Guardado como {nombre_archivo}")
        return f"Código QR generado correctamente: {nombre_archivo}"

    except pymysql.MySQLError as e:
        print(f"Error al ejecutar la consulta para {nombre} {apellido} {segundo_apellido or ''}: {e}")
        return f"Error al generar el QR: {e}"

    finally:
        # Cerrar la conexión a la base de datos
        if 'conexion' in locals() and conexion.open:
            conexion.close()

# Función para verificar si un código QR ya existe en la base de datos
def verificar_codigo_qr_existente(cursor, codigo_qr):
    cursor.execute("SELECT 1 FROM usuarios WHERE qr_code = %s", (codigo_qr,))
    return cursor.fetchone() is not None