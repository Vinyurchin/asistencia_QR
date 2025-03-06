import qrcode
import random
import string
import pymysql
import os

# Conexión a MySQL
try:
    conexion = pymysql.connect(
        host="localhost",
        user="root",
        password="1234",
        database="asistencia"
    )
    cursor = conexion.cursor()
except pymysql.MySQLError as e:
    print(f"Error al conectar a la base de datos: {e}")
    exit(1)

# Función para generar un código QR único
def generar_codigo_qr():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))  # Código de 8 caracteres

# Crear la carpeta /imagenes_qr si no existe
output_dir = "imagenes_qr"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Lista de usuarios de ejemplo
usuarios = [
    ("Jonathan", "Hernandez")
]

for nombre, apellido in usuarios:
    try:
        # Verificar si ya existe un código QR para este usuario
        cursor.execute("SELECT qr_code FROM usuarios WHERE nombre = %s AND apellido = %s", (nombre, apellido))
        resultado = cursor.fetchone()

        if resultado:
            codigo_qr = resultado[0]  # Ya tiene un código QR
            if codigo_qr is None:
                codigo_qr = generar_codigo_qr()  # Generar nuevo QR si está vacío
                cursor.execute("UPDATE usuarios SET qr_code = %s WHERE nombre = %s AND apellido = %s", 
                               (codigo_qr, nombre, apellido))
                conexion.commit()
            print(f"{nombre} {apellido} ya tiene un código QR: {codigo_qr}")
        else:
            codigo_qr = generar_codigo_qr()  # Generar un nuevo código QR
            # Insertar usuario en MySQL con su código QR
            cursor.execute("INSERT INTO usuarios (nombre, apellido, qr_code, foto) VALUES (%s, %s, %s, NULL)",
                           (nombre, apellido, codigo_qr))
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

        # Guardar la imagen con el código QR en la carpeta /imagenes_qr
        nombre_archivo = os.path.join(output_dir, f"QR_{codigo_qr}.png")
        imagen_qr.save(nombre_archivo)

        print(f"Código QR para {nombre} {apellido}: {codigo_qr} → Guardado como {nombre_archivo}")

    except pymysql.MySQLError as e:
        print(f"Error al ejecutar la consulta para {nombre} {apellido}: {e}")

# Cerrar conexión
cursor.close()
conexion.close()