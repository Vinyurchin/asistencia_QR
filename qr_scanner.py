import cv2
import pymysql
from pyzbar.pyzbar import decode
from datetime import datetime
import time

# Conectar a MySQL
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

# Inicializa la cámara
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        break

    # Decodifica los QR en la imagen
    for qr in decode(frame):
        qr_data = qr.data.decode("utf-8")  # Extrae el contenido del QR
        qr_rect = qr.rect  # Obtiene la posición del QR

        # Dibuja un rectángulo alrededor del QR
        cv2.rectangle(frame, (qr_rect.left, qr_rect.top),
                      (qr_rect.left + qr_rect.width, qr_rect.top + qr_rect.height),
                      (0, 255, 0), 3)

        # Buscar el código QR en la base de datos
        cursor.execute("SELECT id, nombre, apellido FROM usuarios WHERE qr_code = %s", (qr_data,))
        usuario = cursor.fetchone()

        if usuario:
            usuario_id, nombre, apellido = usuario
            fecha_actual = datetime.now().strftime('%Y-%m-%d')

            # Verificar si ya tiene asistencia hoy
            cursor.execute("SELECT * FROM asistencia WHERE usuario_id = %s AND DATE(fecha_asistencia) = %s", (usuario_id, fecha_actual))
            asistencia_existente = cursor.fetchone()

            if asistencia_existente:
                mensaje = f"El usuario {nombre} {apellido} ya tiene asistencia registrada hoy."
                print(mensaje)
            else:
                # Registrar asistencia
                cursor.execute("INSERT INTO asistencia (usuario_id, fecha_asistencia) VALUES (%s, NOW())", (usuario_id,))
                conexion.commit()
                mensaje = f"Asistencia registrada: {nombre} {apellido}"
                print(mensaje)

        else:
            mensaje = "Código QR no registrado"

        # Muestra el mensaje en la imagen
        cv2.putText(frame, mensaje, (qr_rect.left, qr_rect.top - 10),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 0), 2)

        # Pausa de 3 segundos
        time.sleep(3)

    # Muestra la cámara en tiempo real
    cv2.imshow("Escáner QR", frame)

    # Presiona 'q' para salir
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

# Libera la cámara y cierra la ventana
cap.release()
cv2.destroyAllWindows()

# Cierra la conexión a la base de datos
cursor.close()
conexion.close()