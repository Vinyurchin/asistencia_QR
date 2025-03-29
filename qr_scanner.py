import cv2
import pymysql
from pyzbar.pyzbar import decode
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

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

# Nombre del archivo de Excel
excel_file = "asistencias.xlsx"

# Intentar cargar el archivo, si no existe, crearlo
try:
    wb = load_workbook(excel_file)
    ws = wb.active
except FileNotFoundError:
    wb = Workbook()
    ws = wb.active
    ws.append(["Nombre", "Apellido", "Número de Alumno"])  # Encabezados base
    wb.save(excel_file)
except PermissionError:
    print(f"Error: El archivo {excel_file} está abierto. Por favor, ciérralo e intenta de nuevo.")
    exit(1)

# Función para registrar asistencia correctamente en la columna de la fecha actual
def registrar_asistencia_excel(usuario_id, nombre, apellido):
    fecha_actual = datetime.now().strftime("%Y-%m-%d")

    # Obtener la lista de encabezados
    encabezados = [cell.value for cell in ws[1] if cell.value]  # Ignorar celdas vacías

    # Verificar si la fecha ya existe en los encabezados
    if fecha_actual not in encabezados:
        col_index = len(encabezados) + 1  # Nueva columna para la fecha
        ws.cell(row=1, column=col_index, value=fecha_actual)
        wb.save(excel_file)
    else:
        col_index = encabezados.index(fecha_actual) + 1  # Buscar la columna correcta

    # Buscar la fila del usuario
    for row in ws.iter_rows(min_row=2, max_col=len(encabezados), values_only=False):
        if row[0].value == nombre and row[1].value == apellido and row[2].value == usuario_id:
            ws.cell(row=row[0].row, column=col_index, value="✔")  # Marcar asistencia con un check
            ws.cell(row=row[0].row, column=col_index).fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            wb.save(excel_file)
            print(f"Asistencia guardada en Excel: {nombre} {apellido}")
            return

    # Si el usuario no está en la lista, agregarlo
    new_row = [nombre, apellido, usuario_id] + [None] * (col_index - 3) + ["✔"]
    ws.append(new_row)
    ws.cell(row=ws.max_row, column=col_index).fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    wb.save(excel_file)
    print(f"Asistencia guardada en Excel: {nombre} {apellido}")

# Validación de datos
def validar_datos(usuario_id, nombre, apellido):
    if not usuario_id or not nombre or not apellido:
        print("Error: Datos incompletos del usuario.")
        return False
    return True

# Inicializa la cámara
cap = cv2.VideoCapture(0)

# Lista para evitar múltiples registros del mismo QR
procesados = set()

while True:
    ret, frame = cap.read()
    if not ret:
        break

    # Decodifica los QR en la imagen
    for qr in decode(frame):
        qr_data = qr.data.decode("utf-8")  # Extrae el contenido del QR
        qr_rect = qr.rect  # Obtiene la posición del QR

        # Evitar procesar el mismo QR repetidamente
        if qr_data in procesados:
            continue

        # Dibuja un rectángulo alrededor del QR
        cv2.rectangle(frame, (qr_rect.left, qr_rect.top),
                      (qr_rect.left + qr_rect.width, qr_rect.top + qr_rect.height),
                      (0, 255, 0), 3)

        # Buscar el código QR en la base de datos
        try:
            cursor.execute("SELECT id, nombre, apellido FROM usuarios WHERE qr_code = %s", (qr_data,))
            usuario = cursor.fetchone()

            if usuario:
                usuario_id, nombre, apellido = usuario

                # Validar datos del usuario
                if not validar_datos(usuario_id, nombre, apellido):
                    continue

                fecha_actual = datetime.now().strftime('%Y-%m-%d')

                # Verificar si ya tiene asistencia hoy
                cursor.execute("SELECT * FROM asistencia WHERE usuario_id = %s AND DATE(fecha_asistencia) = %s", (usuario_id, fecha_actual))
                asistencia_existente = cursor.fetchone()

                if asistencia_existente:
                    mensaje = f"El usuario {nombre} {apellido} ya tiene asistencia registrada hoy."
                    print(mensaje)
                else:
                    # Registrar asistencia en MySQL
                    cursor.execute("INSERT INTO asistencia (usuario_id, fecha_asistencia) VALUES (%s, NOW())", (usuario_id,))
                    conexion.commit()

                    # Guardar en Excel
                    registrar_asistencia_excel(usuario_id, nombre, apellido)

                    mensaje = f"Asistencia registrada: {nombre} {apellido}"
                    print(mensaje)

                # Agregar el QR procesado a la lista
                procesados.add(qr_data)
            else:
                mensaje = "Código QR no registrado"
                print(mensaje)

            # Mostrar mensaje en la ventana sin detener la cámara
            cv2.putText(frame, mensaje, (qr_rect.left, qr_rect.top - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 0), 2)

        except pymysql.MySQLError as e:
            print(f"Error al consultar la base de datos: {e}")

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