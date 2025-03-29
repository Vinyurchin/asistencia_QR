import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import cv2
from qr_scanner import registrar_asistencia_excel, validar_datos  # Importar funciones necesarias
import os

# Ruta del archivo Excel
excel_file = "asistencias.xlsx"

# Función para iniciar la cámara y escanear QR
def iniciar_camara():
    def ejecutar_escaneo():
        try:
            # Inicializa la cámara
            cap = cv2.VideoCapture(0)
            procesados = set()  # Lista para evitar múltiples registros del mismo QR

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

        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar la cámara: {e}")

    # Ejecutar el escaneo en un hilo separado para no bloquear la interfaz
    hilo = threading.Thread(target=ejecutar_escaneo)
    hilo.daemon = True
    hilo.start()

# Función para descargar el archivo Excel
def descargar_excel():
    if not os.path.exists(excel_file):
        messagebox.showerror("Error", "El archivo de Excel no existe.")
        return

    # Abrir un cuadro de diálogo para guardar el archivo
    ruta_destino = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx")],
        title="Guardar archivo de Excel"
    )
    if ruta_destino:
        try:
            os.replace(excel_file, ruta_destino)
            messagebox.showinfo("Éxito", f"Archivo guardado en: {ruta_destino}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

# Crear la ventana principal
root = tk.Tk()
root.title("Sistema de Asistencia con QR")
root.geometry("400x300")

# Etiqueta de título
label_titulo = tk.Label(root, text="Sistema de Asistencia con QR", font=("Arial", 16))
label_titulo.pack(pady=10)

# Botón para iniciar la cámara
btn_iniciar_camara = tk.Button(root, text="Iniciar Cámara", command=iniciar_camara, width=20, height=2, bg="green", fg="white")
btn_iniciar_camara.pack(pady=10)

# Botón para descargar el archivo Excel
btn_descargar_excel = tk.Button(root, text="Descargar Excel", command=descargar_excel, width=20, height=2, bg="blue", fg="white")
btn_descargar_excel.pack(pady=10)

# Ejecutar la ventana principal
root.mainloop()