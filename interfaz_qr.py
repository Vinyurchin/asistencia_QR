import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel, Listbox
import threading
from qr_scanner import iniciar_escaneo_qr, descargar_excel_desde_qr_scanner
from generar_qrs import generar_qr_para_usuario
from PIL import Image, ImageTk
import os
import shutil

# Ruta del Excel
excel_file = "asistencias.xlsx"
qr_folder = "imagenes_qr"

# Crear la carpeta de imágenes QR
if not os.path.exists(qr_folder):
    os.makedirs(qr_folder)

# Función para iniciar la cámara y escanear QR
def iniciar_camara():
    def ejecutar_escaneo():
        try:
            iniciar_escaneo_qr()
            actualizar_estado("Escaneo completado", "green")
        except Exception as e:
            print(f"Error al iniciar la cámara: {e}")
            messagebox.showerror("Error", f"Error al iniciar la cámara: {e}")
            actualizar_estado("Error al iniciar la cámara", "red")

    hilo = threading.Thread(target=ejecutar_escaneo)
    hilo.daemon = True
    hilo.start()

# Función para descargar el archivo Excel
def descargar_excel():
    try:
        descargar_excel_desde_qr_scanner()
        ruta_destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx")],
            title="Guardar archivo de Excel"
        )
        if ruta_destino:
            shutil.copy(excel_file, ruta_destino)
            messagebox.showinfo("Éxito", f"Archivo guardado en: {ruta_destino}")
    except PermissionError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

# Función para generar un nuevo usuario
def generar_usuario_desde_interfaz():
    def ejecutar_generacion():
        try:
            nombre = entry_nombre.get()
            apellido = entry_apellido.get()
            if not nombre or not apellido:
                messagebox.showwarning("Advertencia", "Por favor, ingresa el nombre y apellido.")
                return
            resultado = generar_qr_para_usuario(nombre, apellido)
            messagebox.showinfo("Éxito", resultado)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el usuario: {e}")

    hilo = threading.Thread(target=ejecutar_generacion)
    hilo.daemon = True
    hilo.start()

# Función para mostrar todos los QRs existentes
def mostrar_todos_qrs():
    if not os.path.exists(qr_folder):
        messagebox.showerror("Error", f"La carpeta '{qr_folder}' no existe. Por favor, genera un QR primero.")
        return

    archivos_qr = [f for f in os.listdir(qr_folder) if f.endswith(".png")]

    if not archivos_qr:
        messagebox.showinfo("Información", "No hay códigos QR disponibles. Por favor, genera un QR primero.")
        return

    ventana_qrs = Toplevel(root)
    ventana_qrs.title("Códigos QR Disponibles")
    ventana_qrs.geometry("400x400")

    lista_qrs = Listbox(ventana_qrs, width=50, height=20)
    lista_qrs.pack(pady=10)

    for archivo in archivos_qr:
        lista_qrs.insert(tk.END, archivo)

    def abrir_qr():
        seleccion = lista_qrs.curselection()
        if seleccion:
            archivo_seleccionado = archivos_qr[seleccion[0]]
            ruta_archivo = os.path.join(qr_folder, archivo_seleccionado)
            try:
                ventana_imagen = Toplevel(ventana_qrs)
                ventana_imagen.title(archivo_seleccionado)

                imagen = Image.open(ruta_archivo)
                imagen_tk = ImageTk.PhotoImage(imagen)

                label_imagen = tk.Label(ventana_imagen, image=imagen_tk)
                label_imagen.image = imagen_tk
                label_imagen.pack()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un QR para abrir.")

    btn_abrir_qr = tk.Button(ventana_qrs, text="Abrir QR", command=abrir_qr, bg="purple", fg="white")
    btn_abrir_qr.pack(pady=10)

# Estructura de la interfaz
root = tk.Tk()
root.title("Sistema de Asistencia con QR")
root.geometry("400x700")

label_titulo = tk.Label(root, text="Sistema de Asistencia con QR", font=("Arial", 16))
label_titulo.pack(pady=10)

btn_iniciar_camara = tk.Button(root, text="Iniciar Cámara", command=iniciar_camara, width=20, height=2, bg="green", fg="white")
btn_iniciar_camara.pack(pady=10)

btn_descargar_excel = tk.Button(root, text="Descargar Excel", command=descargar_excel, width=20, height=2, bg="blue", fg="white")
btn_descargar_excel.pack(pady=10)

btn_abrir_excel = tk.Button(root, text="Abrir Excel", command=lambda: os.startfile(excel_file), width=20, height=2, bg="gray", fg="white")
btn_abrir_excel.pack(pady=10)

frame_generar_usuario = tk.Frame(root)
frame_generar_usuario.pack(pady=20)

label_nombre = tk.Label(frame_generar_usuario, text="Nombre:")
label_nombre.grid(row=0, column=0, padx=5, pady=5)
entry_nombre = tk.Entry(frame_generar_usuario)
entry_nombre.grid(row=0, column=1, padx=5, pady=5)

label_apellido = tk.Label(frame_generar_usuario, text="Apellido:")
label_apellido.grid(row=1, column=0, padx=5, pady=5)
entry_apellido = tk.Entry(frame_generar_usuario)
entry_apellido.grid(row=1, column=1, padx=5, pady=5)

btn_generar_usuario = tk.Button(frame_generar_usuario, text="Generar Usuario", command=generar_usuario_desde_interfaz, bg="orange", fg="white")
btn_generar_usuario.grid(row=2, column=0, columnspan=2, pady=10)

btn_mostrar_qrs = tk.Button(root, text="Mostrar Todos los QRs", command=mostrar_todos_qrs, width=20, height=2, bg="purple", fg="white")
btn_mostrar_qrs.pack(pady=10)

# Create a status bar
frame_estado = tk.Frame(root, bg="lightgray", height=50)
frame_estado.pack(fill=tk.X)

label_estado = tk.Label(frame_estado, text="Estado: Esperando...", font=("Arial", 12), bg="lightgray")
label_estado.pack(pady=10)

# Function to update the status bar
def actualizar_estado(mensaje, color="black"):
    label_estado.config(text=f"Estado: {mensaje}", fg=color)

root.mainloop()