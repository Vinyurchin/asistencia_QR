import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel, Listbox
import threading
from qr_scanner import iniciar_escaneo_qr, descargar_excel_desde_qr_scanner
from generar_qrs import generar_qr_para_usuario
from PIL import Image, ImageTk
import os
import shutil
from openpyxl import load_workbook

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
        except Exception as e:
            print(f"Error al iniciar la cámara: {e}")
            messagebox.showerror("Error", f"Error al iniciar la cámara: {e}")

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
            nombres = entry_nombre.get()
            apellido = entry_apellido.get()
            segundo_apellido = entry_segundo_apellido.get()

            if not nombres or not apellido:
                messagebox.showwarning("Advertencia", "Por favor, ingresa el nombre(s) y apellido(s).")
                return

            # Combine the second surname if provided
            nombre_completo = f"{nombres} {apellido}"
            if segundo_apellido:
                nombre_completo += f" {segundo_apellido}"

            resultado = generar_qr_para_usuario(nombres, apellido)
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
    ventana_qrs.geometry("400x500")

    # Add a search bar
    def filtrar_qrs():
        query = entry_buscar.get().lower()
        lista_qrs.delete(0, tk.END)
        for archivo in archivos_qr:
            if query in archivo.lower():
                lista_qrs.insert(tk.END, archivo)

    label_buscar = tk.Label(ventana_qrs, text="Buscar QR:", font=("Arial", 12))
    label_buscar.pack(pady=5)
    entry_buscar = tk.Entry(ventana_qrs, width=30)
    entry_buscar.pack(pady=5)
    entry_buscar.bind("<KeyRelease>", lambda event: filtrar_qrs())

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

# Improve the interface design
root.configure(bg="lightblue")

label_titulo = tk.Label(root, text="Sistema de Asistencia con QR", font=("Arial", 20, "bold"), bg="lightblue", fg="darkblue")
label_titulo.pack(pady=10)

btn_iniciar_camara = tk.Button(root, text="Iniciar Cámara", command=iniciar_camara, width=20, height=2, bg="green", fg="white", font=("Arial", 12, "bold"))
btn_iniciar_camara.pack(pady=10)

btn_descargar_excel = tk.Button(root, text="Descargar Excel", command=descargar_excel, width=20, height=2, bg="blue", fg="white", font=("Arial", 12, "bold"))
btn_descargar_excel.pack(pady=10)

btn_visualizar_excel = tk.Button(root, text="Visualizar Excel", command=lambda: os.startfile(excel_file), width=20, height=2, bg="gray", fg="white", font=("Arial", 12, "bold"))
btn_visualizar_excel.pack(pady=10)

# Move the definition of limpiar_excel above the button creation
def limpiar_excel():
    def confirmar_limpiar():
        if messagebox.askyesno("Confirmación", "¿Estás seguro de que deseas limpiar el archivo Excel? Esta acción no se puede deshacer."):
            try:
                # Crear un respaldo antes de limpiar
                ruta_respaldo = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Archivos de Excel", "*.xlsx")],
                    title="Guardar respaldo del archivo de Excel"
                )
                if ruta_respaldo:
                    shutil.copy(excel_file, ruta_respaldo)
                    messagebox.showinfo("Éxito", f"Respaldo guardado en: {ruta_respaldo}")

                # Limpiar el archivo Excel
                wb = Workbook()
                ws = wb.active
                ws.append(["Nombre", "Apellido", "Número de Alumno"])  # Encabezados base
                wb.save(excel_file)
                messagebox.showinfo("Éxito", "El archivo Excel ha sido limpiado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo limpiar el archivo Excel: {e}")

    # Ventana de confirmación
    confirmar_limpiar()

# Ensure the button is created after the function definition
btn_limpiar_excel = tk.Button(root, text="Limpiar Excel", command=limpiar_excel, width=20, height=2, bg="red", fg="white", font=("Arial", 12, "bold"))
btn_limpiar_excel.pack(pady=10)

# Botón para importar el Excel
btn_importar_excel = tk.Button(root, text="Importar Excel", command=importar_excel, width=20, height=2, bg="orange", fg="white", font=("Arial", 12, "bold"))
btn_importar_excel.pack(pady=10)

frame_generar_usuario = tk.Frame(root, bg="lightblue")
frame_generar_usuario.pack(pady=20)

label_nombre = tk.Label(frame_generar_usuario, text="Nombre(s):", bg="lightblue", fg="darkblue", font=("Arial", 12))
label_nombre.grid(row=0, column=0, padx=10, pady=5)
entry_nombre = tk.Entry(frame_generar_usuario, width=30)
entry_nombre.grid(row=0, column=1, padx=10, pady=5)

label_apellido = tk.Label(frame_generar_usuario, text="Apellido(s):", bg="lightblue", fg="darkblue", font=("Arial", 12))
label_apellido.grid(row=1, column=0, padx=10, pady=5)
entry_apellido = tk.Entry(frame_generar_usuario, width=30)
entry_apellido.grid(row=1, column=1, padx=10, pady=5)

entry_segundo_apellido = tk.Entry(frame_generar_usuario, width=30)
entry_segundo_apellido.grid(row=2, column=1, padx=10, pady=5)
label_segundo_apellido = tk.Label(frame_generar_usuario, text="Segundo Apellido:", bg="lightblue", fg="darkblue", font=("Arial", 12))
label_segundo_apellido.grid(row=2, column=0, padx=10, pady=5)

btn_generar_usuario = tk.Button(frame_generar_usuario, text="Generar Usuario", command=generar_usuario_desde_interfaz, bg="orange", fg="white", font=("Arial", 12, "bold"))
btn_generar_usuario.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

btn_mostrar_qrs = tk.Button(root, text="Mostrar Todos los QRs", command=mostrar_todos_qrs, width=20, height=2, bg="purple", fg="white", font=("Arial", 12, "bold"))
btn_mostrar_qrs.pack(pady=10)

# Add a footer
footer = tk.Label(root, text="Sistema de Asistencia con QR - 2025", bg="lightblue", fg="darkblue", font=("Arial", 10))
footer.pack(side=tk.BOTTOM, pady=10)

root.mainloop()

# Add a button to import an Excel file and validate its format
def importar_excel():
    try:
        # Seleccionar el archivo Excel
        ruta_excel = filedialog.askopenfilename(
            filetypes=[("Archivos de Excel", "*.xlsx")],
            title="Seleccionar archivo de Excel"
        )
        if not ruta_excel:
            return

        # Cargar el archivo Excel
        wb = load_workbook(ruta_excel)
        ws = wb.active

        # Verificar las columnas esperadas
        columnas_esperadas = ["Nombre", "Apellido", "Número de Alumno"]
        columnas_excel = [cell.value for cell in ws[1] if cell.value]

        if not all(col in columnas_excel for col in columnas_esperadas):
            messagebox.showerror("Error", "El archivo Excel no tiene el formato adecuado.")
            return

        # Sobrescribir el archivo actual con el nuevo
        shutil.copy(ruta_excel, excel_file)
        messagebox.showinfo("Éxito", "El archivo Excel ha sido importado correctamente.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo importar el archivo Excel: {e}")