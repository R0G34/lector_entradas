import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import usb.core
import usb.util
from datetime import datetime

# Configuración de la ventana principal
root = tk.Tk()
root.title("Gestión Entradas WTC24")
root.geometry("800x600")
root.configure(bg='white')

# Variables para almacenar la ruta del archivo y los datos
file_path = None
datos_excel = None

# Función para seleccionar el archivo Excel
def seleccionar_archivo():
    global file_path, datos_excel
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    if file_path:
        # Leer el archivo Excel y cargarlo en datos_excel
        datos_excel = pd.read_excel(file_path, dtype=str)
        messagebox.showinfo("Archivo Cargado", "El archivo ha sido cargado correctamente.")

# Función para encontrar automáticamente una impresora Zebra conectada
def encontrar_impresora_zebra():
    # Buscar todos los dispositivos USB y encontrar uno de la marca Zebra
    dispositivos = usb.core.find(find_all=True)
    for dispositivo in dispositivos:
        # Verificar si el fabricante es "Zebra" (o usa el ID de fabricante de Zebra si lo conoces)
        if dispositivo.manufacturer and "Zebra" in dispositivo.manufacturer:
            return dispositivo
    raise ValueError("Impresora Zebra no encontrada. Asegúrate de que está conectada.")

# Función para enviar el comando ZPL a la impresora real
"""def enviar_a_impresora_real(zpl_command):
    try:
        # Buscar e inicializar la impresora Zebra
        dispositivo = encontrar_impresora_zebra()
        
        # Si es necesario, desactivar el driver de kernel y configurar la impresora
        if dispositivo.is_kernel_driver_active(0):
            dispositivo.detach_kernel_driver(0)
        dispositivo.set_configuration()

        # Enviar el comando ZPL a la impresora
        bytes_written = dispositivo.write(1, zpl_command.encode('iso-8859-1'))
        if bytes_written == len(zpl_command):
            print("Etiqueta enviada a la impresora correctamente.")
        else:
            print("Error: No se enviaron todos los datos a la impresora.")

    except usb.core.USBError as e:
        print(f"Error de USB: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")
    finally:
        # Liberar recursos de la impresora
        if 'dispositivo' in locals():
            usb.util.dispose_resources(dispositivo)"""
def enviar_a_impresora_real(zpl_command):
    try:
        dispositivo = encontrar_impresora_zebra()  # Función que encuentra la impresora
        # Enviar el comando ZPL
        bytes_written = dispositivo.write(1, zpl_command.encode('iso-8859-1'))
        if bytes_written == len(zpl_command):
            print("Etiqueta enviada a la impresora correctamente.")
        else:
            print("Error: No se enviaron todos los datos a la impresora.")

    except usb.core.USBError as e:
        print(f"Error de USB: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")
    finally:
        if 'dispositivo' in locals():
            usb.util.dispose_resources(dispositivo)

# Función para verificar si la impresora está conectada
def verificar_impresora():
    try:
        dispositivo = encontrar_impresora_zebra()
        messagebox.showinfo("Detección de Impresora", "¡Impresora Zebra detectada exitosamente!")
    except ValueError as e:
        messagebox.showerror("Detección de Impresora", "Impresora no detectada.")

# Función para buscar el QR en el Excel y enviar a la impresora
def buscar_qr():
    global datos_excel
    if datos_excel is None:
        messagebox.showwarning("Archivo no seleccionado", "Por favor, selecciona un archivo Excel primero.")
        return

    qr_code = entry_qr.get().strip()
    if qr_code:
        # Buscar el QR en la primera columna
        fila = datos_excel[datos_excel.iloc[:, 0] == qr_code]
        if not fila.empty:
            if pd.isna(fila.iloc[0].get('FECHA', '')) or fila.iloc[0]['FECHA'] == "":
                fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
                datos_excel.loc[datos_excel.iloc[:, 0] == qr_code, 'FECHA'] = fecha_actual
                datos_excel.to_excel(file_path, index=False)
                print("Fecha registrada:", fecha_actual)
            else:
                messagebox.showinfo("Registro existente", "Esta entrada ya está registrada.")
                return
            # Extraer los valores de las columnas
            nombre = fila.iloc[0]['NOMBRE']
            apellidos = fila.iloc[0]['APELLIDOS']
            empresa = fila.iloc[0]['EMPRESA']
            cargo = fila.iloc[0]['CARGO']
            
            # Actualizar las etiquetas con la información
            label_nombre_val.config(text=nombre)
            label_apellidos_val.config(text=apellidos)
            label_empresa_val.config(text=empresa)
            label_cargo_val.config(text=cargo)
            entry_qr.delete(0, tk.END)
            
            def remove_accents(text):
                accents = {
                    'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
                    'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U'
                }
                for accented_char, replacement in accents.items():
                    text = text.replace(accented_char, replacement)
                return text

            nombre = remove_accents(nombre)
            apellidos = remove_accents(apellidos)
            empresa = remove_accents(empresa)
            cargo = remove_accents(cargo)

            # Crear el comando ZPL codificado en ISO-8859-1
            nombre = nombre.replace("ñ", "n")  # Reemplaza la "ñ" con el código hexadecimal específico para ZPL
            apellidos = apellidos.replace("ñ", "n")
            empresa = empresa.replace("ñ", "n")
            cargo = cargo.replace("ñ", "n")

            zpl_command = f"""
            ^XA
            ^CI28
            ^FO20,70^A0N,70,70^FDNombre:^FS
            ^FO300,70^A0N,70,70^FD{nombre}^FS
            ^FO20,150^A0N,70,70^FDApellido:^FS
            ^FO300,150^A0N,70,70^FD{apellidos}^FS
            ^FO20,230^A0N,70,70^FDEmpresa:^FS
            ^FO300,230^A0N,70,70^FD{empresa}^FS
            ^FO20,310^A0N,70,70^FDCargo:^FS
            ^FO300,310^A0N,70,70^FD{cargo}^FS
            ^XZ
            """

            # Enviar el comando ZPL a la impresora real
            enviar_a_impresora_real(zpl_command)

        else:
            messagebox.showinfo("No encontrado", "El código QR no se encontró en el archivo.")
            # Limpiar las etiquetas si no se encuentra el código
            label_nombre_val.config(text="")
            label_apellidos_val.config(text="")
            label_empresa_val.config(text="")
            label_cargo_val.config(text="")
            entry_qr.delete(0, tk.END)
    else:
        messagebox.showwarning("QR vacío", "Por favor, ingresa un código QR.")
        entry_qr.delete(0, tk.END)
"""
label_nombre_manual = tk.Label(root, text="Nombre:", font=("Arial", 10), bg='white')
label_nombre_manual.pack(anchor="w", padx=10, pady=(5, 0))
entry_nombre = tk.Entry(root, font=("Arial", 12), width=30, relief="solid", bd=1)
entry_nombre.pack(anchor="w", padx=10)

label_apellido_manual = tk.Label(root, text="Apellido:", font=("Arial", 10), bg='white')
label_apellido_manual.pack(anchor="w", padx=10, pady=(5, 0))
entry_apellido = tk.Entry(root, font=("Arial", 12), width=30, relief="solid", bd=1)
entry_apellido.pack(anchor="w", padx=10)"""

def abrir_ventana_busqueda():
    ventana_busqueda = tk.Toplevel(root)
    ventana_busqueda.title("Búsqueda Manual")
    ventana_busqueda.geometry("300x200")
    
    # Etiqueta y campo de entrada para Nombre
    label_nombre = tk.Label(ventana_busqueda, text="Nombre:", font=("Arial", 10))
    label_nombre.pack(pady=(10, 5))
    entry_nombre = tk.Entry(ventana_busqueda, font=("Arial", 12), width=25)
    entry_nombre.pack()

    # Etiqueta y campo de entrada para Apellido
    label_apellido = tk.Label(ventana_busqueda, text="Apellido:", font=("Arial", 10))
    label_apellido.pack(pady=(10, 5))
    entry_apellido = tk.Entry(ventana_busqueda, font=("Arial", 12), width=25)
    entry_apellido.pack()

    # Función de búsqueda manual
    def buscar_manual():
        global datos_excel
        nombre = entry_nombre.get().strip()
        apellido = entry_apellido.get().strip()

        if not nombre or not apellido:
            messagebox.showwarning("Campos vacíos", "Por favor, ingresa tanto el Nombre como el Apellido.")
            return

        if datos_excel is None:
            messagebox.showwarning("Archivo no seleccionado", "Por favor, selecciona un archivo Excel primero.")
            return

        # Buscar la fila que coincida con el Nombre y Apellido
        fila = datos_excel[(datos_excel['NOMBRE'] == nombre) & (datos_excel['APELLIDOS'] == apellido)]
        if not fila.empty:
            # Verificar si la columna 'FECHA' está vacía y actualizar con la fecha y hora actual
            if pd.isna(fila.iloc[0].get('FECHA', '')) or fila.iloc[0]['FECHA'] == "":
                fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
                datos_excel.loc[(datos_excel['NOMBRE'] == nombre) & (datos_excel['APELLIDOS'] == apellido), 'FECHA'] = fecha_actual
                datos_excel.to_excel(file_path, index=False)  # Guardar cambios en el archivo
                print("Fecha registrada:", fecha_actual)
            else:
                messagebox.showinfo("Registro existente", "Esta entrada ya está registrada.")
                ventana_busqueda.destroy()
                return

            # Extraer los valores de las columnas para mostrar en la interfaz principal
            empresa = fila.iloc[0]['EMPRESA']
            cargo = fila.iloc[0]['CARGO']
            label_nombre_val.config(text=nombre)
            label_apellidos_val.config(text=apellido)
            label_empresa_val.config(text=empresa)
            label_cargo_val.config(text=cargo)

            # Crear y enviar el comando ZPL con los datos
            zpl_command = f"""
            ^XA
            ^CI28
            ^FO20,70^A0N,70,70^FDNombre:^FS
            ^FO300,70^A0N,70,70^FD{nombre}^FS
            ^FO20,150^A0N,70,70^FDApellido:^FS
            ^FO300,150^A0N,70,70^FD{apellido}^FS
            ^FO20,230^A0N,70,70^FDEmpresa:^FS
            ^FO300,230^A0N,70,70^FD{empresa}^FS
            ^FO20,310^A0N,70,70^FDCargo:^FS
            ^FO300,310^A0N,70,70^FD{cargo}^FS
            ^XZ
            """
            enviar_a_impresora_real(zpl_command)

            ventana_busqueda.destroy()  # Cerrar la ventana después de la búsqueda
        else:
            messagebox.showinfo("No encontrado", "No se encontró ningún registro con ese Nombre y Apellido.")
            ventana_busqueda.destroy()

    # Botón para ejecutar la búsqueda en la ventana emergente
    button_buscar = tk.Button(ventana_busqueda, text="Buscar", font=("Arial", 10), command=buscar_manual)
    button_buscar.pack(pady=(20, 10))

    ventana_busqueda.bind("<Return>", lambda event: buscar_manual())
# Título principal
label_title = tk.Label(root, text="GESTIÓN ENTRADAS WTC24", font=("Arial", 20, "bold"), bg='white')
label_title.pack(pady=(20, 10))

# Botón para verificar si la impresora está conectada (ubicado en la esquina inferior derecha)
"""button_verificar = tk.Button(root, text="Verificar Impresora", font=("Arial", 10), bg='#D3D3D3', relief="flat", command=verificar_impresora)
button_verificar.place(relx=1.0, rely=1.0, anchor="se")  # Coloca el botón en la esquina inferior derecha"""

# Botón para la búsqueda manual por Nombre y Apellido
button_buscar_manual = tk.Button(root, text="Búsqueda Manual", font=("Arial", 10), bg='#D3D3D3', relief="flat", command=abrir_ventana_busqueda)
button_buscar_manual.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

# Campo para el código QR
label_qr = tk.Label(root, text="CÓDIGO QR", font=("Arial", 12), bg='white')
label_qr.pack(pady=(10, 5))

entry_qr = tk.Entry(root, font=("Arial", 14), width=30, relief="solid", bd=1)
entry_qr.pack(pady=(0, 20))

# Asociar el evento "Enter" y "Tab" al campo de entrada
entry_qr.bind("<Return>", lambda event: buscar_qr())
entry_qr.bind("<Tab>", lambda event: buscar_qr())

# Botón para asignar la base de datos
button_db = tk.Button(root, text="ASIGNAR BASE DE DATOS", font=("Arial", 10), bg='#D3D3D3', relief="flat", command=seleccionar_archivo)
button_db.pack(pady=(0, 10))

# Botón para buscar el código QR
button_search = tk.Button(root, text="BUSCAR", font=("Arial", 10), bg='#D3D3D3', relief="flat", command=buscar_qr)
button_search.pack(pady=(0, 20))

# Etiquetas para mostrar los resultados
label_nombre = tk.Label(root, text="Nombre:", font=("Arial", 12), bg='white')
label_nombre.pack(anchor="w", padx=50)
label_nombre_val = tk.Label(root, text="", font=("Arial", 12, "bold"), bg='white')
label_nombre_val.pack(anchor="w", padx=50)

label_apellidos = tk.Label(root, text="Apellidos:", font=("Arial", 12), bg='white')
label_apellidos.pack(anchor="w", padx=50)
label_apellidos_val = tk.Label(root, text="", font=("Arial", 12, "bold"), bg='white')
label_apellidos_val.pack(anchor="w", padx=50)

label_empresa = tk.Label(root, text="Empresa:", font=("Arial", 12), bg='white')
label_empresa.pack(anchor="w", padx=50)
label_empresa_val = tk.Label(root, text="", font=("Arial", 12, "bold"), bg='white')
label_empresa_val.pack(anchor="w", padx=50)

label_cargo = tk.Label(root, text="Cargo:", font=("Arial", 12), bg='white')
label_cargo.pack(anchor="w", padx=50)
label_cargo_val = tk.Label(root, text="", font=("Arial", 12, "bold"), bg='white')
label_cargo_val.pack(anchor="w", padx=50)

# Iniciar el bucle principal
verificar_impresora()
root.mainloop()
