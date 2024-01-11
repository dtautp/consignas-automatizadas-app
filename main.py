import tkinter as tk
from tkinter import ttk, filedialog
from convert_outcomes import excel_to_outcomes
from convert_word import excel_to_word
from create_folder import create_folder_and_subfolder
import os

def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        ruta_label.config(text=f"Ruta del archivo seleccionado:\n{archivo}")
        etiqueta.config(text="")  # Limpiar el mensaje de resultados
        boton_convertir_outcomes.config(state=tk.NORMAL)  # Habilitar el botón de conversión a Outcomes
        boton_convertir_word.config(state=tk.NORMAL)  # Habilitar el botón de conversión a Word
        ventana.update_idletasks()  # Actualizar la interfaz gráfica

    else:
        ruta_label.config(text="Ruta del archivo seleccionado:\nNinguna")
        etiqueta.config("")
        boton_convertir_outcomes.config(state=tk.DISABLED)  # Deshabilitar el botón de conversión a Outcomes
        boton_convertir_word.config(state=tk.DISABLED)  # Deshabilitar el botón de conversión a Word si no se selecciona un archivo

def convertir_a_outcomes():
    archivo = ruta_label.cget("text").split(":\n")[1]  # Obtener la ruta del archivo seleccionado desde la etiqueta
    nombre_archivo = os.path.basename(archivo).split('.')[0]

    # Ejecutar la función de crear carpeta
    ruta_destino = create_folder_and_subfolder(nombre_archivo)

    # Ejecutar la función que convierte excel a csv_outcomes
    response = excel_to_outcomes(archivo, ruta_destino, nombre_archivo)
    nombre_archivo_response = response.split('_outcomes')[0]

    # Imprimir mensaje si todo va bien o si se encontró algún error
    if nombre_archivo == nombre_archivo_response:
        etiqueta.config(text=f"La conversión a outcomes csv ha finalizado.\nPuedes ubicarlo en la ruta: {ruta_destino}\nCon el nombre de: {response}\nGracias. !!!! :3")
    else:
        etiqueta.config(text=f"Advertencia: {response}\nRevisa el excel y vuelve a intentarlo, Gracias :)")

def convertir_a_word():
    archivo = ruta_label.cget("text").split(":\n")[1]  # Obtener la ruta del archivo seleccionado desde la etiqueta
    nombre_archivo = os.path.basename(archivo).split('.')[0]

    # Ejecutar la función de crear carpeta
    ruta_destino = create_folder_and_subfolder(nombre_archivo)
 
    # Ejecutar la función que convierte excel a Word
    response = excel_to_word(archivo, ruta_destino)

    if response == 'ok':
        etiqueta.config(text=f"La conversión a Word ha finalizado.\nPuedes ubicarlo en la ruta: {ruta_destino}\nGracias. !!!! :3")
    else:
        etiqueta.config(text=f"Advertencia: {response}\nRevisalo y vuelve a intentarlo, Gracias :)")

# Crear una ventana con un tamaño específico
ventana = tk.Tk()
ventana.title("Conversor Excel a Outcomes y Word")
ventana.geometry("600x450")  # Ancho x Alto

# Crear un botón para seleccionar el archivo
boton_seleccionar = tk.Button(ventana, text="Seleccionar archivo", command=seleccionar_archivo)
boton_seleccionar.pack(pady=20)

# Etiqueta para mostrar la ruta del archivo seleccionado con saltos de línea
ruta_label = tk.Label(ventana, text="Ruta del archivo seleccionado:\nNinguna", wraplength=500, justify=tk.LEFT)
ruta_label.pack(pady=10)

# Crear un marco para los botones de conversión
frame_botones = tk.Frame(ventana)
frame_botones.pack(pady=10)

# Crear un botón para convertir a Outcomes (inicialmente deshabilitado)
boton_convertir_outcomes = tk.Button(frame_botones, text="Convertir a Outcomes", command=convertir_a_outcomes, state=tk.DISABLED)
boton_convertir_outcomes.pack(side=tk.LEFT, padx=5)

# Crear un botón para convertir a Word (inicialmente deshabilitado)
boton_convertir_word = tk.Button(frame_botones, text="Convertir a Word", command=convertir_a_word, state=tk.DISABLED)
boton_convertir_word.pack(side=tk.LEFT, padx=5)

# Etiqueta para mostrar mensajes de resultados
etiqueta = tk.Label(ventana, text="", wraplength=500, justify=tk.CENTER)
etiqueta.pack(pady=10)

# Ejecutar el bucle principal de la interfaz gráfica
ventana.mainloop()

