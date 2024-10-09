import pandas as pd
import tkinter as tk
from tkinter import ttk

# Leer el archivo XLS
file_path = r'C:\Users\User\Desktop\App.xlsx'  # Ruta a tu archivo
data = pd.read_excel(file_path)

# Crear la ventana principal
root = tk.Tk()
root.title('Cotizador Aspen')

# Etiqueta de búsqueda
search_label = tk.Label(root, text="Buscar por Código o Artículo")
search_label.grid(row=0, column=0, padx=10, pady=10)

# Campo de entrada para la búsqueda
search_entry = tk.Entry(root)
search_entry.grid(row=0, column=1, padx=10, pady=10)

# Etiquetas para mostrar los resultados
result_label = tk.Label(root, text="", fg="blue")
result_label.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Función que se ejecutará al hacer clic en el botón de búsqueda
def search():
    search_text = search_entry.get()
    result = data[(data['CODIGO'] == search_text) | (data['ARTICULO'].str.contains(search_text, case=False, na=False))]
    
    if not result.empty:
        # Extraer los datos necesarios
        articulo = result.iloc[0]['ARTICULO']
        precio_efectivo = result.iloc[0]['PRECIO EFECTIVO']
        
        # Formatear el precio con separador de miles como punto
        precio_efectivo_formatted = "${:,.0f}".format(precio_efectivo).replace(',', '.')
        
        # Mostrar los resultados
        result_text = f"{articulo} {precio_efectivo_formatted} precio contado efectivo. Casco + Formulario 01."
        result_label.config(text=result_text)
    else:
        result_label.config(text="Artículo no encontrado.")

# Función para copiar el texto al portapapeles
def copy_to_clipboard():
    result_text = result_label.cget("text")
    if result_text:  # Solo copiar si hay texto
        root.clipboard_clear()  # Limpiar el portapapeles
        root.clipboard_append(result_text)  # Agregar el texto
        root.update()  # Actualizar el portapapeles

# Botón de búsqueda
search_button = tk.Button(root, text="Buscar", command=search)
search_button.grid(row=1, column=1, padx=10, pady=10)

# Botón para copiar el texto
copy_button = tk.Button(root, text="Copiar", command=copy_to_clipboard)
copy_button.grid(row=3, column=2, padx=10, pady=10)

# Iniciar la aplicación
root.mainloop()
