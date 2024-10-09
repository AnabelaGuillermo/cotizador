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

# Función que se ejecutará al hacer clic en el botón de búsqueda
def search():
    search_text = search_entry.get()
    result = data[(data['CODIGO'] == search_text) | (data['ARTICULO'].str.contains(search_text, case=False, na=False))]
    if not result.empty:
        print(result)
    else:
        print("Artículo no encontrado.")

# Botón de búsqueda
search_button = tk.Button(root, text="Buscar", command=search)
search_button.grid(row=1, column=1, padx=10, pady=10)

# Iniciar la aplicación
root.mainloop()
