import pandas as pd
import tkinter as tk
from tkinter import ttk

# Leer el archivo XLS
file_path = r'C:\Users\User\Desktop\App.xlsx'
data = pd.read_excel(file_path)

# Crear la ventana principal
root = tk.Tk()
root.title('Cotizador Aspen')

# Ajustar el tamaño de la ventana (ancho x alto)
root.geometry("600x400")

# Configurar las columnas para que se expandan
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)

# Etiqueta de búsqueda
search_label = tk.Label(root, text="Buscar por Código o Artículo")
search_label.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

# Campo de entrada para la búsqueda
search_entry = tk.Entry(root)
search_entry.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

# Crear un frame para contener el Listbox y el Scrollbar
listbox_frame = tk.Frame(root)
listbox_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

# Listbox para mostrar las opciones de búsqueda
listbox = tk.Listbox(listbox_frame, height=4, width=70, selectmode=tk.SINGLE)
listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Scrollbar para el Listbox
scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Vincular el Scrollbar al Listbox
listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# Etiqueta para mostrar los resultados seleccionados
result_label = tk.Label(root, text="", fg="blue", wraplength=500)
result_label.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

# Función para actualizar la búsqueda en tiempo real
def search(event):
    search_text = search_entry.get()
    
    # Limpiar la lista actual de resultados
    listbox.delete(0, tk.END)
    
    if not search_text.strip():  # Si el campo de búsqueda está vacío
        listbox.insert(tk.END, "Indica arriba la moto por la que consultas.")
        listbox.itemconfig(0, {'fg': 'light grey'})  # Color gris claro
    else:
        result = data[(data['CODIGO'].astype(str).str.contains(search_text, case=False, na=False)) | 
                      (data['ARTICULO'].str.contains(search_text, case=False, na=False))]
        
        if not result.empty:
            for index, row in result.iterrows():
                articulo = row['ARTICULO']
                codigo = row['CODIGO']
                # Mostrar "Código - Artículo" en el listbox
                listbox.insert(tk.END, f"{codigo} - {articulo}")
        else:
            listbox.insert(tk.END, "Artículo no encontrado.")

# Función para mostrar los detalles del artículo seleccionado
def show_selected(event):
    selected_index = listbox.curselection()  # Obtener el índice del elemento seleccionado
    if selected_index:
        selected_item = listbox.get(selected_index)  # Obtener el texto del artículo seleccionado
        codigo = selected_item.split(' - ')[0]  # Extraer el código del artículo
        
        # Buscar el artículo seleccionado en los datos originales
        result = data[data['CODIGO'].astype(str) == codigo]
        
        if not result.empty:
            # Extraer los datos necesarios
            articulo = result.iloc[0]['ARTICULO']
            precio_efectivo = result.iloc[0]['PRECIO EFECTIVO']
            
            # Formatear el precio con separador de miles como punto
            precio_efectivo_formatted = "${:,.0f}".format(precio_efectivo).replace(',', '.')
            
            # Mostrar los resultados
            result_text = f"{articulo} {precio_efectivo_formatted} precio contado efectivo. Casco + Formulario 01."
            result_label.config(text=result_text)

# Función para copiar el texto al portapapeles
def copy_to_clipboard():
    result_text = result_label.cget("text")
    if result_text:  # Solo copiar si hay texto
        root.clipboard_clear()  # Limpiar el portapapeles
        root.clipboard_append(result_text)  # Agregar el texto
        root.update()  # Actualizar el portapapeles

# Vincular la función de búsqueda con el evento de liberación de tecla
search_entry.bind('<KeyRelease>', search)

# Vincular la selección del Listbox para mostrar detalles del artículo
listbox.bind('<<ListboxSelect>>', show_selected)

# Botón para copiar el texto
copy_button = tk.Button(root, text="Copiar", command=copy_to_clipboard)
copy_button.grid(row=3, column=2, padx=10, pady=10, sticky='ew')

# Mostrar mensaje inicial al cargar la ventana
listbox.insert(tk.END, "Indica arriba la moto por la que consultas.")
listbox.itemconfig(0, {'fg': 'light grey'})  # Color gris claro

# Iniciar la aplicación
root.mainloop()
