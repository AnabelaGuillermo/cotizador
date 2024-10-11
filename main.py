import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
import os

# Función para leer la ruta del archivo desde un archivo de configuración
def read_file_path():
    if os.path.exists('config.txt'):
        with open('config.txt', 'r') as file:
            return file.read().strip()
    return ''

# Función para guardar la ruta del archivo en un archivo de configuración
def save_file_path(file_path):
    with open('config.txt', 'w') as file:
        file.write(file_path)

# Leer la ruta del archivo desde el archivo de configuración
file_path = read_file_path()

# Si no hay una ruta guardada, pedir al usuario que la seleccione
if not file_path:
    file_path = simpledialog.askstring("Ruta del archivo", "Introduce la ruta del archivo Excel:", initialvalue=r'C:\Users\User\Desktop\App.xlsx')
    if file_path:
        save_file_path(file_path)

data = pd.read_excel(file_path)

# Crear la ventana principal
root = tk.Tk()
root.title('Cotizador Aspen')

# Ajustar el tamaño de la ventana (ancho x alto)
root.geometry("800x600")

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
result_label.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

# Etiqueta y campo para el anticipo
anticipo_label = tk.Label(root, text="Anticipo:")
anticipo_label.grid(row=2, column=0, padx=10, pady=10, sticky='w')

anticipo_entry = tk.Entry(root)
anticipo_entry.grid(row=2, column=1, padx=10, pady=10, sticky='ew')

# Funciones para las opciones de financiación
financing_options = ["VISA/MASTERCARD", "NARANJA", "SUCREDITO", "SOL"]

# Crear checkboxes para las opciones de financiación
checkbox_vars = {}
row_index = 4
for option in financing_options:
    var = tk.BooleanVar()
    checkbox_vars[option] = var
    checkbox = tk.Checkbutton(root, text=option, variable=var, command=lambda: update_results())
    checkbox.grid(row=row_index, column=0, padx=10, sticky='w')
    row_index += 1

# Función para calcular los precios de las cuotas con anticipo
def calculate_financing(precio_lista, anticipo):
    results = []
    precio_lista -= anticipo  # Restar el anticipo al precio de lista
    
    # Asegurarse de que el precio sea al menos 0
    precio_lista = max(0, precio_lista)
    
    # Mostrar todas las cuotas seleccionadas por tarjeta
    for option in financing_options:
        if checkbox_vars[option].get():
            if option == "VISA/MASTERCARD":
                results.append("VISA/MASTERCARD:")
                for cuotas in ["3 CUOTAS", "6 CUOTAS", "12 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"VISA/MASTERCARD BANCO {cuotas}"].values[0]
                    cuota = (precio_lista * financing_multiplier) / int(cuotas.split()[0])
                    results.append(f"  {cuotas} de ${cuota:.2f}")
                results.append("")  # Agregar una línea en blanco después de VISA/MASTERCARD
            elif option == "NARANJA":
                results.append("NARANJA:")
                for cuotas in ["PLAN Z 3 CUOTAS", "6 CUOTAS", "10 CUOTAS", "12 CUOTAS", "18 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"NARANJA {cuotas}"].values[0]
                    cuota = (precio_lista * financing_multiplier) / int(cuotas.split()[-2])
                    results.append(f"  {cuotas} de ${cuota:.2f}")
                results.append("")  # Agregar una línea en blanco después de NARANJA
            elif option == "SUCREDITO":
                results.append("SUCREDITO:")
                for cuotas in ["3 CUOTAS", "6 CUOTAS", "12 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"SUCREDITO {cuotas}"].values[0]
                    cuota = (precio_lista * financing_multiplier) / int(cuotas.split()[0])
                    results.append(f"  {cuotas} de ${cuota:.2f}")
                results.append("")  # Agregar una línea en blanco después de SUCREDITO
            elif option == "SOL":
                results.append("SOL:")
                financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, "SOL 12 CUOTAS"].values[0]
                cuota = (precio_lista * financing_multiplier) / 12
                results.append(f"  12 CUOTAS de ${cuota:.2f}")
                results.append("")  # Agregar una línea en blanco después de SOL
    
    return results

# Función para mostrar los detalles del artículo seleccionado
def show_selected(event):
    global selected_codigo  # Variable global para acceder en otras funciones
    selected_index = listbox.curselection()  # Obtener el índice del elemento seleccionado
    if selected_index:
        selected_item = listbox.get(selected_index)  # Obtener el texto del artículo seleccionado
        selected_codigo = selected_item.split(' - ')[0]  # Extraer el código del artículo
        
        # Buscar el artículo seleccionado en los datos originales
        result = data[data['CODIGO'].astype(str) == selected_codigo]
        
        if not result.empty:
            # Extraer los datos necesarios
            articulo = result.iloc[0]['ARTICULO']
            precio_efectivo = result.iloc[0]['PRECIO EFECTIVO']
            precio_lista = result.iloc[0]['PRECIO LISTA']
            
            # Formatear el precio con separador de miles como punto
            precio_efectivo_formatted = "${:,.0f}".format(precio_efectivo).replace(',', '.')
            
            # Mostrar los resultados con un salto de línea después del artículo y precio efectivo
            result_text = f"{articulo} {precio_efectivo_formatted} precio contado efectivo. Casco + Formulario 01.\n\n"
            
            result_label.config(text=result_text)
            update_results()  # Llamar a la función de actualización de resultados

# Función para actualizar los resultados de cuotas cuando cambia el anticipo
def update_results():
    anticipo = anticipo_entry.get()
    anticipo = float(anticipo) if anticipo else 0
    result = data[data['CODIGO'].astype(str) == selected_codigo]
    if not result.empty:
        precio_lista = result.iloc[0]['PRECIO LISTA']
        financing_results = calculate_financing(precio_lista, anticipo)
        
        # Actualiza el texto del result_label en lugar de agregar más texto
        result_label.config(text=result_label.cget("text").split("\n\n")[0] + "\n\n" + "\n".join(financing_results))

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

# Función para copiar el texto al portapapeles
def copy_to_clipboard():
    result_text = result_label.cget("text")
    root.clipboard_clear()
    root.clipboard_append(result_text)

# Vincular la búsqueda con el evento KeyRelease
search_entry.bind("<KeyRelease>", search)

# Vincular la selección del Listbox con la función show_selected
listbox.bind('<<ListboxSelect>>', show_selected)

# Vincular la entrada del anticipo con la función de actualización de resultados
anticipo_entry.bind("<KeyRelease>", lambda event: update_results())

# Botón para copiar al portapapeles
copy_button = tk.Button(root, text="Copiar información", command=copy_to_clipboard)
copy_button.grid(row=row_index, column=0, padx=10, pady=10, sticky='w')

# Ejecutar el bucle principal de la ventana
root.mainloop()