import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from PIL import Image, ImageTk
import os

def read_file_path():
    if os.path.exists('config.txt'):
        with open('config.txt', 'r') as file:
            return file.read().strip()
    return ''

def save_file_path(file_path):
    with open('config.txt', 'w') as file:
        file.write(file_path)

file_path = read_file_path()
if not file_path:
    file_path = simpledialog.askstring("Ruta del archivo", "Introduce la ruta del archivo Excel:", initialvalue=r'C:\Users\User\Desktop\App.xlsx')
    if file_path:
        save_file_path(file_path)

data = pd.read_excel(file_path)

root = tk.Tk()
root.title('Cotizador Aspen')
root.geometry("800x600")

favicon_image = Image.open('assets/logoAspenCotizador.png')
favicon_image = favicon_image.resize((32, 32), Image.LANCZOS)
favicon_photo = ImageTk.PhotoImage(favicon_image)
root.iconphoto(False, favicon_photo)

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)

logo_image = Image.open('assets/logoAspenCotizador.png')
logo_image = logo_image.resize((120, 80), Image.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(root, image=logo_photo)
logo_label.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky='n')

search_label = tk.Label(root, text="Buscar por Código o Artículo")
search_label.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

search_entry = tk.Entry(root)
search_entry.grid(row=1, column=1, padx=10, pady=10, sticky='ew')

listbox_frame = tk.Frame(root)
listbox_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

listbox = tk.Listbox(listbox_frame, height=4, width=70, selectmode=tk.SINGLE)
listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# Cambiado el result_label por un Text
result_frame = tk.Frame(root)
result_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

result_text = tk.Text(result_frame, height=10, width=70, wrap=tk.WORD)
result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

result_scrollbar = tk.Scrollbar(result_frame, orient=tk.VERTICAL, command=result_text.yview)
result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
result_text.config(yscrollcommand=result_scrollbar.set)

anticipo_label = tk.Label(root, text="Anticipo:")
anticipo_label.grid(row=4, column=0, padx=10, pady=10, sticky='w')

anticipo_entry = tk.Entry(root)
anticipo_entry.grid(row=4, column=1, padx=10, pady=10, sticky='ew')

financing_options = ["VISA/MASTERCARD", "NARANJA", "SUCREDITO", "SOL"]

checkbox_vars = {}
row_index = 5
for option in financing_options:
    var = tk.BooleanVar()
    checkbox_vars[option] = var
    checkbox = tk.Checkbutton(root, text=option, variable=var, command=lambda: update_results())
    checkbox.grid(row=row_index, column=0, padx=10, sticky='w')
    row_index += 1

def format_currency(value):
    return "{:,.0f}".format(value).replace(',', 'X').replace('.', ',').replace('X', '.')

def calculate_financing(precio_lista, anticipo):
    results = []
    precio_lista -= anticipo
    precio_lista = max(0, precio_lista)
    
    for option in financing_options:
        if checkbox_vars[option].get():
            if option == "VISA/MASTERCARD":
                results.append("VISA/MASTERCARD:")
                for cuotas in ["3 CUOTAS", "6 CUOTAS", "12 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"VISA/MASTERCARD BANCO {cuotas}"].values[0]
                    cuota = round((precio_lista * financing_multiplier) / int(cuotas.split()[0]))
                    results.append(f"  {cuotas} de ${format_currency(cuota)}")
                results.append("")
            elif option == "NARANJA":
                results.append("NARANJA:")
                for cuotas in ["PLAN Z 3 CUOTAS", "6 CUOTAS", "10 CUOTAS", "12 CUOTAS", "18 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"NARANJA {cuotas}"].values[0]
                    cuota = round((precio_lista * financing_multiplier) / int(cuotas.split()[-2]))
                    results.append(f"  {cuotas} de ${format_currency(cuota)}")
                results.append("")
            elif option == "SUCREDITO":
                results.append("SUCREDITO:")
                for cuotas in ["3 CUOTAS", "6 CUOTAS", "12 CUOTAS"]:
                    financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, f"SUCREDITO {cuotas}"].values[0]
                    cuota = round((precio_lista * financing_multiplier) / int(cuotas.split()[0]))
                    results.append(f"  {cuotas} de ${format_currency(cuota)}")
                results.append("")
            elif option == "SOL":
                results.append("SOL:")
                financing_multiplier = data.loc[data['CODIGO'] == selected_codigo, "SOL 12 CUOTAS"].values[0]
                cuota = round((precio_lista * financing_multiplier) / 12)
                results.append(f"  12 CUOTAS de ${format_currency(cuota)}")
                results.append("")
    
    return results

def show_selected(event):
    global selected_codigo
    selected_index = listbox.curselection()
    if selected_index:
        selected_item = listbox.get(selected_index)
        selected_codigo = selected_item.split(' - ')[0]
        
        result = data[data['CODIGO'].astype(str) == selected_codigo]
        if not result.empty:
            articulo = result.iloc[0]['ARTICULO']
            precio_efectivo = result.iloc[0]['PRECIO EFECTIVO']
            precio_lista = result.iloc[0]['PRECIO LISTA']
            precio_efectivo_formatted = "${:,.0f}".format(precio_efectivo).replace(',', '.')
            result_text_content = f"{articulo} ${precio_efectivo_formatted} precio contado efectivo. Casco + Formulario 01.\n\n"
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, result_text_content)
            update_results()

def update_results():
    anticipo = anticipo_entry.get()
    anticipo = round(float(anticipo)) if anticipo else 0
    result = data[data['CODIGO'].astype(str) == selected_codigo]
    
    if not result.empty:
        precio_lista = result.iloc[0]['PRECIO LISTA']
        articulo = result.iloc[0]['ARTICULO']
        precio_efectivo = result.iloc[0]['PRECIO EFECTIVO']
        precio_efectivo_formatted = format_currency(round(precio_efectivo))

        result_text_content = f"{articulo} ${precio_efectivo_formatted} precio contado efectivo. Casco + Formulario 01.\n\n"
        
        if anticipo > 0:
            result_text_content += f"Anticipo de ${format_currency(anticipo)} +\n\n"

        financing_results = calculate_financing(precio_lista, anticipo)

        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, result_text_content)
        
        if any(var.get() for var in checkbox_vars.values()):
            result_text.insert(tk.END, "\n".join(financing_results))

def search(event):
    search_text = search_entry.get()
    listbox.delete(0, tk.END)
    
    if not search_text.strip():
        listbox.insert(tk.END, "Indica arriba la moto por la que consultas.")
        listbox.itemconfig(0, {'fg': 'light grey'})
    else:
        result = data[(data['CODIGO'].astype(str).str.contains(search_text, case=False, na=False)) | 
                      (data['ARTICULO'].str.contains(search_text, case=False, na=False))]
        
        if not result.empty:
            for index, row in result.iterrows():
                articulo = row['ARTICULO']
                codigo = row['CODIGO']
                listbox.insert(tk.END, f"{codigo} - {articulo}")
        else:
            listbox.insert(tk.END, "Artículo no encontrado.")

def copy_to_clipboard():
    result_text_content = result_text.get(1.0, tk.END)
    root.clipboard_clear()
    root.clipboard_append(result_text_content)

search_entry.bind("<KeyRelease>", search)
listbox.bind('<<ListboxSelect>>', show_selected)
anticipo_entry.bind("<KeyRelease>", lambda event: update_results())

copy_button = tk.Button(root, text="Copiar información", command=copy_to_clipboard)
copy_button.grid(row=row_index, column=0, padx=10, pady=10, sticky='w')

root.mainloop()