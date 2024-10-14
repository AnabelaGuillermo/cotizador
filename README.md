# Cotizador

Este proyecto es una aplicación de escritorio en Python que permite a los usuarios cotizar precios y opciones de financiación de productos a partir de un archivo Excel. Utiliza la biblioteca `Tkinter` para la interfaz gráfica y `Pandas` para la manipulación de datos.

## Requisitos

Para ejecutar este proyecto, necesitas tener Python instalado en tu máquina. Además, es recomendable usar un entorno virtual para manejar las dependencias. Puedes crear un entorno virtual utilizando `venv`.

### Instalación de Dependencias

1. Clona este repositorio:
   git clone https://github.com/AnabelaGuillermo/cotizador.git
2. Crea un entorno virtual:
   python -m venv venv
3. Activa el entorno virtual:
   venv\Scripts\activate
4. Instala las dependencias:
   pip install pandas openpyxl

### Archivos necesarios

1. config.txt: Este archivo se utiliza para almacenar la ruta del último archivo Excel seleccionado.
2. logoAspenCotizador.png: Logo que se muestra en la interfaz gráfica. Puedes modificar su nombre pero si lo haces recuerda modificarlo en el código.
3. Archivo Excel: Asegúrate de tener un archivo Excel con las columnas como las del ejemplo cargado en este repositorio.

## Uso

1. Ejecuta la aplicación:

   python main.py

   Se abrirá una ventana donde podrás seleccionar un archivo Excel (si no hay uno guardado en config.txt).
3. Ingresa un código o artículo en el campo de búsqueda para filtrar los resultados.
4. Selecciona un artículo de la lista para ver su precio y opciones de financiación.
5. Ingresa un anticipo si deseas calcular las cuotas de financiación.
6. Puedes copiar la información mostrada en el cuadro de resultados al portapapeles haciendo clic en el botón "Copiar información".

## Funcionalidades

1. Búsqueda de productos por código o nombre.
2. Cálculo de opciones de financiación basadas en el anticipo ingresado.
3. Interfaz gráfica amigable para facilitar la navegación y selección de productos.

## Contribuciones

Las contribuciones son bienvenidas. Siéntete libre de abrir un problema o enviar un pull request para mejorar el proyecto.
