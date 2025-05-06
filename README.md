# ST-HONORE
TF Saint Honoré - YOEC
TF Saint Honoré es una aplicación web diseñada para optimizar el proceso de gestión de información en el entorno comercial de productos y ventas. Este proyecto permite a los usuarios interactuar con datos de productos de manera eficiente, procesando información directamente desde archivos PDF y generando resultados procesados en formato Excel.

Funcionalidades
Carga de archivos PDF: Los usuarios pueden subir archivos PDF con datos de productos para procesarlos automáticamente.

Procesamiento de datos: La aplicación extrae información de los PDF, como números de factura, descripciones de productos, cantidades y precios, entre otros.

Generación de archivo Excel: Los datos extraídos son organizados y convertidos en un archivo Excel descargable para su análisis posterior.

Optimización de procesos: La aplicación ayuda a agilizar el flujo de trabajo, eliminando la necesidad de ingresar manualmente la información desde los archivos PDF.

Tecnologías
Este proyecto está construido utilizando las siguientes tecnologías:

Frontend:

HTML

CSS (para estilo y diseño responsivo)

JavaScript (para interactividad y manejo de eventos)

React (para la construcción de la interfaz de usuario dinámica)

Backend:

Python (con Flask para la creación de la API)

PDFplumber (para la extracción de texto y datos de los archivos PDF)

Openpyxl (para la generación del archivo Excel)

Despliegue:

Vercel para el despliegue de la aplicación web en la nube.

Cómo usarlo
Sube el archivo PDF: Usa el botón de carga de archivos en la página para seleccionar un archivo PDF que contenga los datos de los productos.

Procesamiento automático: La aplicación procesará automáticamente el archivo, extrayendo la información relevante.

Descarga el archivo Excel: Una vez procesados los datos, podrás descargar el archivo Excel generado.

Instalación
Para correr este proyecto localmente, sigue estos pasos:

Clona el repositorio:

bash
Copiar
git clone https://github.com/tu-usuario/tf-sthonore-yoec.git
Instala las dependencias:

bash
Copiar
pip install -r requirements.txt
Corre el servidor de desarrollo:

bash
Copiar
python app.py
Abre tu navegador y visita http://localhost:5000.

Contribuciones
Si deseas contribuir a este proyecto, sigue estos pasos:

Haz un fork del repositorio.

Crea una nueva rama (git checkout -b nueva-rama).

Realiza tus cambios.

Haz un commit de tus cambios (git commit -am 'Descripción de los cambios').

Sube tu rama (git push origin nueva-rama).

Abre un pull request.
