Prueba RPA

Este repositorio contiene el proyecto prueba_rpa, desarrollado para realizar tareas de automatización y scraping utilizando Python, de la página del DANE.
Tiene como finalidad la descarga y extracción de información del archivo que se encuentra dentro de la página

## Contenido

El proyecto incluye los siguientes archivos principales:

prueba_scrapping.py: Script principal que contiene la lógica del scraping.
requirements.txt: Archivo con las dependencias necesarias para ejecutar el proyecto.
.gitignore: Lista de archivos y directorios ignorados por Git.

## Requisitos previos

1. Tener Python instalado en tu máquina (se recomienda la versión 3.8 o superior).
2. Contar con `pip` para manejar las dependencias de Python.
3. Disponer de un entorno virtual configurado.


## Instalación

1. Clona este repositorio en tu máquina local:

   ```bash
   git clone git@github.com:SebsMontero/prueba_rpa.git

   cd prueba_rpa

   .\venv\Scripts\activate

   source venv/bin/activate

   pip install -r requirements.txt


Uso
Asegúrate de tener el entorno virtual activado.

Ejecuta el script principal:

python prueba_scrapping.py

Estructura del proyecto:

prueba_rpa/
│
├── .gitignore             # Archivos ignorados por Git
├── prueba_scrapping.py    # Script principal del scraping
├── requirements.txt       # Dependencias del proyecto
├── resultados_procesados/ # Carpeta para los resultados generados
└── evidencias/            # Carpeta para almacenar evidencias

## Instalación

Importante hacer la configuración del correo al cuál se debe hacer el envío del correo en la siguiente línea

destinatario = "ingjaviermonterot@gmail.com"
        asunto = "Resumen de Ventas - Top 10 Productos"
        correo = Correo(remitente, contraseña)
        correo.enviar(destinatario, asunto, resumen, top_10_path)

Esto, para garantizar el envío del correo con la información solicitada, dado que en la documentación, no se indica a qué correo se debe realizar el envío.
