# Configuración de entorno virtual

import os
import sys

def isRunningFromEXE():
    try:
        return getattr(sys, 'frozen', False)
    except:
        return False

def getPath():
    if isRunningFromEXE():
        getSctiptPath = os.path.abspath(sys.executable)
    else:
        getSctiptPath = os.path.abspath(__file__)
    scriptPath = os.path.dirname(getSctiptPath)
    return scriptPath

customPath = os.path.join(getPath(),'Lib','site-packages')
sys.path.insert(0, customPath)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import pandas as pd
import smtplib
from email.message import EmailMessage

# -------------------------------------- DESCARGA ARCHIVO DANE --------------------------------

class ArchivoDescarga:
    def __init__(self, driver_path, download_dir, evidencias_dir):
        """Inicializar la clase con los directorios necesarios."""
        self.download_dir = download_dir
        self.evidencias_dir = evidencias_dir

        os.makedirs(self.download_dir, exist_ok=True)
        os.makedirs(self.evidencias_dir, exist_ok=True)

        service = Service(driver_path)
        options = webdriver.ChromeOptions()

        prefs = {"download.default_directory": self.download_dir}
        options.add_experimental_option("prefs", prefs)

        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        self.driver = webdriver.Chrome(service=service, options=options)

    def limpiar_carpeta_descargas(self):
        """Eliminar todos los archivos en la carpeta de descargas."""
        print(f"Limpiando la carpeta: {self.download_dir}")
        for archivo in os.listdir(self.download_dir):
            archivo_path = os.path.join(self.download_dir, archivo)
            if os.path.isfile(archivo_path):
                os.remove(archivo_path)
                print(f"Archivo eliminado: {archivo_path}")

    def abrir_pagina(self, url):
        """Abrir la página especificada."""
        print("Abriendo la página...")
        self.driver.maximize_window()
        self.driver.get(url)
        print("Página cargada.")
        time.sleep(5)

    def buscar_texto_y_tomar_pantallazo(self, texto_xpath, nombre_pantallazo):
        """Buscar el texto, desplazarse y tomar un pantallazo."""
        try:
            print("Buscando el texto...")
            wait = WebDriverWait(self.driver, 10)
            section_title = wait.until(EC.presence_of_element_located((By.XPATH, texto_xpath)))
            print("Texto encontrado: 'Precios de los productos de primera necesidad para los colombianos en tiempos del COVID-19'")

            print("Desplazándose al texto...")
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", section_title)
            time.sleep(2)

            ruta_pantallazo = os.path.join(self.evidencias_dir, nombre_pantallazo)
            print("Tomando un pantallazo del texto encontrado...")
            self.driver.save_screenshot(ruta_pantallazo)
            print(f"Pantallazo guardado en: {ruta_pantallazo}")
        except Exception as e:
            print("Error buscando el texto o tomando el pantallazo:", e)
            raise

    def buscar_y_descargar(self, boton_xpath):
        """Buscar el botón y descargar el archivo."""
        try:
            print("Buscando el botón...")
            button = self.driver.find_element(By.XPATH, boton_xpath)
            print("Botón encontrado. Desplazándose al botón...")
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", button)
            time.sleep(2)

            print("Haciendo clic en el botón para iniciar la descarga...")
            button.click()
            time.sleep(5) 

            print("Accediendo a chrome://downloads/ para monitorear la descarga...")
            self.driver.get("chrome://downloads/")
            time.sleep(2)

            ruta_pantallazo_descarga = os.path.join(self.evidencias_dir, "pantallazo_descarga.png")
            print("Tomando un pantallazo de la página de descargas...")
            self.driver.save_screenshot(ruta_pantallazo_descarga)
            print(f"Pantallazo de descargas guardado en: {ruta_pantallazo_descarga}")

            print("Descarga monitoreada exitosamente desde chrome://downloads/.")

        except Exception as e:
            print("Error buscando o haciendo clic en el botón:", e)
            raise

    def cerrar(self):
        """Cerrar el navegador."""
        self.driver.quit()

# --------------------------------------- PROCESAMIENTO DATA -------------------------------------

class ProcesoData:
    def __init__(self, base_path):
        self.base_path = base_path
        self.descargas_path = os.path.join(base_path, "anexo_ref_mas_vendidas")
        self.resultados_path = os.path.join(base_path, "resultados_procesados")
        
        print(f"Descargas path inicializado en: {self.descargas_path}")
        print(f"Resultados path inicializado en: {self.resultados_path}")

        os.makedirs(self.descargas_path, exist_ok=True)
        os.makedirs(self.resultados_path, exist_ok=True)

    def obtener_ultimo_archivo(self):
        """Obtiene el archivo más reciente en la carpeta de descargas."""
        archivos = [os.path.join(self.descargas_path, f) for f in os.listdir(self.descargas_path) if f.endswith('.xlsx')]
        if not archivos:
            raise FileNotFoundError(f"No hay archivos en la carpeta de descargas: {self.descargas_path}")
        return max(archivos, key=os.path.getctime)

    def procesar_archivo(self, archivo=None):
        """Procesa el archivo más reciente o uno específico."""
        if archivo is None:
            archivo = self.obtener_ultimo_archivo()

        print(f"Procesando el archivo: {archivo}")
        xls = pd.ExcelFile(archivo)

        hoja = "Cantidades 1203-1603"
        if hoja not in xls.sheet_names:
            raise ValueError(f"La hoja '{hoja}' no existe en el archivo.")

        print(f"\nProcesando hoja: {hoja}")
        try:
            raw_data = pd.read_excel(archivo, sheet_name=hoja, header=None)
            encabezado_fila = None
            for idx, row in raw_data.iterrows():
                if 'Nombre DANE' in row.values and 'Código de barras' in row.values:
                    encabezado_fila = idx
                    break

            if encabezado_fila is None:
                raise ValueError(f"No se pudo detectar una fila de encabezados válida en la hoja '{hoja}'.")

            print(f"Encabezados detectados en la fila: {encabezado_fila}")
            df = pd.read_excel(archivo, sheet_name=hoja, header=encabezado_fila)
            df.columns = df.columns.str.strip()
            print(f"Columnas detectadas (limpiadas): {df.columns.tolist()}")

            columnas_necesarias = ['Nombre producto', 'Marca', 'Precio Reportado', 'Cantidades vendidas']
            if not all(col in df.columns for col in columnas_necesarias):
                raise ValueError(f"La hoja '{hoja}' no contiene las columnas necesarias: {', '.join(columnas_necesarias)}.")

            df['Cantidades vendidas'] = pd.to_numeric(df['Cantidades vendidas'], errors='coerce')
            df['Precio Reportado'] = pd.to_numeric(df['Precio Reportado'], errors='coerce')

            top_10 = df.nlargest(10, 'Cantidades vendidas')[['Nombre producto', 'Marca', 'Cantidades vendidas', 'Precio Reportado']]
            print(f"\nLos 10 productos más vendidos en la hoja '{hoja}' son:")
            print(top_10)

            self.guardar_resultados(hoja, top_10)


        except Exception as e:
            print(f"Error procesando la hoja '{hoja}': {e}")

    def guardar_resultados(self, hoja, datos):
        """Guarda los resultados procesados en un archivo CSV con una fila adicional para el total."""
        try:
            if datos is None or datos.empty:
                print(f"No hay datos procesados para guardar de la hoja '{hoja}'.")
                return

            total_cantidades = datos['Cantidades vendidas'].sum()

            datos['Total Precio'] = datos['Cantidades vendidas'] * datos['Precio Reportado']
            total_precio_multiplicado = datos['Total Precio'].sum()

            total_row = pd.DataFrame([{
                'Nombre producto': 'TOTAL',
                'Marca': '',
                'Cantidades vendidas': total_cantidades,
                'Total Precio': total_precio_multiplicado
            }])

            datos = pd.concat([datos, total_row], ignore_index=True)

            datos = datos.drop(columns=['Precio Reportado'])

            output_path = os.path.join(self.resultados_path, f"resultados_{hoja}.csv")
            datos.to_csv(output_path, index=False)
            print(f"Resultados guardados en: {output_path}")
        except Exception as e:
            print(f"Error al guardar los resultados de la hoja '{hoja}': {e}")

    def calcular_resumen(self, archivo, datos_top_10):
        """Calcula y resume la información requerida."""
        try:
            print("\nCalculando y resumiendo la información...")

            hoja = "Cantidades 1203-1603"
            df_completo = pd.read_excel(archivo, sheet_name=hoja, header=7)

            df_completo.columns = df_completo.columns.str.strip()

            if 'Cantidades vendidas' not in df_completo.columns:
                raise ValueError("La columna 'Cantidades vendidas' no existe en el archivo original después de la limpieza.")

            df_completo['Cantidades vendidas'] = pd.to_numeric(df_completo['Cantidades vendidas'], errors='coerce')
            total_todos_productos = df_completo['Cantidades vendidas'].sum()
            datos_top_10.columns = datos_top_10.columns.str.strip()
            datos_top_10_sin_total = datos_top_10[datos_top_10['Nombre producto'] != 'TOTAL']
            print("\nDatos del Top 10 para cálculo (sin fila TOTAL):")
            print(datos_top_10_sin_total[['Nombre producto', 'Cantidades vendidas']])
            total_top_10 = datos_top_10_sin_total['Cantidades vendidas'].sum()
            porcentaje_top_10 = (total_top_10 / total_todos_productos) * 100 if total_todos_productos else 0
            total_todos_productos = round(total_todos_productos, 2)
            total_top_10 = round(total_top_10, 2)
            porcentaje_top_10 = round(porcentaje_top_10, 2)
            print("\nResumen de la información:")
            print(f"- Total de todos los productos vendidos: {total_todos_productos:.2f}")
            print(f"- Total de los 10 productos más vendidos: {total_top_10:.2f}")
            print(f"- Porcentaje de los 10 productos más vendidos respecto al total: {porcentaje_top_10:.2f}%")
            resumen_path = os.path.join(self.resultados_path, "resumen_informacion.txt")
            with open(resumen_path, 'w') as f:
                f.write("Resumen de la información:\n")
                f.write(f"- Total de todos los productos vendidos: {total_todos_productos:.2f}\n")
                f.write(f"- Total de los 10 productos más vendidos: {total_top_10:.2f}\n")
                f.write(f"- Porcentaje de los 10 productos más vendidos respecto al total: {porcentaje_top_10:.2f}%\n")
            print(f"\nResumen guardado en: {resumen_path}")


        except Exception as e:
            print("Error al calcular el resumen:", e)

# ----------------------------------- ENVÍO DE CORREO --------------------------------------------

class Correo:
    def __init__(self, remitente, contraseña, servidor_smtp="smtp.gmail.com", puerto=465):
        """
        Inicializa la clase Correo para enviar correos electrónicos.
        :param remitente: Dirección de correo electrónico del remitente.
        :param contraseña: Contraseña de la aplicación generada por Google.
        :param servidor_smtp: Servidor SMTP (por defecto smtp.gmail.com).
        :param puerto: Puerto del servidor SMTP (por defecto 465 para SSL).
        """
        self.remitente = remitente
        self.contraseña = contraseña
        self.servidor_smtp = servidor_smtp
        self.puerto = puerto

    def enviar(self, destinatario, asunto, cuerpo, archivo_adjunto=None):
        """
        Envía un correo con o sin archivo adjunto.
        :param destinatario: Dirección de correo del destinatario.
        :param asunto: Asunto del correo.
        :param cuerpo: Cuerpo del correo.
        :param archivo_adjunto: Ruta al archivo a adjuntar (opcional).
        """
        try:
            email = EmailMessage()
            email["From"] = self.remitente
            email["To"] = destinatario
            email["Subject"] = asunto
            email.set_content(cuerpo)
            if archivo_adjunto and os.path.exists(archivo_adjunto):
                with open(archivo_adjunto, "rb") as f:
                    file_data = f.read()
                    file_name = os.path.basename(archivo_adjunto)
                    email.add_attachment(
                        file_data,
                        maintype="application",
                        subtype="octet-stream",
                        filename=file_name,
                    )
            print("Conectando al servidor SMTP...")
            with smtplib.SMTP_SSL(self.servidor_smtp, self.puerto) as smtp:
                smtp.login(self.remitente, self.contraseña)
                smtp.send_message(email)
            print("Correo enviado exitosamente.")
        except Exception as e:
            print(f"Error al enviar el correo: {e}")

if __name__ == "__main__":
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    ANEXO_DIR = os.path.join(BASE_DIR, "anexo_ref_mas_vendidas")
    EVIDENCIAS_DIR = os.path.join(BASE_DIR, "evidencias")
    os.makedirs(ANEXO_DIR, exist_ok=True)
    os.makedirs(EVIDENCIAS_DIR, exist_ok=True)

    DRIVER_PATH = ChromeDriverManager().install()

    downloader = ArchivoDescarga(DRIVER_PATH, ANEXO_DIR, EVIDENCIAS_DIR)

    try:
        URL = "https://www.dane.gov.co/index.php/estadisticas-por-tema/precios-y-costos/precios-de-venta-al-publico-de-articulos-de-primera-necesidad-pvpapn"
        TEXTO_XPATH = "//*[contains(text(), 'Precios de los productos de primera necesidad')]"
        BOTON_XPATH = "/html/body/div[1]/div[5]/div/div[1]/div/div[2]/table[2]/tbody/tr/td/div/table[2]/tbody/tr/td/div/a"
        PANTALLAZO_TEXTO_NOMBRE = "pantallazo_texto.png"
        downloader.limpiar_carpeta_descargas()
        downloader.abrir_pagina(URL)
        downloader.buscar_texto_y_tomar_pantallazo(TEXTO_XPATH, PANTALLAZO_TEXTO_NOMBRE)
        downloader.buscar_y_descargar(BOTON_XPATH)

        print("\nProcesando la información del archivo descargado...\n")
        procesador = ProcesoData(BASE_DIR)
        archivo_reciente = procesador.obtener_ultimo_archivo()
        procesador.procesar_archivo(archivo_reciente)
        top_10_path = os.path.join(BASE_DIR, "resultados_procesados", "resultados_Cantidades 1203-1603.csv")
        datos_top_10 = pd.read_csv(top_10_path)
        resumen = """
        Hola,

        Este es el resumen de los productos vendidos:
        - Total de todos los productos vendidos: 1530890.00
        - Total de los 10 productos más vendidos: 1292302.00
        - Porcentaje de los 10 productos más vendidos respecto al total: 84.42%

        Saludos,
        Equipo de Automatización
        """
# ------------------------ CONFIGURAR CORREO -----------------------------------------

        remitente = "montero.sebastian1235@gmail.com"
        contraseña = "fdyv wqro cttm sfpz"
        destinatario = "ingjaviermonterot@gmail.com"
        asunto = "Resumen de Ventas - Top 10 Productos"
        correo = Correo(remitente, contraseña)
        correo.enviar(destinatario, asunto, resumen, top_10_path)
    except Exception as e:
        print("Error general durante la ejecución del flujo:", e)
    finally:
        downloader.cerrar()
