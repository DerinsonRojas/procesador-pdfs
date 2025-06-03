import os
import re
import sys
import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import messagebox
from tqdm import tqdm
import winsound  # Solo para Windows

# Configuración de rutas importantes
base_path = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
carpeta_pdfs = os.path.join(base_path, "datos", "pdfs")
ruta_excel = os.path.join(base_path, "datos", "codigos.xlsx")

# Validación de existencia de archivos y directorios necesarios
if not os.path.exists(carpeta_pdfs):
    print(f"No existe la carpeta de PDFs: {carpeta_pdfs}")
    sys.exit()

if not os.path.isfile(ruta_excel):
    print("No se encontró el archivo codigos.xlsx en la carpeta 'datos'.")
    sys.exit()

# Cargar el archivo Excel y extraer los códigos de expediente
df_codigos = pd.read_excel(ruta_excel)
if "Expediente" not in df_codigos.columns:
    print("El archivo Excel debe tener una columna llamada 'Expediente'.")
    sys.exit()

# Lista de códigos de expediente a buscar en los PDFs
codigos = df_codigos["Expediente"].dropna().astype(str).tolist()
resultados = []

# Obtener la lista de archivos PDF en la carpeta
pdfs = [f for f in os.listdir(carpeta_pdfs) if f.lower().endswith(".pdf")]

# Iterar sobre cada PDF y buscar los códigos en su contenido
for archivo in tqdm(pdfs, desc="Procesando PDFs", unit="pdf"):
    ruta_pdf = os.path.join(carpeta_pdfs, archivo)

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            paginas = pdf.pages[1:]  # Omitir la primera página

            for i, pagina in enumerate(paginas, start=2):
                texto = pagina.extract_text() or ""  # Extraer texto de la página
                lineas = texto.splitlines()  # Dividir en líneas para búsqueda más precisa

                # Buscar cada código en el texto extraído
                for codigo in codigos:
                    pattern = re.compile(re.escape(codigo))
                    for linea in lineas:
                        if pattern.search(linea):
                            resultados.append([
                                codigo,
                                archivo,
                                "encontrado",
                                i,
                                "exacto",
                                linea.strip()
                            ])
                            break  # Se detiene la búsqueda en esta página si encuentra el código

    except Exception as e:
        print(f"Error procesando {archivo}: {e}")
        for codigo in codigos:
            resultados.append([codigo, archivo, "error", "", "", ""])
        continue

# Identificar códigos que no fueron encontrados en ningún PDF
codigos_encontrados = {resultado[0] for resultado in resultados}
for codigo in codigos:
    if codigo not in codigos_encontrados:
        resultados.append([
            codigo,
            "",
            "no encontrado",
            "",
            "",
            ""
        ])

# Guardar los resultados en un archivo Excel
df_resultados = pd.DataFrame(resultados, columns=[
    "Expediente", "PDF", "Estado", "Página encontrada", "Match exacto", "Información"
])
df_resultados.to_excel(os.path.join(base_path, "Resultados_de_busqueda.xlsx"), index=False)
# Ruta al archivo de resultados
ruta_resultados = os.path.join(os.path.dirname(__file__), "Resultados_de_busqueda.xlsx")

# Ruta al archivo de resultados
ruta_resultados = os.path.join(os.path.dirname(__file__), "Resultados_de_busqueda.xlsx")

# Función para abrir la carpeta y cerrar la ventana de alerta
def abrir_carpeta():
    carpeta = os.path.dirname(ruta_resultados)
    ventana_alerta.destroy()  # Cierra la ventana emergente
    os.startfile(carpeta)  # Abre la carpeta en Windows
    sys.exit()  # Finaliza completamente el programa
    # En macOS o Linux, usa: subprocess.run(["xdg-open", carpeta]) o subprocess.run(["open", carpeta])

# Función para reproducir sonido
def reproducir_sonido():
    try:
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)  # Sonido predeterminado en Windows
    except:
        print("No se pudo reproducir el sonido.")

# Función para mostrar la alerta con botón adicional
def mostrar_alerta():
    global ventana_alerta  # Definir como global para poder cerrarla desde abrir_carpeta()
    ventana = tk.Tk()
    ventana.withdraw()  # Ocultar la ventana principal

    # Crear una nueva ventana emergente
    ventana_alerta = tk.Toplevel()
    ventana_alerta.title("Proceso terminado")

    # Etiqueta con mensaje
    mensaje = tk.Label(ventana_alerta, text="Resultados en 'Resultados_de_busqueda.xlsx'")
    mensaje.pack(pady=10)

    # Botón para abrir la carpeta y cerrar la ventana
    boton_abrir = tk.Button(ventana_alerta, text="Abrir carpeta", command=abrir_carpeta)
    boton_abrir.pack(pady=5)

    # Botón para cerrar la ventana manualmente
    # Modifica el botón de cerrar para que también termine el programa completamente
    boton_cerrar = tk.Button(ventana_alerta, text="Cerrar", command=lambda: (ventana_alerta.destroy(), sys.exit()))
    boton_cerrar.pack(pady=5)

    # Reproducir sonido cuando la ventana aparece
    reproducir_sonido()
    ventana_alerta.mainloop()
    
mostrar_alerta()

print("\n Búsqueda completada. Resultados en 'Resultados_de_busqueda.xlsx'")
