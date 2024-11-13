import clr
import os
import pythoncom
from win32com.client import Dispatch
import tkinter as tk
from tkinter import filedialog

# Función para elegir un archivo usando un cuadro de diálogo
def seleccionar_archivo(tipo_archivo):
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de tkinter
    archivo = filedialog.askopenfilename(title=f"Seleccionar archivo {tipo_archivo}",
                                         filetypes=[(f"{tipo_archivo} files", f"*.{tipo_archivo}")])
    return archivo

# Inicializar Inventor COM
try:
    pythoncom.CoInitialize()  # Inicializar la API COM
    inventor_app = Dispatch("Inventor.Application")
    inventor_app.Visible = True  # Mostrar Inventor
except Exception as e:
    inventor_app = None
    print(f"No se pudo iniciar Autodesk Inventor: {e}")

# Elegir la ruta del archivo de plano (IDW) y el archivo de ensamblaje (IAM)
ruta_plano = seleccionar_archivo("idw")
archivo_modelo = seleccionar_archivo("iam")

if inventor_app and os.path.exists(ruta_plano) and os.path.exists(archivo_modelo):
    # Intentar abrir el archivo de dibujo (IDW)
    try:
        plano_doc = inventor_app.Documents.Open(ruta_plano)
    except Exception as e:
        print(f"Error al abrir el archivo de dibujo: {e}")
        plano_doc = None
    
    # Verificar si el documento abierto es un DrawingDocument (IDW)
    if plano_doc:
        print(f"Tipo de documento abierto: {plano_doc.DocumentType}")  # Mostrar el tipo de documento
        print(f"Nombre del documento: {plano_doc.DisplayName}")
        print(f"Ruta completa del documento: {plano_doc.FullFileName}")
        
        if plano_doc.DocumentType == 12292:  # Verificar si es un DrawingDocument (IDW)
            print("Se ha abierto un archivo de dibujo (IDW).")

            # Abrir el archivo de ensamblaje (IAM)
            try:
                ensamblaje_doc = inventor_app.Documents.Open(archivo_modelo)
                print(f"Se ha abierto el archivo de ensamblaje (IAM).")
            except Exception as e:
                print(f"Error al abrir el archivo de ensamblaje (IAM): {e}")
                ensamblaje_doc = None
            
            if ensamblaje_doc:
                # Verificar si el archivo de ensamblaje es del tipo correcto (IAM)
                if ensamblaje_doc.DocumentType == 12291:  # 12291 significa AssemblyDocument (IAM)
                    print("El archivo de ensamblaje (IAM) es válido.")
                    
                    # Intentar crear una nueva hoja si no existe una hoja activa
                    try:
                        if plano_doc.Sheets.Count == 0:
                            print("No hay hojas, creando una nueva hoja.")
                            hoja = plano_doc.Sheets.Add("Hoja 1")
                        else:
                            hoja = plano_doc.Sheets.Item(1)  # Obtener la primera hoja
                            print("Se ha obtenido la primera hoja del plano.")
                    except Exception as e:
                        print(f"Error al acceder a las hojas del documento: {e}")
                        hoja = None
                    
                    if hoja:
                        # Definir constantes para la orientación de la vista
                        kIsoTopRightViewOrientation = 133  # Orientación de vista isométrica superior derecha
                        kModelSourceType = 1  # Fuente del modelo (BaseView)
                        kFromBaseDrawingViewStyle = 104  # Estilo de vista basado en origen
                        
                        # Crear posición de la vista como punto 2D
                        try:
                            posicion = inventor_app.TransientGeometry.CreatePoint2d(10, 10)
                        except Exception as e:
                            print(f"Error al crear punto 2D: {e}")
                            posicion = None
                        
                        if posicion:
                            # Agregar vista basada en el ensamblaje (IAM)
                            try:
                                vista = hoja.DrawingViews.AddBaseView(
                                    posicion,              # Posición
                                    archivo_modelo,        # Ruta del archivo de ensamblaje (IAM)
                                    kIsoTopRightViewOrientation,  # Orientación
                                    1.0,                   # Escala
                                    kModelSourceType,      # Tipo de fuente (BaseView)
                                    kFromBaseDrawingViewStyle  # Estilo de vista
                                )
                                
                                # Ajustar otros parámetros si es necesario
                                vista.Scale = 0.5  # Ajustar la escala
                                vista.Label = "Plano Puerta"
                                
                                # Guardar el documento
                                plano_doc.Save()
                                print("El plano se ha actualizado y guardado con éxito.")
                            except Exception as e:
                                print(f"Error al agregar la vista base: {e}")
                        else:
                            print("No se pudo crear la posición 2D.")
                else:
                    print("El archivo de ensamblaje no es un documento de ensamblaje (IAM).")
            else:
                print("No se pudo abrir el archivo de ensamblaje (IAM).")
        else:
            print("El archivo abierto no es un documento de dibujo (IDW). DocumentType:", plano_doc.DocumentType)
    else:
        print("No se pudo abrir el archivo de dibujo (IDW).")
else:
    print("No se pudo encontrar el archivo de plano o el archivo de ensamblaje (IAM) no está disponible.")

# Finalizar sesión COM
pythoncom.CoUninitialize()