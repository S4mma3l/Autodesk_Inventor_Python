import win32com.client as wc
import os

print("Creando vistas en el archivo .idw actualmente abierto en Autodesk Inventor...")

# Conexión con Autodesk Inventor
inv = wc.GetActiveObject("Inventor.Application")
inv.Visible = True

# Obtener el archivo activo en Inventor
active_doc = inv.ActiveDocument

# Imprimir tipo de documento para depuración
print(f"Tipo de documento activo: {active_doc.DocumentType}")

# Verificar si el archivo activo es un dibujo (.idw)
if active_doc.DocumentType != 5 and not active_doc.FullFileName.endswith(".idw"):
    print("El documento activo no es un archivo de dibujo (.idw). Por favor, abre un archivo .idw y vuelve a ejecutar el script.")
else:
    print("Archivo .idw detectado correctamente.")

    # Obtener el archivo de dibujo activo
    drawing_doc = active_doc

    # Establecer el formato de hoja deseado
    sheet_format_name = "C size, 4 view"  # Cambia esto al nombre de tu formato de hoja
    try:
        sheet_format = drawing_doc.SheetFormats.Item(sheet_format_name)
    except Exception as e:
        print(f"No se pudo encontrar el formato de hoja '{sheet_format_name}': {e}")
        sheet_format = None

    # Abrir el modelo de documento (invisible)
    model_file_path = "C:\\temp\\block.ipt"  # Cambia esto a la ruta de tu modelo
    model_doc = None  # Inicializar la variable aquí

    if not os.path.isfile(model_file_path):
        print(f"El archivo '{model_file_path}' no se encuentra. Verifica la ruta.")
    else:
        try:
            model_doc = inv.Documents.Open(model_file_path, False)
            print(f"Modelo '{model_file_path}' abierto correctamente.")
        except Exception as e:
            print(f"Ocurrió un error al abrir el modelo: {e}")
            model_doc = None

    # Crear una nueva hoja basada en el formato de hoja
    if sheet_format and model_doc:
        new_sheet = drawing_doc.Sheets.AddUsingSheetFormat(sheet_format, model_doc)
        print(f"Nueva hoja creada utilizando el formato '{sheet_format_name}'.")
    else:
        print("No se pudo crear la nueva hoja debido a que el formato o el modelo no están disponibles.")

    print("Proceso completado: se ha creado una nueva hoja en el archivo .idw actual.")