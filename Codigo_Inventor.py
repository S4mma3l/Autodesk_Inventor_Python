import win32com.client as wc

print("Intentando acceder al archivo .idw actualmente abierto en Autodesk Inventor...")

# Conexión con Autodesk Inventor
inv = wc.GetActiveObject("Inventor.Application")
inv.Visible = True

# Obtener el archivo activo en Inventor
active_doc = inv.ActiveDocument

# Verificar si el archivo activo es un archivo de dibujo (.idw)
if active_doc.DocumentType != 5 and not active_doc.FullFileName.endswith(".idw"):
    print("El documento activo no es un archivo de dibujo (.idw). Por favor, abre un archivo .idw y vuelve a ejecutar el script.")
else:
    print("Archivo .idw detectado correctamente.")

    # Obtener el archivo de dibujo activo
    drawing_doc = active_doc

    # Listar y verificar todos los documentos abiertos
    print("Documentos abiertos:")
    open_assemblies = []
    for doc in inv.Documents:
        print(f"- {doc.DisplayName} (Tipo: {doc.DocumentType})")
        if doc.DocumentType == 2:  # Tipo 2 es ensamblaje
            open_assemblies.append(doc)

    print("Proceso completado: verificación de documentos finalizada.")