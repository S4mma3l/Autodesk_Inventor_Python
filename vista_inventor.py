import win32com.client as wc

# Conectar a Autodesk Inventor
inv = wc.GetActiveObject("Inventor.Application")

# Acceder al documento activo
active_doc = inv.ActiveDocument

# Imprimir el tipo de documento para depuración
print(f"Tipo de documento activo: {active_doc.DocumentType}")

# Verificar si es un documento de dibujo
if active_doc.DocumentType == 12291:  # Valor para kDrawingDocumentObject
    # Intentar acceder a DrawingViews
    try:
        # Forzar actualización (opcional)
        inv.ActiveDocument.Update()

        # Verificar si se puede acceder a las DrawingViews
        if hasattr(active_doc, "DrawingViews"):
            drawing_views = active_doc.DrawingViews
            
            # Verificar si hay vistas
            if drawing_views.Count > 0:
                # Recorrer las vistas y mostrar sus nombres y orientaciones
                for view in drawing_views:
                    print(f"Nombre de la vista: {view.Name}")
                    print(f"Orientación de la vista: {view.ViewOrientation}")
            else:
                print("No hay vistas en este documento de dibujo.")
        else:
            print("No se puede acceder a DrawingViews.")
    except Exception as e:
        print(f"Ocurrió un error al acceder a las vistas: {e}")
else:
    print("El documento activo no es un dibujo.")