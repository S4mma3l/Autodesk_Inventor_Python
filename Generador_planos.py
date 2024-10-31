import win32com.client as wc

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

    # Obtener ensamblajes (.iam) abiertos en Inventor
    open_assemblies = [doc for doc in inv.Documents if doc.DocumentType == 12291]  # Tipo de documento 12291 es ensamblaje

    # Verificar si hay ensamblajes abiertos
    if not open_assemblies:
        print("No se encontraron ensamblajes (.IAM) abiertos en Autodesk Inventor.")
    else:
        # Crear hoja y vistas para cada ensamblaje abierto
        for i, assembly_doc in enumerate(open_assemblies, start=1):
            # Agregar una nueva hoja en el archivo de dibujo activo
            new_sheet = drawing_doc.Sheets.Add()  # Agrega una nueva hoja
            new_sheet.Name = f"Hoja_{i}"  # Renombrar la hoja

            # Crear un punto 2D para la posición de la vista
            tg = inv.TransientGeometry
            view_point = tg.CreatePoint2d(10, 10)  # Punto para la vista base
            
            # Crear un NameValueMap para las opciones de la vista base
            base_view_options = inv.TransientObjects.CreateNameValueMap()
            base_view_options.Add("PositionalRepresentation", "MyPositionalRep")  # Cambia "MyPositionalRep" al nombre real
            base_view_options.Add("DesignViewRepresentation", "MyDesignViewRep")  # Cambia "MyDesignViewRep" al nombre real
            base_view_options.Add("DesignViewAssociative", True)

            # Crear vista base
            try:
                # Comprobar que el ensamblaje esté bien cargado
                if assembly_doc is not None:
                    # Crear la vista base
                    view_base = new_sheet.DrawingViews.AddBaseView(
                        Model=assembly_doc,
                        Position=view_point,
                        Scale=1,
                        ViewOrientation=wc.constants.kFrontViewOrientation,  # Asegúrate de que esto esté bien definido
                        HiddenLineStyle=wc.constants.kHiddenLineRemoved,
                        BaseViewOptions=base_view_options  # Añadir opciones de la vista base
                    )
                    print(f"Vista base creada en Hoja_{i} para el ensamblaje '{assembly_doc.DisplayName}'.")

                    # Crear vista de sección
                    section_view_position = tg.CreatePoint2d(30, 10)  # Posición para la vista de sección
                    section_view = new_sheet.DrawingViews.AddSectionView(
                        BaseView=view_base,
                        Position=section_view_position,
                        HiddenLineStyle=wc.constants.kHiddenLineRemoved
                    )
                    print(f"Vista de sección creada en Hoja_{i}.")

                    # Agregar información del cajetín
                    title_block = new_sheet.TitleBlock  # Obtener el cajetín
                    title_block.SetTextValue("Nombre del Proyecto", assembly_doc.DisplayName)  # Establecer el nombre del proyecto
                    title_block.SetTextValue("Creado Por", "Tu Nombre")
                    title_block.SetTextValue("Fecha", inv.TransientGeometry.CreateDateTime())  # Obtener la fecha actual
                else:
                    print(f"El ensamblaje '{assembly_doc.DisplayName}' no está disponible para crear vistas.")

            except Exception as e:
                print(f"Ocurrió un error al crear vistas para el ensamblaje '{assembly_doc.DisplayName}': {e}")

    print("Proceso completado: todas las vistas se han generado en el archivo .idw actual.")