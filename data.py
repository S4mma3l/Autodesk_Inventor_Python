import win32com.client
from openpyxl import Workbook

def export_bom_to_excel():
    # Inicializar la aplicación de Inventor
    inventor_app = win32com.client.Dispatch("Inventor.Application")
    inventor_app.Visible = True

    # Obtener el documento activo
    oDoc = inventor_app.ActiveDocument

    # Obtener la referencia al BOM
    oBOM = oDoc.ComponentDefinition.BOM

    # Establecer la vista estructurada a 'todos los niveles'
    oBOM.StructuredViewFirstLevelOnly = False
    oBOM.StructuredViewEnabled = True

    # Crear un nuevo libro de trabajo de Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM Data"

    # Definir los encabezados para el archivo de Excel
    headers = ["ITEM", "QTY", "PART NUMBER", "DESCRIPTION", "Largo", "Ancho", "Espesor"]
    ws.append(headers)

    # Obtener la vista BOM estructurada
    oBOMView = oBOM.BOMViews.Item("Parts Only (legacy)")

    # Iterar a través de cada fila en la vista BOM
    for i in range(1, oBOMView.BOMRows.Count + 1):
        try:
            oRow = oBOMView.BOMRows.Item(i)
            
            # Obtener el número de artículo y cantidad
            item_number = oRow.ItemNumber
            quantity = oRow.ItemQuantity
            
            # Obtener el componente correspondiente
            oCompDef = oRow.ComponentDefinition
            
            # Obtener el conjunto de propiedades de diseño
            oPropSet = oCompDef.Document.PropertySets.Item("Design Tracking Properties")
            
            # Obtener el número de parte y la descripción
            part_number = oPropSet.Item("Part Number").Value
            description = oPropSet.Item("Description").Value

            # Inicializar Largo, Ancho y Espesor
            largo = ancho = espesor = ""

            # Verificar si el componente es una pieza
            if oRow.Component.DefinitionType == 1:  # 1 representa kPartObject
                oPart = oRow.Component
                try:
                    largo = oPart.PropertySets.Item("Design Tracking Properties").Item("Length").Value
                    ancho = oPart.PropertySets.Item("Design Tracking Properties").Item("Width").Value
                    espesor = oPart.PropertySets.Item("Design Tracking Properties").Item("Thickness").Value
                except Exception as e:
                    print(f"Error al obtener propiedades de la pieza: {e}")

            # Agregar los datos a la hoja de cálculo
            ws.append([item_number, quantity, part_number, description, largo, ancho, espesor])

        except Exception as e:
            print(f"Error al procesar la fila: {e}")

    # Guardar el archivo de Excel
    wb.save("Z:\\Autocad Files\\PROYECTOS 2024\\RESIDENCE LOTE #7\\SET DE PUERTAS\\Lista\\BOM-ExtractedData.xlsx")
    print("Exportación completada. Archivo guardado en Z:\\Autocad Files\\PROYECTOS 2024\\RESIDENCE LOTE #7\\SET DE PUERTAS\\Lista\\BOM-ExtractedData.xlsx")

# Ejecutar la función
export_bom_to_excel()