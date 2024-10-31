import win32com.client as wc

inv = wc.GetActiveObject('Inventor.Application')
# print(inv)
print("--------------------------------------------------------")
print("Todas las dimensiones se deben de indicar en centimetros")
print("--------------------------------------------------------")

Largo = float(input("Largo: "))
Radio = float(input("Radio: "))
Nombre = str(input("Nombre de la pieza: "))

# crear una parte

inv_part_document = inv.Documents.Add(12290, inv.FileManager.GetTemplateFile(12290, 8962))
pin_comp_def = inv_part_document.ComponentDefinition

pin_sketch = pin_comp_def.Sketches.Add(pin_comp_def.WorkPlanes.Item(3))

tg = inv.TransientGeometry

circle_sketch =pin_sketch.SketchCircles

circle_sketch.AddByCenterRadius(tg.CreatePoint2d(0, 0), Radio)
solid_profile = pin_sketch.Profiles.AddforSolid()

extrude_solid_def = pin_comp_def.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile, 20481)
extrude_solid_def.SetDistanceExtent(Largo, 20995)

pin_comp_def.Features.ExtrudeFeatures.Add(extrude_solid_def)

inv.ActiveView.GoHome()

inv.ActiveDocument.SaveAs(f"C:/Users/anhernandez/Desktop/hello/{Nombre}.ipt", False)