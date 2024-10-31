import win32com.client as wc


print("todas las dimensiones se deben de indicar en centimetros")

Largo = int(input("Largo: "))
Ancho = int(input("Ancho: ")) 
Espesor = int(input("Espesor: "))
perforacion = int(input("Perforacion: "))
Redondo = float(input("Redondo: "))
Nombre = str(input("Nombre de la pieza: "))

Largo_Div = Largo/2
Ancho_Div = Ancho/2


inv = wc.GetActiveObject('Inventor.Application')
# print(inv)

# crear una parte

inv_part_document = inv.Documents.Add(12290, inv.FileManager.GetTemplateFile(12290, 8962))
# inv_part_document = inv.ActiveDocument
# print(inv_part_document)

# definir dimensiones del componente

part_comp_definition = inv_part_document.ComponentDefinition
# print(dir(part_comp_definition))

# Crear y definir plano del sketch

sketch = part_comp_definition.Sketches.Add(part_comp_definition.WorkPlanes.Item(3))
# print(part_comp_definition.WorkPlanes.Item(3).Name)  //ver el nombre de los planos

# crear sketch

tg = inv.TransientGeometry

# definir dimensiones, los digitos los toma de los centros en 8 y el 3 quedando un rectangulo de 160 x 60 mm las dimensiones las toma en centimetros

sketch.SketchLines.AddAsTwoPointCenteredRectangle(tg.CreatePoint2d(0, 0), tg.CreatePoint2d(Largo_Div, Ancho_Div))

# poner la vista de home

inv.ActiveView.GoHome()

# crear una extrucion del sketch

solid_profile = sketch.Profiles.AddforSolid()

# Crear las caracteristiscas de la extrucion

ext_solid_def = part_comp_definition.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile, 20481)
ext_solid_def.SetDistanceExtent(Espesor, 20995)

part_comp_definition.Features.extrudeFeatures.Add(ext_solid_def)


# crear perforaciones

# definir plano para el sketch
hole_sketch = part_comp_definition.Sketches.Add(part_comp_definition.WorkPlanes.Item(3))

# Cear una colleccion de objetos
hole_centers = inv.TransientObjects.CreateObjectCollection()

# definir los centros de los huecos
hole_centers.Add(hole_sketch.SketchPoints.Add(tg.CreatePoint2d(-7, -2)))
hole_centers.Add(hole_sketch.SketchPoints.Add(tg.CreatePoint2d(-7, 2)))
hole_centers.Add(hole_sketch.SketchPoints.Add(tg.CreatePoint2d(7, -2)))
hole_centers.Add(hole_sketch.SketchPoints.Add(tg.CreatePoint2d(7, 2)))

# definir  las caracteristicas de los huecos en la parte

part_comp_definition.Features.HoleFeatures.AddDrilledByThroughAllExtent(hole_centers, perforacion, 20995)

# Redondos en la esquinas

edges1 = part_comp_definition.SurfaceBodies.Item(1).Faces.Item(6).Edges.Item(3)
edges2 = part_comp_definition.SurfaceBodies.Item(1).Faces.Item(6).Edges.Item(1)
edges3 = part_comp_definition.SurfaceBodies.Item(1).Faces.Item(8).Edges.Item(3)
edges4 = part_comp_definition.SurfaceBodies.Item(1).Faces.Item(8).Edges.Item(1)

edge_colletion_side = inv.TransientObjects.CreateEdgeCollection()
side_fillet_definition = part_comp_definition.Features.FilletFeatures.CreateFilletDefinition()

edge_colletion_side.Add(edges1)
edge_colletion_side.Add(edges2)
edge_colletion_side.Add(edges3)
edge_colletion_side.Add(edges4)

side_fillet_definition.AddConstantRadiusEdgeSet(edge_colletion_side, Redondo)
part_comp_definition.Features.FilletFeatures.Add(side_fillet_definition)


inv.ActiveDocument.SaveAs(f"C:/Users/anhernandez/Desktop/hello/{Nombre}.ipt", False)