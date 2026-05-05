Sub reflectAirfoilSetPointThroughOrigin()

'Set Active Document
Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

'Set part
Dim part1 As part
Set part1 = partDocument1.part

'Set HybridShapeFactory (Generador de entidades geométricas)
Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory


'Set de Hybridbodies (collección de entidades geométricas existentes de la parte)
Dim hybridBodies1 As hybridBodies
Set hybridBodies1 = part1.hybridBodies



'Set de Geometric set de "Construction-Airfoil"
Dim airfoilGeomSet As HybridBody
Set airfoilGeomSet = hybridBodies1.Item("Construction-Airfoil")


'Set de target point(s)
'Dim targetPoint As HybridShapePointCoord
'Set targetPoint = airfoilGeomSet.hybridShapes.Item("TestPoint")


' -------Definir target Point abajo y crear un for que recorra Construction-Airfoil------------------


'Set de geometric set existente dentro de la parte
Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Construction-Wing")

'Set de entidades geométricas dentro del geometric set definido en el paso anterior
Dim hybridShapes1 As hybridShapes
Set hybridShapes1 = hybridBody1.hybridShapes

'Set de la entidad geométrica (nuevo point origin) dentro de la familia de entidades del geometric set definido
Dim targetOriginPoint As HybridShapePointCoord
Set targetOriginPoint = hybridShapes1.Item("PointWingOrigin")

'Set reference para cambiar el punto de referencia en el target point
Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(targetOriginPoint)

'Set de nuevo punto de referencia para el target point

Dim targetPoint As HybridShapePointCoord
For i = 1 To airfoilGeomSet.hybridShapes.Count
    Set targetPoint = airfoilGeomSet.hybridShapes.Item(i)
    targetPoint.x.Value = -targetPoint.x.Value
Next

part1.Update

End Sub


    MsgBox "Point coordinates updated successfully"

End Sub
