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

Sub UpdatePointsFromExcel()

    ' === Excel Initialization ===
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim filePath As String

    filePath = "C:\Users\ELC\Documents\Projects\Avion Scale\ExcelFiles\GenerativeShapeDesign_SplineFromExcel.xls" ' <<< CHANGE THIS TO YOUR FILE

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Set to True if you want to see Excel open

    On Error Resume Next
    Set xlWorkbook = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0

    If xlWorkbook Is Nothing Then
        MsgBox "Could not open Excel file: " & filePath
        Exit Sub
    End If

    Set xlSheet = xlWorkbook.Sheets(1)

    ' === CATIA Initialization ===
    Dim CATIA As Object
    Set CATIA = GetObject(, "CATIA.Application")

    If CATIA Is Nothing Then
        MsgBox "CATIA is not running."
        Exit Sub
    End If

    ' === Get the active document ===
    Dim partDoc As PartDocument
    Set partDoc = CATIA.ActiveDocument

    Dim part As part
    Set part = partDoc.part

    Dim hybridBodies As hybridBodies
    Set hybridBodies = part.hybridBodies

    ' === Get the target Geometric Set ===
    Dim geoSet As HybridBody
    Set geoSet = hybridBodies.Item("Construction-Airfoil") ' Change this if needed

    ' === Loop through Excel rows ===
    row = 3 ' Assuming headers in row 1
    Do While xlSheet.Cells(row, 1).Value <> "EndCurve"

        Dim pointName As String
        Dim xCoord As Double
        Dim yCoord As Double
        Dim zCoord As Double

        pointName = xlSheet.Cells(row, 5).Value
        xCoord = CDbl(xlSheet.Cells(row, 1).Value)
        yCoord = CDbl(xlSheet.Cells(row, 2).Value)
        zCoord = CDbl(xlSheet.Cells(row, 3).Value)

        ' === Search for the point in the Geometric Set ===
        Dim i As Integer
        Dim pointFound As Boolean
        pointFound = False

        For i = 1 To geoSet.hybridShapes.Count
            If geoSet.hybridShapes.Item(i).Name = pointName Then
                Dim point As HybridShapePointCoord
                Set point = geoSet.hybridShapes.Item(i)

                point.x.Value = xCoord
                point.y.Value = yCoord
                ' Optionally update Z as well: point.Z.Value = zCoord

                pointFound = True
                Exit For
            End If
        Next

        If Not pointFound Then
            MsgBox "Point not found: " & pointName
        End If

        row = row + 1
    Loop

    ' === Update the part ===
    part.Update

    MsgBox "Point coordinates updated successfully"

End Sub
