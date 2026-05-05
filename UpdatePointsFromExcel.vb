Sub UpdatePointsFromExcel()

    ' === Excel Initialization ===
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim filePath As String

    filePath = "C:\Users\ELC\Documents\Projects\Avion Scale\ExcelFiles\GenerativeShapeDesign_SplineFromExcel.xls"

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
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
    Set geoSet = hybridBodies.Item("Construction-Airfoil") '

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

    MsgBox "Point coordinates updated successfully!"

End Sub
