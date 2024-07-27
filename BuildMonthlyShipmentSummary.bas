Sub BuildMonthlyShipmentSummary()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim sheetName As String
    Dim shipmentDate As Date
    Dim monthKey As Variant
    Dim monthDict As Object
    Dim shipmentSheetsDict As Object
    Dim uniqueItemsDict As Object
    Dim i As Long
    Dim lastRow As Long
    Dim itemName As String
    Dim dateParts() As String
    
    ' Create dictionaries to store month keys and counts
    Set monthDict = CreateObject("Scripting.Dictionary")
    Set shipmentSheetsDict = CreateObject("Scripting.Dictionary")
    Set uniqueItemsDict = CreateObject("Scripting.Dictionary")
    
    ' Create the destination worksheet
    On Error Resume Next
    Set destWs = ThisWorkbook.Sheets("Monthly Shipment Summary")
    On Error GoTo 0
    
    If Not destWs Is Nothing Then
        Application.DisplayAlerts = False
        destWs.Delete
        Application.DisplayAlerts = True
    End If
    
    Set destWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    destWs.Name = "Monthly Shipment Summary"
    
    ' Add headers to the new summary sheet
    destWs.Cells(1, 1).Value = "Month"
    destWs.Cells(1, 2).Value = "Shipment Count"
    destWs.Cells(1, 3).Value = "Unique Item Count"
    destWs.Cells(1, 4).Value = "Shipment Sheets"
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        
        ' Check if the sheet name starts with "Shipment"
        If Left(sheetName, 9) = "Shipment-" Then
            ' Extract the date part from the sheet name
            dateParts = Split(sheetName, "-")
            shipmentDate = DateSerial(CInt(dateParts(3)), CInt(dateParts(1)), CInt(dateParts(2)))
            monthKey = Format(shipmentDate, "mmmm yyyy")
            
            ' Increment the count for the month key in the dictionary
            If monthDict.exists(monthKey) Then
                monthDict(monthKey) = monthDict(monthKey) + 1
            Else
                monthDict.Add monthKey, 1
            End If
            
            ' Store sheet names in the dictionary
            If shipmentSheetsDict.exists(monthKey) Then
                shipmentSheetsDict(monthKey) = shipmentSheetsDict(monthKey) & ", " & sheetName
            Else
                shipmentSheetsDict.Add monthKey, sheetName
            End If
            
            ' Track unique item names for the month
            If Not uniqueItemsDict.exists(monthKey) Then
                Set uniqueItemsDict(monthKey) = CreateObject("Scripting.Dictionary")
            End If
            
            ' Loop through items in the current sheet to count unique item names
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastRow ' Assuming item names start from the second row
                itemName = ws.Cells(i, 1).Value ' Assuming item names are in the first column
                If Len(Trim(itemName)) > 0 Then
                    uniqueItemsDict(monthKey)(itemName) = 1
                End If
            Next i
        End If
    Next ws
    
    ' Populate the destination sheet with the data from the dictionaries
    i = 2 ' Start populating from the second row
    For Each monthKey In monthDict.Keys
        destWs.Cells(i, 1).Value = monthKey
        destWs.Cells(i, 2).Value = monthDict(monthKey)
        destWs.Cells(i, 3).Value = uniqueItemsDict(monthKey).Count
        destWs.Cells(i, 4).Value = shipmentSheetsDict(monthKey)
        i = i + 1
    Next monthKey
    
    ' Ensure the date format is full month name and year
    destWs.Columns("A").NumberFormat = "mmmm yyyy"
    
    ' Autofit columns
    destWs.Columns("A:D").AutoFit
    
    ' Activate the summary worksheet
    destWs.Activate
    
    MsgBox "Monthly shipment summary built successfully in Monthly Shipment Summary tab!", vbInformation
End Sub

