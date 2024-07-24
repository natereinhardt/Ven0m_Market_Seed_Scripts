Sub CombineFitData()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim typeIdWs As Worksheet
    Dim itemDict As Object
    Dim typeIdDict As Object
    Dim item As String
    Dim qty As Variant
    Dim typeID As String
    Dim lastRow As Long
    Dim i As Long
    Dim key As Variant
    Dim startTime As Double
    Dim endTime As Double
    Dim sheetExists As Boolean
    
    ' Start timing the script
    startTime = Timer
    
    ' Create dictionaries to store item quantities and type IDs
    Set itemDict = CreateObject("Scripting.Dictionary")
    Set typeIdDict = CreateObject("Scripting.Dictionary")
    
    ' Check if the destination sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "All_Fits_RAW" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If the sheet exists, clear it. Otherwise, create a new sheet.
    If sheetExists Then
        Set destWs = ThisWorkbook.Sheets("All_Fits_RAW")
        destWs.Cells.Clear
    Else
        Set destWs = ThisWorkbook.Sheets.Add
        destWs.Name = "All_Fits_RAW"
    End If
    
    ' Set the Type_ids sheet
    Set typeIdWs = ThisWorkbook.Sheets("Type_ids")
    
    ' Read the Type_ids into the dictionary
    lastRow = typeIdWs.Cells(typeIdWs.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        typeIdDict(typeIdWs.Cells(i, 2).Value) = typeIdWs.Cells(i, 3).Value
    Next i
    Debug.Print "Type IDs loaded into dictionary"
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Only include sheets that start with "Fit-"
        If Left(ws.Name, 4) = "Fit-" Then
            Debug.Print "Processing sheet: " & ws.Name
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Loop through each row in the current sheet
            For i = 2 To lastRow
                item = ws.Cells(i, 1).Value
                qty = ws.Cells(i, 2).Value
                
                ' Check for valid item and quantity
                If Len(Trim(item)) > 0 Then
                    If Not IsNumeric(qty) Or IsEmpty(qty) Then qty = 1
                    
                    ' Update the dictionary with item quantities
                    If itemDict.exists(item) Then
                        itemDict(item) = itemDict(item) + qty
                    Else
                        itemDict.Add item, qty
                    End If
                End If
            Next i
        End If
    Next ws
    
    ' Add headers to the destination sheet
    destWs.Cells(1, 1).Value = "Item"
    destWs.Cells(1, 2).Value = "Type ID"
    destWs.Cells(1, 3).Value = "Total Quantity"
    
    ' Output the combined data to the destination sheet
    i = 2
    For Each key In itemDict.Keys
        destWs.Cells(i, 1).Value = key
        If typeIdDict.exists(key) Then
            destWs.Cells(i, 2).Value = typeIdDict(key)
        Else
            destWs.Cells(i, 2).Value = "N/A"
        End If
        destWs.Cells(i, 3).Value = itemDict(key)
        i = i + 1
    Next key
    
    ' Autofit columns
    destWs.Columns("A:C").AutoFit
    
    ' End timing the script
    endTime = Timer
    
    ' Print the time taken to the Immediate Window
    Debug.Print "Time taken: " & (endTime - startTime) & " seconds"
    
    MsgBox "Shipping data combined successfully into All_Fits_RAW tab!", vbInformation
End Sub

