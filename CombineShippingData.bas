Attribute VB_Name = "CombineShippingData"
Sub CombineShippingData()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim typeIdWs As Worksheet
    Dim itemDict As Object
    Dim itemCounts As Object
    Dim typeIdDict As Object
    Dim sheetTracker As Object
    Dim item As String
    Dim qty As Variant
    Dim typeID As String
    Dim lastRow As Long
    Dim i As Long
    Dim key As Variant
    Dim startTime As Double
    Dim endTime As Double

    ' Start timing the script
    startTime = Timer
    
    ' Create dictionaries to store item quantities, counts, and type IDs
    Set itemDict = CreateObject("Scripting.Dictionary")
    Set itemCounts = CreateObject("Scripting.Dictionary")
    Set typeIdDict = CreateObject("Scripting.Dictionary")
    Set sheetTracker = CreateObject("Scripting.Dictionary")
    
    ' Set the destination sheet
    Set destWs = ThisWorkbook.Sheets("All_Shipping_Items_RAW")
    destWs.Cells.Clear
    
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
        ' Only include sheets that start with "Shipment"
        If Left(ws.Name, 8) = "Shipment" Then
            Debug.Print "Processing sheet: " & ws.Name
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Track sheets where items have been seen
            Set sheetTracker = CreateObject("Scripting.Dictionary")
            
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
                    
                    ' Track unique occurrences per sheet
                    If Not sheetTracker.exists(item) Then
                        sheetTracker.Add item, True
                        If itemCounts.exists(item) Then
                            itemCounts(item) = itemCounts(item) + 1
                        Else
                            itemCounts.Add item, 1
                        End If
                    End If
                End If
            Next i
        End If
    Next ws
    
    ' Add headers to the destination sheet
    destWs.Cells(1, 1).Value = "Item"
    destWs.Cells(1, 2).Value = "Type ID"
    destWs.Cells(1, 3).Value = "Total Quantity"
    destWs.Cells(1, 4).Value = "Occurrences"
    
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
        destWs.Cells(i, 4).Value = itemCounts(key)
        i = i + 1
    Next key
    
    ' Autofit columns
    destWs.Columns("A:D").AutoFit
    
    ' End timing the script
    endTime = Timer
    
    ' Print the time taken to the Immediate Window
    Debug.Print "Time taken: " & (endTime - startTime) & " seconds"
    
    MsgBox "Shipping data combined successfully into All_Shipping_Items_RAW tab!", vbInformation
End Sub

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                