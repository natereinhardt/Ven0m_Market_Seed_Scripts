Attribute VB_Name = "PullMarketDataFits"
Sub PullMarketData_FITS()
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim itemCounts As Object
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim typeID As String
    Dim stationId As String
    Dim sellFormula As String
    Dim priceFormula As String
    Dim item As Variant
    Dim qty As Long
    Dim tbl As ListObject

    ' Create dictionary to store item counts
    Set itemCounts = CreateObject("Scripting.Dictionary")
    
    ' Set the source and destination sheets
    Set srcWs = ThisWorkbook.Sheets("All_Fits_RAW")
    
    ' Create the destination sheet if it doesn't exist
    On Error Resume Next
    Set destWs = ThisWorkbook.Sheets("All_Fits_Data")
    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destWs.Name = "All_Fits_Data"
    End If
    On Error GoTo 0
    
    ' Clear the destination sheet
    destWs.Cells.Clear
    
    ' Set station ID
    stationId = "1035466617946"
    
    ' Find the last row with data in column A of the source sheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' Add headers for new data in the destination sheet
    destWs.Cells(1, 1).Value = "Item"
    destWs.Cells(1, 2).Value = "Type ID"
    destWs.Cells(1, 3).Value = "Total Qty"
    destWs.Cells(1, 4).Value = "Sell Orders Qty"
    destWs.Cells(1, 5).Value = "Required Qty (5)"
    destWs.Cells(1, 6).Value = "Difference (Sell - Required)"
    destWs.Cells(1, 7).Value = "Cheapest Price"
    
    ' Loop through each row in the source sheet to count quantities
    For i = 2 To lastRow
        item = srcWs.Cells(i, 1).Value
        typeID = srcWs.Cells(i, 2).Value
        qty = srcWs.Cells(i, 3).Value
        
        ' Check if the item already exists in the dictionary
        If itemCounts.exists(item) Then
            itemCounts(item)(0) = itemCounts(item)(0) + qty ' Increment quantity
        Else
            ' Add new item to the dictionary
            itemCounts.Add item, Array(qty, typeID)
        End If
    Next i
    
    ' Populate the destination sheet with the data from the dictionary
    j = 2 ' Start populating from the second row
    For Each item In itemCounts.Keys
        destWs.Cells(j, 1).Value = item
        destWs.Cells(j, 2).Value = itemCounts(item)(1) ' Type ID
        destWs.Cells(j, 3).Value = itemCounts(item)(0) ' Total Quantity
        
        ' Create the formula for sell orders quantity with IFERROR
        sellFormula = "=IFERROR(SUM(EVEONLINE.MARKET_STRUCTURE_ORDERS(" & stationId & ", " & itemCounts(item)(1) & ", FALSE).volume_remain), 0)"
        
        ' Create the formula for the cheapest price
        priceFormula = "=IFERROR(MIN(EVEONLINE.MARKET_STRUCTURE_ORDERS(" & stationId & ", " & itemCounts(item)(1) & ", FALSE).price), 0)"
        
        ' Set the formulas in the destination sheet using Formula2
        destWs.Cells(j, 4).Formula2 = sellFormula
        destWs.Cells(j, 7).Formula2 = priceFormula
        destWs.Cells(j, 7).NumberFormat = "$#,##0.00" ' Format as currency
        
        ' Calculate the required quantity (Total Qty * 5)
        destWs.Cells(j, 5).Formula2 = "=" & destWs.Cells(j, 3).Address & " * 5"
        
        ' Calculate the difference as Sell Orders Qty - Required Qty
        destWs.Cells(j, 6).Formula2 = "=IFERROR(" & destWs.Cells(j, 4).Address & " - " & destWs.Cells(j, 5).Address & ", 0)"
        
        j = j + 1
    Next item
    
    ' Convert range to table to enable sorting
    Set tbl = destWs.ListObjects.Add(xlSrcRange, destWs.Range("A1:G" & destWs.Cells(destWs.Rows.Count, "A").End(xlUp).Row), , xlYes)
    tbl.Name = "MarketComparisonTable"
    
    ' Activate the destination worksheet before freezing panes
    destWs.Activate
    
    ' Freeze the header row
    destWs.Rows("2:2").Select
    ActiveWindow.FreezePanes = True

    ' Ensure the worksheet is activated
    destWs.Activate
    
    MsgBox "Market data formulas set successfully into All_Fits_Data tab!", vbInformation
End Sub

  �MV  ����     #yV  p1MV  ����    ��Z|V  `B�|V  ����    ��Z|V  0�FMV  ����    �Z|V  `A�|V  ����    @�Z|V  @T�LV  �       `9yV  0�|V  ����    P�Z|V  @�MV  ����    �2yV  ��|V  �       �Z|V  @	�MV  �   3 \  5yV  ��|V  ����3 , P�Z|V   �vV  ����t h ��Z|V   �vV  ����" N �>yV  ���|V  ����} } ��Z|V   �vV  �   " " �8yV  ���~V  ����" " ��Z|V   �vV  �   l l 