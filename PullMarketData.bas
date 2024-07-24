Attribute VB_Name = "PullMarketData"
Sub PullMarketData()
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim itemCounts As Object
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim typeID As String
    Dim stationId As String
    Dim sellFormula As String
    Dim item As Variant
    Dim qty As Long
    Dim occurrences As Long
    Dim tbl As ListObject

    ' Create dictionary to store item counts
    Set itemCounts = CreateObject("Scripting.Dictionary")
    
    ' Set the source and destination sheets
    Set srcWs = ThisWorkbook.Sheets("All_Shipping_Items_RAW")
    Set destWs = ThisWorkbook.Sheets("M-M Banestar Market Comparison")
    
    ' Clear the destination sheet
    destWs.Cells.Clear
    
    ' Set station ID
    stationId = "1035466617946"
    
    ' Find the last row with data in column A of the source sheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' Add headers for new data in the destination sheet
    destWs.Cells(1, 1).Value = "Item"
    destWs.Cells(1, 2).Value = "Type ID"
    destWs.Cells(1, 3).Value = "Shipping Qty"
    destWs.Cells(1, 4).Value = "Occurrences"
    destWs.Cells(1, 5).Value = "Sell Orders Qty"
    destWs.Cells(1, 6).Value = "Difference (Sell - Ship)"
    
    ' Loop through each row in the source sheet to count occurrences and total quantities
    For i = 2 To lastRow
        item = srcWs.Cells(i, 1).Value
        typeID = srcWs.Cells(i, 2).Value
        qty = srcWs.Cells(i, 3).Value
        occurrences = srcWs.Cells(i, 4).Value
        
        ' Check if the item already exists in the dictionary
        If itemCounts.exists(item) Then
            itemCounts(item)(0) = itemCounts(item)(0) + occurrences ' Increment occurrences
            itemCounts(item)(1) = itemCounts(item)(1) + qty ' Increment quantity
        Else
            ' Add new item to the dictionary
            itemCounts.Add item, Array(occurrences, qty, typeID)
        End If
    Next i
    
    ' Populate the destination sheet with the data from the dictionary
    j = 2 ' Start populating from the second row
    For Each item In itemCounts.Keys
        destWs.Cells(j, 1).Value = item
        destWs.Cells(j, 2).Value = itemCounts(item)(2) ' Type ID
        destWs.Cells(j, 3).Value = itemCounts(item)(1) ' Total Quantity
        destWs.Cells(j, 4).Value = itemCounts(item)(0) ' Occurrences
        
        ' Create the formula for sell orders quantity with IFERROR
        sellFormula = "=IFERROR(SUM(EVEONLINE.MARKET_STRUCTURE_ORDERS(" & stationId & ", " & itemCounts(item)(2) & ", TRUE).volume_remain), 0)"
        
        ' Set the formula in the destination sheet
        destWs.Cells(j, 5).Formula2 = sellFormula
        
        ' Calculate the difference as Sell Orders Qty - Shipping Qty
        destWs.Cells(j, 6).Formula2 = "=" & destWs.Cells(j, 5).Address & " - " & destWs.Cells(j, 3).Address
        
        j = j + 1
    Next item
    
    ' Convert range to table to enable sorting
    Set tbl = destWs.ListObjects.Add(xlSrcRange, destWs.Range("A1:F" & destWs.Cells(destWs.Rows.Count, "A").End(xlUp).Row), , xlYes)
    tbl.Name = "MarketComparisonTable"
    
    ' Activate the destination worksheet before freezing panes
    destWs.Activate
    
    ' Freeze the header row
    destWs.Rows("2:2").Select
    ActiveWindow.FreezePanes = True

    ' Ensure the worksheet is activated
    destWs.Activate
    
    MsgBox "Market data formulas set successfully into M-M Banestar Market Comparison tab!", vbInformation
End Sub

�      �?                      �                �����      �?                      �                �����      �?                      �                �����      �?                      �                �����      �?                      �                                                                                                                                                                �GwV  �KwV                                       