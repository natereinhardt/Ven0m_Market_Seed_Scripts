Attribute VB_Name = "SortedMarketComparisonSheet"
Sub CreateSortedMarketComparisonSheet()
    Dim srcWs As Worksheet
    Dim sortedWs As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    Dim i As Long
    Dim typeID As String
    Dim item As Variant
    Dim tbl As ListObject

    ' Set the source worksheet
    Set srcWs = ThisWorkbook.Sheets("M-M Banestar Market Comparison")

    ' Create a new worksheet
    On Error Resume Next
    Set sortedWs = ThisWorkbook.Sheets("M-M Market Comparison Sorted")
    On Error GoTo 0

    If Not sortedWs Is Nothing Then
        Application.DisplayAlerts = False
        sortedWs.Delete
        Application.DisplayAlerts = True
    End If

    Set sortedWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    sortedWs.Name = "M-M Market Comparison Sorted"

    ' Add headers to the new sorted sheet
    sortedWs.Cells(1, 1).Value = "Item"
    sortedWs.Cells(1, 2).Value = "Diff"
    sortedWs.Cells(1, 3).Value = "Min 15% Mkup"
    sortedWs.Cells(1, 4).Value = "Med 15% Mkup"
    sortedWs.Cells(1, 5).Value = "Max 15% Mkup"
    sortedWs.Cells(1, 6).Value = "Min Profit"
    sortedWs.Cells(1, 7).Value = "Med Profit"
    sortedWs.Cells(1, 8).Value = "Max Profit"
    sortedWs.Cells(1, 9).Value = "Ship Qty"
    sortedWs.Cells(1, 10).Value = "Occur"
    sortedWs.Cells(1, 11).Value = "Sell Ord Qty"
    sortedWs.Cells(1, 12).Value = "Jita Min"
    sortedWs.Cells(1, 13).Value = "Jita Med"
    sortedWs.Cells(1, 14).Value = "Jita Max"
    sortedWs.Cells(1, 15).Value = "Type ID"
    sortedWs.Cells(1, 16).Value = "Raw Jita Data"

    ' Find the last row with data in the source sheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    newRow = 2 ' Start populating from the second row

    ' Copy rows with negative difference to the new sorted sheet and add prices
    For i = 2 To lastRow
        If srcWs.Cells(i, 6).Value < 0 Then
            sortedWs.Cells(newRow, 1).Value = srcWs.Cells(i, 1).Value
            sortedWs.Cells(newRow, 15).Value = srcWs.Cells(i, 2).Value
            sortedWs.Cells(newRow, 9).Value = srcWs.Cells(i, 3).Value
            sortedWs.Cells(newRow, 10).Value = srcWs.Cells(i, 4).Value
            sortedWs.Cells(newRow, 11).Value = srcWs.Cells(i, 5).Value
            sortedWs.Cells(newRow, 2).Value = srcWs.Cells(i, 6).Value

            ' Get the raw Jita data
            typeID = srcWs.Cells(i, 2).Value
            sortedWs.Cells(newRow, 16).Formula = "=EVEONLINE.MARKET_ORDERS_STATS(10000002, " & typeID & ", 60003760)"
            sortedWs.Cells(newRow, 12).Formula = "=IFERROR(RC[4].sell.min, 0)"
            sortedWs.Cells(newRow, 13).Formula = "=IFERROR(RC[3].sell.median, 0)"
            sortedWs.Cells(newRow, 14).Formula = "=IFERROR(RC[2].sell.max, 0)"
            sortedWs.Cells(newRow, 3).Formula = "=IFERROR(L" & newRow & "*1.15, 0)" ' Min 15% Mkup
            sortedWs.Cells(newRow, 4).Formula = "=IFERROR(M" & newRow & "*1.15, 0)" ' Med 15% Mkup
            sortedWs.Cells(newRow, 5).Formula = "=IFERROR(N" & newRow & "*1.15, 0)" ' Max 15% Mkup
            sortedWs.Cells(newRow, 6).Formula = "=IFERROR(C" & newRow & "-L" & newRow & ", 0)" ' Min Profit
            sortedWs.Cells(newRow, 7).Formula = "=IFERROR(D" & newRow & "-M" & newRow & ", 0)" ' Med Profit
            sortedWs.Cells(newRow, 8).Formula = "=IFERROR(E" & newRow & "-N" & newRow & ", 0)" ' Max Profit

            newRow = newRow + 1
        End If
    Next i

    ' Sort data by Occurrences (column 10) descending, then by Difference (column 2) ascending
    With sortedWs.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortedWs.Range("J2:J" & newRow - 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SortFields.Add key:=sortedWs.Range("B2:B" & newRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortedWs.Range("A1:P" & newRow - 1)
        .Header = xlYes
        .Apply
    End With

    ' Highlight cells
    sortedWs.Range("C2:C" & newRow - 1).Interior.Color = RGB(144, 238, 144) ' Light green for Min 15% Mkup
    sortedWs.Range("F2:F" & newRow - 1).Interior.Color = RGB(173, 255, 47) ' Very light green for Min Profit
    sortedWs.Range("D2:D" & newRow - 1).Interior.Color = RGB(255, 255, 224) ' Light yellow for Med 15% Mkup
    sortedWs.Range("G2:G" & newRow - 1).Interior.Color = RGB(255, 255, 153) ' Very light yellow for Med Profit
    sortedWs.Range("E2:E" & newRow - 1).Interior.Color = RGB(255, 182, 193) ' Light red for Max 15% Mkup
    sortedWs.Range("H2:H" & newRow - 1).Interior.Color = RGB(255, 204, 204) ' Very light red for Max Profit

    ' Convert the data range to a table
    Set tbl = sortedWs.ListObjects.Add(xlSrcRange, sortedWs.Range("A1:P" & newRow - 1), , xlYes)
    tbl.Name = "MarketComparisonTable"

    ' Apply table style
    tbl.TableStyle = "TableStyleLight9"

    ' Hide the Raw Jita Data and Type ID columns
    sortedWs.Columns("P").Hidden = True
    sortedWs.Columns("O").Hidden = True

    ' Set default width for price columns to accommodate values up to the billions
    sortedWs.Columns("C:H").ColumnWidth = 18
    sortedWs.Columns("L:N").ColumnWidth = 18

    ' Autofit all columns except the price columns
    sortedWs.Columns("A:B").AutoFit
    sortedWs.Columns("I:K").AutoFit
    sortedWs.Rows("2:" & newRow - 1).AutoFit

    ' Format the price columns as currency with decimals
    sortedWs.Range("C2:C" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("D2:D" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("E2:E" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("F2:F" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("G2:G" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("H2:H" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("L2:L" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("M2:M" & newRow - 1).NumberFormat = "$#,##0.00"
    sortedWs.Range("N2:N" & newRow - 1).NumberFormat = "$#,##0.00"

    MsgBox "Sorted market comparison table with negative differences created successfully in M-M Market Comparison Sorted tab!", vbInformation
End Sub

 a t e m   U S E R P R O F I L E = C : \ U s e r s \ n a t e m   W E B V I E W 2 _ A D D I T I O N A L _ B R O W S E R _ A R G U M E N T S =   - - e n a b l e - f e a t u r e s = m s A b y d o s , m s A b y d o s G e s t u r e S u p p o r t , m s A b y d o s H a n d w r i t i n g A t t r   W e c V e r s i o n F o r R o s e b u d . 6 B 5 C = 6   w i n d 