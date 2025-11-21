Sub ConsolidateSalesData()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsConsolidated As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, i As Long, nextRow As Long
    Dim mainRange As Range, dupRange As Range

    ' Set worksheets
    Set ws1 = ThisWorkbook.Sheets("Sales_Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sales_Sheet2")

    ' Create/clear consolidated sheet
    On Error Resume Next
    Set wsConsolidated = ThisWorkbook.Sheets("Consolidated")
    If wsConsolidated Is Nothing Then
        Set wsConsolidated = ThisWorkbook.Sheets.Add
        wsConsolidated.Name = "Consolidated"
    End If
    wsConsolidated.Cells.Clear
    On Error GoTo 0

    ' Copy headers from first sheet
    ws1.Rows(1).Copy wsConsolidated.Rows(1)
    nextRow = 2

    ' Copy data from Sheet1
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    ws1.Rows("2:" & lastRow1).Copy wsConsolidated.Rows(nextRow)
    nextRow = nextRow + (lastRow1 - 1)

    ' Copy data from Sheet2
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    ws2.Rows("2:" & lastRow2).Copy wsConsolidated.Rows(nextRow)

    ' Remove duplicate records
    lastRow2 = wsConsolidated.Cells(wsConsolidated.Rows.Count, 1).End(xlUp).Row
    Set mainRange = wsConsolidated.Range("A1:E" & lastRow2)
    mainRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5), Header:=xlYes
End Sub

Sub CalculateSalesMetrics()
    Dim wsConsolidated As Worksheet
    Dim lastRow As Long, i As Long
    Dim totalSales As Double, salesDays As Object
    Dim bestProduct As String, dict As Object, avgDailySales As Double
    Dim maxSale As Double, prod As Variant

    Set wsConsolidated = ThisWorkbook.Sheets("Consolidated")
    lastRow = wsConsolidated.Cells(wsConsolidated.Rows.Count, 1).End(xlUp).Row

    Set dict = CreateObject("Scripting.Dictionary")
    Set salesDays = CreateObject("Scripting.Dictionary")
    totalSales = 0

    For i = 2 To lastRow
        ' Add Date to unique day list
        salesDays(wsConsolidated.Cells(i, 1).Value) = 1
        
        ' Sum Total Sales
        totalSales = totalSales + wsConsolidated.Cells(i, 5).Value
        prod = wsConsolidated.Cells(i, 2).Value
        If dict.Exists(prod) Then
            dict(prod) = dict(prod) + wsConsolidated.Cells(i, 5).Value
        Else
            dict.Add prod, wsConsolidated.Cells(i, 5).Value
        End If
    Next i

    avgDailySales = totalSales / salesDays.Count

    maxSale = 0
    bestProduct = ""
    For Each prod In dict.Keys
        If dict(prod) > maxSale Then
            maxSale = dict(prod)
            bestProduct = prod
        End If
    Next prod

    ' Store metrics temporarily in Consolidated sheet
    wsConsolidated.Range("G1").Value = "Total Sales"
    wsConsolidated.Range("G2").Value = totalSales
    wsConsolidated.Range("G3").Value = "Avg Daily Sales"
    wsConsolidated.Range("G4").Value = avgDailySales
    wsConsolidated.Range("G5").Value = "Best Product"
    wsConsolidated.Range("G6").Value = bestProduct
    wsConsolidated.Range("G7").Value = "Best Product Sales"
    wsConsolidated.Range("G8").Value = maxSale
End Sub

Sub GenerateSalesSummaryReport()
    Dim wsReport As Worksheet, wsConsolidated As Worksheet

    Set wsConsolidated = ThisWorkbook.Sheets("Consolidated")

    ' Create report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("SummaryReport")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "SummaryReport"
    End If
    wsReport.Cells.Clear
    On Error GoTo 0

    ' Write metrics
    With wsReport
        .Range("A1").Value = "Sales Summary Report"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1:B1").Merge

        .Range("A3").Value = "Total Sales"
        .Range("B3").Value = wsConsolidated.Range("G2").Value
        .Range("A4").Value = "Average Daily Sales"
        .Range("B4").Value = wsConsolidated.Range("G4").Value
        .Range("A5").Value = "Best-Selling Product"
        .Range("B5").Value = wsConsolidated.Range("G6").Value
        .Range("A6").Value = "Best Product Sales"
        .Range("B6").Value = wsConsolidated.Range("G8").Value

        ' Formatting
        .Range("A3:B6").Font.Bold = True
        .Range("B3").Interior.Color = RGB(198, 239, 206)
        .Range("B4").Interior.Color = RGB(255, 235, 156)
        .Range("B5:B6").Interior.Color = RGB(189, 215, 238)
        .Columns("A:B").AutoFit
        .Range("A3:B6").HorizontalAlignment = xlCenter
    End With
End Sub

Sub RunFullSalesAutomation()
    On Error GoTo SafeExit
    Call ConsolidateSalesData
    Call CalculateSalesMetrics
    Call GenerateSalesSummaryReport
    Exit Sub
SafeExit:
    MsgBox "Error: " & Err.Description, vbCritical, "Sales Automation"
End Sub

Sub EmailSalesSummary()
    Dim olApp As Object, olMail As Object
    Dim wsReport As Worksheet, msgBody As String
    Dim i As Long

    Set wsReport = ThisWorkbook.Sheets("SummaryReport")
    msgBody = "Sales Summary Report:" & vbCrLf & vbCrLf

    ' Build message body from summary
    For i = 3 To 6
        msgBody = msgBody & wsReport.Cells(i, 1).Value & ": " & wsReport.Cells(i, 2).Value & vbCrLf
    Next i

    ' Requires Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    olMail.To = "recipient@example.com"          ' <-- Change recipient(s)
    olMail.Subject = "Automated Sales Report"
    olMail.Body = msgBody
    olMail.Send
End Sub
