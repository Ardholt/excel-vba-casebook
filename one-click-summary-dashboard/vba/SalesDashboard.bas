Attribute VB_Name = "Module1"
Option Explicit

' Integer-cent arithmetic; no external references, automatic events, or file access.
' Derived sheets are rebuilt. Sales is never changed by this module.
Public Sub InstallRefreshButton()
    Dim ws As Worksheet, button As Shape
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    Set button = ws.Shapes("RefreshSales")
    On Error GoTo 0
    If button Is Nothing Then
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, 385, 8, 200, 38)
        button.Name = "RefreshSales"
    End If
    button.TextFrame.Characters.Text = "Refresh sales"
    button.Fill.ForeColor.RGB = RGB(8, 127, 115)
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!RefreshDashboard"
End Sub

Public Sub RefreshDashboard()
    Dim result As String
    result = RefreshSalesCore(True)
    If Left$(result, 4) <> "PASS" Then MsgBox result, vbExclamation, "Sales dashboard"
End Sub

Public Function RefreshDashboardHeadless() As String
    RefreshDashboardHeadless = RefreshSalesCore(False)
End Function

Private Function RefreshSalesCore(ByVal interactive As Boolean) As String
    Dim src As Worksheet, valid As Worksheet, audit As Worksheet, dash As Worksheet
    Dim headers As Variant, col(0 To 4) As Long, found As Long, headerRow As Long
    Dim last As Range, lastRow As Long, lastCol As Long, r As Long, c As Long, j As Long
    Dim n As Long, i As Long, k As Long, pos As Long, validRow As Long, reason As String
    Dim data() As String, sourceRow() As Long, bad() As String, dupe() As Boolean
    Dim validatedCents() As Double
    Dim ordering() As Long, paid As Double, pending As Double, amount As Double
    Dim paidN As Long, pendingN As Long, cancelN As Long, rejectedN As Long
    Dim months() As String, categories() As String, monthN As Long, catN As Long
    Dim counts(0 To 4) As Long, nonblank As Boolean, anyFormula As Boolean, anyError As Boolean
    Dim endRow As Long, catStart As Long, prevScreen As Boolean, prevEvents As Boolean
    Dim previousCalculation As XlCalculation, totalByMonth As Double, totalByCategory As Double
    On Error GoTo Failed
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    previousCalculation = Application.Calculation
    Set src = ThisWorkbook.Worksheets("Sales")
    headers = Array("sale_id", "date", "category", "amount", "status")
    Set last = src.Cells.Find(What:="*", After:=src.Cells(1, 1), LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If last Is Nothing Then Err.Raise 2101, , "Sales has no headers."
    lastRow = last.Row
    Set last = src.Cells.Find(What:="*", After:=src.Cells(1, 1), LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lastCol = last.Column
    For r = 1 To WorksheetFunction.Min(lastRow, 50)
        For j = 0 To 4: counts(j) = 0: col(j) = 0: Next j
        For c = 1 To lastCol
            For j = 0 To 4
                If SafeText(src.Cells(r, c).Value2) = headers(j) Then
                    counts(j) = counts(j) + 1: col(j) = c
                End If
            Next j
        Next c
        found = 0
        For j = 0 To 4
            If counts(j) > 0 Then found = found + 1
        Next j
        If found = 5 Then
            For j = 0 To 4
                If counts(j) <> 1 Then Err.Raise 2102, , "Repeated required header."
            Next j
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Err.Raise 2103, , "Required headers not found in first 50 rows."
    ReDim data(1 To lastRow + 1, 0 To 4)
    ReDim sourceRow(1 To lastRow + 1)
    ReDim bad(1 To lastRow + 1)
    ReDim validatedCents(1 To lastRow + 1)
    ReDim dupe(1 To lastRow + 1)
    ReDim ordering(1 To lastRow + 1)
    ReDim months(1 To lastRow + 1)
    ReDim categories(1 To lastRow + 1)
    For r = headerRow + 1 To lastRow
        nonblank = False: anyFormula = False: anyError = False
        For j = 0 To 4
            If SafeText(src.Cells(r, col(j)).Value2) <> "" Then nonblank = True
            If src.Cells(r, col(j)).HasFormula Then anyFormula = True
            If IsError(src.Cells(r, col(j)).Value2) Then anyError = True
        Next j
        If nonblank Or anyFormula Or anyError Then
            n = n + 1: sourceRow(n) = r: ordering(n) = n
            For j = 0 To 4: data(n, j) = SafeText(src.Cells(r, col(j)).Value2): Next j
            If anyError Then bad(n) = "Invalid cell value"
            If anyFormula Then bad(n) = "Formula in source row: use literal values"
        End If
    Next r
    If n > 1 Then
        SortIndices ordering, data, 1, n
        For i = 2 To n
            If data(ordering(i), 0) <> "" Then
                If StrComp(data(ordering(i), 0), data(ordering(i - 1), 0), vbBinaryCompare) = 0 Then
                    dupe(ordering(i)) = True: dupe(ordering(i - 1)) = True
                End If
            End If
        Next i
    End If
    ' Validate everything before changing derived sheets.
    For i = 1 To n
        reason = ""
        If data(i, 0) = "" Then
            reason = "Missing sale ID"
        ElseIf dupe(i) Then
            reason = "Duplicate sale ID: all occurrences excluded"
        ElseIf bad(i) <> "" Then
            reason = bad(i)
        ElseIf Not IsISODate(data(i, 1)) Then
            reason = "Invalid ISO date"
        ElseIf data(i, 2) = "" Then
            reason = "Missing category"
        ElseIf Not ReadAmount(src.Cells(sourceRow(i), col(3)).Value2, amount) Then
            reason = "Invalid amount: decimal with up to two places required"
        ElseIf data(i, 4) <> "Paid" And data(i, 4) <> "Pending" And data(i, 4) <> "Cancelled" Then
            reason = "Invalid status"
        End If
        bad(i) = reason
        If reason = "" Then validatedCents(i) = amount
    Next i
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set valid = ThisWorkbook.Worksheets("ValidSales")
    Set audit = ThisWorkbook.Worksheets("Audit")
    Set dash = ThisWorkbook.Worksheets("Dashboard")
    dash.UsedRange.UnMerge
    valid.Cells.Clear: audit.Cells.Clear: dash.Cells.Clear
    Do While dash.ChartObjects.Count > 0: dash.ChartObjects(1).Delete: Loop
    valid.Cells.NumberFormat = "@": audit.Cells.NumberFormat = "@"
    valid.Range("A1:G1").Value = Array("source_row", "sale_id", "date", "category", "amount", "status", "month")
    audit.Range("A1:H1").Value = Array("source_row", "sale_id", "date", "category", "amount", "status", "disposition", "reason")
    validRow = 1
    For i = 1 To n
        audit.Cells(i + 1, 1).Value2 = sourceRow(i)
        For j = 0 To 4: audit.Cells(i + 1, j + 2).Value2 = data(i, j): Next j
        audit.Cells(i + 1, 8).Value2 = bad(i)
        If bad(i) <> "" Then
            audit.Cells(i + 1, 7).Value2 = "Rejected": rejectedN = rejectedN + 1
        Else
            audit.Cells(i + 1, 7).Value2 = data(i, 4)
            amount = validatedCents(i)
            validRow = validRow + 1
            valid.Cells(validRow, 1).Value2 = sourceRow(i)
            For j = 0 To 4: valid.Cells(validRow, j + 2).Value2 = data(i, j): Next j
            valid.Cells(validRow, 5).NumberFormat = "0.00"
            valid.Cells(validRow, 5).Value2 = CDbl(amount / 100)
            valid.Cells(validRow, 7).Value2 = Left$(data(i, 1), 7)
            Select Case data(i, 4)
                Case "Paid"
                    paid = paid + amount: paidN = paidN + 1
                    If Abs(paid) > 9007199254740991# Then Err.Raise 2107, , "Paid total exceeds exact-cent range."
                    AddLabel months, monthN, Left$(data(i, 1), 7)
                    AddLabel categories, catN, data(i, 2)
                Case "Pending"
                    pending = pending + amount: pendingN = pendingN + 1
                    If Abs(pending) > 9007199254740991# Then Err.Raise 2108, , "Pending total exceeds exact-cent range."
                Case "Cancelled": cancelN = cancelN + 1
            End Select
        End If
    Next i
    SortLabels months, monthN: SortLabels categories, catN
    endRow = WorksheetFunction.Max(2, validRow)
    With dash
        .Columns("A").ColumnWidth = 29: .Columns("B").ColumnWidth = 21
        .Range("A1:H" & monthN + catN + 20).Font.Name = "Calibri"
        .Range("A1:H" & monthN + catN + 20).Font.Size = 14
        .Rows("1:" & monthN + catN + 20).RowHeight = 25
        .Range("A1:B1").Merge: .Range("A1").Value2 = "SALES / REFRESHED"
        .Range("A1").Font.Size = 22: .Range("A1").Font.Bold = True
        .Range("A2:B2").Merge: .Range("A2").Value2 = "Paid revenue - GBP - returns included"
        .Range("A2").Font.Size = 12
        .Range("A3:B4").Merge
        .Range("A3").Formula = "=SUMIF(ValidSales!F2:F" & endRow & ",""Paid"",ValidSales!E2:E" & endRow & ")"
        .Range("A3").Font.Size = 30: .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(232, 243, 241)
        .Range("A3").Font.Color = RGB(8, 127, 115)
        .Range("A6").Value2 = "Paid transactions"
        .Range("B6").Formula = "=COUNTIF(ValidSales!F2:F" & endRow & ",""Paid"")"
        .Range("A7").Value2 = "Pending - kept separate"
        .Range("B7").Formula = "=SUMIF(ValidSales!F2:F" & endRow & ",""Pending"",ValidSales!E2:E" & endRow & ")"
        .Range("A8").Value2 = "Cancelled - excluded": .Range("B8").Value2 = cancelN
        .Range("A9").Value2 = "Rows to review": .Range("B9").Value2 = rejectedN
        .Range("A11:B11").Value = Array("Paid by month", "Revenue")
        For i = 1 To monthN
            r = 11 + i
            .Cells(r, 1).NumberFormat = "@": .Cells(r, 1).Value2 = months(i)
            .Cells(r, 2).Formula = SummaryFormula(endRow, "G", r)
        Next i
        catStart = 13 + WorksheetFunction.Max(1, monthN)
        .Cells(catStart, 1).Value2 = "Paid by category": .Cells(catStart, 2).Value2 = "Revenue"
        For i = 1 To catN
            r = catStart + i
            .Cells(r, 1).NumberFormat = "@": .Cells(r, 1).Value2 = categories(i)
            .Cells(r, 2).Formula = SummaryFormula(endRow, "D", r)
        Next i
        .Range("A3").NumberFormat = MoneyFormat()
        .Range("B7").NumberFormat = MoneyFormat()
        .Range("B12:B" & catStart + catN).NumberFormat = MoneyFormat()
        StyleHeading .Range("A11:B11")
        StyleHeading .Range(.Cells(catStart, 1), .Cells(catStart, 2))
        Application.Calculate
        If Round(CDbl(.Range("A3").Value2) * 100, 0) <> paid Or CLng(.Range("B6").Value2) <> paidN Or Round(CDbl(.Range("B7").Value2) * 100, 0) <> pending Then
            Err.Raise 2104, , "Dashboard totals do not reconcile: paid=" & CStr(paid / 100) & ", cell=" & CStr(.Range("A3").Value2) & ", count=" & CStr(paidN) & "/" & CStr(.Range("B6").Value2) & ", pending=" & CStr(pending / 100) & "/" & CStr(.Range("B7").Value2)
        End If
        For i = 1 To monthN: totalByMonth = totalByMonth + Round(CDbl(.Cells(11 + i, 2).Value2) * 100, 0): Next i
        For i = 1 To catN: totalByCategory = totalByCategory + Round(CDbl(.Cells(catStart + i, 2).Value2) * 100, 0): Next i
        If totalByMonth <> paid Or totalByCategory <> paid Then Err.Raise 2105, , "Breakdowns do not reconcile."
        If n <> paidN + pendingN + cancelN + rejectedN Then Err.Raise 2106, , "Row counts do not reconcile."
        .Range("A10:B10").Merge: .Range("A10").Value2 = "Checked: all source rows accounted for"
        .Range("A10").Font.Size = 12
        .PageSetup.PrintArea = "$A$1:$B$9"
    End With
    If monthN > 0 Then AddChart dash, 11, 11 + monthN, "Paid by month", 70
    If catN > 0 Then AddChart dash, catStart, catStart + catN, "Paid by category", 380
    valid.Columns.AutoFit: audit.Columns.AutoFit
    StyleHeading valid.Range("A1:G1"): StyleHeading audit.Range("A1:H1")
    If interactive Then
        InstallRefreshButton
        dash.Activate
    End If
    Application.Calculation = previousCalculation
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    RefreshSalesCore = "PASS"
    Exit Function
Failed:
    Dim detail As String
    detail = Err.Description
    On Error Resume Next
    If Not dash Is Nothing Then dash.Range("A10").Value2 = "REFRESH FAILED - correct source and retry"
    Application.Calculation = previousCalculation
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    RefreshSalesCore = "FAIL: " & detail
End Function


Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then
        SafeText = "#CELL-ERROR"
    ElseIf IsEmpty(value) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(value))
    End If
End Function

Private Function IsISODate(ByVal text As String) As Boolean
    Dim y As Long, m As Long, d As Long, result As Date
    On Error GoTo Invalid
    If Len(text) <> 10 Then Exit Function
    If Mid$(text, 5, 1) <> "-" Or Mid$(text, 8, 1) <> "-" Then Exit Function
    If Not OnlyDigits(Left$(text, 4) & Mid$(text, 6, 2) & Right$(text, 2)) Then Exit Function
    y = CLng(Left$(text, 4)): m = CLng(Mid$(text, 6, 2)): d = CLng(Right$(text, 2))
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, m, d)
    If Year(result) <> y Then Exit Function
    If Month(result) <> m Then Exit Function
    If Day(result) <> d Then Exit Function
    IsISODate = True
Invalid:
End Function

Private Function OnlyDigits(ByVal text As String) As Boolean
    Dim i As Long, code As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next i
    OnlyDigits = True
End Function

Private Function ReadAmount(ByVal value As Variant, ByRef result As Double) As Boolean
    Dim text As String, whole As String, fraction As String, digits As String
    Dim dot As Long, i As Long, negative As Boolean, number As Double
    On Error GoTo Invalid
    If IsError(value) Or IsEmpty(value) Or VarType(value) = vbBoolean Then Exit Function
    If VarType(value) <> vbString Then
        If Not IsNumeric(value) Then Exit Function
        number = CDbl(value)
        If Abs(number) > 999999999 Then Exit Function
        If number <> Round(number, 2) Then Exit Function
        result = Round(number * 100, 0)
        ReadAmount = True
        Exit Function
    End If
    text = Trim$(CStr(value))
    If Left$(text, 1) = "-" Then negative = True: text = Mid$(text, 2)
    dot = InStr(1, text, ".", vbBinaryCompare)
    If dot = 0 Then
        whole = text: fraction = "00"
    Else
        whole = Left$(text, dot - 1): fraction = Mid$(text, dot + 1)
        If Len(fraction) < 1 Or Len(fraction) > 2 Then Exit Function
        fraction = Left$(fraction & "0", 2)
    End If
    If Not OnlyDigits(whole) Or Not OnlyDigits(fraction) Then Exit Function
    digits = whole & fraction
    result = 0
    For i = 1 To Len(digits)
        result = result * 10 + AscW(Mid$(digits, i, 1)) - 48
        If result > 99999999900# Then Exit Function
    Next i
    If negative Then result = -result
    ReadAmount = True
Invalid:
End Function

Private Sub SortIndices(ByRef a() As Long, ByRef data() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As String, swap As Long
    i = lo: j = hi: pivot = data(a((lo + hi) \ 2), 0)
    Do While i <= j
        Do While StrComp(data(a(i), 0), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(data(a(j), 0), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            swap = a(i): a(i) = a(j): a(j) = swap: i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortIndices a, data, lo, j
    If i < hi Then SortIndices a, data, i, hi
End Sub

Private Sub AddLabel(ByRef labels() As String, ByRef count As Long, ByVal label As String)
    Dim i As Long
    For i = 1 To count
        If StrComp(labels(i), label, vbBinaryCompare) = 0 Then Exit Sub
    Next i
    count = count + 1: labels(count) = label
End Sub

Private Sub SortLabels(ByRef labels() As String, ByVal count As Long)
    Dim i As Long, j As Long, value As String
    For i = 2 To count
        value = labels(i): j = i - 1
        Do While j >= 1
            If StrComp(labels(j), value, vbBinaryCompare) <= 0 Then Exit Do
            labels(j + 1) = labels(j): j = j - 1
        Loop
        labels(j + 1) = value
    Next i
End Sub

Private Function SummaryFormula(ByVal lastRow As Long, ByVal keyCol As String, ByVal row As Long) As String
    SummaryFormula = "=SUMPRODUCT(--EXACT(ValidSales!" & keyCol & "2:" & keyCol & lastRow & ",A" & row & _
                     "),--(ValidSales!F2:F" & lastRow & "=""Paid""),ValidSales!E2:E" & lastRow & ")"
End Function

Private Function MoneyFormat() As String
    MoneyFormat = ChrW(163) & "#,##0.00;[Red]-" & ChrW(163) & "#,##0.00"
End Function

Private Sub StyleHeading(ByVal target As Range)
    target.Interior.Color = RGB(232, 243, 241)
    target.Font.Color = RGB(20, 47, 67)
    target.Font.Bold = True
End Sub

Private Sub AddChart(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal title As String, ByVal top As Double)
    Dim item As ChartObject
    Set item = ws.ChartObjects.Add(385, top, 480, 280)
    item.Chart.ChartType = xlColumnClustered
    item.Chart.SetSourceData Source:=ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 2))
    item.Chart.HasTitle = True: item.Chart.ChartTitle.Text = title
    item.Chart.HasLegend = False
End Sub
