Attribute VB_Name = "Module1"
Option Explicit
Option Compare Binary

' Run manually. Config!B2 is a folder path; relative paths start beside this workbook.
' Input workbooks are opened read-only with links disabled. No files are deleted or saved.
Public Sub RefreshFolder()
    Dim root As String, sep As String, name As String, names() As String
    Dim n As Long, i As Long, j As Long, swap As String
    Dim dest As Worksheet, audit As Worksheet, cfg As Worksheet
    Dim nextRow As Long, nextAudit As Long, oldEvents As Boolean, oldScreen As Boolean
    Dim oldSecurity As Long
    On Error GoTo Fatal
    Set cfg = ThisWorkbook.Worksheets("Config")
    Set dest = ThisWorkbook.Worksheets("Combined")
    Set audit = ThisWorkbook.Worksheets("Manifest")
    root = Trim(CStr(cfg.Cells(2, 2).Value))
    If Len(root) = 0 Then Err.Raise 5, , "Set Config B2 to the input folder"
    sep = "/"
    If InStr(ThisWorkbook.Path, "\") > 0 Then sep = "\"
    If InStr(root, "/") > 0 Then sep = "/"
    If InStr(root, "\") > 0 Then sep = "\"
    If Left(root, 1) <> "/" And InStr(root, ":") = 0 And Left(root, 2) <> "\\" Then
        root = ThisWorkbook.Path & sep & root
    End If
    If Right(root, 1) <> sep Then root = root & sep
    oldScreen = Application.ScreenUpdating
    ' Some compatibility hosts omit Excel's event/security properties.
    ' Only XLSX inputs are accepted; source books cannot contain a VBA project.
    On Error Resume Next
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    oldSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Err.Clear
    On Error GoTo Fatal
    Application.ScreenUpdating = False
    cfg.Cells(3, 2).Value = "Running"
    ' Collect names before opening workbooks: Workbooks.Open may reset Dir's cursor.
    name = Dir(root & "*")
    Do While Len(name) > 0
        If (GetAttr(root & name) And vbDirectory) = 0 Then
            n = n + 1
            ReDim Preserve names(1 To n)
            names(n) = name
        End If
        name = Dir()
    Loop
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(names(i), names(j), vbBinaryCompare) > 0 Then
                swap = names(i): names(i) = names(j): names(j) = swap
            End If
        Next j
    Next i
    dest.Cells.ClearContents
    audit.Cells.ClearContents
    Dim fixed As Variant
    fixed = Array("record_id", "date", "category", "amount", "source_file", "source_row")
    For i = 0 To UBound(fixed)
        dest.Cells(1, i + 1).Value = fixed(i)
    Next i
    fixed = Array("file", "status", "rows", "reason")
    For i = 0 To UBound(fixed)
        audit.Cells(1, i + 1).Value = fixed(i)
    Next i
    nextRow = 2: nextAudit = 2
    For i = 1 To n
        ProcessFile root, names(i), dest, audit, nextRow, nextAudit
    Next i
    cfg.Cells(3, 2).Value = "Complete"
    cfg.Cells(4, 2).Value = nextRow - 2
    dest.Rows(1).Font.Bold = True
    audit.Rows(1).Font.Bold = True
    dest.Columns(4).NumberFormat = "#,##0.00"
Restore:
    On Error Resume Next
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Fatal:
    If Not cfg Is Nothing Then cfg.Cells(3, 2).Value = "Stopped: error " & CStr(Err.Number)
    Resume Restore
End Sub

Private Function HeaderColumn(ws As Worksheet, key As String) As Long
    Dim c As Long, last As Long
    last = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For c = 1 To last
        If Trim(CStr(ws.Cells(1, c).Value)) = key Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Sub AddAudit(ws As Worksheet, ByRef pos As Long, file As String, status As String, count As Long, reason As String)
    ws.Cells(pos, 1).Value = file
    ws.Cells(pos, 2).Value = status
    ws.Cells(pos, 3).Value = count
    ws.Cells(pos, 4).Value = reason
    pos = pos + 1
End Sub

Private Sub ProcessFile(root As String, file As String, dest As Worksheet, audit As Worksheet, ByRef pos As Long, ByRef auditPos As Long)
    Dim book As Workbook, src As Worksheet, sheet As Worksheet
    Dim cols As Long, last As Long, c As Long, r As Long, k As Long, count As Long
    Dim key As String, reason As String, status As String, mapping() As Long, hasValue As Boolean
    Dim firstRow As Long, v As Variant, sourceData As Variant, resultData() As Variant
    Dim outCols As Long, outRow As Long
    firstRow = pos
    If Left(file, 2) = "~$" Or LCase(file) = "consolidated.xlsx" Or LCase(file) = "vba-result.xlsx" Or LCase(file) = "combine.xlsm" Then
        AddAudit audit, auditPos, file, "ignored", 0, "Temporary or output file"
        Exit Sub
    End If
    If LCase(Right(file, 5)) <> ".xlsx" Then
        AddAudit audit, auditPos, file, "ignored", 0, "Not an .xlsx workbook"
        Exit Sub
    End If
    On Error GoTo Unreadable
    Set book = Workbooks.Open(root & file, 0, True)
    For Each sheet In book.Worksheets
        If sheet.Name = "Data" Then Set src = sheet
    Next sheet
    If src Is Nothing Then
        reason = "Missing Data sheet": GoTo Reject
    End If
    cols = src.UsedRange.Column + src.UsedRange.Columns.Count - 1
    last = src.UsedRange.Row + src.UsedRange.Rows.Count - 1
    ReDim mapping(1 To cols)
    For c = 1 To cols
        key = Trim(CStr(src.Cells(1, c).Value))
        If Len(key) = 0 Then reason = "Blank header": GoTo Reject
        If key = "source_file" Or key = "source_row" Then reason = "Reserved provenance header": GoTo Reject
        For k = 1 To c - 1
            If Trim(CStr(src.Cells(1, k).Value)) = key Then reason = "Duplicate header": GoTo Reject
        Next k
    Next c
    v = Array("record_id", "date", "category", "amount")
    For k = 0 To UBound(v)
        If HeaderColumn(src, CStr(v(k))) = 0 Then reason = "Missing required columns": GoTo Reject
    Next k
    For c = 1 To cols
        key = Trim(CStr(src.Cells(1, c).Value))
        k = HeaderColumn(dest, key)
        If k = 0 Then
            k = 7
            Do While Len(CStr(dest.Cells(1, k).Value)) > 0
                k = k + 1
            Loop
            dest.Cells(1, k).Value = key
        End If
        mapping(c) = k
    Next c
    If last >= 2 Then
        sourceData = src.Range(src.Cells(1, 1), src.Cells(last, cols)).Value
        For r = 2 To last
            hasValue = False
            For c = 1 To cols
                If Len(CStr(sourceData(r, c))) > 0 Then hasValue = True
            Next c
            If hasValue Then count = count + 1
        Next r
        If count > 0 Then
            outCols = 6
            For c = 1 To cols
                If mapping(c) > outCols Then outCols = mapping(c)
            Next c
            ReDim resultData(1 To count, 1 To outCols)
            outRow = 0
            For r = 2 To last
                hasValue = False
                For c = 1 To cols
                    If Len(CStr(sourceData(r, c))) > 0 Then hasValue = True
                Next c
                If hasValue Then
                    outRow = outRow + 1
                    For c = 1 To cols
                        resultData(outRow, mapping(c)) = sourceData(r, c)
                    Next c
                    resultData(outRow, 5) = file
                    resultData(outRow, 6) = r
                End If
            Next r
            dest.Range(dest.Cells(pos, 1), dest.Cells(pos + count - 1, outCols)).NumberFormat = "@"
            dest.Range(dest.Cells(pos, 1), dest.Cells(pos + count - 1, outCols)).Value = resultData
            pos = pos + count
        End If
    End If
    status = "accepted"
    If count = 0 Then status = "empty"
    book.Close False
    AddAudit audit, auditPos, file, status, count, ""
    Exit Sub
Unreadable:
    reason = "Unreadable workbook"
    Resume Reject
Reject:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If pos > firstRow Then dest.Range(dest.Cells(firstRow, 1), dest.Cells(pos - 1, dest.UsedRange.Columns.Count)).ClearContents
    pos = firstRow
    On Error GoTo 0
    AddAudit audit, auditPos, file, "rejected", 0, reason
End Sub
