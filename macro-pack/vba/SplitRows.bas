Attribute VB_Name = "SplitRowsModule"
Option Explicit

Public Function HeaderRow(ByVal ws As Object, ByVal names As String) As Long
    Dim r As Long, c As Long, part As Variant, ok As Boolean, found As Long
    For r = 1 To ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
        ok = True
        For Each part In Split(names, "|")
            found = 0
            For c = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
                If Trim(CStr(ws.Cells(r, c).Value)) = CStr(part) Then found = found + 1
            Next c
            If found <> 1 Then ok = False
        Next part
        If ok Then HeaderRow = r: Exit Function
    Next r
    Err.Raise 5, , "Required headers missing or repeated"
End Function

Public Function Col(ByVal ws As Object, ByVal r As Long, ByVal name As String) As Long
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If Trim(CStr(ws.Cells(r, c).Value)) = name Then Col = c: Exit Function
    Next c
    Err.Raise 5, , "Header not found"
End Function

Public Function LastRow(ByVal ws As Object) As Long
    LastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

Public Function Exists(ByVal name As String) As Boolean
    Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) = LCase(name) Then Exists = True: Exit Function
    Next ws
End Function

Public Function OutSheet(ByVal name As String) As Object
    Dim ws As Object
    If Exists(name) Then
        Set ws = ThisWorkbook.Worksheets(name)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    Set OutSheet = ws
End Function

Public Sub Head(ByVal ws As Object, ByVal names As String)
    Dim p As Variant, c As Long
    c = 1
    For Each p In Split(names, "|")
        ws.Cells(1, c).Value = p
        c = c + 1
    Next p
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:C").ColumnWidth = 24
End Sub

Public Function IsBlankRow(ByVal ws As Object, ByVal r As Long) As Boolean
    Dim c As Long
    IsBlankRow = True
    For c = 1 To ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        If CStr(ws.Cells(r, c).Value) <> "" Then IsBlankRow = False: Exit Function
    Next c
End Function

Public Sub CheckData(ByVal ws As Object)
    Dim hr As Long, r As Long, id As Long, g As Long, v As Long
    hr = HeaderRow(ws, "record_id|group|value")
    id = Col(ws, hr, "record_id"): g = Col(ws, hr, "group"): v = Col(ws, hr, "value")
    For r = hr + 1 To LastRow(ws)
        If Not IsBlankRow(ws, r) Then
            If CStr(ws.Cells(r, id).Value) = "" Or CStr(ws.Cells(r, g).Value) = "" Or Not IsNumeric(ws.Cells(r, v).Value) Then Err.Raise 5, , "Invalid data row"
        End If
    Next r
End Sub

Public Sub SplitRows()
    Dim result As String
    result = RunSplit()
    If Left(result, 5) = "ERROR" Then Err.Raise 5, , result
End Sub

Public Function SafeGroupFile(ByVal name As String) As Boolean
    Dim i As Long, digits As String
    If Left(name, 6) <> "group-" Or Right(name, 5) <> ".xlsx" Then Exit Function
    digits = Mid(name, 7, Len(name) - 11)
    If Len(digits) < 3 Then Exit Function
    For i = 1 To Len(digits)
        If InStr("0123456789", Mid(digits, i, 1)) = 0 Then Exit Function
    Next i
    SafeGroupFile = True
End Function

Public Sub SaveGroup(ByVal filename As String, ByRef data As Variant, ByVal count As Long)
    Dim book As Object, document As Object, desktop As Object, sheet As Object
    Dim priorAlerts As Boolean, changedAlerts As Boolean, message As String
    On Error GoTo FailedSave
    ' Excel excludes this adapter. Calc 24.2 ignores conditional directives,
    ' executes the native export below, and exits before the Excel SaveAs.
#If False Then
    Dim args(0) As New com.sun.star.beans.PropertyValue
    Dim settings(1) As New com.sun.star.beans.PropertyValue
    Dim r As Long, c As Long
    args(0).Name = "Hidden": args(0).Value = True
    Set desktop = CreateUnoService("com.sun.star.frame.Desktop")
    Set document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, args())
    Set sheet = document.Sheets.getByIndex(0): sheet.Name = "Data"
    For r = 1 To count + 1
        For c = 1 To 3
            If r > 1 And c = 3 Then
                sheet.getCellByPosition(c - 1, r - 1).setValue(CDbl(data(r, c)))
            Else
                sheet.getCellByPosition(c - 1, r - 1).setString(CStr(data(r, c)))
            End If
        Next c
    Next r
    settings(0).Name = "FilterName": settings(0).Value = "Calc MS Excel 2007 XML"
    settings(1).Name = "Overwrite": settings(1).Value = True
    document.storeAsURL(ConvertToURL(filename), settings())
    document.close(True)
    Exit Sub
#End If
    priorAlerts = Application.DisplayAlerts
    Set book = Workbooks.Add(-4167)
    book.Worksheets(1).Name = "Data"
    book.Worksheets(1).Range("A1").Resize(count + 1, 3).Value = data
    book.Worksheets(1).Rows(1).Font.Bold = True
    book.Worksheets(1).Columns("A:C").ColumnWidth = 24
    Application.DisplayAlerts = False: changedAlerts = True
    book.SaveAs Filename:=filename, FileFormat:=51
    book.Close False
    Application.DisplayAlerts = priorAlerts
    Exit Sub
FailedSave:
    message = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not document Is Nothing Then document.close(True)
    If changedAlerts Then Application.DisplayAlerts = priorAlerts
    On Error GoTo 0
    Err.Raise 5, , message
End Sub

Public Function RunSplit() As String
    On Error GoTo Failed
    Dim src As Object, index As Object, groups As New Collection
    Dim g As String, name As String, folder As String, sep As String
    Dim r As Long, j As Long, k As Long, n As Long, count As Long, found As Boolean
    Dim hr As Long, id As Long, gc As Long, vc As Long, data As Variant
    Set src = ThisWorkbook.Worksheets("Data"): CheckData src
    hr = HeaderRow(src, "record_id|group|value")
    id = Col(src, hr, "record_id"): gc = Col(src, hr, "group"): vc = Col(src, hr, "value")
    If ThisWorkbook.Path = "" Then Err.Raise 5, , "Save the macro workbook before splitting"
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "output"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    folder = folder & sep & "split"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    ' Only remove files recorded by the previous split run.
    If Exists("__MP_Split") Then
        Set index = ThisWorkbook.Worksheets("__MP_Split")
        For r = 2 To LastRow(index)
            name = CStr(index.Cells(r, 2).Value)
            If SafeGroupFile(name) Then
                If Dir(folder & sep & name) <> "" Then Kill folder & sep & name
            End If
        Next r
    End If
    For r = hr + 1 To LastRow(src)
        If Not IsBlankRow(src, r) Then
            g = CStr(src.Cells(r, gc).Value): found = False
            For j = 1 To groups.Count
                If CStr(groups(j)) = g Then found = True: Exit For
            Next j
            If Not found Then groups.Add g
        End If
    Next r
    Set index = OutSheet("__MP_Split"): Head index, "group|file|rows"
    For j = 1 To groups.Count
        g = CStr(groups(j)): count = 0
        For r = hr + 1 To LastRow(src)
            If Not IsBlankRow(src, r) And CStr(src.Cells(r, gc).Value) = g Then count = count + 1
        Next r
        ReDim data(1 To count + 1, 1 To 3)
        data(1, 1) = "record_id": data(1, 2) = "group": data(1, 3) = "value": k = 1
        For r = hr + 1 To LastRow(src)
            If Not IsBlankRow(src, r) And CStr(src.Cells(r, gc).Value) = g Then
                k = k + 1
                data(k, 1) = src.Cells(r, id).Value
                data(k, 2) = src.Cells(r, gc).Value
                data(k, 3) = src.Cells(r, vc).Value
            End If
        Next r
        name = "group-" & Format(j, "000") & ".xlsx"
        SaveGroup folder & sep & name, data, count
        index.Cells(j + 1, 1).Value = g
        index.Cells(j + 1, 2).Value = name
        index.Cells(j + 1, 3).Value = count
        n = n + count
    Next j
    RunSplit = "OK split " & n: Exit Function
Failed:
    RunSplit = "ERROR " & Err.Description
End Function
