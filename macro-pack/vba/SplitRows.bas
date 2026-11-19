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

Public Function RunSplit() As String
    On Error GoTo Failed
    Dim src As Object, dst As Object, index As Object, ws As Object
    Dim groups As New Collection, g As String, item As Variant, n As Long, r As Long, j As Long
    Dim hr As Long, id As Long, gc As Long, vc As Long, found As Long, nextrow As Long
    Set src = ThisWorkbook.Worksheets("Data"): CheckData src
    hr = HeaderRow(src, "record_id|group|value")
    id = Col(src, hr, "record_id"): gc = Col(src, hr, "group"): vc = Col(src, hr, "value")
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "__MP_G" Then ws.Cells.Clear
    Next ws
    Set index = OutSheet("__MP_Split"): Head index, "group|sheet|rows"
    For r = hr + 1 To LastRow(src)
        If Not IsBlankRow(src, r) Then
            g = CStr(src.Cells(r, gc).Value): found = 0
            For j = 1 To groups.Count
                If CStr(groups(j)) = g Then found = j: Exit For
            Next j
            If found = 0 Then
                groups.Add g: found = groups.Count
                Set dst = OutSheet("__MP_G" & Format(found, "000")): Head dst, "record_id|group|value"
                index.Cells(found + 1, 1).Value = g
                index.Cells(found + 1, 2).Value = dst.Name
                index.Cells(found + 1, 3).Value = 0
            Else
                Set dst = ThisWorkbook.Worksheets("__MP_G" & Format(found, "000"))
            End If
            nextrow = CLng(index.Cells(found + 1, 3).Value) + 2
            dst.Cells(nextrow, 1).Value = src.Cells(r, id).Value
            dst.Cells(nextrow, 2).Value = src.Cells(r, gc).Value
            dst.Cells(nextrow, 3).Value = src.Cells(r, vc).Value
            index.Cells(found + 1, 3).Value = nextrow - 1
            n = n + 1
        End If
    Next r
    RunSplit = "OK split " & n: Exit Function
Failed:
    RunSplit = "ERROR " & Err.Description
End Function
