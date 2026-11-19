Attribute VB_Name = "RenameSheetsModule"
Option Explicit
Public Function GoodName(ByVal s As String) As Boolean
    Dim c As Long, ch As String
    If Len(s) = 0 Or Len(s) > 31 Then Exit Function
    If Left(s, 1) = "'" Or Right(s, 1) = "'" Or LCase(s) = "history" Or UCase(Left(s, 5)) = "__MP_" Then Exit Function
    For c = 1 To Len(s)
        ch = Mid(s, c, 1)
        If InStr("[]:*?/\", ch) > 0 Or Asc(ch) < 32 Then Exit Function
    Next c
    GoodName = True
End Function
Public Sub RenameSheets()
    Dim result As String
    result = RunRename()
    If Left(result, 5) = "ERROR" Then Err.Raise 5, , result
End Sub
Public Function RunRename() As String
    On Error GoTo Failed
    Dim src As Object, dst As Object, hr As Long, oc As Long, nc As Long, r As Long, n As Long
    Dim old As String, newname As String, reason As String, seen As New Collection, item As Variant, repeated As Boolean
    Set src = ThisWorkbook.Worksheets("RenameMap")
    hr = HeaderRow(src, "old|new"): oc = Col(src, hr, "old"): nc = Col(src, hr, "new")
    Set dst = OutSheet("__MP_Rename"): Head dst, "old|new|status": n = 1
    For r = hr + 1 To LastRow(src)
        If Not IsBlankRow(src, r) Then
            old = CStr(src.Cells(r, oc).Value): newname = CStr(src.Cells(r, nc).Value): repeated = False
            For Each item In seen
                If LCase(CStr(item)) = LCase(old) Then repeated = True
            Next item
            If repeated Then
                reason = "rejected: repeated source"
            ElseIf InStr("|data|mergea|mergeb|renamemap|recipients|instructions|template|config|", "|" & LCase(old) & "|") > 0 Or UCase(Left(old, 5)) = "__MP_" Then
                reason = "rejected: protected sheet"
            ElseIf Not Exists(old) Then
                reason = "rejected: missing source"
            ElseIf Not GoodName(newname) Then
                reason = "rejected: invalid name"
            ElseIf Exists(newname) And LCase(old) <> LCase(newname) Then
                reason = "rejected: destination exists"
            Else
                ThisWorkbook.Worksheets(old).Name = newname
                reason = "renamed"
            End If
            seen.Add old: n = n + 1
            dst.Cells(n, 1).Value = old: dst.Cells(n, 2).NumberFormat = "@": dst.Cells(n, 2).Value = "'" & newname: dst.Cells(n, 3).Value = reason
        End If
    Next r
    RunRename = "OK rename " & n - 1: Exit Function
Failed:
    RunRename = "ERROR " & Err.Description
End Function
