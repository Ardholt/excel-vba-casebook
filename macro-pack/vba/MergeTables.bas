Attribute VB_Name = "MergeTablesModule"
Option Explicit
Public Sub MergeTables()
    Dim result As String
    result = RunMerge()
    If Left(result, 5) = "ERROR" Then Err.Raise 5, , result
End Sub
Public Function RunMerge() As String
    On Error GoTo Failed
    Dim src As Object, dst As Object, name As Variant, r As Long, hr As Long, n As Long
    Dim id As Long, g As Long, v As Long
    CheckData ThisWorkbook.Worksheets("MergeA")
    CheckData ThisWorkbook.Worksheets("MergeB")
    Set dst = OutSheet("__MP_Merged"): Head dst, "record_id|group|value": n = 1
    For Each name In Array("MergeA", "MergeB")
        Set src = ThisWorkbook.Worksheets(CStr(name))
        hr = HeaderRow(src, "record_id|group|value")
        id = Col(src, hr, "record_id"): g = Col(src, hr, "group"): v = Col(src, hr, "value")
        For r = hr + 1 To LastRow(src)
            If Not IsBlankRow(src, r) Then
                n = n + 1
                dst.Cells(n, 1).Value = src.Cells(r, id).Value
                dst.Cells(n, 2).Value = src.Cells(r, g).Value
                dst.Cells(n, 3).Value = src.Cells(r, v).Value
            End If
        Next r
    Next name
    RunMerge = "OK merge " & n - 1: Exit Function
Failed:
    RunMerge = "ERROR " & Err.Description
End Function
