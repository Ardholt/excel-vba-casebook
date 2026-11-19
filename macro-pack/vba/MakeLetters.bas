Attribute VB_Name = "MakeLettersModule"
Option Explicit
Public Function TokenLetter(ByVal t As String, ByVal id As String, ByVal name As String, ByVal amt As String) As String
    Dim pos As Long, closing As Long, key As String, result As String
    Do
        pos = InStr(t, "{{")
        If pos = 0 Then
            If InStr(t, "}}") > 0 Then Err.Raise 5, , "Unmatched template token"
            result = result & t: Exit Do
        End If
        result = result & Left(t, pos - 1): closing = InStr(pos + 2, t, "}}")
        If closing = 0 Then Err.Raise 5, , "Unclosed template token"
        key = Mid(t, pos + 2, closing - pos - 2)
        Select Case key
            Case "recipient_id": result = result & id
            Case "display_name": result = result & name
            Case "amount": result = result & amt
            Case Else: Err.Raise 5, , "Unknown template token"
        End Select
        t = Mid(t, closing + 2)
    Loop
    TokenLetter = result
End Function
Public Function SafeRecipientID(ByVal id As String) As Boolean
    Dim c As Long
    If Len(id) = 0 Or Len(id) > 64 Then Exit Function
    If InStr("|CON|PRN|AUX|NUL|COM1|COM2|COM3|COM4|COM5|COM6|COM7|COM8|COM9|LPT1|LPT2|LPT3|LPT4|LPT5|LPT6|LPT7|LPT8|LPT9|", "|" & UCase(id) & "|") > 0 Then Exit Function
    For c = 1 To Len(id)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-", Mid(id, c, 1)) = 0 Then Exit Function
    Next c
    SafeRecipientID = True
End Function
Public Sub MakeLetters()
    Dim result As String
    result = RunLetters()
    If Left(result, 5) = "ERROR" Then Err.Raise 5, , result
End Sub
Public Function RunLetters() As String
    On Error GoTo Failed
    Dim src As Object, dst As Object, hr As Long, ic As Long, nc As Long, ac As Long, r As Long, n As Long, c As Long
    Dim id As String, name As String, amount As String, template As String, body As String, folder As String, sep As String
    Dim f As Integer, opened As Boolean, ids As New Collection, item As Variant, ch As String
    Set src = ThisWorkbook.Worksheets("Recipients")
    hr = HeaderRow(src, "recipient_id|display_name|amount")
    ic = Col(src, hr, "recipient_id"): nc = Col(src, hr, "display_name"): ac = Col(src, hr, "amount")
    template = CStr(ThisWorkbook.Worksheets("Template").Cells(1, 1).Value)
    body = TokenLetter(template, "", "", "")
    For c = 1 To Len(template)
        If AscW(Mid(template, c, 1)) < 0 Or AscW(Mid(template, c, 1)) > 127 Then Err.Raise 5, , "ASCII template required"
    Next c
    For r = hr + 1 To LastRow(src)
        If Not IsBlankRow(src, r) Then
            id = CStr(src.Cells(r, ic).Value): name = CStr(src.Cells(r, nc).Value)
            If Not SafeRecipientID(id) Or name = "" Or Not IsNumeric(src.Cells(r, ac).Value) Then Err.Raise 5, , "Invalid recipient"
            If Abs(CDbl(src.Cells(r, ac).Value) * 100 - Round(CDbl(src.Cells(r, ac).Value) * 100, 0)) > 0.000001 Then Err.Raise 5, , "Amount needs at most two decimals"
            For c = 1 To Len(id)
                ch = Mid(id, c, 1)
                If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-", ch) = 0 Then Err.Raise 5, , "Unsafe recipient ID"
            Next c
            For c = 1 To Len(name)
                If AscW(Mid(name, c, 1)) < 0 Or AscW(Mid(name, c, 1)) > 127 Then Err.Raise 5, , "ASCII recipient name required"
            Next c
            For Each item In ids
                If LCase(CStr(item)) = LCase(id) Then Err.Raise 5, , "Duplicate recipient ID"
            Next item
            ids.Add id
        End If
    Next r
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "output"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    folder = folder & sep & "letters"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    If Exists("__MP_Letters") Then
        Set dst = ThisWorkbook.Worksheets("__MP_Letters")
        For r = 2 To LastRow(dst)
            id = CStr(dst.Cells(r, 1).Value)
            If SafeRecipientID(id) Then
                If Dir(folder & sep & id & ".txt") <> "" Then Kill folder & sep & id & ".txt"
            End If
        Next r
    End If
    Set dst = OutSheet("__MP_Letters"): Head dst, "recipient_id|amount|letter": n = 1
    For r = hr + 1 To LastRow(src)
        If Not IsBlankRow(src, r) Then
            id = CStr(src.Cells(r, ic).Value): name = CStr(src.Cells(r, nc).Value)
            amount = Format(src.Cells(r, ac).Value, "0.00")
            amount = Replace(amount, ",", ".")
            body = TokenLetter(template, id, name, amount)
            f = FreeFile: Open folder & sep & id & ".txt" For Output As #f: opened = True
            Print #f, body
            Close #f: opened = False
            n = n + 1: dst.Cells(n, 1).Value = id: dst.Cells(n, 2).Value = amount: dst.Cells(n, 3).Value = body
        End If
    Next r
    RunLetters = "OK letters " & n - 1: Exit Function
Failed:
    If opened Then Close #f
    RunLetters = "ERROR " & Err.Description
End Function
