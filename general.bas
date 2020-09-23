Attribute VB_Name = "general"

Option Explicit
Public Crystal_Report_Application As New CRAXDRT.Application
Public The_Crystal_Report As CRAXDRT.Report
Public Const DLLName As String = "pdsodbc.dll"
'

Public Sub center_form(the_form As Form)
    the_form.Left = (Screen.Width / 2) - (the_form.Width / 2)
    the_form.Top = (Screen.Height / 2) - (the_form.Height / 2)
    the_form.Refresh
End Sub
Sub NumbersOnly(t As Control, KeyAscii As Integer)
    If KeyAscii < Asc(" ") Then     ' Is this Control char?
        Exit Sub                    ' Yes, let it pass
    End If
    CheckPeriod t                   ' Remove excess periods
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        ' keep digit
    ElseIf KeyAscii = Asc(".") Then
        ' keep .
    ElseIf KeyAscii = Asc("-") And t.SelStart = 0 Then
        ' Keep - only if first char
    Else
        KeyAscii = 0                ' Discard all other chars
    End If
    ' This code keeps you from typing any characters in front of
    ' a minus sign.
    '
    If Mid$(t.Text, t.SelStart + t.SelLength + 1, 1) = "-" Then
        KeyAscii = 0                ' Discard chars before -
    End If
End Sub

Public Function GetFilename(ByVal TempPath As String, ReturnType As Integer)

'function to get the filename and its parts

    Dim DriveLetter As String
    Dim DirPath As String
    Dim fname As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean


    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 And ReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If

    DriveLetter = ""
    DirPath = ""
    fname = ""
    Extension = ""


    If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
        DriveLetter = Left(TempPath, 2)
        TempPath = Mid(TempPath, 3)
    End If

    PathLength = Len(TempPath)


    For Offset = PathLength To 1 Step -1 ' Find the next delimiter.
        Select Case Mid(TempPath, Offset, 1)
        Case ".": ' This indicates either an extension or a . or a ..
        ThisLength = Len(TempPath) - Offset


        If ThisLength >= 1 Then ' Extension
            Extension = Mid(TempPath, Offset, ThisLength + 1)
        End If

        TempPath = Left(TempPath, Offset - 1)
        Case "\": ' This indicates a path delimiter.
        ThisLength = Len(TempPath) - Offset


        If ThisLength >= 1 Then ' Filename
            fname = Mid(TempPath, Offset + 1, ThisLength)
            TempPath = Left(TempPath, Offset)
            FileNameFound = True
            Exit For
        End If

        Case Else
        End Select
    Next Offset
    If FileNameFound = False Then
        fname = TempPath
    Else
        DirPath = TempPath
    End If
    If ReturnType = 0 Then
        GetFilename = DriveLetter
    ElseIf ReturnType = 1 Then
        GetFilename = DirPath
    ElseIf ReturnType = 2 Then
        GetFilename = fname
    ElseIf ReturnType = 3 Then
        GetFilename = Extension
    End If

End Function
Sub CheckPeriod(t As Control)
    Dim i As Integer
    
    i = InStr(1, t.Text, ".")   ' Look for a period
    If i > 0 And InStr(i + 1, t.Text, ".") > 0 Then
        t.SelStart = t.SelStart - 1
        t.SelLength = 1         ' Select new period
        t.SelText = ""          ' Remove new period
    End If
End Sub

