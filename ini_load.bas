Attribute VB_Name = "ini"
#If Win32 Then


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long


Private Declare Function GetDesktopWindow Lib "user32" () As Long
#Else


Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
    ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
    ByVal lpszDir$, ByVal fsShowCmd%) As Integer


Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If


Public Const SW_SHOWNORMAL = 1

'functions use to read and write to an ini file..as taken from the web

Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
    Public Const PROCESS_QUERY_INFORMATION = &H400

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
        ByVal lpFileName As String) As Long


Public Function ReadINI(strKey As String, strName As String) As String
    Dim intLen As Integer
    Dim strText As String, strIniFile As String
    Dim response
    Dim edit_password
    
    strIniFile = App.Path & "\" & App.EXEName & ".dat" 'sets where the ini file is
    strText = Space(256)
    intLen = GetPrivateProfileString(strKey, strName, "", strText, Len(strText), strIniFile)
    If intLen > -1 Then
        strText = Left(strText, intLen)
    End If
    ReadINI = strText
End Function
Public Sub WriteINI(strKey As String, strName As String, strText As String)
    Dim intLen As Integer
    Dim strIniFile As String
    strIniFile = App.Path & "\" & App.EXEName & ".dat"
    intLen = WritePrivateProfileString(strKey, strName, strText, strIniFile)

End Sub

