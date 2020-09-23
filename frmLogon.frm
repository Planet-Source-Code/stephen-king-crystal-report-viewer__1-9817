VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon/Logoff Server"
   ClientHeight    =   2235
   ClientLeft      =   5175
   ClientTop       =   3570
   ClientWidth     =   3780
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsetup 
      Caption         =   "&Setup"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame frasetup 
      Caption         =   "Setup"
      Height          =   1695
      Left            =   3840
      TabIndex        =   7
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbserver 
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDatabaseName 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image imgsetup 
         Height          =   480
         Left            =   120
         Picture         =   "frmLogon.frx":0442
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDataBaseName 
         Caption         =   "Database Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblServerName 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fralogon 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      Begin VB.CheckBox chksave 
         Caption         =   "Save LogOn Info"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblUserID 
         Caption         =   "User ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdLogOn 
      Caption         =   "&LogOn"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim attempts As Integer     'counts the number of logon attempts

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdLogOn_Click()

On Error Resume Next
    
    'decide, according to command button caption, wheter to logon/logoff..
    
    If cmdLogOn.Caption = "&LogOn" Then
        Crystal_Report_Application.LogOnServer Trim(DLLName), Trim(cmbserver.Text), _
                    Trim(txtDatabaseName), Trim(txtUserID), _
                    Trim(txtPassword)
        If Err.Number <> 0 Then
            MsgBox "Logon Failed!" & vbCrLf & "Please try again", vbCritical + vbOKOnly + vbApplicationModal, "Logon Failed"
            cmdLogOn.Caption = "&LogOn"
            attempts = attempts + 1
        Else
            cmdLogOn.Caption = "&LogOff"
            fralogon.Enabled = False
            write_logon_info
            frmMain.Show
            Me.Hide
        End If
    ElseIf cmdLogOn.Caption = "&LogOff" Then
       On Error Resume Next
            Crystal_Report_Application.LogOffServer Trim(DLLName), Trim(cmbserver.Text), _
                    Trim(txtDatabaseName), Trim(txtUserID), _
                    Trim(txtPassword)
            If Err.Number <> 0 Then
                MsgBox "Logoff Failed!"
            Else
                cmdLogOn.Caption = "&LogOn"
                frmLogon.Enabled = True
                Me.Hide
            End If
    End If
    If attempts > 3 Then
        MsgBox "You have tried 3 times to logon" & vbCrLf & "I am exiting the program", vbCritical + vbInformation + vbOKOnly, "Stop Trying"
        End
    End If
    
End Sub

Private Sub cmdOK_Click()
    Me.Width = 3900
    center_form Me
End Sub

Private Sub cmdsetup_Click()
    Me.Width = 7965
    center_form Me
End Sub

Private Sub Form_Load()
    read_logon_info
End Sub

Private Sub txtUserID_GotFocus()
    txtUserID.SelStart = 4
End Sub
Private Sub write_logon_info()
   
   'save the info to file..
   
    WriteINI "Logon Info", "Server", cmbserver.Text
    WriteINI "Logon Info", "Database", txtDatabaseName
    WriteINI "Logon Info", "UserID", txtUserID
    WriteINI "Logon Info", "Password", txtPassword
    
    If chksave.Value = 1 Then
        WriteINI "Logon Info", "Auto", "YES"
    Else
        WriteINI "Logon Info", "Auto", "NO"
    End If
    
End Sub
Private Sub read_logon_info()
    'read from file
    
Dim check_for_auto As String    'see if the user wants to logon automatically

    cmbserver.Text = ReadINI("Logon Info", "Server")
    txtDatabaseName = ReadINI("Logon Info", "Database")
    check_for_auto = ReadINI("Logon Info", "Auto")
    If check_for_auto = "YES" Then
        txtUserID = ReadINI("Logon Info", "UserID")
        txtPassword = ReadINI("Logon Info", "Password")
        chksave.Value = 1
        cmdLogOn_Click
    Else
        txtUserID = ""
        txtPassword = ""
        chksave.Value = 0
    End If
    
End Sub
