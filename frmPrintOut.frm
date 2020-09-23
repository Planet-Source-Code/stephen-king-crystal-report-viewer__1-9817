VERSION 5.00
Begin VB.Form frmPrintOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Report"
   ClientHeight    =   1980
   ClientLeft      =   5400
   ClientTop       =   3690
   ClientWidth     =   4545
   Icon            =   "frmPrintOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtStopPageNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Text            =   "1"
      Top             =   960
      Width           =   500
   End
   Begin VB.TextBox txtStartPageNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "1"
      Top             =   600
      Width           =   500
   End
   Begin VB.CheckBox chkCollated 
      Caption         =   "Collated"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtNumberOfCopy 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "1"
      Top             =   240
      Width           =   500
   End
   Begin VB.CheckBox chkPromptUser 
      Caption         =   "Prompt User"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintReportWithParameters 
      Caption         =   "Print Report"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintReportWithoutParameters 
      Caption         =   "Print Report"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblStopPageNumber 
      Caption         =   "Stop Page Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Start Page Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblNumberOfCopies 
      Caption         =   "Number Of Copies:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblWithoutParameters 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "With out Parameters:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblWithParameters 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "With Parameters:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrintReportWithoutParameters_Click()
     The_Crystal_Report.PrintOut
End Sub

Private Sub cmdPrintReportWithParameters_Click()
The_Crystal_Report.PrintOut False, CInt(txtNumberOfCopy.Text), CBool(chkCollated), CInt(txtStartPageNumber.Text), CInt(txtStopPageNumber.Text)
Unload Me

End Sub

Private Sub Form_Load()

'Hiding unneccessary information
Me.Height = 2460

End Sub

