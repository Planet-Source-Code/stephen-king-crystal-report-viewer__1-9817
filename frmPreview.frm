VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close Report"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   6960
      Width           =   2295
   End
   Begin CRVIEWERLibCtl.CRViewer CRVReport 
      Height          =   6720
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9000
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    CRVReport.ReportSource = The_Crystal_Report
    CRVReport.ViewReport
    Screen.MousePointer = vbDefault
   
    
End Sub

Private Sub Form_Resize()
  
    CRVReport.Top = 20
    CRVReport.Left = 20
    CRVReport.Height = Me.Height - 1000
    CRVReport.Width = Me.Width - 200
    cmdclose.Top = CRVReport.Top + CRVReport.Height + 100
  
    
End Sub
