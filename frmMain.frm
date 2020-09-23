VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voice Service Center: Crystal Reports Viewer"
   ClientHeight    =   2205
   ClientLeft      =   2895
   ClientTop       =   2850
   ClientWidth     =   5940
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdchangeparameter 
      Caption         =   "Change Selection/Parameter Values"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdpreview2 
      Caption         =   "Preview Report in &Separate Window"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   18
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame frareports 
      Height          =   2175
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdremove 
         Caption         =   "&Remove Report"
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add Report"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ListBox lstreports 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1425
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image imgreport 
         Height          =   960
         Left            =   2520
         Picture         =   "frmMain.frx":030A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblreports 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Available Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdpreview 
      Caption         =   "Pre&view Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close Preview"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Frame frareportstatus 
      Caption         =   "Report Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtRecordsPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox txtRecordsRead 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox txtRecordsSelected 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox txtNumberOfPages 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblPSNumOfRecordsPrinted 
         Caption         =   "Number of Records Printed:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblPSNumOfRecordsRead 
         Caption         =   "Number of Records Read:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblPSNumOfRecordsSelected 
         Caption         =   "Number of Records Selected:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblpages 
         Caption         =   "Number of Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frareport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   9375
      Begin CRVIEWERLibCtl.CRViewer CRVReport 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   9135
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   0   'False
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   0   'False
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   0   'False
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   5880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Reports"
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Pre&view..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPreviewApplication 
         Caption         =   "Preview (&Separate Window)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSavedData 
         Caption         =   "Save Data &with Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Sa&ve As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Begin VB.Menu mnuFilePrintExport 
            Caption         =   "Expor&t..."
         End
         Begin VB.Menu mnuFilePrintPrinter 
            Caption         =   "Pri&nter"
         End
      End
      Begin VB.Menu mnuVerifyOnEveryPrint 
         Caption         =   "&Verify On Every Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogonServer 
         Caption         =   "&Logon\Logoff Server"
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reports() As String 'array of reports -- for user loaded..

Private Sub cmdadd_Click()
  
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  
  CommonDialog1.DialogTitle = "Add Report to List"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.InitDir = App.Path & "\Reports\"
  CommonDialog1.Filter = "Crystal Report Files (*.rpt)|*.rpt|All Files" & "(*.*)|*.*"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  add_reports
  save_report_data
  DoEvents
    
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub cmdchangeparameter_Click()
    CRVReport.Refresh       'change the selection formula(s)
    set_report_status
End Sub

Private Sub cmdclose_Click()
    Me.Height = 2805
    Me.Width = 6030
    cmdpreview.Enabled = True
    center_form Me
    
End Sub

Private Sub cmdend_Click()
    DoEvents
    Unload Me
    End
End Sub

Private Sub show_report()
    
    Screen.MousePointer = vbArrowHourglass
    Set The_Crystal_Report = Nothing
    Set The_Crystal_Report = Crystal_Report_Application.OpenReport(reports(lstreports.ListIndex), 1)  'Opening report, non-exclusively (shared)
    mnuFileSaveAs.Visible = True
    'Check to see if there is saved data with report
    mnuFileSavedData.Checked = The_Crystal_Report.HasSavedData
    Screen.MousePointer = vbDefault
       
End Sub
Private Sub cmdpreview_Click()
    
    If Dir(reports(lstreports.ListIndex)) <> "" Then 'check to make sure the report is a valid/existing file
        cmdpreview.Enabled = False
        Me.Enabled = False
    '  On Error GoTo error_check
        show_report
        previewinframe
    Else
         MsgBox "Problems with finding that report" & vbCrLf & _
                "Report has been removed from list" & vbCrLf & _
                "Please add the report again ", vbCritical + vbApplicationModal + vbOKOnly, "Problem Loading Report"
        cmdremove_Click
        DoEvents
        cmdadd_Click
        Screen.MousePointer = vbDefault
        Me.Refresh
        Exit Sub
    End If
    
'error_check:
'       If Err.Number = 440 Then
'        MsgBox "Problems with finding that report" & vbCrLf & _
'                "Report has been removed from list" & vbCrLf & _
'                "Please add the report again ", vbCritical + vbApplicationModal + vbOKOnly, "Problem Loading Report"
'        cmdremove_Click
'        DoEvents
'        cmdadd_Click
'        Screen.MousePointer = vbDefault
'        Me.Refresh
'        Exit Sub
'     Else
'        Resume
'    End If
    
End Sub

Private Sub cmdpreview2_Click()
    Screen.MousePointer = vbArrowHourglass
    show_report
    Load frmPreview
    frmPreview.Caption = "Report :" & lstreports.List(lstreports.ListIndex) & ".rpt"
    frmPreview.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdprint_Click()
    CRVReport.PrintReport

End Sub

Private Sub cmdremove_Click()
Dim steve As Integer

        lstreports.RemoveItem lstreports.ListIndex
        lstreports.Refresh
        ReDim reports(lstreports.ListCount + 1)
        Do Until steve = lstreports.ListCount + 1
            reports(steve) = lstreports.List(steve)
            steve = steve + 1
        Loop
        
        save_report_data
        check_reports
           
End Sub

Private Sub Form_Load()
Dim temp_report_number As String
    ReDim reports(1)
    temp_report_number = ReadINI("Reports", "Number")
    If temp_report_number <> "" Then
        temp_report_number = CInt(temp_report_number)
        If temp_report_number > 0 Then
            load_report_data
        End If
    End If
    check_reports
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Clean up
Set The_Crystal_Report = Nothing

End

End Sub

Private Sub lstreports_Click()
    check_reports
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    cmdclose_Click
End Sub

Private Sub mnuFileLogonServer_Click()
    frmLogon.Show vbModal
End Sub

Private Sub mnuFileOpenSubreport_Click()
    frmOpenSubreport.Show vbModal
End Sub

Private Sub mnuFilePrintExport_Click()
    frmExport.Show vbModal
End Sub

Private Sub mnuFilePrintPrinter_Click()
    frmPrintOut.Show vbModal
End Sub

Private Sub mnuFileSaveAs_Click()

With CommonDialog1
    .DefaultExt = "*.rpt"
    .DialogTitle = "Save As..."
    .FileName = "*.rpt"
    .Filter = "Crystal Reports (*.rpt)"
    .InitDir = App.Path & "\Reports\"
    .ShowSave
End With

The_Crystal_Report.Save CommonDialog1.FileName

End Sub

Private Sub mnuFileSavedData_Click()

    If mnuFileSavedData.Checked Then
        mnuFileSavedData.Checked = False
    Else
        mnuFileSavedData.Checked = True
    End If

End Sub

Private Sub mnuFileSummaryInfo_Click()
    frmSummaryInfo.Show vbModal
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuLineObject_Click()
    frmLineObject.Show vbModal
End Sub
Private Sub set_report_status()
    
    Set CrystalPrintingStatus = The_Crystal_Report.PrintingStatus
    'give some report status
    With CrystalPrintingStatus
       txtNumberOfPages = .NumberOfPages
       txtRecordsPrinted = .NumberOfRecordPrinted
       txtRecordsRead = .NumberOfRecordRead
       txtRecordsSelected = .NumberOfRecordSelected
    End With
    
End Sub
Private Sub previewinframe()
    
    Me.Height = 8100
    Me.Width = 9540
    Screen.MousePointer = vbArrowHourglass
    center_form Me
    frareportstatus.Visible = True
    
    'Pass the report to the viewer to display it
    CRVReport.ReportSource = The_Crystal_Report
    frareport.Caption = "Report : " & lstreports.List(lstreports.ListIndex) & ".rpt"
    set_report_status
    
    'Preview the report
    CRVReport.ViewReport
    CRVReport.Visible = True
    CRVReport.Width = frareport.Width - 300
    CRVReport.Height = frareport.Height - 500
    Screen.MousePointer = vbDefault
    Me.Enabled = True 'allow form functions once the report is shown
    
End Sub

Private Sub mnuPrinterSetup_Click()
    frmPrinterSetup.Show vbModal
End Sub

Private Sub mnuReportInfo_Click()
    frmReportInfo.Show vbModal
End Sub

Private Sub mnuPreview_Click()
    cmdpreview_Click
End Sub

Private Sub mnuPreviewApplication_Click()
    cmdpreview2_Click
End Sub

Private Sub mnuVerifyOnEveryPrint_Click()
    
    If mnuVerifyOnEveryPrint.Checked = True Then
        mnuVerifyOnEveryPrint.Checked = False
    Else
        mnuVerifyOnEveryPrint.Checked = True
    End If
    The_Crystal_Report.VerifyOnEveryPrint = mnuVerifyOnEveryPrint

End Sub

Public Sub ShowMenus(the_value As Boolean)
    
    mnuPreview.Enabled = the_value
    mnuPreviewApplication.Enabled = the_value
    mnuFileSavedData.Enabled = the_value
    mnuFileSaveAs.Enabled = the_value
    mnuPrint.Enabled = the_value
    mnuVerifyOnEveryPrint.Enabled = the_value
   
End Sub
Private Sub check_reports()
    
    lstreports.Refresh
    
    If lstreports.ListCount <= 0 Then
        cmdremove.Enabled = False
        cmdremove.Enabled = False
        cmdpreview.Enabled = False
        cmdremove.Enabled = False
        cmdpreview2.Enabled = False
        Me.Caption = "Voice Service Center : Crystal Reports Viewer"
        ShowMenus False
    Else    'if there are items in the list, then have to check for if one is selected
        If lstreports.ListIndex >= 0 Then
            cmdremove.Enabled = True
            cmdpreview.Enabled = True
            cmdremove.Enabled = True
            cmdpreview2.Enabled = True
            ShowMenus True
            Me.Caption = "Voice Service Center : " & lstreports.List(lstreports.ListIndex)
        Else
            cmdremove.Enabled = False
            cmdpreview.Enabled = False
            cmdremove.Enabled = False
            cmdpreview2.Enabled = False
            Me.Caption = "Voice Service Center : Crystal Reports Viewer"
            ShowMenus False
        End If
    End If
    
End Sub
Private Sub add_reports()

Dim steve As Integer
    
    Do While reports(steve) <> ""
        If reports(steve) = CommonDialog1.FileName Then
            MsgBox "Report " & CommonDialog1.FileName & " already exists in the list", vbOKOnly + vbApplicationModal + vbInformation, "Report Already Exists"
            Exit Sub
        End If
        steve = steve + 1
    Loop
         
    ReDim Preserve reports(lstreports.ListCount + 1)
    reports(lstreports.ListCount) = CommonDialog1.FileName
    lstreports.AddItem Mid(CommonDialog1.FileTitle, 1, Len(CommonDialog1.FileTitle) - 4), lstreports.ListCount
    lstreports.Refresh
    check_reports
End Sub
Private Sub save_report_data()

Dim steve As Integer    'loop counter
    
    WriteINI "Reports", "Number", lstreports.ListCount
    Do Until steve = lstreports.ListCount
        WriteINI "Report #" & steve + 1, "Name", lstreports.List(steve)
        WriteINI "Report #" & steve + 1, "Location", reports(steve)
        steve = steve + 1
    Loop
 
End Sub
Private Sub load_report_data()

Dim steve As Integer
Dim jack
    
    lstreports.Clear
    jack = ReadINI("Reports", "Number")
    steve = 1
    ReDim reports(jack)
    Do Until steve = jack + 1
        lstreports.AddItem ReadINI("Report #" & steve, "Name")
        reports(steve - 1) = ReadINI("Report #" & steve, "Location")
        steve = steve + 1
    Loop
        
End Sub

