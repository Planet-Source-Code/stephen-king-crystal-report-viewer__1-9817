VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   Caption         =   "Export Options..."
   ClientHeight    =   6360
   ClientLeft      =   3840
   ClientTop       =   2055
   ClientWidth     =   6435
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset..."
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   5880
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmExport.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDestFileName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHTMLFilename"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblNumOfLinesPerPage"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDestPath"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCharDelimiters"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDestPath"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdHTMLFilePath"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtHTMLFilePath"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNumOfLinesPerPage"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkUseRptDateFormat"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkUseRptNumFormat"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraExportTypeDest"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Exchange Options"
      TabPicture(1)   =   "frmExport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblExchangeType"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblExchangeFolderPath"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPassword"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblXProfile"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbDestType"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtXFolderPath"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtXPassword"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtXProfile"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkXTabHasColumnHeadings"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Mail Options"
      TabPicture(2)   =   "frmExport.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblBCCList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblCCList"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblSubject"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblToList"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblMessage"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtBCCList"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtCCList"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtSubject"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtToList"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtMessage"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "ODBC Options"
      TabPicture(3)   =   "frmExport.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblDSN"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblDSNPassword"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblDSNUserID"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblExportTableName"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtDSN"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtDSNPassword"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtDSNUserID"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtTableName"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.Frame fraExportTypeDest 
         Caption         =   "Exporting Format and Destination"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   47
         Top             =   600
         Width           =   5895
         Begin VB.ComboBox cmbDestination 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   480
            Width           =   2655
         End
         Begin VB.ComboBox cmbFormat 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblFormat 
            Caption         =   "Format:"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblDestination 
            Caption         =   "Destination:"
            Height          =   255
            Left            =   3000
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblFormatDLL 
            Caption         =   "DLL:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblDestDLL 
            Caption         =   "DLL:"
            Height          =   255
            Left            =   3000
            TabIndex        =   50
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.CheckBox chkUseRptNumFormat 
         Caption         =   "Use Report Number Format"
         Height          =   255
         Left            =   -72600
         TabIndex        =   45
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CheckBox chkUseRptDateFormat 
         Caption         =   "Use Report Date Format"
         Height          =   255
         Left            =   -72600
         TabIndex        =   44
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtTableName 
         Height          =   285
         Left            =   1800
         TabIndex        =   43
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtDSNUserID 
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtDSNPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtDSN 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtNumOfLinesPerPage 
         Height          =   285
         Left            =   -72600
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtMessage 
         Height          =   1095
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2640
         Width           =   4455
      End
      Begin VB.TextBox txtToList 
         Height          =   285
         Left            =   -73800
         TabIndex        =   31
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   -73800
         TabIndex        =   29
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtCCList 
         Height          =   285
         Left            =   -73800
         TabIndex        =   27
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtBCCList 
         Height          =   285
         Left            =   -73800
         TabIndex        =   25
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtHTMLFilePath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74880
         TabIndex        =   23
         Top             =   4320
         Width           =   5535
      End
      Begin VB.CommandButton cmdHTMLFilePath 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69360
         TabIndex        =   22
         Top             =   4320
         Width           =   375
      End
      Begin VB.CheckBox chkXTabHasColumnHeadings 
         Caption         =   "Tab Has Column Headings"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtXProfile 
         Height          =   285
         Left            =   -73560
         TabIndex        =   19
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtXPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73560
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtXFolderPath 
         Height          =   285
         Left            =   -73560
         TabIndex        =   15
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbDestType 
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdDestPath 
         Caption         =   "..."
         Height          =   255
         Left            =   -69360
         TabIndex        =   11
         Top             =   3600
         Width           =   375
      End
      Begin VB.Frame fraCharDelimiters 
         Caption         =   "Character Delimiters"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74880
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
         Begin VB.TextBox txtFieldDel 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtStringDel 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblFieldDel 
            Caption         =   "Field Delimiter:"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblStringDel 
            Caption         =   "String Delimiter:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.TextBox txtDestPath 
         Height          =   285
         Left            =   -74880
         TabIndex        =   4
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Label lblExportTableName 
         Caption         =   "Export Table Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblDSNUserID 
         Caption         =   "User ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblDSNPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblDSN 
         Caption         =   "Datasource Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblNumOfLinesPerPage 
         Caption         =   "Number of Lines Per Page:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   34
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblMessage 
         Caption         =   "Message:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblToList 
         Caption         =   "To List:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblSubject 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblCCList 
         Caption         =   "Cc List:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBCCList 
         Caption         =   "Bcc List:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblHTMLFilename 
         Caption         =   "HTML File Name:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblXProfile 
         Caption         =   "Profile:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblExchangeFolderPath 
         Caption         =   "Folder Path:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblExchangeType 
         Caption         =   "Destination Type:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDestFileName 
         Caption         =   "Destination File Path\Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   3360
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog cmdlg1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkPromptUser 
      Caption         =   "Prompt For Export Options"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code snippets from Crystal Reports Developers Help
Private Sub chkPromptUser_Click()

If chkPromptUser.Value = Checked Then
    SSTab1.Enabled = False
Else
    SSTab1.Enabled = True
End If

End Sub

Private Sub cmbDestination_Click()

'Read the corresponding Dll
'This only returns a dllName after a report is run
lblDestDLL.Caption = "DLL: " & CrystalExportOptions.DestinationDllName

'The following enables certain items only if they are required
If cmbDestination = "crEDTDiskFile" Then
    lblDestFileName.Enabled = True
    txtDestPath.Enabled = True
    cmdDestPath.Enabled = True
Else
    lblDestFileName.Enabled = False
    txtDestPath.Enabled = False
    cmdDestPath.Enabled = False
End If


If cmbFormat = "crEFTHTML32Standard" And cmbDestination = "crEDTDiskFile" Then
    lblHTMLFilename.Enabled = True
    txtHTMLFilePath.Enabled = True
    cmdHTMLFilePath.Enabled = True
Else
    lblHTMLFilename.Enabled = False
    txtHTMLFilePath.Enabled = False
    cmdHTMLFilePath.Enabled = False
End If

End Sub

Private Sub cmbFormat_Click()

'Read the corresponding Dll
'This only returns a dllName after a report is run
lblFormatDLL.Caption = "DLL: " & CrystalExportOptions.FormatDllName


'The following enables certain items only if they are required
If (cmbFormat = "crEFTCharSeparatedValues" Or _
    cmbFormat = "crEFTCommaSeparatedValues") Then
    
    fraCharDelimiters.Enabled = True
Else
    fraCharDelimiters.Enabled = False
End If


If cmbFormat = "crEFTHTML32Standard" And cmbDestination = "crEDTDiskFile" Then
    lblHTMLFilename.Enabled = True
    txtHTMLFilePath.Enabled = True
    cmdHTMLFilePath.Enabled = True
Else
    lblHTMLFilename.Enabled = False
    txtHTMLFilePath.Enabled = False
    cmdHTMLFilePath.Enabled = False
End If


End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdDestPath_Click()

cmdlg1.ShowSave
txtDestPath = cmdlg1.FileName

End Sub

Private Sub cmdExport_Click()
On Error Resume Next

'Value returned by a Msgbox
Dim Result As String

With CrystalExportOptions

    'GENERAL TAB
    .DestinationType = cmbDestination.ListIndex
    .FormatType = cmbFormat.ListIndex
    
    .CharFieldDelimiter = txtFieldDel
    .CharStringDelimiter = txtStringDel

    .NumberOfLinesPerPage = txtNumOfLinesPerPage
    
    .UseReportDateFormat = CBool(chkUseRptDateFormat.Value)
    .UseReportNumberFormat = CBool(chkUseRptNumFormat.Value)
    
    .DiskFileName = txtDestPath
    
    .HTMLFileName = txtHTMLFilePath
    
    
    'EXCHANGE OPTIONS TAB
    .ExchangeDestinationType = cmbDestType.ListIndex
    .ExchangeFolderPath = txtfolderpath
    .ExchangePassword = txtXPassword
    .ExchangeProfile = txtXProfile
    .ExchangeTabHasColumnHeadings = CBool(chkXTabHasColumnHeadings.Value)
    
    
    'MAIL OPTIONS TAB
    .MailBccList = txtBCCList
    .MailCcList = txtCCList
    .MailToList = txtToList
    .MailSubject = txtSubject
    .MailMessage = txtMessage
    
    
    'ODBC OPTIONS TAB
    .ODBCDataSourceName = txtDSN
    .ODBCDataSourcePassword = txtDSNPassword
    .ODBCDataSourceUserID = txtDSNUserID
    .ODBCExportTableName = txtTableName
    
End With


If chkPromptUser.Value = Checked Then
    CrystalExportOptions.PromptForExportOptions
    The_Crystal_Report.Export False
Else
    'Error checking to let user know that an error will occur if
    'all values are not set
    Result = MsgBox("Each export Format and Destination requires certain " & Chr(10) & _
            "values to be entered.  If any one of these values " & Chr(10) & _
            "is missing, then an error will occur!" & Chr(10) & Chr(10) & _
            "Do you still want to continue?", vbYesNo, "Export...")
    
    If Result = vbYes Then
        The_Crystal_Report.Export False
    Else
        Exit Sub
    End If
    
End If

'If any kind of error occurs, trap it and display message
If Err.Number <> 0 Then
    MsgBox "Export failed!"
Else
    MsgBox "Exported!"
End If

End Sub

Private Sub cmdHTMLFilePath_Click()

cmdlg1.ShowSave
txtHTMLFilePath = cmdlg1.FileName

End Sub

Private Sub cmdReset_Click()

'Just for code reuse in other routines
Call Read_All_Values

cmbFormat.Text = cmbFormat.List(1)
cmbDestination.Text = cmbDestination.List(1)

End Sub

Private Sub Form_Load()

'Get the ExportOptions object
Set CrystalExportOptions = The_Crystal_Report.ExportOptions

'Populate the Format combobox with all possibilities
'These items will show whether you have the export
'dll or not
cmbFormat.AddItem "crEFTNoFormat"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 0
cmbFormat.AddItem "crEFTCrystalReport"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 1
cmbFormat.AddItem "crEFTDataInterchange"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 2
cmbFormat.AddItem "crEFTRecordStyle"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 3
cmbFormat.AddItem "crEFTRichText"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 4
cmbFormat.AddItem "crEFTCommaSeparatedValues"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 5
cmbFormat.AddItem "crEFTTabSeparatedValues"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 6
cmbFormat.AddItem "crEFTCharSeparatedValues"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 7
cmbFormat.AddItem "crEFTText"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 8
cmbFormat.AddItem "crEFTTabSeparatedText"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 9
cmbFormat.AddItem "crEFTPaginatedText"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 10
cmbFormat.AddItem "crEFTLotus123WKS"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 11
cmbFormat.AddItem "crEFTLotus123WK1"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 12
cmbFormat.AddItem "crEFTLotus123WK3"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 13
cmbFormat.AddItem "crEFTWordForWindows"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 14
cmbFormat.AddItem "crEFTWordForDOS"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 15
cmbFormat.AddItem "crEFTWordPerfect"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 16
cmbFormat.AddItem "crEFTQuattroPro50"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 17
cmbFormat.AddItem "crEFTExcel21"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 18
cmbFormat.AddItem "crEFTExcel30"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 19
cmbFormat.AddItem "crEFTExcel40"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 20
cmbFormat.AddItem "crEFTExcel50"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 21
cmbFormat.AddItem "crEFTExcel50Tabular"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 22
cmbFormat.AddItem "crEFTODBC"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 23
cmbFormat.AddItem "crEFTHTML32Standard"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 24
cmbFormat.AddItem "crEFTExplorer32Extend"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 25
cmbFormat.AddItem "crEFTNetScape20"
    cmbFormat.ItemData(cmbFormat.NewIndex) = 26
'Set the text to the CRYSTAL report format
cmbFormat.Text = cmbFormat.List(1)

'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'Populate the Destination combobox with all possibilities
'These items will show whether you have the destination
'dll or not
cmbDestination.AddItem "crEDTNoDestination"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 0
cmbDestination.AddItem "crEDTDiskFile"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 1
cmbDestination.AddItem "crEDTEMailMAPI"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 2
cmbDestination.AddItem "crEDTEMailVIM"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 3
cmbDestination.AddItem "crEDTMicrosoftExchange"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 4
cmbDestination.AddItem "crEDTApplication"
    cmbDestination.ItemData(cmbDestination.NewIndex) = 5
'Set the text to the DISK Destination
cmbDestination.Text = cmbDestination.List(1)

'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/


cmbDestType.AddItem "crExhangeFolderType"
    cmbDestType.ItemData(cmbDestType.NewIndex) = 0
cmbDestType.AddItem "crExchangePostDocMessage"
    cmbDestType.ItemData(cmbDestType.NewIndex) = 1011
'Set the text to the Exchange Folder Type
cmbDestType.Text = cmbDestType.List(0)

'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'Just for code reuse in other routines
'Call this routine to read any set information
Call Read_All_Values

End Sub

Public Sub Read_All_Values()

With CrystalExportOptions

    'READ GENERAL INFORMATION
    cmbFormat.Text = cmbFormat.List(1) '.FormatType -- READ ONLY
    'Read the corresponding Dll
    'This only returns a dllName after a report is run
    lblFormatDLL.Caption = "DLL: " & CrystalExportOptions.FormatDllName
    
    cmbDestination.Text = cmbDestination.List(1) '.DestinationType -- READ ONLY
    'Read the corresponding Dll
    'This only returns a dllName after a report is run
    lblDestDLL.Caption = "DLL: " & CrystalExportOptions.DestinationDllName

    txtFieldDel = .CharFieldDelimiter
    txtStringDel = .CharStringDelimiter
    
    txtNumOfLinesPerPage = .NumberOfLinesPerPage
    
    chkUseRptDateFormat = Abs(.UseReportDateFormat)
    chkUseRptNumFormat = Abs(.UseReportNumberFormat)
    
    txtDestPath = .DiskFileName
    txtHTMLFilePath = .HTMLFileName
    
    
    'READ EXCHANGE OPTIONS
    cmbDestType = cmbDestType.List(0) '.ExchangeDestinationType -- READ ONLY
    txtXFolderPath = .ExchangeFolderPath
    txtXPassword = ""
    txtXProfile = .ExchangeProfile
    chkXTabHasColumnHeadings = Abs(.ExchangeTabHasColumnHeadings)
    
    'READ MAIL OPTIONS
    txtBCCList = .MailBccList
    txtCCList = .MailCcList
    txtToList = .MailToList
    txtSubject = .MailSubject
    txtMessage = .MailMessage
    
    'READ ODBC OPTIONS
    txtDSN = .ODBCDataSourceName
    txtDSNPassword = ""
    txtDSNUserID = .ODBCDataSourceUserID
    txtTableName = .ODBCExportTableName
    
End With

End Sub

