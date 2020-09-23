VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Data Environment Example Program"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5220
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3572
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/11/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "7:12 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
      Begin VB.Menu mnuAddAddContact 
         Caption         =   "Add Contact"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditContact 
         Caption         =   "Edit Contact"
      End
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
      Begin VB.Menu mnuDeleteContact 
         Caption         =   "Delete Contact"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchName 
         Caption         =   "By Name"
      End
      Begin VB.Menu mnuSearchPhone 
         Caption         =   "By Phone Number"
      End
      Begin VB.Menu mnuSearchZipCode 
         Caption         =   "By Zip Code"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu MnuReportAll 
         Caption         =   "All Contacts"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
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
'Author:        Lord Nauction (lordnauction@hotmail.com)
'Code Purpose:  Examples how to manipulate a microsoft data
'               environment and other ADO functions.
'Version Type:  Open Source
'Comments:      This is free to use in anyway you like.
'               As a lot of other people's reasons for posting source
'               I post this because i couldn't find any help with the
'               MS Data Environment. So after playing around with it i finally
'               Got it to do what i need it to.
'               This Example is composed of basic ways to manipulate the
'               Data environment. Hope you enjoy it and learn something.
'
'               The example database is a basic contact database.
'               The easiest test data i could think of is the
'               age old use of the contact list.
'
'
'Some Subjects Covered: Adding, Editing, Deleting, and searching records
'                       using the DE and ADO commands.
'                       Learn how to use transactions with a MS Access database.

Option Explicit

Private Sub MDIForm_Load()
    'load the windowstate and location
    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
    
    'Get database path from app settings in registry
    sDatabasePath = GetSetting(App.title, "Database", "Path", App.Path & "\nauction.mdb")
    
    'check if database exists
    If Dir(sDatabasePath) = "" Then
        MsgBox "Database not found. Please locate the database file.", vbOKOnly + vbExclamation, "Database not found"
        LocateDatabase
        
        'Since database is not found we must reset
        'the connectionstring that will be used for
        'the dataenvironment and any ado code/controls used
        'in this example
        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=brhn1aca2"

    End If
    
    'incase of any changes re-save the path and connstring
    SaveSetting App.title, "Database", "Path", sDatabasePath
    SaveSetting App.title, "Database", "ConnString", sConnString
    
    'Set all DE connections to current connectionstring
    DEContact.ConnAdd.ConnectionString = sConnString
    DEContact.ConnEdit.ConnectionString = sConnString
    DEContact.ConnDelete.ConnectionString = sConnString
    DEContact.ConnSearch.ConnectionString = sConnString
    DEContact.ConnReport.ConnectionString = sConnString
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'save the windowstate and location
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
        SaveSetting App.title, "Database", "Path", sDatabasePath
        SaveSetting App.title, "Database", "ConnString", sConnString

    End If
End Sub

Private Sub mnuAbout_Click()
'load and show the about frame to a vbmoal state
'vbmodal makes the program only allow focus to
'that window until it is closed.

frmAbout.Show vbModal
End Sub

Private Sub mnuAddAddContact_Click()
FrmAdd.Show 'load and show the add form
FrmAdd.Top = 0
FrmAdd.Left = 0
End Sub

Private Sub mnuDeleteContact_Click()
FrmDelete.Show
FrmDelete.Top = 0
FrmDelete.Left = 0
End Sub



Private Sub mnuEditContact_Click()
FrmEdit.Show
FrmEdit.Top = 0
FrmEdit.Left = 0
End Sub

Private Sub MnuReportAll_Click()
'run the report command for the freshest data
DEContact.ConnReport.Open sConnString
DEContact.CmdReport
RptContact.Show
End Sub

Private Sub mnuSearchName_Click()
FrmSearchName.Show
End Sub

Private Sub mnuSearchPhone_Click()
FrmSearchPhone.Show
End Sub

Private Sub mnuSearchZipCode_Click()
FrmSearchZip.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub LocateDatabase()
  Dim s1 As Integer
  Dim s2 As String
  
  dlgCommonDialog.CancelError = True
  On Error GoTo ErrHandler
  ' Set flags
  dlgCommonDialog.Flags = cdlOFNHideReadOnly
  ' Set filters
  dlgCommonDialog.Filter = "User DB (nauction.mdb)|nauction.mdb"
  ' Specify default filter
  dlgCommonDialog.FilterIndex = 2
  ' Display the Open dialog box
  dlgCommonDialog.ShowOpen
  ' Display name of selected file
  s1 = Len(dlgCommonDialog.FileName)
  s1 = s1 - 12
  s2 = Mid(dlgCommonDialog.FileName, s1 + 1, 12)
  If s2 <> "nauction.mdb" Then
    MsgBox "Please select the correct database with the name nauction.mdb"
    LocateDatabase
    Exit Sub
  End If
  sDatabasePath = dlgCommonDialog.FileName
  Exit Sub
ErrHandler:
  'User pressed the Cancel button
  LocateDatabase
End Sub

