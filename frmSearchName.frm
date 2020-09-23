VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSearchName 
   Caption         =   "Search Contact - By Name"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   5895
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   11895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSGridResults 
         Height          =   5295
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   10
         _Band(0)._MapCol(0)._Name=   "ID"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(0)._Alignment=   7
         _Band(0)._MapCol(0)._Hidden=   -1  'True
         _Band(0)._MapCol(1)._Name=   "FirstName"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "MI"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "LastName"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(4)._Name=   "Address1"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(0)._MapCol(5)._Name=   "Address2"
         _Band(0)._MapCol(5)._RSIndex=   5
         _Band(0)._MapCol(6)._Name=   "City"
         _Band(0)._MapCol(6)._RSIndex=   6
         _Band(0)._MapCol(7)._Name=   "State"
         _Band(0)._MapCol(7)._RSIndex=   7
         _Band(0)._MapCol(8)._Name=   "Zip"
         _Band(0)._MapCol(8)._RSIndex=   8
         _Band(0)._MapCol(9)._Name=   "Phone"
         _Band(0)._MapCol(9)._RSIndex=   9
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   5880
         TabIndex        =   4
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   5880
         TabIndex        =   5
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   1365
         TabIndex        =   0
         Top             =   360
         Width           =   2835
      End
      Begin VB.TextBox txtMi 
         Height          =   285
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   1365
         TabIndex        =   2
         Top             =   720
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MI :"
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   405
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   765
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmSearchName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
'clear boxes setfocus to txtfirst
txtFirst.Text = ""
txtMi.Text = ""
txtLast.Text = ""
MSGridResults.Visible = False
DEContact.rsCmdSearch.Close
DEContact.ConnSearch.Close
txtFirst.SetFocus
End Sub

Private Sub CmdClose_Click()
Unload Me 'unload form
Set FrmSearchName = Nothing 'set form = nothing
End Sub

Private Sub CmdSearch_Click()
'Dim variable for sql query
Dim sSQL As String

On Error GoTo ErrHandler

sSQL = "Select * from contact where FirstName LIKE'%" & txtFirst.Text & "%' and LastName LIKE'%" & txtLast.Text & "%' and MI LIKE'%" & txtMi.Text & "%'"
'open connection and set connectionstring
DEContact.ConnSearch.Open sConnString
DEContact.rsCmdSearch.Open sSQL, DEContact.ConnSearch, adOpenDynamic, adLockPessimistic

If DEContact.rsCmdSearch.EOF = True Or DEContact.rsCmdSearch.RecordCount = 0 Then
    'Change content of headers and show no results found
    MSGridResults.TextMatrix(0, 0) = "First Name"
    MSGridResults.TextMatrix(0, 1) = "MI"
    MSGridResults.TextMatrix(0, 2) = "Last Name"
    MSGridResults.TextMatrix(0, 3) = "Address 1"
    MSGridResults.TextMatrix(0, 4) = "Address 2"
    MSGridResults.TextMatrix(0, 5) = "City"
    MSGridResults.TextMatrix(0, 6) = "State"
    MSGridResults.TextMatrix(0, 7) = "Zip Code"
    MSGridResults.TextMatrix(0, 8) = "Phone"
    MSGridResults.TextMatrix(1, 0) = "No"
    MSGridResults.TextMatrix(1, 1) = "Results"
    MSGridResults.TextMatrix(1, 2) = "Found"
    MSGridResults.Visible = True
    Exit Sub
End If
    'set grid to display recordset
    Set MSGridResults.Recordset = DEContact.rsCmdSearch
    'Change content of headers
    MSGridResults.TextMatrix(0, 0) = "First Name"
    MSGridResults.TextMatrix(0, 1) = "MI"
    MSGridResults.TextMatrix(0, 2) = "Last Name"
    MSGridResults.TextMatrix(0, 3) = "Address 1"
    MSGridResults.TextMatrix(0, 4) = "Address 2"
    MSGridResults.TextMatrix(0, 5) = "City"
    MSGridResults.TextMatrix(0, 6) = "State"
    MSGridResults.TextMatrix(0, 7) = "Zip Code"
    MSGridResults.TextMatrix(0, 8) = "Phone"
    
    'close the connection
    DEContact.rsCmdSearch.Close
    DEContact.ConnSearch.Close
    
    'show the results
    MSGridResults.Visible = True
    Exit Sub

ErrHandler:
    ErrHandler
End Sub

Private Sub ErrHandler()
    If Err.Number <> 0 Then
       msg = "An Error has occured." & Chr(13) & Chr(13) & _
       "Error # " & Str(Err.Number) & " was generated by " _
             & Err.Source & Chr(13) & Err.Description
       MsgBox msg, vbOKOnly + vbCritical, "Error"

        On Error Resume Next
        DEContact.rsCmdSearch.Close
        DEContact.ConnSearch.Close
        MSGridResults.Visible = False
        Exit Sub
    End If
End Sub

