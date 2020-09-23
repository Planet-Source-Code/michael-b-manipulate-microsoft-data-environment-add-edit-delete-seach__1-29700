VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSearchZip 
   Caption         =   "Search Contact - By Zip Code"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   11895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSGridResults 
         Height          =   6015
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10610
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
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   5760
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
      Begin MSMask.MaskEdBox MebZip 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   375
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#####-####"
         Mask            =   "#####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Zip Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmSearchZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
'clear mebphone and set focus to it
MebZip.Text = "_____-____"
MSGridResults.Visible = False
DEContact.rsCmdSearch.Close
DEContact.ConnSearch.Close
MebPhone.SetFocus
End Sub

Private Sub CmdSearch_Click()
'Dim variable for sql query
Dim sSQL As String

On Error GoTo ErrHandler

sSQL = "Select * from contact where Zip ='" & MebZip.Text & "'"
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


