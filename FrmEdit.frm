VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Contact"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   4590
   Begin VB.Frame FrmTos 
      Caption         =   "Edit Criteria"
      Height          =   1800
      Left            =   30
      TabIndex        =   27
      Top             =   45
      Width           =   4530
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   510
         TabIndex        =   3
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   3030
         TabIndex        =   5
         Top             =   1350
         Width           =   1050
      End
      Begin VB.TextBox txtFirstEdit 
         Height          =   285
         Left            =   1380
         TabIndex        =   0
         Top             =   195
         Width           =   2835
      End
      Begin VB.TextBox txtMIEdit 
         Height          =   285
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   1
         Top             =   585
         Width           =   495
      End
      Begin VB.TextBox txtLastEdit 
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   975
         Width           =   2835
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "First Name :"
         Height          =   195
         Left            =   255
         TabIndex        =   30
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "MI :"
         Height          =   195
         Left            =   255
         TabIndex        =   29
         Top             =   630
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Last Name :"
         Height          =   195
         Left            =   255
         TabIndex        =   28
         Top             =   1020
         Width           =   855
      End
   End
   Begin VB.Frame FrmEdits 
      Caption         =   "Edit Contact"
      Height          =   4410
      Left            =   30
      TabIndex        =   17
      Top             =   1905
      Visible         =   0   'False
      Width           =   4530
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Top             =   255
         Width           =   2835
      End
      Begin VB.TextBox txtMi 
         Height          =   285
         Left            =   1365
         MaxLength       =   1
         TabIndex        =   7
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   1365
         TabIndex        =   8
         Top             =   1035
         Width           =   2835
      End
      Begin VB.TextBox txtAddress1 
         Height          =   285
         Left            =   1365
         TabIndex        =   9
         Top             =   1425
         Width           =   2835
      End
      Begin VB.TextBox txtAddress2 
         Height          =   285
         Left            =   1365
         TabIndex        =   10
         Top             =   1830
         Width           =   2835
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1365
         TabIndex        =   11
         Top             =   2220
         Width           =   2835
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         ItemData        =   "FrmEdit.frx":0000
         Left            =   1365
         List            =   "FrmEdit.frx":00A3
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   405
         TabIndex        =   15
         Top             =   3960
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1665
         TabIndex        =   16
         Top             =   3960
         Width           =   1050
      End
      Begin MSMask.MaskEdBox MebZip 
         Height          =   285
         Left            =   1365
         TabIndex        =   13
         Top             =   3015
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
      Begin MSMask.MaskEdBox MebPhone 
         Height          =   285
         Left            =   1365
         TabIndex        =   14
         Top             =   3435
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "(###)-###-####"
         Mask            =   "(###)-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MI :"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   690
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address 1 :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address 2 :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1875
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "City :"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2265
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Phone:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   3480
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Zip Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3060
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'msgbox variables
    Dim msg, title, style, response As String
    
Private Sub CmdCancel_Click()
On Error Resume Next
DEContact.ConnEdit.RollbackTrans
FrmEdits.Visible = False
DEContact.ConnEdit.Close
FrmTos.Enabled = True
End Sub


Private Sub CmdClose_Click()
Unload Me
Set FrmEdit = Nothing
End Sub

Private Sub CmdEdit_Click()
On Error GoTo ErrHandler

msg = "Are you sure you want to commit the changed?"
style = vbYesNo + vbQuestion
title = "Confirmation"
response = MsgBox(msg, style, title)

If response = vbYes Then
    'commit changes
    DEContact.rsCmdEdit.Update
    DEContact.ConnEdit.CommitTrans
    'clear the boxes
    CmdCancel_Click
    Exit Sub
Else
    Exit Sub
End If

ErrHandler:
    ErrHandler1
End Sub

Private Sub CmdSearch_Click()
    ' Dim sSQL query string
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    txtFirstEdit.Text = UCase(txtFirstEdit.Text)
    txtMIEdit.Text = UCase(txtMIEdit.Text)
    txtLastEdit.Text = UCase(txtLastEdit.Text)
    
    'Required Fields validation
    If txtFirstEdit.Text = "" Then
        MsgBox "First name is a required field.", vbOKOnly, "Required Field"
        txtFirstEdit.SetFocus
        Exit Sub
    End If
    
    If txtLastEdit.Text = "" Then
        MsgBox "Last name is a required field.", vbOKOnly, "Required Field"
        txtLastEdit.SetFocus
        Exit Sub
    End If
    
    sSQL = "Select * from contact where FirstName='" & txtFirstEdit.Text & "' and LastName='" & txtLastEdit.Text & "' and MI='" & txtMIEdit.Text & "'"
    
    'open connection and set connectionstring
    DEContact.ConnEdit.Open sConnString
    DEContact.ConnEdit.BeginTrans
    
    DEContact.rsCmdEdit.Open sSQL, DEContact.ConnEdit, adOpenDynamic, adLockPessimistic
    
    If DEContact.rsCmdEdit.EOF = True Or DEContact.rsCmdEdit.RecordCount = 0 Then
        MsgBox "Contact not found.", vbOKOnly + vbExclamation, "Contact not found."
        DEContact.ConnEdit.RollbackTrans
        DEContact.ConnEdit.Close
        Exit Sub
    End If
    
    Set txtFirst.DataSource = DEContact.rsCmdEdit
    txtFirst.DataField = "FirstName"
    
    Set txtMi.DataSource = DEContact.rsCmdEdit
    txtMi.DataField = "MI"
    
    Set txtLast.DataSource = DEContact.rsCmdEdit
    txtLast.DataField = "LastName"
    
    Set txtAddress1.DataSource = DEContact.rsCmdEdit
    txtAddress1.DataField = "Address1"
    
    Set txtAddress2.DataSource = DEContact.rsCmdEdit
    txtAddress2.DataField = "Address2"

    Set txtCity.DataSource = DEContact.rsCmdEdit
    txtCity.DataField = "City"

    Set cboState.DataSource = DEContact.rsCmdEdit
    cboState.DataField = "State"
    
    Set MebZip.DataSource = DEContact.rsCmdEdit
    MebZip.DataField = "Zip"
    
    Set MebPhone.DataSource = DEContact.rsCmdEdit
    MebPhone.DataField = "Phone"
    
    FrmEdits.Visible = True
    FrmTos.Enabled = False
    Exit Sub
    
ErrHandler:
    ErrHandler
End Sub

Private Sub Command2_Click()
'clear search boxes and setfocus to txtFirstEdit
txtFirstEdit.Text = ""
txtMIEdit.Text = ""
txtLastEdit.Text = ""
txtFirstEdit.SetFocus
End Sub

Private Sub ErrHandler()
    If Err.Number <> 0 Then
       msg = "An Error has occured." & Chr(13) & Chr(13) & _
       "Error # " & Str(Err.Number) & " was generated by " _
             & Err.Source & Chr(13) & Err.Description
       MsgBox msg, vbOKOnly + vbCritical, "Error"

        On Error Resume Next
        DEContact.ConnEdit.RollbackTrans
        DEContact.ConnEdit.Close
        Exit Sub
    End If
End Sub


Private Sub ErrHandler1()
    If Err.Number <> 0 Then
       msg = "An Error has occured." & Chr(13) & Chr(13) & _
       "Error # " & Str(Err.Number) & " was generated by " _
             & Err.Source & Chr(13) & Err.Description
       MsgBox msg, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
End Sub

