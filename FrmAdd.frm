VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Contact"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4290
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2888
      TabIndex        =   11
      Top             =   3855
      Width           =   1050
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1620
      TabIndex        =   10
      Top             =   3855
      Width           =   1050
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   353
      TabIndex        =   9
      Top             =   3855
      Width           =   1050
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      ItemData        =   "FrmAdd.frx":0000
      Left            =   1335
      List            =   "FrmAdd.frx":00A3
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2490
      Width           =   720
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1335
      TabIndex        =   5
      Top             =   2110
      Width           =   2835
   End
   Begin VB.TextBox txtAddress2 
      Height          =   285
      Left            =   1335
      TabIndex        =   4
      Top             =   1718
      Width           =   2835
   End
   Begin VB.TextBox txtAddress1 
      Height          =   285
      Left            =   1335
      TabIndex        =   3
      Top             =   1326
      Width           =   2835
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1335
      TabIndex        =   2
      Top             =   934
      Width           =   2835
   End
   Begin VB.TextBox txtMi 
      Height          =   285
      Left            =   1335
      MaxLength       =   1
      TabIndex        =   1
      Top             =   542
      Width           =   495
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1335
      TabIndex        =   0
      Top             =   150
      Width           =   2835
   End
   Begin MSMask.MaskEdBox MebZip 
      Height          =   285
      Left            =   1335
      TabIndex        =   7
      Top             =   2910
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
      Left            =   1335
      TabIndex        =   8
      Top             =   3330
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Zip Code:"
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   2955
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   195
      Left            =   210
      TabIndex        =   19
      Top             =   3375
      Width           =   510
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "State:"
      Height          =   195
      Left            =   210
      TabIndex        =   18
      Top             =   2550
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "City :"
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   2155
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Address 2 :"
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   1763
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Address 1 :"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   1371
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Last Name :"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   979
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MI :"
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   587
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Name :"
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   195
      Width           =   840
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As New ADODB.Recordset ' dim the recordset
Dim sSQL As String ' This variable is to hold the sql statement

Private Sub CmdAdd_Click()
    'msgbox variables
    Dim msg, title, style, response As String
    
    'Required Fields validation
    If txtFirst.Text = "" Then
        MsgBox "First name is a required field.", vbOKOnly, "Required Field"
        txtFirst.SetFocus
        Exit Sub
    End If
    
    If txtLast.Text = "" Then
        MsgBox "Last name is a required field.", vbOKOnly, "Required Field"
        txtLast.SetFocus
        Exit Sub
    End If
    
    If txtAddress1.Text = "" Then
        MsgBox "Address 1 is a required field.", vbOKOnly, "Required Field"
        txtAddress1.SetFocus
        Exit Sub
    End If
    
    If txtCity.Text = "" Then
        MsgBox "City is a required field.", vbOKOnly, "Required Field"
        txtCity.SetFocus
        Exit Sub
    End If
    
    If cboState.Text = "" Or cboState.Text = "Select a State" Then
        MsgBox "First name is a required field.", vbOKOnly, "Required Field"
        txtFirst.SetFocus
        Exit Sub
    End If

    
    On Error GoTo ErrHandler
    
    DEContact.ConnAdd.ConnectionString = sConnString 'Set the connectionstring for the connection
    DEContact.ConnAdd.Open 'open the Connection
    
    sSQL = "select * from contact where LastName='" & txtLast.Text & "' and FirstName='" & txtFirst.Text & "' order by LastName,FirstName"

    'set the Recordset to the executed sSQL
    Set RS = DEContact.ConnAdd.Execute(sSQL)

   
    'we want to test to see if there ARE results
    If RS.EOF = False Or RS.RecordCount <> 0 Then
        'if results then
        MsgBox "This contact already exists.", vbOKOnly + vbExclamation, ""
        
        'close and clear ado variables
        RS.Close
        DEContact.ConnAdd.Close
        Set RS = Nothing
        
        Exit Sub
    Else
        'if no results then
        
        'close and clear ado variables
        RS.Close
        DEContact.ConnAdd.Close
        Set RS = Nothing
        
        'open the connection to the connectionstring
         DEContact.ConnAdd.Open sConnString
        'open the table contact,set connection,set cursortype and set locktype
        RS.Open "CONTACT", DEContact.ConnAdd, adOpenDynamic, adLockPessimistic
        
        'start the transaction
        'A transaction is used to ensure the data is
        'update correctly before applying the change in
        'in full. Think of a senerio with a bank.
        'You don't want an ATM machine debiting your
        'account and crashing before it gives you the
        'money? Well in the same sense Transactions are
        'used for databases. It basically takes the DB
        'copys it and makes the changes. Then once the
        'changes are made commit then to the actual DB.
        'This prevents the database from being affected
        'before ALL records are changes.
        
        DEContact.ConnAdd.BeginTrans
        RS.AddNew
        
        'Start the adding
        'For Readability I force uppercase to the fields
        RS.Fields("FirstName") = UCase(txtFirst.Text)
        RS.Fields("MI") = UCase(txtMi.Text)
        RS.Fields("LastName") = UCase(txtLast.Text)
        RS.Fields("Address1") = UCase(txtAddress1.Text)
        RS.Fields("Address2") = UCase(txtAddress2.Text)
        RS.Fields("City") = UCase(txtCity.Text)
        RS.Fields("State") = UCase(cboState.Text)
        RS.Fields("Zip") = MebZip.Text
        RS.Fields("Phone") = MebPhone.Text
        RS.Update
        
        'Message to confirm adding of new contact
        msg = "Are you sure you want to add this contact?"
        style = vbYesNo + vbQuestion
        title = "Confirmation"
        response = MsgBox(msg, style, title)
        
        'confirmation validation
        If response = vbYes Then
            'then set the adding in action
            DEContact.ConnAdd.CommitTrans
            'run the command button "Cancel"'s click event
            CmdCancel_Click
            Exit Sub
        Else
            'then don't add
            DEContact.ConnAdd.RollbackTrans
        End If
        
        'close and clear ado variables
        RS.Close
        DEContact.ConnAdd.Close
        Set RS = Nothing
        Exit Sub
    End If

ErrHandler:
    ErrHandler
End Sub

Private Sub CmdCancel_Click()
'Clear all boxes and setfocus to txtfirst
txtFirst.Text = ""
txtMi.Text = ""
txtLast.Text = ""
txtAddress1.Text = ""
txtAddress2.Text = ""
txtCity.Text = ""
cboState.ListIndex = 1
MebZip.Text = "00000-0000"
MebPhone.Text = "(___)-___-____"
txtFirst.SetFocus
End Sub

Private Sub CmdClose_Click()
Unload Me 'unload form
Set FrmAdd = Nothing 'set form = nothing
End Sub

Private Sub Form_Load()
cboState.ListIndex = 1
End Sub

Private Sub ErrHandler()
Dim msg As String

    If Err.Number <> 0 Then
       msg = "An Error has occured." & Chr(13) & Chr(13) & _
       "Error # " & Str(Err.Number) & " was generated by " _
             & Err.Source & Chr(13) & Err.Description
       MsgBox msg, vbOKOnly + vbCritical, "Error"

    On Error Resume Next
    RS.Close
    DEContact.ConnAdd.Close
    Set RS = Nothing
        Exit Sub
    End If
End Sub
