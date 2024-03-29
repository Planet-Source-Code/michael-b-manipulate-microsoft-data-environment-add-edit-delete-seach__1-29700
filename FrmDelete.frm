VERSION 5.00
Begin VB.Form FrmDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Contact"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4200
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   308
      TabIndex        =   3
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1568
      TabIndex        =   4
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2843
      TabIndex        =   5
      Top             =   1410
      Width           =   1050
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   195
      Width           =   2835
   End
   Begin VB.TextBox txtMi 
      Height          =   285
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   1
      Top             =   585
      Width           =   495
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   975
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Name :"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MI :"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   630
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Last Name :"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1020
      Width           =   855
   End
End
Attribute VB_Name = "FrmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDelete_Click()
txtFirst.Text = UCase(txtFirst.Text)
txtMi.Text = UCase(txtMi.Text)
txtLast.Text = UCase(txtLast.Text)

'msgbox variables
    Dim msg, title, style, response As String
    
    ' Dim sSQL query string
    Dim sSQL As String
    
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
    
    DEContact.ConnDelete.Open sConnString
    DEContact.ConnDelete.BeginTrans
    
    sSQL = "DELETE FROM CONTACT WHERE LastName='" & txtLast.Text & "' and FirstName='" & txtFirst.Text & "' and MI='" & txtMi.Text & "'"

    DEContact.ConnDelete.Execute (sSQL)
    
    'Message to confirm adding of new contact
    msg = "Are you sure you want to delete this contact?"
    style = vbYesNo + vbQuestion
    title = "Confirmation"
    response = MsgBox(msg, style, title)
    
    'confirmation validation
    If response = vbYes Then
        'then set the adding in action
        DEContact.ConnDelete.CommitTrans
        'run the command button "Cancel"'s click event
        CmdCancel_Click
        Exit Sub
    Else
        'then don't add
        DEContact.ConnDelete.RollbackTrans
        DEContact.ConnDelete.Close
        Exit Sub
    End If

ErrHandler:
    ErrHandler
End Sub

Private Sub CmdCancel_Click()
'clear boxes and setfocus to txtfirst
txtFirst.Text = ""
txtMi.Text = ""
txtLast.Text = ""
txtFirst.SetFocus
End Sub

Private Sub CmdClose_Click()
Unload Me ' unload me
Set FrmDelete = Nothing 'set for = nothing
End Sub

Private Sub ErrHandler()
Dim msg As String

    If Err.Number <> 0 Then
       msg = "An Error has occured." & Chr(13) & Chr(13) & _
       "Error # " & Str(Err.Number) & " was generated by " _
             & Err.Source & Chr(13) & Err.Description
       MsgBox msg, vbOKOnly + vbCritical, "Error"

        On Error Resume Next
        DEContact.ConnDelete.Close
        Exit Sub
    End If
End Sub

