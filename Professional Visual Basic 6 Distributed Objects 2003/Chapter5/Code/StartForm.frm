VERSION 5.00
Begin VB.Form StartForm 
   Caption         =   "Start"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Client ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdNew_Click()
  Dim objClient As Client
  Dim frmClient As ClientEdit
  
  Set objClient = New Client
  Set frmClient = New ClientEdit
  frmClient.Component objClient
  frmClient.Show vbModal
End Sub

Private Sub cmdOpen_Click()
  Dim objClient As Client
  Dim frmClient As ClientEdit
  
  If Val(txtID) > 0 Then
    Set objClient = New Client
    On Error Resume Next
    objClient.Load Val(txtID)
    If Err Then
      MsgBox "Client ID not on file", vbExclamation
      Exit Sub
    End If
    On Error GoTo 0
    Set frmClient = New ClientEdit
    frmClient.Component objClient
    frmClient.Show vbModal
  Else
    MsgBox "You must supply a value", vbInformation
  End If
End Sub

