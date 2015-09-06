VERSION 5.00
Begin VB.Form ClientEdit 
   Caption         =   "Client"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   8280
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtContactName 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Projects"
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9495
      Begin VB.ListBox lstProjects 
         Height          =   3765
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "ClientEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjClient As Client
Attribute mobjClient.VB_VarHelpID = -1

Public Sub Component(ClientObject As Client)
  Set mobjClient = ClientObject
End Sub

Private Sub Form_Load()
  mflgLoading = True
  With mobjClient
    EnableOK .IsValid
    ' load object values into form controls
    txtName = .Name
    txtContactName = .ContactName
    txtPhone = .Phone
    ListProjects
    .BeginEdit
  End With
  mflgLoading = False
End Sub

Private Sub cmdOK_Click()
  mobjClient.ApplyEdit
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  mobjClient.CancelEdit
  Unload Me
End Sub

Private Sub cmdApply_Click()
  mobjClient.ApplyEdit
  ListProjects
  mobjClient.BeginEdit
End Sub

Private Sub EnableOK(flgValid As Boolean)
  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid
End Sub

Private Sub mobjClient_Valid(IsValid As Boolean)
  EnableOK IsValid
End Sub

Private Sub txtContactName_Change()
  If mflgLoading Then Exit Sub

  TextChange txtContactName, mobjClient, "ContactName"
End Sub

Private Sub txtContactName_LostFocus()
  txtContactName = TextLostFocus(mobjClient, "ContactName")
End Sub

Private Sub txtName_Change()
  If mflgLoading Then Exit Sub

  TextChange txtName, mobjClient, "Name"
End Sub

Private Sub txtName_LostFocus()
  txtName = TextLostFocus(mobjClient, "Name")
End Sub

Private Sub txtPhone_Change()
  If mflgLoading Then Exit Sub

  TextChange txtPhone, mobjClient, "Phone"
End Sub

Private Sub txtPhone_LostFocus()
  txtPhone = TextLostFocus(mobjClient, "Phone")
End Sub

Private Sub ListProjects()
  Dim objProject As Project
  Dim lngIndex As Long
  
  lstProjects.Clear
  For lngIndex = 1 To mobjClient.Projects.Count
    Set objProject = mobjClient.Projects(lngIndex)
    With objProject
      If .IsDeleted Then
        lstProjects.AddItem .Name & " (d)"
      ElseIf .IsNew Then
        lstProjects.AddItem .Name & " (new)"
      Else
        lstProjects.AddItem .Name
      End If
      lstProjects.ItemData(lstProjects.NewIndex) = lngIndex
    End With
  Next
End Sub

Private Sub cmdAdd_Click()
  Dim frmProject As ProjectEdit
  
  Set frmProject = New ProjectEdit
  frmProject.Component mobjClient.Projects.Add
  frmProject.Show vbModal
  ListProjects
End Sub

Private Sub cmdEdit_Click()
  Dim frmProject As ProjectEdit
  
  Set frmProject = New ProjectEdit
  frmProject.Component _
    mobjClient.Projects(lstProjects.ItemData(lstProjects.ListIndex))
  frmProject.Show vbModal
  ListProjects
End Sub

Private Sub cmdRemove_Click()
  With mobjClient.Projects(lstProjects.ItemData(lstProjects.ListIndex))
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  ListProjects
End Sub

