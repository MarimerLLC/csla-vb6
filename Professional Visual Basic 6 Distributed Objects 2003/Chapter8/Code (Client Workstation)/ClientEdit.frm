VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
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
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtContactName 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
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
      Begin MSDataListLib.DataList lstProjects 
         Height          =   3765
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6641
         _Version        =   393216
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
Private mrsClient As Recordset
Private mrsProjects As Recordset

Public Sub Component(ClientObject As Client)
  Set mobjClient = ClientObject
End Sub

Private Sub cmdAdd_Click()
  Dim frmProject As ProjectEdit
  
  Set frmProject = New ProjectEdit
  frmProject.Component mobjClient.Projects.Add
  frmProject.Show vbModal
  mrsProjects.Requery
End Sub

Private Sub cmdEdit_Click()
  Dim frmProject As ProjectEdit
  
  Set frmProject = New ProjectEdit
  frmProject.Component _
    mobjClient.Projects(SelectedItem(lstProjects, mobjClient.Projects))
  frmProject.Show vbModal
  mrsProjects.Requery
End Sub

Private Sub cmdRemove_Click()
  mobjClient.Projects.Remove SelectedItem(lstProjects, mobjClient.Projects)
  mrsProjects.Requery
End Sub

Private Sub Form_Load()
  Dim objDS As BusinessObjects
  
  mflgLoading = True
  With mobjClient
    EnableOK .IsValid
    .BeginEdit
  End With
      
  ' Register our business objects with the
  ' data source
  Set objDS = New BusinessObjects
  objDS.Add mobjClient, "Client"
  objDS.Add mobjClient.Projects, "Projects"
  
  Set mrsClient = New Recordset
  mrsClient.Open "Client:Name,ContactName,Phone", _
    "Provider=ODSOLEDB"
  Set txtName.DataSource = mrsClient
  txtName.DataField = "Name"
  Set txtContactName.DataSource = mrsClient
  txtContactName.DataField = "ContactName"
  Set txtPhone.DataSource = mrsClient
  txtPhone.DataField = "Phone"
  
  Set mrsProjects = New Recordset
  mrsProjects.Open "Projects:Name,ID", "Provider=ODSOLEDB"
  With lstProjects
    Set .RowSource = mrsProjects
    .ListField = "Name"
    .BoundColumn = "ID"
  End With
  
  objDS.RemoveAll
  Set objDS = Nothing

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
  mrsProjects.Requery
  mobjClient.BeginEdit
End Sub

Private Sub EnableOK(flgValid As Boolean)
  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mrsClient.Close
  Set mrsClient = Nothing
  
  mrsProjects.Close
  Set mrsProjects = Nothing
End Sub

Private Sub mobjClient_Valid(IsValid As Boolean)
  EnableOK IsValid
End Sub


