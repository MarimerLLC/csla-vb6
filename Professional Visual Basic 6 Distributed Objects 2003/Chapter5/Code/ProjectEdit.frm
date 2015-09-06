VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ProjectEdit 
   Caption         =   "Project"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tasks"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4095
      Begin MSDataListLib.DataList lstTasks 
         Height          =   2400
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4233
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "A&dd"
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
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
Attribute VB_Name = "ProjectEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjProject As Project
Attribute mobjProject.VB_VarHelpID = -1
Private mrsProject As Recordset
Private mrsTasks As Recordset

Public Sub Component(ProjectObject As Project)
  Set mobjProject = ProjectObject
End Sub

Private Sub Form_Load()
  Dim objDS As BusinessObjects
  
  mflgLoading = True
  With mobjProject
    EnableOK .IsValid
    .BeginEdit
  End With
  
  ' Register our business objects with the
  ' data source
  Set objDS = New BusinessObjects
  objDS.Add mobjProject, "Project"
  objDS.Add mobjProject.Tasks, "Tasks"

  Set mrsProject = New Recordset
  mrsProject.Open "Project", "Provider=ODSOLEDB"
  Set txtName.DataSource = mrsProject
  txtName.DataField = "Name"
  
  Set mrsTasks = New Recordset
  mrsTasks.Open "Tasks:ID,Name", "Provider=ODSOLEDB"
  With lstTasks
    Set .RowSource = mrsTasks
    .ListField = "Name"
    .BoundColumn = "ID"
  End With
  
  objDS.RemoveAll
  Set objDS = Nothing
  mflgLoading = False
End Sub

Private Sub cmdOK_Click()
  mobjProject.ApplyEdit
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  mobjProject.CancelEdit
  Unload Me
End Sub

Private Sub cmdApply_Click()
  mobjProject.ApplyEdit
  mobjProject.BeginEdit
End Sub

Private Sub EnableOK(flgValid As Boolean)
  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid
End Sub

Private Sub mobjProject_Valid(IsValid As Boolean)
  EnableOK IsValid
End Sub

Private Sub cmdAdd_Click()
  Dim frmTask As TaskEdit
  
  Set frmTask = New TaskEdit
  frmTask.Component mobjProject.Tasks.Add
  frmTask.Show vbModal
  mrsTasks.Requery
End Sub

Private Sub cmdEdit_Click()
  Dim frmTask As TaskEdit
  
  Set frmTask = New TaskEdit
  frmTask.Component mobjProject.Tasks(SelectedItem(lstTasks, mobjProject.Tasks))
  frmTask.Show vbModal
  mrsTasks.Requery
End Sub

Private Sub cmdRemove_Click()
  mobjProject.Tasks.Remove SelectedItem(lstTasks, mobjProject.Tasks)
  mrsTasks.Requery
End Sub

