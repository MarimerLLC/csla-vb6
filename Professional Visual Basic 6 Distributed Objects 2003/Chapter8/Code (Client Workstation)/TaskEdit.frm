VERSION 5.00
Begin VB.Form TaskEdit 
   Caption         =   "Task"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtPercent 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtDays 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Percent complete"
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
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Projected days"
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
      Width           =   1455
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
      Width           =   1455
   End
End
Attribute VB_Name = "TaskEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjTask As Task
Attribute mobjTask.VB_VarHelpID = -1
Private mrsTask As Recordset

Public Sub Component(TaskObject As Task)
  Set mobjTask = TaskObject
End Sub

Private Sub Form_Load()
  Dim objDS As BusinessObjects
  
  mflgLoading = True
  With mobjTask
    EnableOK .IsValid
    .BeginEdit
  End With
  
  Set objDS = New BusinessObjects
  objDS.Add mobjTask, "Task"
  
  Set mrsTask = New Recordset
  mrsTask.Open "Task", "Provider=ODSOLEDB"
  
  ' Set mfmtPercent = New StdDataFormat
  ' mfmtPercent.Format = "0.0"
  
  Set txtName.DataSource = mrsTask
  txtName.DataField = "Name"
  Set txtDays.DataSource = mrsTask
  txtDays.DataField = "ProjectedDays"
  Set txtPercent.DataSource = mrsTask
  ' Set txtPercent.DataFormat = mfmtPercent
  txtPercent.DataField = "PercentComplete"
  
  objDS.RemoveAll
  Set objDS = Nothing
  mflgLoading = False
End Sub

Private Sub cmdOK_Click()
  mobjTask.ApplyEdit
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  mobjTask.CancelEdit
  Unload Me
End Sub

Private Sub cmdApply_Click()
  mobjTask.ApplyEdit
  mobjTask.BeginEdit
End Sub

Private Sub EnableOK(flgValid As Boolean)
  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid
End Sub

Private Sub mobjTask_Valid(IsValid As Boolean)
  EnableOK IsValid
End Sub


