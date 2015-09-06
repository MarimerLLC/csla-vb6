VERSION 5.00
Begin VB.Form CustomerSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Search"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Phone"
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "CustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgOK As Boolean

Private Sub cmdCancel_Click()

  mflgOK = False
  Hide

End Sub

Private Sub cmdOK_Click()

  mflgOK = True
  Hide

End Sub

Public Property Get OK() As Boolean

  OK = mflgOK

End Property

Public Property Get ResultName() As String

  ResultName = txtName

End Property

Public Property Get ResultPhone() As String

  ResultPhone = txtPhone

End Property

