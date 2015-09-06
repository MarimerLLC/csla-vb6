VERSION 5.00
Begin VB.Form VideoSearch 
   Caption         =   "Video Search"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtStudio 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Studio"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "VideoSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgOK As Boolean

Public Property Get OK() As Boolean

  OK = mflgOK

End Property

Public Property Get ResultTitle() As String

  ResultTitle = txtTitle

End Property

Public Property Get ResultStudio() As String

  ResultStudio = txtStudio

End Property

Private Sub cmdCancel_Click()

  mflgOK = False
  Hide

End Sub

Private Sub cmdOK_Click()

  mflgOK = True
  Hide

End Sub

