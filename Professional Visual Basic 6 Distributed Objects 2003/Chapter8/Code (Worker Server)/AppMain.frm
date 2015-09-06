VERSION 5.00
Begin VB.Form AppMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "AppMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjClient As Monitor

Public Sub Register(Client As Monitor)
  Set mobjClient = Client
End Sub

Public Sub Deregister()
  Set mobjClient = Nothing
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  mobjClient.SendStatus
  Timer1.Enabled = True
End Sub
