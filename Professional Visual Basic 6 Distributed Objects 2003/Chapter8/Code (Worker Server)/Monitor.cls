VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjCaller As Object
Private mlngRating As Long

Private Const REPORT_INT = 500

Public Sub Initialize(Caller As Object)
  Dim objShell As IWshShell_Class
  
  Set objShell = New IWshShell_Class
  mlngRating = _
    objShell.RegRead("HKEY_LOCAL_MACHINE\Software\Wrox\Cluster\ServerRating")
  Set objShell = Nothing
  
  Set mobjCaller = Caller
  
  Load AppMain
  AppMain.Register Me
  With AppMain.Timer1
    .Interval = REPORT_INT
    .Enabled = True
  End With
  
  SendStatus
End Sub

Public Sub Shutdown()
  AppMain.Deregister
  Set mobjCaller = Nothing
End Sub

Friend Sub SendStatus()
  mobjCaller.UpdateStatus mlngRating
End Sub




