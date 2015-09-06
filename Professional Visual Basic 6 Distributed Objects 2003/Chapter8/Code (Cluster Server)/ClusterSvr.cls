VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClusterSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjServer As ServerMonitor.Monitor
Private mlngStatus As Long
Private mdtUpdate As Date
Private mstrServer As String

Friend Sub Attach(Server As String)
  mstrServer = Server
  mdtUpdate = Now

  On Error Resume Next
  Set mobjServer = CreateObject("ServerMonitor.Monitor", Server)

  If Err Then
    mlngStatus = 0
  Else
    mobjServer.Initialize Me
    If Err Then
      Set mobjServer = Nothing
      mlngStatus = 0
    End If
  End If

End Sub

Public Sub UpdateStatus(NewStatus As Long)
  mlngStatus = NewStatus
  mdtUpdate = Now
End Sub

Friend Property Get Status() As Long
  If DateDiff("s", mdtUpdate, Now) > REPORT_INTERVAL / 1000 + 1 Then
    Status = 0
  Else
    Status = mlngStatus
  End If
End Property

Friend Sub Shutdown()
  If Not mobjServer Is Nothing And Status > 0 Then
    mobjServer.Shutdown
    Set mobjServer = Nothing
  End If
End Sub

Friend Property Get ServerName() As String
  ServerName = mstrServer
End Property


