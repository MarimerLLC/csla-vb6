VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Receiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Arrived()

Private WithEvents mobjQE As MSMQEvent
Attribute mobjQE.VB_VarHelpID = -1
Private mobjQ As MSMQQueue

Private Sub Class_Initialize()
  Dim objQI As MSMQQueueInfo
  
  Set objQI = New MSMQQueueInfo
  With objQI
    .PathName = ".\TaskTracking"
    Set mobjQ = .Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
  End With
  
  Set mobjQE = New MSMQEvent
  mobjQ.EnableNotification mobjQE
End Sub

Private Sub Class_Terminate()
  Set mobjQE = Nothing
  mobjQ.Close
  Set mobjQ = Nothing
End Sub

Public Sub CheckMessages()
  Dim objM As MSMQMessage
  Dim objProxy As Object
  
  On Error Resume Next
  Do While True
    Set objM = mobjQ.Receive(ReceiveTimeout:=10)
    If Err Then Exit Do
    
    Set objProxy = objM.Body
    
    Set objM = Nothing
    If Not objProxy Is Nothing Then RaiseEvent Arrived
    Set objProxy = Nothing
  Loop
  mobjQ.EnableNotification mobjQE
End Sub

Private Sub mobjQE_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
  CheckMessages
End Sub

