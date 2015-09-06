VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrBody As String
Private mstrAction As String

Public Function Fetch(ID As Long) As String
  Dim objPersist As ClientPersist
  
  Set objPersist = New ClientPersist
  Fetch = objPersist.Fetch(ID)
  Set objPersist = Nothing
End Function

Public Function Save(Buffer As String) As String
  mstrBody = Buffer
  mstrAction = "Save"
  SendMessage
  Save = Buffer
End Function

Public Sub DeleteObject(ID As Long)
  mstrBody = CStr(ID)
  mstrAction = "Delete"
  SendMessage
End Sub

Private Sub SendMessage()
  Dim objQI As MSMQQueueInfo
  Dim objQ As MSMQQueue
  Dim objM As MSMQMessage
  
  On Error GoTo Err:
  Set objQI = New MSMQQueueInfo
  MsgBox ("Getting queue")
  objQI.PathName = "ss3test\TaskTracking"
  Set objQ = objQI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
  MsgBox ("got queue")
  Set objM = New MSMQMessage
  With objM
    .Body = Me
    MsgBox ("Sending message")
    .Send objQ
    MsgBox ("Sent message")
  End With
  Exit Sub
  
Err:
  MsgBox (Err.Number)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Body", mstrBody
    .WriteProperty "Action", mstrAction
  End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
  Dim objPersist As ClientPersist
  
  Set objPersist = New ClientPersist
  
  With PropBag
    Select Case .ReadProperty("Action")
    Case "Save"
      objPersist.Save .ReadProperty("Body")
    Case "Delete"
      objPersist.DeleteObject Val(.ReadProperty("Body"))
    End Select
  End With
  
  Set objPersist = Nothing
End Sub

