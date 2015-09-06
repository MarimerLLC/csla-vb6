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

Public Function Fetch(ID As Long) As String
  Dim objPersist As Object
  
  Set objPersist = Cluster.CreateObject("TaskServer.ClientPersist")
  Fetch = objPersist.Fetch(ID)
  Set objPersist = Nothing
End Function

Public Function Save(Buffer As String) As String
  Dim objPersist As Object
  
  Set objPersist = Cluster.CreateObject("TaskServer.ClientPersist")
  Save = objPersist.Save(Buffer)
  Set objPersist = Nothing
End Function

Public Sub DeleteObject(ID As Long)
  Dim objPersist As Object
  
  Set objPersist = Cluster.CreateObject("TaskServer.ClientPersist")
  objPersist.DeleteObject ID
  Set objPersist = Nothing
End Sub


