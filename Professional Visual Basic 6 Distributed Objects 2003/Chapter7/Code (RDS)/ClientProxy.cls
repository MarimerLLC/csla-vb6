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
  Dim objDS As DataSpace
  Dim objPersist As Object
  
  Set objDS = New RDS.DataSpace
  Set objPersist = objDS.CreateObject("TaskServer.ClientPersist", _
    "http://ss3test")
  Fetch = objPersist.Fetch(ID)
  Set objPersist = Nothing
  Set objDS = Nothing
End Function

Public Function Save(Buffer As String) As String
  Dim objDS As DataSpace
  Dim objPersist As Object
  
  Set objDS = New RDS.DataSpace
  Set objPersist = objDS.CreateObject("TaskServer.ClientPersist", _
    "http://myserver")
  Save = objPersist.Save(Buffer)
  Set objPersist = Nothing
  Set objDS = Nothing
End Function

Public Sub DeleteObject(ID As Long)
  Dim objDS As DataSpace
  Dim objPersist As Object
  
  Set objDS = New RDS.DataSpace
  Set objPersist = objDS.CreateObject("TaskServer.ClientPersist", _
    "http://myserver")
  objPersist.DeleteObject ID
  Set objPersist = Nothing
  Set objDS = Nothing
End Sub

