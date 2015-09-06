VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cluster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjMM As MemoryMap

Private Sub Class_Initialize()
  Randomize
  Set mobjMM = New MemoryMap
  mobjMM.Initialize "WroxCluster", 32768
End Sub

Private Sub Class_Terminate()
  Set mobjMM = Nothing
End Sub

Public Function GetServerName() As String
  Dim lngMaxValue As Long
  Dim lngMaxIndex As Long
  Dim lngIndex As Long
  Dim lngStatus As Long
  Dim objPB As PropertyBag
  
  Set objPB = mobjMM.GetData
  lngMaxValue = 0
  For lngIndex = 1 To objPB.ReadProperty("ServerCount")
    lngStatus = objPB.ReadProperty("ServerStatus" & CStr(lngIndex))
    If lngStatus > 0 Then
      lngStatus = Rnd * lngStatus + 1
      If lngStatus > lngMaxValue Then
        lngMaxValue = lngStatus
        lngMaxIndex = lngIndex
      End If
    End If
  Next
  
  If lngMaxValue = 0 Then
    GetServerName = ""
  Else
    GetServerName = objPB.ReadProperty("ServerName" & CStr(lngMaxIndex))
  End If
End Function

