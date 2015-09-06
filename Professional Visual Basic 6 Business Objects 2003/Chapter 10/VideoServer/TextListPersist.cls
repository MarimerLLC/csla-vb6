VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextListPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal ListName As String) As String

  Dim strSQL As String
  Dim rsList As Recordset
  Dim udtProps As TextListProps
  Dim udtData As TextListData

  Set mobjBuffer = New Buffer
  mobjBuffer.Initialize Len(udtData), 50
  
  Set rsList = New Recordset
  rsList.Open ListName, cnVideo

  Do While Not rsList.EOF
    With udtProps
      .Item = rsList(1)
      .Key = rsList(0)
    End With

    LSet udtData = udtProps
    mobjBuffer.Add udtData.Buffer
    rsList.MoveNext
  Loop

  rsList.Close
  Set rsList = Nothing
  Fetch = mobjBuffer.GetState

End Function

