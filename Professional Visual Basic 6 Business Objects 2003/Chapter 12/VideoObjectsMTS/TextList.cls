VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolList As Collection
Private mcolKeys As Collection

Private Sub Class_Initialize()

  Set mcolList = New Collection
  Set mcolKeys = New Collection

End Sub

Public Function Item(ByVal Index As Variant) As String
Attribute Item.VB_UserMemId = 0

  Item = mcolList.Item(Index)

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = mcolList.[_NewEnum]

End Function

Public Function Key(ByVal Text As String) As String

  Key = mcolKeys.Item(Text)

End Function

Friend Sub Load(ListName As String)

  Dim objPersist As TextListPersist
  
  Set objPersist = CreateObject("VideoServerMTS.TextListPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(ListName)
  Set objPersist = Nothing
  
End Sub

Private Sub SetState(Buffer As String)

  Dim objBuffer As Buffer
  Dim lngIndex As Long
  Dim udtProps As TextListProps
  Dim udtData As TextListData

  Set objBuffer = New Buffer
  objBuffer.SetState Buffer

  With objBuffer
    For lngIndex = 1 To objBuffer.Count
      udtData.Buffer = objBuffer.Item(lngIndex)
      LSet udtProps = udtData

      With udtProps
        mcolList.Add Format$(.Item), Format$(.Key)
        mcolKeys.Add Format$(.Key), Format$(.Item)
      End With

    Next

  End With

End Sub





