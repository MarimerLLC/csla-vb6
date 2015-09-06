VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolDisplay As Collection

Private Sub Class_Initialize()

  Set mcolDisplay = New Collection

End Sub

Public Function Count() As Long

  Count = mcolDisplay.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = mcolDisplay.[_NewEnum]

End Function

Public Function Item(ByVal Index As Variant) As CustomerDisplay
Attribute Item.VB_UserMemId = 0

  Set Item = mcolDisplay(Index)
  
End Function

Public Sub Load(Optional Name As String, Optional Phone As String)

  Dim objPersist As CustomersPersist
  
  Set objPersist = CreateObject("VideoServer.CustomersPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(Name, Phone)
  Set objPersist = Nothing
  
End Sub

Private Sub SetState(Buffer As String)

  Dim objBuffer As Buffer
  Dim objDisplay As CustomerDisplay
  Dim lngIndex As Long
  Dim udtData As CustDisplayData
  Dim udtProps As CustDisplayProps
  
  Set objBuffer = New Buffer
  With objBuffer
    .SetState Buffer
    For lngIndex = 1 To .Count
      Set objDisplay = New CustomerDisplay
      udtData.Buffer = .Item(lngIndex)
      LSet udtProps = udtData

      With objDisplay
        .CustomerID = udtProps.CustomerID
        .Name = Trim$(udtProps.Name)
        .Phone = Trim$(udtProps.Phone)
        mcolDisplay.Add objDisplay
        Set objDisplay = Nothing
      End With

    Next

  End With

  Set objBuffer = Nothing

End Sub


