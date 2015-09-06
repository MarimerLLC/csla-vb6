VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Dim objOSP As Object
  Dim objBO As Object
  Dim strObject As String
  Dim arFields() As String
  
  ParseQuery DataMember, strObject, arFields
  Set objBO = gcolObjects.Item(strObject)
  If IsCollection(objBO) Then
    Set objOSP = New CollectionProvider
  Else
    Set objOSP = New SimpleObjectProvider
  End If
  objOSP.LoadData objBO, arFields
  Set Data = objOSP
  Set objOSP = Nothing
End Sub

Private Function IsCollection(BusinessObject As Object) As Boolean
  Dim objTLIApplication As TLI.TLIApplication
  Dim objInterfaceInfo As TLI.InterfaceInfo
  Dim lngIndex As Long
  Dim flgCount As Boolean
  Dim flgItem As Boolean
  Dim flgNewEnum As Boolean
  
  Set objTLIApplication = New TLI.TLIApplication
  Set objInterfaceInfo = objTLIApplication.InterfaceInfoFromObject(BusinessObject)
  With objInterfaceInfo.Members
    For lngIndex = 1 To .Count
      If .Item(lngIndex).Name = "Count" Then
        flgCount = True
      ElseIf .Item(lngIndex).Name = "Item" Then
        flgItem = True
      ElseIf .Item(lngIndex).Name = "_NewEnum" Then
        flgNewEnum = True
      ElseIf .Item(lngIndex).Name = "NewEnum" Then
        flgNewEnum = True
      End If
    Next
  End With
  IsCollection = flgCount And flgItem And flgNewEnum
End Function

Private Sub ParseQuery(ByVal Query As String, _
                       Object As String, Fields() As String)
  Dim lngDelimiter As Long
  
  lngDelimiter = InStr(1, Query, ":")
  If lngDelimiter > 0 Then
    Object = Left$(Query, lngDelimiter - 1)
    Fields = Split(Mid$(Query, lngDelimiter + 1), ",")
    For lngDelimiter = LBound(Fields) To UBound(Fields)
      Fields(lngDelimiter) = Trim$(Fields(lngDelimiter))
    Next
  Else
    Object = Query
    ReDim Fields(0)
    Fields(0) = "*"
  End If
End Sub


