VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function FetchTapes(InvoiceID As Long) As String

  Dim rsItems As Recordset
  Dim strSQL As String
  Dim objBuffer As Buffer
  Dim udtProps As InvoiceTapeProps
  Dim objPersist As InvoiceTapePersist

  strSQL = "SELECT ITEMID FROM INVOICETAPE " & _
    "WHERE INVOICEID=" & InvoiceID
  Set rsItems = New Recordset
  rsItems.Open strSQL, cnVideo
  Set objBuffer = New Buffer
  objBuffer.Initialize Len(udtProps), 10
  Set objPersist = New InvoiceTapePersist

  Do While Not rsItems.EOF
    objBuffer.Add objPersist.Fetch(rsItems("ItemID"))
  Loop

  rsItems.Close
  Set rsItems = Nothing
  FetchTapes = objBuffer.GetState
  
End Function


Public Function FetchFees(CustomerID As Long) As String

  Dim rsItems As Recordset
  Dim strSQL As String
  Dim objBuffer As Buffer
  Dim udtProps As FeeProps
  Dim objPersist As InvoiceFeePersist

  strSQL = "SELECT FEEID FROM LATEFEE " & _
    "WHERE DELETED=FALSE AND PAID=FALSE AND " & _
    "CUSTOMERID=" & CustomerID
  Set rsItems = New Recordset
  rsItems.Open strSQL, cnVideo
  Set objBuffer = New Buffer
  objBuffer.Initialize Len(udtProps), 10
  Set objPersist = New InvoiceFeePersist

  Do While Not rsItems.EOF
    objBuffer.Add objPersist.Fetch(rsItems("FeeID"))
    rsItems.MoveNext
  Loop

  rsItems.Close
  Set rsItems = Nothing
  FetchFees = objBuffer.GetState
  
End Function

Public Function SaveTapes(InvoiceID As Long, Buffer As String) As String

  Dim objBuffer As Buffer
  Dim objUpdatedTapes As Buffer
  Dim objPersist As InvoiceTapePersist
  Dim lngIndex As Long
  Dim strBuffer As String

  Set objBuffer = New Buffer
  Set objUpdatedTapes = New Buffer
  objBuffer.SetState Buffer
  objUpdatedTapes.Initialize objBuffer.Length, objBuffer.Count
  Set objPersist = New InvoiceTapePersist

  For lngIndex = 1 To objBuffer.Count
    With objPersist
      strBuffer = .Save(InvoiceID, objBuffer.Item(lngIndex))
      If Len(strBuffer) > 0 Then _
        objUpdatedTapes.Add strBuffer
    End With
  Next

  Set objPersist = Nothing
  SaveTapes = objUpdatedTapes.GetState

End Function


Public Function SaveFees(Buffer As String) As String

  Dim objBuffer As Buffer
  Dim objUpdatedFees As Buffer
  Dim objPersist As InvoiceFeePersist
  Dim lngIndex As Long
  Dim strBuffer As String

  Set objBuffer = New Buffer
  Set objUpdatedFees = New Buffer
  objBuffer.SetState Buffer
  objUpdatedFees.Initialize objBuffer.Length, objBuffer.Count
  Set objPersist = New InvoiceFeePersist

  For lngIndex = 1 To objBuffer.Count
    With objPersist
      strBuffer = .Save(objBuffer.Item(lngIndex))
      If Len(strBuffer) > 0 Then _
        objUpdatedFees.Add strBuffer
    End With
  Next

  Set objPersist = Nothing
  SaveFees = objUpdatedFees.GetState

End Function




