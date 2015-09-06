VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceTapePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As InvoiceTapeProps

Private Sub SetState(Buffer As String)

  Dim udtData As InvoiceTapeData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Private Function GetState() As String

  Dim udtData As InvoiceTapeData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Private Sub DeleteObject(ItemID As Long)

  cnVideo.Execute "DELETE FROM InvoiceTape WHERE ItemID=" & _
    CStr(ItemID)

End Sub

Friend Function Fetch(ItemID As Long) As String

  Dim rsInvoiceTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT InvoiceTape.*, Video.Title " & _
    "FROM (InvoiceTape INNER JOIN Tape ON " & _
    "InvoiceTape.TapeID = Tape.TapeID) " & _
    "INNER JOIN Video ON Tape.VideoID = Video.VideoID " & _
    "WHERE InvoiceTape.ItemID=" & ItemID
  Set rsInvoiceTape = New Recordset
  rsInvoiceTape.Open strSQL, cnVideo

  With rsInvoiceTape
    mudtProps.ItemID = .Fields("ItemID")
    mudtProps.TapeID = .Fields("TapeID")
    mudtProps.Title = .Fields("Title")
    mudtProps.Price = .Fields("Price")
    .Close
  End With

  Set rsInvoiceTape = Nothing
  Fetch = GetState
  
End Function

Friend Function Save(InvoiceID As Long, Buffer As String) As String

  Dim rsInvoiceTape As Recordset
  Dim strSQL As String
  Dim objPersist As TapePersist
  
  SetState Buffer
  If mudtProps.IsDeleted Then
    DeleteObject mudtProps.ItemID
    Exit Function
  End If
  
  strSQL = "SELECT * FROM InvoiceTape WHERE ItemID=" & _
    CStr(mudtProps.ItemID)
  Set rsInvoiceTape = New Recordset
  rsInvoiceTape.Open strSQL, cnVideo, , adLockOptimistic
  If mudtProps.IsNew Then rsInvoiceTape.AddNew

  With rsInvoiceTape
    .Fields("InvoiceID") = InvoiceID
    .Fields("TapeID") = mudtProps.TapeID
    .Fields("Price") = mudtProps.Price
    .Update
    If mudtProps.IsNew Then mudtProps.ItemID = .Fields("ItemID")
    .Close
  End With
  
  Set objPersist = New TapePersist
  objPersist.CheckOut mudtProps.TapeID, InvoiceID
  Set objPersist = Nothing

  If mudtProps.IsNew Then
    strSQL = "SELECT Video.Title " & _
      "FROM Tape INNER JOIN Video ON " & _
      "Tape.VideoID = Video.VideoID " & _
      "WHERE Tape.TapeID=" & mudtProps.TapeID
    rsInvoiceTape.Open strSQL, cnVideo
    mudtProps.Title = rsInvoiceTape("Title")
    rsInvoiceTape.Close
  End If
  Set rsInvoiceTape = Nothing
  Save = GetState
  
End Function



