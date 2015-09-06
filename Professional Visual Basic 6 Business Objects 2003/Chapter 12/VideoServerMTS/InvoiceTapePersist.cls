VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "InvoiceTapePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
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

  Dim cnVideo As Connection

  cnVideo.Open DB_CONNECT

  cnVideo.Execute "DELETE FROM InvoiceTape WHERE ItemID=" & _
    CStr(ItemID)
  cnVideo.Close
  Set cnVideo = Nothing
  mobjContext.SetComplete

End Sub

Friend Function Fetch(ItemID As Long) As String

  Dim rsInvoiceTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT InvoiceTape.*, Video.Title " & _
    "FROM (InvoiceTape INNER JOIN Tape_ ON " & _
    "InvoiceTape.TapeID = Tape_.TapeID) " & _
    "INNER JOIN Video ON Tape_.VideoID = Video.VideoID " & _
    "WHERE InvoiceTape.ItemID=" & ItemID
  Set rsInvoiceTape = New Recordset
  rsInvoiceTape.Open strSQL, DB_CONNECT

  With rsInvoiceTape
    mudtProps.ItemID = .Fields("ItemID")
    mudtProps.TapeID = .Fields("TapeID")
    mudtProps.Title = .Fields("Title")
    mudtProps.Price = .Fields("Price")
    .Close
  End With

  Set rsInvoiceTape = Nothing
  Fetch = GetState
    mobjContext.SetComplete

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
  rsInvoiceTape.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsInvoiceTape.AddNew

  With rsInvoiceTape
    .Fields("InvoiceID") = InvoiceID
    .Fields("TapeID") = mudtProps.TapeID
    .Fields("Price") = mudtProps.Price
    .Update
    If mudtProps.IsNew Then
        .Bookmark = .Bookmark
        mudtProps.ItemID = .Fields("ItemID")
    End If
    
    .Close
  End With
  
  Set objPersist = New TapePersist
  objPersist.CheckOut mudtProps.TapeID, InvoiceID
  Set objPersist = Nothing

  If mudtProps.IsNew Then
    strSQL = "SELECT Video.Title " & _
      "FROM Tape_ INNER JOIN Video ON " & _
      "Tape_.VideoID = Video.VideoID " & _
      "WHERE Tape_.TapeID=" & mudtProps.TapeID
    rsInvoiceTape.Open strSQL, DB_CONNECT
    mudtProps.Title = rsInvoiceTape("Title")
    rsInvoiceTape.Close
  End If
  Set rsInvoiceTape = Nothing
  Save = GetState
    mobjContext.SetComplete

End Function

Private Function ObjectControl_CanBePooled() As Boolean

  ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

  Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

  Set mobjContext = Nothing

End Sub


