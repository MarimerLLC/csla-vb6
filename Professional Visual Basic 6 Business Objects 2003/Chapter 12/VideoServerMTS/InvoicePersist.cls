VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "InvoicePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
Private mudtProps As InvoiceProps

Private Function GetState() As String

  Dim udtData As InvoiceData

  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As InvoiceData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  On Error GoTo errh
  SetState Buffer
  strSQL = "SELECT * FROM Invoice WHERE InvoiceID=" & mudtProps.InvoiceID
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsInvoice.AddNew

  With rsInvoice
    .Fields("CustomerID") = mudtProps.CustomerID
    .Fields("Subtotal") = mudtProps.SubTotal
    .Fields("Tax") = mudtProps.Tax
    .Fields("Total") = mudtProps.Total
    .Update
    If mudtProps.IsNew Then
        .Bookmark = .Bookmark
        mudtProps.InvoiceID = .Fields("InvoiceID")
    End If
    
    .Close
  End With

  strSQL = "SELECT NAME,PHONE FROM Customer " & _
    "WHERE CustomerID=" & mudtProps.CustomerID
  rsInvoice.Open strSQL, DB_CONNECT
  mudtProps.CustomerName = rsInvoice("Name")
  mudtProps.CustomerPhone = rsInvoice("Phone")
  rsInvoice.Close
  Set rsInvoice = Nothing
  Save = GetState
  mobjContext.SetComplete
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal InvoiceID As Long)

  Dim cnVideo As Connection

  On Error GoTo errh
  cnVideo.Open DB_CONNECT

  cnVideo.Execute "DELETE FROM Invoice WHERE InvoiceID=" & _
    InvoiceID
  cnVideo.Close
  Set cnVideo = Nothing
  mobjContext.SetComplete
  Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Sub

Public Function Fetch(ByVal InvoiceID As Long) As String

  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  On Error GoTo errh
  strSQL = "SELECT Invoice.*,Customer.Name,Customer.Phone " & _
    "FROM Invoice INNER JOIN Customer ON " & _
    "Invoice.CustomerID = Customer.CustomerID " & _
    "WHERE InvoiceID=" & CStr(InvoiceID)
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, DB_CONNECT
  
  With rsInvoice
    mudtProps.InvoiceID = .Fields("InvoiceID")
    mudtProps.CustomerID = .Fields("CustomerID")
    mudtProps.CustomerName = .Fields("Name")
    mudtProps.CustomerPhone = .Fields("Phone")
    .Close
  End With

  Set rsInvoice = Nothing
  Fetch = GetState
  mobjContext.SetComplete
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

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

