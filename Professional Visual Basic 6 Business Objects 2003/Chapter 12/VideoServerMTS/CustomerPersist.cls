VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As CustomerProps

Private Function GetState() As String

  Dim udtData As CustomerData

  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As CustomerData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

  Dim rsCustomer As Recordset
  Dim strSQL As String
  
  SetState Buffer
  
  Set rsCustomer = New Recordset
  strSQL = "SELECT * FROM CUSTOMER WHERE CUSTOMERID=" & _
    mudtProps.CustomerID
  rsCustomer.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsCustomer.AddNew
  
  With rsCustomer
    .Fields("Name") = mudtProps.Name
    .Fields("Address1") = mudtProps.Address1
    .Fields("Address2") = mudtProps.Address2
    .Fields("City") = mudtProps.City
    .Fields("State") = mudtProps.State
    .Fields("ZipCode") = mudtProps.ZipCode
    .Fields("Phone") = mudtProps.Phone
    .Update

    If mudtProps.IsNew Then
        .Bookmark = .Bookmark
        mudtProps.CustomerID = .Fields("CustomerID")
    End If
    .Close
  End With

  Set rsCustomer = Nothing
  Save = GetState
  
End Function

Public Sub DeleteObject(ByVal CustomerID As Long)

  Dim cnVideo As Connection

  cnVideo.Open DB_CONNECT
  cnVideo.Execute "DELETE FROM CUSTOMER WHERE CUSTOMERID=" & _
    CustomerID
  cnVideo.Close
  Set cnVideo = Nothing
  
End Sub

Public Function Fetch(ByVal CustomerID As Long) As String

  Dim rsCustomer As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM CUSTOMER WHERE CUSTOMERID=" & _
    CustomerID
  Set rsCustomer = New Recordset
  rsCustomer.Open strSQL, DB_CONNECT

  With rsCustomer
    mudtProps.CustomerID = .Fields("CustomerID")
    mudtProps.Name = .Fields("Name")
    mudtProps.Address1 = .Fields("Address1")
    mudtProps.Address2 = .Fields("Address2")
    mudtProps.City = .Fields("City")
    mudtProps.State = .Fields("State")
    mudtProps.ZipCode = .Fields("ZipCode")
    mudtProps.Phone = .Fields("Phone")
    .Close
  End With

  Set rsCustomer = Nothing
  Fetch = GetState
  
End Function
