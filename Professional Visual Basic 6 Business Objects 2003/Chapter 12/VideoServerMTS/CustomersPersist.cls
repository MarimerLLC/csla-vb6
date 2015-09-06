VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomersPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal Name As String, ByVal Phone As String) As String

  Dim rsCustomer As Recordset
  Dim strSQL As String
  Dim strWHERE As String
' Dim objDisplay As CustomerDisplay
  Dim udtData As CustDisplayData
  Dim udtProps As CustDisplayProps
  
  strSQL = "SELECT CUSTOMERID, NAME, PHONE FROM CUSTOMER"
  If Len(Name) > 0 Then _
    strWHERE = "NAME LIKE '" & Name & "%' "

  If Len(Phone) > 0 Then

    If Len(strWHERE) > 0 Then
      strWHERE = strWHERE & " AND PHONE LIKE '" & Phone & "%'"

    Else
      strWHERE = "PHONE LIKE '" & Phone & "%'"

    End If

  End If

  If Len(strWHERE) > 0 Then _
    strSQL = strSQL & " WHERE " & strWHERE
  Set rsCustomer = New Recordset
  rsCustomer.Open strSQL, DB_CONNECT
  Set mobjBuffer = New Buffer
  mobjBuffer.Initialize Len(udtData.Buffer), 100

  Do While Not rsCustomer.EOF
    With udtProps
      .CustomerID = rsCustomer("CustomerID")
      .Name = rsCustomer("Name")
      .Phone = rsCustomer("Phone")
    End With
    LSet udtData = udtProps
    mobjBuffer.Add udtData.Buffer
    rsCustomer.MoveNext
  Loop

  rsCustomer.Close
  Set rsCustomer = Nothing
  Fetch = mobjBuffer.GetState
  
End Function



