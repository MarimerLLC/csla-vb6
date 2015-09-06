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

  ' load data from database
  Fetch Name, Phone
  
End Sub

Private Sub Fetch(Name As String, Phone As String)

  Dim rsCustomer As Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As CustomerDisplay
  
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
  rsCustomer.Open strSQL, cnVideo

  Do While Not rsCustomer.EOF
    Set objDisplay = New CustomerDisplay

    With objDisplay
      .CustomerID = rsCustomer("CustomerID")
      .Name = rsCustomer("Name")
      .Phone = rsCustomer("Phone")
      mcolDisplay.Add objDisplay
      Set objDisplay = Nothing
      rsCustomer.MoveNext
    End With

  Loop

  rsCustomer.Close
  Set rsCustomer = Nothing

End Sub


