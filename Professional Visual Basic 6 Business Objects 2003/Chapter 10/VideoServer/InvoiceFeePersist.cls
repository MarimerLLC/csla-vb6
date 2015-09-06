VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceFeePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As FeeProps

Private Sub SetState(Buffer As String)

  Dim udtData As FeeData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Private Function GetState() As String

  Dim udtData As FeeData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Friend Function Fetch(FeeID As Long) As String

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT LateFee.*, Video.Title " & _
    "FROM (LateFee INNER JOIN Tape ON " & _
    "LateFee.TapeID = Tape.TapeID) " & _
    "INNER JOIN Video ON Tape.VideoID = Video.VideoID " & _
    "WHERE LateFee.FeeID=" & Format$(FeeID)
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo

  With rsLateFee
    mudtProps.FeeID = .Fields("FeeID")
    mudtProps.VideoTitle = .Fields("Title")
    mudtProps.EnteredDate = .Fields("FeeDate")
    mudtProps.DaysOver = .Fields("DaysOver")
    mudtProps.Paid = .Fields("Paid")
    If Not IsNull(.Fields("PaidDate")) Then _
      mudtProps.PaidDate = .Fields("PaidDate")
    .Close
  End With

  Set rsLateFee = Nothing
  Fetch = GetState
  
End Function


Private Sub DeleteObject(FeeID As Long)

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT DELETED FROM LateFee WHERE FeeID=" & FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo

  With rsLateFee
    .Fields("Deleted") = True
    .Update
    .Close
  End With

End Sub

Friend Function Save(Buffer As String) As String

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  SetState Buffer
  If mudtProps.IsDeleted Then
    DeleteObject mudtProps.FeeID
    Exit Function
  End If
  
  strSQL = "SELECT * FROM LateFee WHERE FeeID=" & _
    mudtProps.FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo, , adLockOptimistic
  If mudtProps.IsNew Then rsLateFee.AddNew

  With rsLateFee
    .Fields("Paid") = mudtProps.Paid
    .Fields("PaidDate") = mudtProps.PaidDate
    .Update
    If mudtProps.IsNew Then mudtProps.FeeID = .Fields("FeeID")
    .Close
  End With

  Set rsLateFee = Nothing
  Save = GetState
  
End Function




