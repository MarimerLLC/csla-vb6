VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "InvoiceFeePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
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
  
  On Error GoTo errh
  strSQL = "SELECT LateFee.*, Video.Title " & _
    "FROM (LateFee INNER JOIN Tape_ ON " & _
    "LateFee.TapeID = Tape_.TapeID) " & _
    "INNER JOIN Video ON Tape_.VideoID = Video.VideoID " & _
    "WHERE LateFee.FeeID=" & Format$(FeeID)
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, DB_CONNECT

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
    mobjContext.SetComplete
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function


Private Sub DeleteObject(FeeID As Long)

  Dim rsLateFee As Recordset
  Dim strSQL As String
  Dim cnVideo As Connection

  On Error GoTo errh
  cnVideo.Open DB_CONNECT

  strSQL = "SELECT DELETED FROM LateFee WHERE FeeID=" & FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo

  With rsLateFee
    .Fields("Deleted") = True
    .Update
    .Close
  End With
  
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

Friend Function Save(Buffer As String) As String

  Dim rsLateFee As Recordset
  Dim strSQL As String
  
  On Error GoTo errh
  SetState Buffer
  If mudtProps.IsDeleted Then
    DeleteObject mudtProps.FeeID
    Exit Function
  End If
  
  strSQL = "SELECT * FROM LateFee WHERE FeeID=" & _
    mudtProps.FeeID
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsLateFee.AddNew

  With rsLateFee
    .Fields("Paid") = mudtProps.Paid
    .Fields("PaidDate") = mudtProps.PaidDate
    .Update
    If mudtProps.IsNew Then
        .Bookmark = .Bookmark
        mudtProps.FeeID = .Fields("FeeID")
    End If
    .Close
  End With

  Set rsLateFee = Nothing
  Save = GetState
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

