VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TapePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As TapeProps

Private Sub SetState(Buffer As String)

  Dim udtData As TapeData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Private Function GetState() As String

  Dim udtData As TapeData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal TapeID As Long) As String

  Dim objPersist As TapePersist
  
  Dim rsTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Tape.*,Video.Title " & _
    "FROM Tape INNER JOIN Video ON " & _
    "Tape.VideoID = Video.VideoID " & _
    "WHERE TapeID=" & TapeID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo

  With rsTape
    mudtProps.TapeID = .Fields("TapeID")
    mudtProps.VideoID = .Fields("VideoID")
    mudtProps.Title = .Fields("Title")
    mudtProps.CheckedOut = .Fields("CheckedOut")
    mudtProps.DateAcquired = IIf(IsNull(.Fields("DateAcquired")), 0, _
       .Fields("DateAcquired"))
    mudtProps.DateDue = IIf(IsNull(.Fields("DateDue")), 0, .Fields("DateDue"))
    mudtProps.InvoiceID = .Fields("InvoiceID")
    rsTape.Close
  End With

  Set rsTape = Nothing
  Fetch = GetState
  
End Function

Public Function Save(ByVal Buffer As String) As String

  Dim rsTape As Recordset
  Dim rsVideo As Recordset
  Dim strSQL As String
  
  SetState Buffer
  strSQL = "SELECT * FROM Tape WHERE TapeID=" & mudtProps.TapeID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo, , adLockOptimistic
  If mudtProps.IsNew Then rsTape.AddNew
  
  With rsTape
    .Fields("VideoID") = mudtProps.VideoID
    .Fields("CheckedOut") = mudtProps.CheckedOut
    .Fields("DateAcquired") = mudtProps.DateAcquired
    .Fields("DateDue") = mudtProps.DateDue
    .Fields("InvoiceID") = mudtProps.InvoiceID
    .Update
    If mudtProps.IsNew Then mudtProps.TapeID = .Fields("TapeID")
    .Close
  End With

  Set rsTape = Nothing
  If mudtProps.IsNew Then
    strSQL = "SELECT TITLE FROM Video WHERE VideoID=" & _
      mudtProps.VideoID
    Set rsVideo = New Recordset
    rsVideo.Open strSQL, cnVideo
    mudtProps.Title = rsVideo("Title")
    rsVideo.Close
    Set rsVideo = Nothing
  End If
  Save = GetState
  
  If mudtProps.LateFee Then
    ' add the late fee to the customer
    SaveLateFee
  End If

End Function

Private Sub SaveLateFee()

  Dim rsLateFee As Recordset
  Dim rsInvoice As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CUSTOMERID FROM INVOICE WHERE INVOICEID=" & _
    mudtProps.InvoiceID
  Set rsInvoice = New Recordset
  rsInvoice.Open strSQL, cnVideo
  strSQL = "LateFee"
  Set rsLateFee = New Recordset
  rsLateFee.Open strSQL, cnVideo, , adLockOptimistic
  rsLateFee.AddNew

  With rsLateFee
    .Fields("CustomerID") = rsInvoice("CustomerID")
    .Fields("TapeID") = mudtProps.TapeID
    .Fields("FeeDate") = Now
    .Fields("DaysOver") = DateDiff("d", mudtProps.DateDue, Now)
    .Fields("Paid") = False
    .Update
    .Close
  End With

  rsInvoice.Close
  Set rsInvoice = Nothing
  Set rsLateFee = Nothing

End Sub


Public Sub DeleteObject(ByVal TapeID As Long)

  cnVideo.Execute "DELETE FROM Tape WHERE TapeID=" & TapeID

End Sub

Friend Function SaveChild(VideoID As Long, Buffer As String) As String

  SetState Buffer

  With mudtProps
    If .IsDeleted Then
      DeleteObject .TapeID
      SaveChild = ""

    ElseIf .IsDirty Or .IsNew Then
      .VideoID = VideoID
      SaveChild = Save(GetState)

    Else
      SaveChild = Buffer

    End If

  End With

End Function

Public Sub CheckOut(TapeID As Long, InvoiceID As Long)

  Dim rsTape As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Tape WHERE TapeID=" & TapeID
  Set rsTape = New Recordset
  rsTape.Open strSQL, cnVideo, , adLockOptimistic

  With rsTape
    .Fields("CheckedOut") = True
    .Fields("DateDue") = DateAdd("d", 2, Now)
    .Fields("InvoiceID") = InvoiceID
    .Update
    .Close
  End With

  Set rsTape = Nothing

End Sub


