VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TapesPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal VideoID As Long) As String

  Dim strSQL As String
  Dim rsTape As Recordset
  Dim objBuffer As Buffer
  Dim objPersist As TapePersist
  Dim udtProps As TapeProps

  On Error GoTo errh
  Set objBuffer = New Buffer
  strSQL = "SELECT TapeID FROM Tape_ WHERE VideoID=" & VideoID
  Set rsTape = New Recordset
  rsTape.Open strSQL, DB_CONNECT
  Set objPersist = mobjContext.CreateInstance("VideoServerMTS.TapePersist")

  With objPersist
    objBuffer.Initialize Len(udtProps), 20

    Do While Not rsTape.EOF
      objBuffer.Add .Fetch(rsTape("TapeID"))
      rsTape.MoveNext
    Loop

  End With

  Set objPersist = Nothing
  rsTape.Close
  Set rsTape = Nothing
  Fetch = objBuffer.GetState
  mobjContext.SetComplete
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function Save(ByVal VideoID As Long, ByVal Buffer As String) As String

  Dim objPersist As TapePersist
  Dim objBuffer As Buffer
  Dim objUpdatedTapes As Buffer
  Dim lngIndex As Long
  Dim strBuffer As String

  On Error GoTo errh
  Set objBuffer = New Buffer
  objBuffer.SetState Buffer
  Set objUpdatedTapes = New Buffer
  objUpdatedTapes.Initialize objBuffer.Length, objBuffer.Count
  Set objPersist = mobjContext.CreateInstance("VideoServerMTS.TapePersist")

  For lngIndex = 1 To objBuffer.Count
    With objPersist
      strBuffer = .SaveChild(VideoID, objBuffer.Item(lngIndex))
      If Len(strBuffer) > 0 Then _
        objUpdatedTapes.Add strBuffer
    End With

  Next

  Set objPersist = Nothing
  Save = objUpdatedTapes.GetState
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

