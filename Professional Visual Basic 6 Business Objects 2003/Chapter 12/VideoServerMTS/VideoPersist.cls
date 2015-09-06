VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "VideoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
Private mudtProps As VideoProps

Public Function Save(ByVal Buffer As String, TapeBuffer As String) As String

  Dim rsVideo As Recordset
  Dim strSQL As String
  Dim objPersist As TapesPersist

  On Error GoTo errh
  SetState Buffer
  
  Set rsVideo = New Recordset
  strSQL = "SELECT * FROM Video WHERE VideoID=" & mudtProps.VideoID
  rsVideo.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsVideo.AddNew
  
  With rsVideo
    .Fields("Title") = mudtProps.Title
    .Fields("ReleaseDate") = mudtProps.ReleaseDate
    .Fields("Studio") = mudtProps.Studio
    .Fields("Rating") = mudtProps.Rating
    .Fields("Category") = mudtProps.Category
    .Update
    If mudtProps.IsNew Then
      .Bookmark = .Bookmark
      mudtProps.VideoID = .Fields("VideoID")
    End If
    .Close
  End With

  Set rsVideo = Nothing
  Save = GetState
  Set objPersist = mobjContext.CreateInstance("VideoServerMTS.TapesPersist")
  TapeBuffer = objPersist.Save(mudtProps.VideoID, TapeBuffer)
  Set objPersist = Nothing
  mobjContext.SetComplete
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal VideoID As Long)

  Dim cnVideo As Connection

  On Error GoTo errh
  'Set cnVideo = New Connection
  cnVideo.Open DB_CONNECT
  cnVideo.Execute "DELETE FROM Video WHERE VideoID=" & VideoID
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

Public Function Fetch(ByVal VideoID As Long) As String

  Dim rsVideo As Recordset
  Dim strSQL As String
  
  On Error GoTo errh
  strSQL = "SELECT * FROM Video WHERE VideoID=" & VideoID
  Set rsVideo = New Recordset
  rsVideo.Open strSQL, DB_CONNECT

  With rsVideo
    mudtProps.VideoID = .Fields("VideoID")
    mudtProps.Title = .Fields("Title")
    mudtProps.ReleaseDate = IIf(IsNull(.Fields("ReleaseDate")), 0, _
      .Fields("ReleaseDate"))
    mudtProps.Studio = .Fields("Studio")
    mudtProps.Rating = .Fields("Rating")
    mudtProps.Category = .Fields("Category")
    rsVideo.Close
  End With

  Set rsVideo = Nothing
  Fetch = GetState
    mobjContext.SetComplete
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As VideoData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData

End Sub

Private Function GetState() As String

  Dim udtData As VideoData
    
  LSet udtData = mudtProps
  GetState = udtData.Buffer

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
