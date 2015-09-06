VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As VideoProps

Public Function Save(ByVal Buffer As String) As String

  Dim rsVideo As Recordset
  Dim strSQL As String

  SetState Buffer
  
  Set rsVideo = New Recordset
  strSQL = "SELECT * FROM Video WHERE VideoID=" & mudtProps.VideoID
  rsVideo.Open strSQL, cnVideo, , adLockOptimistic
  If mudtProps.IsNew Then rsVideo.AddNew
  
  With rsVideo
    .Fields("Title") = mudtProps.Title
    .Fields("ReleaseDate") = mudtProps.ReleaseDate
    .Fields("Studio") = mudtProps.Studio
    .Fields("Rating") = mudtProps.Rating
    .Fields("Category") = mudtProps.Category
    .Update
    If mudtProps.IsNew Then mudtProps.VideoID = .Fields("VideoID")
    .Close
  End With

  Set rsVideo = Nothing
  Save = GetState
  
End Function

Public Sub DeleteObject(ByVal VideoID As Long)

  cnVideo.Execute "DELETE FROM Video WHERE VideoID=" & VideoID

End Sub

Public Function Fetch(ByVal VideoID As Long) As String

  Dim rsVideo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM Video WHERE VideoID=" & VideoID
  Set rsVideo = New Recordset
  rsVideo.Open strSQL, cnVideo

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

