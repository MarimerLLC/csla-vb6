VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideosPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal Title As String, ByVal Studio As String) As String

  Dim rsVideo As Recordset
  Dim strSQL As String
  Dim strWHERE As String
' Dim objDisplay As VideoDisplay
  Dim udtData As VideoDisplayData
  Dim udtProps As VideoDisplayProps
  
  strSQL = "SELECT VideoID, TITLE,RELEASEDATE FROM Video"
  If Len(Title) > 0 Then _
    strWHERE = "TITLE LIKE '" & Title & "%' "

  If Len(Studio) > 0 Then
 
    If Len(strWHERE) > 0 Then
      strWHERE = strWHERE & " AND STUDIO LIKE '" & Studio & "%'"

    Else
      strWHERE = "STUDIO LIKE '" & Studio & "%'"

    End If

  End If

  If Len(strWHERE) > 0 Then _
    strSQL = strSQL & " WHERE " & strWHERE
  Set rsVideo = New Recordset
  rsVideo.Open strSQL, cnVideo
  Set mobjBuffer = New Buffer
  mobjBuffer.Initialize Len(udtData.Buffer), 100

  Do While Not rsVideo.EOF
     With udtProps
        .VideoID = rsVideo("VideoID")
        .Title = rsVideo("Title")
        .ReleaseDate = rsVideo("ReleaseDate")
     End With
     LSet udtData = udtProps
     mobjBuffer.Add udtData.Buffer
     rsVideo.MoveNext
  Loop

  rsVideo.Close
  Set rsVideo = Nothing
  Fetch = mobjBuffer.GetState
  
End Function



