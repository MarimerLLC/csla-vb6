VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Videos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolDisplay As Collection

Private Sub Class_Initialize()

  Set mcolDisplay = New Collection

End Sub

Public Sub Load(Optional Title As String, Optional Studio As String)

  ' load data from database
  Fetch Title, Studio

End Sub

Public Function Item(ByVal Index As Variant) As VideoDisplay
Attribute Item.VB_UserMemId = 0

  Set Item = mcolDisplay(Index)
  
End Function

Public Function Count() As Long

  Count = mcolDisplay.Count

End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

  Set NewEnum = mcolDisplay.[_NewEnum]

End Function

Private Sub Fetch(Title As String, Studio As String)

  Dim rsVideo As Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As VideoDisplay
  
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

  Do While Not rsVideo.EOF
    Set objDisplay = New VideoDisplay

    With objDisplay
      .VideoID = rsVideo("VideoID")
      .Title = rsVideo("Title")
      .ReleaseDate = rsVideo("ReleaseDate")
      mcolDisplay.Add objDisplay
      Set objDisplay = Nothing
      rsVideo.MoveNext
    End With

  Loop

  rsVideo.Close
  Set rsVideo = Nothing

End Sub

