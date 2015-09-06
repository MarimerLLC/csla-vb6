VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ProjectPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
Private mflgInMTS As Boolean

Private Sub ObjectControl_Activate()
  Set mobjContext = GetObjectContext
  mflgInMTS = True
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
  ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
  Set mobjContext = Nothing
End Sub

Public Function Fetch(ByVal Client As Long) As String
  Dim rsProject As Recordset
  Dim strSQL As String
  Dim udtProps As ProjectProps
  Dim udtData As ProjectData
  Dim objBuffer As Buffer
  
  Set rsProject = New Recordset
  strSQL = "SELECT * FROM PROJECTS WHERE Client=" & CStr(Client)
  rsProject.Open strSQL, DB_CONN, _
    adOpenForwardOnly, adLockReadOnly
  Set objBuffer = New Buffer
  objBuffer.Initialize Len(udtData.Buffer), 10
  On Error GoTo ERRH
  Do While Not rsProject.EOF
    With udtProps
      .ID = rsProject("ID")
      .Name = rsProject("Name")
      .IsNew = False
      .IsDirty = False
      .IsDeleted = False
    End With
    LSet udtData = udtProps
    objBuffer.Add udtData.Buffer
    rsProject.MoveNext
  Loop
  rsProject.Close
  Set rsProject = Nothing
  Fetch = objBuffer.GetState
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  If Not rsProject Is Nothing Then
    rsProject.Close
    Set rsProject = Nothing
  End If
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Function Save(ByVal Buffer As String, ByVal ClientID As Long) As String
  Dim rsProject As Recordset
  Dim strSQL As String
  Dim udtProps As ProjectProps
  Dim udtData As ProjectData
  Dim objBuffer As Buffer
  Dim objNewBuffer As Buffer
  Dim lngIndex As Long
  
  Set objBuffer = New Buffer
  With objBuffer
    .SetState Buffer
  
    Set objNewBuffer = New Buffer
    objNewBuffer.Initialize .Length, .Count
  End With
  
  Set rsProject = New Recordset
  
  On Error GoTo ERRH
  For lngIndex = 1 To objBuffer.Count
    udtData.Buffer = objBuffer.Item(lngIndex)
    LSet udtProps = udtData
    
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM PROJECTS WHERE ID=" & CStr(udtProps.ID)
      rsProject.Open strSQL, DB_CONN, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rsProject.AddNew
      
      With udtProps
        rsProject("Name") = .Name
        rsProject("Client") = ClientID
        rsProject.Update
        If .IsNew Then
          rsProject.Bookmark = rsProject.Bookmark
          .ID = rsProject("ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      objNewBuffer.Add udtData.Buffer
      rsProject.Close
    Else
      DeleteObject udtProps.ID
    End If
  Next
  
  Set objBuffer = Nothing
  Set rsProject = Nothing
  
  Save = objNewBuffer.GetState
  Set objNewBuffer = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  Set objBuffer = Nothing
  Set objNewBuffer = Nothing
  Set rsProject = Nothing
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Sub DeleteObject(ByVal ProjectID As Long)
  Dim cnProject As Connection
  Dim strSQL As String
  
  Set cnProject = New Connection
  On Error GoTo ERRH
  With cnProject
    .Open DB_CONN
    strSQL = "DELETE FROM PROJECTS WHERE ID=" & CStr(ProjectID)
    .Execute strSQL
    .Close
  End With
  Set cnProject = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Sub
  
ERRH:
  Set cnProject = Nothing
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub


