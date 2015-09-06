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

Public Function Fetch(ByVal Client As Long) As String
  Dim rsProject As Recordset
  Dim strSQL As String
  Dim udtProps As ProjectProps
  Dim udtData As ProjectData
  Dim objPB As PropertyBag
  Dim objProjectPB As PropertyBag
  Dim lngCount As Long
  Dim objPersist As TaskPersist
  
  Set rsProject = New Recordset
  strSQL = "SELECT * FROM PROJECTS WHERE Client=" & CStr(Client)
  rsProject.Open strSQL, DB_CONN, _
    adOpenForwardOnly, adLockReadOnly
  Set objPB = New PropertyBag
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
    Set objProjectPB = New PropertyBag
    With objProjectPB
      .WriteProperty "State", udtData.Buffer
      If mobjContext Is Nothing Then
        Set objPersist = New TaskPersist
      Else
        Set objPersist = mobjContext.CreateInstance("TaskServer.TaskPersist")
      End If
      .WriteProperty "Tasks", objPersist.Fetch(udtProps.ID)
      Set objPersist = Nothing
      lngCount = lngCount + 1
      objPB.WriteProperty "Item" & CStr(lngCount), .Contents
    End With
    Set objProjectPB = Nothing
    rsProject.MoveNext
  Loop
  rsProject.Close
  Set rsProject = Nothing
  objPB.WriteProperty "Count", lngCount
  Fetch = objPB.Contents
  Set objPB = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  rsProject.Close
  Set rsProject = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Function Save(ByVal Buffer As String, ByVal ClientID As Long) As String
  Dim rsProject As Recordset
  Dim strSQL As String
  Dim udtProps As ProjectProps
  Dim udtData As ProjectData
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim objProjectPB As PropertyBag
  Dim objProjectPBOut As PropertyBag
  Dim objPBOut As PropertyBag
  Dim lngIndex As Long
  Dim lngCount As Long
  Dim objPersist As TaskPersist
  
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = Buffer
  objPB.Contents = arBuffer
  
  Set rsProject = New Recordset
  
  On Error GoTo ERRH
  For lngIndex = 1 To objPB.ReadProperty("Count")
    Set objProjectPB = New PropertyBag
    With objProjectPB
      arBuffer = objPB.ReadProperty("Item" & CStr(lngIndex))
      .Contents = arBuffer
      udtData.Buffer = .ReadProperty("State")
      LSet udtProps = udtData
    End With
    
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
      
      Set objProjectPBOut = New PropertyBag
      
      LSet udtData = udtProps
      objProjectPBOut.WriteProperty "State", udtData.Buffer
      
      If mobjContext Is Nothing Then
        Set objPersist = New TaskPersist
      Else
        Set objPersist = mobjContext.CreateInstance("TaskServer.TaskPersist")
      End If
      objProjectPBOut.WriteProperty "Tasks", _
        objPersist.Save(objProjectPB.ReadProperty("Tasks"), udtProps.ID)
      Set objPersist = Nothing
      
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & lngCount, objProjectPBOut.Contents
      
      Set objProjectPB = Nothing
      Set objProjectPBOut = Nothing
      rsProject.Close
    Else
      DeleteProject udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rsProject = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  Set objPB = Nothing
  Set objPBOut = Nothing
  Set rsProject = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Sub DeleteProject(ProjectID As Long)
  Dim cnProject As Connection
  Dim strSQL As String
  
  Set cnProject = New Connection
  With cnProject
    .Open DB_CONN
    strSQL = "DELETE FROM PROJECTS WHERE ID=" & CStr(ProjectID)
    .Execute strSQL
    .Close
  End With
  Set cnProject = Nothing
End Sub

Public Sub DeleteObject(ByVal ClientID As Long)
  Dim cnProject As Connection
  Dim strSQL As String
  Dim rsProject As Recordset
  Dim objPersist As TaskPersist
  
  If mobjContext Is Nothing Then
    Set objPersist = New TaskPersist
  Else
    Set objPersist = mobjContext.CreateInstance("TaskServer.TaskPersist")
  End If
  strSQL = "SELECT ID FROM PROJECTS WHERE Client=" & CStr(ClientID)
  rsProject.Open strSQL, DB_CONN, _
    adOpenForwardOnly, adLockReadOnly
  Do While Not rsProject.EOF
    objPersist.DeleteObject rsProject("ID")
    rsProject.MoveNext
  Loop
  rsProject.Close
  Set rsProject = Nothing
  Set objPersist = Nothing
  
  Set cnProject = New Connection
  On Error GoTo ERRH
  With cnProject
    .Open DB_CONN
    strSQL = "DELETE FROM PROJECTS WHERE CLIENT=" & CStr(ClientID)
    .Execute strSQL
    .Close
  End With
  Set cnProject = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Sub
  
ERRH:
  Set cnProject = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub

Private Sub ObjectControl_Activate()
  Set mobjContext = GetObjectContext
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
  ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
  Set mobjContext = Nothing
End Sub


