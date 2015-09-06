VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ClientPersist"
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

Public Function Fetch(ByVal ID As Long, Projects As String) As String
  Dim rsClient As Recordset
  Dim strSQL As String
  Dim udtProps As ClientProps
  Dim udtData As ClientData
  Dim objPersist As ProjectPersist
  
  Set rsClient = New Recordset
  strSQL = "SELECT * FROM CLIENTS WHERE ID=" & CStr(ID)
  rsClient.Open strSQL, DB_CONN, _
    adOpenForwardOnly, adLockReadOnly
  On Error GoTo ERRH
  If Not rsClient.EOF Then
    With udtProps
      .ID = rsClient("ID")
      .Name = rsClient("Name")
      .ContactName = rsClient("ContactName")
      .Phone = rsClient("Phone")
      .IsNew = False
      .IsDirty = False
      .IsDeleted = False
    End With
    rsClient.Close
    Set rsClient = Nothing
    LSet udtData = udtProps
    Fetch = udtData.Buffer
    
    If mflgInMTS Then
      Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
    Else
      Set objPersist = CreateObject("TaskServer.ProjectPersist")
    End If
    
    Projects = objPersist.Fetch(ID)
    Set objPersist = Nothing
  Else
    ' force an error
    rsClient.MoveNext
  End If
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  If Not rsClient Is Nothing Then
    rsClient.Close
    Set rsClient = Nothing
  End If
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Function Save(ByVal Buffer As String, Projects As String) As String
  Dim rsClient As Recordset
  Dim strSQL As String
  Dim udtProps As ClientProps
  Dim udtData As ClientData
  Dim objPersist As ProjectPersist
  
  udtData.Buffer = Buffer
  LSet udtProps = udtData
  
  strSQL = "SELECT * FROM CLIENTS WHERE ID=" & CStr(udtProps.ID)
  Set rsClient = New Recordset
  rsClient.Open strSQL, DB_CONN, adOpenKeyset, adLockOptimistic
  On Error GoTo ERRH
  If udtProps.IsNew Then rsClient.AddNew
  
  With udtProps
    rsClient("Name") = .Name
    rsClient("ContactName") = .ContactName
    rsClient("Phone") = .Phone
    rsClient.Update
    If .IsNew Then
      rsClient.Bookmark = rsClient.Bookmark
      .ID = rsClient("ID")
    End If
    .IsNew = False
    .IsDirty = False
  End With
  rsClient.Close
  Set rsClient = Nothing
  
  If mflgInMTS Then
    Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
  Else
    Set objPersist = CreateObject("TaskServer.ProjectPersist")
  End If
  
  Projects = objPersist.Save(Projects, udtProps.ID)
  Set objPersist = Nothing
  
  LSet udtData = udtProps
  Save = udtData.Buffer
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  If Not rsClient Is Nothing Then
    rsClient.Close
    Set rsClient = Nothing
  End If
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Sub DeleteObject(ByVal ID As Long)
  Dim cnClient As Connection
  Dim strSQL As String
  Dim objPersist As ProjectPersist
  
  Set cnClient = New Connection
  On Error GoTo ERRH
  With cnClient
    .Open DB_CONN
    strSQL = "DELETE FROM CLIENTS WHERE ID=" & CStr(ID)
    .Execute strSQL
    .Close
  End With
  Set cnClient = Nothing
  
  If mflgInMTS Then
    Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
  Else
    Set objPersist = CreateObject("TaskServer.ProjectPersist")
  End If
  
  objPersist.DeleteObject ID
  Set objPersist = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Sub
  
ERRH:
  Set cnClient = Nothing
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub




