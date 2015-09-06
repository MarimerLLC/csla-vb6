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

Public Function Fetch(ByVal ID As Long) As String
  Dim rsClient As Recordset
  Dim strSQL As String
  Dim udtProps As ClientProps
  Dim udtData As ClientData
  Dim objPersist As ProjectPersist
  Dim objPB As PropertyBag
  
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
    
    If mobjContext Is Nothing Then
      Set objPersist = New ProjectPersist
    Else
      Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
    End If
    Set objPB = New PropertyBag
    With objPB
      .WriteProperty "State", udtData.Buffer
      .WriteProperty "Projects", objPersist.Fetch(ID)
      Fetch = .Contents
    End With
    Set objPB = Nothing
    Set objPersist = Nothing
  Else
    ' force an error
    rsClient.MoveNext
  End If
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  rsClient.Close
  Set rsClient = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
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
  
  If mobjContext Is Nothing Then
    Set objPersist = New ProjectPersist
  Else
    Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
  End If
  objPersist.DeleteObject ID
  Set objPersist = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Sub
  
ERRH:
  Set cnClient = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub

Public Function Save(ByVal Buffer As String) As String
  Dim rsClient As Recordset
  Dim strSQL As String
  Dim udtProps As ClientProps
  Dim udtData As ClientData
  Dim objPersist As ProjectPersist
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    udtData.Buffer = .ReadProperty("State")
  End With
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
  
  Set objPBOut = New PropertyBag
  
  LSet udtData = udtProps
  If mobjContext Is Nothing Then
    Set objPersist = New ProjectPersist
  Else
    Set objPersist = mobjContext.CreateInstance("TaskServer.ProjectPersist")
  End If
  With objPBOut
    .WriteProperty "State", udtData.Buffer
    .WriteProperty "Projects", _
      objPersist.Save(objPB.ReadProperty("Projects"), udtProps.ID)
  End With
  Set objPB = Nothing
  Set objPersist = Nothing

  Save = objPBOut.Contents
  Set objPBOut = Nothing
  
  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  rsClient.Close
  Set rsClient = Nothing
  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Private Sub ObjectControl_Activate()
  Set mobjContext = GetObjectContext
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
  ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
  Set mobjContext = Nothing
End Sub


