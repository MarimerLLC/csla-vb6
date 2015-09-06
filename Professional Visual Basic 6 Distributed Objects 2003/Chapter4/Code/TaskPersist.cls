VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TaskPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext
Private mflgInMTS As Boolean

Public Function Fetch(ByVal Project As Long) As String
  Dim rsTask As Recordset
  Dim strSQL As String
  Dim udtProps As TaskProps
  Dim udtData As TaskData
  Dim objPB As PropertyBag
  Dim lngCount As Long
  
  Set rsTask = New Recordset
  strSQL = "SELECT * FROM TASKS WHERE Project=" & CStr(Project)
  rsTask.Open strSQL, DB_CONN, _
    adOpenForwardOnly, adLockReadOnly
  Set objPB = New PropertyBag
  On Error GoTo ERRH
  Do While Not rsTask.EOF
    With udtProps
      .ID = rsTask("ID")
      .Name = rsTask("Name")
      .PercentComplete = IIf(IsNull(rsTask("PercentComplete")), 0, _
        rsTask("PercentComplete"))
      .ProjectedDays = IIf(IsNull(rsTask("ProjectedDays")), 0, _
        rsTask("ProjectedDays"))
      .IsNew = False
      .IsDirty = False
      .IsDeleted = False
    End With
    LSet udtData = udtProps
    lngCount = lngCount + 1
    objPB.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
    rsTask.MoveNext
  Loop
  rsTask.Close
  Set rsTask = Nothing
  With objPB
    .WriteProperty "Count", lngCount
    Fetch = .Contents
  End With
  Set objPB = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  If Not rsTask Is Nothing Then
    rsTask.Close
    Set rsTask = Nothing
  End If
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Public Function Save(ByVal Buffer As String, ByVal ProjectID As Long) As String
  Dim rsTask As Recordset
  Dim strSQL As String
  Dim udtProps As TaskProps
  Dim udtData As TaskData
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngCount As Long
  
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = Buffer
  objPB.Contents = arBuffer
  Set rsTask = New Recordset
  
  On Error GoTo ERRH
  For lngIndex = 1 To objPB.ReadProperty("Count")
    udtData.Buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM TASKS WHERE ID=" & CStr(udtProps.ID)
      rsTask.Open strSQL, DB_CONN, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rsTask.AddNew
      
      With udtProps
        rsTask("Name") = .Name
        rsTask("PercentComplete") = .PercentComplete
        rsTask("ProjectedDays") = .ProjectedDays
        rsTask("Project") = ProjectID
        rsTask.Update
        If .IsNew Then
          rsTask.Bookmark = rsTask.Bookmark
          .ID = rsTask("ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
      rsTask.Close
    Else
      DeleteTask udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rsTask = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Function
  
ERRH:
  Set objPB = Nothing
  Set objPBOut = Nothing
  Set rsTask = Nothing
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function

Private Sub DeleteTask(TaskID As Long)
  Dim cnClient As Connection
  Dim strSQL As String
  
  strSQL = "DELETE FROM TASKS WHERE ID=" & CStr(TaskID)
  Set cnClient = New Connection
  cnClient.Open DB_CONN
  cnClient.Execute strSQL
  cnClient.Close
  Set cnClient = Nothing
End Sub

Public Sub DeleteObject(ByVal ProjectID As Long)
  Dim cnClient As Connection
  Dim strSQL As String
  
  strSQL = "DELETE FROM TASKS WHERE PROJECT=" & CStr(ProjectID)
  Set cnClient = New Connection
  On Error GoTo ERRH
  cnClient.Open DB_CONN
  cnClient.Execute strSQL
  cnClient.Close
  Set cnClient = Nothing
  If mflgInMTS Then mobjContext.SetComplete
  Exit Sub

ERRH:
  Set cnClient = Nothing
  If mflgInMTS Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub

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
