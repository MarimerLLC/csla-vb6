VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClusterClient"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjCluster As Object
Private mstrServerName As String
Private mdtLast As Date
Private mcolServers As Collection

Private Const ERR_NO_SERVERS = vbObjectError + 1001
Private Const ERR_CANT_FIND_SERVER = 462
Private Const ERR_CANT_CREATE_OBJECT = 429

Private Sub Class_Initialize()
  Dim objShell As IWshShell_Class
  Dim lngCount As Long
  Dim lngIndex As Long
  
  Set mcolServers = New Collection
  
  On Error Resume Next
  Set objShell = New IWshShell_Class
  With objShell
    lngCount = .RegRead("HKEY_LOCAL_MACHINE\Software\Wrox\Cluster\ClusterCount")
    For lngIndex = 1 To lngCount
      mcolServers.Add _
        objShell.RegRead("HKEY_LOCAL_MACHINE\Software\Wrox\Cluster\" & _
          "ClusterServer" & CStr(lngIndex))
    Next
  End With
  Set objShell = Nothing
End Sub

Public Function GetServerName() As String
  If DateDiff("s", mdtLast, Now) > 3 Then

    If mobjCluster Is Nothing Then _
      GetClusterReference

    If Not mobjCluster Is Nothing Then
      mstrServerName = mobjCluster.GetServerName
      mdtLast = Now
    End If

  End If
  GetServerName = mstrServerName
End Function

Private Sub GetClusterReference(Optional Index As Long)
  If Index + 1 > mcolServers.Count Then Exit Sub

  On Error Resume Next
  Set mobjCluster = VBA.CreateObject("ClusterServer.Cluster", _
    mcolServers(Index + 1))

  If Err = ERR_CANT_FIND_SERVER Or Err = ERR_CANT_CREATE_OBJECT Then _
      GetClusterReference (Index + 1)

End Sub

Public Function CreateObject(ProgID As String) As Object
  Dim strName As String
  
  strName = GetServerName
  If Len(strName) > 0 Then

    On Error Resume Next
    Set CreateObject = VBA.CreateObject(ProgID, strName)

    If Err = ERR_CANT_FIND_SERVER Then
      On Error GoTo 0

      strName = GetServerName
      If Len(strName) > 0 Then
        On Error Resume Next
        Set CreateObject = VBA.CreateObject(ProgID, strName)
      End If

    Else
      On Error GoTo 0
      If Err <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
    End If

  Else
    Err.Raise ERR_NO_SERVERS, "Cluster", "No servers available"
  End If
End Function

Private Sub Class_Terminate()
  Set mobjCluster = Nothing
End Sub


