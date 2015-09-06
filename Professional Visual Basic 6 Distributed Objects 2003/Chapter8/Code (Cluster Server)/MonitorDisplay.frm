VERSION 5.00
Begin VB.Form MonitorDisplay 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Restart 
      Enabled         =   0   'False
      Left            =   0
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   720
   End
   Begin VB.ListBox lstServers 
      Height          =   2985
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "MonitorDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Initialize
End Sub

Private Sub Form_Resize()
  lstServers.Move 1, 1, ScaleWidth, ScaleHeight
End Sub

Private Sub Initialize()
  Dim objShell As IWshShell_Class
  Dim lngCount As Long
  Dim lngIndex As Long
  
  Set gcolServers = New Collection
  
  Set objShell = New IWshShell_Class
  With objShell
    lngCount = .RegRead("HKEY_LOCAL_MACHINE\Software\Wrox\Cluster\ServerCount")
    For lngIndex = 1 To lngCount
      AddServer _
        objShell.RegRead("HKEY_LOCAL_MACHINE\Software\Wrox\Cluster\Server" & _
        CStr(lngIndex))
    Next
  End With
  Set objShell = Nothing
  
  Set gobjMM = New MemoryMap
  gobjMM.Initialize "WroxCluster", 32768
  
  With Timer1
    .Interval = REPORT_INTERVAL
    .Enabled = True
  End With
  Timer1_Timer
  With Restart
    .Interval = 60000
    .Enabled = True
  End With
End Sub

Private Sub AddServer(Server As String)
  Dim objServer As ClusterSvr
  
  Set objServer = New ClusterSvr
  objServer.Attach Server
  gcolServers.Add objServer, Server
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim lngIndex As Long
  Dim objServer As ClusterSvr
  
  For lngIndex = gcolServers.Count To 1 Step -1
    Set objServer = gcolServers(lngIndex)
    objServer.Shutdown
    gcolServers.Remove lngIndex
  Next
End Sub

Private Sub Timer1_Timer()
  Dim objPB As PropertyBag
  Dim objServer As ClusterSvr
  Dim lngCount As Long
  Static intDisplay As Integer
  
  Timer1.Enabled = False
  Set objPB = New PropertyBag
  If intDisplay = 0 Then lstServers.Clear
  With objPB
    .WriteProperty "ServerCount", gcolServers.Count
    lngCount = 0
    For Each objServer In gcolServers
      lngCount = lngCount + 1
      .WriteProperty "ServerName" & CStr(lngCount), _
        objServer.ServerName
      .WriteProperty "ServerStatus" & CStr(lngCount), _
        objServer.Status
      If intDisplay = 0 Then _
        lstServers.AddItem Format$(objServer.Status, "0000") & _
          " (" & objServer.ServerName & ")"
    Next
  End With
  gobjMM.SetData objPB
  intDisplay = intDisplay + 1
  If intDisplay = 5 Then intDisplay = 0
  Timer1.Enabled = True
End Sub

Private Sub Restart_Timer()
  Dim lngIndex As Long
  Dim objServer As ClusterSvr
  Static intCount As Integer
  
  intCount = intCount + 1
  If intCount < 5 Then Exit Sub
  intCount = 0
  
  Timer1.Enabled = False
  Restart.Enabled = False
  
  For lngIndex = gcolServers.Count To 1 Step -1
    Set objServer = gcolServers(lngIndex)
    If objServer.Status = 0 Then
      gcolServers.Remove objServer.ServerName
      AddServer objServer.ServerName
    End If
    Set objServer = Nothing
  Next
  
  Restart.Enabled = True
  Timer1.Enabled = True
End Sub

