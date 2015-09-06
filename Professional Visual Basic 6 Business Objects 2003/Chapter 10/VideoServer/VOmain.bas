Attribute VB_Name = "VOmain"
Option Explicit

Public cnVideo As Connection

Public Sub Main()

  Set cnVideo = New Connection
  cnVideo.Provider = "Microsoft.Jet.OLEDB.3.51"
  cnVideo.Mode = adModeReadWrite
  cnVideo.Open "C:\Wrox\VB6 Pro Objects\Video.mdb"

End Sub

