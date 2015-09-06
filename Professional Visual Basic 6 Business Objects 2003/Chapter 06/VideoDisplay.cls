VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
  VideoID As Long
  Title As String * 30
  ReleaseDate As Variant
End Type

Private mudtProps As DisplayProps

Public Property Get VideoID() As Long

  VideoID = mudtProps.VideoID

End Property

Public Property Get Title() As String

  Title = Trim$(mudtProps.Title)

End Property

Public Property Get ReleaseDate() As Variant

  ReleaseDate = Trim$(mudtProps.ReleaseDate)

End Property

Friend Property Let VideoID(Value As Long)

  mudtProps.VideoID = Value

End Property

Friend Property Let Title(Value As String)

  mudtProps.Title = Value

End Property

Friend Property Let ReleaseDate(Value As Variant)

  mudtProps.ReleaseDate = Value

End Property





