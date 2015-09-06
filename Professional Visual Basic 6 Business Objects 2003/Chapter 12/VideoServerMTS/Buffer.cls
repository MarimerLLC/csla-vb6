VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BufferProps
  Length As Integer
  EstCount As Long
  MaxCount As Long
  Count As Long
End Type

Private Type BufferData
  Buffer As String * 8
End Type

Private Const BUFFER_START = 9
Private mstrBuffer As String
Private mudtProps As BufferProps
Private mlngPos As Long

Public Sub Initialize(Length As Integer, EstimatedCount As Long)

  With mudtProps
    .Length = Length
    .EstCount = EstimatedCount
    .MaxCount = EstimatedCount
    .Count = 0
    mstrBuffer = Space$(BUFFER_START + .MaxCount * .Length)
    mlngPos = BUFFER_START
  End With

End Sub

Public Sub Add(Data As String)

  With mudtProps
    If .Count = .MaxCount Then
      mstrBuffer = mstrBuffer & _
        Space$(mudtProps.EstCount / 2 * mudtProps.Length)
      .MaxCount = .MaxCount + mudtProps.EstCount / 2
    End If

    Mid$(mstrBuffer, mlngPos, .Length) = Data
    mlngPos = mlngPos + .Length
    .Count = .Count + 1
  End With

End Sub

Public Function GetState() As String

  Dim udtData As BufferData

  LSet udtData = mudtProps
  Mid$(mstrBuffer, 1, Len(udtData.Buffer)) = udtData.Buffer
  GetState = Left$(mstrBuffer, mlngPos)

End Function

Public Sub SetState(Buffer As String)

  Dim udtData As BufferData
  
  udtData.Buffer = Mid$(Buffer, 1, Len(udtData.Buffer))
  LSet mudtProps = udtData
  mstrBuffer = Buffer

End Sub

Public Property Get Item(Index As Long) As String

  Item = Mid$(mstrBuffer, BUFFER_START + (Index - 1) * _
    mudtProps.Length, mudtProps.Length)

End Property

Public Function Count() As Long

  Count = mudtProps.Count

End Function

Public Function Length() As Long

  Length = mudtProps.Length

End Function

Public Property Let Item(Index As Long, Buffer As String)

  Mid$(mstrBuffer, BUFFER_START + (Index - 1) * _
    mudtProps.Length, mudtProps.Length) = Buffer

End Property


