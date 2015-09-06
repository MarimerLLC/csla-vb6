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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private strBuffer As String
Private udtProps As BufferProps
Private lngPos As Long

Public Sub Initialize(Length As Integer, EstimatedCount As Long)
  With udtProps
    .Length = Length
    .EstCount = EstimatedCount
    .MaxCount = EstimatedCount
    .Count = 0
    strBuffer = Space$(BUFFER_START + .MaxCount * .Length)
    lngPos = BUFFER_START
  End With
End Sub

Public Sub Add(Data As String)
  With udtProps
    If .Count = .MaxCount Then
      strBuffer = strBuffer & _
        Space$(udtProps.EstCount / 2 * udtProps.Length)
      .MaxCount = .MaxCount + udtProps.EstCount / 2
    End If
    Mid$(strBuffer, lngPos, .Length) = Data
    lngPos = lngPos + .Length
    .Count = .Count + 1
  End With
End Sub

Public Property Get Item(Index As Long) As String
  Item = Mid$(strBuffer, BUFFER_START + (Index - 1) * _
    udtProps.Length, udtProps.Length)
End Property

Public Property Let Item(Index As Long, Buffer As String)
  Mid$(strBuffer, BUFFER_START + (Index - 1) * _
    udtProps.Length, udtProps.Length) = Buffer
End Property

Public Function Count() As Long
  Count = udtProps.Count
End Function

Public Function Length() As Long
  Length = udtProps.Length
End Function

Public Function GetState() As String
  Dim udtData As BufferData

  LSet udtData = udtProps
  Mid$(strBuffer, 1, Len(udtData.Buffer)) = udtData.Buffer
  GetState = Left$(strBuffer, lngPos)
End Function

Public Sub SetState(Buffer As String)
  Dim udtData As BufferData
  
  udtData.Buffer = Mid$(Buffer, 1, Len(udtData.Buffer))
  LSet udtProps = udtData
  strBuffer = Buffer
End Sub


