VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngMappingHandle As Long
Private mlngMappingAddress As Long
Private mlngMutexHandle As Long
Private mlngSemaphoreHandle As Long

Private Const FILE_MAP_WRITE = &H2
Private Const PAGE_READWRITE = 4&
Private Const ERROR_ALREADY_EXISTS = 183&

Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0&
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&
Private Const WAIT_TIMEOUT = &H102&

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" _
  (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal _
  flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow _
  As Long, ByVal lpName As String) As Long

Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress _
  As Long) As Long

Private Declare Function MapViewOfFile Lib "kernel32" (ByVal _
  hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal _
  dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal _
  dwNumberOfBytesToMap As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) _
  As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
  (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, _
  ByVal lpName As String) As Long

Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) _
  As Long

Private Declare Function CreateSemaphore Lib "kernel32" Alias _
  "CreateSemaphoreA" (ByVal lpSemaphoreAttributes As Long, _
  ByVal lInitialCount As Long, ByVal lMaximumCount As Long, _
  ByVal lpName As String) As Long

Private Declare Function ReleaseSemaphore Lib "kernel32" _
  (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, _
  lpPreviousCount As Long) As Long

Private Declare Function OpenSemaphore Lib "kernel32" Alias _
  "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle _
  As Long, ByVal lpName As String) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
  hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Sub Initialize(ShareName As String, MaxLength As Long)
  Dim lngMapErr As Long
  
  mlngMappingHandle = CreateFileMapping(-1, 0, PAGE_READWRITE, 0, _
    MaxLength, ShareName & ".map")
  If mlngMappingHandle = 0 Then Exit Sub
  lngMapErr = Err.LastDllError
  
  mlngMappingAddress = MapViewOfFile(mlngMappingHandle, FILE_MAP_WRITE, 0, 0, 0)
  If mlngMappingAddress = 0 Then
    CloseHandle mlngMappingHandle
    Exit Sub
  End If
  If Not lngMapErr = ERROR_ALREADY_EXISTS Then _
    CopyMemory ByVal mlngMappingAddress, 0, 4

  mlngMutexHandle = CreateMutex(0, False, ShareName & ".mtx")
  If mlngMutexHandle = 0 Then
    UnmapViewOfFile mlngMappingAddress
    CloseHandle mlngMappingHandle
    Exit Sub
  End If
  
  mlngSemaphoreHandle = OpenSemaphore(-1, False, ShareName & ".sem")
  If mlngSemaphoreHandle = 0 Then
    mlngSemaphoreHandle = CreateSemaphore(0, 100, 100, ShareName & ".sem")
    If mlngSemaphoreHandle = 0 Then
      UnmapViewOfFile mlngMappingAddress
      CloseHandle mlngMappingHandle
      CloseHandle mlngMutexHandle
      Exit Sub
    End If
  End If
End Sub

Private Sub Class_Terminate()
  UnmapViewOfFile mlngMappingAddress
  CloseHandle mlngMappingHandle
  CloseHandle mlngMutexHandle
  CloseHandle mlngSemaphoreHandle
End Sub

Public Sub SetData(PB As PropertyBag)
  Dim arContents() As Byte
  Dim lngLen As Long
  Dim retval As Long
  
  arContents = PB.Contents
  lngLen = UBound(arContents) + 1
  
  Do
    retval = WaitForSingleObject(mlngMutexHandle, 5000)
  Loop Until retval = WAIT_OBJECT_0
  
  Do
    Do
    Loop Until WaitForSingleObject(mlngSemaphoreHandle, SYNC_WAIT) = _
      WAIT_OBJECT_0
    ReleaseSemaphore mlngSemaphoreHandle, 1, lngSemCount
  Loop Until lngSemCount = 99

  CopyMemory ByVal mlngMappingAddress, lngLen, 4
  CopyMemory ByVal (mlngMappingAddress + 4), arContents(0), lngLen
  ReleaseMutex mlngMutexHandle
End Sub

Public Function GetData() As PropertyBag
  Dim objPB As PropertyBag
  Dim lngLen As Long
  Dim arContents() As Byte
  Dim lngSemCount As Long
  Dim retval As Long
  
  Do
    retval = WaitForSingleObject(mlngMutexHandle, SYNC_WAIT)
    If retval <> WAIT_OBJECT_0 Then DoEvents
  Loop Until retval = WAIT_OBJECT_0
  
  Do
  Loop Until WaitForSingleObject(mlngSemaphoreHandle, SYNC_WAIT) = WAIT_OBJECT_0
  
  ReleaseMutex mlngMutexHandle

  Set objPB = New PropertyBag
  CopyMemory lngLen, ByVal mlngMappingAddress, 4
  If lngLen = 0 Then GoTo leave
  ReDim arContents(lngLen - 1)
  CopyMemory arContents(0), ByVal (mlngMappingAddress + 4), lngLen
  objPB.Contents = arContents
  Set GetData = objPB

leave:
  ReleaseSemaphore mlngSemaphoreHandle, 1, lngSemCount
End Function


