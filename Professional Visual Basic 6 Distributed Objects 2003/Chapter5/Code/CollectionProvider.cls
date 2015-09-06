VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements OLEDBSimpleProvider

Private mcolFields As Collection
Private mcolListeners As Collection
Private mobjOSPL As OLEDBSimpleProviderListener
Private mobjBO As Object

Private Sub Class_Initialize()
  Set mcolListeners = New Collection
End Sub

Private Sub Class_Terminate()
  Set mobjBO = Nothing
End Sub

Public Sub LoadData(DataObject As Object, Fields() As String)
  Dim intIndex As Integer
  Dim intCol As Integer
  Dim objField As Field
  Dim objTA As TLIApplication
  Dim objTI As InterfaceInfo
  Dim objBO As Object
  
  Set mobjBO = DataObject
  Set mcolFields = New Collection
  If Fields(0) <> "*" Then
    For intCol = LBound(Fields) To UBound(Fields)
      Set objField = New Field
      objField.Name = Fields(intCol)
      objField.Writable = False
      mcolFields.Add objField, Fields(intCol)
    Next
    If mobjBO.Count > 0 Then
      Set objBO = mobjBO.Item(1)
    Else
      Exit Sub
    End If
  Else
    MsgBox ("Pre-access.B")
    If mobjBO.Count > 0 Then
      Set objBO = mobjBO.Item(1)
    Else
      Err.Raise DB_E_NOTABLE, "CustomData", _
        "No items in collection - unable to access schema"
      Exit Sub
    End If
    MsgBox ("Post-access.B")
  End If
  On Error GoTo ErrorTrap
  
  Set objTA = New TLIApplication
  Set objTI = objTA.InterfaceInfoFromObject(objBO)
  
  With objTI
    If mcolFields.Count = 0 Then
      For intIndex = 1 To .Members.Count
        If .Members(intIndex).InvokeKind = INVOKE_PROPERTYGET Then
          Set objField = New Field
          objField.Name = .Members(intIndex).Name
          objField.Writable = False
          On Error Resume Next
          mcolFields.Add objField, .Members(intIndex).Name
        End If
        On Error GoTo ErrorTrap
      Next
    End If
    On Error Resume Next
    For intIndex = 1 To .Members.Count
      If .Members(intIndex).InvokeKind = INVOKE_PROPERTYPUT Then
        Set objField = mcolFields(.Members(intIndex).Name)
        objField.Writable = True
      End If
    Next
    On Error GoTo ErrorTrap
  End With
  Set objBO = Nothing
  Set objTI = Nothing
  Set objTA = Nothing
  Exit Sub

ErrorTrap:
  Err.Raise (E_FAIL)
End Sub

Private Sub OLEDBSimpleProvider_addOLEDBSimpleProviderListener _
            (ByVal pospIListener As OLEDBSimpleProviderListener)
  
  If Not (pospIListener Is Nothing) Then
    Set mobjOSPL = pospIListener
    mcolListeners.Add mobjOSPL
  End If
End Sub

Private Function OLEDBSimpleProvider_deleteRows _
    (ByVal iRow As Long, ByVal cRows As Long) As Long
  
  Err.Raise (E_FAIL)
End Function

Private Function OLEDBSimpleProvider_find(ByVal ilngStart As Long, _
        ByVal iColumn As Long, ByVal val As Variant, _
        ByVal findFlags As OSPFIND, ByVal compType As OSPCOMP) As Long
  
  Dim lngStop As Long
  Dim intStep As Integer
  Dim lngIndex As Long
  Dim intCaseSensitive As Integer
  Dim flgStringComp As Boolean
  Dim lngResult As Long
  Dim objBO As Object
  Dim vntValue As Variant
  Dim intCompResult As Integer
  
  If findFlags And OSPFIND_UP Then
    lngStop = 1
    intStep = -1
  Else
    lngStop = mobjBO.Count
    intStep = 1
  End If

  If findFlags And OSPFIND_CASESENSITIVE Then
    intCaseSensitive = vbTextCompare
  Else
    intCaseSensitive = vbBinaryCompare
  End If
  
  If VarType(val) = vbString Then
    flgStringComp = True
  Else
    flgStringComp = False
  End If
  
  lngResult = -1
  For lngIndex = ilngStart To lngStop Step intStep
    objBO = mobjBO.Item(lngIndex)
    vntValue = vntValue
    
    If flgStringComp Then       ' doing string compares
      intCompResult = StrComp(vntValue, val, intCaseSensitive)
      Select Case compType
      Case OSPCOMP_DEFAULT, OSPCOMP_EQ
        If intCompResult = 0 Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_GE
        If intCompResult >= 0 Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_GT
        If intCompResult > 0 Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_LE
        If intCompResult <= 0 Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_LT
        If intCompResult < 0 Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_NE
        If intCompResult <> 0 Then
          lngResult = lngIndex
          Exit For
        End If
      End Select
    Else                        ' normal (non-string) compare
      Select Case compType
      Case OSPCOMP_DEFAULT, OSPCOMP_EQ
        If vntValue = val Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_GE
        If vntValue >= val Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_GT
        If vntValue > val Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_LE
        If vntValue <= val Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_LT
        If vntValue < val Then
          lngResult = lngIndex
          Exit For
        End If
      Case OSPCOMP_NE
        If vntValue <> val Then
          lngResult = lngIndex
          Exit For
        End If
      End Select
    End If
  Next lngIndex
  
  OLEDBSimpleProvider_find = lngResult
End Function

Private Function OLEDBSimpleProvider_getColumnCount() As Long
  OLEDBSimpleProvider_getColumnCount = mcolFields.Count
End Function

Private Function OLEDBSimpleProvider_getEstimatedRows() As Long
  OLEDBSimpleProvider_getEstimatedRows = mobjBO.Count
End Function

Private Function OLEDBSimpleProvider_getLocale() As String
  OLEDBSimpleProvider_getLocale = ""
End Function

Private Function OLEDBSimpleProvider_getRowCount() As Long
  OLEDBSimpleProvider_getRowCount = mobjBO.Count
End Function

Private Function OLEDBSimpleProvider_getRWStatus _
          (ByVal iRow As Long, ByVal iColumn As Long) As OSPRW

  If mcolFields(iColumn).Writable Then
    OLEDBSimpleProvider_getRWStatus = OSPRW_READWRITE
  Else
    OLEDBSimpleProvider_getRWStatus = OSPRW_READONLY
  End If
End Function

Private Function OLEDBSimpleProvider_getVariant _
          (ByVal iRow As Long, ByVal iColumn As Long, _
          ByVal format As OSPFORMAT) As Variant
  
  Dim objBO As Object
  
  If iRow = 0 Then
    OLEDBSimpleProvider_getVariant = mcolFields(iColumn).Name
  Else
    Set objBO = mobjBO.Item(iRow)
    OLEDBSimpleProvider_getVariant = _
      CallByName(objBO, mcolFields(iColumn).Name, VbGet)
  End If
End Function

Private Function OLEDBSimpleProvider_insertRows _
        (ByVal iRow As Long, ByVal cRows As Long) As Long
  
  Err.Raise (E_FAIL)
End Function

Private Function OLEDBSimpleProvider_isAsync() As Long
  OLEDBSimpleProvider_isAsync = False
End Function

Private Sub OLEDBSimpleProvider_removeOLEDBSimpleProviderListener _
          (ByVal pospIListener As OLEDBSimpleProviderListener)

  Dim intIndex As Long
  
  For intIndex = 1 To mcolListeners.Count
    If mcolListeners(intIndex) Is pospIListener Then
      mcolListeners.Remove intIndex
    End If
  Next
End Sub

Private Sub OLEDBSimpleProvider_setVariant(ByVal iRow As Long, _
          ByVal iColumn As Long, ByVal format As OSPFORMAT, _
          ByVal Var As Variant)
  
  Dim objListener As OLEDBSimpleProviderListener
  Dim vntItem As Variant
  Dim objBO As Object

  For Each vntItem In mcolListeners
    Set objListener = vntItem
    objListener.aboutToChangeCell iRow, iColumn    ' Pre-notification
  Next
  
  Set objBO = mobjBO.Item(iRow)
  CallByName objBO, mcolFields(iColumn).Name, VbLet, CStr(Var)
  
  For Each vntItem In mcolListeners
    Set objListener = vntItem
    objListener.cellChanged iRow, iColumn          ' Post-notification
  Next
End Sub

Private Sub OLEDBSimpleProvider_stopTransfer()
  ' Do nothing because we already have all our data
End Sub



