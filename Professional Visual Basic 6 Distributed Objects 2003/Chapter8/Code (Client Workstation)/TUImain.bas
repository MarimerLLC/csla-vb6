Attribute VB_Name = "TUImain"
Option Explicit

Public Sub TextChange(Ctl As TextBox, Obj As Object, Prop As String)
  Dim lngPos As Long
  
  On Error GoTo INPUTERR
  CallByName Obj, Prop, VbLet, Ctl.Text
  Exit Sub
  
INPUTERR:
  Beep
  lngPos = Ctl.SelStart
  Ctl = CallByName(Obj, Prop, VbGet)
  Ctl.SelStart = lngPos - 1
End Sub

Public Function TextLostFocus(Obj As Object, Prop As String) As String
  TextLostFocus = CallByName(Obj, Prop, VbGet)
End Function

Public Function SelectedItem(List As DataList, _
                             BoundCollection As Object) As Long
  Dim lngIndex As Long
  
  With BoundCollection
    For lngIndex = 1 To .Count
      If .Item(lngIndex).ID = Val(List.BoundText) Then
        SelectedItem = lngIndex
        Exit For
      End If
    Next
  End With
End Function


