Attribute VB_Name = "VUImain"
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

Public Function TextLostFocus(Ctl As TextBox, Obj As Object, Prop As String)
  
  TextLostFocus = CallByName(Obj, Prop, VbGet)

End Function

