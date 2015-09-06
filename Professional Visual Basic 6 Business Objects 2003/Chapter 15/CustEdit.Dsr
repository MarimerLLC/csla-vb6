VERSION 5.00
Begin {90290CCD-F27D-11D0-8031-00C04FB6C701} CustEdit 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   _ExtentX        =   13467
   _ExtentY        =   8864
   SourceFile      =   ""
   BuildFile       =   ""
   BuildMode       =   0
   TypeLibCookie   =   302
   AsyncLoad       =   0   'False
   id              =   "DHTMLPage1"
   ShowBorder      =   -1  'True
   ShowDetail      =   0   'False
   AbsPos          =   -1  'True
   HTMLDocument    =   "CustEdit.dsx":0000
End
Attribute VB_Name = "CustEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mobjCustomer As Customer
Attribute mobjCustomer.VB_VarHelpID = -1
Private mflgLoading As Boolean


Private Sub mobjCustomer_Valid(IsValid As Boolean)
  cmdOK.disabled = Not IsValid
  cmdApply.disabled = Not IsValid

End Sub

Private Sub txtAddr1_onblur()
  txtAddr1.Value = mobjCustomer.Address1
End Sub

Private Sub txtAddr2_onblur()
  txtAddr2.Value = mobjCustomer.Address2
End Sub

Private Sub txtCity_onblur()
  txtCity.Value = mobjCustomer.City
End Sub

Private Sub txtName_onblur()
  txtName.Value = mobjCustomer.Name
End Sub

Private Function txtName_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.Name = txtName.Value

  If Err Then
    Beep
    txtName.Value = mobjCustomer.Name
  End If

End Function

Private Function txtAddr1_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.Address1 = txtAddr1.Value

  If Err Then
    Beep
    txtAddr1.Value = mobjCustomer.Address1
  End If

End Function

Private Function txtAddr2_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.Address2 = txtAddr2.Value

  If Err Then
    Beep
    txtAddr2.Value = mobjCustomer.Address2
  End If

End Function

Private Function txtCity_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.City = txtCity.Value

  If Err Then
    Beep
    txtCity.Value = mobjCustomer.City
  End If

End Function

Private Sub txtPhone_onblur()
  txtPhone.Value = mobjCustomer.Phone
End Sub

Private Sub txtState_onblur()
  txtState.Value = mobjCustomer.State
End Sub

Private Function txtState_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.State = txtState.Value

  If Err Then
    Beep
    txtState.Value = mobjCustomer.State
  End If

End Function

Private Sub txtZipCode_onblur()
  txtZipCode.Value = mobjCustomer.ZipCode
End Sub

Private Function txtZipCode_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.ZipCode = txtZipCode.Value

  If Err Then
    Beep
    txtZipCode.Value = mobjCustomer.ZipCode
  End If

End Function

Private Function txtPhone_onchange() As Boolean

  If mflgLoading Then Exit Function
  On Error Resume Next
  mobjCustomer.Phone = txtPhone.Value

  If Err Then
    Beep
    txtPhone.Value = mobjCustomer.Phone
  End If

End Function

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.disabled = Not flgValid
  cmdApply.disabled = Not flgValid

End Sub

Private Sub DHTMLPage_Load()

  Dim lngID As Long
  
  mflgLoading = True
  lngID = GetProperty(BaseWindow.Document, "CustomerID")
  Set mobjCustomer = New Customer
  mobjCustomer.Load lngID

  With mobjCustomer
    EnableOK .IsValid
    If .IsNew Then
      BaseWindow.Document.Title = "Customer [(new)]"

    Else
      BaseWindow.Document.Title = "Customer [" & .Name & "]"

    End If

    txtName.Value = .Name
    txtAddr1.Value = .Address1
    txtAddr2.Value = .Address2
    txtCity.Value = .City
    txtState.Value = .State
    txtZipCode.Value = .ZipCode
    txtPhone.Value = .Phone
    .BeginEdit
  End With

  mflgLoading = False

End Sub

Private Function cmdApply_onclick() As Boolean

  mobjCustomer.ApplyEdit
  mobjCustomer.BeginEdit

End Function

Private Function cmdOK_onclick() As Boolean

  mobjCustomer.ApplyEdit
  BaseWindow.navigate "VideoDHTML_CustSearch.htm"

End Function

Private Function cmdCancel_onclick() As Boolean

  mobjCustomer.CancelEdit
  BaseWindow.navigate "VideoDHTML_CustSearch.htm"

End Function

