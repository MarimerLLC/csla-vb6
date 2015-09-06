VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin {90290CCD-F27D-11D0-8031-00C04FB6C701} CustList 
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   _ExtentX        =   18627
   _ExtentY        =   12091
   SourceFile      =   ""
   BuildFile       =   ""
   BuildMode       =   0
   TypeLibCookie   =   228
   AsyncLoad       =   0   'False
   id              =   "CustList"
   ShowBorder      =   -1  'True
   ShowDetail      =   0   'False
   AbsPos          =   -1  'True
   HTMLDocument    =   "CustList.dsx":0000
End
Attribute VB_Name = "CustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function cmdCancel_onclick() As Boolean
  BaseWindow.navigate "VideoDHTML_CustSearch.html"
End Function

Private Function cmdOK_onclick() As Boolean
  PutProperty BaseWindow.Document, "CustomerID", Val(lvwItems.SelectedItem.Key)
  BaseWindow.navigate "VideoDHTML_CustEdit.html"

End Function

Private Sub DHTMLPage_Load()

  Dim strName As String
  Dim strPhone As String
  Dim objCustomers As Customers
  Dim divCustInfo As Object
  
  strName = GetProperty(BaseWindow.Document, "SearchName")
  strPhone = GetProperty(BaseWindow.Document, "SearchPhone")
  
  Set objCustomers = New Customers
  objCustomers.Load strName, strPhone

  ListCustomers objCustomers
  Set objCustomers = Nothing

End Sub

Private Sub ListCustomers(objCustomers As Customers)

  Dim objItem As CustomerDisplay
  Dim itmList As ListItem
  Dim lngIndex As Long
  
  With lvwItems
    .View = lvwReport
    .FullRowSelect = True
    .LabelEdit = lvwManual
    .ColumnHeaders(1).Text = "Name"
    .ColumnHeaders.Add Text:="Phone"
  End With

  For lngIndex = 1 To objCustomers.Count
    Set objItem = objCustomers.Item(lngIndex)
    Set itmList = _
      lvwItems.ListItems.Add(Key:= _
      Format$(objItem.CustomerID) & " K")

    With itmList
      .Text = objItem.Name
      .SubItems(1) = objItem.Phone
    End With

  Next

End Sub


