VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CustomerList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer List"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "CustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjCustomers As Customers
Private mlngID As Long

Public Sub Component(objComponent As Customers)

  Dim objItem As CustomerDisplay
  Dim itmList As ListItem
  Dim lngIndex As Long
  
  Set mobjCustomers = objComponent
  For lngIndex = 1 To mobjCustomers.Count
    With objItem
      Set objItem = mobjCustomers.Item(lngIndex)
      Set itmList = _
        lvwItems.ListItems.Add(Key:= _
        Format$(objItem.CustomerID) & " K")

      With itmList
        .Text = objItem.Name
        .SubItems(1) = objItem.Phone
      End With

    End With

  Next

End Sub

Private Sub cmdCancel_Click()

  mlngID = 0
  Hide

End Sub

Private Sub cmdOK_Click()

  On Error Resume Next
  mlngID = Val(lvwItems.SelectedItem.Key)
  Hide

End Sub

Public Property Get CustomerID() As Long

  CustomerID = mlngID

End Property

