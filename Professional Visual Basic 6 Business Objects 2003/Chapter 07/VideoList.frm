VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VideoList 
   Caption         =   "Video List"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3413
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
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Release date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "VideoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjVideos As Videos
Private mlngID As Long

Public Sub Component(objComponent As Videos)

  Dim objItem As VideoDisplay
  Dim itmList As ListItem
  Dim lngIndex As Long
  
  Set mobjVideos = objComponent
  For lngIndex = 1 To mobjVideos.Count
    With objItem
      Set objItem = mobjVideos.Item(lngIndex)
      Set itmList = _
        lvwItems.ListItems.Add(Key:= _
        Format$(objItem.VideoID) & " K")

      With itmList
        .Text = objItem.Title
        .SubItems(1) = objItem.ReleaseDate
      End With

    End With

  Next

End Sub

Public Property Get VideoID() As Long

  VideoID = mlngID

End Property

Private Sub cmdCancel_Click()

  mlngID = 0
  Hide

End Sub

Private Sub cmdOK_Click()

  On Error Resume Next
  mlngID = Val(lvwItems.SelectedItem.Key)
  Hide

End Sub


