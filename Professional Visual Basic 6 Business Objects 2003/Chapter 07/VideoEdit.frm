VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VideoEdit 
   Caption         =   "Video"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      Caption         =   "Tapes"
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   8055
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwTapes 
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2355
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TapeID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Purchased"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rented"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cboRating 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtStudio 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox txtRelease 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Rating"
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   255
      Left            =   135
      TabIndex        =   10
      Top             =   1755
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Studio"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Release date"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "VideoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjVideo As Video
Attribute mobjVideo.VB_VarHelpID = -1

Public Sub Component(VideoObject As Video)

  Set mobjVideo = VideoObject

End Sub

Private Sub cmdAdd_Click()
  Dim frmTape As TapeEdit
  
  Set frmTape = New TapeEdit
  frmTape.Component mobjVideo.Tapes.Add
  frmTape.Show vbModal
  LoadTapes
End Sub

Private Sub cmdApply_Click()

  mobjVideo.ApplyEdit
  mobjVideo.BeginEdit
  LoadTapes

End Sub

Private Sub cmdCancel_Click()

  mobjVideo.CancelEdit
  Unload Me


End Sub

Private Sub cmdEdit_Click()

  Dim frmTape As TapeEdit
  
  Set frmTape = New TapeEdit
  frmTape.Component _
    mobjVideo.Tapes(Val(lvwTapes.SelectedItem.Key))
  frmTape.Show vbModal
  LoadTapes

End Sub

Private Sub cmdOK_Click()

  mobjVideo.ApplyEdit
  Unload Me


End Sub

Private Sub cmdRemove_Click()
  mobjVideo.Tapes.Remove Val(lvwTapes.SelectedItem.Key)
  LoadTapes
End Sub

Private Sub Form_Load()

  mflgLoading = True
  With mobjVideo
     EnableOK .IsValid
    If .IsNew Then
      Caption = "Video [(new)]"

    Else
      Caption = "Video [" & .Title & "]"

    End If

    txtTitle = .Title
    txtRelease = .ReleaseDate
    txtStudio = .Studio
    ' LoadCombo cboCategory, .Categories
    ' cboCategory.Text = .Category
    ' LoadCombo cboRating, .Ratings
    ' cboRating.Text = .Rating
    .BeginEdit
  End With
  LoadTapes
  mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjVideo_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub txtTitle_Change()

  If Not mflgLoading Then _
    TextChange txtTitle, mobjVideo, "Title"

End Sub

Private Sub txtTitle_LostFocus()

  txtTitle = TextLostFocus(txtTitle, mobjVideo, "Title")

End Sub

Private Sub txtStudio_Change()

  If Not mflgLoading Then _
    TextChange txtStudio, mobjVideo, "Studio"

End Sub

Private Sub txtStudio_LostFocus()

  txtStudio = TextLostFocus(txtStudio, mobjVideo, "Studio")

End Sub

Private Sub txtRelease_Change()

  If Not mflgLoading Then _
    TextChange txtRelease, mobjVideo, "ReleaseDate"

End Sub

Private Sub txtRelease_LostFocus()

  txtRelease = TextLostFocus(txtRelease, mobjVideo, "ReleaseDate")

End Sub

Private Sub LoadCombo(Combo As ComboBox, List As TextList)

  Dim vntItem As Variant
  
  With Combo
    .Clear
    For Each vntItem In List
      .AddItem vntItem
    Next
    If .ListCount > 0 Then .ListIndex = 0
  End With

End Sub

Private Sub cboCategory_Click()

  If mflgLoading Then Exit Sub
  mobjVideo.Category = cboCategory.Text

End Sub

Private Sub cboRating_Click()

  If mflgLoading Then Exit Sub
  mobjVideo.Rating = cboRating.Text

End Sub

Private Sub LoadTapes()

  Dim objTape As Tape
  Dim itmList As ListItem
  Dim lngIndex As Long
  
  lvwTapes.ListItems.Clear
  For lngIndex = 1 To mobjVideo.Tapes.Count
    Set itmList = lvwTapes.ListItems.Add _
      (Key:=Format$(lngIndex) & "K")
    Set objTape = mobjVideo.Tapes(lngIndex)

    With itmList
      If objTape.IsNew Then
        .Text = "(new)"

      Else
        .Text = objTape.TapeID

      End If

      If objTape.IsDeleted Then .Text = .Text & " (d)"
      .SubItems(1) = objTape.DateAcquired
      .SubItems(2) = IIf(objTape.CheckedOut, "Yes", "No")
    End With

  Next

End Sub


