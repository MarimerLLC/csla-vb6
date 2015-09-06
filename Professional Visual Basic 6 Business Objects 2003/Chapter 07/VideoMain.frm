VERSION 5.00
Begin VB.MDIForm VideoMain 
   BackColor       =   &H8000000C&
   Caption         =   "Video Rental System"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7950
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCust 
         Caption         =   "&Customers"
         Begin VB.Menu mnuFileCustSearch 
            Caption         =   "&Search"
         End
         Begin VB.Menu mnuFileCustLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileCustNew 
            Caption         =   "&New Customer"
         End
      End
      Begin VB.Menu mnuFileVideo 
         Caption         =   "&Videos"
         Begin VB.Menu mnuFileVideoSearch 
            Caption         =   "&Search"
         End
         Begin VB.Menu mnuFileVideoLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileVideoNew 
            Caption         =   "&New video"
         End
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInvoiceNew 
         Caption         =   "&New invoice"
      End
      Begin VB.Menu mnuFileCheckIn 
         Caption         =   "Chec&k in tape"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "VideoMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuFileCheckIn_Click()

  Dim strID As String
  Dim objTape As Tape
  
  strID = InputBox$("Scan tape ID", "Check in tape")
  If Val(strID) = 0 Then Exit Sub
  Set objTape = New Tape
  On Error Resume Next
  objTape.Load Val(strID)
  If Err Then
    MsgBox "Invalid tape ID", vbExclamation
    Exit Sub

  End If

  On Error GoTo 0
  With objTape
    If .CheckedOut Then
      .BeginEdit
      .CheckIn
      .ApplyEdit

    Else
      MsgBox "Tape is not checked out", vbExclamation

    End If

  End With

  Set objTape = Nothing

End Sub


Private Sub mnuFileCustNew_Click()
  Dim objCustomer As Customer
  Dim frmCustomer As CustomerEdit
  
  Set objCustomer = New Customer
  Set frmCustomer = New CustomerEdit
  
  frmCustomer.Component objCustomer
  frmCustomer.Show
End Sub

Private Sub mnuFileCustSearch_Click()
  Dim frmSearch As CustomerSearch
  Dim frmList As CustomerList
  Dim frmCustomer As CustomerEdit
  Dim objCustomers As Customers
  Dim objCustomer As Customer

  Set frmSearch = New CustomerSearch
  With frmSearch
    .Show vbModal

    If .OK Then
      Set frmList = New CustomerList
      Set objCustomers = New Customers
      objCustomers.Load .ResultName, .ResultPhone
      With frmList
        .Component objCustomers
        .Show vbModal

        If .CustomerID > 0 Then
          Set frmCustomer = New CustomerEdit
          Set objCustomer = New Customer
          objCustomer.Load .CustomerID
          frmCustomer.Component objCustomer
          frmCustomer.Show
          Set objCustomer = Nothing
        End If

      End With

      Unload frmList
      Set frmList = Nothing
      Set objCustomers = Nothing
    End If

  End With

  Unload frmSearch
  Set frmSearch = Nothing

End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuFileInvoiceNew_Click()
  Dim frmSearch As CustomerSearch
  Dim frmList As CustomerList
  Dim objCustomers As Customers
  Dim objCustomer As Customer
  Dim objInvoice As Invoice
  Dim frmInvoice As InvoiceEdit
  
  Set frmSearch = New CustomerSearch
  With frmSearch
    .Show vbModal

    If .OK Then
      Set frmList = New CustomerList
      Set objCustomers = New Customers
      objCustomers.Load .ResultName, .ResultPhone

      With frmList
        .Component objCustomers
        .Show vbModal
        If .CustomerID > 0 Then
          Set frmInvoice = New InvoiceEdit
          
          Set objCustomer = New Customer
          objCustomer.Load .CustomerID
          Set objInvoice = objCustomer.CreateInvoice
          
          frmInvoice.Component objInvoice
          frmInvoice.Show
        End If

      End With

      Unload frmList
      Set frmList = Nothing
      Set objCustomers = Nothing

    End If

  End With

  Unload frmSearch
  Set frmSearch = Nothing


End Sub

Private Sub mnuFileVideoNew_Click()
  Dim objVideo As Video
  Dim frmVideo As VideoEdit
  
  Set objVideo = New Video
  Set frmVideo = New VideoEdit
  
  frmVideo.Component objVideo
  frmVideo.Show

End Sub

Private Sub mnuFileVideoSearch_Click()

  Dim frmSearch As VideoSearch
  Dim frmList As VideoList
  Dim frmVideo As VideoEdit
  Dim objVideos As Videos
  Dim objVideo As Video
  
  Set frmSearch = New VideoSearch
  With frmSearch
    .Show vbModal

    If .OK Then
      Set frmList = New VideoList
      Set objVideos = New Videos
      objVideos.Load .ResultTitle, .ResultStudio

      With frmList
        .Component objVideos
        .Show vbModal

        If .VideoID > 0 Then
          Set frmVideo = New VideoEdit
          Set objVideo = New Video
          objVideo.Load .VideoID
          frmVideo.Component objVideo
          frmVideo.Show
          Set objVideo = Nothing
        End If

      End With

      Unload frmList
      Set frmList = Nothing
      Set objVideos = Nothing

    End If

  End With

  Unload frmSearch
  Set frmSearch = Nothing

End Sub
