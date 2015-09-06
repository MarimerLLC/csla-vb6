VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ItemTypes
  ITEM_TAPE = 0
  ITEM_FEE = 1
End Enum

Public Sub BeginEdit()

End Sub

Public Sub CancelEdit()

End Sub

Public Sub ApplyEdit()

End Sub

Public Sub Delete()

End Sub

Public Property Get IsValid() As Boolean

End Property

Public Property Get IsNew() As Boolean

End Property

Public Property Get IsDirty() As Boolean

End Property

Public Property Get IsDeleted() As Boolean

End Property

Public Property Get Total() As Double

End Property

Public Property Get ItemDescription() As String

End Property

Public Property Get ItemType() As ItemTypes

End Property

