VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrName As String
Private mflgWritable As Boolean

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(Value As String)
  mstrName = Value
End Property

Public Property Get Writable() As Boolean
  Writable = mflgWritable
End Property

Public Property Let Writable(Value As Boolean)
  mflgWritable = Value
End Property

