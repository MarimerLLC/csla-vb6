VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Load(ByVal SSN As String, Person As Object)

  Dim rsPerson As Recordset
  Dim strConnect As String
  Dim strSQL As String
  Dim udtPerson As PersonProps
  Dim udtBuffer As PersonData

  strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
    "Persist Security Info=False;" & _
    "Data Source=C:\Wrox\VB6 Pro Objects\Person.mdb"
  strSQL = "SELECT * FROM Person WHERE SSN='" & SSN & "'"
  Set rsPerson = New Recordset
  rsPerson.Open strSQL, strConnect

  With rsPerson
    If Not .EOF And Not .BOF Then
      udtPerson.SSN = .Fields("SSN")
      udtPerson.Name = .Fields("Name")
      udtPerson.Birthdate = .Fields("Birthdate")
      
      LSet udtBuffer = udtPerson
      Person.SetState udtBuffer.Buffer

    Else
      rsPerson.Close
      Err.Raise vbObjectError + 1002, "Person", "SSN not on file"

    End If

  End With

  rsPerson.Close

End Sub

Public Sub Save(Person As Object)

  Dim rsPerson As Recordset
  Dim strConnect As String
  Dim strSQL As String
  Dim udtPerson As PersonProps
  Dim udtBuffer As PersonData
  
  udtBuffer.Buffer = Person.GetState
  LSet udtPerson = udtBuffer
  
  strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
    "Persist Security Info=False;" & _
    "Data Source=C:\Wrox\VB6 Pro Objects\Person.mdb"
  strSQL = "SELECT * FROM Person WHERE SSN='" & udtPerson.SSN & "'"
  Set rsPerson = New Recordset
  rsPerson.Open strSQL, strConnect, , adLockPessimistic

  With rsPerson
    If Person.IsNew Then .AddNew
     
    .Fields("SSN") = udtPerson.SSN
    .Fields("Name") = udtPerson.Name
    .Fields("Birthdate") = udtPerson.Birthdate
    .Update

  End With

  rsPerson.Close

End Sub

Public Sub Delete(SSN As String)

  Dim cnPerson As Connection
  Dim strConnect As String
  Dim strSQL As String
  
  strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
    "Persist Security Info=False;" & _
    "Data Source=C:\Wrox\VB6 Pro Objects\Person.mdb"
  strSQL = "DELETE * FROM Person WHERE SSN='" & SSN & "'"
  Set cnPerson = New Connection
  cnPerson.Open strConnect
  cnPerson.Execute strSQL
  cnPerson.Close
  Set cnPerson = Nothing

End Sub


