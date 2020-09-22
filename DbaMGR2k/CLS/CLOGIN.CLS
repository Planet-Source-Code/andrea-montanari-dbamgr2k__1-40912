VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private collDB As Collection

Private Sub Class_Initialize()
    Set collDB = New Collection
End Sub

Private Sub Class_Terminate()
    Set collDB = Nothing
End Sub
Public Function AddDB(ByVal sDB As String, ByVal bCan As Boolean) As Boolean
    
    Dim i As Integer
    Dim cDBo As New cDBLog
    
    cDBo.DBName() = sDB
    cDBo.CanDB() = bCan
    On Local Error GoTo ErrTrap
    With objServer.Databases(sDB).DatabaseRoles
        For i = 1 To .Count
            If .Item(i).IsFixedRole Then
                cDBo.AddRole .Item(i).Name, IIf(.Item(i).Name = "public", True, False)
            End If
        Next
        For i = 1 To .Count
            If Not .Item(i).IsFixedRole Then
                If Not .Item(i).AppRole Then cDBo.AddRole .Item(i).Name, False
            End If
        Next
    End With
    
    collDB.Add cDBo, sDB
    Set cDBo = Nothing
    AddDB = True
    Exit Function
ErrTrap:
    AddDB = False
    Set cDBo = Nothing
End Function
Public Property Get DBCount() As Integer
    DBCount = collDB.Count
End Property
Public Property Get DbRole(ByVal vItem As Variant) As cDBLog
    On Local Error Resume Next
    Set DbRole = collDB.Item(vItem)
End Property

