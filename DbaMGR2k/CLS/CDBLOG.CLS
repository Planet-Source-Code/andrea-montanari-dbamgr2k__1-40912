VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sDBName As String
Private bCanDB As Boolean

Private collRole As Collection
Public Property Get RoleCount() As Integer
    RoleCount = collRole.Count
End Property
Public Property Let CanDB(ByVal bVal As Boolean)
    bCanDB = bVal
End Property
Public Property Get CanDB() As Boolean
    CanDB = bCanDB
End Property

Public Property Let DBName(ByVal sVal As String)
    sDBName = sVal
End Property
Public Property Get DBName() As String
    DBName = sDBName
End Property
Public Sub AddRole(ByVal sRole As String, ByVal bCan As Boolean)
    Dim cRole As New cDBItem
    cRole.Role() = sRole
    cRole.IsMember() = bCan
    collRole.Add cRole, sRole
    Set cRole = Nothing
End Sub

Private Sub Class_Initialize()
    Set collRole = New Collection
End Sub

Private Sub Class_Terminate()
    Set collRole = Nothing
End Sub
Public Property Get GetRole(ByVal vItem As Variant) As cDBItem
    Set GetRole = collRole.Item(vItem)
End Property

