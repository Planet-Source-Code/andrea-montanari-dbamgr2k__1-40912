VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private iSpid As Integer
Private coll As Collection
Public Property Let Spid(ByVal iVal As Integer)
    iSpid = iVal
End Property
Public Property Get Spid() As Integer
    Spid = iSpid
End Property
Public Sub AddItem(obj As clsLock)
    coll.add obj, "*" & obj.Key()
End Sub
Public Property Get Count() As Integer
    Count = coll.Count
End Property
Public Property Get Item(ByVal vKey As Variant) As clsLock
    Set Item = coll.Item(vKey)
End Property

Private Sub Class_Initialize()
    Set coll = New Collection
End Sub

Private Sub Class_Terminate()
    Set coll = Nothing
End Sub
