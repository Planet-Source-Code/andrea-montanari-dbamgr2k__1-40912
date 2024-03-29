VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRelDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sRelName As String
Private sTbName As String
Private iPtrCtrl As Integer
Private oCollObj As Collection

Private Sub Class_Initialize()
    Set oCollObj = New Collection
End Sub

Private Sub Class_Terminate()
    Set oCollObj = Nothing
End Sub
Public Property Let RelName(ByVal sVal As String)
    sRelName = sVal
End Property
Public Property Get RelName() As String
    RelName = sRelName
End Property
Public Property Let TbName(ByVal sVal As String)
    sTbName = sVal
End Property
Public Property Get TbName() As String
    TbName = sTbName
End Property
Public Property Set CollItems(coll As Collection)
    Set oCollObj = coll
End Property
Public Property Get CollItems() As Collection
    Set CollItems = oCollObj
End Property

Public Property Let PtrCtrl(ByVal iVal As Integer)
    iPtrCtrl = iVal
End Property
Public Property Get PtrCtrl() As Integer
    PtrCtrl = iPtrCtrl
End Property
