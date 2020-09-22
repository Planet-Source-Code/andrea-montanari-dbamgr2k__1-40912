VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cViewIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sObjName As String
Private sObjAlias As String
Private sObjOwner As String
Private bTbl As Boolean
Private iCtrl As Integer

Public Property Let sName(ByVal sVal As String)
    sObjName = sVal
End Property
Public Property Get sName() As String
    sName = sObjName
End Property
Public Property Let isTbl(ByVal bVal As Boolean)
    bTbl = bVal
End Property
Public Property Get isTbl() As Boolean
    isTbl = bTbl
End Property
Public Property Let iControl(ByVal iVal As Integer)
    iCtrl = iVal
End Property
Public Property Get iControl() As Integer
    iControl = iCtrl
End Property
Public Property Let sAlias(ByVal sVal As String)
    sObjAlias = sVal
End Property
Public Property Get sAlias() As String
    sAlias = sObjAlias
End Property
Public Property Get sOwner() As String
    sOwner = sObjOwner
End Property
Public Property Let sOwner(ByVal sVal As String)
    sObjOwner = sVal
End Property

