VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sField As String
Private sRelatedTB As String
Private sRelatedField As String
Private iCtrl As Integer

Public Property Let Field(ByVal sVal As String)
    sField = sVal
End Property
Public Property Get Field() As String
    Field = sField
End Property
Public Property Let RelatedField(ByVal sVal As String)
    sRelatedField = sVal
End Property
Public Property Get RelatedField() As String
    RelatedField = sRelatedField
End Property
Public Property Let RelatedTbl(ByVal sVal As String)
    sRelatedTB = sVal
End Property
Public Property Get RelatedTbl() As String
    RelatedTbl = sRelatedTB
End Property
Public Property Let iControl(ByVal iVal As Integer)
    iCtrl = iVal
End Property
Public Property Get iControl() As String
    iControl = iCtrl
End Property

