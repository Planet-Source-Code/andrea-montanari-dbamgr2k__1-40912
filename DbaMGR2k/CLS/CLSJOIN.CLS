VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum JoinT
    jInner = k_ResJoin '0
    jLeft = k_ResLJoin '1
    jRight = k_ResRJoin '2
    jFull = k_ResFJoin '3
End Enum
Public Enum JoinOp
    jEqual = 0 '"="
    jDiffer = 1 '"<>"
    jMinor = 2 ' "<"
    jMinEq = 3 '"<="
    jSup = 4 '">"
    jSupEq = 5 '"=>"
End Enum
Private sId As String
Private sField As String
Private sBaseTbl As String
Private sAliasBaseTbl As String
Private sRelatedTB As String
Private sAliasRelatedTB As String
Private sRelatedField As String
Private iCtrl As Integer
Private iBaseCtrl As Integer
Private iRefCtrl As Integer
Private iJoinType As JoinT
Private iJoinOp As JoinOp
Public Property Let Id(ByVal sVal As String)
    sId = sVal
End Property
Public Property Get Id() As String
    Id = sId
End Property

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
Public Property Let BaseTbl(ByVal sVal As String)
    sBaseTbl = sVal
End Property
Public Property Get BaseTbl() As String
    BaseTbl = sBaseTbl
End Property
Public Property Let AliasBaseTbl(ByVal sVal As String)
    sAliasBaseTbl = sVal
End Property
Public Property Get AliasBaseTbl() As String
    AliasBaseTbl = sAliasBaseTbl
End Property
Public Property Let AliasRelatedTbl(ByVal sVal As String)
    sAliasRelatedTB = sVal
End Property
Public Property Get AliasRelatedTbl() As String
    AliasRelatedTbl = sAliasRelatedTB
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
Public Property Get iControl() As Integer
    iControl = iCtrl
End Property
Public Property Let iBaseControl(ByVal iVal As Integer)
    iBaseCtrl = iVal
End Property
Public Property Get iBaseControl() As Integer
    iBaseControl = iBaseCtrl
End Property

Public Property Let iRefControl(ByVal iVal As Integer)
    iRefCtrl = iVal
End Property
Public Property Get iRefControl() As Integer
    iRefControl = iRefCtrl
End Property
Public Property Let JoinType(ByVal iVal As JoinT)
    iJoinType = iVal
End Property
Public Property Get JoinType() As JoinT
    JoinType = iJoinType
End Property

Private Sub Class_Initialize()
    iJoinOp = jEqual
End Sub
Public Property Let JoinOperator(ByVal iVal As JoinOp)
    iJoinOp = iVal
End Property
Public Property Get JoinOperator() As JoinOp
    JoinOperator = iJoinOp
End Property

