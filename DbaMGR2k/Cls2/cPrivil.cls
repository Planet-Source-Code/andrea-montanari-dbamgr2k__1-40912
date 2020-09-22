VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum iPrivType
    NotSet = 0
    Grant = 1
    Deny = 2
    Unavailable = 3
    GrantCol = 4
    DenyCol = 5
    GrantDenyCol = 9
End Enum
Public Enum odbType
    oNull = 0
    oTbl = SQLDMOObj_UserTable
    oView = SQLDMOObj_View
    oSP = SQLDMOObj_StoredProcedure
    oDb = SQLDMOObj_Database
    oFunc = SQLDMOObj_UserDefinedFunction
End Enum
Private objType As odbType
Private sName As String
Private sGroup As String
Private bIsGroup As Boolean
Private bIsChanged As Boolean
Private bIsTableFunction As Boolean

Private iPrivil(7) As iPrivType
Public Property Get IsChanged() As Boolean
    IsChanged = bIsChanged
End Property
Public Property Let IsChanged(ByVal bVal As Boolean)
    bIsChanged = bVal
End Property

Public Property Get isGroup() As Boolean
    isGroup = bIsGroup
End Property
Public Property Let isGroup(ByVal bVal As Boolean)
    bIsGroup = bVal
End Property

Public Property Get IsSet() As Boolean
    'IsSet = (iPrivil(0) <> NotSet) Or (iPrivil(1) <> NotSet) Or (iPrivil(2) <> NotSet) Or (iPrivil(3) <> NotSet) Or (iPrivil(4) <> NotSet) Or (iPrivil(5) <> NotSet)
    'IsSet = (iPrivil(0) > NotSet And iPrivil(0) < Unavailable) Or (iPrivil(1) > NotSet And iPrivil(1) < Unavailable) Or (iPrivil(2) > NotSet And iPrivil(2) < Unavailable) Or (iPrivil(3) > NotSet And iPrivil(3) < Unavailable) Or (iPrivil(4) > NotSet And iPrivil(4) < Unavailable) Or (iPrivil(5) > NotSet And iPrivil(5) < Unavailable) Or (iPrivil(6) > NotSet And iPrivil(6) < Unavailable) Or (iPrivil(7) > NotSet And iPrivil(7) < Unavailable)
    
    IsSet = (iPrivil(0) > NotSet And iPrivil(0) <> Unavailable) Or (iPrivil(1) > NotSet And iPrivil(1) <> Unavailable) Or (iPrivil(2) > NotSet And iPrivil(2) <> Unavailable) Or (iPrivil(3) > NotSet And iPrivil(3) <> Unavailable) Or (iPrivil(4) > NotSet And iPrivil(4) <> Unavailable) Or (iPrivil(5) > NotSet And iPrivil(5) <> Unavailable) Or (iPrivil(6) > NotSet And iPrivil(6) <> Unavailable) Or (iPrivil(7) > NotSet And iPrivil(7) <> Unavailable)
End Property
Public Property Get oType() As odbType
    oType = objType
End Property
Public Property Let oType(ByVal lVal As odbType)
    objType = lVal
End Property
Public Property Get oName() As String
    oName = sName
End Property
Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Get oGroup() As String
    oGroup = sGroup
End Property
Public Property Let oGroup(ByVal sVal As String)
    sGroup = sVal
End Property

Public Property Let ItemVal(ByVal iItem As Integer, ByVal iVal As iPrivType)
    iPrivil(iItem) = iVal
End Property
Public Property Get ItemVal(ByVal iItem As Integer) As iPrivType
    ItemVal = iPrivil(iItem)
End Property
Public Property Get IsTableFunction() As Boolean
    IsTableFunction = bIsTableFunction
End Property
Public Property Let IsTableFunction(ByVal bVal As Boolean)
    bIsTableFunction = bVal
End Property

Public Sub ResetSet()
    Dim i As Integer
    For i = 0 To UBound(iPrivil)
        iPrivil(i) = NotSet
    Next
    bIsChanged = False
End Sub
