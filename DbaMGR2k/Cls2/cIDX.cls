VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum idxType
    db_idx = 0
    db_key = 1
    db_check = 2
End Enum

Private bIsPrimaryKey As Boolean
Private sName As String
Private sFields() As String
Private bDescending() As Boolean
Private Attrib As Long
Private FillFactor As Long
Private FileGroup As String
Private NoRecompute  As Boolean
Private Table As String
Private ReferencedTable As String
Private Checked As Boolean
Private ExcludeReplication As Boolean
Private CheckText As String
Private bUpdateCascade As Boolean
Private bDeleteCascade As Boolean

Private sFieldsEX() As String
Private colFld As Collection
Private colFldEx As Collection

Private Sub Class_Initialize()
    Set colFld = New Collection
    ReDim sFieldsEX(0) As String
    ReDim sFields(0) As String
    ReDim bDescending(0) As Boolean
End Sub

Private Sub Class_Terminate()
    Set colFld = Nothing
    Set colFldEx = Nothing
End Sub
Public Sub LoadFld(ByVal sVal As String, ByVal bIsDescending As Boolean, ByVal iItem As Integer)
    If UBound(sFields) < iItem Then
        ReDim Preserve sFields(iItem) As String
        ReDim Preserve bDescending(iItem) As Boolean
    End If
    sFields(iItem) = sVal
    bDescending(iItem) = bIsDescending
End Sub
Public Function GetFld(ByVal iElem As Integer, ByRef bIsDescending As Boolean) As String
    GetFld = sFields(iElem)
    bIsDescending = bDescending(iElem)
End Function

'Public Function GetFld(ByVal iElem As Integer) As String
'    GetFld = sFields(iElem)
'End Function
Public Sub LoadFldEx(ByVal sVal As String, ByVal iItem As Integer)
    If UBound(sFieldsEX) < iItem Then ReDim Preserve sFieldsEX(iItem) As String
    sFieldsEX(iItem) = sVal
End Sub
Public Function GetFldEx(ByVal iElem As Integer) As String
    GetFldEx = sFieldsEX(iElem)
End Function
Public Property Get oType() As idxType
    oType = Attrib
End Property
Public Property Let oType(ByVal lVal As idxType)
    Attrib = lVal
End Property
Public Property Get oFillFactor() As Long
    oFillFactor = FillFactor
End Property
Public Property Let oFillFactor(ByVal lVal As Long)
    FillFactor = lVal
End Property
Public Property Get idxNoRecompute() As Boolean
    idxNoRecompute = NoRecompute
End Property
Public Property Let idxNoRecompute(ByVal bVal As Boolean)
    NoRecompute = bVal
End Property
Public Property Get oExcludeReplication() As Boolean
    oExcludeReplication = ExcludeReplication
End Property
Public Property Let oExcludeReplication(ByVal bVal As Boolean)
    ExcludeReplication = bVal
End Property
Public Property Get relReferencedTable() As String
    relReferencedTable = ReferencedTable
End Property
Public Property Let relReferencedTable(ByVal sVal As String)
    ReferencedTable = sVal
End Property
Public Property Get chkCheckText() As String
    chkCheckText = CheckText
End Property
Public Property Let chkCheckText(ByVal sVal As String)
    CheckText = sVal
End Property
Public Property Get oName() As String
    oName = sName
End Property
Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Get oFileGroup() As String
    oFileGroup = FileGroup
End Property
Public Property Let oFileGroup(ByVal sVal As String)
    FileGroup = sVal
End Property
Public Property Get oChecked() As Boolean
    oChecked = Checked
End Property
Public Property Let oChecked(ByVal bVal As Boolean)
    Checked = bVal
End Property
Public Property Get oAttrib() As Long
    oAttrib = Attrib
End Property
Public Property Let oAttrib(ByVal lVal As Long)
    Attrib = lVal
End Property
Public Property Get oFldCount() As Integer
    oFldCount = UBound(sFields)
End Property
Public Property Get oFldCountEx() As Integer
    oFldCountEx = UBound(sFieldsEX)
End Property
Public Property Get relTable() As String
    relTable = Table
End Property
Public Property Let relTable(ByVal sVal As String)
    Table = sVal
End Property

Public Property Get oIsPrimaryKey() As Boolean
    oIsPrimaryKey = bIsPrimaryKey
End Property
Public Property Let oIsPrimaryKey(ByVal bVal As Boolean)
    bIsPrimaryKey = bVal
End Property
Public Property Get UpdateCascade() As Boolean
    UpdateCascade = bUpdateCascade
End Property
Public Property Let UpdateCascade(ByVal bVal As Boolean)
    bUpdateCascade = bVal
End Property
Public Property Get DeleteCascade() As Boolean
    DeleteCascade = bDeleteCascade
End Property
Public Property Let DeleteCascade(ByVal bVal As Boolean)
    bDeleteCascade = bVal
End Property

