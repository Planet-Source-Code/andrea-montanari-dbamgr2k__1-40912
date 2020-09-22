VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKeyCasc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_bUpdCascade As Boolean
Private m_bDelCascade As Boolean

Public Property Get UpdateCascade() As Boolean
    UpdateCascade = m_bUpdCascade
End Property
Public Property Get DeleteCascade() As Boolean
    DeleteCascade = m_bDelCascade
End Property

Public Sub InitStruct(ByVal sDbName As String, ByVal sRefConstraint As String, ByRef bUpdateCascade As Boolean, ByRef bDeleteCascade As Boolean)

    Dim sSql As String
    Dim oDb As SQLDMO.Database2
    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    
    sSql = Replace("select OBJECTPROPERTY ( object_id('1%'),'CnstIsUpdateCascade') as 'UpdCascade' , OBJECTPROPERTY ( object_id('1%'),'CnstIsDeleteCascade') as 'DelCascade'", "1%", sRefConstraint)
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    Set oQry = oDb.ExecuteWithResults(sSql, Len(sSql))
    On Local Error GoTo 0
    
    If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
    Set oQry = Nothing
    
    If IsArray(vVar) Then
        m_bUpdCascade = vVar(0, 0)
        m_bDelCascade = vVar(1, 0)
    End If
    Set oDb = Nothing
    bUpdateCascade = m_bUpdCascade
    bDeleteCascade = m_bDelCascade
    'select OBJECTPROPERTY ( object_id('FK_Tabella4_Tabella3'),'CnstIsUpdateCascade') as 'UpdCascade' , OBJECTPROPERTY ( object_id('FK_Tabella4_Tabella3'),'CnstIsDeleteCascade') as 'DelCascade'
End Sub
