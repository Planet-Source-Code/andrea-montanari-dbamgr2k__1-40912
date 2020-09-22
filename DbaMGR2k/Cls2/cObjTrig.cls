VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObjTrig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sName As String
Private m_sFileName As String

Public Property Let oName(ByVal sVal As String)
    m_sName = sVal
End Property
Public Property Get oName() As String
    oName = m_sName
End Property
Public Property Let oFileName(ByVal sVal As String)
    m_sFileName = sVal
End Property
Public Property Get oFileName() As String
    oFileName = m_sFileName
End Property

