VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColIDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sCol As String
Private m_bDescending As Boolean
Public Property Let colName(ByVal sVal As String)
    sCol = sVal
End Property
Public Property Get colName() As String
    colName = sCol
End Property
Public Property Let isDescending(ByVal bVal As Boolean)
    m_bDescending = bVal
End Property
Public Property Get isDescending() As Boolean
    isDescending = m_bDescending
End Property

