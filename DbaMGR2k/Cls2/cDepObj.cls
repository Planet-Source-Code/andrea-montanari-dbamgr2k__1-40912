VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDepObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iObjType As SQLDMO_OBJECT_TYPE
Private sOwner As String
Private sName As String
Private lngSequence As Long

Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Get oName() As String
    oName = sName
End Property
Public Property Let oOwner(ByVal sVal As String)
    sOwner = sVal
End Property
Public Property Get oOwner() As String
    oOwner = sOwner
End Property
Public Property Let oSequence(ByVal lVal As Long)
    lngSequence = lVal
End Property
Public Property Get oSequence() As Long
    oSequence = lngSequence
End Property
Public Property Let oType(ByVal iVal As SQLDMO_OBJECT_TYPE)
    iObjType = iVal
End Property
Public Property Get oType() As SQLDMO_OBJECT_TYPE
    oType = iObjType
End Property

