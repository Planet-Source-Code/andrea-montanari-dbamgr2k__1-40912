VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrpO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cType As optScript
Private sName As String
Private sOwner As String


Public Property Let oType(ByVal iVal As optScript)
    cType = iVal
End Property
Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Let oOwner(ByVal sVal As String)
    sOwner = sVal
End Property
Public Property Get oType() As optScript
    oType = cType
End Property
Public Property Get oName() As String
    oName = sName
End Property
Public Property Get oOwner() As String
    oOwner = sOwner
End Property

