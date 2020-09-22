VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sName As String
Private sGroup As String
Private bSkip As Boolean
Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Get oName() As String
    oName = sName
End Property
Public Property Let oGroup(ByVal sVal As String)
    sGroup = sVal
End Property
Public Property Get oGroup() As String
    oGroup = sGroup
End Property
Public Property Let Skip(ByVal bVal As Boolean)
    bSkip = bVal
End Property
Public Property Get Skip() As Boolean
    Skip = bSkip
End Property

