VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sRole As String
Private bIsMember As Boolean
Public Property Let Role(ByVal sVal As String)
    sRole = sVal
End Property
Public Property Get Role() As String
    Role = sRole
End Property
Public Property Let IsMember(ByVal bVal As Boolean)
    bIsMember = bVal
End Property
Public Property Get IsMember() As Boolean
    IsMember = bIsMember
End Property

