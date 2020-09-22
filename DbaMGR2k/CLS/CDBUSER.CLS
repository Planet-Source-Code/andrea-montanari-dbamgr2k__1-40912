VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sName As String
Private sLoginName As String
Private bIsMapped As Boolean
Private bIsChanged As Boolean
Public Property Let Name(ByVal sVal As String)
    sName = Trim$(sVal)
End Property
Public Property Get Name() As String
    Name = sName
End Property
Public Property Let LogName(ByVal sVal As String)
    sLoginName = Trim$(sVal)
End Property
Public Property Get LogName() As String
    LogName = sLoginName
End Property
Public Property Let Mapped(ByVal bVal As Boolean)
    bIsMapped = bVal
End Property
Public Property Get Mapped() As Boolean
    Mapped = bIsMapped
End Property
Public Property Let Changed(ByVal bVal As Boolean)
    bIsChanged = bVal
End Property
Public Property Get Changed() As Boolean
    Changed = bIsChanged
End Property


