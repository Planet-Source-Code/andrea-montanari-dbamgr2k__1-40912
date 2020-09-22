VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDetacDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sFileName As String
Private bFileExist As Boolean
Private m_Fi As CFileInfo

Public Property Set oFileFinfo(ByVal oFI As CFileInfo)
    Set m_Fi = oFI
End Property
Public Property Get oFileFinfo() As CFileInfo
    Set oFileFinfo = m_Fi
End Property
Public Property Let PhysicalName(ByVal sVal As String)
    sFileName = sVal
End Property
Public Property Get PhysicalName() As String
    PhysicalName = sFileName
End Property

Public Property Let FileExist(ByVal bVal As Boolean)
    bFileExist = bVal
End Property
Public Property Get FileExist() As Boolean
    FileExist = bFileExist
End Property

Private Sub Class_Terminate()
    Set m_Fi = Nothing
End Sub
