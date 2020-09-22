VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cViewPrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sName As String
Private iTop As Integer
Private bOpt(4) As Boolean
Private bOptEnab(2) As Boolean

Private bIsGroup As Boolean
Private bExt As Boolean
Private iOpt As Integer

Public Property Let oName(ByVal sVal As String)
    sName = sVal
End Property
Public Property Get oName() As String
    oName = sName
End Property

Public Property Let TopRec(ByVal iVal As Integer)
    iTop = iVal
End Property
Public Property Get TopRec() As Integer
    TopRec = iTop
End Property
Public Property Let bOption(ByVal iElem As Integer, bVal As Boolean)
    bOpt(iElem) = bVal
End Property
Public Property Get bOption(ByVal iElem As Integer) As Boolean
    bOption = bOpt(iElem)
End Property
Public Property Let bOptionEnab(ByVal iElem As Integer, bVal As Boolean)
    bOptEnab(iElem) = bVal
End Property
Public Property Get bOptionEnab(ByVal iElem As Integer) As Boolean
    bOptionEnab = bOptEnab(iElem)
End Property

Public Property Let Extention(bVal As Boolean)
    bExt = bVal
End Property
Public Property Get Extention() As Boolean
    Extention = bExt
End Property
Public Property Let ExtOpt(ByVal iVal As Integer)
    iOpt = iVal
End Property
Public Property Get ExtOpt() As Integer
    ExtOpt = iOpt
End Property

Public Property Let isGroup(ByVal bVal As Boolean)
    bIsGroup = bVal
End Property
Public Property Get isGroup() As Boolean
    isGroup = bIsGroup
End Property

Private Sub Class_Initialize()
    iOpt = -1
End Sub
