VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cErrMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum iErrDbaMgr
    errNum = 0
    errDescription
    errSource
    errNative
End Enum
Private avntErr() As Variant

Public Property Get Count() As Integer

    If Not ArrayVuoto(avntErr) Then
        Count = UBound(avntErr, 2) + 1
    End If
End Property

Public Sub PushError(ByVal Err_Number As Long, ByVal Err_Description As String, ByVal Err_Source As String, Optional ByVal Err_Native As Long = 0)

    Dim i As Integer
    
    If ArrayVuoto(avntErr) Then
        i = 0
    Else
        i = UBound(avntErr, 2) + 1
    End If
    ReDim Preserve avntErr(errNative, i)
    
    avntErr(errNum, i) = Err_Number
    avntErr(errDescription, i) = Err_Description
    avntErr(errSource, i) = Err_Source
    avntErr(errNative, i) = Err_Native
    
End Sub
Public Function PopError(ByVal iItem As Integer) As Variant()
    
    Dim vRet(errNative) As Variant
    
    If Not ArrayVuoto(avntErr) Then
        If iItem <= UBound(avntErr, 2) + 1 Then
            vRet(errNum) = avntErr(errNum, iItem - 1)
            vRet(errDescription) = avntErr(errDescription, iItem - 1)
            vRet(errSource) = avntErr(errSource, iItem - 1)
            vRet(errNative) = avntErr(errNative, iItem - 1)
        End If
    End If
    PopError = vRet
    
End Function

Public Function ArrayVuoto(avnt As Variant) As Boolean

    Dim bRet As Boolean
    Dim i As Integer
    Dim localVnt() As Variant
    
    On Local Error Resume Next
    i = UBound(avnt, 2)
    bRet = Err.Number
    On Local Error GoTo 0
    If bRet Then avnt = localVnt
    
    ArrayVuoto = bRet
    
End Function

