VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Status(ByVal iPerc As Integer)
Public Event Init(ByVal szMsg As String)


Public Sub WriteRes2Ini(ByVal bItalian As Boolean, Optional ByVal bForce As Boolean = False)

    Dim i As Integer
    Dim iOff As Integer, iInit As Integer, iKey As Integer
    Dim sFile As String
    Dim sBuf As String
    Dim bWrite As Boolean
            
    If bItalian Then
        sFile = App.Path & "\Italiano.ln2"
        iInit = 0
        iOff = 1999
    Else
        sFile = App.Path & "\English.ln2"
        iInit = 2000
        iOff = 3999
    End If
           
        
    On Local Error Resume Next
    If Len(Dir$(sFile)) = 0 Then
        bWrite = True
    End If
    If bForce Then bWrite = True
    If Not bWrite And Err.Number <> 0 Then bWrite = True
    If bWrite Then Kill sFile
    On Local Error GoTo 0
    
    If bWrite Then
        Dim f As New fPercent
        f.InitializeLanguage Me
        
        f.Show
        RaiseEvent Init(sFile)
        RaiseEvent Status(0)
        On Local Error Resume Next
        For i = iInit To iOff
            iKey = IIf(i >= 2000, i - 2000, i)
            RaiseEvent Status(iKey \ 20)
            sBuf = ""
            sBuf = VB.LoadResString(i)
            If Len(sBuf) Then
                iKey = IIf(i >= 2000, i - 2000, i)
                WriteIni kLng, Format$(iKey), sBuf, sFile
            End If
        Next
        On Local Error GoTo 0
        Unload f
        Set f = Nothing
    End If
End Sub

