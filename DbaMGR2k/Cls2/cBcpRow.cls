VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBcpRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvaroName As String 'Copia locale.
Private mvaroType As String 'Copia locale.
Private mvaroSize As Long 'Copia locale.
Private mSize2 As Long 'Copia locale.
Private msSep As String 'Copia locale.
Private m_sCollation As String

Private miPos As Integer
Private bString As Boolean

Public Property Get IsString() As Boolean
    IsString = bString
End Property

Public Property Let Separ(ByVal vData As String)
    msSep = vData
End Property
Public Property Get Separ() As String
    Separ = msSep
End Property
Public Property Let iPos(ByVal vData As Integer)
    miPos = vData
End Property

Public Property Get iPos() As Integer
    iPos = miPos
End Property
Public Property Let oSize2(ByVal vData As Long)
    mSize2 = vData
End Property

Public Property Get oSize2() As Long
    oSize2 = mSize2
End Property

Public Property Let oSize(ByVal vData As Long)
    mvaroSize = vData
End Property

Public Property Get oSize() As Long
    oSize = mvaroSize
End Property

Public Property Let oType(ByVal vData As String)
    Dim sBaseType As String
    Dim iPos As Integer
    
    sBaseType = vData
    mvaroType = vData
    iPos = InStr(sBaseType, "{")
    If iPos <> 0 Then
        sBaseType = Mid$(sBaseType, iPos + 1)
        sBaseType = Trim$(Replace(sBaseType, "}", ""))
        sBaseType = LCase$(sBaseType)
    End If
    Select Case Trim$(sBaseType)
        Case "decimal", "numeric"
            bString = False
        Case "float", "money", "real", "smallmoney"
            bString = False
        Case "int", "smallint", "tinyint"
            bString = False
        Case "char", "varchar", "nchar", "nvarchar"
            bString = True
        Case "datetime", "smalldatetime", "timestamp"
            bString = False
        Case Else
            bString = False
    End Select
    If Not bString Then
        m_sCollation = ""
    Else
        m_sCollation = objServer.Collation
    End If
    
End Property

Public Property Get oType() As String
    oType = mvaroType
End Property

Public Property Let oName(ByVal vData As String)
    mvaroName = vData
End Property
Public Property Get oName() As String
    oName = mvaroName
End Property
Private Sub Class_Initialize()
    Me.oType = "varchar"
    Me.oSize = 10
End Sub
Public Property Let Collation(ByVal sVal As String)
    If bString Then
        If Len(sVal) = 0 Then sVal = objServer.Collation
        m_sCollation = sVal
    Else
        m_sCollation = ""
    End If
End Property
Public Property Get Collation() As String
    If bString Then Collation = m_sCollation
End Property

