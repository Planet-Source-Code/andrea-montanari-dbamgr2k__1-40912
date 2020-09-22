VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum fStat
    NoVar = 0
    Changed = 1
    Delete = 2
    Added = 3
    Inserted = 4
End Enum

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
'Private mvaroNew As Boolean 'Copia locale.

Private mvaroStat As fStat 'Copia locale.
Private mvaroName As String 'Copia locale.
Private mvaroPkey As Boolean 'Copia locale.
Private mvaroType As String 'Copia locale.
Private mvaroSize As Long 'Copia locale.
Private mvaroNull As Boolean 'Copia locale.
Private mvaroDefault As String 'Copia locale.
Private mvaroPrecision As Integer 'Copia locale.
Private mvaroScale As Integer 'Copia locale.
Private mvaroGuid As Boolean 'Copia locale.
Private mvaroIdentity As Boolean 'Copia locale.
Private mvaroIDInit As Long 'Copia locale.
Private mvaroIDInc As Long 'Copia locale.
Private mbComputed As Boolean
Private msComputedText As String
Private bDecimal As Boolean
Private bUDT As Boolean
Private sUDTName As String
Private bNumeric As Boolean
Private bDate As Boolean
Private bSizeble As Boolean
Private bCanInt As Boolean
Private miPoint2 As Integer 'Copia locale.
Private m_sCollation As String
Private m_bIsString As Boolean
Private m_bIsBinary As Boolean

Public Property Get IsBinary() As Boolean
    IsBinary = m_bIsBinary
End Property
Public Property Get IsString() As Boolean
    IsString = m_bIsString
End Property
Public Property Get Collation() As String
    If m_bIsString Then Collation = m_sCollation
End Property
Public Property Let Collation(ByVal sVal As String)
    If m_bIsString Then
        If Len(sVal) = 0 Then sVal = objServer.Collation
        m_sCollation = sVal
    Else
        m_sCollation = ""
    End If
End Property
Public Property Get ComputedText() As String
    ComputedText = msComputedText
End Property
Public Property Let ComputedText(ByVal sVal As String)
    msComputedText = sVal
    mbComputed = Len(sVal)
End Property
Public Property Get UdtName() As String
    UdtName = sUDTName
End Property
Public Property Let UdtName(ByVal sVal As String)
    sUDTName = sVal
End Property

Public Property Get IsUDT() As Boolean
    IsUDT = bUDT
End Property
Public Property Let IsUDT(ByVal bVal As Boolean)
    bUDT = bVal
End Property

Public Property Get IsComputed() As Boolean
    IsComputed = mbComputed
End Property
Public Property Let IsComputed(ByVal bVal As Boolean)
    mbComputed = bVal
End Property

Public Property Get CanBeInt() As Boolean
    CanBeInt = bCanInt
End Property

Public Property Get isODate() As Boolean
    isODate = bDate
End Property

Public Property Get isOSizeble() As Boolean
    isOSizeble = bSizeble
End Property

Public Property Get isODecimal() As Boolean
    isODecimal = bDecimal
End Property
Public Property Get isONumeric() As Boolean
    isONumeric = bNumeric
End Property

Public Property Let oIDInc(ByVal vData As Long)
    mvaroIDInc = vData
End Property

Public Property Get oIDInc() As Long
    oIDInc = mvaroIDInc
End Property

Public Property Let oIDInit(ByVal vData As Long)
    mvaroIDInit = vData
End Property

Public Property Get oIDInit() As Long
    oIDInit = mvaroIDInit
End Property

Public Property Let oIdentity(ByVal vData As Boolean)
    mvaroIdentity = vData
End Property

Public Property Get oIdentity() As Boolean
    oIdentity = mvaroIdentity
End Property
Public Property Let oStatus(ByVal vData As fStat)
    mvaroStat = vData
End Property

Public Property Get oStatus() As fStat
    oStatus = mvaroStat
End Property

Public Property Let oGuid(ByVal vData As Boolean)
    mvaroGuid = vData
End Property

Public Property Get oGuid() As Boolean
    oGuid = mvaroGuid
End Property

Public Property Let oScale(ByVal vData As Integer)
    mvaroScale = vData
End Property

Public Property Get oScale() As Integer
    oScale = mvaroScale
End Property

Public Property Let oPrecision(ByVal vData As Integer)
    mvaroPrecision = vData
End Property

Public Property Get oPrecision() As Integer
    oPrecision = mvaroPrecision
End Property

Public Property Let oDefault(ByVal vData As String)
    mvaroDefault = vData
End Property

Public Property Get oDefault() As String
    oDefault = mvaroDefault
End Property

Public Property Let oNull(ByVal vData As Boolean)
    mvaroNull = vData
End Property

Public Property Get oNull() As Boolean
    oNull = mvaroNull
End Property

Public Property Let oSize(ByVal vData As Long)
    mvaroSize = vData
End Property

Public Property Get oSize() As Long
    oSize = mvaroSize
End Property

Public Property Let oType(ByVal vData As String)
        
    mvaroType = vData
    m_sCollation = ""
    m_bIsString = False
    m_bIsBinary = False
    Select Case mvaroType
        Case "binary", "varbinary"
            m_bIsBinary = True
            bCanInt = False
            bNumeric = False
            bDecimal = False
            bSizeble = True
            bDate = False
        Case "decimal", "numeric"
            bCanInt = True
            bNumeric = True
            bDecimal = True
            bSizeble = False
            bDate = False
        Case "float", "money", "real", "smallmoney"
            bCanInt = False
            bNumeric = True
            bDecimal = True
            bSizeble = False
            bDate = False
        Case "int", "smallint", "tinyint", "bigint"
            bCanInt = True
            bNumeric = True
            bDecimal = False
            bSizeble = False
            bDate = False
        Case "char", "varchar", "nchar", "nvarchar"
            bCanInt = False
            bSizeble = True
            bNumeric = False
            bDecimal = False
            bDate = False
            m_bIsString = True
            m_sCollation = objServer.Collation
        Case "datetime", "smalldatetime", "timestamp", "rowversion"
            bCanInt = False
            bDate = True
            bNumeric = False
            bDecimal = False
            bSizeble = False
        Case "image", "text", "ntext"
            bCanInt = False
            bDate = False
            bNumeric = False
            bDecimal = False
            bSizeble = True
        Case Else
            bCanInt = False
            bSizeble = False
            bDate = False
            bNumeric = False
            bDecimal = False
    End Select
    
End Property

Public Property Get oType() As String
    oType = mvaroType
End Property

Public Property Let oPkey(ByVal vData As Boolean)
    mvaroPkey = vData
End Property

Public Property Get oPkey() As Boolean
    oPkey = mvaroPkey
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
    Me.oNull = True
End Sub
Public Property Let Point2(ByVal vData As Integer)
    miPoint2 = vData
End Property

Public Property Get Point2() As Integer
    Point2 = miPoint2
End Property

Public Property Get InSelectList() As Boolean
    InSelectList = Not StrComp(mvaroType, "timestamp", vbTextCompare) = 0 And Not mbComputed
End Property
Public Property Get CanBeRowGuid() As Boolean
    CanBeRowGuid = StrComp(mvaroType, "uniqueidentifier", vbTextCompare) = 0
End Property

