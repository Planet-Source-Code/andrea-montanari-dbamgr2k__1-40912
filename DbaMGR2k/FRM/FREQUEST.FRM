VERSION 5.00
Begin VB.Form fRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tDati 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox tDati 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "fRequest.frx":000C
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdFunz 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdFunz 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "fRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bCancel As Boolean
Private vMinVal As Variant, vMaxVal As Variant
Private iMode As iPropG
Private RetVal As Variant
Private sPattern As String
Private sDefaultLike As String
Private bPattern As Boolean
Private bUpper As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property
Public Property Get ReturnValue() As Variant
    ReturnValue = RetVal
End Property
Public Property Let UpperCase(ByVal bval As Boolean)
    bUpper = bval
End Property

Public Property Let Mode(ByVal iVal As iPropG)
    iMode = iVal
    If iMode = am_Dec Or iMode = am_Long Then
        tDati(1).Visible = True
        tDati(1).ZOrder
        tDati(0).Visible = False
    Else
        tDati(0).Visible = True
        tDati(0).ZOrder
        tDati(1).Visible = False
    End If
End Property
Public Property Let Messages(ByVal sLabel As String, ByVal sForm As String)
    Me.Caption = sForm
    lbl.Caption = sLabel
End Property
Public Property Let Pattern(ByVal sVal As String, ByVal sDefPattern As String)
    sPattern = sVal
    sDefaultLike = sDefPattern
    bPattern = True
End Property

Public Property Let MaxVal(ByVal vVal As Variant)
    vMaxVal = vVal
End Property
Public Property Let ShowValue(ByVal vValue As Variant, ByVal iMaxLenght As Integer)

    Dim iIndex
    Dim lW As Long
    iIndex = IIf(iMode = am_Long Or iMode = am_Dec, 1, 0)
    
    tDati(iIndex).Text = vValue
    tDati(iIndex).MaxLength = iMaxLenght
    'lW = Me.TextWidth(String$(iMaxLenght, "W"))
    lW = (120 * iMaxLenght) + 150
    
    If lW > Me.ScaleWidth - 50 Then lW = Me.ScaleWidth - 50
    'lW = 1000
    tDati(iIndex).Width = lW 'Me.TextWidth(String$(iMaxLenght, "W"))
    tDati(iIndex).Left = (Me.ScaleWidth - tDati(iIndex).Width) \ 2
    
End Property

Public Property Let MinVal(ByVal vVal As Variant)
    vMinVal = vVal
End Property

Private Sub cmdFunz_Click(Index As Integer)

    Dim bFalse As Boolean
    Dim iIndex As Integer
    
    Select Case Index
        Case 0
            bCancel = True
            Me.Hide
        Case 1
            iIndex = IIf(iMode = am_Long Or iMode = am_Dec, 1, 0)
            If iMode = am_Date Then bFalse = Not IsDateValid(tDati(iIndex).Text, tDati(iIndex).MaxLength)
            If Not bFalse Then
                If iMode <> am_String Then
                    If Not IsEmpty(vMinVal) Then
                    'If Not vMinVal Is Null Then
                        If iMode = am_Long Then
                            bFalse = Val(tDati(iIndex).Text) < CLng(vMinVal)
                        ElseIf iMode = am_Dec Then
                            bFalse = ValF(tDati(iIndex).Text) < CLng(vMinVal)
                        ElseIf iMode = am_Date Then
                            If tDati(iIndex).MaxLength = 10 Or tDati(iIndex).MaxLength = 8 Then
                                bFalse = DateValue(tDati(iIndex).Text) < DateValue(vMinVal)
                            ElseIf tDati(iIndex).MaxLength = 7 Then
                                bFalse = DateValue("01/" & tDati(iIndex).Text) < DateValue("01/" & vMinVal)
                            End If
                        End If
                    End If
                    If Not bFalse Then
                        'If Not vMaxVal Is Nothing Then
                        If Not IsEmpty(vMaxVal) Then
                            If iMode = am_Long Then
                                bFalse = ValInt(tDati(iIndex).Text) > CLng(vMaxVal)
                            ElseIf iMode = am_Dec Then
                                bFalse = ValF(tDati(iIndex).Text) > CLng(vMaxVal)
                            ElseIf iMode = am_Date Then
                                If tDati(iIndex).MaxLength = 10 Or tDati(iIndex).MaxLength = 8 Then
                                    bFalse = DateValue(tDati(iIndex).Text) > DateValue(vMaxVal)
                                ElseIf tDati(iIndex).MaxLength = 7 Then
                                    bFalse = DateValue("01/" & tDati(iIndex).Text) > DateValue("01/" & vMaxVal)
                                End If
                            End If
                        
                        End If
                    End If
                End If
            End If
            If Not bFalse Then
                If bPattern Then bFalse = Not (tDati(iIndex).Text Like sPattern)
            End If
            If Not bFalse Then
                Select Case iMode
                    Case am_Dec
                        RetVal = CCur(tDati(iIndex).Text)
                    Case am_Long
                        RetVal = CLng(tDati(iIndex).Text)
                    Case Else 'am_String, am_Date
                        RetVal = CStr(tDati(iIndex).Text)
                End Select
                Me.Hide
            Else
                Beep
                tDati(iIndex).SelStart = 0
            End If
    End Select
End Sub

Private Sub Form_Initialize()
    vMinVal = Empty
    vMaxVal = Empty 'Null
End Sub

Private Sub Form_Load()

    cmdFunz(0).Caption = MyLoadResString(k_Cancel)
    CenterForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fRequest = Nothing
End Sub

Private Sub tDati_GotFocus(Index As Integer)
    If bPattern Then
        tDati(Index).SelStart = Get1AlfaNumChar(sDefaultLike)
    Else
        tDati(Index).SelStart = 0
    End If
End Sub

Private Sub tDati_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Debug.Print Index
    If iMode = am_Date Then
        Data_KDown tDati(Index), KeyCode, Left$(MASKYEAR, tDati(Index).MaxLength)
    ElseIf iMode = am_Long Then
        Digit_KDown tDati(Index), KeyCode
    ElseIf iMode = am_Dec Then
        Decimal_kDown tDati(Index), KeyCode
    ElseIf bPattern Then
        Mask_KDown tDati(Index), KeyCode, sPattern, sDefaultLike
    Else
        If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then KeyCode = 0
    End If
End Sub

Private Sub tDati_KeyPress(Index As Integer, KeyAscii As Integer)

    Debug.Print Index
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyEscape
            KeyAscii = 0
        Case vbKeyBACK, 32, 47 To 58, 65 To 122
            If iMode = am_Date Then
                Data_KPress tDati(Index), KeyAscii, Left$(MASKYEAR, tDati(Index).MaxLength)
            ElseIf iMode = am_Long Then
                Digit_KPress tDati(Index), KeyAscii
            ElseIf iMode = am_Dec Then
                Decimal_kPress tDati(Index), KeyAscii
            ElseIf bPattern Then
                If bUpper Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
                Mask_KPress tDati(Index), KeyAscii, sPattern, sDefaultLike
            Else
                If bUpper Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
                OverStrike tDati(Index)
            End If
        Case Else
            KeyAscii = 0
    End Select

End Sub
