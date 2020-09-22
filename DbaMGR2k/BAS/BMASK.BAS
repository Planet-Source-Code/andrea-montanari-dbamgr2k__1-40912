Attribute VB_Name = "bMask"
Attribute VB_Description = "Routine Associate Gestione MASK Date e Interi"
Option Explicit

Const PUNTO = 46, SLASH = 47, kSpace = 32
Const vbKey0 = 48, vbKey9 = 57
Const vbKeyUP = 38, vbKeyDOWN = 40, vbKeyDELETE = 46, vbKeyBACK = 8
Const vbKeyLEFT = 37, vbKeyRIGHT = 39
Const kComma = 44
Const kHome = 36, kEnd = 35
'
' Num con routine.
'1) Settare MaxkLenght=Num.Char desiderato + i Punti
'2) Allineamento Right
'3) MultiLine=True
'4) Select Case KeyAscii (intercetta Numeri e BackSpace)
    '    Case 13, 27
    '        Beep
    '        KeyAscii = 0
    '    Case Else
    '        Digit_KPress Text1, KeyAscii
    'End Select
'5) Select Case KeyCode (Intercetta Mov.Cursore e Canc)
    'Digit_KDown Text1, KeyCode

Public Const MASKMY = "  /    ", MASKYEAR = "  /  /    "

' Data con routine.
'1) Settare MaxkLenght=Num.Char desiderato + i /
'2) nel Form_Load Settare i Mask
'3) Select Case KeyAscii (intercetta Numeri e BackSpace)
    '    Case 13, 27
    '        Beep
    '        KeyAscii = 0
    '    Case Else
    '        Data_KPress Text1, KeyAscii, MASK
    'End Select
'4) Select Case KeyCode (Intercetta Mov.Cursore e Canc)
    'Data_KDown Text1, KeyCode, MASK

Public Sub Data_KDown(txt As TextBox, iKeyCode As Integer, ByVal sMask As String)

    Dim i As Integer, iPos As Integer, sTmp As String
    
    If iKeyCode <> vbKeyDELETE Then
        If txt.SelLength > 0 Then txt.SelLength = 0
    End If
        
    Select Case iKeyCode
        Case vbKeyUP, vbKeyDOWN
            iKeyCode = 0
        Case vbKeyRIGHT, vbKeyLEFT
            iPos = txt.SelStart + (iKeyCode - 38)
            If iPos >= Len(sMask) Or iPos < 1 Then Exit Sub
            sTmp = txt
            i = Asc(Mid$(sTmp, iPos, 1))
            If i = kSpace Then
                iPos = 0
                For i = 0 To Len(sTmp)
                    If Asc(Mid$(sTmp, i + 1, 1)) >= SLASH Then
                        iPos = iPos + 1
                    Else
                        Exit For
                    End If
                Next
            ElseIf Asc(Mid$(sTmp, iPos + 1, 1)) = SLASH Then
                iPos = iPos + (iKeyCode - 38)
            End If
            txt.SelStart = iPos
            iKeyCode = 0
        Case vbKeyDELETE         'CANC 46
            iKeyCode = 0
            If txt.SelLength = Len(sMask) Then
                txt = sMask
                Exit Sub
            End If
            sTmp = txt
            iPos = txt.SelStart + 1
            If iPos > Len(sMask) Then Exit Sub
            If txt.SelLength > 0 Then
                Mid$(sTmp, iPos) = Mid$(sMask, iPos)
            ElseIf Asc(Mid$(sTmp, iPos, 1)) = SLASH Then
                iPos = iPos - 1
            Else
                Mid$(sTmp, iPos, 1) = "0"
            End If
            txt = sTmp
            txt.SelStart = iPos - 1
    End Select
    
End Sub

Public Sub Data_KPress(txt As TextBox, iKeyAscii As Integer, ByVal sMask As String)
    
    Dim iPos As Integer, sTmp As String

    Select Case iKeyAscii
        Case vbKeyBACK
            iKeyAscii = 0
            iPos = txt.SelStart
            If iPos = 0 Then Exit Sub
            sTmp = txt
            If Mid$(sTmp, iPos, 1) = Chr$(SLASH) Then
                iPos = iPos - 1
            End If
            If txt.SelLength = Len(sTmp) Then
                txt = sMask
            Else
                Mid$(sTmp, iPos, 1) = "0"
                txt = sTmp
                txt.SelStart = iPos - 1
            End If
        Case vbKey0 To vbKey9
            iPos = txt.SelStart + 1         ' posn = 0 prior to this statement
            If iPos > Len(sMask) Then
                iPos = Len(sMask)
            End If
            sTmp = txt
            If txt.SelLength > 0 Then
                Mid$(sTmp, iPos) = Mid$(sMask, iPos)
            End If
            Mid$(sTmp, iPos, 1) = Chr$(iKeyAscii)
            txt = sTmp
            If Mid$(sTmp, iPos + 1, 1) = Chr$(SLASH) Then
                iPos = iPos + 1
            End If
            txt.SelStart = iPos
            iKeyAscii = 0
        Case Else
            Beep
            iKeyAscii = 0
    End Select

End Sub

Public Sub Digit_KDown(txt As TextBox, iKeyCode As Integer)

    Dim iPos As Integer, sTmp As String, iLen As Integer

    If txt.SelLength > 1 Then txt.SelLength = 1
    Select Case iKeyCode
        Case vbKeyUP, vbKeyDOWN
            iKeyCode = 0
        Case vbKeyRIGHT
            If Mid$(txt.Text, txt.SelStart + 2, 1) = Chr$(PUNTO) Then
                txt.SelStart = txt.SelStart + 1
            End If
        Case vbKeyLEFT
            If txt.SelStart > 0 Then
                If Mid$(txt.Text, txt.SelStart, 1) = Chr$(PUNTO) Then
                    txt.SelStart = txt.SelStart - 1
                End If
            End If
        Case vbKeyDELETE
            If txt.SelLength = Len(txt.Text) Then
                txt.Text = 0
                txt.SelLength = 0
            Else
                sTmp = txt.Text
                iPos = txt.SelStart + 1
                iLen = Len(sTmp)
                If iPos > iLen Then iPos = iLen
                Mid$(sTmp, iPos, 1) = " "
                sTmp = FrmDigit(sTmp)
                iPos = iPos - 1
                iLen = iLen - Len(sTmp) - 1
                If iPos >= iLen Then iPos = iPos - iLen
                If Mid$(sTmp, iPos + 1, 1) = Chr$(PUNTO) Then
                    iPos = iPos + 1
                End If
                txt.Text = sTmp
                txt.SelStart = iPos
            End If
            iKeyCode = 0
    End Select

End Sub

Public Sub Digit_KPress(txt As TextBox, iKeyAscii As Integer)
    
    Dim iPos As Integer, sTmp As String

    Select Case iKeyAscii
        Case vbKeyBACK
            iPos = txt.SelStart
            If iPos > 0 Then
                sTmp = txt.Text
                If iPos = Len(sTmp) Then
                    Mid$(sTmp, iPos, 1) = " "
                Else
                    If Mid$(sTmp, iPos, 1) = Chr$(PUNTO) Then
                        iPos = iPos - 1
                    End If
                    Mid$(sTmp, iPos, 1) = "0"
                End If
                txt.Text = FrmDigit(sTmp)
                txt.SelStart = iPos - 1
            Else
                Beep
            End If
            iKeyAscii = 0
            
        Case vbKey0 To vbKey9
            sTmp = txt.Text
            iPos = txt.SelStart + 1
            If iPos > Len(sTmp) Then
                If iPos <= txt.MaxLength Then
                    txt.Text = FrmDigit(sTmp & Chr$(iKeyAscii))
                    txt.SelStart = Len(txt.Text)
                End If
            Else
                Mid$(sTmp, iPos, 1) = Chr$(iKeyAscii)
                If Mid$(sTmp, iPos + 1, 1) = Chr$(PUNTO) Then
                    iPos = iPos + 1
                End If
                txt.Text = sTmp
                txt.SelStart = iPos
            End If
            If iKeyAscii = vbKey0 And Asc(Left$(txt, 1)) = 48 Then
                txt.Text = FrmDigit(txt)
            End If
            iKeyAscii = 0
        Case Else
            Beep
            iKeyAscii = 0
    End Select

End Sub

Public Function frm(ByVal l As Long, ByVal s As String) As String
    '*-----------------------------------
    ' FORMATTA NUMERO A DESTRA SENZA SEP.
    '*-----------------------------------

    Dim i As Integer

    i = Len(s)
    frm = Right$(Space$(i) & Format$(l, s), i)

End Function

Public Function FrmDigit(ByVal s As String) As String

    FrmDigit = Format$(ValInt(s), "##,###,##0")

End Function

Public Function InvertData(ByVal sData As String) As String

    InvertData = Format$(DateValue(sData), kInvertYEAR)

End Function
Public Function IsDateValid(ByVal sData As String, ByVal iLenData As Integer) As Integer

'    Dim sStr As String
'    sStr = sData
    
    If iLenData = 5 Then
        If Len(sData) = iLenData Then
            If (Val(Left$(sData, 2)) > 0 _
                And Val(Left$(sData, 2)) <= 31) _
                And (Val(Mid$(sData, 4, 2)) > 0 _
                And Val(Mid$(sData, 4, 2)) < 13) Then
                
                IsDateValid = True
            End If
        End If
    ElseIf iLenData <> 7 Then
        If Len(sData) = iLenData Then
            If (Val(Left$(sData, 2)) > 0 _
                And Val(Left$(sData, 2)) <= 31) _
                And (Val(Mid$(sData, 4, 2)) > 0 _
                And Val(Mid$(sData, 4, 2)) < 13) _
                And (Val(Right$(sData, 4)) > 1900 _
                And Val(Right$(sData, 4)) < 2051) Then
                
                If IsDate(sData) Then IsDateValid = True
            End If

'
'        If IsDate(sStr) And Len(Trim$(sData)) = iLenData Then
'            If InStr(sData, " ") = 0 _
'                And (Right$(sData, 4) <> "0000") Then
'                    IsDateValid = True
'            End If
'        Else
'            IsDateValid = False
        End If
    Else
        If (Val(Left$(sData, 2)) > 0 _
            And Val(Left$(sData, 2)) < 13) _
            And (Val(Right$(sData, 4)) > 1979 _
            And Val(Right$(sData, 4)) < 2051) Then
                IsDateValid = True
'        Else
'            IsDateValid = False
        End If
    End If

End Function

Public Function ValChar(ByVal sStr As String) As String
    '*--------------------------------------------
    ' OTTIENE NUMERO COME CARATTERE NON FORMATTATO
    '*--------------------------------------------

    Dim sTmp As String, iChar As Integer, i As Integer
    
    sTmp = ""
    For i = 1 To Len(sStr)
        iChar = Asc(Mid$(sStr, i, 1))
        If iChar = 44 Then Exit For     '44=","
        If iChar > SLASH Then
            sTmp = sTmp & Chr$(iChar)
        End If
    Next
    ValChar = sTmp
    
End Function

Public Function ValF(ByVal sBuff As String) As Double
    '*-----------------------------------------------
    ' RICAVA NUMERO DOUBLE DA NUMERO TESTO FORMATTATO
    '*-----------------------------------------------

    Dim i As Integer

    For i = 1 To Len(sBuff)
        Select Case Mid$(sBuff, i, 1)
            Case "."
                Mid$(sBuff, i, 1) = " "
            Case ","
                Mid$(sBuff, i, 1) = "."
        End Select
   Next
   ValF = Val(sBuff)

End Function

Public Function ValInt(ByVal sBuff As String) As Long
'*------------------------------------------------
' OTTIENE VALORE INTERO DA NUMERO TESTO FORMATTATO
'*------------------------------------------------

    Dim i As Integer

    For i = 1 To Len(sBuff)
        If Asc(Mid$(sBuff, i, 1)) = PUNTO Then
            Mid$(sBuff, i, 1) = " "
        End If
    Next
    ValInt = Val(sBuff)

End Function

Public Sub Mask_KPress(txt As TextBox, iKeyAscii As Integer, _
                        ByVal sMask As String, ByVal sDef As String)
    
    Dim iPos As Integer, sTmp As String
    Dim iLen As Integer
    Dim sBuf As String

    Select Case iKeyAscii
        Case vbKeyBACK
            iKeyAscii = 0
            iPos = txt.SelStart
            If iPos = 0 Then Exit Sub

            sBuf = txt.Text
            sTmp = sBuf
         
            If txt.SelLength = Len(sTmp) Then
                txt = sMask
            Else
                If iPos < 1 Then iPos = 1
                Do While iPos > 1 And iPos < Len(sDef)
                    If IsAlfaNum(Mid$(sDef, iPos, 1)) Then Exit Do
                    iPos = iPos - 1
                Loop
                
                Mid$(sTmp, iPos, 1) = Mid$(sDef, iPos, 1) '"0"
                txt = sTmp
                iPos = iPos - 1
                Do While iPos > 1 And iPos < Len(sDef)
                    If IsAlfaNum(Mid$(sDef, iPos + 1, 1)) Then Exit Do
                    iPos = iPos - 1
                Loop
                If iPos = 0 Then iPos = Get1AlfaNumChar(sDef)
                'txt.SelStart = iPos - 1
                txt.SelStart = iPos
            End If
        Case Else
            iLen = Len(txt.Text)
            sBuf = txt.Text
            sTmp = sBuf
            
            iPos = txt.SelStart + 1         ' posn = 0 prior to this statement
            If iPos > txt.MaxLength Then iPos = txt.MaxLength

            If txt.SelLength > 0 Then
                iKeyAscii = 0
                Exit Sub
                Mid$(sTmp, iPos) = Mid$(sMask, iPos)
            End If
            If iLen >= txt.MaxLength Then
                Mid$(sTmp, iPos, 1) = Chr$(iKeyAscii)
            Else
                sTmp = sTmp & Chr$(iKeyAscii)
            End If
            If sTmp Like sMask Then
                txt = sTmp
                iPos = iPos + 1
                Do While iPos < Len(sDef)
                    If IsAlfaNum(Mid$(sDef, iPos, 1)) Then Exit Do
                    iPos = iPos + 1
                Loop
                iPos = iPos - 1
            Else
                Beep
                txt.Text = sBuf
                iPos = iPos - 1
            End If
            
'            If Not IsAlfaNum(Mid$(sDef, iPos + 1, 1)) Then
'            'If Mid$(sDef, iPos + 1, 1) = Chr$(SLASH) Then
'                iPos = iPos + 1
'            End If
            
            txt.SelStart = iPos
            iKeyAscii = 0
    End Select

End Sub

Public Sub Mask_KDown(txt As TextBox, iKeyCode As Integer, _
                        ByVal sMask As String, ByVal sDef As String)

    Dim iPos As Integer, sTmp As String
                    
    If txt.SelLength > 0 Then txt.SelLength = 0
    
    Select Case iKeyCode
        Case vbKeyUP, vbKeyDOWN
            iKeyCode = 0
        Case kHome
            iKeyCode = 0
            txt.SelStart = Get1AlfaNumChar(sDef)
        Case kEnd
            iKeyCode = 0
            iPos = Len(sDef) '- 1
            Do While iPos > 1 And iPos <= Len(sDef)
                If IsAlfaNum(Mid$(sDef, iPos, 1)) Then Exit Do
                iPos = iPos - 1
            Loop
            txt.SelStart = iPos
            
            
        Case vbKeyRIGHT, vbKeyLEFT
            iPos = txt.SelStart + (iKeyCode - 38)
            If iPos >= Len(sDef) Or iPos < 1 Then
                If iKeyCode = vbKeyLEFT Then
                    iKeyCode = 0
                    txt.SelStart = Get1AlfaNumChar(sDef)
                End If
                Exit Sub
            End If
            sTmp = txt
            'i = Asc(Mid$(sTmp, iPos, 1))
            Do While iPos > 1 And iPos < Len(sDef)
                If IsAlfaNum(Mid$(sTmp, iPos + 1, 1)) Then Exit Do
                iPos = iPos + (iKeyCode - 38)
            Loop
            txt.SelStart = iPos
            iKeyCode = 0
        
        Case vbKeyDELETE         'CANC 46
            iKeyCode = 0
            If txt.SelLength = txt.MaxLength Then
                Exit Sub
            End If
            sTmp = txt.Text

            iPos = txt.SelStart + 1
            If iPos > txt.MaxLength Then Exit Sub
            Mid$(sTmp, iPos, 1) = Mid$(sDef, iPos, 1) '"0"

            txt = sTmp
            txt.SelStart = iPos - 1
    End Select
    
End Sub

'Private Function GetMaskChar(ByVal iPos As Integer, ByVal sMask As String) As String
'
'    Dim i As Integer
'    ReDim sBuf(0) As String
'    Dim iCount As Integer
'    Dim bOpen As Boolean
'    Dim bFreeChar As Boolean
'
'    iCount = -1
'    For i = 1 To Len(sMask)
'        Select Case Mid$(sMask, i, 1)
'            Case "["
'                bFreeChar = False
'                If Not bOpen Then
'                    iCount = iCount + 1
'                    ReDim Preserve sBuf(iCount) As String
'                    bOpen = True
'                End If
'            Case "]"
'                bOpen = False
'                bFreeChar = False
'            Case "-"
'            Case "*"
'            Case "?"
'                iCount = iCount + 1
'                ReDim Preserve sBuf(iCount) As String
'                sBuf(iCount) = "Z"
'            Case "#"
'                iCount = iCount + 1
'                ReDim Preserve sBuf(iCount) As String
'                sBuf(iCount) = "0"
'            Case "!"
'            Case "A" To "Z"
'
'                If bOpen And Not bFreeChar Then
'                    sBuf(iCount) = "Z"
'                    bFreeChar = True
'                ElseIf Not bOpen And Not bFreeChar Then
'                    bFreeChar = True
'                    iCount = iCount + 1
'                    ReDim Preserve sBuf(iCount) As String
'                    sBuf(iCount) = "Z"
'                End If
'        End Select
'    Next
'    GetMaskChar = sBuf(iPos - 1)
'
'End Function
Public Sub Decimal_kDown(txt As TextBox, iKeyCode As Integer)

    Dim iPos As Integer, iLen As Integer, i As Integer
    Dim sTmp As String

    If txt.SelLength > 1 Then txt.SelLength = 1
    sTmp = txt
    Select Case iKeyCode
        Case vbKeyUP, vbKeyDOWN
            iKeyCode = 0
        Case vbKeyRIGHT
            If GetAsc(Mid$(sTmp, txt.SelStart + 2, 1)) = PUNTO Then
                txt.SelStart = txt.SelStart + 1
            End If
        Case vbKeyLEFT
            If txt.SelStart > 0 Then
                i = GetAsc(Mid$(sTmp, txt.SelStart, 1))
                If i = PUNTO Or i = kComma Then
                    txt.SelStart = txt.SelStart - 1
                End If
            End If
        Case vbKeyDELETE
            i = txt.SelLength
            If i = Len(txt.Text) Then
                txt.Text = "0,00"
                txt.SelLength = 0
            Else
                iPos = txt.SelStart + 1
                iLen = Len(sTmp)
                If iPos > iLen Then
                    iPos = iLen
                End If
                If Mid$(sTmp, iPos, 1) <> "," Then
                    Mid$(sTmp, iPos, 1) = "0"
                End If
                sTmp = FrmDecim(sTmp)
                iPos = iPos - 1
                iLen = iLen - Len(sTmp) - 1
                If iPos >= iLen Then iPos = iPos - iLen
                i = GetAsc(Mid$(sTmp, iPos + 1, 1))
                If i = PUNTO Or i = kComma Then
                    iPos = iPos + 1
                End If
                txt.Text = sTmp
                txt.SelStart = iPos
            End If
            iKeyCode = 0
    End Select

End Sub

Public Sub Decimal_kPress(txt As TextBox, iKeyAscii As Integer)
    
    Dim sTmp As String
    Dim iComma As Integer, iPos As Integer

    sTmp = txt.Text
    iComma = InStr(sTmp, ",")
    
    Select Case iKeyAscii
        Case vbKeyBACK
            iPos = txt.SelStart
            If iPos > 0 Then
                If iPos = Len(sTmp) Then
                    Mid$(sTmp, iPos, 1) = " "
                Else
                    iComma = GetAsc(Mid$(sTmp, txt.SelStart, 1))
                    If iComma = PUNTO Or iComma = kComma Then
                        iPos = iPos - 1
                    End If
                    Mid$(sTmp, iPos, 1) = "0"
                End If
                txt.Text = FrmDecim(sTmp)
                txt.SelStart = iPos - 1
            Else
                Beep
            End If
            iKeyAscii = 0
            
        Case vbKey0 To vbKey9
            iPos = txt.SelStart + 1
            iComma = Len(sTmp) - 2
            If iPos = iComma Then
                If Len(sTmp) < txt.MaxLength Then
                    sTmp = Mid$(sTmp, 1, iComma - 1) & Chr$(iKeyAscii) & Mid$(sTmp, iComma)
                    If Len(sTmp) = 7 Or Len(sTmp) = 11 Then
                        iPos = iPos + 1
                    End If
                Else
                    Beep
                    iPos = iPos - 1
                End If
            Else
                If iPos <= Len(sTmp) Then
                    Mid$(sTmp, iPos, 1) = Chr$(iKeyAscii)
                    If Mid$(sTmp, iPos + 1, 1) = Chr$(PUNTO) Then
                        iPos = iPos + 1
                    End If
                End If
            End If
            sTmp = FrmDecim(sTmp)
            If Mid$(sTmp, iPos, 1) = Chr$(kComma) Then iPos = iPos - 1
            
            txt.Text = sTmp
            txt.SelStart = iPos
            iKeyAscii = 0
        Case kComma
            iKeyAscii = 0
            txt.SelStart = iComma
        Case Else
            Beep
            iKeyAscii = 0
    End Select

End Sub

Public Function FrmDecim(ByVal s As String) As String

    FrmDecim = Format$(ValCur(s), "#,0.00")

End Function
Private Function GetAsc(ByVal szBuf As String) As Integer

    '//ottiene codice ASCII senza ERRORE
    If Len(szBuf) Then GetAsc = Asc(szBuf)

End Function
Public Function ValDbl(ByVal sBuff As String) As Double
'*------------------------------------------------
' OTTIENE VALORE INTERO DA NUMERO TESTO FORMATTATO
'*------------------------------------------------

    Dim i As Integer
    Dim iAsc As Integer

    For i = 1 To Len(sBuff)
        iAsc = Asc(Mid$(sBuff, i, 1))
        If iAsc = PUNTO Then
            Mid$(sBuff, i, 1) = " "
        ElseIf iAsc = kComma Then
            Mid$(sBuff, i, 1) = "."
        End If
    Next
    ValDbl = Val(sBuff)

End Function

Public Function frmCur(ByVal CurVal As Currency, ByVal s As String) As String
    '*-----------------------------------
    ' FORMATTA NUMERO A DESTRA SENZA SEP.
    '*-----------------------------------

    Dim i As Integer

    i = Len(s)
    frmCur = Right$(Space$(i) & Format$(CurVal, s), i)

End Function

Public Function ValCur(ByVal sBuff As String) As Currency
'*------------------------------------------------
' OTTIENE VALORE INTERO DA NUMERO TESTO FORMATTATO
'*------------------------------------------------

    Dim i As Integer
    Dim iAsc As Integer

    For i = 1 To Len(sBuff)
        iAsc = Asc(Mid$(sBuff, i, 1))
        If iAsc = PUNTO Then
            Mid$(sBuff, i, 1) = " "
        ElseIf iAsc = kComma Then
            Mid$(sBuff, i, 1) = "."
        End If
    Next
    ValCur = Val(sBuff)

End Function
Private Function IsAlfaNum(ByVal sChar As String) As Boolean
    
    If Len(sChar) Then
        Select Case Asc(sChar)
            Case 48 To 57, 65 To 90, 97 To 122
                IsAlfaNum = True
        End Select
    Else
        'IsAlfaNum = True
    End If
    
End Function
Public Function Get1AlfaNumChar(ByVal sDefaultPattern As String) As Integer

    Dim iPos As Integer
    iPos = 0
    Do While iPos < Len(sDefaultPattern)
        If IsAlfaNum(Mid$(sDefaultPattern, iPos + 1, 1)) Then Exit Do
        iPos = iPos + 1
    Loop
    Get1AlfaNumChar = iPos
End Function
