VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fUDT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "udt"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5145
   ClipControls    =   0   'False
   Icon            =   "fUDT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tFlex 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "used"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "fUDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum amUdtFlexRow
    rDataType = 1
    rCollation = 7
End Enum
Private Type udtSTR
    sField As String
    iLen As Integer
    vVal As Variant
    vOtherVal As Variant
    iType As iPropG
    bCombo As Boolean
    bDisab As Boolean
End Type
Private Type udtType
    sName As String
    bDis As Boolean
End Type

Private sDbName As String
Private sStuff As String
Private bNew As Boolean
Private bLoading As Boolean
Private m_blnAnnullaMod As Boolean
Private mbDirty As Boolean
Private xElem() As udtSTR
Private miRow As Integer

Private sData() As udtType
Private sRules() As String
Private sDef() As String
Private iTmpType As iPropG

Public Property Let DBName(ByVal sVal As String)
    sDbName = sVal
End Property

Public Sub loadUDT(ByVal sVal As String)

    Dim oType As SQLDMO.SystemDatatype2
    Dim oUdt As SQLDMO.UserDefinedDatatype2
    sStuff = sVal
    bNew = Len(sVal) = 0
    
    mbDirty = False
    Regen
    
    If Not bNew Then
        Set oUdt = GetUdtByName(objServer, sDbName, sStuff)
        With oUdt
            xElem(0).vVal = .Name
            xElem(0).vOtherVal = vbEmpty
            xElem(0).bDisab = True
                
            xElem(1).vVal = .BaseType
            xElem(1).vOtherVal = vbEmpty
            xElem(1).bDisab = True
            
            Set oType = objServer.Databases(sDbName).SystemDatatypes(.BaseType)
            xElem(2).vVal = .MaxSize
            
            xElem(2).vOtherVal = GetItemArr(xElem(1).vVal)
            xElem(2).bDisab = True
            
            If oType.IsNumeric Then
                xElem(2).vVal = .NumericPrecision
                xElem(3).vVal = .NumericScale
            End If
            xElem(3).bDisab = True
            
            xElem(4).vVal = .AllowNulls
            xElem(4).vOtherVal = vbEmpty
            xElem(4).bDisab = True
            
            
            If Len(.Rule) Then
                xElem(5).vVal = objServer.Databases(sDbName).Rules(.Rule).Owner & "." & .Rule
'            Else
'                xElem(5).vVal = "(none)"
            End If
            xElem(5).vOtherVal = 0
            
            If Len(.Default) Then
                xElem(6).vVal = objServer.Databases(sDbName).Defaults(.Default).Owner & "." & .Default
'            Else
'                xElem(6).vVal = "(none)"
            End If
            xElem(6).vOtherVal = 0
            
            xElem(7).vVal = .Collation
            If Len(xElem(7).vVal) = 0 Then xElem(7).vVal = MyLoadResString(k_UDT_Collation_Not_Applicable)
        
        End With
    End If
    Set oType = Nothing
    mbDirty = False
    
    LoadFlex
    cmd(1).Enabled = False
    cmd(2).Enabled = Not bNew
    
End Sub

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub cmd_Click(Index As Integer)

    If cmd(Index).Enabled Then
        TestActiveControl
        Select Case Index
            Case 0
                Unload Me
            Case 1  'save
                SaveUDT
            Case 2  'where used
                ShowUsed
        End Select
    End If
End Sub
Private Sub ShowUsed()

    Dim Lst As ListBox
    Dim oList As SQLDMO.SQLObjectList
    Dim oColumn As SQLDMO.Column
    Dim oTbl As SQLDMO.Table
    
    Screen.MousePointer = vbHourglass
    Set Lst = fUsed.Lst
    Set oList = objServer.Databases(sDbName).UserDefinedDatatypes(sStuff).ListBoundColumns

    For Each oColumn In oList
        Set oTbl = oColumn.Parent
        Lst.AddItem sStuff & ". " & oTbl.Name & ".[ " & oColumn.Name & " ]"
    Next
    
    Set oTbl = Nothing
    Set oColumn = Nothing
    Set oList = Nothing
    Set Lst = Nothing
    Screen.MousePointer = vbDefault
    fUsed.Show vbModal, Me
    
End Sub
Private Sub Regen()
    
    Dim i As Integer
    Dim iC As Integer
    Dim iLoop As Integer
    Dim sBuf As String
    Dim oDb As SQLDMO.Database2
    
    iC = objServer.Databases(sDbName).SystemDatatypes.Count
    ReDim sData(1 To 5) As udtType
    With objServer.Databases(sDbName)
        iLoop = 1
        For i = 1 To iC
            sBuf = .SystemDatatypes(i).Name
            If StrComp(sBuf, "sysname", vbTextCompare) <> 0 And StrComp(sBuf, "timestamp", vbTextCompare) <> 0 And StrComp(sBuf, "rowversion", vbTextCompare) <> 0 Then
                If UBound(sData) < iLoop Then ReDim Preserve sData(1 To iLoop) As udtType
                sData(iLoop).sName = .SystemDatatypes(i).Name
                sData(iLoop).bDis = Not .SystemDatatypes(i).AllowLength And Not (sData(iLoop).sName = "numeric" Or sData(iLoop).sName = "decimal")
                iLoop = iLoop + 1
            End If
        Next
    End With

    ReDim sRules(objServer.Databases(sDbName).Rules.Count) As String
    sRules(0) = "(none)"
    With objServer.Databases(sDbName)
        For i = 1 To UBound(sRules)
            sRules(i) = .Rules(i).Owner & "." & .Rules(i).Name
        Next
    End With


    ReDim sDef(objServer.Databases(sDbName).Defaults.Count) As String
    sDef(0) = "(none)"
    With objServer.Databases(sDbName)
        For i = 1 To objServer.Databases(sDbName).Defaults.Count
            sDef(i) = .Defaults(i).Owner & "." & .Defaults(i).Name
        Next
'        For i = 1 To UBound(sDef)
'            With objServer.Databases(sDBName).UserDefinedDatatypes(.Defaults(i).Name)
'                sDef(i) = .Owner & "." & .Name '.Defaults(i).Name
'            End With
'        Next
    End With
    
    ReDim xElem(7) As udtSTR
    xElem(0).sField = MyLoadResString(k_Name)
    xElem(0).iType = am_String
    xElem(0).iLen = 50
    xElem(0).vVal = ""
    xElem(0).vOtherVal = vbEmpty
        
    xElem(1).sField = MyLoadResString(k_Data_Type)
    xElem(1).iType = am_String
    xElem(1).iLen = 0
    xElem(1).vVal = "binary"
    xElem(1).vOtherVal = vbEmpty 'GetItemArr("binary")
    xElem(1).bCombo = True
    
    xElem(2).sField = MyLoadResString(k_Size)
    xElem(2).iType = am_Long
    xElem(2).iLen = 5
    xElem(2).vVal = 8000
    xElem(2).vOtherVal = GetItemArr("binary")
    
    xElem(3).sField = MyLoadResString(k_Scale)
    xElem(3).iType = am_Long
    xElem(3).iLen = 2
    xElem(3).vVal = 0
    xElem(3).vOtherVal = vbEmpty
    xElem(3).bDisab = True
    
    xElem(4).sField = MyLoadResString(k_Allow_Null)
    xElem(4).iType = am_Bool
    xElem(4).iLen = 0
    xElem(4).vVal = False
    xElem(4).vOtherVal = vbEmpty
    
    xElem(5).sField = MyLoadResString(k_Rule) 'Rule
    xElem(5).iType = am_String
    xElem(5).iLen = 0
    xElem(5).vVal = "(none)"
    xElem(5).vOtherVal = 0
    xElem(5).bCombo = True
    
    xElem(6).sField = "Default"
    xElem(6).iType = am_String
    xElem(6).iLen = 0
    xElem(6).vVal = "(none)"
    xElem(6).vOtherVal = 0
    xElem(6).bCombo = True
    
    Set oDb = objServer.Databases(sDbName)
    xElem(7).sField = MyLoadResString(k_NewDbSort)
    xElem(7).iType = am_String
    xElem(7).iLen = 0
    xElem(7).vVal = objServer.Collation
    xElem(7).vOtherVal = oDb.Collation 'objServer.Collation
    xElem(7).bCombo = False
    xElem(7).bDisab = True
    Set oDb = Nothing
    
End Sub
Private Sub Form_Load()

    
    Flex.RowHeightMin = cbo.Height + 10
    CenterForm Me
    Me.Caption = MyLoadResString(k_User_Defined_Data_Type_Properties)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    cmd(2).Caption = MyLoadResString(k_Where_Used)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fUDT = Nothing
End Sub
Private Sub LoadFlex()

    Dim sBuf As String
    Dim i As Integer
    Dim iX As Integer
    
    With Flex
        bLoading = True
        .Redraw = False
        .Clear
        .FixedCols = 0
        .Cols = 3
        .Rows = 1
        .Clear
        
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        .TextMatrix(0, 1) = MyLoadResString(k_Property)
        .TextMatrix(0, 2) = MyLoadResString(k_Value)
        
        
        For i = 0 To UBound(xElem)
            sBuf = vbTab & xElem(i).sField
            If (xElem(i).iType And am_String) = am_String Then
                sBuf = sBuf & vbTab & xElem(i).vVal
            ElseIf (xElem(i).iType And am_Long) = am_Long Then
                sBuf = sBuf & vbTab & Format$(xElem(i).vVal, "#,0")
            End If
            
            Flex.AddItem sBuf
            
            If (xElem(i).iType And am_Bool) = am_Bool Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                Set Flex.CellPicture = LoadResPicture(IIf(Abs(xElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            End If
            If (xElem(i).bDisab) Then
                .Row = .Rows - 1
                For iX = 2 To .Cols - 1
                    .Col = iX
                    .CellBackColor = vbInactiveBorder
                Next
            End If
            If xElem(i).bCombo Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                .CellPictureAlignment = flexAlignRightTop
            End If
            
            If i = 2 Then
                If sData(xElem(i).vOtherVal).bDis Then
                    .Row = .Rows - 1
                    For iX = 2 To .Cols - 1
                        .Col = iX
                        .CellBackColor = vbInactiveBorder
                    Next
                End If
            End If
        Next
        
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(2) = flexAlignLeftCenter
        
        .FixedRows = 1
        .FixedCols = 2
        
        .Redraw = True
        .ColWidth(1) = .Width \ 3
        SettaLargFlex Flex, 2, 300
        miRow = 0
        .TopRow = 1
        .Row = 1
        miRow = 1
        bLoading = False
        Flex_RowColChange
    End With

End Sub

Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
       
    If Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            With Flex
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex.CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1

                    .Col = iCol
                    miRow = iRowP
                End If

            End With
            b = False
        End If
    End If

End Sub
Private Sub Flex_GotFocus()
    TestActiveControl
End Sub
Private Sub Flex_Scroll()
    cbo.Visible = False
    tFlex.Visible = False
End Sub
Private Sub LoadCbo(sArray() As String)
    
    Dim i As Integer
    With cbo
        .Clear
        For i = LBound(sArray) To UBound(sArray)
            .AddItem sArray(i)
        Next
    End With
    
End Sub
Private Sub Flex_DblClick()
    
    If Flex.Col = Flex.Cols - 1 Then
        With Flex
            If .CellBackColor = vbInactiveBorder Then Exit Sub
            If xElem(miRow - 1).bCombo Then
                Select Case miRow - 1
                    Case rDataType
                        ReDim sArr(0) As String
                        GetArrData sArr()
                        LoadCbo sArr
                    Case 5
                        LoadCbo sRules()
                    Case 6
                        LoadCbo sDef()
                End Select
                
                cbo.ListIndex = 0
                cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                cbo.Visible = True
                cbo.SetFocus
            Else
                If .MouseRow > 0 Then FlexEdit vbKeySpace
            End If
        End With
    End If

End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)

    Dim iCode As Integer
    
    If Not xElem(miRow - 1).bCombo Then
        iCode = KeyAscii
        KeyAscii = 0
        FlexEdit iCode
    Else
        KeyAscii = 0
        Flex_DblClick
    End If

End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iRow As Integer, iCol As Integer

    Dim iImg As Integer
    
    iRow = Flex.Row - 1
    iCol = Flex.Col
        
    iTmpType = am_String
    
    bEdit = True
    If Flex.CellBackColor = vbInactiveBorder Or iRow < 0 Then bEdit = False
    
    If bEdit Then
        iTmpType = xElem(iRow).iType
        miRow = Flex.Row
        
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If (iTmpType And am_Bool) = am_Bool Then
                    iImg = 1 - (-xElem(iRow).vVal)
                    xElem(iRow).vVal = CBool(iImg)
                    
                    Set Flex.CellPicture = LoadResPicture(IIf(iImg, k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex.CellPictureAlignment = flexAlignCenterCenter
                    bEdit = False
                    mbDirty = True
                    cmd(1).Enabled = Len(xElem(0).vVal) <> 0
                    
                End If
            Case Else
                If (iTmpType And am_Bool) = am_Bool Then bEdit = False
        End Select
    End If

    If bEdit Then
        sBuf = Flex.TextMatrix(Flex.Row, 2)
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = xElem(iRow).iLen
        tFlex.Text = sBuf
        Debug.Print iTmpType = am_Long
        tFlex.SelStart = 0
        With Flex
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If (iTmpType And am_String) = am_String Then If KeyAscii = vbKeySpace Then KeyAscii = 0
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUP Then
        KeyCode = 0
        If miRow > 0 Then
            Flex.Row = miRow
            Flex.SetFocus
        End If
    ElseIf KeyCode = vbKeyDOWN Then
        If miRow < Flex.Rows - 2 Then
            Flex.Row = miRow + 2
            Flex.SetFocus
        End If
    Else
        If (iTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (iTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (iTmpType And am_String) = am_String Then
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93  '<>|[]
            KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex.Text
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyBACK
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then

            End If
        Case vbKeySpace
            KeyAscii = 0
        Case Else
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then
                Select Case KeyAscii
                    Case 33, 34, 39 To 47, 61, 63
                        If Flex.Col = 2 Then KeyAscii = 0
                End Select
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
        
    Dim lVal As Long
    Dim bUpdate As Boolean
    Dim sBuf As String
    
    If Not m_blnAnnullaMod Then
        sBuf = Trim$(tFlex.Text)
        tFlex.Text = sBuf
        Select Case miRow - 1
            Case 0  'Nome
                'If Len(sBuf) Then
                    xElem(miRow - 1).vVal = sBuf
                    bUpdate = True
                'End If
            Case 2  'Len
                lVal = ValInt(sBuf)
                If lVal <> 0 Then
                    xElem(miRow - 1).vVal = lVal
                    bUpdate = True
                End If
            Case 3 'scale
                lVal = ValInt(sBuf)
                'If lVal <> 0 Then
                    xElem(miRow - 1).vVal = lVal
                    bUpdate = True
                'End If
        End Select
    End If
    If bUpdate Then
        Debug.Print Flex.Rows
        Debug.Print Flex.Cols
        
        Flex.TextMatrix(miRow, 2) = tFlex
        mbDirty = True
        cmd(1).Enabled = Len(xElem(0).vVal) <> 0
        
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    
    m_blnAnnullaMod = True
    tFlex.Visible = False
        
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            m_blnAnnullaMod = True
            KeyAscii = 0
            Flex.SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex.SetFocus
    End Select

End Sub

Private Sub cbo_LostFocus()
    
    Dim oType As SQLDMO.SystemDatatype
    Dim iRow As Integer
    
    If Not m_blnAnnullaMod Then
    
        With xElem(miRow - 1)
            .vVal = cbo.Text
            mbDirty = True
            Flex.TextMatrix(miRow, 2) = .vVal
            If miRow - 1 = rDataType Then
                Set oType = objServer.Databases(sDbName).SystemDatatypes(.vVal)
                If oType.Name = "image" Or oType.Name Like "*text" Then
                    xElem(2).vVal = oType.MaximumLength
                Else
                    xElem(2).vVal = oType.MaximumChar
                End If
                xElem(2).vOtherVal = GetItemArr(.vVal)
                                
                xElem(3).vVal = IIf(oType.IsNumeric, 2, 0)
                xElem(3).bDisab = Not oType.IsNumeric

                Set oType = Nothing
                m_blnAnnullaMod = True
                iRow = Flex.Row
                LoadFlex
                Flex.Row = iRow
                Flex.Col = Flex.Cols - 1
            End If
        End With
    End If

    cbo.Visible = False
    cmd(1).Enabled = mbDirty And Len(xElem(0).vVal) <> 0
    m_blnAnnullaMod = True

End Sub

Private Sub SaveUDT()

    Dim lErr As Long, sErr As String
    Dim oUdt As SQLDMO.UserDefinedDatatype2
    Dim oType As SQLDMO.SystemDatatype
    Dim oDb As SQLDMO.Database2
    
    Screen.MousePointer = vbHourglass
    
    If bNew Then
        
        Set oUdt = New SQLDMO.UserDefinedDatatype2
        With oUdt
            .Name = xElem(0).vVal
            
            .BaseType = xElem(1).vVal
            Set oType = objServer.Databases(sDbName).SystemDatatypes(.BaseType)
            If oType.AllowLength Then .Length = CLng(xElem(2).vVal)
            If oType.AllowNulls Then .AllowNulls = CBool(xElem(4).vVal)
            If oType.IsNumeric Then
                .NumericPrecision = CLng(xElem(2).vVal)
                .NumericScale = CLng(xElem(3).vVal)
            End If
'            If xElem(5).vVal <> "(none)" Then .Rule = xElem(5).vVal
'            If xElem(6).vVal <> "(none)" Then .Default = xElem(6).vVal
        
            If xElem(5).vVal <> "(none)" Then
                oUdt.Rule = IIf(InStr(xElem(5).vVal, "."), Mid$(xElem(5).vVal, InStr(xElem(5).vVal, ".") + 1), xElem(5).vVal)
            Else
                oUdt.Rule = ""
            End If
            If xElem(6).vVal <> "(none)" Then
                oUdt.Default = IIf(InStr(xElem(6).vVal, "."), Mid$(xElem(6).vVal, InStr(xElem(6).vVal, ".") + 1), xElem(6).vVal)
            Else
                oUdt.Default = ""
            End If
            
        End With
        On Local Error Resume Next
        Set oDb = objServer.Databases(sDbName)
        lErr = Err.Number
        If lErr = 0 Then
            oDb.UserDefinedDatatypes.Add oUdt
            lErr = Err.Number
        End If
        sErr = Err.Description
        On Local Error GoTo 0
    Else
        On Local Error Resume Next
        Set oUdt = objServer.Databases(sDbName).UserDefinedDatatypes(sStuff)
'        If xElem(5).vVal <> "(none)" Then oUDT.Rule = xElem(5).vVal
'        If xElem(6).vVal <> "(none)" Then .Default = xElem(6).vVal
        If xElem(5).vVal <> "(none)" Then
            oUdt.Rule = IIf(InStr(xElem(5).vVal, "."), Mid$(xElem(5).vVal, InStr(xElem(5).vVal, ".") + 1), xElem(5).vVal)
        Else
            oUdt.Rule = ""
        End If
        If xElem(6).vVal <> "(none)" Then
            oUdt.Default = IIf(InStr(xElem(6).vVal, "."), Mid$(xElem(6).vVal, InStr(xElem(6).vVal, ".") + 1), xElem(6).vVal)
        Else
            oUdt.Default = ""
        End If

        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
    End If
    
    Set oType = Nothing
    Set oUdt = Nothing
    Set oUdt = Nothing
    Screen.MousePointer = vbDefault
    
    If lErr <> 0 Then
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        loadUDT xElem(0).vVal
    End If
End Sub
Private Sub GetArrData(sArr() As String)

    Dim i As Integer
    i = UBound(sData) - 1
    ReDim sArr(i) As String
    For i = LBound(sData) To UBound(sData)
        sArr(i - 1) = sData(i).sName
    Next
    
End Sub
Private Function GetItemArr(ByVal sVal As String) As Integer

    Dim i As Integer, iRet As Integer
    For i = LBound(sData) To UBound(sData)
        If sData(i).sName = sVal Then
            iRet = i
            Exit For
        End If
    Next
    GetItemArr = iRet
End Function
Private Sub TestActiveControl()

    If cbo.Visible Then
        cbo_LostFocus
    ElseIf tFlex.Visible Then
        tFlex_LostFocus
    End If

   On Local Error Resume Next
   If TypeOf Me.ActiveControl Is TextBox Then
      If Me.ActiveControl Is tFlex Then
         tFlex_LostFocus
      End If
   ElseIf TypeOf Me.ActiveControl Is ComboBox Then
      If Me.ActiveControl Is cbo Then
         cbo_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub

