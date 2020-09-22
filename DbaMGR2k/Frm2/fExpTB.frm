VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fExpTB 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "fExpTB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboDefault 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "CboDefault"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   60
      Left            =   0
      ScaleHeight     =   26.127
      ScaleMode       =   0  'User
      ScaleWidth      =   26988
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   2595
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
      Left            =   5160
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4425
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13970
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "16/11/2002"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLsit 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":0168
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":05BC
            Key             =   "FitGrid"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":0716
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpTB.frx":0CC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLsit"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Explorer"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Stop"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   180
         Left            =   2520
         TabIndex        =   2
         Top             =   60
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2280
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1800
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Image imgSplitter 
      Height          =   105
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   2000
      Width           =   4095
   End
End
Attribute VB_Name = "fExpTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sStmGO As String
Private m_sTbName As String
Private m_sFullTbName As String
Private m_sDataBase As String
Private m_sDefOrderBy As String

Const lk_sqlProp_Value_VALUE As String = " -) 1% = 2%"

Private Const k_ColDropTbl = 0
Private Const k_ColSetDateFormat = 1
Private Const k_ColSetNoCount = 2
Private Const k_ColIdentityPreserve = 3
Private Const k_ColIsolationLevel = 4
Private Const k_ColBatchSize = 5
Private Const k_ColExpFile = 6
Private Const k_ColWhere = 7
Private Const k_ColOrderBy = 8
Private Const k_ColTbAlias = 9
Private Const k_ColTOP = 10
Private Const k_ColTimeOut = 11
Private Const k_ColSplitFile = 12
Private Const k_ColScriptTbGeneration = 13
Private Const k_ColScriptIDX = 14
Private Const k_ColScriptFullTextIDX = 15
Private Const k_ColScriptTrigger = 16
Private Const k_ColScriptKeys = 17
    
Private Const k_ColInclude = 7
Private Const k_ColReplaceNull = 8
Private Const k_ColReplaceNullValue = 9
Private Const k_ColPosition = 10
Private Const k_ColAlias = 11


Private Enum cboLoad
    x_DropTable = 0
    x_DateFormat
    x_TbSetting
End Enum

Private Type udtTableExp
    ElementData As elemFlexG
    sDesHelp As String
    EditType As TypeEnum
    bDisab As Boolean
End Type

Private Type udtColumn
    sName As String
    sType As String
    sUdt As String
    bIsIdentity As Boolean
    bIsComputed As Boolean
    bIsTimeStamp As Boolean
    bIsRowGuid As Boolean
    bInclude As Boolean
    bReplaceNull As Boolean
    sReplaceNull As String
    sAlias As String
    iPos As Integer
    bDisabled As Boolean
    bDisableNULL As Boolean
    bExistDefault As Boolean
    bCanBeNull As Boolean
'    bDisableAlias As Boolean
End Type

Private flexColumns() As udtColumn
Private flexTBotions() As udtTableExp
Private flexDesHelp() As String

Private m_blnStopExport As Boolean
Private m_blnAnnullaMod As Boolean
Private mbMoving As Boolean

Private m_Width As Long
Private m_Height As Long
Private bWork As Boolean
Private bLoading As Boolean
Private miFlex As Integer
Private miRow(1)  As Integer
Private miCol As Integer
Private miTmpType As iPropG

Const sglSplitLimit = 500

Public Property Let ObjectsNames(ByVal sDbName As String, ByVal sTbName As String)

    Screen.MousePointer = vbHourglass
    
    m_sTbName = sTbName
    m_sDataBase = sDbName
    LoadDbStructure
    tBar.Buttons(keyExplorer).Enabled = ButtonEnabled()

End Property

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Flex(miFlex).SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex(miFlex).SetFocus
    End Select

End Sub
Private Sub cbo_LostFocus()

    If Not m_blnAnnullaMod Then
        If miFlex = 0 Then
            If miRow(miFlex) - 1 = k_ColDropTbl Or miRow(miFlex) - 1 = k_ColSetDateFormat Or miRow(miFlex) - 1 = k_ColIsolationLevel Then
                flexTBotions(miRow(miFlex) - 1).ElementData.vVal = cbo.Text
                Flex(miFlex).TextMatrix(miRow(miFlex), 2) = flexTBotions(miRow(miFlex) - 1).ElementData.vVal
            End If
        End If
    End If
    cbo.Visible = False
    m_blnAnnullaMod = True

End Sub

Private Sub CboDefault_GotFocus()
    m_blnAnnullaMod = False
    StBar.Panels(1).Text = MyLoadResString(k_sqlValidSqlSyntax)
End Sub
Private Sub CboDefault_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Flex(miFlex).SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex(miFlex).SetFocus
    End Select
End Sub
Private Sub CboDefault_LostFocus()

    If Not m_blnAnnullaMod Then
        If miFlex = 1 Then
            If miCol = k_ColReplaceNullValue Then
                flexColumns(miRow(miFlex) - 1).sReplaceNull = CboDefault.Text
                Flex(miFlex).TextMatrix(miRow(miFlex), miCol) = flexColumns(miRow(miFlex) - 1).sReplaceNull
            End If
        End If
    End If
    
    CboDefault.Visible = False
    m_blnAnnullaMod = True

End Sub

Private Sub Flex_DblClick(Index As Integer)
    
    With Flex(Index)
        If Index = 0 Then
            If .Col = .Cols - 1 Then FlexEdit 0
        Else
            If .Col >= k_ColInclude Then FlexEdit 0
        End If
    End With

End Sub
Private Sub Flex_GotFocus(Index As Integer)

    TestActiveControl
    tFlex.Visible = False
    cbo.Visible = False
    CboDefault.Visible = False
    m_blnAnnullaMod = True
    miFlex = Index

End Sub
Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    iCode = KeyCode
    If iCode = vbKeyF2 Then
        KeyCode = 0
        FlexEdit 0
    End If

End Sub
Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyEscape, vbKeyBACK
        Case vbKeySpace
            FlexEdit KeyAscii
        Case Else
            FlexEdit KeyAscii
    End Select

End Sub
Private Sub Flex_RowColChange(Index As Integer)
    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
       
    If Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            With Flex(Index)
                If .Row <> miRow(Index) Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                                        
                    iRowP = .Row
                    If miRow(Index) > .Rows - 1 Then miRow(Index) = .Rows - 1
                    .Row = miRow(Index)
                    .Col = 0
                    Set Flex(Index).CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1
                    
                    .Col = iCol
                    miRow(Index) = iRowP
                End If
                
                If Index = 0 Then
                    StBar.Panels(1).Text = flexTBotions(miRow(Index) - 1).sDesHelp
                Else
                    StBar.Panels(1).Text = flexDesHelp(.Col - 1)
                End If

            End With
            b = False
        End If
    End If

End Sub
Private Sub Flex_Scroll(Index As Integer)
    tFlex.Visible = False
    cbo.Visible = False
    CboDefault.Visible = False
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    With imgSplitter
        picSplitter.ZOrder
        picSplitter.Move .Left, .Top, .Width, .Height
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = y + imgSplitter.Top
        'picSplitter.ZOrder
        If sglPos < sglSplitLimit Then
            picSplitter.Top = sglSplitLimit
        ElseIf sglPos > Me.Height - sglSplitLimit Then
            picSplitter.Top = Me.Height - sglSplitLimit
        Else
            picSplitter.Top = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    SizeControls picSplitter.Top
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub Form_Activate()
    SizeControls imgSplitter.Top
    bLoading = True
    LoadFlex 0
    LoadFlex 1
    bLoading = False
    Flex_RowColChange 1
    Flex_RowColChange 0
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()

    Flex(0).RowHeightMin = cbo.Height + 10
    Flex(1).RowHeightMin = CboDefault.Height + 10
    LoadInitTB
    imgSplitter.Move 0, sglSplitLimit, Me.ScaleWidth
    
    m_sStmGO = "GO -- " & MyLoadResString(k_sqlStatementGO)
    
    ReDim flexDesHelp(10) As String
    flexDesHelp(0) = MyLoadResString(k_Column_Name)
    flexDesHelp(1) = MyLoadResString(k_Column_Type)
    flexDesHelp(2) = MyLoadResString(k_Is_Identity)
    flexDesHelp(3) = MyLoadResString(k_Computed_Field)
    flexDesHelp(4) = MyLoadResString(k_sqlIsRowGuid)
    flexDesHelp(5) = MyLoadResString(k_sqlIsTimeStamp)
    flexDesHelp(6) = MyLoadResString(k_sqlInclude)
    flexDesHelp(7) = MyLoadResString(k_sqlReplaceNull)
    flexDesHelp(8) = MyLoadResString(k_sqlFieldReplace4Null)
    flexDesHelp(9) = MyLoadResString(k_sqlPosition)
    flexDesHelp(10) = MyLoadResString(k_sqlFieldAlias)
    
        
    m_Width = Me.Width
    m_Height = Me.Height
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keyRun).ToolTipText = MyLoadResString(k_Run) & " {F10}"
    tBar.Buttons(keyExplorer).ToolTipText = MyLoadResString(k_sqlTbarResult)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_RES_FitGrid)
    tBar.Buttons(keyStop).ToolTipText = MyLoadResString(k_HtmKillExport)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource) & " {F12}"
            
    CenterForm Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF10        'Run
            If Shift = 0 Then tBar_ButtonClick tBar.Buttons(keyRun)
        Case vbKeyF12        'refresh
            If Shift = 0 Then tBar_ButtonClick tBar.Buttons(keyRefresh)
    End Select

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = bWork
End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Me.WindowState <> vbMinimized Then
        If Not b Then
            b = True
            SizeControls imgSplitter.Top
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set fExpTB = Nothing
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Enabled Then
        TestActiveControl
        Select Case Button.Key
            Case keyEnd
                If Not bWork Then Unload Me
            Case keyFitGrid
                If Not bWork Then
                    SettaLargFlex Flex(0), 2, 200
                    SettaLargFlex Flex(1), 1, 200
                End If
            Case keyRun
                If Not bWork Then RunExport
            Case keyExplorer
                If Not bWork Then ShowHtm
            Case keyStop
                If bWork Then m_blnStopExport = True
            Case keyRefresh
                bLoading = True
                LoadInitTB
                LoadDbStructure
                LoadFlex 0
                LoadFlex 1
                bLoading = False
                
                Flex_RowColChange 0
                Flex_RowColChange 1
                tBar.Buttons(keyExplorer).Enabled = ButtonEnabled()
        End Select
    End If
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then
        KeyCode = 0
    Else
        If (miTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (miTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (miTmpType And am_String) = am_String Then
        End If
    End If

End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
'        Case 32, 35, 37, 39 To 47, 60 To 62, 63, 91, 93, 94
'
'        Case 33, 34, 36, 38, 58, 59, 64, 92, 96, 123 To 255
'            KeyAscii = 0
'        Case 32
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex.Text = Flex(miFlex).Text
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyBACK
            If (miTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (miTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (miTmpType And am_String) = am_String Then
            End If
        Case Else
            If (miTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (miTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (miTmpType And am_String) = am_String Then
                
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()

    Dim bUpdate As Boolean
    Dim sBuf As String
    Dim iVal As Integer
    Dim iCol As Integer
    Dim iLoop As Integer
    
    If Not m_blnAnnullaMod Then
        If miFlex = 0 Then
            tFlex = Trim$(tFlex.Text)
            sBuf = tFlex.Text
            bUpdate = True
        ElseIf miFlex = 1 Then
            If miCol = k_ColPosition Then
                iVal = ValInt(tFlex.Text)
                If iVal <> 0 Then
                    flexColumns(miRow(miFlex) - 1).iPos = iVal
                    For iCol = 0 To UBound(flexColumns)
                        If iCol <> miRow(miFlex) - 1 Then
                            If flexColumns(iCol).iPos >= iVal Then flexColumns(iCol).iPos = flexColumns(iCol).iPos + 1
                        End If
                    Next
                    bLoading = True
                    BB_SortPosition
                    LoadFlex 1
                    
                    iCol = 1
                    For iLoop = 0 To UBound(flexColumns)
                        If flexColumns(iLoop).iPos = iVal Then
                            iCol = iLoop + 1
                            Exit For
                        End If
                    Next
                    bLoading = False
                    Flex(miFlex).Row = iCol
                    Flex(miFlex).Col = k_ColPosition
                    
                    bUpdate = False
                End If
            ElseIf miCol = k_ColAlias Then
                tFlex = Trim$(tFlex.Text)
                sBuf = tFlex.Text
                flexColumns(miRow(miFlex) - 1).sAlias = sBuf
                bUpdate = True
            End If
        End If
    End If
    If bUpdate Then
        If miFlex = 0 Then
            If miRow(miFlex) - 1 = k_ColBatchSize Then
                flexTBotions(k_ColBatchSize).ElementData.vVal = ValInt(sBuf)
            Else
                flexTBotions(miRow(miFlex) - 1).ElementData.vVal = sBuf
            End If
        End If
        Flex(miFlex).TextMatrix(miRow(miFlex), IIf(miFlex = 0, 2, k_ColAlias)) = sBuf
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    
    m_blnAnnullaMod = True
    tFlex.Visible = False
    Flex(miFlex).SetFocus

End Sub

Private Sub TestActiveControl()

    If cbo.Visible Then
        cbo_LostFocus
    ElseIf CboDefault.Visible Then
        CboDefault_LostFocus
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
      ElseIf Me.ActiveControl Is CboDefault Then
         CboDefault_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub

Private Sub SizeControls(y As Single)

    If Me.Width < m_Width Then Me.Width = m_Width
    If Me.Height < m_Height Then Me.Height = m_Height

    'set the width
    tBar.Refresh

    If y < 1500 Then y = 1500
    If y > (Me.Height - 1500) Then y = Me.Height - 1500
        
    pBar.Width = Me.ScaleWidth - pBar.Left - (k_FlexBevel \ 2)
        
    imgSplitter.Move 0, y, Me.ScaleWidth
    picSplitter.Move imgSplitter.Left, imgSplitter.Top, imgSplitter.Width, imgSplitter.Height
    picSplitter.Width = Me.ScaleWidth
                  
    Flex(0).Move Flex(0).Left, tBar.Height + 20, Me.ScaleWidth - 20, y - Flex(0).Top
    Flex(1).Move Flex(0).Left, picSplitter.Top + picSplitter.Height, Me.ScaleWidth - 20, Me.ScaleHeight - StBar.Height - 100 - y
        
End Sub

Private Sub LoadInitTB()

    ReDim flexTBotions(17) As udtTableExp
    
    flexTBotions(0).ElementData.sField = MyLoadResString(k_sqlTB_opt_Delete)
    flexTBotions(0).ElementData.vVal = ""
    flexTBotions(0).ElementData.iType = am_String
    flexTBotions(0).sDesHelp = MyLoadResString(k_sqlTB_opt_Delete)
    flexTBotions(0).EditType = iCombo

    flexTBotions(1).ElementData.sField = "Set DATEFORMAT"
    flexTBotions(1).ElementData.vVal = "ymd"
    flexTBotions(1).ElementData.iType = am_String
    flexTBotions(1).sDesHelp = MyLoadResString(k_sqlTB_opt_DateFormatHelp)
    flexTBotions(1).EditType = iCombo

    flexTBotions(2).ElementData.sField = "SET NOCOUNT ON"
    flexTBotions(2).ElementData.vVal = True
    flexTBotions(2).ElementData.iType = am_Bool
    flexTBotions(2).sDesHelp = MyLoadResString(k_sqlTB_opt_NoCountHelp)
    flexTBotions(2).EditType = iNormalEdit

    flexTBotions(3).ElementData.sField = MyLoadResString(k_sqlProp_PreserveIdentity)
    flexTBotions(3).ElementData.vVal = False
    flexTBotions(3).bDisab = True
    flexTBotions(3).ElementData.iType = am_Bool
    flexTBotions(3).sDesHelp = MyLoadResString(k_sqlTB_opt_IdentityHelp)
    flexTBotions(3).EditType = iNormalEdit

    flexTBotions(4).ElementData.sField = MyLoadResString(k_sqlProp_TranIsolationLevel)
    flexTBotions(4).ElementData.vVal = ""
    flexTBotions(4).ElementData.iType = am_String
    flexTBotions(4).sDesHelp = MyLoadResString(k_sqlTB_opt_IsolationHelp)
    flexTBotions(4).EditType = iCombo

    flexTBotions(5).ElementData.sField = MyLoadResString(k_sqlProp_BatchSize)
    flexTBotions(5).ElementData.vVal = 1000
    flexTBotions(5).ElementData.iType = am_Long
    flexTBotions(5).ElementData.iLen = 9
    flexTBotions(5).sDesHelp = MyLoadResString(k_sqlTB_opt_BatchHelp)
    flexTBotions(5).EditType = iNormalEdit

    flexTBotions(6).ElementData.sField = MyLoadResString(k_sqlTB_opt_File)
    flexTBotions(6).ElementData.iType = am_String
    flexTBotions(6).sDesHelp = MyLoadResString(k_sqlTB_opt_FileHelp)
    'flexTBotions(6).ElementData.vVal = GetSpecialFolderLocation(Me.hwnd, CSIDL_PERSONAL) & "\" & ReplaceMsg(m_sFullTbName, Array("[", "]", "."), Array("", "", "_")) & ".Sql"
    flexTBotions(6).EditType = iEllipsis

    flexTBotions(7).ElementData.sField = MyLoadResString(k_sqlTB_opt_Where)
    flexTBotions(7).ElementData.vVal = ""
    flexTBotions(7).ElementData.iType = am_String
    flexTBotions(7).ElementData.iLen = 200
    flexTBotions(7).sDesHelp = MyLoadResString(k_sqlTB_opt_WhereHelp)
    flexTBotions(7).EditType = iNormalEdit

    flexTBotions(8).ElementData.sField = MyLoadResString(k_sqlTB_opt_OrderBy)
    flexTBotions(8).ElementData.vVal = ""
    flexTBotions(8).ElementData.iType = am_String
    flexTBotions(8).ElementData.iLen = 200
    flexTBotions(8).sDesHelp = MyLoadResString(k_sqlTB_opt_OrderByHelp)
    flexTBotions(8).EditType = iNormalEdit

    flexTBotions(9).ElementData.sField = MyLoadResString(k_sqlTB_opt_TbAlias)
    flexTBotions(9).ElementData.vVal = ""
    flexTBotions(9).ElementData.iType = am_String
    flexTBotions(9).ElementData.iLen = 200
    flexTBotions(9).sDesHelp = MyLoadResString(k_sqlTB_opt_TbAliasHelp)
    flexTBotions(9).EditType = iNormalEdit

    flexTBotions(10).ElementData.sField = MyLoadResString(k_sqlTB_opt_Top)
    flexTBotions(10).ElementData.vVal = 0
    flexTBotions(10).ElementData.iType = am_Long
    flexTBotions(10).ElementData.iLen = 9
    flexTBotions(10).sDesHelp = MyLoadResString(k_sqlTB_opt_TopHelp)
    flexTBotions(10).EditType = iNormalEdit

    flexTBotions(11).ElementData.sField = MyLoadResString(k_sqlTB_opt_TimeOut)
    flexTBotions(11).ElementData.vVal = 30
    flexTBotions(11).ElementData.iType = am_Long
    flexTBotions(11).ElementData.iLen = 3
    flexTBotions(11).sDesHelp = MyLoadResString(k_sqlTB_opt_TimeOutHelp)
    flexTBotions(11).EditType = iNormalEdit
    
    flexTBotions(12).ElementData.sField = MyLoadResString(k_sqlTB_opt_SlitFile)
    flexTBotions(12).ElementData.vVal = 0
    flexTBotions(12).ElementData.iType = am_Long
    flexTBotions(12).ElementData.iLen = 9
    flexTBotions(12).sDesHelp = MyLoadResString(k_sqlTB_opt_SlitFileHelp)
    flexTBotions(12).EditType = iNormalEdit
    
'<-- Script
    flexTBotions(13).ElementData.sField = MyLoadResString(k_sqlTB_opt_ScriptTable)
    flexTBotions(13).ElementData.vVal = False
    flexTBotions(13).bDisab = False
    flexTBotions(13).ElementData.iType = am_Bool
    flexTBotions(13).sDesHelp = MyLoadResString(k_sqlTB_opt_ScriptTableHelp)
    flexTBotions(13).EditType = iNormalEdit
    
    flexTBotions(14).ElementData.sField = MyLoadResString(k_Script_Indexes)
    flexTBotions(14).ElementData.vVal = False
    flexTBotions(14).bDisab = True
    flexTBotions(14).ElementData.iType = am_Bool
    flexTBotions(14).sDesHelp = flexTBotions(14).ElementData.sField
    flexTBotions(14).EditType = iNormalEdit
    
    flexTBotions(15).ElementData.sField = MyLoadResString(k_Script_Full_Text_Indexes)
    flexTBotions(15).ElementData.vVal = False
    flexTBotions(15).bDisab = True
    flexTBotions(15).ElementData.iType = am_Bool
    flexTBotions(15).sDesHelp = flexTBotions(15).ElementData.sField
    flexTBotions(15).EditType = iNormalEdit
    
    flexTBotions(16).ElementData.sField = MyLoadResString(k_Script_Triggers)
    flexTBotions(16).ElementData.vVal = False
    flexTBotions(16).bDisab = True
    flexTBotions(16).ElementData.iType = am_Bool
    flexTBotions(16).sDesHelp = flexTBotions(16).ElementData.sField
    flexTBotions(16).EditType = iNormalEdit
    
    flexTBotions(17).ElementData.sField = MyLoadResString(k_sqlTB_opt_ScriptKeys)
    flexTBotions(17).ElementData.vVal = False
    flexTBotions(17).bDisab = True
    flexTBotions(17).ElementData.iType = am_Bool
    flexTBotions(17).sDesHelp = MyLoadResString(k_sqlTB_opt_ScriptKeysHelp)
    flexTBotions(17).EditType = iNormalEdit
       
End Sub
Private Sub LoadFlex(ByVal iFlex As Integer)

    Dim i As Integer, l As Integer
    Dim sBuf As String
    
    With Flex(iFlex)
        .Redraw = False
        .Clear
        .FixedCols = 0
        
        .Rows = 1
        .Clear
        If iFlex = 0 Then
            .Cols = 3
            .TextMatrix(0, 1) = MyLoadResString(k_Property)
            .TextMatrix(0, 2) = MyLoadResString(k_Value)
            
            .ColWidth(0) = k_FlexPARAM_cellaSelector
            .ColWidth(1) = 2000
        Else
            .Cols = 12
            .ColWidth(0) = k_FlexPARAM_cellaSelector
            .ColWidth(1) = 1200
            .ColWidth(2) = 1000
            .ColWidth(3) = 600
            .ColWidth(4) = 600
            .ColWidth(5) = 600
            .ColWidth(6) = 600
            .ColWidth(7) = 600
            .ColWidth(8) = 600
            .ColWidth(9) = 1200
            .ColWidth(10) = 800
            .ColWidth(11) = 1200
            
            .TextMatrix(0, 1) = MyLoadResString(k_Column_Name)
            .TextMatrix(0, 2) = MyLoadResString(k_Column_Type)
            .TextMatrix(0, 3) = MyLoadResString(k_Is_Identity)
            .TextMatrix(0, 4) = MyLoadResString(k_Computed_Field)
            .TextMatrix(0, 5) = MyLoadResString(k_sqlIsRowGuid)
            .TextMatrix(0, 6) = MyLoadResString(k_sqlIsTimeStamp)
            .TextMatrix(0, 7) = MyLoadResString(k_sqlInclude)
            .TextMatrix(0, 8) = MyLoadResString(k_sqlReplaceNull)
            .TextMatrix(0, 9) = MyLoadResString(k_sqlFieldReplace4Null)
            .TextMatrix(0, 10) = MyLoadResString(k_sqlPosition)
            .TextMatrix(0, 11) = MyLoadResString(k_sqlFieldAlias)
        End If
        
        If iFlex = 0 Then
            For i = 0 To UBound(flexTBotions)
                sBuf = ""
                If flexTBotions(i).ElementData.iType = am_String Then
                    sBuf = sBuf & flexTBotions(i).ElementData.vVal
                ElseIf flexTBotions(i).ElementData.iType = am_Long Then
                    sBuf = sBuf & Format$(flexTBotions(i).ElementData.vVal, "#,0")
                End If
                Flex(iFlex).AddItem vbTab & flexTBotions(i).ElementData.sField & vbTab & sBuf
                
                If flexTBotions(i).ElementData.iType = am_Bool Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexTBotions(i).ElementData.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                
                If flexTBotions(i).EditType = iCombo Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(iFlex).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                    .CellPictureAlignment = flexAlignRightTop
                ElseIf flexTBotions(i).EditType = iEllipsis Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(iFlex).CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
                    .CellPictureAlignment = flexAlignRightTop
                End If
                
                If flexTBotions(i).bDisab Then
                    .Col = .Cols - 1
                    .CellBackColor = vbInactiveBorder
                End If
            Next
        Else
            For i = 0 To UBound(flexColumns)
                sBuf = IIf(Len(flexColumns(i).sUdt) <> 0, flexColumns(i).sUdt & " (", "") & flexColumns(i).sType & IIf(Len(flexColumns(i).sUdt) <> 0, ")", "")
                sBuf = flexColumns(i).sName & vbTab & sBuf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flexColumns(i).sReplaceNull & vbTab & flexColumns(i).iPos & vbTab & flexColumns(i).sAlias
                .AddItem vbTab & sBuf
                
                .Row = .Rows - 1
                .Col = 1
                .CellBackColor = vbInactiveBorder
                .Col = 2
                .CellBackColor = vbInactiveBorder
                
                .Col = 3
                .CellBackColor = vbInactiveBorder
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bIsIdentity), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 4
                .CellBackColor = vbInactiveBorder
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bIsComputed), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 5
                .CellBackColor = vbInactiveBorder
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bIsRowGuid), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 6
                .CellBackColor = vbInactiveBorder
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bIsTimeStamp), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 7
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bInclude), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 8
                Set Flex(iFlex).CellPicture = LoadResPicture(IIf(Abs(flexColumns(i).bReplaceNull), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If flexColumns(i).bDisableNULL Then .CellBackColor = vbInactiveBorder
                
                .Col = 9
                Set Flex(iFlex).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                .CellPictureAlignment = flexAlignRightTop
                If Not flexColumns(i).bReplaceNull Then .CellBackColor = vbInactiveBorder
                .ColAlignment(k_ColReplaceNullValue) = flexAlignLeftCenter
                
                If flexColumns(i).bDisabled Then
                    For l = 1 To .Cols - 1
                        .Col = l
                        .CellBackColor = vbInactiveBorder
                    Next
                End If
                
            Next
            If .Rows = 1 Then
                .AddItem ""
                For l = 1 To .Cols - 1
                    .Col = l
                    .CellBackColor = vbInactiveBorder
                Next
            End If
        End If
        
        .FixedRows = 1
        If iFlex = 0 Then
            .FixedCols = 2
        Else
            .FixedCols = 1
            miCol = 0
        End If
        
        .Redraw = True
        SettaLargFlex Flex(iFlex), IIf(iFlex = 0, 2, 1), .ColWidth(2)
        miRow(iFlex) = 0
        
        .TopRow = 1
        .Row = 1
        .Col = IIf(iFlex = 0, 2, 1)
    End With
    
End Sub
Private Sub FlexEdit(ByVal iKeyAscii As Integer)
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iImg As Integer
    Dim bCombo As Boolean, bPush As Boolean
    
    miTmpType = am_String
    iRow = Flex(miFlex).Row
    iCol = Flex(miFlex).Col
    miCol = iCol
    bEdit = True

    If Flex(miFlex).CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        If miFlex = 0 Then
            miTmpType = flexTBotions(iRow - 1).ElementData.iType
            If miTmpType = am_Long Then
                sBuf = Format$(flexTBotions(iRow - 1).ElementData.vVal, "#,0")
            Else
                sBuf = flexTBotions(iRow - 1).ElementData.vVal
            End If
            iLen = flexTBotions(iRow - 1).ElementData.iLen
        Else
            Select Case iCol
                Case Is < k_ColInclude
                    bEdit = False
                    Exit Sub
                Case k_ColInclude To k_ColReplaceNull
                    miTmpType = am_Bool
                Case k_ColReplaceNullValue
                    miTmpType = am_String
                    sBuf = flexColumns(iRow - 1).sReplaceNull
                Case k_ColPosition
                    miTmpType = am_Long
                    iLen = 3
                    sBuf = flexColumns(iRow - 1).iPos
                Case k_ColAlias
                    miTmpType = am_String
                    sBuf = flexColumns(iRow - 1).sAlias
                    iLen = 50
            End Select
        End If
        
        miRow(miFlex) = iRow
        Select Case iKeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0, vbKeySpace
                If miFlex = 0 Then
                    If (miTmpType And am_Bool) = am_Bool Then
                        iImg = 1 - (-flexTBotions(iRow - 1).ElementData.vVal)
                        flexTBotions(iRow - 1).ElementData.vVal = CBool(iImg)
                        Set Flex(miFlex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        bEdit = False
                        
                        If iRow - 1 = k_ColScriptTbGeneration Then
                            Flex(miFlex).Redraw = False
                            bLoading = True
                            For iCol = k_ColScriptTbGeneration + 1 To k_ColScriptKeys
                                flexTBotions(iCol).bDisab = Not CBool(iImg)
                                Flex(miFlex).Row = iCol + 1
                                Flex(miFlex).Col = Flex(miFlex).Cols - 1
                                Flex(miFlex).CellBackColor = IIf(Not CBool(iImg), vbInactiveBorder, vbWindowBackground)
                            Next
                            
                            'Alias Tabella
                            flexTBotions(k_ColTbAlias).bDisab = CBool(iImg)
                            If flexTBotions(k_ColTbAlias).bDisab Then
                                flexTBotions(k_ColTbAlias).ElementData.vVal = ""
                                Flex(miFlex).TextMatrix(k_ColTbAlias + 1, Flex(miFlex).Cols - 1) = ""
                            End If
                            Flex(miFlex).Row = k_ColTbAlias + 1
                            Flex(miFlex).Col = Flex(miFlex).Cols - 1
                            Flex(miFlex).CellBackColor = IIf(flexTBotions(k_ColTbAlias).bDisab, vbInactiveBorder, vbWindowBackground)
                            
                            'Alias Campi
                            Flex(1).Redraw = False
                            iLen = Flex(1).Row
                            For iCol = 0 To UBound(flexColumns)
                                If CBool(iImg) Then
                                    flexColumns(iCol).sAlias = ""
                                    Flex(1).TextMatrix(iCol + 1, Flex(1).Cols - 1) = ""
                                End If
                                Flex(1).Row = iCol + 1
                                Flex(1).Col = Flex(1).Cols - 1
                                Flex(1).CellBackColor = IIf(CBool(iImg), vbInactiveBorder, vbWindowBackground)
                            Next
                            Flex(1).Row = iLen
                            Flex(1).Redraw = True
                            
                            Flex(miFlex).Redraw = True
                            Flex(miFlex).Row = iRow
                            bLoading = False
                        End If
                    End If
                    If flexTBotions(iRow - 1).EditType = iCombo Then
                    
                        If iRow - 1 = k_ColSetDateFormat Then
                            LoadCombo x_DateFormat
                        ElseIf iRow - 1 = k_ColIsolationLevel Then
                            LoadCombo x_TbSetting
                        ElseIf iRow - 1 = k_ColDropTbl Then
                            LoadCombo x_DropTable
                        End If

                        bCombo = True
                    ElseIf flexTBotions(iRow - 1).EditType = iEllipsis Then
                        bPush = True
                    End If
                Else
                    If (miTmpType And am_Bool) = am_Bool Then
                        Select Case iCol
                            Case k_ColInclude
                                iImg = 1 - (-flexColumns(iRow - 1).bInclude)
                                flexColumns(iRow - 1).bInclude = CBool(iImg)
                            Case k_ColReplaceNull
                                iImg = 1 - (-flexColumns(iRow - 1).bReplaceNull)
                                flexColumns(iRow - 1).bReplaceNull = CBool(iImg)
                                bLoading = True
                                Flex(miFlex).Col = miCol + 1
                                Flex(miFlex).CellBackColor = IIf(flexColumns(iRow - 1).bReplaceNull, vbWindowBackground, vbInactiveBorder)
                                Flex(miFlex).Col = miCol
                                bLoading = False
                        End Select
                            
                        Set Flex(miFlex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        bEdit = False
                    ElseIf miCol = k_ColReplaceNullValue Then
                        bCombo = True
                    End If
                End If
            Case Else
                If (miTmpType And am_Bool) = am_Bool Then bEdit = False
                
                If miFlex = 0 Then
                    If flexTBotions(iRow - 1).EditType = iCombo Then
                        If iRow - 1 = k_ColSetDateFormat Then
                            LoadCombo x_DateFormat
                        ElseIf iRow - 1 = k_ColIsolationLevel Then
                            LoadCombo x_TbSetting
                        ElseIf iRow - 1 = k_ColDropTbl Then
                            LoadCombo x_DropTable
                        End If
                        bCombo = True
                    End If
                Else
                    If miCol = k_ColReplaceNullValue Then bCombo = True
                End If
        End Select
    End If

    If bEdit Then
        If Not bCombo And Not bPush Then
            tFlex.PasswordChar = ""
            tFlex.MaxLength = 0
            tFlex.Text = ""
            tFlex.MaxLength = iLen
            tFlex.Text = sBuf
            tFlex.SelStart = 0
            With Flex(miFlex)
                tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
            End With
            tFlex.Visible = True
            tFlex.ZOrder
            tFlex.SetFocus
            If (miTmpType And am_String) = am_String Then If iKeyAscii = vbKeySpace Then iKeyAscii = 0
            If iKeyAscii <> 0 Then
                On Local Error Resume Next
                SendKeys Chr$(iKeyAscii)
                On Local Error GoTo 0
            End If
        ElseIf bCombo Then
            If miFlex = 0 Then
                iLen = GetItem(sBuf, cbo)
            
                cbo.ListIndex = iLen
                With Flex(miFlex)
                    cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                End With
                cbo.Visible = True
                cbo.ZOrder
                cbo.SetFocus
            Else
                LoadComboDefault miRow(miFlex) - 1
                iLen = GetItem(sBuf, CboDefault)
                CboDefault.ListIndex = iLen
                With Flex(miFlex)
                    CboDefault.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                End With
                CboDefault.Visible = True
                CboDefault.ZOrder
                CboDefault.SetFocus
            End If
            If (miTmpType And am_String) = am_String Then If iKeyAscii = vbKeySpace Then iKeyAscii = 0
            If iKeyAscii <> 0 Then
                On Local Error Resume Next
                SendKeys Chr$(iKeyAscii)
                On Local Error GoTo 0
            End If
        ElseIf bPush Then
            SaveFileRequest
        End If
    End If

End Sub

Private Sub LoadCombo(ByVal iType As cboLoad)

    With cbo
        .Clear
        If iType = x_TbSetting Then
            .AddItem ""
            .AddItem "READ COMMITTED"
            .AddItem "REPEATABLE READ"
            .AddItem "SERIALIZABLE"
        ElseIf iType = x_DateFormat Then
            .AddItem ""
            .AddItem "ymd"
            .AddItem "ydm"
            .AddItem "mdy"
        ElseIf iType = x_DropTable Then
            .AddItem ""
            .AddItem "DELETE TABLE"
            .AddItem "TRUNCATE TABLE"
        End If
        
    End With
End Sub
Private Sub LoadComboDefault(ByVal iElem As Integer)

    Dim iItem As Integer
    Dim sBuf As String
    
    With CboDefault
        .Clear
        '.AddItem ""
        If Len(flexColumns(iElem).sReplaceNull) <> 0 Then .AddItem flexColumns(iElem).sReplaceNull
        Select Case LCase$(flexColumns(iElem).sType)
            Case "char", "varchar", "nchar", "nvarchar", "text", "ntext"
                sBuf = "''"
                iItem = GetItem(sBuf, CboDefault)
                If iItem = -1 Then .AddItem sBuf
            Case "datetime", "smalldatetime"
                sBuf = "getdate()"
                iItem = GetItem(sBuf, CboDefault)
                If iItem = -1 Then .AddItem sBuf
            Case "uniqueidentifier"
                sBuf = "newid()"
                iItem = GetItem(sBuf, CboDefault)
                If iItem = -1 Then .AddItem sBuf
            Case "binary", "varbinary"
            Case "timestamp", "rowversion"
            Case Else
                sBuf = "0"
                iItem = GetItem(sBuf, CboDefault)
                If iItem = -1 Then .AddItem sBuf
        End Select
        If flexColumns(iElem).bExistDefault Then
            sBuf = "DEFAULT"
            iItem = GetItem(sBuf, CboDefault)
            If iItem = -1 Then .AddItem sBuf
        End If
        If flexColumns(iElem).bCanBeNull Then
            sBuf = "NULL"
            iItem = GetItem(sBuf, CboDefault)
            If iItem = -1 Then .AddItem sBuf
        End If
    End With
    
End Sub

Private Sub LoadDbStructure()

    Dim iCol As Integer
    Dim bUDT As Boolean
    Dim sBuf As String
    
    Dim oType As SQLDMO.SystemDatatype
    ReDim flexColumns(0) As udtColumn
    flexColumns(0).bDisabled = True
        
    Dim oDb As SQLDMO.Database2
    Dim oTb As SQLDMO.Table
    Dim oCol As SQLDMO.Column2
    Set oTb = GetTbByName(objServer, m_sDataBase, m_sTbName)
    
    
    m_sDefOrderBy = ""
    If Not oTb Is Nothing Then
        Set oDb = oTb.Parent
        m_sFullTbName = "[" & oTb.Owner & "].[" & oTb.Name & "]"
        iCol = oTb.Columns.Count - 1
        ReDim flexColumns(iCol) As udtColumn
                
        iCol = 0
        For Each oCol In oTb.Columns
            flexColumns(iCol).sName = oCol.Name
            flexColumns(iCol).bInclude = True
            flexColumns(iCol).bIsComputed = oCol.IsComputed
            flexColumns(iCol).bIsIdentity = oCol.Identity
            flexColumns(iCol).bIsRowGuid = oCol.IsRowGuidCol
            
            flexColumns(iCol).bCanBeNull = oCol.AllowNulls
            
            flexColumns(iCol).bReplaceNull = False
            flexColumns(iCol).sAlias = ""
            flexColumns(iCol).iPos = iCol + 1
                                                
            On Local Error Resume Next
            Set oType = oDb.SystemDatatypes(oCol.DataType)
            bUDT = Err.Number <> 0
            On Local Error GoTo 0
            
            If Not bUDT Then
                flexColumns(iCol).sType = oCol.DataType
                sBuf = Trim$(oCol.DRIDefault.Text)
            Else
                flexColumns(iCol).sType = oDb.UserDefinedDatatypes(oCol.DataType).BaseType
                flexColumns(iCol).sUdt = oDb.UserDefinedDatatypes(oCol.DataType).Owner & "." & oCol.DataType
                                
                sBuf = Trim$(oDb.UserDefinedDatatypes(oCol.DataType).Default)
            End If
            'Default
            If Len(sBuf) Then
                If Left$(sBuf, 1) = "(" Then sBuf = Mid$(sBuf, 2)
                If Right$(sBuf, 1) = ")" Then
                    If Right$(sBuf, 2) <> "()" Then sBuf = Left$(sBuf, Len(sBuf) - 1)
                End If
                flexColumns(iCol).sReplaceNull = sBuf
            End If
            flexColumns(iCol).bExistDefault = Len(sBuf) <> 0
            
            
            'If StrComp(flexColumns(iCol).sType, "uniqueidentifier", vbTextCompare) = 0 Then
            '    flexColumns(iCol).bDisableNULL = True
            'End If
            
            flexColumns(iCol).bIsTimeStamp = (StrComp(flexColumns(iCol).sType, "timestamp", vbTextCompare) = 0) Or (StrComp(flexColumns(iCol).sType, "ROWVERSION", vbTextCompare) = 0)
            
            If flexColumns(iCol).bIsIdentity Then flexTBotions(k_ColIdentityPreserve).bDisab = False
            
            If flexColumns(iCol).bIsIdentity Or flexColumns(iCol).bIsComputed Or flexColumns(iCol).bIsTimeStamp Or flexColumns(iCol).bIsRowGuid Then
                flexColumns(iCol).bInclude = False
            End If
            
            If (StrComp(flexColumns(iCol).sType, "image", vbTextCompare) = 0) Or (StrComp(flexColumns(iCol).sType, "sql_variant", vbTextCompare) = 0) Then
                flexColumns(iCol).bDisabled = True
                flexColumns(iCol).bInclude = False
            End If
                        
            iCol = iCol + 1
        Next
        
        If oTb.Keys.Count Then
            Dim oKey As SQLDMO.Key
            For Each oKey In oTb.Keys
                If (oKey.Type = SQLDMOKey_Primary) Or (oKey.Clustered) Then
                    For iCol = 1 To oKey.KeyColumns.Count
                        m_sDefOrderBy = m_sDefOrderBy & IIf(Len(m_sDefOrderBy) <> 0, ", ", "")
                        m_sDefOrderBy = m_sDefOrderBy & "[" & oKey.KeyColumns(iCol) & "]"
                    Next
                    Exit For
                End If
            Next
            Set oKey = Nothing
        End If
        If Len(m_sDefOrderBy) = 0 Then
            If oTb.HasClusteredIndex Then
                Dim oIx As SQLDMO.Index
                For Each oIx In oTb.Indexes
                    If (oIx.Type And SQLDMOIndex_Clustered) = SQLDMOIndex_Clustered Then
                        For Each oCol In oIx.ListIndexedColumns
                            m_sDefOrderBy = m_sDefOrderBy & IIf(Len(m_sDefOrderBy) <> 0, ", ", "")
                            m_sDefOrderBy = m_sDefOrderBy & "[" & oCol.Name & "]"
                        Next
                        Exit For
                    End If
                Next
                Set oIx = Nothing
            End If
        End If
    End If
    flexTBotions(k_ColScriptTbGeneration).bDisab = oTb.SystemObject
    flexTBotions(k_ColOrderBy).ElementData.vVal = m_sDefOrderBy
    flexTBotions(k_ColExpFile).ElementData.vVal = GetSpecialFolderLocation(Me.hWnd, CSIDL_PERSONAL) & "\" & ReplaceMsg(m_sFullTbName, Array("[", "]", "."), Array("", "", "_")) & ".Sql"
    
    Set oCol = Nothing
    Set oTb = Nothing
    Set oDb = Nothing
    
End Sub
Public Sub BB_SortPosition()

    Dim Value As Integer
    Dim Index As Long
    Dim fItem As Long
    Dim indLimit As Long, lastSwap As Long
    Dim numEl As Long

        
    Dim TmpUDT As udtColumn
    
    numEl = UBound(flexColumns)
            
    fItem = 0
    lastSwap = UBound(flexColumns)
    Do
        indLimit = lastSwap - 1
        lastSwap = 0
        For Index = fItem To indLimit
            TmpUDT = flexColumns(Index)
            Value = TmpUDT.iPos
            If Value > flexColumns(Index + 1).iPos Then
                flexColumns(Index) = flexColumns(Index + 1)
                flexColumns(Index + 1) = TmpUDT

                lastSwap = Index
            End If
        Next
    Loop While lastSwap
    
    For fItem = 0 To UBound(flexColumns)
        flexColumns(fItem).iPos = fItem + 1
    Next
            
End Sub
Private Sub SaveFileRequest()

    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim obj As cOpen
    Set obj = New cOpen
    
    Dim sTitle As String
    Dim sFilter As String
    Dim sDefDir As String
    Dim sDefFileName As String
    
    sDefDir = flexTBotions(k_ColExpFile).ElementData.vVal
    If Len(sDefDir) Then
        sDefDir = Left$(sDefDir, InStrRev(sDefDir, "\"))
    End If
    
    sTitle = MyLoadResString(k_sqlDlgFileEport)
    sFilter = "Sql File (*.Sql)" & Chr$(0) & "*.sql" & Chr$(0) & "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    
    
    sDefFileName = ReplaceMsg(m_sFullTbName, Array("[", "]", "."), Array("", "", "_")) & ".Sql"
    
    obj.Title = sTitle
    obj.Def_Dir() = sDefDir
    obj.Filter() = sFilter
    obj.f_Name = sDefFileName
    obj.formHwnd() = Me.hWnd

    obj.Do_Save
    lErr = obj.ErrCode()
    If lErr = 0 Then
        sFile = obj.FullPath()
    Else
        If lErr = 1 Then
            lErr = 0
        Else
            sErr = obj.ErrMsg()
        End If
    End If
    Set obj = Nothing
    
    If lErr = 0 Then
        If Len(sFile) Then
            flexTBotions(k_ColExpFile).ElementData.vVal = sFile
            Flex(0).TextMatrix(k_ColExpFile + 1, 2) = sFile
        End If
    End If
    tBar.Buttons(keyExplorer).Enabled = ButtonEnabled()
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If

End Sub

Private Sub RunExport()
    
    Dim iFile As Integer
    Dim lErr As Long, sErr As String
    Dim sBuf As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sConnString As String
    
    Dim sSqlCommand As String
    Dim sSqlInsert As String
    Dim lTotRows As Long, lCurRow As Long
    Dim iTotCols As Integer, iCol As Integer
    Dim lBatchSize As Long
    Dim tInit As Date
    Dim iFileNum As Integer, lngSlitNew As Long
    
    
    m_blnStopExport = False
        
    lngSlitNew = flexTBotions(k_ColSplitFile).ElementData.vVal
    lBatchSize = flexTBotions(k_ColBatchSize).ElementData.vVal
    
    If Len(flexTBotions(k_ColExpFile).ElementData.vVal) = 0 Then
        lErr = 52
        On Local Error Resume Next
        Err.Raise lErr
        Open "" For Input As iFile
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    Else
        sSqlCommand = GetSqlCommand(sSqlInsert)
        If Len(sSqlCommand) = 0 Then
            lErr = Err_Free
            sErr = MyLoadResString(k_sqlErrSqlStatement)
        End If
    End If
    
    If lErr = 0 Then
        Screen.MousePointer = vbHourglass
        bWork = True
        tBar.Buttons(keyEnd).Enabled = False
        tBar.Buttons(keyRun).Enabled = False
        tBar.Buttons(keyFitGrid).Enabled = False
        tBar.Buttons(keyRefresh).Enabled = False
        tBar.Buttons(keyExplorer).Enabled = False
        
        pBar.Value = 0
    
        
        If CBool(Val(GetIni("CONNECT", "TRUSTED", App.Path & "\" & App.EXEName & ".INI"))) Then
            sConnString = "Initial Catalog=" & m_sDataBase & "; Integrated Security = SSPI;"
        Else
            'Standard
            sConnString = "Initial Catalog=" & m_sDataBase & ";User Id= " & objServer.Login & ";" & IIf(Len(objServer.Password), " Password= " & objServer.Password & ";", "")
        End If
        lErr = ErrMyOpenConn(adoConn, sConnString, Nothing, iTimeOut:=IIf(flexTBotions(k_ColTimeOut).ElementData.vVal <> 0, flexTBotions(k_ColTimeOut).ElementData.vVal, 30))
        
        If lErr = 0 And Not adoConn Is Nothing Then
            If adoConn.State = adStateOpen Then
                Flex(0).Enabled = False
                Flex(1).Enabled = False
                tInit = Now()
                
                StBar.Panels(1).Text = MyLoadResString(k_sqlGettingResult)
                Set adoRS = New ADODB.Recordset
                
                On Local Error Resume Next
                adoRS.Open sSqlCommand, adoConn, adOpenKeyset, adLockReadOnly, adCmdText
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                
                If lErr = 0 And (Not adoRS Is Nothing) Then
                    If adoRS.State <> adStateOpen Then
                        lErr = Err_Free
                        sErr = MyLoadResString(k_sqlErrAdoConnOpening)
                    End If
                End If
            End If
        End If
    End If
    
    If lErr = 0 Then
        lErr = ErrOpenNewFile(iFile, iFileNum, sErr)
        
        If lErr = 0 Then
            tBar.Buttons(keyStop).Enabled = True
        
            If CBool(flexTBotions(k_ColScriptTbGeneration).ElementData.vVal) Then lErr = ErrScriptTable(iFile, sErr)
            
            If lErr = 0 Then
                If Len(flexTBotions(k_ColIsolationLevel).ElementData.vVal) Then
                    Print #iFile, "SET TRANSACTION ISOLATION LEVEL "; flexTBotions(k_ColIsolationLevel).ElementData.vVal
                End If
                
                If Len(flexTBotions(k_ColDropTbl).ElementData.vVal) Then
                    If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT ON"
                    
                    Print #iFile, ""
                    Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlAction_Deleting), 74) & " */"
                    Print #iFile, ""
                    sBuf = "IF EXISTS ( SELECT * FROM " & m_sFullTbName & " ) BEGIN" & vbCrLf & Space(5)
                    If StrComp(flexTBotions(k_ColDropTbl).ElementData.vVal, "DELETE TABLE", vbTextCompare) = 0 Then
                        sBuf = sBuf & "DELETE FROM "
                    Else
                        sBuf = sBuf & "TRUNCATE TABLE "
                    End If
                    sBuf = sBuf & m_sFullTbName & vbCrLf & "END"
                    Print #iFile, sBuf
                    
                    If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT OFF"
                    Print #iFile, m_sStmGO
                End If
            
                If Len(flexTBotions(k_ColSetDateFormat).ElementData.vVal) <> 0 Then Print #iFile, "SET DATEFORMAT " & flexTBotions(k_ColSetDateFormat).ElementData.vVal
                
                Print #iFile, ""
                Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlAction_Loading), 74) & " */"
                Print #iFile, ""
                If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT ON"
                Print #iFile, ""
                    
                If flexTBotions(k_ColIdentityPreserve).ElementData.vVal Then Print #iFile, "SET IDENTITY_INSERT ON"
                Print #iFile, ""
                    
                With adoRS
                    lTotRows = .RecordCount
                    iTotCols = adoRS.Fields.Count
                    StBar.Panels(1).Text = MyLoadResString(k_sqlWritingResult)
                    DoEvents
                                    
                    While Not .EOF And Not m_blnStopExport And lErr = 0
                        lCurRow = lCurRow + 1
                        sBuf = ""
                        For iCol = 0 To iTotCols - 1
                            sBuf = sBuf & IIf(Len(sBuf) = 0, "", ", ")
                            sBuf = sBuf & .Fields(iCol).Value
                        Next
                    
                        Print #iFile, sSqlInsert
                        Print #iFile, " VALUES ( " & sBuf & " )"
                        Print #iFile, ""
                        
                        If lBatchSize <> 0 Then
                            If lCurRow Mod lBatchSize = 0 Then
                                Print #iFile, m_sStmGO
                                Print #iFile, ""
                            End If
                        End If
                        If lngSlitNew <> 0 Then
                            If lCurRow Mod lngSlitNew = 0 Then
                                Print #iFile, m_sStmGO
                                
                                If flexTBotions(k_ColIdentityPreserve).ElementData.vVal Then Print #iFile, "SET IDENTITY_INSERT OFF"
                                Print #iFile, ""
                            
                                If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT OFF"
                                Print #iFile, ""
                                Print #iFile, "/* " & sPad(False, "", 74) & " */"
                                Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlWritingClosingFileNum), "1%", iFileNum), 74) & " */"
                                Print #iFile, "/* " & sPad(False, "", 74) & " */"
                                
                                Print #iFile, ""
                                Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlAction_End), 74) & " */"
                                Print #iFile, "/* " & sPad(False, "Sql Insert Script " & Replace(MyLoadResString(k_HtmProducedBy), "1%", App.EXEName & "  " & App.Major & "." & App.Minor & "." & App.Revision), 74) & " */"
                                Print #iFile, ""
                                Close #iFile
                                iFile = 0
                                
                                lErr = ErrOpenNewFile(iFile, iFileNum, sErr)
                            End If
                        End If
                        If lErr = 0 Then
                            pBar.Value = ((.AbsolutePosition) * 100) \ lTotRows
                            .MoveNext
                            DoEvents
                        End If
                    Wend
                End With
            End If
            If lErr = 0 Then
                Print #iFile, m_sStmGO
                Print #iFile, ""
            
                If m_blnStopExport Then Print #iFile, "/* " & sPad(False, "<-------- " & MyLoadResString(k_sqlAction_Break), 74) & " */"
            
                If flexTBotions(k_ColIdentityPreserve).ElementData.vVal Then Print #iFile, "SET IDENTITY_INSERT OFF"
                Print #iFile, ""
            
                If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT OFF"
                
                If lngSlitNew <> 0 Then
                    Print #iFile, ""
                    Print #iFile, "/* " & sPad(False, "", 74) & " */"
                    Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlWritingClosingFileNum), "1%", iFileNum), 74) & " */"
                    Print #iFile, "/* " & sPad(False, "", 74) & " */"
                End If
                
                Print #iFile, ""
                Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlAction_End), 74) & " */"
                Print #iFile, "/* " & sPad(False, "Sql Insert Script " & Replace(MyLoadResString(k_HtmProducedBy), "1%", App.EXEName & "  " & App.Major & "." & App.Minor & "." & App.Revision), 74) & " */"
                Print #iFile, ""
            End If
        End If
        StBar.Panels(1).Text = MyLoadResString(k_HtmReady) & Format(Now() - tInit, "hh:mm:ss")

    End If
    
    Close

    If Not adoRS Is Nothing Then
        If adoRS.State = adStateOpen Then adoRS.Close
    End If
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then adoConn.Close
    End If
    Set adoRS = Nothing
    Set adoConn = Nothing
    
    Flex(0).Enabled = True
    Flex(1).Enabled = True
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
    tBar.Buttons(keyStop).Enabled = False
    tBar.Buttons(keyEnd).Enabled = True
    tBar.Buttons(keyRun).Enabled = True
    tBar.Buttons(keyFitGrid).Enabled = True
    tBar.Buttons(keyRefresh).Enabled = True
    tBar.Buttons(keyExplorer).Enabled = ButtonEnabled()
    bWork = False

End Sub

Private Function GetSqlCommand(ByRef sSqlInsert As String) As String

    Dim i As Integer
    Dim sSql As String
    Dim sBuf As String
        
    'sSql = ""
    'sSqlInsert = ""
    For i = 0 To UBound(flexColumns)
        If flexColumns(i).bInclude Then
            sBuf = ""
            If Len(sSqlInsert) <> 0 Then sSqlInsert = sSqlInsert & ", "
            sSqlInsert = sSqlInsert & IIf(Len(flexColumns(i).sAlias) <> 0, flexColumns(i).sAlias, "[" & flexColumns(i).sName & "]")
            
            If Len(sSql) <> 0 Then sSql = sSql & ", "
            sSql = sSql & "CASE WHEN ( [" & flexColumns(i).sName & "] IS NOT NULL) THEN "
            Select Case LCase$(flexColumns(i).sType)
                Case "char", "varchar", "nchar", "nvarchar"
                    sSql = sSql & "'''' + REPLACE(RTRIM( [" & flexColumns(i).sName & "] ),'''','''''') +'''' ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & "'" & flexColumns(i).sReplaceNull & "'"
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
                                    
                Case "datetime", "smalldatetime" '121
                    sSql = sSql & "'''' + RTRIM(CONVERT(char, [" & flexColumns(i).sName & "] ,121)) + '''' ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & "'" & flexColumns(i).sReplaceNull & "'"
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
                Case "uniqueidentifier"
                    sSql = sSql & "'''' + REPLACE(CONVERT(char(255),RTRIM( [" & flexColumns(i).sName & "] ),'''','''''')) +'''' ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & "'" & flexColumns(i).sReplaceNull & "'"
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'
'                    'sSql = sSql & "'NULL'"
'                    sSql = sSql & " END"
                    
                Case "text", "ntext"
                    sSql = sSql & "'''' + REPLACE(CONVERT(char(8000), [" & flexColumns(i).sName & "] ),'''','''''') +'''' ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & "'" & flexColumns(i).sReplaceNull & "'"
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
                
                Case "binary", "varbinary"
                    sSql = sSql & "RTRIM(CONVERT(char ,CONVERT(int, [" & flexColumns(i).sName & "] ))) ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & "'" & flexColumns(i).sReplaceNull & "'"
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
                Case "timestamp", "rowversion"
                    sSql = sSql & "RTRIM(CONVERT(char(1),CONVERT(int, [" & flexColumns(i).sName & "] ))) ELSE "
'                    sSql = sSql & "'NULL'"
'                    sSql = sSql & " END"
                    
                Case "float", "real", "money", "smallmoney"
                    sSql = sSql & "LTRIM(RTRIM(CONVERT(char ,CONVERT(int, [" & flexColumns(i).sName & "] , 2)))) ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & flexColumns(i).sReplaceNull
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
                
                Case Else
                    sSql = sSql & "LTRIM(RTRIM(CONVERT(char , [" & flexColumns(i).sName & "] ))) ELSE "
'                    If flexColumns(i).bReplaceNull Then
'                        sSql = sSql & flexColumns(i).sReplaceNull
'                    Else
'                        sSql = sSql & "'NULL'"
'                    End If
'                    sSql = sSql & " END"
            End Select
            
            If flexColumns(i).bReplaceNull Then
                If Len(flexColumns(i).sReplaceNull) <> 0 Then sBuf = "'" & flexColumns(i).sReplaceNull & "'"
            End If
            sSql = sSql & IIf(Len(sBuf) <> 0, sBuf, "'NULL'") & " END"
        
        End If
    Next
    
    If Len(sSql) <> 0 Then
        sSqlInsert = "INSERT INTO " & IIf(Len(flexTBotions(k_ColTbAlias).ElementData.vVal) <> 0, flexTBotions(k_ColTbAlias).ElementData.vVal, m_sFullTbName) & " (" & sSqlInsert & " )"
        sSql = "SELECT " & IIf(flexTBotions(k_ColTOP).ElementData.vVal <> 0, "TOP " & ValInt(flexTBotions(k_ColTOP).ElementData.vVal) & " ", "") & sSql
        sSql = sSql & " FROM " & m_sFullTbName
        If Len(flexTBotions(k_ColWhere).ElementData.vVal) <> 0 Then sSql = sSql & " WHERE " & flexTBotions(k_ColWhere).ElementData.vVal
        If Len(flexTBotions(k_ColOrderBy).ElementData.vVal) <> 0 Then sSql = sSql & " ORDER BY " & flexTBotions(k_ColOrderBy).ElementData.vVal
    End If
'Debug.Print sSql
    GetSqlCommand = sSql
    
End Function
Private Function ButtonEnabled() As Boolean

    Dim lErr As Long
    Dim bRet As Boolean
    
    On Local Error Resume Next
    bRet = Len(Dir$(flexTBotions(k_ColExpFile).ElementData.vVal))
    lErr = Err.Number
    On Local Error GoTo 0
    If lErr <> 0 Then bRet = False
    ButtonEnabled = bRet
    
End Function
Private Sub ShowHtm()

    Dim sTopic As String
    Dim sParams As Variant
    Dim sDirectory As Variant
    Dim l_strFile As String
    
    l_strFile = flexTBotions(k_ColExpFile).ElementData.vVal
    
    sTopic = "Open"
    sParams = 0&
    sDirectory = 0&
    Call RunShellExecute(sTopic, l_strFile, sParams, sDirectory, SW_SHOWNORMAL, False)

End Sub

Private Function ErrOpenNewFile(ByRef iFile, ByRef iFileNum, ByRef sErr) As Long

    Dim sBuf As String
    Dim iPos As Integer
    Dim lErr As Long
    
    If iFile <> 0 Then Close #iFile
    
    sBuf = flexTBotions(k_ColExpFile).ElementData.vVal
    If flexTBotions(k_ColSplitFile).ElementData.vVal <> 0 Then
        iFileNum = iFileNum + 1
        iPos = InStrRev(sBuf, ".")
        sBuf = Left$(sBuf, iPos - 1) & "_" & iFileNum & Mid$(sBuf, iPos)
    End If

    iFile = FreeFile
    On Local Error Resume Next
    Open sBuf For Output As iFile
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
        
    If lErr = 0 Then
        Print #iFile, "/* " & sPad(False, "Sql Insert Script " & Replace(MyLoadResString(k_HtmProducedBy), "1%", App.EXEName & "  " & App.Major & "." & App.Minor & "." & App.Revision), 74) & " */"
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
        Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlTimeGenerated), "1%", Now()), 74) & " */"
        Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlInputDB), "1%", m_sDataBase), 74) & " */"
        Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlInputTbl), "1%", m_sFullTbName), 74) & " */"
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlSettings), 74) & " */"
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlProp_Value), 74) & " */"
        If Len(flexTBotions(k_ColIsolationLevel).ElementData.vVal) <> 0 Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(MyLoadResString(k_sqlProp_TranIsolationLevel), flexTBotions(k_ColIsolationLevel).ElementData.vVal)), 74) & " */"
        If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array("SET NOCOUNT ON", flexTBotions(k_ColSetNoCount).ElementData.vVal)), 74) & " */"
        If flexTBotions(k_ColIdentityPreserve).ElementData.vVal Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(MyLoadResString(k_sqlProp_PreserveIdentity), flexTBotions(k_ColIdentityPreserve).ElementData.vVal)), 74) & " */"
        If flexTBotions(k_ColBatchSize).ElementData.vVal <> 0 Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(MyLoadResString(k_sqlProp_BatchSize), flexTBotions(k_ColBatchSize).ElementData.vVal)), 74) & " */"
        If Len(flexTBotions(k_ColTbAlias).ElementData.vVal) Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(MyLoadResString(k_sqlProp_DestinationTable), flexTBotions(k_ColTbAlias).ElementData.vVal)), 74) & " */"
        If Len(flexTBotions(k_ColDropTbl).ElementData.vVal) <> 0 Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(MyLoadResString(k_sqlProp_Preliminary), flexTBotions(k_ColDropTbl).ElementData.vVal)), 74) & " */"
        If Len(flexTBotions(k_ColSetDateFormat).ElementData.vVal) <> 0 Then Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array("SET DATE FORMAT", flexTBotions(k_ColSetDateFormat).ElementData.vVal)), 74) & " */"
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlWarning0), 74) & " */"
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlWarning1), 74) & " */"
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlWarning2), 74) & " */"
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
        
        If iFileNum <> 0 Then
            Print #iFile, "/* " & sPad(False, "", 74) & " */"
            Print #iFile, "/* " & sPad(False, Replace(MyLoadResString(k_sqlWritingFileNum), "1%", iFileNum), 74) & " */"
            Print #iFile, "/* " & sPad(False, "", 74) & " */"
        End If
        
        If iFileNum > 1 Then
            If Len(flexTBotions(k_ColIsolationLevel).ElementData.vVal) Then
                Print #iFile, "SET TRANSACTION ISOLATION LEVEL "; flexTBotions(k_ColIsolationLevel).ElementData.vVal
            End If
        
            If Len(flexTBotions(k_ColSetDateFormat).ElementData.vVal) <> 0 Then Print #iFile, "SET DATEFORMAT " & flexTBotions(k_ColSetDateFormat).ElementData.vVal
            
            Print #iFile, ""
            Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlAction_Loading), 74) & " */"
            Print #iFile, ""
            If flexTBotions(k_ColSetNoCount).ElementData.vVal Then Print #iFile, "SET NOCOUNT ON"
            Print #iFile, ""
                
            If flexTBotions(k_ColIdentityPreserve).ElementData.vVal Then Print #iFile, "SET IDENTITY_INSERT ON"
            Print #iFile, ""
        
        End If
    End If
    ErrOpenNewFile = lErr

End Function
Private Function ErrScriptTable(ByVal iFile As String, ByRef sErr) As Long

    Dim sBuf As String
    Dim lErr As Long
    Dim oTrans As SQLDMO.Transfer
    Dim oDb As SQLDMO.Database2
    Dim oTb As SQLDMO.Table
    Dim x As SQLDMO_SCRIPT_TYPE
    Dim y As SQLDMO_SCRIPT2_TYPE
        
    On Local Error Resume Next
    Set oDb = objServer.Databases(m_sDataBase)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then Set oTb = GetTbByName(objServer, m_sDataBase, m_sTbName)
    
    If Not oTb Is Nothing Then
        y = 0
        x = SQLDMOScript_Aliases Or SQLDMOScript_Default
        y = SQLDMOScript2_AnsiFile
        x = x Or SQLDMOScript_OwnerQualify Or SQLDMOScript_IncludeIfNotExists
        'X = X Or SQLDMOScript_IncludeHeaders
        If CBool(flexTBotions(k_ColScriptIDX).ElementData.vVal) Then x = x Or SQLDMOScript_Indexes Or SQLDMOScript_NonClusteredIndexes Or SQLDMOScript_ClusteredIndexes
        If CBool(flexTBotions(k_ColScriptFullTextIDX).ElementData.vVal) Then y = y Or SQLDMOScript2_FullTextCat Or SQLDMOScript2_FullTextIndex
        If CBool(flexTBotions(k_ColScriptTrigger).ElementData.vVal) Then x = x Or SQLDMOScript_Triggers
        If CBool(flexTBotions(k_ColScriptKeys).ElementData.vVal) Then x = x Or SQLDMOScript_DRI_UniqueKeys Or SQLDMOScript_DRI_Checks Or SQLDMOScript_DRI_Clustered Or SQLDMOScript_DRI_Defaults Or SQLDMOScript_DRI_NonClustered Or SQLDMOScript_DRI_PrimaryKey Or SQLDMOScript_DRI_UniqueKeys Or SQLDMOScript_DRIIndexes
        
        Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlTB_opt_ScriptGenerate), 74) & " */"
        Print #iFile, "/* " & sPad(False, String$(Len(MyLoadResString(k_sqlTB_opt_ScriptGenerate)) + 5, "-"), 74) & " */"
        
        Set oTrans = New SQLDMO.Transfer
        oTrans.AddObjectByName m_sFullTbName, SQLDMOObj_UserTable
        
        oTrans.CopyData = SQLDMOCopyData_False
        oTrans.ScriptType = x
        oTrans.Script2Type = y
        
        On Local Error Resume Next
        sBuf = oDb.ScriptTransfer(oTrans)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        'Debug.Print sBuf

        If lErr = 0 Then
            Print #iFile, "/* " & sPad(False, "", 74) & " */"
            Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlTB_opt_ScriptGenerateWarning0), 74) & " */"
            Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlTB_opt_ScriptGenerateWarning1), 74) & " */"
            Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlTB_opt_ScriptGenerateWarning2), 74) & " */"
            Print #iFile, "/* " & sPad(False, "", 74) & " */"
            
            Print #iFile, "/* " & sPad(False, MyLoadResString(k_sqlSettings), 74) & " */"
            Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(flexTBotions(k_ColScriptIDX).ElementData.sField, CBool(flexTBotions(k_ColScriptIDX).ElementData.vVal))), 74) & " */"
            Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(flexTBotions(k_ColScriptFullTextIDX).ElementData.sField, CBool(flexTBotions(k_ColScriptFullTextIDX).ElementData.vVal))), 74) & " */"
            Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(flexTBotions(k_ColScriptTrigger).ElementData.sField, CBool(flexTBotions(k_ColScriptTrigger).ElementData.vVal))), 74) & " */"
            Print #iFile, "/* " & sPad(False, ReplaceMsg(lk_sqlProp_Value_VALUE, Array("1%", "2%"), Array(flexTBotions(k_ColScriptKeys).ElementData.sField, CBool(flexTBotions(k_ColScriptKeys).ElementData.vVal))), 74) & " */"
            Print #iFile, "/* " & sPad(False, "", 74) & " */"
            
            
'            '**** BAD IDEA
'            'table alias
'            If Len(flexTBotions(k_ColTbAlias).ElementData.vVal) <> 0 Then
'                sBuf = Replace(sBuf, m_sTbName, flexTBotions(k_ColTbAlias).ElementData.vVal)
'            End If
'
'            'column alias
'            For i = 0 To UBound(flexColumns)
'                If Len(flexColumns(i).sAlias) <> 0 Then
'                    sBuf = Replace(sBuf, flexColumns(i).sName, flexColumns(i).sAlias)
'                End If
'            Next
            
            'Debug.Print sBuf
            'Debug.Print Replace("if not exists (select * from sysobjects where id = object_id(N'1%') and OBJECTPROPERTY(id, N'IsUserTable') = 1)", "1%", IIf(Len(flexTBotions(k_ColTbAlias).ElementData.vVal) <> 0, flexTBotions(k_ColTbAlias).ElementData.vVal, m_sFullTbName))
            'Print #iFile, Replace("if not exists (select * from sysobjects where id = object_id(N'1%') and OBJECTPROPERTY(id, N'IsUserTable') = 1)", "1%", IIf(Len(flexTBotions(k_ColTbAlias).ElementData.vVal) <> 0, flexTBotions(k_ColTbAlias).ElementData.vVal, m_sFullTbName))
            'Print #iFile, "BEGIN"
            'if not exists (select * from sysobjects where id = object_id(N'[dbo].[CO_AGENZIE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)


            Print #iFile, sBuf
            'Print #iFile, "END"
        Else
            Print #iFile, "/* " & sPad(False, MyLoadResString(kMsgBoxError) & ": " & lErr, 74) & " */"
            Print #iFile, "/* " & sPad(False, "@: " & sErr, 74) & " */"
        End If
        Print #iFile, ""
        Print #iFile, "/* " & sPad(False, "", 74) & " */"
    End If
    
    Set oTrans = Nothing
    Set oTb = Nothing
    Set oDb = Nothing
    ErrScriptTable = lErr

End Function
