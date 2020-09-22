VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fExpHTM 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6060
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   180
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
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
      Left            =   4320
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Tbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6060
      _ExtentX        =   10689
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Explorer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Stop"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1470
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7594
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/11/2002"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1905
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSComctlLib.ImageList imgLsit 
      Left            =   4800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpHTM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpHTM.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpHTM.frx":05B0
            Key             =   "XML"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpHTM.frx":070A
            Key             =   "FitGrid"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fExpHTM.frx":0864
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fExpHTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const kGetProperties As String = "SELECT [name] , convert(VARCHAR(8000),[value]) AS vCharValue, convert(VARCHAR(256),SQL_VARIANT_PROPERTY([value],'basetype')) as [VariantType]  FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
Private m_blnStopExport As Boolean

Private m_blnAnnullaMod As Boolean
Private Type udtExport
    ElemData As elemFlexG
    EditType As TypeEnum
End Type
Private m_sProduced As String
Private m_sDbName As String
Private m_Width As Long
Private m_Height As Long
Private bWork As Boolean
Private bLoading As Boolean
Private mType() As udtExport
Private miRow  As Integer
Private miTmpType As iPropG

Private Const kDefPage As String = "\Default.htm"
Private Const kCloseTB As String = "<TBODY></TBODY></TABLE></P>"
Private Const kRefMain As String = "<BR><A href='1@%\Default.htm'> 2@%</A><BR>"      'Pagina Iniziale
Private Const kObjPage As String = "\pg_1%.htm"

Private m_sLngRes() As String
Private Enum am_LangExport
    enScanHOME
    enScanName
    enScanning
    enScan_Object
    enScan_Type
    enScanOwner
    enScanOwnership
    enScanCreateDate
    enScanFileGroup
    enScanRows
    enScanSysObj
    enScanTb
    enScanStoredProc
    enScanViews
    enScanUDT
    enScanRULE
    enScanDEF
    enScan_UdFunc
    enScanUsers
    enScanLogFile
    enScanDatabase
    enScanErr_OnlyOwner
    enScanLastBackup
    enScanAvailableSpace
    enScanFileName
    enScanPhysicalName
    enScanFileGrowth
    enScanMaxSize
    enScanFileSizeKb
    enScanFileGrowthType
    enScanCompatLevel
    enScanCollation
    enScanExtendedProperties
    enScanExtendedPropertiesBaseType
    'tb
    enScanTbDef
    enScanTb_inPK
    enScanTb_FieldName
    enScanTb_Data_Type
    enScanTb_Size
    enScanTb_AllowNull
    enScanTb_DefVal
    enScanTb_Precision
    enScanTb_Scale
    enScanTb_IsRowGuid
    enScanTb_IsIdentity
    enScanTb_InitialVal
    enScanTb_Increment
    enScanTb_Computed
    enScanTb_Index_Key_Constr
    enScanTb_FK
    enScanTb_FKtype
    enScanTb_FKclustered
    enScanTb_FKfillfactor
    enScanTb_FKcheck
    enScanTb_FKcolumns
    enScanTb_Htm4TB
    enScanTb_Checkconstr
    enScanTb_CheckText
    enScanScript
    enScan_Index
    'Col
    enScanColDef
    enScanColComputedText
    enScanColList
    enScanColPriTB
    enScanCol_PK
    enScanCol_UK
    
    enScan_Table
    enScan_View
    enScan_StoredP
    enScan_UDT
    enScan_Parameter
    
    enScan_DedendON
    enScan_DedendFROM
    
    enScan_ViewColumn
    enScan_Lenght
    enScan_BoundCols
    enScan_Text
    enScan_BoundUDT
    
    enScanProperty
    enScanValue
End Enum

Public Property Let DBName(ByVal sDbName As String)

    
    m_sDbName = sDbName
    
    Me.Caption = ReplaceMsg(MyLoadResString(k_HtmDocumInfo), Array("1%"), Array(sDbName))
    ReDim mType(2) As udtExport
    
    mType(0).ElemData.sField = MyLoadResString(k_DbName)
    mType(0).ElemData.vVal = sDbName
    mType(0).ElemData.iType = am_String
    mType(0).EditType = iNoEdit
    
    mType(1).ElemData.sField = MyLoadResString(k_HtmLocation)
    mType(1).ElemData.vVal = GetSpecialFolderLocation(Me.hWnd, CSIDL_PERSONAL) & "\" & sDbName
    mType(1).ElemData.iType = am_String
    mType(1).ElemData.iLen = 255
    mType(1).EditType = iEllipsis
    
    mType(2).ElemData.sField = MyLoadResString(k_HtmIncludeSYSobj)
    mType(2).ElemData.vVal = False
    mType(2).ElemData.iType = am_Bool
    mType(2).EditType = iNormalEdit
    
    tBar.Buttons(keyExplorer).Enabled = IsShowEnabled(mType(1).ElemData.vVal)
    LoadFlex
    

End Property

Private Sub Flex_DblClick()
    
    Dim bEllipsis As Boolean
    
    With Flex
        If Flex.Col = Flex.Cols - 1 Then 'And Flex.MouseRow > 0 And Flex.MouseCol > 1 Then
            bEllipsis = mType(miRow - 1).EditType = iEllipsis
    
            If bEllipsis Then
                LocationIsCalled
            Else
                FlexEdit 0
            End If
        
        End If
    End With

End Sub

Private Sub Flex_GotFocus()

    TestActiveControl
    tFlex.Visible = False
    m_blnAnnullaMod = True
    
End Sub

Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    iCode = KeyCode
    If iCode = vbKeyF2 Then
        KeyCode = 0
        FlexEdit 0
    End If

End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyEscape, vbKeyBACK
        Case vbKeySpace
            If mType(miRow - 1).EditType = iEllipsis Then
                LocationIsCalled
            Else
                FlexEdit KeyAscii
            End If
            
        Case Else
            FlexEdit KeyAscii
    End Select
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
                    StBar.Panels(1).Text = MyLoadResString(k_Property) & ": " & .TextMatrix(miRow, 1)
                
                End If
                
            End With
            b = False
        End If
    End If

End Sub

Private Sub Flex_Scroll()
    tFlex.Visible = False
End Sub

Private Sub Form_Load()
    
    m_Width = Me.Width
    m_Height = Me.Height
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keyRun).ToolTipText = MyLoadResString(k_Run)
    tBar.Buttons(keyExplorer).ToolTipText = MyLoadResString(k_HtmShow)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_RES_FitGrid)
    tBar.Buttons(keyStop).ToolTipText = MyLoadResString(k_HtmKillExport)
        
    CenterForm Me

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = bWork
End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Me.WindowState <> vbMinimized Then
        If Not b Then
            b = True
            If Me.Width < m_Width Then Me.Width = m_Width
            If Me.Height <> m_Height Then Me.Height = m_Height
            Flex.Width = Me.ScaleWidth - 2 * (Flex.Left)
            pBar.Width = Me.ScaleWidth - pBar.Left - (k_FlexBevel \ 2)
            SettaLargFlex Flex, 2, Flex.ColWidth(1)
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set fExpHTM = Nothing
    
End Sub
Private Sub LoadFlex()

    Dim i As Integer
    Dim sBuf As String
    
    With Flex
        bLoading = True
        .Redraw = False
        .Clear
        .FixedCols = 0
        .Cols = 3
        .Rows = 1
        .Clear
                
        .TextMatrix(0, 1) = MyLoadResString(k_Property)
        .TextMatrix(0, 2) = MyLoadResString(k_Value)
        
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        .ColWidth(1) = .Width / 3
        .ColWidth(2) = .Width / 2
                
        For i = 0 To UBound(mType)
            sBuf = vbTab & mType(i).ElemData.sField
            
            Select Case mType(i).ElemData.iType
                Case am_String
                    sBuf = sBuf & vbTab & mType(i).ElemData.vVal
                Case am_Bool
            End Select
            Flex.AddItem sBuf
            
            
            .Row = .Rows - 1
            .Col = .Cols - 1
            If mType(i).EditType = iEllipsis Then
                Set Flex.CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
                .CellPictureAlignment = flexAlignRightBottom
            ElseIf mType(i).EditType = iNoEdit Then
                .CellBackColor = vbInactiveBorder
            ElseIf mType(i).ElemData.iType = am_Bool Then
                Set Flex.CellPicture = LoadResPicture(IIf(Abs(mType(i).ElemData.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                Flex.CellPictureAlignment = flexAlignCenterCenter
            End If
        Next
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .FixedRows = 1
        .FixedCols = 2
        
        .Redraw = True
        
        SettaLargFlex Flex, 2, .ColWidth(1)
        miRow = 0
        .TopRow = 1
        .Row = 1
        miRow = 1
        bLoading = False
        Flex_RowColChange
    End With
End Sub
Private Function GetLocation() As String

    Dim a As BROWSEINFO
    Dim sBuf As String
    
    Dim pidlBrowse As Long
    
    a.lpszTitle = MyLoadResString(k_Select_Full_Data_Path)   '"Seleziona Percorso Dati"
        
    a.hOwner = Me.hWnd
    pidlBrowse = SHBrowseForFolder(a)
    sBuf = GetPathFromPIDL(pidlBrowse)
    GetLocation = sBuf

End Function
Private Sub LocationIsCalled()

    Dim sNewLoc As String
    
    sNewLoc = GetLocation() 'GetSpecialFolderLocation(Me.hWnd, CSIDL_PERSONAL)
    If Len(sNewLoc) <> 0 Then
        If Right$(sNewLoc, 1) <> "\" Then sNewLoc = sNewLoc & "\"
        mType(1).ElemData.vVal = sNewLoc & mType(0).ElemData.vVal
    Else
        mType(1).ElemData.vVal = GetSpecialFolderLocation(Me.hWnd, CSIDL_PERSONAL) & "\" & m_sDbName
    End If
    Flex.TextMatrix(2, 2) = mType(1).ElemData.vVal
    tBar.Buttons(keyExplorer).Enabled = IsShowEnabled(mType(1).ElemData.vVal)
    
End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iImg As Integer
    
    miTmpType = am_String
    iRow = Flex.Row
    iCol = Flex.Col
    
    bEdit = True

    If Flex.CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        miTmpType = mType(iRow - 1).ElemData.iType
        sBuf = mType(iRow - 1).ElemData.vVal
        iLen = mType(iRow - 1).ElemData.iLen
        
        miRow = iRow
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0, vbKeySpace
                If (miTmpType And am_Bool) = am_Bool Then
                    iImg = 1 - (-mType(iRow - 1).ElemData.vVal)
                    mType(iRow - 1).ElemData.vVal = CBool(iImg)
                    Set Flex.CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    bEdit = False
                End If
            Case Else
                If (miTmpType And am_Bool) = am_Bool Then bEdit = False
                
        End Select
    End If

    If bEdit Then
        tFlex.PasswordChar = ""
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf
        tFlex.SelStart = 0
        With Flex
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If (miTmpType And am_String) = am_String Then If KeyAscii = vbKeySpace Then KeyAscii = 0
        If KeyAscii <> 0 Then
            On Local Error Resume Next
            SendKeys Chr$(KeyAscii)
            On Local Error GoTo 0
        End If
    End If

End Sub

Private Sub TestActiveControl()

    If tFlex.Visible Then
        tFlex_LostFocus
    End If

   On Local Error Resume Next
   If TypeOf Me.ActiveControl Is TextBox Then
      If Me.ActiveControl Is tFlex Then
         tFlex_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case keyEnd
            If Not bWork Then Unload Me
        Case keyFitGrid
            If Not bWork Then LoadFlex
        Case keyRun
            If Not bWork Then RunExport
        Case keyExplorer
            If Not bWork Then ShowHtm
        Case keyStop
            If bWork Then
                m_blnStopExport = True
            End If
    End Select
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
        'Case 60, 64, 124, 91,92, 93    '<>|[\]
        '    KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex.Text = Flex.Text
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex.SetFocus
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
        Case vbKeySpace
            If miRow <> 2 Then
                KeyAscii = 0
            Else
                OverStrike tFlex
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
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()

    Dim bUpdate As Boolean
    Dim sBuf As String
    
    If Not m_blnAnnullaMod Then
        tFlex = Trim$(tFlex.Text)
        sBuf = tFlex.Text
        bUpdate = True
    End If
    If bUpdate Then
        If miRow = 2 Then
            sBuf = Test4Directory(sBuf)
            mType(miRow - 1).ElemData.vVal = sBuf
            tBar.Buttons(keyExplorer).Enabled = IsShowEnabled(mType(1).ElemData.vVal)
        End If
        Flex.TextMatrix(miRow, 2) = sBuf
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    
    m_blnAnnullaMod = True
    tFlex.Visible = False
    Flex.SetFocus

End Sub

Private Function Test4Directory(ByVal sDir As String) As String

    Dim bOk As Boolean
    sDir = Trim$(sDir)
        
    If Len(sDir) Then
        If Right$(sDir, 1) = "\" Then sDir = Left$(sDir, Len(sDir) - 1)
        On Local Error Resume Next
        bOk = (GetAttr(sDir) And vbDirectory) = vbDirectory
        If Err.Number Then bOk = False
        On Local Error GoTo 0
        If bOk Then Test4Directory = sDir
    End If
    
End Function
Private Function IsShowEnabled(ByVal sOutDir As String) As Boolean

    Dim bExist As Boolean

    On Local Error Resume Next
    bExist = Len(Dir$(sOutDir & kDefPage)) <> 0
    If Err <> 0 Then bExist = False
    On Local Error GoTo 0
    IsShowEnabled = bExist

End Function

Private Sub RunExport()

    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim sDir As String
    Dim bExist As Boolean
    Dim bIsDir As Boolean
    Dim iC As Integer
    Dim sBuf As String
    Dim bDone As Boolean
    Dim oColl As Collection
    Dim tInit As Date
        
    
    If bWork Then Exit Sub
    bWork = True
    tBar.Buttons(keyExplorer).Enabled = False
    tBar.Buttons(keyRun).Enabled = False
    tBar.Buttons(keyFitGrid).Enabled = False
    Flex.Enabled = False
    
    sDir = mType(1).ElemData.vVal
    If Len(sDir) <> 0 Then
        m_sProduced = ReplaceMsg("<P><A href='1%'>2%</A> -3%<P>", Array("1%", "2%", "3%"), Array(k_URLwww, Replace(MyLoadResString(k_HtmProducedBy), "1%", App.EXEName & "  " & App.Major & "." & App.Minor & "." & App.Revision), Now))
    
        On Local Error Resume Next
        bExist = Len(Dir$(sDir)) <> 0
        If Err <> 0 Then bExist = False
        bIsDir = (GetAttr(sDir) And vbDirectory) = vbDirectory
        If Err <> 0 Then bIsDir = False
        On Local Error GoTo 0
        
        If bIsDir Then
            iC = MsgBox(MyLoadResString(k_HtmClearDir), vbQuestion Or vbYesNoCancel, App.EXEName)
            If iC = vbCancel Then
                tBar.Buttons(keyExplorer).Enabled = IsShowEnabled(mType(1).ElemData.vVal)
                tBar.Buttons(keyRun).Enabled = True
                tBar.Buttons(keyFitGrid).Enabled = True
                Flex.Enabled = True
                bWork = False
                Exit Sub
            End If
            
            If iC = vbYes Then
                StBar.Panels(1).Text = MyLoadResString(k_HtmDeleting)
                KillDir sDir, lErr, sErr
            End If
            
        Else
            If Not bExist Then
                On Local Error Resume Next
                MkDir sDir
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
            End If
        End If
        
        If lErr = 0 Then
            LoadAll_RES_string
            
            tInit = Now()
            Screen.MousePointer = vbHourglass
            
            On Local Error Resume Next
            Set oDb = objServer.Databases(m_sDbName)
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            If lErr = 0 Then
                StBar.Panels(1).Text = MyLoadResString(k_HtmWorking)
                ShowProgress True
                m_blnStopExport = False
                
                StBar.Panels(1).Text = Replace(MyLoadResString(k_HtmScanning), "1%", MyLoadResString(k_HtmScanDB))
                
                lErr = ErrWriteDbStuff(oDb, sDir, sErr)
                
'-------Tabelle
                sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=20% bgcolor=#ffffcc> 2@% </TD><TD width=20% bgcolor=#ffffcc> 3@% </TD><TD width=30% bgcolor=#ffffcc> 4@% </TD><TD width=5% bgcolor=#ffffcc align=right> 5@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                bDone = False
                Set oColl = New Collection
                
                StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.Tables.Count & " # " & m_sLngRes(enScanTb))
                
                For iC = 1 To oDb.Tables.Count
                    IncrementProgres
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.Tables.Count & " # " & m_sLngRes(enScanTb))
                    If Not oDb.Tables(iC).SystemObject Then
                        If Not bDone Then
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(m_sLngRes(enScanTb), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate), m_sLngRes(enScanFileGroup), m_sLngRes(enScanRows))))
                            sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD><TD>5@%</TD><TD align=right>6@%</TD></TR>"
                            bDone = True
                        End If
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), Array(Replace(oDb.Tables(iC).Owner & "_" & oDb.Tables(iC).Name, " ", "_"), oDb.Tables(iC).Name, oDb.Tables(iC).Owner, oDb.Tables(iC).CreateDate, oDb.Tables(iC).FileGroup, oDb.Tables(iC).Rows)))
                        lErr = ErrWriteTB(oDb, oDb.Tables(iC), sDir, sErr)
                    Else
                        If mType(2).ElemData.vVal Then oColl.Add oDb.Tables(iC)
                    End If
                    
                    DoEvents
                    If m_blnStopExport Then Exit For
                Next
                If Not m_blnStopExport Then
                    If oColl.Count Then
                        If Not bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(m_sLngRes(enScanTb), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate), m_sLngRes(enScanFileGroup), m_sLngRes(enScanRows))))
                        bDone = True
                        
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, Replace("<TR><TD><B><I>1@%</I></B></TD></TR>", "1@%", Replace(m_sLngRes(enScanSysObj), "1%", m_sLngRes(enScanTb))))
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD><TD>5@%</TD><TD align=right>6@%</TD></TR>"
                        For iC = 1 To oColl.Count
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), Array(Replace(oColl.Item(iC).Owner & "_" & oColl.Item(iC).Name, " ", "_"), oColl.Item(iC).Name, oColl.Item(iC).Owner, oColl.Item(iC).CreateDate, oColl.Item(iC).FileGroup, oColl.Item(iC).Rows)))
                            IncrementProgres
                            lErr = ErrWriteTB(oDb, oColl.Item(iC), sDir, sErr)
                            
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                    End If
                    If bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                End If
                oDb.Tables.Refresh True
                
'-------Stored Procedures
                If Not m_blnStopExport Then
                    sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                    sBuf = sBuf & "<TD width=45% bgcolor=#ffffcc> 1@% </TD><TD width=30% bgcolor=#ffffcc> 2@% </TD><TD width=30% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                
                    bDone = False
                    Set oColl = New Collection
                    
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.StoredProcedures.Count & " # " & m_sLngRes(enScanStoredProc))
                    For iC = 1 To oDb.StoredProcedures.Count
                        IncrementProgres
                        StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.StoredProcedures.Count & " # " & m_sLngRes(enScanStoredProc))
                        If Not oDb.StoredProcedures(iC).SystemObject Then
                            If Not bDone Then
                                lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanStoredProc), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                                sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                                bDone = True
                            End If
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oDb.StoredProcedures(iC).Owner & "_" & oDb.StoredProcedures(iC).Name, " ", "_"), oDb.StoredProcedures(iC).Name, oDb.StoredProcedures(iC).Owner, oDb.StoredProcedures(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteSP(oDb, oDb.StoredProcedures(iC), sDir, sErr)
                        Else
                            If mType(2).ElemData.vVal Then
                                oColl.Add oDb.StoredProcedures(iC)
                            End If
                        End If
                        
                        DoEvents
                        If m_blnStopExport Then Exit For
                    Next
                    
                    If Not m_blnStopExport Then
                        If oColl.Count Then
                            If Not bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanStoredProc), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                            bDone = True
                    
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, Replace("<TR><TD><B><I>1@%</I></B></TD></TR>", "1@%", Replace(m_sLngRes(enScanSysObj), "1%", m_sLngRes(enScanStoredProc))))
                            sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                                                        
                            For iC = 1 To oColl.Count
                                lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oColl.Item(iC).Owner & "_" & oColl.Item(iC).Name, " ", "_"), oColl.Item(iC).Name, oColl.Item(iC).Owner, oColl.Item(iC).CreateDate)))
                                IncrementProgres
                                lErr = ErrWriteSP(oDb, oColl.Item(iC), sDir, sErr)
                                
                                DoEvents
                                If m_blnStopExport Then Exit For
                            Next
                        End If
                        If bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                    End If
                End If
                oDb.StoredProcedures.Refresh True
                
'-------Viste
                If Not m_blnStopExport Then
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.Views.Count & " # " & m_sLngRes(enScanViews))
                    sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                    sBuf = sBuf & "<TD width=45% bgcolor=#ffffcc> 1@% </TD><TD width=30% bgcolor=#ffffcc> 2@% </TD><TD width=30% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                
                    bDone = False
                    Set oColl = New Collection
                
                    For iC = 1 To oDb.Views.Count
                        IncrementProgres
                        StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.Views.Count & " # " & m_sLngRes(enScanViews))
                        If Not oDb.Views(iC).SystemObject Then
                            If Not bDone Then
                                lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanViews), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                                sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                                bDone = True
                            End If
                            
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oDb.Views(iC).Owner & "_" & oDb.Views(iC).Name, " ", "_"), oDb.Views(iC).Name, oDb.Views(iC).Owner, oDb.Views(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteView(oDb, oDb.Views(iC), sDir, sErr)
                        Else
                            If mType(2).ElemData.vVal Then oColl.Add oDb.Views(iC)
                        End If
                        
                        DoEvents
                        If m_blnStopExport Then Exit For
                    Next
                End If
                If Not m_blnStopExport Then
                    If oColl.Count Then
                        If Not bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanViews), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                        bDone = True
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, Replace("<TR><TD><B><I>1@%</I></B></TD></TR>", "1@%", Replace(m_sLngRes(enScanSysObj), "1%", m_sLngRes(enScanViews))))
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                        
                        For iC = 1 To oColl.Count
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oColl.Item(iC).Owner & "_" & oColl.Item(iC).Name, " ", "_"), oColl.Item(iC).Name, oColl.Item(iC).Owner, oColl.Item(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteView(oDb, oColl.Item(iC), sDir, sErr)
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                    End If
                
                    If bDone Then lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                End If
                bDone = False
                Set oColl = New Collection
                oDb.Views.Refresh True

'-------UDT
                If Not m_blnStopExport Then
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.UserDefinedDatatypes.Count & " # " & m_sLngRes(enScanUDT))
                    If oDb.UserDefinedDatatypes.Count Then
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=80% bgcolor=#ffffcc> 1@% </TD><TD width=20% bgcolor=#ffffcc> 2@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanUDT), m_sLngRes(enScanOwner))))
                                    
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD></TR>"
                
                        For iC = 1 To oDb.UserDefinedDatatypes.Count
                            IncrementProgres
                            
                            StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.UserDefinedDatatypes.Count & " # " & m_sLngRes(enScanUDT))
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(Replace(oDb.UserDefinedDatatypes(iC).Owner & "_" & oDb.UserDefinedDatatypes(iC).Name, " ", "_"), oDb.UserDefinedDatatypes(iC).Name, oDb.UserDefinedDatatypes(iC).Owner)))
                            IncrementProgres
                            lErr = ErrWriteUDT(oDb.UserDefinedDatatypes(iC), sDir, sErr)
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                        
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                    End If
                End If
                oDb.UserDefinedDatatypes.Refresh True

'-------Rules
                If Not m_blnStopExport Then
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.Rules.Count & " # " & m_sLngRes(enScanRULE))
                    If oDb.Rules.Count Then
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=60% bgcolor=#ffffcc> 1@% </TD><TD width=20% bgcolor=#ffffcc> 2@% </TD><TD width=20% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanRULE), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                                    
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                        For iC = 1 To oDb.Rules.Count
                            IncrementProgres
                            StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.Rules.Count & " # " & m_sLngRes(enScanRULE))
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oDb.Rules(iC).Owner & "_" & oDb.Rules(iC).Name, " ", "_"), oDb.Rules(iC).Name, oDb.Rules(iC).Owner, oDb.Rules(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteRULE(oDb.Rules(iC), sDir, sErr)
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                    End If
                End If
                oDb.Rules.Refresh True
                
'-------Default
                If Not m_blnStopExport Then
                
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.Defaults.Count & " # " & m_sLngRes(enScanDEF))
                    If oDb.Defaults.Count Then
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=60% bgcolor=#ffffcc> 1@% </TD><TD width=20% bgcolor=#ffffcc> 2@% </TD><TD width=20% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanDEF), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                                    
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                        For iC = 1 To oDb.Defaults.Count
                            IncrementProgres
                            StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.Defaults.Count & " # " & m_sLngRes(enScanDEF))
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oDb.Defaults(iC).Owner & "_" & oDb.Defaults(iC).Name, " ", "_"), oDb.Defaults(iC).Name, oDb.Defaults(iC).Owner, oDb.Defaults(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteDEFAULT(oDb.Defaults(iC), sDir, sErr)
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                    End If
                End If
    
                sBuf = "<TBODY></TBODY></TABLE></P>"
                lErr = ErrWriteDbStuff(Nothing, sDir, sErr, sBuf)
                
                lErr = ErrWriteDbStuff(Nothing, sDir, sErr, "", True)

                ShowProgress False
            End If
            oDb.Defaults.Refresh True
            
            
'-------User Function
                If Not m_blnStopExport Then
                    StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & oDb.UserDefinedFunctions.Count & " # " & m_sLngRes(enScan_UdFunc))
                    If oDb.UserDefinedFunctions.Count Then
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=60% bgcolor=#ffffcc> 1@% </TD><TD width=20% bgcolor=#ffffcc> 2@% </TD><TD width=20% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScan_UdFunc), m_sLngRes(enScanOwner), m_sLngRes(enScanCreateDate))))
                                    
                        sBuf = "<TR><TD><A href='pg_1@%.htm'>2@%</A></TD><TD>3@%</TD><TD>4@%</TD></TR>"
                        For iC = 1 To oDb.UserDefinedFunctions.Count
                            IncrementProgres
                            StBar.Panels(1).Text = Replace(m_sLngRes(enScanning), "1%", " # " & iC & " / " & oDb.UserDefinedFunctions.Count & " # " & m_sLngRes(enScan_UdFunc))
                            lErr = ErrWriteDbStuff(Nothing, sDir, sErr, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(Replace(oDb.UserDefinedFunctions(iC).Owner & "_" & oDb.UserDefinedFunctions(iC).Name, " ", "_"), oDb.UserDefinedFunctions(iC).Name, oDb.UserDefinedFunctions(iC).Owner, oDb.UserDefinedFunctions(iC).CreateDate)))
                            IncrementProgres
                            lErr = ErrWriteUserFunc(oDb, oDb.UserDefinedFunctions(iC), sDir, sErr)
                            
                            DoEvents
                            If m_blnStopExport Then Exit For
                        Next
                        lErr = ErrWriteDbStuff(Nothing, sDir, sErr, kCloseTB)
                    End If
                End If
                oDb.UserDefinedFunctions.Refresh True
            
                
                StBar.Panels(1).Text = MyLoadResString(k_HtmReady) & Format(Now() - tInit, "hh:mm:ss")
        
                Screen.MousePointer = vbDefault
        End If
        
    End If
    
    If lErr <> 0 Then MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    
    Set oColl = Nothing
    ReDim m_sLngRes(0) As String
    
    Close   'close all
    
    If m_blnStopExport Then
        KillDir sDir, lErr, sErr
        If lErr = 0 Then KillDir sDir, 0, "", True
        StBar.Panels(1).Text = MyLoadResString(k_HtmKilled)
    End If
        
    tBar.Buttons(keyExplorer).Enabled = IsShowEnabled(mType(1).ElemData.vVal)
    tBar.Buttons(keyRun).Enabled = True
    tBar.Buttons(keyFitGrid).Enabled = True
    Flex.Enabled = True
    
    m_blnStopExport = False
    bWork = False
    
End Sub
Private Sub ShowProgress(ByVal bVisible As Boolean)

    If bVisible Then
        pBar.Value = 0
        tBar.Buttons(keyStop).Enabled = True
        'show stop
    Else
        tBar.Buttons(keyStop).Enabled = False
    End If
    
    pBar.Visible = bVisible
    
End Sub
Private Sub IncrementProgres()

    If pBar.Value < pBar.Max Then
        pBar.Value = pBar.Value + 1
    Else
        pBar.Value = 1
    End If
    DoEvents
    
End Sub
Private Function ErrWriteDbStuff(ByVal oDb As SQLDMO.Database2, ByVal sDir As String, ByRef sErr As String, Optional ByVal szBuf As String = "", Optional ByVal bClose As Boolean) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sType As String
    Dim bOwner As Boolean
    Dim i As Integer, l As Integer
    Dim vVarProp As Variant
    Dim sSql As String
    
    iOut = FreeFile
    On Local Error Resume Next
    If Len(Trim$(szBuf)) = 0 And Not bClose Then
        Open sDir & kDefPage For Output As iOut
    Else
        Open sDir & kDefPage For Append As iOut
    End If
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        If bClose Then
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
        ElseIf Len(Trim$(szBuf)) = 0 Then
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% [2@%]</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScanDatabase), m_sDbName))
            
            Print #iOut, m_sProduced

            
            If Not oDb Is Nothing Then
                
                With oDb
                    sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                    sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD><TD width=50% align=right bgcolor=#ffffcc> 2@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                
                    'ObjectName
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
                     
                    sBuf = "<TR><TD>1@%</TD><TD align=right>2@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanDatabase), .Name))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanOwnership), .Owner))
                    
                                
                    Select Case .Status
                        Case SQLDMODBStat_Normal
                            sType = "Normal"
                        Case SQLDMODBStat_Offline
                            sType = "Off Line"
                        Case SQLDMODBStat_Recovering
                            sType = "Recovering"
                        Case SQLDMODBStat_Standby
                            sType = "Stand By"
                        Case SQLDMODBStat_Suspect
                            sType = "Suspect"
                        Case SQLDMODBStat_Inaccessible
                            sType = "Inaccessible"
                        Case Else
                            sType = "Unknown"
                    End Select
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Status", sType))
                                        
                    
                    On Local Error Resume Next
                    bOwner = .Isdb_owner
                    If Err.Number <> 0 Then bOwner = False
                    On Local Error GoTo 0
                    
                    If bOwner Then
                        sType = .TransactionLog.LastBackup
                    Else
                        sType = m_sLngRes(enScanErr_OnlyOwner)
                    End If
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanLastBackup), sType))
                    
                    
                    If bOwner Then
                        sType = .Users.Count
                    Else
                        sType = m_sLngRes(enScanErr_OnlyOwner)
                    End If
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanUsers), sType))
                    
                    If bOwner Then
                        sType = .Size
                    Else
                        sType = m_sLngRes(enScanErr_OnlyOwner)
                    End If
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Size (Mb)", sType))
                                    
                    If bOwner Then
                        sType = .SpaceAvailableInMB
                    Else
                        sType = m_sLngRes(enScanErr_OnlyOwner)
                    End If
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanAvailableSpace), sType))
                                    
                    If bOwner Then
                        If .CompatibilityLevel <> SQLDMOCompLevel_70 Then
                            sType = .CompatibilityLevel
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCompatLevel), sType))
                        End If
                    End If
                    
                    If bOwner Then
                        sType = .Collation
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCollation), sType))
                    End If
                                    
                    Print #iOut, kCloseTB
                    
                    sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL"))
                    vVarProp = GetObjectProperties(oDb, sSql)
                    If IsArray(vVarProp) Then
                        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                    
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                        For i = 0 To UBound(vVarProp, 2)
                            sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarProp(0, i), vVarProp(1, i), vVarProp(2, i)))
                        Next
                        Print #iOut, kCloseTB
                    End If
                    vVarProp = vbEmpty
                    
                    Print #iOut, Replace("<P>1%<P>", "1%", m_sLngRes(enScanFileGroup))
    
                    If bOwner Then
                        For i = 1 To .FileGroups.Count
                            Print #iOut, Replace("<P>1%<P>", "1%", m_sLngRes(enScanFileGroup) & " " & .FileGroups(i).Name)
                            
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=35% bgcolor=#ffffcc> 2@% </TD><TD width=10% bgcolor=#ffffcc> 3@% </TD><TD width=10% bgcolor=#ffffcc align=right> 4@% </TD><TD width=7.5% bgcolor=#ffffcc align=right> 5@% </TD><TD width=7.5% bgcolor=#ffffcc align=right> 6@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), Array(m_sLngRes(enScanFileName), m_sLngRes(enScanPhysicalName), m_sLngRes(enScanFileGrowth), m_sLngRes(enScanMaxSize), m_sLngRes(enScanFileSizeKb), m_sLngRes(enScanFileGrowthType)))
                            
                            For l = 1 To .FileGroups(i).DBFiles.Count
                                sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD><TD align=right>5@%</TD><TD align=right>6@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), _
                                            Array(.FileGroups(i).DBFiles(l).Name, .FileGroups(i).DBFiles(l).PhysicalName, .FileGroups(i).DBFiles(l).FileGrowth, .FileGroups(i).DBFiles(l).MaximumSize, .FileGroups(i).DBFiles(l).SizeInKB, IIf(.FileGroups(i).DBFiles(l).FileGrowthType = SQLDMOGrowth_MB, "MB", "%")))
                            Next
                            Print #iOut, kCloseTB
                        Next
                        
                        
                        Print #iOut, Replace("<P>1%<P>", "1%", m_sLngRes(enScanLogFile))
                            
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=35% bgcolor=#ffffcc> 2@% </TD><TD width=10% bgcolor=#ffffcc> 3@% </TD><TD width=10% bgcolor=#ffffcc align=right> 4@% </TD><TD width=7.5% bgcolor=#ffffcc align=right> 5@% </TD><TD width=7.5% bgcolor=#ffffcc align=right> 6@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), Array(m_sLngRes(enScanLogFile), m_sLngRes(enScanPhysicalName), m_sLngRes(enScanFileGrowth), m_sLngRes(enScanMaxSize), m_sLngRes(enScanFileSizeKb), m_sLngRes(enScanFileGrowthType)))
                        
                        sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD><TD align=right>5@%</TD><TD align=right>6@%</TD></TR>"
                        For i = 1 To .TransactionLog.LogFiles.Count
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%"), _
                                        Array(.TransactionLog.LogFiles(i).Name, .TransactionLog.LogFiles(i).PhysicalName, .TransactionLog.LogFiles(i).FileGrowth, .TransactionLog.LogFiles(i).MaximumSize, .TransactionLog.LogFiles(i).SizeInKB, IIf(.TransactionLog.LogFiles(i).FileGrowthType = SQLDMOGrowth_MB, "MB", "%")))
                        Next
                        Print #iOut, kCloseTB
                        
                    Else
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=100%><THEAD><TR>"
                        sBuf = sBuf & "<TD width=100%> 1@% </TD></THEAD><TBODY><TR></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanErr_OnlyOwner)))
                    End If
                    Print #iOut, kCloseTB
                End With
            End If
            
        Else
            Print #iOut, szBuf
        End If
    End If
    
    Close iOut
    ErrWriteDbStuff = lErr
    
End Function
Private Sub ShowHtm()

      Dim sTopic As String
      Dim sParams As Variant
      Dim sDirectory As Variant
      Dim l_strFile As String
      
      l_strFile = mType(1).ElemData.vVal & kDefPage
      
      sTopic = "Open"
      sParams = 0&
      sDirectory = 0&
      Call RunShellExecute(sTopic, l_strFile, sParams, sDirectory, SW_SHOWNORMAL, False)

End Sub
Private Function ErrWriteTB(ByVal oDb As SQLDMO.Database2, ByVal oTb As SQLDMO.Table, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim i As Integer, l As Integer
    Dim sFile As String
    Dim oCol As SQLDMO.Column2
    Dim sType As String
    Dim bUDT As Boolean
    Dim oUdt As SQLDMO.UserDefinedDatatype2
    Dim oType As SQLDMO.SystemDatatype
    Dim sComputedT As String
    Dim oTmpTB As SQLDMO.Table
    Dim sReferTB As String
    Dim sReferTBOwn  As String
    Dim sStripTb As String
    Dim iPos As Integer
    Dim sSql As String
    Dim vVarExtProp As Variant
    Dim x As Integer
    Dim sIndexType As String
    
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
        
    sFile = Replace(kObjPage, "1%", Replace(oTb.Owner & "_" & oTb.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oTb
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScan_Table), ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            Print #iOut, Replace("<P>1@%", "1@%", m_sLngRes(enScanTbDef))
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "<TD width=15% bgcolor=#ffffcc> 3@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc align=right> 4@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 5@% </TD>"
            sBuf = sBuf & "<TD width=15% bgcolor=#ffffcc> 6@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc align=right> 7@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc align=right> 8@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 9@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 10@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc align=right> 1A@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc align=right> 1B@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 1C@% </TD>"
            sBuf = sBuf & "<TD width=5% bgcolor=#ffffcc> 1D@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            
            ReDim vVarProp(0 To 13) As Variant
            vVarProp(0) = m_sLngRes(enScanTb_inPK)
            vVarProp(1) = m_sLngRes(enScanTb_FieldName)
            vVarProp(2) = m_sLngRes(enScanTb_Data_Type)
            vVarProp(3) = m_sLngRes(enScanTb_Size)
            vVarProp(4) = m_sLngRes(enScanTb_AllowNull)
            vVarProp(5) = m_sLngRes(enScanTb_DefVal)
            vVarProp(6) = m_sLngRes(enScanTb_Precision)
            vVarProp(7) = m_sLngRes(enScanTb_Scale)
            vVarProp(8) = m_sLngRes(enScanTb_IsRowGuid)
            vVarProp(9) = m_sLngRes(enScanTb_IsIdentity)
            vVarProp(10) = m_sLngRes(enScanTb_InitialVal)
            vVarProp(11) = m_sLngRes(enScanTb_Increment)
            vVarProp(12) = m_sLngRes(enScanTb_Computed)
            vVarProp(13) = m_sLngRes(enScanCollation)
            
            
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%", "7@%", "8@%", "9@%", "10@%", "1A@%", "1B@%", "1C@%", "1D@%"), vVarProp)
                        
            sBuf = "<TR><TD>1@%</TD><TD><A href='pg_1Z@%.htm'>2@%</A></TD><TD>3@%</TD><TD align=right>4@%</TD><TD>5@%</TD><TD>6@%</TD>"
            sBuf = sBuf & "<TD align=right>7@%</TD><TD align=right>8@%</TD><TD>9@%</TD><TD>10@%</TD><TD align=right>1A@%</TD><TD align=right>1B@%</TD><TD>1C@%</TD><TD>1D@%</TD></TR>"
            For Each oCol In .Columns
                
                sComputedT = ""
                On Local Error Resume Next
                Set oType = objServer.Databases(m_sDbName).SystemDatatypes(oCol.DataType)
                bUDT = Err.Number <> 0
                On Local Error GoTo 0
                ReDim vVarArray(0 To 14) As Variant

                vVarArray(UBound(vVarArray)) = Replace(oTb.Owner & "_" & oTb.Name & "_" & oCol.Name, " ", "_")
                vVarArray(0) = oCol.InPrimaryKey
                vVarArray(1) = oCol.Name
                
                vVarArray(9) = oCol.Identity
                vVarArray(8) = oCol.IsRowGuidCol
                vVarArray(10) = oCol.IdentitySeed
                vVarArray(11) = oCol.IdentityIncrement
                vVarArray(12) = oCol.IsComputed
                
                If Not bUDT Then
                    vVarArray(2) = oCol.DataType
                    vVarArray(4) = oCol.AllowNulls
                    vVarArray(5) = oCol.DRIDefault.Text
                    vVarArray(6) = oCol.NumericPrecision
                    vVarArray(7) = oCol.NumericScale
                    vVarArray(3) = oCol.Length
                    vVarArray(13) = oCol.Collation
                Else
                    
'                    vVarArray(2) = oCol.DataType & " (" & objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).BaseType & ")"
'                    vVarArray(4) = objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).AllowNulls
'                    vVarArray(5) = objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).Default
'
'
'                    If objServer.Databases(m_sDbName).SystemDatatypes(objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).BaseType).IsNumeric Then
'                        vVarArray(6) = objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).NumericPrecision
'                        vVarArray(7) = objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).NumericScale
'                    Else
'                        vVarArray(6) = ""
'                        vVarArray(7) = ""
'                    End If
'                    vVarArray(3) = objServer.Databases(m_sDbName).UserDefinedDatatypes(oCol.DataType).MaxSize

                    Set oUdt = oDb.UserDefinedDatatypes(oCol.DataType)
                    vVarArray(2) = oCol.DataType & " (" & oUdt.BaseType & ")"
                    vVarArray(4) = oUdt.AllowNulls
                    vVarArray(5) = oUdt.Default
                                        
                    
                    If objServer.Databases(m_sDbName).SystemDatatypes(oUdt.BaseType).IsNumeric Then
                        vVarArray(6) = oUdt.NumericPrecision
                        vVarArray(7) = oUdt.NumericScale
                    Else
                        vVarArray(6) = ""
                        vVarArray(7) = ""
                    End If
                    vVarArray(3) = oUdt.MaxSize
                    vVarArray(13) = oUdt.Collation
                    Set oUdt = Nothing
                End If
                If oCol.IsComputed Then sComputedT = oCol.ComputedText
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%", "6@%", "7@%", "8@%", "9@%", "10@%", "1A@%", "1B@%", "1C@%", "1D@%", "1Z@%"), vVarArray)
                
                lErr = ErrWriteCol(oDb, oTb, oCol, sDir, vVarProp, vVarArray, sComputedT, sErr)
                
            Next
            
            Print #iOut, kCloseTB
            
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oTb.Owner)), "'table'", SqlFinalS(ApiciSQL(oTb.Name)), "NULL", "NULL"))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For i = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, i), vVarExtProp(1, i), vVarExtProp(2, i)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            
            If oTb.Keys.Count Then
                For i = 1 To oTb.Keys.Count
                    If (oTb.Keys(i).Type And SQLDMOKey_Foreign) = SQLDMOKey_Foreign Then
                    
                        Print #iOut, "<BR>" & m_sLngRes(enScanTb_Index_Key_Constr)
                        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                        sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
                        sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
                    
                        sReferTB = ""
                        sReferTBOwn = ""
                        
                        'Set oTmpTB = objServer.Databases(m_sdbName).Tables(oTB.Keys(i).ReferencedTable)
                        
                        sStripTb = oTb.Keys(i).ReferencedTable
                        iPos = InStr(sStripTb, ".")
                        If iPos Then sStripTb = Mid$(sStripTb, iPos + 1)
                        
                        Set oTmpTB = GetTbByName(objServer, m_sDbName, sStripTb)
                        
                        
                        If Not oTmpTB Is Nothing Then
                            sReferTB = oTmpTB.Name
                            sReferTBOwn = oTmpTB.Owner
                            sType = m_sLngRes(enScanTb_FK)
                            sBuf = "<TR><TD><B><I>1@%</I></B></TD><TD><B><I>2@%</I></B></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanName), oTb.Keys(i).Name))
                            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKtype), sType))
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKclustered), oTb.Keys(i).Clustered))
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKfillfactor), oTb.Keys(i).FillFactor))
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcheck), oTb.Keys(i).Checked))
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Exclude Replication", oTb.Keys(i).ExcludeReplication))
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanFileGroup), oTb.Keys(i).FileGroup))
                        
                            For l = 1 To oTb.Keys(i).KeyColumns.Count
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcolumns), oTb.Keys(i).KeyColumns(l)))
                            Next
    
                            Print #iOut, ReplaceMsg("<TR><TD>1@%</TD><TD><A href='pg_1Z@%_1Y@%.htm'>2@%</A></TR>", Array("1@%", "2@%", "1Z@%", "1Y@%"), Array(m_sLngRes(enScanTb_Htm4TB), sReferTB, Replace(sReferTBOwn, " ", "_"), Replace(sReferTB, " ", "_")))
                            For l = 1 To oTb.Keys(i).ReferencedColumns.Count
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcolumns), oTb.Keys(i).ReferencedColumns(l)))
                            Next
                        End If
                        Print #iOut, kCloseTB
                        
                        
                        sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oTb.Owner)), "'table'", SqlFinalS(ApiciSQL(oTb.Name)), "'constraint'", SqlFinalS(ApiciSQL(oTb.Keys(i).Name))))
                        vVarExtProp = GetObjectProperties(oDb, sSql)
                        If IsArray(vVarExtProp) Then
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                        
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                            For x = 0 To UBound(vVarExtProp, 2)
                                sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, x), vVarExtProp(1, x), vVarExtProp(2, x)))
                            Next
                            Print #iOut, kCloseTB
                        End If
                        vVarExtProp = vbEmpty
                        
                    End If
                Next
                
                Set oTmpTB = Nothing
            End If
                                
            If oTb.Indexes.Count Then
                For i = 1 To oTb.Indexes.Count
                    Print #iOut, "<BR>" & m_sLngRes(enScan_Index)
                    sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                    sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                    sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
                    sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
                
                    If ((oTb.Indexes(i).Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex) And ((oTb.Indexes(i).Type And SQLDMOIndex_Hypothetical) <> SQLDMOIndex_Hypothetical) Then
                        Select Case True
                            Case (oTb.Indexes(i).Type And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey
                                sType = m_sLngRes(enScanCol_PK)
                            Case (oTb.Indexes(i).Type And SQLDMOIndex_DRIUniqueKey) = SQLDMOIndex_DRIUniqueKey, (oTb.Indexes(i).Type And SQLDMOIndex_Unique) = SQLDMOIndex_Unique
                                sType = m_sLngRes(enScanCol_UK)
                            Case (oTb.Indexes(i).Type And SQLDMOIndex_PadIndex) = SQLDMOIndex_PadIndex
                                sType = "PadIndex"
                        End Select
                        
                        
                        sBuf = "<TR><TD><B><I>1@%</I></B></TD><TD><B><I>2@%</I></B></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanName), oTb.Indexes(i).Name))
                        sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKtype), sType))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKfillfactor), oTb.Indexes(i).FillFactor))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcheck), oTb.Indexes(i).CheckIndex))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("No Recompute", oTb.Indexes(i).NoRecompute))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanFileGroup), oTb.Indexes(i).FileGroup))
                        
                        For Each oCol In oTb.Indexes(i).ListIndexedColumns
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcolumns), oCol.Name))
                        Next
                        Print #iOut, kCloseTB
                                                

                        If (oTb.Indexes(i).Type And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey Then
                            sIndexType = "'constraint'"
                        Else
                            sIndexType = "'index'"
                        End If
                        sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oTb.Owner)), "'table'", SqlFinalS(ApiciSQL(oTb.Name)), sIndexType, SqlFinalS(ApiciSQL(oTb.Indexes(i).Name))))
                        vVarExtProp = GetObjectProperties(oDb, sSql)
                        If IsArray(vVarExtProp) Then
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                        
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                            For x = 0 To UBound(vVarExtProp, 2)
                                sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, x), vVarExtProp(1, x), vVarExtProp(2, x)))
                            Next
                            Print #iOut, kCloseTB
                        End If
                        vVarExtProp = vbEmpty
                        
                    End If
                Next
                
                Set oTmpTB = Nothing
            End If
            
            If oTb.Checks.Count Then
                Print #iOut, "<BR>" & m_sLngRes(enScanTb_Checkconstr)
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
                
                
                For i = 1 To oTb.Checks.Count
                    sBuf = "<TR><TD><B><I>1@%</I></B></TD><TD><B><I>2@%</I></B></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanName), oTb.Checks(i).Name))
                    sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcheck), oTb.Checks(i).Checked))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_CheckText), oTb.Checks(i).Text))
                Next
                Print #iOut, kCloseTB
            End If
            
'------------ Triggers
            If oTb.Triggers.Count Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 1@% </TD>"
                sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 2@% </TD>"
                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Trigger", m_sLngRes(enScanOwner)))
            
                
                sBuf = "<TR><TD><A href='pg_1Z@%.htm'>1@%</A></TD><TD>2@%</TD></TR>"
                For i = 1 To oTb.Triggers.Count
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "1Z@%"), Array(oTb.Triggers(i).Name, oTb.Triggers(i).Owner, Replace(oTb.Triggers(i).Owner & "_" & oTb.Triggers(i).Name, " ", "_")))
                    lErr = ErrWriteTRIG(oDb, oTb, oTb.Triggers(i), sDir, sErr, True)
                Next
                Print #iOut, kCloseTB
            
'                Print #iOut, "<BR>Triggers"
'                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
'                sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
'                sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
'                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
'                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Property), MyLoadResString(k_Value)))
'
'                sBuf = "<TR><TD bgcolor=cyan>1@%</TD><TD bgcolor=cyan>2@%</TD></TR>"
'                For i = 1 To oTB.Triggers.Count
'                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Name), oTB.Triggers(i).Name))
'                    sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
'                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Create_Date), oTB.Triggers(i).CreateDate))
'                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Quoted Identifier", oTB.Triggers(i).QuotedIdentifierStatus))
'                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Ansi Nulls", oTB.Triggers(i).AnsiNullsStatus))
'                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_HtmText), oTB.Triggers(i).Text))
'                Next
'                Print #iOut, kCloseTB
            End If
                        
                
            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, tableDrop, oCollChild, oCollParent, sErr)
            
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
            'sType = GetScript(oTB, opt_Tbl)
            'sType = Replace(sType, vbCrLf, "<BR>")
                        
            sBuf = "<TR><TD>1@%</TD></TR>"
            'Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oTb, opt_Tbl) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteTB = lErr

End Function
Private Function ErrWriteCol(ByVal oDb As SQLDMO.Database2, ByVal oTb As SQLDMO.Table, ByVal oCol As SQLDMO.Column, _
                                ByVal sDir As String, _
                                vVarProp As Variant, vVarData As Variant, ByVal sComputedT As String, _
                                ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim i As Integer
    Dim sFile As String, sParentFile As String
    Dim oList As SQLDMO.SQLObjectList
    Dim oKey As SQLDMO.Key, oTmpTB As SQLDMO.Table
    Dim sType As String, sParent As String
    Dim sReferTB As String, sReferTBOwn As String
    Dim sStripName As String, iPos As Integer
    Dim sSql As String, vVarExtProp As Variant
    
        
    sFile = Replace(kObjPage, "1%", Replace(oTb.Owner & "_" & oTb.Name & "_" & oCol.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        sParentFile = sDir & Replace(kObjPage, "1%", Replace(oTb.Owner & "_" & oTb.Name, " ", "_"))
        With oCol
            Print #iOut, "<HTML><HEAD><TITLE>" & Replace(m_sLngRes(enScanColDef), "1@%", .Name) & "</HEAD></TITLE><H2>" & Replace(m_sLngRes(enScanColDef), "1@%", .Name) & "</H2><BR>"
                       
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            Print #iOut, ReplaceMsg("<A href='1@%'>2@%</A>", Array("1@%", "2@%"), Array(sParentFile, oTb.Name))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For i = 0 To UBound(vVarProp)
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(vVarProp(i), vVarData(i)))
            Next
            Print #iOut, kCloseTB
            
            If .IsComputed Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc> 1@% </TD>"
                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanColComputedText)))
                
                Print #iOut, ReplaceMsg("<TR><TD>1@%</TD></TR>", Array("1@%"), Array(.ComputedText))
                Print #iOut, kCloseTB
            End If
                        
            Set oList = oCol.ListKeys
            If Not oList Is Nothing Then
                If oList.Count <> 0 Then
                    sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                    sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD>"
                    sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 2@% </TD>"
                    sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 3@% </TD>"
                    sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 4@% </TD>"
                    sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(m_sLngRes(enScanColList), m_sLngRes(enScanColPriTB), m_sLngRes(enScanTb_Htm4TB), m_sLngRes(enScanTb_FKtype)))

                    For i = 1 To oList.Count
                        Set oKey = oList.Item(i)
                        Set oTmpTB = oKey.Parent
                        sParent = oTmpTB.Name
                                                
                        sReferTB = ""
                        sReferTBOwn = ""
                        Select Case oKey.Type
                            Case SQLDMOKey_Unknown
                                sType = MyLoadResString(k_HtmUNKNOWN)
                            Case SQLDMOKey_Primary
                                sType = m_sLngRes(enScanCol_PK)
                            Case SQLDMOKey_Unique
                                sType = m_sLngRes(enScanCol_UK)
                            Case SQLDMOKey_Foreign
                                sType = m_sLngRes(enScanTb_FK)
                                'Set oTmpTB = objServer.Databases(m_sdbName).Tables(oKey.ReferencedTable)
                                sStripName = oKey.ReferencedTable
                                iPos = InStr(sStripName, ".")
                                If iPos <> 0 Then sStripName = Mid$(sStripName, iPos + 1)
                                Set oTmpTB = GetTbByName(objServer, m_sDbName, sStripName)
                                
                                If Not oTmpTB Is Nothing Then
                                    sReferTB = oTmpTB.Name
                                    sReferTBOwn = oTmpTB.Owner
                                End If
                        End Select
                        If oKey.Type = SQLDMOKey_Foreign Then
                            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD><A href='pg_1Z@%_1Y@%.htm'>3@%</A></TD><TD>4@%</TD></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "1Z@%", "1Y@%"), Array(oKey.Name, sParent, sReferTB, sType, Replace(sReferTBOwn, " ", "_"), Replace(sReferTB, " ", "_")))
                        Else
                            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD>3@%</TD><TD>4@%</TD></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(oKey.Name, sParent, sReferTB, sType))
                        End If
                    Next
                    Set oTmpTB = Nothing
                    Set oKey = Nothing
                    Print #iOut, kCloseTB
                
                End If
            End If
            Set oList = Nothing
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oTb.Owner)), "'table'", SqlFinalS(ApiciSQL(oTb.Name)), "'column'", SqlFinalS(ApiciSQL(oCol.Name))))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For i = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, i), vVarExtProp(1, i), vVarExtProp(2, i)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If

    Close iOut
    ErrWriteCol = lErr

End Function


Private Function ErrWriteDependent(ByVal iOutFile As Integer, ByVal sObj As String, ByVal bChild As Boolean, ByVal oColl As Collection, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim i As Integer
    Dim sBuf As String
    Dim sType As String
    
'Public Const k_Objects_obj_THAT_Depends_on = 561            ' Oggetti che dipendono DA
'Public Const k_Objects_THAT_obj_Depends_on = 562            ' Oggetti dai quali % Dipende
    
    If oColl.Count Then
        If bChild Then
            sBuf = m_sLngRes(enScan_DedendON) & " [" & sObj & "]"
        Else
            sBuf = Replace(m_sLngRes(enScan_DedendFROM), "%", " [" & sObj & "]")
        End If
        Print #iOutFile, "<BR>" & sBuf
        
        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
        sBuf = sBuf & "<TD width=15% bgcolor=#ffffcc> 1@% </TD>"
        sBuf = sBuf & "<TD width=60% bgcolor=#ffffcc> 2@% </TD>"
        sBuf = sBuf & "<TD width=15% bgcolor=#ffffcc> 3@% </TD>"
        sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc align=right> 4@% </TD>"
        sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
        Print #iOutFile, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(m_sLngRes(enScanOwner), m_sLngRes(enScan_Object), m_sLngRes(enScan_Type), "Sequence"))

        
        sBuf = "<TR><TD>1@%</TD><TD><A href='pg_1@%_1Z@%.htm'>2@%</A></TD><TD>3@%</TD><TD align=right>4@%</TD></TR>"
        For i = 1 To oColl.Count
            With oColl.Item(i)
                Debug.Print .oType, .oName
                If .oType = SQLDMOObj_SystemTable Or .oType = SQLDMOObj_UserTable Then
                    sType = m_sLngRes(enScan_Table)
                ElseIf .oType = SQLDMOObj_View Then
                    sType = m_sLngRes(enScan_View)
                ElseIf .oType = SQLDMOObj_StoredProcedure Then
                    sType = m_sLngRes(enScan_StoredP)
                ElseIf .oType = SQLDMOObj_Default Then
                    sType = "Default"
                ElseIf .oType = SQLDMOObj_Rule Then
                    sType = "Rule"
                ElseIf .oType = SQLDMOObj_UserDefinedDatatype Then
                    sType = m_sLngRes(enScan_UDT)
                ElseIf .oType = SQLDMOObj_Trigger Then
                   sType = "Trigger"
                ElseIf .oType = SQLDMOObj_UserDefinedFunction Then
                    'funzione
                    sType = m_sLngRes(enScan_UdFunc)
                Else
                    sType = MyLoadResString(k_HtmUNKNOWN)
                End If
                Print #iOutFile, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "1Z@%"), Array(Replace(.oOwner(), " ", "_"), .oName(), sType, .oSequence(), Replace(.oName(), " ", "_")))
            End With
            
        Next
        Print #iOutFile, kCloseTB
    End If
    ErrWriteDependent = lErr
    
End Function

Private Function ErrWriteSP(ByVal oDb As SQLDMO.Database2, ByVal oSP As SQLDMO.StoredProcedure2, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim i As Integer, l As Integer
    Dim sFile As String
    Dim oQry As SQLDMO.QueryResults
    Dim sSql As String, vVarExtProp As Variant
    Dim sParameter As String
    Dim sDirection As String, sName As String
    
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
        
    sFile = Replace(kObjPage, "1%", Replace(oSP.Owner & "_" & oSP.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oSP
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScan_StoredP), ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Ansi Nulls", .AnsiNullsStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Quoted Identifiers", .QuotedIdentifierStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("StartUp", .Startup))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Param_Encrypted), .Encrypted))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Type), StoredProcType(.Type)))
            Print #iOut, kCloseTB
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oSP.Owner)), "'procedure'", SqlFinalS(ApiciSQL(oSP.Name)), "NULL", "NULL"))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For l = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, l), vVarExtProp(1, l), vVarExtProp(2, l)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            Set oQry = .EnumParameters
            If Not oQry Is Nothing Then
                If oQry.Rows > 0 Then
                    For i = 1 To oQry.Rows
                        Print #iOut, "<BR>" & m_sLngRes(enScan_Parameter)
                        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 2@% </TD>"
                        sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 3@% </TD>"
                        sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc> 4@% </TD>"
                        sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(oQry.ColumnName(1), oQry.ColumnName(2), oQry.ColumnName(3), oQry.ColumnName(4)))
                        
                        sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD></TR>"
                    
                        sParameter = oQry.GetColumnString(i, 1)
                        sDirection = oQry.GetColumnLong(i, 4)
                        If Len(sDirection) <> 0 Then
                            sDirection = IIf(sDirection = "1", "input/output", "input")
                        Else
                            sDirection = "@RETURN_VALUE"
                        End If
                        
                        sName = oQry.GetColumnString(i, 1)
                        If Len(sName) = 0 Then sName = "@RetParam " & i + 1
                        
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(sName, oQry.GetColumnString(i, 2), oQry.GetColumnLong(i, 3), sDirection))
                        Print #iOut, kCloseTB
                        
                        sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oSP.Owner)), "'procedure'", SqlFinalS(ApiciSQL(oSP.Name)), "'parameter'", SqlFinalS(ApiciSQL(sParameter))))
                        vVarExtProp = GetObjectProperties(oDb, sSql)
                        If IsArray(vVarExtProp) Then
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                        
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                            For l = 0 To UBound(vVarExtProp, 2)
                                sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, l), vVarExtProp(1, l), vVarExtProp(2, l)))
                            Next
                            Print #iOut, kCloseTB
                        End If
                        vVarExtProp = vbEmpty
                        
                    Next
                End If
            End If
                           
            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, spDrop, oCollChild, oCollParent, sErr)
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
            sBuf = "<TR><TD>1@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oSP, opt_SP) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteSP = lErr

End Function

Private Function ErrWriteView(ByVal oDb As SQLDMO.Database2, ByVal oView As SQLDMO.View2, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sFile As String
    Dim oCol As SQLDMO.Column2
    Dim oIdx As SQLDMO.Index2
    Dim oTrig As SQLDMO.Trigger2
    Dim oList As SQLDMO.SQLObjectList
    Dim lType As Long, sType As String
    Dim x As Integer
    
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
    Dim i As Integer
    Dim sSql As String, vVarExtProp As Variant
        
    sFile = Replace(kObjPage, "1%", Replace(oView.Owner & "_" & oView.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oView
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScan_View), ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Ansi Nulls", .AnsiNullsStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Quoted Identifiers", .QuotedIdentifierStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Func_Param_IsSchemaBound), .IsSchemaBound))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Param_Encrypted), .Encrypted))
                   
            
            Print #iOut, kCloseTB
            
            
            For Each oCol In .ListColumns
                Print #iOut, "<BR>" & m_sLngRes(enScan_ViewColumn)
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
                sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FieldName), m_sLngRes(enScanTb_Data_Type)))
                
                sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oCol.Name, oCol.DataType))
                Print #iOut, kCloseTB
                
                sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oView.Owner)), "'view'", SqlFinalS(ApiciSQL(oView.Name)), "'column'", SqlFinalS(ApiciSQL(oCol.Name))))
                vVarExtProp = GetObjectProperties(oDb, sSql)
                If IsArray(vVarExtProp) Then
                    sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                    sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                    For i = 0 To UBound(vVarExtProp, 2)
                        sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, i), vVarExtProp(1, i), vVarExtProp(2, i)))
                    Next
                    Print #iOut, kCloseTB
                End If
                vVarExtProp = vbEmpty
                
            Next
            
                          
            
'***        'Indici
            If oView.Indexes.Count <> 0 Then

                For i = 1 To oView.Indexes.Count
                    Print #iOut, "<BR>" & m_sLngRes(enScan_Index)
                    sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                    sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                    sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
                    sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
                
                    lType = oView.Indexes(i).Type
                    Select Case True
                        Case (lType And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey
                            sType = m_sLngRes(enScanCol_PK)
                        Case (lType And SQLDMOIndex_DRIUniqueKey) = SQLDMOIndex_DRIUniqueKey, (lType And SQLDMOIndex_Unique) = SQLDMOIndex_Unique
                            sType = m_sLngRes(enScanCol_UK)
                        Case (lType And SQLDMOIndex_PadIndex) = SQLDMOIndex_PadIndex
                            sType = "PadIndex"
                    End Select
                
                    Debug.Print oView.Indexes(i).Type
                    sBuf = "<TR><TD><B><I>1@%</I></B></TD><TD><B><I>2@%</I></B></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanName), oView.Indexes(i).Name))
                    sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKtype), sType))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKfillfactor), oView.Indexes(i).FillFactor))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcheck), oView.Indexes(i).CheckIndex))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("No Recompute", oView.Indexes(i).NoRecompute))
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanFileGroup), oView.Indexes(i).FileGroup))
                        
                        
                    Set oList = oView.Indexes(i).ListIndexedColumns
                    For Each oCol In oView.Indexes(i).ListIndexedColumns
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FKcolumns), oCol.Name))
                    Next
                    Print #iOut, kCloseTB
        
                    sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oView.Owner)), "'view'", SqlFinalS(ApiciSQL(oView.Name)), "'index'", SqlFinalS(ApiciSQL(oView.Indexes(i).Name))))
                    vVarExtProp = GetObjectProperties(oDb, sSql)
                    If IsArray(vVarExtProp) Then
                        sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                    
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                        For x = 0 To UBound(vVarExtProp, 2)
                            sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, x), vVarExtProp(1, x), vVarExtProp(2, x)))
                        Next
                        Print #iOut, kCloseTB
                    End If
                    vVarExtProp = vbEmpty
        
                Next
                
            End If

'*****      Trigger
            If oView.Triggers.Count Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 1@% </TD>"
                sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 2@% </TD>"
                sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Trigger", m_sLngRes(enScanOwner)))
            
                
                sBuf = "<TR><TD><A href='pg_1Z@%.htm'>1@%</A></TD><TD>2@%</TD></TR>"
                For i = 1 To oView.Triggers.Count
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "1Z@%"), Array(oView.Triggers(i).Name, oView.Triggers(i).Owner, Replace(oView.Triggers(i).Owner & "_" & oView.Triggers(i).Name, " ", "_")))
                    lErr = ErrWriteTRIG(oDb, oView, oView.Triggers(i), sDir, sErr, False)
                Next
                Print #iOut, kCloseTB
            
            End If



            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, viewDrop, oCollChild, oCollParent, sErr)
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oView.Owner)), "'view'", SqlFinalS(ApiciSQL(oView.Name)), "NULL", "NULL"))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For i = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, i), vVarExtProp(1, i), vVarExtProp(2, i)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
            'sType = GetScript(oView, opt_View)
            'sType = Replace(sType, vbCrLf, "<BR>")
            
            sBuf = "<TR><TD>1@%</TD></TR>"
            'Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oView, opt_View) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteView = lErr

End Function
Private Function ErrWriteUDT(ByVal oObj As SQLDMO.UserDefinedDatatype, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sFile As String
    Dim oCol As SQLDMO.Column
    Dim oTmpTB As SQLDMO.Table
    
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
        
    sFile = Replace(kObjPage, "1%", Replace(oObj.Owner & "_" & oObj.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oObj
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScan_UDT), ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_Data_Type), .BaseType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Allow Identity", .AllowIdentity))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_AllowNull), .AllowNulls))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("IsVariableLength", .IsVariableLength))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Lenght), .Length))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Max Size", .MaxSize))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_Precision), .NumericPrecision))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_Scale), .NumericScale))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Rule", .Rule))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_DefVal), .Default))
            Print #iOut, kCloseTB
            
            
            Print #iOut, "<BR>" & m_sLngRes(enScan_BoundCols)
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FieldName), m_sLngRes(enScan_Table)))
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For Each oCol In .ListBoundColumns
                Set oTmpTB = oCol.Parent
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oCol.Name, oTmpTB.Name))
            Next
            Print #iOut, kCloseTB
            Set oTmpTB = Nothing
                           
            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, udtDrop, oCollChild, oCollParent, sErr)
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
'            sType = GetScript(oObj, opt_UDT)
'            sType = Replace(sType, vbCrLf, "<BR>")
            
            sBuf = "<TR><TD>1@%</TD></TR>"
            'Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oObj, opt_UDT) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteUDT = lErr

End Function
Private Function ErrWriteRULE(ByVal oObj As SQLDMO.Rule, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sFile As String
    Dim oCol As SQLDMO.Column
    Dim oTmpTB As SQLDMO.Table
        
    Dim oUdt As SQLDMO.UserDefinedDatatype
    
        
    sFile = Replace(kObjPage, "1%", Replace(oObj.Owner & "_" & oObj.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oObj
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array("Rule", ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Text), .Text))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
            Print #iOut, kCloseTB
            
            
            Print #iOut, "<BR>" & m_sLngRes(enScan_BoundCols)
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FieldName), m_sLngRes(enScan_Table)))
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For Each oCol In .ListBoundColumns
                Set oTmpTB = oCol.Parent
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oCol.Name, oTmpTB.Name))
            Next
            Print #iOut, kCloseTB
            Set oTmpTB = Nothing
                           
                           
            Print #iOut, "<BR>" & m_sLngRes(enScan_BoundUDT)
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_UDT), m_sLngRes(enScanOwner)))
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For Each oUdt In .ListBoundDatatypes
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oUdt.Name, oUdt.Owner))
            Next
            Print #iOut, kCloseTB
            Set oUdt = Nothing
                           
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
'            sType = GetScript(oObj, opt_Rule)
'            sType = Replace(sType, vbCrLf, "<BR>")
            
            sBuf = "<TR><TD>1@%</TD></TR>"
'            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oObj, opt_Rule) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If

    Close iOut
    ErrWriteRULE = lErr

End Function
Private Function ErrWriteDEFAULT(ByVal oObj As SQLDMO.Default, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sFile As String
    Dim oCol As SQLDMO.Column
    Dim oTmpTB As SQLDMO.Table
        
    Dim oUdt As SQLDMO.UserDefinedDatatype
    
        
    sFile = Replace(kObjPage, "1%", Replace(oObj.Owner & "_" & oObj.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oObj
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array("Default", ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Text), .Text))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
            Print #iOut, kCloseTB
            
            
            Print #iOut, "<BR>" & m_sLngRes(enScan_BoundCols)
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanTb_FieldName), m_sLngRes(enScan_Table)))
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For Each oCol In .ListBoundColumns
                Set oTmpTB = oCol.Parent
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oCol.Name, oTmpTB.Name))
            Next
            Print #iOut, kCloseTB
            Set oTmpTB = Nothing
                           
                           
            Print #iOut, "<BR>" & m_sLngRes(enScan_BoundUDT)
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=50% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_UDT), m_sLngRes(enScanOwner)))
            
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            For Each oUdt In .ListBoundDatatypes
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(oUdt.Name, oUdt.Owner))
            Next
            Print #iOut, kCloseTB
            Set oUdt = Nothing
                           
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
'            sType = GetScript(oObj, opt_Def)
'            sType = Replace(sType, vbCrLf, "<BR>")
            
            sBuf = "<TR><TD>1@%</TD></TR>"
'            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oObj, opt_Def) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If

    Close iOut
    ErrWriteDEFAULT = lErr

End Function

Private Function ErrWriteTRIG(ByVal oDb As SQLDMO.Database2, ByVal objParent As Object, ByVal oObj As SQLDMO.Trigger2, ByVal sDir As String, ByRef sErr As String, ByVal bIsTable As Boolean) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim sFile As String
    Dim sParentFile As String
    Dim i As Integer
    Dim sSql As String, vVarExtProp As Variant
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
            
       
    sFile = Replace(kObjPage, "1%", Replace(oObj.Owner & "_" & oObj.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oObj
            sParentFile = sDir & Replace(kObjPage, "1%", Replace(objParent.Owner & "_" & objParent.Name, " ", "_"))
            
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array("Trigger", ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            Print #iOut, ReplaceMsg("<A href='1@%'>2@%</A>", Array("1@%", "2@%"), Array(sParentFile, objParent.Name))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanName), oObj.Name))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), oObj.CreateDate))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Quoted Identifier", oObj.QuotedIdentifierStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Ansi Nulls", oObj.AnsiNullsStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Param_Encrypted), oObj.Encrypted))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Instead Of", oObj.InsteadOfTrigger))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Type), TriggerType(oObj.Type)))
            Print #iOut, kCloseTB
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(objParent.Owner)), IIf(bIsTable, "'table'", "'view'"), SqlFinalS(ApiciSQL(objParent.Name)), "'trigger'", SqlFinalS(ApiciSQL(oObj.Name))))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For i = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, i), vVarExtProp(1, i), vVarExtProp(2, i)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScan_Text)))
            sBuf = "<TR><TD>1@%</TD></TR>"
            
            'Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(Replace(oObj.Text, vbCrLf, "<BR>")))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & oObj.Text & "</PRE>"))
            Print #iOut, kCloseTB
            
            
            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, triggerDROP, oCollChild, oCollParent, sErr, objParent)
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteTRIG = lErr

End Function
Private Function ErrWriteUserFunc(ByVal oDb As SQLDMO.Database2, ByVal oFunc As SQLDMO.UserDefinedFunction, ByVal sDir As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim iOut As Integer
    Dim sBuf As String
    Dim i As Integer, l As Integer
    Dim sFile As String
    Dim oQry As SQLDMO.QueryResults
    Dim sSql As String, vVarExtProp As Variant
    Dim sParameter As String
    Dim sDirection As String
    Dim sName As String
        
    
    Dim oObjDep As clsDepen
    Dim oCollChild As Collection, oCollParent As Collection
        
    sFile = Replace(kObjPage, "1%", Replace(oFunc.Owner & "_" & oFunc.Name, " ", "_"))
    
    iOut = FreeFile
    On Local Error Resume Next
    Open sDir & sFile For Output As iOut
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With oFunc
            Print #iOut, ReplaceMsg("<HTML><HEAD><TITLE>1@% [2@%]</HEAD></TITLE><H2>1@% 2@%</H2><BR>", Array("1@%", "2@%"), Array(m_sLngRes(enScan_UdFunc), ReplaceMsg("1%.[2%]", Array("1%", "2%"), Array(.Owner, .Name))))
            
            Print #iOut, ReplaceMsg(kRefMain, Array("1@%", "2@%"), Array(sDir, m_sLngRes(enScanHOME)))
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
            sBuf = sBuf & "<TD width=70% bgcolor=#ffffcc> 2@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanProperty), m_sLngRes(enScanValue)))
            sBuf = "<TR><TD>1@%</TD><TD>2@%</TD></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScanCreateDate), .CreateDate))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Ansi Nulls", .AnsiNullsStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array("Quoted Identifiers", .QuotedIdentifierStatus))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Param_Encrypted), .Encrypted))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(m_sLngRes(enScan_Type), UDFtype(.Type)))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Func_Param_IsDeterministic), .IsDeterministic))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%"), Array(MyLoadResString(k_Func_Param_IsSchemaBound), .IsSchemaBound))
            
            Print #iOut, kCloseTB
            
            sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oFunc.Owner)), "'function'", SqlFinalS(ApiciSQL(oFunc.Name)), "NULL", "NULL"))
            vVarExtProp = GetObjectProperties(oDb, sSql)
            If IsArray(vVarExtProp) Then
                sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                            
                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                For l = 0 To UBound(vVarExtProp, 2)
                    sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                    Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, l), vVarExtProp(1, l), vVarExtProp(2, l)))
                Next
                Print #iOut, kCloseTB
            End If
            vVarExtProp = vbEmpty
            
            
            Set oQry = oFunc.EnumOutputs
            If Not oQry Is Nothing Then
                If oQry.Rows > 0 Then
                    For i = 1 To oQry.Rows
                        Print #iOut, "<BR>" & m_sLngRes(enScan_Parameter)
                        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 2@% </TD>"
                        sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 3@% </TD>"
                        sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc> 4@% </TD>"
                        'sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc> 5@% </TD>"
                        sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        'Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(oQry.ColumnName(1), oQry.ColumnName(2), oQry.ColumnName(3), oQry.ColumnName(4), ""))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(oQry.ColumnName(1), oQry.ColumnName(2), oQry.ColumnName(3), oQry.ColumnName(4)))
                        
                        
                        'sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD><TD align=right>5@%</TD></TR>"
                        sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD></TR>"
                    
                    
                        sParameter = oQry.GetColumnString(i, 1)
                        sDirection = oQry.GetColumnLong(i, 4)
                        If Len(sDirection) <> 0 Then
                            sDirection = IIf(sDirection = "1", "input/output", "input")
                        Else
                            sDirection = "@RETURN_VALUE"
                        End If
                        
                        sName = oQry.GetColumnString(i, 1)
                        If Len(sName) = 0 Then sName = "@RetParam " & i + 1
                        
                        'Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(sName, oQry.GetColumnString(i, 2), oQry.GetColumnLong(i, 3), sDirection))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(sName, oQry.GetColumnString(i, 2), oQry.GetColumnLong(i, 3), sDirection))
                        Print #iOut, kCloseTB
                        
                        sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oFunc.Owner)), "'function'", SqlFinalS(ApiciSQL(oFunc.Name)), "'parameter'", SqlFinalS(ApiciSQL(sParameter))))
                        vVarExtProp = GetObjectProperties(oDb, sSql)
                        If IsArray(vVarExtProp) Then
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                        
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                            For l = 0 To UBound(vVarExtProp, 2)
                                sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, l), vVarExtProp(1, l), vVarExtProp(2, l)))
                            Next
                            Print #iOut, kCloseTB
                        End If
                        vVarExtProp = vbEmpty
                        
                    Next
                    'Print #iOut, kCloseTB
                End If
            End If
            
            
            
            Set oQry = .EnumParameters
            If Not oQry Is Nothing Then
                If oQry.Rows > 0 Then
                    For i = 1 To oQry.Rows
                        Print #iOut, "<BR>" & m_sLngRes(enScan_Parameter)
                        sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 1@% </TD>"
                        sBuf = sBuf & "<TD width=30% bgcolor=#ffffcc> 2@% </TD>"
                        sBuf = sBuf & "<TD width=20% bgcolor=#ffffcc> 3@% </TD>"
                        sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc> 4@% </TD>"
                        'sBuf = sBuf & "<TD width=10% bgcolor=#ffffcc> 5@% </TD>"
                        sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                        'Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(oQry.ColumnName(1), oQry.ColumnName(2), oQry.ColumnName(3), oQry.ColumnName(4), oQry.ColumnName(5)))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(oQry.ColumnName(1), oQry.ColumnName(2), oQry.ColumnName(3), oQry.ColumnName(4)))
                        
                        'sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD><TD align=right>5@%</TD></TR>"
                        sBuf = "<TR><TD>1@%</TD><TD>2@%</TD><TD align=right>3@%</TD><TD align=right>4@%</TD></TR>"
                    
                        
                        sParameter = oQry.GetColumnString(i, 1)
                        sDirection = oQry.GetColumnLong(i, 4)
                        If Len(sDirection) <> 0 Then
                            sDirection = IIf(sDirection = "1", "input/output", "input")
                        Else
                            sDirection = "@RETURN_VALUE"
                        End If
                        
                        sName = oQry.GetColumnString(i, 1)
                        If Len(sName) = 0 Then sName = "@RetParam " & i + 1
                        
                        'Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%", "5@%"), Array(sName, oQry.GetColumnString(i, 2), oQry.GetColumnLong(i, 3), sDirection, oQry.GetColumnLong(i, 5)))
                        Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%", "4@%"), Array(sName, oQry.GetColumnString(i, 2), oQry.GetColumnLong(i, 3), sDirection))
                        Print #iOut, kCloseTB
                        
                        sSql = ReplaceMsg(kGetProperties, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(oFunc.Owner)), "'function'", SqlFinalS(ApiciSQL(oFunc.Name)), "'parameter'", SqlFinalS(ApiciSQL(sParameter))))
                        vVarExtProp = GetObjectProperties(oDb, sSql)
                        If IsArray(vVarExtProp) Then
                            sBuf = "<TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
                            sBuf = sBuf & "<TD width=25% bgcolor=#ffffcc> 1@% </TD><TD width=70% align=left bgcolor=#ffffcc> 2@% </TD><TD width=5% bgcolor=#ffffcc> 3@% </TD></THEAD><TBODY><TR bgColor=#f0fff0></TR>"
                                        
                            Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(m_sLngRes(enScanExtendedProperties), m_sLngRes(enScanValue), m_sLngRes(enScanExtendedPropertiesBaseType)))
                            For l = 0 To UBound(vVarExtProp, 2)
                                sBuf = "<TR><TD>1@%</TD><TD align=left>2@%</TD><TD align=left>3@%</TD></TR>"
                                Print #iOut, ReplaceMsg(sBuf, Array("1@%", "2@%", "3@%"), Array(vVarExtProp(0, l), vVarExtProp(1, l), vVarExtProp(2, l)))
                            Next
                            Print #iOut, kCloseTB
                        End If
                        vVarExtProp = vbEmpty
                        
                    Next
                    'Print #iOut, kCloseTB
                End If
            End If
                           
            Set oObjDep = New clsDepen
            lErr = oObjDep.ErrGetDependent(m_sDbName, .Name, funcDrop, oCollChild, oCollParent, sErr)
            Set oObjDep = Nothing
            If lErr = 0 Then
                lErr = ErrWriteDependent(iOut, .Name, True, oCollChild, sErr)
                lErr = ErrWriteDependent(iOut, .Name, False, oCollParent, sErr)
            End If
            
            
            sBuf = "<BR><TABLE border=1 cellPadding=2 cellSpacing=1 width=""100%""><THEAD><TR>"
            sBuf = sBuf & "<TD width=100% bgcolor=#ffffcc align=center> 1@% </TD>"
            sBuf = sBuf & "</THEAD><TBODY><TR bgColor=#f0fff0></TR>"
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(m_sLngRes(enScanScript)))
            
            'sType = GetScript(oSP, opt_SP)
            'sType = Replace(sType, vbCrLf, "<BR>")
            
            sBuf = "<TR><TD>1@%</TD></TR>"
            'Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array(sType))
            Print #iOut, ReplaceMsg(sBuf, Array("1@%"), Array("<PRE>" & GetScript(oFunc, opt_Func) & "</PRE>"))
            
            Print #iOut, kCloseTB
            
            Print #iOut, m_sProduced
            Print #iOut, "</BODY></HTML>"
            
        End With
        
    End If
    Set oObjDep = Nothing
    Set oCollChild = Nothing
    Set oCollParent = Nothing
    

    Close iOut
    ErrWriteUserFunc = lErr

End Function
Private Function GetScript(ByVal obj As Object, ByVal objType As optScript) As String
        
    Dim oDb As SQLDMO.Database2
    Dim oTb As SQLDMO.Table
    Dim oView As SQLDMO.View
    Dim oSP As SQLDMO.StoredProcedure
    Dim x As SQLDMO_SCRIPT_TYPE
    Dim y As SQLDMO_SCRIPT2_TYPE
    Dim lErr As Long, sErr As String
    
    Dim sOutPut As String
        
    Set oDb = objServer.Databases(m_sDbName)
    
    Dim objTrans As New SQLDMO.Transfer2
    
    x = SQLDMOScript_Default

'    objTrans.CopyAllDefaults = True
'    objTrans.CopyAllRules = True
'    objTrans.CopyAllUserDefinedDatatypes = True
    
    
    x = x Or SQLDMOScript_AppendToFile
    x = x Or SQLDMOScript_Drops
    objTrans.IncludeDependencies = False
    x = x Or SQLDMOScript_IncludeHeaders
    
    x = x Or SQLDMOScript_ObjectPermissions Or SQLDMOScript_DatabasePermissions Or SQLDMOScript_Permissions
    x = x Or SQLDMOScript_Indexes Or SQLDMOScript_ClusteredIndexes
    y = y Or SQLDMOScript2_FullTextCat Or SQLDMOScript2_FullTextIndex
    y = y Or SQLDMOScript2_ExtendedProperty
    x = x Or SQLDMOScript_Triggers
    x = x Or SQLDMOScript_DRI_All
    
'    objTrans.ScriptType = x
'    objTrans.Script2Type = Y
    If objType = opt_Tbl Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_UserTable, obj.Owner
    ElseIf objType = opt_View Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_View, obj.Owner
    ElseIf objType = opt_SP Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_StoredProcedure, obj.Owner
    ElseIf objType = opt_Def Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_Default, obj.Owner
    ElseIf objType = opt_Rule Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_Rule, obj.Owner
    ElseIf objType = opt_UDT Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_UserDefinedDatatype, obj.Owner
    ElseIf objType = opt_TRI Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_Trigger, obj.Owner
    ElseIf objType = opt_Func Then
        objTrans.AddObjectByName obj.Name, SQLDMOObj_UserDefinedFunction, obj.Owner
    End If

    objTrans.CopyData = SQLDMOCopyData_False
    objTrans.ScriptType = x
    objTrans.Script2Type = y
    
    On Local Error Resume Next
    sOutPut = oDb.ScriptTransfer(objTrans, SQLDMOXfrFile_SingleFile)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    If lErr <> 0 Then sOutPut = MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
    
        
    Set oTb = Nothing
    Set oView = Nothing
    Set oSP = Nothing
    Set oDb = Nothing
    Set objTrans = Nothing
    
    GetScript = sOutPut

End Function

Private Sub KillDir(ByVal sDir As String, ByRef lErr As Long, ByRef sErr As String, Optional ByVal bDirectory As Boolean = False)
    
    Dim sBuf As String
    
    On Local Error Resume Next
    If Not bDirectory Then
        sBuf = Dir$(sDir & "\*.*")
        While Len(sBuf) <> 0 And lErr = 0
            Kill sDir & "\" & sBuf
            lErr = Err.Number
            sErr = Err.Description
            sBuf = Dir$
        Wend
    Else
        RmDir sDir
        lErr = Err.Number
        sErr = Err.Description
        If lErr = 53 Or lErr = 76 Then
            lErr = 0
            sErr = ""
        End If
    End If
    On Local Error GoTo 0

End Sub
Private Sub LoadAll_RES_string()

    ReDim m_sLngRes(enScanValue) As String
    m_sLngRes(enScanHOME) = MyLoadResString(k_HtmHome)
    m_sLngRes(enScanName) = MyLoadResString(k_Name)
    m_sLngRes(enScanning) = MyLoadResString(k_HtmScanning)
    m_sLngRes(enScanTb) = MyLoadResString(k_HtmScanTB)
    m_sLngRes(enScanOwner) = MyLoadResString(k_Owner)
    m_sLngRes(enScanOwnership) = MyLoadResString(k_Ownership)
    m_sLngRes(enScanCreateDate) = MyLoadResString(k_Create_Date)
    m_sLngRes(enScanFileGroup) = MyLoadResString(k_HtmFileGroup)
    m_sLngRes(enScanRows) = MyLoadResString(k_Rows)
    m_sLngRes(enScanSysObj) = MyLoadResString(k_HtmSysObj)
    m_sLngRes(enScanStoredProc) = MyLoadResString(k_HtmScanSP)
    m_sLngRes(enScanViews) = MyLoadResString(k_HtmScanV)
    m_sLngRes(enScanUDT) = MyLoadResString(k_HtmScanUDT)
    m_sLngRes(enScanRULE) = MyLoadResString(k_HtmScanR)
    m_sLngRes(enScanDEF) = MyLoadResString(k_HtmScanDef)
    m_sLngRes(enScan_UdFunc) = MyLoadResString(k_HtmScanFunction)
    
    m_sLngRes(enScanDatabase) = MyLoadResString(k_RES_Database)
    m_sLngRes(enScanErr_OnlyOwner) = MyLoadResString(k_Only_DB_Owner_has_access_to_this_function)
    m_sLngRes(enScanLastBackup) = MyLoadResString(k_Last_BackUp)
    m_sLngRes(enScanUsers) = MyLoadResString(k_Users)
    m_sLngRes(enScanAvailableSpace) = MyLoadResString(k_Available_Space_Mb)
    m_sLngRes(enScanCompatLevel) = MyLoadResString(k_CompatLevel)
    m_sLngRes(enScanCollation) = MyLoadResString(k_NewDbSort)
    m_sLngRes(enScanExtendedProperties) = MyLoadResString(k_Extended_Property)
    m_sLngRes(enScanExtendedPropertiesBaseType) = MyLoadResString(k_Data_Type)
    
    
        
    m_sLngRes(enScanFileName) = MyLoadResString(k_HtmFileName)
    m_sLngRes(enScanPhysicalName) = MyLoadResString(k_HtmPhysicalName)
    m_sLngRes(enScanFileGrowth) = MyLoadResString(k_HtmFileGrowth)
    m_sLngRes(enScanMaxSize) = MyLoadResString(k_HtmFileMaxSize)
    m_sLngRes(enScanFileSizeKb) = MyLoadResString(k_HtmFileSizeKb)
    m_sLngRes(enScanFileGrowthType) = MyLoadResString(k_HtmFileGrowthType)
    m_sLngRes(enScanLogFile) = MyLoadResString(k_HtmLogFile)
    m_sLngRes(enScanTbDef) = MyLoadResString(k_HtmTBDef)
    m_sLngRes(enScanTb_inPK) = MyLoadResString(k_In_Primary_Key)
    m_sLngRes(enScanTb_FieldName) = MyLoadResString(k_Field_Name)
    m_sLngRes(enScanTb_Data_Type) = MyLoadResString(k_Data_Type)
    m_sLngRes(enScanTb_Size) = MyLoadResString(k_Size)
    m_sLngRes(enScanTb_AllowNull) = MyLoadResString(k_Allow_Null)
    m_sLngRes(enScanTb_DefVal) = MyLoadResString(k_Default_Value)
    m_sLngRes(enScanTb_Precision) = MyLoadResString(k_Precision)
    m_sLngRes(enScanTb_Scale) = MyLoadResString(k_Scale)
    m_sLngRes(enScanTb_IsRowGuid) = MyLoadResString(k_Is_RowGuid)
    m_sLngRes(enScanTb_IsIdentity) = MyLoadResString(k_Is_Identity)
    m_sLngRes(enScanTb_InitialVal) = MyLoadResString(k_Initial_Value)
    m_sLngRes(enScanTb_Increment) = MyLoadResString(k_Increment)
    m_sLngRes(enScanTb_Computed) = MyLoadResString(k_Computed_Field)
    m_sLngRes(enScanTb_Index_Key_Constr) = MyLoadResString(k_Indexes_Keys_constraints)
    m_sLngRes(enScanTb_FK) = MyLoadResString(k_HtmFK)
    m_sLngRes(enScanTb_FKtype) = MyLoadResString(k_HtmKType)
    m_sLngRes(enScanTb_FKclustered) = MyLoadResString(k_HtmKClustered)
    m_sLngRes(enScanTb_FKfillfactor) = MyLoadResString(k_HtmKFillFactor)
    m_sLngRes(enScanTb_FKcheck) = MyLoadResString(k_HtmKCheck)
    m_sLngRes(enScanTb_FKcolumns) = MyLoadResString(k_HtmKCcolumns)
    m_sLngRes(enScanTb_Htm4TB) = MyLoadResString(k_HtmForTB)
    m_sLngRes(enScanTb_Checkconstr) = MyLoadResString(k_Check_Constraints)
    m_sLngRes(enScanTb_CheckText) = MyLoadResString(k_HtmChkText)
    m_sLngRes(enScanScript) = MyLoadResString(k_HtmScript)
    m_sLngRes(enScanColDef) = MyLoadResString(k_HtmColDef)
    m_sLngRes(enScanColComputedText) = MyLoadResString(k_Computed_Text)
    m_sLngRes(enScanColList) = MyLoadResString(k_HtmColList)
    m_sLngRes(enScanColPriTB) = MyLoadResString(k_HtmPriTB)
    m_sLngRes(enScanCol_PK) = MyLoadResString(k_HtmPK)
    m_sLngRes(enScanCol_UK) = MyLoadResString(k_HtmUK)
    m_sLngRes(enScan_DedendON) = MyLoadResString(k_Objects_obj_THAT_Depends_on)
    m_sLngRes(enScan_DedendFROM) = MyLoadResString(k_Objects_THAT_obj_Depends_on)
    m_sLngRes(enScan_Object) = MyLoadResString(k_Object)
    m_sLngRes(enScan_Type) = MyLoadResString(k_Type)
    m_sLngRes(enScan_Table) = MyLoadResString(k_Table)
    m_sLngRes(enScan_View) = MyLoadResString(k_View)
    m_sLngRes(enScan_StoredP) = MyLoadResString(k_Stored_Procedure)
    m_sLngRes(enScan_UDT) = MyLoadResString(k_User_Defined_Data_Type)
    m_sLngRes(enScan_Parameter) = MyLoadResString(k_HtmSpParameter)
        
    m_sLngRes(enScan_ViewColumn) = MyLoadResString(k_HtmViewColumns)
    m_sLngRes(enScan_Lenght) = MyLoadResString(k_Lenght)
    m_sLngRes(enScan_BoundCols) = MyLoadResString(k_HtmBoundColumns)
    m_sLngRes(enScan_Text) = MyLoadResString(k_HtmText)
    m_sLngRes(enScan_BoundUDT) = MyLoadResString(k_HtmBoundUDT)
    m_sLngRes(enScan_Index) = MyLoadResString(k_Indexes)
    m_sLngRes(enScanProperty) = MyLoadResString(k_Property)
    m_sLngRes(enScanValue) = MyLoadResString(k_Value)
            
End Sub
'Private Sub Write2Log(ByVal sMsg As String, Optional ByVal bCreate As Boolean = False)
'
'    Dim lErr As Long
'    Dim iFile As Integer
'
'    iFile = FreeFile
'    On Local Error Resume Next
'    If bCreate Then
'        Open App.Path & "\exp.log" For Output As iFile
'    Else
'        Open App.Path & "\exp.log" For Append As iFile
'    End If
'    lErr = Err.Number
'    On Local Error GoTo 0
'    If lErr <> 0 Then iFile = 0
'
'    On Local Error Resume Next
'    If iFile <> 0 And bCreate Then
'        Print #iFile, "File created on: " & NomeComputer() & " at: " & Now
'    End If
'
'    If iFile <> 0 Then Print #iFile, sMsg
'    On Local Error GoTo 0
'
'    Close #iFile
'End Sub
Private Function GetObjectProperties(ByVal oDb As SQLDMO.Database2, ByVal sSql As String) As Variant

    Dim oQry As SQLDMO.QueryResults
    Dim lErr As Long, sErr As String
    Dim vVar As Variant
    
    On Local Error Resume Next
    Set oQry = oDb.ExecuteWithResults(sSql, Len(sSql))
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If Not oQry Is Nothing And lErr = 0 Then
        vVar = Qry2Var(oQry)
    Else
        ReDim vVar(2, 0) As Variant
        vVar(0, 0) = MyLoadResString(kMsgBoxError) & ": " & lErr
        vVar(1, 0) = sErr
        vVar(2, 0) = ""
    End If
    Set oQry = Nothing
    Set oDb = Nothing
    GetObjectProperties = vVar
        
End Function
