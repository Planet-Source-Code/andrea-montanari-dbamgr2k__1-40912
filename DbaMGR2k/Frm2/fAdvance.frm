VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fAdvance 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6885
   ClipControls    =   0   'False
   Icon            =   "fAdvance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
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
      Left            =   5160
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   3840
      Top             =   0
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
            Picture         =   "fAdvance.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdvance.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdvance.frx":05BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdvance.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdvance.frx":0874
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Param"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Security"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Connections"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Server Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Client Components"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
End
Attribute VB_Name = "fAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const kCON_ConstraintCheck = 1
Const kCON_ImplicitTrans = 2
Const kCON_CloseCursors = 4
Const kCON_ANSI_warn = 8
Const kCON_ANSI_pad = 16
Const kCON_ANSI_nulls = 32
Const kCON_Arit_aborth = 64
Const kCON_Arit_ignore = 128
Const kCON_quoted_ident = 256
Const kCON_NO_count = 512
Const kCON_NULLS_definedON = 1024
Const kCON_NULLS_definedOFF = 2048

Private m_blnAnnullaMod As Boolean
Private bLoading As Boolean
Private lngWidth As Long
Private lngHeight As Long
Private miRow As Integer
Private miFlex As Integer
Private miTmpType As iPropG
Private mlngUserOptions As Long
Private Enum cboLoad
    xcb_none = 0
    x_Security
    x_Audit
    x_Language
End Enum

Private Enum xBoundary
    xb_None = 0
    xb_Min = 1
    xb_Max = 2
End Enum


Private Type udtSTR
    Elem As elemFlexG
    EditType As TypeEnum
    bDisab As Boolean
    vOrigValue As Variant
    iBoundary As xBoundary
    vMaxVal As Variant
    vMinVal As Variant
    sConfigName As String
End Type

Private xEnumSecurity() As udtLANG
Private xEnumAudit() As udtLANG
Private xEnumLanguage() As udtLANG

Private xGeneral() As udtSTR
Private xSecurity() As udtSTR
Private xConnections() As udtSTR
Private xServerSet() As udtSTR
Private xClient() As udtSTR


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
        
    Dim iType As cboLoad
    If Not m_blnAnnullaMod Then
        Select Case miFlex
            Case 1
                If miRow = 1 Then
                    iType = x_Security
                ElseIf miRow = 2 Then
                    iType = x_Audit
                End If
                If iType <> xcb_none Then
                    xSecurity(miRow - 1).Elem.iLen = cbo.ItemData(cbo.ListIndex)
                    Flex(miFlex).TextMatrix(miRow, 2) = GetCBODescriptionByID(iType, xSecurity(miRow - 1).Elem.iLen)
                End If
            Case 2
                'xConnections()
            Case 3
                'xServerSet
                If miRow = 1 Then
                    iType = x_Language
                End If
                If iType <> xcb_none Then
                    xServerSet(miRow - 1).Elem.iLen = cbo.ItemData(cbo.ListIndex)
                    Flex(miFlex).TextMatrix(miRow, 2) = GetCBODescriptionByID(iType, xServerSet(miRow - 1).Elem.iLen)
                End If
        
        End Select
    End If
    cbo.Visible = False
    Flex(miFlex).Refresh
    fra(miFlex).Refresh
    m_blnAnnullaMod = True
    
End Sub

Private Sub Flex_DblClick(Index As Integer)

    Dim bPush As Boolean
    Dim bCombo As Boolean
    Dim bEllipsis As Boolean
    
    If Flex(Index).Col = Flex(Index).Cols - 1 Then ' And Flex(index).MouseRow > 0 And Flex(index).MouseCol > 1 Then
        Select Case miFlex
            Case 0
                bPush = xGeneral(miRow - 1).EditType = iPush
                bCombo = xGeneral(miRow - 1).EditType = iCombo
                bEllipsis = xGeneral(miRow - 1).EditType = iEllipsis
            Case 1
                bPush = xSecurity(miRow - 1).EditType = iPush
                bCombo = xSecurity(miRow - 1).EditType = iCombo
                bEllipsis = xSecurity(miRow - 1).EditType = iEllipsis
            Case 2
                bPush = xConnections(miRow - 1).EditType = iPush
                bCombo = xConnections(miRow - 1).EditType = iCombo
                bEllipsis = xConnections(miRow - 1).EditType = iEllipsis
            Case 3
                bPush = xServerSet(miRow - 1).EditType = iPush
                bCombo = xServerSet(miRow - 1).EditType = iCombo
                bEllipsis = xServerSet(miRow - 1).EditType = iEllipsis
        End Select
        
        If bCombo Then
            FlexEdit 0, Index
        ElseIf bEllipsis Then
            FlexEdit 0, Index
        Else
            FlexEdit 0, Index
        End If
    End If
    
End Sub

Private Sub Flex_GotFocus(Index As Integer)
    
    TestActiveControl
    cbo.Visible = False
    tFlex.Visible = False
    m_blnAnnullaMod = True

End Sub

Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    
    iCode = KeyCode
    KeyCode = 0
    If Index <> 0 Then
        If iCode = vbKeyF2 Then FlexEdit 0, Index
    End If

End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    
    iCode = KeyAscii
    KeyAscii = 0
    If miFlex <> 0 Then FlexEdit iCode, Index

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
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex(Index).CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1

                    .Col = iCol
                    miRow = iRowP
                
                End If
                
                .ToolTipText = IIf(Len(.TextMatrix(.Row, .Col)) <> 0, .TextMatrix(.Row, .Col), "")
            End With
            b = False
        End If
    End If

End Sub

Private Sub Flex_Scroll(Index As Integer)

    cbo.Visible = False
    tFlex.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo
    'spostamento sulla scheda successiva.
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tStrip.SelectedItem.Index
        If i = tStrip.Tabs.Count Then
            'È stata raggiunta l'ultima scheda e quindi
            'torna alla scheda 1.
            Set tStrip.SelectedItem = tStrip.Tabs(1)
        Else
            'Incrementa l'indice della scheda
            Set tStrip.SelectedItem = tStrip.Tabs(i + 1)
        End If
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer, iCount As Integer
    
    lngWidth = Me.Width
    lngHeight = Me.Height
    For i = 0 To 4
        Flex(i).RowHeightMin = cbo.Height
    Next
        
    For i = 1 To objServer.Languages.Count
        ReDim Preserve xEnumLanguage(iCount) As udtLANG
        xEnumLanguage(iCount).sLang = objServer.Languages(i).Alias
        xEnumLanguage(iCount).idLang = objServer.Languages(i).Id
        iCount = iCount + 1
    Next

    Me.Caption = ReplaceMsg(MyLoadResString(k_SqlProperty), Array("1%"), Array(objServer.NetName))
    tStrip.Tabs(1).Caption = MyLoadResString(k_GeneralTab)
    tStrip.Tabs(2).Caption = MyLoadResString(k_SecurityTab)
    tStrip.Tabs(3).Caption = MyLoadResString(k_ConnectionTab)
    tStrip.Tabs(4).Caption = MyLoadResString(k_ServerSettingsTab)
    tStrip.Tabs(5).Caption = MyLoadResString(k_DMOTab)
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keyRun).ToolTipText = MyLoadResString(k_ApplyChanges)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_Fit_Grid)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_Refresh)
    tBar.Buttons(keyParam).ToolTipText = MyLoadResString(k_Param)
    
    CenterForm Me
    
    Screen.MousePointer = vbHourglass
    For i = 0 To 4
        RefreshData i
        LoadFlex i
    Next
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Me.WindowState <> vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < lngWidth Then Me.Width = lngWidth
            If Me.Height < lngHeight Then Me.Height = lngHeight
            SizeControls
            b = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fAdvance = Nothing
End Sub
Private Sub SizeControls()

    Dim i As Integer
    tStrip.Width = Me.ScaleWidth - (2 * tStrip.Left)
    tStrip.Height = Me.ScaleHeight - tStrip.Top - 100
    fra(0).Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
    
    For i = 0 To 4
        fra(i).Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
        Flex(i).Width = fra(0).Width
        Flex(i).Height = fra(0).Height
    Next
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    TestActiveControl
    If Button.Enabled Then
        Select Case Button.Key
            Case keyRun
                Flex(tStrip.SelectedItem.Index - 1).SetFocus
                SaveData
            Case keyEnd
                Unload Me
            Case keyFitGrid
                LoadFlex (tStrip.SelectedItem.Index - 1)
            Case keyRefresh
                RefreshData tStrip.SelectedItem.Index - 1
                LoadFlex tStrip.SelectedItem.Index - 1
                Flex(tStrip.SelectedItem.Index - 1).SetFocus
            Case keyParam
                fParam.Show vbModal, Me
        End Select
    End If
    
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iLen As Integer
    If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then
        KeyCode = 0
    Else
        If (miTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (miTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        ElseIf (miTmpType And am_Date) = am_Date Then
            If miFlex = 1 Then
                iLen = xSecurity(miRow - 1).Elem.iLen
            ElseIf miFlex = 2 Then
                iLen = xConnections(miRow - 1).Elem.iLen
            ElseIf miFlex = 3 Then
                iLen = xServerSet(miRow - 1).Elem.iLen
            End If
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
            Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (miTmpType And am_String) = am_String Then
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)
    
    Dim iLen As Integer
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93  '<>|[]
            KeyAscii = 0
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
            ElseIf (miTmpType And am_Date) = am_Date Then
                If miFlex = 1 Then
                    iLen = xSecurity(miRow - 1).Elem.iLen
                ElseIf miFlex = 2 Then
                    iLen = xConnections(miRow - 1).Elem.iLen
                ElseIf miFlex = 3 Then
                    iLen = xServerSet(miRow - 1).Elem.iLen
                End If
               'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
                Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (miTmpType And am_String) = am_String Then

            End If
        Case vbKeySpace
'            If Not bAcceptSpace Then KeyAscii = 0
        Case Else
            If (miTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (miTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            ElseIf (miTmpType And am_Date) = am_Date Then
                If miFlex = 1 Then
                    iLen = xSecurity(miRow - 1).Elem.iLen
                ElseIf miFlex = 2 Then
                    iLen = xConnections(miRow - 1).Elem.iLen
                ElseIf miFlex = 3 Then
                    iLen = xServerSet(miRow - 1).Elem.iLen
                End If
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen
                Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (miTmpType And am_String) = am_String Then
                Select Case KeyAscii
                    Case 33, 34, 39 To 47, 61, 63
                        If Flex(miFlex).Col = 2 Then KeyAscii = 0
                End Select
                OverStrike tFlex
            End If
    End Select

End Sub

Private Sub tFlex_LostFocus()

    Dim bUpdate As Boolean
    Dim vVal As Variant
                
    vVal = ValF(tFlex.Text)
    Select Case miFlex
        Case 2
            'xConnections()
            If Len(vVal) Then
                bUpdate = True
                If xConnections(miRow - 1).iBoundary <> xb_None Then
                    If (xConnections(miRow - 1).iBoundary And xb_Min) = xb_Min Then
                        bUpdate = Val(vVal) >= xConnections(miRow - 1).vMinVal
                    End If
                    If bUpdate Then
                        If (xConnections(miRow - 1).iBoundary And xb_Max) = xb_Max Then
                            bUpdate = Val(vVal) <= xConnections(miRow - 1).vMaxVal
                        End If
                    End If
                End If
                If bUpdate Then
                    If xConnections(miRow - 1).Elem.iType = am_Long Then
                        xConnections(miRow - 1).Elem.vVal = Val(vVal)
                    ElseIf xConnections(miRow - 1).Elem.iType = am_Date Then
                        xConnections(miRow - 1).Elem.vVal = Val(vVal)
                    Else
                        xConnections(miRow - 1).Elem.vVal = CStr(vVal)
                    End If
                End If
            End If
        Case 3
            'xServerSet
            If Len(vVal) Then
                bUpdate = True
                If xServerSet(miRow - 1).iBoundary <> xb_None Then
                    If (xServerSet(miRow - 1).iBoundary And xb_Min) = xb_Min Then
                        bUpdate = Val(vVal) >= xServerSet(miRow - 1).vMinVal
                    End If
                    If bUpdate Then
                        If (xServerSet(miRow - 1).iBoundary And xb_Max) = xb_Max Then
                            bUpdate = Val(vVal) <= xServerSet(miRow - 1).vMaxVal
                        End If
                    End If
                End If
                
                If bUpdate Then
                    If xServerSet(miRow - 1).Elem.iType = am_Long Then
                        xServerSet(miRow - 1).Elem.vVal = Val(vVal)
                    ElseIf xServerSet(miRow - 1).Elem.iType = am_Date Then
                        xServerSet(miRow - 1).Elem.vVal = Val(vVal)
                    Else
                        xServerSet(miRow - 1).Elem.vVal = CStr(vVal)
                    End If
                    If miRow = Flex(miFlex).Rows - 1 Then
                        xServerSet(miRow - 1).Elem.sField = ReplaceMsg(MyLoadResString(k_2yerCutoff), Array("1%"), Array(Format$(xServerSet(miRow - 1).Elem.vVal - 99)))
                        Flex(miFlex).TextMatrix(miRow, 1) = xServerSet(miRow - 1).Elem.sField
                    End If
                End If
            End If
    
    End Select
    
    If bUpdate Then
        Flex(miFlex).TextMatrix(miRow, 2) = tFlex.Text
    Else
        Beep
    End If

    tFlex.Visible = False
    Flex(miFlex).Refresh
    fra(miFlex).Refresh

End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    Dim x As Integer
    TestActiveControl
         
    i = tStrip.SelectedItem.Index - 1
    
    For x = 0 To 3
        fra(x).Visible = x = i
    Next
    fra(i).ZOrder
    Flex(i).SetFocus
    Flex_RowColChange i
    fra(i).Refresh

End Sub
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

Private Sub RefreshData(ByVal iItem As Integer)

    Dim i As Integer, iOffset As Integer
    Dim lngUser As Long
    Dim oQryEnum As SQLDMO.QueryResults
    Dim oApp As SQLDMO.Application
    Dim sBuf As String
    Dim sConnString As String, sDefDb As String
    Dim lErr As Long
    
        
    Screen.MousePointer = vbHourglass
    
    
'    Dim cv As SQLDMO.ConfigValue
'    With objServer.Configuration
'        .ShowAdvancedOptions = True
'        For Each cv In .ConfigValues
'            Debug.Print cv.Name, cv.RunningValue, cv.CurrentValue
'        Next
'
'    End With
    
        
    objServer.Configuration.ShowAdvancedOptions = True
    Select Case iItem
        Case 0
            Set oQryEnum = objServer.EnumVersionInfo
            iOffset = 10
            i = oQryEnum.Rows + iOffset
            ReDim xGeneral(i) As udtSTR
            
            xGeneral(0).EditType = iNormalEdit
            xGeneral(0).bDisab = True
            xGeneral(0).Elem.sField = "Name"
            xGeneral(0).Elem.vVal = objServer.NetName
            
            xGeneral(1).EditType = iNormalEdit
            xGeneral(1).bDisab = True
            xGeneral(1).Elem.sField = "Instance Name"
            xGeneral(1).Elem.vVal = objServer.InstanceName
            
            Debug.Print objServer.IsClustered
            xGeneral(2).bDisab = True
            xGeneral(2).Elem.sField = "Package Version"
            xGeneral(2).Elem.vVal = GetPackage()
            
            xGeneral(3).bDisab = True
            xGeneral(3).Elem.sField = "Code Page"
            xGeneral(3).Elem.vVal = objServer.CodePage
            
            xGeneral(4).bDisab = True
            xGeneral(4).Elem.sField = "Collation"
            xGeneral(4).Elem.vVal = objServer.Collation
            
            
            xGeneral(5).EditType = iNormalEdit
            xGeneral(5).bDisab = True
            xGeneral(5).Elem.sField = "Is Cluster"
            xGeneral(5).Elem.vVal = objServer.IsClustered
            xGeneral(5).Elem.iType = am_Bool
            
            xGeneral(6).EditType = iNormalEdit
            xGeneral(6).bDisab = True
            xGeneral(6).Elem.sField = "Regional Setting"
            xGeneral(6).Elem.vVal = objServer.RegionalSetting
            xGeneral(6).Elem.iType = am_Bool
            
            xGeneral(7).EditType = iNormalEdit
            xGeneral(7).bDisab = True
            xGeneral(7).Elem.sField = "Character Set"
            xGeneral(7).Elem.vVal = objServer.Registry.CharacterSet
            
            xGeneral(8).EditType = iNormalEdit
            xGeneral(8).bDisab = True
            xGeneral(8).Elem.sField = "Sort Order"
            xGeneral(8).Elem.vVal = objServer.Registry.SortOrder
            
            xGeneral(9).EditType = iNormalEdit
            xGeneral(9).bDisab = True
            xGeneral(9).Elem.sField = "Error Log Path"
            xGeneral(9).Elem.vVal = objServer.Registry.ErrorLogPath
            
            xGeneral(10).EditType = iNormalEdit
            xGeneral(10).bDisab = True
            xGeneral(10).Elem.sField = "Server Root"
            xGeneral(10).Elem.vVal = objServer.Registry.SQLRootPath
            
                        
            For i = 1 To oQryEnum.Rows
                xGeneral(i + iOffset).bDisab = True
                xGeneral(i + iOffset).Elem.sField = oQryEnum.GetColumnString(i, 2)
                xGeneral(i + iOffset).Elem.vVal = oQryEnum.GetColumnString(i, 4)
            Next
            Set oQryEnum = Nothing
                        
        
        Case 1
            ReDim xEnumSecurity(2) As udtLANG
            xEnumSecurity(0).idLang = SQLDMOSecurity_Mixed
            xEnumSecurity(0).sLang = MyLoadResString(k_MixedSecurity)
            xEnumSecurity(1).idLang = SQLDMOSecurity_Normal
            xEnumSecurity(1).sLang = MyLoadResString(k_SqlSecurity)
            xEnumSecurity(2).idLang = SQLDMOSecurity_Integrated
            xEnumSecurity(2).sLang = MyLoadResString(k_WinSecurity)
        
            ReDim xEnumAudit(3) As udtLANG
            xEnumAudit(0).idLang = SQLDMOAudit_None
            xEnumAudit(0).sLang = MyLoadResString(k_Audit_None)
            xEnumAudit(1).idLang = SQLDMOAudit_Success
            xEnumAudit(1).sLang = MyLoadResString(k_Audit_Success)
            xEnumAudit(2).idLang = SQLDMOAudit_Failure
            xEnumAudit(2).sLang = MyLoadResString(k_Audit_Failure)
            xEnumAudit(3).idLang = SQLDMOAudit_All
            xEnumAudit(3).sLang = MyLoadResString(k_Audit_All)
        
        
            i = 4
            ReDim xSecurity(i) As udtSTR
            
            xSecurity(0).EditType = iCombo
            xSecurity(0).Elem.iLen = objServer.IntegratedSecurity.SecurityMode
            xSecurity(0).Elem.sField = MyLoadResString(k_Security_Mode)
            xSecurity(0).Elem.vVal = GetCBODescriptionByID(x_Security, xSecurity(0).Elem.iLen)
            xSecurity(0).vOrigValue = xSecurity(0).Elem.iLen
            xSecurity(0).Elem.iType = am_Long
            xSecurity(0).sConfigName = ""
            
            
            xSecurity(1).EditType = iCombo
            xSecurity(1).Elem.iLen = objServer.IntegratedSecurity.AuditLevel
            xSecurity(1).Elem.vVal = GetCBODescriptionByID(x_Audit, xSecurity(1).Elem.iLen)
            xSecurity(1).Elem.sField = MyLoadResString(k_Audit_Level)
            xSecurity(1).vOrigValue = xSecurity(1).Elem.iLen
            xSecurity(1).Elem.iType = am_Long
            xSecurity(1).sConfigName = ""
                                   
            xSecurity(2).bDisab = True
            xSecurity(2).EditType = iNormalEdit
            xSecurity(2).Elem.sField = "Start SqlServer Using Specific Account"
            xSecurity(2).Elem.vVal = IIf(Len(objServer.StartupAccount) = 0, True, False)
            xSecurity(2).Elem.iType = am_Bool
            xSecurity(2).vOrigValue = IIf(Len(objServer.StartupAccount) = 0, True, False)
            
            
            xSecurity(3).EditType = iNormalEdit
            xSecurity(3).Elem.sField = "Specific Account"
            xSecurity(3).Elem.vVal = objServer.StartupAccount
            xSecurity(3).Elem.iType = am_String
            xSecurity(3).Elem.iLen = 50
            xSecurity(3).bDisab = True
            xSecurity(3).vOrigValue = objServer.StartupAccount
            
            
            xSecurity(4).EditType = iEllipsis
            xSecurity(4).Elem.sField = "Default Data Root"
            xSecurity(4).Elem.vVal = objServer.Registry.SQLDataRoot
            xSecurity(4).Elem.iType = am_String
            xSecurity(4).Elem.iLen = 256
            xSecurity(4).vOrigValue = objServer.Registry.SQLDataRoot
                        
                       
        Case 2
            lngUser = objServer.Configuration.ConfigValues("user options").RunningValue
            mlngUserOptions = lngUser
        
            i = 12
                        
            ReDim xConnections(i) As udtSTR
            
            xConnections(0).sConfigName = "user connections"
            xConnections(0).Elem.iLen = 6
            xConnections(0).Elem.iType = am_Long
            xConnections(0).Elem.vVal = objServer.Configuration.ConfigValues(xConnections(0).sConfigName).RunningValue
            xConnections(0).Elem.sField = MyLoadResString(kMaxConcurrentUser)
            xConnections(0).vOrigValue = xConnections(0).Elem.vVal
            xConnections(0).iBoundary = xb_Max
            xConnections(0).vMaxVal = objServer.Configuration.ConfigValues(xConnections(0).sConfigName).MaximumValue
                        
            'bit mask
            xConnections(1).Elem.iType = am_Bool
            xConnections(1).Elem.vVal = (lngUser And kCON_ConstraintCheck) = kCON_ConstraintCheck
            xConnections(1).Elem.sField = MyLoadResString(kConstraintCheck)
            xConnections(1).vOrigValue = xConnections(1).Elem.vVal
                    
            xConnections(2).Elem.iType = am_Bool
            xConnections(2).Elem.vVal = (lngUser And kCON_ImplicitTrans) = kCON_ImplicitTrans
            xConnections(2).Elem.sField = MyLoadResString(kImplicitTrans)
            xConnections(2).vOrigValue = xConnections(2).Elem.vVal
                    
            xConnections(3).Elem.iType = am_Bool
            xConnections(3).Elem.vVal = (lngUser And kCON_CloseCursors) = kCON_CloseCursors
            xConnections(3).Elem.sField = MyLoadResString(kCloseCursors)
            xConnections(3).vOrigValue = xConnections(3).Elem.vVal
                    
            xConnections(4).Elem.iType = am_Bool
            xConnections(4).Elem.vVal = (lngUser And kCON_ANSI_warn) = kCON_ANSI_warn
            xConnections(4).Elem.sField = MyLoadResString(kANSI_warn)
            xConnections(4).vOrigValue = xConnections(4).Elem.vVal
            
            xConnections(5).Elem.iType = am_Bool
            xConnections(5).Elem.vVal = (lngUser And kCON_ANSI_pad) = kCON_ANSI_pad
            xConnections(5).Elem.sField = MyLoadResString(kANSI_pad)
            xConnections(5).vOrigValue = xConnections(5).Elem.vVal
        
            xConnections(6).Elem.iType = am_Bool
            xConnections(6).Elem.vVal = (lngUser And kCON_ANSI_nulls) = kCON_ANSI_nulls
            xConnections(6).Elem.sField = MyLoadResString(kANSI_nulls)
            xConnections(6).vOrigValue = xConnections(6).Elem.vVal
            
            xConnections(7).Elem.iType = am_Bool
            xConnections(7).Elem.vVal = (lngUser And kCON_Arit_aborth) = kCON_Arit_aborth
            xConnections(7).Elem.sField = MyLoadResString(kArit_aborth)
            xConnections(7).vOrigValue = xConnections(7).Elem.vVal
            
            xConnections(8).Elem.iType = am_Bool
            xConnections(8).Elem.vVal = (lngUser And kCON_Arit_ignore) = kCON_Arit_ignore
            xConnections(8).Elem.sField = MyLoadResString(kArit_ignore)
            xConnections(8).vOrigValue = xConnections(8).Elem.vVal
            
            xConnections(9).Elem.iType = am_Bool
            xConnections(9).Elem.vVal = (lngUser And kCON_quoted_ident) = kCON_quoted_ident
            xConnections(9).Elem.sField = MyLoadResString(kQuoted_ident)
            xConnections(9).vOrigValue = xConnections(9).Elem.vVal
            
            xConnections(10).Elem.iType = am_Bool
            xConnections(10).Elem.vVal = (lngUser And kCON_NO_count) = kCON_NO_count
            xConnections(10).Elem.sField = MyLoadResString(kNO_count)
            xConnections(10).vOrigValue = xConnections(10).Elem.vVal
            
            xConnections(11).Elem.iType = am_Bool
            xConnections(11).Elem.vVal = (lngUser And kCON_NULLS_definedON) = kCON_NULLS_definedON
            xConnections(11).Elem.sField = MyLoadResString(kNULLS_definedON)
            xConnections(11).vOrigValue = xConnections(11).Elem.vVal
            
            xConnections(12).Elem.iType = am_Bool
            xConnections(12).Elem.vVal = (lngUser And kCON_NULLS_definedOFF) = kCON_NULLS_definedOFF
            xConnections(12).Elem.sField = MyLoadResString(kNULLS_definedOFF)
            xConnections(12).vOrigValue = xConnections(12).Elem.vVal
                       
            For i = 1 To 12
                xConnections(i).sConfigName = "user options"
            Next
            
        Case 3
           'Private xServerSet() As udtSTR
            i = 3
            
            ReDim xServerSet(i) As udtSTR
            
            xServerSet(0).sConfigName = "default language"
            xServerSet(0).EditType = iCombo
            xServerSet(0).Elem.iLen = objServer.Configuration(xServerSet(0).sConfigName).CurrentValue
            xServerSet(0).Elem.vVal = GetCBODescriptionByID(x_Language, xServerSet(0).Elem.iLen)
            xServerSet(0).Elem.sField = MyLoadResString(k_ServerLanguage)
            xServerSet(0).vOrigValue = xServerSet(0).Elem.iLen
            xServerSet(0).Elem.iType = am_Long
                                    
            xServerSet(1).sConfigName = "allow updates"
            xServerSet(1).Elem.vVal = objServer.Configuration(xServerSet(1).sConfigName).CurrentValue
            xServerSet(1).Elem.sField = MyLoadResString(k_AllowChanges)
            xServerSet(1).vOrigValue = xServerSet(1).Elem.vVal
            xServerSet(1).Elem.iType = am_Bool
                        
            xServerSet(2).sConfigName = "nested triggers"
            xServerSet(2).Elem.vVal = objServer.Configuration(xServerSet(2).sConfigName).CurrentValue
            xServerSet(2).Elem.sField = MyLoadResString(k_NestedTriggers)
            xServerSet(2).vOrigValue = xServerSet(2).Elem.vVal
            xServerSet(2).Elem.iType = am_Bool
            
            xServerSet(3).sConfigName = "two digit year cutoff"
            xServerSet(3).Elem.vVal = objServer.Configuration(xServerSet(3).sConfigName).CurrentValue
            xServerSet(3).Elem.sField = ReplaceMsg(MyLoadResString(k_2yerCutoff), Array("1%"), Array(Format$(xServerSet(3).Elem.vVal - 99)))
            xServerSet(3).vOrigValue = xServerSet(3).Elem.vVal
            xServerSet(3).Elem.iType = am_Date
            xServerSet(3).Elem.iLen = 4
            xServerSet(3).iBoundary = xb_Min Or xb_Max
            xServerSet(3).vMinVal = objServer.Configuration.ConfigValues(xServerSet(3).sConfigName).MinimumValue
            xServerSet(3).vMaxVal = objServer.Configuration.ConfigValues(xServerSet(3).sConfigName).MaximumValue
        Case 4
            Set oApp = New SQLDMO.Application
            i = 8
        
            ReDim xClient(i) As udtSTR
            xClient(0).bDisab = True
            xClient(0).Elem.sField = MyLoadResString(kODBCVersionString)
            xClient(0).Elem.vVal = oApp.ODBCVersionString
                                    
            xClient(1).bDisab = True
            xClient(1).Elem.sField = MyLoadResString(kGroupRegistrationServer)
            xClient(1).Elem.vVal = oApp.GroupRegistrationServer
                                    
            xClient(2).bDisab = True
            xClient(2).Elem.sField = MyLoadResString(kBlockingTimeout)
            xClient(2).Elem.vVal = oApp.BlockingTimeout
                                    
            xClient(3).bDisab = True
            xClient(3).Elem.sField = MyLoadResString(kFullName)
            xClient(3).Elem.vVal = oApp.FullName
                                    
            xClient(4).bDisab = True
            xClient(4).Elem.sField = MyLoadResString(kVersion)
            xClient(4).Elem.vVal = oApp.VersionMajor & "." & oApp.VersionMinor & "." & oApp.VersionBuild
            
            Dim FI As cFileVer
            Set FI = New cFileVer
            FI.FullPathName() = oApp.FullName
            
            xClient(5).bDisab = True
            xClient(5).Elem.sField = MyLoadResString(kFileVersion)
            xClient(5).Elem.vVal = FI.FileVersion
            
            sBuf = FI.ProductVersion
            xClient(6).bDisab = True
            xClient(6).Elem.sField = MyLoadResString(kProductVersion)
            xClient(6).Elem.vVal = sBuf
            
            xClient(7).bDisab = True
            xClient(7).Elem.sField = MyLoadResString(kServicePack)
            If InStr(sBuf, "8.00.534") Then
                sBuf = " Service Pack 2"
            ElseIf InStr(sBuf, "8.00.384") Then
                sBuf = " Service Pack 1"
            ElseIf InStr(sBuf, "8.00.194") Then
                sBuf = " Retail"
            Else
                sBuf = ""
            End If
            
            xClient(7).Elem.vVal = sBuf
            
            Set FI = Nothing
            
            Dim adoConn As ADODB.Connection
            Dim objErr As New cErrMgr
            
            xClient(8).bDisab = True
            xClient(8).Elem.sField = MyLoadResString(kMdacVersion)
            xClient(8).Elem.vVal = MyLoadResString(kNotAvailable)
            
            sDefDb = objServer.Logins(objServer.Login).Database
                        
            
            If CBool(Val(GetIni("CONNECT", "TRUSTED", App.Path & "\" & App.EXEName & ".INI"))) Then
                'trusted connection
                sConnString = "Initial Catalog=" & sDefDb & "; Integrated Security = SSPI;"
            Else
                'Standard connection
                sConnString = "Initial Catalog=" & sDefDb & ";User Id= " & objServer.Login & ";" & IIf(Len(objServer.Password), " Password= " & objServer.Password & ";", "")
            End If
            
            lErr = ErrMyOpenConn(adoConn, sConnString, objErr, adUseClient)
            If lErr = 0 And adoConn.State = adStateOpen Then xClient(8).Elem.vVal = adoConn.Version
            If lErr <> 0 Then Call ExceptionHandler(objErr, "", True)
            Set adoConn = Nothing
            Set objErr = Nothing
            
    End Select
    Set oApp = Nothing
    objServer.Configuration.ShowAdvancedOptions = False
    Screen.MousePointer = vbDefault

End Sub
Private Function GetCBODescriptionByID(ByVal iType As cboLoad, ByVal lId As Long) As String

    Dim i As Integer
    If iType = x_Security Then
        For i = 0 To UBound(xEnumSecurity)
            If xEnumSecurity(i).idLang = lId Then
                GetCBODescriptionByID = xEnumSecurity(i).sLang
                Exit For
            End If
        Next
    ElseIf iType = x_Audit Then
        For i = 0 To UBound(xEnumAudit)
            If xEnumAudit(i).idLang = lId Then
                GetCBODescriptionByID = xEnumAudit(i).sLang
                Exit For
            End If
        Next
    ElseIf iType = x_Language Then
        For i = 0 To UBound(xEnumLanguage)
            If xEnumLanguage(i).idLang = lId Then
                GetCBODescriptionByID = xEnumLanguage(i).sLang
                Exit For
            End If
        Next
    
    End If
    
End Function

Private Sub LoadFlex(ByVal iItem As Integer)

    Dim sBuf As String
    Dim i As Integer
    Dim iX As Integer
    Dim bDisab As Boolean, iMaxElem As Integer
    Dim bBoolean As Boolean
    Dim bPush As Boolean, bCombo As Boolean, bEllipsis As Boolean
    Dim idBitmap As Integer
    Dim iCol As Integer, iCol2Enlarge As Integer
       
    With Flex(iItem)
        iCol = 1
        iCol2Enlarge = 2
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
        .ColWidth(1) = 0.1
        .ColWidth(2) = .Width \ 2
        
        Select Case iItem
            Case 0
                iMaxElem = UBound(xGeneral)
            Case 1
                iMaxElem = UBound(xSecurity)
            Case 2
                iMaxElem = UBound(xConnections)
            Case 3
                iMaxElem = UBound(xServerSet)
            Case 4
                iMaxElem = UBound(xClient)
        End Select
                
        For i = 0 To iMaxElem
            bCombo = False
            Select Case iItem
                Case 0
                    sBuf = vbTab & xGeneral(i).Elem.sField
                    If (xGeneral(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & xGeneral(i).Elem.vVal
                    bBoolean = (xGeneral(i).Elem.iType And am_Bool) = am_Bool
                    If bBoolean Then idBitmap = IIf(Abs(CBool(xGeneral(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                
                    bDisab = xGeneral(i).bDisab
                    bPush = xGeneral(i).EditType = iPush
                    bCombo = xGeneral(i).EditType = iCombo
                    bEllipsis = xGeneral(i).EditType = iEllipsis
                
                Case 1
                    sBuf = vbTab & xSecurity(i).Elem.sField
                    If (xSecurity(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & xSecurity(i).Elem.vVal
                    bBoolean = (xSecurity(i).Elem.iType And am_Bool) = am_Bool
                    If bBoolean Then idBitmap = IIf(Abs(CBool(xSecurity(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                
                    bDisab = xSecurity(i).bDisab
                    bPush = xSecurity(i).EditType = iPush
                    bCombo = xSecurity(i).EditType = iCombo
                    bEllipsis = xSecurity(i).EditType = iEllipsis
                Case 2
                    sBuf = vbTab & xConnections(i).Elem.sField
                    If (xConnections(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & xConnections(i).Elem.vVal
                    bBoolean = (xConnections(i).Elem.iType And am_Bool) = am_Bool
                    If bBoolean Then idBitmap = IIf(Abs(CBool(xConnections(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                    bDisab = xConnections(i).bDisab
                    
                    bDisab = xConnections(i).bDisab
                    bPush = xConnections(i).EditType = iPush
                    bCombo = xConnections(i).EditType = iCombo
                    bEllipsis = xConnections(i).EditType = iEllipsis
                    
                Case 3
                    sBuf = vbTab & xServerSet(i).Elem.sField
                    If (xServerSet(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & xServerSet(i).Elem.vVal
                    bBoolean = (xServerSet(i).Elem.iType And am_Bool) = am_Bool
                    If bBoolean Then idBitmap = IIf(Abs(CBool(xServerSet(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                    bDisab = xServerSet(i).bDisab
                    
                    bDisab = xServerSet(i).bDisab
                    bPush = xServerSet(i).EditType = iPush
                    bCombo = xServerSet(i).EditType = iCombo
                    bEllipsis = xServerSet(i).EditType = iEllipsis
                    
                Case 4
                    sBuf = vbTab & xClient(i).Elem.sField & vbTab & xClient(i).Elem.vVal
                    bDisab = xClient(i).bDisab
                    
            End Select
            Flex(iItem).AddItem sBuf
            If bBoolean Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                Set Flex(iItem).CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter

            End If
            If bDisab Then
                .Row = .Rows - 1
                For iX = 2 To .Cols - 1
                    .Col = iX
                    .CellBackColor = vbInactiveBorder
                Next
            End If
            If bPush Or bCombo Or bEllipsis Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                
                idBitmap = Switch(bPush, k_ResPush, bCombo, k_ResCombo, bEllipsis, k_ResEllipsis)
                Set Flex(iItem).CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = IIf(bEllipsis, flexAlignRightBottom, flexAlignRightTop)
            End If
        Next
        If iItem = 2 Then
            .ColWidth(2) = 1000
        ElseIf iItem = 3 Then
            .ColWidth(2) = .Width \ 4
        End If
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(1) = flexAlignLeftCenter
        '.ColAlignment(2) = flexAlignLeftCenter
        
        .FixedRows = 1
        .FixedCols = 2
        .Redraw = True
        
        SettaLargFlex Flex(iItem), 1, .ColWidth(1)
        miRow = 0
        .TopRow = 1
        .Row = 1
        miRow = 1
        bLoading = False
        Flex_RowColChange iItem
    End With

End Sub

Private Sub FlexEdit(KeyAscii As Integer, ByVal iIndex As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iImg As Integer
    Dim bCombo As Boolean
    Dim bElipses As Boolean
    
    miTmpType = am_String

    iRow = Flex(iIndex).Row
    iCol = Flex(iIndex).Col
    
    miFlex = iIndex
    
    bEdit = True

    If Flex(iIndex).CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        If miFlex = 1 Then
            miTmpType = xSecurity(iRow - 1).Elem.iType
            sBuf = xSecurity(iRow - 1).Elem.vVal
            iLen = xSecurity(iRow - 1).Elem.iLen
            bCombo = xSecurity(iRow - 1).EditType = iCombo
            bElipses = xSecurity(iRow - 1).EditType = iEllipsis
        ElseIf miFlex = 2 Then
            miTmpType = xConnections(iRow - 1).Elem.iType
            sBuf = xConnections(iRow - 1).Elem.vVal
            iLen = xConnections(iRow - 1).Elem.iLen
            bCombo = xConnections(iRow - 1).EditType = iCombo
            
        ElseIf miFlex = 3 Then
            miTmpType = xServerSet(iRow - 1).Elem.iType
            sBuf = xServerSet(iRow - 1).Elem.vVal
            iLen = xServerSet(iRow - 1).Elem.iLen
            bCombo = xServerSet(iRow - 1).EditType = iCombo
            
        End If
        miRow = iRow
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If (miTmpType And am_Bool) = am_Bool Then
                    If miFlex = 1 Then
                        iImg = 1 - (-xSecurity(iRow - 1).Elem.vVal)
                        xSecurity(iRow - 1).Elem.vVal = CBool(iImg)
                        Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    ElseIf miFlex = 2 Then
                        iImg = 1 - (-xConnections(iRow - 1).Elem.vVal)
                        xConnections(iRow - 1).Elem.vVal = CBool(iImg)
                        Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        
                        If iRow >= Flex(miFlex).Rows - 2 Then
                            'autoesclusione reciproca di
                            'NULLS_definedON e NULLS_definedOFF
                            If iImg <> 0 Then
                                bLoading = True
                                If iRow = Flex(miFlex).Rows - 1 Then
                                    xConnections(iRow - 2).Elem.vVal = CBool(0)
                                    Flex(miFlex).Row = iRow - 1
                                ElseIf iRow = Flex(miFlex).Rows - 2 Then
                                    xConnections(iRow).Elem.vVal = CBool(0)
                                    Flex(miFlex).Row = iRow + 1
                                End If
                                iImg = 0
                                Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                Flex(miFlex).Row = iRow
                                
                                bLoading = False
                            End If
                        End If
                    ElseIf miFlex = 3 Then
                        iImg = 1 - (-xServerSet(iRow - 1).Elem.vVal)
                        xServerSet(iRow - 1).Elem.vVal = CBool(iImg)
                        Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    End If
                    bEdit = False
                End If
            Case Else
                If (miTmpType And am_Bool) = am_Bool Then bEdit = False
        End Select
    End If

    If bEdit Then
        If bElipses Then
            If miFlex = 1 Then
                SelectDataRoot miFlex, iRow - 1
            End If
        ElseIf bCombo Then
            If miFlex = 1 Then
                If Flex(iIndex).Row = 1 Then
                    LoadCombo x_Security
                    bCombo = True
                ElseIf Flex(iIndex).Row = 2 Then
                    LoadCombo x_Audit
                    bCombo = True
                End If
            ElseIf miFlex = 3 Then
                If Flex(iIndex).Row = 1 Then
                    LoadCombo x_Language
                    bCombo = True
                End If
            End If
            If bCombo Then
                cbo.ListIndex = 0
                cbo.ListIndex = GetItem(Flex(iIndex).TextMatrix(Flex(iIndex).Row, Flex(iIndex).Col), cbo)
                With Flex(miFlex)
                    cbo.Move .Left + .CellLeft + fra(iIndex).Left, .Top + .CellTop + fra(iIndex).Top, .CellWidth ', .CellHeight
                End With
                cbo.Visible = True
                cbo.ZOrder
                cbo.SetFocus
                If KeyAscii <> 0 Then
                    On Local Error Resume Next
                    SendKeys Chr$(KeyAscii)
                    On Local Error GoTo 0
                End If
            End If
        Else
            tFlex.PasswordChar = ""
            tFlex.MaxLength = 0
            tFlex.Text = ""
            tFlex.MaxLength = iLen
            If miTmpType = am_Long Then
                tFlex.Text = Format$(sBuf, kLNum)
            Else
                tFlex.Text = sBuf
            End If
            Debug.Print miTmpType = am_Long
            tFlex.SelStart = 0
            With Flex(iIndex)
                tFlex.Move .Left + .CellLeft + fra(iIndex).Left, .Top + .CellTop + fra(iIndex).Top, .CellWidth, .CellHeight
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
    End If

End Sub

Private Sub LoadCombo(ByVal iType As cboLoad)

    Dim i As Integer
    With cbo
        .Clear
        If iType = x_Audit Then
            For i = 0 To UBound(xEnumAudit)
                .AddItem xEnumAudit(i).sLang
                .ItemData(.NewIndex) = xEnumAudit(i).idLang
            Next
        ElseIf iType = x_Security Then
            For i = 0 To UBound(xEnumSecurity)
                .AddItem xEnumSecurity(i).sLang
                .ItemData(.NewIndex) = xEnumSecurity(i).idLang
            Next
        ElseIf iType = x_Language Then
            For i = 0 To UBound(xEnumLanguage)
                .AddItem xEnumLanguage(i).sLang
                .ItemData(.NewIndex) = xEnumLanguage(i).idLang
            Next
        End If
    End With
End Sub

Private Sub SaveData()

    Dim bProced As Boolean
    Dim i As Integer, lVal As Long
    Dim bSave As Boolean
    Dim lErr As Long, sErr As String
    
    Screen.MousePointer = vbHourglass
    
    For i = 0 To UBound(xSecurity)
        If Not xSecurity(i).bDisab Then
            If i <> 4 Then
                bProced = xSecurity(i).vOrigValue <> xSecurity(i).Elem.iLen
            Else
                bProced = xSecurity(i).vOrigValue <> xSecurity(i).Elem.vVal
            End If
            If bProced Then Exit For
        End If
    Next
    If Not bProced Then
        For i = 0 To UBound(xConnections)
            bProced = xConnections(i).vOrigValue <> xConnections(i).Elem.vVal
            If bProced Then Exit For
        Next
    End If
    If Not bProced Then
        For i = 0 To UBound(xServerSet)
            If i = 0 Then
                bProced = xServerSet(i).vOrigValue <> xServerSet(i).Elem.iLen
            Else
                bProced = xServerSet(i).vOrigValue <> xServerSet(i).Elem.vVal
            End If
            If bProced Then Exit For
        Next
    End If

    If bProced Then
        On Local Error GoTo ErrHandler
        objServer.Configuration.ShowAdvancedOptions = True
        
        For i = 0 To UBound(xSecurity)
            If i <> 4 Then
                bSave = xSecurity(i).vOrigValue <> xSecurity(i).Elem.iLen
            Else
                bSave = xSecurity(i).vOrigValue <> xSecurity(i).Elem.vVal
            End If
        
            If bSave Then
                If i = 0 Then
                    objServer.IntegratedSecurity.SecurityMode = xSecurity(i).Elem.iLen
                ElseIf i = 1 Then
                    objServer.IntegratedSecurity.AuditLevel = xSecurity(i).Elem.iLen
                ElseIf i = 2 Then
                    'donothing objServer.StartupAccount
                ElseIf i = 3 Then
                    'account
                ElseIf i = 4 Then
                    objServer.Registry.SQLDataRoot = xSecurity(i).Elem.vVal
                End If
            End If
        Next
        
        If xConnections(0).vOrigValue <> xConnections(0).Elem.vVal Then
            objServer.Configuration.ConfigValues(xConnections(0).sConfigName).CurrentValue = xConnections(0).Elem.vVal
        End If
        
        For i = 1 To UBound(xConnections)
            If xConnections(i).Elem.vVal Then
                lVal = lVal Or (2 ^ (i - 1))
            End If
        Next
        If mlngUserOptions <> lVal Then
            objServer.Configuration.ConfigValues(xConnections(1).sConfigName).CurrentValue = lVal
        End If
            
            
        If xServerSet(0).vOrigValue <> xServerSet(0).Elem.iLen Then
            objServer.Configuration.ConfigValues(xServerSet(0).sConfigName).CurrentValue = xServerSet(0).Elem.iLen
        End If
        For i = 1 To UBound(xServerSet)
            If xServerSet(i).vOrigValue <> xServerSet(i).Elem.vVal Then
                objServer.Configuration.ConfigValues(xServerSet(i).sConfigName).CurrentValue = Abs(xServerSet(i).Elem.vVal)
            End If
        Next
            
        objServer.Configuration.ReconfigureWithOverride
        
        objServer.Configuration.ShowAdvancedOptions = False
        
        For i = 0 To 4
            RefreshData i
            LoadFlex i
        Next
        Screen.MousePointer = vbDefault
        MsgBox ReplaceMsg(MyLoadResString(kExecutedOverrides), Array("1%"), Array(objServer.NetName)), vbInformation, App.EXEName
    End If
        
    
ExitHere:
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
    lErr = Err.Number
    sErr = Err.Description
    Screen.MousePointer = vbDefault
    
    objServer.Configuration.ShowAdvancedOptions = False
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    Resume ExitHere
    
End Sub
Private Sub SelectDataRoot(ByVal iFlex As Integer, ByVal iRow As Integer)

    Dim bCancel As Boolean
    Dim bErr As Boolean
    Dim sTarget As String
    Dim lErr As Long, sErr As String
    
    Dim fBrowse As fBr4Fold
    Set fBrowse = New fBr4Fold
        
    With fBrowse
        .Init enB4isSelectDir, MyLoadResString(kDefaultDataRoot), xSecurity(4).Elem.vVal
                    
        bErr = .ErrLoadingInit()
        If Not bErr Then
            .Show vbModal, Me
            bCancel = .Cancelled()
            If Not bCancel Then sTarget = .FullPath()
        Else
            .GetError lErr, sErr
        End If
    End With
    Unload fBrowse
    Set fBrowse = Nothing
    If Not bErr And Not bCancel Then
        If Right$(sTarget, 1) = "\" Then sTarget = Left$(sTarget, Len(sTarget) - 1)
        xSecurity(4).Elem.vVal = sTarget
        Flex(iFlex).TextMatrix(iRow + 1, 2) = sTarget
        Flex(iFlex).ToolTipText = sTarget
        Debug.Print objServer.Registry.SQLDataRoot
    End If
    
    If lErr <> 0 Then MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)

End Sub
