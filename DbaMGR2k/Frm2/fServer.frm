VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Connetti"
      Default         =   -1  'True
      Height          =   360
      Index           =   2
      Left            =   4200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Annulla"
      Height          =   360
      Index           =   1
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strLng As String
Private miLang As Integer
Private Type udtSTR
    Elem As elemFlexG
    EditType As TypeEnum
    bDisab As Boolean
    bPwd As Boolean
End Type
Private Enum cboLoad
    x_Lang = 0
    x_Lock
    x_Detach
End Enum

Private xConnect() As udtSTR
Private xLang() As udtLANG
Private xLock() As udtLANG
Private xDetach() As udtLANG
Private xElem() As udtSTR
Private miTmpType As iPropG
Private mbPwd As Boolean
Private bLoading As Boolean
Private miRow As Integer
Private miFlex As Integer
Private m_blnAnnullaMod As Boolean
Private bChangeConn As Boolean
Private m_lngVSBarW  As Long

Public Property Get ChangedConn() As Boolean
    ChangedConn = bChangeConn
End Property

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub cmd_Click(Index As Integer)

    TestActiveControl
    Select Case Index
        Case 1
            Me.Hide
        Case 2
            Connect2Sql
        Case 0
            ListServer
    End Select
End Sub
Private Sub ListServer()
    
    Dim a As BROWSEINFO
    Dim pidlRoot As Long
    Dim sBuf As String
    
    Dim pidlBrowse As Long

    a.lpszTitle = MyLoadResString(k_Select_Server, strLng)   '"Seleziona Server"
    pidlRoot = GetPIDLFromFolderID(Me.hWnd, CSIDL_NETWORK)
    a.pidlRoot = pidlRoot
    a.ulFlags = BIF_BROWSEFORCOMPUTER
        
    a.hOwner = Me.hWnd
    pidlBrowse = SHBrowseForFolder(a)
    
    sBuf = GetFileDisplayName(pidlBrowse)
    'If Len(sBuf) = 0 Then sBuf = NomeComputer()
    If (Len(sBuf) = 0) Or (UCase$(sBuf) = NomeComputer()) Then sBuf = "(local)"
    
    xConnect(0).Elem.vVal = sBuf
    Flex.TextMatrix(1, 2) = sBuf

End Sub
Private Sub Connect2Sql()
  
    Dim sServer As String
    Dim sUser As String
    Dim vPwd As Variant
    Dim iTime As Integer
    Dim bTrusted As Boolean, bAutoStart As Boolean
    Dim iVal As Integer
    Dim lErr As Long, sErr As String
  
    cbo.Visible = False
    tFlex.Visible = False

    Set tStrip.SelectedItem = tStrip.Tabs(1)
    Flex.Refresh
    Me.Refresh
    iTime = Val(xConnect(3).Elem.vVal)
    If iTime = 0 Then
        MsgBox MyLoadResString(k_Time_Out_Not_Defined, strLng), vbInformation Or vbOKOnly, App.EXEName
        Exit Sub
    End If
  
    frmMain.stbMain.Panels(1).Text = "Connecting..."
    DoEvents
  
    sServer = Trim$(xConnect(0).Elem.vVal)
    sUser = Trim$(xConnect(1).Elem.vVal)
    vPwd = Trim$(xConnect(2).Elem.vVal)
    If Len(vPwd) = 0 Then vPwd = Null
  
    If StrComp(sServer, NomeComputer(), vbTextCompare) = 0 Then sServer = "(local)"
  
    Screen.MousePointer = vbHourglass
    If Not objServer Is Nothing Then
        On Local Error Resume Next
        objServer.DisConnect
        objServer.Close
        On Local Error GoTo 0
        Set objServer = Nothing
    End If
    
    On Local Error GoTo LoginError
    Set objServer = New SQLDMO.SQLServer2
    objServer.AutoReConnect = True      'enables auto-reconnect when connection is broken
   
  
    'objServer.Connect txtServer.Text, txtUserID.Text, txtPassword.Text
    objServer.LoginTimeout = iTime '60
    
    bTrusted = CBool(xElem(1).Elem.vVal)
    bAutoStart = CBool(xElem(2).Elem.vVal)
    If bTrusted Then objServer.LoginSecure = True
    
    If StrComp(sServer, "(local)", vbTextCompare) = 0 And (bAutoStart) Then
        objServer.Start True, sServer, sUser, vPwd
    Else
        objServer.Connect sServer, sUser, vPwd
    End If
    bChangeConn = True
    frmMain.stbMain.Panels(1).Text = "Ready"
    'DoEvents
    Me.Hide
    frmMain.Refresh
    DoEvents
  
    Call WriteIni("CONNECT", "SERVER", Trim$(xConnect(0).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "USER", Trim$(xConnect(1).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    
    If Len(GetProperName(Trim$(xElem(0).Elem.vVal))) Then
        g_strLNG = Trim$(xElem(0).Elem.vVal)
    Else
        If Len(GetProperName(g_strLNG)) = 0 Then g_strLNG = "English"
    End If
    'Call WriteIni("CONNECT", "LCID", Format$(Abs(giLang)), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "LCID", g_strLNG, App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "AUTOSTART", Format$(Abs(xElem(2).Elem.vVal)), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "SHOWALL", Format$(1 - Abs(xElem(3).Elem.vVal)), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "TIMEOUT", Format$((xConnect(3).Elem.vVal)), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "TRUSTED", Format$(Abs(xElem(1).Elem.vVal)), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("CONNECT", "SLEEP", Format$(Val(xElem(4).Elem.vVal)), App.Path & "\" & App.EXEName & ".INI")
    
        
    Call WriteIni("DEFAULT", "BOL", CStr(xElem(6).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("DEFAULT", "DETACHMODE", Format$(Val(xElem(7).Elem.iLen)), App.Path & "\" & App.EXEName & ".INI")
        
    iVal = Val(xElem(5).Elem.iLen)
    If iVal = 0 Then iVal = adLockPessimistic
    Call WriteIni("CONNECT", "LOCK_TYPE", Format$(iVal), App.Path & "\" & App.EXEName & ".INI")
      
    Exit Sub

LoginError:
    lErr = Err.Number
    sErr = Err.Description
    frmMain.stbMain.Panels(1).Text = "Ready"

    DoEvents
    Set objServer = Nothing
    Screen.MousePointer = vbDefault
    MsgBox lErr & "- " & sErr, vbCritical, "Login Error"
    
End Sub

Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iCode As Integer
    Dim bPush As Boolean
    Dim bCombo As Boolean
    
    If miFlex = 0 Then
        bPush = xConnect(miRow - 1).EditType = iPush And KeyCode = vbKeyF2
        bCombo = xConnect(miRow - 1).EditType = iCombo And KeyCode = vbKeyF2
    Else
        bCombo = xElem(miRow - 1).EditType = iCombo And KeyCode = vbKeyF2
    End If
    
    If bPush Or bCombo Then
        KeyCode = 0
        Flex_DblClick
    Else
        iCode = KeyCode
        KeyCode = 0
        If iCode = vbKeyF2 Then FlexEdit 0
    End If

End Sub

Private Sub Form_Load()
    
    Dim sBuf As String
    Dim i As Integer
    
    strLng = g_strLNG
       
    m_lngVSBarW = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CYVTHUMB)
    Flex.RowHeightMin = cbo.Height + 10
    
    On Local Error Resume Next
    sBuf = Dir$(App.Path & "\*.Ln2")
    Do While Len(sBuf) <> 0 And Err.Number = 0
        sBuf = Left$(sBuf, InStr(sBuf, ".") - 1)
        If Len(sBuf) Then
            ReDim Preserve xLang(i) As udtLANG
            xLang(i).sLang = StrConv(sBuf, vbProperCase)
            xLang(i).idLang = i
            i = i + 1
        End If
        sBuf = Dir$
    Loop
    On Local Error GoTo 0
    
    
    'ReDim xLang(1) As udtLANG
    
            
'    xLang(0).sLang = "Italiano"
'    xLang(0).idLang = 0
'
'    xLang(1).sLang = "English"
'    xLang(1).idLang = 1
    
    Me.Move ((frmMain.Width - Me.Width) \ 2) + frmMain.Left, ((frmMain.Height - Me.Height) \ 2) + frmMain.Top
    'CenterForm Me
        
    ReDim xConnect(3) As udtSTR
    
    sBuf = Trim$(GetIni("CONNECT", "SERVER", App.Path & "\" & App.EXEName & ".INI"))
    'If (Len(sBuf) = 0) Then sBuf = "(local)" 'NomeComputer()
    If (Len(sBuf) = 0) Or (StrComp(sBuf, NomeComputer(), vbTextCompare) = 0) Then sBuf = "(local)"
    xConnect(0).EditType = iEllipsis
    'xConnect(0).bEllipsis = True
    xConnect(0).Elem.vVal = sBuf
    xConnect(0).Elem.iLen = 25
    xConnect(0).Elem.sField = "Server"
    xConnect(0).Elem.iType = am_String
    
    
    sBuf = Trim$(GetIni("CONNECT", "USER", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) = 0 Then sBuf = "sa"
    xConnect(1).Elem.vVal = sBuf
    xConnect(1).Elem.iLen = 25
    xConnect(1).Elem.sField = "User"
    xConnect(1).Elem.iType = am_String
    
    xConnect(2).Elem.vVal = ""
    xConnect(2).Elem.iLen = 25
    xConnect(2).Elem.sField = "Password"
    xConnect(2).Elem.iType = am_String
    xConnect(2).bPwd = True
    
    sBuf = Trim$(GetIni("CONNECT", "TIMEOUT", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) = 0 Then sBuf = "60"
    xConnect(3).Elem.vVal = Val(sBuf)
    xConnect(3).Elem.iLen = 3
    xConnect(3).Elem.sField = "Login Timeout"
    xConnect(3).Elem.iType = am_Long

    LoadString True
    LoadFlex
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fServer = Nothing
End Sub

Private Sub LoadFlex()

    Dim sBuf As String
    Dim i As Integer
    Dim iX As Integer
    Dim bDisab As Boolean, iMaxElem As Integer
    Dim bBoolean As Boolean
    Dim bPush As Boolean, bCombo As Boolean, bEllipsis As Boolean
    Dim idBitmap As Integer
    
    With Flex
        bLoading = True
        .Redraw = False
        .Clear
        .FixedCols = 0
        .Cols = 3
        .Rows = 1
        .Clear
                
        .TextMatrix(0, 1) = MyLoadResString(k_Property, strLng)
        .TextMatrix(0, 2) = MyLoadResString(k_Value, strLng)
        
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        .ColWidth(1) = 0.1
        .ColWidth(2) = .Width \ 2
        
        iMaxElem = IIf(miFlex = 0, UBound(xConnect), UBound(xElem))
        For i = 0 To iMaxElem
            bPush = False
            bCombo = False
            If miFlex = 0 Then
                sBuf = vbTab & xConnect(i).Elem.sField
                If (xConnect(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & IIf(xConnect(i).bPwd, String$(Len(xConnect(i).Elem.vVal), "*"), xConnect(i).Elem.vVal)
                bBoolean = (xConnect(i).Elem.iType And am_Bool) = am_Bool
                If bBoolean Then idBitmap = IIf(Abs(CBool(xConnect(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                
                bDisab = xConnect(i).bDisab
                xElem(0).EditType = iCombo
                bPush = xConnect(i).EditType = iPush
                bCombo = xConnect(i).EditType = iCombo
                bEllipsis = xConnect(i).EditType = iEllipsis
            Else
                sBuf = vbTab & xElem(i).Elem.sField
                If (xElem(i).Elem.iType And am_Bool) <> am_Bool Then sBuf = sBuf & vbTab & xElem(i).Elem.vVal
                bBoolean = (xElem(i).Elem.iType And am_Bool) = am_Bool
                If bBoolean Then idBitmap = IIf(Abs(CBool(xElem(i).Elem.vVal)), k_ResChkON, k_ResChkOFF)
                
                bDisab = xElem(i).bDisab
                bPush = xElem(i).EditType = iPush
                bCombo = xElem(i).EditType = iCombo
                bEllipsis = xElem(i).EditType = iEllipsis
                
            End If
            Flex.AddItem sBuf
            If bBoolean Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                Set Flex.CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter

            End If
            If bDisab Then
                For iX = 2 To .Cols - 1
                    .Col = iX
                    .CellBackColor = vbInactiveBorder
                Next
            End If
            If bPush Or bCombo Or bEllipsis Then
                .Col = .Cols - 1
                .Row = .Rows - 1
                
                idBitmap = Switch(bPush, k_ResPush, bCombo, k_ResCombo, bEllipsis, k_ResEllipsis)
                'idBitmap = IIf(bPush, k_ResPush, k_ResCombo)
                Set Flex.CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = IIf(bEllipsis, flexAlignRightBottom, flexAlignRightTop)
                '.CellPictureAlignment = flexAlignRightTop
            End If
        Next
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .FixedRows = 1
        .FixedCols = 2
        If miFlex = 0 Then
            .Height = 1750
        Else
            .Height = 2500
        End If
        .Redraw = True
        
        SettaLargFlex Flex, 1, .ColWidth(1)
        miRow = 0
        .TopRow = 1
        .Row = 1
        miRow = 1
        bLoading = False
        Flex_RowColChange
    End With

End Sub
Private Sub Flex_DblClick()

    Dim bPush As Boolean
    Dim bCombo As Boolean
    Dim bEllipsis As Boolean
    
    'If Flex.Col = Flex.Cols - 1 And Flex.MouseRow > 0 And Flex.MouseCol > 1 Then
    If Flex.Col = Flex.Cols - 1 Then 'And Flex.MouseRow > 0 And Flex.MouseCol > 1 Then
        If miFlex = 0 Then
            bPush = xConnect(miRow - 1).EditType = iPush
            bCombo = xConnect(miRow - 1).EditType = iCombo
            bEllipsis = xConnect(miRow - 1).EditType = iEllipsis
        Else
            bCombo = xElem(miRow - 1).EditType = iCombo
            bEllipsis = xElem(miRow - 1).EditType = iEllipsis
        End If
    
        With Flex
            If bCombo Then
                If miFlex = 1 Then
                    If .Row = 1 Then
                        LoadCombo x_Lang
                    ElseIf .Row = 6 Then
                        LoadCombo x_Lock
                    ElseIf .Row = 8 Then
                        LoadCombo x_Detach
                    End If
                    cbo.ListIndex = 0
                    cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                    cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                    cbo.Visible = True
                    cbo.SetFocus
                End If
            ElseIf bEllipsis Then
                If miFlex = 0 Then
                    ListServer
                Else
                    Search4Help
                End If
            Else
                FlexEdit vbKeySpace
            End If
        End With
    End If

End Sub

Private Sub Flex_GotFocus()
    
    TestActiveControl
    cbo.Visible = False
    tFlex.Visible = False
    m_blnAnnullaMod = True
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)

    Dim iCode As Integer
    Dim bPush As Boolean
    Dim bCombo As Boolean
    
    If miFlex = 0 Then
        bPush = xConnect(miRow - 1).EditType = iPush And KeyAscii = vbKeySpace
        bCombo = xConnect(miRow - 1).EditType = iCombo
    Else
        bCombo = xElem(miRow - 1).EditType = iCombo
    End If
    
    If bPush Or bCombo Then
        KeyAscii = 0
        Flex_DblClick
    Else
        iCode = KeyAscii
        KeyAscii = 0
        If miFlex = 0 Then
            If Not xConnect(miRow - 1).EditType = iPush Then FlexEdit iCode
        Else
            FlexEdit iCode
        End If
    End If
    
End Sub

Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
    
    If Not bLoading Then
        If Not b Then
            b = True
            
            Debug.Print m_blnAnnullaMod
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

Private Sub Flex_Scroll()
    cbo.Visible = False
    tFlex.Visible = False
End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyRIGHT, vbKeyLEFT
                iCode = KeyCode
                KeyCode = 0
                cbo.Visible = False
                Flex.Col = Flex.Col - (38 - iCode)
        End Select
    End If
End Sub
Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Flex.SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex.SetFocus
    End Select

End Sub
Private Sub cbo_LostFocus()
    
    If Not m_blnAnnullaMod Then
        With xElem(miRow - 1)
            If miRow - 1 = 0 Then
                miLang = cbo.ItemData(cbo.ListIndex)
                If .Elem.iLen <> miLang Then
                    strLng = cbo.List(cbo.ListIndex)
                    .Elem.iLen = miLang
                End If
            ElseIf miRow - 1 = 5 Then
                .Elem.iLen = cbo.ItemData(cbo.ListIndex)
                Flex.TextMatrix(miRow, 2) = GetCBODescriptionByID(x_Lock, .Elem.iLen)
            ElseIf miRow - 1 = 7 Then
                .Elem.iLen = cbo.ItemData(cbo.ListIndex)
                Flex.TextMatrix(miRow, 2) = GetCBODescriptionByID(x_Detach, .Elem.iLen)
            End If
        End With
        
    End If
    
    cbo.Visible = False
    
    If Not m_blnAnnullaMod Then Flex.Col = 2
    If (miRow - 1 = 0) And Not m_blnAnnullaMod Then LoadString
        
    m_blnAnnullaMod = True

End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iImg As Integer
    
    miTmpType = am_String
    mbPwd = False
    iRow = Flex.Row
    iCol = Flex.Col
    
    bEdit = True
    mbPwd = False
    If Flex.CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        If miFlex = 0 Then
            miTmpType = xConnect(iRow - 1).Elem.iType
            sBuf = xConnect(iRow - 1).Elem.vVal
            iLen = xConnect(iRow - 1).Elem.iLen
            mbPwd = xConnect(iRow - 1).bPwd
        Else
            miTmpType = xElem(iRow - 1).Elem.iType
            sBuf = xElem(iRow - 1).Elem.vVal
            iLen = xElem(iRow - 1).Elem.iLen
        End If
        miRow = iRow
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If (miTmpType And am_Bool) = am_Bool And miFlex = 1 Then
                    iImg = 1 - (-xElem(iRow - 1).Elem.vVal)
                    xElem(iRow - 1).Elem.vVal = CBool(iImg)
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
        If mbPwd Then tFlex.PasswordChar = "*"
        tFlex.Text = sBuf
        Debug.Print miTmpType = am_Long
        tFlex.SelStart = 0
        With Flex
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If (miTmpType And am_String) = am_String Then If KeyAscii = vbKeySpace Then KeyAscii = 0
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    TestActiveControl
    
    i = tStrip.SelectedItem.Index - 1
    miRow = 0
    miFlex = i
    If i = 1 Then LoadString False
    LoadFlex
    Flex.SetFocus
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

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUP Then
        KeyCode = 0
        If Flex.Row > 1 Then
            Flex.Row = Flex.Row - 1
        End If
        Flex.SetFocus
    ElseIf KeyCode = vbKeyDOWN Then
        If Flex.Row < Flex.Rows - 1 Then
            Flex.Row = Flex.Row + 1
        End If
        Flex.SetFocus
'
'    If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then
'        KeyCode = 0
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
            KeyAscii = 0
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
'        If miFlex = 0 Then
'            bUpdate = True
'        Else
'            Select Case miRow
'                Case 0, 1
'                    If Len(sBuf) Then bUpdate = True
'                Case 2
'            End Select
'        End If
'        If bUpdate Then
'        End If
    End If
    If bUpdate Then
        If miFlex = 0 Then
            xConnect(miRow - 1).Elem.vVal = sBuf
        Else
            xElem(miRow - 1).Elem.vVal = sBuf
        End If
        Flex.TextMatrix(miRow, 2) = IIf(mbPwd, String$(Len(sBuf), "*"), sBuf)
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    
    m_blnAnnullaMod = True
    tFlex.Visible = False
    Flex.SetFocus
    
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub LoadString(Optional ByVal b1Time As Boolean = False)

    Dim sComp As String
    
    Screen.MousePointer = vbHourglass
    
    Me.Caption = MyLoadResString(k_SqlServer_Connection_Properties, strLng)
    
    cmd(1).Caption = MyLoadResString(k_Cancel, strLng)
    cmd(2).Caption = IIf(objServer Is Nothing, MyLoadResString(k_Connect, strLng), MyLoadResString(k_Reconnect, strLng))
    tStrip.Tabs(1).Caption = MyLoadResString(k_Connection_Propertyes, strLng)
    tStrip.Tabs(2).Caption = MyLoadResString(k_Settings, strLng)
    
    If b1Time Then ReDim xElem(7) As udtSTR
    
    xElem(0).EditType = iCombo
    xElem(0).Elem.vVal = strLng
    xElem(0).Elem.iLen = Val(GetCBODescriptionByID(x_Lang, strLng))
    xElem(0).Elem.sField = ReplaceMsg(MyLoadResString(k_DbaMGR_Language, strLng), Array("1%"), Array(App.EXEName))
    xElem(0).Elem.iType = am_Long
    
    
    sComp = Trim$(GetIni("CONNECT", "SERVER", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sComp) = 0 Then sComp = "(local)" 'NomeComputer()
    If b1Time Then xElem(1).Elem.vVal = CBool(Val(GetIni("CONNECT", "TRUSTED", App.Path & "\" & App.EXEName & ".INI")))
    xElem(1).Elem.sField = MyLoadResString(k_Trusted_NT, strLng)
    xElem(1).Elem.iType = am_Bool

    
    If b1Time Then xElem(2).Elem.vVal = CBool(Val(GetIni("CONNECT", "AUTOSTART", App.Path & "\" & App.EXEName & ".INI")))
    'xElem(2).bDisab = IsWindowsNT() Or (UCase$(sComp) <> UCase$(NomeComputer()))
    xElem(2).bDisab = IsWindowsNT() Or (UCase$(sComp) <> "(LOCAL)" And UCase$(sComp) <> NomeComputer())
    xElem(2).Elem.sField = "AutoStart SqlServer"
    xElem(2).Elem.iType = am_Bool
    If xElem(2).bDisab Then xElem(2).Elem.vVal = False
        
  
    If b1Time Then xElem(3).Elem.vVal = Not CBool(Val(GetIni("CONNECT", "SHOWALL", App.Path & "\" & App.EXEName & ".INI")))
    xElem(3).Elem.sField = MyLoadResString(k_Hide_Non_Granted_DB, strLng)
    xElem(3).Elem.iType = am_Bool
    
    If b1Time Then xElem(4).Elem.vVal = Val(GetIni("CONNECT", "SLEEP", App.Path & "\" & App.EXEName & ".INI"))
    xElem(4).Elem.sField = MyLoadResString(k_Seconds_to_wait_for_Detach, strLng)
    xElem(4).Elem.iType = am_Long
    xElem(4).Elem.iLen = 2
    
    ReDim xLock(1) As udtLANG
    xLock(0).sLang = MyLoadResString(k_Pessimistic_Lock_default, strLng)
    xLock(0).idLang = adLockPessimistic
    xLock(1).sLang = MyLoadResString(k_Optimistic_Lock, strLng)
    xLock(1).idLang = adLockOptimistic
    
    If b1Time Then xElem(5).Elem.iLen = Val(GetIni("CONNECT", "LOCK_TYPE", App.Path & "\" & App.EXEName & ".INI"))
    If xElem(5).Elem.iLen = 0 Then xElem(5).Elem.iLen = adLockPessimistic
    xElem(5).Elem.sField = MyLoadResString(k_Lock_Type, strLng)
    xElem(5).Elem.vVal = GetCBODescriptionByID(x_Lock, xElem(5).Elem.iLen)
    xElem(5).Elem.iType = am_Long
    xElem(5).EditType = iCombo
    
    
    If b1Time Then xElem(6).Elem.vVal = Trim$(GetIni("DEFAULT", "BOL", App.Path & "\" & App.EXEName & ".INI"))
    If Len(Trim$(xElem(6).Elem.vVal)) = 0 Then xElem(6).Elem.vVal = App.HelpFile
    xElem(6).Elem.sField = MyLoadResString(k_Help_File, strLng)
    xElem(6).Elem.iType = am_String
    xElem(6).EditType = iEllipsis
    
    
    ReDim xDetach(2) As udtLANG
    xDetach(0).sLang = MyLoadResString(k_DetachMode_Check, strLng)
    xDetach(0).idLang = amCheckDb
    xDetach(1).sLang = MyLoadResString(k_DetachMode_NoCheck, strLng)
    xDetach(1).idLang = amNoCheck
    xDetach(2).sLang = MyLoadResString(k_DetachMode_Ask, strLng)
    xDetach(2).idLang = amAskEach
    
    If b1Time Then xElem(7).Elem.iLen = Val(Trim$(GetIni("DEFAULT", "DETACHMODE", App.Path & "\" & App.EXEName & ".INI")))
    xElem(7).EditType = iCombo
    xElem(7).Elem.vVal = GetCBODescriptionByID(x_Detach, xElem(7).Elem.iLen)
    xElem(7).Elem.sField = MyLoadResString(k_Detach_Setting, strLng)
    xElem(7).Elem.iType = am_Long
    
    If Not b1Time Then LoadFlex
    Screen.MousePointer = vbDefault
            
End Sub
Private Sub LoadCombo(ByVal iType As cboLoad)

    Dim i As Integer
    With cbo
        .Clear
        If iType = x_Lock Then
            For i = 0 To UBound(xLock)
                .AddItem xLock(i).sLang
                .ItemData(.NewIndex) = xLock(i).idLang
            Next
        ElseIf iType = x_Lang Then
            For i = 0 To UBound(xLang)
                .AddItem xLang(i).sLang
                .ItemData(.NewIndex) = xLang(i).idLang
            Next
        ElseIf iType = x_Detach Then
            For i = 0 To UBound(xDetach)
                .AddItem xDetach(i).sLang
                .ItemData(.NewIndex) = xDetach(i).idLang
            Next
        End If
    End With
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


Private Function GetCBODescriptionByID(ByVal iType As cboLoad, ByVal vId As Variant) As String

    Dim i As Integer
    If iType = x_Lock Then
        vId = CInt(vId)
        For i = 0 To UBound(xLock)
            If xLock(i).idLang = vId Then
                GetCBODescriptionByID = xLock(i).sLang
                Exit For
            End If
        Next
    ElseIf iType = x_Lang Then
        If VarType(vId) = vbString Then
            vId = CStr(vId)
            For i = 0 To UBound(xLang)
                If xLang(i).sLang = vId Then
                    GetCBODescriptionByID = xLang(i).idLang
                    Exit For
                End If
            Next
        
        Else
            vId = CInt(vId)
            For i = 0 To UBound(xLang)
                If xLang(i).idLang = vId Then
                    GetCBODescriptionByID = xLang(i).sLang
                    Exit For
                End If
            Next
        End If
    ElseIf iType = x_Detach Then
        vId = CInt(vId)
        For i = 0 To UBound(xDetach)
            If xDetach(i).idLang = vId Then
                GetCBODescriptionByID = xDetach(i).sLang
                Exit For
            End If
        Next
        
    End If
    
End Function
Private Sub Search4Help()

    Dim sBuf As String
    Dim sFile As String
    Dim lErr As Long, sErr As String
    
    sBuf = xElem(6).Elem.vVal
    If InStrRev(sBuf, "\") <> 0 Then
        sBuf = Left$(sBuf, InStrRev(sBuf, "\") - 1)
    End If
    'If Len(sBuf) = 0 Then sBuf = App.Path
    If Len(sBuf) = 0 Then
        sBuf = App.HelpFile
        If InStrRev(sBuf, "\") <> 0 Then
            sBuf = Left$(sBuf, InStrRev(sBuf, "\") - 1)
        End If
    End If
    
    Dim obj As cOpen
    Set obj = New cOpen
    
    obj.Title = MyLoadResString(k_Help_File_Browse)
    obj.Def_Dir() = sBuf
    obj.Filter() = "Help File (*.Hlp)" & Chr$(0) & "*.Hlp" & Chr$(0) & "Chm File (*.Chm)" & Chr$(0) & "*.Chm" & Chr$(0) & "Col File (*.Col)" & Chr$(0) & "*.Col" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    obj.formHwnd() = Me.hWnd
    obj.Do_Open
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
            xElem(6).Elem.vVal = sFile
            Flex.TextMatrix(7, 2) = sFile
        End If
    End If
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If
            
End Sub
