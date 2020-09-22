VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fGenOpt 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Opzioni Database"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5145
   Icon            =   "fGenOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1815
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmd 
      Caption         =   "#"
      Height          =   360
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
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
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "#"
      Height          =   360
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "#"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "fGenOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type elSetting
    Elem As PageElement
    EditType As TypeEnum
    bDisab As Boolean
    iComboType As amCbo2Read
End Type

Private m_bIsObjectOwner As Boolean
Private m_lWidth As Long
Private m_lHeight As Long
Private miFlex As Integer       'Indice Flex
Private miRow As Integer        'Riga Flex
Private m_blnAnnullaMod As Boolean

Private m_sDbName As String
Private sStuff As String
Private Elem() As elSetting
Private UserElem() As PageElement
Private ElemSetting() As elSetting
Private bDirty As Boolean
Private iActionMode As optAction

Public Property Let Action(ByVal i As optAction)
    iActionMode = i
        
    Select Case iActionMode
        Case act_DBOption
            cmd(2).Visible = False
            'cmd(2).Caption = MyLoadResString(k_Property )
            LoadProp
        Case act_Q_Options
            cmd(2).Visible = False
            LoadOptions
            
        Case act_User
            cmd(2).Visible = False
            LoadUser
        Case act_Tbl
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadTable
        Case act_User_Function
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadFuncParameter
        Case act_View
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadViewProperties
        Case act_SP
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadSpParameter
        Case act_Extended_VIEW_Idx_Trig
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadView_Idx_Trig
        Case act_Extended_TABLE_Idx_Rel
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadTable_Idx_Rel
        Case act_Extended_TABLE_Trig
            cmd(1).Visible = False
            cmd(2).Visible = False
            m_lWidth = Me.Width * 1.5
            LoadTable_Triggers
        Case Else
    End Select
    Me.Caption = MyLoadResString(k_Properties_for) & " " & Me.Caption
    
End Property
Public Property Let DBName(ByVal sVal As String)
    m_sDbName = sVal
End Property
Public Property Let Stuff(ByVal sVal As String)
    sStuff = sVal
End Property

Private Sub LoadProp()

    Dim oDb As SQLDMO.Database2
    Dim i As Integer, sBuf As String
    Dim bBool As Boolean
        
    If objServer.SaLogin Then
        Flex(1).RowHeightMin = cbo.Height
        
        Screen.MousePointer = vbHourglass
        Set oDb = objServer.Databases(m_sDbName)
        
        With Flex(0)
            .Redraw = False
            .ScrollBars = flexScrollBarBoth
            .Clear
            .FixedCols = 0
            .Cols = 2
            .Rows = 1
            .ColWidth(0) = .Width / 2
            .ColWidth(1) = 2 * .ColWidth(0)
            .ScrollBars = flexScrollBarBoth
            .ColAlignment(1) = flexAlignLeftCenter
                        
            .TextMatrix(0, 0) = MyLoadResString(k_Property)
            .TextMatrix(0, 1) = MyLoadResString(k_Value)
        End With
                    
        Dim l As Integer
        With oDb
            Debug.Print .Status
            On Local Error Resume Next
            sBuf = Format$(.DBOption.Offline)
            On Local Error GoTo 0
            
            If .Status = SQLDMODBStat_Normal Then
                For i = 1 To .FileGroups.Count
                    Flex(0).AddItem "File Group" & vbTab & .FileGroups(i).Name
                    Flex(0).Col = 1
                    Flex(0).Row = Flex(0).Rows - 1
                    
                    Flex(0).CellBackColor = vbHighlight
                    Flex(0).CellForeColor = vbHighlightText
    
                    For l = 1 To .FileGroups(i).DBFiles.Count
                        Debug.Print .FileGroups(i).DBFiles(l).Name
                        Debug.Print .FileGroups(i).DBFiles(l).PhysicalName
                        'Flex(0).AddItem "DB Name" & vbTab & .FileGroups(i).DBFiles(l).Name
                        'Flex(0).AddItem "Phisical Name" & vbTab & .FileGroups(i).DBFiles(l).PhysicalName
                        
                        Flex(0).AddItem MyLoadResString(k_DbName) & vbTab & .FileGroups(i).DBFiles(l).Name
                        Flex(0).AddItem MyLoadResString(k_HtmPhysicalName) & vbTab & .FileGroups(i).DBFiles(l).PhysicalName
                        
                    Next
                    For l = 1 To .TransactionLog.LogFiles.Count
                        Debug.Print .TransactionLog.LogFiles(l).Name
                        Debug.Print .TransactionLog.LogFiles(l).PhysicalName
'                        Flex(0).AddItem "Log Name" & vbTab & .TransactionLog.LogFiles(l).Name
'                        Flex(0).AddItem "Phisical Name" & vbTab & .TransactionLog.LogFiles(l).PhysicalName
                        Flex(0).AddItem MyLoadResString(k_Log_Name) & vbTab & .TransactionLog.LogFiles(l).Name
                        Flex(0).AddItem MyLoadResString(k_HtmPhysicalName) & vbTab & .TransactionLog.LogFiles(l).PhysicalName
                    
                    Next
                    Flex(0).AddItem MyLoadResString(k_Last_BackUp) & vbTab & .TransactionLog.LastBackup
                    
                Next
            End If
            If .Status = SQLDMODBStat_Normal Then
                For i = 1 To .Properties.Count - 1
                    Select Case VarType(.Properties(i).Value)
                        Case vbString
                            sBuf = .Properties(i).Value
                        Case vbInteger, vbLong
                            sBuf = Format$(.Properties(i).Value, "#,0")
                        Case vbSingle, vbDouble, vbDecimal
                            sBuf = Format$(.Properties(i).Value, "#,0.0000")
                        Case vbDate
                            sBuf = Format$(.Properties(i).Value, "dd/mm/yyyy")
                        Case vbBoolean, vbByte
                            bBool = True
                            sBuf = ""
                    End Select
                    Flex(0).AddItem .Properties(i).Name & vbTab & sBuf
                    If bBool Then
                        Flex(0).Col = Flex(0).Cols - 1
                        Flex(0).Row = Flex(0).Rows - 1
                        
                        Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(.Properties(i).Value), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        Flex(0).CellPictureAlignment = flexAlignCenterCenter
                    End If
                    bBool = False
                Next
            Else
                Dim statMode(4) As PageElement
                
                statMode(0).sName = MyLoadResString(k_DBStatus_OffLine)
                statMode(0).vVal = .DBOption.Offline
                statMode(0).iType = am_Bool
                
                statMode(1).sName = MyLoadResString(k_DBStatus_Suspect)
                statMode(1).vVal = .Status = SQLDMODBStat_Suspect
                statMode(1).iType = am_Bool
                
                statMode(2).sName = MyLoadResString(k_DBStatus_Loading)
                statMode(2).vVal = .Status = SQLDMODBStat_Loading
                statMode(2).iType = am_Bool
                
                statMode(3).sName = MyLoadResString(k_DBStatus_Recovering)
                statMode(3).vVal = .Status = SQLDMODBStat_Recovering
                statMode(3).iType = am_Bool
                
                statMode(4).sName = MyLoadResString(k_DBStatus_Stand_By)
                statMode(4).vVal = .Status = SQLDMODBStat_Standby
                statMode(4).iType = am_Bool
                
                For i = 0 To UBound(statMode)
                    Flex(0).AddItem statMode(i).sName
                    Flex(0).Col = Flex(0).Cols - 1
                    Flex(0).Row = Flex(0).Rows - 1
                    
                    Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(statMode(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex(0).CellPictureAlignment = flexAlignCenterCenter
                    
                    If statMode(i).vVal = True Then
                        Flex(0).CellBackColor = vbHighlight
                        Flex(0).CellForeColor = vbHighlightText
                    End If
                Next
            End If
        End With
        With Flex(0)
            .Row = 1
            .ScrollBars = flexScrollBarBoth
            .FixedRows = 1
            
            .Redraw = True
            SettaLargFlex Flex(0), 1, 2 * .ColWidth(0)
            .TopRow = 1
        End With
        With Flex(1)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .Cols = 2
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_Property)
            .TextMatrix(0, 1) = MyLoadResString(k_Value)
            
            .ColWidth(0) = 2 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0)
        End With
        LoadDBOption oDb
        
        sBuf = ""
        With oDb.DBOption
            For i = 0 To UBound(Elem)
                sBuf = ""
                If Elem(i).Elem.iType <> am_Bool Then
                    If Elem(i).EditType = iCombo Then
                        sBuf = GetComboDescription(Elem(i).Elem.iLen, Elem(i).iComboType)
                    End If
                End If
                Flex(1).AddItem Elem(i).Elem.sName & vbTab & sBuf
            
                
                Flex(1).Col = Flex(1).Cols - 1
                Flex(1).Row = Flex(1).Rows - 1
                
                If Elem(i).Elem.iType = am_Bool Then
                    Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(i).Elem.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex(1).CellPictureAlignment = flexAlignCenterCenter
                End If
                
                If Elem(i).EditType = iCombo Then
                    Set Flex(1).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                    Flex(1).CellPictureAlignment = flexAlignRightTop
                End If
                
                If Elem(i).bDisab Then Flex(1).CellBackColor = vbInactiveBorder
            Next
        End With
                    
        With Flex(1)
            .ColAlignment(1) = flexAlignLeftCenter
            .FixedRows = 1
            .FixedCols = 1
            .Row = 1
            .Redraw = True
            SettaLargFlex Flex(1), 0, 2 * (.Width / 3)
            .TopRow = 1
        End With
        Screen.MousePointer = vbDefault
    End If
    Set oDb = Nothing
    
End Sub

Private Sub LoadDBOption(db As SQLDMO.Database2)
        
    Dim iAccessType As am_sqlAccessType
    Dim i As Integer
    Dim bOffLine As Boolean
    
    Dim oOpt As SQLDMO.DBOption2
    Set oOpt = db.DBOption
                
    ReDim Elem(11) As elSetting
    With db.DBOption
        bOffLine = .Offline
        Elem(11).EditType = iNormalEdit
        Elem(11).Elem.sName = MyLoadResString(k_DBStatus_OffLine)
        Elem(11).Elem.vVal = bOffLine
        Elem(11).Elem.iType = am_Bool
    
        If .DBOUseOnly Then iAccessType = amAccessDBO
        If .SingleUser Then iAccessType = iAccessType + amAccessSingle
        Elem(0).Elem.sName = MyLoadResString(k_DbOpt_AccessGrantedTo) '"Accesso consentito a"
        Elem(0).Elem.iLen = iAccessType
        Elem(0).Elem.vVal = GetComboDescription(iAccessType, amRestricAccess)
        Elem(0).Elem.iType = am_String
        Elem(0).EditType = iCombo
        Elem(0).iComboType = amRestricAccess
        
        Elem(1).EditType = iNormalEdit
        Elem(1).Elem.sName = MyLoadResString(k_DBStatus_ReadOnly)
        Elem(1).Elem.vVal = .ReadOnly
        Elem(1).Elem.iType = am_Bool
        
                
        Elem(2).Elem.sName = MyLoadResString(k_DbOpt_RecoveryModel)
        If Not bOffLine Then Elem(2).Elem.iLen = oOpt.RecoveryModel
        Elem(2).Elem.vVal = GetComboDescription(Elem(2).Elem.iLen, amRecovery)
        Elem(2).Elem.iType = am_String
        Elem(2).EditType = iCombo
        Elem(2).iComboType = amRecovery
                    
        Elem(3).EditType = iNormalEdit
        Elem(3).Elem.sName = MyLoadResString(k_DbOpt_AnsiNullDefault) '"Valore Predefinito ANSI NULL"
        Elem(3).Elem.vVal = .ColumnsNullByDefault
        Elem(3).Elem.iType = am_Bool
                    
        Elem(4).EditType = iNormalEdit
        Elem(4).Elem.sName = MyLoadResString(k_DbOpt_RecursiveTriggers) '"Trigger Ricorsivi"
        Elem(4).Elem.vVal = .RecursiveTriggers
        Elem(4).Elem.iType = am_Bool
                
        Elem(5).EditType = iNormalEdit
        Elem(5).Elem.sName = MyLoadResString(k_DbOpt_AutoCreateStat) '"Creazione Automatica Statistiche"
        Elem(5).Elem.vVal = .AutoCreateStat
        Elem(5).Elem.iType = am_Bool

        Elem(6).EditType = iNormalEdit
        Elem(6).Elem.sName = MyLoadResString(k_DbOpt_AutoUpdateStat) '"Aggiornamento Automatico Statistiche"
        Elem(6).Elem.vVal = .AutoUpdateStat
        Elem(6).Elem.iType = am_Bool

        Elem(7).EditType = iNormalEdit
        Elem(7).Elem.sName = MyLoadResString(k_DbOpt_TornPageDetection) '"Rilevamento pagine incomplete"    '"TornPageDetection"
        Elem(7).Elem.vVal = .TornPageDetection
        Elem(7).Elem.iType = am_Bool

        Elem(8).EditType = iNormalEdit
        Elem(8).Elem.sName = MyLoadResString(k_DbOpt_AutoClose) '"Chiusura Automatica"
        Elem(8).Elem.vVal = .AutoClose
        Elem(8).Elem.iType = am_Bool

        Elem(9).EditType = iNormalEdit
        Elem(9).Elem.sName = MyLoadResString(k_DbOpt_AutoShrink) '"Compattazione Automatica"     'Auto Shrink
        Elem(9).Elem.vVal = .AutoShrink
        Elem(9).Elem.iType = am_Bool

        Elem(10).EditType = iNormalEdit
        Elem(10).Elem.sName = MyLoadResString(k_DbOpt_UseQuotedIdentifier) '"Usa Identificatori tra virgolette"     'Use quoted identifier
        Elem(10).Elem.vVal = .QuoteDelimiter
        Elem(10).Elem.iType = am_Bool

        Elem(11).EditType = iNormalEdit
        Elem(11).Elem.sName = MyLoadResString(k_DBStatus_OffLine)
        Elem(11).Elem.vVal = .Offline
        Elem(11).Elem.iType = am_Bool
        
        
        For i = 0 To UBound(Elem) - 1
            Elem(i).bDisab = bOffLine
        Next
    End With
    Set oOpt = Nothing

End Sub


Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub cbo_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyEscape
            m_blnAnnullaMod = True
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(miFlex).SetFocus
            KeyAscii = 0
    End Select

End Sub
Private Sub cbo_LostFocus()
    
    Dim iElem As Integer
    Dim iCol As Integer
    Dim iRow As Long
    Dim bUpdate As Boolean
        
    If Not m_blnAnnullaMod Then
        iElem = miRow
        iRow = miRow
        iCol = 1
        If miFlex = 1 Then
            If iActionMode <> act_Q_Options Then
                bUpdate = True
            End If
            
            If bUpdate Then
                If iActionMode <> act_Q_Options Then
                    Elem(iElem).Elem.iLen = cbo.ItemData(cbo.ListIndex)
                    Elem(iElem).Elem.vVal = cbo.Text
                End If
            End If
        End If
    End If
    If bUpdate Then
        
        Flex(miFlex).TextMatrix(iRow + 1, iCol) = cbo.Text
        bDirty = True
        cmd(1).Enabled = True
        
    Else
        If Not m_blnAnnullaMod Then Beep
    End If

    cbo.Visible = False
    Flex(miFlex).SetFocus
    m_blnAnnullaMod = True

End Sub

Private Sub cmd_Click(Index As Integer)
    
    TestActiveControl
    If Index = 0 Then
        Unload Me
    Else
        Select Case iActionMode
            Case act_DBOption
                If Index = 1 Then
                    If bDirty Then UpdateSQLDMO
                ElseIf Index = 2 Then
                    fNewDB.LoadDB m_sDbName
                    fNewDB.Show vbModal, Me
                End If
                
            Case act_User
                If bDirty Then
                    If Index = 1 Then UpdateUser
                End If
            Case act_Q_Options
                If bDirty Then
                    UpdateSetting
                    Unload Me
                End If
        End Select
    End If
    
End Sub
Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        
    If KeyCode = vbKeyF2 Then
        If (Index = 0 And (iActionMode = act_Extended_VIEW_Idx_Trig And Flex(Index).Col = Flex(Index).Cols - 1)) Then
            If Flex(Index).Row <> 0 Then ManageExtendedProp 0
        Else
            FlexEdit 0
        End If
    End If

End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            TestActiveControl
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            SizeControls
            b = False
        End If
    End If
End Sub
Private Sub Form_Load()
    m_lWidth = Me.Width
    m_lHeight = Me.Height

    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    cmd(2).Caption = MyLoadResString(k_Indexes_caption)
    
    CenterForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fGenOpt = Nothing
End Sub

Private Sub UpdateSQLDMO()

    Dim oDb As SQLDMO.Database2
    Dim oOpt As SQLDMO.DBOption2
    Dim iVal As Integer
    Dim lErr As Long, sErr As String
       
    bDirty = False
    Screen.MousePointer = vbHourglass
    Set oDb = objServer.Databases(m_sDbName)
    
    On Local Error GoTo ErrStat
    With oDb.DBOption
        If .Offline = False Then
            Set oOpt = oDb.DBOption
            If .DBOUseOnly Then iVal = amAccessDBO
            If .SingleUser Then iVal = iVal + amAccessSingle
            If iVal <> Elem(0).Elem.iLen Then
                .DBOUseOnly = Elem(0).Elem.iLen = amAccessDBO
                .SingleUser = Elem(0).Elem.iLen = amAccessSingle
            End If
            If Abs(.ReadOnly) <> Abs(Elem(1).Elem.vVal) Then .ReadOnly = Abs(Elem(1).Elem.vVal)
            
            If oOpt.RecoveryModel <> Elem(2).Elem.iLen Then oOpt.RecoveryModel = Elem(2).Elem.iLen
            If Abs(.ColumnsNullByDefault) <> Abs(Elem(3).Elem.vVal) Then .ColumnsNullByDefault = Abs(Elem(3).Elem.vVal)
            If Abs(.RecursiveTriggers) <> Abs(Elem(4).Elem.vVal) Then .RecursiveTriggers = Abs(Elem(4).Elem.vVal)
            If Abs(.AutoCreateStat) <> Abs(Elem(5).Elem.vVal) Then .AutoCreateStat = Abs(Elem(5).Elem.vVal)
            If Abs(.AutoUpdateStat) <> Abs(Elem(6).Elem.vVal) Then .AutoUpdateStat = Abs(Elem(6).Elem.vVal)
            If Abs(.TornPageDetection) <> Abs(Elem(7).Elem.vVal) Then .TornPageDetection = Abs(Elem(7).Elem.vVal)
            If Abs(.AutoClose) <> Abs(Elem(8).Elem.vVal) Then .AutoClose = Abs(Elem(8).Elem.vVal)
            If Abs(.AutoShrink) <> Abs(Elem(9).Elem.vVal) Then .AutoShrink = Abs(Elem(9).Elem.vVal)
            If Abs(.QuoteDelimiter) <> Abs(Elem(10).Elem.vVal) Then .QuoteDelimiter = Abs(Elem(10).Elem.vVal)
        End If
        
        If Abs(.Offline) <> Abs(Elem(UBound(Elem)).Elem.vVal) Then .Offline = Abs(Elem(UBound(Elem)).Elem.vVal)
    End With
    
    cmd(1).Enabled = False
    
ExitHere:
    Set oOpt = Nothing
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrStat:
    lErr = Err.Number
    sErr = Err.Description
    Screen.MousePointer = vbDefault
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    
    GoTo ExitHere
End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iImg As Integer, iLen As Integer, iElem As Integer
    Dim iTmpType As iPropG
    Dim bPush As Boolean, bCombo As Boolean
        
            
    
    If iActionMode = act_SP Or iActionMode = act_User_Function Or iActionMode = act_Tbl Or iActionMode = act_Extended_VIEW_Idx_Trig Or iActionMode = act_Extended_TABLE_Idx_Rel Or iActionMode = act_Extended_TABLE_Trig Then
        If Flex(1).Col <> Flex(1).Cols - 1 Then Exit Sub
    End If
    If Flex(1).CellBackColor = vbInactiveBorder Then Exit Sub
    
    If Not objServer.SaLogin Then
        Beep
        Exit Sub
    End If
    If sStuff = "dbo" Then
        Exit Sub
    End If
    miRow = Flex(1).Row - 1
    sBuf = Flex(1).Text
    
    bEdit = True

    iElem = Flex(1).Row - 1
    If iActionMode <> act_Q_Options Then
        If iActionMode = act_DBOption Then
            iTmpType = Elem(iElem).Elem.iType
            iLen = Elem(iElem).Elem.iLen
            bCombo = Elem(iElem).EditType = iCombo
        ElseIf iActionMode = act_User Then
            iTmpType = UserElem(iElem).iType
            iLen = UserElem(iElem).iLen
        ElseIf iActionMode = act_SP Or iActionMode = act_View Or iActionMode = act_User_Function Or iActionMode = act_Tbl Or iActionMode = act_Extended_VIEW_Idx_Trig Or iActionMode = act_Extended_TABLE_Idx_Rel Or iActionMode = act_Extended_TABLE_Trig Then
            If KeyAscii = 0 Or KeyAscii = vbKeySpace Then ManageExtendedProp 1
            bEdit = False
        End If
    Else
        iTmpType = ElemSetting(iElem).Elem.iType
        iLen = ElemSetting(iElem).Elem.iLen
        bPush = ElemSetting(iElem).EditType = iEllipsis
    End If
    
    If Flex(1).CellBackColor = vbInactiveBorder Then bEdit = False
    
    If iElem < 0 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                
                If (iTmpType And am_Bool) = am_Bool Then
                    If iActionMode = act_DBOption Then
                        iImg = 1 - (-CBool(Elem(iElem).Elem.vVal))
                        Elem(iElem).Elem.vVal = iImg
                        
                        Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        Flex(1).CellPictureAlignment = flexAlignCenterCenter
                        
                        'offline
                        If iElem = UBound(Elem) Then
                            Flex(1).Col = Flex(1).Cols - 1
                            For iLen = 0 To iElem - 1
                                Flex(1).Row = iLen + 1
                                Elem(iLen).bDisab = CBool(iImg)
                                Flex(1).Row = iLen + 1
                                Flex(1).CellBackColor = IIf(CBool(iImg), vbInactiveBorder, vbWindowBackground)
                            Next
                            Flex(1).Row = iElem + 1
                        End If
                        bDirty = True
                        cmd(1).Enabled = True
                    
                    ElseIf iActionMode = act_User Then
                        iImg = 1 - (-CBool(UserElem(iElem).vVal))
                        UserElem(iElem).vVal = iImg
                        
                        Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        Flex(1).CellPictureAlignment = flexAlignCenterCenter
                        bDirty = True
                        cmd(1).Enabled = True
                    
                    End If
                    bEdit = False
                End If
            Case Else
                If (iTmpType And am_Bool) = am_Bool Then bEdit = False
        End Select
    End If
    
    If bEdit Then
        If Not bPush And Not bCombo Then
            tFlex.MaxLength = 0
            tFlex.Text = ""
            tFlex.MaxLength = iLen
            tFlex.Text = sBuf
    
            tFlex.SelStart = 0
            With Flex(1)
                tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
            End With
            tFlex.Visible = True
            tFlex.ZOrder
            tFlex.SetFocus
    
            If (iTmpType And am_String) = am_String Then
                If KeyAscii = vbKeySpace Then KeyAscii = 0
            End If
            On Local Error Resume Next
            SendKeys Chr$(KeyAscii)
            On Local Error GoTo 0
        ElseIf bPush Then
            LoadFont
        ElseIf bCombo Then
            PreloadCBO Elem(iElem).iComboType
            cbo.ListIndex = GetItem(Elem(iElem).Elem.vVal, cbo)
                        
            With Flex(1)
                cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            End With
            cbo.Visible = True
            cbo.ZOrder
            cbo.SetFocus
                        
            If (iTmpType And am_String) = am_String Then
                If KeyAscii = vbKeySpace Then KeyAscii = 0
            End If
            On Local Error Resume Next
            SendKeys Chr$(KeyAscii)
            On Local Error GoTo 0
        End If
    End If
End Sub

Private Sub Flex_DblClick(Index As Integer)
    
    If Index = 1 Then
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then
            FlexEdit vbKeySpace
        End If
    Else
        If (Index = 0 And (iActionMode = act_Extended_VIEW_Idx_Trig And Flex(Index).Col = Flex(Index).Cols - 1) Or (iActionMode = act_Extended_TABLE_Idx_Rel And Flex(Index).Col = Flex(Index).Cols - 1) Or (iActionMode = act_Extended_TABLE_Idx_Rel And Flex(Index).Col = Flex(Index).Cols - 1)) Then
            If Flex(Index).Row <> 0 Then ManageExtendedProp 0
        End If
    End If
End Sub
Private Sub Flex_GotFocus(Index As Integer)
    TestActiveControl
    tFlex.Visible = False
    cbo.Visible = False
    m_blnAnnullaMod = True
    miFlex = Index
End Sub
Private Sub Flex_Scroll(Index As Integer)
    tFlex.Visible = False
    cbo.Visible = False
End Sub
Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    
    If Index = 1 Then
        iCode = KeyAscii
        KeyAscii = 0
                
        FlexEdit iCode
    Else
        If (Index = 0 And (iActionMode = act_Extended_VIEW_Idx_Trig And Flex(Index).Col = Flex(Index).Cols - 1)) Then
            If KeyAscii = vbKeySpace Then
                If Flex(Index).Row <> 0 Then ManageExtendedProp 0
            End If
        End If
    End If
    
End Sub
Private Sub LoadOptions()
    
    Dim i As Integer, sBuf As String
    Dim bBool As Boolean
        
    Screen.MousePointer = vbHourglass
    Flex(0).Visible = False
    Flex(1).Top = Flex(0).Top
    Flex(1).Height = Me.ScaleHeight - (cmd(0).Height + 200)

    ReDim ElemSetting(3) As elSetting

    ElemSetting(0).Elem.sName = MyLoadResString(k_Font_Qry)
    ElemSetting(0).EditType = iEllipsis
    ElemSetting(0).Elem.vVal = ""
    ElemSetting(0).Elem.iType = am_String
    ElemSetting(0).Elem.iLen = 0
    
    ElemSetting(1).Elem.sName = MyLoadResString(k_Font_Grd)
    ElemSetting(1).EditType = iEllipsis
    ElemSetting(1).Elem.vVal = ""
    ElemSetting(1).Elem.iType = am_String
    ElemSetting(1).Elem.iLen = 0
    
    ElemSetting(2).Elem.sName = MyLoadResString(k_Font_RES)
    ElemSetting(2).EditType = iEllipsis
    ElemSetting(2).Elem.vVal = ""
    ElemSetting(2).Elem.iType = am_String
    ElemSetting(2).Elem.iLen = 0
    
    ElemSetting(3).Elem.sName = MyLoadResString(k_Font_GoStm)
    ElemSetting(3).Elem.vVal = "GO"
    ElemSetting(3).Elem.iType = am_String
    ElemSetting(3).Elem.iLen = 15
        
        
    sBuf = Trim$(GetIni("QUERY", "FONT-QRY", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then ElemSetting(0).Elem.vVal = sBuf
    
    sBuf = Trim$(GetIni("QUERY", "FONT-GRD", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then ElemSetting(1).Elem.vVal = sBuf
    
    sBuf = Trim$(GetIni("QUERY", "FONT-RES", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then ElemSetting(2).Elem.vVal = sBuf
    
    sBuf = Trim$(GetIni("QUERY", "TERM", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then ElemSetting(3).Elem.vVal = sBuf
            
    With Flex(1)
        .Redraw = False
        
        .Clear
        .FixedCols = 0
        .Cols = 2
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 0) = MyLoadResString(k_Property)
        .TextMatrix(0, 1) = MyLoadResString(k_Value)
        
        .ColWidth(0) = 2 * (.Width / 3)
        .ColWidth(1) = .Width - .ColWidth(0)
    End With

    sBuf = ""

    For i = 0 To UBound(ElemSetting)
        Select Case VarType(ElemSetting(i).Elem.vVal)
            Case vbString
                sBuf = ElemSetting(i).Elem.vVal
            Case vbInteger, vbLong
                sBuf = Format$(ElemSetting(i).Elem.vVal, "#,0")
            Case vbSingle, vbDouble, vbDecimal
                sBuf = Format$(ElemSetting(i).Elem.vVal, "#,0.0000")
            Case vbDate
                sBuf = Format$(ElemSetting(i).Elem.vVal, "dd/mm/yyyy")
            Case vbBoolean, vbByte
                bBool = True
                sBuf = ""
        End Select
        Flex(1).AddItem ElemSetting(i).Elem.sName & vbTab & sBuf
        
        If ElemSetting(i).bDisab Then
            Flex(1).Col = Flex(1).Cols - 1
            Flex(1).Row = Flex(1).Rows - 1
            Flex(1).CellBackColor = vbInactiveBorder
        End If
        
        If bBool Then
            Flex(1).Col = Flex(1).Cols - 1
            Flex(1).Row = Flex(1).Rows - 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(ElemSetting(i).Elem.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex(1).CellPictureAlignment = flexAlignCenterCenter
        End If
        bBool = False
        
        If ElemSetting(i).EditType = iEllipsis Then
            Flex(1).Col = Flex(1).Cols - 1
            Flex(1).Row = Flex(1).Rows - 1
            Set Flex(1).CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
            Flex(1).CellPictureAlignment = flexAlignRightBottom
        End If
    Next

                    
    With Flex(1)
        .ColAlignment(1) = flexAlignLeftCenter
        .FixedRows = 1
        .FixedCols = 1
        .Row = 1
        .Redraw = True
        .TopRow = 1
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LoadUser()
    
    Dim oUser As SQLDMO.User
    Dim oDb As SQLDMO.Database2
    Dim i As Integer, sBuf As String
    Dim bBool As Boolean
    
        
    Screen.MousePointer = vbHourglass
    Flex(0).Visible = False
    Flex(1).Top = Flex(0).Top
    Flex(1).Height = Me.ScaleHeight - (cmd(0).Height + 200)

    ReDim UserElem(11) As PageElement

    UserElem(11).sName = "Public"
    UserElem(11).vVal = True
    UserElem(11).iType = am_Bool
    
    UserElem(0).sName = "Id"
    UserElem(0).vVal = 0
    UserElem(0).iType = am_Long
    UserElem(0).iLen = 2
    
    UserElem(1).sName = MyLoadResString(k_Login) '"Login"
    UserElem(1).vVal = ""
    UserElem(1).iType = am_String
    UserElem(1).iLen = 10
    
    UserElem(2).sName = MyLoadResString(k_Name) '"Name"
    UserElem(2).vVal = ""
    UserElem(2).iType = am_String
    UserElem(2).iLen = 10

    UserElem(3).sName = "db_owner"
    UserElem(3).vVal = False
    UserElem(3).iType = am_Bool

    UserElem(4).sName = "db_accessadmin"
    UserElem(4).vVal = False
    UserElem(4).iType = am_Bool
    
    UserElem(5).sName = "db_backupoperator"
    UserElem(5).vVal = False
    UserElem(5).iType = am_Bool
    
    UserElem(6).sName = "db_datareader"
    UserElem(6).vVal = False
    UserElem(6).iType = am_Bool
    
    UserElem(7).sName = "db_datawriter"
    UserElem(7).vVal = False
    UserElem(7).iType = am_Bool

    UserElem(8).sName = "db_ddladmin"
    UserElem(8).vVal = False
    UserElem(8).iType = am_Bool

    UserElem(9).sName = "db_denydatareader"
    UserElem(9).vVal = False
    UserElem(9).iType = am_Bool

    UserElem(10).sName = "db_denydatawriter"
    UserElem(10).vVal = False
    UserElem(10).iType = am_Bool
    
    If Len(sStuff) Then
        Set oDb = objServer.Databases(m_sDbName)
        Set oUser = oDb.Users(sStuff)

        With oUser
            UserElem(0).vVal = .Id
            UserElem(1).vVal = .Login
            UserElem(2).vVal = .Name
        
            UserElem(3).vVal = .IsMember("db_owner")
            UserElem(4).vVal = .IsMember("db_accessadmin")
            UserElem(5).vVal = .IsMember("db_backupoperator")
            UserElem(6).vVal = .IsMember("db_datareader")
            UserElem(7).vVal = .IsMember("db_datawriter")
            UserElem(8).vVal = .IsMember("db_ddladmin")
            UserElem(9).vVal = .IsMember("db_denydatareader")
            UserElem(10).vVal = .IsMember("db_denydatawriter")
        End With
    End If
      
    With Flex(1)
        .Redraw = False
        
        .Clear
        .FixedCols = 0
        .Cols = 2
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 0) = MyLoadResString(k_Property)
        .TextMatrix(0, 1) = MyLoadResString(k_Value)
        
        .ColWidth(0) = 2 * (.Width / 3)
        .ColWidth(1) = .Width - .ColWidth(0)
    End With

    sBuf = ""

    For i = 0 To UBound(UserElem)
        Debug.Print UserElem(i).sName; UserElem(i).vVal
        Select Case VarType(UserElem(i).vVal)
            Case vbString
                sBuf = UserElem(i).vVal
            Case vbInteger, vbLong
                sBuf = Format$(UserElem(i).vVal, "#,0")
            Case vbSingle, vbDouble, vbDecimal
                sBuf = Format$(UserElem(i).vVal, "#,0.0000")
            Case vbDate
                sBuf = Format$(UserElem(i).vVal, "dd/mm/yyyy")
            Case vbBoolean, vbByte
                bBool = True
                sBuf = ""
        End Select
        Flex(1).AddItem UserElem(i).sName & vbTab & sBuf
        If i = 0 Then
            Flex(1).Col = Flex(1).Cols - 1
            Flex(1).Row = Flex(1).Rows - 1
            Flex(1).CellBackColor = vbInactiveBorder
        ElseIf i < 3 Then
            If Len(sStuff) <> 0 Then
                Flex(1).Col = Flex(1).Cols - 1
                Flex(1).Row = Flex(1).Rows - 1
                Flex(1).CellBackColor = vbInactiveBorder
            End If
        End If
        If bBool Then
            Flex(1).Col = Flex(1).Cols - 1
            Flex(1).Row = Flex(1).Rows - 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(UserElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex(1).CellPictureAlignment = flexAlignCenterCenter
            
            If i = 11 Then Flex(1).CellBackColor = vbInactiveBorder
        End If
        bBool = False
    Next

                
    With Flex(1)
        .ColAlignment(1) = flexAlignLeftCenter
        .FixedRows = 1
        .FixedCols = 1
        .Row = 1
        .Redraw = True
        .TopRow = 1
    End With
    Screen.MousePointer = vbDefault

    Set oUser = Nothing
    Set oDb = Nothing
    cmd(1).Enabled = True

End Sub

Private Sub UpdateUser()

    Dim lErr As Long, sErr As String
    Dim oUser As SQLDMO.User
    Dim oDb As SQLDMO.Database2
    Dim bIsGuest As Boolean
    
    bIsGuest = Len(Trim$(UserElem(1).vVal)) = 0 And (StrComp(Trim$(UserElem(2).vVal), "guest", vbTextCompare) = 0)
    If Not bIsGuest Then
        If Len(Trim$(UserElem(1).vVal)) = 0 Or Len(Trim$(UserElem(2).vVal)) = 0 Then
            MsgBox MyLoadResString(kErr_Insuf_Data), vbInformation Or vbOKOnly, App.EXEName
            Exit Sub
        End If
    End If
    
    
    bDirty = False
    cmd(1).Enabled = False
    Screen.MousePointer = vbHourglass
    
    On Local Error GoTo ErrStat
    Set oDb = objServer.Databases(m_sDbName)
    
    If Len(sStuff) Then
        Set oUser = oDb.Users(sStuff)
    Else
        Set oUser = New SQLDMO.User
    End If
    
    With oUser
        If Len(sStuff) = 0 Then
            If Not bIsGuest Then .Login = Trim$(UserElem(1).vVal)
            .Name = Trim$(UserElem(2).vVal)
            oDb.Users.Add oUser
            sStuff = Trim$(UserElem(2).vVal)
            
            If UserElem(3).vVal Then oDb.DatabaseRoles("db_owner").AddMember sStuff
            If UserElem(4).vVal Then oDb.DatabaseRoles("db_accessadmin").AddMember sStuff
            If UserElem(5).vVal Then oDb.DatabaseRoles("db_backupoperator").AddMember sStuff
            If UserElem(6).vVal Then oDb.DatabaseRoles("db_datareader").AddMember sStuff
            If UserElem(7).vVal Then oDb.DatabaseRoles("db_datawriter").AddMember sStuff
            If UserElem(8).vVal Then oDb.DatabaseRoles("db_ddladmin").AddMember sStuff
            If UserElem(9).vVal Then oDb.DatabaseRoles("db_denydatareader").AddMember sStuff
            If UserElem(10).vVal Then oDb.DatabaseRoles("db_denydatawriter").AddMember sStuff

        Else
            If Not sStuff = "dbo" Then
                If UserElem(3).vVal = False Then
                    If .IsMember("db_owner") Then oDb.DatabaseRoles("db_owner").DropMember sStuff
                Else
                    If Not .IsMember("db_owner") Then oDb.DatabaseRoles("db_owner").AddMember sStuff
                End If
                If UserElem(4).vVal = False Then
                    If .IsMember("db_accessadmin") Then oDb.DatabaseRoles("db_accessadmin").DropMember sStuff
                Else
                    If Not .IsMember("db_accessadmin") Then oDb.DatabaseRoles("db_accessadmin").AddMember sStuff
                End If
                If UserElem(5).vVal = False Then
                    If .IsMember("db_backupoperator") Then oDb.DatabaseRoles("db_backupoperator").DropMember sStuff
                Else
                    If Not .IsMember("db_backupoperator") Then oDb.DatabaseRoles("db_backupoperator").AddMember sStuff
                End If
                If UserElem(6).vVal = False Then
                    If .IsMember("db_datareader") Then oDb.DatabaseRoles("db_datareader").DropMember sStuff
                Else
                    If Not .IsMember("db_datareader") Then oDb.DatabaseRoles("db_datareader").AddMember sStuff
                End If
                If UserElem(7).vVal = False Then
                    If .IsMember("db_datawriter") Then oDb.DatabaseRoles("db_datawriter").DropMember sStuff
                Else
                    If Not .IsMember("db_datawriter") Then oDb.DatabaseRoles("db_datawriter").AddMember sStuff
                End If
                If UserElem(8).vVal = False Then
                    If .IsMember("db_ddladmin") Then oDb.DatabaseRoles("db_ddladmin").DropMember sStuff
                Else
                    If Not .IsMember("db_ddladmin") Then oDb.DatabaseRoles("db_ddladmin").AddMember sStuff
                End If
                If UserElem(9).vVal = False Then
                    If .IsMember("db_denydatareader") Then oDb.DatabaseRoles("db_denydatareader").DropMember sStuff
                Else
                    If Not .IsMember("db_denydatareader") Then oDb.DatabaseRoles("db_denydatareader").AddMember sStuff
                End If
                If UserElem(10).vVal = False Then
                    If .IsMember("db_denydatawriter") Then oDb.DatabaseRoles("db_denydatawriter").DropMember sStuff
                Else
                    If Not .IsMember("db_denydatawriter") Then oDb.DatabaseRoles("db_denydatawriter").AddMember sStuff
                End If
            End If
        End If
    End With
    
ExitHere:
    Set oUser = Nothing
    Set oDb = Nothing

    If lErr = 0 Then LoadUser
    Screen.MousePointer = vbDefault
    Exit Sub
ErrStat:
    lErr = Err.Number
    sErr = Err.Description
    Screen.MousePointer = vbDefault
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    LoadUser
    GoTo ExitHere
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iElem As Integer
    Dim iTmpType As iPropG
    
    iElem = miRow
    
    If miFlex = 1 Then
        If iActionMode = act_DBOption Then
            iTmpType = Elem(iElem).Elem.iType
        ElseIf iActionMode = act_User Then
            iTmpType = UserElem(iElem).iType
        ElseIf iActionMode = act_Q_Options Then
            iTmpType = ElemSetting(iElem).Elem.iType
        End If
    End If
    
    
    If KeyCode = vbKeyUP Then
        KeyCode = 0
        If miRow > 0 Then
            Flex(miFlex).Row = miRow
            Flex(miFlex).SetFocus
        End If
    ElseIf KeyCode = vbKeyDOWN Then
        If miRow < Flex(miFlex).Rows - 2 Then
            Flex(miFlex).Row = miRow + 2
            Flex(miFlex).SetFocus
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
            If miFlex = 0 And miRow = 1 Then
                'Input_KDown tFlex, KeyCode
            End If
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Dim iTmpType As iPropG
    Dim iElem As Integer
    
    iElem = miRow
    If miFlex = 1 Then
        If iActionMode = act_DBOption Then
            iTmpType = Elem(iElem).Elem.iType
        ElseIf iActionMode = act_User Then
            iTmpType = UserElem(iElem).iType
        ElseIf iActionMode = act_Q_Options Then
            iTmpType = ElemSetting(iElem).Elem.iType
        End If
    End If
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93    '<>|[]
            KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex(miFlex).Text
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(miFlex).SetFocus
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
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
    
    Dim iElem As Integer
    Dim iCol As Integer
    Dim iRow As Long
    Dim bUpdate As Boolean
    
    If Not m_blnAnnullaMod Then
        iElem = miRow
        iRow = miRow
        iCol = 1
            
        tFlex.Text = Trim$(tFlex)
        If miFlex = 1 Then
            If iActionMode <> act_Q_Options Then
                Select Case miRow
                    Case 0
                    Case 1, 2
                        If Len(tFlex) Then bUpdate = True
                End Select
            Else
                bUpdate = Len(tFlex.Text)
            End If
            
            If bUpdate Then
                If iActionMode = act_DBOption Then
                    Elem(iElem).Elem.vVal = tFlex.Text
                ElseIf iActionMode = act_User Then
                    UserElem(iElem).vVal = tFlex.Text
                ElseIf iActionMode = act_Q_Options Then
                    ElemSetting(iElem).Elem.vVal = tFlex.Text
                End If
            
            End If
        End If
    End If
    If bUpdate Then
        Debug.Print Flex(1).Rows
        Debug.Print Flex(1).Cols
        
        Flex(miFlex).TextMatrix(iRow + 1, iCol) = tFlex
        bDirty = True
        cmd(1).Enabled = True
        
    Else
        If Not m_blnAnnullaMod Then Beep
    End If

    tFlex.Visible = False
    Flex(miFlex).SetFocus
    m_blnAnnullaMod = True
    
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub SizeControls()
    
    Dim l As Long, i As Integer
    l = Me.ScaleWidth \ 4 - cmd(0).Width
    cmd(0).Top = Me.ScaleHeight - (cmd(0).Height + 50)
    
    For i = 1 To 2
        cmd(i).Top = cmd(0).Top
        cmd(i).Left = cmd(0).Left + i * (cmd(0).Width + 25)
    Next
            
    If Flex(0).Visible Then
        Flex(0).Height = (cmd(0).Top - 50 - Flex(0).Top) \ 2
        Flex(1).Height = Flex(0).Height
            
        Flex(0).Width = Me.ScaleWidth - (2 * Flex(0).Left)
        Flex(1).Top = Flex(0).Height + 25 + Flex(0).Top
    Else
        Flex(1).Top = Flex(0).Top
        Flex(1).Height = cmd(0).Top - 50 - Flex(1).Top
    End If
    Flex(1).Width = Me.ScaleWidth - (2 * Flex(1).Left)
    
    If Flex(0).Visible Then SettaLargFlex Flex(0), 1, 0
    Select Case iActionMode
        Case act_Tbl, act_View
            i = 2
        Case act_SP, act_User_Function
            i = 1
        Case Else
            i = 0
    End Select
    SettaLargFlex Flex(1), i, 0
            
End Sub
Private Sub TestActiveControl()

    If tFlex.Visible Then
        tFlex_LostFocus
    ElseIf cbo.Visible Then
        cbo_LostFocus
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

Private Sub LoadTable()

    Dim sBuf As String
    Dim oTb As SQLDMO.Table
    Dim oCol As SQLDMO.Column2
    Dim oDb As SQLDMO.Database2
    Dim i As Integer
        
    Screen.MousePointer = vbHourglass
    Flex(0).Visible = True
    Flex(1).Top = Flex(0).Top
    Flex(1).Height = Me.ScaleHeight - (cmd(0).Height + 200)

    Dim TbElem(4) As PageElement
    Dim ColElem() As ColProperty
    

    TbElem(0).sName = MyLoadResString(k_Table)
    TbElem(0).iType = am_String
    
    TbElem(1).sName = MyLoadResString(k_Owner)
    TbElem(1).iType = am_String
    
    TbElem(2).sName = MyLoadResString(k_Create_Date)
    TbElem(2).iType = am_Date
        
    TbElem(3).sName = "FileGroup"
    TbElem(3).iType = am_String
    
    TbElem(4).sName = MyLoadResString(k_Rows)
    TbElem(4).iType = am_Long

    Dim sColHeader(8) As String
    sColHeader(0) = MyLoadResString(k_In_Primary_Key) '"Key"
    sColHeader(1) = MyLoadResString(k_Param_ColId)
    sColHeader(2) = MyLoadResString(k_Field_Name) '"Name"
    sColHeader(3) = MyLoadResString(k_Data_Type) '"DataType"
    sColHeader(4) = MyLoadResString(k_Size) '"Size"
    sColHeader(5) = MyLoadResString(k_Allow_Null) '"Nulls"
    sColHeader(6) = MyLoadResString(k_Default_Value) '"Default"
    sColHeader(7) = MyLoadResString(k_NewDbSort)  'Collation
    sColHeader(8) = MyLoadResString(k_Extended_Property)
                
    If Len(sStuff) Then
        Set oDb = objServer.Databases(m_sDbName)
        
        Set oTb = GetTbByName(objServer, m_sDbName, sStuff)
        'Set oTB = oDB.Tables(sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnTable)
        
        TbElem(0).vVal = oTb.Name
        TbElem(1).vVal = oTb.Owner
        TbElem(2).vVal = oTb.CreateDate
        TbElem(3).vVal = oTb.FileGroup
        TbElem(4).vVal = oTb.Rows
        
        i = -1
        For Each oCol In oTb.Columns
            i = i + 1
            ReDim Preserve ColElem(i) As ColProperty
            With oCol
                ColElem(i).Key = .InPrimaryKey
                ColElem(i).Id = .Id
                ColElem(i).Name = .Name
                ColElem(i).DataType = .DataType
                ColElem(i).Size = .Length
                ColElem(i).Nulls = .AllowNulls
                ColElem(i).Default = .DRIDefault.Text
                ColElem(i).Collation = .Collation
            End With
        Next
    End If
                    
    With Flex(0)
        .Redraw = False
            
        .Clear
        .FixedCols = 0
        .Cols = 2
        .Rows = 1
        
        .TextMatrix(0, 0) = MyLoadResString(k_Property)
        .TextMatrix(0, 1) = MyLoadResString(k_Value)
            
        .ColWidth(0) = (.Width / 3)
        For i = 0 To UBound(TbElem)
            sBuf = TbElem(i).sName & vbTab & TbElem(i).vVal
            .AddItem sBuf
        Next
        .FixedRows = 1
        .Redraw = True
        SettaLargFlex Flex(0), 1, 200
    End With

    LoadColumns sColHeader, ColElem

    Screen.MousePointer = vbDefault
    Set oCol = Nothing
    Set oTb = Nothing
    Set oDb = Nothing
    
End Sub

Private Sub UpdateSetting()

    bDirty = False
    cmd(1).Enabled = False
    Screen.MousePointer = vbHourglass
        
    Call WriteIni("QUERY", "FONT-QRY", CStr(ElemSetting(0).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-GRD", CStr(ElemSetting(1).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-RES", CStr(ElemSetting(2).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "TERM", CStr(ElemSetting(3).Elem.vVal), App.Path & "\" & App.EXEName & ".INI")

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LoadFont()

    Dim sFontName As String, sgFontSize As Single
    Dim iPos As Integer
    Dim sFont As SelectedFont
    Dim oFont As StdFont
    
    sFontName = ElemSetting(miRow).Elem.vVal
    iPos = InStr(sFontName, "\")
    If iPos <> 0 Then
        sgFontSize = ValF(Mid$(sFontName, iPos + 1))
        sFontName = Left$(sFontName, iPos - 1)
    Else
        Beep
        Exit Sub
    End If
    
    On Error GoTo e_Trap
    'FontDialog.iPointSize = 12 * 10
    FontDialog.iPointSize = sgFontSize * 10
    sFont = ShowFont(Me.hWnd, sFontName)
    
    If Not sFont.bCanceled Then
    
        Set oFont = New StdFont
    
        oFont.Name = sFont.sSelectedFont
        oFont.Size = sFont.nSize
                
        Flex(1).TextMatrix(Flex(1).Row, Flex(1).Cols - 1) = oFont.Name & "\" & oFont.Size
        ElemSetting(miRow).Elem.vVal = oFont.Name & "\" & oFont.Size
        
        Set oFont = Nothing
        bDirty = True
    End If
    
    Exit Sub
e_Trap:
    Exit Sub
        
End Sub
Private Function GetComboDescription(ByVal iVal As Integer, ByVal iOpt As amCbo2Read) As String

    Dim sBuf As String
    If iOpt = amRestricAccess Then
        If iVal = amAccessAll Then
            sBuf = MyLoadResString(k_DbOpt_Access_ALL) '"All"
        ElseIf iVal = amAccessDBO Then
            sBuf = MyLoadResString(k_DbOpt_Access_Dbo)  '"Dbo"
        ElseIf iVal = amAccessSingle Then
            sBuf = MyLoadResString(k_DBStatus_SingleUser) '"Single"
        End If
    ElseIf iOpt = amRecovery Then
        If iVal = SQLDMORECOVERY_Simple Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Simple) '"Semplice"
        ElseIf iVal = SQLDMORECOVERY_BulkLogged Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Bulklogged) '"Con Registrazioni di Massa"
        ElseIf iVal = SQLDMORECOVERY_Full Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Full) '"Completo"
        End If
    End If
    
    GetComboDescription = sBuf
End Function
Private Sub PreloadCBO(ByVal iOpt As amCbo2Read)

    Dim i As Integer
    Dim sBuf As String
    
    With cbo
        .Clear
        If iOpt = amRecovery Then
            For i = SQLDMORECOVERY_Simple To SQLDMORECOVERY_Full
                sBuf = GetComboDescription(i, amRecovery)
                .AddItem sBuf
                .ItemData(.NewIndex) = i
            Next
        ElseIf iOpt = amRestricAccess Then
            For i = amAccessAll To amAccessSingle
                sBuf = GetComboDescription(i, amRestricAccess)
                .AddItem sBuf
                .ItemData(.NewIndex) = i
            Next
        End If
    End With
    
End Sub
Private Sub LoadFuncParameter()
    
    Dim oFunc As SQLDMO.UserDefinedFunction
    
    Dim i As Integer, sBuf As String
    Dim y As Integer
    Dim bBool As Boolean
    Dim oQry As SQLDMO.QueryResults
    Dim vVar As Variant
    Dim vVarOut As Variant
    Dim tmpVar As Variant
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oFunc = GetFuncByName(objServer, m_sDbName, sStuff)
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnFunction)
            
        ReDim UserElem(5) As PageElement
        
        UserElem(0).sName = MyLoadResString(k_Param_Encrypted)
        UserElem(0).vVal = oFunc.Encrypted
        UserElem(0).iType = am_Bool
            
        UserElem(1).sName = MyLoadResString(k_Func_Param_IsDeterministic)
        UserElem(1).vVal = oFunc.IsDeterministic
        UserElem(1).iType = am_Bool
        
        UserElem(2).sName = MyLoadResString(k_Func_Param_IsSchemaBound)
        UserElem(2).vVal = oFunc.IsSchemaBound
        UserElem(2).iType = am_Bool
        
        UserElem(3).sName = MyLoadResString(k_Param_QuotedIdentifier)
        UserElem(3).vVal = oFunc.QuotedIdentifierStatus
        UserElem(3).iType = am_Bool
        
        UserElem(4).sName = MyLoadResString(k_Type)
        UserElem(4).vVal = UDFtype(oFunc.Type)
        UserElem(4).iType = am_String
        
        UserElem(5).sName = MyLoadResString(k_Param_ANSI_nulls)
        UserElem(5).vVal = oFunc.AnsiNullsStatus
        UserElem(5).iType = am_String
    
      
        With Flex(0)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 2
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_Property)
            .TextMatrix(0, 1) = MyLoadResString(k_Value)
            
            .ColWidth(0) = 2 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0)
            
    
            sBuf = ""
            For i = 0 To UBound(UserElem)
                Debug.Print UserElem(i).sName; UserElem(i).vVal
                Select Case VarType(UserElem(i).vVal)
                    Case vbString
                        sBuf = UserElem(i).vVal
                    Case vbInteger, vbLong
                        sBuf = Format$(UserElem(i).vVal, "#,0")
                    Case vbSingle, vbDouble, vbDecimal
                        sBuf = Format$(UserElem(i).vVal, "#,0.0000")
                    Case vbDate
                        sBuf = Format$(UserElem(i).vVal, "dd/mm/yyyy")
                    Case vbBoolean, vbByte
                        bBool = True
                        sBuf = ""
                End Select
                Flex(0).AddItem UserElem(i).sName & vbTab & sBuf
                If bBool Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(UserElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                bBool = False
            Next
            
            .FixedCols = 1
            .FixedRows = 1
            .Row = 1
            .TopRow = 1
            .Redraw = True
            SettaLargFlex Flex(0), 1, 200
        End With

        Set oQry = oFunc.EnumOutputs
        If Not oQry Is Nothing Then vVarOut = Qry2Var(oQry)
        If IsArray(vVarOut) Then
            i = UBound(vVarOut, 2)
            'ReDim Preserve vVarOut(5, i) As Variant
            ReDim tmpVar(5, i) As Variant
            For y = 0 To i
                tmpVar(0, y) = vVarOut(0, y)
                tmpVar(1, y) = vVarOut(1, y)
                tmpVar(2, y) = vVarOut(2, y)
                tmpVar(3, y) = vVarOut(3, y)
                tmpVar(4, y) = ""
                tmpVar(5, y) = False
            Next
        End If
        vVarOut = tmpVar
        
        Set oQry = oFunc.EnumParameters
        If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
        Set oQry = Nothing
        
        If IsArray(vVarOut) Then
            If IsArray(vVar) Then
                vVar = AddVar2Var(vVarOut, vVar)
            Else
                vVar = vVarOut
            End If
        End If
        
        For i = 0 To UBound(vVar, 2)
            If Len(vVar(5, i)) = 0 Then vVar(5, i) = True
        Next
        
        LoadFlexParam vVar
    
    End If
    Screen.MousePointer = vbDefault

    Set oFunc = Nothing

End Sub

Private Sub LoadSpParameter()
    
    Dim oProc As SQLDMO.StoredProcedure2
    Dim i As Integer, sBuf As String
    Dim bBool As Boolean
    Dim oQry As SQLDMO.QueryResults
    Dim vVar As Variant
    Dim vVarOut As Variant
    
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oProc = GetSpByName(objServer, m_sDbName, sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnSP)
            
        ReDim UserElem(3) As PageElement
        
        UserElem(0).sName = MyLoadResString(k_Param_Encrypted)
        UserElem(0).vVal = oProc.Encrypted
        UserElem(0).iType = am_Bool
        
        UserElem(1).sName = MyLoadResString(k_Param_QuotedIdentifier)
        UserElem(1).vVal = oProc.QuotedIdentifierStatus
        UserElem(1).iType = am_Bool
        
        UserElem(2).sName = MyLoadResString(k_Type)
        UserElem(2).vVal = StoredProcType(oProc.Type)
        UserElem(2).iType = am_String
        
        UserElem(3).sName = MyLoadResString(k_Param_ANSI_nulls)
        UserElem(3).vVal = oProc.AnsiNullsStatus
        UserElem(3).iType = am_String
    
      
        With Flex(0)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 2
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_Property)
            .TextMatrix(0, 1) = MyLoadResString(k_Value)
            
            .ColWidth(0) = 2 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0)
            
    
            sBuf = ""
            For i = 0 To UBound(UserElem)
                Debug.Print UserElem(i).sName; UserElem(i).vVal
                Select Case VarType(UserElem(i).vVal)
                    Case vbString
                        sBuf = UserElem(i).vVal
                    Case vbInteger, vbLong
                        sBuf = Format$(UserElem(i).vVal, "#,0")
                    Case vbSingle, vbDouble, vbDecimal
                        sBuf = Format$(UserElem(i).vVal, "#,0.0000")
                    Case vbDate
                        sBuf = Format$(UserElem(i).vVal, "dd/mm/yyyy")
                    Case vbBoolean, vbByte
                        bBool = True
                        sBuf = ""
                End Select
                Flex(0).AddItem UserElem(i).sName & vbTab & sBuf
                If bBool Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(UserElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                bBool = False
            Next
            
            .FixedCols = 1
            .FixedRows = 1
            .Row = 1
            .TopRow = 1
            .Redraw = True
            SettaLargFlex Flex(0), 1, 200
        End With

'        Set oQry = oProc.EnumParameters
'        If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
'        Set oQry = Nothing
        
        
        'vVarOut = Array("Return", "", 1)
        ReDim vVarOut(5, 0) As Variant
        vVarOut(0, 0) = "@RETURN_VALUE"
        vVarOut(1, 0) = ""
        vVarOut(2, 0) = ""
        vVarOut(3, 0) = 1
        vVarOut(4, 0) = ""
        vVarOut(5, 0) = False
        
        
        Set oQry = oProc.EnumParameters
        If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
        Set oQry = Nothing
                
        If IsArray(vVarOut) Then
            If IsArray(vVar) Then
                vVar = AddVar2Var(vVarOut, vVar)
            Else
                vVar = vVarOut
            End If
        End If
        For i = 0 To UBound(vVar, 2)
            If Len(vVar(5, i)) = 0 Then vVar(5, i) = True
        Next
        
        LoadFlexParam vVar
    
    End If
    Screen.MousePointer = vbDefault

    Set oProc = Nothing

End Sub
Private Sub LoadViewProperties()
    
    Dim oView As SQLDMO.View2
    Dim i As Integer, sBuf As String
    Dim bBool As Boolean
    Dim oList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column2
    Dim ColElem() As ColProperty
    
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oView = GetViewByName(objServer, m_sDbName, sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnView)
            
        ReDim UserElem(3) As PageElement
        
        UserElem(0).sName = MyLoadResString(k_Param_Encrypted)
        UserElem(0).vVal = oView.Encrypted
        UserElem(0).iType = am_Bool
        
        UserElem(1).sName = MyLoadResString(k_Param_QuotedIdentifier)
        UserElem(1).vVal = oView.QuotedIdentifierStatus
        UserElem(1).iType = am_Bool
        
        UserElem(2).sName = MyLoadResString(k_Func_Param_IsSchemaBound)
        UserElem(2).vVal = oView.IsSchemaBound
        UserElem(2).iType = am_String
        
        UserElem(3).sName = MyLoadResString(k_Param_ANSI_nulls)
        UserElem(3).vVal = oView.AnsiNullsStatus
        UserElem(3).iType = am_String
    
      
        With Flex(0)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 2
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_Property)
            .TextMatrix(0, 1) = MyLoadResString(k_Value)
            
            .ColWidth(0) = 2 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0)
            
    
            sBuf = ""
            For i = 0 To UBound(UserElem)
                Debug.Print UserElem(i).sName; UserElem(i).vVal
                Select Case VarType(UserElem(i).vVal)
                    Case vbString
                        sBuf = UserElem(i).vVal
                    Case vbInteger, vbLong
                        sBuf = Format$(UserElem(i).vVal, "#,0")
                    Case vbSingle, vbDouble, vbDecimal
                        sBuf = Format$(UserElem(i).vVal, "#,0.0000")
                    Case vbDate
                        sBuf = Format$(UserElem(i).vVal, "dd/mm/yyyy")
                    Case vbBoolean, vbByte
                        bBool = True
                        sBuf = ""
                End Select
                Flex(0).AddItem UserElem(i).sName & vbTab & sBuf
                If bBool Then
                    .Col = .Cols - 1
                    .Row = .Rows - 1
                    Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(UserElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                bBool = False
            Next
            
            .FixedCols = 1
            .FixedRows = 1
            .Row = 1
            .TopRow = 1
            .Redraw = True
            SettaLargFlex Flex(0), 1, 200
        End With

        Dim sColHeader(8) As String
        sColHeader(0) = MyLoadResString(k_In_Primary_Key) '"Key"
        sColHeader(1) = MyLoadResString(k_Param_ColId)
        sColHeader(2) = MyLoadResString(k_Field_Name) '"Name"
        sColHeader(3) = MyLoadResString(k_Data_Type) '"DataType"
        sColHeader(4) = MyLoadResString(k_Size) '"Size"
        sColHeader(5) = MyLoadResString(k_Allow_Null) '"Nulls"
        sColHeader(6) = MyLoadResString(k_Default_Value) '"Default"
        sColHeader(7) = MyLoadResString(k_NewDbSort)  'Collation
        sColHeader(8) = MyLoadResString(k_Extended_Property)
                
        Set oList = oView.ListColumns
            
        If Not oList Is Nothing Then
            
            i = -1
            For Each oCol In oList
                i = i + 1
                ReDim Preserve ColElem(i) As ColProperty
                With oCol
                    'ColElem(i).Key = .InPrimaryKey
                    ColElem(i).Id = .Id
                    ColElem(i).Name = .Name
                    ColElem(i).DataType = .DataType
                    ColElem(i).Size = .Length
                    ColElem(i).Nulls = .AllowNulls
                    ColElem(i).Default = .DRIDefault.Text
                    ColElem(i).Collation = .Collation
                End With
            Next
        End If
    
        LoadColumns sColHeader, ColElem
        Flex(1).ColWidth(0) = 0
        
    End If
    Screen.MousePointer = vbDefault

    Set oView = Nothing

End Sub

Private Sub LoadColumns(ByRef sColHeader() As String, ByRef ColElem() As ColProperty)

    Dim i As Integer
    Dim sBuf As String
    
    With Flex(1)
        .Redraw = False
        .Clear
        .FixedCols = 0
        .Cols = 9 'IIf(iActionMode = act_View, 9, 8)
        .Rows = 1
        For i = 0 To UBound(sColHeader)
            .TextMatrix(0, i) = sColHeader(i)
        Next
        .ColWidth(0) = 600
        .ColWidth(1) = 400
        .ColWidth(2) = 200
        .ColWidth(3) = 800
        .ColWidth(4) = 500
        .ColWidth(5) = 800
        
        For i = 0 To UBound(ColElem)
            sBuf = vbTab & ColElem(i).Id & vbTab & ColElem(i).Name & vbTab & ColElem(i).DataType & vbTab & ColElem(i).Size & vbTab & vbTab & ColElem(i).Default & vbTab & ColElem(i).Collation
            .AddItem sBuf
            .Row = .Rows - 1
            
            .Col = 0
            Set Flex(1).CellPicture = LoadResPicture(IIf(ColElem(i).Key, k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            
            .Col = 5
            Set Flex(1).CellPicture = LoadResPicture(IIf(ColElem(i).Nulls, k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            
            'If iActionMode = act_View Then
                .Col = .Cols - 1
                Set Flex(1).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
            'End If
        Next
        .FixedRows = 1
        .Redraw = True
        SettaLargFlex Flex(1), 2, 200
    End With

End Sub
Private Sub LoadFlexParam(ByVal vVar As Variant)

    Dim i As Integer
    Dim sBuf As String
    Dim sDirection As String
    Dim sName As String
    
    With Flex(1)
        .Clear
        .FixedCols = 0
        .Cols = 5
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 0) = MyLoadResString(k_Param_ColId)
        .TextMatrix(0, 1) = MyLoadResString(k_Param_Name)
        .TextMatrix(0, 2) = MyLoadResString(k_Type)
        .TextMatrix(0, 3) = MyLoadResString(k_Param_Direction)
        .TextMatrix(0, 4) = MyLoadResString(k_Extended_Property)
        .ColWidth(0) = 500
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400

        
        .ColAlignment(1) = flexAlignLeftCenter
        
        If IsArray(vVar) Then
            For i = 0 To UBound(vVar, 2)
                If Len(vVar(4, i)) <> 0 Then
                    sDirection = IIf(vVar(4, i) = 1, "input/output", "input")
                Else
                    sDirection = "@RETURN_VALUE"
                End If
                'RetParam
                'If Not CBool(vVar(5, i)) Then
                    sName = vVar(0, i)
                'Else
                '    If Len(vVar(0, i)) <> 0 Then sName = vVar(0, i)
                'End If
                If Len(sName) = 0 Then sName = "@RetParam " & i + 1

                '           ColId           name                        data type                                                   input/output
                sBuf = vVar(3, i) & vbTab & sName & vbTab & vVar(1, i) & " (" & vVar(2, i) & ")" & vbTab & sDirection 'IIf(vVar(4, i) = 1, "input/output", "input")
                .AddItem sBuf
                .Row = .Rows - 1
                .Col = .Cols - 1
                
                If CBool(vVar(5, i)) Then
                    Set Flex(1).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                'Extende Properties Not available
                If Not CBool(vVar(5, i)) Or Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
            Next
        End If
        
        If .Rows > 1 Then
            .FixedRows = 1
            .Row = 1
            .TopRow = 1
        Else
            .Row = .Rows - 1
            For i = 0 To .Cols - 1
                .Col = i
                Flex(1).CellBackColor = vbInactiveBorder
            Next
        End If
        .Redraw = True
        SettaLargFlex Flex(1), 1, 200
    End With

End Sub

Private Function AddVar2Var(ByVal vVar1 As Variant, ByVal vVar2 As Variant) As Variant

    Dim iCols As Integer
    Dim iCols2 As Integer
    Dim iRows As Integer, iRows2 As Integer
    Dim i As Integer, iLoop As Integer
    Dim y As Integer
    Dim iPtr As Integer
    Dim vRet As Variant
    
    iCols = UBound(vVar1)
    iCols2 = UBound(vVar2)
    iRows = UBound(vVar1, 2)
    iRows2 = UBound(vVar2, 2)
    
    i = IIf(iCols > iCols2, iCols, iCols2)
    
    ReDim vRet(i, iRows + iRows2 + 1) As Variant
    
    For iLoop = 0 To iRows
        For i = 0 To iCols
            vRet(i, iPtr) = vVar1(i, iLoop)
        Next
        If iCols2 > iCols Then
            For y = iCols + 1 To iCols2
                vRet(y, iPtr) = ""
            Next
        End If
        iPtr = iPtr + 1
    Next

    For iLoop = 0 To iRows2
        For i = 0 To iCols2
            vRet(i, iPtr) = vVar2(i, iLoop)
        Next
        If iCols > iCols2 Then
            For y = iCols2 + 1 To iCols
                vRet(y, iPtr) = ""
            Next
        End If
        
        iPtr = iPtr + 1
    Next
    AddVar2Var = vRet

End Function

Private Sub LoadView_Idx_Trig()

    Dim oView As SQLDMO.View2
    Dim oIdx As SQLDMO.Index2
    Dim oTrig As SQLDMO.Trigger2
    Dim i As Integer, sBuf As String
    Dim y As Integer
    Dim oList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oView = GetViewByName(objServer, m_sDbName, sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnView)
            

        With Flex(0)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 3
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_IndexName)
            .TextMatrix(0, 1) = MyLoadResString(k_Indexed_Columns)
            .TextMatrix(0, 2) = MyLoadResString(k_Extended_Property)
            
            .ColWidth(0) = 1 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2)
            
    
            sBuf = ""
            For i = 1 To oView.Indexes.Count
                Set oIdx = oView.Indexes(i)
                sBuf = oIdx.Name & vbTab
                Set oList = oIdx.ListIndexedColumns
                
                For y = 1 To oList.Count
                    Set oCol = oList.Item(y)
                    If y > 1 Then sBuf = sBuf & ","
                    sBuf = sBuf & oCol.Name
                Next
            
                .AddItem sBuf
                
                .Row = .Rows - 1
                .Col = .Cols - 1
                Set Flex(0).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
            Next
            
            .FixedCols = 1
            If .Rows > 1 Then
                .FixedRows = 1
                .Row = 1
                .TopRow = 1
            Else
                For i = 1 To .Cols - 1
                    .Col = i
                    .CellBackColor = vbInactiveBorder
                Next
                .Row = 0
            End If
            
            .Redraw = True
            SettaLargFlex Flex(0), 0, 200
        End With
        Set oList = Nothing
        Set oCol = Nothing
        Set oIdx = Nothing
        
        With Flex(1)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 4
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_TriggerName)
            .TextMatrix(0, 1) = MyLoadResString(k_TriggerInsteadOf)
            .TextMatrix(0, 2) = MyLoadResString(k_Type)
            .TextMatrix(0, 3) = MyLoadResString(k_Extended_Property)
            
            .ColWidth(1) = 1 * (.Width / 4)
                
            sBuf = ""
            For i = 1 To oView.Triggers.Count
                Set oTrig = oView.Triggers(i)
                sBuf = oTrig.Name
                .AddItem sBuf & vbTab & vbTab & TriggerType(oTrig.Type)
                
                .Row = .Rows - 1
                .Col = 1
                Set Flex(1).CellPicture = LoadResPicture(IIf(oTrig.InsteadOfTrigger, k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                                
                .Col = .Cols - 1
                Set Flex(1).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
            Next
            
            .FixedCols = 1
            If .Rows > 1 Then
                .FixedRows = 1
                .Row = 1
                .TopRow = 1
            Else
                For i = 1 To .Cols - 1
                    .Col = i
                    .CellBackColor = vbInactiveBorder
                Next
                .Row = 0
            End If
            
            .Redraw = True
            SettaLargFlex Flex(1), 0, 200
        End With
        Set oTrig = Nothing
    End If
    Screen.MousePointer = vbDefault

    Set oView = Nothing

End Sub
Private Sub LoadTable_Idx_Rel()

    Dim oTb As SQLDMO.Table2
    Dim oIdx As SQLDMO.Index2
    Dim oKey As SQLDMO.Key
    
    Dim i As Integer, sBuf As String
    Dim y As Integer
    Dim oList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oTb = GetTbByName(objServer, m_sDbName, sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnTable)
            

        With Flex(0)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 4
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_IndexName)
            .TextMatrix(0, 1) = MyLoadResString(k_Indexed_Columns)
            .TextMatrix(0, 2) = MyLoadResString(k_Primary_Key)
            .TextMatrix(0, 3) = MyLoadResString(k_Extended_Property)
            
            .ColWidth(0) = 1 * (.Width / 3)
            .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2)
            
    
            sBuf = ""
            For i = 1 To oTb.Indexes.Count
                Set oIdx = oTb.Indexes(i)
                Debug.Print oIdx.Type = SQLDMOIndex_DRIPrimaryKey
                If ((oIdx.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex) And ((oIdx.Type And SQLDMOIndex_Hypothetical) <> SQLDMOIndex_Hypothetical) Or ((oIdx.Type And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey) Then
                    sBuf = oIdx.Name & vbTab
                    Set oList = oIdx.ListIndexedColumns
                    
                    For y = 1 To oList.Count
                        Set oCol = oList.Item(y)
                        If y > 1 Then sBuf = sBuf & ","
                        sBuf = sBuf & oCol.Name
                    Next
                                    
                    .AddItem sBuf
                    
                    .Row = .Rows - 1
                    If (oIdx.Type And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey Then
                        .RowData(.Row) = 1
                        .Col = .Cols - 2
                        Set Flex(0).CellPicture = LoadResPicture(k_ResChkON, vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                    End If
                    
                    .Col = .Cols - 1
                    Set Flex(0).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
                End If
            Next
            
            .FixedCols = 1
            If .Rows > 1 Then
                .FixedRows = 1
                .Row = 1
                .TopRow = 1
            Else
                For i = 1 To .Cols - 1
                    .Col = i
                    .CellBackColor = vbInactiveBorder
                Next
                .Row = 0
            End If
            
            .Redraw = True
            SettaLargFlex Flex(0), 0, 200
        End With
        Set oList = Nothing
        Set oCol = Nothing
        Set oIdx = Nothing
        
        With Flex(1)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 3
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_Relations)
            .TextMatrix(0, 1) = MyLoadResString(k_HtmForTB)
            .TextMatrix(0, 2) = MyLoadResString(k_Extended_Property)
            
            .ColWidth(1) = 1 * (.Width / 3)
                
            sBuf = ""
            For i = 1 To oTb.Keys.Count
                Set oKey = oTb.Keys(i)
                If (oKey.Type And SQLDMOKey_Foreign) = SQLDMOKey_Foreign Then
                    sBuf = oKey.Name & vbTab & oKey.ReferencedTable
                    .AddItem sBuf
                                
                    .Row = .Rows - 1
                    .Col = .Cols - 1
                    Set Flex(1).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
                End If
            Next
                        
            .FixedCols = 1
            If .Rows > 1 Then
                .FixedRows = 1
                .Row = 1
                .TopRow = 1
            Else
                For i = 1 To .Cols - 1
                    .Col = i
                    .CellBackColor = vbInactiveBorder
                Next
                .Row = 0
            End If
            
            .Redraw = True
            SettaLargFlex Flex(1), 0, 200
        End With
        Set oKey = Nothing
    End If
    Screen.MousePointer = vbDefault

    Set oTb = Nothing

End Sub
Private Sub LoadTable_Triggers()

    Dim oTb As SQLDMO.Table2
    Dim oTrig As SQLDMO.Trigger2
    
    Dim i As Integer, sBuf As String
        
    Screen.MousePointer = vbHourglass
    
    If Len(m_sDbName) Then
        Set oTb = GetTbByName(objServer, m_sDbName, sStuff)
        
        m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuff, am_OwnTable)
            

        Flex(0).Visible = False
        
        With Flex(1)
            .Redraw = False
            
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            .Cols = 4
            .Rows = 1
            .Clear
            
            .TextMatrix(0, 0) = MyLoadResString(k_TriggerName)
            .TextMatrix(0, 1) = MyLoadResString(k_TriggerInsteadOf)
            .TextMatrix(0, 2) = MyLoadResString(k_Type)
            .TextMatrix(0, 3) = MyLoadResString(k_Extended_Property)
            
            .ColWidth(1) = 1 * (.Width / 3)
                
            sBuf = ""
            For i = 1 To oTb.Triggers.Count
                Set oTrig = oTb.Triggers(i)
                sBuf = oTrig.Name & vbTab & vbTab & TriggerType(oTrig.Type)
                .AddItem sBuf
                
                .Row = .Rows - 1
                .Col = 1
                Set Flex(1).CellPicture = LoadResPicture(IIf(oTrig.InsteadOfTrigger, k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                                
                                

                .Col = .Cols - 1
                Set Flex(1).CellPicture = LoadResPicture(k_ResPush, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If Not m_bIsObjectOwner Then .CellBackColor = vbInactiveBorder
            Next
            
            .FixedCols = 1
            If .Rows > 1 Then
                .FixedRows = 1
                .Row = 1
                .TopRow = 1
            Else
                For i = 1 To .Cols - 1
                    .Col = i
                    .CellBackColor = vbInactiveBorder
                Next
                .Row = 0
            End If
            
            .Redraw = True
            SettaLargFlex Flex(1), 0, 200
        End With
        Set oTrig = Nothing
    End If
    Screen.MousePointer = vbDefault

    Set oTb = Nothing

End Sub

Private Sub ManageExtendedProp(ByVal iFlex As Integer)

    Dim iMode As am_SqlPropType
    Dim iOwner As am_SqlPropTypeOwner
    Dim sParam As String
    Dim iCol As Integer
    
    With Flex(iFlex)
        Select Case iActionMode
            Case act_View, act_Tbl
                iCol = 2
            Case act_Extended_VIEW_Idx_Trig, act_Extended_TABLE_Idx_Rel, act_Extended_TABLE_Trig
                iCol = 0
            Case Else
                iCol = 1
        End Select
            
        sParam = .TextMatrix(.Row, iCol)
    End With
    
    If iActionMode = act_SP Then
        iMode = am_ProcParam
        iOwner = am_OwnSP
    ElseIf iActionMode = act_Extended_VIEW_Idx_Trig Then
        If iFlex = 1 Then
            iMode = am_TrigPropVIEW
        Else
            iMode = am_ViewIdxProp
        End If
        iOwner = am_OwnView
    ElseIf iActionMode = act_Extended_TABLE_Idx_Rel Then
        If iFlex = 1 Then
            iMode = am_KeyProp
        Else
            If Flex(iFlex).RowData(Flex(iFlex).Row) <> 0 Then
            'If StrComp(Flex(iFlex).TextMatrix(Flex(iFlex).Row, Flex(iFlex).Cols - 2), "PK", vbTextCompare) = 0 Then
                iMode = am_KeyProp
            Else
                iMode = am_IdxProp
            End If
        End If
        iOwner = am_OwnTable
    ElseIf iActionMode = act_Extended_TABLE_Trig Then
        iMode = am_TrigPropTB
        iOwner = am_OwnTable
    ElseIf iActionMode = act_View Then
        iMode = am_ViewColProp
        iOwner = am_OwnView
    ElseIf iActionMode = act_User_Function Then
        iMode = am_FuncParam
        iOwner = am_OwnFunction
    ElseIf iActionMode = act_Tbl Then
        iMode = am_ColProp
        iOwner = am_OwnTable
    End If
    If Not m_bIsObjectOwner Then Exit Sub
    
    LoadExtendedProp iMode, m_sDbName, GetObjectOwner(m_sDbName, sStuff, iOwner), sStuff, sParam
End Sub

