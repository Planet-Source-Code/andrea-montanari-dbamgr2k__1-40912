VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Utente"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      Caption         =   "Frame1"
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   5895
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Text            =   "cbo"
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkNt 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
      Begin VB.CheckBox chkNt 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label lbl 
         Caption         =   "Lingua"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "Impostazioni di accesso"
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   5895
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lbl 
         Caption         =   "Lingua"
         ForeColor       =   &H8000000D&
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Database Predefinito"
         ForeColor       =   &H8000000D&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox tLog 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox tLog 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1320
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Frame Fra 
      Caption         =   "Database Consensiti per questo Login"
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1695
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1931
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ruoli Database per"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1365
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "Ruoli assegnati al Login"
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   3015
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5318
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   3855
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Server Roles"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Database Access"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bLoad As Boolean        'annulla caricamento iniziale
Private bCboLoad As Boolean     'annulla caricamento iniziale

Private Type udtSerRole
    bVal As Boolean
End Type

Private Type udtUserNEW
    sDefDb As String
    sDefLang As String
    s_Role() As udtSerRole
    cRoles As cLogin
End Type

Private xUser As udtUserNEW
Private bckUserX As udtUserNEW

Private m_bIsNtUser As Boolean
Private m_bIsOnNT As Boolean
Private bDirtyPass As Boolean
Private sLog As String
Private sDB() As String
Private servRole() As PageElement
Private colSRole() As String


Private Sub cbo_Change(Index As Integer)
    
    If Index = 2 And Not bCboLoad Then LoginStripLoad

End Sub
Private Sub cbo_Click(Index As Integer)
    
    If Index = 2 And Not bCboLoad Then LoginStripLoad
    
End Sub

Private Sub chkNt_Click(Index As Integer)

    If Index = 0 Then
        cbo(2).Enabled = m_bIsOnNT And chkNt(0).Value = vbChecked
        lbl(5).Enabled = m_bIsOnNT And chkNt(0).Value = vbChecked
        tLog(1).Enabled = Not m_bIsOnNT Or chkNt(0).Value = vbUnchecked    'chkNt.Value = vbUnchecked
        lbl(1).Enabled = Not m_bIsOnNT Or chkNt(0).Value = vbUnchecked    'chkNt.Value = vbUnchecked
        chkNt(1).Enabled = m_bIsOnNT ' chkNt.Value = vbChecked
        If chkNt(0).Value = vbChecked Then
            tLog(1).Text = ""
            bDirtyPass = False
        End If
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    Else
        'If chkNt(0).Value = vbUnchecked Then
            SaveData
        'Else
        '    SaveNtData
        'End If
    End If
End Sub

Private Sub Flex_RowColChange(Index As Integer)

    If Index = 1 Then FlexLoadData 2
    
End Sub

Public Property Let LogName(ByVal sVal As String)

    Dim bIsOnNT As Boolean
    
    bIsOnNT = objServer.IsOS(SQLDMO_WINNT) And Len(sLog) = 0

    sLog = sVal
    If Len(sLog) Then
        Me.Caption = MyLoadResString(k_Login_properties) & " ['" & sLog & "']"
        tLog(0).Text = sLog
        tLog(0).Locked = True
    Else
        Me.Caption = MyLoadResString(k_New_Login_Properties)
    End If
End Property

Private Sub tLog_Change(Index As Integer)
    
    Dim bVal As Boolean
    If Len(sLog) Then
        bVal = True
    Else
        bVal = Len(Trim$(tLog(0).Text))
    End If
    
    If Index Then bDirtyPass = True
    
    cmd(1).Enabled = bVal
    tStrip.Enabled = bVal
    Fra(0).Enabled = bVal
    
End Sub
Private Sub tLog_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then KeyCode = 0
End Sub
Private Sub tLog_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn
            If Index = 0 Then SendKeys kNEXT
            KeyAscii = 0
        Case vbKeyEscape
            KeyAscii = 0
            If Index Then
                If tLog(0).Enabled Then SendKeys kPREV
            End If
        Case vbKeyBACK
        Case 43, 45, 48 To 57, 65 To 90, 97 To 122
            OverStrike tLog(Index)
        Case 92  '\
            If Index = 0 Then
                If InStr(tLog(Index).Text, "\") = 0 Then
                    OverStrike tLog(Index)
                Else
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
        
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
        
    m_bIsOnNT = objServer.IsOS(SQLDMO_WINNT)
    bCboLoad = True
    LoadNtDomains
    bCboLoad = False
    
    tStrip.Tabs(1).Caption = MyLoadResString(k_General)
    tStrip.Tabs(2).Caption = MyLoadResString(k_Login_Server_Roles)
    tStrip.Tabs(3).Caption = MyLoadResString(k_Login_Database_Access)
    Fra(0).Caption = MyLoadResString(k_Default_Connection_Information)
    lbl(2).Caption = MyLoadResString(k_Default_Database)
    lbl(3).Caption = MyLoadResString(k_Language)
    Fra(1).Caption = MyLoadResString(k_Server_Role_granted)
    Fra(2).Caption = MyLoadResString(k_Access_granted_to_Database_for_this_Login)
    
    Fra(3).Caption = MyLoadResString(k_Log_Authentication)
    chkNt(0).Caption = MyLoadResString(k_Log_NTAuthentication)
    chkNt(1).Caption = MyLoadResString(k_Log_Deny_Grant)
    lbl(5).Caption = MyLoadResString(k_Log_Domain)
    
    lbl(4).Caption = MyLoadResString(k_Database_Roles_for)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    CenterForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub tStrip_BeforeClick(Cancel As Integer)
    Cancel = Len(sLog) = 0
End Sub

Private Sub tStrip_Click()

    Dim i As Integer, y As Integer
    
    i = tStrip.SelectedItem.Index - 1
    Fra(i).Visible = True
    Fra(i).ZOrder 0
    For y = 0 To 3
        If y <> i Then Fra(y).Visible = False
    Next
    If i = 0 Then
        Fra(3).Visible = True
        Fra(i).ZOrder 0
    End If
    Select Case i
        Case 0
            cbo(0).SetFocus
        Case 1
            Flex(0).SetFocus
        Case 2
            Flex(1).SetFocus
    End Select

End Sub
Private Sub Flex_DblClick(Index As Integer)
    If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then FlexEdit Index, vbKeySpace
End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys kNEXT
        Case vbKeyEscape
            KeyAscii = 0
            SendKeys kPREV
        Case vbKeySpace
            FlexEdit Index, KeyAscii
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Private Sub LoadNtDomains()

    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    Dim i As Integer
                
    Set oQry = objServer.ExecuteWithResults("exec xp_ntsec_enumdomains")
    
    If Not oQry Is Nothing Then
        vVar = Qry2Var(oQry)
        If IsArray(vVar) Then
            With cbo(2)
                .Clear
                For i = 0 To UBound(vVar, 2)
                    .AddItem vVar(0, i)
                Next
                'If .ListCount > 0 Then .ListIndex = 0
            End With
        End If
        
    End If
    Set oQry = Nothing
    
End Sub
Private Sub SaveData()
    
    Dim i As Integer, l As Integer
    Dim bPas As Boolean
    Dim sBuf As String, sStream As String
    Dim lErr As Long, sErr As String
    Dim oLog As SQLDMO.Login
    Dim oUser As SQLDMO.User
    Dim bExit As Boolean
    
    Dim oRole As cDBLog, oOld As cDBLog
    Dim oObj As cDBItem, oOldObj As cDBItem
    
            
    If bDirtyPass And Not m_bIsNtUser Then
        fPwd.Action = act_Login
        fPwd.Stuff() = Trim$(tLog(1).Text)
        fPwd.Show vbModal, Me
        bPas = fPwd.Confirm()
        Unload fPwd
        If Not bPas Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    sStream = cbo(1).List(cbo(1).ListIndex)
    If sStream = "<Default>" Then
        sStream = objServer.Language
    Else
        Dim oLang As SQLDMO.Language
        Set oLang = objServer.Languages.ItemByID(cbo(1).ItemData(cbo(1).ListIndex))
        sStream = oLang.Name
        Set oLang = Nothing
    End If
        
    sBuf = Trim$(tLog(0).Text)
    If Len(sLog) = 0 Then       'NewUser
        If chkNt(0).Value = vbUnchecked Then
            Set oLog = New SQLDMO.Login
            
            oLog.Name = sBuf
            oLog.Type = SQLDMOLogin_Standard
            oLog.Database = cbo(0).List(cbo(0).ListIndex)
                
            oLog.Language = sStream
            
            On Local Error GoTo ErrUpdate
            objServer.Logins.Add oLog
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            objServer.Logins(sBuf).SetPassword "", Trim$(tLog(1).Text)
            sLog = sBuf
        Else
            If chkNt(1).Value = vbUnchecked Then
                sBuf = "exec sp_grantlogin '1%'"
                lErr = ErrDoExecCmd(ReplaceMsg(sBuf, Array("1%"), Array(Trim$(tLog(0).Text))), sErr)
            Else
                sBuf = "exec sp_denylogin '1%'"
                lErr = ErrDoExecCmd(ReplaceMsg(sBuf, Array("1%"), Array(Trim$(tLog(0).Text))), sErr)
                bExit = True
            End If
            
            If lErr = 0 And Not bExit Then
                sBuf = Trim$(tLog(0).Text)
                objServer.Logins.Refresh True
            
                Set oLog = objServer.Logins(sBuf)
                
                oLog.Database = cbo(0).List(cbo(0).ListIndex)
                oLog.Language = sStream
                sLog = sBuf
            Else
                If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
            End If
        
        End If
    Else
        Set oLog = objServer.Logins(sBuf)
        With oLog
            On Local Error GoTo ErrUpdate
            If bDirtyPass Then .SetPassword "", Trim$(tLog(1).Text)
            
            'If sLog <> "sa" Then
                'oLog.Database = cbo(0).List(cbo(0).ListIndex)
                oLog.Language = sStream
                                            
                sStream = cbo(0).List(cbo(0).ListIndex)
                If StrComp(bckUserX.sDefDb, sStream, vbTextCompare) <> 0 Then .Database = sStream
                
                                
                For i = 0 To UBound(servRole)
                    If bckUserX.s_Role(i + 1).bVal <> xUser.s_Role(i + 1).bVal Then
                        If xUser.s_Role(i + 1).bVal Then
                            objServer.ServerRoles(colSRole(i)).AddMember sLog
                        Else
                            objServer.ServerRoles(colSRole(i)).DropMember sLog
                        End If
                    End If
                Next
                
                sStream = ""
                For i = 0 To UBound(sDB)
                    Set oRole = xUser.cRoles.DbRole(sDB(i))
                    Set oOld = bckUserX.cRoles.DbRole(sDB(i))
                    If oRole.CanDB() <> oOld.CanDB() Then
                        If oRole.CanDB() Then
                            Set oUser = New SQLDMO.User
                            oUser.Login = sLog
                            objServer.Databases(sDB(i)).Users.Add oUser
                        Else
                            sStream = objServer.Logins(sLog).GetUserName(sDB(i))
                            If oUser Is Nothing Then Set oUser = objServer.Databases(sDB(i)).Users(sStream)
                        
                            objServer.Databases(sDB(i)).Users.Remove oUser.Name 'sLog
                        End If
                    End If
                    If oRole.CanDB() Then
                        sStream = objServer.Logins(sLog).GetUserName(sDB(i))
                        If oUser Is Nothing Then Set oUser = objServer.Databases(sDB(i)).Users(sStream)
                    
                        For l = 1 To oRole.RoleCount
                            Set oObj = oRole.GetRole(l)
                            Set oOldObj = oOld.GetRole(l)
                            
                            If oObj.IsMember() <> oOldObj.IsMember() Then
                                If oObj.IsMember() Then
                                    objServer.Databases(sDB(i)).DatabaseRoles(oObj.Role()).AddMember oUser.Name 'sLog
                                Else
                                    objServer.Databases(sDB(i)).DatabaseRoles(oObj.Role()).DropMember oUser.Name 'sLog
                                End If
                                
                            End If
                        Next
                        
                    End If
                    Set oUser = Nothing
                    
                Next
                On Local Error GoTo 0
            'End If
        End With
    End If
            
    bDirtyPass = False
    
ExitHere:
    Set oObj = Nothing
    Set oOldObj = Nothing
    Set oRole = Nothing
    Set oOld = Nothing
    If Not oUser Is Nothing Then Set oUser = Nothing
    Set oLog = Nothing
    
    bLoad = False
    If Len(sLog) Then
        LoadLoginProp
    Else
        cbo(0).ListIndex = GetItem("Master", cbo(0))
        cbo(1).ListIndex = GetItem("<Default>", cbo(1))
    End If

    cbo(0).ListIndex = GetItem(xUser.sDefDb, cbo(0))
    cbo(1).ListIndex = GetItem(xUser.sDefLang, cbo(1))

    FlexLoadData 0
    FlexLoadData 1
    bLoad = True
    FlexLoadData 2
    
    cmd(1).Enabled = Len(sLog)
    tStrip.Enabled = Len(sLog)
    Fra(0).Enabled = Len(sLog)
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrUpdate:
    lErr = Err.Number
    sErr = Err.Description
    Screen.MousePointer = vbDefault
    Beep
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Screen.MousePointer = vbHourglass
    Resume ExitHere
    
End Sub
Private Sub LoadDB()

    Dim i As Integer
    Dim oL As SQLDMO.Language
    ReDim sDB(10) As String
    Dim iMax As Integer
    
    objServer.Databases.Refresh True
    
    For i = 1 To objServer.Databases.Count
        If Len(objServer.Databases(i).Name) <> 0 Then
            If iMax > UBound(sDB) Then ReDim Preserve sDB(iMax + 10) As String
            sDB(iMax) = objServer.Databases(i).Name
            iMax = iMax + 1
        End If
    Next
     'iMax- 1(ciclo aggiuntivo)
    ReDim Preserve sDB(iMax - 1) As String
        
    
    cbo(1).AddItem "<Default>"
    
    For Each oL In objServer.Languages
        cbo(1).AddItem oL.Properties("Alias").Value
        cbo(1).ItemData(cbo(1).NewIndex) = oL.Id
    Next
    
    ' i-2 = massimo dei db caricati - 1(zero based) - 1(ciclo aggiuntivo)
    'ReDim Preserve sDB(i - 2) As String
    
    For i = 0 To UBound(sDB)
        cbo(0).AddItem sDB(i)
    Next
        
End Sub
Private Sub PrepareLoad_SRole()

    Dim i As Integer
    
    ReDim servRole(10) As PageElement
    ReDim colSRole(10) As String
    
    For i = 1 To objServer.ServerRoles.Count
        ReDim Preserve xUser.s_Role(1 To i) As udtSerRole
        ReDim Preserve bckUserX.s_Role(1 To i) As udtSerRole
        
        If i > UBound(servRole) Then
            ReDim Preserve servRole(i + 10) As PageElement
            ReDim Preserve colSRole(i + 10) As String
        End If
        servRole(i - 1).iType = am_Bool
        servRole(i - 1).sName = objServer.ServerRoles(i).FullName
        servRole(i - 1).vVal = False
        colSRole(i - 1) = objServer.ServerRoles(i).Name
    Next
    ReDim Preserve servRole(i - 2) As PageElement
    ReDim Preserve colSRole(i - 2) As String
            
End Sub

Private Sub LoadLoginProp()

    Dim i As Integer, l As Integer
    Dim sLogName As String

    If Len(sLog) = 0 Then
        m_bIsNtUser = False

        xUser.sDefDb = "Master"
        xUser.sDefLang = "<Default>"
        bckUserX.sDefLang = xUser.sDefLang
        bckUserX.sDefDb = xUser.sDefDb
        For i = 1 To UBound(servRole)
            xUser.s_Role(i).bVal = False
            bckUserX.s_Role(i).bVal = False
        Next
        Set xUser.cRoles = New cLogin
        Set bckUserX.cRoles = New cLogin
        For i = 0 To UBound(sDB)
            xUser.cRoles.AddDB sDB(i), False
            bckUserX.cRoles.AddDB sDB(i), False
        Next
    Else
        Dim oLog As SQLDMO.Login
        Set oLog = objServer.Logins(sLog)
        
        With oLog
            m_bIsNtUser = oLog.Type <> SQLDMOLogin_Standard
            xUser.sDefDb = .Database
            xUser.sDefLang = .LanguageAlias
            bckUserX.sDefLang = xUser.sDefLang
            bckUserX.sDefDb = xUser.sDefDb
                    
            For i = 0 To UBound(servRole)
                xUser.s_Role(i + 1).bVal = .IsMember(colSRole(i))
                bckUserX.s_Role(i + 1).bVal = xUser.s_Role(i + 1).bVal
            Next
            
            Set xUser.cRoles = New cLogin
            Set bckUserX.cRoles = New cLogin
            For i = 0 To UBound(sDB)
                sLogName = objServer.Logins(sLog).GetUserName(sDB(i))
                If xUser.cRoles.AddDB(sDB(i), Len(sLogName)) Then
                    Call bckUserX.cRoles.AddDB(sDB(i), Len(sLogName))
                    If Len(sLogName) Then
                        Dim oRole As cDBLog
                        Dim oObj As cDBItem
                        Set oRole = xUser.cRoles.DbRole(sDB(i))
                        For l = 1 To oRole.RoleCount()
                            Set oObj = oRole.GetRole(l)
                            oObj.IsMember() = objServer.Databases(sDB(i)).Users(sLogName).IsMember(oObj.Role())
                        Next
                        Set oRole = bckUserX.cRoles.DbRole(sDB(i))
                        For l = 1 To oRole.RoleCount()
                            Set oObj = oRole.GetRole(l)
                            oObj.IsMember() = objServer.Databases(sDB(i)).Users(sLogName).IsMember(oObj.Role())
                        Next
                        Set oObj = Nothing
                        Set oRole = Nothing
                    End If
                End If
            Next
        End With
    End If
    
    chkNt(0).Enabled = m_bIsOnNT And Len(sLog) = 0
    cbo(2).Enabled = m_bIsOnNT And (m_bIsNtUser And (Len(sLog) = 0))
    lbl(5).Enabled = m_bIsOnNT And (m_bIsNtUser And (Len(sLog) = 0))
    chkNt(1).Enabled = m_bIsOnNT And m_bIsNtUser
    If m_bIsOnNT And m_bIsNtUser Then
        chkNt(1).Value = Abs(oLog.DenyNTLogin)
        lbl(1).Enabled = False
        tLog(1).Enabled = False
    End If
        
End Sub

Private Sub FlexLoadData(ByVal iIndex As Integer)

    Dim i As Integer
    Dim l As Integer
    Dim sBuf As String
    Dim oRole As cDBLog
    Dim oObj As cDBItem
    
    Screen.MousePointer = vbHourglass
    With Flex(iIndex)
        .Redraw = False
        .FixedRows = 0
        .FixedCols = 0
        
        .Clear
        .Rows = 1
        Select Case iIndex
            Case 0  'ServerRoles
                .TextMatrix(0, 0) = MyLoadResString(k_Property)
                .TextMatrix(0, 1) = MyLoadResString(k_Value)
                
                .ColWidth(0) = 2 * (.Width / 3)
                .ColWidth(1) = .Width - .ColWidth(0) - 100
                
                For i = 0 To UBound(servRole)
                    .AddItem colSRole(i)
                    .Row = .Rows - 1
                    .Col = 1
                    Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(xUser.s_Role(i + 1).bVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    
                    If sLog = "sa" Then .CellBackColor = vbInactiveBorder
                Next
                .FixedCols = 1
            Case 1  'DBAccessRight
                .Cols = 3
                .TextMatrix(0, 0) = MyLoadResString(k_RES_Database) '"Database"
                .TextMatrix(0, 1) = MyLoadResString(k_User) '"User"
                .TextMatrix(0, 2) = MyLoadResString(k_Permit)
                
                .ColWidth(0) = 2 * (.Width / 5)
                .ColWidth(1) = 2 * (.Width / 5)
                '.ColWidth(2) = .Width - 2 * .ColWidth(0)
                
                For i = 0 To UBound(sDB)
                    sBuf = sDB(i)
                    Set oRole = xUser.cRoles.DbRole(i + 1)
                    If Not oRole Is Nothing Then
                        If oRole.CanDB() Then sBuf = sBuf & vbTab & objServer.Logins(sLog).GetUserName(sDB(i))
                        .AddItem sBuf
                        .Row = .Rows - 1
                        .Col = 2
                        Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(oRole.CanDB()), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                        If sLog = "sa" Then .CellBackColor = vbInactiveBorder
                    End If
                Next
                .FixedCols = 2
            Case 2  'DBRoles
                If bLoad Then
                    lbl(4).Caption = MyLoadResString(k_Database_Roles_for) & " ['" & Flex(1).TextMatrix(Flex(1).Row, 0) & "']"
                    i = Flex(1).Row - 1
                    sBuf = sDB(i)
                    .TextMatrix(0, 0) = MyLoadResString(k_Permit_in_Database_role)
                    .TextMatrix(0, 1) = MyLoadResString(k_Value)
                    
                    .ColWidth(0) = 2 * (.Width / 3)
                    '.ColWidth(1) = .Width - .ColWidth(0)
                
                    Set oRole = xUser.cRoles.DbRole(Flex(1).TextMatrix(Flex(1).Row, 0))
                    
'** Francesco Bancalà
                    If Not oRole Is Nothing Then
                        For l = 1 To oRole.RoleCount
                            Set oObj = oRole.GetRole(l)
                            .AddItem oObj.Role()
                            
                            'If sLog = "sa" Or oObj.Role() = "public" Then .CellBackColor = vbInactiveBorder
                            .Row = .Rows - 1
                            .Col = 1
                            If StrComp(sLog, "sa", vbTextCompare) = 0 Or StrComp(oObj.Role(), "public", vbTextCompare) = 0 Then .CellBackColor = vbInactiveBorder
                                                    
                            If Not oRole.CanDB() Then .CellBackColor = vbInactiveBorder
                            If Not oRole.CanDB() And oObj.Role() <> "public" Then
                                Flex(2).RowHeight(.Row) = 0
                            Else
                                .Row = .Rows - 1
                                .Col = 1
                                
                                Set Flex(iIndex).CellPicture = LoadResPicture(IIf(Abs(oObj.IsMember()), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                .CellPictureAlignment = flexAlignCenterCenter
                            End If
                        Next
                        
                        .FixedRows = 1
                        .FixedCols = 1
                    End If
                End If
        End Select
        If iIndex < 2 Then .FixedRows = 1
        
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
        If iIndex = 1 Then
            SettaLargFlex Flex(iIndex), 2, 300
        ElseIf iIndex = 2 Then
            SettaLargFlex Flex(iIndex), 1, 300
        End If
    End With
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub FlexEdit(ByVal Index As Integer, KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim iImg As Integer, iElem As Integer
    Dim oRole As cDBLog
    Dim oObj As cDBItem
                
    bEdit = True
    iElem = Flex(Index).Row - 1
        
    If Flex(Index).CellBackColor = vbInactiveBorder Then bEdit = False
    
    If iElem < 0 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case vbKeySpace
                Select Case Index
                    Case 0  'ServerRoles
                        iImg = 1 + (xUser.s_Role(iElem + 1).bVal)
                        xUser.s_Role(iElem + 1).bVal = iImg
                    Case 1  'DB Access
                        Set oRole = xUser.cRoles.DbRole(Flex(1).TextMatrix(Flex(1).Row, 0))
                        iImg = 1 + (oRole.CanDB())
                        oRole.CanDB() = iImg
                    Case 2  'DB Righ
                        Set oRole = xUser.cRoles.DbRole(Flex(1).TextMatrix(Flex(1).Row, 0))
                        Set oObj = oRole.GetRole(Flex(2).TextMatrix(Flex(2).Row, 0))
                        iImg = 1 + (oObj.IsMember()) '(thisUser.db_Role(Flex(1).Row - 1, iElem).bVal)
                        oObj.IsMember() = iImg
                End Select
                                
                Set Flex(Index).CellPicture = LoadResPicture(IIf(iImg, k_ResChkON, k_ResChkOFF), vbResBitmap)
                Flex(Index).CellPictureAlignment = flexAlignCenterCenter
                
                If Index = 1 Then FlexLoadData 2
        End Select
    End If
    
End Sub

Public Sub Load_Data()
    
    cbo(0).Clear
    cbo(1).Clear
    LoadDB
    
    PrepareLoad_SRole

    LoadLoginProp

    cbo(0).ListIndex = GetItem(xUser.sDefDb, cbo(0))
    cbo(1).ListIndex = GetItem(xUser.sDefLang, cbo(1))

    FlexLoadData 0
    FlexLoadData 1
    bLoad = True
    FlexLoadData 2
    
    cmd(1).Enabled = Len(sLog)
    tStrip.Enabled = Len(sLog)
    Fra(0).Enabled = Len(sLog)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Function ErrDoExecCmd(ByVal sCommand As String, ByRef sErr As String) As Long

    On Local Error Resume Next
    objServer.ExecuteImmediate sCommand
    ErrDoExecCmd = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
End Function
Private Sub LoginStripLoad()
    Dim iPos As Integer
    Dim sLogin As String
        
    sLogin = tLog(0).Text
    iPos = InStrRev(sLogin, "\")
    If iPos <> 0 Then sLogin = Mid$(sLogin, iPos + 1)
    
    tLog(0).Text = cbo(2).Text & "\" & sLogin

End Sub
'Private Sub oldSaveData()
'
'    Dim i As Integer, l As Integer
'    Dim bPas As Boolean
'    Dim sBuf As String, sStream As String
'    Dim lErr As Long, sErr As String
'    Dim oLog As SQLDMO.Login
'    Dim oUser As SQLDMO.User
'
'    Dim oRole As cDBLog, oOld As cDBLog
'    Dim oObj As cDBItem, oOldObj As cDBItem
'
'
'    If bDirtyPass Then
'        fPwd.Action = act_Login
'        fPwd.Stuff() = Trim$(tLog(1).Text)
'        fPwd.Show vbModal, Me
'        bPas = fPwd.Confirm()
'        Unload fPwd
'        If Not bPas Then Exit Sub
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    sStream = cbo(1).List(cbo(1).ListIndex)
'    If sStream = "<Default>" Then
'        sStream = objServer.Language
'    Else
'        Dim oLang As SQLDMO.Language
'        Set oLang = objServer.Languages.ItemByID(cbo(1).ItemData(cbo(1).ListIndex))
'        sStream = oLang.Name
'        Set oLang = Nothing
'    End If
'
'    sBuf = Trim$(tLog(0).Text)
'    If Len(sLog) = 0 Then       'NewUser
'        Set oLog = New SQLDMO.Login
'
'        oLog.Name = sBuf
'        oLog.Type = SQLDMOLogin_Standard
'        oLog.Database = cbo(0).List(cbo(0).ListIndex)
'
'        oLog.Language = sStream
'
'        On Local Error GoTo ErrUpdate
'        objServer.Logins.Add oLog
'        lErr = Err.Number
'        sErr = Err.Description
'        On Local Error GoTo 0
'
'        objServer.Logins(sBuf).SetPassword "", Trim$(tLog(1).Text)
'        sLog = sBuf
'    Else
'        Set oLog = objServer.Logins(sBuf)
'        With oLog
'            On Local Error GoTo ErrUpdate
'            If bDirtyPass Then .SetPassword "", Trim$(tLog(1).Text)
'
'            'If sLog <> "sa" Then
'                sStream = cbo(0).List(cbo(0).ListIndex)
'                If bckUserX.sDefDb <> sStream Then .Database = sStream
'
'                For i = 0 To UBound(servRole)
'                    If bckUserX.s_Role(i + 1).bVal <> xUser.s_Role(i + 1).bVal Then
'                        If xUser.s_Role(i + 1).bVal Then
'                            objServer.ServerRoles(colSRole(i)).AddMember sLog
'                        Else
'                            objServer.ServerRoles(colSRole(i)).DropMember sLog
'                        End If
'                    End If
'                Next
'
'                sStream = ""
'                For i = 0 To UBound(sDB)
'                    Set oRole = xUser.cRoles.DbRole(sDB(i))
'                    Set oOld = bckUserX.cRoles.DbRole(sDB(i))
'                    If oRole.CanDB() <> oOld.CanDB() Then
'                        If oRole.CanDB() Then
'                            Set oUser = New SQLDMO.User
'                            oUser.Login = sLog
'                            objServer.Databases(sDB(i)).Users.Add oUser
'                        Else
'                            sStream = objServer.Logins(sLog).GetUserName(sDB(i))
'                            If oUser Is Nothing Then Set oUser = objServer.Databases(sDB(i)).Users(sStream)
'
'                            objServer.Databases(sDB(i)).Users.Remove oUser.Name 'sLog
'                        End If
'                    End If
'                    If oRole.CanDB() Then
'                        sStream = objServer.Logins(sLog).GetUserName(sDB(i))
'                        If oUser Is Nothing Then Set oUser = objServer.Databases(sDB(i)).Users(sStream)
'
'                        For l = 1 To oRole.RoleCount
'                            Set oObj = oRole.GetRole(l)
'                            Set oOldObj = oOld.GetRole(l)
'
'                            If oObj.IsMember() <> oOldObj.IsMember() Then
'                                If oObj.IsMember() Then
'                                    objServer.Databases(sDB(i)).DatabaseRoles(oObj.Role()).AddMember oUser.Name 'sLog
'                                Else
'                                    objServer.Databases(sDB(i)).DatabaseRoles(oObj.Role()).DropMember oUser.Name 'sLog
'                                End If
'
'                            End If
'                        Next
'
'                    End If
'                    Set oUser = Nothing
'
'                Next
'                On Local Error GoTo 0
'            'End If
'        End With
'    End If
'
'    bDirtyPass = False
'
'ExitHere:
'    Set oObj = Nothing
'    Set oOldObj = Nothing
'    Set oRole = Nothing
'    Set oOld = Nothing
'    If Not oUser Is Nothing Then Set oUser = Nothing
'    Set oLog = Nothing
'
'    bLoad = False
'
'    If Len(sLog) Then
'        LoadLoginProp
'    Else
'        cbo(0).ListIndex = GetItem("Master", cbo(0))
'        cbo(1).ListIndex = GetItem("<Default>", cbo(1))
'    End If
'
'    cbo(0).ListIndex = GetItem(xUser.sDefDb, cbo(0))
'    cbo(1).ListIndex = GetItem(xUser.sDefLang, cbo(1))
'
'    FlexLoadData 0
'    FlexLoadData 1
'    bLoad = True
'    FlexLoadData 2
'
'    cmd(1).Enabled = Len(sLog)
'    tStrip.Enabled = Len(sLog)
'    fra(0).Enabled = Len(sLog)
'
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'ErrUpdate:
'    lErr = Err.Number
'    sErr = Err.Description
'    Screen.MousePointer = vbDefault
'    Beep
'    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
'    Screen.MousePointer = vbHourglass
'    Resume ExitHere
'
'End Sub

'Private Sub SaveNtData()
'
'    Dim lErr As Long, sErr As String
'    Dim sSql As String
'    Dim sStream As String
'    Dim sLoginNt As String
'    Dim i As Integer, l As Integer
'    Dim oQry As SQLDMO.QueryResults
'    Dim oRole As cDBLog, oOld As cDBLog
'    Dim oObj As cDBItem, oOldObj As cDBItem
'    Dim bExit As Boolean
'    Dim bNewUser As Boolean
'
'    Screen.MousePointer = vbHourglass
'
'    sLoginNt = Trim$(tLog(0).Text)
'    sStream = cbo(0).List(cbo(0).ListIndex)
'
'    If StrComp(sStream, "master", vbTextCompare) <> 0 And StrComp(sStream, "<Default>", vbTextCompare) <> 0 Then
'        For i = 0 To UBound(sDB)
'            If StrComp(sDB(i), sStream, vbTextCompare) = 0 Then
'                Set oRole = xUser.cRoles.DbRole(sDB(i))
'
'                If Not oRole.CanDB() Then
'                    lErr = Err_Free
'                    sErr = "Utente"
'                End If
'                Exit For
'            End If
'        Next
'    End If
'    If lErr = 0 Then
'        bNewUser = Len(sLog) = 0
'        If chkNt(1).Value = vbUnchecked Then
'            sSql = "exec sp_grantlogin '1%'"
'            lErr = ErrDoExecCmd(ReplaceMsg(sSql, Array("1%"), Array(sLoginNt)), sErr)
'        Else
'            sSql = "exec sp_denylogin '1%'"
'            lErr = ErrDoExecCmd(ReplaceMsg(sSql, Array("1%"), Array(sLoginNt)), sErr)
'            bExit = True
'        End If
'
'        If lErr = 0 Then
'            If Not bExit Then
'                sStream = cbo(1).List(cbo(1).ListIndex)
'                If StrComp(sStream, "<Default>", vbTextCompare) = 0 Then sStream = objServer.Language
'
'                sSql = "exec sp_defaultlanguage '1%','2%'"
'                lErr = ErrDoExecCmd(ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)), sErr)
'
'                If lErr = 0 Then
'                    sStream = cbo(0).List(cbo(0).ListIndex)
'
'                    sSql = "exec sp_defaultdb  '1%','2%'"
'                    lErr = ErrDoExecCmd(ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)), sErr)
'                End If
'                If bNewUser Then bExit = True
'
'            End If
'        End If
'
'        If lErr = 0 Then
'            ' server roles
'            On Local Error Resume Next
'            For i = 0 To UBound(servRole)
'                If bckUserX.s_Role(i + 1).bVal <> xUser.s_Role(i + 1).bVal Then
'                    If xUser.s_Role(i + 1).bVal Then
'                        sSql = ReplaceMsg("exec sp_addsrvrolemember '1%', '2%'", Array("1%", "2%"), Array(sLoginNt, colSRole(i)))
'                    Else
'                        sSql = ReplaceMsg("exec sp_dropsrvrolemember '1%', '2%'", Array("1%", "2%"), Array(sLoginNt, colSRole(i)))
'                    End If
'                    lErr = ErrDoExecCmd(sSql, sErr)
'                End If
'                If lErr <> 0 Then Exit For
'            Next
'            ' databases
'            If lErr = 0 Then
'                For i = 0 To UBound(sDB)
'                    Set oRole = xUser.cRoles.DbRole(sDB(i))
'                    Set oOld = bckUserX.cRoles.DbRole(sDB(i))
'                    If oRole.CanDB() <> oOld.CanDB() Then
'                        If oRole.CanDB() Then
'                            sSql = ReplaceMsg("exec sp_grantdbaccess '1%'", Array("1%"), Array(sLoginNt))
'                        Else
'                            sSql = ReplaceMsg("exec sp_revokedbaccess '1%'", Array("1%"), Array(sLoginNt))
'                        End If
'                        lErr = ErrDoExecCmd(sSql, sErr)
'                    End If
'
'                    If oRole.CanDB() Then
'                        For l = 1 To oRole.RoleCount
'                            Set oObj = oRole.GetRole(l)
'                            Set oOldObj = oOld.GetRole(l)
'
'                            If oObj.IsMember() <> oOldObj.IsMember() Then
'                                If oObj.IsMember() Then
'                                    sSql = ReplaceMsg("exec sp_addrolemember '1%', '2%'", Array("1%", "2%"), Array(oObj.Role(), sLoginNt))
'                                Else
'                                    sSql = ReplaceMsg("exec sp_droprolemember '1%', '2%'", Array("1%", "2%"), Array(oObj.Role(), sLoginNt))
'                                End If
'                                lErr = ErrDoExecCmd(sSql, sErr)
'                            End If
'                            If lErr <> 0 Then Exit For
'                        Next
'                    End If
'                    If lErr <> 0 Then Exit For
'                Next
'            End If
'        End If
'    End If
'
'
''        If Len(sLog) = 0 Then       'NewUser
''            If chkNt(1).Value = vbUnchecked Then
''                sSql = "exec sp_grantlogin '1%'"
''                lErr = ErrDoExecCmd(ReplaceMsg(sSql, Array("1%"), Array(sLoginNt)))
''            Else
''            End If
''
''            If lErr = 0 Then
''                sStream = cbo(1).List(cbo(1).ListIndex)
''                If StrComp(sStream, "<Default>", vbTextCompare) = 0 Then sStream = objServer.Language
''
''                sSql = "exec sp_defaultlanguage '1%','2%'"
''                objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)))
''                lErr = Err.Number
''            End If
''
''            If lErr = 0 Then
''                sStream = cbo(0).List(cbo(0).ListIndex)
''
''                sSql = "exec sp_defaultdb  '1%','2%'"
''                objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)))
''                lErr = Err.Number
''            End If
''
''            sErr = Err.Description
''            On Local Error GoTo 0
''        Else
''            sSql = "exec sp_denylogin '1%'"
''            On Local Error Resume Next
''            objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%"), Array(sLoginNt)))
''            lErr = Err.Number
''            sErr = Err.Description
''            On Local Error GoTo 0
''        End If
''    Else
''        If chkNt(1).Value = vbUnchecked Then
''
''            If lErr = 0 Then
''                sStream = cbo(1).List(cbo(1).ListIndex)
''                If StrComp(sStream, "<Default>", vbTextCompare) = 0 Then sStream = objServer.Language
''
''                sSql = "exec sp_defaultlanguage '1%','2%'"
''                objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)))
''                lErr = Err.Number
''            End If
''
''            If lErr = 0 Then
''                sStream = cbo(0).List(cbo(0).ListIndex)
''
''                sSql = "exec sp_defaultdb  '1%','2%'"
''                objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%", "2%"), Array(sLoginNt, sStream)))
''                lErr = Err.Number
''            End If
''
''            ' server roles
''            On Local Error Resume Next
''            For i = 0 To UBound(servRole)
''                If bckUserX.s_Role(i + 1).bVal <> xUser.s_Role(i + 1).bVal Then
''                    If xUser.s_Role(i + 1).bVal Then
''                        objServer.ExecuteImmediate (ReplaceMsg("exec sp_addsrvrolemember '1%', '2%'", Array("1%", "2%"), Array(sLoginNt, colSRole(i))))
''                    Else
''                        objServer.ExecuteImmediate (ReplaceMsg("exec sp_dropsrvrolemember '1%', '2%'", Array("1%", "2%"), Array(sLoginNt, colSRole(i))))
''                    End If
''                End If
''                If lErr <> 0 Then Exit For
''            Next
''            ' databases
''            If lErr = 0 Then
''                For i = 0 To UBound(sDB)
''                    Set oRole = xUser.cRoles.DbRole(sDB(i))
''                    Set oOld = bckUserX.cRoles.DbRole(sDB(i))
''                    If oRole.CanDB() <> oOld.CanDB() Then
''                        If oRole.CanDB() Then
''                            objServer.ExecuteImmediate (ReplaceMsg("exec sp_grantdbaccess '1%'", Array("1%"), Array(sLoginNt)))
''                        Else
''                            objServer.ExecuteImmediate (ReplaceMsg("exec sp_revokedbaccess '1%'", Array("1%"), Array(sLoginNt)))
''                        End If
''                    End If
''                    If oRole.CanDB() Then
''
''                        For l = 1 To oRole.RoleCount
''                            Set oObj = oRole.GetRole(l)
''                            Set oOldObj = oOld.GetRole(l)
''
''                            If oObj.IsMember() <> oOldObj.IsMember() Then
''                                If oObj.IsMember() Then
''                                    objServer.ExecuteImmediate (ReplaceMsg("exec sp_addrolemember '1%', '2%", Array("1%", "2%"), Array(oObj.Role(), sLoginNt)))
''                                Else
''                                    objServer.ExecuteImmediate (ReplaceMsg("exec sp_droprolemember '1%', '2%", Array("1%", "2%"), Array(oObj.Role(), sLoginNt)))
''                                End If
''                            End If
''                            If lErr <> 0 Then Exit For
''                        Next
''                    End If
''                    If lErr <> 0 Then Exit For
''                Next
''            End If
''
''            sErr = Err.Description
''            On Local Error GoTo 0
''        Else
''            sSql = "exec sp_denylogin '1%'"
''            On Local Error Resume Next
''            objServer.ExecuteImmediate (ReplaceMsg(sSql, Array("1%"), Array(sLoginNt)))
''            lErr = Err.Number
''            sErr = Err.Description
''            On Local Error GoTo 0
''        End If
''
''    End If
'    Set oQry = Nothing
'    Set oRole = Nothing
'    Set oOld = Nothing
'    Set oObj = Nothing
'    Set oOldObj = Nothing
'
'    If lErr = 0 Then
'        sLog = sLoginNt
'        objServer.Logins.Refresh True
'        If Len(sLog) Then
'            LoadLoginProp
'        Else
'            cbo(0).ListIndex = GetItem("Master", cbo(0))
'            cbo(1).ListIndex = GetItem("<Default>", cbo(1))
'        End If
'
'        cbo(0).ListIndex = GetItem(xUser.sDefDb, cbo(0))
'        cbo(1).ListIndex = GetItem(xUser.sDefLang, cbo(1))
'
'        FlexLoadData 0
'        FlexLoadData 1
'        bLoad = True
'        FlexLoadData 2
'
'        cmd(1).Enabled = Len(sLog)
'        tStrip.Enabled = Len(sLog)
'        fra(0).Enabled = Len(sLog)
'
'    End If
'    Screen.MousePointer = vbDefault
'    If lErr <> 0 Then
'        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
'    End If
'
'End Sub

