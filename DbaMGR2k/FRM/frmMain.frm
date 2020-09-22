VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "DbaMGR"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8025
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   7920
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   72
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2175
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":0A8A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2055
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Frame fra 
      Caption         =   "Posizione Dati"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmd 
         Caption         =   ".."
         Height          =   300
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5250
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8493
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11.57"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "18/11/2002"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tView 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   50
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5953
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7200
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B0B
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10A7
            Key             =   "Backup"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1643
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BDF
            Key             =   "ServerGroup"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":217B
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2717
            Key             =   "Logs"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CB3
            Key             =   "Logins"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":324F
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37EB
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D87
            Key             =   "Tables"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4323
            Key             =   "Stored"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48BF
            Key             =   "udt"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E5B
            Key             =   "Roles"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53F7
            Key             =   "ProcInfo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5553
            Key             =   "Activity"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AEF
            Key             =   "detached"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":608B
            Key             =   "LockID"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6627
            Key             =   "Func"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2760
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu mnu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuF 
         Caption         =   "connetti"
         Index           =   0
      End
      Begin VB.Menu mnuF 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuF 
         Caption         =   "Includi Oggetti di Sistema"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuF 
         Caption         =   "&Aggiorna"
         Index           =   3
      End
      Begin VB.Menu mnuF 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuF 
         Caption         =   "&Query"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuF 
         Caption         =   "BCP"
         Index           =   6
      End
      Begin VB.Menu mnuF 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuF 
         Caption         =   "&Esci"
         Index           =   9
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Info"
      Index           =   1
      Begin VB.Menu mnuInfo 
         Caption         =   "Sql Engine &Version"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "lng"
         Index           =   1
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
         Index           =   2
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Depend"
         Index           =   3
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "cli"
         Index           =   5
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "srv"
         Index           =   6
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Credits"
         Index           =   9
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "About"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lWidth As Long, m_lHeight As Long
Private Enum SPID_Action
    enSPIDDelete = 0
End Enum
Private Type dbDetach
    sName As String
    bDetach As Boolean
End Type
Private m_sVntDetached() As String

Private mNode As Node
Private mbMoving As Boolean
Const sglSplitLimit = 500

Private collLock As Collection
Private m_bNeedReattach As Boolean

Public Sub CallBackReattach()
    m_bNeedReattach = True
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim bCancel As Boolean
    Dim bErr As Boolean
    Dim sTarget As String
    Dim lErr As Long, sErr As String
    
    If Index = 0 Then
        Dim fBrowse As fBr4Fold
        Set fBrowse = New fBr4Fold
        
        With fBrowse
            .Init enB4isSelectDir, MyLoadResString(k_Full_Data_PathToolTip), txtServer.Text
                        
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
            If Right$(sTarget, 1) <> "\" Then sTarget = sTarget & "\"
            Call WriteIni("DATA", "DEFDIR", sTarget, App.Path & "\" & App.EXEName & ".INI")
        End If
        If Len(sTarget) = 0 Then sTarget = txtServer.Text
        If Len(sTarget) = 0 Then sTarget = objServer.Registry.MasterDBPath()
        txtServer.Text = sTarget
        
        
        If lErr <> 0 Then MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
        
'
'        Dim a As BROWSEINFO
'        Dim sBuf As String
'
'        Dim pidlBrowse As Long
'
'        a.lpszTitle = MyLoadResString(k_Select_Full_Data_Path)   '"Seleziona Percorso Dati"
'
'        a.hOwner = Me.hwnd
'        pidlBrowse = SHBrowseForFolder(a)
'        sBuf = GetPathFromPIDL(pidlBrowse)
'        If Len(sBuf) Then
'            'If Len(sbuf) = 3 And Right$(sbuf, 1) = "\" Then sbuf = Left$(sbuf, 2)
'            'If Right$(sBuf, 1) <> "\" Then sBuf = sBuf & "\"
''            On Local Error Resume Next
''            If Len(Dir$(sBuf & "*.mdf")) = 0 Then sBuf = "..."
''            On Local Error GoTo 0
'        Else
''            sBuf = "..."
'            sBuf = App.Path
'        End If
'        If Right$(sBuf, 1) <> "\" Then sBuf = sBuf & "\"
'        'If sBuf <> "..." Then Call WriteIni("DATA", "DEFDIR", sBuf, App.Path & "\" & App.EXEName & ".INI")
'        'Call WriteIni("DATA", "DEFDIR", IIf(sBuf <> "...", sBuf, ""), App.Path & "\" & App.EXEName & ".INI")
'        Call WriteIni("DATA", "DEFDIR", sBuf, App.Path & "\" & App.EXEName & ".INI")
'
'        lblServer(0).Caption = sBuf
    End If
    
End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer
    Dim sBuf As String
    Dim sStuf As String, sDB As String
    Dim iRow As Integer, iSubMenu As Integer
    Dim vParam As Variant
    Dim bMultiSel As Boolean
    Dim l_intLoop As Integer
    Dim intLow As Integer, intHigh As Integer
    Dim intCount As Integer
    Dim lSpid As Long
   
    Dim iGo As optAction
    
    iRow = Flex.Row
    
    If tView.SelectedItem Is Nothing Then Exit Sub
    
    If Button = vbRightButton Then
        bMultiSel = Flex.RowSel <> Flex.Row
            
        'If bMultiSel Then Exit Sub
        iGo = act_Null
            
        On Local Error Resume Next
        sBuf = tView.SelectedItem.Parent.Parent
        'questo NON e' necessario ma lo metto per sicurezza
        If Err.Number <> 0 Or sBuf = "SqlServer" Then sBuf = ""
        On Local Error GoTo 0
                
        Dim M As New cPUpMenu
            
           
        If Len(sBuf) <> 0 Then
            Select Case sBuf
                Case MyLoadResString(kTviewDatabases) '"Databases"
                    iGo = act_DBOption
                    sDB = tView.SelectedItem.Parent
                    sBuf = tView.SelectedItem
                    Select Case sBuf
                        Case MyLoadResString(kTviewDatabasesUsers) '"Database Users"
                            sStuf = Flex.TextMatrix(iRow, 0)
                            iGo = act_User
                                
                            M.Add MyLoadResString(k_New_User), 1, bDisable:=Not GetMembership(db_AccessAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, , bDisable:=Not GetMembership(db_AccessAdmin, sDB) Or iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Manage_Permissions), 5, , bDisable:=Not GetMembership(db_AccessAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 6, bSeparator:=True
                            M.Add ReplaceMsg(MyLoadResString(k_Search4Orphaned), Array("1%"), Array(sDB)), 7, , bDisable:=Not GetMembership(SvrAdm, sDB) Or bMultiSel Or Not IsDbAvailable(sDB)
                            M.Add ReplaceMsg(MyLoadResString(k_ObjectOwnedby), Array("1%"), Array(sStuf)), 8, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not GetMembership(db_SecurAdm, sDB) Or bMultiSel Or Not IsDbAvailable(sDB)
                            M.Add "", 9, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 10, , bDisable:=Not GetMembership(db_AccessAdmin, sDB) Or iRow < 1 Or bMultiSel
                            
                        Case MyLoadResString(kTviewDatabasesTables) '"Tables"
                            sStuf = Flex.TextMatrix(iRow, 0)
                            iGo = act_Tbl
                            
                            M.Add MyLoadResString(k_New_Table), 1, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Table_Design), 3, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Show_all_Rows), 4, bDisable:=Not GetMembership(db_DataReader, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Show_Max), 5, bDisable:=Not GetMembership(db_DataReader, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 6, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 7, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1
                            M.Add "", 8, bSeparator:=True
                            M.Add MyLoadResString(k_Generate_SQL_Script), 9, bDisable:=iRow < 1 Or bMultiSel   'Script
                            M.Add MyLoadResString(k_Manage_Permissions), 10, , bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 11, bSeparator:=True
                            M.Add MyLoadResString(k_Show_Dependencies), 12, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 13, bSeparator:=True
                            M.Add MyLoadResString(k_ManageTriggers), 14, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Indexes_Management), 15, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_PropertyIdxRel), 16, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnTable) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_PropertyTriggers), 17, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnTable) Or iRow < 1 Or bMultiSel
                            M.Add "", 18, bSeparator:=True
                            M.Add MyLoadResString(k_MainRelViewer), 19, bDisable:=Not GetMembership(db_DataReader, sDB) Or iRow < 1
                            M.Add "", 20, bSeparator:=True
                            M.Add MyLoadResString(k_ModifyObjectOwner), 21, bDisable:=(Not GetMembership(db_DDLAdmin, sDB) Or Not GetMembership(db_SecurAdm, sDB)) Or iRow < 1
                            M.Add MyLoadResString(k_sqlMnuGenerateInsert), 22, bDisable:=(Not GetMembership(db_DataReader, sDB) Or Not GetMembership(db_SecurAdm, sDB)) Or iRow < 1 Or bMultiSel
                            M.Add "", 23, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 24, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_Property), 25, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnTable) Or iRow < 1 Or bMultiSel
                            
                            
                            vParam = Array(sDB, sStuf)
                        Case MyLoadResString(kTviewDatabasesViews) '"Views"
                            iGo = act_View
                            
                            M.Add MyLoadResString(k_New_View), 1, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Show_all_Rows), 5, bDisable:=Not GetMembership(db_DataReader, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Show_Max), 6, bDisable:=Not GetMembership(db_DataReader, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 7, bSeparator:=True
                            M.Add MyLoadResString(k_Generate_SQL_Script), 8, bDisable:=iRow < 1 Or bMultiSel   'Script
                            M.Add MyLoadResString(k_Manage_Permissions), 9, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 10, bSeparator:=True
                            M.Add MyLoadResString(k_ManageTriggers), 11, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Indexes_Management), 12, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_PropertyIdxTrig), 13, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnView) Or iRow < 1 Or bMultiSel
                            M.Add "", 14, bSeparator:=True
                            M.Add MyLoadResString(k_Show_Dependencies), 15, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_ModifyObjectOwner), 16, bDisable:=(Not GetMembership(db_DDLAdmin, sDB) Or Not GetMembership(db_SecurAdm, sDB)) Or iRow < 1
                            M.Add "", 17, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 18, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_Property), 19, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnView) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Text_Property), 20, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            
                            vParam = Array(sDB, Flex.TextMatrix(iRow, 0))

                        Case MyLoadResString(kTviewDatabasesStoredProc) '"Stored Procedures"
                            iGo = act_SP
                            M.Add MyLoadResString(k_New_Stored_Procedure), 1, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1   'bDisable:=iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Generate_SQL_Script), 5, bDisable:=iRow < 1 Or bMultiSel   'Script
                            M.Add MyLoadResString(k_Manage_Permissions), 6, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 7, bSeparator:=True
                            M.Add MyLoadResString(k_Show_Dependencies), 8, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_ModifyObjectOwner), 9, bDisable:=(Not GetMembership(db_DDLAdmin, sDB) Or Not GetMembership(db_SecurAdm, sDB)) Or iRow < 1
                            M.Add "", 10, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 11, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_Property), 12, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnSP) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Text_Property), 13, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel

                            vParam = Array(sDB, Flex.TextMatrix(iRow, 0))
                            
                        Case MyLoadResString(kTviewFunction)  'Funzioni
                            iGo = act_User_Function
                            M.Add MyLoadResString(k_Func_NewFunction), 1, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1   'bDisable:=iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Generate_SQL_Script), 5, bDisable:=iRow < 1 Or bMultiSel   'Script
                            M.Add MyLoadResString(k_Manage_Permissions), 6, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1 Or bMultiSel
                            M.Add "", 7, bSeparator:=True
                            M.Add MyLoadResString(k_Show_Dependencies), 8, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_ModifyObjectOwner), 9, bDisable:=(Not GetMembership(db_DDLAdmin, sDB) Or Not GetMembership(db_SecurAdm, sDB)) Or iRow < 1
                            M.Add "", 10, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 11, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Extended_Property), 12, bDisable:=Not IsObjectOwner(sDB, Flex.TextMatrix(iRow, 0), am_OwnFunction) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Text_Property), 13, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel

                            vParam = Array(sDB, Flex.TextMatrix(iRow, 0))
                            
                        Case MyLoadResString(kTviewDatabasesRoles) '"Roles"
                            iGo = act_Roles
                            M.Add MyLoadResString(k_New_DB_Role), 1, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1   'bDisable:=iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Property), 5, bDisable:=Not GetMembership(db_SecurAdm, sDB) Or iRow < 1 Or bMultiSel
                            
                        Case MyLoadResString(kTviewDatabasesUDT) '"User Defined Data Types"
                            iGo = act_UDT
                            M.Add MyLoadResString(k_New_User_Defined_Data_Type), 1, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or Not IsDbAvailable(sDB)
                            M.Add "", 2, bSeparator:=True
                            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1   'bDisable:=iRow < 1
                            M.Add "", 4, bSeparator:=True
                            M.Add MyLoadResString(k_Generate_SQL_Script), 5, bDisable:=iRow < 1 Or bMultiSel   'Script
                            M.Add "", 6, bSeparator:=True
                            M.Add MyLoadResString(k_Show_Dependencies), 7, , bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel
                            M.Add MyLoadResString(k_Property), 8, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1 Or bMultiSel   'bDisable:=iRow < 1
                            M.Add MyLoadResString(k_Extended_Property), 9, bDisable:=Not GetMembership(db_DDLAdmin, sDB) Or iRow < 1
                            
                    End Select
                Case MyLoadResString(k_RES_Management)
                    'verifico che lo spid > system con
                    '"EXEC sp_MSget_current_activity @id=7,@option=1"
                    
                    If StrComp(tView.SelectedItem.Key, "procinfo", vbTextCompare) = 0 Then

                        If Not bMultiSel Then
                            iGo = act_ProcInfo
                            lSpid = Val(Flex.TextMatrix(Flex.Row, 0))

                            M.Add MyLoadResString(k_PROC_KillSpid), 1, bDisable:=Not GetMembership(ProcAdm, "") Or Not TestProcInfo(lSpid, enSPIDDelete)

                            Call GetCursorPos(P)
                            i = M.Popup(Me.hWnd, P.x, P.y)
                            M.Clear
                            Set M = Nothing
                            If i <> 0 Then
                                ActMenu i, iGo, "", Array(lSpid)
                            End If

                        End If
                        iGo = act_Null

                    End If

            End Select
            If iGo <> act_Null Then
                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
                Set M = Nothing
                
                If i <> 0 Then
                    ReDim vParam(intCount) As Variant
                    SetLowHigh intLow, intHigh
        
                    For l_intLoop = intLow To intHigh
                        ReDim Preserve vParam(intCount) As Variant
                        vParam(intCount) = Flex.TextMatrix(l_intLoop, 0)
                        intCount = intCount + 1
                    Next
                    ActMenu i, iGo, sDB, vParam
                End If
            End If
        ElseIf tView.SelectedItem = MyLoadResString(kTviewDatabases) Then '"Databases" Then
            sBuf = Flex.TextMatrix(iRow, 0)
            i = objServer.Databases(sBuf).SystemObject
            iGo = act_DBOption

            M.Add MyLoadResString(k_New_DB), 1, bDisable:=Not GetMembership(DBCreat)
            M.Add "", 2, bSeparator:=True
            M.Add MyLoadResString(k_Detach_DB), 3, bDisable:=Not GetMembership(SysAdm) Or Not GetMembership(db_dbo, sBuf) Or iRow < 1
            M.Add "", 4, bSeparator:=True
            M.Add MyLoadResString(k_Remove), 5, bDisable:=Not GetMembership(SysAdm) Or Not GetMembership(db_dbo, sBuf) Or iRow < 1
            M.Add "", 6, bSeparator:=True
            M.Add "Backup Database", 7, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "Restore Database", 8, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "Truncate Log", 9, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "", 10, bSeparator:=True
            M.Add MyLoadResString(k_Shrink_DB), 11, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "Check Database Integrity >>", 12, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "CheckPoint Database", 13, bDisable:=Not GetMembership(db_BckOper, sBuf) Or iRow < 1 Or bMultiSel
            M.Add "", 14, bSeparator:=True
            M.Add MyLoadResString(k_Generate_SQL_Script), 15, bDisable:=iRow < 1 Or bMultiSel Or Not (IsDbUser(sBuf))
            M.Add "", 16, bSeparator:=True
            M.Add MyLoadResString(k_HtmDocum), 17, bDisable:=iRow < 1 Or bMultiSel Or Not (IsDbUser(sBuf))
            M.Add "", 18, bSeparator:=True
            'M.Add MyLoadResString(k_ModifyDB_Owner), 19, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
            'sp3 requires sysadmin privileges
            M.Add MyLoadResString(k_ModifyDB_Owner), 19, bDisable:=Not GetMembership(SysAdm) Or iRow < 1
            
            M.Add "", 20, bSeparator:=True
            M.Add MyLoadResString(k_Extended_Property), 21, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
            M.Add MyLoadResString(k_Property), 22, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
                        

            Call GetCursorPos(P)
            i = M.Popup(Me.hWnd, P.x, P.y)
            M.Clear
            Set M = Nothing

            If i = 11 Then
                iSubMenu = 30

                M.Add MyLoadResString(k_Shrink_Default5%), 1 ' "Shrink_Default (prova a 5%)"
                M.Add "Shrink_TruncateOnly", 2
                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
            ElseIf i = 12 Then
                iSubMenu = 40

                M.Add "Check DB Repair_None", 1
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Fast", 2
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Rebuild", 3
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Allow_DataLoss", 4

                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
            End If

            If i <> 0 Then
                ReDim vParam(intCount) As Variant
                SetLowHigh intLow, intHigh
                
                For l_intLoop = intLow To intHigh
                    ReDim Preserve vParam(intCount) As Variant
                    vParam(intCount) = Flex.TextMatrix(l_intLoop, 0)
                    intCount = intCount + 1
                Next
                'vParam = sBuf
                'ActMenu i + iSubMenu, iGo, vParam
                ActMenu i + iSubMenu, iGo, "", vParam
            End If
                    
        ElseIf tView.SelectedItem = MyLoadResString(kTviewLogin) Then '"Login" Then
            iGo = act_Login
            
            M.Add MyLoadResString(k_New_Login), 1, bDisable:=Not GetMembership(SecurAdm)
            M.Add "", 2, bSeparator:=True
            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(SecurAdm) Or iRow < 1
            M.Add "", 4, bSeparator:=True
            M.Add MyLoadResString(k_Property), 5, bDisable:=Not GetMembership(SecurAdm) Or iRow < 1 Or bMultiSel
            Call GetCursorPos(P)
            i = M.Popup(Me.hWnd, P.x, P.y)
            M.Clear
            Set M = Nothing
            
            If i <> 0 Then
                ReDim vParam(intCount) As Variant
                SetLowHigh intLow, intHigh
                intCount = 0
                
                For l_intLoop = intLow To intHigh
                    ReDim Preserve vParam(intCount) As Variant
                    vParam(intCount) = Flex.TextMatrix(l_intLoop, 0)
                    intCount = intCount + 1
                Next
                        
                ActMenu i, iGo, "", vParam
            End If
        ElseIf tView.SelectedItem = MyLoadResString(kTviewDevices) Then '"BackUp Device" Then
            iGo = act_Bck
            M.Add MyLoadResString(k_New_BackUp_Device), 1, bDisable:=Not GetMembership(DiskAdm)
            M.Add "", 2, bSeparator:=True
            M.Add MyLoadResString(k_Remove), 3, bDisable:=Not GetMembership(DiskAdm) Or iRow < 1 Or bMultiSel
            M.Add "", 4, bSeparator:=True
            M.Add MyLoadResString(k_Property), 5, bDisable:=Not GetMembership(DiskAdm) Or iRow < 1 Or bMultiSel

            Call GetCursorPos(P)
            i = M.Popup(Me.hWnd, P.x, P.y)
            M.Clear
            Set M = Nothing
            
            If i <> 0 Then
                ReDim vParam(intCount) As Variant
                SetLowHigh intLow, intHigh
                intCount = 0
                
                For l_intLoop = intLow To intHigh
                    ReDim Preserve vParam(intCount) As Variant
                    vParam(intCount) = Flex.TextMatrix(l_intLoop, 0)
                    intCount = intCount + 1
                Next
                        
                ActMenu i, iGo, "", vParam
            End If
        ElseIf tView.SelectedItem = MyLoadResString(k_Detached_Databases) Then
            If iRow > 0 Then
            
                M.Add MyLoadResString(k_Attach_DB), 1, bDisable:=Not GetMembership(DBCreat) Or bMultiSel
            
                vParam = Flex.TextMatrix(iRow, 0)
        
                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
                Set M = Nothing
                
                If i <> 0 Then AttachDB Flex.TextMatrix(iRow, 0), m_sVntDetached(iRow - 1)
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    
    Dim iL As Integer, iT As Integer, iW As Integer, iH As Integer
    Dim sBuf As String
    
    On Local Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    On Local Error GoTo 0
    
    If App.PrevInstance = True Then End
    m_lWidth = Me.Width
    m_lHeight = Me.Height
    
    'default = Show SysObjects
    If Len(Trim$(GetIni("CONNECT", "SYSOBJ", App.Path & "\" & App.EXEName & ".INI"))) = 0 Then
        Call WriteIni("CONNECT", "SYSOBJ", "1", App.Path & "\" & App.EXEName & ".INI")
    End If
    mnuF(2).Checked = Val(GetIni("CONNECT", "SYSOBJ", App.Path & "\" & App.EXEName & ".INI"))
    
    
    g_strLNG = GetIni("CONNECT", "LCID", App.Path & "\" & App.EXEName & ".INI")
    If Len(g_strLNG) = 0 Or Len(Dir$(LngFileName())) = 0 Then g_strLNG = "English"
        
    LoadStrings
  
    'default = Show All Objects
    If Len(Trim$(GetIni("CONNECT", "SHOWALL", App.Path & "\" & App.EXEName & ".INI"))) = 0 Then
        Call WriteIni("CONNECT", "SHOWALL", "1", App.Path & "\" & App.EXEName & ".INI")
    End If
    
    iL = Val(GetIni("DEFAULT", "LEFT", App.Path & "\" & App.EXEName & ".INI"))
    iT = Val(GetIni("DEFAULT", "TOP", App.Path & "\" & App.EXEName & ".INI"))
    iW = Val(GetIni("DEFAULT", "WIDTH", App.Path & "\" & App.EXEName & ".INI"))
    iH = Val(GetIni("DEFAULT", "HEIGHT", App.Path & "\" & App.EXEName & ".INI"))
    If iL = 0 Then iL = 1000
    If iT = 0 Then iT = 1000
    If iW = 0 Then iW = 8150
    If iH = 0 Then iH = 6200
    Me.Move iL, iT, iW, iH
    tView.ImageList = imgList 'imgDatabase

    sBuf = Trim$(GetIni("DATA", "DEFDIR", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) = 0 Then sBuf = "..."
    txtServer.Text = sBuf
    LoadRoot
    
    Flex.Redraw = False
    Flex.Clear
    Flex.Rows = 1
    Flex.Cols = 1
    Flex.ColWidth(0) = Flex.Width
    Flex.Redraw = True
    
    sBuf = GetIni("CONNECT", "WELCOME", App.Path & "\" & App.EXEName & ".INI")
    If Val(sBuf) = 0 Then
        Call WriteIni("CONNECT", "WELCOME", "1", App.Path & "\" & App.EXEName & ".INI")
        Beep
        Screen.MousePointer = vbDefault
        mnuInfo_Click 2
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set collLock = Nothing
    
    If Not objServer Is Nothing Then objServer.DisConnect
    Set objServer = Nothing
    Dim sBuf As String
    If Me.WindowState <> vbMinimized Then
        Call WriteIni("DEFAULT", "LEFT", Format$(Me.Left), App.Path & "\" & App.EXEName & ".INI")
        Call WriteIni("DEFAULT", "TOP", Format$(Me.Top), App.Path & "\" & App.EXEName & ".INI")
        Call WriteIni("DEFAULT", "WIDTH", Format$(Me.Width), App.Path & "\" & App.EXEName & ".INI")
        Call WriteIni("DEFAULT", "HEIGHT", Format$(Me.Height), App.Path & "\" & App.EXEName & ".INI")
        Call WriteIni("DEFAULT", "BOL", App.HelpFile, App.Path & "\" & App.EXEName & ".INI")
        sBuf = txtServer.Text
        If sBuf <> "..." Then Call WriteIni("DATA", "DEFDIR", sBuf, App.Path & "\" & App.EXEName & ".INI")
  End If
  Set frmMain = Nothing

End Sub
Private Sub Form_Resize()
    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            SizeControls imgSplitter.Left
            
            'SizeControls
            b = False
        End If
    End If
End Sub

Private Sub Init_TView()
      
    Dim sOs As String
    Dim sErr As String
    
    sOs = " (Windows " & IIf(objServer.IsOS(SQLDMO_WINNT), "NT", "95/98") & ")"
        
    On Error GoTo ErrorHandle
    
    objServer.DisConnect
    objServer.ReConnect
  
    'add Root node - add server node
    LoadRoot
    Set mNode = tView.Nodes.Add("ServerGroup", tvwChild, "Server")
    With mNode
      .Tag = "Root" ' Set Tag property.
      .Image = "Server"  ' Set Image property.
      .Expanded = True
    End With
  '****************************************************************
  'add default child nodes
 
    Set mNode = tView.Nodes.Add("Server", tvwChild, "Database")
    mNode.Text = MyLoadResString(kTviewDatabases) '"Databases"
    mNode.Image = "Folder"   ' Set Image property.
    mNode.Expanded = True
  
    LoadDatabaseNodes  'load database nodes
      
    tView.Nodes("Server").Text = objServer.Name & sOs  'First node
    'tView.Nodes(1).Text = objServer.Name & sOs  'First node
  Exit Sub

ErrorHandle:
    sErr = Err.Description
 
    MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)

End Sub
Private Sub LoadDatabaseNodes()
    
    Dim bSysObj As Boolean
    Dim bAdd As Boolean, bErr As Boolean
    Dim bShowAll As Boolean
    Dim sUser As String
        
    sUser = objServer.TrueLogin
    bShowAll = CBool(Val(GetIni("CONNECT", "SHOWALL", App.Path & "\" & App.EXEName & ".INI")))
        
    Dim dmoDatabase As SQLDMO.Database2
    bSysObj = mnuF(2).Checked

    On Local Error Resume Next
    bErr = objServer.Databases.Count
    On Local Error GoTo 0
    
    For Each dmoDatabase In objServer.Databases     'Do for each Database in the Server
        Debug.Print dmoDatabase.Name
        
        If bSysObj Or dmoDatabase.SystemObject = 0 Then
            
            bAdd = True
            If Not bShowAll Then
                On Local Error Resume Next
                bAdd = dmoDatabase.IsUser(sUser) Or objServer.SaLogin
                bErr = Err.Number <> 0
                If bErr Then bAdd = False
                On Local Error GoTo 0
                If dmoDatabase.DBOption.Offline Then bAdd = True
            End If
            
            If bAdd Then
                                
                Set mNode = tView.Nodes.Add("Database", tvwChild, dmoDatabase.Name)   ' add child node
                mNode.Text = dmoDatabase.Name
                mNode.Image = "Database"  'Database icon
                'load sub nodes

                'If objServer.SaLogin Then
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Database Users")   ' add child node
                    mNode.Text = MyLoadResString(kTviewDatabasesUsers) '"Database Users"
                    mNode.Image = "Users"  'Folder icon
            
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Tables")
                    mNode.Text = MyLoadResString(kTviewDatabasesTables)  '"Tables"
                    mNode.Image = "Tables"  'Folder icon
            
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Views")
                    mNode.Text = MyLoadResString(kTviewDatabasesViews) '"Views"
                    mNode.Image = "Views"  'Folder icon
            
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Stored Procedures")
                    mNode.Text = MyLoadResString(kTviewDatabasesStoredProc) '"Stored Procedures"
                    mNode.Image = "Stored"  'Folder icon]
                    
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Roles")
                    mNode.Text = MyLoadResString(kTviewDatabasesRoles) '"Roles"
                    mNode.Image = "Roles"  'Folder icon
                
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "UDTs")
                    mNode.Text = MyLoadResString(kTviewDatabasesUDT) '"User Defined Data Types"
                    mNode.Image = "udt"  'Folder icon
                
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "UDFs")
                    mNode.Text = MyLoadResString(kTviewFunction) '"User Defined Data Types"
                    mNode.Image = "Func"  'Folder icon
                
                'End If
            End If
        End If
    Next
    'On Local Error GoTo 0
    Set dmoDatabase = Nothing
  
End Sub
Private Sub LoadLogsNodes()
    
    Dim oQry As SQLDMO.QueryResults
    Dim iCols As Integer, i As Integer
    Dim l As Long, lRows As Long
    Dim iLog As Integer
    Dim sBuf As String
    
    
    Set mNode = tView.Nodes.Add("Management", tvwChild, "Err Logs")
    mNode.Text = MyLoadResString(kTviewLogs) '"Sql Server Logs"
    mNode.Image = "Logs"   ' Set Image property.
    'mNode.Expanded = True
    
    On Local Error Resume Next
    Set oQry = objServer.EnumErrorLogs
    On Local Error GoTo 0
    
    If Not oQry Is Nothing Then
        With oQry
            If .ResultSets <> 0 Then
                'load sub nodes
                .CurrentResultSet = 1
                iCols = .Columns
                lRows = .Rows

                On Local Error Resume Next
                For l = 1 To lRows
                    Debug.Print .GetColumnLong(l, 1)
                    For i = 1 To iCols - 1
                        If i > 1 Then sBuf = sBuf & Space(5)
                        Select Case oQry.ColumnType(i)
'                            Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
'                                sBuf = sBuf & Format$(.GetColumnDate(1, i), "General Date")
                            
                            Case SQLDMO_DTypeUChar, SQLDMO_DTypeChar, SQLDMO_DTypeUVarchar, SQLDMO_DTypeVarchar
                                sBuf = sBuf & .GetColumnString(l, i)
                            'Aggiunto Case 2 perche' altrimenti non carica colonne Identity
                            'anche se NON e' un valore listato tra le costanti
                            Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                                iLog = .GetColumnLong(l, i)
                                sBuf = sBuf & IIf(iLog = 0, "Current - ", "Archive #") & iLog
                                'sBuf = sBuf & Format$(iLog, "#,0")
                        End Select
                    Next
                    Set mNode = tView.Nodes.Add("Err Logs", tvwChild, iLog & "*")  ' add child node
                    mNode.Text = sBuf
                    sBuf = ""
                Next
                
                On Local Error GoTo 0
            End If
        
        End With
    End If
    Set oQry = Nothing
  
End Sub

Private Sub mnuF_Click(Index As Integer)
    
    Select Case Index
        Case 0
            ActMenu 1, act_Server, "", Nothing
        Case 2
            mnuF(2).Checked = Not mnuF(2).Checked
            Call WriteIni("CONNECT", "SYSOBJ", Abs(mnuF(2).Checked), App.Path & "\" & App.EXEName & ".INI")
            LoadINIT
        Case 3
            LoadINIT
        Case 5
            fQuery.LoadDB
            fQuery.Show vbModal, Me
            'fQuery.Show vbModeless, Me
            'Refresh
            'LoadINIT
        Case 6
            fBCP.Show vbModal, Me
        Case 9
            Unload Me
        Case Else
    End Select
End Sub

Private Sub ListDevice()

    Set mNode = tView.Nodes.Add("Server", tvwChild, "Backup")
    mNode.Text = MyLoadResString(kTviewDevices) '"BackUp Device"
    mNode.Image = "Backup"   ' Set Image property.
    mNode.Expanded = True
    
End Sub

Private Sub mnuInfo_Click(Index As Integer)
    
    Dim sBuf As String
    Dim sPath As String
    Dim iFile As Integer
    Dim FP As FILE_PARAMS
    
    Select Case Index
        Case 0
            sBuf = "SELECT CASE " & _
                    "WHEN CHARINDEX( 'mssh' , p.net_library ) > 0 THEN 'Shared Memory' " & _
                    "WHEN CHARINDEX( 'nmpn' , p.net_library ) > 0 THEN 'Named Pipes' " & _
                    "WHEN CHARINDEX( 'msrp' , p.net_library ) > 0 THEN 'Multi-Protocol' " & _
                    "WHEN CHARINDEX( 'msso' , p.net_library ) > 0 THEN 'TCP/IP Sockets' " & _
                    "WHEN CHARINDEX( 'mssp' , p.net_library ) > 0 THEN 'Novell IPX/SPX' " & _
                    "Else: P.net_library END As ConnectionNetworkLibrary " & _
                    "FROM master..sysprocesses AS p WHERE p.spid = @@SPID"
            On Local Error Resume Next
            Dim oQry As SQLDMO.QueryResults
            Set oQry = objServer.Databases(1).ExecuteWithResults(sBuf)
            sBuf = ""
            sBuf = oQry.GetColumnString(1, 1)
            Set oQry = Nothing
            On Local Error GoTo 0
            If Len(sBuf) Then sBuf = vbCrLf & "Network Library: " & sBuf
            
            MsgBox GetPackage() & " Ver. " & objServer.VersionMajor & "." & objServer.VersionMinor & vbLf & objServer.VersionString & sBuf, vbInformation
        Case 1
            Screen.MousePointer = vbHourglass
            Dim oLang As cLang
            Set oLang = New cLang
            oLang.WriteRes2Ini True, True
            oLang.WriteRes2Ini False, True
            Set oLang = Nothing
'
'            WriteRes2Ini True, True
'            WriteRes2Ini False, True
            Screen.MousePointer = vbDefault
            MsgBox MyLoadResString(k_RegenLngFilesDONE)
            LoadStrings
        Case 2
            iFile = FreeFile
            sPath = App.Path
            If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
            On Local Error Resume Next
            Open sPath & "DbaMgr2k-EULA.txt" For Input As iFile
            If Err.Number = 0 Then sBuf = Trim$(Input$(LOF(iFile), iFile))
            On Local Error GoTo 0
            Close #iFile
            
            sBuf = Trim$(sBuf)
            If Len(sBuf) = 0 Then
                sBuf = Replace(MyLoadResString(k_EulaNotFound), "1%", App.EXEName) & vbCrLf & vbCrLf
                sBuf = sBuf & Replace(k_DbaMgr_Welcome1, "|", vbCrLf)
                sBuf = sBuf & Replace(k_DbaMgr_Welcome2, "|", vbCrLf)
                sBuf = sBuf & Replace(k_DbaMgr_Welcome3, "|", vbCrLf)
            End If
            
            fResult.Action() = act_Null
            fResult.WWrapVisible() = False
            fResult.tRes.Text = sBuf
                        
            fResult.Caption = Replace(MyLoadResString(k_EulaWelcome), "1%", App.EXEName & " - " & App.Major & "." & App.Minor & "." & App.Revision)
            Screen.MousePointer = vbDefault
            fResult.Show vbModal, Me
        Case 3
                        
            iFile = FreeFile
            sPath = App.Path
            If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
            
            On Local Error Resume Next
            Open sPath & "Depend2k.txt" For Input As iFile
            If Err.Number = 0 Then sBuf = Trim$(Input$(LOF(iFile), iFile))
            On Local Error GoTo 0
            Close #iFile
            
            sBuf = Trim$(sBuf)
            If Len(sBuf) = 0 Then sBuf = Replace(MyLoadResString(k_DependenciesNotFound), "1%", App.EXEName)
        
            fResult.Action() = act_Null
            fResult.WWrapVisible() = False
            
            fResult.tRes.Text = sBuf
            
            fResult.Caption = Replace(MyLoadResString(k_mnuDependencies), "1%", App.EXEName)
            Screen.MousePointer = vbDefault
            fResult.Show vbModal, Me
        Case 5
            FP.sFileRoot = "C:\"
            FP.sFileNameExt = "CliConfg.exe"
            If SearchSystemForFile(FP) Then
                Call Shell(FP.sResult, vbNormalFocus)
            End If
        
        Case 6
            FP.sFileRoot = "C:\"
            FP.sFileNameExt = "svrnetcn.exe"
            If SearchSystemForFile(FP) Then
                Call Shell(FP.sResult, vbNormalFocus)
            Else
                sBuf = MyLoadResString(kErr_AvalilableOnlyOnHostServer)
                If Not objServer Is Nothing Then
                    sBuf = Replace(sBuf, "1%", objServer.NetName)
                Else
                    sBuf = Replace(sBuf, "# 1% #", "")
                End If
                MsgBox sBuf, vbInformation, App.EXEName
            End If
            
        Case 9
            fCredits.Show vbModal, Me
        Case 10
            AbSplash.About Me
            'frmAbout.Show vbModal
    End Select
End Sub

Private Sub tView_DblClick()

  'check to see of server is running
    On Error Resume Next
    If tView.SelectedItem Is tView.Nodes("ServerGroup") Then
        If objServer Is Nothing Then ActMenu 1, act_Server, "", Nothing
    End If
End Sub

Private Sub tView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer ', bEnab As Boolean
    Dim iSubMenu As Integer
    Dim sBuf As String, sDB As String
    Dim iRow As Integer

    Dim iGo As optAction
        
    If Button = vbRightButton Then
        iRow = Flex.Row
        On Local Error Resume Next
        sBuf = tView.SelectedItem
        On Local Error GoTo 0
                
        'bEnab = Not objServer.SaLogin
        iGo = act_Null
                
        Dim M As New cPUpMenu

        Select Case sBuf
            Case "SqlServer"
                iGo = act_Server
                M.Add MyLoadResString(k_Connect_Edit_Connection_Properties), 1 ', bDisable:=Not objServer Is Nothing
                M.Add "", 2, bSeparator:=True
                M.Add MyLoadResString(k_Disconnect), 3, bDisable:=objServer Is Nothing
                M.Add "", 4, bSeparator:=True
                M.Add MyLoadResString(k_Property), 5, bDisable:=objServer Is Nothing Or Not GetMembership(SysAdm)
            
            Case MyLoadResString(kTviewDatabases) '"Databases"
                iGo = act_DBTasks
                M.Add MyLoadResString(k_New_DB), 1, bDisable:=Not GetMembership(DBCreat)
                M.Add "", 2, bSeparator:=True
                M.Add "BackUp Database", 3, bDisable:=Not GetMembership(SysAdm)     'prob. necessario
                M.Add "Restore Database", 4, bDisable:=Not GetMembership(SysAdm)    'meno privilegi, ma...
            Case MyLoadResString(kTviewDevices) '"BackUp Device"
                iGo = act_Bck
                M.Add MyLoadResString(k_New_BackUp_Device), 1, bDisable:=Not GetMembership(DiskAdm)
                M.Add "", 2, bSeparator:=True
                M.Add MyLoadResString(k_Clear_BackUp_History), 3, bDisable:=Not GetMembership(SysAdm)
            Case MyLoadResString(kTviewLogin) '"Login"
                iGo = act_Login
                M.Add MyLoadResString(k_New_Login), 1, bDisable:=Not GetMembership(SecurAdm)
            Case Else
                On Local Error Resume Next
                sBuf = tView.SelectedItem.Parent
                On Local Error GoTo 0
                If sBuf = MyLoadResString(kTviewDatabases) Then '"Databases" Then
                    sDB = tView.SelectedItem
                    iGo = act_DBOption
                    M.Add MyLoadResString(k_New_DB), 1, bDisable:=Not GetMembership(DBCreat)
                    M.Add "", 2, bSeparator:=True
                    M.Add MyLoadResString(k_Detach_DB), 3, bDisable:=Not GetMembership(SysAdm) Or Not GetMembership(db_dbo, sDB) Or iRow < 1
                    M.Add "", 4, bSeparator:=True
                    M.Add MyLoadResString(k_Remove), 5, bDisable:=Not GetMembership(SysAdm) Or Not GetMembership(db_dbo, sDB) Or iRow < 1
                    M.Add "", 6, bSeparator:=True
                    M.Add "Backup Database", 7, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "Restore Database", 8, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "Truncate Log", 9, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "", 10, bSeparator:=True
                    M.Add MyLoadResString(k_Shrink_DB), 11, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "Check Database Integrity >>", 12, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "CheckPoint Database", 13, bDisable:=Not GetMembership(db_BckOper, sDB) Or iRow < 1
                    M.Add "", 14, bSeparator:=True
                    M.Add MyLoadResString(k_Generate_SQL_Script), 15, bDisable:=iRow < 1 Or Not (IsDbUser(sDB))
                    M.Add "", 16, bSeparator:=True
                    M.Add MyLoadResString(k_HtmDocum), 17, bDisable:=iRow < 1 Or Not (IsDbUser(sDB))
                    M.Add "", 18, bSeparator:=True
                    'M.Add MyLoadResString(k_ModifyDB_Owner), 19, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
                    'sp3 requires sysadmin privileges
                    M.Add MyLoadResString(k_ModifyDB_Owner), 19, bDisable:=Not GetMembership(SysAdm) Or iRow < 1
                    
                    M.Add "", 20, bSeparator:=True
                    M.Add MyLoadResString(k_Extended_Property), 21, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
                    M.Add MyLoadResString(k_Property), 22, bDisable:=Not GetMembership(db_dbo, sDB) Or iRow < 1
                End If
        End Select
        
        If iGo <> act_Null Then
            Call GetCursorPos(P)
            i = M.Popup(Me.hWnd, P.x, P.y)
            M.Clear
            If i = 11 Then
                iSubMenu = 30
                
                M.Add MyLoadResString(k_Shrink_Default5%), 1    '"Shrink_Default (prova a 5%)"
                M.Add "Shrink_TruncateOnly", 2
                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
            ElseIf i = 12 Then
                iSubMenu = 40
                                
                M.Add "Check DB Repair_None", 1
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Fast", 2
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Rebuild", 3
                M.Add MyLoadResString(k_Shrink_SingleUser) & " Check DB Repair_Allow_DataLoss", 4

                Call GetCursorPos(P)
                i = M.Popup(Me.hWnd, P.x, P.y)
                M.Clear
            End If
            
            Set M = Nothing
            Refresh
            If i <> 0 Then
                If iGo = act_DBOption Then
                    ActMenu i + iSubMenu, iGo, tView.SelectedItem, Array(tView.SelectedItem)
                ElseIf iGo = act_Server Then
                    ActMenu i, iGo, "", Nothing
                ElseIf iGo = act_DBTasks Then
                    ActMenu i, iGo, "", Nothing
                ElseIf iGo = act_Bck Then
                    If i = 3 Then i = 9
                    'ActMenu i, iGo, tView.SelectedItem
                    ActMenu i, iGo, "", Array(tView.SelectedItem)
                ElseIf iGo = act_Login Then
                    ActMenu i, iGo, "", Nothing
                    'ActMenu i, iGo, ""
                End If
            End If
        End If
    End If

End Sub

Private Sub tView_NodeClick(ByVal Node As MSComctlLib.Node)

'    Dim i As Integer
  
    'On Error GoTo ErrorHandle
    'reset all open folder images
'    For i = 1 To tView.Nodes.Count
'        If tView.Nodes(i).Image = "OpenFolder" And tView.Nodes(i).Expanded = False Then
'            tView.Nodes(i).Image = "Folder"
'        End If
'    Next
    Erase m_sVntDetached
    Flex.Redraw = False
    Flex.Clear
    Flex.Rows = 1
    Flex.Cols = 1
    Flex.ColWidth(0) = Flex.Width
    Flex.Visible = True
    rtf.Visible = False
    
    Flex.Redraw = True
    Flex.MergeCells = flexMergeNever
    
    On Error GoTo ErrorHandle
    Select Case tView.SelectedItem
        Case MyLoadResString(kTviewFunction)
            Flex.Visible = True
            rtf.Visible = False
            'tView.SelectedItem.Image = "OpenFolder"
            tView.SelectedItem.Image = "Func"
            'load the table for the selected item
            LoadFunctions (tView.SelectedItem.Parent)
            
        Case MyLoadResString(kTviewDatabasesTables)  '"Tables"
            Flex.Visible = True
            rtf.Visible = False
            'tView.SelectedItem.Image = "OpenFolder"
            tView.SelectedItem.Image = "Tables"
            'load the table for the selected item
            LoadTables (tView.SelectedItem.Parent)
        Case MyLoadResString(kTviewDatabasesViews)  '"Views"
            Flex.Visible = True
            rtf.Visible = False
            tView.SelectedItem.Image = "Views"
            'load the views for the selected item
            LoadView (tView.SelectedItem.Parent)
        Case MyLoadResString(kTviewDatabasesUDT)  '"User Defined Data Types"
            Flex.Visible = True
            rtf.Visible = False
            tView.SelectedItem.Image = "udt"
            'load the views for the selected item
            loadUDT (tView.SelectedItem.Parent)
        
        Case MyLoadResString(kTviewDatabasesStoredProc)   '"Stored Procedures"
            Flex.Visible = True
            rtf.Visible = False
            tView.SelectedItem.Image = "Stored"
            'load the stored procedures for the selected item
            LoadStoredProc (tView.SelectedItem.Parent)
        Case MyLoadResString(kTviewDatabasesUsers)   '"Database Users"
            Flex.Visible = True
            rtf.Visible = False
            tView.SelectedItem.Image = "Users"
            'load the users for the selected item
            LoadDatabaseUsers (tView.SelectedItem.Parent)
        Case MyLoadResString(kTviewDatabasesRoles)   '"Roles"
            Flex.Visible = True
            rtf.Visible = False
            tView.SelectedItem.Image = "Roles"
            'load the stored procedures for the selected item
            LoadRoles (tView.SelectedItem.Parent)
        Case MyLoadResString(kTviewDatabases) '"Databases"
            Flex.Visible = True
            rtf.Visible = False
            'tView.SelectedItem.Image = "OpenFolder"
            LoadDatabases
        Case MyLoadResString(kTviewLogin) '"Login"
            Flex.Visible = True
            rtf.Visible = False
            'tView.SelectedItem.Image = "OpenFolder"
            LoadLogin
        Case MyLoadResString(kTviewDevices) '"BackUp Device"
            Flex.Visible = True
            rtf.Visible = False
            'tView.SelectedItem.Image = "OpenFolder"
            LoadBackup
        Case MyLoadResString(k_Detached_Databases)
            Flex.Visible = True
            rtf.Visible = False
            LoadDetachedNode
        Case MyLoadResString(k_RES_Process_Info)
            Flex.Visible = True
            rtf.Visible = False
            LoadProcInfo
        Case Else
            Flex.Clear
            Flex.Rows = 1
            If Not tView.SelectedItem.Parent Is Nothing Then
                If tView.SelectedItem.Parent = MyLoadResString(kTviewDatabases) Then '"Databases" Then
                    Flex.Visible = True
                    rtf.Visible = False
                    Load_DBFlex tView.SelectedItem
                ElseIf tView.SelectedItem.Parent = MyLoadResString(kTviewLogs) Then '"Sql Server Logs" Then
                    rtf.Visible = True
                    Flex.Visible = False
                    Load_ErrLogNo Val(tView.SelectedItem.Key)
                ElseIf tView.SelectedItem.Parent = MyLoadResString(kTviewLockID) Then '"Locks /ID" Then
                    Flex.Visible = True
                    rtf.Visible = False
                    Load_Spid tView.SelectedItem.Key
                End If
            End If
            'do nothing
    End Select
    Exit Sub
    
ErrorHandle:
    Debug.Print Err.Description
    Select Case Err.Number
        Case 91  'tree view not initialized
            'Exit Sub
    End Select
    If Flex.Visible Then Flex.Redraw = True
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadDatabases()
    
    Dim bSysObj As Boolean
    Dim dmoDB As SQLDMO.Database2
    Dim bShowAll As Boolean, bAdd As Boolean
    Dim sUser As String
    Dim sDbName As String
    Dim lErr As Long, sErr As String
        
    sUser = objServer.TrueLogin
    bShowAll = CBool(Val(GetIni("CONNECT", "SHOWALL", App.Path & "\" & App.EXEName & ".INI")))
    
    bSysObj = mnuF(2).Checked
  
                
    With Flex
        .Clear
        .Redraw = False
        .Cols = 5
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 2500
        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = MyLoadResString(k_Create_Date) '"Create Date"
        .ColWidth(2) = 3500
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(3) = 1000
        .TextArray(3) = MyLoadResString(k_CompatLevel) '"Compatibility Level"
        .ColWidth(4) = 3500
        .TextArray(4) = MyLoadResString(k_NewDbSort) 'SortOrder

        .ToolTipText = MyLoadResString(k_RES_Database) '"User Databases"
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        On Local Error GoTo ErrorHandle
        For Each dmoDB In objServer.Databases
            If bSysObj Or dmoDB.SystemObject = 0 Then
                bAdd = True
                
                If Not bShowAll Then
                    On Local Error Resume Next
                    bAdd = dmoDB.IsUser(sUser) Or objServer.SaLogin
                    If Err.Number <> 0 Then bAdd = False
                    On Local Error GoTo 0
                End If
                                
                If bAdd Then
                    
                    On Error Resume Next
                    sErr = dmoDB.CompatibilityLevel
                    lErr = Err.Number
                    On Local Error GoTo 0
                    sDbName = dmoDB.Name
                    If lErr = 0 Then
                        .AddItem sDbName & vbTab & dmoDB.Owner & vbTab & dmoDB.CreateDate & vbTab & dmoDB.CompatibilityLevel & vbTab & dmoDB.Collation
                    Else
                        .AddItem sDbName & vbTab & dmoDB.Owner & vbTab & dmoDB.CreateDate & vbTab & Replace(MyLoadResString(k_DB_is_OffLine), "1%", sDbName) & vbTab & Replace(MyLoadResString(k_DB_is_OffLine), "1%", sDbName)
                    End If
                End If
            End If
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    
    Set dmoDB = Nothing
    Exit Sub
    
ErrorHandle:
    lErr = Err.Number
    sErr = Err.Description
    Set dmoDB = Nothing
    SetFixedRows Flex
    Flex.Redraw = True
    
    Select Case lErr
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub

Private Sub LoadDatabaseUsers(ByVal strDatabaseName As String)
     
    Dim dmoUser As SQLDMO.User
    Dim sErr As String
    
    Flex.Clear
    On Error GoTo ErrorHandle
        
    With Flex
        .Redraw = False
        .Cols = 2
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_LoginName) '"Login Name"
        .ColWidth(1) = 1500
        .ToolTipText = strDatabaseName & " " & MyLoadResString(k_Users) '" Users"
        'extra colmun to fill out rest of space
        '.ColWidth(2) = Me.Width
        objServer.Databases(strDatabaseName).Users.Refresh True
        For Each dmoUser In objServer.Databases(strDatabaseName).Users
            .AddItem dmoUser.Name & vbTab & dmoUser.Login ', 1
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    Set dmoUser = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoUser = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub
Private Sub Load_ErrLogNo(ByVal iLogNo As Integer)
     
    Dim oQry As SQLDMO.QueryResults
    Dim lW As Long
    Dim sErr As String
    
    Screen.MousePointer = vbHourglass
    With rtf
        .Text = ""
        
        On Error GoTo ErrorHandle
        Set oQry = objServer.ReadErrorLog(iLogNo)
            
        If Not oQry Is Nothing Then
            If oQry.ResultSets <> 0 Then
                            
                oQry.CurrentResultSet = 1
                .Text = oQry.GetRangeString()
                lW = InStr(.Text, vbCrLf)
                If lW > 0 Then .RightMargin = Me.TextWidth("w") * (lW + 2)
                
            End If
        End If
    End With
    Set oQry = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandle:
    Print Err.Number, Err.Description, Err.Source
    sErr = Err.Description
    
    Flex.Redraw = True
    Set oQry = Nothing
    Screen.MousePointer = vbDefault
    
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub

Private Sub LoadStoredProc(ByVal strDatabaseName As String)

    Dim bSysObj As Boolean
    Dim dmoStoredProcedure As SQLDMO.StoredProcedure
    Dim sErr As String
    
    bSysObj = mnuF(2).Checked

    Flex.Clear
    On Error GoTo ErrorHandle


    With Flex
        .Redraw = False
        .Cols = 3
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_Create_Date)  '"Create Date"
        .ColWidth(2) = 2500
        .ColAlignment(2) = flexAlignLeftCenter
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewDatabasesStoredProc) '" Stored Procedures"
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        objServer.Databases(strDatabaseName).StoredProcedures.Refresh True
        For Each dmoStoredProcedure In objServer.Databases(strDatabaseName).StoredProcedures
            If bSysObj Or dmoStoredProcedure.SystemObject = 0 Then
                .AddItem dmoStoredProcedure.Name & vbTab & dmoStoredProcedure.Owner & vbTab & dmoStoredProcedure.CreateDate ', 1
            End If
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    Set dmoStoredProcedure = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoStoredProcedure = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub
Private Sub LoadView(strDatabaseName As String)
    
    Dim dmoView As SQLDMO.View2
    Dim bSysObj As Boolean
    Dim sErr As String
    
    bSysObj = mnuF(2).Checked
  
    Flex.Clear
    On Error GoTo ErrorHandle
  
    With Flex
        .Redraw = False
        .Cols = 4
        .Rows = 1
        If .Rows > 1 Then .FixedRows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_Func_Param_IsSchemaBound)
        .TextArray(3) = MyLoadResString(k_Create_Date) '"Create Date"
        .ColWidth(3) = 2500
        .ColAlignment(2) = flexAlignLeftCenter
        'extra colmun to fill out rest of space
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewDatabasesViews) '" Views"
        '.ColWidth(3) = Me.Width
        objServer.Databases(strDatabaseName).Views.Refresh True
        For Each dmoView In objServer.Databases(strDatabaseName).Views
            If bSysObj Or dmoView.SystemObject = 0 Then
                Debug.Print dmoView.IsSchemaBound
                .AddItem dmoView.Name & vbTab & dmoView.Owner & vbTab & vbTab & dmoView.CreateDate ', 1
                .Row = .Rows - 1
                .Col = 2
                Set Flex.CellPicture = LoadResPicture(IIf(dmoView.IsSchemaBound, k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            End If
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    Set dmoView = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoView = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub
Private Sub loadUDT(ByVal strDatabaseName As String)
    
    Dim dmoUDT As SQLDMO.UserDefinedDatatype2
    Dim bIsNum As Boolean
    Dim sErr As String
      
    Flex.Clear
    On Error GoTo ErrorHandle
  
    With Flex
        .Redraw = False
        .Cols = 8
        .Rows = 1
        If .Rows > 1 Then .FixedRows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 2500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_BaseType) '"Base Type"
        .ColWidth(2) = 1500
        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = MyLoadResString(k_Lenght) '"Lenght"
        .ColWidth(3) = 1000
        .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = MyLoadResString(k_Allow_Null) '"Allow Null"
        .ColWidth(4) = 1000
        .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = MyLoadResString(k_Default_Value) '"Default"
        .ColWidth(5) = 1000
        .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = MyLoadResString(k_Rule) '"Rule"
        .ColWidth(6) = 1000
        .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = MyLoadResString(k_NewDbSort)
        .ColWidth(7) = 1000
        .ColAlignment(7) = flexAlignLeftCenter
        
        'extra colmun to fill out rest of space
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewDatabasesUDT) '" UDTs"
        '.ColWidth(3) = Me.Width
        objServer.Databases(strDatabaseName).UserDefinedDatatypes.Refresh True
        For Each dmoUDT In objServer.Databases(strDatabaseName).UserDefinedDatatypes
                        
            bIsNum = objServer.Databases(strDatabaseName).SystemDatatypes(dmoUDT.BaseType).IsNumeric
            
            .AddItem dmoUDT.Name & vbTab & dmoUDT.Owner & vbTab & dmoUDT.BaseType & vbTab & IIf(bIsNum, Format$(dmoUDT.NumericPrecision) & "," & Format$(dmoUDT.NumericScale), dmoUDT.MaxSize) & vbTab & vbTab & dmoUDT.Default & vbTab & dmoUDT.Rule & vbTab & dmoUDT.Collation
            .Row = .Rows - 1
            .Col = 4
            Set Flex.CellPicture = LoadResPicture(IIf(dmoUDT.AllowNulls, k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    
    Set dmoUDT = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoUDT = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub

Private Sub LoadTables(ByVal strDatabaseName As String)
    
    Dim dmoTable As SQLDMO.Table
    Dim bSysObj As Boolean
    Dim sErr As String
    
    bSysObj = mnuF(2).Checked
  
    Flex.Clear
    On Error GoTo ErrorHandle
      
    With Flex
        .Redraw = False
        .Cols = 3
        .Rows = 1
        If .Rows > 1 Then .FixedRows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_Create_Date)  '"Create Date"
        .ColWidth(2) = 2500
        .ColAlignment(2) = flexAlignLeftCenter
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewDatabasesTables) '" Tables"
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        objServer.Databases(strDatabaseName).Tables.Refresh True
        For Each dmoTable In objServer.Databases(strDatabaseName).Tables
            If bSysObj Or dmoTable.SystemObject = 0 Then
                .AddItem dmoTable.Name & vbTab & dmoTable.Owner & vbTab & dmoTable.CreateDate ', 1
            End If
        Next
        
        
        SetFixedRows Flex
        .Redraw = True

    End With
 
    Set dmoTable = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoTable = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub
Private Sub LoadFunctions(ByVal strDatabaseName As String)
        
    Dim dmoFunc As SQLDMO.UserDefinedFunction
    Dim bSysObj As Boolean
    Dim oDb As SQLDMO.Database2
    Dim sErr As String
    
    bSysObj = mnuF(2).Checked
  
    Flex.Clear
    On Error GoTo ErrorHandle
      
    With Flex
        .Redraw = False
        .Cols = 4
        .Rows = 1
        If .Rows > 1 Then .FixedRows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Owner) '"Owner"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_Create_Date)  '"Create Date"
        .ColWidth(2) = 2500
        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = MyLoadResString(k_Type)
        .ColWidth(3) = 1000
        
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewFunction)
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        Set oDb = objServer.Databases(strDatabaseName)
        oDb.UserDefinedFunctions.Refresh True
        
        For Each dmoFunc In oDb.UserDefinedFunctions
            If bSysObj Or dmoFunc.SystemObject = 0 Then
                .AddItem dmoFunc.Name & vbTab & dmoFunc.Owner & vbTab & dmoFunc.CreateDate & vbTab & UDFtype(dmoFunc.Type)
            End If
        Next
        
        SetFixedRows Flex
        .Redraw = True

    End With
 
    Set dmoFunc = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoFunc = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub


Private Sub ListLogin()

    Set mNode = tView.Nodes.Add("Server", tvwChild, "Users")
    mNode.Text = MyLoadResString(kTviewLogin) '"Login"
    mNode.Image = "Logins"   ' Set Image property.
    mNode.Expanded = True

End Sub

Private Sub LoadLogin()
        
    Dim oLog As SQLDMO.Login
    Dim sType As String
    Dim sAccess As String
    Dim sErr As String
     
    Flex.Clear
    On Error GoTo ErrorHandle
      
    With Flex
        .Redraw = False
        .Cols = 5
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3000
        .TextArray(1) = MyLoadResString(k_Type)  '"Type"
        .ColWidth(1) = 1500
        .TextArray(2) = MyLoadResString(k_ServerAccess)
        .ColWidth(2) = 1200
        
        .TextArray(3) = MyLoadResString(k_Default_Database) '"Default Database"
        .ColWidth(3) = 2500
        .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = MyLoadResString(k_Language) '"Language"
        .ColWidth(3) = 2500
        .ToolTipText = ""
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        objServer.Logins.Refresh True
        For Each oLog In objServer.Logins
            With oLog
                Select Case .Type
                    Case SQLDMOLogin_NTGroup
                        sType = "NTGroup"
                    Case SQLDMOLogin_NTUser
                        sType = "NTUser"
                    Case SQLDMOLogin_Standard
                        sType = "Standard"
                End Select
                sAccess = IIf(oLog.DenyNTLogin, MyLoadResString(k_AccessDeny), MyLoadResString(k_AccessPermit))
                Flex.AddItem .Name & vbTab & sType & vbTab & sAccess & vbTab & .Database & vbTab & .LanguageAlias '.Language
            End With
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    Set oLog = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Set oLog = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub
Private Sub LoadBackup()
        
    Dim oBack As SQLDMO.BackupDevice
    Dim sType As String
    Dim sErr As String
  
    Flex.Clear
    On Error GoTo ErrorHandle
      
    With Flex
        .Redraw = False
        .Cols = 3
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 1500
        .TextArray(1) = MyLoadResString(k_Physical_Location) '"Physical Location"
        .ColWidth(1) = 3500
        .TextArray(2) = MyLoadResString(k_Type) '"Type"
        .ColWidth(2) = 2500
        .ColAlignment(2) = flexAlignLeftCenter
        .ToolTipText = ""
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        objServer.BackupDevices.Refresh True
        For Each oBack In objServer.BackupDevices
        
            With oBack
                Select Case .Type
                    Case SQLDMODevice_CDROM
                        sType = "CDROM"
                    Case SQLDMODevice_DiskDump
                        sType = "DiskDump"
                    Case SQLDMODevice_FloppyADump
                        sType = "FloppyADump"
                    Case SQLDMODevice_FloppyBDump
                        sType = "FloppyBDump"
                    Case SQLDMODevice_PipeDump
                        sType = "PipeDump"
                    Case SQLDMODevice_TapeDump
                        sType = "TapeDump"
                    Case SQLDMODevice_Unknown
                        sType = "Unknown"
                        
                End Select
                Flex.AddItem .Name & vbTab & .PhysicalLocation & vbTab & sType
            End With
        Next

        SetFixedRows Flex
        .Redraw = True
    End With
    Set oBack = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Set oBack = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub

Private Sub AttachDB(ByVal sDbName As String, ByVal sPhysicalName As String)

    Dim bExec As Boolean
    Dim sDbstr As String
    Dim lErr As Long, sErr As String
    Dim objDataF As cDetGrp
    
    If InStr(sDbName, ".") Then sDbstr = Left$(sDbName, InStr(sDbName, ".") - 1)
    
    Screen.MousePointer = vbHourglass
    Set objDataF = New cDetGrp
    objDataF.PrimaryMdfFile() = sPhysicalName
    If objDataF.AreErrors() Then
        lErr = objDataF.Errors(sErr)
    Else
        If objDataF.AreWarnings() Then
            lErr = objDataF.Warnings(sErr)
            Screen.MousePointer = vbDefault
            MsgBox MyLoadResString(k_GenWarning) & vbCrLf & lErr & " - " & sErr, vbInformation, App.EXEName
            lErr = 0
            sErr = ""
        End If
        Screen.MousePointer = vbHourglass
        fAttach.Stuff(sPhysicalName, objDataF) = sDbstr
        fAttach.Show vbModal, Me
    
        bExec = fAttach.Done()
        Unload fAttach
    End If
    Set objDataF = Nothing
    
    DoEvents
    If bExec Then
        LoadINIT
        stbMain.Panels(1).Text = "Ready"
        Refresh
    End If
    Screen.MousePointer = vbDefault
    If Len(sErr) <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
    
End Sub
Private Sub Load_DBFlex(ByVal sDbName As String)
    
    Dim oDb As SQLDMO.Database2
    Dim sType As String
    Dim bRed As Boolean
    Dim bOwner As Boolean
    Dim i As Integer
    Dim sErr As String
    
    Flex.Clear
    On Error GoTo ErrorHandle
        
    With Flex
        .Redraw = False
        .Cols = 2
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Property)  ' "Proprietà"
        .ColWidth(0) = 1500
        .TextArray(1) = MyLoadResString(k_Value)  '"Valore"
        .ColWidth(1) = 3500
        .ColAlignment(1) = flexAlignLeftCenter
        .ToolTipText = ""
        'extra colmun to fill out rest of space
        '.ColWidth(3) = Me.Width
        'For Each oBack In objServer.BackupDevices
        Set oDb = objServer.Databases(sDbName)
        With oDb
            'Flex.AddItem "DataBase" & vbTab & .Name
            Flex.AddItem MyLoadResString(k_RES_Database) & vbTab & .Name
            Flex.AddItem MyLoadResString(k_Create_Date) & vbTab & .CreateDate
            Flex.AddItem MyLoadResString(k_Owner) & vbTab & .Owner
                        
            Select Case .Status
                Case SQLDMODBStat_Normal
                    'sType = "Normal"
                    i = k_DBStatus_Normal
                Case SQLDMODBStat_Offline
                    'sType = "Off Line"
                    i = k_DBStatus_OffLine
                Case SQLDMODBStat_Recovering
                    'sType = "Recovering"
                    i = k_DBStatus_Recovering
                Case SQLDMODBStat_Standby
                    'sType = "Stand By"
                    i = k_DBStatus_StandBy
                Case SQLDMODBStat_Suspect
                    'sType = "Suspect"
                    i = k_DBStatus_Suspect
                    bRed = True
                Case SQLDMODBStat_Inaccessible
                    'sType = "Inaccessible"
                    i = k_DBStatus_Inaccessible
                    bRed = True
                Case Else
                    'sType = "Unknown"
                    i = k_DBStatus_UnKnown
                    bRed = True
            End Select
            
            'Flex.AddItem "Status" & vbTab & sType
            Flex.AddItem "Status" & vbTab & MyLoadResString(i)
            If bRed Then
                Flex.Col = 1
                Flex.Row = Flex.Rows - 1
                Flex.CellBackColor = vbHighlight
                Flex.CellForeColor = vbHighlightText
            End If

            If .DBOption.DBOUseOnly Then
                Flex.AddItem "Status" & vbTab & MyLoadResString(k_DBStatus_dboUseOnly)
                Flex.Col = 1
                Flex.Row = Flex.Rows - 1
                Flex.CellBackColor = vbHighlight
                Flex.CellForeColor = vbHighlightText
            End If
            
            If .DBOption.ReadOnly Then
                Flex.AddItem "Status" & vbTab & MyLoadResString(k_DBStatus_ReadOnly)
                Flex.Col = 1
                Flex.Row = Flex.Rows - 1
                Flex.CellBackColor = vbHighlight
                Flex.CellForeColor = vbHighlightText
            End If
            If .DBOption.SingleUser Then
                Flex.AddItem "Status" & vbTab & MyLoadResString(k_DBStatus_SingleUser)
                Flex.Col = 1
                Flex.Row = Flex.Rows - 1
                Flex.CellBackColor = vbHighlight
                Flex.CellForeColor = vbHighlightText
            End If
            
            On Local Error Resume Next
            bOwner = .Isdb_owner Or objServer.SaLogin
            If Err.Number <> 0 Then bOwner = False
            On Local Error GoTo 0
            
            
            If bOwner Then
                sType = .PrimaryFilePath
            Else
                sType = MyLoadResString(k_Only_DB_Owner_has_access_to_this_function)
            End If
            Flex.AddItem MyLoadResString(k_DB_PrimaryFilePath) & vbTab & sType  '.PrimaryFilePath
            
            If bOwner Then
                sType = .TransactionLog.LastBackup
            Else
                sType = MyLoadResString(k_Only_DB_Users_has_access_to_this_function)
            End If
            Flex.AddItem MyLoadResString(k_Last_BackUp) & vbTab & sType
            
            If bOwner Then
                sType = .Users.Count
            Else
                sType = MyLoadResString(k_Only_DB_Users_has_access_to_this_function)
            End If
            Flex.AddItem MyLoadResString(k_Users) & vbTab & sType
            
            If bOwner Then
                sType = .Size
            Else
                sType = MyLoadResString(k_Only_DB_Users_has_access_to_this_function)
            End If
            Flex.AddItem MyLoadResString(k_Size) & " (Mb)" & vbTab & sType
            
            If bOwner Then
                sType = .SpaceAvailableInMB
            Else
                sType = MyLoadResString(k_Only_DB_Users_has_access_to_this_function)
            End If
            Flex.AddItem MyLoadResString(k_Available_Space_Mb) & vbTab & sType
            
            
            If bOwner Then
                sType = .CurrentCompatibility
                Flex.AddItem MyLoadResString(k_CompatLevel) & vbTab & sType
            End If
            
            If bOwner Then
                'Dim oDbOpt As SQLDMO.DBOption2
                'Set oDbOpt = oDb.DBOption
                sType = .Collation
                Flex.AddItem MyLoadResString(k_NewDbSort) & vbTab & sType
            End If
            
        End With
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        For i = 1 To .Rows - 1
            .MergeRow(i) = True
        Next
        
        If .Rows > 1 Then .FixedRows = 1
        .Redraw = True
    End With
    Set oDb = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set oDb = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select

End Sub
Private Sub LoadINIT()
    
    Screen.MousePointer = vbHourglass
    With Flex
        .Redraw = False
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 1
        .Cols = 1
        .Clear
        .Redraw = True
    End With
    
    Init_TView
    ListDevice
    ListLogin
    LoadActivity
        
    Set mNode = tView.Nodes.Add("ServerGroup", tvwChild, "Detached")
    mNode.Text = MyLoadResString(k_Detached_Databases)
    mNode.Image = "detached"   ' Set Image property.
    Flex.ToolTipText = ""

    Screen.MousePointer = vbDefault
End Sub
Private Sub TruncateLog(ByVal sDbName As String)

    Dim lErr As Long, sErr As String
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    objServer.Databases(sDbName).TransactionLog.Truncate
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If

End Sub

Private Sub GenerateSQL(ByVal sDbName As String, ByVal sObjectName As String, ByVal oType As optAction)

    Dim lErr As Long, sErr As String
    Dim oDb As Database2
    Dim oObj As Object
    Dim oScr As cScrpO
    Dim bLoadParam As Boolean
    Dim iType As optScript
        
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    lErr = Err.Number
    If lErr = 0 Then
        Select Case oType
            Case act_SP
                oDb.StoredProcedures.Refresh True
                Set oObj = GetSpByName(objServer, sDbName, sObjectName)
                'Set oObj = oDB.StoredProcedures(sObjectName)
                iType = opt_SP
            Case act_View
                oDb.Views.Refresh True
                'Set oObj = oDB.Views(sObjectName)
                Set oObj = GetViewByName(objServer, sDbName, sObjectName)
                iType = opt_View
            Case act_Tbl
                oDb.Tables.Refresh True
                Set oObj = GetTbByName(objServer, sDbName, sObjectName)
                'Set oObj = oDB.Tables(sObjectName)
                iType = opt_Tbl
            Case act_UDT
                oDb.UserDefinedDatatypes.Refresh True
                Set oObj = GetUdtByName(objServer, sDbName, sObjectName)
                'Set oObj = oDB.UserDefinedDatatypes(sObjectName)
                iType = opt_UDT
            Case act_User_Function
                oDb.UserDefinedFunctions.Refresh True
                Set oObj = GetFuncByName(objServer, sDbName, sObjectName)
                'Set oObj = oDB.UserDefinedDatatypes(sObjectName)
                iType = opt_Func
                
        End Select
        lErr = Err.Number
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        If oType <> act_DBOption Then
            bLoadParam = False
            If iType <> opt_UDT Then
                If oObj.SystemObject = True Then
                    On Local Error Resume Next
                    lErr = 425
                    Err.Raise lErr, Description:=MyLoadResString(k_Unavailable_for_System_Objects)
                    sErr = Err.Description
                    On Local Error GoTo 0
                End If
            End If
            If lErr = 0 Then
                With oObj
                    Set oScr = New cScrpO
                    oScr.oName() = .Name
                    oScr.oOwner() = .Owner
                    oScr.oType() = iType
                End With
            End If
        Else
            bLoadParam = True
        End If
    End If
    Set oObj = Nothing
    Set oDb = Nothing
    
    If lErr = 0 Then
        fScript.Caption = MyLoadResString(k_Generate_SQL_Script) & " - DB: [" & sDbName & "]"
        fScript.TypeOfScript() = oType
        If oType <> act_DBOption Then fScript.ScriptThis() = oScr
        Set oScr = Nothing

        fScript.DB2Script() = sDbName
        fScript.Init bLoadParam, bLoadParam, bLoadParam, bLoadParam, bLoadParam, bLoadParam, bLoadParam, bJustLoaded:=Not bLoadParam
    
        If Not fScript.ErrOccured() Then
            fScript.Show vbModal
        Else
            lErr = fScript.ErrCode()
            sErr = fScript.ErrMsg()
            Unload fScript
        End If
    End If
    Set oScr = Nothing
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub

Private Sub LoadRoles(ByVal strDatabaseName As String)
        
    Dim dmoR As SQLDMO.DatabaseRole
    Dim sErr As String
    
    Flex.Clear
    On Error GoTo ErrorHandle
        
    With Flex
        .Redraw = False
        .Cols = 2
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .ColWidth(0) = 3500
        .TextArray(1) = MyLoadResString(k_Role_Type)  '"Role Type"
        .ColWidth(1) = 1500
        .ToolTipText = strDatabaseName & " " & MyLoadResString(kTviewDatabasesRoles) '" Roles"
        'extra colmun to fill out rest of space
        '.ColWidth(2) = Me.Width
        objServer.Databases(strDatabaseName).DatabaseRoles.Refresh True
        For Each dmoR In objServer.Databases(strDatabaseName).DatabaseRoles
            .AddItem dmoR.Name & vbTab & IIf(dmoR.AppRole, "Application", "Standard") ',1
        Next
        
        SetFixedRows Flex
        .Redraw = True
    End With
    Set dmoR = Nothing
    Exit Sub
    
ErrorHandle:
    sErr = Err.Description
    Flex.Redraw = True
    Set dmoR = Nothing
    Select Case Err.Number
        Case 381
            Exit Sub   'no data to load
        Case Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End Select
End Sub

Sub SizeControls(x As Single)
    'APP WIZARD
    On Error Resume Next
    
    'set the width
    If Me.WindowState <> vbMinimized Then
        If x < 1500 Then x = 1500
        If x > (Me.Width - 1500) Then x = Me.Width - 1500
        tView.Width = x
        imgSplitter.Left = x
    End If
      
    Flex.Left = x + 40
    Flex.Width = Me.Width - (tView.Width + 240)
    rtf.Left = Flex.Left
    rtf.Width = Flex.Width
    fra.Left = Flex.Left
    fra.Width = Flex.Width
    txtServer.Width = fra.Width - 2 * cmd(0).Width - txtServer.Left
    cmd(0).Left = txtServer.Left + txtServer.Width + 50
        
    'set the height
    
    tView.Height = Me.ScaleHeight - (fra.Top + stbMain.Height + 100)
        
    'Flex.Height = Me.ScaleHeight - (fra.Top + fra.Height + stbMain.Height + 100)
    Flex.Height = tView.Height - (fra.Top + fra.Height + 40)
    Flex.Top = fra.Top + fra.Height + 60
    rtf.Top = fra.Top + fra.Height + 60
    rtf.Height = Flex.Height
    imgSplitter.Top = fra.Top
    imgSplitter.Height = Flex.Height
'    stbMain.Panels(3).Width = stbMain.Width \ 8
'    stbMain.Panels(2).Width = stbMain.Panels(3).Width
    
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    With imgSplitter
        picSplitter.Move .Left, tView.Top, .Width \ 2, Flex.Height + fra.Height '- 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub LoadRoot()
    Dim bMenu As Boolean
    Dim bLocalHost As Boolean
    
    bMenu = Not objServer Is Nothing
    
    
    Me.Caption = App.EXEName
    Me.Caption = Caption & " (Ver." & App.Major & "." & App.Minor & "." & App.Revision & ")"
    If Not bMenu Then
        'Me.Caption = Caption & " (Ver." & App.Major & "." & App.Minor & "." & App.Revision & ")"
        bLocalHost = True
    Else
        'Caption = Caption & " running on .. " & GetPackage()
        Me.Caption = Me.Caption & " " & MyLoadResString(k_Running_on) & " " & GetPackage()
        bLocalHost = StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0
    End If
    tView.Nodes.Clear
    'add Root node - add server node
    Set mNode = tView.Nodes.Add(, , "ServerGroup")
    With mNode
      .Text = "SqlServer"
      .Tag = "MasterRoot" ' Set Tag property.
      .Image = "ServerGroup"  ' Set Image property.
      .Expanded = True
    End With
    mnuF(2).Enabled = bMenu
    mnuF(3).Enabled = bMenu
    mnuF(5).Enabled = bMenu
    mnuF(6).Enabled = bMenu
    mnuInfo(0).Enabled = bMenu
    
End Sub

Private Sub LoadStrings()

    Dim sBuf As String
    sBuf = GetIni("DEFAULT", "BOL", App.Path & "\" & App.EXEName & ".INI")
    If Len(Trim$(sBuf)) = 0 Then sBuf = App.HelpFile
    App.HelpFile = sBuf

    mnu(0).Caption = MyLoadResString(k_mnuActivity)
    mnuF(0).Caption = MyLoadResString(k_Connect_Edit_Connection_Properties)
    mnuF(2).Caption = MyLoadResString(k_mnuInclude_System_Objects)
    mnuF(3).Caption = MyLoadResString(k_Refresh)
    mnuF(5).Caption = MyLoadResString(k_mnuQuery)
    mnuF(6).Caption = MyLoadResString(k_mnuBCP)
    
    mnuF(9).Caption = MyLoadResString(k_End)
    mnu(1).Caption = MyLoadResString(k_mnuInfo)
    mnuInfo(0).Caption = MyLoadResString(k_mnuEngineVersion)
    mnuInfo(1).Caption = MyLoadResString(k_mnuRegenLngFiles)
    mnuInfo(2).Caption = MyLoadResString(k_mnuLicense)
    mnuInfo(3).Caption = Replace(MyLoadResString(k_mnuDependencies), "1%", App.EXEName)
    
    mnuInfo(5).Caption = MyLoadResString(k_mnuCliConfg)
    mnuInfo(6).Caption = MyLoadResString(k_mnuSrvConfg)
    
    mnuInfo(9).Caption = MyLoadResString(k_Credits)
    mnuInfo(10).Caption = MyLoadResString(k_mnuAbout)
    
    txtServer.ToolTipText = MyLoadResString(k_Full_Data_PathToolTip)
    cmd(0).ToolTipText = MyLoadResString(k_Full_Data_PathToolTip)
    
    fra.Caption = MyLoadResString(k_Full_Data_Path)

End Sub
Private Sub ActMenu(ByVal iAct As Integer, ByVal iOpt As optAction, ByVal sDbName As String, ByVal vDati As Variant)

    Dim bRefresh As Boolean
    Dim bConfirm As Boolean
    Dim bCheckPoint As Boolean
    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim objDependent As clsDepen
    Dim sRet As String
    Dim iEl As Integer, iCount As Integer, iLoop As Integer, iIndex As Integer
    Dim lWait As Long
    Dim sAttachSQL As String
    Dim sAttachMSG As String
    Dim oView As SQLDMO.View
    Dim oTb As SQLDMO.Table
    Dim objFunc As SQLDMO.UserDefinedFunction
    Dim objSp As SQLDMO.StoredProcedure
    Dim sOwner As String
    Dim iEnDetachMode As amDetachingMode
    Dim sCurError As String
    Dim bWarning As Boolean
    Dim bIsReadOnly As Boolean
            
    Select Case iOpt
        Case act_ProcInfo
            Select Case iAct
                Case 1  'Kill Process
                    Debug.Print vDati(0)
                    fLista.Stuff_Exec spidDrop, "", vDati
                    fLista.Show vbModal, Me
                    bConfirm = fLista.Confirm()
                    Unload fLista
                    Me.Refresh
                    
                    If bConfirm Then
                        If TestProcInfo(CLng(vDati(0)), enSPIDDelete) Then
                            Screen.MousePointer = vbHourglass
                            On Local Error Resume Next
                            objServer.KillProcess CLng(vDati(0))
                            sErr = Err.Description
                            lErr = Err.Number
                            On Local Error GoTo 0
                            If lErr <> 0 Then
                                Screen.MousePointer = vbDefault
                                MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                            End If
                        End If
                        LoadProcInfo
                        Screen.MousePointer = vbDefault
                    End If
            End Select
            
        Case act_Server
            Select Case iAct
                Case 1  'Connect/Edit
                    bConfirm = objServer Is Nothing
                    fServer.Show vbModal, Me
                    'DoEvents
                    Me.Refresh
                    If Not objServer Is Nothing Then
                        If fServer.ChangedConn() Or bConfirm Then
                            LoadStrings
                            If objServer.VersionMajor >= kSqlServerSupported Then
                                LoadINIT
                                'LoadStrings

                                If objServer.VersionMajor <> kSqlServerSupported Then
                                    bWarning = True
                                    sCurError = ReplaceMsg(MyLoadResString(k_Not_Supported_Sql_Version), Array("1%", "2%", "3%", "|"), Array(objServer.VersionMajor, App.EXEName, kSqlServerSupported, vbCrLf))
                                End If
                            Else
                                bWarning = True
                                sCurError = ReplaceMsg(MyLoadResString(k_Obsolete_Sql_Version), Array("1%", "2%", "3%", "4%", "|"), Array(objServer.VersionMajor, App.EXEName, kSqlServerSupported, "DbaMgr", vbCrLf))
                                If Not objServer Is Nothing Then
                                    objServer.DisConnect
                                    objServer.Close
                                    Set objServer = Nothing
                                End If
                                
                            End If
                            
                            If bWarning Then
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                                               
                                fResult.tRes.Text = sCurError
                                fResult.Caption = MyLoadResString(k_GenWarning)
                                Beep
                                Screen.MousePointer = vbDefault
                                fResult.Show vbModal, Me
                            End If
                        End If
                    Else
                        LoadRoot
                    End If
                    Unload fServer
                    Screen.MousePointer = vbDefault
                    
                Case 3  'Disconnect
                    objServer.DisConnect
                    objServer.Close
                    Set objServer = Nothing
                    LoadRoot
                Case 5
                    fAdvance.Show vbModal, Me
                    Me.Refresh
                    DoEvents
                
            End Select
        Case act_DBOption, act_DBTasks
            Select Case iAct
                Case 1
                    fNewDB.Show vbModal, Me
                    DoEvents
                    Screen.MousePointer = vbHourglass
                    objServer.Databases.Refresh True
                    LoadINIT
                    'LoadDatabases
                    Screen.MousePointer = vbDefault
                Case 15     'Generate Script
                    If IsDbAvailable(sDbName) Then GenerateSQL vDati(0), "", act_DBOption
                Case 17
                    If IsDbAvailable(sDbName) Then ExportStruct2Html vDati(0)
                Case 19 'Change DB Owner
                    If IsDbAvailable(sDbName) Then
                        Screen.MousePointer = vbHourglass
                        fDbOwner.DatabaseProp vDati(0)
                        Screen.MousePointer = vbDefault
                        fDbOwner.Show vbModal, Me
                        Screen.MousePointer = vbHourglass
                        tView_NodeClick tView.SelectedItem
                        Screen.MousePointer = vbDefault
                    End If
                Case 21
                    LoadExtendedProp am_DbProp, CStr(vDati(0))
                Case 22     'Property
                    On Local Error Resume Next
                    Set oDb = objServer.Databases(CStr(vDati(0)))
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    
                    If lErr = 0 Then
                        If Not oDb.DBOption.Offline = True Then
                            fNewDB.LoadDB vDati(0)
                            fNewDB.Show vbModal, Me
                            DoEvents
                            objServer.Databases.Refresh True
                        Else
                            'fGenOpt.Caption = "Database '" & vDati(0) & "'"
                            fGenOpt.Caption = MyLoadResString(k_RES_Database) & "'" & vDati(0) & "'"
                            fGenOpt.DBName() = vDati(0)
                            fGenOpt.Action() = act_DBOption
                            fGenOpt.Show vbModal, Me
                        End If
                        Screen.MousePointer = vbHourglass
                        objServer.Databases.Refresh True
                        tView_NodeClick tView.SelectedItem
                        'LoadINIT
                        'LoadDatabases
                        Screen.MousePointer = vbDefault
                    Else
                        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                    End If
                Case Else
                    If iOpt = act_DBTasks Then
                        If iAct = 3 Then
                            iAct = 7
                        Else
                            iAct = 8
                        End If
                    End If
                    Select Case iAct
                        Case 3      'Detach
                        
                            iEnDetachMode = Val(GetIni("DEFAULT", "DETACHMODE", App.Path & "\" & App.EXEName & ".INI"))
                            
                            fLista.Stuff_Exec dbDetach, "", vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                fResult.Caption = MyLoadResString(k_Detach_Setting)
                                fResult.Action = act_Null
                                fResult.tRes.Text = ReplaceMsg(MyLoadResString(k_Detach_Warning), Array("|", "1%"), Array(vbCrLf, MyLoadResString(k_DetachMode_Check + iEnDetachMode)))
                                fResult.Show vbModal, Me
                                Unload fResult
                                Set fResult = Nothing
                                
                                Screen.MousePointer = vbHourglass
                                
                                fResult.Action() = act_Null
                                lWait = Val(GetIni("CONNECT", "SLEEP", App.Path & "\" & App.EXEName & ".INI")) * 1000
                                stbMain.Panels(1).Text = MyLoadResString(k_Detach_DB)
                                Refresh
                                
                                For iEl = 0 To UBound(vDati)
                                    On Local Error Resume Next
                                    Set oDb = objServer.Databases(CStr(vDati(iEl)))
                                    lErr = Err.Number
                                    sErr = Err.Description
                                    On Local Error GoTo 0
                                                                                                        
'iEnDetachMode = amAskEach
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    If lErr = 0 Then
                                        bIsReadOnly = oDb.DBOption.ReadOnly
                                        If oDb.SystemObject Then
                                            fResult.tRes.Text = fResult.tRes.Text & "Detach SYSTEM Database '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                        Else
                                            
                                            If oDb.DBOption.Offline = False Then
                                                fDet_MSG.Caption = ReplaceMsg(MyLoadResString(k_Detach_Of_DB), Array("1%"), Array(vDati(iEl)))
                                                'Il Database [1%] è costituito da più Files, FileGrops o LogFiles:[#]
                                                'l'operazione di RiAttacco del Database sarà possibile tramite questa Interfaccia
                                                'oppure sarà possibile eseguire la Stored Procedure [#]
                                                'sp_attach_db [@dbname =] 'dbname', [@filename1 =] 'filename_n' [,...16][#]
                                                'dalla gestione Comandi SQL con sintassi simile a[#][#]


                                                '-- Database [1%] is composed of more DataFiles, FileGroup(s) or LogFile(s):[#]
                                                'it will be possible to ReAttach this Database via this Graphical Interface or you can execute the Stored Procedure [#]
                                                'sp_attach_db [@dbname =] 'dbname', [@filename1 =] 'filename_n' [,...16][#]
                                                'from the Sql Command Line utility with a syntax similar to[#][#]

                                                sAttachSQL = ReplaceMsg("EXEC sp_attach_db @dbname = N'1%'", Array("1%"), Array(oDb.Name))
                                                iIndex = 0
                                                For iCount = 1 To oDb.FileGroups.Count
                                                    For iLoop = 1 To oDb.FileGroups(iCount).DBFiles.Count
                                                        iIndex = iIndex + 1
                                                        sAttachSQL = sAttachSQL & "," & vbCrLf & vbTab & ReplaceMsg("@filename[1%] = N'[2%]'", Array("[1%]", "[2%]"), Array(Format$(iIndex), Trim$(oDb.FileGroups(iCount).DBFiles(iLoop).PhysicalName)))
                                                    Next
                                                Next
                                                    
                                                For iCount = 1 To oDb.TransactionLog.LogFiles.Count
                                                    iIndex = iIndex + 1
                                                    sAttachSQL = sAttachSQL & "," & vbCrLf & vbTab & ReplaceMsg("@filename[1%] = N'[2%]'", Array("[1%]", "[2%]"), Array(Format$(iIndex), Trim$(oDb.TransactionLog.LogFiles(iCount).PhysicalName)))
                                                Next
                                                sAttachMSG = ReplaceMsg(MyLoadResString(k_ReAttach_Of_MultiFile_DBMSG), Array("1%", "[#]"), Array(oDb.Name, vbCrLf & "-- "))
                                                fDet_MSG.tRes.Text = sAttachMSG & vbCrLf & sAttachSQL
                                                                                                        
                                                Screen.MousePointer = vbDefault
                                                fDet_MSG.DBName() = vDati(iEl)
                                                fDet_MSG.Show vbModal, Me
                                                bConfirm = fDet_MSG.Confirm()
                                                Unload fDet_MSG
                                                Set fDet_MSG = Nothing
                                                DoEvents
                                                Screen.MousePointer = vbHourglass
                                                
                                                If bConfirm Then
                                                    If Not bIsReadOnly Then
                                                        bCheckPoint = True
    
                                                        If iEnDetachMode = amNoCheck Or oDb.DBOption.ReadOnly Then
                                                            bCheckPoint = False
                                                        ElseIf iEnDetachMode = amAskEach Then
                                                            Screen.MousePointer = vbDefault
                                                            bCheckPoint = MsgBox(Replace(MyLoadResString(k_Detach_Prompt), "1%", vDati(iEl)), vbQuestion Or vbOKCancel Or vbDefaultButton2, MyLoadResString(k_Detach_Setting)) <> vbOK
                                                            Screen.MousePointer = vbHourglass
                                                        End If
                                                    
                                                        stbMain.Panels(1).Text = Replace(MyLoadResString(k_Detach_CheckDB_stbar), "1%", vDati(iEl)) & " " & IIf(Not bCheckPoint, "-> " & MyLoadResString(k_Detach_Skipped_stbar), "")
                                                        If bCheckPoint Then
                                                            Refresh
                                                            oDb.Checkpoint
                                                        End If
                                                    End If
                                                    Set oDb = Nothing
                                                    objServer.Databases.Refresh True
                                                    If lWait <> 0 Then
                                                        stbMain.Panels(1).Text = stbMain.Panels(1).Text & " - Pause " & lWait / 1000 & """ sec."
                                                        Refresh
                                                        Sleep lWait
                                                    End If
                                                    sErr = ""
                                                    'On Local Error Resume Next
                                                    'If Not bIsReadOnly Then
                                                    '    sErr = objServer.DetachDB(vDati(iEl), True)    'bCheckPoint)
                                                    'Else
                                                        sErr = DetachWithSP_DETACH_DB(vDati(iEl))
                                                    'End If
                                                    'If Err.Number <> 0 Then sErr = sErr & IIf(Len(sErr) <> 0, vbCrLf, "") & Err.Description
                                                    'On Local Error GoTo 0
                                                    fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(k_DBerror_wait_4_retry) & vbCrLf & sErr
                                                    If Not bCheckPoint Then fResult.tRes.Text = fResult.tRes.Text & vbCrLf & "--> " & vbCrLf & Replace(MyLoadResString(k_Detach_Skipped_warning), "|", vbCrLf)
                                                Else
                                                    Set oDb = Nothing
                                                    fResult.tRes.Text = fResult.tRes.Text & ReplaceMsg(MyLoadResString(k_Skipped_Detach_Of_DB), Array("1%"), Array(vDati(iEl))) & vbCrLf & sErr
                                                End If
                                            Else
                                                fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(k_RES_Database) & " " & oDb.Name & vbLf & " OffLine"
                                            End If
                                        End If
                                    Else
                                        fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                    End If
                                    Set oDb = Nothing
                                Next
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_DB_Name_for_reattach)
                                fResult.Action = act_Null
                                fResult.WWrapVisible() = False
                                Refresh
                                fResult.Show vbModal, Me
                                Unload fResult
                                Set fResult = Nothing
                                Refresh
                                LoadINIT
                                stbMain.Panels(1).Text = "Ready"
                                Refresh
                            End If
                        Case 5      'Drop
                            fLista.Stuff_Exec dbDrop, "", vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Refresh
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    stbMain.Panels(1).Text = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_RES_Database) & " '" & vDati(iEl) & "'"
                                
                                    If objServer.Databases(CStr(vDati(iEl))).SystemObject Then
                                        fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Database '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                    Else
                                        On Local Error Resume Next
                                        objServer.Databases.Remove (CStr(vDati(iEl)))
                                    End If
                                    lErr = Err.Number
                                    sErr = Err.Description
                                    On Local Error GoTo 0
                                    If lErr <> 0 Then
                                        fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                    Else
                                        fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                    End If
                                Next
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_RES_Database)  '" Database"
                                Refresh
                                fResult.Show vbModal, Me
                                Refresh
                                LoadINIT
                                stbMain.Panels(1).Text = "Ready"
                                Refresh
                            End If
                        Case 7      'Backup
                            If IsArray(vDati) Then
                                'bConfirm = MsgBox(MyLoadResString(k_Confirm) & " " & "BackUp" & vbLf & "Database '" & vDati(0) & "'", vbQuestion Or vbOKCancel, App.EXEName) = vbOK
                                bConfirm = True
                                sRet = vDati(0)
                            Else
                                bConfirm = True
                                sRet = ""
                            End If
                            If bConfirm Then
                                fBckMain.InitDB sRet
                                fBckMain.Show ownerform:=Me
                            End If
                            
                        Case 8      'Restore
                            If IsArray(vDati) Then
                                'bConfirm = MsgBox(MyLoadResString(k_Confirm) & " " & MyLoadResString(k_Restore) & vbLf & "Database '" & vDati(0) & "'", vbQuestion Or vbOKCancel, App.EXEName) = vbOK
                                bConfirm = True
                                sRet = vDati(0)
                            Else
                                bConfirm = True
                                sRet = ""
                            End If
                            If bConfirm Then
                                fResMain.InitDB sRet
                                fResMain.Show ownerform:=Me
                            End If
                        Case 9      'Truncate Log
                            If MsgBox(MyLoadResString(k_Confirm) & " " & "Truncate Log" & vbLf & MyLoadResString(k_RES_Database) & " '" & vDati(0) & "'", vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then
                                TruncateLog vDati(0)
                            End If
                        Case 13     'CheckPoint
                            Set oDb = objServer.Databases(CStr(vDati(0)))
                            If Not oDb.DBOption.Offline Then
                                oDb.Checkpoint
                                sErr = MyLoadResString(k_Executed) & " CheckPoint:" & vbLf & "Database ['" & oDb.Name & "']"
                            Else
                                sErr = MyLoadResString(k_RES_Database) & " ['" & oDb.Name & "']" & vbLf & " OffLine"
                            End If
                            Set oDb = Nothing
                            MsgBox sErr, vbInformation Or vbOKOnly, App.EXEName
                        Case 31, 32 'Shrink,Shrink_TruncateOnly,CheckDB,CheckPoint
                        
                            Set oDb = objServer.Databases(CStr(vDati(0)))
                            If oDb.DBOption.Offline = False Then
                                On Local Error Resume Next
                                oDb.Checkpoint
                                lErr = Err.Number
                                If lErr = 0 Then
                                    '               SQLDMOShrink_Default  -- SQLDMOShrink_TruncateOnly
                                    oDb.Shrink 5, IIf(iAct = 31, SQLDMOShrink_Default, SQLDMOShrink_TruncateOnly)
                                    sAttachMSG = MyLoadResString(k_Executed) & " Pack" & vbLf & "Database ['" & oDb.Name & "']" & vbLf & MyLoadResString(k_Test_For) & " "
                                    If iAct = 31 Then
                                        sAttachMSG = sAttachMSG & MyLoadResString(k_Space_Gain_5)
                                    Else
                                        sAttachMSG = sErr & MyLoadResString(k_With_NO_Space_Gain)
                                    End If
                                End If
                                sErr = Err.Description
                                On Local Error GoTo 0
                            Else
                                sAttachMSG = MyLoadResString(k_RES_Database) & " " & oDb.Name & vbLf & " OffLine"
                            End If
                            Set oDb = Nothing
                            If lErr = 0 Then
                                MsgBox sAttachMSG, vbInformation Or vbOKOnly, App.EXEName
                            Else
                                MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
                            End If
                        
                        Case 41, 42, 43, 44
                            Set oDb = objServer.Databases(CStr(vDati(0)))
                            Screen.MousePointer = vbHourglass
                            On Local Error Resume Next
                            sErr = oDb.CheckTables(iAct - 41)
                            sCurError = Err.Description
                            lErr = Err.Number
                            If lErr <> 0 Then sErr = MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sCurError & vbLf & sErr
                            On Local Error GoTo 0
                            If Len(sErr) = 0 Then sErr = MyLoadResString(k_DBCC_Comleted)
                            
                                                            
                            fResult.tRes.Text = sErr
                            fResult.Action = act_Null
                            fResult.WWrapVisible() = False
                            fResult.Caption = MyLoadResString(k_Executed) & " CheckDB: Database '" & vDati(0) & "' " & sRet
                            Screen.MousePointer = vbDefault
                            fResult.Show vbModal, Me
                    End Select
            End Select
        Case act_Roles
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 3  'drop
                        If IsDbAvailableReadWrite(sDbName) Then
                            fLista.Stuff_Exec roleDrop, sDbName, vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Set oDb = objServer.Databases(sDbName)
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(k_Role) & " '" & vDati(iEl) & "'"
                                                                                            
                                    If oDb.DatabaseRoles(CStr(vDati(iEl))).IsFixedRole Then
                                        fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Role '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                    Else
                                        On Local Error Resume Next
                                        oDb.DatabaseRoles.Remove vDati(iEl)
                                        If Err.Number <> 0 Then
                                            lErr = Err.Number
                                            sErr = Err.Description
                                        End If
                                        On Local Error GoTo 0
                                        If lErr <> 0 Then
                                            fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                        Else
                                            fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                        End If
                                    End If
                                Next
                                
                                bRefresh = True
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_Role) '" Role"
                                Refresh
                                fResult.Show vbModal, Me
                                
                            End If
                        End If
                    Case 1, 5 'new/property
                        If iAct = 1 Then
                            If Not IsDbAvailableReadWrite(sDbName) Then Exit Sub
                        End If
                        fRoles.LoadRole sDbName, IIf(iAct = 1, "", vDati(0))
                        lErr = fRoles.ErrStat(sErr)
                        If lErr = 0 Then
                            fRoles.Show vbModal, Me
                        Else
                            Unload fRoles
                            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                        End If
                        bRefresh = True
                End Select
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    objServer.Databases(sDbName).DatabaseRoles.Refresh True
                    LoadRoles sDbName
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case act_SP
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 1  'New
                        If IsDbAvailableReadWrite(sDbName) Then
                            fPwd.Action() = act_Tbl
                            fPwd.Caption = MyLoadResString(k_New_Stored_Procedure)
                            fPwd.Stuff() = "ProcName"
                            fPwd.Show vbModal, Me
                            bConfirm = fPwd.Confirm()
                            ReDim vDati(0) As Variant
                            If bConfirm Then vDati(0) = fPwd.Stuff()
                            Unload fPwd
                            If bConfirm Then
                                fResult.Caption = MyLoadResString(k_RES_Database) & " '" & sDbName & "'" & "- " & MyLoadResString(k_New_Stored_Procedure) & " '" & vDati(0) & "'"
                                fResult.Stuff_SP sDbName, vDati(0) ', vDati(1)
        
                                sOwner = GetCreateOwner(sDbName)
                                fResult.tRes.Text = ReplaceMsg("CREATE PROCEDURE 1%.[2%] AS", Array("1%", "2%"), Array(sOwner, vDati(0)))
                                fResult.Action() = act_SP
                                fResult.Show vbModal, Me
                            End If
                            bRefresh = True
                        End If
                    Case 3  'Delete
                        If IsDbAvailableReadWrite(sDbName) Then
                            fLista.Stuff_Exec spDrop, sDbName, vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Set oDb = objServer.Databases(sDbName)
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(k_Stored_Procedure) & " '" & vDati(iEl) & "'"
                                    
                                                            
                                    Set objSp = GetSpByName(objServer, sDbName, CStr(vDati(iEl)))
                                    
                                    'If oDB.StoredProcedures(CStr(vDati(iEl))).SystemObject = 1 Then
                                    If objSp.SystemObject = 1 Then
                                        fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Object '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                    Else
                                        On Local Error Resume Next
                                        'oDB.StoredProcedures.Remove vDati(iEl)
                                        oDb.StoredProcedures.Remove objSp.Name
                                        If Err.Number <> 0 Then
                                            lErr = Err.Number
                                            sErr = Err.Description
                                        End If
                                        On Local Error GoTo 0
                                        Set objSp = Nothing
                                        
                                        If lErr <> 0 Then
                                            fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                        Else
                                            fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                        End If
                                    End If
                                Next
                                
                                bRefresh = True
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_Stored_Procedure)
                                Refresh
                                fResult.Show vbModal, Me
                            End If
                        End If
                    Case 5
                        GenerateSQL sDbName, vDati(0), act_SP
                    Case 6
                        LoadPermission sDbName, vDati(0), act_SP
                    Case 8
                        Set objDependent = New clsDepen
                        objDependent.ShowDependent sDbName, vDati(0), spDrop
                        Set objDependent = Nothing
                    Case 9
                        If IsDbAvailableReadWrite(sDbName) Then ChangeObjectOwner sDbName, vDati, amObjStoredProc
                        
                    Case 11 'Property
                        fGenOpt.Caption = MyLoadResString(k_Stored_Procedure) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_SP
                        fGenOpt.Show vbModal, Me
                    Case 12 'Extended Property
                        LoadExtendedProp am_SpProp, sDbName, GetObjectOwner(sDbName, CStr(vDati(0)), am_OwnSP), CStr(vDati(0))
                    Case 13  'Text
                        Screen.MousePointer = vbHourglass
                        'Set oDB = objServer.Databases(sDbName)
                        Set objSp = GetSpByName(objServer, sDbName, CStr(vDati(0)))
            
                        fResult.Caption = MyLoadResString(k_RES_Database) & " '" & sDbName & "'" & "- " & MyLoadResString(k_Stored_Procedure) & "'" & vDati(0) & "'"
                        'fResult.tRes.Text = oDB.StoredProcedures(CStr(vDati(0))).Text
                        fResult.tRes.Text = objSp.Text
                        Set objSp = Nothing
                        
                        fResult.Action() = act_Null
                        Screen.MousePointer = vbDefault
                        fResult.Show vbModal, Me
                End Select
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    objServer.Databases(sDbName).StoredProcedures.Refresh True
                    LoadStoredProc sDbName
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case act_Tbl
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 1
                        If IsDbAvailableReadWrite(sDbName) Then
                            fTBDes.DBName() = sDbName
                            fTBDes.NewTbl() = True
                            fTBDes.LoadTbl
                            fTBDes.Show vbModal, Me
                            bRefresh = True
                        End If
                    Case 3
                        fTBDes.DBName() = sDbName
                        fTBDes.Stuff() = "[" & vDati(0) & "]"
                        fTBDes.LoadTbl
                        fTBDes.Show vbModal, Me
                        bRefresh = True
                    
                        
                    Case 7  'DropTable
                        If IsDbAvailableReadWrite(sDbName) Then
                            fLista.Stuff_Exec tableDrop, sDbName, vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Set oDb = objServer.Databases(sDbName)
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    fResult.tRes.Text = fResult.tRes.Text & " " & MyLoadResString(k_Table) & " '" & vDati(iEl) & "'"
                                    
                                    Set oTb = GetTbByName(objServer, sDbName, vDati(iEl))
                                    'If oDB.Tables(CStr(vDati(iEl))).SystemObject = 1 Then
                                    If oTb.SystemObject = 1 Then
                                        fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Object '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                        Set oTb = Nothing
                                    Else
                                        On Local Error Resume Next
                                        'oDB.Tables.Remove vDati(iEl)
                                        oDb.Tables.Remove oTb.Name, oTb.Owner
                                        If Err.Number <> 0 Then
                                            lErr = Err.Number
                                            sErr = Err.Description
                                        End If
                                        On Local Error GoTo 0
                                        Set oTb = Nothing
                                        If lErr <> 0 Then
                                            fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                        Else
                                            fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                        End If
                                    End If
                                Next
                                
                                bRefresh = True
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_Table)
                                Refresh
                                fResult.Show vbModal, Me
                            End If
                        End If
                    Case 9
                        GenerateSQL sDbName, vDati(0), act_Tbl
                    Case 10
                        LoadPermission sDbName, vDati(0), act_Tbl
                        
                    Case 12
                        Set objDependent = New clsDepen
                        objDependent.ShowDependent sDbName, "[" & vDati(0) & "]", tableDrop
                        Set objDependent = Nothing
                    
                    Case 14     'trigger
                        fResult.Caption = MyLoadResString(k_Trigger_for_Table) & " '" & vDati(0) & "'"
                        fResult.Action = act_Tri
                        fResult.Stuff_Trigger sDbName, vDati(0)
                        fResult.Show vbModal, Me

                    Case 15
                        fIndex.DataNames sDbName, vDati(0)
                        fIndex.Caption = fIndex.Caption & MyLoadResString(k_for_Table) & "  '" & vDati(0) & "' - Database '" & sDbName & "'"
                        fIndex.Show vbModal, Me
                    Case 16     'Extended Properties Idx+Relations
                        fGenOpt.Caption = MyLoadResString(k_Table) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_Extended_TABLE_Idx_Rel
                        fGenOpt.Show vbModal, Me

                    Case 17     'Extended Properties Triggers
                        fGenOpt.Caption = MyLoadResString(k_Table) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_Extended_TABLE_Trig
                        fGenOpt.Show vbModal, Me

                    Case 19     'Relation Designer
                        lErr = fRelView.ErrSetDBName(sDbName, sErr)
                        If lErr = 0 Then
                            fRelView.Show vbModal, Me
                        Else
                            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                        End If
                        Set fRelView = Nothing
                    Case 21     'Change Object Owner
                        If IsDbAvailableReadWrite(sDbName) Then ChangeObjectOwner sDbName, vDati, amObjTable
                    Case 22     'generate insert
                        fExpTB.Caption = ReplaceMsg(MyLoadResString(k_sqlFrmGenerateInsert), Array("1%", "2%"), Array(vDati(0), sDbName))
                        fExpTB.ObjectsNames(sDbName) = vDati(0)
                        
                        fExpTB.Show vbModal, Me
                    
                    Case 24
                        fGenOpt.Caption = MyLoadResString(k_Table) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_Tbl
                        fGenOpt.Show vbModal, Me
                    Case 25 'Extended Property
                        LoadExtendedProp am_TbProp, sDbName, GetObjectOwner(sDbName, CStr(vDati(0)), am_OwnTable), CStr(vDati(0))
                    
                    Case Else
                        If iAct = 5 Then
                        
                            Screen.MousePointer = vbDefault
                            With fRequest
                                .Mode() = am_Long
                                .Messages("") = MyLoadResString(k_Num_of_Rows_to_return)
                                .ShowValue(100) = 5
                            End With
                            fRequest.Show vbModal, Me
                            bConfirm = Not fRequest.Cancelled()
                            If bConfirm Then
                                bConfirm = CLng(fRequest.ReturnValue()) > 0
                                If bConfirm Then sRet = " TOP " & CLng(fRequest.ReturnValue()) & " "
                            End If
                            Unload fRequest
                        Else
                            Screen.MousePointer = vbDefault
                            bConfirm = MsgBox(MyLoadResString(k_Confirm) & " " & MyLoadResString(k_Show_all_Rows) & vbLf & "Table: ['" & vDati(0) & "']", vbQuestion Or vbOKCancel Or vbDefaultButton2, App.EXEName) = vbOK
                        End If
                        Screen.MousePointer = vbDefault
                        
                        If bConfirm Then
                            Screen.MousePointer = vbHourglass
                            stbMain.Panels(1).Text = "Fetching..."
                            Refresh
                            sRet = "SELECT" & sRet & " * FROM " & GetTableOwner(objServer, sDbName, vDati(0)) & ".[" & vDati(0) & "]"
                            TableSet.DBName() = sDbName
                            TableSet.TableName() = vDati(0)
                            TableSet.SqlStatement = sRet
                            TableSet.Execute
                            If TableSet.ErrReturned() = 0 Then
                                TableSet.Show vbModal, Me
                            Else
                                Unload TableSet
    '                            If MsgBox(MyLoadResString(k_Showing_Results_via_DMO ), vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then
    '                                fRetSet.DBName() = sDbName
    '                                fRetSet.SQLstr() = sRet
    '                                fRetSet.Execute
    '                                fRetSet.Show vbModal, Me
    '                            End If
                            End If
                            
                        End If
                End Select
                
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    On Local Error Resume Next
                    objServer.Databases(sDbName).Tables.Refresh True
                    lErr = Err.Number
                    If lErr <> 0 Then
                        Err.Clear
                        objServer.ReConnect
                        lErr = Err.Number
                        If lErr = 0 Then objServer.Databases(sDbName).Tables.Refresh True
                        lErr = Err.Number
                        sErr = Err.Description
                    End If
                    On Local Error GoTo 0
                    If lErr = 0 Then
                        LoadTables sDbName
                    Else
                        MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
                        LoadRoot
                    End If
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case act_UDT
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 1, 8 'new/property
                        If iAct = 1 Then
                            If Not IsDbAvailableReadWrite(sDbName) Then Exit Sub
                        End If
                        bRefresh = True
                        fUDT.DBName = sDbName
                        fUDT.loadUDT IIf(iAct = 1, "", CStr(vDati(0)))
                        
                        fUDT.Show vbModal, Me
                    Case 3  'drop
                        If IsDbAvailableReadWrite(sDbName) Then
                            fLista.Stuff_Exec udtDrop, sDbName, vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Set oDb = objServer.Databases(sDbName)
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    fResult.tRes.Text = fResult.tRes.Text & " " & MyLoadResString(k_User_Defined_Data_Type) & " '" & vDati(iEl) & "'"
                                                            
                                    On Local Error Resume Next
                                    oDb.UserDefinedDatatypes.Remove vDati(iEl)
                                    If Err.Number <> 0 Then
                                        lErr = Err.Number
                                        sErr = Err.Description
                                    End If
                                    On Local Error GoTo 0
                                    If lErr <> 0 Then
                                        fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                    Else
                                        fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                    End If
                                Next
                                
                                bRefresh = True
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_User_Defined_Data_Type)
                                Refresh
                                fResult.Show vbModal, Me
                            End If
                        End If
                    Case 5  'script
                        GenerateSQL sDbName, vDati(0), act_UDT
                    Case 7
                        Set objDependent = New clsDepen
                        objDependent.ShowDependent sDbName, vDati(0), udtDrop
                        Set objDependent = Nothing
                    Case 9 'Extended Property
                        LoadExtendedProp am_UdtProp, sDbName, "", CStr(vDati(0))
                        
                End Select
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    objServer.Databases(sDbName).UserDefinedDatatypes.Refresh True
                    loadUDT sDbName
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case act_User
            If IsDbAvailable(sDbName) Then
                If iAct = 1 Or iAct = 10 Then
                    If iAct = 1 Then
                        If Not IsDbAvailableReadWrite(sDbName) Then Exit Sub
                    End If
                    fGenOpt.Caption = MyLoadResString(k_User) & " '" & IIf(iAct = 10, vDati(0), MyLoadResString(k_New_User)) & "'"
                    fGenOpt.DBName() = sDbName
                    If iAct = 10 Then fGenOpt.Stuff() = vDati(0)
                    fGenOpt.Action() = act_User
                    fGenOpt.Show vbModal, Me
                ElseIf iAct = 5 Then
                    LoadPermission sDbName, vDati(0), act_User
                ElseIf iAct = 7 Then
                    LoadUsersProp sDbName
                ElseIf iAct = 8 Then   'Objects Owned by ..
                    fLista.Stuff_Exec EnumUserObj, sDbName, vDati
                    fLista.Show vbModal, Me
                    Unload fLista
                    If m_bNeedReattach Then
                        Beep
                        MsgBox Replace(MyLoadResString(k_SqlObjInvalid), "|", vbCrLf), vbInformation Or vbOKOnly, App.EXEName
                        Set objServer = Nothing
                        LoadRoot
                        
                        Flex.Redraw = False
                        Flex.Clear
                        Flex.Rows = 1
                        Flex.Cols = 1
                        Flex.ColWidth(0) = Flex.Width
                        Flex.Redraw = True
    
                        Exit Sub
                    End If
                Else
                    If IsDbAvailableReadWrite(sDbName) Then
                        fLista.Stuff_Exec userDrop, sDbName, vDati
                        fLista.Show vbModal, Me
                        bConfirm = fLista.Confirm()
                        Unload fLista
                        
                        If bConfirm Then
                            Screen.MousePointer = vbHourglass
                            fResult.Action() = act_Null
                            fResult.WWrapVisible() = False
                            stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                            Set oDb = objServer.Databases(sDbName)
                            
                            For iEl = 0 To UBound(vDati)
                                fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                fResult.tRes.Text = fResult.tRes.Text & " " & MyLoadResString(k_User) & " '" & vDati(iEl) & "'"
                                                        
                                On Local Error Resume Next
                                oDb.Users.Remove vDati(iEl)
                                If Err.Number <> 0 Then
                                    lErr = Err.Number
                                    sErr = Err.Description
                                End If
                                On Local Error GoTo 0
                                If lErr <> 0 Then
                                    fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                Else
                                    fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                End If
                            Next
                            
                            bRefresh = True
                            Screen.MousePointer = vbDefault
                            fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_Users)
                            Refresh
                            fResult.Show vbModal, Me
                        End If
                    End If
                End If
                Screen.MousePointer = vbHourglass
                objServer.Databases(sDbName).Users.Refresh True
                LoadDatabaseUsers sDbName
                stbMain.Panels(1).Text = "Ready"
                Screen.MousePointer = vbDefault
            End If
        Case act_Login
            If iAct = 1 Or iAct = 5 Then
                Screen.MousePointer = vbHourglass
                stbMain.Panels(1).Text = "Scanning..."
                DoEvents
                If IsArray(vDati) Then sRet = vDati(0)
                frmLogin.LogName() = IIf(iAct = 1, "", sRet)
                
                frmLogin.Load_Data
                stbMain.Panels(1).Text = "Ready"
                DoEvents
                frmLogin.Show vbModal, Me
            Else
                fLista.Stuff_Exec loginDrop, "", vDati
                fLista.Show vbModal, Me
                bConfirm = fLista.Confirm()
                Unload fLista
                
                If bConfirm Then
                    Screen.MousePointer = vbHourglass
                    fResult.Action() = act_Null
                    fResult.WWrapVisible() = False
                    stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                        
                    For iEl = 0 To UBound(vDati)
                        fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                        fResult.tRes.Text = fResult.tRes.Text & " " & MyLoadResString(k_Login) & " '" & vDati(iEl) & "'"
                
                        On Local Error Resume Next
                        objServer.Logins.Remove vDati(iEl)
                        If Err.Number <> 0 Then
                            lErr = Err.Number
                            sErr = Err.Description
                        End If
                        On Local Error GoTo 0
                        If lErr <> 0 Then
                            fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                        Else
                            fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                        End If
                    Next
                    
                    bRefresh = True
                    Screen.MousePointer = vbDefault
                    fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_Users)
                    Refresh
                    fResult.Show vbModal, Me
                End If
            End If
            Screen.MousePointer = vbHourglass
            objServer.Logins.Refresh True
            LoadLogin
            stbMain.Panels(1).Text = "Ready"
            Screen.MousePointer = vbDefault
            
        Case act_View
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 1
                        fPwd.Action() = act_View
                        fPwd.Caption = MyLoadResString(k_New_View)
                        fPwd.Stuff() = MyLoadResString(k_View)
                        fPwd.Show vbModal, Me
                        bConfirm = fPwd.Confirm()
                        ReDim vDati(0) As Variant
                        If bConfirm Then vDati(0) = fPwd.Stuff()
                        Unload fPwd
                        If bConfirm Then
                            'fQB.ViewStuff() = oViewObject
                            
                            fQB.DBStuff(sDbName) = vDati(0)
                            fQB.Show vbModal, Me
                            'fQB.Show vbModeless, Me
                            bRefresh = True
        
                        End If
                    Case 3  'Drop
                        fLista.Stuff_Exec viewDrop, sDbName, vDati
                        fLista.Show vbModal, Me
                        bConfirm = fLista.Confirm()
                        Unload fLista
                        
                        If bConfirm Then
                            Screen.MousePointer = vbHourglass
                            fResult.Action() = act_Null
                            fResult.WWrapVisible() = False
                            stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                            Set oDb = objServer.Databases(sDbName)
                            
                            For iEl = 0 To UBound(vDati)
                                fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                fResult.tRes.Text = fResult.tRes.Text & " " & MyLoadResString(k_View) & " '" & vDati(iEl) & "'"
                                Set oView = GetViewByName(objServer, sDbName, vDati(iEl))
                                
                                If oView.SystemObject = 1 Then
                                    fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Object '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                    Set oView = Nothing
                                Else
                                    On Local Error Resume Next
                                    oDb.Views.Remove oView.Name, oView.Owner
                                    If Err.Number <> 0 Then
                                        lErr = Err.Number
                                        sErr = Err.Description
                                    End If
                                    On Local Error GoTo 0
                                    Set oView = Nothing
                                    If lErr <> 0 Then
                                        fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                    Else
                                        fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                    End If
                                End If
                            Next
                            
                            bRefresh = True
                            Screen.MousePointer = vbDefault
                            fResult.Caption = MyLoadResString(k_Views) & " " & MyLoadResString(k_View)
                            Refresh
                            fResult.Show vbModal, Me
                        End If
                    
                    Case 5, 6
                        If iAct = 6 Then
                        
                            Screen.MousePointer = vbDefault
                            With fRequest
                                .Mode() = am_Long
                                .Messages("") = MyLoadResString(k_Num_of_Rows_to_return)
                                .ShowValue(100) = 5
                            End With
                            fRequest.Show vbModal, Me
                            bConfirm = Not fRequest.Cancelled()
                            If bConfirm Then
                                bConfirm = CLng(fRequest.ReturnValue()) > 0
                                If bConfirm Then sRet = " TOP " & CLng(fRequest.ReturnValue()) & " "
                            End If
                            Unload fRequest
                        Else
                            Screen.MousePointer = vbDefault
                            bConfirm = MsgBox(MyLoadResString(k_Confirm) & " " & MyLoadResString(k_Show_all_Rows) & vbLf & "View: ['" & vDati(0) & "']", vbQuestion Or vbOKCancel Or vbDefaultButton2, App.EXEName) = vbOK
                        End If
                        Screen.MousePointer = vbDefault
                        
                        If bConfirm Then
                            Screen.MousePointer = vbHourglass
                            stbMain.Panels(1).Text = "Fetching..."
                            Refresh
                            sRet = "SELECT" & sRet & " * FROM [" & GetViewOwner(objServer, sDbName, vDati(0)) & "].[" & vDati(0) & "]"
                            TableSet.DBName() = sDbName
                            TableSet.TableName() = vDati(0)
                            TableSet.SqlStatement = sRet
                            TableSet.Execute
                            If TableSet.ErrReturned() = 0 Then
                                TableSet.Show vbModal, Me
                            Else
                                Unload TableSet
                                If MsgBox(MyLoadResString(k_Showing_Results_via_DMO), vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then
                                    fRetSet.DBName() = sDbName
                                    fRetSet.SQLstr() = sRet
                                    fRetSet.Execute
                                    fRetSet.Show vbModal, Me
                                End If
                            End If
                        End If
    
                    Case 8
                        GenerateSQL sDbName, vDati(0), act_View
                    Case 9
                        LoadPermission sDbName, vDati(0), act_View
                    Case 11     'trigger
                        fResult.Caption = MyLoadResString(k_Trigger_for_View) & " '" & vDati(0) & "'"
                        fResult.IsTrig4View() = True
                        fResult.Action = act_Tri
                        fResult.Stuff_Trigger sDbName, vDati(0)
                        fResult.Show vbModal, Me
                    Case 12     'Index
                        fIdxView.DataNames sDbName, vDati(0)
                        fIdxView.Caption = fIdxView.Caption & " " & MyLoadResString(k_for_View) & "  '" & vDati(0) & "' - Database '" & sDbName & "'"
                        fIdxView.Show vbModal, Me
                    
                    Case 13     'Extended Prop for Idx And Triggers
                        fGenOpt.Caption = MyLoadResString(k_View) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_Extended_VIEW_Idx_Trig
                        fGenOpt.Show vbModal, Me

                    
                    Case 15
                        Set objDependent = New clsDepen
                        objDependent.ShowDependent sDbName, vDati(0), viewDrop
                        Set objDependent = Nothing
                    Case 16
                        ChangeObjectOwner sDbName, vDati, amObjView
                    
                    Case 18
                        fGenOpt.Caption = MyLoadResString(k_View) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_View
                        fGenOpt.Show vbModal, Me
                    
                    Case 19 'Extended Property
                        LoadExtendedProp am_ViewProp, sDbName, GetObjectOwner(sDbName, CStr(vDati(0)), am_OwnView), CStr(vDati(0))
                    Case 20 'Text
                        Screen.MousePointer = vbHourglass
                        Set oDb = objServer.Databases(sDbName)
                        
                        Set oView = GetViewByName(objServer, sDbName, CStr(vDati(0)))
            
                        fResult.Caption = MyLoadResString(k_RES_Database) & " '" & sDbName & "'" & "- " & MyLoadResString(k_View) & " '" & vDati(0) & "'"
                        fResult.tRes.Text = oView.Text
                        Set oView = Nothing
                        fResult.Action() = act_Null
                        Screen.MousePointer = vbDefault
                        fResult.Show vbModal, Me
                End Select
                
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    objServer.Databases(sDbName).Views.Refresh True
                    LoadView sDbName
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case act_Bck
            If iAct = 9 Then
                With fRequest
                    .Mode() = am_Date
                    .Messages(MyLoadResString(k_Clear_BackUp_History) & MyLoadResString(k_Till_Date)) = MyLoadResString(k_Clear_BackUp_History)
                    .ShowValue(Format$(Now - 30, kYEAR)) = 10
                End With
                fRequest.Show vbModal, Me
                bConfirm = Not fRequest.Cancelled()
                If bConfirm Then sRet = InvertData(fRequest.ReturnValue())
                Unload fRequest
                DoEvents
                If Len(sRet) Then
                    Screen.MousePointer = vbHourglass
                    sRet = "USE msdb EXEC sp_delete_backuphistory @oldest_date = '" & sRet & "'"
                    On Local Error Resume Next
                    Call objServer.ExecuteWithResultsAndMessages(sRet, Messages:=sErr)
                    lErr = Err.Number
                    If lErr <> 0 Then sErr = sErr & Err.Description
                    On Local Error GoTo 0
                    
                    If lErr <> 0 Then sErr = MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                    Screen.MousePointer = vbDefault
                    MsgBox sErr, vbInformation Or vbOKOnly, App.EXEName

                End If
                Exit Sub
            
            
            ElseIf iAct = 1 Or iAct = 5 Then
                Screen.MousePointer = vbHourglass
                If iAct = 5 Then fBckD.LoadDevice vDati(0)
                Screen.MousePointer = vbDefault
                fBckD.Show vbModal, Me
            Else
            
                If MsgBox(MyLoadResString(k_Confirm_Remove) & " " & MyLoadResString(k_Device) & vbLf & vDati(0) & vbLf & "(" & MyLoadResString(k_Not_restorable_Operation) & ")", vbCritical Or vbOKCancel, App.EXEName) = vbOK Then
                
                    Screen.MousePointer = vbHourglass
                    
                    On Local Error Resume Next
                    Call objServer.ExecuteWithResultsAndMessages("sp_dropdevice @logicalname ='" & vDati(0) & "', @delfile = 'delfile'", Messages:=sErr)
                                        
                    lErr = Err.Number
                    If lErr <> 0 Then sErr = sErr & Err.Description
                    On Local Error GoTo 0
                    
                    If lErr <> 0 Then
                        Screen.MousePointer = vbDefault
                        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                    End If
                    Screen.MousePointer = vbDefault
                End If
            End If
            Screen.MousePointer = vbHourglass
            objServer.BackupDevices.Refresh True
            LoadBackup
            stbMain.Panels(1).Text = "Ready"
            Screen.MousePointer = vbDefault
        Case act_User_Function
            If IsDbAvailable(sDbName) Then
                Select Case iAct
                    Case 1  'New
                        If IsDbAvailableReadWrite(sDbName) Then
                            fPwd.Action() = act_Tbl
                            fPwd.Caption = MyLoadResString(k_Func_NewFunction)
                            fPwd.Stuff() = "FunctionName"
                            fPwd.Show vbModal, Me
                            bConfirm = fPwd.Confirm()
                            ReDim vDati(0) As Variant
                            If bConfirm Then vDati(0) = fPwd.Stuff()
                            Unload fPwd
                            If bConfirm Then
                                fResult.Caption = MyLoadResString(k_RES_Database) & " '" & sDbName & "'" & "- " & MyLoadResString(k_New_Stored_Procedure) & " '" & vDati(0) & "'"
                                fResult.Stuff_SP sDbName, vDati(0) ', vDati(1)
        
                                sOwner = GetCreateOwner(sDbName)
                                '"CREATE FUNCTION [1%].[2%] (PARAMETER LIST)|RETURNS (return_type_spec) AS|BEGIN|(FUNCTION BODY)|End
                                fResult.tRes.Text = ReplaceMsg("CREATE FUNCTION [1%].[2%] (PARAMETER LIST)|RETURNS (return_type_spec) AS|BEGIN|(FUNCTION BODY)|End", Array("1%", "2%", "|"), Array(sOwner, vDati(0), vbCrLf))
                                fResult.Action() = act_User_Function
                                fResult.Show vbModal, Me
                            End If
                            bRefresh = True
                        End If
                    Case 3  'Delete
                        If IsDbAvailableReadWrite(sDbName) Then
                            fLista.Stuff_Exec funcDrop, sDbName, vDati
                            fLista.Show vbModal, Me
                            bConfirm = fLista.Confirm()
                            Unload fLista
                            
                            If bConfirm Then
                                Screen.MousePointer = vbHourglass
                                fResult.Action() = act_Null
                                fResult.WWrapVisible() = False
                                stbMain.Panels(1).Text = MyLoadResString(k_Dropping_Objects)
                                Set oDb = objServer.Databases(sDbName)
                                
                                For iEl = 0 To UBound(vDati)
                                    fResult.tRes.Text = fResult.tRes.Text & IIf(Len(fResult.tRes.Text), vbCrLf, "") & String$(20, "=") & vbCrLf
                                    fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(k_objFunction) & " '" & vDati(iEl) & "'"
                                    
                                    Set objFunc = GetFuncByName(objServer, sDbName, CStr(vDati(iEl)))
                                    
                                    If objFunc.SystemObject = 1 Then
                                        fResult.tRes.Text = fResult.tRes.Text & "Drop SYSTEM Object '" & vDati(iEl) & "' " & MyLoadResString(k_System_Object_not_removable)
                                    Else
                                        On Local Error Resume Next
                                        oDb.UserDefinedFunctions.Remove objFunc.Name
                                        If Err.Number <> 0 Then
                                            lErr = Err.Number
                                            sErr = Err.Description
                                        End If
                                        On Local Error GoTo 0
                                        Set objFunc = Nothing
                                        
                                        If lErr <> 0 Then
                                            fResult.tRes.Text = fResult.tRes.Text & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr
                                        Else
                                            fResult.tRes.Text = fResult.tRes.Text & vbTab & "ok."
                                        End If
                                    End If
                                Next
                                
                                bRefresh = True
                                Screen.MousePointer = vbDefault
                                fResult.Caption = MyLoadResString(k_Dropping_Objects) & " " & MyLoadResString(k_objFunction)
                                Refresh
                                fResult.Show vbModal, Me
                            End If
                        End If
                    Case 5
                        GenerateSQL sDbName, vDati(0), act_User_Function
                    Case 6
                        LoadPermission sDbName, vDati(0), act_User_Function
                    Case 8
                        Set objDependent = New clsDepen
                        objDependent.ShowDependent sDbName, vDati(0), funcDrop
                        Set objDependent = Nothing
                    Case 9
                        If IsDbAvailableReadWrite(sDbName) Then ChangeObjectOwner sDbName, vDati, amObjFunction
                    Case 11  'Property
                        fGenOpt.Caption = MyLoadResString(k_objFunction) & " '" & vDati(0) & "' Database: '" & sDbName & "'"
                        fGenOpt.DBName() = sDbName
                        fGenOpt.Stuff() = vDati(0)
                        fGenOpt.Action() = act_User_Function
                        fGenOpt.Show vbModal, Me
                    
                    Case 12 'Extended Property
                        LoadExtendedProp am_FuncProp, sDbName, GetObjectOwner(sDbName, CStr(vDati(0)), am_OwnFunction), CStr(vDati(0))
                    
                    Case 13  'Text
                        Screen.MousePointer = vbHourglass
                        'Set oDB = objServer.Databases(sDbName)
                        Set objFunc = GetFuncByName(objServer, sDbName, CStr(vDati(0)))
            
                        fResult.Caption = MyLoadResString(k_RES_Database) & " '" & sDbName & "'" & "- " & MyLoadResString(k_objFunction) & "'" & vDati(0) & "'"
                        
                        fResult.tRes.Text = objFunc.Text
                        Set objFunc = Nothing
                        
                        fResult.Action() = act_Null
                        Screen.MousePointer = vbDefault
                        fResult.Show vbModal, Me
                End Select
                If bRefresh Then
                    Screen.MousePointer = vbHourglass
                    Set oDb = objServer.Databases(sDbName)
                    oDb.UserDefinedFunctions.Refresh True
                    LoadFunctions sDbName
                    stbMain.Panels(1).Text = "Ready"
                    Screen.MousePointer = vbDefault
                End If
            End If
    End Select
    Set oTb = Nothing
    Set oView = Nothing
    Set objSp = Nothing
    Set objFunc = Nothing
    
    Set oDb = Nothing
    Set objDependent = Nothing
    
End Sub

Private Sub SetLowHigh(ByRef intLow As Integer, ByRef intHigh As Integer)

    Dim intX As Integer
    
    intLow = Flex.Row
    intX = Flex.RowSel
    If intLow > intX Then
        intHigh = intLow
        intLow = intX
    Else
        intHigh = intX
    End If

End Sub
Public Sub ReInit()
    LoadINIT
End Sub
Private Sub LoadActivity()


    Set mNode = tView.Nodes.Add("Server", tvwChild, "Management")
    mNode.Text = MyLoadResString(k_RES_Management)
    mNode.Image = "Folder"   ' Set Image property.
    
    
    Set mNode = tView.Nodes.Add("Management", tvwChild, "Activity")
    mNode.Text = MyLoadResString(k_RES_CurrentActivity) & " - " & Now
    mNode.Image = "Activity"   ' Set Image property.
    
    Set mNode = tView.Nodes.Add("Activity", tvwChild, "ProcInfo")
    mNode.Text = MyLoadResString(k_RES_Process_Info)
    mNode.Image = "ProcInfo"   ' Set Image property.

    Set mNode = tView.Nodes.Add("Activity", tvwChild, "locksID")
    mNode.Text = MyLoadResString(kTviewLockID) '"Locks /ID" 'MyLoadResString(k_RES_Process_Info )
    mNode.Image = "LockID"   ' Set Image property.
    LoadLocks

    LoadLogsNodes
End Sub
Private Sub LoadProcInfo()

    Const kSpid = 1
    Const kUser = 2
    Const kDBName = 3
    Const kStatus = 4
    Const kOpenTran = 5
    Const kCmd = 6
    Const kPrgName = 7
    Const kWaitTime = 8
    Const kWaitType = 9
    Const kCPU = 11
    Const kIO = 12
    Const kMemUsage = 13
    Const kLoginTime = 14
    Const kLastBatch = 15
    Const kHostName = 16
    Const kNetLib = 17
    Const kNetAddress = 18
    Const kBlocked = 19
    Const kBlocking = 20
    
    Dim iSet As Integer
    Dim iCol As Integer ', icLoop As Integer
    Dim iRow As Integer, irLoop As Integer
    Dim oQry As QueryResults
    Dim iIco As Integer
        
    Dim sBuf As String
    
    Screen.MousePointer = vbHourglass
    Set oQry = GetSpidQry()

'    'Set oQry = objServer.EnumProcesses
'    sBuf = "EXEC sp_MSget_current_activity @id=7,@option=1"
'    On Local Error Resume Next
'    Set oQry = objServer.ExecuteWithResults(sBuf, Len(sBuf))
'    On Local Error GoTo 0

    On Local Error Resume Next

    With Flex
        .Clear
        .Redraw = False
        .Cols = 19
        .Rows = 1

        .TextMatrix(0, 0) = MyLoadResString(k_RES_Spid)
        .TextMatrix(0, 1) = MyLoadResString(k_RES_User)
        .TextMatrix(0, 2) = MyLoadResString(k_RES_Database)
        .TextMatrix(0, 3) = MyLoadResString(k_RES_Status)
        .TextMatrix(0, 4) = MyLoadResString(k_RES_openTran)
        .TextMatrix(0, 5) = MyLoadResString(k_RES_Command)
        .TextMatrix(0, 6) = MyLoadResString(k_RES_Application)
        .TextMatrix(0, 7) = MyLoadResString(k_RES_WaitTime)
        .TextMatrix(0, 8) = MyLoadResString(k_RES_WaitType)
        .TextMatrix(0, 9) = MyLoadResString(k_RES_Cpu)
        .TextMatrix(0, 10) = MyLoadResString(k_RES_IO)
        .TextMatrix(0, 11) = MyLoadResString(k_RES_MemUsage)
        .TextMatrix(0, 12) = MyLoadResString(k_RES_LoginTime)
        .TextMatrix(0, 13) = MyLoadResString(k_RES_LastBatch)
        .TextMatrix(0, 14) = MyLoadResString(k_RES_Host)
        .TextMatrix(0, 15) = MyLoadResString(k_RES_NetLib)
        .TextMatrix(0, 16) = MyLoadResString(k_RES_NetAddress)
        .TextMatrix(0, 17) = MyLoadResString(k_RES_Blocked)
        .TextMatrix(0, 18) = MyLoadResString(k_RES_Blocking)
        .ColWidth(0) = 600
        .ToolTipText = ""

        If Not oQry Is Nothing Then
            If oQry.ResultSets <> 0 Then

                For iSet = 1 To oQry.ResultSets
                    oQry.CurrentResultSet = iSet
                    iCol = oQry.Columns
                    iRow = oQry.Rows

                    On Local Error Resume Next
                    For irLoop = 1 To iRow

                        sBuf = Format$(oQry.GetColumnLong(irLoop, kSpid)) & vbTab & oQry.GetColumnString(irLoop, kUser) & vbTab & oQry.GetColumnString(irLoop, kDBName) & vbTab & oQry.GetColumnString(irLoop, kStatus)
                        sBuf = sBuf & vbTab & oQry.GetColumnLong(irLoop, kOpenTran) & vbTab & oQry.GetColumnString(irLoop, kCmd) & vbTab & oQry.GetColumnString(irLoop, kPrgName) & vbTab & oQry.GetColumnString(irLoop, kWaitTime) & vbTab & oQry.GetColumnString(irLoop, kWaitType)
                        sBuf = sBuf & vbTab & Format$(oQry.GetColumnLong(irLoop, kCPU)) & vbTab & Format$(oQry.GetColumnLong(irLoop, kIO)) & vbTab & Format$(oQry.GetColumnLong(irLoop, kMemUsage))
                        sBuf = sBuf & vbTab & oQry.GetColumnString(irLoop, kLoginTime) & vbTab & oQry.GetColumnString(irLoop, kLastBatch) & vbTab & oQry.GetColumnString(irLoop, kHostName)
                        sBuf = sBuf & vbTab & oQry.GetColumnString(irLoop, kNetLib) & vbTab & oQry.GetColumnString(irLoop, kNetAddress) & vbTab & Format$(oQry.GetColumnLong(irLoop, kBlocked)) & vbTab & Format$(oQry.GetColumnLong(irLoop, kBlocking))

                        Debug.Print sBuf
                        .AddItem sBuf
                        .Row = .Rows - 1
                        .Col = 0
                        sBuf = Trim$(oQry.GetColumnString(irLoop, kStatus))
                        If StrComp(sBuf, "sleeping", vbTextCompare) = 0 Then
                            iIco = 120
                        ElseIf StrComp(sBuf, "background", vbTextCompare) = 0 Then
                            iIco = 121
                        ElseIf StrComp(sBuf, "runnable", vbTextCompare) = 0 Then
                            iIco = 121
                        End If
                        Set Flex.CellPicture = LoadResPicture(iIco, vbResBitmap)
                        .CellPictureAlignment = flexAlignLeftCenter

                    Next
                    On Local Error GoTo 0
                Next
            End If
        End If

        SetFixedRows Flex
        .Redraw = True
    End With
    On Local Error GoTo 0
    Set oQry = Nothing

    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadLocks()

    Dim iSet As Integer
    Dim iRow As Integer, irLoop As Integer
    Dim oQry As QueryResults
    Dim oColLocks As clsLocks
    Dim oLock As clsLock
    Dim iSpid As Integer
    
        
    Set collLock = New Collection
    
    
    Set oQry = objServer.EnumLocks
    
    If Not oQry Is Nothing Then
        If oQry.ResultSets <> 0 Then

            For iSet = 1 To oQry.ResultSets
                oQry.CurrentResultSet = iSet
                iRow = oQry.Rows

                On Local Error Resume Next
                For irLoop = 1 To iRow
                
'Private iReq_Spid As Integer ' Process ID of the process requesting the lock.
'Private sLockType As String '  nvarchar(36)    A text description of a locking mode. For more information about interpreting values, see the description of the system table syslockinfo column req_mode.
'Private sTableName As String 'nvarchar(129)   If applicable, the name of the table against which the lock is applied.
'Private sIndexName As String 'nvarchar(129)   If applicable, the name of the index against which the lock is applied.
'Private sDbName As String 'nvarchar(129)   Name of the database in which the locked resource is defined.
'Private iStatus As Integer 'tinyint An integer indicating lock application status. For more information about interpreting values, see the description of the system table syslockinfo column req_status.
'1=Granted
'2=Converting
'3=Waiting
                    iSpid = oQry.GetColumnLong(irLoop, 1)
                    
                    On Local Error Resume Next
                    Set oColLocks = collLock.Item("*" & iSpid)
                    If Err.Number <> 0 Then
                        Set oColLocks = New clsLocks
                        oColLocks.Spid() = iSpid
                        collLock.Add oColLocks, "*" & iSpid
                        Set oColLocks = collLock.Item("*" & iSpid)
                    End If
                    Debug.Print oQry.GetColumnString(irLoop, 4)
                    Set oLock = New clsLock
                    With oLock
                        .Spid() = iSpid
                        .LockType() = oQry.GetColumnString(irLoop, 2)
                        .LockTbName() = oQry.GetColumnString(irLoop, 3)
                        .LockIdxName() = oQry.GetColumnString(irLoop, 4)
                        .LockDbName() = oQry.GetColumnString(irLoop, 5)
                        .LockStatus() = oQry.GetColumnLong(irLoop, 6)
                        .Key() = irLoop
                    End With
                    
                    oColLocks.AddItem oLock
                Next
            Next
        End If
    End If
    Set oLock = Nothing
    Set oQry = Nothing
    

    For irLoop = 1 To collLock.Count
        Set mNode = tView.Nodes.Add("locksID", tvwChild, "*" & collLock.Item(irLoop).Spid())
        mNode.Text = "spid " & collLock.Item(irLoop).Spid()
        mNode.Image = "LockID"   ' Set Image property.
        mNode.Expanded = True
    Next
        
    Screen.MousePointer = vbDefault

End Sub
Private Sub Load_Spid(ByVal szKey As String)

'Private iReq_Spid As Integer ' Process ID of the process requesting the lock.
'Private sLockType As String '  nvarchar(36)    A text description of a locking mode. For more information about interpreting values, see the description of the system table syslockinfo column req_mode.
'Private sTableName As String 'nvarchar(129)   If applicable, the name of the table against which the lock is applied.
'Private sIndexName As String 'nvarchar(129)   If applicable, the name of the index against which the lock is applied.
'Private sDbName As String 'nvarchar(129)   Name of the database in which the locked resource is defined.
'Private iStatus As Integer 'tinyint An integer indicating lock application status. For more information about interpreting values, see the description of the system table syslockinfo column req_status.
'1=Granted
'2=Converting
'3=Waiting

    Dim sBuf As String
    Dim oCollLock As clsLocks
    Dim i As Integer
    

    Set oCollLock = collLock.Item(szKey)
    
    With Flex
        .Clear
        .Redraw = False
        .Cols = 5
        .Rows = 1
        
        .TextMatrix(0, 0) = MyLoadResString(k_DbName)
        .TextMatrix(0, 1) = MyLoadResString(k_Mode)
        .TextMatrix(0, 2) = MyLoadResString(k_Status)
        .TextMatrix(0, 3) = MyLoadResString(k_TableName)
        .TextMatrix(0, 4) = MyLoadResString(k_IndexName)
        .ToolTipText = ""
        
        .ColWidth(0) = TextWidth(MyLoadResString(k_DbName)) + 120

        For i = 1 To oCollLock.Count
            sBuf = oCollLock.Item(i).LockDbName() & vbTab & oCollLock.Item(i).LockType() & vbTab & oCollLock.Item(i).strLockStatus() & vbTab & oCollLock.Item(i).LockTbName() & vbTab & oCollLock.Item(i).LockIdxName()
            .AddItem sBuf
        Next
                        
        
        .AllowUserResizing = flexResizeColumns
        SetFixedRows Flex
        .Redraw = True
    End With
    On Local Error GoTo 0

End Sub
Private Sub LoadUsersProp(ByVal sDbName As String)

    Dim colAvailUsers As Collection
        
    Dim lErr As Long, sErr As String
    
    Screen.MousePointer = vbHourglass
    lErr = ErrGetUsersMapping(sDbName, colAvailUsers, sErr)
    If lErr = 0 Then
        Screen.MousePointer = vbDefault
        fUserP.UsersData(colAvailUsers) = sDbName
        fUserP.Show vbModal, Me
    Else
        Screen.MousePointer = vbDefault
        MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End If
    
End Sub
Private Sub ExportStruct2Html(ByVal sDB As String)

    Dim oF As New fExpHTM
    oF.DBName() = sDB
    oF.Show vbModal, Me
    Set oF = Nothing
End Sub
Private Function TestProcInfo(ByVal lSpid As Long, ByVal iStatus As SPID_Action) As Boolean

    Const kSpid = 1
    Const kUser = 2
'    Const kDBName = 3
'    Const kStatus = 4
'    Const kOpenTran = 5
'    Const kCmd = 6
'    Const kPrgName = 7
'    Const kWaitTime = 8
'    Const kWaitType = 9
'    Const kCPU = 11
'    Const kIO = 12
'    Const kMemUsage = 13
'    Const kLoginTime = 14
'    Const kLastBatch = 15
'    Const kHostName = 16
'    Const kNetLib = 17
'    Const kNetAddress = 18
'    Const kBlocked = 19
'    Const kBlocking = 20
'
    
    Dim iWrkSpid As Integer
    Dim ltmpSpid As Long
    Dim bRet As Boolean
    Dim iRow As Integer, irLoop As Integer
    Dim sBuf As String, sUser As String
    
    Dim oQry As QueryResults


    Screen.MousePointer = vbHourglass
    iWrkSpid = SetSpidInfo()
    If iWrkSpid <> 0 Then
        sBuf = Replace("EXEC dbo.sp_MSget_current_activity @id=1%,@option=1", "1%", iWrkSpid)
        On Local Error Resume Next
        Set oQry = objServer.ExecuteWithResults(sBuf, Len(sBuf))
        On Local Error GoTo 0
    End If

    If Not oQry Is Nothing And iWrkSpid <> 0 Then
        If oQry.ResultSets <> 0 Then
            oQry.CurrentResultSet = 1
            iRow = oQry.Rows
            For irLoop = 1 To iRow
                ltmpSpid = oQry.GetColumnLong(irLoop, kSpid)
                If ltmpSpid = lSpid Then
                    sUser = Trim$(oQry.GetColumnString(irLoop, kUser))
                    If iStatus = enSPIDDelete Then
                        If StrComp(sUser, "system", vbTextCompare) <> 0 Then bRet = True
                    Else
                        bRet = True
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    Set oQry = Nothing
    Screen.MousePointer = vbDefault
    TestProcInfo = bRet

End Function

Private Function GetSpidQry() As SQLDMO.QueryResults

    Dim sBuf As String
    Dim iSpid As Integer
    Dim oQry As SQLDMO.QueryResults
    
    Const sSql As String = "EXEC dbo.sp_MSget_current_activity @id=1%,@option=1"
    

    iSpid = SetSpidInfo()
    
    If iSpid <> 0 Then
        sBuf = Replace(sSql, "1%", iSpid)
        On Local Error Resume Next
        Set oQry = objServer.ExecuteWithResults(sBuf, Len(sBuf))
        On Local Error GoTo 0
    End If
    If Not oQry Is Nothing Then Set GetSpidQry = oQry
    
End Function
Private Function SetSpidInfo() As Integer
    'Return Spid

    Dim iSpid As Integer
    Dim oQry As SQLDMO.QueryResults
    Dim vVar As Variant
    Const sSql As String = "declare @spid int, @rc int EXEC @rc=dbo.sp_MSset_current_activity @spid OUTPUT"
    
    On Local Error Resume Next
    Set oQry = objServer.ExecuteWithResults(sSql, Len(sSql))
    Debug.Print Err.Description
    On Local Error GoTo 0
    
    If Not oQry Is Nothing Then
        vVar = Qry2Var(oQry)
        Set oQry = Nothing
        If IsArray(vVar) Then
            iSpid = vVar(0, 0)
        End If
    End If
    SetSpidInfo = iSpid

End Function
Private Sub ChangeObjectOwner(ByVal sDbName As String, vDati As Variant, ByVal enType As amChangingOwnerType)

    Dim vntObject As Variant
    Dim iLoop As Integer
    Dim sDbo As String
    Dim sErr As String
    Dim sOwner As String
    Dim sCurOwner As String
    Dim bAdd As Boolean, bRefresh As Boolean
    Dim lType As SQLDMO.SQLDMO_OBJECT_TYPE
    Dim oObj As Object
    Dim oDb As SQLDMO.Database2
           
        
    If IsArray(vDati) Then
        Screen.MousePointer = vbHourglass
        sDbo = WhoIsDBO(objServer.Databases(sDbName))
        For iLoop = 0 To UBound(vDati)
            Select Case enType
                Case amObjTable
                    Set oObj = GetTbByName(objServer, sDbName, vDati(iLoop))
                    lType = SQLDMOObj_UserTable
                Case amObjStoredProc
                    Set oObj = GetSpByName(objServer, sDbName, vDati(iLoop))
                    lType = SQLDMOObj_StoredProcedure
                Case amObjView
                    Set oObj = GetViewByName(objServer, sDbName, vDati(iLoop))
                    lType = SQLDMOObj_View
                Case amObjFunction
                    Set oObj = GetFuncByName(objServer, sDbName, vDati(iLoop))
                    lType = SQLDMOObj_UserDefinedFunction
                    
                'Case amObjDefault
                '    sCurOwner = GetDefaultOwner(objServer, sDbName, vDati(iLoop))
                '    lType = SQLDMOObj_Default
                'Case amObjRule
                '    sCurOwner = GetRuleOwner(objServer, sDbName, vDati(iLoop))
                '    lType = SQLDMOObj_Rule
                Case Else
                    sErr = MyLoadResString(k_RES_Invalid_Value)
                    Exit For
            End Select
            
            If Not oObj Is Nothing Then
                sCurOwner = oObj.Owner
                If StrComp(sCurOwner, "dbo", vbTextCompare) = 0 Then sCurOwner = sDbo
                
                bAdd = Not oObj.SystemObject
                If bAdd Then
                    If Len(sOwner) = 0 Then sOwner = sCurOwner
                    bAdd = StrComp(sCurOwner, sOwner) = 0
                End If
                If bAdd Then
                    vntObject = AddArray2List(Array(vDati(iLoop), sOwner, lType), vntObject)
                End If
            End If
        Next
        Set oObj = Nothing
        
        Screen.MousePointer = vbDefault
        
        If Len(sErr) = 0 And IsArray(vntObject) Then
            bRefresh = True
            
            Dim fWarning As fResult
            Set fWarning = New fResult
            With fWarning
                .Action = act_Null
                .WWrapVisible() = False
                .Caption = Me.Caption
                .tRes.Text = ReplaceMsg(MyLoadResString(k_ChangingPreWarning) & MyLoadResString(k_ChangingPreWarning2), Array("|"), Array(vbCrLf))
                .Show vbModal, Me
            End With
            Unload fWarning
            Set fWarning = Nothing
            Me.Refresh
            
            Dim fChangeOwner As fDbOwner
            Set fChangeOwner = New fDbOwner
            With fChangeOwner
                .DatabaseObjsProp sDbName, vntObject, sOwner
                .Show vbModal, Me
            End With
            Unload fChangeOwner
            Set fChangeOwner = Nothing
            
            On Local Error Resume Next
            sOwner = objServer.Databases(sDbName).Owner
            m_bNeedReattach = Err.Number <> 0
            On Local Error GoTo 0
        Else
            If Len(sErr) <> 0 Then MsgBox sErr, vbInformation Or vbOKOnly, MyLoadResString(kMsgBoxError)
        End If
    End If
    
    If m_bNeedReattach Then
        Beep
        MsgBox Replace(MyLoadResString(k_SqlObjInvalid), "|", vbCrLf), vbInformation Or vbOKOnly, App.EXEName
        Set objServer = Nothing
        LoadRoot
        
        Flex.Redraw = False
        Flex.Clear
        Flex.Rows = 1
        Flex.Cols = 1
        Flex.ColWidth(0) = Flex.Width
        Flex.Redraw = True
    
        Exit Sub
    End If
    
    If bRefresh Then
        Screen.MousePointer = vbHourglass
        Select Case enType
            Case amObjTable
                objServer.Databases(sDbName).Tables.Refresh True
                LoadTables sDbName
            Case amObjStoredProc
                objServer.Databases(sDbName).StoredProcedures.Refresh True
                LoadStoredProc sDbName
            Case amObjView
                objServer.Databases(sDbName).Views.Refresh True
                LoadView sDbName
            Case amObjFunction
                Set oDb = objServer.Databases(sDbName)
                oDb.UserDefinedFunctions.Refresh True
                Set oDb = Nothing
                LoadFunctions sDbName

'            Case amObjDefault
'                objServer.Databases(sDbName).Defaults.Refresh True
'                'Loaddefaults sDbName
'            Case amObjRule
'                objServer.Databases(sDbName).Rules.Refresh True
'                'LoadRoles sDbName
        End Select
        stbMain.Panels(1).Text = "Ready"
        Screen.MousePointer = vbDefault
    End If
    
    
End Sub
Private Function AddArray2List(ByVal vArr2Add As Variant, ByVal vInitialArr As Variant) As Variant

    Dim iCols As Integer
    Dim iRows As Integer
    Dim iLoopC As Integer, iLoopR As Integer
    
    
    iRows = 0
    iCols = UBound(vArr2Add)
    
    If Not IsEmpty(vInitialArr) Then iRows = UBound(vInitialArr, 1) + 1
    
    ReDim vntRet(iRows, iCols)
    
    If Not IsEmpty(vInitialArr) Then
        For iLoopR = 0 To iRows - 1
            For iLoopC = 0 To iCols
                vntRet(iLoopR, iLoopC) = vInitialArr(iLoopR, iLoopC)
            Next
        Next
    End If
    
    For iLoopC = 0 To iCols
        vntRet(iRows, iLoopC) = vArr2Add(iLoopC)
    Next
    
    AddArray2List = vntRet
    
End Function
Private Sub LoadDetachedNode()
    
    Dim arr_Bflags(1) As Boolean
    Dim sBuf As String
    Dim bShowAll As Boolean
    Dim vBuf As Variant
    ReDim oUdt(0) As dbDetach
    Dim bContinue As Boolean
    Dim sCommand As String
    Dim vArrFiles As Variant
    Dim sErrorReading As String
    Dim sErr As String
    
    Dim bErr As Boolean
    Dim dmoDatabase As New SQLDMO.Database2
    Dim i As Integer, iC As Integer, iLoop As Integer, iFiles As Integer
    Dim bRemove As Boolean
    Dim sUser As String
    Dim bIsLocalHost As Boolean
    Dim sNotAvailable As String
    
    
    Screen.MousePointer = vbHourglass
    
    bIsLocalHost = StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0

        
    If bIsLocalHost Then  'LocalHost
        bShowAll = CBool(Val(GetIni("CONNECT", "SHOWALL", App.Path & "\" & App.EXEName & ".INI")))
                
        ReDim oUdt(i) As dbDetach
        On Local Error Resume Next
            
        sBuf = Dir$(txtServer.Text & "*.mdf")
        Do While Len(sBuf) And Err.Number = 0
            sBuf = Left$(sBuf, InStr(sBuf, ".") - 1)
            If Len(sBuf) Then
                If StrComp(sBuf, "DISTMDL", vbTextCompare) <> 0 Then
                    If i > 0 Then ReDim Preserve oUdt(i) As dbDetach
                    oUdt(i).sName = sBuf
                    oUdt(i).bDetach = False
                    i = i + 1
                    bContinue = True
                End If
            End If
            sBuf = Dir$
        Loop
        On Local Error GoTo 0
    Else
        sNotAvailable = MyLoadResString(kNotAvailable)
        sCommand = txtServer.Text
        sCommand = Replace("EXECUTE dbo.xp_dirtree N'1%',1,1", "1%", sCommand)
        
        Dim oQry As SQLDMO.QueryResults
        On Local Error Resume Next
        Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
        On Local Error GoTo 0
        
        If Not oQry Is Nothing Then vArrFiles = Qry2Var(oQry)
        Set oQry = Nothing
        
        If IsArray(vArrFiles) Then
            BB_Sort vArrFiles
            PurgeFilesDetached vArrFiles, "*.MDF"
        End If
        
        If IsArray(vArrFiles) Then
            ReDim oUdt(iLoop) As dbDetach
            For iLoop = 0 To UBound(vArrFiles, 2)
                sBuf = vArrFiles(0, iLoop)
                If InStr(sBuf, ".") <> 0 Then sBuf = Left$(sBuf, InStr(sBuf, ".") - 1)
                If Len(sBuf) <> 0 Then
                    If StrComp(sBuf, "DISTMDL", vbTextCompare) <> 0 Then
                        If i > 0 Then ReDim Preserve oUdt(i) As dbDetach
                        oUdt(i).sName = sBuf
                        oUdt(i).bDetach = False
                        i = i + 1
                        bContinue = True
                    End If
                End If
            Next
        End If
    End If
    
    If bContinue Then
        sUser = objServer.TrueLogin
        iC = UBound(oUdt)
        For Each dmoDatabase In objServer.Databases     'Do for each Database in the Server
            Debug.Print dmoDatabase.Name
            
            sBuf = ""
            bRemove = True
            On Local Error Resume Next
            'test for errors referencing dbfiles (off-line)
            iFiles = dmoDatabase.FileGroups.Count
            bErr = Err.Number
            sErr = Err.Description
            If bErr Then sErrorReading = sErrorReading & IIf(Len(sErrorReading) <> 0, vbCrLf, "") & ReplaceMsg(MyLoadResString(k_DbDetacheRefWarning), Array("1%", "2%"), Array(dmoDatabase.Name, sErr))
            On Local Error GoTo 0
            If Not bErr Then
                For iLoop = 1 To dmoDatabase.FileGroups.Count
                    For iFiles = 1 To dmoDatabase.FileGroups(iLoop).DBFiles.Count
                        On Local Error Resume Next
                        sBuf = Trim$(dmoDatabase.FileGroups(iLoop).DBFiles(iFiles).PhysicalName)
                        Debug.Print Err.Description
                        bErr = Err.Number <> 0
                        If bErr Then sBuf = dmoDatabase.Name
                        On Local Error GoTo 0
                        
                        If Not bErr And Len(sBuf) <> 0 Then
                            vBuf = Split(sBuf, "\")
                            sBuf = vBuf(UBound(vBuf))
                            sBuf = UCase$(sBuf)
                            If InStr(1, sBuf, ".MDF", vbTextCompare) <> 0 Then
                                sBuf = Left$(sBuf, InStr(1, sBuf, ".MDF", vbTextCompare) - 1)
                                bRemove = True
                            Else
                                bRemove = False
                            End If
                        End If
                        
                        If bErr Then
                            bRemove = True
                            bErr = False
                        End If
                        
                        If bErr = 0 And Not bShowAll Then
                            On Local Error Resume Next
                            bRemove = Not dmoDatabase.IsUser(sUser) Or Not objServer.SaLogin
                            bErr = Err.Number <> 0
                            On Local Error GoTo 0
                        End If
                        
                        If bRemove Then RemoveFromList sBuf, oUdt()
                        
                    Next
                Next
            Else
                On Local Error Resume Next
                sBuf = Trim$(dmoDatabase.FileGroups("PRIMARY").DBFiles(1).PhysicalName)
                Debug.Print Err.Description
                bErr = Err.Number <> 0
                If bErr Then sBuf = dmoDatabase.Name
                On Local Error GoTo 0

                If Not bErr And Len(sBuf) <> 0 Then
                    vBuf = Split(sBuf, "\")
                    sBuf = vBuf(UBound(vBuf))
                    sBuf = UCase$(sBuf)
                    If InStr(1, sBuf, ".MDF", vbTextCompare) <> 0 Then
                        sBuf = Left$(sBuf, InStr(1, sBuf, ".MDF", vbTextCompare) - 1)
                        bRemove = True
                    Else
                        bRemove = False
                    End If
                End If
    
                If bErr Then
                    bRemove = True
                    bErr = False
                End If
    
                If bErr = 0 And Not bShowAll Then
                    On Local Error Resume Next
                    bRemove = Not dmoDatabase.IsUser(sUser) Or Not objServer.SaLogin
                    bErr = Err.Number <> 0
                    On Local Error GoTo 0
                End If
    
                If bRemove Then RemoveFromList sBuf, oUdt()
            End If
        Next
        Set dmoDatabase = Nothing
    End If
    
    Dim FI As CFileInfo
   
    
    With Flex
        .Clear
        .Redraw = False
        .Cols = 7
        .Rows = 1
        .TextArray(0) = MyLoadResString(k_Name) '"Name"
        .TextArray(1) = MyLoadResString(k_File_Size)
        .TextArray(2) = MyLoadResString(k_Created)
        .TextArray(3) = MyLoadResString(k_Last_Modified)
        .TextArray(4) = MyLoadResString(k_Last_Access)
        .TextArray(5) = MyLoadResString(k_Read_Only)
        .TextArray(6) = MyLoadResString(k_Archive)
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColAlignment(0) = flexAlignLeftCenter
        .ToolTipText = MyLoadResString(k_Detached_Databases)
        iC = UBound(oUdt)
        
        iLoop = 0
        ReDim m_sVntDetached(iLoop) As String
        
        For i = 0 To iC
            'If Not oUDT(i).bDetach And Len(oUDT(i).sName) Then .AddItem oUDT(i).sName
            If Not oUdt(i).bDetach And Len(oUdt(i).sName) Then
            
                If iLoop > UBound(m_sVntDetached) Then ReDim Preserve m_sVntDetached(iLoop) As String
                            
                sBuf = UCase$(txtServer.Text & oUdt(i).sName & ".MDF")
                If bIsLocalHost Then
                    Set FI = New CFileInfo
                    FI.FullPathName = sBuf
                    '
                    ' Fill controls with attributes.
                    '
                    .AddItem FI.DisplayName & vbTab & FI.FormatFileSize(FI.FileSize) & vbTab & FI.FormatFileDate(FI.CreationTime) & vbTab & FI.FormatFileDate(FI.ModifyTime) & vbTab & FI.FormatFileDate(FI.LastAccessTime)
                    arr_Bflags(0) = FI.attrReadOnly
                    arr_Bflags(1) = FI.attrArchive
                    m_sVntDetached(iLoop) = FI.FullPathName
                Else
                    .AddItem oUdt(i).sName & ".MDF" & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable
                    m_sVntDetached(iLoop) = txtServer.Text & oUdt(i).sName & ".MDF"
                    
                End If
                iLoop = iLoop + 1
                .Row = .Rows - 1
                .Col = 5
                
                
                Set Flex.CellPicture = LoadResPicture(IIf(arr_Bflags(0), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter

                .Col = 6
                .Row = .Rows - 1
                Set Flex.CellPicture = LoadResPicture(IIf(arr_Bflags(1), k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            End If
            
        Next
        
        .AllowUserResizing = flexResizeColumns
        SetFixedRows Flex
        .Redraw = True
    End With
    Set FI = Nothing
    
    Set dmoDatabase = Nothing

    If Len(sErrorReading) <> 0 Then
        fResult.Action() = act_Null
        fResult.WWrapVisible() = False
        
        fResult.tRes.Text = MyLoadResString(k_DbDetachedScanWarning) & vbCrLf & vbCrLf & sErrorReading
        fResult.Caption = MyLoadResString(k_GenWarning)
        Beep
        Screen.MousePointer = vbDefault
        fResult.Show vbModal, Me
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub RemoveFromList(ByVal sDbName As String, ByRef oUdt() As dbDetach)

    Dim i As Integer
    'If Not bShowAll Then
        For i = 0 To UBound(oUdt)
            If Not oUdt(i).bDetach Then
                If UCase$(oUdt(i).sName) Like UCase$(sDbName) Then
                    oUdt(i).bDetach = True
                    Exit For
                End If
            End If
        Next
'    Else
'        For i = 0 To UBound(oUdt)
'            If Not oUdt(i).bDetach Then
'                If UCase$(oUdt(i).sName) Like UCase$(sDbName) Then
'                    oUdt(i).bDetach = True
'                    Exit For
'                End If
'            End If
'        Next
'    End If

End Sub

Private Sub txtServer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu txtServer
End Sub
Private Sub SetFixedRows(Flex As MSFlexGrid)

    Dim i As Integer
    
    With Flex
        If .Rows > 1 Then
            .FixedRows = 1
        Else
            .Row = .Rows - 1
            For i = 0 To .Cols - 1
                .Col = i
                .CellBackColor = vbInactiveBorder
            Next
        End If
    End With

End Sub
Private Function DetachWithSP_DETACH_DB(ByVal sDbName As String) As String

    Dim lErr As Long, sErr As String
    Dim sSql As String
    Dim oQry As SQLDMO.QueryResults
'    Dim vVar As Variant
    Dim sRet As String, sMsg As String
    Dim i As Integer, x As Integer
    
    sSql = Replace("EXEC sp_detach_db '1%', 'true'", "1%", sDbName)
    
    On Local Error Resume Next
    Set oQry = objServer.ExecuteWithResultsAndMessages(sSql, Len(sSql), sMsg)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
'    If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
    Set oQry = Nothing
    
'    If IsArray(vVar) Then
'        For i = 0 To UBound(vVar, 2)
'            For x = 0 To UBound(vVar)
'                sRet = sRet & IIf(Len(sRet) <> 0, " ", "") & vVar(x, i)
'            Next
'        Next
'    End If
    If lErr <> 0 Then sRet = MyLoadResString(kMsgBoxError) & " - " & sErr & IIf(Len(sRet) <> 0, vbCrLf, "") & sRet
    sRet = sRet & IIf(Len(sRet) <> 0, vbCrLf, "") & sMsg
    DetachWithSP_DETACH_DB = sRet
    
End Function

