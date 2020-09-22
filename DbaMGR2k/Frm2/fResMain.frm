VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fResMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Restore Database"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
   Icon            =   "fResMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6405
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
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   5040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6135
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Text            =   "cbo"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Frame fraParam 
         Caption         =   "Parameter - Restore da Device"
         Height          =   2295
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   5895
         Begin VB.TextBox tDati 
            Height          =   285
            Index           =   5
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   3735
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Aggiungi"
            Height          =   360
            Index           =   3
            Left            =   4320
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Elimina"
            Enabled         =   0   'False
            Height          =   360
            Index           =   4
            Left            =   4320
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Contenuto"
            Enabled         =   0   'False
            Height          =   360
            Index           =   5
            Left            =   4320
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            Caption         =   "DataBase Complete"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Value           =   -1  'True
            Width           =   3615
         End
         Begin VB.OptionButton opt 
            Caption         =   "DataBase Differential"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   1560
            Width           =   3615
         End
         Begin VB.OptionButton opt 
            Caption         =   "Transaction Log"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   8
            Top             =   1920
            Width           =   3615
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   5
            Text            =   "fResMain.frx":000C
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lbl 
            Caption         =   "Backup Numero"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Label lbl 
         Caption         =   "Restore come DB"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Frame fraParam 
         Caption         =   "Recovery Completation State"
         Height          =   855
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   6015
         Begin VB.OptionButton opt 
            Caption         =   "Imposta DB NON operativo - altri Transaction log Restore seguono"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   5655
         End
         Begin VB.OptionButton opt 
            Caption         =   "Imposta DB operativo - nessun altro Transaction log Restore successivo"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   5655
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "Forza Restore a sovrascrivere DB esistenti"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1931
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         Caption         =   "Nomi dei File di Restore"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   4335
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Opzioni"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fResMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LogicLoc() As String
Private PhisicLoc() As String

Private bFileBak As Boolean
Private m_bisOnServerFileSystem As Boolean

Private m_blnAnnullaMod As Boolean
Private miRow As Integer        'Riga Flex
Private miCol As Integer
Private blnDone As Boolean

Private Sub RestoreDB()
    
    Dim lErr As Long, sErr As String
    Dim lAction As Long
    Dim sFileMove As String
    Dim iFileN As Integer
    
    blnDone = True
    
    sFileMove = GetFileLocation()
    iFileN = Val(tDati(0).Text)
        
    If opt(0).Value Then
        lAction = SQLDMORestore_Database
    ElseIf opt(1).Value Then
        lAction = SQLDMORestore_Database '?? Differential Restore
    ElseIf opt(2).Value Then
        lAction = SQLDMORestore_Log
    End If
    
    If Len(Trim$(cbo(0).Text)) = 0 Then
        lErr = 424
        On Local Error Resume Next
        Err.Raise lErr
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    
    If lErr = 0 Then
        If Len(sFileMove) = 0 Then
            lErr = 76
            On Local Error Resume Next
            Err.Raise lErr
            sErr = Err.Description
            On Local Error GoTo 0
        End If
    End If
    If lErr = 0 Then
        If iFileN = 0 Then
            lErr = 52
            On Local Error Resume Next
            Err.Raise lErr
            sErr = Err.Description
            On Local Error GoTo 0
        End If
    End If
    
    If lErr = 0 Then
        Screen.MousePointer = vbHourglass
        With fPercent
            fPercent.Caption = "Restore DB ['" & cbo(0).Text & "']"
            .DBName() = cbo(0).Text
            If Not bFileBak Then
                .BckDevice() = tDati(5).Text
            Else
                .BckFiles() = tDati(5).Text
            End If
            .OtherRestore() = opt(6).Value
            .ReplaceDB() = chk(0).Value <> vbUnchecked
            .Rest_FileN = iFileN
            .Rest_MoveFile() = sFileMove
                    
            .Action() = lAction
            .Show
            .ExecRest m_bisOnServerFileSystem
            lErr = .ErrCode()
            sErr = .ErrMsg()
        End With
        Unload fPercent
    End If
    
    Screen.MousePointer = vbDefault
    Me.Show
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
End Sub
Private Function GetFileLocation() As String

    Dim i As Integer, sVal As String
    For i = 0 To UBound(PhisicLoc)
        If Len(sVal) Then sVal = sVal & ","
        If Len(PhisicLoc(i)) Then sVal = sVal & "[" & LogicLoc(i) & "],[" & PhisicLoc(i) & "]"
    Next
    GetFileLocation = sVal
    
End Function
Public Sub InitDB(ByVal sVal As String)
    
    Dim i As Integer
    For i = 1 To objServer.Databases.Count
        cbo(0).AddItem objServer.Databases(i).Name
    Next
    If Len(sVal) = 0 Then sVal = "master"
    cbo(0).ListIndex = GetItem(sVal, cbo(0))
    
End Sub

Private Sub cbo_Click(Index As Integer)
    cmd(3).Enabled = True
    cmd(4).Enabled = False
    cmd(5).Enabled = False
    tDati(5).Text = ""
    tDati(0).Text = 1
    tDati(0).Enabled = False
    
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    If Index = 0 Then ComboLimit cbo(Index), k_MaxDBLen
End Sub

Private Sub cmd_Click(Index As Integer)
    
    TestActiveControl
    Select Case Index
        Case 0  'Esci
            If blnDone Then frmMain.ReInit
            Unload Me
        Case 1  'Ok
            RestoreDB
        Case 3  'AddDeviceName
            AddDeviceName
        Case 4  'DropDeviceName
            bFileBak = False
            tDati(5).Text = ""
            cmd(3).Enabled = True
            cmd(5).Enabled = False
            cmd(4).Enabled = False
            tDati(0).Enabled = False
            tDati(0).Text = 1
        Case 5  'Showcontent
            ShowContent
    End Select

End Sub
Private Sub ShowRestoreFiles()

    Dim i As Integer, l As Integer
    Dim iCols As Integer, lRows As Long, iFileN As Integer
    Dim sBuf As String
    Dim sBuf2 As String
    Dim lMaxSize As Long
    Dim sFiles As String

    Dim oRest As SQLDMO.Restore2
    Set oRest = New SQLDMO.Restore2
    
    If Not bFileBak Then
        Dim oDev As SQLDMO.BackupDevice
        On Local Error Resume Next
        Set oDev = objServer.BackupDevices(tDati(5).Text)
        On Local Error GoTo 0
        If Not oDev Is Nothing Then sFiles = oDev.PhysicalLocation()
        Set oDev = Nothing
    Else
        If Not m_bisOnServerFileSystem Then
            sFiles = GetShortPath(tDati(5).Text)
        Else
            sFiles = "[" & tDati(5).Text & "]"
        End If
    End If
    
    ReDim LogicLoc(0) As String
    ReDim PhisicLoc(0) As String
    
    With Flex(1)
        .Redraw = False
        .FixedCols = 0
        .FixedRows = 0
        .Clear
        .Rows = 1
        .Cols = 2
    End With
    
    With Flex(0)
        .Redraw = False
        .FixedCols = 0
        .FixedRows = 0
        .Clear
        .Rows = 1
        .Cols = 6
    
        If Len(sFiles) <> 0 And (Val(tDati(0).Text) > 0) Then
            oRest.Files = sFiles
            oRest.FileNumber = Val(tDati(0).Text)
            If Not oRest Is Nothing Then
                Dim oQry As SQLDMO.QueryResults
                On Local Error Resume Next
                Set oQry = oRest.ReadFileList(objServer)
                On Local Error GoTo 0
                
                If Not oQry Is Nothing Then
                    If oQry.ResultSets <> 0 Then
                        oQry.CurrentResultSet = 1
                        iCols = oQry.Columns
                        lRows = oQry.Rows
                        .Cols = iCols
                        For i = 0 To iCols - 1
                            .TextMatrix(0, i) = oQry.ColumnName(i + 1)
                            If i < 2 Then Flex(1).TextMatrix(0, i) = oQry.ColumnName(i + 1)
                        Next
            
                        On Local Error Resume Next
                        For l = 1 To lRows
                            For i = 1 To iCols
                                If i > 1 Then sBuf = sBuf & vbTab
                                If i < 3 Then
                                    If i > 1 Then sBuf2 = sBuf2 & vbTab
                                End If
                                Select Case oQry.ColumnType(i)
                                    Case SQLDMO_DTypeUnknown
                                        sBuf = sBuf & MyLoadResString(k_Unknown)
                                    Case SQLDMO_DTypeBinary
                                        sBuf = sBuf & "..Binary"
                                    Case SQLDMO_DTypeImage
                                        sBuf = sBuf & "..BLOB"
                                    Case SQLDMO_DTypeText, SQLDMO_DTypeNText
                                        sBuf = sBuf & "..MEMO"
                                    Case SQLDMO_DTypeGUID
                                        sBuf = sBuf & "..Guid"
                                    Case SQLDMO_DTypeUChar, SQLDMO_DTypeChar, SQLDMO_DTypeUVarchar, SQLDMO_DTypeVarchar
                                        sBuf = sBuf & oQry.GetColumnString(l, i)
                                        If i < 3 Then sBuf2 = sBuf2 & oQry.GetColumnString(l, i)
                                    Case SQLDMO_DTypeMoney
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.0000")
                                    Case SQLDMO_DTypeFloat4, SQLDMO_DTypeMoney
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.0000")
                                    Case SQLDMO_DTypeFloat8
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.00000000")
                                    Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
                                        sBuf = sBuf & Format$(oQry.GetColumnDate(l, i), kYEAR)
                                    Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                                        If i = 6 Then
                                            lMaxSize = oQry.GetColumnLong(l, i)
                                            If lMaxSize < 0 Then
                                                'sBuf = sBuf & MyLoadResString(k_Unlimited) & "/" & MyLoadResString(k_Unknown)
                                                sBuf = sBuf & MyLoadResString(k_Unknown)
                                            Else
                                                sBuf = sBuf & Format$(lMaxSize, "#,0")
                                            End If
                                        Else
                                            sBuf = sBuf & Format$(oQry.GetColumnLong(l, i), "#,0")
                                        End If
                                    Case SQLDMO_DTypeBit
                                        sBuf = sBuf & Format$(oQry.GetColumnBool(l, i))
                                End Select
                                If i = 1 Then
                                    ReDim Preserve LogicLoc(iFileN) As String
                                    LogicLoc(iFileN) = oQry.GetColumnString(l, i)
                                End If
                                If i = 2 Then
                                    ReDim Preserve PhisicLoc(iFileN) As String
                                    PhisicLoc(iFileN) = oQry.GetColumnString(l, i)
                                    iFileN = iFileN + 1
                                End If
                            Next
    
                            .AddItem sBuf
                            Flex(1).AddItem sBuf2
                            sBuf = ""
                            sBuf2 = ""
                        Next
                    End If
                End If
            End If
        End If
        If .Rows = 1 Then .AddItem ""
        .FixedRows = 1
        .ScrollBars = flexScrollBarBoth
        .Row = 1
        .TopRow = 1
        .Redraw = True
    End With
    With Flex(1)
        If .Rows = 1 Then .AddItem ""
        .ColWidth(0) = 1 / 2 * .Width
        .ColWidth(1) = .Width - .ColWidth(0) - 300
        .FixedRows = 1
        '.FixedCols = 1
        .ScrollBars = flexScrollBarBoth
        .Row = 1
        .TopRow = 1
        .Redraw = True
    End With
    SettaLargFlex Flex(0), 1, 500
    SettaLargFlex Flex(1), 1, 500
    
    Set oQry = Nothing
    Set oRest = Nothing
End Sub

Private Sub Flex_DblClick(Index As Integer)
    If Index Then
        'If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then FlexEdit vbKeySpace
        If Flex(Index).MouseRow > 0 Then FlexEdit vbKeySpace
    End If
End Sub

Private Sub Flex_GotFocus(Index As Integer)
    TestActiveControl
End Sub

Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then
        FlexEdit 0
    End If

End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        SendKeys kPREV
    Else
        If Index Then FlexEdit KeyAscii
    End If

End Sub

Private Sub tDati_GotFocus(Index As Integer)
    tDati(Index).SelStart = 0
End Sub

Private Sub tDati_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then Digit_KDown tDati(Index), KeyCode
End Sub

Private Sub tDati_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then Digit_KPress tDati(Index), KeyAscii
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
    ReDim LogicLoc(0) As String
    ReDim PhisicLoc(0) As String
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(3).Caption = MyLoadResString(k_Add)
    cmd(4).Caption = MyLoadResString(k_Remove)
    cmd(5).Caption = MyLoadResString(k_Contents)
    lbl(0).Caption = MyLoadResString(k_Restore_as_DB)
    fraParam(1).Caption = MyLoadResString(k_Parameter_Restore_from_Device)
    lbl(2).Caption = MyLoadResString(k_BackUp_Number)
    
    tStrip.Tabs.Item(1).Caption = MyLoadResString(k_General)
    tStrip.Tabs.Item(2).Caption = MyLoadResString(k_Options)
    chk(0).Caption = MyLoadResString(k_Force_Restore_to_overwrite_existing_DB)
    opt(5).Caption = MyLoadResString(k_Let_DB_operational_no_other_Transaction_log_Restore)
    opt(6).Caption = MyLoadResString(k_Let_DB_NOT_operational_other_Transaction_log_follow)
    
    opt(0).Caption = MyLoadResString(k_Bck_Complete)
    opt(1).Caption = MyLoadResString(k_Bck_Differential)
    opt(2).Caption = MyLoadResString(k_Bck_Log)
    
    frmMain.Enabled = False
    CenterForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.Show
    Set fResMain = Nothing
    
End Sub

Private Sub tDati_LostFocus(Index As Integer)
    If Index = 0 Then ShowRestoreFiles
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
       
    If KeyCode = vbKeyUP Then
        KeyCode = 0
        If miRow > 0 Then
            Flex(1).Row = miRow
            Flex(1).SetFocus
        End If
    ElseIf KeyCode = vbKeyDOWN Then
        If miRow < Flex(1).Rows - 2 Then
            Flex(1).Row = miRow + 2
            Flex(1).SetFocus
        End If
    End If

End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Dim iElem As Integer
    
    iElem = miRow
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93    '<>|[]
            KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex(1).Text
            Flex(1).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(1).SetFocus
            KeyAscii = 0
        Case vbKeyBACK ', vbKeyDELETE
        Case Else
           OverStrike tFlex
    End Select

End Sub
Private Sub tFlex_LostFocus()

    Dim iElem As Integer
    Dim bUpdate As Boolean
      
    iElem = miRow
    If Not m_blnAnnullaMod Then
        tFlex = Trim$(tFlex.Text)
        bUpdate = Len(tFlex.Text)
        
        If bUpdate Then
            If miCol = 0 Then
                LogicLoc(iElem) = Trim$(tFlex.Text)
            Else
                PhisicLoc(iElem) = Trim$(tFlex.Text)
            End If
        End If
    End If
    If bUpdate Then
        Flex(1).TextMatrix(miRow + 1, miCol) = tFlex
        Flex(1).Row = Flex(1).Row - 1
        Flex(1).Row = Flex(1).Row + 1
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    
    m_blnAnnullaMod = True
    miCol = 0
    tFlex.Visible = False
    'Flex(1).SetFocus
    
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    fraMain(0).Visible = i = 0
    fraMain(1).Visible = i <> 0
    
    fraMain(i).ZOrder
    If i = 0 Then
        cbo(0).SetFocus
    Else
        chk(0).SetFocus
    End If
    
End Sub

Private Sub AddDeviceName()
    
    m_bisOnServerFileSystem = False
    
    fChoseBK.MustExist() = True
    fChoseBK.InitDevice
    fChoseBK.Show vbModal, Me
    If fChoseBK.Confirm() Then
        bFileBak = fChoseBK.DeviceType()
        If bFileBak Then
            If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) <> 0 Then m_bisOnServerFileSystem = fChoseBK.IsOnServerFileSystem()
        End If
        
        tDati(5).Text = fChoseBK.DeviceName
        cmd(3).Enabled = False
        cmd(4).Enabled = True
        cmd(5).Enabled = True
        tDati(0).Enabled = True
    End If
    Unload fChoseBK
    ShowRestoreFiles
    
End Sub
Private Sub ShowContent()

    Dim oRest As SQLDMO.Restore2
    Set oRest = New SQLDMO.Restore2
    Dim sFiles As String
    
    If bFileBak Then
        If Not m_bisOnServerFileSystem Then
            sFiles = GetShortPath(tDati(5).Text)
        Else
            sFiles = "[" & tDati(5).Text & "]"
'            Set oRest = Nothing
'            MsgBox Replace(MyLoadResString(kErr_AvalilableOnlyOnHostServer), "1%", objServer.NetName), vbInformation Or vbOKOnly, App.EXEName
'            Exit Sub
        
        End If
    Else
        Dim oDev As SQLDMO.BackupDevice
        Set oDev = objServer.BackupDevices(tDati(5).Text)
        'sFiles = GetShortPath(oDev.PhysicalLocation)
        sFiles = oDev.PhysicalLocation
        Set oDev = Nothing
    End If
    If Len(sFiles) Then
        oRest.Database = cbo(0).Text
        'If bFileBak Then
            'oRest.Files = GetShortPath(tDati(5).Text)
            oRest.Files = sFiles
        'Else
        '    'oRest.Devices = tDati(5).Text
        '    Dim oDev As SQLDMO.BackupDevice
        '    Set oDev = objServer.BackupDevices(tDati(5).Text)
        '    oRest.Files = GetShortPath(oDev.PhysicalLocation)
        '    Set oDev = Nothing
        '
        'End If
        Debug.Print oRest Is Nothing
        If Len(tDati(5).Text) <> 0 And errTestBackUpSet(oRest) = 0 Then
            fBckCont.LoadFromDisk oRest, Val(tDati(0).Text)
            fBckCont.Show vbModal, Me
            If Not fBckCont.Cancelled() Then
                tDati(0).Text = fBckCont.SelectedRestoreFile()
            End If
            Unload fBckCont
            Set fBckCont = Nothing
            ShowRestoreFiles
        Else
            If Len(tDati(5).Text) <> 0 Then MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
        End If
        Set oRest = Nothing
    Else
        MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
    End If
    
End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iElem As Integer
    
    miRow = Flex(1).Row - 1
    sBuf = Flex(1).Text
    
    bEdit = True
    iElem = miRow
    miCol = Flex(1).Col

        
    If miCol = 0 Then
        iLen = k_MaxDBLen
    ElseIf miCol = 1 Then
        iLen = k_MaxPath_DBLen
    End If
    If Flex(1).CellBackColor = vbInactiveBorder Then bEdit = False
        
    bEdit = Len(PhisicLoc(iElem)) <> 0

    If miRow < 0 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
        End Select
    End If
    
    If bEdit Then
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf
        
        tFlex.SelStart = 0
        With Flex(1)
            tFlex.Move fraMain(1).Left + .Left + .CellLeft, fraMain(1).Top + .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
                
        If KeyAscii = vbKeySpace Then KeyAscii = 0
        If KeyAscii <> 0 Then
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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

