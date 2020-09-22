VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fBckMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BackUp"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   Icon            =   "fBckMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   5655
      Begin VB.Frame fraDati 
         Caption         =   "Destinazione:  Disk"
         Height          =   1935
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   5415
         Begin VB.TextBox tDati 
            Height          =   285
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   480
            Width           =   3735
         End
         Begin VB.CheckBox chk 
            Caption         =   "Aggiungi al Media"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Contenuto"
            Enabled         =   0   'False
            Height          =   360
            Index           =   5
            Left            =   4080
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Elimina"
            Enabled         =   0   'False
            Height          =   360
            Index           =   4
            Left            =   4080
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Aggiungi"
            Height          =   360
            Index           =   3
            Left            =   4080
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraDati 
         Caption         =   "BackUp"
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   5415
         Begin VB.OptionButton opt 
            Caption         =   "Transaction Log"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   4575
         End
         Begin VB.OptionButton opt 
            Caption         =   "DataBase Differential"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   4575
         End
         Begin VB.OptionButton opt 
            Caption         =   "DataBase Complete"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   4575
         End
      End
      Begin VB.TextBox tDati 
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox tDati 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lbl 
         Caption         =   "Descrizione"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Nome"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Database"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   4560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   5655
      Begin VB.CheckBox chk 
         Caption         =   "Rimuovi Dati Scaduti da Log Transazioni"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.Frame fraDati 
         Enabled         =   0   'False
         Height          =   1095
         Index           =   2
         Left            =   0
         TabIndex        =   28
         Top             =   2160
         Width           =   5415
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   14
            Text            =   "fBckMain.frx":000C
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox tDati 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "  /  /    "
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chk 
            Caption         =   "BackUp Set Scade:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton opt 
            Caption         =   "tra giorni:"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt 
            Caption         =   "il:"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   960
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.TextBox tDati 
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   128
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "Verifica Nome MediaSet e Data Scadenza"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Verifica BackUp"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5520
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lbl 
         Caption         =   "Nome Media Set"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   27
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   5775
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10186
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
Attribute VB_Name = "fBckMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bFileBak As Boolean
Private blnDone As Boolean
Private m_bisOnServerFileSystem As Boolean

Public Sub InitDB(ByVal sVal As String)
    
    Dim i As Integer
    For i = 1 To objServer.Databases.Count
        cbo.AddItem objServer.Databases(i).Name
    Next
    If Len(sVal) = 0 Then sVal = "master"
    cbo.ListIndex = GetItem(sVal, cbo)
    
End Sub

Private Sub cbo_Click()
    
    Dim lErr As Long, sErr As String
    Dim bVal As Boolean
    Dim oDB As SQLDMO.Database2
    Dim oOpt2 As SQLDMO.DBOption2
    
    On Local Error Resume Next
    Set oDB = objServer.Databases(cbo.Text)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        Set oOpt2 = oDB.DBOption
        bVal = Not oOpt2.RecoveryModel = SQLDMORECOVERY_Simple
        Set oOpt2 = Nothing
        
        tDati(0).Text = cbo.Text & " BackUp"
        opt(2).Enabled = bVal
        If Not bVal Then opt(0).Value = True
        cmd(3).Enabled = True
        cmd(4).Enabled = False
        cmd(5).Enabled = False
        tDati(5).Text = ""
    Else
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
        
End Sub

Private Sub chk_Click(index As Integer)

    Dim bVal As Boolean
    bVal = chk(index).Value
    Select Case index
        Case 0
            fraDati(2).Enabled = bVal
            If Not bVal Then chk(3).Value = vbUnchecked
        Case 2
            tDati(2).Enabled = bVal
        Case 3
            opt(4).Enabled = bVal
            opt(5).Enabled = bVal
            tDati(3).Enabled = bVal
            tDati(4).Enabled = bVal
    End Select
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo
    'spostamento sulla scheda successiva.
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tStrip.SelectedItem.index
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

    tStrip.Tabs.Item(1).Caption = MyLoadResString(k_General)
    tStrip.Tabs.Item(2).Caption = MyLoadResString(k_Options)
    lbl(1).Caption = MyLoadResString(k_Name)
    lbl(2).Caption = MyLoadResString(k_Description)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    
    cmd(3).Caption = MyLoadResString(k_Add)
    cmd(4).Caption = MyLoadResString(k_Remove)
    cmd(5).Caption = MyLoadResString(k_Contents)
    
    fraDati(1).Caption = MyLoadResString(k_Destination_Disk)
    chk(0).Caption = MyLoadResString(k_Add_to_Media)
    
    lbl(3).Caption = MyLoadResString(k_Media_Set_Name)
    chk(1).Caption = MyLoadResString(k_Verify_BackUp)
    chk(4).Caption = MyLoadResString(k_Remove_old_data_from_Transaction_Log)
    chk(2).Caption = MyLoadResString(k_Verify_MediaSet_Name_and_Expiration_Date)
    
    chk(3).Caption = MyLoadResString(k_BackUp_Set_Expires)
    opt(4).Caption = MyLoadResString(k_In_Days)
    opt(5).Caption = MyLoadResString(k_On)
    
    opt(0).Caption = MyLoadResString(k_Bck_Complete)
    opt(1).Caption = MyLoadResString(k_Bck_Differential)
    opt(2).Caption = MyLoadResString(k_Bck_Log)
    
    
    tDati(4).Text = Format$(Now, kYEAR)
    frmMain.Enabled = False
    CenterForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.Show
    Set fBckMain = Nothing
End Sub

Private Sub opt_GotFocus(index As Integer)

    If index > 3 Then
        tDati(3).Enabled = index = 4
        tDati(4).Enabled = index = 5
    End If
    chk(4).Enabled = index = 2
    
End Sub

Private Sub tDati_Change(index As Integer)
    Static b As Boolean
    If index = 3 Then
        If Val(tDati(index).Text) < 1 Then
            If Not b Then
                b = True
                tDati(index).Text = 1
                b = False
            End If
        End If
    End If
End Sub

Private Sub tDati_GotFocus(index As Integer)
    tDati(index).SelStart = 0
End Sub

Private Sub tDati_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 3 Then
        If KeyCode = vbKeyUP Then
            KeyCode = 0
            If Val(tDati(index).Text) < 99 Then tDati(index).Text = Val(tDati(index).Text) + 1
        ElseIf KeyCode = vbKeyDOWN Then
            KeyCode = 0
            If Val(tDati(index).Text) > 1 Then tDati(index).Text = Val(tDati(index).Text) - 1
        Else
            Digit_KDown tDati(index), KeyCode
        End If
    ElseIf index = 4 Then
        Data_KDown tDati(index), KeyCode, MASKYEAR
    Else
        If KeyCode = vbKeyDOWN Or KeyCode = vbKeyUP Then KeyCode = 0
    End If
End Sub

Private Sub tDati_KeyPress(index As Integer, KeyAscii As Integer)

    Select Case index
        Case 3
            Digit_KPress tDati(index), KeyAscii
        Case 4
            Data_KPress tDati(index), KeyAscii, MASKYEAR
        Case Else
            Select Case KeyAscii
                Case vbKeyEscape, vbKeyReturn
                    KeyAscii = 0
                Case vbKeyBACK, vbKeySpace
                Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95
                Case Else
                    KeyAscii = 0
            End Select
    End Select
    
End Sub

Private Sub tDati_LostFocus(index As Integer)
    If index = 4 Then
        If Not IsDateValid(tDati(index), 10) Then tDati(index).Text = MASKYEAR
    End If
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    
    i = tStrip.SelectedItem.index - 1
    
    fra(0).Visible = i = 0
    fra(1).Visible = i <> 0
    
    fra(i).ZOrder
    If i = 0 Then
        cbo.SetFocus
    Else
        chk(1).SetFocus
    End If
    
End Sub

Private Sub cmd_Click(index As Integer)
    
    Select Case index
        Case 0  'Esci
            If blnDone Then frmMain.ReInit
            Unload Me
        Case 1  'Ok
            If CanBackUp() Then BackUpDB
        Case 3  'AddDeviceName
            AddDeviceName
        Case 4  'DropDeviceName
            bFileBak = False
            tDati(5).Text = ""
            cmd(3).Enabled = True
            cmd(5).Enabled = False
            cmd(4).Enabled = False
        Case 5  'Showcontent
            ShowContent
    End Select
        
End Sub

Private Function CanBackUp() As Boolean
    
    Dim bRet As Boolean
    Dim sErr As String
    bRet = True
    
    If Len(Trim$(tDati(0).Text)) = 0 Then
        bRet = False
        sErr = MyLoadResString(k_Insert_BackUp_Name)
    ElseIf Len(Trim$(tDati(5).Text)) = 0 Then
        bRet = False
        sErr = MyLoadResString(k_Insert_BackUp_Destination)
    End If
        
    If Not bRet Then MsgBox sErr, vbInformation Or vbOKOnly, App.EXEName
    CanBackUp = bRet
    
End Function
Private Sub BackUpDB()
    
    Dim lErr As Long, sErr As String
    Dim lAction As SQLDMO.SQLDMO_BACKUP_TYPE
    
    blnDone = True
    If opt(0).Value Then
        lAction = SQLDMOBackup_Database
    ElseIf opt(1).Value Then
        lAction = SQLDMOBackup_Differential
    ElseIf opt(2).Value Then
        lAction = SQLDMOBackup_Log
    End If
    tDati(0).Text = Trim$(tDati(0).Text)
    tDati(1).Text = Trim$(tDati(1).Text)
    tDati(2).Text = Trim$(tDati(2).Text)
    
    Screen.MousePointer = vbHourglass
    
    With fPercent
        fPercent.Caption = "BackUp DB ['" & cbo.Text & "']"
        .DBName() = cbo.Text
        .BckSetName() = tDati(0).Text
        
        If Len(tDati(1).Text) Then .BckSetDesc() = tDati(1).Text
        If Len(tDati(2).Text) Then .MediaName() = tDati(2).Text
        If Not bFileBak Then
            .BckDevice() = tDati(5).Text
        Else
            .BckFiles() = "[" & tDati(5).Text & "]"
        End If
        
        .Init() = chk(0).Value = vbUnchecked
        
        If chk(2).Value Then .SkipTapeHead() = True
        If chk(3).Value Then
            .VerifyExp = True
            If opt(4).Value Then
                .RetainDay() = Val(tDati(3).Text)
            Else
                .ExpireDate() = Format$(DateValue(tDati(4)), "yyyy/mm/dd")
            End If
        End If
        If chk(4).Enabled Then .DropLog() = chk(4).Value
                
        .Action() = lAction
        .Show
        .ExecBck
        lErr = .ErrCode()
        sErr = .ErrMsg()
    End With
    Me.Show
    Unload fPercent
    
    If lErr = 0 Then
        If chk(1).Value Then
            
            Screen.MousePointer = vbHourglass
            Dim oRest As SQLDMO.Restore
            Set oRest = New SQLDMO.Restore

            oRest.Database = cbo.Text
            If bFileBak Then
                If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0 Then
                    oRest.Files = GetShortPath(tDati(5).Text)
                Else
                    oRest.Files = "[" & tDati(5).Text & "]"
                End If
            Else
                Dim oDev As SQLDMO.BackupDevice
                Set oDev = objServer.BackupDevices(tDati(5).Text)
                'oRest.Files = GetShortPath(oDev.PhysicalLocation)
                oRest.Files = oDev.PhysicalLocation
                Set oDev = Nothing
            End If

            On Local Error Resume Next
            oRest.SQLVerify objServer
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0

            Set oRest = Nothing
        End If
    End If
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
End Sub
Private Sub AddDeviceName()
    
    m_bisOnServerFileSystem = False
    
    fChoseBK.MustExist() = False
    fChoseBK.IsBackupOp() = True
    fChoseBK.InitDevice
    fChoseBK.Show vbModal, Me
    If fChoseBK.Confirm() Then
        
        bFileBak = fChoseBK.DeviceType()
        If bFileBak Then
            If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) <> 0 Then m_bisOnServerFileSystem = fChoseBK.IsOnServerFileSystem()
        End If
        tDati(5).Text = fChoseBK.DeviceName()
        cmd(3).Enabled = False
        cmd(4).Enabled = True
        cmd(5).Enabled = True
    End If
    Unload fChoseBK
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
            'Set oRest = Nothing
            'MsgBox Replace(MyLoadResString(kErr_AvalilableOnlyOnHostServer), "1%", objServer.NetName), vbInformation Or vbOKOnly, App.EXEName
            'Exit Sub
        End If
    Else
        Dim oDev As SQLDMO.BackupDevice
        Set oDev = objServer.BackupDevices(tDati(5).Text)
        'sFiles = GetShortPath(oDev.PhysicalLocation)
        sFiles = oDev.PhysicalLocation
        Set oDev = Nothing
    End If
    
    If Len(sFiles) Then
        oRest.Database = cbo.Text
        oRest.Files = sFiles
        Debug.Print oRest Is Nothing
        If Len(tDati(5).Text) And Not (oRest Is Nothing) Then
            If Len(oRest.Files) <> 0 And errTestBackUpSet(oRest) = 0 Then
                fBckCont.LoadFromDisk oRest, -1
                fBckCont.Show vbModal, Me
                Set fBckCont = Nothing
            Else
                If Len(tDati(5).Text) <> 0 Then MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
            End If
        End If
        Set oRest = Nothing
    Else
        MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
    End If
    
End Sub

