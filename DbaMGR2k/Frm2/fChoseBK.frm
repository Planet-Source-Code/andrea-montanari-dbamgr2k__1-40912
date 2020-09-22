VERSION 5.00
Begin VB.Form fChoseBK 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Destinazione BackUp"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox tBck 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   300
      Index           =   3
      Left            =   5290
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      Caption         =   "Device"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton opt 
      Caption         =   "File"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "fChoseBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bMustExist As Boolean
Private bConfirm As Boolean
Private bIsBackupOp As Boolean
Private m_bIsOnServer As Boolean

Public Property Get IsOnServerFileSystem() As Boolean
    IsOnServerFileSystem = m_bIsOnServer
End Property
Public Property Let IsBackupOp(ByVal bVal As Boolean)
    bIsBackupOp = bVal
End Property

Public Property Let MustExist(ByVal bVal As Boolean)
    bMustExist = bVal
End Property
Public Sub InitDevice()
    Dim i As Integer
    Dim oRegistry As SQLDMO.Registry2
    
    Set oRegistry = objServer.Registry
    tBck.Text = oRegistry.BackupDirectory & "\"
    Set oRegistry = Nothing
      
    For i = 1 To objServer.BackupDevices.Count
        cbo.AddItem objServer.BackupDevices(i).Name
    Next
    
End Sub
Public Property Get Confirm() As Boolean
    Confirm = bConfirm
End Property

Public Property Get DeviceType() As Boolean
    DeviceType = opt(0).Value
End Property
Public Property Get DeviceName() As String
    Dim sBuf As String
    Dim oDev As SQLDMO.BackupDevice
    Dim lErr As Long, sErr As String
    
    If opt(0).Value Then
        sBuf = Trim$(tBck.Text)
    Else
        sBuf = cbo.Text
    End If
    Set oDev = Nothing
    
    DeviceName = sBuf
    
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    
End Property

Private Sub cmd_Click(Index As Integer)

    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim iPos As Integer
    Dim bTypeServers As Boolean
    Dim bErr As Boolean, bCancel As Boolean

    m_bIsOnServer = False
    If Index < 2 Then
        If Index Then
            If opt(0).Value Then
                sFile = Trim$(tBck.Text)
                Dim objNet As cNetMap
                Set objNet = New cNetMap
                sFile = objNet.GetNetworkMappedFileName(sFile)
                Set objNet = Nothing
                tBck.Text = sFile
                
                If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) <> 0 Then bTypeServers = True
                                
                If Len(sFile) Then
                    If bMustExist Then
                        If Not bTypeServers Then
                            On Local Error Resume Next
                            bConfirm = (Len(Dir$(sFile)) <> 0) And Not ((GetAttr(sFile) And vbDirectory) = vbDirectory) '(GetAttr(sFile) = vbDirectory)
                            On Local Error GoTo 0
                        Else
                            bConfirm = True
                            m_bIsOnServer = True
                        End If
                    Else
                        If Right$(sFile, 1) <> "\" Then
                            iPos = InStrRev(sFile, "\")
                            If iPos <> 0 Then
                                'bConfirm = CBool(Len(Dir$(Left$(sFile, iPos))) = 0) And (GetAttr(Left$(sFile, iPos)) = vbDirectory)
                                If Not bTypeServers Then
                                    bConfirm = CBool(Len(Left$(sFile, iPos)) <> 0) And ((GetAttr(Left$(sFile, iPos)) And vbDirectory) = vbDirectory)
                                Else
                                    bConfirm = True
                                    m_bIsOnServer = True
                                End If
                            End If
                        Else
                            On Local Error Resume Next
                            bConfirm = Not CBool((GetAttr(sFile) And vbDirectory) = vbDirectory)
                            On Local Error GoTo 0
                        End If
                    End If
                End If
            Else
                bConfirm = cbo.ListIndex <> -1
            End If
        End If
        Me.Hide
    Else
            
        If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) <> 0 Then
            bTypeServers = True
        End If
        
        If Not bTypeServers Then
        
            Dim obj As cOpen
            Set obj = New cOpen
        
            obj.Title = MyLoadResString(k_BackUp_File_Location)
            obj.Def_Dir() = objServer.Registry.SQLDataRoot & "\Backup\"
            obj.Filter() = "Bak File (*.bak)" & Chr$(0) & "*.bak" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
            obj.formHwnd() = Me.hWnd
        
            If bMustExist Then
                obj.Do_Open
            Else
                obj.Do_Save
            End If
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
        
        Else
            Dim fBrowse As fBr4Fold
            Set fBrowse = New fBr4Fold
            
            With fBrowse
                .FileMustExist() = Not bIsBackupOp
                .Init enB4isSelectFile, MyLoadResString(k_Server_File_System_PathToolTip), objServer.Registry.SQLDataRoot & "\Backup\"
                            
                bErr = .ErrLoadingInit()
                If Not bErr Then
                    .Show vbModal, Me
                    bCancel = .Cancelled()
                    If Not bCancel Then sFile = .FullPath()
                Else
                    .GetError lErr, sErr
                End If
            End With
            Unload fBrowse
            Set fBrowse = Nothing
        End If
        
        If lErr = 0 Then
            If Len(sFile) Then
                If InStr(sFile, ".") = 0 Then sFile = sFile & ".Bak"
                tBck.Text = sFile
            Else
                tBck.Text = objServer.Registry.SQLDataRoot & "\Backup\"
            End If
        End If
        
        If lErr <> 0 Then
            Beep
            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
        End If
        
    End If
End Sub

Private Sub Form_Load()

    Me.Caption = MyLoadResString(k_BackUp_Destination)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    opt(0).Caption = MyLoadResString(k_File_Name)
    opt(1).Caption = MyLoadResString(k_Device)
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fChoseBK = Nothing
End Sub

Private Sub opt_GotFocus(Index As Integer)
    tBck.Enabled = Index = 0
    cmd(3).Enabled = Index = 0
    cbo.Enabled = Index = 1
End Sub

'Private Function GetNetworkMappedFileName(ByVal sFileName As String) As String
'
'    Dim sFile           As String
'    Dim sDrive          As String
'    Dim i               As Long
'    Dim sRegKey         As String
'    Dim sDummy          As String
'    Dim sDrivetoSearch  As String
'
'    'Suppose that j is a mapped network drive
'
'    sDrive = Left$(sFileName, 3)
'    i = GetDriveType(sDrive)
'    If i = DRIVE_REMOTE Then
'        'Yes this is a mapped remote drive
'        'Then we have to use \\computer_name\share_name
'        'instead of using drive name
'        'Let's get remote computer's name and share name
'        sDrivetoSearch = Left$(sDrive, 1)
'        If IsWindowsNT Then
'            sRegKey = "HKEY_CURRENT_USER\NETWORK\" & sDrivetoSearch
'        Else
'            sRegKey = "HKEY_CURRENT_USER\NETWORK\Persistent\" & sDrivetoSearch
'        End If
'        sDummy = GetStringValue(sRegKey, "RemotePath")
'
'        i = InStr(sDummy, Chr$(0))
'        If i <> 0 Then sDummy = Left$(sDummy, i - 1)
'        sFile = sDummy & Mid$(sFileName, 3)
'        'Now you can backup to j
'    Else
'        sFile = sFileName
'    End If
'    GetNetworkMappedFileName = sFile
'
'End Function
