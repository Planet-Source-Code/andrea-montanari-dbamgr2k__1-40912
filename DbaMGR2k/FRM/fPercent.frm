VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fPercent 
   Caption         =   "*"
   ClientHeight    =   300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   ControlBox      =   0   'False
   Icon            =   "fPercent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "fPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oBack As SQLDMO.Backup2
Attribute oBack.VB_VarHelpID = -1
Private WithEvents oRest As SQLDMO.Restore2
Attribute oRest.VB_VarHelpID = -1
Private WithEvents oLang As cLang
Attribute oLang.VB_VarHelpID = -1

Private lErr As Long
Private sErr As String
Private lAction As Long

'BackUp
Private sDatabase As String
Private sBackupSetName As String
Private sBackupSetDescription As String
Private sMediaName As String
Private sDevices  As String
Private sFiles As String
Private bInitialize As Boolean
Private bSkipTapeHeader As Boolean
Private bVerifyExp As Boolean
Private iRetainDays As Integer
Private sExpirationDate As String
Private bDropLog As Boolean

'Restore
Private bOtherRestore As Boolean
Private bReplaceDB As Boolean
Private iFileNum As Integer
Private sMoveFiles As String

Public Sub InitializeLanguage(objLang As cLang)
    Set oLang = objLang
End Sub
Public Sub ExecRest(ByVal bIsOnServerFileSystem As Boolean)
    
    Set oRest = New SQLDMO.Restore2
    DoEvents
    With oRest
        .Database = sDatabase
        If Len(sDevices) Then
            Dim oDev As SQLDMO.BackupDevice
            Set oDev = objServer.BackupDevices(sDevices)
            Debug.Print oDev.PhysicalLocation
            .Files = oDev.PhysicalLocation
            Set oDev = Nothing
        
        Else
            If Not bIsOnServerFileSystem Then
                .Files = GetShortPath(sFiles)
            Else
                .Files = "[" & sFiles & "]"
            End If
        End If
        .LastRestore = Not bOtherRestore
        .FileNumber = iFileNum
        .ReplaceDatabase = bReplaceDB
        .RelocateFiles = sMoveFiles
        .Action = lAction
    End With
    On Local Error Resume Next
    oRest.SQLRestore objServer
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    Set oRest = Nothing
    
End Sub

Public Property Let OtherRestore(ByVal bVal As Boolean)
    bOtherRestore = bVal
End Property
Public Property Let ReplaceDB(ByVal bVal As Boolean)
    bReplaceDB = bVal
End Property
Public Property Let Rest_FileN(ByVal iVal As Integer)
    iFileNum = iVal
End Property
Public Property Let Rest_MoveFile(ByVal sVal As String)
    sMoveFiles = sVal
End Property

Public Property Let DropLog(ByVal bVal As Boolean)
    bDropLog = bVal
End Property

Public Property Let VerifyExp(ByVal bVal As Boolean)
    bVerifyExp = bVal
End Property

Public Property Let DBName(ByVal sVal As String)
    sDatabase = sVal
End Property
Public Property Let BckSetName(ByVal sVal As String)
    sBackupSetName = sVal
End Property
Public Property Let BckSetDesc(ByVal sVal As String)
    sBackupSetDescription = sVal
End Property
Public Property Let MediaName(ByVal sVal As String)
    sMediaName = sVal
End Property
Public Property Let BckDevice(ByVal sVal As String)
    sDevices = sVal
End Property
Public Property Let BckFiles(ByVal sVal As String)
    sFiles = sVal
End Property
Public Property Let Init(ByVal bVal As Boolean)
    bInitialize = bVal
End Property
Public Property Let SkipTapeHead(ByVal bVal As Boolean)
    bSkipTapeHeader = bVal
End Property
Public Property Let RetainDay(ByVal iVal As Integer)
    iRetainDays = iVal
End Property
Public Property Let ExpireDate(ByVal sVal As String)
    sExpirationDate = sVal
End Property
Public Property Let Action(ByVal lVal As SQLDMO.SQLDMO_BACKUP_TYPE)
    lAction = lVal
End Property
Public Sub ExecBck()
    
    'Dim bErrExist As Boolean
'    Dim iFile As Integer
    Dim oDev As SQLDMO.BackupDevice
    
    Set oBack = New SQLDMO.Backup2
    DoEvents
    With oBack
        .Database = sDatabase
        If Len(sBackupSetName) Then .BackupSetName = sBackupSetName
        If Len(sBackupSetDescription) Then .BackupSetDescription = sBackupSetDescription
        If Len(sMediaName) Then .MediaName = sMediaName
        
        If Len(sDevices) Then
            '.Devices = sDevices
            Set oDev = objServer.BackupDevices(sDevices)
            Debug.Print oDev.PhysicalLocation
            sFiles = oDev.PhysicalLocation
            .Files = sFiles
        Else
'            If StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0 Then
'                On Local Error Resume Next
'                bErrExist = Len(Dir$(sFiles)) = 0
'                lErr = Err.Number
'                If lErr <> 0 Then bErrExist = True
'                On Local Error GoTo 0
'                If Not bErrExist Then .Files = GetShortPath(sFiles)
'            Else
'                .Files = sFiles
'            End If
            ' sFiles ="[" + physical_position + "]"
            .Files = sFiles
        End If
        
            
'        If bErrExist Then
'            iFile = FreeFile
'            'create an empty file else GetShortPathName will fail
'            On Local Error Resume Next
'            Open sFiles For Binary Access Write As #iFile
'            ' Chiude prima di riaprire in una modalità diversa.
'            Close #iFile
'            Err.Clear
'            bErrExist = Len(Dir$(sFiles)) = 0
'            lErr = Err.Number
'            If lErr <> 0 Then bErrExist = True
'            On Local Error GoTo 0
'            If Not bErrExist Then .Files = GetShortPath(sFiles)
'        End If
        '.Files = sFiles
        
        
        .Initialize = bInitialize
        .SkipTapeHeader = Not bSkipTapeHeader
        If bVerifyExp Then
            If Val(iRetainDays) Then
                .RetainDays = iRetainDays
            Else
                .ExpirationDate = sExpirationDate
            End If
        End If
        If lAction = SQLDMOBackup_Log Then
            If bDropLog Then
                .TruncateLog = SQLDMOBackup_Log_NoLog 'SQLDMOBackup_Log_Truncate
            Else
                .TruncateLog = SQLDMOBackup_Log_NoTruncate
            End If
        End If
        .Action = lAction
    End With
    
    lErr = 0
    On Local Error Resume Next
    oBack.SQLBackup objServer
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    Set oBack = Nothing
End Sub

Public Property Get ErrCode() As Long
    ErrCode = lErr
End Property
Public Property Get ErrMsg() As String
    ErrMsg = sErr
End Property

Private Sub Form_Load()
    CenterForm Me
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fPercent = Nothing
End Sub

Private Sub oBack_Complete(ByVal Message As String)
    Screen.MousePointer = vbDefault
    MsgBox Message, vbInformation Or vbOKOnly, App.EXEName
    Screen.MousePointer = vbHourglass
End Sub

Private Sub oBack_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    pBar.Value = Percent
End Sub

Private Sub oLang_Init(ByVal szMsg As String)
    Me.Caption = szMsg
End Sub

Private Sub oLang_Status(ByVal iPerc As Integer)
    pBar.Value = iPerc
End Sub

Private Sub oRest_Complete(ByVal Message As String)
    Screen.MousePointer = vbDefault
    MsgBox Message, vbInformation Or vbOKOnly, App.EXEName
    Screen.MousePointer = vbHourglass
End Sub

Private Sub oRest_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    pBar.Value = Percent
End Sub
'Public Sub oldExecBck()
'
'    Dim lErr As Long
'    Dim bErrExist As Boolean
'    Dim iFile As Integer
'    Dim oDev As SQLDMO.BackupDevice
'
'    Set oBack = New SQLDMO.Backup
'
'    With oBack
'        .Database = sDatabase
'        If Len(sBackupSetName) Then .BackupSetName = sBackupSetName
'        If Len(sBackupSetDescription) Then .BackupSetDescription = sBackupSetDescription
'        If Len(sMediaName) Then .MediaName = sMediaName
'
'        If Len(sDevices) Then
'            '.Devices = sDevices
'            Set oDev = objServer.BackupDevices(sDevices)
'            Debug.Print oDev.PhysicalLocation
'            On Local Error Resume Next
'            bErrExist = Len(Dir$(oDev.PhysicalLocation)) = 0
'            lErr = Err.Number
'            If lErr <> 0 Then bErrExist = False
'            On Local Error GoTo 0
'            If bErrExist Then
'                iFile = FreeFile
'                'create an empty file else GetShortPathName will fail
'                On Local Error Resume Next
'                Open oDev.PhysicalLocation For Binary Access Write As #iFile
'                ' Chiude prima di riaprire in una modalità diversa.
'                Close #iFile
'                On Local Error GoTo 0
'            End If
'            Set oDev = Nothing
'            .Devices = sDevices
'
'        Else
'            On Local Error Resume Next
'            bErrExist = Len(Dir$(sFiles)) = 0
'            lErr = Err.Number
'            If lErr <> 0 Then bErrExist = False
'            On Local Error GoTo 0
'
'            If bErrExist Then
'                iFile = FreeFile
'                'create an empty file else GetShortPathName will fail
'                On Local Error Resume Next
'                Open sFiles For Binary Access Write As #iFile
'                ' Chiude prima di riaprire in una modalità diversa.
'                Close #iFile
'                On Local Error GoTo 0
'            End If
'            '.Files = sFiles
'            .Files = GetShortPath(sFiles)
'        End If
'        .Initialize = bInitialize
'        .SkipTapeHeader = Not bSkipTapeHeader
'        If bVerifyExp Then
'            If Val(iRetainDays) Then
'                .RetainDays = iRetainDays
'            Else
'                .ExpirationDate = sExpirationDate
'            End If
'        End If
'        If lAction = SQLDMOBackup_Log Then
'            If bDropLog Then
'                .TruncateLog = SQLDMOBackup_Log_NoLog 'SQLDMOBackup_Log_Truncate
'            Else
'                .TruncateLog = SQLDMOBackup_Log_NoTruncate
'            End If
'        End If
'        .Action = lAction
'    End With
'
'    On Local Error Resume Next
'    oBack.SQLBackup objServer
'    lErr = Err.Number
'    sErr = Err.Description
'    On Local Error GoTo 0
'    Set oBack = Nothing
'End Sub
