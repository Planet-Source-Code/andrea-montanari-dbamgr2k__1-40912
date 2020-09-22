VERSION 5.00
Begin VB.Form fBckD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BackUp Device Property"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPhisLoc 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   300
      Index           =   3
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Contenuto"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   3585
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox tBck 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Nome File"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Nome"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "fBckD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bNew As Boolean
Private oMedia As SQLDMO.BackupDevice
Private sInitPath As String

Public Sub LoadDevice(ByVal sVal As String)
    cmd(2).Enabled = True
    cmd(3).Visible = False
    Set oMedia = objServer.BackupDevices(sVal)
    With oMedia
        tBck.Text = oMedia.Name
        txtPhisLoc.Text = .PhysicalLocation
    End With
    bNew = False
    tBck.Locked = True
    
End Sub

Private Sub cmd_Click(index As Integer)

    Select Case index
        Case 0  'unload
            Unload Me
        Case 1  'Ok
            If Not bNew Then
                Unload Me
            Else
                SaveDevice
            End If
        Case 2  'Show Content
            ShowContent
        Case 3  'Select PhisicalLocation
            Select_PhisicalLocation
    End Select
End Sub
Private Sub ShowContent()

    Dim sFiles As String
    Dim oDev As SQLDMO.BackupDevice
    Dim lErr As Long, sErr As String

'    If objServer.NetName = NomeComputer() Then
'
'        Set oDev = objServer.BackupDevices(tBck.Text)
'        sFiles = GetShortPath(oDev.PhysicalLocation)
'        sFiles = oDev.PhysicalLocation
'        Set oDev = Nothing
'
'        If Len(sFiles) Then
'
'            fBckCont.LoadData tBck.Text
'            lErr = fBckCont.ErrReadingDevice(sErr)
'            If lErr = 0 Then
'                fBckCont.Show vbModal, Me
'            Else
'                Unload fBckCont
'                Set fBckCont = Nothing
'                MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
'            End If
'        Else
'            MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
'        End If
'    Else
'        MsgBox Replace(MyLoadResString(kErr_AvalilableOnlyOnHostServer), "1%", objServer.NetName), vbInformation Or vbOKOnly, App.EXEName
'    End If
'    Set oDev = Nothing
    

    Set oDev = objServer.BackupDevices(tBck.Text)
    sFiles = GetShortPath(oDev.PhysicalLocation)
    sFiles = oDev.PhysicalLocation
    Set oDev = Nothing

    If Len(sFiles) Then

        fBckCont.LoadData tBck.Text
        lErr = fBckCont.ErrReadingDevice(sErr)
        If lErr = 0 Then
            fBckCont.Show vbModal, Me
        Else
            Unload fBckCont
            Set fBckCont = Nothing
            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
        End If
    Else
        MsgBox MyLoadResString(k_This_Device_does_not_contain_any_BackUp_sets), vbOKOnly Or vbInformation, App.EXEName
    End If
    Set oDev = Nothing
    
End Sub
Private Sub SaveDevice()
    
    Dim sFile As String
    Dim sName As String
    Dim lErr As Long, sErr As String
    Dim iPos As Integer
    
    sName = Trim$(tBck.Text)
    sFile = Trim$(txtPhisLoc.Text)
    If Len(sFile) = 0 Then
        lErr = Err_Free
        sErr = MyLoadResString(k_B4F_Err_NoFileSelected)
    End If
    If lErr = 0 Then
        If Len(sName) = 0 Then
            lErr = Err_Free
            sErr = MyLoadResString(kErr_No_Name)
        End If
    End If
    
    If lErr = 0 Then
        iPos = InStr(sFile, ".")
        If iPos = 0 Then
            sFile = sFile & ".Bak"
        Else
            If Len(Mid$(sFile, iPos + 1)) = 0 Then
                sFile = sFile & ".Bak"
            End If
        End If
        
        Set oMedia = New SQLDMO.BackupDevice
        oMedia.Name = sName
        oMedia.PhysicalLocation = sFile
        oMedia.Type = SQLDMODevice_DiskDump
        On Local Error Resume Next
        objServer.BackupDevices.Add oMedia
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    
    If lErr <> 0 Then
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        LoadDevice sName
    End If
    
End Sub
Private Sub Select_PhisicalLocation()
    
    Dim bErr As Boolean, bCancel As Boolean
    Dim lErr As Long, sErr As String
    Dim sTarget As String
    'Dim sBuf As String
    
    Dim fBrowse As fBr4Fold
    Set fBrowse = New fBr4Fold
    If Len(sInitPath) = 0 Then sInitPath = objServer.Registry.SQLRootPath
    With fBrowse
    
        Dim vntFiltri(1, 2) As Variant
        
        vntFiltri(0, 0) = "*.Bak"
        vntFiltri(1, 0) = "Back File (*.Bak)"
        
        vntFiltri(0, 1) = "*.Bck"
        vntFiltri(1, 1) = "Bck File (*.Bck)"
        
        vntFiltri(0, 2) = "*.*"
        vntFiltri(1, 2) = MyLoadResString(k_All_Files)
    
        '.Init enB4isSelectDir, MyLoadResString(k_Select_Full_Data_Path) & " BackUp", sInitPath
        .Init enB4SelectFileFilters, MyLoadResString(k_Select_Data_File) & " BackUp", sInitPath, , vntFiltri
                    
        bErr = .ErrLoadingInit()
        If Not bErr Then
            .Show vbModal, Me
            bCancel = .Cancelled()
            If Not bCancel Then sTarget = .FullPath()
        Else
            .GetError lErr, sErr
        End If
    End With
    If Not bErr And Not bCancel Then
        'If Right$(sTarget, 1) <> "\" Then sTarget = sTarget & "\"
        'sInitPath = sTarget
        'sBuf = sTarget 'Trim$(tBck.Text)
        If Len(sTarget) <> 0 Then
            If InStr(sTarget, ".") = 0 Then sTarget = sTarget & ".bak"
        End If
        'txtPhisLoc.Text = sInitPath & sBuf
        txtPhisLoc.Text = sTarget
    Else
        If bErr Then sInitPath = objServer.Registry.SQLDataRoot
    End If
    
    
    If lErr <> 0 Then MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    
'    Exit Sub
'    Dim a As BROWSEINFO
'    'Dim sBuf As String
'
'    Dim pidlBrowse As Long
'    a.lpszTitle = MyLoadResString(k_Select_Full_Data_Path) & " BackUp"
'
'    a.hOwner = Me.hwnd
'    pidlBrowse = SHBrowseForFolder(a)
'    sBuf = GetPathFromPIDL(pidlBrowse)
'    If Len(sBuf) Then
'        If Right$(sBuf, 1) <> "\" Then sBuf = sBuf & "\"
'        sInitPath = sBuf
'    Else
'        sInitPath = objServer.Registry.SQLDataRoot & "\Backup\"
'    End If
'    txtPhisLoc.Text = sInitPath
'
'    If Len(tBck.Text) Then sBuf = tBck.Text & ".bak"
'    txtPhisLoc.Text = sInitPath & sBuf

End Sub
Private Sub Form_Load()
    
    lbl(0).Caption = MyLoadResString(k_Name)
    lbl(1).Caption = MyLoadResString(k_Data_File_Name)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(2).Caption = MyLoadResString(k_Contents)

    
    sInitPath = objServer.Registry.SQLDataRoot & "\Backup\"
    txtPhisLoc.Text = sInitPath
    bNew = True
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oMedia = Nothing
    Set fBckD = Nothing
End Sub

Private Sub tBck_Change()
    txtPhisLoc.Text = sInitPath & Trim$(tBck.Text)
End Sub
