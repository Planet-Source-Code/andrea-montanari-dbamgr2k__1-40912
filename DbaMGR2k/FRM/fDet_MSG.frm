VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fDet_MSG 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   Icon            =   "fDet_MSG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "salva"
      Height          =   360
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox tRes 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5106
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fDet_MSG.frx":000C
   End
End
Attribute VB_Name = "fDet_MSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sDBStuff As String
Private bConfirm As Boolean
Public Property Let DBName(ByVal sVal As String)
    sDBStuff = sVal
End Property
Public Property Get Confirm() As Boolean
    Confirm = bConfirm
End Property

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            Me.Hide
        Case 1  'Ok
            bConfirm = True
            Me.Hide
        Case 2  'Save
            Save2File
    End Select

End Sub

Private Sub Form_Load()
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(2).Caption = MyLoadResString(k_Save)
    
    CenterForm Me

End Sub

Private Sub Form_Resize()
    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < 6030 Then Me.Width = 6030
            If Me.Height < 4245 Then Me.Height = 4245
            SizeControls
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fDet_MSG = Nothing
End Sub
Private Sub SizeControls()

    cmd(0).Top = Me.ScaleHeight - (cmd(0).Height + 50)
    cmd(1).Top = cmd(0).Top
    cmd(2).Top = cmd(0).Top
    cmd(2).Left = Me.ScaleWidth - cmd(0).Left - cmd(0).Width
    cmd(1).Left = cmd(2).Left - cmd(0).Width - 25
        
    tRes.Top = 25
    tRes.Height = cmd(0).Top - 50
    tRes.Width = Me.ScaleWidth - (2 * tRes.Left)
        
End Sub

Private Sub Save2File()

    Dim sFileName As String
    Dim sFilePath As String
    Dim lErr As Long, sErr As String
    Dim bCancel As Boolean
    Dim iDot As Integer
    Dim sCurDir As String
    
    sCurDir = CurDir$
    
    sFilePath = GetSpecialFolderLocation(Me.hwnd, CSIDL_PERSONAL)

    If Len(sFilePath) = 0 Then sFilePath = App.Path
    sFileName = "Reattach_" & sDBStuff & ".Sql"
   
    
    Dim obj As New cOpen
    obj.Title = MyLoadResString(k_Select_Destination_File)
    obj.Def_Dir() = sFilePath
    obj.Filter() = "*.Sql"
    obj.formHwnd() = Me.hwnd
    obj.f_Name() = sFileName
    
    obj.Do_Save
    
    lErr = obj.ErrCode()
    If lErr = 0 Then
        sFilePath = obj.FullPath()
        
        If Len(sFilePath) Then
            iDot = InStrRev(sFilePath, ".")
            If iDot Then sFilePath = Left$(sFilePath, iDot - 1)
            sFilePath = sFilePath & ".Sql"
        Else
            bCancel = True
        End If
    Else
        If lErr = 1 Then
            lErr = 0
        Else
            sErr = obj.ErrMsg()
        End If
    End If
    Set obj = Nothing
    
    On Local Error Resume Next
    ChDrive sCurDir
    ChDir sCurDir
    On Local Error GoTo 0
    
    Screen.MousePointer = vbHourglass
    If lErr = 0 And Not bCancel Then
        On Local Error Resume Next
        tRes.SaveFile sFilePath, rtfText
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub

