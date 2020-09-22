VERSION 5.00
Begin VB.Form fPwd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conferma Nuova Password"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   Icon            =   "fPwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFileG 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Annulla"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox tLog 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblFileG 
      Caption         =   "&FileGroup:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "fPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sStuff As String
Private bConfirm As Boolean
Private ActMode As optAction
Public Property Let Action(ByVal Val As optAction)
    ActMode = Val
    If ActMode = act_Tbl Then
        tLog.PasswordChar = ""
        If cboFileG.ListCount > 1 Then
            Me.Height = 2025
            cboFileG.Visible = True
            lblFileG.Visible = True
            cmd(0).Top = cboFileG.Top + cboFileG.Height + 40
            cmd(1).Top = cmd(0).Top
        End If
    ElseIf ActMode = act_View Then
        tLog.PasswordChar = ""
    ElseIf ActMode = act_Login Then
        tLog.PasswordChar = "*"
    End If
End Property
Public Property Get Stuff() As String
    Stuff = sStuff
End Property
Public Property Let collFileGroups(collFileGroups As Collection)

    Dim i As Integer
    cboFileG.Clear
    
    For i = 1 To collFileGroups.Count
        cboFileG.AddItem collFileGroups.Item(i)
    Next
    cboFileG.ListIndex = 0

End Property

Public Property Get FileGroup() As String
    FileGroup = cboFileG.Text
End Property

Public Property Let Stuff(ByVal sVal As String)
    sStuff = sVal
    'If ActMode = act_Tbl Or ActMode = act_ATTACH Then tLog.Text = sVal
    If ActMode = act_Tbl Then tLog.Text = sVal
End Property
Public Property Get Confirm() As Boolean
    Confirm = bConfirm
End Property

Private Sub cmd_Click(Index As Integer)
    If Index Then
        If ActMode = act_Login Then
            If Trim$(tLog.Text) <> sStuff Then
                MsgBox MyLoadResString(k_Password_confirmation_aborted), vbInformation Or vbOKOnly, App.EXEName
            Else
                bConfirm = True
                Me.Hide
            End If
        ElseIf ActMode = act_Tbl Or ActMode = act_View Then
            sStuff = Trim$(tLog.Text)
            If Len(sStuff) <> 0 Then
                bConfirm = True
                Me.Hide
            End If
        End If
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = MyLoadResString(k_Confirm_New_Password)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fPwd = Nothing
End Sub

Private Sub tLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then KeyCode = 0
End Sub

Private Sub tLog_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
        Case vbKeyEscape
            KeyAscii = 0
        Case vbKeyBACK
        Case 43, 45, 48 To 57, 65 To 90, 95, 97 To 122
            OverStrike tLog
        Case Else
            KeyAscii = 0
    End Select
End Sub
