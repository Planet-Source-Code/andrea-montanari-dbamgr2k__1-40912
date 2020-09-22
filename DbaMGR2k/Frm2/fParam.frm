VERSION 5.00
Begin VB.Form fParam 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6705
   ClipControls    =   0   'False
   Icon            =   "fParam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "remove"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   5400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "add"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   5400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox Lst 
      Height          =   1815
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "fParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(index As Integer)

    Dim i As Integer
    If cmd(index).Enabled Then
        Select Case index
            Case 0
                Unload Me
            Case 2
                'add
                objServer.Configuration.Parameters.Add Trim$(txt.Text)
                RefreshData
            Case 3
                'remove
                For i = 0 To Lst.ListCount - 1
                    If Lst.Selected(i) Then
                        objServer.Configuration.Parameters.Remove (i + 1)
                    End If
                Next
                RefreshData
        End Select
    End If
    
End Sub

Private Sub Form_Load()
    
    CenterForm Me
    Me.Caption = MyLoadResString(k_Param) & " - " & objServer.NetName
    
    lbl(0).Caption = MyLoadResString(k_Parameter)
    lbl(1).Caption = MyLoadResString(k_Existing_Parameter)
    
    lbl(2).Caption = ReplaceMsg(MyLoadResString(k_ParameterWarning), Array("1%"), Array(vbCrLf))
    cmd(0).Caption = MyLoadResString(k_Exit)
    
    cmd(2).Caption = MyLoadResString(k_Add)
    cmd(3).Caption = MyLoadResString(k_Remove)
    
    RefreshData

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set fParam = Nothing
End Sub

Private Sub Lst_Click()

    Dim i As Integer, bEnab As Boolean
    Dim sBuf As String
    If Lst.ListCount <> 0 Then
        For i = 0 To Lst.ListCount - 1
            bEnab = Lst.Selected(i)
            If bEnab Then Exit For
        Next
        sBuf = Lst.Text
    End If
    Lst.ToolTipText = sBuf
    cmd(3).Enabled = bEnab
    
End Sub

Private Sub txt_Change()

    cmd(2).Enabled = Len(Trim$(txt.Text)) <> 0
End Sub
Private Sub RefreshData()

    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    objServer.Configuration.Parameters.Refresh
    txt.Text = ""
    Lst.Clear
    For i = 1 To objServer.Configuration.Parameters.Count
        Lst.AddItem objServer.Configuration.Parameters(i)
    Next
    cmd(2).Enabled = False
    cmd(3).Enabled = False
    Screen.MousePointer = vbDefault

End Sub
