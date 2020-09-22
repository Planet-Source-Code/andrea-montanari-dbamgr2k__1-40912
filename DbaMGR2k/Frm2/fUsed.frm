VERSION 5.00
Begin VB.Form fUsed 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "used"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4740
   ClipControls    =   0   'False
   Icon            =   "fUsed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "fUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    CenterForm Me
    Me.Caption = Replace(MyLoadResString(k_Where_Used), "&", "")
    lbl.Caption = MyLoadResString(k_the_following_columns_use_this_UDT)
    cmd(0).Caption = MyLoadResString(k_Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fUsed = Nothing
End Sub
