VERSION 5.00
Begin VB.Form fAddR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   Icon            =   "fAddR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   2760
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "fAddR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bCancel As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Private Sub cmd_Click(Index As Integer)

    bCancel = Index = 0
    Me.Hide
    
End Sub

Private Sub Form_Load()

    Me.Caption = MyLoadResString(k_Add_Role_Members)
    lbl.Caption = MyLoadResString(k_Select_Users_to_Add)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    
    CenterForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fAddR = Nothing
End Sub
