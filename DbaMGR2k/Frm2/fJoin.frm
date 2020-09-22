VERSION 5.00
Begin VB.Form fJoin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Join"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6375
   ClipControls    =   0   'False
   Icon            =   "fJoin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMain 
      Caption         =   "Frame1"
      Height          =   3135
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkProp 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   24
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   1455
         Index           =   2
         Left            =   3600
         TabIndex        =   25
         Top             =   1440
         Width           =   2535
         Begin VB.OptionButton opt 
            Caption         =   "ALL"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton opt 
            Caption         =   "WITH CUBE"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            Caption         =   "WITH ROLLUP"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.TextBox tTop 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "fJoin.frx":000C
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame fra 
         Height          =   1815
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   3135
         Begin VB.CheckBox chkProp 
            Caption         =   "Check1"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CheckBox chkProp 
            Caption         =   "Check1"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkProp 
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   2655
         End
         Begin VB.CheckBox chkProp 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   2655
         End
         Begin VB.CheckBox chkProp 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Label lblT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "table"
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lbl 
         Caption         =   "Top:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "viewName"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Frame1"
      Height          =   3135
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cbo 
         Height          =   315
         ItemData        =   "fJoin.frx":000E
         Left            =   3360
         List            =   "fJoin.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame fra 
         Height          =   1215
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   1560
         Width           =   5175
         Begin VB.CheckBox chkJoin 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   5
            Top             =   840
            Width           =   3495
         End
         Begin VB.CheckBox chkJoin 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   4
            Top             =   360
            Width           =   3375
         End
         Begin VB.Image imgJoin 
            Height          =   240
            Left            =   240
            Picture         =   "fJoin.frx":003D
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Label lbl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "table"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "table"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "field"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "field"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Height          =   360
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "fJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oJoin As clsJoin
Private oView As cViewPrp
Private bCancel As Boolean
Private bLoading As Boolean
Private bIsJoin As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property
Public Property Let LetJoin(ByVal cJoin As clsJoin)
    
    Dim iJoinType As JoinT
    Set oJoin = New clsJoin
    
    bIsJoin = True
    
    fraMain(0).Visible = False
    fraMain(1).Caption = MyLoadResString(k_Join_Line)
    fraMain(1).Visible = True
    fraMain(1).ZOrder
    lbl(0).Caption = MyLoadResString(k_Table)
    lbl(1).Caption = MyLoadResString(k_Column_Name)
    fra(0).Caption = MyLoadResString(k_Include_Row)
    
    With oJoin
        .Id() = cJoin.Id()
        .BaseTbl = cJoin.BaseTbl()
        .Field() = cJoin.Field()
        
        .RelatedTbl() = cJoin.RelatedTbl()
        .RelatedField() = cJoin.RelatedField()
        
        .iBaseControl() = cJoin.iBaseControl()
        .iControl() = cJoin.iControl()
        .iRefControl() = cJoin.iRefControl()
        
        .JoinOperator() = cJoin.JoinOperator()
        .JoinType() = cJoin.JoinType()
        
        lblT(0).Caption = .BaseTbl()
        lblT(1).Caption = .RelatedTbl()
        lblT(2).Caption = .Field()
        lblT(3).Caption = .RelatedField()
        cbo.ListIndex = .JoinOperator()
        iJoinType = .JoinType
        imgJoin.Picture = LoadResPicture(iJoinType, vbResIcon)
        bLoading = True
        chkJoin(0).Caption = MyLoadResString(k_All_Rows_from) & .BaseTbl()
        chkJoin(1).Caption = MyLoadResString(k_All_Rows_from) & .RelatedTbl()
        If iJoinType = jLeft Then
            chkJoin(0).Value = 1
        ElseIf iJoinType = jRight Then
            chkJoin(1).Value = 1
        ElseIf iJoinType = jFull Then
            chkJoin(0).Value = 1
            chkJoin(1).Value = 1
        End If
        
        bLoading = False
    End With
    
End Property
Public Sub GetJoin(cJoin As clsJoin)

    With cJoin
        .JoinOperator() = oJoin.JoinOperator()
        .JoinType() = oJoin.JoinType()
    End With
    
End Sub
Public Sub GetView(coView As cViewPrp)

    Dim i As Integer
    With coView
        .TopRec() = Val(tTop.Text)
        .bOption(0) = chkProp(0).Value
        .bOption(1) = chkProp(1).Value
        .bOption(2) = chkProp(2).Value
        .bOption(3) = chkProp(4).Value
        .bOption(4) = chkProp(5).Value
        .Extention() = chkProp(3).Enabled
        .ExtOpt() = -1
        If chkProp(3).Value Then
            For i = 0 To 2
                If opt(i).Value Then
                    .ExtOpt() = i
                End If
            Next
        End If
    End With
    
End Sub

Private Sub chkJoin_Click(Index As Integer)
    
    If Not bLoading Then
        With oJoin
            If chkJoin(0).Value = 0 And chkJoin(1).Value = 0 Then
                .JoinType = jInner
            ElseIf chkJoin(0).Value = 1 And chkJoin(1).Value = 0 Then
                .JoinType = jLeft
            ElseIf chkJoin(0).Value = 0 And chkJoin(1).Value = 1 Then
                .JoinType = jRight
            Else
                .JoinType = jFull
            End If
            imgJoin.Picture = LoadResPicture(.JoinType(), vbResIcon)
        End With
    End If
End Sub

Private Sub chkProp_Click(Index As Integer)

    If Index = 3 Then
        fra(2).Enabled = chkProp(Index).Value
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
        bCancel = True
        Me.Hide
    Else
        If bIsJoin Then oJoin.JoinOperator = cbo.ListIndex
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Caption = MyLoadResString(k_Property)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Frm_Ok)
    lbl(2).Caption = MyLoadResString(k_View_Name)
    
    imgJoin.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set oJoin = Nothing
    Set oView = Nothing
    Set fJoin = Nothing
End Sub
Public Sub ViewOptions(cView As cViewPrp)

    Set oView = New cViewPrp
    
    oView.oName() = cView.oName()
    oView.TopRec() = cView.TopRec()
    oView.bOption(0) = cView.bOption(0)
    oView.bOption(1) = cView.bOption(1)
    oView.bOption(2) = cView.bOption(2)
    oView.bOption(3) = cView.bOption(3)
    oView.bOption(4) = cView.bOption(4)

    oView.bOptionEnab(0) = cView.bOptionEnab(0)
    oView.bOptionEnab(1) = cView.bOptionEnab(1)
    
    'oView.bOptionEnab(3) = cView.bOptionEnab(3)
    oView.isGroup() = cView.isGroup()
    oView.Extention() = cView.Extention()
    oView.ExtOpt() = cView.ExtOpt()
    
    lblT(4).Caption = oView.oName()
    tTop.Text = oView.TopRec()
    chkProp(0).Value = Abs(oView.bOption(0))
    chkProp(1).Value = Abs(oView.bOption(1))
    
    chkProp(2).Value = Abs(oView.bOption(2))
    chkProp(4).Value = Abs(oView.bOption(3))
    chkProp(5).Value = Abs(oView.bOption(4))
    
    chkProp(0).Enabled = oView.bOptionEnab(0)
    chkProp(1).Enabled = oView.bOptionEnab(1)
    
    
    chkProp(3).Value = Abs(oView.Extention())
    chkProp(3).Enabled = oView.isGroup()
    If oView.isGroup() Then
        If oView.ExtOpt() <> -1 Then opt(oView.ExtOpt()).Value = True
    End If
    
    fraMain(1).Visible = False
    'fraMain(0).Caption = MyLoadResString(k_Join_Line)
    fraMain(0).Caption = ""
    fraMain(0).Visible = True
    fraMain(0).ZOrder
    lbl(0).Caption = MyLoadResString(k_Table)
    lbl(1).Caption = MyLoadResString(k_Column_Name)
    fra(1).Caption = MyLoadResString(k_Options)
    'chkProp(0).Caption = MyLoadResString(k_All_Columns  )
    chkProp(0).Caption = MyLoadResString(k_DISTINCT_Value)
    chkProp(1).Caption = MyLoadResString(k_View_Encryption)
    chkProp(2).Caption = MyLoadResString(k_Func_Param_IsSchemaBound)
    chkProp(4).Caption = MyLoadResString(k_Explicit_AnsiNull)
    chkProp(5).Caption = MyLoadResString(k_Explicit_QuotedIdentifier)
    
    chkProp(3).Caption = MyLoadResString(k_GROUP_BY_Extensions)

End Sub

Private Sub tTop_KeyDown(KeyCode As Integer, Shift As Integer)
    Digit_KDown tTop, KeyCode
End Sub

Private Sub tTop_KeyPress(KeyAscii As Integer)
    Digit_KPress tTop, KeyAscii
End Sub
