VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TblView 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ScaleHeight     =   3375
   ScaleWidth      =   4695
   Begin VB.CommandButton cmd 
      Caption         =   "close"
      Height          =   360
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   2400
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.ListBox lst 
      Height          =   2400
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmd 
      Caption         =   "add"
      Height          =   360
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "t"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "v"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      Height          =   165
      Left            =   4440
      Picture         =   "TblView.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H80000003&
      Caption         =   "Label1"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "TblView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum LstType
    isTbl = 0
    isView = 1
End Enum
Private mbMove As Boolean

Public Event ItemSelected(ByVal iItem As Integer, ByVal iLst As LstType)
Public Event QueryUnload(bCancel As Boolean)
Public Event BeginMove(ByVal X As Integer, ByVal Y As Integer)
Public Event Moving(ByVal X As Integer, ByVal Y As Integer)
Public Event EndMoving(ByVal X As Integer, ByVal Y As Integer)

Public Property Let CurrentItem(ByVal iLst As LstType)
    Set tStrip.SelectedItem = tStrip.Tabs(iLst + 1)
End Property
Public Property Get CurrentItem() As LstType
    CurrentItem = tStrip.SelectedItem.Index - 1
End Property

Public Sub Clear(ByVal iLst As LstType)
    lst(iLst).Clear
End Sub
Public Property Get ListCount(ByVal iLst As LstType) As Integer
    ListCount = lst(iLst).ListCount
End Property
Public Function GetItem(ByVal iElem As Integer, ByVal iLst As LstType) As String
    If lst(iLst).ListCount >= iElem Then GetItem = lst(iLst).List(iElem)
End Function

Public Sub AddItem(ByVal sVal As String, ByVal iLst As LstType)
    lst(iLst).AddItem sVal
End Sub

Public Sub InitString(ByVal sStrip0 As String, ByVal sStrip1 As String, ByVal sCmd0 As String, ByVal sCmd1 As String)
    tStrip.Tabs(1).Caption = sStrip0
    tStrip.Tabs(2).Caption = sStrip1
    cmd(0).Caption = sCmd0
    cmd(1).Caption = sCmd1
End Sub
Public Property Let Caption(ByVal sVal As String)
    lblCaption.Caption = sVal
End Property
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property


Private Sub cmd_Click(Index As Integer)
    Dim i As Integer
    If cmd(Index).Enabled Then
        If Index = 0 Then
            i = tStrip.SelectedItem.Index - 1
            If lst(i).ListCount And lst(i).ListIndex >= 0 Then lst_DblClick i
        Else
            img_Click
        End If
    End If

End Sub

Private Sub img_Click()
    Dim bCancel As Boolean
    RaiseEvent QueryUnload(bCancel)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mbMove = True
        RaiseEvent BeginMove(X, Y)
    End If
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbMove Then RaiseEvent Moving(X, Y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbMove Then
        mbMove = False
        RaiseEvent EndMoving(X, Y)
    End If

End Sub

Private Sub lst_DblClick(Index As Integer)

    If lst(Index).ListCount Then
        RaiseEvent ItemSelected(lst(Index).ListIndex, Index)
    End If
    
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    i = tStrip.SelectedItem.Index - 1
    lst(i).Visible = True
    lst(1 - Abs(i)).Visible = False
    lst(i).ZOrder
    cmd(0).Enabled = lst(i).ListCount
    
End Sub

Private Sub UserControl_EnterFocus()
    lblCaption.BackColor = vbActiveTitleBar
    lblCaption.ForeColor = vbActiveTitleBarText
    tStrip_Click
End Sub

Private Sub UserControl_ExitFocus()
    lblCaption.BackColor = vbInactiveTitleBar
    lblCaption.ForeColor = vbInactiveTitleBarText
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()

    Static b As Boolean
    If Not b Then
        b = True
        If UserControl.ScaleWidth <> 4695 Then UserControl.Width = 4740
        If UserControl.ScaleHeight <> 3375 Then UserControl.Height = 3420
                
        lblCaption.Move 0, 0, UserControl.ScaleWidth
        img.Left = UserControl.ScaleWidth - img.Width
        img.ZOrder
        b = False
    End If

End Sub
