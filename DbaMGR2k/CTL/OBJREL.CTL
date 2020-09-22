VERSION 5.00
Begin VB.UserControl objRel 
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ScaleHeight     =   2775
   ScaleWidth      =   2910
   Begin VB.ListBox lstField 
      Height          =   2115
      IntegralHeight  =   0   'False
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmd 
      Caption         =   "..."
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "objRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bIsSecondary As Boolean
Private bShowMore As Boolean
Private sObjName As String
Private sObjReference As String
Private collKeys As Collection
Private collReference As Collection
Private mbMove As Boolean
Private iReferenced As Integer
Private bLoading As Boolean

Public Event Selected()
Public Event Click()
Public Event BeginMove(ByVal x As Integer, ByVal Y As Integer)
Public Event Moving(ByVal x As Integer, ByVal Y As Integer)
Public Event EndMoving(ByVal x As Integer, ByVal Y As Integer)
Public Event RequestMenu(ByVal x As Integer, ByVal Y As Integer)

Public Property Let ReferencedCtrl(ByVal iVal As Integer)
    iReferenced = iVal
End Property
Public Property Get ReferencedCtrl() As Integer
    ReferencedCtrl = iReferenced
End Property

Public Property Let IsSecondary(ByVal bVal As Boolean)
    bIsSecondary = bVal
    If bIsSecondary Then
        lblName(0).BackColor = vbInfoBackground
        lblName(0).ForeColor = vbInfoText
    End If
End Property
Public Property Let ToolTipToggle(ByVal sVal As String)
    cmd.ToolTipText = sVal
End Property
Public Property Let ToolTipCaption(ByVal sVal As String)
    lblName(0).ToolTipText = sVal
End Property
Public Property Let ToolTipFields(ByVal sVal As String)
    lstField.ToolTipText = sVal
End Property
Public Property Let ToolTipOwner(ByVal sVal As String)
    lblName(1).ToolTipText = sVal
End Property

Public Property Let ObjName(ByVal ObjOwner As String, ByVal ObjName As String)
    sObjName = Trim$(ObjName)
    sObjReference = Trim$(ObjOwner)
    lblName(0).Caption = sObjName
    lblName(1).Caption = sObjReference
End Property
Public Sub CheckThis(ByVal sItem As String)

    Dim i As Integer
    bLoading = True
    sItem = UCase$(sItem)
    With lstField
        For i = 0 To .ListCount - 1
            If UCase$(.List(i)) = sItem Then
                .Selected(i) = True
                Exit For
            End If
        Next
    End With
    bLoading = False
End Sub
Public Sub AddKeyReferences(ByVal sField As String)

    If collReference Is Nothing Then Set collReference = New Collection
    collReference.Add sField, sField
End Sub

Public Sub AddKeyColItem(ByVal sField As String)

    If collKeys Is Nothing Then Set collKeys = New Collection
    
    collKeys.Add sField, sField
End Sub
Public Sub AddItem(ByVal sField As String)
    
    lstField.AddItem sField
    
End Sub

Private Sub cmd_Click()
    RaiseEvent Selected
    bShowMore = Not bShowMore
    UserControl_Resize
    RaiseEvent Click
    
End Sub

Private Sub lblName_Click(Index As Integer)
    RaiseEvent Selected
End Sub


Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        mbMove = True
        RaiseEvent BeginMove(x, Y)
    ElseIf Button = vbRightButton Then
        mbMove = False
        RaiseEvent EndMoving(x, Y)
        RaiseEvent RequestMenu(x, Y)
    End If

End Sub

Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mbMove Then RaiseEvent Moving(x, Y)
End Sub

Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mbMove Then
        mbMove = False
        RaiseEvent EndMoving(x, Y)
    End If
End Sub

Private Sub lstField_GotFocus()
    RaiseEvent Selected
End Sub

Private Sub lstField_ItemCheck(Item As Integer)

    If Not bLoading Then
        lstField.Selected(Item) = Not lstField.Selected(Item)
    End If
    
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Width = 2970
End Sub

Private Sub UserControl_Resize()

    Static b As Boolean
    If Not b Then
        b = True
        If bShowMore Then
            If UserControl.Height < 2150 Then UserControl.Height = 2150
            lstField.Move 0, lblName(0).Height + lblName(1).Height + 5, UserControl.Width, UserControl.Height - lstField.Top
        Else
            If UserControl.Height <> lblName(0).Height Then UserControl.Height = lblName(0).Height
        End If
        
        If UserControl.Width <> 2970 Then UserControl.Width = 2970
        lblName(0).Width = UserControl.ScaleWidth - cmd.Width - 10
        lblName(1).Width = UserControl.ScaleWidth
        cmd.Left = lblName(0).Width + 10
        
        b = False
    End If
End Sub

Private Sub UserControl_Terminate()
    Set collKeys = Nothing
    Set collReference = Nothing
End Sub
Public Sub Clear()

    Dim i As Integer
    bLoading = True
    With lstField
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
    bLoading = False
End Sub
Public Property Get RelName() As String
    RelName = sObjReference
End Property
Public Property Get ReferencedCollection() As Collection
    Set ReferencedCollection = collReference
End Property
Public Property Get KeysCollection() As Collection
    Set KeysCollection = collKeys
End Property
Public Property Get TableName() As String
    TableName = sObjName
End Property

