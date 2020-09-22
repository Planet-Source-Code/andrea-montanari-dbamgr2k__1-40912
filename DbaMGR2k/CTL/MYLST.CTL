VERSION 5.00
Begin VB.UserControl myLst 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   ScaleHeight     =   1455
   ScaleWidth      =   2250
   Begin VB.ListBox lst 
      DragIcon        =   "myLst.ctx":0000
      Height          =   1185
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image imgGroup 
      Height          =   240
      Left            =   360
      Picture         =   "myLst.ctx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgT 
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "myLst.ctx":0454
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgT 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "myLst.ctx":059E
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgT 
      Height          =   240
      Index           =   1
      Left            =   120
      Picture         =   "myLst.ctx":06E8
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgType 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   165
      Left            =   1440
      Picture         =   "myLst.ctx":0832
      Stretch         =   -1  'True
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000003&
      Caption         =   "Table Name"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "myLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private iType As Integer
Private sTbName As String
Private iDragInit As Integer
Private bIsGrouped As Boolean
Private mbMove As Boolean
'Dichiarazioni di eventi:
Public Event ItemCheck(Item As Integer)
Public Event QueryUnload(bCancel As Boolean)
Public Event BeginMove(ByVal X As Integer, ByVal Y As Integer)
Public Event Moving(ByVal X As Integer, ByVal Y As Integer)
Public Event EndMoving(ByVal X As Integer, ByVal Y As Integer)

Public Event BeginRefer()
Public Event DropRefer(iItem As Integer)
Public Enum CtrlType
    None = -1
    Tbl = 0
    View = 1
End Enum
Public Property Get DragInit() As Integer
    DragInit = iDragInit
End Property
Public Sub TopIndex(Optional ByVal iIndex As Integer = 0)
    lst.TopIndex = iIndex
End Sub
Public Sub AddItem(ByVal sVal As String, Optional ByVal bSelected As Boolean = False, Optional ByVal iIndex As Integer = -1, Optional ByVal bIsKey As Boolean = False, Optional ByVal lItemData As Long = -1)

    With lst
        If bIsKey Then sVal = "(*) " & sVal
        If iIndex >= 0 Then
            .AddItem sVal, iIndex
        Else
            .AddItem sVal
        End If
        .Selected(lst.NewIndex) = bSelected
        If lItemData <> -1 Then .ItemData(.NewIndex) = lItemData
    End With
End Sub
Public Function GetItem(ByVal iIndex As Integer) As String
    Dim sBuf As String
    Dim iPos As Integer
    If iIndex <= lst.ListCount Then
        sBuf = lst.List(iIndex)
        iPos = InStr(sBuf, ")")
        If iPos > 0 Then sBuf = Mid$(sBuf, iPos + 2)
        GetItem = sBuf
    End If
End Function
Public Property Get IsSelected(ByVal iIndex As Integer) As Boolean
    If iIndex <= lst.ListCount Then IsSelected = lst.Selected(iIndex)
End Property
Public Property Let IsSelected(ByVal iIndex As Integer, ByVal bval As Boolean)
    If iIndex <= lst.ListCount Then lst.Selected(iIndex) = bval
End Property

Private Sub img_Click()
    Dim bCancel As Boolean
    RaiseEvent QueryUnload(bCancel)
End Sub
Public Property Get ListCount() As Integer
    ListCount = lst.ListCount
End Property


Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mbMove = True
        RaiseEvent BeginMove(X, Y)
    End If
End Sub
Public Sub Referred(ByVal X As Long, ByVal Y As Long)

    Dim listit As Integer
    listit = Fix(((Y - lst.Top) / 210)) + lst.TopIndex
    
    On Error Resume Next
    Debug.Print Err.Description
    Debug.Print lst.ListCount
    lst.ListIndex = listit
    If Err.Number = 0 And lst.ListIndex > 0 Then RaiseEvent DropRefer(lst.ListIndex)

End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mbMove Then RaiseEvent Moving(X, Y)

End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbMove Then
        mbMove = False
        RaiseEvent EndMoving(X, Y)
    End If
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    RaiseEvent ItemCheck(Item)
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim listit As Integer
    If Button = vbRightButton Then
        iDragInit = -1
        listit = Fix(((Y + lst.Top) / 210)) - 1 + lst.TopIndex
        On Error Resume Next
        lst.ListIndex = listit
        If listit > 0 Then
            iDragInit = listit
            RaiseEvent BeginRefer
        End If
    End If
    
End Sub

Private Sub UserControl_EnterFocus()
    lbl.BackColor = vbActiveTitleBar
    lbl.ForeColor = vbActiveTitleBarText
End Sub

Private Sub UserControl_ExitFocus()
    lbl.BackColor = vbInactiveTitleBar
    lbl.ForeColor = vbInactiveTitleBarText
End Sub

Private Sub UserControl_Initialize()
    iType = -1
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()

    Static b As Boolean
    
    If Not b Then
        b = True
        If UserControl.Width < 2500 Then UserControl.Width = 2510
        If UserControl.Height < 1455 Then UserControl.Height = 1455
        
        lst.Move 5, lbl.Height + 5, UserControl.ScaleWidth - (2 * lst.Left), UserControl.ScaleHeight - lbl.Height - 5
        UserControl.Height = lst.Height + lbl.Height + 80
        imgType.Height = lbl.Height
        'img.Height = lbl.Height
        
        imgGroup.Visible = bIsGrouped
        If iType <> -1 Then
            imgType.Picture = imgT(iType).Picture
            imgType.Left = 0
            imgType.Visible = True
            imgGroup.Height = lbl.Height
            imgGroup.Left = imgType.Width
            
            'lbl.Move imgType.Width + IIf(bIsGrouped, imgType.Width, 0), 0, UserControl.ScaleWidth - imgType.Width - (img.Width + IIf(bIsGrouped, imgType.Width, 0))
            lbl.Move imgType.Width + IIf(bIsGrouped, imgType.Width, 0), 0, UserControl.ScaleWidth - imgType.Width - (IIf(bIsGrouped, imgType.Width, 0))
            
        Else
            'lbl.Move 0, 0, UserControl.ScaleWidth
            lbl.Move 0, 0, UserControl.ScaleWidth '- img.Width
            imgType.Visible = False
        End If
        img.Left = UserControl.ScaleWidth - img.Width
        img.ZOrder
        b = False
    End If
    
End Sub
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=lbl,lbl,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lbl.Caption() = New_Caption
    sTbName = New_Caption
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=lst,lst,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Cancella il contenuto di un controllo o degli Appunti di sistema."
    lst.Clear
End Sub

Public Property Get DragPic() As Object
    Set DragPic = lst.DragIcon
End Property
Public Property Let TType(ByVal iVal As CtrlType)
    iType = iVal
    PropertyChanged "iType"
    UserControl_Resize
End Property
Public Property Get TType() As CtrlType
    TType = iType
End Property

Public Property Let ItemData(ByVal iItem As Integer, ByVal lVal As Long)
    lst.ItemData(iItem) = lVal
End Property
Public Property Get ItemData(ByVal iItem As Integer) As Long
    ItemData = lst.ItemData(iItem)
End Property
Public Property Let IsGrouped(ByVal bval As Boolean)
    bIsGrouped = bval
    UserControl_Resize
End Property
Public Property Get IsGrouped() As Boolean
    IsGrouped = bIsGrouped
End Property
