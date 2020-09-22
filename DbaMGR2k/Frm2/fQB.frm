VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fQB 
   Caption         =   "-"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   ClipControls    =   0   'False
   Icon            =   "fQB.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin DbaMGR2k.TblView TblView 
      Height          =   3420
      Left            =   1200
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   6033
   End
   Begin VB.TextBox tFlex 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tSql 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flexG 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.CheckBox chkDBO 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   0
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.PictureBox picC 
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   360
      Width           =   7335
      Begin VB.PictureBox picTBL 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawStyle       =   1  'Dash
         Height          =   10000
         Left            =   0
         ScaleHeight     =   10005
         ScaleWidth      =   10005
         TabIndex        =   4
         Top             =   0
         Width           =   10000
         Begin DbaMGR2k.myLst myLst 
            Height          =   1290
            Index           =   0
            Left            =   600
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   2275
         End
         Begin VB.Image imgJoin 
            Height          =   240
            Index           =   0
            Left            =   0
            Picture         =   "fQB.frx":014A
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.VScrollBar VScr 
      Height          =   1215
      Left            =   6480
      Max             =   32000
      SmallChange     =   10
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.HScrollBar HScr 
      Height          =   255
      Left            =   0
      Max             =   32000
      SmallChange     =   10
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flexG 
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0294
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":06E8
            Key             =   "property"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0844
            Key             =   "diag"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":09A0
            Key             =   "grid"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0AFC
            Key             =   "sql"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0C58
            Key             =   "result"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0DB4
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":0F10
            Key             =   "testSql"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":106C
            Key             =   "group"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":11C8
            Key             =   "erase"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":14E4
            Key             =   "addTbl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":1640
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQB.frx":179C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWidth As Long
Private m_lHeight As Long

Private Const kAlias = 2
Private Const kShow = 4
Private Const kGroup = 5
Private Const kCondition = 6
Private Const kOr = 7

Private Type sFieldUdt
    iRow As Integer
    iCtrl As Integer
    sField As String
    sAlias As String
    sTb As String
    sTbAlias As String
    bShow As Boolean
    iGroup As Integer
    sCriteria As String
    sOr As String
    sOwner As String
End Type
Private iOldX As Integer, iOldY As Integer

Private bFlexLoad As Boolean
Private miRow As Integer
Private miCol As Integer
Private m_blnAnnullaMod As Boolean
Private mbIsGroup As Boolean
Private mbOldView As Boolean

Private sDbName As String
Private sStuff As String
Private oDb As Database
Private iCurDrag As Integer
Private iIncY As Integer, iIncX As Integer
Private bVisible(3) As Boolean  'Diag,Grid,Sql,Result
Private bTbl As Boolean
Private collItem As Collection
Private collJoin As Collection

Private arrField() As sFieldUdt
Private arrGroup() As String

Private cViewP As cViewPrp

Public Property Let ViewStuff(ByVal oView As SQLDMO.View)
    mbOldView = True
    cViewP.oName = oView.Name
End Property
Public Property Let DBStuff(ByVal sDB As String, ByVal sView As String)
    sDbName = sDB
    sStuff = sView
    
    cViewP.oName() = sStuff
    cViewP.bOptionEnab(0) = True
    cViewP.bOptionEnab(1) = True
    
    cViewP.isGroup() = 0
    cViewP.Extention() = False
    cViewP.ExtOpt() = -1
    
    Set oDb = objServer.Databases(sDbName)
    Me.Caption = MyLoadResString(k_New_View) & " DB: [" & sStuff & "]"
    
    Dim oObj As Object
    Screen.MousePointer = vbHourglass
    With TblView
        .Caption = MyLoadResString(k_Tables) & "/" & MyLoadResString(k_Views)
        .InitString MyLoadResString(k_Tables), MyLoadResString(k_Views), MyLoadResString(k_Add_All), MyLoadResString(k_Close)
        .Clear isTbl
        .Clear isView
        For Each oObj In oDb.Tables
            If Not oObj.SystemObject Then .AddItem "[" & oObj.Name & "]", isTbl
        Next
        For Each oObj In oDb.Views
            If Not oObj.SystemObject Then .AddItem "[" & oObj.Name & "]", isView
        Next
        .Visible = False
    End With
    Set oObj = Nothing
    Screen.MousePointer = vbDefault
    
End Property

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub flexG_DblClick(Index As Integer)
    
    Dim bCombo As Boolean
    Dim iCol As Integer
    If Index = 0 And (flexG(Index).MouseRow > 0 And flexG(Index).MouseCol > 0) Then
        If Len(arrField(miRow - 1).sField) = 0 Then Exit Sub
        
        iCol = flexG(Index).Col
        If iCol = kGroup Then bCombo = True
    
        With flexG(Index)
            If bCombo Then
                LoadCbo iCol
                cbo.ListIndex = 0
                If iCol = kGroup Then
                    cbo.ListIndex = arrField(miRow - 1).iGroup
                Else
                    cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                End If
                cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                miCol = .Col
                cbo.Visible = True
                cbo.SetFocus
            Else
                FlexEdit vbKeySpace
            End If
        End With
    End If

End Sub

Private Sub flexG_GotFocus(Index As Integer)
    
    TestActiveControl
    m_blnAnnullaMod = True
End Sub

Private Sub flexG_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Index = 0 And KeyCode = vbKeyF2 Then
        KeyCode = 0
        If flexG(Index).Col = kGroup Then
            flexG_DblClick Index
        Else
            FlexEdit 0
        End If
    End If

End Sub

Private Sub flexG_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    Dim bCombo As Boolean

    If Index = 0 Then
        bCombo = flexG(Index).Col = kGroup
        
        If bCombo Then
            KeyAscii = 0
            flexG_DblClick Index
        Else
            iCode = KeyAscii
            KeyAscii = 0
            FlexEdit iCode
        End If
    End If
    
End Sub

Private Sub flexG_RowColChange(Index As Integer)

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean

    If Not bFlexLoad And Index = 0 Then
        If Not b Then
            b = True
            TestActiveControl
            With flexG(Index)
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set flexG(Index).CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
'                    If mbIsGroup Then
'                        .Col = kGroup
'                        Set flexG(index).CellPicture = Nothing
'                    End If

                    .Row = iRowP
                    .Col = 0
                    Set flexG(Index).CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1
                    
'                    If mbIsGroup Then
'                        .Col = kGroup
'                        Set flexG(index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
'                        .CellPictureAlignment = flexAlignRightTop
'                    End If
                    .Col = iCol
                    miRow = iRowP
                End If
                
            End With
            b = False
        End If
    End If

End Sub

Private Sub flexG_Scroll(Index As Integer)
    cbo.Visible = False
    tFlex.Visible = False
End Sub

Private Sub flexG_SelChange(Index As Integer)
    flexG(Index).Row = flexG(Index).RowSel
    flexG(Index).Col = flexG(Index).ColSel
End Sub

Private Sub Form_Activate()
    Static b As Boolean
    Dim i As Integer
    If Not b Then
        b = True
        For i = 4 To 7
            SetState i, bVisible(i - 4)
        Next
    End If
    
End Sub

Private Sub Form_Initialize()
    Set collItem = New Collection
    Set collJoin = New Collection
    Set cViewP = New cViewPrp
    
    cViewP.bOption(3) = True
    cViewP.bOption(4) = True
End Sub

Private Sub Form_Load()
    
    m_lWidth = Me.Width
    m_lHeight = Me.Height
    
    flexG(0).RowHeightMin = cbo.Height
    
    ReDim arrField(0) As sFieldUdt
    ReDim arrGroup(11) As String
    arrGroup(0) = ""
    arrGroup(1) = "GROUP BY"
    arrGroup(2) = "AVG"
    arrGroup(3) = "COUNT"
    arrGroup(4) = "MIN"
    arrGroup(5) = "MAX"
    arrGroup(6) = "SUM"
    arrGroup(7) = "AVG DISTINCT"
    arrGroup(8) = "COUNT DISTINCT"
    arrGroup(9) = "MIN DISTINCT"
    arrGroup(10) = "MAX DISTINCT"
    arrGroup(11) = "SUM DISTINCT"
    
    
    Dim sBuf As String
'    For i = 0 To UBound(bVisible)
'        bVisible(i) = True
'    Next
    bVisible(0) = True
    bVisible(2) = True
    tBar.Buttons(1).ToolTipText = MyLoadResString(k_Save_View)
    tBar.Buttons(2).ToolTipText = MyLoadResString(k_Property)
    
    sBuf = MyLoadResString(k_Show_Hide)
    tBar.Buttons(4).ToolTipText = sBuf & MyLoadResString(k_Diagram_Pane)
    tBar.Buttons(5).ToolTipText = sBuf & MyLoadResString(k_Grid_Pane)
    tBar.Buttons(6).ToolTipText = sBuf & MyLoadResString(k_Sql_Pane)
    tBar.Buttons(7).ToolTipText = sBuf & MyLoadResString(k_Result_Pane)
    
    tBar.Buttons(9).ToolTipText = MyLoadResString(k_Run)
    tBar.Buttons(10).ToolTipText = MyLoadResString(k_Verify_Sql)
    tBar.Buttons(11).ToolTipText = MyLoadResString(k_Use_Group_By)
    tBar.Buttons(12).ToolTipText = MyLoadResString(k_Clear_Result)
    
    tBar.Buttons(14).ToolTipText = MyLoadResString(k_Show_Table_View_List)
    tBar.Buttons(15).ToolTipText = MyLoadResString(k_Automatic_Arrange_Tables)
    tBar.Buttons(16).ToolTipText = MyLoadResString(k_Fit_Grid)
    
    chkDBO.Caption = MyLoadResString(k_OBJ_Create_As_DBO)
    chkDBO.ToolTipText = MyLoadResString(k_OBJ_Create_As_DBO_tolTip)
    
    LoadFlexG
    iIncY = 1
    iIncX = 1
    CenterForm Me
    
End Sub

Private Sub Form_Resize()
    Dim lH As Long, iItem As Integer, i As Integer
    Dim lfHeight As Long
    Static b As Boolean
    
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            
            
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            
            For i = 0 To UBound(bVisible)
                If bVisible(i) Then iItem = iItem + 1
            Next
            If iItem > 0 Then lfHeight = ((Me.ScaleHeight - tBar.Height) \ iItem)
    
                
            picC.Visible = bVisible(0)
            If bVisible(0) Then
                picC.Move 0, tBar.Height, Me.ScaleWidth - VScr.Width - 10, lfHeight - HScr.Height
                HScr.Move 0, tBar.Height + picC.Height, Me.ScaleWidth
                VScr.Move Me.ScaleWidth - VScr.Width, tBar.Height
                VScr.Height = picC.Height + 10
                If picTBL.Width < picC.ScaleWidth + 1000 Then picTBL.Width = picC.ScaleWidth + 1000
                If picTBL.Height < picC.ScaleHeight + 1000 Then picTBL.Height = picC.ScaleHeight + 1000
            End If
    
            HScr.Visible = bVisible(0)
            VScr.Visible = bVisible(0)
            HScr.Max = picTBL.Width - picC.ScaleWidth
            VScr.Max = picTBL.Width - picC.ScaleHeight
            
            lH = IIf(bVisible(0), HScr.Top - (HScr.Height \ 2), 0) + tBar.Height
            flexG(0).Move 0, lH, Me.ScaleWidth, lfHeight
            flexG(0).Visible = bVisible(1)
            
            lH = 0
            If bVisible(0) And bVisible(1) Then
                lH = flexG(0).Height + flexG(0).Top
            ElseIf bVisible(0) Then
                lH = HScr.Height + HScr.Top
            ElseIf bVisible(1) Then
                lH = flexG(0).Height + flexG(0).Top
            Else
                lH = tBar.Height
            End If
            tSql.Move 0, lH, Me.ScaleWidth, lfHeight
            tSql.Visible = bVisible(2)
            
            lH = 0
            For i = 0 To 2
                If bVisible(i) Then
                    Select Case i
                        Case 0
                            lH = HScr.Height + HScr.Top
                        Case 1
                            lH = flexG(0).Height + flexG(0).Top
                        Case 2
                            lH = tSql.Height + tSql.Top
                    End Select
                End If
            Next
            If lH = 0 Then lH = tBar.Height
    
            flexG(1).Move 0, lH, Me.ScaleWidth, lfHeight
            flexG(1).Visible = bVisible(3)
            With TblView
                If .Visible Then
                    If .Left + .Width > Me.ScaleWidth Or .Top + .Height > Me.ScaleHeight Then .Move Me.ScaleWidth - .Width - 100, Me.ScaleHeight - .Height - 100
                End If
            End With
            SettaLargFlex flexG(0), 1, 1300
            tBar.Align = vbAlignTop
            tBar.Top = 0
            chkDBO.Left = tBar.Width - chkDBO.Width - k_MargineFRM
            chkDBO.Refresh
            b = False
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set collJoin = Nothing
    Set collItem = Nothing
    Set oDb = Nothing
    Set fQB = Nothing
End Sub

Private Sub imgJoin_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer
    Dim sBuf As String
        
    If Button = vbRightButton Then
        sBuf = " Join"
        Dim M As New cPUpMenu
                                    
        M.Add MyLoadResString(k_Remove) & sBuf, 1
        M.Add "", True, 2
        M.Add MyLoadResString(k_Property) & sBuf, 3

        Call GetCursorPos(P)
        i = M.Popup(Me.hWnd, P.x, P.y)
        M.Clear
        Set M = Nothing
        If i = 1 Then
            DropJoin Index
        ElseIf i = 3 Then
            ShowJoin Index
        End If
    End If

End Sub
Private Sub DropJoin(ByVal iCtrl As Integer)

    Dim i As Integer

    Dim bDone As Boolean
    Screen.MousePointer = vbHourglass
    
    For i = 1 To collJoin.Count
        If iCtrl = collJoin(i).iControl() Then
            Unload imgJoin(iCtrl)
            collJoin.Remove i
            bDone = True
            Exit For
        End If
    Next
    Screen.MousePointer = vbDefault
    If bDone Then DrawJoin
    
End Sub
Private Sub ShowJoin(ByVal iCtrl As Integer)

    Dim i As Integer
    Dim bDone As Boolean, bCancel As Boolean
    Dim iNewJoinT As JoinT
        
    Screen.MousePointer = vbHourglass
    Dim oJoin As clsJoin
    
    For i = 1 To collJoin.Count
        If iCtrl = collJoin(i).iControl() Then
            Set oJoin = collJoin(i)
            bDone = True
            Exit For
        End If
    Next
    
    Screen.MousePointer = vbDefault
    If bDone Then
        fJoin.LetJoin() = oJoin
        fJoin.Show vbModal, Me
        bCancel = fJoin.Cancelled()
        If Not bCancel Then
            fJoin.GetJoin oJoin
        Else
            bDone = False
        End If
        Unload fJoin
    End If
    If bDone Then
        Screen.MousePointer = vbHourglass
        iNewJoinT = oJoin.JoinType()
        ReDim sTables(3) As String
        
        sTables(0) = oJoin.BaseTbl()
        sTables(1) = oJoin.RelatedTbl()
        
        For i = 1 To collJoin.Count
            If iCtrl <> collJoin(i).iControl() Then
                sTables(2) = collJoin(i).BaseTbl()
                sTables(3) = collJoin(i).RelatedTbl()
                If (sTables(0) = sTables(2) And sTables(1) = sTables(3)) _
                    Or (sTables(0) = sTables(3) And sTables(1) = sTables(2)) Then
                    collJoin(i).JoinType() = iNewJoinT
                End If
            End If
        
        Next
        Screen.MousePointer = vbDefault
        DrawJoin
    End If
    
End Sub

Private Sub myLst_BeginMove(Index As Integer, ByVal x As Integer, ByVal y As Integer)
    iCurDrag = Index
    iOldX = x
    iOldY = y
    myLst(Index).ZOrder
End Sub

Private Sub myLst_BeginRefer(Index As Integer)
    iCurDrag = Index
    picTBL.DragIcon = myLst(Index).DragPic
    picTBL.Drag vbBeginDrag
End Sub

Private Sub myLst_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    
    If Index <> iCurDrag And myLst(iCurDrag).Visible Then
        myLst(Index).Referred x, y
    End If
    iOldY = 0
    iOldX = 0

End Sub

Private Sub myLst_DropRefer(Index As Integer, iItem As Integer)
    
    Dim oJoin As clsJoin
    Dim oObj As cViewIt
    Dim sBaseTB As String, sBaseF As String, sRelTB As String, sRelF As String
    
    Dim iOr As Integer, iCtrl As Integer
    
    iOr = collItem(GetTblItem(myLst(iCurDrag).Caption)).iControl()
    Set oObj = collItem(GetTblItem(myLst(Index).Caption))
        
    sBaseTB = myLst(Index).Caption
    sBaseF = myLst(Index).GetItem(iItem)
    sRelTB = myLst(iOr).Caption
    sRelF = myLst(iOr).GetItem(myLst(iOr).DragInit)
    
    If Not ExistJoin(sBaseTB, sBaseF, sRelTB, sRelF) Then
        iCtrl = NextSlot_Join()
        
        Set oJoin = New clsJoin
        oJoin.BaseTbl() = sBaseTB
        oJoin.Field() = sBaseF
        oJoin.iBaseControl() = Index
        oJoin.RelatedTbl() = sRelTB
        oJoin.AliasBaseTbl() = GetAliasedTbl(iCtrl)
        oJoin.AliasRelatedTbl() = GetAliasedTbl(iOr)
        oJoin.RelatedField() = sRelF
        oJoin.iRefControl() = iOr
        oJoin.iControl() = iCtrl
        oJoin.JoinType() = jInner
        oJoin.Id() = Now()
        collJoin.Add oJoin, oJoin.Id()
        Load imgJoin(iCtrl)
    
        imgJoin(iCtrl).Picture = LoadResPicture(oJoin.JoinType(), vbResIcon)
        imgJoin(iCtrl).Visible = True
    
        Set oJoin = Nothing
        DrawJoin
    End If
    iCurDrag = 0
    
End Sub

Private Sub myLst_EndMoving(Index As Integer, ByVal x As Integer, ByVal y As Integer)
    iOldX = 0
    iOldY = 0
    DrawJoin
End Sub

Private Sub myLst_ItemCheck(Index As Integer, Item As Integer)
    
    Dim udtTMP As sFieldUdt
    
    Dim i As Integer, l As Integer
    Dim bDone As Boolean
    Dim sBuf As String
    
    sBuf = myLst(Index).GetItem(Item)
    
    If myLst(Index).IsSelected(Item) Then
        'add
        With udtTMP
            .iRow = Item
            .iCtrl = Index
            .sField = sBuf
            If mbIsGroup Then .iGroup = 1
            .sTb = myLst(Index).Caption
            .sTbAlias = GetAliasTB(.sTb)
                        
            If myLst(Index).TType = Tbl Then
                .sOwner = GetTableOwner(objServer, sDbName, IIf(Len(.sTbAlias) <> 0, .sTbAlias, .sTb))
            Else
                .sOwner = GetViewOwner(objServer, sDbName, IIf(Len(.sTbAlias) <> 0, .sTbAlias, .sTb))
            End If
            .bShow = True
            If sBuf <> "*" Then .sAlias = GetAlias(sBuf, "")
        End With
        
        For i = 0 To UBound(arrField)
            If arrField(i).iCtrl = Index Then
                For l = i + 1 To UBound(arrField)
                    If arrField(l).iCtrl <> Index Then
                        bDone = True
                        InsertUDT l, udtTMP
                        Exit For
                    End If
                Next
            End If
            If bDone Then Exit For
        Next
        If Not bDone Then
            InsertUDT UBound(arrField), udtTMP
        End If
    Else
        'remove
        For i = 0 To UBound(arrField)
            If arrField(i).iCtrl = Index And arrField(i).sField = sBuf Then
                RemoveUDT i
                Exit For
            End If
        Next
    End If
    
    LoadFlexG
    WriteSql
    
End Sub

Private Sub myLst_Moving(Index As Integer, ByVal x As Integer, ByVal y As Integer)

    Dim iX As Integer, iY As Integer
    
    With myLst(Index)
        iX = .Left + x - iOldX
        iY = .Top + y - iOldY
        If iX < 0 Then
            iX = 0
        ElseIf iX + myLst(Index).Width > picTBL.ScaleWidth Then
            iX = picTBL.ScaleWidth - myLst(Index).Width
        End If
        If iY < 0 Then
            iY = 0
        ElseIf iY + myLst(Index).Height > picTBL.ScaleHeight Then
            iY = picTBL.ScaleHeight - myLst(Index).Height
        End If
        .Move iX, iY
    End With

End Sub

Private Sub myLst_QueryUnload(Index As Integer, bCancel As Boolean)
    
    Dim s As String
    If Index > 0 Then
        s = myLst(Index).Caption
        Debug.Print myLst(Index).TType = Tbl
        If MsgBox(MyLoadResString(k_Remove_Object_from_current_View), vbQuestion Or vbOKCancel, s) = vbOK Then
            RemoveItem s, myLst(Index).TType
            RemoveMassUDT Index
            LoadFlexG
        End If
    End If
End Sub


Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Static b As Boolean

    TestActiveControl
    Select Case Button.Index
        Case 1  'save
            SaveView
        Case 2  'property
            ViewProperty
        Case 4, 5, 6, 7
            If Not b Then
                b = True
            'Button.Value = IIf(Button.Value = tbrUnpressed, tbrPressed, tbrUnpressed)
                bVisible(Button.Index - 4) = Not bVisible(Button.Index - 4)
                SetState Button.Index, bVisible(Button.Index - 4)
                Form_Resize
                b = False
            End If
        Case 9  'Run
            ExecuteSql False
        Case 10 'Verify
            ExecuteSql True
        Case 11 'Group
            GroupObject
        Case 12 'Clear
            ClearFlex
        Case 14  'ShowTbl
            ShowTbl
        Case 15
            ArrangeCtrl
        Case 16
            LoadFlexG
    End Select
End Sub

Private Sub TblView_BeginMove(ByVal x As Integer, ByVal y As Integer)
    iOldX = x
    iOldY = y
    TblView.ZOrder
'    iCurDrag = -2
'    TblView.Drag vbBeginDrag
'    TblView.Visible = False
End Sub

Private Sub TblView_EndMoving(ByVal x As Integer, ByVal y As Integer)
    iOldX = 0
    iOldY = 0

End Sub

Private Sub TblView_GotFocus()
    cbo.Visible = False
    tFlex.Visible = False
End Sub

Private Sub TblView_ItemSelected(ByVal iItem As Integer, ByVal iLst As LstType)
    LoadItem iItem, iLst
End Sub

Private Sub TblView_Moving(ByVal x As Integer, ByVal y As Integer)

    Dim iX As Integer, iY As Integer
    
    With TblView
        iX = .Left + x - iOldX
        iY = .Top + y - iOldY
        If iX < 0 Then
            iX = 0
        ElseIf iX + .Width > Me.ScaleWidth Then
            iX = Me.ScaleWidth - .Width
        End If
        If iY < 0 Then
            iY = 0
        ElseIf iY + .Height > Me.ScaleHeight Then
            iY = Me.ScaleHeight - .Height
        End If
        .Move iX, iY
    End With

End Sub

Private Sub TblView_QueryUnload(bCancel As Boolean)
    TblView.Visible = False
    bTbl = False
    SetState 14, bTbl
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub VScr_Change()
    vScr_Scroll
End Sub

Private Sub vScr_Scroll()
    picTBL.Top = -iIncY * VScr.Value
End Sub
Private Sub HScr_Change()
    HScr_Scroll
End Sub
Private Sub HScr_Scroll()
    picTBL.Left = -iIncX * HScr.Value
End Sub

Private Sub ShowTbl()

    If Not bTbl Then
        bTbl = True
        Screen.MousePointer = vbHourglass
        With TblView
            .Move Me.ScaleWidth - .Width - 200, Me.ScaleHeight - .Height - 200
            .Visible = True
            .ZOrder
        End With
        Screen.MousePointer = vbDefault
        TblView.SetFocus
    Else
        TblView.Visible = False
        bTbl = False
    End If
    SetState 14, bTbl
    
End Sub
Private Sub LoadItem(ByVal iItem As Integer, ByVal iLst As LstType)

    Dim iC As Integer
    
    Dim sOName As String, bIsTbl As Boolean
    Dim oTb As SQLDMO.Table
    Dim oSqlList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column
    Dim iType As LstType
    Dim sBuf As String
    Dim bIsAlias As Boolean
    Dim oView As SQLDMO.View
        
    Screen.MousePointer = vbHourglass
    If collItem.Count Then
        iC = NextSlot_TBView()
    Else
        iC = 1
    End If
    
    Dim oObj As cViewIt
    Set oObj = New cViewIt
    iType = TblView.CurrentItem()
    sOName = TblView.GetItem(iItem, iLst)
    bIsTbl = iType = isTbl
    sBuf = ""
    bIsAlias = GetResolvedObj(sOName, sBuf)
    oObj.sName = IIf(bIsAlias, sBuf, sOName)
    oObj.sAlias = IIf(bIsAlias, sOName, sBuf)
    oObj.isTbl() = bIsTbl
    If bIsTbl Then
        Set oTb = GetTbByName(objServer, sDbName, sOName)
        oObj.sOwner = oTb.Owner
        Set oTb = Nothing
    Else
        Set oView = GetViewByName(objServer, sDbName, sOName)
        oObj.sOwner = oView.Owner
        Set oView = Nothing
    End If
    oObj.iControl() = iC
    collItem.Add oObj, oObj.sName
    
    Load myLst(iC)
    With myLst(iC)
        '.Move -1000, -1000
        .Visible = False
        DoEvents
        .Caption = IIf(bIsAlias, sBuf, sOName)
        .IsGrouped = mbIsGroup
        .AddItem "*" '& vbTab & MyLoadResString(k_All_Columns )
        .TType = 1 - Abs(bIsTbl)
        If bIsTbl Then
            Set oTb = GetTbByName(objServer, sDbName, IIf(bIsAlias, oObj.sAlias, oObj.sName))
            'Set oTb = oDB.Tables(IIf(bIsAlias, oObj.sAlias, oObj.sName))
            
            For Each oCol In oTb.Columns
                .AddItem oCol.Name, bIsKey:=oCol.InPrimaryKey
            Next
            Set oTb = Nothing
        Else
            'Set oSqlList = oDB.Views(sOName).ListColumns
            Set oView = GetViewByName(objServer, sDbName, sOName)
            
            Set oSqlList = oView.ListColumns
            For Each oCol In oSqlList
                .AddItem oCol.Name
            Next
            Set oView = Nothing
        End If
        '.Move 200, 200
        .Visible = True
        .ZOrder
    End With
    Set oCol = Nothing
    Set oSqlList = Nothing
    Set oTb = Nothing
    Set oView = Nothing
    WriteSql
    
    Screen.MousePointer = vbDefault

End Sub
Private Sub RemoveItem(ByVal sVal As String, ByVal iType As CtrlType)
    
    Dim i As Integer, iC As Integer, iJoin As Integer
    Dim bDone As Boolean
    Dim bTbl As Boolean
    
    bTbl = iType = Tbl
    
    For i = 1 To collItem.Count
        If sVal = collItem(i).sName() And bTbl = collItem(i).isTbl() Then
            iC = collItem(i).iControl()
            bDone = True
            Exit For
        End If
    Next
    If bDone Then
        Unload myLst(iC)
        collItem.Remove sVal

        Dim oJoin As clsJoin
        For Each oJoin In collJoin
            If oJoin.iBaseControl() = iC Or oJoin.iRefControl() = iC Then
                iJoin = oJoin.iControl()
                Unload imgJoin(iJoin)
                collJoin.Remove oJoin.Id()
            End If
        Next
        Set oJoin = Nothing
        DrawJoin
    
    End If
    
End Sub
Private Function NextSlot_TBView() As Integer
    
    Screen.MousePointer = vbHourglass
   
    Dim Index As Long
    Dim fItem As Long
    Dim indLimit As Long, lastSwap As Long
    Dim numEl As Long
    Dim tmp As Integer
        
    ReDim iMatrix(collItem.Count - 1) As Integer
    For Index = 1 To collItem.Count
        iMatrix(Index - 1) = collItem(Index).iControl()
    Next
        
    numEl = UBound(iMatrix)
        
    fItem = LBound(iMatrix)
    lastSwap = numEl
    Do
        indLimit = lastSwap - 1
        lastSwap = 0
        For Index = fItem To indLimit
            tmp = iMatrix(Index)
            If tmp > iMatrix(Index + 1) Then
                iMatrix(Index) = iMatrix(Index + 1)
                iMatrix(Index + 1) = tmp
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
    tmp = 1
    lastSwap = 1
    Do
        For Index = 0 To numEl
            If iMatrix(Index) = tmp Then
                tmp = tmp + 1
                Exit For
            End If
        Next
        If tmp = lastSwap Then
            Exit Do
        Else
            lastSwap = tmp
        End If
    Loop
    NextSlot_TBView = tmp

End Function
Private Sub LoadFlexG()

    Dim i As Integer
    bFlexLoad = True
    With flexG(0)
        .Redraw = False
        .FixedCols = 0
        .FixedRows = 0
        .Clear
        .Cols = 8
        .Rows = 1
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        For i = 1 To 7
            .TextMatrix(0, i) = MyLoadResString(620 + i)
        Next
        For i = 0 To UBound(arrField)
            .AddItem vbTab & arrField(i).sField & vbTab & arrField(i).sAlias & vbTab & arrField(i).sTb & vbTab & vbTab & arrGroup(arrField(i).iGroup) & vbTab & arrField(i).sCriteria & vbTab & arrField(i).sOr
            .Col = 4
            .Row = .Rows - 1
                        
            Set flexG(0).CellPicture = LoadResPicture(IIf(Abs(arrField(i).bShow), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            
            If mbIsGroup Then
                .Col = kGroup
                Set flexG(0).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                .CellPictureAlignment = flexAlignRightTop
            End If
            
        Next
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns
        .ColWidth(kGroup) = IIf(mbIsGroup, .ColWidth(1), 0)
        .AllowBigSelection = False
        
        .FixedRows = 1
        .FixedCols = 1
        .Redraw = True
        SettaLargFlex flexG(0), 1, 400
        
    End With
    bFlexLoad = False
    flexG_RowColChange 0
    
End Sub
Private Function NextSlot_Join() As Integer
    
    Screen.MousePointer = vbHourglass
   
    Dim Index As Long
    Dim fItem As Long
    Dim indLimit As Long, lastSwap As Long
    Dim numEl As Long
    Dim tmp As Integer
       
    fItem = -1
    ReDim iMatrix(0) As Integer
    For Index = 1 To collJoin.Count
        fItem = fItem + 1
        ReDim Preserve iMatrix(fItem) As Integer
        iMatrix(fItem) = collJoin(Index).iControl()
    Next
    
    numEl = UBound(iMatrix)
        
    fItem = LBound(iMatrix)
    lastSwap = numEl
    Do
        indLimit = lastSwap - 1
        lastSwap = 0
        For Index = fItem To indLimit
            tmp = iMatrix(Index)
            If tmp > iMatrix(Index + 1) Then
                iMatrix(Index) = iMatrix(Index + 1)
                iMatrix(Index + 1) = tmp
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
    tmp = 1
    lastSwap = 1
    Do
        For Index = 0 To numEl
            If iMatrix(Index) = tmp Then
                tmp = tmp + 1
                Exit For
            End If
        Next
        If tmp = lastSwap Then
            Exit Do
        Else
            lastSwap = tmp
        End If
    Loop
    
    NextSlot_Join = tmp
    Screen.MousePointer = vbDefault

End Function
Private Function GetTblItem(ByVal sVal As String) As Integer
    
    Dim i As Integer
    
    For i = 1 To collItem.Count
        If sVal = collItem(i).sName() Then
            GetTblItem = i
            Exit Function
        End If
    Next
    GetTblItem = -1

End Function

Private Sub DrawJoin()

    Dim i As Integer, iCtrl As Integer, iDest As Integer
    Dim iLst As Integer
    Dim lFromX As Long, lFromY As Long
    Dim lX As Long, lY As Long
    ReDim xJoin(collItem.Count) As Integer
    Dim lyOff As Long
    Dim iOffLstJoin As Integer
    Dim lErr As Long
        
    Screen.MousePointer = vbHourglass
    picTBL.Cls
    
    iOffLstJoin = (imgJoin(0).Width / 2) + 20
    For i = 1 To collJoin.Count
        With collJoin(i)
            iLst = .iBaseControl()
            lX = myLst(iLst).Left - 50 - imgJoin(0).Width
            lY = myLst(iLst).Top + 150
            If lX < 20 Then
                myLst(iLst).Left = 320
                lX = myLst(iLst).Left - 50 - imgJoin(0).Width
            End If
            iCtrl = 0
            On Local Error Resume Next
            lyOff = lY + ((xJoin(iLst) * imgJoin(0).Height) + 20)
            lErr = Err.Number
            On Local Error GoTo 0
            
            If lErr = 0 Then
                xJoin(iLst) = xJoin(iLst) + 1
                iCtrl = .iControl()
                iDest = .iRefControl()
                
                imgJoin(iCtrl).Move lX, lyOff
                imgJoin(iCtrl).Picture = LoadResPicture(.JoinType(), vbResIcon)
                
                lFromY = imgJoin(iCtrl).Top + (imgJoin(0).Height \ 2)
                lFromX = lX + (imgJoin(0).Height \ 2)
                picTBL.Line (lFromX, lFromY)-(myLst(iDest).Left - iOffLstJoin, myLst(iDest).Top + (myLst(iDest).Height \ 2))
            End If
        End With
    Next
    Screen.MousePointer = vbDefault
    WriteSql

End Sub
Private Sub ClearFlex()
    
    Screen.MousePointer = vbHourglass
    With flexG(1)
        .Redraw = False
        .Clear
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 1
        .Rows = 1
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
End Sub
Private Sub ViewProperty()

    Dim bCancel As Boolean
                          
    fJoin.ViewOptions cViewP
    fJoin.Show vbModal, Me
    bCancel = fJoin.Cancelled()
    If Not bCancel Then
        fJoin.GetView cViewP
    End If
    Unload fJoin
    
    If Not bCancel Then WriteSql
    
End Sub
Private Function ExistJoin(ByVal sBaseTB As String, ByVal sBaseF As String, ByVal sRefTb As String, ByVal sRefF As String) As Boolean

    Dim i As Integer
    sBaseTB = UCase$(sBaseTB)
    sBaseF = UCase$(sBaseF)
    sRefTb = UCase$(sRefTb)
    sRefF = UCase$(sRefF)
    For i = 1 To collJoin.Count
        If UCase$(collJoin(i).BaseTbl()) = sBaseTB And UCase$(collJoin(i).Field()) = sBaseF And UCase$(collJoin(i).RelatedTbl()) = sRefTb And UCase$(collJoin(i).RelatedField()) = sRefF Then
            ExistJoin = True
            Exit For
        End If
    Next

End Function
Private Function ExistInArray(ByVal sSearch4 As String, sArray() As String) As Boolean

    Dim i As Integer
    
    For i = 0 To UBound(sArray)
        If StrComp(sArray(i), sSearch4, vbTextCompare) = 0 Then
            ExistInArray = True
            Exit For
        End If
    Next
End Function

Private Sub RebaseJoinColl(iOrd() As Integer)

    Dim i As Integer, x As Integer, z As Integer
    ReDim sTables(3) As String
    
    ReDim bDone(1 To collJoin.Count) As Boolean
    ReDim iOrd(1 To collJoin.Count) As Integer
    Dim iLoop As Integer
    bDone(1) = True

    For i = 1 To collJoin.Count
        iOrd(i) = i
    Next
    x = 1
    iLoop = 1
    Do
        For i = x + 1 To collJoin.Count
            sTables(0) = collJoin(iLoop).BaseTbl()
            sTables(1) = collJoin(iLoop).RelatedTbl()
            sTables(2) = collJoin(i).BaseTbl()
            sTables(3) = collJoin(i).RelatedTbl()

            If Not bDone(i) Then
                If (sTables(0) = sTables(2) And sTables(1) = sTables(3)) _
                    Or (sTables(0) = sTables(3) And sTables(1) = sTables(2)) Then
                    x = x + 1
                    z = iOrd(x)
                    iOrd(x) = i
                    iOrd(i) = z
                    bDone(i) = True
                End If
            End If
        Next
        iLoop = iLoop + 1
    Loop Until iLoop > collJoin.Count
    
End Sub
Private Sub GroupObject()
    
    Dim i As Integer
    Dim objLst As myLst
    Dim bDone As Boolean
    Dim iCol As Integer, iRow As Integer
        
    If TypeOf Screen.ActiveControl Is myLst Then
        Set objLst = Screen.ActiveControl
        mbIsGroup = Not objLst.IsGrouped()
        
        cViewP.isGroup() = mbIsGroup
        If Not mbIsGroup Then cViewP.Extention() = 0
        cViewP.ExtOpt() = -1
        
        For i = 1 To collItem.Count
            myLst(collItem(i).iControl).IsGrouped = mbIsGroup
        Next
        bDone = True

        Set objLst = Nothing
        If bDone Then
            For i = 0 To UBound(arrField) - 1
                arrField(i).iGroup = 0
            Next
            With flexG(0)
                iCol = .Col
                iRow = .Row
                If mbIsGroup Then
                    For i = 0 To UBound(arrField) - 1
                        arrField(i).iGroup = 1
                    Next
                    LoadFlexG
                Else
                    .Redraw = False
                    bFlexLoad = True
                    .Row = 1
                    .Col = kGroup
                    .RowSel = .Rows - 1
                    .ColSel = kGroup
                    .Clip = ""
                    .Redraw = True
                    flexG(0).Row = miRow
                    flexG(0).Col = kGroup
                    Set flexG(0).CellPicture = Nothing
                    miRow = 0
                    .ColWidth(kGroup) = 0
                    .Col = iCol
                    bFlexLoad = False
                    .Row = iRow
                    miRow = iRow
                    'flexG_RowColChange
                End If
            End With
            WriteSql
        End If
    End If
    SetState 11, mbIsGroup
        
End Sub
Private Sub InsertUDT(ByVal iElem As Integer, udtTMP As sFieldUdt)

    Dim i As Integer, iBound As Integer
    
    iBound = UBound(arrField) + 1
    ReDim Preserve arrField(iBound) As sFieldUdt
    
    For i = iBound To iElem + 1 Step -1
        LSet arrField(i) = arrField(i - 1)
    Next
    LSet arrField(iElem) = udtTMP
    
End Sub
Private Sub RemoveUDT(ByVal iElem As Integer)

    Dim i As Integer
   
    For i = iElem To UBound(arrField) - 1
        LSet arrField(i) = arrField(i + 1)
    Next
    i = UBound(arrField) - 1
    If i < 0 Then i = 0
    ReDim Preserve arrField(i) As sFieldUdt
    
End Sub
Private Sub RemoveMassUDT(ByVal iCtrl As Integer)

    Dim iLoop As Integer, iCounter As Integer
    ReDim tmpField(10) As sFieldUdt
    
    iCounter = 0
    For iLoop = 0 To UBound(arrField)
        If arrField(iLoop).iCtrl <> iCtrl Then
            LSet tmpField(iCounter) = arrField(iLoop)
            iCounter = iCounter + 1
            If iCounter > UBound(tmpField) Then ReDim Preserve tmpField(iCounter + 10) As sFieldUdt
        End If
    Next
    iCounter = iCounter - 1
    ReDim Preserve tmpField(iCounter) As sFieldUdt
    
    ReDim arrField(iCounter) As sFieldUdt
    For iLoop = 0 To UBound(arrField)
        LSet arrField(iLoop) = tmpField(iLoop)
    Next

End Sub
Private Sub LoadCbo(ByVal iCol As Integer)

    Dim i As Integer
    With cbo
        .Clear
        If iCol = kGroup Then
            For i = 0 To UBound(arrGroup)
                .AddItem arrGroup(i)
            Next
        End If
    End With
End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyRIGHT, vbKeyLEFT
                iCode = KeyCode
                KeyCode = 0
                cbo.Visible = False
                flexG(0).Col = flexG(0).Col - (38 - iCode)
        End Select
    End If
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            m_blnAnnullaMod = True
            KeyAscii = 0
            flexG(0).SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            flexG(0).SetFocus
    End Select

End Sub

Private Sub cbo_LostFocus()

    ReDim lW(flexG(0).Cols - 1) As Long
    If Not m_blnAnnullaMod Then
        With arrField(miRow - 1)
            If miCol = kGroup Then
                arrField(miRow - 1).iGroup = cbo.ListIndex
                arrField(miRow - 1).sAlias = GetAlias(arrField(miRow - 1).sField, IIf(arrField(miRow - 1).iGroup > 1, "EXPR1", ""), miRow - 1)
                flexG(0).TextMatrix(miRow, miCol) = arrGroup(arrField(miRow - 1).iGroup)
                flexG(0).TextMatrix(miRow, 2) = arrField(miRow - 1).sAlias
            End If
        End With
    End If
    
    cbo.Visible = False
    If Not m_blnAnnullaMod Then WriteSql
    m_blnAnnullaMod = True
    miCol = -1

End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iImg As Integer
    
    iRow = flexG(0).Row
    iCol = flexG(0).Col
    
    bEdit = True
    If flexG(0).CellBackColor = vbInactiveBorder Then bEdit = False
    
    bEdit = Len(arrField(iRow - 1).sField)
    
    If bEdit Then
        If iCol = kAlias Then
            sBuf = arrField(iRow - 1).sAlias
            iLen = 25
        ElseIf iCol = kCondition Then
            sBuf = arrField(iRow - 1).sCriteria
            iLen = 50
        ElseIf iCol = kOr Then
            sBuf = arrField(iRow - 1).sOr
            iLen = 50
        ElseIf iCol = kShow Then
        Else
            bEdit = False
        
        End If
        
        miRow = iRow
        
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If iCol = kShow Then
                    iImg = 1 - (-arrField(iRow - 1).bShow)
                    arrField(iRow - 1).bShow = CBool(iImg)
                    If iImg <> 0 Then
                        arrField(iRow - 1).sAlias = GetAlias(arrField(iRow - 1).sField, arrField(iRow - 1).sAlias, iRow - 1)
                    Else
                        If Len(arrField(miRow - 1).sCriteria) + Len(arrField(miRow - 1).sOr) Then
                            arrField(iRow - 1).sAlias = ""
                        Else
                            iImg = arrField(miRow - 1).iRow
                            iLen = arrField(miRow - 1).iCtrl
                            'RemoveUDT miRow - 1
                            myLst(iLen).IsSelected(iImg) = False
                            iRow = IIf(iRow - 1 = 0, 1, iRow - 1)
                        End If
                    End If
                    LoadFlexG
                    WriteSql
                    flexG(0).Row = iRow
                    flexG(0).Col = iCol

                    bEdit = False
                End If
        End Select
    End If

    If bEdit Then
        miCol = iCol
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf
        tFlex.SelStart = 0
        With flexG(0)
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If KeyAscii = vbKeySpace Then KeyAscii = 0
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        'Case 60, 64, 124, 91,92, 93    '<>|[\]
        '    KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex.Text = flexG(0).Text
            flexG(0).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            flexG(0).SetFocus
            KeyAscii = 0
    End Select

End Sub
Private Sub tFlex_LostFocus()
    
    Dim bUpdate As Boolean
    Dim sBuf As String
    
    If Not m_blnAnnullaMod Then
        tFlex = Trim$(tFlex.Text)
        sBuf = tFlex.Text
        bUpdate = True
        If miCol = kAlias Then
            arrField(miRow - 1).sAlias = ""
            If Len(sBuf) = 0 And arrField(miRow - 1).iGroup > 1 Then sBuf = "EXPR1"
            
            sBuf = GetAlias(arrField(miRow - 1).sField, sBuf, miRow - 1)
            arrField(miRow - 1).sAlias = sBuf
            If Len(sBuf) And Not arrField(miRow - 1).bShow Then
                arrField(miRow - 1).bShow = True
                flexG(0).Row = miRow
                flexG(0).Col = kShow
                Set flexG(0).CellPicture = LoadResPicture(k_ResChkON, vbResBitmap)
                flexG(0).CellPictureAlignment = flexAlignCenterCenter
            End If
        ElseIf miCol = kCondition Then
            arrField(miRow - 1).sCriteria = sBuf
        ElseIf miCol = kOr Then
            arrField(miRow - 1).sOr = sBuf
        Else
            bUpdate = False
        End If
    End If
    If bUpdate Then
        flexG(0).TextMatrix(miRow, miCol) = sBuf
        WriteSql
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
    m_blnAnnullaMod = True
    tFlex.Visible = False
    miCol = -1
    
End Sub
Private Function GetAlias(ByVal sTest4 As String, ByVal sAlias As String, Optional ByVal iEsclude As Integer = -1) As String

    Dim l As Integer, i As Integer
    Dim bDone As Boolean, bRepeat As Boolean
        
    Do
        bDone = True
        For i = 0 To UBound(arrField)
            bRepeat = False
            If iEsclude <> i Then
                If arrField(i).bShow Then
                    If Len(sAlias) = 0 Then
                        bRepeat = sTest4 = arrField(i).sField And Len(arrField(i).sAlias) = 0
                    Else
                        bRepeat = (sTest4 = arrField(i).sField And sAlias = arrField(i).sAlias) Or (sAlias = arrField(i).sAlias)
                    End If
                End If
                If bRepeat Then
                    l = l + 1
                    sAlias = "EXPR" & l
                    bDone = False
                    
                    Exit For
                End If
            End If
        Next
    Loop Until bDone
    GetAlias = sAlias
End Function
Private Sub WriteSql()

    Dim i As Integer
    Dim iJoinT As Integer
    Dim sSql As String
    Dim bAddF As Boolean
    Dim bAddTB As Boolean
    Dim sJoin As String, sFilter As String
    Dim bDescJoin As Boolean
    Dim iFldC As Integer
    Dim iCurrTB As Integer
    Dim sTmpTB As String
    Dim bTest As Boolean
    Dim sSumFilter As String
    Dim bNewLine As Boolean
    Dim sTmpJoin As String
    Dim bInvert As Boolean
    Dim sOwner As String
    Dim sTmp As String
    
    Screen.MousePointer = vbHourglass
    iJoinT = collJoin.Count
    Dim bIsTb As Boolean
    
    ReDim sTBn(0) As String
    ReDim sFldN(0) As String
    
    iFldC = -1
        
    sSql = "SELECT " & IIf(cViewP.bOption(0), "DISTINCT ", "")
    sSql = sSql & IIf(cViewP.TopRec(), "TOP " & Format$(cViewP.TopRec()) & " ", "")
    iCurrTB = -1
    
    bNewLine = True
    For i = 1 To collItem.Count
        bIsTb = collItem(i).isTbl()
        If bIsTb Then
            sOwner = GetTableOwner(objServer, sDbName, IIf(Len(collItem(i).sAlias()), collItem(i).sAlias(), collItem(i).sName()))
        Else
            sOwner = GetViewOwner(objServer, sDbName, IIf(Len(collItem(i).sAlias()), collItem(i).sAlias(), collItem(i).sName()))
        End If
        sTmpTB = sOwner & "." & IIf(Len(collItem(i).sAlias()), collItem(i).sAlias() & " ", "") & collItem(i).sName()
        iCurrTB = iCurrTB + 1
        ReDim Preserve sTBn(iCurrTB) As String
        sTBn(iCurrTB) = sTmpTB
    Next
    
    For i = 0 To UBound(arrField) - 1
        sOwner = ""
        If sTmpTB <> arrField(i).sTb Then
            If Len(arrField(i).sTbAlias) = 0 Then
                sOwner = arrField(i).sOwner
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
            End If
            sTmpTB = sOwner & IIf(Len(arrField(i).sTbAlias), arrField(i).sTbAlias & " ", "") & arrField(i).sTb
            bNewLine = False
        End If
        If bNewLine Then sSql = sSql & vbCrLf
        
        If Len(arrField(i).sField) <> 0 And arrField(i).bShow Then
            If bAddF Then sSql = sSql & ", "
            If arrField(i).iGroup > 1 Then sSql = sSql & " " & arrGroup(arrField(i).iGroup) & "("
            If arrField(i).iRow = 0 Then
                sSql = sSql & sOwner & arrField(i).sTb & ".*"
            Else
                sSql = sSql & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "]"
            End If
            If arrField(i).iGroup > 1 Then sSql = sSql & ")"
            If Len(arrField(i).sAlias) Then sSql = sSql & " AS " & arrField(i).sAlias
                        
            If Not bAddF Then bAddF = arrField(i).bShow
            If Not bNewLine Then bNewLine = arrField(i).bShow
            
            bAddTB = True
        End If
    Next
    
    If iJoinT = 0 Then
        bAddTB = False
        sSql = sSql & vbCrLf & " FROM "
        For i = 0 To UBound(sTBn)
            If Len(sTBn(i)) Then
                If Not bAddTB Then
                    bAddTB = True
                Else
                    sSql = sSql & ", "
                End If
                sSql = sSql & sTBn(i)
            End If
        Next
    Else
        ReDim iOrd(0) As Integer
        
        RebaseJoinColl iOrd
        
        bAddTB = True
        sSql = sSql & vbCrLf & " FROM "
        If Len(sTBn(0)) <> 0 Then sSql = sSql & sTBn(0)
        
        iFldC = 0
        ReDim sTBnotJoin(0) As String
        sTBnotJoin(0) = sTBn(0)
        ReDim sTables(3) As String
        ReDim bAlias(1) As Boolean
        For i = 1 To UBound(iOrd)
            bDescJoin = False

            With collJoin(iOrd(i))
                Debug.Print .AliasBaseTbl(), .BaseTbl(), .AliasRelatedTbl(), .RelatedTbl()
                
                If Len(.AliasBaseTbl()) Then
                    sTables(0) = .AliasBaseTbl()
                    bAlias(0) = True
                Else
                    bAlias(0) = False
                    sTables(0) = .BaseTbl()
                End If
                
                If Len(.AliasRelatedTbl()) Then
                    bAlias(1) = True
                    sTables(1) = .AliasRelatedTbl()
                Else
                    bAlias(1) = False
                    sTables(1) = .RelatedTbl()
                End If
                
                sOwner = GetTableOwner(objServer, sDbName, sTables(0))
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTables(0))
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                sTables(0) = sOwner & sTables(0)
                
                sOwner = GetTableOwner(objServer, sDbName, sTables(1))
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTables(1))
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                sTables(1) = sOwner & sTables(1)
                            
                iFldC = iFldC + 1
                ReDim Preserve sTBnotJoin(iFldC) As String
                                
                If sTables(1) = sTBn(0) And bAlias(1) = False Then
                    sTmpJoin = sTables(1)
                    sTables(1) = sTables(0)
                    sTables(0) = sTmpJoin
                    
                    sTmp = .AliasBaseTbl()
                    If Len(sTmp) = 0 Then sTmp = .BaseTbl()
                    sOwner = GetTableOwner(objServer, sDbName, sTmp)
                    If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTmp)
                    If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                    
                    sTBnotJoin(iFldC) = sOwner & IIf(bAlias(0), .AliasBaseTbl() & " ", "") & .BaseTbl()
                    bInvert = True
                Else
                    Debug.Print .AliasRelatedTbl()
                    Debug.Print .RelatedTbl()
                    'sOwner = GetTableOwner(objServer, sDbName, IIf(Len(.AliasRelatedTbl()) <> 0, .AliasRelatedTbl(), .RelatedTbl()))
                    'If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, IIf(Len(.AliasRelatedTbl()) <> 0, .AliasRelatedTbl(), .RelatedTbl()))
                    
                    sTmp = .AliasRelatedTbl()
                    If Len(sTmp) = 0 Then sTmp = .RelatedTbl()
                    sOwner = GetTableOwner(objServer, sDbName, sTmp)
                    If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTmp)
                    
                    If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                    sTBnotJoin(iFldC) = sOwner & IIf(bAlias(1), .AliasRelatedTbl() & " ", "") & .RelatedTbl()
                    bInvert = False
                End If
                'sTBnotJoin(iFldC) = IIf(bAlias(1), .AliasRelatedTbl() & " ", "") & .RelatedTbl()
                
                If i > 1 Then
                    If (sTables(0) = sTables(2) And sTables(1) = sTables(3)) _
                        Or (sTables(0) = sTables(3) And sTables(1) = sTables(2)) Then
                        sJoin = " AND "
                    Else
                        bDescJoin = True
                    End If
                    sTables(2) = sTables(0)
                    sTables(3) = sTables(1)

                Else
                    bDescJoin = True
                    sTables(2) = sTables(0)
                    sTables(3) = sTables(1)
                End If
                If bDescJoin Then
                    'sJoin = " " & Choose(.JoinType() + 1, "INNER", "LEFT OUTER", "RIGHT OUTER", "FULL OUTER") & " JOIN "
                    sJoin = " " & Choose((.JoinType()) - 100, "INNER", "LEFT OUTER", "RIGHT OUTER", "FULL OUTER") & " JOIN "
                    
                    If Not bInvert Then
                        Debug.Print IIf(Len(.AliasRelatedTbl()), .AliasRelatedTbl() & " ", "") & .RelatedTbl() & " ON "
                                                
                        'sOwner = GetTableOwner(objServer, sDbName, IIf(Len(.AliasRelatedTbl()) <> 0, .AliasRelatedTbl(), .RelatedTbl()))
                        'If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, IIf(Len(.AliasRelatedTbl()) <> 0, .AliasRelatedTbl(), .RelatedTbl()))
                        sTmp = .AliasRelatedTbl()
                        If Len(sTmp) = 0 Then sTmp = .RelatedTbl()
                        sOwner = GetTableOwner(objServer, sDbName, sTmp)
                        If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTmp)
                        
                        
                        If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                        
                        sJoin = sJoin & sOwner & IIf(Len(.AliasRelatedTbl()), .AliasRelatedTbl() & " ", "") & .RelatedTbl() & " ON "
                    Else
                        'sOwner = GetTableOwner(objServer, sDbName, IIf(Len(.AliasBaseTbl()) <> 0, .AliasBaseTbl(), .BaseTbl()))
                        'If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, IIf(Len(.AliasBaseTbl()) <> 0, .AliasBaseTbl(), .BaseTbl()))
                        sTmp = .AliasBaseTbl()
                        If Len(sTmp) = 0 Then sTmp = .BaseTbl()
                        sOwner = GetTableOwner(objServer, sDbName, sTmp)
                        If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTmp)
                                                
                        If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                    
                        sJoin = sJoin & sOwner & IIf(Len(.AliasBaseTbl()), .AliasBaseTbl() & " ", "") & .BaseTbl() & " ON "
                    End If
                End If
                Debug.Print sJoin
                Debug.Print .BaseTbl() & "." & .Field()
                
                'sOwner = GetTableOwner(objServer, sDbName, IIf(Len(.AliasBaseTbl()), .AliasBaseTbl(), .BaseTbl()))
                'If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, IIf(Len(.AliasBaseTbl()), .AliasBaseTbl(), .BaseTbl()))
                
                sTmp = .AliasBaseTbl()
                If Len(sTmp) = 0 Then sTmp = .BaseTbl()
                sOwner = GetTableOwner(objServer, sDbName, sTmp)
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, sTmp)
                
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                
                sJoin = sJoin & sOwner & .BaseTbl() & "." & .Field()
                
                sJoin = sJoin & " " & Choose(.JoinOperator() + 1, "=", "<>", "<", "<=", ">", ">=")
                'sJoin = sJoin & " " & sTBnotJoin(iFldC) & "." & .RelatedField()
                
                sOwner = GetTableOwner(objServer, sDbName, .RelatedTbl())
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, .RelatedTbl())
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
                
                sJoin = sJoin & sOwner & " " & .RelatedTbl() & "." & .RelatedField()
            End With
            
            sSql = sSql & IIf(i > 1, vbCrLf & " ", "") & sJoin
        Next
        
        For i = 0 To UBound(sTBn)
            If Len(sTBn(i)) Then
                If Not ExistInArray(sTBn(i), sTBnotJoin) Then sSql = sSql & ", " & sTBn(i)
            End If
        Next
    End If
    
    'GroupBy
    For i = 0 To UBound(arrField)
        sFilter = ""
        sOwner = ""
        If Len(arrField(i).sTbAlias) = 0 Then
            sOwner = arrField(i).sOwner
            If Len(sOwner) <> 0 Then sOwner = sOwner & "."
        End If
        
        If arrField(i).iGroup = 1 Then
            sFilter = sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "] "
        End If
        If Len(sFilter) Then
            If Not bTest Then
                sFilter = IIf(cViewP.Extention() And cViewP.ExtOpt() = 2, " ALL ", "") & sFilter
                sFilter = vbCrLf & " " & arrGroup(1) & " " & sFilter
                bTest = True
            Else
                sFilter = ", " & sFilter
            End If
        End If
        If Len(sFilter) Then
            If Len(sSumFilter) Then sSumFilter = sSumFilter & vbCrLf
            sSumFilter = sSumFilter & sFilter
        End If
    Next
    If Len(sSumFilter) Then
        sFilter = ""
        If cViewP.Extention() Then
            If cViewP.ExtOpt() = 0 Then
                sFilter = "WITH CUBE"
            ElseIf cViewP.ExtOpt() = 1 Then
                sFilter = "WITH ROLLUP"
            End If
        End If
        sSql = sSql & sSumFilter & IIf(Len(sFilter), " " & sFilter, "")
    End If
    
    
    'where filters
    bTest = False
    sSumFilter = ""
    If Not mbIsGroup Then
        For i = 0 To UBound(arrField)
            sFilter = ""
            sOwner = ""
            If Len(arrField(i).sTbAlias) = 0 Then
                sOwner = GetTableOwner(objServer, sDbName, arrField(i).sTb)
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, arrField(i).sTb)
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
            End If
            
            If Len(arrField(i).sCriteria) Then
                sFilter = "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "] " & arrField(i).sCriteria & ")"
            End If
            If Len(arrField(i).sOr) Then
                If Len(sFilter) Then sFilter = sFilter & " OR "
                sFilter = sFilter & "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "] " & arrField(i).sOr & ")"
            End If
            If Len(sFilter) Then
                If Not bTest Then
                    sFilter = vbCrLf & " WHERE " & sFilter
                    bTest = True
                Else
                    sFilter = " AND " & sFilter
                End If
            End If
            If Len(sFilter) Then
                If Len(sSumFilter) Then sSumFilter = sSumFilter & vbCrLf
                sSumFilter = sSumFilter & sFilter
            End If
        Next
    Else
        For i = 0 To UBound(arrField)
            sFilter = ""
            sOwner = ""
            If Len(arrField(i).sTbAlias) = 0 Then
                sOwner = GetTableOwner(objServer, sDbName, arrField(i).sTb)
                If Len(sOwner) = 0 Then sOwner = GetViewOwner(objServer, sDbName, arrField(i).sTb)
                If Len(sOwner) <> 0 Then sOwner = sOwner & "."
            End If

            
            If Len(arrField(i).sCriteria) Then
                If arrField(i).iGroup > 1 Then
                    sFilter = "(" & arrGroup(arrField(i).iGroup) & "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "]) " & arrField(i).sCriteria & ")"
                Else
                    sFilter = "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "] " & arrField(i).sCriteria & ")"
                End If
            End If
            If Len(arrField(i).sOr) Then
                If Len(sFilter) Then sFilter = sFilter & " OR "
                If arrField(i).iGroup > 1 Then
                    sFilter = sFilter & "(" & arrGroup(arrField(i).iGroup) & "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "]) " & arrField(i).sOr & ")"
                Else
                    sFilter = sFilter & "(" & sOwner & arrField(i).sTb & ".[" & arrField(i).sField & "] " & arrField(i).sOr & ")"
                End If
            End If
            If Len(sFilter) Then
                If Not bTest Then
                    sFilter = vbCrLf & " HAVING " & sFilter
                    bTest = True
                Else
                    sFilter = " AND " & sFilter
                End If
            End If
            If Len(sFilter) Then
                If Len(sSumFilter) Then sSumFilter = sSumFilter & vbCrLf
                sSumFilter = sSumFilter & sFilter
            End If
        Next
    End If
    If Len(sSumFilter) Then sSql = sSql & sSumFilter
    If Not bAddTB Then sSql = "SELECT * FROM"
    tSql.Text = sSql
    Screen.MousePointer = vbDefault
End Sub
Public Sub ExecuteSql(ByVal bTestOnly As Boolean)
        
    Dim lErr As Long, sErr As String
    Dim oQry As SQLDMO.QueryResults
    
    If Not bTestOnly Then
        With flexG(1)
            .Redraw = False
            .Clear
            .Rows = 1
            .FixedRows = 0
            .Cols = 1
            .Redraw = True
        End With
    End If
    Screen.MousePointer = vbHourglass

    If lErr = 0 And Len(tSql.Text) Then
        On Local Error Resume Next
        With oDb
            If bTestOnly Then
                .ExecuteImmediate tSql.Text, SQLDMOExec_ParseOnly
            Else
                Set oQry = .ExecuteWithResults(tSql.Text)
            End If
            lErr = Err.Number
            If lErr = 0 And Not bTestOnly Then LoadGrid oQry
            Set oQry = Nothing
        End With
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Or bTestOnly Then
        If bTestOnly And lErr = 0 Then sErr = MyLoadResString(k_Sql_Syntax_verified_successfully)
        
        MsgBox sErr, vbOKOnly Or vbInformation, App.EXEName
    End If

End Sub

Private Sub LoadGrid(oQry As SQLDMO.QueryResults)
    
    Dim iCols As Integer, lRows As Long
    Dim sBuf As String
    Dim x As Integer
    
    If Not oQry Is Nothing Then
        If oQry.ResultSets <> 0 Then
            
            With flexG(1)
                .Redraw = False
                oQry.CurrentResultSet = 1
                iCols = oQry.Columns
                lRows = oQry.Rows
                .Cols = iCols
                For x = 1 To iCols
                    .TextMatrix(0, x - 1) = oQry.ColumnName(x)
                Next
                
                If lRows > 0 Then
                    sBuf = oQry.GetRangeString(Top:=1, Bottom:=100, rowdelim:=vbCr, coldelim:=vbTab)
                    .Cols = iCols
                    .Rows = IIf(lRows > 100, 100, lRows) + 1
                    .Row = 1
                    .Col = 0
                    .RowSel = flexG(1).Rows - 1
                    .ColSel = flexG(1).Cols - 1
                    .Clip = sBuf
                    .RowSel = 0
                    .ColSel = 0
                End If
                .Row = 0
                .Col = 0
                Do While .Rows < 2
                    .AddItem ""
                Loop
                .Row = 1
                .ScrollBars = flexScrollBarBoth
                .AllowUserResizing = flexResizeColumns
                .FixedRows = 1
                
                .Redraw = True
                .TopRow = 1
            End With
        End If
    End If
        
End Sub

Private Sub SaveView()

    Dim objView As SQLDMO.View2
    Dim lErr As Long, sErr As String
    Dim sBuf As String
    Dim sOwner As String
    Dim sAttribute As String
    Dim sQuoted As String
    Dim bStatus As Boolean
    Dim bSetExplicit As Boolean
        
    If Not mbOldView Then
        bSetExplicit = cViewP.bOption(4)
        If chkDBO.Value Then
            sOwner = "dbo"
        Else
            sOwner = GetCreateOwner(sDbName)
        End If
        If cViewP.bOption(1) Then sAttribute = " ENCRYPTION "
        If cViewP.bOption(2) Then sAttribute = sAttribute & IIf(Len(sAttribute) <> 0, ",", "") & " SCHEMABINDING "
        If Len(sAttribute) Then sAttribute = " WITH" & sAttribute
        sBuf = "CREATE VIEW " & sOwner & "." & sStuff & vbCrLf & sAttribute & "AS" & vbCrLf
        Debug.Print sBuf
        Screen.MousePointer = vbHourglass
                
        
        If bSetExplicit Then
            bStatus = objServer.QuotedIdentifier
            sQuoted = Replace("SET QUOTED_IDENTIFIER 1%", "1%", "on")
            oDb.ExecuteImmediate sQuoted
        End If
        Set objView = New SQLDMO.View
        On Local Error Resume Next
        objView.Name = sStuff
        objView.Text = sBuf & tSql.Text
        objView.AnsiNullsStatus = cViewP.bOption(3)
        
        oDb.Views.Add objView
        lErr = Err.Number
        sErr = Err.Description
        If bSetExplicit Then
            sQuoted = Replace("SET QUOTED_IDENTIFIER 1%", "1%", IIf(bStatus, "on", "off"))
            oDb.ExecuteImmediate sQuoted
        End If

        On Local Error GoTo 0
    Else
        Beep
    End If
    
    Set objView = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        mbOldView = True
    End If
    
End Sub
Private Function GetResolvedObj(ByVal sTest4 As String, ByRef sAlias As String, Optional ByVal iEsclude As Integer = -1) As Boolean

    Dim l As Integer, i As Integer
    Dim bDone As Boolean, bRepeat As Boolean
    Dim sName As String, sAliased As String, sTmp As String
        
    sTmp = sAlias
    sTest4 = Mid$(sTest4, 2)
    sTest4 = Left$(sTest4, Len(sTest4) - 1)
    If Len(sAlias) Then
        sAlias = Mid$(sAlias, 2)
        sAlias = Left$(sAlias, Len(sAlias) - 1)
    End If
    
    Do
        bDone = True
        For i = 1 To collItem.Count
            If iEsclude <> i Then
                sName = Mid$(collItem(i).sName(), 2)
                sName = Left$(sName, Len(sName) - 1)
                sAliased = collItem(i).sAlias()
                If Len(sAliased) Then
                    sAliased = Mid$(sAliased, 2)
                    sAliased = Left$(sAliased, Len(sAliased) - 1)
                End If
                
                If Len(sAlias) = 0 Then
                    bRepeat = sTest4 = sName And Len(sAliased) = 0
                Else
                    bRepeat = (sTest4 = sAliased And sAlias = sName) Or (sTest4 = sName And sAlias = sAliased) Or (sAlias = sAliased)
                End If
                If bRepeat Then
                    l = l + 1
                    sAlias = sTest4 & "x" & Format$(l)
                    bDone = False
                    
                    Exit For
                End If
            End If
        Next
    Loop Until bDone
    If Len(sAlias) Then sAlias = "[" & sAlias & "]"
    GetResolvedObj = sTmp <> sAlias
End Function

Private Function GetAliasTB(ByVal sAliased As String) As String

    Dim i As Integer
    For i = 1 To collItem.Count
        If sAliased = collItem(i).sName() Then
            GetAliasTB = collItem(i).sAlias()
            Exit For
        End If
    Next
End Function
Private Function GetAliasedTbl(ByVal iControl As Integer) As String
    Dim i As Integer
    For i = 1 To collItem.Count
        If collItem(i).iControl = iControl Then
            GetAliasedTbl = collItem(i).sAlias
            Exit For
        End If
    Next
End Function

Private Sub ArrangeCtrl()

    Dim i As Integer

    Dim iX As Integer, iY As Integer
    Dim iW As Integer, iH As Integer
    Dim iPicW As Integer
    Dim iPicTH As Integer
    
    iPicTH = picTBL.ScaleHeight
    iPicW = picC.ScaleWidth

    iW = myLst(0).Width + 320
    iH = myLst(0).Height + 50
    
    iX = 320
    iY = 100
    For i = 1 To collItem.Count
        With myLst(collItem(i).iControl())
            .Move iX, iY
            .Visible = True
        End With
        iX = iX + iW
        
        If iX + iW > iPicW Then
            iX = 320
            iY = iY + iH
            If iY + iH > iPicTH Then
                If iY + (2 * iH) + 50 < 32000 Then
                    picTBL.Height = iY + iH + 50
                    VScr.Max = picTBL.ScaleHeight - picC.Height + 100
                End If
            End If
        End If
        
    Next
    DrawJoin
    
End Sub
Private Sub SetState(ByVal iItem As Integer, ByVal bVal As Boolean)
    tBar.Buttons(iItem).Value = Abs(bVal)
End Sub
Private Sub TestActiveControl()

    If cbo.Visible Then
        cbo_LostFocus
    ElseIf tFlex.Visible Then
        tFlex_LostFocus
    End If

   On Local Error Resume Next
   If TypeOf Me.ActiveControl Is TextBox Then
      If Me.ActiveControl Is tFlex Then
         tFlex_LostFocus
      End If
   ElseIf TypeOf Me.ActiveControl Is ComboBox Then
      If Me.ActiveControl Is cbo Then
         cbo_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub


