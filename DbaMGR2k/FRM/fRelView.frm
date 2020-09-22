VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fRelView 
   Caption         =   "-"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   Icon            =   "fRelView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCont 
      ClipControls    =   0   'False
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   8235
      TabIndex        =   7
      Top             =   600
      Width           =   8295
      Begin VB.PictureBox picPane 
         AutoRedraw      =   -1  'True
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   8115
         TabIndex        =   8
         Top             =   0
         Width           =   8175
         Begin DbaMGR2k.objRel objRel 
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   9
            Top             =   1560
            Visible         =   0   'False
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   450
         End
      End
   End
   Begin VB.HScrollBar HScr 
      Height          =   255
      Left            =   120
      Max             =   32000
      SmallChange     =   10
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.VScrollBar VScr 
      Height          =   1215
      Left            =   6600
      Max             =   32000
      SmallChange     =   10
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin MSComctlLib.ImageList imList 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fRelView.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fRelView.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fRelView.frx":05BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fRelView.frx":0718
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " "
            Key             =   "End"
            Description     =   "Run"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Description     =   "Open"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Index"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CheckBox chkRecur 
         Caption         =   "Mostra Oggetti Dipendenti"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4680
         ScaleHeight     =   375
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   0
         Width           =   3615
         Begin VB.ComboBox cbo 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   2895
         End
         Begin VB.Label lbl 
            Caption         =   "Object:"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "fRelView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIncY As Integer, iIncX As Integer
Private mlWidth As Long
Private mlHeight As Long
Private sDbName As String
Private iOldX As Integer, iOldY As Integer
Private objColRel As cRelDep

Dim oDB As SQLDMO.Database

Public Function ErrSetDBName(ByVal sVal As String, ByRef sErr As String) As Long
    
    Dim lErr As Long
    Dim oTb As SQLDMO.Table
    Dim oView As SQLDMO.View
    
    Screen.MousePointer = vbHourglass
    
    sDbName = sVal
    tBar.Buttons(keyIndex).Enabled = GetMembership(db_DDLAdmin, sDbName)
    On Local Error Resume Next
    Set oDB = objServer.Databases(sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        With cbo
            .Clear
            For Each oTb In oDB.Tables
                If Not oTb.SystemObject Then .AddItem oTb.Name
            Next
        End With
    End If
    
    Set oTb = Nothing
    Set oView = Nothing
    Screen.MousePointer = vbDefault
    
    ErrSetDBName = lErr
    If lErr = 0 Then Me.Caption = ReplaceMsg(MyLoadResString(k_RelGrapficalView), Array("1%"), Array(sDbName))
    
    
End Function
Private Sub InitLoad()

    Dim sBuf As String
    Dim i As Integer, iMax As Integer
    
    If cbo.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    picPane.Cls
    sBuf = "[" & cbo.Text & "]"
    
    iMax = objRel.Count

    On Local Error Resume Next
    For i = iMax To 1 Step -1
        Unload objRel(i)
    Next
    On Local Error GoTo 0
    
    Set objColRel = New cRelDep
    LoadCtrlFromTable objColRel, sBuf, 0, Nothing
    ShowCtrl objColRel
    PaintLines objColRel
    HScr.Value = 0
    VScr.Value = 0
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub LoadCtrlFromTable(ByRef initObj As cRelDep, ByVal sTableName As String, ByRef iObjIdx As Integer, ByRef RootObj As cRelDep)

    
    Dim i As Integer, l As Integer
    Dim iItem As Integer
    Dim oQry As SQLDMO.QueryResults
    Dim oColl As Collection
    Dim oCol As SQLDMO.Column
    Dim oTb As SQLDMO.Table
    Dim oK As SQLDMO.Key
    Dim iPtrRefereced As Integer
    Dim bRecursive As Boolean
    
    
        
'    ' blocca recursione
'    iStack = iStack + 1
'    If iStack = 8 Then
'        iStack = iStack - 1
'        Exit Sub
'    End If
    
    ' Modello ad oggetti relativo
    ' ===========================
    '
    '
    '   Rel0 (base)
    '   =====
    '   .TbName
    '   .RelName
    '   .PtrCtrl
    '   .oColl    ----- Rel1
    '               |   ====
    '               |   .TbName
    '               |   .RelName
    '               |   .PtrCtrl
    '               |   .oColl      ----------------------  RelN+x
    '               |                                       ====
    '               |                                       .TbName
    '               |                                       .RelName
    '               |                                       .PtrCtrl
    '               |                                       .oColl      ---- RelN+x
    '               |
    '               |-- Rel2
    '               |   ====
    '               |   .TbName
    '               |   .RelName
    '               |   .PtrCtrl
    '               |   .oColl      ----------------------  RelN+x
    '               |                                       ====
    '               |                                       .TbName
    '               |                                       .RelName
    '               |                                       .PtrCtrl
    '               |                                       .oColl      ---- RelN+x
    
    
    
    'Set oQry = oDB.Tables(sTableName, GetTableOwner(objServer, sDbName, sTableName)).EnumReferencingKeys(includeallcandidates:=False)
    
    Set oTb = GetTbByName(objServer, sDbName, sTableName)
    Set oQry = oTb.EnumReferencingKeys(includeallcandidates:=False)
    Set oColl = LoadCollFromQRY(oQry)
    Set oQry = Nothing

    If oColl.Count Then
        'Debug.Print oDB.Tables(sTableName).Name
        initObj.TbName = sTableName
        
        'Set oTB = oDB.Tables(sTableName)
        Set oTb = GetTbByName(objServer, sDbName, sTableName)
        
        If initObj Is objColRel Then
            iObjIdx = iObjIdx + 1
            initObj.PtrCtrl() = iObjIdx
            initObj.TbName() = sTableName
            
            Load objRel(iObjIdx)
            With objRel(iObjIdx)
            
                .ToolTipCaption = ReplaceMsg(MyLoadResString(k_RelOrigTbl), Array("1%"), Array(sTableName))
                .ToolTipFields = MyLoadResString(k_RelReferencedTblCols)
                .ToolTipToggle = MyLoadResString(k_RelShowDetails)
                '.ObjName("") = sTableName
                .ObjName("") = oTb.Owner & "." & sTableName
                For Each oCol In oTb.Columns
                    .AddItem oCol.Name
                Next
            End With
        End If
        iPtrRefereced = initObj.PtrCtrl
        
        For i = 1 To oColl.Count
            
            'Debug.Print oColl.Item(i).TbName(), initObj.TbName(), StrComp(oColl.Item(i).TbName(), initObj.TbName(), vbTextCompare) <> 0
            'bRecursive = StrComp(oColl.Item(i).TbName(), initObj.TbName(), vbTextCompare) = 0
            'Debug.Print oColl.Item(i).TbName
            sTableName = oColl.Item(i).TbName 'initObj.CollItems(iItem).TbName()
            
            RecursiveScan RootObj, sTableName, bRecursive
            
            initObj.CollItems.Add oColl.Item(i)
            iItem = initObj.CollItems.Count
            
            
            'Set oTB = oDB.Tables(sTableName)
            Set oTb = GetTbByName(objServer, sDbName, sTableName)
                        
            iObjIdx = iObjIdx + 1
            initObj.CollItems(iItem).PtrCtrl() = iObjIdx
            Load objRel(iObjIdx)
            With objRel(iObjIdx)
                .ObjName(initObj.CollItems(iItem).RelName()) = GetTableOwner(objServer, sDbName, initObj.CollItems(iItem).TbName()) & "." & initObj.CollItems(iItem).TbName()
                .IsSecondary() = True
                .ReferencedCtrl() = iPtrRefereced
                Set oK = oTb.Keys(initObj.CollItems(iItem).RelName())
                
                .ToolTipCaption = ReplaceMsg(MyLoadResString(k_RelReferencedTbl), Array("1%"), Array(sTableName)) & IIf(bRecursive, MyLoadResString(k_RelSelfRefer), "")
                .ToolTipOwner = ReplaceMsg(MyLoadResString(k_RelName), Array("1%"), Array(oK.Name)) & IIf(bRecursive, MyLoadResString(k_RelSelfRefer), "")
                .ToolTipFields = MyLoadResString(k_RelReferencingTblCols) & IIf(bRecursive, MyLoadResString(k_RelSelfRefer), "")
                .ToolTipToggle = MyLoadResString(k_RelShowDetails)
                For Each oCol In oTb.Columns
                    .AddItem oCol.Name
                Next
                Set oK = oTb.Keys(initObj.CollItems(iItem).RelName())
                For l = 1 To oK.KeyColumns.Count
                    .AddKeyColItem oK.KeyColumns(l)
                Next
                For l = 1 To oK.ReferencedColumns.Count
                    .AddKeyReferences oK.ReferencedColumns(l)
                Next
                Set oK = Nothing
            End With
            
            If chkRecur.Value And Not bRecursive Then LoadCtrlFromTable initObj.CollItems(iItem), sTableName, iObjIdx, IIf(RootObj Is Nothing, initObj, RootObj)
            
            Set oTb = Nothing
        Next
    End If
   
    Set oColl = Nothing
    Set oCol = Nothing
    Set oK = Nothing
    Set oTb = Nothing
   
End Sub

Private Sub cbo_Click()

    If cbo.ListCount <> -1 Then tBar.Buttons(keyIndex).ToolTipText = ReplaceMsg(MyLoadResString(k_RelManageIdx), Array("1%"), Array(cbo.Text))
End Sub

Private Sub Form_Load()

    mlWidth = Me.Width
    mlHeight = Me.Height
    iIncY = 1
    iIncX = 1
    picPane.ToolTipText = MyLoadResString(k_RelPaneMSG)
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keyRun).ToolTipText = MyLoadResString(k_RelSearch)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_Refresh)
    tBar.Buttons(keyIndex).ToolTipText = MyLoadResString(k_RelSelectTable)
        
    cbo.ToolTipText = MyLoadResString(k_RelSelectTB)
    chkRecur.Caption = MyLoadResString(k_RelRecurviveSearch)
    chkRecur.ToolTipText = MyLoadResString(k_RelRecurviveShow)
    lbl(0).Caption = MyLoadResString(k_Table)
   
    CenterForm Me
End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        b = True
        If Me.Width < mlWidth Then Me.Width = mlWidth
        If Me.Height < mlHeight Then Me.Height = mlHeight
            
        SizeControls
        
        b = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objColRel = Nothing
    Set oDB = Nothing
    Set fRelView = Nothing
End Sub
Private Sub SizeControls()
        
    picCont.Width = Me.ScaleWidth - VScr.Width
    picCont.Height = Me.ScaleHeight - picCont.Top - HScr.Height
    
    If picCont.ScaleWidth > picPane.Width Then picPane.Width = picCont.ScaleWidth
    If picCont.ScaleHeight > picPane.Height Then picPane.Height = picCont.ScaleHeight
    
    VScr.Move picCont.Width, picCont.Top, VScr.Width, picCont.Height
    HScr.Move picCont.Left, picCont.Top + picCont.Height, picCont.Width
    
    
    HScr.Max = picPane.Width - picCont.ScaleWidth
    VScr.Max = picPane.Height - picCont.ScaleHeight
    
    Picture1.Width = Me.ScaleWidth - Picture1.Left - k_MargineFRM
    cbo.Width = Picture1.ScaleWidth - cbo.Left - k_FlexBevel
    
    PaintLines objColRel

End Sub

Private Sub objRel_Click(index As Integer)

    Dim lMaxY As Integer
    lMaxY = objRel(index).Top + objRel(index).Height + k_MargineFRM
    If lMaxY > picPane.ScaleHeight Then picPane.Height = lMaxY
    SizeControls
End Sub

Private Sub objRel_RequestMenu(index As Integer, ByVal x As Integer, ByVal y As Integer)

    Dim i As Integer
    Dim P As POINTAPI
    Dim M As New cPUpMenu
    
    M.Add MyLoadResString(k_Indexes_Management), 1, bDisable:=Not GetMembership(db_DDLAdmin, sDbName)
    Call GetCursorPos(P)
    i = M.Popup(Me.hWnd, P.x, P.y)
    M.Clear
    Set M = Nothing
    
    If i <> 0 Then RunRelation objRel(index).TableName
    
End Sub

Private Sub VScr_Change()
    vScr_Scroll
End Sub

Private Sub vScr_Scroll()
    picPane.Top = -iIncY * VScr.Value
End Sub

Private Sub HScr_Change()
    HScr_Scroll
End Sub
Private Sub HScr_Scroll()
    picPane.Left = -iIncX * HScr.Value
End Sub

Private Sub objRel_BeginMove(index As Integer, ByVal x As Integer, ByVal y As Integer)
    iOldX = x
    iOldY = y
    objRel(index).ZOrder
End Sub

Private Sub objRel_EndMoving(index As Integer, ByVal x As Integer, ByVal y As Integer)
    iOldX = 0
    iOldY = 0
    
    PaintLines objColRel
End Sub

Private Sub objRel_Moving(index As Integer, ByVal x As Integer, ByVal y As Integer)

    Dim iX As Integer, iY As Integer
    
    With objRel(index)
        iX = .Left + x - iOldX
        iY = .Top + y - iOldY
        If iX < 0 Then
            iX = 0
        ElseIf iX + .Width > picPane.ScaleWidth Then
            iX = picPane.ScaleWidth - .Width
        End If
        If iY < 0 Then
            iY = 0
        ElseIf iY + .Height > picPane.ScaleHeight Then
            iY = picPane.ScaleHeight - .Height
        End If
        .Move iX, iY
    End With

End Sub

Private Sub objRel_Selected(index As Integer)
    
    Dim i As Integer, iRefer As Integer, iMax As Integer
    Dim oCol As Collection
    objRel(index).ZOrder
    
    Screen.MousePointer = vbHourglass
    
    iMax = objRel.Count
    On Local Error Resume Next
    For i = iMax To 1 Step -1
        objRel(i).Clear
    Next
    On Local Error GoTo 0
    If index > 1 Then
        iRefer = objRel(index).ReferencedCtrl
        Set oCol = objRel(index).ReferencedCollection
        For i = 1 To oCol.Count
            objRel(iRefer).CheckThis oCol.Item(i)
        Next
        Set oCol = objRel(index).KeysCollection
        For i = 1 To oCol.Count
            objRel(index).CheckThis oCol.Item(i)
        Next
    End If
    DoEvents

    PaintLines objColRel
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case keyEnd
            Unload Me
        Case keyRefresh
            RefreshPic
        Case keyRun
            InitLoad
        Case keyIndex
            If cbo.ListIndex <> -1 Then RunRelation cbo.Text
                
    End Select
End Sub
Private Sub RefreshPic()
        
    Screen.MousePointer = vbHourglass
    ShowCtrl objColRel
    SizeControls
    PaintLines objColRel
    Screen.MousePointer = vbDefault

End Sub
Private Function LoadCollFromQRY(oQry As SQLDMO.QueryResults) As Collection

    Dim x As Integer
    Dim lRows As Long, l As Long
    Dim oObj As cRelDep
    Dim sBuf As String
    Dim sOwner As String
    Dim iPos As Integer
    
    Dim retColl As New Collection
    If oQry.ResultSets <> 0 Then
        For x = 1 To oQry.ResultSets
            oQry.CurrentResultSet = x
            lRows = oQry.Rows
            For l = 1 To lRows
                If oQry.GetColumnBool(l, 3) Then
                    Set oObj = New cRelDep
                    sBuf = oQry.GetColumnString(l, 1)
                    iPos = InStr(sBuf, ".")
                    If iPos <> 0 Then
                        sOwner = Left$(sBuf, iPos - 1)
                        sBuf = Mid$(sBuf, iPos + 1)
                    End If
                    oObj.TbName() = sBuf 'oQry.GetColumnString(l, 1)
                    oObj.RelName() = oQry.GetColumnString(l, 2)
                    
                    On Local Error Resume Next
                    retColl.Add oObj, oObj.RelName()
                    On Local Error GoTo 0
                End If
            Next
        Next
    End If
    Set oObj = Nothing
    
    Set LoadCollFromQRY = retColl
             
End Function

Private Sub ShowCtrl(initObj As cRelDep)

    Dim i As Integer
    Dim bInit As Boolean
    Dim lY As Integer, lX As Integer
    Dim lW As Integer, lH As Integer
    Dim lMaxX As Integer, lMaxY As Integer
    Dim iPrev As Integer
        
    If initObj Is Nothing Then Exit Sub
    bInit = initObj Is objColRel
    lW = objRel(0).Width

    If initObj.PtrCtrl() = 0 Then Exit Sub
    If bInit Then
        Screen.MousePointer = vbHourglass
        objRel(initObj.PtrCtrl()).Move 200, 200
        objRel(initObj.PtrCtrl()).Visible = True
        objRel(initObj.PtrCtrl()).ZOrder
    End If
    
    
    lX = objRel(initObj.PtrCtrl()).Left
    lY = objRel(initObj.PtrCtrl()).Top
    
    For i = 1 To initObj.CollItems.Count
        iPrev = initObj.CollItems(i).PtrCtrl() - 1
        If i = 1 Then
            lH = objRel(iPrev).Top
        Else
            lH = objRel(iPrev).Top + objRel(iPrev).Height + (2 * k_MargineFRM)
        End If
        
        objRel(iPrev + 1).Move lX + lW + (2 * k_MargineFRM), lH
        objRel(iPrev + 1).Visible = True
        objRel(iPrev + 1).ZOrder
        lMaxX = objRel(iPrev + 1).Left + objRel(iPrev + 1).Width + k_MargineFRM
        lMaxY = objRel(iPrev + 1).Top + objRel(iPrev + 1).Height + k_MargineFRM
        If lMaxX > picPane.ScaleWidth Then picPane.Width = lMaxX
        If lMaxY > picPane.ScaleHeight Then picPane.Height = lMaxY
        ShowCtrl initObj.CollItems(i)
    Next
    
    If bInit Then
        SizeControls
        Screen.MousePointer = vbDefault
    End If
    
End Sub
Private Sub PaintLines(ByVal initObj As cRelDep)

    Dim i As Integer
    Dim iX As Integer, iY As Integer
    Dim zX As Integer, zY As Integer
    

    If initObj Is objColRel Then picPane.Cls
    If initObj Is Nothing Then Exit Sub

    iX = objRel(initObj.PtrCtrl).Left + objRel(initObj.PtrCtrl).Width
    iY = objRel(initObj.PtrCtrl).Top + 120

    For i = 1 To initObj.CollItems.Count
        zX = objRel(initObj.CollItems(i).PtrCtrl).Left
        zY = objRel(initObj.CollItems(i).PtrCtrl).Top + 120
        picPane.Line (iX, iY)-(zX, zY)
        
        PaintLines initObj.CollItems(i)
    Next
    
End Sub

Private Sub RunRelation(ByVal sTbName As String)

    fIndex.DataNames sDbName, sTbName
    fIndex.Caption = fIndex.Caption & MyLoadResString(k_for_Table) & "  '" & sTbName & "' - Database '" & sDbName & "'"
    
    fIndex.Show vbModal, Me

End Sub
Private Sub RecursiveScan(ByRef RootObj As cRelDep, ByVal sTb2Test As String, ByRef bRecursive As Boolean)

    ' Modello ad oggetti relativo
    ' ===========================
    '
    '
    '   Rel0 (base)
    '   =====
    '   .TbName
    '   .RelName
    '   .PtrCtrl
    '   .oColl    ----- Rel1
    '               |   ====
    '               |   .TbName
    '               |   .RelName
    '               |   .PtrCtrl
    '               |   .oColl      ----------------------  RelN+x
    '               |                                       ====
    '               |                                       .TbName
    '               |                                       .RelName
    '               |                                       .PtrCtrl
    '               |                                       .oColl      ---- RelN+x
    '               |
    '               |-- Rel2
    '               |   ====
    '               |   .TbName
    '               |   .RelName
    '               |   .PtrCtrl
    '               |   .oColl      ----------------------  RelN+x
    '               |                                       ====
    '               |                                       .TbName
    '               |                                       .RelName
    '               |                                       .PtrCtrl
    '               |                                       .oColl      ---- RelN+x

    Dim oColl As Collection
    Dim sTblMaster As String
    Dim bRet As Boolean
    Dim i As Integer
    
    If Not RootObj Is Nothing Then
        sTblMaster = RootObj.TbName()
        sTblMaster = Replace(sTblMaster, "[", "")
        sTblMaster = Replace(sTblMaster, "]", "")
        
        sTb2Test = Replace(sTb2Test, "[", "")
        sTb2Test = Replace(sTb2Test, "]", "")
        
        If StrComp(sTblMaster, sTb2Test, vbTextCompare) = 0 Then
            bRet = True
        Else
            Set oColl = RootObj.CollItems
            For i = 1 To oColl.Count
                RecursiveScan oColl.Item(i), sTb2Test, bRet
                If bRet Then Exit For
            Next
            Set oColl = Nothing
        End If
    End If
    
    bRecursive = bRet
        
End Sub

