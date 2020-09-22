VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fIdxView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7200
   ClipControls    =   0   'False
   Icon            =   "fIdxView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Mod. SQL"
      Height          =   360
      Index           =   4
      Left            =   5400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox cIndex 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   315
      Index           =   5
      Left            =   6720
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4965
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Height          =   360
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Nuovo"
      Height          =   360
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
      Height          =   360
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
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
      Left            =   1200
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label lbl 
      Caption         =   "Elemento Selezionato"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblNewI 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "fIdxView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type idxElem
    pgElement As PageElement
    EditType As TypeEnum
End Type
Private m_blnIdxProp As Boolean
Private m_sDbName As String
Private m_sViewName As String
Private miFlex As Integer       'Indice Flex
Private miRow(1) As Integer        'Riga Flex
Private miCol As Integer
Private m_blnAnnullaMod As Boolean
Private bLoading As Boolean

Private m_oView As SQLDMO.View2
Private Elem(8) As idxElem 'PageElement
Private native(8) As Long
Private colIdxCol As Collection
Private allColumn As Collection
Private bNewIndex As Boolean
Private m_sOwner As String
Private m_bIsReadOnly As Boolean
Private m_bIsOwner As Boolean
Private m_lHeight As Long
Private m_lWidth As Long

Private Sub Load_Elem()
        
    native(0) = SQLDMOIndex_Clustered
    native(1) = SQLDMOIndex_Unique
    native(2) = SQLDMOIndex_IgnoreDupKey
    native(3) = SQLDMOIndex_NoRecompute
    
        
    native(4) = SQLDMOIndex_PadIndex
    native(5) = -1  'FillFactor
    
    native(6) = -2  'SpaceUsed
    native(7) = -3 'FileGroup
    native(8) = SQLDMOIndex_Default
    
    
    Elem(0).pgElement.sName = "Clustered"
    Elem(0).pgElement.iType = am_Bool
    
    Elem(1).pgElement.sName = "Unique"
    Elem(1).pgElement.iType = am_Bool
    
    Elem(2).pgElement.sName = "IgnoreDupKey"
    Elem(2).pgElement.iType = am_Bool
    
    Elem(3).pgElement.sName = MyLoadResString(k_Not_refresh_statistics)
    Elem(3).pgElement.iType = am_Bool

    
    Elem(4).pgElement.sName = "PadIndex"
    Elem(4).pgElement.iType = am_Bool
    
    Elem(5).pgElement.sName = "Fill Factor"
    Elem(5).pgElement.iType = am_Long
    Elem(5).pgElement.iLen = 3
    
        
    Elem(6).pgElement.sName = "Space Used Kb"
    Elem(6).pgElement.iType = am_Long
    Elem(6).pgElement.iLen = -1

    Elem(7).pgElement.sName = MyLoadResString(k_HtmFileGroup) '"File Group"
    Elem(7).pgElement.iType = am_String
    Elem(7).EditType = iCombo
    Elem(7).pgElement.iLen = -1

    Elem(8).pgElement.sName = "Default"
    Elem(8).pgElement.iType = am_Bool
    Elem(8).pgElement.iLen = -1
          
End Sub

Public Sub DataNames(ByVal sDB As String, ByVal sViewName As String)
    
    m_sDbName = sDB
    m_sViewName = sViewName
    
    
    Dim i As Integer
    Dim iC As Integer
    Dim oList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column
    
    m_bIsReadOnly = Not IsDbAvailableReadWrite(m_sDbName)
    Set m_oView = GetViewByName(objServer, m_sDbName, m_sViewName)
    If Not m_oView Is Nothing Then
        m_sOwner = m_oView.Owner
        m_bIsOwner = IsObjectOwner(m_sDbName, m_sViewName, am_OwnView)
    End If
    
    
    Set allColumn = New Collection
    Set oList = m_oView.ListColumns
    For Each oCol In oList
        allColumn.Add oCol.Name
    Next
    Set oCol = Nothing
    Set oList = Nothing
    
    miCol = 0
            
    cbo.Visible = False
    
    LoadIndex ""
    
    SizeControls
    If m_bIsReadOnly Then
        cmd(1).Enabled = False
        cmd(2).Enabled = False
        cmd(3).Enabled = False
        cmd(4).Visible = False
    End If
    
    
End Sub
Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub cbo_LostFocus()

    Dim obj As cColIDX
            
    If Not m_blnAnnullaMod Then
        If Not m_blnIdxProp Then
            Debug.Print Flex(0).TextMatrix(miRow(0), miCol)
            Flex(0).TextMatrix(miRow(0), miCol) = cbo.Text
            If miCol = 0 Then
                If colIdxCol.Count >= miRow(0) Then
                    Set obj = colIdxCol.Item(miRow(0))
                Else
                    m_blnAnnullaMod = True
                End If
            Else
                m_blnAnnullaMod = True
            End If
            
            If Not m_blnAnnullaMod Then
                obj.colName() = cbo.Text
                If Len(obj.colName) = 0 Then
                    obj.isDescending = False
                    bLoading = True
                    Flex(0).Row = miRow(0)
                    Flex(0).Col = 1
                    Set Flex(0).CellPicture = LoadResPicture(IIf(obj.isDescending(), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex(0).CellPictureAlignment = flexAlignCenterCenter
                    
                    Flex(0).Col = 0
                    Flex(0).Refresh
                    bLoading = False
                End If
            End If
        Else
            Flex(1).TextMatrix(miRow(1) + 1, 1) = cbo.Text
            Elem(miRow(1)).pgElement.vVal = cbo.Text
        End If
    End If
    Set obj = Nothing
    
    cbo.Visible = False
    m_blnAnnullaMod = True
    m_blnIdxProp = False

End Sub
Private Sub cIndex_Click()
    LoadIndex cIndex.List(cIndex.ListIndex)
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            Unload Me
        Case 1  'Save
            TestActiveControl
            Flex(0).SetFocus
            SaveIX
        Case 2  'New
            cIndex.SetFocus
            If Not bNewIndex Then
                SetNewIndex
            Else
                Beep
            End If
        Case 3  'Drop
            If bNewIndex Then
                bNewIndex = False
                LoadIndex ""
            Else
                DropIndex
            End If
        Case 4  'Modify Sql
            Flex(1).SetFocus
            ModifySQL
        Case 5  'Extended Properties
            Flex(0).SetFocus
            ManageExtendedProp
    End Select

End Sub
Private Sub SetNewIndex()

    Dim bConfirm As Boolean
    Dim sBuf As String
    Dim i As Integer
    Dim oCol As cColIDX
    
    Screen.MousePointer = vbDefault
    fPwd.Action() = act_Tbl
    
    fPwd.Caption = MyLoadResString(k_New_Index) & "..."
    fPwd.Stuff = "IX_" & UCase$(m_sViewName)
    
    fPwd.Show vbModal, Me
    bConfirm = fPwd.Confirm()
    If bConfirm Then sBuf = fPwd.Stuff()
    Unload fPwd
    
    If bConfirm Then
        bNewIndex = True
        lblNewI.Caption = sBuf
        lblNewI.Visible = True
        cmd(4).Visible = True
        cmd(5).Enabled = False
        Set colIdxCol = New Collection
    
        For i = 0 To UBound(native)
            Select Case native(i)
                Case -1
                    Elem(i).pgElement.vVal = 0
                Case -2
                    Elem(i).pgElement.vVal = 0
                Case -3
                    Elem(i).pgElement.vVal = "PRIMARY"
                    
                Case Else
                    Elem(i).pgElement.vVal = 0
            End Select
        Next
        
        For i = 1 To allColumn.Count
            Set oCol = New cColIDX
            colIdxCol.Add oCol
        Next
        LoadGrids 0
        LoadGrids 1
        SizeControls
        miRow(0) = 0
        miCol = 0
        Flex(0).Row = 1
        
        Flex(0).SetFocus
        Flex_RowColChange 0
        cmd(1).Visible = True
        cmd(2).Visible = False

    End If
    
End Sub
Private Sub DropIndex()
    
    Dim sBuf As String, sChar As String
'    Dim i As Integer
    Dim iElem As Integer
    Dim lErr As Long, sErr As String
    
    If cIndex.ListCount Then
        sBuf = cIndex.Text
        iElem = cIndex.ListIndex
        sChar = MyLoadResString(k_Indexes_Keys_constraints)
        
        If MsgBox(MyLoadResString(k_Confirm_Remove) & vbLf & sChar & " " & sBuf, vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then

            On Local Error Resume Next
            m_oView.Indexes.Remove sBuf
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            If lErr = 0 Then
                cIndex.RemoveItem iElem
                If cIndex.ListCount Then
                    cIndex.ListIndex = 0
                Else
                    cIndex.ListIndex = -1
                    LoadIndex ""
                End If
            Else
                MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
            End If
        End If
    End If
End Sub

Private Sub Flex_DblClick(Index As Integer)

    Dim iElem As Integer
    Dim iImg As Integer
    Dim obj As cColIDX
    
    If Index Then
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then
            FlexEdit vbKeySpace
        End If
    Else
        If bNewIndex Then
            With Flex(0)
                If .Col < .Cols - 1 Then
                    LoadCbo .Row
                    If cbo.ListCount Then cbo.ListIndex = 0
                    miRow(0) = Flex(Index).Row
                    miCol = Flex(Index).Col
                    
                    cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                    cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                    cbo.ZOrder
                    cbo.Visible = True
                    cbo.SetFocus
                Else
                    iElem = Flex(0).Row
                    If colIdxCol.Count >= iElem Then
                        Set obj = colIdxCol.Item(iElem)
                        If Len(obj.colName) Then
                            iImg = 1 - (-CBool(obj.isDescending))
                        
                            obj.isDescending = iImg
                            Set Flex(0).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        End If
                    End If
                End If
            End With
        End If
    End If

End Sub
Private Sub Flex_GotFocus(Index As Integer)

    TestActiveControl
    If Index Then
        miRow(1) = Flex(1).Row - 1
        If bNewIndex Then Flex_RowColChange Index
    Else
        miRow(0) = Flex(0).Row
    End If
    m_blnAnnullaMod = True
    
End Sub
Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    
    If Index Then
        iCode = KeyAscii
        KeyAscii = 0
        FlexEdit iCode
    Else
        If bNewIndex Then
            KeyAscii = 0
            Flex_DblClick Index
        End If
    End If
End Sub
Private Sub Flex_RowColChange(Index As Integer)

    Static b As Boolean

    If (Index = 0) And (bNewIndex = True) And Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            b = False
        End If
    End If

End Sub
Private Sub Flex_Scroll(Index As Integer)
    cbo.Visible = False
    tFlex.Visible = False
    TestActiveControl
End Sub

Private Sub Form_Initialize()
    Load_Elem
End Sub
Private Sub Form_Load()
    m_lHeight = Me.Height
    m_lWidth = Me.Width

    Me.Caption = MyLoadResString(k_Indexes)

    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    cmd(2).Caption = MyLoadResString(k_New)
    cmd(3).Caption = MyLoadResString(k_Remove)
    cmd(4).Caption = MyLoadResString(k_ModifySQL)
    cmd(5).ToolTipText = MyLoadResString(k_Extended_Property)
    lbl(0).Caption = MyLoadResString(k_Selected_Object)
    CenterForm Me
    Flex(0).RowHeightMin = cbo.Height
    Flex(1).RowHeightMin = cbo.Height

End Sub
Private Sub Form_Resize()
    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            SizeControls
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set colIdxCol = Nothing
    Set m_oView = Nothing
    Set fIdxView = Nothing
End Sub

Private Sub SizeControls()

    Dim l As Long, i As Integer
    
    l = Me.ScaleWidth \ 4 - cmd(0).Width
    cmd(0).Top = Me.ScaleHeight - (cmd(0).Height + 50)
    
    For i = 1 To 4
        cmd(i).Top = cmd(0).Top
        cmd(i).Left = cmd(0).Left + i * (cmd(0).Width + 25)
    Next
    cmd(5).Top = cmd(0).Top + cmd(0).Height - cmd(5).Height
    cmd(5).Left = cmd(0).Left + i * (cmd(0).Width + 25)
    
    Flex(0).Width = Me.ScaleWidth - (2 * Flex(0).Left)
    Flex(1).Top = Flex(0).Height + 25 + Flex(0).Top
    Flex(1).Width = Flex(0).Width
    
    lblNewI.Width = Me.ScaleWidth \ 2 - 120
    lblNewI.Left = Me.ScaleWidth \ 2
    
    With Flex(0)
        .Height = ((cmd(0).Top - .Top - 200) \ 2) - k_FlexBevel
        .ColWidth(0) = .Width \ 2
    End With
            
    With Flex(1)
        .Top = Flex(0).Top + Flex(0).Height + (k_FlexBevel \ 2)
        .Height = Flex(0).Height
        .ColWidth(0) = .Width / 2
        .ColWidth(1) = .ColWidth(0) - 350
    End With
    
    cIndex.Width = Me.ScaleWidth - lbl(0).Width - lbl(0).Left - 200
    SettaLargFlex Flex(0), 0, 200
    SettaLargFlex Flex(1), 0, 200
    
End Sub
Private Sub LoadIndex(ByVal sIndexName As String)

    Dim oIndex As SQLDMO.Index2
    'Dim oList As SQLDMO.SQLObjectList
    Dim oCol As SQLDMO.Column
    Dim objCol As cColIDX

    Dim i As Integer
    
    
    lblNewI.Visible = False
    lblNewI.Caption = ""
    bNewIndex = False
    
    cmd(1).Visible = False
    cmd(2).Visible = True
    cmd(4).Visible = False
    cmd(5).Enabled = False
            
    Screen.MousePointer = vbHourglass
    If Len(sIndexName) = 0 Then
        Set m_oView = Nothing
        
        Set m_oView = GetViewByName(objServer, m_sDbName, m_sViewName)
        m_oView.Indexes.Refresh True
    End If
    
    
    If Len(sIndexName) <> 0 Then
        cmd(5).Enabled = m_bIsOwner And Not bNewIndex
        
        On Local Error GoTo exitRoutine
        Set oIndex = m_oView.Indexes(sIndexName)
        With oIndex
            'Set oList = oIndex.ListIndexedColumns
            Set colIdxCol = New Collection
            
            For Each oCol In oIndex.ListIndexedColumns
                Set objCol = New cColIDX
                objCol.colName() = oCol.Name
                objCol.isDescending() = oIndex.GetIndexedColumnDESC(oCol.Name)
                colIdxCol.Add objCol
            Next
            Set objCol = Nothing
            
            For i = 0 To UBound(native)
                Select Case native(i)
                    Case -1
                        Elem(i).pgElement.vVal = oIndex.FillFactor
                    Case -2
                        Elem(i).pgElement.vVal = oIndex.SpaceUsed
                    Case -3
                        Elem(i).pgElement.vVal = oIndex.FileGroup
                    Case Else
                        Elem(i).pgElement.vVal = (.Type And native(i)) = native(i)
                End Select
            Next
        End With
        LoadGrids 0
        LoadGrids 1
    Else
        LoadGrids -1
        Set colIdxCol = Nothing
        cIndex.Clear
        Debug.Print m_oView.Indexes.Count
        For Each oIndex In m_oView.Indexes
            Debug.Print (oIndex.Type And SQLDMOIndex_Hypothetical) = SQLDMOIndex_Hypothetical
            'If (oIndex.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex Then cIndex.AddItem oIndex.Name
            If ((oIndex.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex) And ((oIndex.Type And SQLDMOIndex_Hypothetical) <> SQLDMOIndex_Hypothetical) Then cIndex.AddItem oIndex.Name
        Next
    End If
    
finish:
    
    Set objCol = Nothing
    'Set oList = Nothing
    Set oCol = Nothing
    Set oIndex = Nothing
    If Len(sIndexName) = 0 Then
        If cIndex.ListCount Then cIndex.ListIndex = 0
    End If
    SizeControls
    Screen.MousePointer = vbDefault
    Exit Sub

exitRoutine:
    Debug.Print Err.Description
    Resume finish

End Sub
Private Sub LoadGrids(ByVal Index As Integer)
        
    Dim i As Integer
    Dim sBuf As String
    Dim obj As cColIDX

    bLoading = True
    
    If Index <> -1 Then
        With Flex(Index)
            .Redraw = False
            .ScrollBars = flexScrollBarBoth
            .Clear
            .FixedCols = 0
            .FixedRows = 0
            
            .Cols = 2
            .Rows = 1
            If Index = 1 Then .ColAlignment(1) = flexAlignLeftCenter
            
            .ScrollBars = flexScrollBarBoth
            If Index = 0 Then
                .TextMatrix(0, 0) = MyLoadResString(k_Indexed_Columns)
                .TextMatrix(0, 1) = MyLoadResString(k_Descending)
            Else
                .TextMatrix(0, 0) = MyLoadResString(k_Property)
                .TextMatrix(0, 1) = MyLoadResString(k_Value)
            End If
            .Redraw = True
            
            If Index = 0 Then
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    .AddItem obj.colName
                    .Row = .Rows - 1
                    .Col = 0
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                    .CellPictureAlignment = flexAlignRightTop
                
                    .Col = 1
                    Set Flex(Index).CellPicture = LoadResPicture(IIf(obj.isDescending(), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                Next
                Set obj = Nothing
            Else
                For i = 0 To UBound(Elem)
                    sBuf = Elem(i).pgElement.sName
                    Select Case Elem(i).pgElement.iType
                        Case am_Bool
                        Case am_Long
                            sBuf = sBuf & vbTab & Format$(Elem(i).pgElement.vVal, "0")
                        Case am_String
                            sBuf = sBuf & vbTab & Elem(i).pgElement.vVal
                    End Select
                    .AddItem sBuf
                    If Elem(i).pgElement.iType = am_Bool Then
                        .Row = .Rows - 1
                        .Col = .Cols - 1
                        Debug.Print Elem(i).pgElement.vVal
                        Set Flex(Index).CellPicture = LoadResPicture(IIf(Abs(Elem(i).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                    End If
                    If Elem(i).EditType = iCombo Then
                        .Row = i + 1
                        .Col = .Cols - 1
                        Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                        .CellPictureAlignment = flexAlignRightTop
                        
                    End If
                    If i > 4 And i <> 7 Then
                        .Row = i + 1
                        .Col = .Cols - 1
                        .CellBackColor = vbInactiveBorder
                    End If
                Next
            End If
            
            .FixedCols = Index
            .AllowUserResizing = flexResizeColumns
            .ScrollBars = flexScrollBarBoth
            If .Rows > 1 Then .FixedRows = 1
            .Row = 1
            .Redraw = True

            .TopRow = 1
            .Col = .FixedCols
            If Index = 1 Then SetEnabled
        End With
    Else
        Flex(0).FixedRows = 0
        Flex(1).FixedRows = 0
        Flex(0).Clear
        Flex(1).Clear
        Flex(0).Rows = 1
        Flex(0).Cols = 2
        Flex(0).TextMatrix(0, 0) = MyLoadResString(k_Indexed_Columns)
        Flex(0).TextMatrix(0, 1) = MyLoadResString(k_Descending)
        Flex(0).AddItem ""
        Flex(1).Rows = 1
        Flex(1).Cols = 2
        Flex(1).TextMatrix(0, 0) = MyLoadResString(k_Property)
        Flex(1).TextMatrix(0, 1) = MyLoadResString(k_Value)
        Flex(1).AddItem ""
        Flex(0).FixedRows = 1
        Flex(1).FixedRows = 1
        Flex(1).ColAlignment(1) = flexAlignLeftCenter
        Flex(1).FixedCols = 1
    End If
    Set obj = Nothing
    bLoading = False
    
End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iImg As Integer, iLen As Integer, iElem As Integer
    Dim iTmpType As iPropG
    
    m_blnIdxProp = False
        
    iElem = Flex(1).Row - 1
    miRow(1) = iElem
    If Not bNewIndex Then Exit Sub
    sBuf = Flex(1).Text

    bEdit = True

    iTmpType = Elem(iElem).pgElement.iType
    iLen = Elem(iElem).pgElement.iLen
    If Not Flex(1).CellBackColor = vbInactiveBorder Then
        If Elem(iElem).EditType = iCombo Then
            If iElem = 7 Then 'filegroups
                miFlex = 1
                cbo.Clear
                For iLen = 1 To objServer.Databases(m_sDbName).FileGroups.Count
                    On Local Error Resume Next
                    cbo.AddItem objServer.Databases(m_sDbName).FileGroups(iLen).Name
                    On Local Error GoTo 0
                Next
            End If
            
            If cbo.ListCount <> 0 Then
                With Flex(1)
                    cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                    cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                End With
                m_blnIdxProp = True
                
                cbo.Visible = True
                cbo.ZOrder
                cbo.SetFocus
                If (iTmpType And am_String) = am_String Then
                    If KeyAscii = vbKeySpace Then KeyAscii = 0
                End If
                On Local Error Resume Next
                SendKeys Chr$(KeyAscii)
                On Local Error GoTo 0
            End If
            bEdit = False
        End If
    End If
    
    If Flex(1).CellBackColor = vbInactiveBorder Then bEdit = False

    If iElem < 0 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If (iTmpType And am_Bool) = am_Bool Then
                    iImg = 1 - (-CBool(Elem(iElem).pgElement.vVal))
                    Elem(iElem).pgElement.vVal = iImg
                    Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(iElem).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    SetEnabled
                    Flex(1).Row = iElem + 1
                    bEdit = False
                End If
            Case Else
                If (iTmpType And am_Bool) = am_Bool Then bEdit = False
        End Select
    End If

    If bEdit Then
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf
        Debug.Print iTmpType = am_Long
        tFlex.SelStart = 0
        With Flex(1)
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If (iTmpType And am_String) = am_String Then
            If KeyAscii = vbKeySpace Then KeyAscii = 0
        End If
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub

Private Sub tFlex_Change()
    If miRow(1) = 5 Then
        If Val(tFlex.Text) > 100 Then tFlex.Text = 100
    End If
End Sub
Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iElem As Integer
    Dim iTmpType As iPropG
  
    iElem = miRow(1)

    iTmpType = Elem(iElem).pgElement.iType
            
    If KeyCode = vbKeyUP Then
        KeyCode = 0
    ElseIf KeyCode = vbKeyDOWN Then
    Else
        If (iTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (iTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (iTmpType And am_String) = am_String Then
            If miFlex = 0 And miRow(1) = 1 Then
                'Input_KDown tFlex, KeyCode
            End If
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Dim iTmpType As iPropG
    Dim iElem As Integer
    
    iElem = miRow(1)
    iTmpType = Elem(iElem).pgElement.iType
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93    '<>|[]
            KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex(miFlex).Text
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(miFlex).SetFocus
            KeyAscii = 0
        Case vbKeyBACK
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then

            End If
        Case vbKeySpace
            KeyAscii = 0
        Case Else
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
    
    Dim iElem As Integer
    Dim lVal As Long
    Dim bUpdate As Boolean

    If Not m_blnAnnullaMod Then
        iElem = miRow(1)
        tFlex = Trim$(tFlex)
        If iElem = 5 Then
            lVal = Val(tFlex)
            If lVal <= 100 Then
                Elem(5).pgElement.vVal = lVal
                bUpdate = True
            Else
                tFlex = 0
            End If
        End If
    End If
    If bUpdate Then
        Flex(1).TextMatrix(iElem + 1, 1) = tFlex
        Flex(1).Row = iElem + 1
    Else
        If Not m_blnAnnullaMod Then Beep
    End If

    tFlex.Visible = False
    Flex(1).SetFocus
    m_blnAnnullaMod = True
    
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Flex(miFlex).SetFocus
            m_blnAnnullaMod = True
        Case vbKeyReturn
            KeyAscii = 0
            Flex(miFlex).SetFocus
    End Select

End Sub

Private Sub LoadCbo(ByVal iElem As Integer)

    Dim i As Integer, l As Integer
  
    cbo.Clear
    cbo.AddItem ""
    
    Dim obj As cColIDX

    If Flex(0).Col = 0 Then
        For i = 1 To allColumn.Count
            cbo.AddItem allColumn(i)
        Next
        For l = 1 To colIdxCol.Count
            Set obj = colIdxCol.Item(l)
            
            If l <> iElem Then
                For i = 1 To cbo.ListCount - 1
                    If obj.colName = cbo.List(i) Then
                        cbo.RemoveItem i
                        Exit For
                    End If
                Next
            End If
        Next
    End If
End Sub

Private Sub SaveIX()

    Dim i As Integer
    Dim bSave As Boolean
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim obj As cColIDX
    Dim sSql As String
    Dim sAttrib As String
    Dim sColumns As String
    Dim sSetOptions As String
    'Dim oView As SQLDMO.View2
        
    '30/10/2002
    'index creation for views is done via a t-sql execution becouse
    'standard DMO view.indexes.add objIndex method always fails
    'on Sql2k sp2 with error "no associated clustered index for view ..."
    'don't know why....
        
    Screen.MousePointer = vbHourglass
    sSql = ""
    sBuf = ""
    For i = 1 To colIdxCol.Count
        Set obj = colIdxCol.Item(i)
        If Len(obj.colName) <> 0 Then
            If Len(sColumns) Then sColumns = sColumns & ", "
            sColumns = sColumns & obj.colName
            If Len(sBuf) Then sBuf = sBuf & ", "
            sBuf = sBuf & "[" & Trim$(obj.colName) & "] "
            If obj.isDescending Then sBuf = sBuf & "DESC "
        End If
    Next
    
    bSave = Len(sBuf) <> 0
    
    sAttrib = ""

    If bSave Then
        sSql = "CREATE "
        If Abs(Elem(1).pgElement.vVal) Then sSql = sSql & "UNIQUE "
        If Abs(Elem(0).pgElement.vVal) Then sSql = sSql & "CLUSTERED "
        
        If Abs(Elem(4).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "PAD_INDEX "
        If Val(Elem(5).pgElement.vVal) <> 0 Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & " FILLFACTOR = " & Elem(5).pgElement.vVal
        If Abs(Elem(2).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "IGNORE_DUP_KEY "
        If Abs(Elem(3).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "STATISTICS_NORECOMPUTE "
        
        
        
        sSql = sSql & "INDEX [" & lblNewI.Caption & "] ON [" & GetViewOwner(objServer, m_sDbName, m_sViewName) & "].[" & m_sViewName & "] "
        sSql = sSql & "( " & sBuf & " ) "
        sSql = sSql & IIf(Len(sAttrib), " WITH ", "") & sAttrib
        
        sSql = sSql & " ON [" & Elem(7).pgElement.vVal & "]"
        
    Else
        lErr = vbObjectError + 2001
        'sErr = "Campi Non Definiti"
    End If
            
    If lErr = 0 Then
        'get SET options
        Dim oIdx As SQLDMO.Index2
        Set oIdx = New SQLDMO.Index2
        oIdx.Name = lblNewI.Caption
        oIdx.IndexedColumns = sColumns
        
        For i = 1 To colIdxCol.Count
            Set obj = colIdxCol.Item(i)
            If Len(obj.colName) <> 0 And obj.isDescending Then
                oIdx.SetIndexedColumnDESC obj.colName, True
            End If
        Next
        For i = 0 To UBound(Elem)
            If i <> 3 And i < 5 Then
                If Elem(i).pgElement.vVal Then
                    oIdx.Type = oIdx.Type Or native(i)
                    Debug.Print Elem(i).pgElement.sName
                End If
            End If
        Next
        oIdx.NoRecompute = Elem(3).pgElement.vVal
        oIdx.FileGroup = Elem(7).pgElement.vVal
        

        
        sSetOptions = oIdx.GenerateCreationSQLOnView(m_oView)
        Debug.Print sSetOptions
        Set oIdx = Nothing
        sBuf = Left$(sSetOptions, InStr(1, sSetOptions, "CREATE", vbTextCompare) - 1)
        Debug.Print sBuf
        
                
        Dim oDb As SQLDMO.Database2
        Set oDb = objServer.Databases(m_sDbName)
        'Set oView = GetViewByName(objServer, m_sDbName, m_sViewName)
        
'        If Not oView Is Nothing Then
'            Debug.Print oView.QuotedIdentifierStatus
'            oidx.
'            On Local Error Resume Next
'            oView.Indexes.Add oIdx
'            lErr = Err.Number
'            sErr = Err.Description
'            On Local Error GoTo 0
'
'        End If
        
        
        'sSql = sBuf & " " & sSql
        Debug.Print sSql
        On Local Error Resume Next
        oDb.ExecuteImmediate sSql, SQLDMOExec_Default, Len(sSql)
        lErr = Err.Number
        
        If lErr = -2147219570 Then
            oDb.ExecuteImmediate sBuf & " " & sSql, SQLDMOExec_Default, Len(sSql)
            lErr = Err.Number
        End If
        sErr = Err.Description
        On Local Error GoTo 0
        
        
        oDb.Views.Refresh True
        Set oDb = Nothing
        Set m_oView = Nothing
        Set m_oView = GetViewByName(objServer, m_sDbName, m_sViewName)
        sBuf = lblNewI.Caption
        Debug.Print m_oView Is Nothing
    End If
    
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        If lErr = vbObjectError + 2001 Then sErr = MyLoadResString(k_Columns_not_defined)
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        cIndex.AddItem sBuf
        cIndex.ListIndex = GetItem(sBuf, cIndex)
    End If
    
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

Private Sub SetEnabled()
    
    Dim i As Integer
    With Flex(1)
        .Col = .Cols - 1
        For i = 2 To 6
            .Row = i
            .CellBackColor = vbWindowBackground
        Next
        
        If Not CBool(Elem(1).pgElement.vVal) Then    'Unique
        
            .Row = 2 + 1
            Elem(2).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(2).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
        
        End If
    End With
                    
End Sub

Private Sub ManageExtendedProp()

    Dim iMode As am_SqlPropType
    Dim iOwner As am_SqlPropTypeOwner
    Dim sParam As String
        
    
    sParam = cIndex.Text
    iMode = am_ViewIdxProp
    iOwner = am_OwnView
    
    If Not m_bIsOwner Then Exit Sub
    
    LoadExtendedProp iMode, m_sDbName, GetObjectOwner(m_sDbName, m_sViewName, iOwner), m_sViewName, sParam
End Sub

Private Sub ModifySQL()

    Dim i As Integer
    Dim bSave As Boolean
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim obj As cColIDX
    Dim sSql As String
    Dim sAttrib As String
    Dim sColumns As String
    Dim sSetOptions As String
    Dim sRet As String
        
    Screen.MousePointer = vbHourglass
    sSql = ""
    sBuf = ""
    For i = 1 To colIdxCol.Count
        Set obj = colIdxCol.Item(i)
        If Len(obj.colName) <> 0 Then
            If Len(sColumns) Then sColumns = sColumns & ", "
            sColumns = sColumns & obj.colName
            If Len(sBuf) Then sBuf = sBuf & ", "
            sBuf = sBuf & "[" & Trim$(obj.colName) & "] "
            If obj.isDescending Then sBuf = sBuf & "DESC "
        End If
    Next
    
    bSave = Len(sBuf) <> 0
    
    sAttrib = ""

    If bSave Then
        sSql = "CREATE "
        If Abs(Elem(1).pgElement.vVal) Then sSql = sSql & "UNIQUE "
        If Abs(Elem(0).pgElement.vVal) Then sSql = sSql & "CLUSTERED "
        
        If Abs(Elem(4).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "PAD_INDEX "
        If Val(Elem(5).pgElement.vVal) <> 0 Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & " FILLFACTOR = " & Elem(5).pgElement.vVal
        If Abs(Elem(2).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "IGNORE_DUP_KEY "
        If Abs(Elem(3).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "STATISTICS_NORECOMPUTE "
        
        
        
        sSql = sSql & "INDEX [" & lblNewI.Caption & "] ON [" & GetViewOwner(objServer, m_sDbName, m_sViewName) & "].[" & m_sViewName & "] "
        sSql = sSql & vbCrLf & "( " & sBuf & " ) "
        sSql = sSql & IIf(Len(sAttrib), vbCrLf & " WITH ", "") & sAttrib
        
        sSql = sSql & vbCrLf & " ON [" & Elem(7).pgElement.vVal & "]"
        
    Else
        lErr = vbObjectError + 2001
        'sErr = "Campi Non Definiti"
    End If
            
    If lErr = 0 Then
        'get SET options
        Dim oIdx As SQLDMO.Index2
        Set oIdx = New SQLDMO.Index2
        oIdx.Name = lblNewI.Caption
        oIdx.IndexedColumns = sColumns
        
        For i = 1 To colIdxCol.Count
            Set obj = colIdxCol.Item(i)
            If Len(obj.colName) <> 0 And obj.isDescending Then
                oIdx.SetIndexedColumnDESC obj.colName, True
            End If
        Next
        
        sSetOptions = oIdx.GenerateCreationSQLOnView(m_oView)
        Debug.Print sSetOptions
        Set oIdx = Nothing
        sBuf = Left$(sSetOptions, InStr(1, sSetOptions, "CREATE", vbTextCompare) - 1)
        Debug.Print sBuf
        
        Screen.MousePointer = vbDefault
        With fResult
            .Action() = act_Modify_Idx_Sql
            .Caption = MyLoadResString(k_Frm_ModifySQL)
            .tRes.Text = sBuf & vbCrLf & sSql
            .Show vbModal, Me
            bSave = .Confirm()
            If bSave Then
                sRet = .tRes.Text
            End If
        End With
        Unload fResult
        
        If bSave Then
            Screen.MousePointer = vbHourglass
            Dim oDb As SQLDMO.Database2
            Set oDb = objServer.Databases(m_sDbName)
        
            On Local Error Resume Next
            oDb.ExecuteImmediate sRet, SQLDMOExec_Default, Len(sRet)
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
        
        
            oDb.Views.Refresh True
            Set oDb = Nothing
            Set m_oView = Nothing
            Set m_oView = GetViewByName(objServer, m_sDbName, m_sViewName)
            sBuf = lblNewI.Caption
            Debug.Print m_oView Is Nothing
        End If
    End If
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        If lErr = vbObjectError + 2001 Then sErr = MyLoadResString(k_Columns_not_defined)
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        If bSave Then
            cIndex.AddItem sBuf
            cIndex.ListIndex = GetItem(sBuf, cIndex)
        End If
    End If
    
End Sub


