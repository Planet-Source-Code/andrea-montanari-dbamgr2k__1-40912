VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fIndex 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Indici"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7140
   Icon            =   "fIndex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Mod. SQL"
      Height          =   360
      Index           =   4
      Left            =   5400
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   315
      Index           =   5
      Left            =   6720
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4440
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3000
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
      TabIndex        =   2
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Frame fra 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   5655
      Begin VB.TextBox tCheck 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.ComboBox cboTB 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cIndex 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "Tab.Chiave Primaria"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl 
         Caption         =   "Elemento Selezionato"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblNewI 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
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
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
      Height          =   360
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Nuovo"
      Height          =   360
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   4335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vincoli Check"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relazioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Indici/Chiavi"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fIndex"
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
Private m_sTbName As String
Private miFlex As Integer       'Indice Flex
Private miRow(1) As Integer        'Riga Flex
Private miCol As Integer
Private m_blnAnnullaMod As Boolean
Private bLoading As Boolean

Private m_oTbl As SQLDMO.Table
Private Elem(11) As idxElem 'PageElement
Private native(11) As Long
Private colIdxCol As Collection
Private colKey As Collection            'Colonne di chiave esterna
Private allColumn As Collection
Private bNewIndex As Boolean
Private allTB() As String
Private ReferTB As String               'Tabella di Chiave esterna
Private keyElem() As PageElement        'attributi di FK
Private allRefCol As Collection         'Tutte i campi della Tabella esterna
Private bLoadingFK As Boolean
Private m_sOwner As String
Private m_bIsReadOnly As Boolean
Private m_lWidth As Long
Private m_lHeight As Long
Private m_bIsOwner As Boolean

Private Sub Load_Elem()
        
    native(0) = SQLDMOIndex_DRIPrimaryKey
    native(1) = SQLDMOIndex_Clustered
    native(2) = SQLDMOIndex_Unique
    native(3) = SQLDMOIndex_DRIUniqueKey
    native(4) = SQLDMOIndex_NoRecompute
    native(5) = SQLDMOIndex_DRIIndex
    native(6) = SQLDMOIndex_IgnoreDupKey
    native(7) = -1  'FillFactor
    native(8) = SQLDMOIndex_PadIndex
    native(9) = -2  'SpaceUsed
    native(10) = -3 'FileGroup
    native(11) = SQLDMOIndex_Default
    
    
    Elem(0).pgElement.sName = "PrimaryKey"
    Elem(0).pgElement.iType = am_Bool
    
    Elem(1).pgElement.sName = "Clustered"
    Elem(1).pgElement.iType = am_Bool
    
    Elem(2).pgElement.sName = "Unique"
    Elem(2).pgElement.iType = am_Bool
    
    Elem(3).pgElement.sName = MyLoadResString(k_Constraint)
    Elem(3).pgElement.iType = am_Bool
    
    Elem(4).pgElement.sName = MyLoadResString(k_Not_refresh_statistics)
    Elem(4).pgElement.iType = am_Bool

    Elem(5).pgElement.sName = MyLoadResString(k_DRI_Key)
    Elem(5).pgElement.iType = am_Bool
    
    Elem(6).pgElement.sName = "IgnoreDupKey"
    Elem(6).pgElement.iType = am_Bool
    
    Elem(7).pgElement.sName = "Fill Factor"
    Elem(7).pgElement.iType = am_Long
    Elem(7).pgElement.iLen = 3
    
    Elem(8).pgElement.sName = "PadIndex"
    Elem(8).pgElement.iType = am_Bool
        
    Elem(9).pgElement.sName = "Space Used Kb"
    Elem(9).pgElement.iType = am_Long
    Elem(9).pgElement.iLen = -1

    Elem(10).pgElement.sName = MyLoadResString(k_HtmFileGroup) '"File Group"
    Elem(10).pgElement.iType = am_String
    Elem(10).EditType = iCombo
    Elem(10).pgElement.iLen = -1

    Elem(11).pgElement.sName = "Default"
    Elem(11).pgElement.iType = am_Bool
    Elem(11).pgElement.iLen = -1
          
End Sub
Public Sub DataNames(ByVal sDB As String, ByVal sTb As String)
    
    m_sDbName = sDB
    m_sTbName = sTb
    
    
    Dim i As Integer
    Dim iC As Integer
    
    'Set oTbl = objServer.Databases(sDbName).Tables(sTbName)
    m_bIsReadOnly = Not IsDbAvailableReadWrite(sDB)
    Set m_oTbl = GetTbByName(objServer, m_sDbName, sTb)
    If Not m_oTbl Is Nothing Then
        m_sOwner = m_oTbl.Owner
        m_bIsOwner = IsObjectOwner(m_sDbName, sTb, am_OwnTable)
    End If
    
    
    Set allColumn = New Collection
    For i = 1 To m_oTbl.Columns.Count
        allColumn.Add m_oTbl.Columns(i).Name
    Next
    
    With objServer.Databases(m_sDbName)
        For i = 1 To .Tables.Count
            If .Tables(i).SystemObject = 0 Then
                ReDim Preserve allTB(iC) As String
                allTB(iC) = "[" & .Tables(i).Name & "]"
                iC = iC + 1
            End If
        Next
    End With
    Set tStrip.SelectedItem = tStrip.Tabs(3)
    
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
            If tStrip.SelectedItem.Index - 1 = 2 Or miCol = 1 Then
                If colIdxCol.Count >= miRow(0) Then
                    Set obj = colIdxCol.Item(miRow(0))
                Else
                    m_blnAnnullaMod = True
                End If
            Else
                If colKey.Count >= miRow(0) Then
                    Set obj = colKey.Item(miRow(0))
                Else
                    m_blnAnnullaMod = True
                End If
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
    'Flex(miFlex).SetFocus
    m_blnAnnullaMod = True
    m_blnIdxProp = False
    
End Sub

Private Sub cboTB_Click(Index As Integer)

    Static b As Boolean
    Dim oC As SQLDMO.Column
    Dim i As Integer
    Dim oTb As SQLDMO.Table
    
    If Index = 0 Then
        If Not b And Not bLoadingFK Then
            b = True
            ReferTB = cboTB(Index).Text
            Flex(0).TextMatrix(0, 0) = ReferTB
            Set allRefCol = New Collection
            
            Dim oCol As cColIDX
            If Len(cboTB(Index).Text) Then
                Set colKey = New Collection
                
                Set oTb = GetTbByName(objServer, m_sDbName, ReferTB)
                'Debug.Print objServer.Databases(sDbName).Tables(ReferTB).Columns.Count
                If Not oTb Is Nothing Then
                    For i = 1 To oTb.Columns.Count
                        Set oC = oTb.Columns(i)
                    
                        allRefCol.Add oC.Name
                        Set oCol = New cColIDX
                        colKey.Add oCol
                    Next
                    
'                For i = 1 To objServer.Databases(sDbName).Tables(ReferTB).Columns.Count
'                    Set oC = objServer.Databases(sDbName).Tables(ReferTB).Columns(i)
'
'                    allRefCol.Add oC.Name
'                    Set oCol = New cColIDX
'                    colKey.Add oCol
'                Next
                
                    LoadGrids 0
                    SettaLargFlex Flex(0), 0, 0
                End If
            End If
            b = False
        End If
    End If
    Set oC = Nothing
    
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
            If tStrip.SelectedItem.Index - 1 > 0 Then
                Flex(0).SetFocus
            Else
                tCheck.SetFocus
            End If
            If tStrip.SelectedItem.Index - 1 <> 1 Then
                SaveIX
            Else
                SaveRefKey
            End If
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
            If tStrip.SelectedItem.Index - 1 = 0 Then
                tCheck.SetFocus
            Else
                Flex(0).SetFocus
            End If
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
    i = tStrip.SelectedItem.Index - 1
    fPwd.Caption = MyLoadResString(k_New_Check_Constraint + i) & "..."
    fPwd.Stuff = Choose(i + 1, "CH_", "FK_", "IX_")
    
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
        Select Case i
            Case 2                  'New Indice
                Set colIdxCol = New Collection
            
                For i = 0 To 11
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
            Case 1          'Nuova Relazione
                Set colKey = New Collection
            
                Set colIdxCol = New Collection
                For i = 1 To allColumn.Count
                    Set oCol = New cColIDX
                    colIdxCol.Add oCol
                Next
    
                ReDim keyElem(3) As PageElement
                                
                keyElem(0).sName = MyLoadResString(k_Check_existing_data)
                keyElem(0).iType = am_Bool
                keyElem(0).vVal = True
        
                keyElem(1).sName = MyLoadResString(k_Activate_Relation_for_Replica)
                keyElem(1).iType = am_Bool
                keyElem(1).vVal = True
                
                keyElem(2).sName = MyLoadResString(k_Update_Cascade)
                keyElem(2).iType = am_Bool
                keyElem(2).vVal = False
                
                keyElem(3).sName = MyLoadResString(k_Delete_Cascade)
                keyElem(3).iType = am_Bool
                keyElem(3).vVal = False
                
                ReferTB = ""
            Case 0          'Check
                ReDim keyElem(1) As PageElement
                                
                keyElem(0).sName = MyLoadResString(k_Activate_Constraint_for_INSERT_and_UPDATE)
                keyElem(0).iType = am_Bool
                keyElem(0).vVal = True
        
                keyElem(1).sName = MyLoadResString(k_Activate_Constraint_for_Replica)
                keyElem(1).iType = am_Bool
                keyElem(1).vVal = True
                tCheck.Locked = False
                tCheck.Text = ""
        
        End Select
        LoadGrids 0
        LoadGrids 1
        SizeControls
        miRow(0) = 0
        miCol = 0
        Flex(0).Row = 1
        Flex(0).Col = 0
        
        Select Case tStrip.SelectedItem.Index - 1
            Case 0      'Check
                tCheck.SetFocus
            Case 1      'FK
                LoadTBcombo 0
                cboTB(0).SetFocus
            Case 2      'Ix
                Flex(0).SetFocus
                Flex_RowColChange 0
        End Select
        cmd(1).Visible = True
        cmd(2).Visible = False
        
    End If
    
End Sub
Private Sub DropIndex()
    
    Dim sBuf As String, sChar As String
    Dim i As Integer
    Dim iElem As Integer
    Dim lErr As Long, sErr As String
    Dim bDrop As Boolean
    
    If cIndex.ListCount Then
        sBuf = cIndex.Text
        iElem = cIndex.ListIndex
        sChar = MyLoadResString(k_Check_Constraints + tStrip.SelectedItem.Index - 1)
        
        If MsgBox(MyLoadResString(k_Confirm_Remove) & vbLf & sChar & " " & sBuf, vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then

            On Local Error Resume Next
            If tStrip.SelectedItem.Index - 1 > 0 Then
                For i = 1 To m_oTbl.Keys.Count
                    If m_oTbl.Keys(i).Name = sBuf Then
                        m_oTbl.Keys.Remove sBuf
                        lErr = Err.Number
                        sErr = Err.Description
                        bDrop = True
                        Exit For
                    End If
                Next
                If tStrip.SelectedItem.Index - 1 = 2 Then
                    If lErr = 0 Then
                        If Not bDrop Then
                            m_oTbl.Indexes.Remove sBuf
                            lErr = Err.Number
                            sErr = Err.Description
                        End If
                    End If
                End If
            Else
                m_oTbl.Checks.Remove sBuf
                lErr = Err.Number
                sErr = Err.Description
            
            End If
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
        If bNewIndex And Not Flex(0).CellBackColor = vbInactiveBorder Then
            iImg = tStrip.SelectedItem.Index - 1
            With Flex(0)
                If (.Col < .Cols - 1 And iImg = 2) Or iImg = 1 Then
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo
    'spostamento sulla scheda successiva.
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tStrip.SelectedItem.Index
        If i = tStrip.Tabs.Count Then
            'È stata raggiunta l'ultima scheda e quindi
            'torna alla scheda 1.
            Set tStrip.SelectedItem = tStrip.Tabs(1)
        Else
            'Incrementa l'indice della scheda
            Set tStrip.SelectedItem = tStrip.Tabs(i + 1)
        End If
    End If

End Sub
Private Sub Form_Load()
    
    m_lWidth = Me.ScaleWidth
    m_lHeight = Me.ScaleHeight
    
    Me.Caption = MyLoadResString(k_Indexes)
    tStrip.Tabs.Item(1).Caption = MyLoadResString(k_Check_Constraints)
    tStrip.Tabs.Item(2).Caption = MyLoadResString(k_Relations)
    tStrip.Tabs.Item(3).Caption = MyLoadResString(k_Indexes_Keys_constraints)
    

    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    cmd(2).Caption = MyLoadResString(k_New)
    cmd(3).Caption = MyLoadResString(k_Remove)
    cmd(4).Caption = MyLoadResString(k_ModifySQL)
    cmd(5).ToolTipText = MyLoadResString(k_Extended_Property)
        
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
    Set colKey = Nothing
    Set colIdxCol = Nothing
    Set m_oTbl = Nothing
    Set fIndex = Nothing
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
    
    tStrip.Width = Me.ScaleWidth - 2 * (tStrip.Left)
    tStrip.Height = cmd(0).Top - 50
    
    fra.Width = Me.ScaleWidth - 2 * (fra.Left)
        
    Flex(0).Width = Me.ScaleWidth - (2 * Flex(0).Left)
    Flex(1).Top = Flex(0).Height + 25 + Flex(0).Top
    Flex(1).Width = Me.ScaleWidth - (2 * Flex(1).Left)
    
    Select Case tStrip.SelectedItem.Index - 1
        Case 0
            lblNewI.Width = fra.Width \ 2 - 120
            lblNewI.Left = fra.Width \ 2
        
            With Flex(1)
                .Height = 1050
                .Top = cmd(0).Top - .Height - 100
                .ColWidth(0) = 2 * (.Width / 3)
                .ColWidth(1) = .Width - .ColWidth(0) - 75
            End With
        
            fra.Height = Flex(1).Top - fra.Top - 10
            tCheck.Height = fra.Height - cboTB(0).Top - cboTB(0).Height - 150
            tCheck.Width = fra.Width - 2 * tCheck.Left
        
        Case 1      'FK
            cboTB(0).Width = fra.Width \ 2 - 120
            lblNewI.Width = fra.Width \ 2 - 120
            lblNewI.Move cboTB(0).Width + cboTB(0).Left + 50
            
            With Flex(0)
                .Height = (cmd(0).Top - 150 - .Top) \ 2
                .Height = (cmd(0).Top - 150 - .Top) - Flex(1).Height
                .ColWidth(1) = .Width \ 2
            End With
            
            With Flex(1)
                .Height = 1050
                .Top = Flex(0).Top + Flex(0).Height
                .ColWidth(0) = 2 * (.Width / 3)
                If .RowIsVisible(.Rows - 1) Then
                    .ColWidth(1) = .Width - .ColWidth(0) - 75
                Else
                    .ColWidth(1) = .Width - .ColWidth(0) - 350
                End If
            End With
        Case 2
            lblNewI.Width = fra.Width \ 2 - 120
            lblNewI.Left = fra.Width \ 2
            fra.Height = Flex(0).Top - fra.Top
            With Flex(0)
                .Height = (cmd(0).Top - 150 - .Top) \ 2
                .ColWidth(0) = .Width \ 2
            End With
            
            With Flex(1)
                .Top = Flex(0).Top + Flex(0).Height
                .Height = Flex(0).Height
                .ColWidth(0) = .Width / 2
                .ColWidth(1) = .ColWidth(0) - 350
            End With
    End Select
    cIndex.Width = fra.Width - lbl(0).Width - lbl(0).Left - 200
    SettaLargFlex Flex(0), 0, 200
    SettaLargFlex Flex(1), 0, 200
    
End Sub

Private Sub LoadIndex(ByVal sIndexName As String)

    Dim oIndex As SQLDMO.Index2
    Dim oCol As SQLDMO.Column
    Dim oKey As SQLDMO.Key
    Dim oChk As SQLDMO.Check
    Dim objCol As cColIDX
    Dim objKey As cKeyCasc

    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    lblNewI.Visible = False
    lblNewI.Caption = ""
    bNewIndex = False
    
    bLoadingFK = True
    cboTB(0).Clear
    bLoadingFK = False
    
    cmd(1).Visible = False
    cmd(2).Visible = True
    cmd(4).Visible = False
    cmd(5).Enabled = False

            
    Screen.MousePointer = vbHourglass
    If Len(sIndexName) = 0 Then
        Set m_oTbl = Nothing
        
        Set m_oTbl = GetTbByName(objServer, m_sDbName, m_sTbName)
        m_oTbl.Indexes.Refresh True
        m_oTbl.Keys.Refresh True
        m_oTbl.Checks.Refresh True
        
        Set m_oTbl = Nothing
        Set m_oTbl = GetTbByName(objServer, m_sDbName, m_sTbName)
    End If
    
    
    Select Case i
        Case 2          'Indici
            If Len(sIndexName) <> 0 Then
                cmd(5).Enabled = m_bIsOwner And Not bNewIndex And i <> 0
                
                On Local Error GoTo exitRoutine
                Set oIndex = m_oTbl.Indexes(sIndexName)
                With oIndex
                    'Set oList = oIndex.ListIndexedColumns
                    Set colIdxCol = New Collection
                    
                    For Each oCol In oIndex.ListIndexedColumns
                        Set objCol = New cColIDX
                        objCol.colName = oCol.Name
                        objCol.isDescending = oIndex.GetIndexedColumnDESC(oCol.Name)
                        colIdxCol.Add objCol
                    Next
                    Set objCol = Nothing
                    
                    For i = 0 To 11
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
                Debug.Print m_oTbl.Indexes.Count
                For Each oIndex In m_oTbl.Indexes
                    Debug.Print (oIndex.Type And SQLDMOIndex_Hypothetical) = SQLDMOIndex_Hypothetical
                    'If (oIndex.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex Then cIndex.AddItem oIndex.Name
                    If ((oIndex.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex) And ((oIndex.Type And SQLDMOIndex_Hypothetical) <> SQLDMOIndex_Hypothetical) Then cIndex.AddItem oIndex.Name
                Next
            End If
        
        Case 1          'Integrita' Refer.
            ReDim keyElem(3) As PageElement
            
            keyElem(0).sName = MyLoadResString(k_Activate_Relation_for_INSERT_and_UPDATE)
            keyElem(0).iType = am_Bool
            keyElem(0).vVal = True
    
            keyElem(1).sName = MyLoadResString(k_Activate_Relation_for_Replica)
            keyElem(1).iType = am_Bool
            keyElem(1).vVal = True
            
            keyElem(2).sName = MyLoadResString(k_Update_Cascade)
            keyElem(2).iType = am_Bool
            keyElem(2).vVal = False
            
            keyElem(3).sName = MyLoadResString(k_Delete_Cascade)
            keyElem(3).iType = am_Bool
            keyElem(3).vVal = False
            
            ReferTB = ""
            If Len(sIndexName) <> 0 Then    'Load
                cmd(5).Enabled = m_bIsOwner And Not bNewIndex And i <> 0
                On Local Error GoTo exitRoutine
                Set oKey = m_oTbl.Keys(sIndexName)
                With oKey
                    keyElem(0).vVal = .Checked
                    keyElem(1).vVal = Not .ExcludeReplication
                    
                    Set colIdxCol = New Collection
                    Set colKey = New Collection
                                                    
                    For i = 1 To .KeyColumns.Count
                        Set objCol = New cColIDX
                        objCol.colName = .KeyColumns(i)
                        colIdxCol.Add objCol
                    Next
                    Set objCol = Nothing
                    For i = 1 To .ReferencedColumns.Count
                        Set objCol = New cColIDX
                        objCol.colName = .ReferencedColumns(i)
                        colKey.Add objCol
                    Next
                
                    
                    ReferTB = .ReferencedTable
                    i = InStr(ReferTB, ".")
                    If i <> 0 Then ReferTB = Mid$(ReferTB, i + 1)
                    Set objCol = Nothing
                    
                    Set objKey = New cKeyCasc
                    objKey.InitStruct m_sDbName, oKey.Name, 0, 0
                    'Set objKey.KeyObject() = oKey
                    
                    keyElem(2).vVal = objKey.UpdateCascade
                    keyElem(3).vVal = objKey.DeleteCascade
                    
                    Set objKey = Nothing
                End With
                bLoadingFK = True
                LoadGrids 0
                LoadGrids 1
                bLoadingFK = False
    
            Else                            'Load All
                LoadGrids -1
                Set colIdxCol = Nothing
                cIndex.Clear
                For Each oKey In m_oTbl.Keys
                    If (oKey.Type And SQLDMOKey_Foreign) = SQLDMOKey_Foreign Then cIndex.AddItem oKey.Name
                Next
            End If
        Case 0          'Check
            tCheck.Text = ""
            tCheck.Locked = True
            ReDim keyElem(1) As PageElement
                            
            keyElem(0).sName = MyLoadResString(k_Activate_Constraint_for_INSERT_and_UPDATE)
            keyElem(0).iType = am_Bool
            keyElem(0).vVal = False
    
            keyElem(1).sName = MyLoadResString(k_Activate_Constraint_for_Replica)
            keyElem(1).iType = am_Bool
            keyElem(1).vVal = False
            If Len(sIndexName) <> 0 Then    'Load
                                
                On Local Error GoTo exitRoutine
                Set oChk = m_oTbl.Checks(sIndexName)
                With oChk
                    tCheck.Text = .Text
                    keyElem(0).vVal = .Checked
                    keyElem(1).vVal = Not .ExcludeReplication
                End With
                bLoadingFK = True
                LoadGrids 0
                LoadGrids 1
                bLoadingFK = False
    
            Else                            'Load All
                tCheck.Text = ""
                LoadGrids -1
                cIndex.Clear
                For Each oChk In m_oTbl.Checks
                    cIndex.AddItem oChk.Name
                Next
            End If
        
    End Select
    
finish:
    
    Set objCol = Nothing
    'Set oList = Nothing
    Set oCol = Nothing
    Set oIndex = Nothing
    Set oKey = Nothing
    Set oChk = Nothing
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
        
    Dim i As Integer, y As Integer
    Dim sBuf As String
    Dim obj As cColIDX
    
    bLoading = True
    
    i = tStrip.SelectedItem.Index - 1
    Select Case i
        Case 2          'Indici
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
                        For i = 0 To 11
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
                                Set Flex(Index).CellPicture = LoadResPicture(IIf(Abs(Elem(i).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                .CellPictureAlignment = flexAlignCenterCenter
                            End If
                            If Elem(i).EditType = iCombo Then
                                .Row = i + 1
                                .Col = .Cols - 1
                                Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                                .CellPictureAlignment = flexAlignRightTop
                                
                            End If
                            If i > 8 And i <> 10 Then
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
        Case 1              'Relazioni
            If Index <> -1 Then
                With Flex(Index)
                    .Redraw = False
                    .ScrollBars = flexScrollBarBoth
                    .Clear
                    .FixedCols = 0
                    .FixedRows = 0
                    
                    .Cols = 2
                    .Rows = 1
            
                    .ScrollBars = flexScrollBarBoth
                    If Index = 0 Then
                        LoadTBcombo 0
                        cboTB(0).ListIndex = GetItem(ReferTB, cboTB(0))
                        
                        .TextMatrix(0, 0) = ReferTB
                        .TextMatrix(0, 1) = m_sTbName
    
                    Else
                        .TextMatrix(0, 0) = MyLoadResString(k_Property)
                        .TextMatrix(0, 1) = MyLoadResString(k_Value)
                    End If
                    If Index = 0 Then
                        y = colKey.Count
                        i = colIdxCol.Count
                        
                        If i > y Then y = i
                        
                        .Rows = y + 1
    
                        For i = 1 To y
                            If i <= colKey.Count Then .TextMatrix(i, 0) = colKey.Item(i).colName
                            If i <= colIdxCol.Count Then .TextMatrix(i, 1) = colIdxCol.Item(i).colName
                        Next
                        For i = 1 To .Rows - 1
                            .Row = i
                            .Col = 0
                            Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                            .CellPictureAlignment = flexAlignRightTop
                            .Col = 1
                            Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                            .CellPictureAlignment = flexAlignRightTop
                            
                        Next
                    Else
                        For i = 0 To UBound(keyElem)
                            If keyElem(i).iType <> am_Bool Then sBuf = keyElem(i).vVal
                            .AddItem keyElem(i).sName & vbTab & sBuf
                            sBuf = ""
                            .Row = Flex(1).Rows - 1
                            If keyElem(i).iType = am_Bool Then
                                .Col = 1
                                Set Flex(Index).CellPicture = LoadResPicture(IIf(Abs(keyElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                .CellPictureAlignment = flexAlignCenterCenter
                                
                            End If
                            If Not bNewIndex Then
                                If i > 1 Then .CellBackColor = vbInactiveBorder
                            End If
                        Next
                        
                        .ColAlignment(1) = flexAlignLeftCenter
                        .FixedCols = 1
                    End If
                    .FixedRows = 1
                    .Redraw = True
                End With
            Else
                With Flex(0)
                    .Redraw = False
                    .FixedRows = 0
                    .FixedCols = 0
                    .Clear
                    .Rows = 1
                    .Cols = 2
                    .TextMatrix(0, 0) = ""
                    .TextMatrix(0, 1) = ""
                    .AddItem ""
                    .FixedRows = 1
                    .Redraw = True
                End With
            
                With Flex(1)
                    .Redraw = False
                    .FixedRows = 0
                    .FixedCols = 0
                    .Clear
                    .Rows = 1
                    .Cols = 2
                    
                    .TextMatrix(0, 0) = MyLoadResString(k_Property)
                    .TextMatrix(0, 1) = MyLoadResString(k_Value)
                    For i = 0 To UBound(keyElem)
                        If keyElem(i).iType <> am_Bool Then sBuf = keyElem(i).vVal
                        .AddItem keyElem(i).sName & vbTab & sBuf
                        sBuf = ""
                        .Row = Flex(1).Rows - 1
                        If keyElem(i).iType = am_Bool Then
                            .Col = 1
                            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(keyElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                            .CellPictureAlignment = flexAlignCenterCenter
                            
                        End If
                        .CellBackColor = vbInactiveBorder
                    Next
                    
                    .ColAlignment(1) = flexAlignLeftCenter
                    
                    .FixedRows = 1
                    .FixedCols = 1
                    .Redraw = True
                End With
            End If
        Case 0          'Check
                
            If Index = -1 Then
                ReDim keyElem(1) As PageElement
                                
                keyElem(0).sName = MyLoadResString(k_Activate_Constraint_for_INSERT_and_UPDATE)
                keyElem(0).iType = am_Bool
        
                keyElem(1).sName = MyLoadResString(k_Activate_Constraint_for_Replica)
                keyElem(1).iType = am_Bool
            End If
                        
            With Flex(1)
                .Redraw = False
                .ScrollBars = flexScrollBarBoth
                .Clear
                .FixedCols = 0
                .FixedRows = 0
                
                .Cols = 2
                .Rows = 1
            
                .ScrollBars = flexScrollBarBoth
                .TextMatrix(0, 0) = MyLoadResString(k_Property)
                .TextMatrix(0, 1) = MyLoadResString(k_Value)
                For i = 0 To UBound(keyElem)
                    If keyElem(i).iType <> am_Bool Then sBuf = keyElem(i).vVal
                    .AddItem keyElem(i).sName & vbTab & sBuf
                    sBuf = ""
                    .Row = Flex(1).Rows - 1
                    If keyElem(i).iType = am_Bool Then
                        .Col = 1
                        Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(keyElem(i).vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                    End If
                    If Index = -1 Then .CellBackColor = vbInactiveBorder
                Next
                        
                .ColAlignment(1) = flexAlignLeftCenter
                .FixedCols = 1
        
                .FixedRows = 1
                .Redraw = True
            End With
    End Select
    Set obj = Nothing
    bLoading = False
    
End Sub
Private Sub SetEnabled()
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim bIsPK As Boolean
    Dim i As Integer
    
    With Flex(0)
        iRow = Flex(0).Row
        iCol = Flex(0).Col
        bIsPK = CBool(Elem(0).pgElement.vVal) 'Primary Key
        bLoading = True
        For i = 1 To .Rows - 1
            .Row = i
            .Col = .Cols - 1
            If bIsPK Then colIdxCol.Item(i).isDescending() = False
                        
            Set Flex(0).CellPicture = LoadResPicture(IIf(colIdxCol.Item(i).isDescending(), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = IIf(bIsPK, vbInactiveBorder, vbWindowBackground)
        Next
        .Row = iRow
        .Col = iCol
        bLoading = False
    End With
    
    With Flex(1)
        .Col = .Cols - 1
        For i = 1 To 9
            .Row = i
            If i = 6 Then
                .CellBackColor = vbInactiveBorder
            Else
                .CellBackColor = vbWindowBackground
            End If
        Next
        
        If Elem(0).pgElement.vVal Then    'Primary Key
            Elem(2).pgElement.vVal = False
            .Row = 2 + 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(2).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            .Row = 3 + 1
            Elem(3).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(3).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            .Row = 4 + 1
            Elem(4).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(4).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            .Row = 6 + 1
            Elem(6).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(6).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            '.Row = 8 + 1
            'Elem(8).pgElement.vVal = False
            'Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(8).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            '.CellPictureAlignment = flexAlignCenterCenter
            '.CellBackColor = vbInactiveBorder
        ElseIf Elem(2).pgElement.vVal = False Then    'Not Unique
            .Row = 3 + 1
            Elem(3).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(3).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            .Row = 6 + 1
            Elem(6).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(6).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            .Row = 4 + 1
            Elem(4).pgElement.vVal = False
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(4).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
        ElseIf Elem(3).pgElement.vVal Then
            Elem(0).pgElement.vVal = False
            .Row = 0 + 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(0).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            Elem(4).pgElement.vVal = False
            .Row = 4 + 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(4).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
            
            Elem(6).pgElement.vVal = False
            .Row = 6 + 1
            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(6).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = vbInactiveBorder
             
'             Elem(8).pgElement.vVal = 0
'            .Row = 8 + 1
'            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(8).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
'            .CellPictureAlignment = flexAlignCenterCenter
'            .CellBackColor = vbInactiveBorder
        End If
'        If Elem(7).pgElement.vVal = 0 Then
'            Elem(8).pgElement.vVal = False
'            .Row = 8 + 1
'            Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(Elem(8).pgElement.vVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
'            .CellPictureAlignment = flexAlignCenterCenter
'            .CellBackColor = vbInactiveBorder
'        End If
    End With
                    
End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iImg As Integer, iLen As Integer, iElem As Integer
    Dim iTmpType As iPropG
    Dim iTab As Integer
    

    m_blnIdxProp = False
    iTab = tStrip.SelectedItem.Index - 1
    iElem = Flex(1).Row - 1
    If iTab > 0 Then miRow(iTab - 1) = iElem
    If iTab = 2 Then
        If Not bNewIndex Then Exit Sub
    Else
        If Not bNewIndex Then
            If iElem > 1 Then Exit Sub
        End If
    End If
    
    sBuf = Flex(1).Text

    bEdit = True
    
    Select Case iTab
        Case 2
            iTmpType = Elem(iElem).pgElement.iType
            iLen = Elem(iElem).pgElement.iLen
            If Not Flex(1).CellBackColor = vbInactiveBorder Then
                If Elem(iElem).EditType = iCombo Then
                    If iElem = 10 Then 'filegroups
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
            
        Case Else
            iTmpType = keyElem(iElem).iType
            iLen = keyElem(iElem).iLen
    End Select
    If Flex(1).CellBackColor = vbInactiveBorder Then bEdit = False

    If iElem < 0 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If (iTmpType And am_Bool) = am_Bool Then
                
                    If iTab = 2 Then
                        iImg = 1 - (-CBool(Elem(iElem).pgElement.vVal))
                        Elem(iElem).pgElement.vVal = iImg
                    Else
                        If Not bNewIndex Then
                            If Not UpdateConstraint(iTab, iElem) Then Exit Sub
                        End If
                        iImg = 1 - (-CBool(keyElem(iElem).vVal))
                        keyElem(iElem).vVal = iImg
                    End If
                    
                    Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex(1).CellPictureAlignment = flexAlignCenterCenter
                    
                    bEdit = False
                    If iTab = 2 Then SetEnabled
                    Flex(1).Row = iElem + 1
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
'            If KeyAscii Then
'                'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'                'sBuf = Chr$(KeyAscii) & Left$(sBuf, iLen - 1)
'                If Len(sBuf) Then
'                    sBuf = Chr$(KeyAscii) & Mid$(sBuf, 2, iLen - 2)
'                Else
'                    sBuf = Chr$(KeyAscii)
'                End If
'                iLen = 1
'            Else
'                iLen = 0
'            End If
'            tFlex = sBuf
'            tFlex.SelStart = iLen
        'Else
        '    tFlex_KeyPress KeyAscii
        End If
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub

Private Sub tFlex_Change()
    Select Case tStrip.SelectedItem.Index - 1
        Case 2
            If miRow(1) = 7 Then
                If Val(tFlex) > 100 Then tFlex = 100
            End If
        Case 1
            If miRow(1) = 2 Then
                If Val(tFlex.Text) > 100 Then tFlex.Text = 100
            End If
        Case Else
            Exit Sub
    End Select
End Sub
Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iElem As Integer
    Dim iTmpType As iPropG
  
    iElem = miRow(1)
    
    Select Case tStrip.SelectedItem.Index - 1
        Case 2
            iTmpType = Elem(iElem).pgElement.iType
        Case 1
            iTmpType = keyElem(iElem).iType
        Case Else
            Exit Sub
    End Select
    
            
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
    
    Select Case tStrip.SelectedItem.Index - 1
        Case 2
            iTmpType = Elem(iElem).pgElement.iType
        Case 1
            iTmpType = keyElem(iElem).iType
        Case Else
            Exit Sub
    End Select
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
    Dim iTab As Integer

    If Not m_blnAnnullaMod Then
        iElem = miRow(1)
        tFlex = Trim$(tFlex)
        iTab = tStrip.SelectedItem.Index - 1
        If iTab = 2 Then
            If iElem = 7 Then
                lVal = Val(tFlex)
                If lVal <= 100 Then
                    Elem(7).pgElement.vVal = lVal
                    bUpdate = True
                Else
                    tFlex = 0
                End If
            End If
        ElseIf iTab = 1 Then
            If iElem = 2 Then
                lVal = Val(tFlex.Text)
                If lVal <= 100 Then
                    keyElem(2).vVal = lVal
                    bUpdate = True
                Else
                    tFlex = 0
                End If
            End If
        End If
    End If
    If bUpdate Then
        Flex(1).TextMatrix(iElem + 1, 1) = tFlex
        'If iTab = 2 Then
        '    If iElem = 7 Then SetEnabled
        'End If
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

    Dim i As Integer, l As Integer, iTab As Integer
  
    cbo.Clear
    cbo.AddItem ""
    
    Dim obj As cColIDX
    iTab = tStrip.SelectedItem.Index - 1

    If iTab > 1 Or Flex(0).Col = 1 Then
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
    Else
        For i = 1 To allRefCol.Count
            cbo.AddItem allRefCol(i)
        Next
        For l = 1 To colKey.Count
            Set obj = colKey.Item(l)
            
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
Private Sub tStrip_Click()

    Dim i As Integer
            
    i = tStrip.SelectedItem.Index - 1
    If i = 2 Then miCol = 0
    If bNewIndex Then bNewIndex = False
    lbl(0).Caption = MyLoadResString(k_Selected_Object)
    lbl(1).Visible = i < 2
    lbl(1).Caption = MyLoadResString(IIf(i = 0, k_Constraint_Expression, k_Primary_Key_Table))
    tCheck.Visible = i = 0
    Flex(0).Visible = i <> 0
    cboTB(0).Visible = i = 1
    cbo.Visible = False
    cmd(5).Visible = i <> 0
    
    
    LoadIndex ""
    tFlex.Visible = False
    SizeControls
    If m_bIsReadOnly Then
        cmd(1).Enabled = False
        cmd(2).Enabled = False
        cmd(3).Enabled = False
        cmd(4).Visible = False
    End If
End Sub
Private Sub LoadTBcombo(ByVal iIndex As Integer)

    Dim i As Integer
            
    cboTB(iIndex).Clear
    cboTB(iIndex).AddItem ""
    If iIndex = 0 Then
        For i = 0 To UBound(allTB)
            cboTB(iIndex).AddItem allTB(i)
        Next
    Else
        cboTB(1).AddItem m_sTbName
    End If
End Sub
Private Function UpdateConstraint(ByVal iTab As Integer, ByVal iElem As Integer) As Boolean

    Dim oObj As Object
    Dim lErr As Long, sErr As String
    Dim oTb As SQLDMO.Table
    
    Screen.MousePointer = vbHourglass
    Set oTb = GetTbByName(objServer, m_sDbName, m_sTbName)
    On Local Error Resume Next
    If iTab = 0 Then
        'Set oObj = objServer.Databases(sDbName).Tables(sTbName).Checks(cIndex.Text)
        Set oObj = oTb.Checks(cIndex.Text)
    Else
        'Set oObj = objServer.Databases(sDbName).Tables(sTbName).Keys(cIndex.Text)
        Set oObj = oTb.Keys(cIndex.Text)
    End If
    lErr = Err.Number
    
    If lErr = 0 Then
        Select Case iElem
            Case 0
                oObj.Checked = Not CBool(keyElem(iElem).vVal)
            Case 1
                oObj.ExcludeReplication = Not CBool(keyElem(iElem).vVal)
        End Select
        lErr = Err.Number
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    
    Set oObj = Nothing
    Set oTb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        UpdateConstraint = True
    End If
   
End Function
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

Private Sub ManageExtendedProp()

    Dim iMode As am_SqlPropType
    Dim iOwner As am_SqlPropTypeOwner
    Dim sParam As String
    Dim iTab As Integer
        
    iTab = tStrip.SelectedItem.Index - 1
    
    sParam = cIndex.Text
    iMode = am_ViewIdxProp
    iOwner = am_OwnTable
    
    
    If iTab = 1 Then
        iMode = am_KeyProp
    Else
        If Elem(0).pgElement.vVal Then
            iMode = am_KeyProp
        Else
            iMode = am_IdxProp
        End If
    End If
        
    If Not m_bIsOwner Or iTab = 0 Then Exit Sub
    
    LoadExtendedProp iMode, m_sDbName, GetObjectOwner(m_sDbName, m_sTbName, iOwner), m_sTbName, sParam
    
End Sub
Private Sub SaveIX()

    Dim i As Integer, iTab As Integer
    Dim bSave As Boolean, bKey As Boolean
    Dim sColumn As String
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim obj As cColIDX
    Dim oN_Key As New SQLDMO.Key
    
    Screen.MousePointer = vbHourglass
    iTab = tStrip.SelectedItem.Index - 1
    
    Select Case iTab
        Case 2
            bKey = Elem(0).pgElement.vVal Or Elem(3).pgElement.vVal
            
            If bKey Then
                
                oN_Key.Name = lblNewI.Caption
                If Elem(0).pgElement.vVal Then
                    oN_Key.Type = SQLDMOKey_Primary
                Else
                    oN_Key.Type = SQLDMOKey_Unique
                End If
                
                oN_Key.FileGroup = Elem(10).pgElement.vVal '"PRIMARY"
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    sBuf = Trim$(obj.colName)
                    If Len(sBuf) <> 0 Then oN_Key.KeyColumns.Add sBuf
                Next
                If oN_Key.KeyColumns.Count = 0 Then
                    lErr = vbObjectError + 2001
                    'sErr = MyLoadResString(k_Columns_not_defined) '"Campi Non Definiti"
                Else
                    oN_Key.Clustered = Elem(1).pgElement.vVal
                    oN_Key.FillFactor = Elem(7).pgElement.vVal
                End If
                'segue dopo
                'oTbl.Keys.Add oN_Key
            Else
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    sBuf = Trim$(obj.colName)
                    If Len(sBuf) <> 0 Then
                        bSave = True
                        sColumn = sColumn & IIf(Len(sColumn) <> 0, "," & sBuf, sBuf)
                    End If
                Next
                Set obj = Nothing
        
                bSave = Len(sColumn) <> 0
                If bSave Then
                    Dim oN_Ix As New SQLDMO.Index2
                                
                    With oN_Ix
                        .Name = lblNewI.Caption
                        .IndexedColumns = sColumn
                        
                        For i = 1 To colIdxCol.Count
                            Set obj = colIdxCol.Item(i)
                            sBuf = Trim$(obj.colName)
                            If Len(colIdxCol.Item(i).colName) <> 0 And colIdxCol.Item(i).isDescending Then
                                .SetIndexedColumnDESC colIdxCol.Item(i).colName, True
                            End If
                        Next
                                                
                        
                        '.FileGroup = "PRIMARY"
                        .FileGroup = Elem(10).pgElement.vVal
                                    
                        For i = 1 To 8
                            If i <> 7 And i <> 4 And i <> 3 Then
                                If Elem(i).pgElement.vVal Then .Type = native(i) Or .Type
                            End If
                        Next
                        .NoRecompute = Elem(4).pgElement.vVal
                        .FillFactor = Elem(7).pgElement.vVal
                    End With
                    'segue dopo
                    'oTbl.Indexes.Add oN_iX
                Else
                    lErr = vbObjectError + 2001
                    'sErr = "Campi Non Definiti"
                End If
            End If
        Case 1
            lErr = Err_Free
            sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
        Case 0      'Check
            tCheck.Text = Trim$(tCheck.Text)
            If Len(tCheck.Text) Then
                Dim oN_Check As New SQLDMO.Check
                
                oN_Check.Name = lblNewI.Caption
                oN_Check.Text = tCheck.Text
                oN_Check.ExcludeReplication = Not keyElem(1).vVal
            Else
                On Local Error Resume Next
                lErr = kErrProperty
                Err.Raise lErr
                sErr = Err.Description
                On Local Error GoTo 0
            End If
    End Select
            
    If lErr = 0 Then
        Dim oixType As Object
        
        If iTab <> 0 Then
            If bKey Then
                Set oixType = oN_Key
            Else
                Set oixType = oN_Ix
            End If
        Else
            Set oixType = oN_Check
        End If
        On Local Error Resume Next
        If iTab <> 0 Then
            If bKey Then
                m_oTbl.Keys.Add oixType
            Else
                m_oTbl.Indexes.Add oixType
            End If
        Else
            m_oTbl.Checks.Add oixType
        End If
        lErr = Err.Number
        If lErr = 0 Then
            sBuf = oixType.Name
            If iTab < 2 Then oixType.Checked = keyElem(0).vVal
            lErr = Err.Number
        End If
        sErr = Err.Description
        On Local Error GoTo 0
        sBuf = oixType.Name
    End If
    
    Set oN_Ix = Nothing
    Set oN_Key = Nothing
    Set oN_Check = Nothing
    Set oixType = Nothing
    
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
Private Sub SaveRefKey()

    Dim sSql As String
    Dim oDb As SQLDMO.Database2
    Dim obj As cColIDX
    Dim i As Integer
    Dim oTb As SQLDMO.Table
    
    Dim sColumns As String, sRefColumns As String
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    
    Dim sTables(1) As String
    Dim sOwner(1) As String
    Dim sKeyName As String
    Const kAddKey As String = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT [3%]" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    Const kNoCheckKey As String = "ALTER TABLE [1%].[2%] NOCHECK CONSTRAINT [3%]"

    'sSql = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT 3%" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    '        [ ON DELETE { CASCADE | NO ACTION } ]
    '        [ ON UPDATE { CASCADE | NO ACTION } ]
    '        [ NOT FOR REPLICATION ]

    
    sKeyName = lblNewI.Caption
    Screen.MousePointer = vbHourglass
    sRefColumns = ""
    sColumns = ""
    
    For i = 1 To colKey.Count
        Set obj = colKey.Item(i)
        sBuf = Trim$(obj.colName)
        If Len(sBuf) <> 0 Then
            If Len(sRefColumns) <> 0 Then sRefColumns = sRefColumns & " , "
            sRefColumns = sRefColumns & "[" & sBuf & "]"
        End If
    Next
    If Len(sRefColumns) = 0 Then lErr = vbObjectError + 2001
    
    If lErr = 0 Then
        For i = 1 To colIdxCol.Count
            Set obj = colIdxCol.Item(i)
            sBuf = Trim$(obj.colName)
            If Len(sBuf) <> 0 Then
                If Len(sColumns) <> 0 Then sColumns = sColumns & " , "
                sColumns = sColumns & "[" & sBuf & "]"
            End If
        Next
        If Len(sColumns) = 0 Then lErr = vbObjectError + 2001
    End If
    
    If lErr = 0 Then
        sTables(0) = Replace(m_sTbName, "[", "")
        sTables(0) = Replace(sTables(0), "]", "")
        sOwner(0) = GetObjectOwner(m_sDbName, sTables(0), am_OwnTable)
        
        sTables(1) = Replace(ReferTB, "[", "")
        sTables(1) = Replace(sTables(1), "]", "")
        sOwner(1) = GetObjectOwner(m_sDbName, sTables(1), am_OwnTable)
        
        If Len(sOwner(0)) = 0 Or Len(sOwner(1)) = 0 Then lErr = vbObjectError + 2002
    End If
    If lErr = 0 Then
    
        sSql = ReplaceMsg(kAddKey, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), _
                Array(sOwner(0), sTables(0), sKeyName, sColumns, _
                sOwner(1), sTables(1), sRefColumns))
                
        Debug.Print sSql
        If keyElem(2).vVal Then 'Update Cascade
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON UPDATE CASCADE"
        End If
        If keyElem(3).vVal Then 'Delete Cascade
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON DELETE CASCADE"
        End If
        If keyElem(1).vVal Then 'Not 4 Replication
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "NOT FOR REPLICATION"
        End If
        
        Debug.Print sSql
    End If
                
    If lErr = 0 Then
        Set oDb = objServer.Databases(m_sDbName)
        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
    End If
    If lErr = 0 Then
        If Not keyElem(0).vVal Then
            sSql = ReplaceMsg(kNoCheckKey, Array("1%", "2%", "3%"), Array(sOwner(0), sTables(0), sKeyName))
            lErr = ErrExexSqlDirect(oDb, sSql, sErr)
        End If
    End If
    
    Set obj = Nothing
    If lErr = 0 Then
        Set oTb = GetTbByName(objServer, m_sDbName, sTables(0))
        If Not oTb Is Nothing Then oTb.Keys.Refresh True
        Set oTb = Nothing
    End If
    Set oDb = Nothing
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        If lErr = vbObjectError + 2001 Then
            sErr = MyLoadResString(k_Columns_not_defined)
        ElseIf lErr = vbObjectError + 2002 Then
            sErr = MyLoadResString(k_Tables_not_defined)
        End If
        
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        
        cIndex.AddItem sKeyName
        cIndex.ListIndex = GetItem(sKeyName, cIndex)
    End If

End Sub
Private Sub ModifySQL()

    Dim i As Integer
    Dim bSave As Boolean
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim obj As cColIDX
    Dim sSql As String
    Dim sAttrib As String
    Dim sColumns As String, sRefColumns As String
    Dim sRet As String
    Dim iMode As Integer
    Dim bKey As Boolean
    Dim sNewObjName As String
    
    Dim sTables(1) As String
    Dim sOwner(1) As String
    
    Screen.MousePointer = vbHourglass
        
    iMode = tStrip.SelectedItem.Index - 1
    
    sRefColumns = ""
    sColumns = ""
    sSql = ""
    sBuf = ""
    sNewObjName = lblNewI.Caption
    
    Select Case iMode
        Case 0          'Check constraint
            sColumns = Trim$(tCheck.Text)
            If Len(tCheck.Text) = 0 Then
                lErr = vbObjectError + 2003
            Else
                sSql = "ALTER TABLE [" & GetTableOwner(objServer, m_sDbName, m_sTbName) & "].[" & m_sTbName & "] WITH NOCHECK ADD" & vbCrLf
                sSql = sSql & vbTab & "CONSTRAINT [" & sNewObjName & "] CHECK "
                If Not keyElem(1).vVal Then sSql = sSql & "NOT FOR REPLICATION "
                sSql = sSql & "( " & sColumns & " )"
            End If
        Case 1          'FK Relation

    Const kAddKey As String = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT [3%]" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    Const kNoCheckKey As String = "ALTER TABLE [1%].[2%] NOCHECK CONSTRAINT [3%]"

    'sSql = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT 3%" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    '        [ ON DELETE { CASCADE | NO ACTION } ]
    '        [ ON UPDATE { CASCADE | NO ACTION } ]
    '        [ NOT FOR REPLICATION ]

    
            For i = 1 To colKey.Count
                Set obj = colKey.Item(i)
                sBuf = Trim$(obj.colName)
                If Len(sBuf) <> 0 Then
                    If Len(sRefColumns) <> 0 Then sRefColumns = sRefColumns & " , "
                    sRefColumns = sRefColumns & "[" & sBuf & "]"
                End If
            Next
            If Len(sRefColumns) = 0 Then lErr = vbObjectError + 2001
            
            If lErr = 0 Then
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    sBuf = Trim$(obj.colName)
                    If Len(sBuf) <> 0 Then
                        If Len(sColumns) <> 0 Then sColumns = sColumns & " , "
                        sColumns = sColumns & "[" & sBuf & "]"
                    End If
                Next
                If Len(sColumns) = 0 Then lErr = vbObjectError + 2001
            End If
            
            If lErr = 0 Then
                sTables(0) = Replace(m_sTbName, "[", "")
                sTables(0) = Replace(sTables(0), "]", "")
                sOwner(0) = GetObjectOwner(m_sDbName, sTables(0), am_OwnTable)
                
                sTables(1) = Replace(ReferTB, "[", "")
                sTables(1) = Replace(sTables(1), "]", "")
                sOwner(1) = GetObjectOwner(m_sDbName, sTables(1), am_OwnTable)
                
                If Len(sOwner(0)) = 0 Or Len(sOwner(1)) = 0 Then lErr = vbObjectError + 2002
            End If
            If lErr = 0 Then
            
                sSql = ReplaceMsg(kAddKey, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), _
                        Array(sOwner(0), sTables(0), sNewObjName, sColumns, _
                        sOwner(1), sTables(1), sRefColumns))
                        
                Debug.Print sSql
                If keyElem(2).vVal Then 'Update Cascade
                    sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON UPDATE CASCADE"
                End If
                If keyElem(3).vVal Then 'Delete Cascade
                    sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON DELETE CASCADE"
                End If
                If keyElem(1).vVal Then 'Not 4 Replication
                    sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "NOT FOR REPLICATION"
                End If
                
                
                If Not keyElem(0).vVal Then
                    sSql = sSql & vbCrLf & vbCrLf & ReplaceMsg(kNoCheckKey, Array("1%", "2%", "3%"), Array(sOwner(0), sTables(0), sNewObjName))
                End If
                
            End If
        
        Case 2          'Idx + key
            bKey = Elem(0).pgElement.vVal Or Elem(3).pgElement.vVal
            If Not bKey Then
                sSql = "CREATE "
                If Abs(Elem(2).pgElement.vVal) Then sSql = sSql & "UNIQUE "
                If Abs(Elem(1).pgElement.vVal) Then sSql = sSql & "CLUSTERED "
                If Abs(Elem(8).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "PAD_INDEX "
                If Val(Elem(7).pgElement.vVal) <> 0 Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & " FILLFACTOR = " & Elem(7).pgElement.vVal
                If Abs(Elem(6).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "IGNORE_DUP_KEY "
                If Abs(Elem(4).pgElement.vVal) Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & "STATISTICS_NORECOMPUTE "
            
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    If Len(obj.colName) <> 0 Then
                        If Len(sColumns) Then sColumns = sColumns & ", "
                        sColumns = sColumns & "[" & Trim$(obj.colName) & "] "
                        If obj.isDescending Then sColumns = sColumns & "DESC "
                    End If
                Next
                If Len(sColumns) = 0 Then
                    lErr = vbObjectError + 2001
                    'sErr = "Campi Non Definiti"
                Else
                    sSql = sSql & "INDEX [" & sNewObjName & "] ON [" & GetTableOwner(objServer, m_sDbName, m_sTbName) & "].[" & m_sTbName & "] "
                    sSql = sSql & vbCrLf & "( " & sColumns & " ) "
                    sSql = sSql & IIf(Len(sAttrib), vbCrLf & " WITH ", "") & sAttrib
            
                    sSql = sSql & vbCrLf & " ON [" & Elem(10).pgElement.vVal & "]"
                End If
            Else
                sSql = "ALTER TABLE [" & GetTableOwner(objServer, m_sDbName, m_sTbName) & "].[" & m_sTbName & "] WITH NOCHECK ADD" & vbCrLf
                sSql = sSql & vbTab & "CONSTRAINT [" & sNewObjName & "]" & vbCrLf
                If Abs(Elem(0).pgElement.vVal) Then     'Primary Key
                    sSql = sSql & vbTab & "PRIMARY KEY "
                    If Abs(Elem(1).pgElement.vVal) Then
                        sSql = sSql & "CLUSTERED "
                    Else
                        sSql = sSql & "NONCLUSTERED "
                    End If
                Else                                    'Constraint
                    sSql = sSql & vbTab & "UNIQUE "
                    If Abs(Elem(1).pgElement.vVal) Then
                        sSql = sSql & "CLUSTERED "
                    Else
                        sSql = sSql & "NONCLUSTERED "
                    End If
                End If
                If Val(Elem(7).pgElement.vVal) <> 0 Then sAttrib = sAttrib & IIf(Len(sAttrib) <> 0, ", ", "") & " FILLFACTOR = " & Elem(7).pgElement.vVal
            
                For i = 1 To colIdxCol.Count
                    Set obj = colIdxCol.Item(i)
                    If Len(obj.colName) <> 0 Then
                        If Len(sColumns) Then sColumns = sColumns & ", "
                        sColumns = sColumns & "[" & Trim$(obj.colName) & "] "
                        If Not Abs(Elem(0).pgElement.vVal) Then     'Unique constraint
                            If obj.isDescending Then sColumns = sColumns & "DESC "
                        End If
                    End If
                Next
                
                If Len(sColumns) = 0 Then
                    lErr = vbObjectError + 2001
                    'sErr = "Campi Non Definiti"
                Else
                    sSql = sSql & vbCrLf & vbTab & vbTab & "( " & sColumns & " ) "
                    sSql = sSql & IIf(Len(sAttrib), vbCrLf & " WITH ", "") & sAttrib
            
                    sSql = sSql & vbCrLf & " ON [" & Elem(10).pgElement.vVal & "]"
                End If
            
            End If
    End Select
    
    If lErr = 0 Then
        Screen.MousePointer = vbDefault
        With fResult
            .Action() = act_Modify_Idx_Sql
            .Caption = MyLoadResString(k_Frm_ModifySQL)
            .tRes.Text = sSql
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

            oDb.Tables.Refresh True
            Set oDb = Nothing
            Set m_oTbl = GetTbByName(objServer, m_sDbName, m_sTbName)
        End If
    End If

    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        If lErr = vbObjectError + 2001 Then
            sErr = MyLoadResString(k_Columns_not_defined)
        ElseIf lErr = vbObjectError + 2002 Then
            sErr = MyLoadResString(k_Tables_not_defined)
        ElseIf lErr = vbObjectError + 2003 Then
            sErr = MyLoadResString(k_Check_not_defined)
        End If

        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        If bSave Then
            cIndex.AddItem sNewObjName
            cIndex.ListIndex = GetItem(sNewObjName, cIndex)
        End If
    End If
    
End Sub


