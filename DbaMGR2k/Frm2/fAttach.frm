VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fAttach 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   ClipControls    =   0   'False
   Icon            =   "fAttach.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList iList 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":05BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":0B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAttach.frx":0E88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   7815
      Begin VB.TextBox tDati 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   4
         Top             =   120
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1695
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label lblDB 
         Alignment       =   1  'Right Justify
         Caption         =   "DB &Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   6615
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1935
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Transaction Log"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMsg 
      Caption         =   "Label2"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "fAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type elSetting
    Elem As PageElement
    EditType As TypeEnum
    bDisab As Boolean
    iComboType As amCbo2Read
End Type
Private m_lWidth As Long
Private m_lHeight As Long
Private m_blnAnnullaMod As Boolean
Private miRow(1) As Integer
Private sWorkPath As String
Private sSqlPath As String
Private m_objDataF As cDetGrp
Private m_sPhysicalName As String
Private collFiles As Collection
Private bLoading As Boolean
Private bDone As Boolean
Private sDbName As String
Private optElem() As elSetting
Private m_bIsLocalHost As Boolean
Private m_oCol As Collection

Public Property Get Done() As Boolean
    Done = bDone
End Property
Public Property Let Stuff(ByVal sPhysicalName As String, ByVal objDataF As cDetGrp, ByVal sName4DB As String)
    
    m_bIsLocalHost = StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0
    sDbName = Trim$(sName4DB)
    m_sPhysicalName = sPhysicalName
    tDati.Text = sDbName
    lblMsg.Caption = ReplaceMsg(MyLoadResString(k_Attaching_MSG), Array("1%", "2%", "|"), Array(objServer.Name, sDbName, vbCrLf))
    
    Set m_objDataF = objDataF
    Set collFiles = New Collection

End Property

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim iImg As Integer, iElem As Integer
    Dim iTmpType As iPropG
    Dim bCombo As Boolean
        
    bEdit = True

    iElem = Flex(1).Row - 1
    iTmpType = optElem(iElem).Elem.iType
    bCombo = optElem(iElem).EditType = iCombo
    If Flex(1).CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        Select Case KeyAscii
            Case 0, vbKeySpace
                If (iTmpType And am_Bool) = am_Bool Then
                    iImg = 1 - (-CBool(optElem(iElem).Elem.vVal))
                    optElem(iElem).Elem.vVal = iImg
                                    
                    Set Flex(1).CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                    Flex(1).CellPictureAlignment = flexAlignCenterCenter
                    bEdit = False
                End If
            Case Else
                If (iTmpType And am_Bool) = am_Bool Then bEdit = False
        End Select
    End If
    If bEdit Then
        If bCombo Then
            PreloadCBO optElem(iElem).iComboType
            cbo.ListIndex = GetItem(optElem(iElem).Elem.vVal, cbo)
                        
            With Flex(1)
                cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            End With
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
    End If
End Sub
Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub cbo_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyEscape
            m_blnAnnullaMod = True
            Flex(1).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex(1).SetFocus
            KeyAscii = 0
    End Select

End Sub
Private Sub cbo_LostFocus()
    
    Dim iElem As Integer
    Dim bUpdate As Boolean
    
    If Not m_blnAnnullaMod Then
        iElem = miRow(1) - 1
        bUpdate = True
        
        optElem(iElem).Elem.iLen = cbo.ItemData(cbo.ListIndex)
        optElem(iElem).Elem.vVal = cbo.Text
    End If
    If bUpdate Then
        Flex(1).TextMatrix(iElem + 1, Flex(1).Cols - 1) = cbo.Text
    Else
        If Not m_blnAnnullaMod Then Beep
    End If

    cbo.Visible = False
    Flex(1).SetFocus
    m_blnAnnullaMod = True
    
End Sub

Private Sub Flex_DblClick(Index As Integer)
    If Index Then
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then
            FlexEdit vbKeySpace
        End If
    End If
End Sub

Private Sub Flex_GotFocus(Index As Integer)
    TestActiveControl
    cbo.Visible = False
    m_blnAnnullaMod = True
End Sub

Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If Index Then
        If KeyCode = vbKeyF2 Then FlexEdit 0
    End If
End Sub
Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    
    If Index Then
        iCode = KeyAscii
        KeyAscii = 0
                
        FlexEdit iCode
    End If
    
End Sub
Private Sub Flex_RowColChange(Index As Integer)
    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
       
    If Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            With Flex(Index)
                If .Row <> miRow(Index) Or (.RowData(.Row) = 0) Then
                    
                    iCol = .Col
                    iRowP = .Row
                    If miRow(Index) > .Rows - 1 Then miRow(Index) = .Rows - 1
                    .Row = miRow(Index)
                    .Col = 0
                    Set Flex(Index).CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1

                    .Col = iCol
                    miRow(Index) = iRowP
                    If Index = 0 Then tBar.Buttons(keyDelete).Enabled = miRow(Index) > 1
                End If
            End With
            b = False
        End If
    End If

End Sub
Private Sub Flex_Scroll(Index As Integer)
    cbo.Visible = False
End Sub

Private Sub Form_Activate()

    Static b As Boolean
    If Not b Then
        b = True
        LoadFlex 0
        LoadFlex 1
        tBar.Refresh
        fra(0).Refresh
        Flex(0).Refresh
    End If
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
    
    Dim i As Integer
    
    m_lWidth = Me.Width
    m_lHeight = Me.Height

    tDati.MaxLength = k_MaxDBLen
    Flex(1).RowHeightMin = cbo.Height
    
    InitDBOption
     
    Set collFiles = New Collection
    Me.Caption = MyLoadResString(k_Attach_DB_to_Server)
    tStrip.Tabs(1).Caption = MyLoadResString(k_tab_Files_Selection)
    tStrip.Tabs(2).Caption = MyLoadResString(k_tab_Post_Attach_DB_Options)
        
    tDati.ToolTipText = MyLoadResString(k_Attaching_DB_Name)
    lblDB.Caption = MyLoadResString(k_DB_Name_for_reattach)
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keyRun).ToolTipText = MyLoadResString(k_Exec_ReAttach)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_RES_FitGrid)
    tBar.Buttons(keyNew).ToolTipText = MyLoadResString(k_AddNewFile)
    tBar.Buttons(keyDelete).ToolTipText = MyLoadResString(k_RemoveAddedFile)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource)
    
    sSqlPath = objServer.Registry.SQLDataRoot & "\"
    sWorkPath = GetIni("DATA", "DEFDIR", App.Path & "\" & App.EXEName & ".INI")
    If m_bIsLocalHost Then
        If sWorkPath <> "..." Then
            On Local Error Resume Next
            i = Len(Dir$(sWorkPath))
            i = Err.Number <> 0
            On Local Error GoTo 0
        Else
            i = True
        End If
        If i <> 0 Then sWorkPath = sSqlPath
    End If
    If Len(sWorkPath) = 0 Then sWorkPath = sSqlPath

    CenterForm Me
End Sub
Private Sub Form_Resize()
    
    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            TestActiveControl
            On Error Resume Next
            
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            lblMsg.Width = Me.ScaleWidth - 2 * (lblMsg.Left)
            
            tStrip.Width = Me.ScaleWidth - 2 * (tStrip.Left)
            tStrip.Height = Me.ScaleHeight - tStrip.Top
            fra(0).Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
            fra(1).Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
            
            Flex(0).Width = fra(0).Width - 2 * (Flex(0).Left)
            Flex(1).Width = Flex(0).Width
            
            Flex(0).Height = fra(0).Height - Flex(0).Top - 50
            Flex(1).Height = fra(1).Height - Flex(1).Top - 50
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set m_objDataF = Nothing
    Set m_oCol = Nothing
    Set collFiles = Nothing
    Set fAttach = Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    TestActiveControl
    Select Case Button.Key
        Case keyEnd
            Me.Hide
        Case keyRun
            ExecAttach
        Case keyFitGrid
            LoadFlex 0
            LoadFlex 1
        Case keyNew
            Add1Row
        Case keyDelete
            DeleteRow
        Case keyRefresh
            RefreshData
    End Select
    
End Sub

Private Sub LoadFlex(ByVal Index As Integer)
       
    Dim i As Integer, l As Integer
    Dim sBuf As String
    Dim arr_Bflags(1) As Boolean
    Dim sNotAvailable As String
    Dim iImg As Integer
    Dim iCol2Enlarge As Integer
    
    bLoading = True
    
    Screen.MousePointer = vbHourglass
    
    With Flex(Index)
        .Redraw = False
        .Clear
        .Rows = 1
        If Index = 0 Then
            iCol2Enlarge = 2
            .Cols = 9
            
            '.TextArray(1) = MyLoadResString(k_Name)
            .TextArray(2) = MyLoadResString(k_Name)
            .TextArray(3) = MyLoadResString(k_File_Size)
            .TextArray(4) = MyLoadResString(k_Created)
            .TextArray(5) = MyLoadResString(k_Last_Modified)
            .TextArray(6) = MyLoadResString(k_Last_Access)
            .TextArray(7) = MyLoadResString(k_Read_Only)
            .TextArray(8) = MyLoadResString(k_Archive)
            
            .ColWidth(3) = 2000
            .ColAlignment(2) = flexAlignLeftCenter
            .ColWidth(1) = k_FlexPARAM_cellaSelector
        Else
            iCol2Enlarge = 1
            .Cols = 3
            .TextMatrix(0, 1) = MyLoadResString(k_Property)
            .TextMatrix(0, 2) = MyLoadResString(k_Value)
            .ColWidth(2) = 2000
        End If
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        
        
        If Index = 0 Then
            sNotAvailable = MyLoadResString(kNotAvailable)
            For i = 1 To m_objDataF.m_FileColl.Count
                If m_bIsLocalHost And m_objDataF.m_FileColl.Item(i).FileExist() Then
                
                    sBuf = vbTab & vbTab & m_objDataF.m_FileColl.Item(i).oFileFinfo.FullPathName() & vbTab & m_objDataF.m_FileColl.Item(i).oFileFinfo.FormatFileSize(m_objDataF.m_FileColl.Item(i).oFileFinfo.FileSize) & vbTab & m_objDataF.m_FileColl.Item(i).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(i).oFileFinfo.CreationTime) & vbTab & m_objDataF.m_FileColl.Item(i).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(i).oFileFinfo.ModifyTime) & vbTab & m_objDataF.m_FileColl.Item(i).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(i).oFileFinfo.LastAccessTime)
                    arr_Bflags(0) = m_objDataF.m_FileColl.Item(i).oFileFinfo.attrReadOnly
                    arr_Bflags(1) = m_objDataF.m_FileColl.Item(i).oFileFinfo.attrArchive
                
                Else
                    sBuf = vbTab & vbTab & m_objDataF.m_FileColl.Item(i).PhysicalName & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable
                End If
                .AddItem sBuf
                
                .Row = .Rows - 1
                
                .Col = 1
                Set Flex(Index).CellPicture = LoadResPicture(IIf(m_objDataF.m_FileColl.Item(i).FileExist, k_ResChkOFF, k_ResChkDel), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                If m_bIsLocalHost And m_objDataF.m_FileColl.Item(i).FileExist() Then
                    iImg = IIf(arr_Bflags(0), k_ResChkON, k_ResChkOFF)
                Else
                    iImg = k_ResChkDel
                End If
                .Col = 7
                Set Flex(Index).CellPicture = LoadResPicture(iImg, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                If m_bIsLocalHost And m_objDataF.m_FileColl.Item(i).FileExist() Then
                    iImg = IIf(arr_Bflags(1), k_ResChkON, k_ResChkOFF)
                Else
                    iImg = k_ResChkDel
                End If
                .Col = 8
                Set Flex(Index).CellPicture = LoadResPicture(iImg, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                If i = 1 Then
                    For l = 1 To .Cols - 1
                        .Col = l
                        .CellBackColor = vbInactiveBorder
                    Next
                End If
            Next
            .FixedCols = 1
        Else
            For i = 0 To UBound(optElem)
                sBuf = ""
                If optElem(i).Elem.iType <> am_Bool Then
                    If optElem(i).EditType = iCombo Then
                        sBuf = GetComboDescription(optElem(i).Elem.iLen, optElem(i).iComboType)
                    End If
                End If
                .AddItem vbTab & optElem(i).Elem.sName & vbTab & sBuf
            
                
                .Row = .Rows - 1
                .Col = .Cols - 1
                If optElem(i).Elem.iType = am_Bool Then
                    Set Flex(Index).CellPicture = LoadResPicture(IIf(optElem(i).Elem.vVal, k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                
                If optElem(i).EditType = iCombo Then
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
                    Flex(Index).CellPictureAlignment = flexAlignRightTop
                End If
                                
            Next
            .FixedCols = 2
        End If
        .FixedRows = 1
        .Redraw = True
        SettaLargFlex Flex(Index), iCol2Enlarge, 200
    End With
    miRow(Index) = 0
    
    bLoading = False
    Flex(Index).Row = 1
    Flex(Index).TopRow = 1
    Flex_RowColChange Index
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Add1Row()
        
    Dim iC As Integer, sBuf As String
    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim arr_Bflags(1) As Boolean
    Dim sNotAvailable As String
    
    Dim bCancel As Boolean
    Dim bErr As Boolean
    
    Dim obj As cOpen
    Set obj = New cOpen
    
    If collFiles.Count < 17 Then
    
        If m_bIsLocalHost Then
            obj.Title = MyLoadResString(k_Select_Data_File)
            obj.Def_Dir() = sWorkPath
            
            obj.Filter() = "Data File (*.Mdf)" & Chr$(0) & "*.Mdf" & Chr$(0) & "Sec .Data File (*.Ndf)" & Chr$(0) & "*.Ndf" & Chr$(0) & "Log File (*.Ldf)" & Chr$(0) & "*.Ldf" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
            obj.formHwnd() = Me.hWnd
            obj.Do_Open
            lErr = obj.ErrCode()
            If lErr = 0 Then
                sFile = obj.FullPath()
            Else
                If lErr = 1 Then
                    lErr = 0
                Else
                    sErr = obj.ErrMsg()
                End If
            End If
            Set obj = Nothing
        Else
        
            Dim vntFiltri(1, 3) As Variant
            
            vntFiltri(0, 0) = "*.Mdf"
            vntFiltri(1, 0) = "Data File (*.Mdf)"
            
            vntFiltri(0, 1) = "*.Ndf"
            vntFiltri(1, 1) = "Sec .Data File (*.Ndf)"
            
            vntFiltri(0, 2) = "*.Ldf"
            vntFiltri(1, 2) = "Log File (*.Ldf)"
            
            vntFiltri(0, 3) = "*.*"
            vntFiltri(1, 3) = MyLoadResString(k_All_Files)
            
            Dim fBrowse As fBr4Fold
            Set fBrowse = New fBr4Fold
            
            With fBrowse
                .FileMustExist() = True
                .Init enB4SelectFileFilters, MyLoadResString(k_Select_Data_File), sWorkPath, , vntFiltri
                            
                bErr = .ErrLoadingInit()
                If Not bErr Then
                    .Show vbModal, Me
                    bCancel = .Cancelled()
                    If Not bCancel Then sFile = .FullPath()
                Else
                    .GetError lErr, sErr
                End If
            End With
            Unload fBrowse
            Set fBrowse = Nothing
        
        End If
        If lErr = 0 Then
            If Len(sFile) Then
                m_objDataF.AddFile2Col sFile, True
                If m_objDataF.AreErrors() Then
                    lErr = m_objDataF.Errors(sErr)
                End If
                If lErr = 0 Then
                    iC = m_objDataF.m_FileColl.Count
                    bLoading = True
                    If Not m_bIsLocalHost Then sNotAvailable = MyLoadResString(kNotAvailable)
                                        
                    With Flex(0)
                        If m_bIsLocalHost Then
                            sBuf = vbTab & vbTab & m_objDataF.m_FileColl.Item(iC).oFileFinfo.FullPathName() & vbTab & m_objDataF.m_FileColl.Item(iC).oFileFinfo.FormatFileSize(m_objDataF.m_FileColl.Item(iC).oFileFinfo.FileSize) & vbTab & m_objDataF.m_FileColl.Item(iC).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(iC).oFileFinfo.CreationTime) & vbTab & m_objDataF.m_FileColl.Item(iC).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(iC).oFileFinfo.ModifyTime) & vbTab & m_objDataF.m_FileColl.Item(iC).oFileFinfo.FormatFileDate(m_objDataF.m_FileColl.Item(iC).oFileFinfo.LastAccessTime)
                            arr_Bflags(0) = m_objDataF.m_FileColl.Item(iC).oFileFinfo.attrReadOnly
                            arr_Bflags(1) = m_objDataF.m_FileColl.Item(iC).oFileFinfo.attrArchive
                        Else
                            sBuf = vbTab & vbTab & m_objDataF.m_FileColl.Item(iC).PhysicalName() & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable & vbTab & sNotAvailable
                        End If
                        .AddItem sBuf
                        
                        .Row = .Rows - 1
                        .Col = 1
                        Set Flex(0).CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                        .Col = 7
                        Set Flex(0).CellPicture = LoadResPicture(IIf(arr_Bflags(0), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                        .Col = 8
                        Set Flex(0).CellPicture = LoadResPicture(IIf(arr_Bflags(1), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                    End With
                    bLoading = False
                End If
            End If
        End If
    End If
    fra(tStrip.SelectedItem.Index - 1).Refresh
    Flex(tStrip.SelectedItem.Index - 1).Refresh
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If
   
End Sub
Private Sub DeleteRow()

    Dim iCol As Integer, iRow As Integer
    
    iCol = Flex(0).Col
    iRow = Flex(0).Row
    
    m_objDataF.m_FileColl.Remove iRow
        
    bLoading = True
    Flex(0).RemoveItem iRow
    miRow(0) = miRow(0) - 1
    bLoading = False
    Flex_RowColChange 0
    
End Sub

Private Sub tDati_Change()
    If Not bDone Then
        tBar.Buttons(keyRun).Enabled = Len(Trim$(tDati.Text))
    End If
End Sub
Private Sub ExecAttach()

    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim oOpt As SQLDMO.DBOption2
    Dim sDbName As String
    Dim sFiles As String
    Dim i As Integer
    Dim bNotFind As Boolean
    
    sDbName = Trim$(tDati.Text)
    Screen.MousePointer = vbHourglass
    
        
    For i = 1 To m_objDataF.m_FileColl.Count
        If i > 1 Then sFiles = sFiles & ";"
        If m_bIsLocalHost And m_objDataF.m_FileColl.Item(i).FileExist() Then
            sFiles = sFiles & "[" & m_objDataF.m_FileColl.Item(i).oFileFinfo.FullPathName() & "]"
        Else
            sFiles = sFiles & "[" & m_objDataF.m_FileColl.Item(i).PhysicalName() & "]"
        End If
        
        If Not m_objDataF.m_FileColl.Item(i).FileExist() Then
            lErr = Err_Free
            sErr = IIf(Len(sErr) <> 0, vbCrLf, "") & m_objDataF.m_FileColl.Item(i).PhysicalName()
            bNotFind = True
        End If
    Next
    
    If lErr = 0 Then
        On Local Error Resume Next
        fResult.tRes.Text = objServer.AttachDB(sDbName, sFiles)
        bDone = Err.Number = 0
        On Local Error GoTo 0
        
        tBar.Buttons(keyRun).Enabled = Not bDone
    End If
    
    fResult.Action() = act_Null
    fResult.Caption = "Attach Database '" & sDbName & "'"

    If bDone Then
        On Local Error Resume Next
        Set oDb = objServer.Databases(sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        If lErr = 0 Then
            Set oOpt = oDb.DBOption
            For i = 0 To UBound(optElem)
                If i <> 0 And i <> 2 Then
                    fResult.tRes.Text = fResult.tRes.Text & vbCrLf & ReplaceMsg(MyLoadResString(k_Option_Setting), Array("1%", "2%"), Array(optElem(i).Elem.sName, CBool(optElem(i).Elem.vVal)))
                Else
                    fResult.tRes.Text = fResult.tRes.Text & vbCrLf & ReplaceMsg(MyLoadResString(k_Option_Setting), Array("1%", "2%"), Array(optElem(i).Elem.sName, optElem(i).Elem.vVal))
                End If
                On Local Error Resume Next
                Select Case i
                    Case 0
                        oDb.DBOption.DBOUseOnly = optElem(0).Elem.iLen = amAccessDBO
                        oDb.DBOption.SingleUser = optElem(0).Elem.iLen = amAccessSingle
                    Case 1
                        oDb.DBOption.ReadOnly = Abs(optElem(1).Elem.vVal)
                    Case 2
                        oOpt.RecoveryModel = optElem(2).Elem.iLen
                    Case 3
                        oDb.DBOption.ColumnsNullByDefault = Abs(optElem(3).Elem.vVal)
                    Case 4
                        oDb.DBOption.RecursiveTriggers = Abs(optElem(4).Elem.vVal)
                    Case 5
                        oDb.DBOption.AutoCreateStat = Abs(optElem(5).Elem.vVal)
                    Case 6
                        oDb.DBOption.AutoUpdateStat = Abs(optElem(6).Elem.vVal)
                    Case 7
                        oDb.DBOption.TornPageDetection = Abs(optElem(7).Elem.vVal)
                    Case 8
                        oDb.DBOption.AutoClose = Abs(optElem(8).Elem.vVal)
                    Case 9
                        oDb.DBOption.AutoShrink = Abs(optElem(9).Elem.vVal)
                    Case 10
                        oDb.DBOption.QuoteDelimiter = Abs(optElem(10).Elem.vVal)
                End Select
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                If lErr <> 0 Then
                    fResult.tRes.Text = fResult.tRes.Text & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                End If
            Next
            bDone = True
        Else
            fResult.tRes.Text = fResult.tRes.Text & vbCrLf & MyLoadResString(k_RES_Object_Not_Found_simple) & ": " & sDbName & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        End If
        On Local Error GoTo 0
    Else
        If bNotFind Then
            sErr = Replace(MyLoadResString(k_DbDetached_FileNotFound), "1%", sDbName) & vbCrLf & String$(35, "-") & vbCrLf & sErr
        Else
            sErr = MyLoadResString(k_RES_Object_Not_Found_simple) & ": " & sDbName & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        End If
            
        fResult.tRes.Text = fResult.tRes.Text & vbCrLf & sErr
    End If
    
    Set oOpt = Nothing
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    fResult.Show vbModal, Me
    tBar.Refresh
    DoEvents

End Sub
Private Sub InitDBOption()

    ReDim optElem(10) As elSetting
    
    optElem(0).EditType = iNormalEdit
    optElem(0).Elem.sName = MyLoadResString(k_DbOpt_AccessGrantedTo) '"Accesso consentito a"
    optElem(0).Elem.iLen = 0
    optElem(0).Elem.vVal = GetComboDescription(optElem(0).Elem.iLen, amRestricAccess)
    optElem(0).Elem.iType = am_String
    optElem(0).EditType = iCombo
    optElem(0).iComboType = amRestricAccess
        
    optElem(1).EditType = iNormalEdit
    optElem(1).Elem.sName = MyLoadResString(k_DBStatus_ReadOnly)
    optElem(1).Elem.vVal = False
    optElem(1).Elem.iType = am_Bool
        
        
    optElem(2).Elem.sName = MyLoadResString(k_DbOpt_RecoveryModel)
    optElem(2).Elem.iLen = SQLDMORECOVERY_Simple
    optElem(2).Elem.vVal = GetComboDescription(optElem(2).Elem.iLen, amRecovery)
    optElem(2).Elem.iType = am_String
    optElem(2).EditType = iCombo
    optElem(2).iComboType = amRecovery
                    
    optElem(3).EditType = iNormalEdit
    optElem(3).Elem.sName = MyLoadResString(k_DbOpt_AnsiNullDefault) '"Valore Predefinito ANSI NULL"
    optElem(3).Elem.vVal = False
    optElem(3).Elem.iType = am_Bool
                    
    optElem(4).EditType = iNormalEdit
    optElem(4).Elem.sName = MyLoadResString(k_DbOpt_RecursiveTriggers) '"Trigger Ricorsivi"
    optElem(4).Elem.vVal = False
    optElem(4).Elem.iType = am_Bool
                
    optElem(5).EditType = iNormalEdit
    optElem(5).Elem.sName = MyLoadResString(k_DbOpt_AutoCreateStat) '"Creazione Automatica Statistiche"
    optElem(5).Elem.vVal = True
    optElem(5).Elem.iType = am_Bool

    optElem(6).EditType = iNormalEdit
    optElem(6).Elem.sName = MyLoadResString(k_DbOpt_AutoUpdateStat) '"Aggiornamento Automatico Statistiche"
    optElem(6).Elem.vVal = True
    optElem(6).Elem.iType = am_Bool

    optElem(7).EditType = iNormalEdit
    optElem(7).Elem.sName = MyLoadResString(k_DbOpt_TornPageDetection) '"Rilevamento pagine incomplete"    '"TornPageDetection"
    optElem(7).Elem.vVal = True
    optElem(7).Elem.iType = am_Bool

    optElem(8).EditType = iNormalEdit
    optElem(8).Elem.sName = MyLoadResString(k_DbOpt_AutoClose) '"Chiusura Automatica"
    optElem(8).Elem.vVal = True
    optElem(8).Elem.iType = am_Bool

    optElem(9).EditType = iNormalEdit
    optElem(9).Elem.sName = MyLoadResString(k_DbOpt_AutoShrink) '"Compattazione Automatica"     'Auto Shrink
    optElem(9).Elem.vVal = True
    optElem(9).Elem.iType = am_Bool

    optElem(10).EditType = iNormalEdit
    optElem(10).Elem.sName = MyLoadResString(k_DbOpt_UseQuotedIdentifier) '"Usa Identificatori tra virgolette"     'Use quoted identifier
    optElem(10).Elem.vVal = False
    optElem(10).Elem.iType = am_Bool

End Sub

Private Sub tStrip_Click()
    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    fra(0).Visible = i = 0
    fra(1).Visible = i <> 0
    
    fra(i).ZOrder
    fra(i).Refresh
    Flex(i).Refresh
    If i = 0 Then
        tDati.SetFocus
    Else
        Flex(1).SetFocus
    End If
    tBar.Buttons(keyNew).Enabled = i = 0
    tBar.Buttons(keyDelete).Enabled = i = 0 And miRow(0) > 1
    Flex_RowColChange i

End Sub
Private Function GetComboDescription(ByVal iVal As Integer, ByVal iOpt As amCbo2Read) As String

    Dim sBuf As String
    If iOpt = amRestricAccess Then
        If iVal = amAccessAll Then
            sBuf = MyLoadResString(k_DbOpt_Access_ALL) '"All"
        ElseIf iVal = amAccessDBO Then
            sBuf = MyLoadResString(k_DbOpt_Access_Dbo)  '"Dbo"
        ElseIf iVal = amAccessSingle Then
            sBuf = MyLoadResString(k_DBStatus_SingleUser) '"Single"
        End If
    ElseIf iOpt = amRecovery Then
        If iVal = SQLDMORECOVERY_Simple Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Simple) '"Semplice"
        ElseIf iVal = SQLDMORECOVERY_BulkLogged Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Bulklogged) '"Con Registrazioni di Massa"
        ElseIf iVal = SQLDMORECOVERY_Full Then
            sBuf = MyLoadResString(k_DbOpt_Recovery_Full) '"Completo"
        End If
    End If
    
    GetComboDescription = sBuf
End Function
Private Sub PreloadCBO(ByVal iOpt As amCbo2Read)

    Dim i As Integer
    Dim sBuf As String
    
    With cbo
        .Clear
        If iOpt = amRecovery Then
            For i = SQLDMORECOVERY_Simple To SQLDMORECOVERY_Full
                sBuf = GetComboDescription(i, amRecovery)
                .AddItem sBuf
                .ItemData(.NewIndex) = i
            Next
        ElseIf iOpt = amRestricAccess Then
            For i = amAccessAll To amAccessSingle
                sBuf = GetComboDescription(i, amRestricAccess)
                .AddItem sBuf
                .ItemData(.NewIndex) = i
            Next
        End If
    End With
    
End Sub

Private Sub TestActiveControl()

    If cbo.Visible Then
        cbo_LostFocus
    End If

    On Local Error Resume Next
    If TypeOf Me.ActiveControl Is ComboBox Then
        If Me.ActiveControl Is cbo Then
            cbo_LostFocus
        End If
    End If
    On Local Error GoTo 0

End Sub

Private Sub RefreshData()

    Dim lErr As Long, sErr As String
    Dim tmp_DataF As cDetGrp
    
    Set tmp_DataF = New cDetGrp
    tmp_DataF.PrimaryMdfFile() = m_sPhysicalName
    If tmp_DataF.AreErrors() Then
        lErr = tmp_DataF.Errors(sErr)
    Else
        Me.Stuff(m_sPhysicalName, tmp_DataF) = sDbName
        If Not bDone Then InitDBOption
        LoadFlex 0
        LoadFlex 1
    End If
    Set tmp_DataF = Nothing
    
    If lErr <> 0 Then
        Me.Hide
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If

End Sub
