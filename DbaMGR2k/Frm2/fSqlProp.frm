VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fSqlProp 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   ClipControls    =   0   'False
   Icon            =   "fSqlProp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   7095
      Begin VB.TextBox TxtProp 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox rtfProp 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2355
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"fSqlProp.frx":000C
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2328
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   6240
      Top             =   0
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
            Picture         =   "fSqlProp.frx":008E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSqlProp.frx":04E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSqlProp.frx":0642
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSqlProp.frx":095C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSqlProp.frx":0AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSqlProp.frx":0C10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4110
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13123
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fSqlProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' max col Name = SYSNAME 128
' max col Description= variant 7500 byte
Private Enum am_PropStatus
    am_PropNONE = 0
    am_PropEdit
    am_PropNew
End Enum


Private m_bObjChanged(1) As Boolean
Private m_Lerr As Long
Private m_sErr As String

Private m_bLoading As Boolean
Private m_lWidth As Long
Private m_lHeigth As Long
Private m_sDbName As String
Private m_sOwner As String
Private m_s1LevelObj As String
Private m_s2LevelObj As String
Private m_iobjType As am_SqlPropType
Private mi_Row As Integer
Private m_iStatus As am_PropStatus

Public Sub Stuff(ByVal iObjType As am_SqlPropType, ByVal sDbName As String, ByVal sOwner As String, ByVal s1LevelObj As String, ByVal s2LevelObj As String)
    m_iobjType = iObjType
    m_sDbName = sDbName
    m_sOwner = sOwner
    m_s1LevelObj = s1LevelObj
    m_s2LevelObj = s2LevelObj
    
    LoadProperty
    SetFormCaption
    
End Sub
Public Property Get Errors(ByRef sErr As String) As Long
    Errors = m_Lerr
    sErr = m_sErr
End Property

Private Sub Flex_GotFocus()
    StBar.Panels(1).Text = MyLoadResString(k_PROP_Form_Available_Property)
End Sub

Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
       
    If Not m_bLoading Then
        If Not b Then
            b = True
            
            With Flex
                If .Row <> mi_Row Or (.RowData(.Row) = 0) Then
                    
                    iCol = .Col
                    iRowP = .Row
                    If mi_Row > .Rows - 1 Then mi_Row = .Rows - 1
                    .Row = mi_Row
                    .Col = 0
                    Set Flex.CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    If iRowP > 0 Then
                        Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                    End If
                    .RowData(.Row) = 1

                    .Col = iCol
                    mi_Row = iRowP
                    
                    LoadRowProp mi_Row
                End If

            End With
            b = False
        End If
    End If

End Sub

Private Sub Form_Activate()
    Static b As Boolean
    
    If Not b Then
        b = True
        SettaLargFlex Flex, 1, 0
    End If
End Sub
Private Sub Form_Load()

    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keySave).ToolTipText = MyLoadResString(k_Save)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_RES_FitGrid)
    tBar.Buttons(keyNew).ToolTipText = MyLoadResString(k_PROP_Property_New)
    tBar.Buttons(keyDelete).ToolTipText = MyLoadResString(k_PROP_Property_Delete)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource)


    Frame.Caption = MyLoadResString(k_PROP_Selected_Property)
    Label1(0).Caption = MyLoadResString(k_Name)
    Label1(1).Caption = MyLoadResString(k_PROP_Property_Value)
    
    m_lHeigth = Me.Height
    m_lWidth = Me.Width
    rtfProp.MaxLength = 7500 \ LenB("A")

    CenterForm Me
End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeigth Then Me.Height = m_lHeigth
            
            With Flex
                .Move .Left, .Top, Me.ScaleWidth - (2 * .Left), (Me.ScaleHeight / 2) - tBar.Height - k_FlexBevel
            End With
            With Frame
                .Move .Left, Flex.Top + Flex.Height + 20, Me.ScaleWidth - (2 * .Left), Me.ScaleHeight - (Flex.Top + Flex.Height + k_FlexBevel) - StBar.Height
                TxtProp.Width = .Width - TxtProp.Left - rtfProp.Left
                rtfProp.Move rtfProp.Left, rtfProp.Top, .Width - (2 * rtfProp.Left), .Height - rtfProp.Top - k_FlexBevel
                Label1(1).Move rtfProp.Left, Label1(1).Top, rtfProp.Width
            End With
            
            b = False
        End If
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fSqlProp = Nothing
End Sub

Private Sub rtfProp_Change()
    m_bObjChanged(1) = True
End Sub
Private Sub rtfProp_GotFocus()
    StBar.Panels(1).Text = MyLoadResString(k_PROP_Form_Property_Value)
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Enabled Then
        Select Case Button.Key
            Case keySave
                SaveProperty
                'Test4Save False
            Case keyEnd
                Unload Me
            Case keyFitGrid
                SettaLargFlex Flex, 1, 0
            Case keyRefresh
                Me.Stuff m_iobjType, m_sDbName, m_sOwner, m_s1LevelObj, m_s2LevelObj
            Case keyNew
                NewProperty
            Case keyDelete
                DeleteProp
        End Select
    End If

End Sub
Private Sub LoadProperty()

    Dim oVar As Variant
    Dim sBuf As String
    Dim i As Integer
    Dim lErr As Long, sErr As String
    Dim sSql As String
    Dim oDb As SQLDMO.Database2
    Dim oQry As SQLDMO.QueryResults2
        
    sSql = "SELECT [name] , SQL_VARIANT_PROPERTY([value],'basetype') as [VariantType] FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
        
    Select Case m_iobjType
        Case am_DbProp
            'sSql = "SELECT [name] FROM ::fn_listextendedproperty(NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL"))
        Case am_TbProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
        Case am_ColProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))

        Case am_IdxProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        Case am_KeyProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'constraint'", SqlFinalS(ApiciSQL(m_s2LevelObj))))

        Case am_TrigPropTB, am_TrigPropVIEW
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), IIf(m_iobjType = am_TrigPropTB, "'table'", "'view'"), SqlFinalS(ApiciSQL(m_s1LevelObj)), "'trigger'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        
        Case am_ViewProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
        Case am_SpProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
        Case am_FuncProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
        Case am_FuncParam
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        Case am_ProcParam
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        Case am_ViewColProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        Case am_ViewIdxProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        Case am_UdtProp
            sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'type'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL", "NULL", "NULL"))
    End Select
    
    Debug.Print sSql
    m_Lerr = 0
    m_sErr = ""
    On Local Error Resume Next
    Set oDb = objServer.Databases(m_sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 And Not oDb Is Nothing Then
        On Local Error Resume Next
        Set oQry = oDb.ExecuteWithResults(sSql, Len(sSql))
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If lErr = 0 And Not oQry Is Nothing Then
            
            oVar = Qry2Var(oQry)
            Set oQry = Nothing
            Frame.Enabled = False
            TxtProp.Text = ""
            rtfProp.Text = ""
            Erase m_bObjChanged
            
            With Flex
                m_bLoading = True
                .Redraw = False
                .FixedRows = 0
                .FixedCols = 0
                .Cols = 3
                .Rows = 1
                .Clear
                .TextMatrix(0, 1) = MyLoadResString(k_Property)
                .TextMatrix(0, 2) = MyLoadResString(k_Data_Type)
                .ColWidth(0) = k_FlexPARAM_cellaSelector
                
                .FixedAlignment(0) = flexAlignCenterCenter
            
                If IsArray(oVar) Then
                    For i = 0 To UBound(oVar, 2)
                        sBuf = oVar(0, i) & vbTab & oVar(0, i)
                        .AddItem vbTab & sBuf
                        .Row = .Rows - 1
                        .Col = .Cols - 2
                        .CellAlignment = flexAlignLeftCenter
                    Next
                End If
                .Row = .Rows - 1
                .AllowUserResizing = flexResizeColumns
                .AllowBigSelection = False
                .ScrollBars = flexScrollBarBoth

                If .Rows > 1 Then
                    .FixedRows = 1
                Else
                    For i = 1 To .Cols - 1
                        .Col = i
                        .CellBackColor = .BackColorFixed
                    Next
                End If
                .FixedCols = 1
                .Redraw = True
                SettaLargFlex Flex, 1, 0
                m_bLoading = False
        
                mi_Row = 0
                .TopRow = IIf(.Rows > 1, 1, 0)
                .Row = IIf(.Rows > 1, 1, 0)
                .Col = 1
                Flex_RowColChange
            End With
        End If
    End If
    
    Set oDb = Nothing
    If lErr <> 0 Then
        m_Lerr = lErr
        m_sErr = sErr
    End If
    
End Sub
Private Sub LoadRowProp(ByVal iRow As Integer)

    Dim oVar As Variant
    Dim lErr As Long, sErr As String
    Dim sSql As String
    Dim oDb As SQLDMO.Database2
    Dim oQry As SQLDMO.QueryResults2
    Dim bFound As Boolean
    Dim iStat As am_PropStatus
    'Dim iAsciz0 As Integer
    'Dim sValue As String
        
    If iRow <> 0 Then
        sSql = "SELECT [name], convert(VARCHAR(8000),[value]) AS vCharValue FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
        'sSql = "SELECT [name], [value] FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
        
        Select Case m_iobjType
            Case am_DbProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "NULL", "NULL", "NULL", "NULL", "NULL", "NULL"))
            Case am_TbProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
            Case am_ColProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                
            Case am_SpProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
            Case am_ViewProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
            Case am_UdtProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'type'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL", "NULL", "NULL"))
            Case am_FuncProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
            Case am_FuncParam
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
            Case am_ProcParam
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
            Case am_ViewColProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                                
            Case am_TrigPropTB, am_TrigPropVIEW
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), IIf(m_iobjType = am_TrigPropTB, "'table'", "'view'"), SqlFinalS(ApiciSQL(m_s1LevelObj)), "'trigger'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                                
                                
            Case am_ViewIdxProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                                
            Case am_IdxProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
            Case am_KeyProp
                sSql = ReplaceMsg(sSql, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array(SqlFinalS(ApiciSQL(Flex.TextMatrix(iRow, 1))), "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'constraint'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        End Select
        Debug.Print sSql
        
        Screen.MousePointer = vbHourglass
        m_Lerr = 0
        m_sErr = ""
        
        On Local Error Resume Next
        Set oDb = objServer.Databases(m_sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        
        If lErr = 0 And Not oDb Is Nothing Then
            On Local Error Resume Next
            Set oQry = oDb.ExecuteWithResults(sSql, Len(sSql))
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            If lErr = 0 And Not oQry Is Nothing Then
                
                oVar = Qry2Var(oQry)
                Set oQry = Nothing
                If IsArray(oVar) Then
                    TxtProp.Text = oVar(0, 0)
                        
'                   direct cast to VARCHAR(8000)
'                    sValue = StrConv(oVar(1, 0), vbUnicode)
'                    iAsciz0 = InStr(sValue, vbNullChar)
'                    If iAsciz0 <> 0 Then sValue = Left(sValue, iAsciz0 - 1)
'                    rtfProp.Text = sValue
                    rtfProp.Text = oVar(1, 0)
                    Erase m_bObjChanged
    
                    bFound = True
                    iStat = am_PropEdit
                    
                End If
            End If
        End If
    End If
    
    Frame.Enabled = bFound
    If Not bFound Then
        TxtProp.Text = ""
        rtfProp.Text = ""
    End If
    tBar.Buttons(keySave).Enabled = bFound
    tBar.Buttons(keyDelete).Enabled = bFound
    m_iStatus = iStat

    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub
Private Sub SetFormCaption()
    Dim sBuf As String
'Public Const k_PROP_Form_Extended_Property = 1851           ' Proprietà Estese oggetto # 1% # ( 2% )

    sBuf = MyLoadResString(k_PROP_Form_Extended_Property)
    Select Case m_iobjType
        Case am_DbProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName, MyLoadResString(k_RES_Database)))
        Case am_TbProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_sOwner & "." & m_s1LevelObj, MyLoadResString(k_Table)))
        Case am_ColProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Column)))
        Case am_IdxProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Indexes)))
        Case am_KeyProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Relations)))
            
        Case am_ViewProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_sOwner & "." & m_s1LevelObj, MyLoadResString(k_View)))
        Case am_SpProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_sOwner & "." & m_s1LevelObj, MyLoadResString(k_Stored_Procedure)))
        Case am_FuncProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_sOwner & "." & m_s1LevelObj, MyLoadResString(k_objFunction)))
        Case am_UdtProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj, MyLoadResString(k_User_Defined_Data_Type)))

        Case am_FuncParam, am_ProcParam
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Param_Name)))

        Case am_ViewColProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Column)))

        Case am_ViewIdxProp
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, MyLoadResString(k_Indexes)))
            
        Case am_TrigPropTB, am_TrigPropVIEW
            sBuf = ReplaceMsg(sBuf, Array("1%", "2%"), Array(m_sDbName & "." & m_s1LevelObj & "." & m_s2LevelObj, "Trigger"))

    End Select
    Me.Caption = sBuf

End Sub
Private Sub SaveProperty()

    Dim lErr As Long, sErr As String
    Dim sOldName As String
    Dim sCommand As String
    Dim sSqlAdd As String
    Dim sSqlUpdate As String
    Dim bDropAdd As Boolean
    Dim sBuf As String
    Dim oDb As SQLDMO.Database2
    Dim sPropertyName As String
    Dim sPropWrite As String
    Dim iRow As Integer
    
    sSqlUpdate = "EXEC sp_updateextendedproperty 1%, 2%, 3%, 4%, 5%, 6%, 7%, 8%"
    sSqlAdd = "EXEC sp_addextendedproperty 1% , 2% , 3% , 4% , 5% , 6% , 7% , 8%"

    
    Screen.MousePointer = vbHourglass

    If m_bObjChanged(0) Or m_bObjChanged(1) And Len(Trim$(TxtProp.Text)) <> 0 Then
    
        On Local Error Resume Next
        Set oDb = objServer.Databases(m_sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If lErr = 0 Then
            If m_iStatus = am_PropEdit Then
                sOldName = Flex.TextMatrix(mi_Row, 1)
                bDropAdd = m_bObjChanged(0) = True
                
                If bDropAdd Then
                    sCommand = sSqlAdd
                Else
                    sCommand = sSqlUpdate
                End If
            Else
                sCommand = sSqlAdd
            End If
            
        
            If bDropAdd Then
                Select Case m_iobjType
                    Case am_DbProp
                        sBuf = SqlFinalS(ApiciSQL(sOldName))
                        Case am_TbProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'table' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
                        Case am_ColProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'column', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                        
                        Case am_SpProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'procedure' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
                        Case am_ViewProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'view' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
                        Case am_FuncProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'function' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
                        Case am_UdtProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", @level0type ='type' ,@level0name =" & SqlFinalS(ApiciSQL(m_s1LevelObj))
                        
                        Case am_FuncParam
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'function', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'parameter', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                        Case am_ProcParam
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'procedure', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'parameter', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                            
                        Case am_ViewColProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'view', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'column', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                        Case am_ViewIdxProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'view', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'index', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                    
                        Case am_IdxProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'index', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                        Case am_KeyProp
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'constraint', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                        Case am_TrigPropTB, am_TrigPropVIEW
                            sBuf = SqlFinalS(ApiciSQL(sOldName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", " & IIf(m_iobjType = am_TrigPropTB, "'table'", "'view'") & ", " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'trigger', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                    
                End Select
                Debug.Print sBuf
                sBuf = "EXEC sp_dropextendedproperty " & sBuf
                On Local Error Resume Next
                oDb.ExecuteImmediate sBuf, SQLDMOExec_Default, Len(sBuf)
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
            End If
        End If
        
        If lErr = 0 Then
            sPropertyName = Trim$(TxtProp.Text)
            sPropWrite = SqlFinalS(ApiciSQL(Trim$(rtfProp.Text)))
            'If Len(sPropWrite) = 0 Then sPropWrite = "''"
            If Len(sPropWrite) = 0 Then sPropWrite = "NULL"
            Select Case m_iobjType
                Case am_DbProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "NULL", "NULL", "NULL", "NULL", "NULL", "NULL"))
                Case am_TbProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
                Case am_ColProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
        
                Case am_IdxProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                Case am_KeyProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'table'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'constraint'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                
                        
                Case am_ViewProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
                Case am_SpProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
                Case am_FuncProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL"))
                Case am_UdtProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'type'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "NULL", "NULL", "NULL", "NULL"))
                
                Case am_FuncParam
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'function'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                Case am_ProcParam
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'procedure'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'parameter'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                Case am_ViewColProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'column'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                Case am_ViewIdxProp
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), "'view'", SqlFinalS(ApiciSQL(m_s1LevelObj)), "'index'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
                    
                Case am_TrigPropTB, am_TrigPropVIEW
                    sCommand = ReplaceMsg(sCommand, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%"), Array(SqlFinalS(ApiciSQL(sPropertyName)), sPropWrite, "'user'", SqlFinalS(ApiciSQL(m_sOwner)), IIf(m_iobjType = am_TrigPropTB, "'table'", "'view'"), SqlFinalS(ApiciSQL(m_s1LevelObj)), "'trigger'", SqlFinalS(ApiciSQL(m_s2LevelObj))))
            End Select
                
            Debug.Print sCommand
            On Local Error Resume Next
            oDb.ExecuteImmediate sCommand, SQLDMOExec_Default, Len(sCommand)
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            If lErr = 0 Then
                If m_iStatus = am_PropEdit Then
                    iRow = Flex.Row
                    Flex.TextMatrix(iRow, 1) = sPropertyName
                Else
                    Flex.AddItem vbTab & sPropertyName
                    Flex.Row = Flex.Rows - 1
                    Flex.CellAlignment = flexAlignLeftCenter
                    If Flex.Rows = 2 Then
                        Flex.Row = Flex.Rows - 2
                        Flex.CellBackColor = Flex.BackColorFixed
                        Flex.FixedRows = 1
                        mi_Row = 0
                        Flex_RowColChange
                    End If
                End If
            End If

        End If
    End If
    
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
End Sub
Private Sub NewProperty()
    
    m_iStatus = am_PropNew
    TxtProp.Text = ""
    rtfProp.Text = ""
    Erase m_bObjChanged
    Frame.Enabled = True
    TxtProp.SetFocus
    tBar.Buttons(keySave).Enabled = True
    
End Sub

Private Sub TxtProp_Change()
    m_bObjChanged(0) = True
End Sub
Private Sub TxtProp_GotFocus()
    StBar.Panels(1).Text = MyLoadResString(k_PROP_Form_Property_Name)
End Sub

Private Sub DeleteProp()

    Dim lErr As Long, sErr As String
    Dim sCommand As String
    Dim sPropName As String
    Dim iRow As Integer
    Dim i As Integer
    
    Dim oDb As SQLDMO.Database2
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    Set oDb = objServer.Databases(m_sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        sPropName = Flex.TextMatrix(Flex.Row, 1)
        Select Case m_iobjType
            Case am_DbProp
                sCommand = SqlFinalS(ApiciSQL(sPropName))
            Case am_TbProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'table' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
            Case am_ColProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'column', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
                
            Case am_ViewProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'view' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
            Case am_SpProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'procedure' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
            Case am_FuncProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user' ," & SqlFinalS(ApiciSQL(m_sOwner)) & " , 'function' , " & SqlFinalS(ApiciSQL(m_s1LevelObj))
            Case am_UdtProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", @level0type ='type' ,@level0name =" & SqlFinalS(ApiciSQL(m_s1LevelObj))
                
                
            Case am_FuncParam
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'function', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'parameter', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
            Case am_ProcParam
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'procedure', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'parameter', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
            Case am_ViewColProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'view', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'column', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
            Case am_ViewIdxProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'view', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'index', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
        
            Case am_IdxProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'index', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
            Case am_KeyProp
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", 'table', " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'constraint', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
            
            Case am_TrigPropTB, am_TrigPropVIEW
                sCommand = SqlFinalS(ApiciSQL(sPropName)) & ", 'user', " & SqlFinalS(ApiciSQL(m_sOwner)) & ", " & IIf(m_iobjType = am_TrigPropTB, "'table'", "'view'") & ", " & SqlFinalS(ApiciSQL(m_s1LevelObj)) & ", 'trigger', " & SqlFinalS(ApiciSQL(m_s2LevelObj))
        
        End Select
        sCommand = "EXEC sp_dropextendedproperty " & sCommand
        On Local Error Resume Next
        oDb.ExecuteImmediate sCommand, SQLDMOExec_Default, Len(sCommand)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If lErr = 0 Then
            iRow = Flex.Row
            If Flex.Rows = 2 Then
                m_bLoading = True
                Flex.FixedRows = 0
            End If
            Flex.RemoveItem iRow
            
            If Flex.Rows = 1 Then
                For i = 1 To Flex.Cols - 1
                    Flex.Col = i
                    Flex.CellBackColor = Flex.BackColorFixed
                Next
                m_bLoading = False
                Flex.Col = 1
            End If
            If iRow >= Flex.Rows - 1 Then
                iRow = Flex.Rows - 1
            End If
            mi_Row = 0
            Flex.Row = iRow
            Flex_RowColChange
        End If
    End If
    
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
End Sub

