VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fQuery 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Query"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   60
      Left            =   0
      ScaleHeight     =   26.127
      ScaleMode       =   0  'User
      ScaleWidth      =   26988
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   2595
   End
   Begin MSComctlLib.ImageList imList 
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":00EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":0408
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":085C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":0CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":0E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fQuery.frx":0F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7435
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Results"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Messages"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfQry 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"fQuery.frx":10C2
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin RichTextLib.RichTextBox tRes 
      Height          =   3375
      Index           =   0
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fQuery.frx":113C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox tRes 
      Height          =   3375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fQuery.frx":11B5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
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
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " "
            Description     =   "Run"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Del"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Open"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WWrap"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Default"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   3615
         TabIndex        =   6
         Top             =   0
         Width           =   3615
         Begin VB.ComboBox cbo 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label lbl 
            Caption         =   "DB:"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin VB.Image imgSplitter 
      Height          =   105
      Left            =   180
      MousePointer    =   7  'Size N S
      Top             =   2160
      Width           =   4095
   End
End
Attribute VB_Name = "fQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum amSaveMode
    enSql = 0
    enRPT
    enTAB
End Enum
Private mbFlex As Boolean
Private mbMoving As Boolean
Private m_iSaveMode As amSaveMode

Private m_lWidth As Long
Private m_lHeight As Long

Private m_bExecuting As Boolean
Private m_sGOstm As String
Private m_oFont(2) As StdFont

Const sglSplitLimit = 500

Private Sub Flex_GotFocus()
    m_iSaveMode = enTAB
    ChangeTBarSaveMode
End Sub

Private Sub Form_Activate()

    Static b As Boolean
    If Not b Then
        Set tStrip.SelectedItem = tStrip.Tabs(1)
        b = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyF5        'Run
            If Shift = 0 Then
                tBar_ButtonClick tBar.Buttons(1)
                Exit Sub
            End If
    End Select

    If tStrip.Visible Then
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
    End If

End Sub
Private Sub Form_Load()
    
    m_lWidth = Me.Width
    m_lHeight = Me.Height

    Set m_oFont(0) = New StdFont
    Set m_oFont(1) = New StdFont
    Set m_oFont(2) = New StdFont
    m_oFont(0) = rtfQry.Font
    m_oFont(1) = Flex.Font
    m_oFont(2) = tRes(0).Font

    GetOptions

    tBar.Buttons(1).ToolTipText = MyLoadResString(k_Execute_Query_F5)
    tBar.Buttons(2).ToolTipText = MyLoadResString(k_Qry_Delete)
    tBar.Buttons(4).ToolTipText = MyLoadResString(k_Qry_OpenQry)
    tBar.Buttons(5).ToolTipText = MyLoadResString(k_ShowGrid_Text)
    tBar.Buttons(keyWWrap).ToolTipText = MyLoadResString(k_WordWrap)
    tBar.Buttons(7).ToolTipText = MyLoadResString(k_Settings)
    tBar.Buttons(7).ButtonMenus(keyDefault).Text = MyLoadResString(k_Qry_Opt_Reset)
    
    
    picSplitter.BackColor = vb3DFace
    CenterForm Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = m_bExecuting
End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            SizeControls imgSplitter.Top
            b = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fQuery = Nothing
End Sub
Private Sub SizeControls(y As Single)

    'set the width
    tBar.Refresh

    If y < 1500 Then y = 1500
    If y > (Me.Height - 1500) Then y = Me.Height - 1500
    
    imgSplitter.Move 0, y, Me.ScaleWidth
    picSplitter.Move imgSplitter.Left, imgSplitter.Top, imgSplitter.Width, imgSplitter.Height
    picSplitter.Width = Me.ScaleWidth
    
    Picture1.Width = Me.ScaleWidth - Picture1.Left - k_MargineFRM
    cbo.Width = Picture1.ScaleWidth - cbo.Left - k_FlexBevel

        
    rtfQry.Move 20, tBar.Height + 20, Me.ScaleWidth - 20
        
    rtfQry.Height = y - rtfQry.Top
                
    tStrip.Move 0, y + picSplitter.Height + 20, Me.ScaleWidth, Me.ScaleHeight - y - k_MargineFRM
    tRes(1).Move 20, tStrip.Top + 20, tStrip.Width - 40, tStrip.Height - 350
    Flex.Move 20, tStrip.Top + 20, tStrip.Width - 40, tStrip.Height - 350
    tRes(0).Move 20, tStrip.Top + 20, tStrip.Width - 40, tStrip.Height - 350
    'Flex.Visible = mbFlex
    'tRes(0).Visible = not mbFlex
    rtfQry.RightMargin = TextWidth("W") * 200
        
End Sub
'Private Sub LoadGrid(oQry As SQLDMO.QueryResults, ByVal bFirst As Boolean, ByRef lW2 As Long)
Private Sub LoadGrid(oQry As SQLDMO.QueryResults2, ByVal iRetSet As Integer, ByRef lW2 As Long)
    
    Dim iCols As Integer, lRows As Long
    Dim i As Integer, y As Integer, l As Long
    Dim sBuf As String
    ReDim iBool(0) As Boolean, fBool(0) As Boolean
    Dim x As Integer
    Dim sColName As String
    Dim lErr As Long, sErr As String
    Dim bFirst As Boolean
    Dim lW As Long
    
    Dim lChunk As Long, lRest As Long
    
    Dim sHeader As String
    
    ' chunk of data to load each time
    Const kChunkSize = 1000
    
    
    bFirst = iRetSet = 1
    
    Debug.Print oQry Is Nothing
    If Not oQry Is Nothing Then
        If oQry.ResultSets <> 0 Then
            Flex.Redraw = False
            If bFirst Then Flex.ColWidth(0) = 650
            
            For x = 1 To oQry.ResultSets
                oQry.CurrentResultSet = x
                iCols = oQry.Columns + 1
                lRows = oQry.Rows
                
                On Local Error Resume Next
                If mbFlex Then
                    If Flex.Cols < iCols Then Flex.Cols = iCols
                End If
                                    
                If x > 0 Or Not bFirst Then
                    If mbFlex Then
                        'Flex.AddItem vbTab & "Resultset " & vbTab & x
                        Flex.AddItem vbTab & "Resultset " & x
                    Else
                        'tRes(0).Text = tRes(0).Text & "Resultset " & vbTab & x & vbCrLf
                        tRes(0).Text = tRes(0).Text & "Resultset " & x & vbCrLf
                    End If
                End If
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                If lErr <> 0 And Len(sErr) <> 0 Then
                    Beep
                    tRes(1).Text = tRes(1).Text & vbCrLf & String$(15, "-") & vbCrLf & ReplaceMsg(MyLoadResString(kErrorBatch), Array("1%", "2%", "3%"), Array(lErr, sErr, iRetSet))
                End If
                
                If mbFlex Then
                    With Flex
                        For i = 1 To iCols - 1
                            sColName = oQry.ColumnName(i)
                            If Len(sColName) = 0 Then sColName = MyLoadResString(k_Column) & i
                            If x = 1 And bFirst Then
                                .TextMatrix(0, i) = sColName
                            Else
                                sBuf = sBuf & vbTab & sColName
                            End If
                        Next
                        If x > 1 Then
                            .AddItem sBuf
                            sBuf = ""
                        End If
                    
                        On Local Error Resume Next
                        For l = 1 To lRows
                             For i = 1 To iCols - 1
                                If i > 1 Then sBuf = sBuf & vbTab
                                Select Case oQry.ColumnType(i)
                                    Case SQLDMO_DTypeUnknown
                                        sBuf = sBuf & MyLoadResString(k_Unknown)
                                    Case SQLDMO_DTypeBinary
                                        sBuf = sBuf & "..Binary"
                                    Case SQLDMO_DTypeImage
                                        sBuf = sBuf & "..BLOB"
                                    Case SQLDMO_DTypeText, SQLDMO_DTypeNText
                                        sBuf = sBuf & "..MEMO"
                                    Case SQLDMO_DTypeGUID
                                        sBuf = sBuf & "{..Guid..}"
                                    Case SQLDMO_DTypeSQLVariant
                                        sBuf = sBuf & "..Sql_Variant.."
                                    Case SQLDMO_DTypeUChar, SQLDMO_DTypeChar, SQLDMO_DTypeUVarchar, SQLDMO_DTypeVarchar
                                        sBuf = sBuf & oQry.GetColumnString(l, i)
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignLeftCenter
                                    Case SQLDMO_DTypeMoney
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.0000")
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignRightCenter
                                    Case SQLDMO_DTypeFloat4, SQLDMO_DTypeMoney
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.0000")
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignRightCenter
                                    Case SQLDMO_DTypeFloat8
                                        sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i), "#,0.00000000")
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignRightCenter
                                    Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
                                        sBuf = sBuf & Format$(oQry.GetColumnDate(l, i), kYEAR)
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignLeftCenter
                                    
                                    Case SQLDMO_DTypeBigint
                                        sBuf = sBuf & oQry.GetColumnBigInt(l, i)
                                    
                                    'Aggiunto Case 2 perche' altrimenti non carica colonne Identity
                                    'anche se NON e' un valore listato tra le costanti
                                    Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                                        sBuf = sBuf & Format$(oQry.GetColumnLong(l, i), "#,0")
                                        If x = 1 And bFirst Then .ColAlignment(i) = flexAlignRightCenter
                                    Case SQLDMO_DTypeBit
                                        sBuf = sBuf & Format$(oQry.GetColumnBool(l, i))
                                End Select
                            Next
                            
                            .AddItem l & vbTab & sBuf
                            lErr = Err.Number
                            sErr = Err.Description
                            
                            sBuf = ""
                            If lErr <> 0 Then Exit For
                        Next
                        On Local Error GoTo 0
                        
                        If lErr <> 0 And Len(sErr) <> 0 Then
                            Beep
                            tRes(1).Text = tRes(1).Text & vbCrLf & String$(15, "-") & vbCrLf & ReplaceMsg(MyLoadResString(kErrorBatch), Array("1%", "2%", "3%"), Array(lErr, sErr, iRetSet))
                        End If
                    End With
                    
                Else
                    sBuf = ""
                    iCols = oQry.Columns
                    lRows = oQry.Rows
                    sHeader = oQry.GetRangeString(0, , 1, , vbCrLf, " ")
                    
                    sBuf = sHeader
                    i = 0
                    i = InStr(sBuf, vbCrLf)
                    If i <> 0 Then y = InStr(i + 1, sBuf, vbCrLf)
                    If y <> 0 Then i = y
                    If i <> 0 Then
                        sHeader = Left$(sBuf, i) & vbCrLf
                    Else
                        sHeader = sBuf
                    End If
                    sBuf = sHeader
                                        
                    lChunk = lRows \ kChunkSize
                    lRest = lRows Mod kChunkSize
                    Debug.Print sBuf

                    On Local Error Resume Next
                    For l = 0 To lChunk - 1
                        On Local Error Resume Next
                        sBuf = sBuf & oQry.GetRangeString((l * kChunkSize) + 1, , (l * kChunkSize) + 1 + kChunkSize, , vbCrLf, " ") & vbCrLf
                        lErr = Err.Number
                        sErr = Err.Description
                        If lErr <> 0 Then Exit For
                    Next
                    On Local Error GoTo 0
                    If lErr = 0 Then
                        If l > 0 Then l = l - 1
                        If lChunk <> 0 Then lChunk = kChunkSize
                        If lRest <> 0 Then sBuf = sBuf & oQry.GetRangeString((l * kChunkSize) + 1 + lChunk, , , , vbCrLf, " ") & vbCrLf
                    End If
                    
                    If lErr <> 0 And Len(sErr) <> 0 Then
                        Beep
                        tRes(1).Text = tRes(1).Text & vbCrLf & String$(15, "-") & vbCrLf & ReplaceMsg(MyLoadResString(kErrorBatch), Array("1%", "2%", "3%"), Array(lErr, sErr, iRetSet))
                    End If
                                        
                    On Local Error Resume Next
                    tRes(0).Text = tRes(0).Text & sBuf
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    
                    If lErr <> 0 And Len(sErr) <> 0 Then
                        Beep
                        tRes(1).Text = tRes(1).Text & vbCrLf & String$(15, "-") & vbCrLf & ReplaceMsg(MyLoadResString(kErrorBatch), Array("1%", "2%", "3%"), Array(lErr, sErr, iRetSet))
                    End If
                    
                    lW = InStr(sHeader, vbCrLf)
                    If lW < (Len("Resultset " & x)) Then lW = Len("Resultset " & x)
                    If lW > lW2 Then lW2 = lW
                End If
            Next
            With Flex
                If .Cols = 1 Then .Cols = 2
                If .Rows = 1 Then .AddItem ""
                .Row = 1
                .ScrollBars = flexScrollBarBoth
                .FixedRows = 1
                .FixedCols = 1
                
                .Redraw = True
                .TopRow = 1
            End With
            If Not mbFlex Then tRes(0).RightMargin = Me.TextWidth("w") * lW2

        End If
    End If
    
End Sub
Private Sub OpenSql()

    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim obj As cOpen
    Set obj = New cOpen
    
    obj.Title = MyLoadResString(k_Select_File_Sql)
    obj.Def_Dir() = App.Path
    obj.Filter() = "Query File (*.Sql)" & Chr$(0) & "*.sql" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
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
    
    If lErr = 0 Then
        If Len(sFile) Then
            On Local Error Resume Next
            rtfQry.LoadFile sFile, rtfText
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
        End If
    End If
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If

End Sub
Public Sub LoadDB()

    Dim i As Integer
      
    Screen.MousePointer = vbHourglass
    cbo.Clear
    For i = 1 To objServer.Databases.Count
        If Len(objServer.Databases(i).Name) <> 0 Then
            cbo.AddItem objServer.Databases(i).Name
        End If
    Next
    cbo.ListIndex = GetItem("master", cbo)
    Screen.MousePointer = vbDefault

End Sub

Private Sub rtfQry_GotFocus()
    m_iSaveMode = enSql
    ChangeTBarSaveMode
End Sub

Private Sub rtfQry_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer
    
    If Button = vbRightButton Then
        Dim M As New cPUpMenu
        M.Add MyLoadResString(k_MnuCut), 1
        M.Add MyLoadResString(k_MnuCopy), 2
        M.Add MyLoadResString(k_MnuPaste), 3, , bDisable:=Len(Clipboard.GetText) = 0
        M.Add "", 4, bSeparator:=True
        M.Add MyLoadResString(k_MnuSelectAll), 5
        
        Call GetCursorPos(P)
        i = M.Popup(Me.hWnd, P.x, P.y)
        M.Clear
        Set M = Nothing
        
        Select Case i
            Case 1
                Clipboard.SetText rtfQry.SelRTF, vbCFText
                rtfQry.SelRTF = ""
            Case 2
                Clipboard.SetText rtfQry.SelRTF, vbCFText
            Case 3
                rtfQry.SelRTF = Clipboard.GetText(vbCFText)
            Case 5
                rtfQry.SelStart = 0
                rtfQry.SelLength = Len(rtfQry.Text)
            
        End Select
    End If
       
End Sub

Private Sub rtfQry_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim lErr As Long, sErr As String
    If Data.GetFormat(vbCFFiles) Then
        rtfQry.Text = ""
        On Local Error Resume Next
        rtfQry.LoadFile Data.Files(1), rtfText
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If lErr <> 0 Then
            Beep
            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
        End If
    End If

End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim bSave As Boolean
    Select Case Button.Index
        Case 1  'Go
            rtfQry.SetFocus
            If Len(rtfQry.Text) <> 0 Then Execute
            
        Case 2  'Delete
            If MsgBox(MyLoadResString(k_Confirm_Remove), vbQuestion Or vbOKCancel, App.EXEName) = vbOK Then
                rtfQry.ZOrder
                rtfQry.Text = ""
                tRes(0).Text = ""
                tRes(1).Text = ""
                With Flex
                    .Redraw = False
                    .Clear
                    .FixedCols = 0
                    .FixedRows = 0
                    .Cols = 1
                    .Rows = 1
                    .Redraw = True
                End With
                'tStrip.Visible = True
                SizeControls imgSplitter.Top
            End If
            rtfQry.SetFocus

        Case 3  'Save
            If m_iSaveMode = enSql Then
                rtfQry.SetFocus
                bSave = Len(rtfQry.Text)
            ElseIf m_iSaveMode = enRPT Then
                Set tStrip.SelectedItem = tStrip.Tabs(1)
                tRes(0).SetFocus
                bSave = Len(tRes(0).Text)
            ElseIf m_iSaveMode = enTAB Then
                Set tStrip.SelectedItem = tStrip.Tabs(1)
                Flex.SetFocus
                bSave = Flex.Rows > 1
            End If
                
            If bSave Then SaveRequest
            
        Case 4  'Open
            rtfQry.SetFocus
            OpenSql
         Case 5
            'mbFlex = Not mbFlex
            'tBar.Buttons(Button.index).Value = Abs(mbFlex)
            mbFlex = tBar.Buttons(Button.Index).Value
        Case 6
            rtfQry.RightMargin = IIf(Button.Value, 40000, 0)
            
        Case 7
            LoadOptions
    End Select
        
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Key = keyDefault Then
        ResetQryDefault
    End If
End Sub

Private Sub tRes_GotFocus(Index As Integer)
    If Index = 0 Then
        m_iSaveMode = enRPT
        ChangeTBarSaveMode
    End If
End Sub

Private Sub tStrip_Click()

    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    If mbFlex Then
        Flex.Visible = i = 0
    Else
        tRes(0).Visible = i = 0
    End If
    tRes(1).Visible = i <> 0 'Not flex.Visible
    If i = 0 Then
        If mbFlex Then
            Flex.ZOrder
            'Flex.SetFocus
        Else
            tRes(0).ZOrder
            'tRes(0).SetFocus
        End If
    Else
        tRes(1).ZOrder
        'tRes(1).SetFocus
    End If

End Sub
Private Sub SaveRequest()

    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim obj As cOpen
    Set obj = New cOpen
    Dim iMsg As Integer
    Dim bErr As Boolean
    Dim iFile As Integer
    Dim sOutPut As String
    Dim iCol As Integer, iRow As Integer
    Dim iFixexCol As Integer, iFixedRow As Integer
    
    
    Dim sTitle As String
    Dim sFilter As String
    Dim sDefDir As String
    Dim sDefFileName As String
    
    sDefDir = App.Path
    If m_iSaveMode = enSql Then
        sTitle = MyLoadResString(k_Save_Query)
        'sDefDir = App.Path
        sFilter = "Query File (*.Sql)" & Chr$(0) & "*.sql" & Chr$(0) & "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        sDefFileName = MyLoadResString(k_New) & ".sql"
        iMsg = k_Save_Query
    Else
        sTitle = MyLoadResString(k_Save_Query)
        'sDefDir = App.Path
        iMsg = k_QRY_Save_Done
        If m_iSaveMode = enRPT Then
            sFilter = "Report File (*.Rpt)" & Chr$(0) & "*.Rpt" & Chr$(0) & "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
            sDefFileName = "Export.Rpt"
        Else
            sFilter = "Tab Delimited Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
            sDefFileName = "Export.txt"
        End If
    End If

    obj.Title = sTitle
    obj.Def_Dir() = sDefDir
    obj.Filter() = sFilter
    obj.f_Name = sDefFileName
    obj.formHwnd() = Me.hWnd

    obj.Do_Save
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
    
    If lErr = 0 Then
        If Len(sFile) Then
            On Local Error Resume Next
            If m_iSaveMode = enSql Then
                rtfQry.SaveFile sFile, rtfText
            ElseIf m_iSaveMode = enRPT Then
                tRes(0).SaveFile sFile, rtfText
            ElseIf m_iSaveMode = enTAB Then
                
                With Flex
                    .Redraw = False
                    iCol = .Col
                    iRow = .Row
                    iFixexCol = .FixedCols
                    iFixedRow = .FixedRows
                    .Col = 0
                    .Row = 0
                    .ColSel = .Cols - 1
                    .RowSel = .Rows - 1
                    sOutPut = .Clip
                    .FixedCols = iFixexCol
                    .FixedCols = .FixedRows
                    .Col = iCol
                    .Row = iRow
                    .Redraw = True
                End With
                iFile = FreeFile
                Open sFile For Output As #iFile
                If Err.Number = 0 Then Print #iFile, sOutPut
                Close #iFile
                sOutPut = ""
            End If
            bErr = Err.Number <> 0
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
                
            If bErr = False Then
                MsgBox MyLoadResString(k_Executed) & " " & MyLoadResString(iMsg) & " " & sFile, vbInformation Or vbOKOnly, App.EXEName
            End If
        End If
    End If
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If

End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    With imgSplitter
        picSplitter.ZOrder
        picSplitter.Move .Left, .Top, .Width, .Height
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = y + imgSplitter.Top
        'picSplitter.ZOrder
        If sglPos < sglSplitLimit Then
            picSplitter.Top = sglSplitLimit
        ElseIf sglPos > Me.Height - sglSplitLimit Then
            picSplitter.Top = Me.Height - sglSplitLimit
        Else
            picSplitter.Top = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'APP WIZARD
    SizeControls picSplitter.Top
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub IsGO_USE_Statement(ByVal sTmp As String, ByRef blnGO As Boolean, ByRef blnIsComment As Boolean, ByRef blnUse As Boolean, ByRef sDbName As String)

    Dim lngStartQuote As Long
    Dim lngEndQuote As Long
    Dim lngStartRem As Long
    Dim lRem(1) As Long
    Dim sBuf As String
    Dim lGOpos As Long
    Dim blnWasREM As Boolean
    
    'byref Variable Init
    blnWasREM = blnIsComment
    
    blnGO = False
    blnUse = False
    sDbName = ""

    lngStartQuote = 0
    lngEndQuote = 0
    lRem(0) = 0
    lRem(1) = 0
    'removes all 'char' included in ''
    Do
        lngStartQuote = InStr(sTmp, "'")
        lngEndQuote = InStr(lngStartQuote + 1, sTmp, "'")
        
        If lngStartQuote <> 0 And lngEndQuote <> 0 Then
            sTmp = Trim$(Left$(sTmp, lngStartQuote - 1) & IIf(Len(sTmp) > lngEndQuote + 1, Mid$(sTmp, lngEndQuote + 1), ""))
        ElseIf lngStartQuote <> 0 And lngEndQuote = 0 Then
            Mid$(sTmp, lngStartQuote, 1) = " "
        End If
    Loop While lngStartQuote <> 0
    
    lngStartQuote = 0
    lngEndQuote = 0
    lRem(0) = 0
    lRem(1) = 0
    'removes all "char" included in ""
    Do
        lngStartQuote = InStr(sTmp, """")
        lngEndQuote = InStr(lngStartQuote + 1, sTmp, """")
        
        If lngStartQuote <> 0 And lngEndQuote <> 0 Then
            sTmp = Trim$(Left$(sTmp, lngStartQuote - 1) & IIf(Len(sTmp) > lngEndQuote + 1, Mid$(sTmp, lngEndQuote + 1), ""))
        ElseIf lngStartQuote <> 0 And lngEndQuote = 0 Then
            Mid$(sTmp, lngStartQuote, 1) = " "
        End If
    Loop While lngStartQuote <> 0
    

    'search for REM
    lngStartQuote = 0
    lngEndQuote = 0
    lRem(0) = 0
    lRem(1) = 0
    lRem(0) = InStr(sTmp, "/*")
    lRem(1) = InStr(lRem(0) + 1, sTmp, "*/")
    If lRem(0) <> 0 And lRem(1) <> 0 Then
        sTmp = Trim$(Left$(sTmp, lRem(0) - 1) & IIf(Len(sTmp) > lRem(1) + 2, Mid$(sTmp, lRem(1) + 2), ""))
    ElseIf lRem(0) <> 0 Then
        If lRem(0) > 1 Then
            sTmp = Trim$(Left$(sTmp, lRem(0) - 2))
        Else
            sTmp = ""
        End If
        blnIsComment = True
    ElseIf lRem(1) <> 0 Then
        sTmp = Trim$(IIf(Len(sTmp) > lRem(1) + 2, Mid$(sTmp, lRem(1) + 2), ""))
        blnIsComment = False
        blnWasREM = False
    End If
                
    'removes after '--
    Do
        lngStartRem = InStr(sTmp, "--")
        If lngStartRem <> 0 Then sTmp = Trim$(Left$(sTmp, lngStartRem - 1))
    Loop While lngStartRem <> 0
        
    'search for GO
    lngStartQuote = 0
    lngEndQuote = 0
    lGOpos = InStr(1, sTmp, m_sGOstm, vbTextCompare)
    If lGOpos <> 0 Then
        blnGO = True
        If Len(sTmp) > Len(m_sGOstm) Then
            lngStartQuote = Len(Mid$(sTmp, lGOpos + Len(m_sGOstm)))
            If lngStartQuote <> 0 Then
                If Mid$(sTmp, lGOpos + Len(m_sGOstm), 1) <> " " Then blnGO = False
            End If
            lngEndQuote = Len(Left$(sTmp, lGOpos - 1))
            If lngEndQuote <> 0 Then
                If Mid$(sTmp, lGOpos - 1, 1) <> " " Then blnGO = False
            End If
        End If
    End If
    
    
    'search for USE database
    lngStartQuote = 0
    lngEndQuote = 0
    lRem(0) = 0
    lRem(1) = 0
    Do
        lngStartQuote = InStr(1, sTmp, "USE ", vbTextCompare)
            
        lRem(0) = lngStartQuote
        lRem(1) = lngStartQuote
        If lngStartQuote <> 0 Then
            If lngStartQuote > 1 Then
                'if preceding char <> " " then null
                If Mid$(sTmp, lngStartQuote - 1, 1) <> " " Then
                    lRem(1) = 0
                End If
            End If
        End If
        If lngStartQuote <> 0 Then
            sTmp = Trim$(Mid$(sTmp, lngStartQuote + 4))
            
            If lRem(1) <> 0 Then blnUse = Len(sTmp)
            
            If lRem(1) <> 0 Then
                sBuf = sTmp
                'init of DB Name
                lngEndQuote = InStr(sTmp, " ")
                If lngEndQuote <> 0 Then
                    sBuf = Left$(sTmp, lngEndQuote - 1)
                    sTmp = Mid$(sTmp, lngEndQuote + 1)
                    blnUse = True
                End If
                    
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "-")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "/")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "*")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "(")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, ")")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "*")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
                If Not blnUse Then
                    lngEndQuote = InStr(sTmp, "#")
                    If lngEndQuote <> 0 Then
                        sBuf = Left$(sTmp, lngEndQuote - 1)
                        sTmp = Mid$(sTmp, lngEndQuote + 1)
                        blnUse = True
                    End If
                End If
            End If
        End If
    Loop While lRem(0) <> 0
        
    
    If blnGO Then
        If blnWasREM Then blnGO = False
    End If
    If Not blnUse Then sTmp = ""
    If blnWasREM Then sTmp = ""
    If Len(sTmp) <> 0 Then
        sDbName = sTmp
        blnUse = True
    End If
    
End Sub

Public Sub Execute()
        
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim db As SQLDMO.Database2
    Dim lW As Long
    Dim strNewDb As String, sSavedDBname As String
    Dim sTmpFileIn As String, strFileOut As String
    Dim iChan As Integer, iOut As Integer
    Dim iRetSet As Integer
    Dim blnGO As Boolean, blnComment As Boolean, blnUse As Boolean
            
        
    Screen.MousePointer = vbHourglass
    m_bExecuting = True
    tBar.Enabled = False
    
    lErr = ErrWriteTmpQuery(sTmpFileIn, sErr)
    Set Me.Font = tRes(0).Font

    If lErr = 0 And Len(sTmpFileIn) Then
        iChan = FreeFile
        On Local Error Resume Next
        Open sTmpFileIn For Input As #iChan
        lErr = Err.Number
        sErr = "File: " & sTmpFileIn & vbLf & Err.Description
        On Local Error GoTo 0
        
        If lErr = 0 Then
            Call LockWindowUpdate(tRes(0).hWnd)
            
            lErr = ErrSetActiveDB(db, cbo.Text, sErr)
            If lErr = 0 Then
                iOut = OutputFileOpen(strFileOut)
                Do While Not EOF(iChan) And iOut <> 0 And lErr = 0
                    Line Input #iChan, sBuf
                    sBuf = Trim$(sBuf)
                    Debug.Print sBuf
                    
                    If Len(sBuf) <> 0 Then
                        IsGO_USE_Statement sBuf, blnGO, blnComment, blnUse, strNewDb
                        If blnUse And Len(strNewDb) <> 0 Then
                            If StrComp(db.Name, strNewDb, vbTextCompare) <> 0 Then sSavedDBname = strNewDb
                        End If
                        
                        If blnGO Then
                            Close #iOut
                            lErr = ErrExecCmd(db, strFileOut, sErr, iRetSet, lW)
                            Call DeleteFile(strFileOut)
                            
                            iOut = 0
                            If lErr = 0 Then iOut = OutputFileOpen(strFileOut)
                            If lErr = 0 Then
                                If Len(sSavedDBname) <> 0 Then
                                    lErr = ErrSetActiveDB(db, sSavedDBname, sErr)
                                    If lErr = 0 Then cbo.ListIndex = GetItem(sSavedDBname, cbo)
                                End If
                            End If
                        Else
                            Print #iOut, sBuf
                        End If
                    End If
                Loop
                
                Close #iOut
                If Len(strFileOut) And lErr = 0 Then
                    lErr = ErrExecCmd(db, strFileOut, sErr, iRetSet, lW)
                    Call DeleteFile(strFileOut)
                    strFileOut = ""
                    iOut = 0
                    If Len(sSavedDBname) <> 0 Then cbo.ListIndex = GetItem(sSavedDBname, cbo)
                End If
            End If
            Call LockWindowUpdate(0)
        End If
    End If
    
    tRes(1).Text = tRes(1).Text & IIf(Len(tRes(1).Text) <> 0, vbCrLf, "") & sErr
    
    If lErr <> 0 Then
        If mbFlex Then
            Flex.TextMatrix(0, 0) = MyLoadResString(kMsgBoxError) & ": " & lErr '& " - " & sErr
            Flex.ColWidth(0) = Flex.Width
            Flex.AllowUserResizing = flexResizeColumns
        Else
            tRes(0).Text = MyLoadResString(kMsgBoxError) & ": " & lErr
        End If
    End If
    'tStrip.Visible = True
    Set tStrip.SelectedItem = tStrip.Tabs(1)
    SizeControls imgSplitter.Top
        
    Set db = Nothing
    
    Close #iChan
    Call DeleteFile(sTmpFileIn)
    
    If Len(strFileOut) Then Call DeleteFile(strFileOut)
    
    tBar.Enabled = True
    m_bExecuting = False
    
    Screen.MousePointer = vbDefault
    Beep
    
End Sub

Private Function ErrExecCmd(ByVal oDb As SQLDMO.Database2, ByVal strFile As String, ByRef sErr As String, ByRef iRetSet As Integer, ByRef lW As Long) As Boolean

    Dim sMsg As String
    Dim lErr As Long
    Dim iIn As Integer
    Dim strCommand As String
    Dim oQry As SQLDMO.QueryResults2
    
    iRetSet = iRetSet + 1
    iIn = FreeFile
    On Local Error Resume Next
    Open strFile For Input As #iIn
    lErr = Err.Number
    If lErr = 0 Then
        strCommand = Trim$(Input$(LOF(iIn), iIn))
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    Close #iIn

    If Len(strCommand) Then
        Err.Clear
        On Local Error Resume Next
        Set oQry = oDb.ExecuteWithResultsAndMessages(strCommand, Messages:=sMsg)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
                
        'If lErr = 0 Then LoadGrid oQry, iRetSet = 1, lW
        If lErr = 0 Then LoadGrid oQry, iRetSet, lW
        tRes(1).Text = tRes(1).Text & IIf(Len(tRes(1).Text) <> 0, vbCrLf, "") & sMsg
        
        Set oQry = Nothing
    End If
    If lErr <> 0 Then
        sErr = sErr & vbCrLf & sMsg
    End If
    
    ErrExecCmd = lErr <> 0
    
End Function

Private Function ErrWriteTmpQuery(ByRef sTmpFile As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim sTmpBase As String, sTmp As String
    Dim sSql As String
    Dim lLngPos As Long, lLngLen As Long
        
    lLngPos = rtfQry.SelStart
    lLngLen = rtfQry.SelLength
    Call LockWindowUpdate(Me.hWnd)
    
    If Len(rtfQry.SelText) Then
        sSql = rtfQry.SelText
        sTmp = rtfQry.Text
        rtfQry.Text = sSql
    Else
        sSql = rtfQry.Text
    End If
    tRes(1).Text = ""
    
    With Flex
        .Redraw = False
        .Clear
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 1
        .Cols = 1
        .Redraw = True
    End With
    tRes(0).Text = ""
    
    sTmpBase = GetTempFile("tmp")
    On Local Error Resume Next
    rtfQry.SaveFile sTmpBase, rtfText
    lErr = Err.Number
    sErr = Err.Description
    If Len(sTmp) Then rtfQry.Text = sTmp
    On Local Error GoTo 0
    
    rtfQry.SelStart = lLngPos
    rtfQry.SelLength = lLngLen
    Call LockWindowUpdate(False)
    If lErr = 0 Then
        If Len(sTmpBase) Then sTmpFile = sTmpBase
    End If
        
    ErrWriteTmpQuery = lErr

End Function
Private Function ErrSetActiveDB(ByRef oActiveDB As SQLDMO.Database2, ByVal sDbName As String, ByRef sErr As String) As Long

    Dim lErr As Long
    
    On Local Error Resume Next
    Set oActiveDB = objServer.Databases(sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    ErrSetActiveDB = lErr

End Function
Private Function OutputFileOpen(ByRef strOutFile As String) As Integer

    Dim iOut As Integer
        
    iOut = FreeFile
    strOutFile = GetTempFile("tmp")
    
    On Local Error Resume Next
    Open strOutFile For Output As #iOut
    If Err = 0 Then OutputFileOpen = iOut
    On Local Error GoTo 0
        
End Function

Private Sub ChangeTBarSaveMode()

    Dim iMsg As Integer
    
    If m_iSaveMode = enSql Then
        iMsg = k_QRY_SaveQry
    Else
        iMsg = k_QRY_SaveResult
    End If
    tBar.Buttons(3).ToolTipText = MyLoadResString(iMsg)
    
End Sub
Private Sub LoadOptions()

    Dim fOpt As fGenOpt
    Set fOpt = New fGenOpt
    fOpt.Caption = Me.Caption
    fOpt.Action() = act_Q_Options
    
    fOpt.Show vbModal, Me
    Set fOpt = Nothing
    
    GetOptions
    
End Sub
Private Sub GetOptions()

    Dim i As Integer
    Dim oFont As StdFont
    Dim sBuf As String
    Dim sFont(2) As String
    
    sFont(0) = rtfQry.Font.Name & "\" & rtfQry.Font.Size
    Set oFont = New StdFont
    oFont = rtfQry.Font
    sBuf = Trim$(GetIni("QUERY", "FONT-QRY", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then
        i = InStr(sBuf, "\")
        If i Then
            On Local Error Resume Next
            rtfQry.Font.Name = Left$(sBuf, i - 1)
            rtfQry.Font.Size = Val(Mid$(sBuf, i + 1))
            If Err.Number = 0 Then
                sFont(0) = rtfQry.Font.Name & "\" & rtfQry.Font.Size
            Else
                rtfQry.Font = oFont
            End If
            On Local Error GoTo 0
        End If
    End If
    Set Me.Font = rtfQry.Font
    
    sFont(1) = Flex.Font.Name & "\" & Flex.Font.Size
    Set oFont = New StdFont
    oFont = Flex.Font
    sBuf = Trim$(GetIni("QUERY", "FONT-GRD", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then
        i = InStr(sBuf, "\")
        If i Then
            On Local Error Resume Next
            Flex.Font.Name = Left$(sBuf, i - 1)
            Flex.Font.Size = Val(Mid$(sBuf, i + 1))
            If Err.Number = 0 Then
                sFont(1) = Flex.Font.Name & "\" & Flex.Font.Size
            Else
                Flex.Font = oFont
            End If
            On Local Error GoTo 0
        End If
    End If
    
    sFont(2) = tRes(0).Font.Name & "\" & tRes(0).Font.Size
    Set oFont = New StdFont
    oFont = tRes(0).Font
    sBuf = Trim$(GetIni("QUERY", "FONT-RES", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then
        i = InStr(sBuf, "\")
        If i Then
            On Local Error Resume Next
            tRes(0).Font.Name = Left$(sBuf, i - 1)
            tRes(0).Font.Size = Val(Mid$(sBuf, i + 1))
            If Err.Number = 0 Then
                sFont(2) = tRes(0).Font.Name & "\" & tRes(0).Font.Size
            Else
                tRes(0).Font = oFont
            End If
            On Local Error GoTo 0
        End If
    End If
    Set tRes(1).Font = tRes(0).Font
    Set oFont = Nothing
    
        
    m_sGOstm = "GO"
    sBuf = Trim$(GetIni("QUERY", "TERM", App.Path & "\" & App.EXEName & ".INI"))
    If Len(sBuf) <> 0 Then m_sGOstm = sBuf


    Call WriteIni("QUERY", "FONT-QRY", sFont(0), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-GRD", sFont(1), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-RES", sFont(2), App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "TERM", m_sGOstm, App.Path & "\" & App.EXEName & ".INI")

End Sub
Private Sub ResetQryDefault()

    Call WriteIni("QUERY", "FONT-QRY", m_oFont(0).Name & "\" & m_oFont(0).Size, App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-GRD", m_oFont(1).Name & "\" & m_oFont(1).Size, App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "FONT-RES", m_oFont(2).Name & "\" & m_oFont(2).Size, App.Path & "\" & App.EXEName & ".INI")
    Call WriteIni("QUERY", "TERM", "GO", App.Path & "\" & App.EXEName & ".INI")
    
    rtfQry.Font.Name = m_oFont(0).Name
    rtfQry.Font.Size = m_oFont(0).Size
    
    Set Me.Font = rtfQry.Font
    Flex.Font.Name = m_oFont(1).Name
    Flex.Font.Size = m_oFont(1).Size
    
    tRes(0).Font.Name = m_oFont(2).Name
    tRes(0).Font.Size = m_oFont(2).Size
    
    Set tRes(1).Font = tRes(0).Font
End Sub
