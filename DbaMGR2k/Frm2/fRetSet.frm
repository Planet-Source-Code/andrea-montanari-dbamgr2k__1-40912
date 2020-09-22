VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fRetSet 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Result"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5625
   Icon            =   "fRetSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Annulla"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"fRetSet.frx":000C
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grid"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Messages"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fRetSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Qry As SQLDMO.QueryResults2
Private sDbName As String
Private sSql As String

Public Property Let DBName(ByVal sVal As String)
    sDbName = sVal
End Property
Public Property Let SQLstr(ByVal sVal As String)
    sSql = sVal
End Property
Public Sub Execute()
    
    Dim sBuf As String
    Dim lErr As Long, sErr As String
    Dim db As SQLDMO.Database2
        
    On Local Error Resume Next
    Set db = objServer.Databases(sDbName)
    
    With db
        Set Qry = .ExecuteWithResultsAndMessages(sSql, Messages:=sBuf)
    End With
    lErr = Err.Number
    If lErr <> 0 Then sErr = Err.Description
    On Local Error GoTo 0
    
    tRes.Text = sBuf
    If lErr <> 0 Then
        Flex.TextMatrix(0, 0) = MyLoadResString(kMsgBoxError) & ": " & lErr '& " - " & sErr
        Flex.ColWidth(0) = Flex.Width
        Flex.AllowUserResizing = flexResizeColumns
    Else
        tRes.Locked = True
        LoadGrid
    End If
    SizeControls
    frmMain.stbMain.Panels(1).Text = "Ready"
    Set db = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub LoadGrid()
    
    Dim iCols As Integer, lRows As Long
    Dim i As Integer, l As Long
    Dim sBuf As String
    ReDim iBool(0) As Boolean, fBool(0) As Boolean
    Dim bBolP As Boolean
    
    Debug.Print Qry Is Nothing
    If Not Qry Is Nothing Then
        If Qry.ResultSets <> 0 Then
            iCols = Qry.Columns + 1
            lRows = Qry.Rows
            With Flex
                .Cols = iCols
                .ColWidth(0) = 500
                        
                Qry.CurrentResultSet = 1
                For i = 1 To iCols - 1
                    Flex.TextMatrix(0, i) = Qry.ColumnName(i)
                Next

                On Local Error Resume Next
                For l = 1 To lRows
                    'If l > 2000 Then Exit For
                    For i = 1 To iCols - 1
                        If i > 1 Then sBuf = sBuf & vbTab
                        Select Case Qry.ColumnType(i)
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
                                sBuf = sBuf & Qry.GetColumnString(l, i)
                                .ColAlignment(i) = flexAlignLeftCenter
                            Case SQLDMO_DTypeMoney
                                sBuf = sBuf & Format$(Qry.GetColumnDouble(l, i), "#,0.0000")
                                .ColAlignment(i) = flexAlignRightCenter
                            Case SQLDMO_DTypeFloat4, SQLDMO_DTypeMoney
                                sBuf = sBuf & Format$(Qry.GetColumnDouble(l, i), "#,0.0000")
                                .ColAlignment(i) = flexAlignRightCenter
                            Case SQLDMO_DTypeFloat8
                                sBuf = sBuf & Format$(Qry.GetColumnDouble(l, i), "#,0.00000000")
                                .ColAlignment(i) = flexAlignRightCenter
                            Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
                                sBuf = sBuf & Format$(Qry.GetColumnDate(l, i), kYEAR)
                                .ColAlignment(i) = flexAlignLeftCenter
                                
                                'Aggiunto Case 2 perche' altrimenti non carica colonne Identity
                                'anche se NON e' un valore listato tra le costanti
                            Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                                sBuf = sBuf & Format$(Qry.GetColumnLong(l, i), "#,0")
                                .ColAlignment(i) = flexAlignRightCenter
                            Case SQLDMO_DTypeBigint
                                sBuf = sBuf & Qry.GetColumnBigInt(l, i)
                            
                            Case SQLDMO_DTypeBit
                                If UBound(iBool) < iCols Then
                                    ReDim Preserve iBool(iCols) As Boolean
                                    ReDim Preserve fBool(iCols) As Boolean
                                End If
                                bBolP = True
                                iBool(i) = True
                                fBool(i) = Qry.GetColumnBool(l, i)
                        End Select
                    Next
                    .AddItem l & vbTab & sBuf
                    If bBolP Then
                        .Row = .Rows - 1
                        For i = 0 To UBound(iBool)
                            If iBool(i) Then
                                Flex.Col = i
                                Set Flex.CellPicture = LoadResPicture(IIf(Abs(fBool(i)), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                .CellPictureAlignment = flexAlignCenterCenter
                            End If
                        Next
                        ReDim fBool(iCols) As Boolean
                    End If
                    sBuf = ""
                Next
                On Local Error GoTo 0
                If .Rows = 1 Then .AddItem ""
                .Row = 1
                .ScrollBars = flexScrollBarBoth
                .FixedRows = 1
                .FixedCols = 1
                
                .Redraw = True
                .TopRow = 1
            End With
        End If
    End If
    Set Qry = Nothing
    
End Sub
Private Sub cmd_Click(Index As Integer)
    
    If Index = 0 Then
        Unload Me
    End If
    
End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < 5235 Then Me.Width = 5235
            If Me.Height < 4770 Then Me.Height = 5235
            SizeControls
            b = False
        End If
    End If
End Sub

Private Sub Form_Load()

    cmd(0).Caption = MyLoadResString(k_Cancel)
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fRetSet = Nothing
End Sub
Private Sub SizeControls()
    
    cmd(0).Top = Me.ScaleHeight - (cmd(0).Height + 50)
    
    tStrip.Width = Me.ScaleWidth
    tStrip.Height = cmd(0).Top - 100
    Flex.Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
    tRes.Move tStrip.ClientLeft, tStrip.ClientTop, tStrip.ClientWidth, tStrip.ClientHeight
    
    'Flex.Left = tStrip.Left + 50
    'tRes.Left = tStrip.Left + 50
    'Flex.Width = tStrip.Width - 2 * (Flex.Left)
    'tRes.Width = tStrip.Width - 2 * (tRes.Left)
    
    'Flex.Height = tStrip.Height - 400
    'tRes.Height = Flex.Height
            
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

Private Sub tStrip_Click()

    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    Flex.Visible = i = 0
    tRes.Visible = i <> 0
    If i = 0 Then
        Flex.ZOrder
        Flex.SetFocus
    Else
        tRes.ZOrder
        tRes.SetFocus
    End If
    
End Sub
