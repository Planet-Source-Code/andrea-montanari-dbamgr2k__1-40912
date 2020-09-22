VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fTBDes 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "-"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6270
   Icon            =   "fTBDes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Index"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Trigger"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Permission"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dependencies"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eProp"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.CheckBox chkDBO 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   5640
      Top             =   360
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
            Picture         =   "fTBDes.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":05C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":0720
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":0880
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":09E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTBDes.frx":0B40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "fTBDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lWidth As Long
Private m_lHeight As Long

Private Enum eColumnProp
    eIdentity = 0
    eGuid
End Enum
Private miRow As Integer        'Riga Flex
Private miCol As Integer
Private m_blnAnnullaMod As Boolean
Private mbNewTbl As Boolean
Private m_sDbName As String
Private sStuff As String
Private Elem() As PageElement
Private iTmpType As iPropG
Private bDirty As Boolean
Private bEditable As Boolean
Private mbExit As Boolean
Private colTBBase As Collection
Private colTB As Collection
Private bLoading As Boolean
Private bAcceptSpace As Boolean
Private colComputed As Collection
Private msFileGroup As String
Private m_bIsOwner As Boolean

Public Property Let DBName(ByVal sVal As String)
    m_sDbName = sVal
End Property
Public Property Let NewTbl(ByVal bVal As Boolean)
    mbNewTbl = bVal
End Property

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
End Sub
Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyRIGHT, vbKeyLEFT
                iCode = KeyCode
                KeyCode = 0
                cbo.Visible = False
                Flex.Col = Flex.Col - (38 - iCode)
        End Select
    End If
End Sub
Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            m_blnAnnullaMod = True
            KeyAscii = 0
            Flex.SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex.SetFocus
    End Select

End Sub
Private Sub cbo_LostFocus()
    
    Dim oType As SQLDMO.SystemDatatype2
    Dim objUdt As SQLDMO.UserDefinedDatatype2
    Dim sBuf As String
    Dim bUDT As Boolean
    Dim iCol As Integer
        
    iCol = Flex.Col
    If Not m_blnAnnullaMod Then
        sBuf = cbo.List(cbo.ListIndex)
        If miCol = 3 Then   'DataType
            bUDT = cbo.ItemData(cbo.ListIndex)
            Debug.Print colTB.Item(miRow).oType
            sBuf = cbo.List(cbo.ListIndex)
            'iCol = Flex.Col
            With colTB.Item(miRow)
                .IsUDT() = bUDT
                If Not bUDT Then
                    .UdtName() = ""
                    .oType() = sBuf
                    Set oType = objServer.Databases(m_sDbName).SystemDatatypes(sBuf)
                    .oIdentity() = False 'oType.AllowIdentity
                    .oIDInc() = 0
                    .oIDInit() = 0
                    If oType.IsNumeric Then
                        .oPrecision() = 18
                        .oScale() = 0
                    End If
                    If oType.AllowLength Then
                        .oSize() = 10
                    Else
                        .oSize() = oType.MaximumLength
                    End If
                    
                Else
                    .UdtName() = Trim$(Left$(sBuf, InStr(sBuf, "(") - 1))
                    Set objUdt = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName())
                    sBuf = Mid$(sBuf, InStr(sBuf, "(") + 1)
                    sBuf = Left$(sBuf, Len(sBuf) - 1)
                    .oType() = objUdt.BaseType
                    .oNull() = objUdt.AllowNulls
                    .oDefault() = objUdt.Default
                    
                    If objServer.Databases(m_sDbName).SystemDatatypes(objUdt.BaseType).IsNumeric Then
                        .oPrecision() = objUdt.NumericPrecision
                        .oScale() = objUdt.NumericScale
                    End If
                    .oSize() = objUdt.MaxSize
                    Debug.Print .Collation()
                    .Collation() = objUdt.Collation
                    Set objUdt = Nothing
                
'                    .UdtName() = Trim$(Left$(sBuf, InStr(sBuf, "(") - 1))
'                    Set oUdt = objServer.Databases(sDbName).UserDefinedDatatypes(.UdtName())
'                    sBuf = Mid$(sBuf, InStr(sBuf, "(") + 1)
'                    sBuf = Left$(sBuf, Len(sBuf) - 1)
'                    .oType() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).BaseType
'                    .oNull() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).AllowNulls
'                    .oDefault() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).Default
'
'                    If objServer.Databases(m_sDbName).SystemDatatypes(objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).BaseType).IsNumeric Then
'                        .oPrecision() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).NumericPrecision
'                        .oScale() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).NumericScale
'                    End If
'                    .oSize() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.UdtName()).MaxSize
                End If
            
                If .oStatus() <= Changed Then .oStatus() = Changed
                Debug.Print .oName()
                bDirty = True
                
                bLoading = True
                Flex.Col = 10
                Set Flex.CellPicture = LoadResPicture(IIf(Abs(.oIdentity), k_ResChkON, k_ResChkOFF), vbResBitmap)
                sBuf = .oType()
                If bUDT Then sBuf = .UdtName() & " (" & sBuf & ")"
                
                Flex.TextMatrix(miRow, 4) = .oSize
                Flex.TextMatrix(miRow, 7) = .oPrecision
                Flex.TextMatrix(miRow, 8) = .oScale
                Flex.TextMatrix(miRow, 11) = .oIDInit
                Flex.TextMatrix(miRow, 12) = .oIDInc
                Flex.TextMatrix(miRow, 14) = .Collation
                
                Flex.Col = 5
                Set Flex.CellPicture = LoadResPicture(IIf(Abs(.oNull), k_ResChkON, k_ResChkOFF), vbResBitmap)
                
                Flex.Col = 9
                Set Flex.CellPicture = LoadResPicture(IIf(Abs(.oGuid), k_ResChkON, k_ResChkOFF), vbResBitmap)
                
                Flex.Col = 14
                Flex.CellBackColor = IIf(.IsString And Not bUDT, vbWindowBackground, vbInactiveBorder)
                bLoading = False
            
            End With
        Else
            colTB.Item(miRow).Collation = sBuf
        End If
        Flex.TextMatrix(miRow, miCol) = sBuf
        
    End If
    Set oType = Nothing
    cbo.Visible = False
'    On Local Error Resume Next
'    If Screen.ActiveForm Is Me Then Flex.SetFocus
'    On Local Error GoTo 0
    
    Flex.Col = iCol
    
    m_blnAnnullaMod = True

End Sub

Private Sub Flex_DblClick()

    If Flex.MouseCol > 0 And (Flex.MouseRow > 0) Then

        If miRow = Flex.Rows - 1 Then
            Exit Sub
        End If
        If Flex.Row <= colTB.Count Then FlexEdit vbKeySpace
    End If

End Sub

Private Sub Flex_GotFocus()
    TestActiveControl
    m_blnAnnullaMod = True
End Sub

Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iRowP As Integer
    Dim bOk As Boolean
    Dim iPoint2 As Integer
    
    If KeyCode = 46 Then
        Debug.Print "Cancello"
        iRowP = Flex.Row
        If colTB.Count And iRowP <= colTB.Count Then
            Debug.Print colTB.Item(iRowP).oName
            bOk = True
            If Len(Trim$(colTB.Item(iRowP).oName)) Then
                bOk = MsgBox(MyLoadResString(k_Confirm_Remove) & vbLf & colTB.Item(iRowP).oName(), vbQuestion Or vbOKCancel, App.EXEName) = vbOK
            End If
            If bOk Then
                If iRowP <= colTB.Count Then
                    If colTB.Item(iRowP).oStatus() < Added Then
                        iPoint2 = colTB.Item(iRowP).Point2()
                        If iPoint2 <> 0 Then colTBBase.Item(iPoint2).oStatus() = Delete
                        Debug.Print colTBBase.Item(iPoint2).oName
                        bDirty = True
                        'colTB.Item(iRowP).oStatus() = Delete
                    End If
                        
                    colTB.Remove iRowP
                    If Flex.Rows <= 3 Then
                        PrepareRow Flex.Rows - 2
                    Else
                        Flex.RemoveItem iRowP
                    End If
                    
                    If Flex.Rows <= 2 Then
                        Flex.Row = Flex.Rows - 1
                    Else
                        Flex.Row = 1
                        Flex_RowColChange
                        If Flex.Rows - 1 = iRowP Then iRowP = iRowP - 1
                        Flex.Row = iRowP
                    End If
                    Flex.Col = 2
                    Flex_RowColChange
                    
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyF2 Then
        FlexEdit 0
    End If

End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)

    Dim iCode As Integer
    
    iCode = KeyAscii
    KeyAscii = 0
    If iCode <> vbKeyReturn And iCode <> vbKeyEscape Then FlexEdit iCode

End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer
    Dim iRow As Integer, iCol As Integer
    Dim vComp As Variant
    Dim iUbound As Integer
    Dim sText As String
    Dim iPos As Integer
            
    If Button = vbRightButton Then
        If bEditable Then
            If (Flex.Rows > 3 Or Flex.Rows = 3 And colTB.Count > 0) And colTB.Item(Flex.Row).oName() <> "" Then
                
                'If Not objServer.SaLogin Then
                If Not bEditable Then
                    MsgBox MyLoadResString(k_ADMIN_Login_requested), vbOKOnly Or vbInformation, App.EXEName
                Else
                    Dim M As New cPUpMenu
                    M.Add MyLoadResString(k_Insert_Field), 1
                    M.Add "", 2, bSeparator:=True
                    M.Add MyLoadResString(k_Remove), 3
                    'If Flex.Col = 13 And colTB.Item(Flex.Row).IsComputed() Then
                    If colTB.Item(Flex.Row).IsComputed() Then
                        M.Add "", 4, bSeparator:=True
                        M.Add MyLoadResString(k_Property) & " " & MyLoadResString(k_Computed_Field), 5
                    End If
                    
                    
                    Call GetCursorPos(P)
                    i = M.Popup(Me.hWnd, P.x, P.y)
                    M.Clear
                    Set M = Nothing
                    
                    If i <> 0 Then
                        Select Case i
                            Case 1
                                With Flex
                                    iCol = .Col
                                    iRow = .Row
                                    .Col = 0
                                    Set Flex.CellPicture = Nothing
                                    .RowData(iRow) = 0
                                    .AddItem "", iRow
                                    PrepareRow iRow, iRow
                                    .Col = iCol
                                End With
                            Case 3
                                Flex.SetFocus
                                SendKeys "{DELETE}"
                            Case 5
                            
                                With fResult
                                    .tRes.Text = colTB.Item(Flex.Row).ComputedText()
                                    '.Action = IIf(colTB.Item(Flex.Row).oStatus() >= Added, act_Computed, act_Null)
                                    .Action = act_Computed
                                    
                                    .Caption = MyLoadResString(k_Computed_Text) & " ['" & colTB.Item(Flex.Row).oName() & "']"
                                End With
                                Screen.MousePointer = vbDefault
                                fResult.Show vbModal, Me
                                If fResult.Confirm() Then
                                    sText = Trim$(fResult.tRes.Text)
                                    vComp = Split(sText, vbLf)
                                    sText = ""
                                    iUbound = UBound(vComp)
                                    For iCol = 0 To iUbound
                                        iPos = InStr(vComp(iCol), vbCr)
                                        If iPos Then Mid$(vComp(iCol), iPos, 1) = " "
                                        vComp(iCol) = Trim$(vComp(iCol))
                                        If Len(vComp(iCol)) Then sText = sText & vComp(iCol) & " "
                                    Next
                                    sText = Trim$(sText)
                                    colTB.Item(Flex.Row).ComputedText() = sText
                                    colTB.Item(Flex.Row).oStatus() = Changed
                                    If Len(sText) Then colTB.Item(Flex.Row).oDefault() = ""
                                    bDirty = True
                                End If
                                Unload fResult
                                
                                With Flex
                                    iCol = .Col
                                    .Col = 13
                                    Set Flex.CellPicture = LoadResPicture(IIf(Abs(colTB.Item(Flex.Row).IsComputed()), k_ResChkON, k_ResChkOFF), vbResBitmap)
                                    .Col = iCol
                                End With
                            
                        End Select
                    End If
                End If
            End If
        Else
            'Non valido per SystemObjects
            MsgBox MyLoadResString(k_Unavailable_for_System_Objects), vbOKOnly Or vbInformation, App.EXEName
        End If
    End If
        
End Sub

Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
    Dim bNLine As Boolean
       
    If Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            With Flex
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                    If .Row = .Rows - 1 Then
                        bNLine = AddField()
                    ElseIf .Row = .Rows - 3 Then
                        Debug.Print colTB.Count
                        If .Rows - 2 <= colTB.Count Then
                            If Trim$(colTB.Item(.Rows - 2).oName()) = "" And colTB.Item(.Rows - 2).oStatus() >= Added Then
                                colTB.Remove .Rows - 2
                                .RemoveItem .Rows - 2
                            End If
                            
                        End If
                    End If
                    If .Row <> miRow Then
                        If miRow > 0 And miRow <= colTB.Count Then
                            If colTB.Item(miRow).oName() = "" And colTB.Item(miRow).oStatus >= Added Then
                                colTB.Remove miRow
                                If .Rows <= 3 Then
                                    PrepareRow .Rows - 2
                                Else
                                    .RemoveItem miRow
                                End If
                                b = False
                                iRowP = miRow
                                If .Rows <= 2 Then
                                    .Row = Flex.Rows - 1
                                Else
                                    .Row = 1
                                    Flex_RowColChange
                                    If .Rows - 1 = iRowP Then iRowP = iRowP - 1
                                    .Row = iRowP
                                End If
                                Flex.Col = iCol
                                Flex_RowColChange
                                Exit Sub
                            End If
                        End If
                    End If
                                        
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex.CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1
                    
                    .Col = iCol
                    miRow = iRowP
                    
                    Debug.Print colTB.Count
                End If

            End With
            b = False
        End If
    End If

End Sub

Private Sub Flex_Scroll()
    cbo.Visible = False
    tFlex.Visible = False
End Sub

Private Sub Flex_SelChange()
    Flex.Row = Flex.RowSel
    Flex.Col = Flex.ColSel
End Sub

Private Sub Form_Activate()

    Static bFirst As Boolean
    Dim i As Integer, iPoint2 As Integer
            
    If Not bFirst Then
        bFirst = True
        If mbExit Then
            Unload Me
            Exit Sub
        Else
            If colComputed.Count Then
                With fResult
                    .tRes.Text = MyLoadResString(k_Table_with_Computed_Fields_warning) & String$(2, vbLf) & String$(20, "-") & vbLf & MyLoadResString(k_Computed_Fields) & String$(2, vbLf)
                    .tRes.Text = .tRes.Text & MyLoadResString(k_Field_Name) & vbTab & MyLoadResString(k_Computed_Text) & vbLf
                    For i = 1 To colComputed.Count
                        iPoint2 = CInt(colComputed.Item(i))
                        .tRes.Text = .tRes.Text & colTBBase.Item(iPoint2).oName() & vbTab & colTBBase.Item(iPoint2).ComputedText() & vbLf & String$(20, "-") & vbLf
                    Next
                    .Action = act_Null
                    .Caption = MyLoadResString(k_Message_about_Tables)
                End With
                Screen.MousePointer = vbDefault
                fResult.Show vbModal, Me
            End If
        End If
    End If
End Sub
Private Sub Form_Initialize()
    bEditable = True
End Sub
Private Sub Form_Load()
    
    m_lWidth = Me.Width
    m_lHeight = Me.Height
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keySave).ToolTipText = MyLoadResString(k_Save)
    tBar.Buttons(keyIndex).ToolTipText = MyLoadResString(k_Indexes_caption)
    tBar.Buttons(keyTrigger).ToolTipText = MyLoadResString(k_Trigger_for_Table)
    tBar.Buttons(keyPermission).ToolTipText = MyLoadResString(k_Permissions)
    tBar.Buttons(keyDependencies).ToolTipText = MyLoadResString(k_Show_Dependencies)
    tBar.Buttons(keyEProp).ToolTipText = MyLoadResString(k_Extended_Property)
    chkDBO.Caption = MyLoadResString(k_OBJ_Create_As_DBO)
    chkDBO.ToolTipText = MyLoadResString(k_OBJ_Create_As_DBO_tolTip)
        
    CenterForm Me

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    TestActiveControl
    Test4Save True

End Sub
Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeight Then Me.Height = m_lHeight
            TestActiveControl
            
            SizeControls
            b = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set colComputed = Nothing
    Set colTBBase = Nothing
    Set colTB = Nothing
    Set fTBDes = Nothing
End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iImg As Integer, iLen As Integer, iElem As Integer
    
    iTmpType = am_String
    bAcceptSpace = False
        
    If StrComp(sStuff, "dbo", vbTextCompare) = 0 Then
        Exit Sub
    ElseIf Not bEditable Then
        Exit Sub
    End If
    miRow = Flex.Row - 1
    sBuf = Flex.Text
    
    bEdit = True

    miRow = Flex.Row
    miCol = Flex.Col
    bEdit = False

    With colTB.Item(miRow)
        ReDim Elem(0) As PageElement
        Select Case Flex.Col
            Case 2, 6 'Name Default
                iTmpType = am_String
                Elem(0).iType = am_String
                If Flex.Col = 2 Then iLen = k_MaxColumnLen
                bEdit = True
                If Flex.Col = 6 Then
                    iLen = k_MaxColumnDefaultLen
                    bAcceptSpace = True
                    bEdit = Not colTB.Item(miRow).IsComputed() And Not colTB.Item(miRow).IsUDT()
                End If
            Case 3  'Type
                If Not colTB.Item(Flex.Row).IsComputed() Then
                    LoadCboDataType True, KeyAscii
                    bEdit = False
                End If
            Case 4  'Size
                If .isOSizeble() Then
                    iTmpType = am_Long
                    Elem(0).iType = am_Long
                    If .IsBinary() Or .IsString() Then
                        iLen = 5
                    Else
                        iLen = 3
                    End If
                    bEdit = True
                    bEdit = Not colTB.Item(miRow).IsUDT()
                End If
            Case 5  'Null
                'If (.oStatus() >= Added And Not mbNewTbl) Or (Len(.oDefault()) <> 0) Then
                '// sempre Nullable se in Add Mode, anche se si
                '   dovrebbe anche testare se non esistono righe
                If (.oStatus() >= Added) Or (Len(.oDefault()) <> 0) Then
                    iTmpType = am_Bool
                    Elem(0).iType = am_Bool
                    bEdit = True
                    bEdit = Not colTB.Item(miRow).IsUDT()
                End If
            Case 7  'Precision
                If .isONumeric Then
                    iTmpType = am_Long
                    Elem(0).iType = am_Long
                    iLen = 2
                    bEdit = True
                    bEdit = Not colTB.Item(miRow).IsUDT()
                End If
            Case 8  'Scale
                If .isODecimal Then
                    iTmpType = am_Long
                    Elem(0).iType = am_Long
                    iLen = 2
                    bEdit = True
                    bEdit = Not colTB.Item(miRow).IsUDT()
                End If
            Case 9  'Guid RowID
                If .CanBeRowGuid() Then
                    iTmpType = am_Bool
                    Elem(0).iType = am_Bool
                    bEdit = True
                Else
                    Beep
                End If
            Case 10 'Identity
                If .oStatus() >= Added Then
                    If .CanBeInt() Then
                        iTmpType = am_Bool
                        Elem(0).iType = am_Bool
                        bEdit = True
                    Else
                        Beep
                    End If
                End If
            Case 11, 12 'IdentitySeed Increment
                If .oStatus() >= Added Then
                    If .oIdentity() Then
                        iTmpType = am_Long
                        Elem(0).iType = am_Long
                        iLen = 7
                        bEdit = True
                    End If
                End If
            Case 13 'Computed Text
                'If .oStatus() >= Added Or (.IsComputed()) Then
                    iTmpType = am_Bool
                    Elem(0).iType = am_Bool
                    bEdit = True
                'End If
            Case 14 'Collation
                LoadCboDataType False, False
                bEdit = False
        End Select
    
        If iElem < 0 Then Exit Sub
        If bEdit Then
            If (iTmpType And am_Bool) = am_Bool Then
                Select Case KeyAscii
                    Case 0, vbKeySpace
                        If Flex.Col = 5 Then
                            iImg = 1 - (-.oNull())
                            .oNull() = iImg
                        ElseIf Flex.Col = 9 Then
                            If NotOtherSameProp(miRow, eGuid) Then
                                iImg = 1 - (-.oGuid())
                                If iImg = 1 Then
                                    colTB.Item(miRow).oDefault() = "(newid())"
                                    Flex.TextMatrix(miRow, 6) = "(newid())"
                                End If
                                .oGuid() = iImg
                            End If
                        ElseIf Flex.Col = 10 Then
                            If NotOtherSameProp(miRow, eIdentity) Then
                                iImg = 1 - (-.oIdentity())
                                If iImg = 0 Then
                                    .oIDInc() = 0
                                    .oIDInit() = 0
                                End If
                                .oIdentity() = iImg
                            End If
                        ElseIf Flex.Col = 13 Then
                            iImg = 1 - (-.IsComputed())
                            .IsComputed() = iImg
                            If iImg Then
                                colTB.Item(miRow).oDefault() = ""
                                Flex.TextMatrix(miRow, 6) = ""
                            End If
                        End If
                        Set Flex.CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        bEdit = False
                        bDirty = True
                        tBar.Buttons(keySave).Enabled = True
                        If colTB.Item(miRow).oStatus() = NoVar Then colTB.Item(miRow).oStatus() = Changed
                    Case Else
                        KeyAscii = 0
                        bEdit = False
                End Select
            End If
        End If
    End With
    If bEdit Then
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf

        tFlex.SelStart = 0
        With Flex
            tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If
End Sub

Public Sub LoadTbl(Optional ByVal b1Time As Boolean = True)

    Dim oDb As SQLDMO.Database2
    Dim oTb As SQLDMO.Table
    Dim oCol As SQLDMO.Column2
    Dim iCount As Integer
    Dim oType As SQLDMO.SystemDatatype2
    Dim bUDT As Boolean
    Dim bIsNum As Boolean
    Dim sTbName As String
    Dim sOwner As String
    Dim oUdt As SQLDMO.UserDefinedDatatype2
   
    Dim oField As cRow, oNField As cRow
           
    Screen.MousePointer = vbHourglass
    bEditable = ((GetMembership(db_DDLAdmin, m_sDbName) And GetMembership(db_DataReader, m_sDbName)) Or GetMembership(db_dbo, m_sDbName)) And IsDbAvailableReadWrite(m_sDbName)
    chkDBO.Visible = mbNewTbl

    Set oDb = objServer.Databases(m_sDbName)
    
    
    bDirty = False
        
    Flex.RowHeightMin = cbo.Height
    
    tBar.Buttons(keyEProp).Enabled = False
            
    Set colTB = Nothing
    Set colTBBase = Nothing
    
    Set colTB = New Collection
    Set colTBBase = New Collection
    Set colComputed = New Collection
    
    If Not mbNewTbl Then
        sTbName = Replace(sStuff, "[", "")
        sTbName = Replace(sTbName, "]", "")
        Set oDb = objServer.Databases(m_sDbName)
        
        Set oTb = GetTbByName(objServer, m_sDbName, sTbName)
        sOwner = oTb.Owner
    
        Me.Caption = MyLoadResString(k_Table_Design) & " '" & sOwner & "." & sStuff
        
        
        If Not oTb Is Nothing Then
            'Set oTB = oDB.Tables(sStuff)
            m_bIsOwner = IsObjectOwner(m_sDbName, sTbName, am_OwnTable)
            If bEditable Then bEditable = oTb.SystemObject = False
            tBar.Buttons(keySave).Enabled = bEditable
            'tBar.Buttons(keySave).Visible = bEditable
            
            iCount = 1
            For Each oCol In oTb.Columns
                Set oField = New cRow
                Set oNField = New cRow
                With oCol
                    bIsNum = False
                    On Local Error Resume Next
                    Set oType = objServer.Databases(m_sDbName).SystemDatatypes(.DataType)
                    bUDT = Err.Number <> 0
                    On Local Error GoTo 0
                    oField.oPkey() = .InPrimaryKey
                    oField.oIdentity() = .Identity
                    oField.oName() = .Name
                    oField.oGuid() = .IsRowGuidCol
                    oField.oIDInit() = .IdentitySeed
                    oField.oIDInc() = .IdentityIncrement
                    oField.IsComputed() = .IsComputed
                    
                    If Not bUDT Then
                        oField.oType() = .DataType
                        oField.oNull() = .AllowNulls
                        oField.oDefault() = .DRIDefault.Text
                        oField.oPrecision() = .NumericPrecision
                        oField.oScale() = .NumericScale
                        oField.oSize() = .Length
                        oField.Collation() = .Collation
                    Else
                        oField.IsUDT() = bUDT
                        oField.UdtName() = .DataType
                        Set oUdt = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType)
                        
                        oField.oType() = oUdt.BaseType
                        oField.oNull() = oUdt.AllowNulls
                        oField.oDefault() = oUdt.Default
                        oField.Collation() = oUdt.Collation
                                                
                        If objServer.Databases(m_sDbName).SystemDatatypes(oUdt.BaseType).IsNumeric Then
                            bIsNum = True
                            oField.oPrecision() = oUdt.NumericPrecision
                            oField.oScale() = oUdt.NumericScale
                        End If
                        oField.oSize() = oUdt.MaxSize
                        Set oUdt = Nothing
                        
'                        oField.oType() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).BaseType
'                        oField.oNull() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).AllowNulls
'                        oField.oDefault() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).Default
'
'
'                        If objServer.Databases(m_sDbName).SystemDatatypes(objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).BaseType).IsNumeric Then
'                            bIsNum = True
'                            oField.oPrecision() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).NumericPrecision
'                            oField.oScale() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).NumericScale
'                        End If
'                        oField.oSize() = objServer.Databases(m_sDbName).UserDefinedDatatypes(.DataType).MaxSize
                    End If
                    
                    If .IsComputed Then
                        oField.ComputedText() = .ComputedText
                        colComputed.Add iCount
                    End If
                    
                    oNField.oPkey() = .InPrimaryKey
                    oNField.oIdentity() = .Identity
                    oNField.oName() = .Name
                    
                    If Not bUDT Then
                        oNField.oType() = oField.oType()
                        oNField.oNull() = oField.oNull()
                        oNField.oDefault() = oField.oDefault()
                        oNField.oPrecision() = oField.oPrecision()
                        oNField.oScale() = oField.oScale()
                    Else
                        oNField.IsUDT() = bUDT
                        oNField.UdtName() = oField.UdtName()
                        oNField.oType() = oField.oType()
                        oNField.oNull() = oField.oNull()
                        oNField.oDefault() = oField.oDefault()
                        
                        If bIsNum Then
                            oNField.oPrecision() = oField.oPrecision()
                            oNField.oScale() = oField.oScale()
                        End If
                        'oNField.oSize() = oField.oSize()
        
                    End If
                    oNField.Collation() = oField.Collation()
                    oNField.oSize() = oField.oSize()
                                    
                    oNField.oGuid() = .IsRowGuidCol
                    oNField.oIDInit() = .IdentitySeed
                    oNField.oIDInc() = .IdentityIncrement
                    oNField.IsComputed() = .IsComputed
                    If .IsComputed Then oNField.ComputedText() = .ComputedText
                    oNField.Point2() = iCount
                End With
                
                colTBBase.Add oField
                colTB.Add oNField
    
                iCount = iCount + 1
            Next
        Else
            mbExit = True
        End If
    Else
        Me.Caption = MyLoadResString(k_New_Table)
        
        Dim bConfirm As Boolean
        Dim oColl As New Collection
        
        For iCount = 1 To objServer.Databases(m_sDbName).FileGroups.Count
            On Local Error Resume Next
            oColl.Add objServer.Databases(m_sDbName).FileGroups(iCount).Name, objServer.Databases(m_sDbName).FileGroups(iCount).Name
            On Local Error GoTo 0
        Next
                
        Screen.MousePointer = vbDefault
        fPwd.collFileGroups() = oColl
        fPwd.Action() = act_Tbl
        fPwd.Caption = MyLoadResString(k_New_Table) & " ..."
        fPwd.Stuff() = "Table_1"
        fPwd.Show vbModal, Me
        bConfirm = fPwd.Confirm()
        If bConfirm Then
            sStuff = fPwd.Stuff()
            msFileGroup = fPwd.FileGroup()
        End If
        Unload fPwd
        Set oColl = Nothing
        If Not bConfirm Then mbExit = True
    End If
    Set oType = Nothing
    
    If b1Time Then
        Me.Width = 6500
        CenterForm Me
    End If
    
    Set oField = Nothing
    Set oNField = Nothing
    Set oCol = Nothing
    Set oTb = Nothing
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
      
    If Not mbNewTbl Then
        tBar.Buttons(keyIndex).Enabled = bEditable
        tBar.Buttons(keyTrigger).Enabled = bEditable
        tBar.Buttons(keyPermission).Enabled = bEditable And GetMembership(db_SecurAdm, m_sDbName)
        tBar.Buttons(keyDependencies).Enabled = bEditable And GetMembership(db_SecurAdm, m_sDbName)
        tBar.Buttons(keyEProp).Enabled = m_bIsOwner
    Else
        tBar.Buttons(keyIndex).Enabled = False
        tBar.Buttons(keyTrigger).Enabled = False
        tBar.Buttons(keyPermission).Enabled = False
        tBar.Buttons(keyDependencies).Enabled = False
    End If
    
    If Not mbExit Then LoadFlex
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub LoadFlex()

    Dim i As Integer, l As Integer
    Dim oCol As cRow
    Dim sBuf As String
    
    With Flex
        bLoading = True
        .Redraw = False
        .Clear
        .FixedCols = 0
        .Cols = 15
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 1) = MyLoadResString(k_In_Primary_Key)
        .TextMatrix(0, 2) = MyLoadResString(k_Field_Name)
        .TextMatrix(0, 3) = MyLoadResString(k_Data_Type)
        .TextMatrix(0, 4) = MyLoadResString(k_Size)
        .TextMatrix(0, 5) = MyLoadResString(k_Allow_Null)
        .TextMatrix(0, 6) = MyLoadResString(k_Default_Value)
        .TextMatrix(0, 7) = MyLoadResString(k_Precision)
        .TextMatrix(0, 8) = MyLoadResString(k_Scale)
        .TextMatrix(0, 9) = MyLoadResString(k_Is_RowGuid)
        .TextMatrix(0, 10) = MyLoadResString(k_Is_Identity)
        .TextMatrix(0, 11) = MyLoadResString(k_Initial_Value)
        .TextMatrix(0, 12) = MyLoadResString(k_Increment)
        .TextMatrix(0, 13) = MyLoadResString(k_Computed_Field)
        .TextMatrix(0, 14) = MyLoadResString(k_NewDbSort)
        
        .ColWidth(0) = 300
        .ColWidth(1) = 400
        .ColWidth(2) = 2000
        .ColWidth(3) = 1200
        .ColWidth(4) = 600
        .ColWidth(5) = 400
        .ColWidth(14) = 1800
    End With
            
    i = colTB.Count
    If i > 0 Then
        For l = 1 To i
            Set oCol = colTB.Item(l)
            sBuf = oCol.oType()
            If oCol.IsUDT() Then sBuf = oCol.UdtName() & " (" & sBuf & ")"
            Flex.AddItem vbTab & vbTab & oCol.oName & vbTab & sBuf & vbTab & Format$(oCol.oSize, "#,0") & vbTab & vbTab & oCol.oDefault & vbTab & oCol.oPrecision & vbTab & oCol.oScale & vbTab & vbTab & vbTab & oCol.oIDInit & vbTab & oCol.oIDInc & vbTab & vbTab & oCol.Collation
            Flex.Col = 1
            Flex.Row = Flex.Rows - 1
            Flex.CellBackColor = vbInactiveBorder
            'Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.oPkey), k_ResChkON, k_ResChkOFF), vbResBitmap)
            
            Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.oPkey), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex.CellPictureAlignment = flexAlignCenterCenter
                        
            Flex.Col = 3
            Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
            Flex.CellPictureAlignment = flexAlignRightTop
            
            Flex.Col = 5
            Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.oNull), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex.CellPictureAlignment = flexAlignCenterCenter
            
            Flex.Col = 9
            Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.oGuid), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex.CellPictureAlignment = flexAlignCenterCenter
            
            Flex.Col = 10
            Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.oIdentity), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex.CellPictureAlignment = flexAlignCenterCenter
            
            Flex.Col = 13
            Set Flex.CellPicture = LoadResPicture(IIf(Abs(oCol.IsComputed), k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex.CellPictureAlignment = flexAlignCenterCenter
        
            Flex.Col = 14
            Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
            Flex.CellPictureAlignment = flexAlignRightTop
            Flex.CellBackColor = IIf(oCol.IsString() And Not oCol.IsUDT(), vbWindowBackground, vbInactiveBorder)
        Next
    Else
        Flex.AddItem ""
        PrepareRow Flex.Rows - 1
    End If

    Flex.AddItem ""
                    
    With Flex
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter

        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignRightCenter
        .FixedRows = 1
        .FixedCols = 1
        
        .Redraw = True
        miRow = 0
        miCol = 0
        .TopRow = 1
        .Row = 1
        bLoading = False
        Flex_RowColChange
    End With
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sBuf As String

    TestActiveControl
    Flex.SetFocus

    If Button.Enabled Then
        Select Case Button.Key
            Case keySave
                Test4Save False
 '               Flex.SetFocus
            Case keyEnd
                Flex.SetFocus
                Unload Me
            Case keyEProp
                ManageExtendedProp

            Case Else
                If bDirty Then
                    If MsgBox(MyLoadResString(k_Uncommitted_Structure_Changes) & vbLf & MyLoadResString(k_Confirm_Loose_of_Changes), vbOKCancel Or vbQuestion, App.EXEName) = vbCancel Then
'                        If Screen.ActiveForm Is Me Then Flex.SetFocus
                        Exit Sub
                    End If
                End If
            
                Select Case Button.Key
                    Case keyIndex
                        ShowIndex
                    Case keyTrigger
                        ShowTrigger
                    Case keyPermission
                        sBuf = Replace(sStuff, "[", "")
                        sBuf = Replace(sBuf, "]", "")
                        LoadPermission m_sDbName, sBuf, act_Tbl
                    Case keyDependencies
                        ShowDependencies
                End Select
        End Select
    End If
    
               
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUP Then
        KeyCode = 0
        If miRow > 0 Then
            Flex.Row = miRow
            Flex.SetFocus
        End If
    ElseIf KeyCode = vbKeyDOWN Then
        If miRow < Flex.Rows - 2 Then
            Flex.Row = miRow + 2
            Flex.SetFocus
        End If
    Else
        If (iTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (iTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (iTmpType And am_String) = am_String Then
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93  '<>|[]
            KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex.Text
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex.SetFocus
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
            If Not bAcceptSpace Then KeyAscii = 0
        Case Else
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then
                Select Case KeyAscii
                    Case 33, 34, 39 To 47, 61, 63
                        If Flex.Col = 2 Then KeyAscii = 0
                End Select
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
        
    Dim lVal As Long
    Dim bUpdate As Boolean
    Dim sBuf As String
    Dim sErr As String
    
    If Not m_blnAnnullaMod Then
        sBuf = Trim$(tFlex.Text)
        tFlex.Text = sBuf
        Select Case miCol
            Case 2  'Nome
                If Len(sBuf) Then
                    If AddFld2Coll(Nothing, sBuf, miRow) Then
                        If colTB.Item(miRow).oName() <> sBuf Then
                            If colTB.Item(miRow).oStatus() < Added Then colTB.Item(miRow).oStatus() = Changed
                        End If
                        colTB.Item(miRow).oName() = sBuf
                        bUpdate = True
                    Else
                        m_blnAnnullaMod = True
                    End If
                End If
            Case 6  'Default
                colTB.Item(miRow).oDefault() = sBuf
                If colTB.Item(miRow).oStatus() < Added Then colTB.Item(miRow).oStatus() = Changed
                If (colTB.Item(miRow).oStatus() < Added And Not mbNewTbl And Len(sBuf) = 0) Then
                    colTB.Item(miRow).oNull() = True
                    bLoading = True
                    Flex.Col = 5
                    Set Flex.CellPicture = LoadResPicture(k_ResChkON, vbResBitmap)
                    Flex.Col = miCol
                    bLoading = False
                End If

                bUpdate = True
            Case 4  'Size
                If ValInt(sBuf) Then
                    colTB.Item(miRow).oSize() = ValInt(sBuf)
                    If colTB.Item(miRow).oStatus() < Added Then colTB.Item(miRow).oStatus() = Changed
                    bUpdate = True
                End If
            Case 7
                lVal = ValInt(sBuf)
                If lVal > 0 And lVal < 29 Then
                    colTB.Item(miRow).oPrecision() = lVal
                    If colTB.Item(miRow).oStatus() < Added Then colTB.Item(miRow).oStatus() = Changed
                    bUpdate = True
                Else
                    sErr = MyLoadResString(k_Precision_value_must_be_between_1_and_28)
                End If
            Case 8
                lVal = ValInt(sBuf)
                If lVal < 29 Then
                    colTB.Item(miRow).oScale() = lVal
                    If colTB.Item(miRow).oStatus() < Added Then colTB.Item(miRow).oStatus() = Changed
                    bUpdate = True
                Else
                    m_blnAnnullaMod = True
                End If
            Case 11
                lVal = ValInt(sBuf)
                colTB.Item(miRow).oIDInit() = lVal
                bUpdate = True
            Case 12
                lVal = ValInt(sBuf)
                colTB.Item(miRow).oIDInc() = lVal
                bUpdate = True

        End Select
    End If
    
    If bUpdate Then
        Debug.Print Flex.Rows
        Debug.Print Flex.Cols
        
        Flex.TextMatrix(miRow, miCol) = tFlex.Text
        bDirty = True
        tBar.Buttons(keySave).Enabled = True
        
    Else
        If Not m_blnAnnullaMod Then Beep
    End If

    
    m_blnAnnullaMod = True
    tFlex.Visible = False
'    On Local Error Resume Next
'    If Screen.ActiveForm Is Me Then Flex.SetFocus
'    On Local Error GoTo 0
    If Len(sErr) <> 0 Then MsgBox sErr, vbInformation Or vbOKOnly, App.EXEName
            
End Sub
Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub SizeControls()
        
    Flex.Move 0, tBar.Height + k_FlexBevel, Me.ScaleWidth - (2 * Flex.Left), Me.ScaleHeight - tBar.Height - k_FlexBevel
    SettaLargFlex Flex, 2, 2000
    tBar.Align = vbAlignTop
    tBar.Refresh
    chkDBO.Left = tBar.Width - chkDBO.Width - k_MargineFRM
            
End Sub
Private Function AddField(Optional ByVal iRow As Integer = -1) As Boolean

    Dim lErr As Long, sErr As String
    Dim iRowP As Integer
    Dim iWork As Integer
    
    If iRow > 0 Then
        iWork = iRow
    Else
        iWork = Flex.Rows - 2
    End If
    If colTB.Count > 100 Then
        MsgBox MyLoadResString(k_Limit_of) & colTB.Count & MyLoadResString(k_Columns_reached_for_Table) & sStuff, vbInformation Or vbOKOnly, App.EXEName
    Else
        With Flex
            iRowP = .Rows - 1
            If Trim$(.TextMatrix(iWork, 2)) <> "" Then
                PrepareRow iWork + 1
                
                If lErr <> 0 Then
                    MsgBox MyLoadResString(k_Adding_Column) & colTB.Item(iWork + 1).oName & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                    If .Rows > 2 Then
                        .RemoveItem .Rows - 1
                    Else
                        .AddItem ""
                        .RemoveItem .Rows - 2
                    End If
                    .Col = 1
                End If
            
                .AddItem ""
                AddField = True
            End If
            .Row = .Rows - 2
        End With
        
    End If

End Function

Private Sub ShowIndex()

    fIndex.DataNames m_sDbName, sStuff
    fIndex.Caption = fIndex.Caption & MyLoadResString(k_for_Table) & "  '" & sStuff & "' - Database '" & m_sDbName & "'"
    
    fIndex.Show vbModal, Me
    LoadTbl False
    
End Sub
Private Sub ShowTrigger()
    fResult.Caption = MyLoadResString(k_Trigger_for_Table) & " '" & sStuff & "'"
    fResult.Action = act_Tri
    fResult.Stuff_Trigger m_sDbName, sStuff
    fResult.Show vbModal, Me
End Sub
Public Property Let Stuff(ByVal sVal As String)
    sStuff = sVal
End Property

Private Sub PrepareRow(ByVal iRowP As Integer, Optional ByVal iForceR As Integer = 0)

    Dim oDb As SQLDMO.Database2
    Dim sCollation As String
    
    Set oDb = objServer.Databases(m_sDbName)
    sCollation = oDb.Collation
    Set oDb = Nothing
    If Len(sCollation) = 0 Then sCollation = objServer.Collation
    
    With Flex
        .TextMatrix(iRowP, 2) = ""
        .TextMatrix(iRowP, 3) = "varchar"
        .TextMatrix(iRowP, 4) = 10
        .TextMatrix(iRowP, 6) = ""
        .TextMatrix(iRowP, 7) = 0
        .TextMatrix(iRowP, 8) = 0
        .TextMatrix(iRowP, 11) = 0
        .TextMatrix(iRowP, 12) = 0
        .TextMatrix(iRowP, 14) = sCollation

        .Row = iRowP
        .Col = 1
        .CellBackColor = vbInactiveBorder
        Set Flex.CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter
        
        .Col = 3
        Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
        .CellPictureAlignment = flexAlignRightTop
        
        .Col = 5
        Set Flex.CellPicture = LoadResPicture(k_ResChkON, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter
                   
        .Col = 9
        Set Flex.CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter
        
        .Col = 10
        Set Flex.CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter
        
        .Col = 13
        Set Flex.CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter
        
        .Col = 14
        Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
        .CellPictureAlignment = flexAlignRightTop
        .CellBackColor = vbWindowBackground
        
        .Col = 0
        Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter

        Dim objF As cRow
        Set objF = New cRow

        objF.oStatus() = IIf(iForceR = 0, Added, Inserted)
        Debug.Print colTB.Count
        Call AddFld2Coll(objF, iForceR:=iForceR)
        Set objF = Nothing
    End With

End Sub
Private Function AddFld2Coll(Optional nObjF As cRow, _
                        Optional sBuffer As String = "", _
                        Optional ByVal iRow As Integer = 0, _
                        Optional ByVal iForceR As Integer = 0) As Boolean

    Dim i As Integer
    Dim lErr As Long, sErr As String
    Dim sTmp As String
    
    sTmp = UCase$(sBuffer)
    For i = 1 To colTB.Count
        Debug.Print i; UCase$(colTB.Item(i).oName())
        If iRow <> 0 Then
            If sTmp = UCase$(colTB.Item(i).oName()) And iRow <> i Then
                lErr = Err_Free
                Exit For
            End If
        Else
            If UCase$(nObjF.oName()) = UCase$(colTB.Item(i).oName()) Then
                lErr = Err_Free
                Exit For
            End If
        End If
    Next

    If lErr Then
        lErr = 457
        sErr = MyLoadResString(k_Object_already_exist_with_same_name)
        If iRow = 0 Then
            sTmp = nObjF.oName()
        Else
            sTmp = sBuffer
        End If
    Else
        If iRow = 0 Then
            On Local Error Resume Next
            If iForceR = 0 Then
                colTB.Add nObjF  ', objF.oName
            Else
                colTB.Add nObjF, before:=iForceR
            End If
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            sTmp = nObjF.oName()
        End If
    End If
    
    If lErr <> 0 Then
        MsgBox MyLoadResString(k_Adding_Column) & " [" & sTmp & "]" & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        AddFld2Coll = True
    End If
    
End Function
Private Function GetFldName(ByVal iFldPos As Integer) As String

    Dim i As Integer
    Dim sFld As String
    
    If Not colTBBase.Item(iFldPos).oStatus = Delete Then
        For i = 1 To colTB.Count
            If colTB.Item(i).Point2() = iFldPos Then
                If colTB.Item(i).InSelectList() Then sFld = colTB.Item(i).oName()
                'If Not colTB.Item(i).IsComputed() And Not StrComp(colTB.Item(i).oType(), "timestamp", vbTextCompare) = 0 Then sFld = colTB.Item(i).oName()
                'sFld = colTB.Item(i).oName()
                Exit For
            End If
        Next
    End If
    GetFldName = sFld
    
End Function

Private Function ScriptTriggers(oDb As SQLDMO.Database, oTable As SQLDMO.Table) As Collection
    
    Dim sFileName As String
    Dim x As SQLDMO_SCRIPT_TYPE
    Dim y As SQLDMO_SCRIPT2_TYPE
    Dim oTrans As SQLDMO.Transfer
    Dim oCol As New Collection
    Dim objTrig As cObjTrig
    
    Dim i As Integer
    
    For i = 1 To oTable.Triggers.Count
        Set oTrans = New SQLDMO.Transfer
        
        y = SQLDMOScript2_AnsiFile
        x = SQLDMOScript_Default
        
        x = x Or SQLDMOScript_AppendToFile
        'oTrans.IncludeDependencies = True
        'x = x Or SQLDMOScript_IncludeHeaders
        
        x = x Or SQLDMOScript_ObjectPermissions Or SQLDMOScript_DatabasePermissions Or SQLDMOScript_Permissions
        x = x Or SQLDMOScript_Triggers
    
        
        sFileName = GetTempFile("tmp")
        If Len(sFileName) Then
        
            oTrans.ScriptType = x
            oTrans.Script2Type = y
            oTrans.AddObjectByName oTable.Triggers.Item(i).Name, SQLDMOObj_Trigger, oTable.Triggers.Item(i).Owner
            Call oDb.ScriptTransfer(oTrans, SQLDMOXfrFile_SingleFile, sFileName)
            Set objTrig = New cObjTrig
            objTrig.oName() = oTable.Triggers.Item(i).Name
            objTrig.oFileName() = sFileName
            oCol.Add objTrig, objTrig.oName()
        End If
    Next
    
    Set objTrig = Nothing
    
    Set ScriptTriggers = oCol

End Function
Private Function GetTmpFile(sTempFile As String) As String

    Dim numOfBytes As Long
    Dim buf As String
    Dim hFile As Long
    
    On Local Error Resume Next
    hFile = lopen(sTempFile, 0)
    numOfBytes = GetFileSize(hFile, ByVal 0&)
    buf = Space(numOfBytes)
    If Err.Number = 0 Then
        lread hFile, buf, numOfBytes
    End If
    Debug.Print buf

    lclose (hFile)
    On Local Error GoTo 0
    GetTmpFile = buf

End Function

Private Function LoadCollIDX(oDb As SQLDMO.Database, oTable As SQLDMO.Table, _
                ByVal iType As idxType, Optional ByVal sErrFile As String = "") As Collection
    
    Dim oCol As New Collection
    Dim oIdx As SQLDMO.Index2
    Dim oKey As SQLDMO.Key
    Dim oCheck As SQLDMO.Check
    Dim oObj As cIDX
    Dim objKey As cKeyCasc
    'Dim oList As SQLDMO.SQLObjectList
    Dim sBuf As String
    Dim bOrphan As Boolean
    Dim oColumn As SQLDMO.Column
    Dim oQry As SQLDMO.QueryResults
    Dim vVar As Variant, iColV As Integer, iRowV As Integer
    Dim iRl As Integer
    Dim oRefTB As SQLDMO.Table
    Dim lErr As Long, sErr As String, sObjSpec As String
    Dim iPos As Integer, sStripName As String
    Dim bDescending As Boolean
    Dim bCascade(1) As Boolean
        
    Dim i As Integer
    
    
    If iType = db_idx Then
        For Each oIdx In oTable.Indexes
            If ((oIdx.Type And SQLDMOIndex_DRIIndex) <> SQLDMOIndex_DRIIndex) And ((oIdx.Type And SQLDMOIndex_Hypothetical) <> SQLDMOIndex_Hypothetical) Then
                bOrphan = False
                Set oObj = New cIDX
                
                oObj.oName() = oIdx.Name
                oObj.oType() = iType
                oObj.oIsPrimaryKey() = (oIdx.Type And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey
                'Set oList = oIdx.ListIndexedColumns
                
                i = 0
                For Each oColumn In oIdx.ListIndexedColumns
                    sBuf = GetIDXFldName(oColumn.Name)
                    If Len(sBuf) Then
                        bDescending = oIdx.GetIndexedColumnDESC(sBuf)
                        oObj.LoadFld sBuf, bDescending, i
                        i = i + 1
                    Else
                        bOrphan = True
                        Exit For
                    End If
                Next
                oObj.oAttrib() = oIdx.Type
                oObj.oFileGroup() = oIdx.FileGroup
                oObj.oFillFactor() = oIdx.FillFactor
                oObj.idxNoRecompute() = oIdx.NoRecompute()
                If Not bOrphan Then oCol.Add oObj
            End If
        Next
    ElseIf iType = db_key Then
    
        For Each oKey In oTable.Keys
            If (oKey.Type And SQLDMOKey_Foreign) = SQLDMOKey_Foreign Then
                bOrphan = False
                Set oObj = New cIDX
                
                oObj.oName() = oKey.Name
                oObj.relTable() = oTable.Name
                oObj.oAttrib() = oKey.Type
                oObj.relReferencedTable() = oKey.ReferencedTable
                oObj.oType() = iType
                oObj.oChecked() = oKey.Checked
                oObj.oExcludeReplication() = oKey.ExcludeReplication
                
                Set objKey = New cKeyCasc
                objKey.InitStruct m_sDbName, oKey.Name, bCascade(0), bCascade(1)
                Set objKey = Nothing
                
                oObj.UpdateCascade() = bCascade(0)
                oObj.DeleteCascade() = bCascade(1)
                
                For i = 1 To oKey.KeyColumns.Count
                    sBuf = GetIDXFldName(oKey.KeyColumns(i))
                    If Len(sBuf) Then
                        'bDescending = oIdx.GetIndexedColumnDESC(sBuf)
                        oObj.LoadFld sBuf, False, i - 1
                    Else
                        bOrphan = True
                        Exit For
                    End If
                Next
                If Not bOrphan Then
                    For i = 1 To oKey.ReferencedColumns.Count
                        oObj.LoadFldEx oKey.ReferencedColumns(i), i - 1
                    Next
                    oCol.Add oObj
                End If
            End If
        Next
        'Set oQry = oTable.EnumReferencingTables
        Set oQry = oTable.EnumReferencingKeys
        If Not oQry Is Nothing Then
            vVar = Qry2Var(oQry)
            If IsArray(vVar) Then
                iColV = UBound(vVar, 1)
                iRowV = UBound(vVar, 2)
                
                For iRl = 0 To iRowV
                    If vVar(2, iRl) = True Then
                        sObjSpec = ""
'                        On Local Error Resume Next
'                        Set oRefTB = oDB.Tables(vVar(0, iRl))
'                        lErr = Err.Number
'                        sErr = Err.Description
'                        On Local Error GoTo 0
                        
                        If Not KeyAlreadyExists(oCol, vVar(1, iRl)) Then
                            sObjSpec = "External Table: " & vVar(0, iRl)
                            sStripName = vVar(0, iRl)
                            iPos = InStr(sStripName, ".")
                            If iPos <> 0 Then sStripName = Mid$(sStripName, iPos + 1)
                            
                        
                            Set oRefTB = GetTbByName(objServer, m_sDbName, sStripName)
                            If oRefTB Is Nothing Then
                                lErr = Err_Free
                                sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
                            End If
                            
                            If lErr = 0 Then
                                sObjSpec = sObjSpec & " - " & "External Key: " & vVar(1, iRl)
                                On Local Error Resume Next
                                Set oKey = oRefTB.Keys(vVar(1, iRl))
                                lErr = Err.Number
                                sErr = Err.Description
                                On Local Error GoTo 0
                            
                                If lErr = 0 Then
                                    If (oKey.Type And SQLDMOKey_Foreign) = SQLDMOKey_Foreign Then
                                        bOrphan = False
                                        Set oObj = New cIDX
                                        
                                        oObj.oName() = oKey.Name
                                        oObj.relTable() = oRefTB.Name
                                        oObj.oAttrib() = oKey.Type
                                        oObj.relReferencedTable() = oKey.ReferencedTable
                                        oObj.oType() = iType
                                        oObj.oChecked() = oKey.Checked
                                        oObj.oExcludeReplication() = oKey.ExcludeReplication
                
                                        Set objKey = New cKeyCasc
                                        objKey.InitStruct m_sDbName, oKey.Name, bCascade(0), bCascade(1)
                                        Set objKey = Nothing
                                        
                                        oObj.UpdateCascade() = bCascade(0)
                                        oObj.DeleteCascade() = bCascade(1)
                                        
                                        For i = 1 To oKey.KeyColumns.Count
                                            oObj.LoadFld oKey.KeyColumns(i), False, i - 1
                                        Next
                                        For i = 1 To oKey.ReferencedColumns.Count
                                            sBuf = GetIDXFldName(oKey.ReferencedColumns(i))
                                            If Len(sBuf) Then
                                                oObj.LoadFldEx sBuf, i - 1
                                            Else
                                                bOrphan = True
                                                Exit For
                                            End If
                                        Next
                                        If Not bOrphan Then oCol.Add oObj
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If lErr <> 0 Then
                        Debug.Print sErr
                        Print2File sErrFile, sObjSpec & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                        Exit For
                    End If
                    
                Next
            End If
        End If
        'Debug.Print oTBOld.EnumReferencingKeys
        
    ElseIf iType = db_check Then
        For Each oCheck In oTable.Checks
            Set oObj = New cIDX
            oObj.oName() = oCheck.Name
            oObj.oType() = iType
            oObj.chkCheckText() = oCheck.Text
            oObj.oChecked() = oCheck.Checked
            oObj.oExcludeReplication() = oCheck.ExcludeReplication
            oCol.Add oObj
        Next
    End If
    
    Set oColumn = Nothing
    Set oCheck = Nothing
    Set oIdx = Nothing
    Set oKey = Nothing
    'Set oList = Nothing
    Set oObj = Nothing
    Set oQry = Nothing
    Set oRefTB = Nothing
    
    Set LoadCollIDX = oCol

End Function

Private Function GetIDXFldName(ByVal sInFld As String) As String

    Dim i As Integer, iFldPos As Integer
    Dim sFld As String
    
    sInFld = UCase$(sInFld)
    For i = 1 To colTBBase.Count
        If UCase$(colTBBase.Item(i).oName()) = sInFld Then
            iFldPos = i
            Exit For
        End If
    Next
    If i <= colTBBase.Count Then
        For i = 1 To colTB.Count
            If colTB.Item(i).Point2() = iFldPos Then
                sFld = colTB.Item(i).oName()
                Exit For
            End If
        Next
    End If
    GetIDXFldName = sFld
    
End Function

Private Function ErrCreateOBJ(oDb As SQLDMO.Database2, ByVal sErrFile As String, _
                    oTb As SQLDMO.Table, oCol As Collection, ByVal iType As idxType) As Long

    Dim lErr As Long, sErr As String
    Dim oIdx As SQLDMO.Index2
    Dim oKey As SQLDMO.Key
    Dim oCheck As SQLDMO.Check
    Dim oObj As cIDX
    Dim bKey As Boolean
    Dim oRelTB As SQLDMO.Table
    Dim bDescending As Boolean

    
    Dim i As Integer, y As Integer
    Dim sBuf As String, sColumn As String
    Dim sObjSpec
    Dim lType As Long
        
    Select Case iType
        Case db_idx
            For i = 1 To oCol.Count
                bKey = False
                Set oObj = oCol.Item(i)
                
                With oObj
                    lType = .oAttrib()
                    If ((lType And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey) Or ((lType And SQLDMOIndex_DRIUniqueKey) = SQLDMOIndex_DRIUniqueKey) Then
                        bKey = True
                        Set oKey = New SQLDMO.Key
                        oKey.Name = .oName()
                        sObjSpec = "Key: " & .oName()

                        If (lType And SQLDMOIndex_DRIPrimaryKey) = SQLDMOIndex_DRIPrimaryKey Then
                            oKey.Type = SQLDMOKey_Primary
                        ElseIf (lType And SQLDMOKey_Unique) = SQLDMOKey_Unique Then
                            oKey.Type = SQLDMOKey_Unique
                        End If
                                                
                        oKey.FileGroup = .oFileGroup()
                        oKey.FillFactor = .oFillFactor()
                        If (lType And SQLDMOIndex_Clustered) = SQLDMOIndex_Clustered Then oKey.Clustered = True
                        For y = 0 To .oFldCount()
                            sBuf = Trim$(.GetFld(y, False))
                            If Len(sBuf) <> 0 Then oKey.KeyColumns.Add sBuf
                        Next
                        
                    Else
                        sColumn = ""
                        For y = 0 To .oFldCount()
                            sBuf = Trim$(.GetFld(y, bDescending))
                            If Len(sBuf) <> 0 Then
                                sColumn = sColumn & IIf(Len(sColumn) <> 0, "," & sBuf, sBuf)
                            End If
                        Next

                
                        If Len(sColumn) <> 0 Then
                            Set oIdx = New SQLDMO.Index
                            If (lType And SQLDMOIndex_Clustered) = SQLDMOIndex_Clustered Then oIdx.Type = SQLDMOIndex_Clustered
                            If (lType And SQLDMOIndex_Unique) = SQLDMOIndex_Unique Then oIdx.Type = oIdx.Type Or SQLDMOIndex_Unique
                            If (lType And SQLDMOIndex_IgnoreDupKey) = SQLDMOIndex_IgnoreDupKey Then oIdx.Type = oIdx.Type Or SQLDMOIndex_IgnoreDupKey
                            If (lType And SQLDMOIndex_PadIndex) = SQLDMOIndex_PadIndex Then oIdx.Type = oIdx.Type Or SQLDMOIndex_PadIndex
                            If (lType And SQLDMOIndex_Default) = SQLDMOIndex_Default Then oIdx.Type = oIdx.Type Or SQLDMOIndex_Default
                            oIdx.Name = .oName()
                            sObjSpec = "Index: " & .oName()
                            
                            oIdx.IndexedColumns = sColumn
                            oIdx.FileGroup = .oFileGroup()
                            oIdx.NoRecompute = .idxNoRecompute()
                            oIdx.FillFactor = .oFillFactor()
                            
                            For y = 0 To .oFldCount()
                                sBuf = Trim$(.GetFld(y, bDescending))
                                If Len(sBuf) <> 0 And bDescending Then oIdx.SetIndexedColumnDESC sBuf, True
                            Next
                            
                        End If
                        
                    End If
                End With
                On Local Error Resume Next
                If bKey Then
                    oTb.Keys.Add oKey
                Else
                    oTb.Indexes.Add oIdx
                End If
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                                
                If lErr <> 0 Then
                    Print2File sErrFile, sObjSpec & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                    Exit For
                End If
            Next
        Case db_key
            For i = 1 To oCol.Count
                Set oObj = oCol.Item(i)
                
                lErr = ErrSaveRefKey(oDb, oObj, sErr)
                If lErr <> 0 Then
                    Print2File sErrFile, sObjSpec & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                    Exit For
                End If
            Next
        Case db_check
            For i = 1 To oCol.Count
                Set oObj = oCol.Item(i)
                With oObj
                    Set oCheck = New SQLDMO.Check
                    oCheck.Name = .oName()
                    sObjSpec = "Check: " & .oName()
                    oCheck.Text = .chkCheckText()
                    oCheck.ExcludeReplication = .oExcludeReplication()
                End With
                On Local Error Resume Next
                oTb.Checks.Add oCheck
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                If lErr <> 0 Then
                    Print2File sErrFile, sObjSpec & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                    Exit For
                End If
            Next
    End Select
    
    Set oCheck = Nothing
    Set oIdx = Nothing
    Set oKey = Nothing
    Set oObj = Nothing
    Set oRelTB = Nothing
    
    ErrCreateOBJ = lErr
    
End Function
Private Sub Print2File(ByVal sFile As String, ByVal sMsg As String)

    Dim iFile As Integer
    
    If Len(sFile) Then
        iFile = FreeFile
    
        On Local Error Resume Next
        Open sFile For Append As iFile
        Print #iFile, sMsg
        On Local Error GoTo 0
        Close iFile
    End If
    Debug.Print sMsg
End Sub

Private Sub SaveTbl()
    
    Dim obj As cRow
    Dim oTb As SQLDMO.Table
    Dim oTBOld As SQLDMO.Table
    Dim oDb As SQLDMO.Database2
    Dim oNewCol As SQLDMO.Column2
    Dim oUdt As SQLDMO.UserDefinedDatatype2
    
    Dim lErr As Long, sErr As String
    Dim iWork As Integer, i As Integer
    Dim iOri As Integer
    Dim bTran As Boolean
    Dim bModif As Boolean
    Dim sColRowGuidDMOError As String
    Dim sTBstruct As String
        
    Dim sNewTBName As String
    Dim sSqlFrom As String
    Dim sSqlIn As String
    Dim sSql As String
    Dim sFieldOut As String, bExport As Boolean
    Dim sBuf As String
    Dim sErrFile As String
    Dim sTmpPropTB As String
    
    
    Dim bHasIdentity As Boolean
    Dim sIdentityOn As String
    Dim sIdentityOFF As String
    Dim sOwner As String
       
    Dim oQry As SQLDMO.QueryResults
    Dim oOList As SQLDMO.SQLObjectList
    'Dim o2List As SQLDMO.SQLObjectList
    Dim oPerm As SQLDMO.Permission
    
    Dim colPriv As New Collection
    Dim colIdx As New Collection
    Dim colKey As New Collection
    Dim colCheck As New Collection
    Dim colTri As New Collection
                    
    Screen.MousePointer = vbHourglass
        
    iWork = colTB.Count
    iOri = colTBBase.Count
        
    For i = 1 To iOri
        If Not bModif Then bModif = colTBBase.Item(i).oStatus() = Delete
        If bModif Then Exit For
    Next
    
    If Not bModif Then
        bModif = iWork <> iOri
        If Not bModif Then
            For i = 1 To iWork
                If Not bModif Then bModif = colTB.Item(i).oStatus() <> NoVar
                If bModif Then Exit For
            Next
        End If
    End If
    
    If mbNewTbl And bModif Then
        If iWork = 1 Then bModif = Trim$(colTB.Item(i).oName()) <> ""
    End If
    
    If bModif Then
        
        sStuff = Replace(sStuff, "[", "")
        sStuff = Replace(sStuff, "]", "")
    
        sErrFile = GetTempFile("err")
        On Local Error Resume Next
        Set oDb = objServer.Databases(m_sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        
        If lErr = 0 Then
            
            If mbNewTbl Then
                sNewTBName = sStuff
            Else
                sNewTBName = "myTmp_" & sStuff
            End If

            If Not mbNewTbl Then
                sOwner = GetTableOwner(objServer, m_sDbName, sStuff)
            Else
                If chkDBO.Value Then
                    sOwner = "dbo"
                Else
                    sOwner = GetCreateOwner(m_sDbName)
                End If
            End If

            Set oTb = New SQLDMO.Table
            oTb.Name = sNewTBName
            oTb.Owner = sOwner
            oTb.FileGroup = msFileGroup

            If Not mbNewTbl Then
                
                Set oTBOld = GetTbByName(objServer, oDb.Name, sStuff)
                If oTBOld Is Nothing Then
                    lErr = Err_Free
                    sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
                End If
                If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                
                
                If lErr = 0 Then
                    Set colTri = ScriptTriggers(oDb, oTBOld)
                    
                    Set colIdx = LoadCollIDX(oDb, oTBOld, db_idx)
                    Set colKey = LoadCollIDX(oDb, oTBOld, db_key, sErrFile)
                    Set colCheck = LoadCollIDX(oDb, oTBOld, db_check)
    
                    Set oOList = oTBOld.ListPermissions
                    Debug.Print oOList.Count
    
                    For Each oPerm In oOList
                        Debug.Print oPerm.Grantee
                        sSql = IIf(oPerm.Granted, "GRANT", "DENY") & " " & oPerm.PrivilegeTypeName & " ON dbo." & oPerm.ObjectName & " TO " & oPerm.Grantee & IIf(oPerm.Granted, " AS " & oPerm.ObjectOwner, "")
                        Debug.Print oPerm.Granted
    
                        Debug.Print oPerm.ObjectTypeName
                        Debug.Print oPerm.PrivilegeTypeName
                        Debug.Print oPerm.PrivilegeType
                        Debug.Print oPerm.ObjectOwner
                        Debug.Print oPerm.ObjectName
                        colPriv.Add sSql
        
                    Next
                End If
            End If
            If lErr = 0 Then
                objServer.BeginTransaction
                bTran = True
            
                With oTb
                    For i = 1 To iWork
                        Set obj = colTB.Item(i)
                        Set oNewCol = New SQLDMO.Column
                        If Len(obj.oName()) Then
                            Debug.Print obj.oName
                            oNewCol.Name = obj.oName()
                            If Not bHasIdentity Then bHasIdentity = obj.oIdentity()
                            
                            If Not obj.IsUDT Then
                                
                                oNewCol.DataType = obj.oType()
                                sBuf = Trim$(obj.oDefault())
                                Debug.Print sBuf
    
                                'If Left$(sBuf, 2) = "('" Then sBuf = Mid$(sBuf, 3)
                                'If Right$(sBuf, 2) = "')" Then sBuf = Left$(sBuf, Len(sBuf) - 2)
                                    
                                'If Left$(sBuf, 1) = """" Then sBuf = Mid$(sBuf, 2)
                                'If Right$(sBuf, 1) = """" Then sBuf = Left$(sBuf, Len(sBuf) - 1)
                                
                                Do While Left$(sBuf, 1) = "(" And Right$(sBuf, 1) = ")"
                                    sBuf = Mid$(sBuf, 2)
                                    sBuf = Left$(sBuf, Len(sBuf) - 1)
                                Loop
'                                If Left$(sBuf, 1) = "(" And Right$(sBuf, 1) = ")" Then
'                                    sBuf = Mid$(sBuf, 2)
'                                    sBuf = Left$(sBuf, Len(sBuf) - 1)
'                                End If
                                    
'                                If Right$(sBuf, 1) = ")" Then
'                                    If Right$(sBuf, 2) <> "()" And Right$(sBuf, 2) <> "))" Then sBuf = Left$(sBuf, Len(sBuf) - 1)
'                                End If
                                'If obj.isOSizeble() And (Left$(sBuf, 1) = "'" And Right$(sBuf, 1) = "'") Then
                                '    sBuf = Mid$(sBuf, 2, Len(sBuf) - 2)
                                'End If
                                'Debug.Print IsSqlFunction(sBuf)
                                'oNewCol.DRIDefault.Text = IIf(obj.isOSizeble(), SqlFinalS(ApiciSQL(sBuf)), sBuf)
                                oNewCol.DRIDefault.Text = sBuf
                                
                                Debug.Print oNewCol.Name
                                Debug.Print oNewCol.DRIDefault.Text
                            Else
                                'must be the owner of the UDT, else will abort, while via
                                'DDL statement will always succed
                                Set oUdt = oDb.UserDefinedDatatypes(obj.UdtName())
                                oNewCol.DataType = obj.UdtName()
                                obj.Collation = oUdt.Collation
                                Set oUdt = Nothing
                            End If
                            
                            If obj.isONumeric() Then
                                oNewCol.Identity = obj.oIdentity()
                                oNewCol.IdentitySeed = obj.oIDInit()
                                oNewCol.IdentityIncrement = obj.oIDInc()
                                oNewCol.NumericPrecision = obj.oPrecision()
                                oNewCol.NumericScale = obj.oScale()
                            End If
                            If obj.isOSizeble() Then oNewCol.Length = obj.oSize()
                            If obj.IsComputed() Then
                                oNewCol.IsComputed = True
                                oNewCol.ComputedText = obj.ComputedText()
                            End If
                            If obj.CanBeRowGuid() Then
                                'dmo error in saving tbl with  IsRowGuidCol=True and AllowNulls=False
                                If obj.oGuid() And Not obj.oNull() Then sColRowGuidDMOError = oNewCol.Name
                                oNewCol.IsRowGuidCol = obj.oGuid()
                                
                            End If
                            oNewCol.AllowNulls = obj.oNull()
                                                        
                            On Local Error Resume Next
                            oTb.Columns.Add oNewCol
                            lErr = Err.Number
                            sErr = Err.Description
                            On Local Error GoTo 0
                            
                            If lErr <> 0 Then
                                Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                                Exit For
                            End If
                        End If
                    Next
                End With
            End If
        End If
        Set oNewCol = Nothing
        If lErr = 0 Then
            On Local Error Resume Next
            If Len(sColRowGuidDMOError) = 0 Then
                oDb.Tables.Add oTb
            Else
                sTBstruct = oTb.GenerateSQL(oDb)
                sTBstruct = Replace(sTBstruct, "[" & sNewTBName & "]", sOwner & ".[" & sNewTBName & "]")
                'adds NOT NULL column validation
                sTBstruct = Replace(sTBstruct, "ROWGUIDCOL ", "ROWGUIDCOL NOT NULL ")
                'removes blank "CONSTRAINT [] name
                sTBstruct = Replace(sTBstruct, "CONSTRAINT []", " ")
                
                oDb.ExecuteImmediate sTBstruct, SQLDMOExec_Default
                If Err.Number = 0 Then
                    oDb.Tables.Refresh
                    'Set oTB = oDB.Tables("dbo.[" & sNewTBName & "]")
                    
                    
                    'Set oTB = oDB.Tables(sOwner & ".[" & sNewTBName & "]")
                    Set oTb = GetTbByName(objServer, oDb.Name, sNewTBName)
                    If oTb Is Nothing Then
                        lErr = Err_Free
                        sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
                    End If
                    If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                    
                End If
            End If
            
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            If lErr <> 0 Then
                Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
            End If
        End If
                               
        If lErr = 0 Then
'            oDb.Tables.Refresh True
'            Set oTb = GetTbByName(objServer, m_sDbName, sNewTBName)
            For Each obj In colTB
                Debug.Print obj.oName
                If obj.IsString() And Not obj.IsUDT Then
                    On Local Error Resume Next
                    Set oNewCol = oTb.Columns(obj.oName)
                    If Len(obj.Collation()) <> 0 Then
                        oNewCol.Collation = obj.Collation()
                    Else
                        oNewCol.Collation = oDb.Collation
                    End If
                    If Len(oNewCol.Collation) = 0 Then oNewCol.Collation = objServer.Collation
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    If lErr <> 0 Then Exit For
                End If
                If Not obj.IsUDT Then
                    On Local Error Resume Next
                    Set oNewCol = oTb.Columns(obj.oName)
                    oNewCol.AllowNulls = obj.oNull()
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    If lErr <> 0 Then Exit For
                End If
            Next
            If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        End If
               
        If lErr = 0 And Not mbNewTbl Then
            lErr = ErrTransferTB_COL_Properties(oDb, sOwner, sStuff, sNewTBName, sErr)
            If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        End If
               
        If lErr = 0 And Not mbNewTbl Then
            sTmpPropTB = "tmp" & Format$(Now, "YYYYMMDD") & "am"
            lErr = ErrDumpProperties(True, oDb, sStuff, sOwner, sTmpPropTB, colTri, colIdx, colKey, sErr)
            If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        End If
               
        If lErr = 0 Then
            sSqlFrom = ""
            sSqlIn = ""
            
            'If bRegen Then
            If Not mbNewTbl Then
                For i = 1 To iOri
                                     
                    sFieldOut = GetFldName(i) 'colTBBase.Item(i).oName())
                    If Len(sFieldOut) Then
                        bExport = True
                        If Len(sSqlFrom) Then sSqlFrom = sSqlFrom & ","
                        If Len(sSqlIn) Then sSqlIn = sSqlIn & ","
                        sSqlIn = sSqlIn & "[" & sFieldOut & "]"
                        sSqlFrom = sSqlFrom & "[" & colTBBase.Item(i).oName() & "]"
                        
                        'sSqlIn = sSqlIn & sOwner & ".[" & sFieldOut & "]"
                        'sSqlFrom = sSqlFrom & sOwner & ".[" & colTBBase.Item(i).oName() & "]"
                        
                    End If
                Next
                Debug.Print sSqlFrom
                Debug.Print sSqlIn
                
                If bExport Then
                    If bHasIdentity Then
                        'SET IDENTITY_INSERT products ON
                        'sIdentityOn = ReplaceMsg("SET IDENTITY_INSERT dbo.[1%] 2%", Array("1%", "2%"), Array(sNewTBName, "ON")) & vbCrLf
                        'sIdentityOFF = vbCrLf & ReplaceMsg("SET IDENTITY_INSERT dbo.[1%] 2%", Array("1%", "2%"), Array(sNewTBName, "OFF"))
                        
                        sIdentityOn = ReplaceMsg("SET IDENTITY_INSERT 1%.[2%] 3%", Array("1%", "2%", "3%"), Array(sOwner, sNewTBName, "ON")) & vbCrLf
                        sIdentityOFF = vbCrLf & ReplaceMsg("SET IDENTITY_INSERT 1%.[2%] 3%", Array("1%", "2%", "3%"), Array(sOwner, sNewTBName, "OFF"))
                        
                    End If
                
                    sSqlIn = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE" & vbCrLf _
                        & "IF EXISTS (SELECT * FROM " & sOwner & ".[" & sStuff & "]) BEGIN" & vbCrLf & sIdentityOn & "EXEC('INSERT INTO " & sOwner & ".[" & sNewTBName & "] (" & sSqlIn & ") "
                        
                    sSqlFrom = "SELECT " & sSqlFrom & " FROM " & sOwner & ".[" & sStuff & "] TABLOCKX')"
                    
                    sSql = sSqlIn & vbCrLf & sSqlFrom & sIdentityOFF & " END"
                    
                    Debug.Print sSql
                    On Local Error Resume Next
                    oDb.ExecuteImmediate sSql, SQLDMOExec_QI_ON
                    lErr = Err.Number
                    sErr = Err.Description
                    Debug.Print Err.Description
                    On Local Error GoTo 0
                    If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                End If
                
                
                If lErr = 0 And Not mbNewTbl Then
                    If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                End If
                
                
                If lErr = 0 And Not mbNewTbl Then
                    For i = 1 To colKey.Count
                    
                        Dim KeyTB As SQLDMO.Table
                        Set KeyTB = GetTbByName(objServer, m_sDbName, colKey(i).relTable())
                        Debug.Print KeyTB.Name
                        If Not KeyTB Is Nothing Then
                            Debug.Print KeyTB.Keys.Count
                            On Local Error Resume Next
                            KeyTB.Keys.Remove colKey(i).oName()
                            lErr = Err.Number
                            sErr = Err.Description
                            Debug.Print Err.Description
                            On Local Error GoTo 0
                        End If
                        Set KeyTB = Nothing
                        If lErr <> 0 Then
                            Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                            Exit For
                        End If
                    Next
                    'If colKey.Count Then oTB.Keys.Refresh True
                End If
                
                If lErr = 0 Then
                    On Local Error Resume Next
                    oDb.Tables.Remove "[" & sStuff & "]", sOwner
                    lErr = Err.Number
                    If lErr = 0 Then
                        oTb.Name = sStuff
                        Debug.Print Err.Description
                        Debug.Print oTb.Owner
                        lErr = Err.Number
                    End If
                    sErr = Err.Description
                    On Local Error GoTo 0
                    If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                End If
                
                If lErr = 0 And bExport Then
                    For i = 1 To colTri.Count
                        sSql = GetTmpFile(colTri.Item(i).oFileName())
                        If Len(sSql) Then
                            On Local Error Resume Next
                            oDb.ExecuteImmediate sSql, SQLDMOExec_QI_ON
                            lErr = Err.Number
                            sErr = Err.Description
                            Debug.Print Err.Description
                            On Local Error GoTo 0
                        End If
                        If lErr <> 0 Then
                            Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                            Exit For
                        End If
                    Next
                    If lErr = 0 Then lErr = ErrCreateOBJ(oDb, sErrFile, oTb, colIdx, db_idx)
                    If lErr = 0 Then lErr = ErrCreateOBJ(oDb, sErrFile, oTb, colCheck, db_check)
                    If lErr = 0 Then lErr = ErrCreateOBJ(oDb, sErrFile, oTb, colKey, db_key)
                
                    If lErr = 0 Then
                        sSql = ""
                        For i = 1 To colPriv.Count
                            If Len(sSql) Then sSql = sSql & vbCrLf
                            sSql = sSql & colPriv.Item(i)
                        Next
                        Debug.Print sSql
                    
                        If Len(sSql) Then sSql = sSql & vbCrLf
                        sSql = sSql & "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
                        On Local Error Resume Next
                        oDb.ExecuteImmediate sSql, SQLDMOExec_QI_ON
                        lErr = Err.Number
                        sErr = Err.Description
                        Debug.Print Err.Description
                        On Local Error GoTo 0
                        If lErr <> 0 Then Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
                    End If
                End If
                
                
            End If
        End If
        
    End If
    
    
    Set obj = Nothing
    Set oQry = Nothing
    Set oOList = Nothing
    'set o2List = Nothing
    Set oPerm = Nothing
    Set colCheck = Nothing
    
    Set oNewCol = Nothing
    Set oTBOld = Nothing
        
    Screen.MousePointer = vbDefault
    If bModif Then
                
        If lErr <> 0 Then
            sStuff = "[" & sStuff & "]"
            lErr = 0
            On Local Error Resume Next
            If bTran Then objServer.RollbackTransaction
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            If lErr <> 0 Then Print2File sErrFile, MyLoadResString(k_Abnormaly_Broken_Transaction) & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
            Beep
            fResult.tRes.Text = MyLoadResString(IIf(mbNewTbl, k_Table_Creation_Not_Allowed, k_Table_Modification_Not_Allowed)) & " '" & sStuff & "'" & vbCrLf & GetTmpFile(sErrFile)
            fResult.Action = act_Null
            fResult.WWrapVisible() = False
            fResult.Caption = MyLoadResString(k_Error_Executing_Table_Definition_for) & sStuff
            Screen.MousePointer = vbDefault
            fResult.Show vbModal, Me
                    
        Else
            objServer.CommitTransaction
            objServer.Databases(m_sDbName).Tables.Refresh True
            
            If lErr = 0 And Not mbNewTbl Then
                lErr = ErrDumpProperties(False, oDb, sStuff, sOwner, sTmpPropTB, colTri, colIdx, colKey, sErr)
                If lErr <> 0 Then
                    Print2File sErrFile, MyLoadResString(kMsgBoxError) & ": " & sErr
                                        
                    fResult.tRes.Text = ReplaceMsg(MyLoadResString(k_Err_Importing_Extended_Properties_4_Tb_Desc), Array("1%", "|"), Array(sStuff, vbCrLf)) & GetTmpFile(sErrFile)
                    fResult.Action = act_Null
                    fResult.WWrapVisible() = False
                    fResult.Caption = Replace(MyLoadResString(k_Err_Importing_Extended_Properties_4_Tb), "1%", sStuff)
                    Screen.MousePointer = vbDefault
                    fResult.Show vbModal, Me
                    
                End If
            End If
            sStuff = "[" & sStuff & "]"
            mbNewTbl = False
            bDirty = False
            LoadTbl False
            
        End If
    End If
    
    If Not mbNewTbl And Len(sTmpPropTB) <> 0 Then
        sSql = ReplaceMsg("if exists (select * from dbo.sysobjects where id = object_id(N'[1%].[2%]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [1%].[2%]", Array("1%", "2%"), Array(sOwner, sTmpPropTB))
        Call ErrExexSqlDirect(oDb, sSql, "")
    End If
    
    Set colPriv = Nothing
    Set colIdx = Nothing
    Set colKey = Nothing
    Set colTri = Nothing
    
    Set oTb = Nothing
    Set oDb = Nothing
    
    For i = 1 To colTri.Count
        Call DeleteFile(colTri.Item(i).oFileName())
    Next
    
    Call DeleteFile(sErrFile)
    
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

Private Sub LoadCboDataType(ByVal bDataType As Boolean, ByVal iKeyAscii As Integer)

    Dim i As Integer
    Dim oType As SQLDMO.SystemDatatype2
    Dim oDb As SQLDMO.Database2
    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    
    Set oDb = objServer.Databases(m_sDbName)
    
    cbo.Clear
    If bDataType Then
        For Each oType In oDb.SystemDatatypes
            cbo.AddItem oType.Name
        Next
        Set oType = Nothing
        
        For i = 1 To oDb.UserDefinedDatatypes.Count
            cbo.AddItem oDb.UserDefinedDatatypes(i).Name & " (" & oDb.UserDefinedDatatypes(i).BaseType & ")"
            cbo.ItemData(cbo.NewIndex) = True
        Next
    Else
        Set oQry = objServer.EnumCollations()
        If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
        If IsArray(vVar) Then
            With cbo
                For i = 0 To UBound(vVar, 2)
                    .AddItem vVar(0, i)
                Next
            End With
        End If
    End If
    Set oDb = Nothing
        
    With Flex
        cbo.ListIndex = 0
        cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
        cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
        cbo.Visible = True
        cbo.SetFocus
    End With
        
    If iKeyAscii <> 0 Then
        On Local Error Resume Next
        SendKeys Chr$(iKeyAscii)
        On Local Error GoTo 0
    End If
    
End Sub
Private Sub Test4Save(ByVal blnAsk As Boolean)

    Dim bConfirm As Boolean
    Dim i As Integer, iPoint2 As Integer
                
    If bDirty Then
        If blnAsk Then
            bConfirm = MsgBox(MyLoadResString(k_Uncommitted_Structure_Changes) & vbLf & MyLoadResString(k_Confirm_Loose_of_Changes), vbOKCancel Or vbQuestion, App.EXEName) = vbCancel
        Else
            bConfirm = True
        End If
        
        If bConfirm Then
            If colComputed.Count Then
                With fResult
                    .tRes.Text = MyLoadResString(k_Committ_Structural_Changes) & String$(2, vbLf)
                    .tRes.Text = .tRes.Text & MyLoadResString(k_Table_with_Computed_Fields_warning) & String$(2, vbLf) & String$(20, "-") & vbLf & MyLoadResString(k_Computed_Fields) & String$(2, vbLf)
                    .tRes.Text = .tRes.Text & MyLoadResString(k_Field_Name) & vbTab & MyLoadResString(k_Computed_Text) & vbLf
                    For i = 1 To colComputed.Count
                        iPoint2 = CInt(colComputed.Item(i))
                        .tRes.Text = .tRes.Text & colTBBase.Item(iPoint2).oName() & vbTab & colTBBase.Item(iPoint2).ComputedText() & vbLf & String$(20, "-") & vbLf
                    Next
                    .Action = act_Response
                    .WWrapVisible() = False
                    .Caption = MyLoadResString(k_Message_about_Tables)
                End With
                fResult.Show vbModal, Me
                bConfirm = fResult.Confirm()
                Unload fResult
            
            End If
            If bConfirm Then SaveTbl
        End If
    End If

End Sub
Private Sub ShowDependencies()

    Dim sBuf As String
    
    sBuf = Replace(sStuff, "[", "")
    sBuf = Replace(sBuf, "]", "")
    
    Dim objDependent As clsDepen
    Set objDependent = New clsDepen
    objDependent.ShowDependent m_sDbName, sBuf, tableDrop
    Set objDependent = Nothing

End Sub
Private Function NotOtherSameProp(ByVal iElem As Integer, ByVal eTest As eColumnProp) As Boolean

    Dim i As Integer
    Dim bRet As Boolean
    
    With colTB
        For i = 1 To .Count
            Debug.Print .Item(i).oName
            If i <> iElem Then
                If eTest = eGuid Then
                    bRet = .Item(i).oGuid()
                ElseIf eTest = eIdentity Then
                    bRet = .Item(i).oIdentity()
                End If
                If bRet Then Exit For
            End If
        Next
    End With
    
    NotOtherSameProp = Not bRet
End Function

Private Sub ManageExtendedProp()

    Dim iMode As am_SqlPropType
    Dim iOwner As am_SqlPropTypeOwner
    Dim sParam As String
            
    sParam = sStuff
    sParam = Replace(sParam, "[", "")
    sParam = Replace(sParam, "]", "")
        
    iMode = am_TbProp
    iOwner = am_OwnTable
    
    If Not m_bIsOwner Then Exit Sub
    
    LoadExtendedProp iMode, m_sDbName, GetObjectOwner(m_sDbName, sStuff, iOwner), sParam, sParam
End Sub

Private Function ErrTransferTB_COL_Properties(ByVal oDb As SQLDMO.Database2, ByVal sOwner As String, ByVal sBaseTable As String, ByVal sNewTable As String, ByRef sErr As String) As Long

    Dim sSql As String
    Dim vVarArray As Variant
    Dim lErr As Long
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim s2LevelObj As String
        
    Const kSqlList As String = "SELECT [name] FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
    Const kSqlAdd As String = "DECLARE @vVar SQL_VARIANT" & vbCrLf & "SELECT @vVar = [value] FROM ::fn_listextendedproperty( 01% , 02% , 03% , 04% , 05% , 06% , 07% )" & vbCrLf & "EXEC sp_addextendedproperty 11% , 12% , 13% , 14% , 15% , 16% , 17% , 18%"
    
        
'    oDb.Tables.Refresh True
    sSql = ReplaceMsg(kSqlList, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "NULL", "NULL"))
    
            
    vVarArray = GetSqlVarArray(oDb, sSql, lErr, sErr)
    
    'table properties
    If lErr = 0 And Not IsEmpty(vVarArray) Then
        For iRow = 0 To UBound(vVarArray, 2)
            Debug.Print vVarArray(0, iRow)
                
            sSql = ReplaceMsg(kSqlAdd, Array("11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%", "01%", "02%", "03%", "04%", "05%", "06%", "07%"), _
                Array(SqlFinalS(ApiciSQL(vVarArray(0, iRow))), _
                "@vVar", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sNewTable)), "NULL", "NULL", _
                SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "NULL", "NULL"))
                
            lErr = ErrExexSqlDirect(oDb, sSql, sErr)
            If lErr <> 0 Then Exit For
        Next
    End If
    
    'columns properties
    If lErr = 0 Then
        For iColumn = 1 To colTB.Count
            s2LevelObj = colTB.Item(iColumn).oName()
            sSql = ReplaceMsg(kSqlList, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'column'", SqlFinalS(ApiciSQL(s2LevelObj))))
            vVarArray = GetSqlVarArray(oDb, sSql, lErr, sErr)
            
            If lErr = 0 And Not IsEmpty(vVarArray) Then
                For iRow = 0 To UBound(vVarArray, 2)
                    Debug.Print vVarArray(0, iRow)
                                                
                    sSql = ReplaceMsg(kSqlAdd, Array("11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%", "01%", "02%", "03%", "04%", "05%", "06%", "07%"), _
                        Array(SqlFinalS(ApiciSQL(vVarArray(0, iRow))), _
                        "@vVar", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sNewTable)), "'column'", SqlFinalS(ApiciSQL(s2LevelObj)), _
                        SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'column'", SqlFinalS(ApiciSQL(s2LevelObj))))
                                                
                    lErr = ErrExexSqlDirect(oDb, sSql, sErr)
                    If lErr <> 0 Then Exit For
                Next
            End If
        Next
    End If
        
    ErrTransferTB_COL_Properties = lErr
    
End Function
Private Function GetSqlVarArray(ByVal oDb As SQLDMO.Database2, ByVal sSql As String, ByRef lErr As Long, ByRef sErr As String) As Variant

    Dim vVarArray As Variant
    Dim oQry As SQLDMO.QueryResults
    On Local Error Resume Next
    Set oQry = oDb.ExecuteWithResults(sSql, Len(sSql))
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
        
    vVarArray = vbEmpty
    If lErr = 0 And Not oQry Is Nothing Then
        vVarArray = Qry2Var(oQry)
    End If
    Set oQry = Nothing

    GetSqlVarArray = vVarArray
End Function
Private Function KeyAlreadyExists(ByVal oCol As Collection, ByVal sKeyName As String) As Boolean

    Dim bExists As Boolean
    Dim i As Integer
    
    For i = 1 To oCol.Count
        If StrComp(oCol.Item(i).oName(), sKeyName, vbTextCompare) = 0 Then
            bExists = True
            Exit For
        End If
    Next
    
    KeyAlreadyExists = bExists
End Function

Private Function ErrDumpProperties(ByVal b2TmpTable As Boolean, ByVal oDb As SQLDMO.Database2, ByVal sBaseTable As String, ByVal sOwner As String, ByVal sTmpTable As String, _
                                        ByVal colTrig As Collection, _
                                        ByVal colIdx As Collection, _
                                        ByVal colKey As Collection, _
                                        ByRef sErr As String) As Long

    Dim sSql As String
    Dim vVarArray As Variant
    Dim lErr As Long
    Dim iRow As Integer
    Dim sObjLev2 As String
    Dim iLoop As Integer
            
    Dim oBaseTb As SQLDMO.Table
    Dim sObjType As String
    Dim lTmpErr As Long, sTmpErr As String
    
    
        
    Const kSqlList2TMP As String = "SELECT [name] FROM ::fn_listextendedproperty( 1% , 2% , 3% , 4% , 5% , 6% , 7% )"
    Const kSqlAdd2TMP As String = "DECLARE @objName SYSNAME, @vVar SQL_VARIANT" & vbCrLf & "SELECT @objName= [name], @vVar = [value] FROM ::fn_listextendedproperty( 01% , 02% , 03% , 04% , 05% , 06% , 07% )" & vbCrLf & "INSERT [11%].[12%] VALUES( 13% , 14%, 15% , 16%)"
    
    Const kSqlList2Base As String = "SELECT [propName] FROM [01%].[02%] WHERE [objType] = '03%' AND [objName] = '04%'"
    Const kSqlAdd2Base As String = "DECLARE @vVar SQL_VARIANT" & vbCrLf & "SELECT @vVar = [propVal] FROM [01%].[02%] WHERE [objType] = '03%' AND [objName] = '04%' AND [propName] = '05%'" & vbCrLf & "EXEC sp_addextendedproperty 11% , 12% , 13% , 14% , 15% , 16% , 17% , 18%"
        
    
    If b2TmpTable Then
'*** Properties 2 TEMP Table
        Set oBaseTb = GetTbByName(objServer, oDb.Name, sBaseTable)
                
        sSql = "if exists (select * from dbo.sysobjects where id = object_id(N'[1%].[2%]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [1%].[2%]" & vbCrLf
        sSql = sSql & "CREATE TABLE [1%].[2%] ( objType VARCHAR(10), objName SYSNAME NOT NULL, propName SYSNAME NOT NULL, propVal SQL_VARIANT )"
        sSql = ReplaceMsg(sSql, Array("1%", "2%"), Array(sOwner, sTmpTable))
        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
        
        
        If lErr = 0 Then
'------ Trigger properties 2 TMP
            For iLoop = 1 To colTrig.Count
                sObjLev2 = SqlFinalS(ApiciSQL(colTrig.Item(iLoop).oName()))
                
                sSql = ReplaceMsg(kSqlList2TMP, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'trigger'", sObjLev2))
                vVarArray = GetSqlVarArray(oDb, sSql, lErr, sErr)
                If lErr = 0 And Not IsEmpty(vVarArray) Then
                    For iRow = 0 To UBound(vVarArray, 2)
                        Debug.Print vVarArray(0, iRow)
                        
                        sSql = ReplaceMsg(kSqlAdd2TMP, Array("11%", "12%", "13%", "14%", "15%", "16%", "01%", "02%", "03%", "04%", "05%", "06%", "07%"), _
                            Array(sOwner, sTmpTable, "'trigger'", sObjLev2, "@objName", "@vVar", _
                            SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'trigger'", sObjLev2))
                            
                        
                        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
                        If lErr <> 0 Then Exit For
                    Next
                End If
                If lErr <> 0 Then Exit For
            Next
        End If
            
        If lErr = 0 Then
'------ Index properties 2 TMP
            For iLoop = 1 To colIdx.Count
                sObjLev2 = SqlFinalS(ApiciSQL(colIdx.Item(iLoop).oName))
                
                If colIdx.Item(iLoop).oIsPrimaryKey() Then
                    sObjType = "'constraint'"
                Else
                    sObjType = "'index'"
                End If
                                
                sSql = ReplaceMsg(kSqlList2TMP, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), sObjType, sObjLev2))
                vVarArray = GetSqlVarArray(oDb, sSql, lErr, sErr)
                If lErr = 0 And Not IsEmpty(vVarArray) Then
                    For iRow = 0 To UBound(vVarArray, 2)
                        Debug.Print vVarArray(0, iRow)
                        
                        sSql = ReplaceMsg(kSqlAdd2TMP, Array("11%", "12%", "13%", "14%", "15%", "16%", "01%", "02%", "03%", "04%", "05%", "06%", "07%"), _
                            Array(sOwner, sTmpTable, sObjType, sObjLev2, "@objName", "@vVar", _
                            SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), sObjType, sObjLev2))
                            
                        
                        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
                        If lErr <> 0 Then Exit For
                    Next
                End If
                If lErr <> 0 Then Exit For
            Next
        End If
           
        If lErr = 0 Then
'------ Keys properties 2 TMP
            For iLoop = 1 To colKey.Count
                sObjLev2 = SqlFinalS(ApiciSQL(colKey.Item(iLoop).oName))
                
                sSql = ReplaceMsg(kSqlList2TMP, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), Array("NULL", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'constraint'", sObjLev2))
                vVarArray = GetSqlVarArray(oDb, sSql, lErr, sErr)
                If lErr = 0 And Not IsEmpty(vVarArray) Then
                    For iRow = 0 To UBound(vVarArray, 2)
                        Debug.Print vVarArray(0, iRow)
                        
                        sSql = ReplaceMsg(kSqlAdd2TMP, Array("11%", "12%", "13%", "14%", "15%", "16%", "01%", "02%", "03%", "04%", "05%", "06%", "07%"), _
                            Array(sOwner, sTmpTable, "'constraint'", sObjLev2, "@objName", "@vVar", _
                            SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'constraint'", sObjLev2))
                                                
                        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
                        If lErr <> 0 Then Exit For
                    Next
                End If
                If lErr <> 0 Then Exit For
            Next
        End If
        
        Set oBaseTb = Nothing
    Else
'*** Properties 2 Base Table

'------ Trigger properties 2 Base Table
        For iLoop = 1 To colTrig.Count
            sObjLev2 = colTrig.Item(iLoop).oName

            sSql = ReplaceMsg(kSqlList2Base, Array("01%", "02%", "03%", "04%"), Array(sOwner, sTmpTable, "trigger", sObjLev2))
            vVarArray = GetSqlVarArray(oDb, sSql, lTmpErr, sTmpErr)
            If lTmpErr <> 0 Then
                sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                lErr = Err_Free
            End If
            
            If lTmpErr = 0 And Not IsEmpty(vVarArray) Then
                For iRow = 0 To UBound(vVarArray, 2)
                    Debug.Print vVarArray(0, iRow)
                                        
                    sSql = ReplaceMsg(kSqlAdd2Base, Array("01%", "02%", "03%", "04%", "05%", "11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%"), _
                        Array(sOwner, sTmpTable, "trigger", sObjLev2, vVarArray(0, iRow), _
                        SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "@vVar", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'trigger'", SqlFinalS(ApiciSQL(sObjLev2))))
                    
                    Debug.Print sSql
                    lTmpErr = ErrExexSqlDirect(oDb, sSql, sTmpErr)
                    
                    If lTmpErr <> 0 Then
                        sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                        lErr = Err_Free
                    End If
                    
                    'If lErr <> 0 Then Exit For
                Next
            End If
            'If lErr <> 0 Then Exit For
        Next
        
        'If lErr = 0 Then
'------ Index properties 2 Base Table
            For iLoop = 1 To colIdx.Count
                sObjLev2 = colIdx.Item(iLoop).oName
                
                If colIdx.Item(iLoop).oIsPrimaryKey() Then
                    sObjType = "constraint"
                Else
                    sObjType = "index"
                End If
    
                sSql = ReplaceMsg(kSqlList2Base, Array("01%", "02%", "03%", "04%"), Array(sOwner, sTmpTable, sObjType, sObjLev2))
                
                vVarArray = GetSqlVarArray(oDb, sSql, lTmpErr, sTmpErr)
                If lTmpErr <> 0 Then
                    sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                    lErr = Err_Free
                End If
                If lTmpErr = 0 And Not IsEmpty(vVarArray) Then
                    For iRow = 0 To UBound(vVarArray, 2)
                        Debug.Print vVarArray(0, iRow)
                                                                        
                        sSql = ReplaceMsg(kSqlAdd2Base, Array("01%", "02%", "03%", "04%", "05%", "11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%"), _
                            Array(sOwner, sTmpTable, sObjType, sObjLev2, vVarArray(0, iRow), _
                            SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "@vVar", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'" & sObjType & "'", SqlFinalS(ApiciSQL(sObjLev2))))
                        
                        lTmpErr = ErrExexSqlDirect(oDb, sSql, sTmpErr)
                        
                        If lTmpErr <> 0 Then
                            sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                            lErr = Err_Free
                        End If
                        'If lErr <> 0 Then Exit For
                    Next
                End If
                'If lErr <> 0 Then Exit For
            Next
        'End If
    
        'If lErr = 0 Then
'------ Key properties 2 Base Table
            For iLoop = 1 To colKey.Count
                sObjLev2 = colKey.Item(iLoop).oName
    
                sSql = ReplaceMsg(kSqlList2Base, Array("01%", "02%", "03%", "04%"), Array(sOwner, sTmpTable, "constraint", sObjLev2))

                
                vVarArray = GetSqlVarArray(oDb, sSql, lTmpErr, sTmpErr)
                If lTmpErr <> 0 Then
                    sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                    lErr = Err_Free
                End If
                If lTmpErr = 0 And Not IsEmpty(vVarArray) Then
                    For iRow = 0 To UBound(vVarArray, 2)
                        Debug.Print vVarArray(0, iRow)
                                                
                        sSql = ReplaceMsg(kSqlAdd2Base, Array("01%", "02%", "03%", "04%", "05%", "11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%"), _
                            Array(sOwner, sTmpTable, "constraint", sObjLev2, vVarArray(0, iRow), _
                            SqlFinalS(ApiciSQL(vVarArray(0, iRow))), "@vVar", "'user'", SqlFinalS(ApiciSQL(sOwner)), "'table'", SqlFinalS(ApiciSQL(sBaseTable)), "'constraint'", SqlFinalS(ApiciSQL(sObjLev2))))
                        
                        lTmpErr = ErrExexSqlDirect(oDb, sSql, sTmpErr)
                        
                        If lTmpErr <> 0 Then
                            sErr = IIf(Len(sErr) <> 0, vbCrLf & String$(10, "-") & vbCrLf, "") & sTmpErr
                            lErr = Err_Free
                        End If
                        'If lErr <> 0 Then Exit For
                    Next
                End If
                'If lErr <> 0 Then Exit For
            Next
        'End If
        sErr = sTmpErr
    
    End If
            
    ErrDumpProperties = lErr
    
End Function

Private Function ErrSaveRefKey(ByRef oDb As SQLDMO.Database2, ByVal oObj As cIDX, ByRef sErr As String) As Long

    Dim sSql As String
    Dim i As Integer

    
    Dim sColumns As String, sRefColumns As String
    Dim sBuf As String
    Dim lErr As Long
    
    Dim sTables(1) As String
    Dim sOwner(1) As String
    Dim sKeyName As String
    Const kAddKey As String = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT [3%]" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    Const kNoCheckKey As String = "ALTER TABLE [1%].[2%] NOCHECK CONSTRAINT [3%]"

    'sSql = "ALTER TABLE [1%].[2%]" & vbCrLf & vbTab & "ADD CONSTRAINT 3%" & vbCrLf & vbTab & vbTab & "Foreign Key" & vbCrLf & vbTab & vbTab & vbTab & "( 4% )" & vbCrLf & vbTab & vbTab & vbTab & "References [5%].[6%] ( 7% )"
    '        [ ON DELETE { CASCADE | NO ACTION } ]
    '        [ ON UPDATE { CASCADE | NO ACTION } ]
    '        [ NOT FOR REPLICATION ]

    
    sRefColumns = ""
    sColumns = ""
        
    With oObj
        sKeyName = .oName
        For i = 0 To .oFldCountEx()
            sBuf = Trim$(.GetFldEx(i))
            If Len(sBuf) <> 0 Then
            If Len(sRefColumns) <> 0 Then sRefColumns = sRefColumns & " , "
            sRefColumns = sRefColumns & "[" & sBuf & "]"
            End If
        Next
        If Len(sRefColumns) = 0 Then lErr = vbObjectError + 2001
        
        If lErr = 0 Then
            For i = 0 To .oFldCount()
                sBuf = Trim$(.GetFld(i, False))
                If Len(sBuf) <> 0 Then
                    If Len(sColumns) <> 0 Then sColumns = sColumns & " , "
                    sColumns = sColumns & "[" & sBuf & "]"
                End If
            Next
            If Len(sColumns) = 0 Then lErr = vbObjectError + 2001
        End If
    End With
    
    If lErr = 0 Then
        sTables(0) = Replace(oObj.relTable, "[", "")
        sTables(0) = Replace(sTables(0), "]", "")
        i = InStr(sTables(0), ".")
        If i <> 0 Then sTables(0) = Trim$(Mid$(sTables(0), i + 1))
        sOwner(0) = GetObjectOwner(m_sDbName, sTables(0), am_OwnTable)
        
        sTables(1) = Replace(oObj.relReferencedTable, "[", "")
        sTables(1) = Replace(sTables(1), "]", "")
        i = InStr(sTables(1), ".")
        If i <> 0 Then sTables(1) = Trim$(Mid$(sTables(1), i + 1))
        sOwner(1) = GetObjectOwner(m_sDbName, sTables(1), am_OwnTable)
        If Len(sOwner(0)) = 0 Or Len(sOwner(1)) = 0 Then lErr = vbObjectError + 2002
    End If
    
    If lErr = 0 Then
        sSql = ReplaceMsg(kAddKey, Array("1%", "2%", "3%", "4%", "5%", "6%", "7%"), _
                Array(sOwner(0), sTables(0), sKeyName, sColumns, _
                sOwner(1), sTables(1), sRefColumns))
                
        Debug.Print sSql
        If oObj.UpdateCascade Then
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON UPDATE CASCADE"
        End If
        If oObj.DeleteCascade Then
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "ON DELETE CASCADE"
        End If
        If oObj.oExcludeReplication Then
            sSql = sSql & vbCrLf & vbTab & vbTab & vbTab & "NOT FOR REPLICATION"
        End If
        
        Debug.Print sSql
    End If
                
    If lErr = 0 Then
        lErr = ErrExexSqlDirect(oDb, sSql, sErr)
    End If
    If lErr = 0 Then
        If Not oObj.oChecked Then
            sSql = ReplaceMsg(kNoCheckKey, Array("1%", "2%", "3%"), Array(sOwner(0), sTables(0), sKeyName))
            lErr = ErrExexSqlDirect(oDb, sSql, sErr)
        End If
    End If
    
    If lErr <> 0 Then
        If lErr = vbObjectError + 2001 Then
            sErr = MyLoadResString(k_Columns_not_defined)
        ElseIf lErr = vbObjectError + 2002 Then
            sErr = MyLoadResString(k_Tables_not_defined)
        End If
    End If
    ErrSaveRefKey = lErr
    
End Function


