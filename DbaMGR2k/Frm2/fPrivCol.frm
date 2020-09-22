VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fPrivCol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2895
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "fPrivCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum am_GrantDeny
    am_Undeclared = 0
    am_Declared_Grant
    am_Declared_Deny
End Enum

Private Type am_colPriv
    iPtr As Integer
    colName As String
    iPrivil(1) As am_GrantDeny
    iPrivilOrig(1) As am_GrantDeny
    IsChaged As Boolean
End Type

Private miRow As Integer
Private m_bLoading As Boolean
Private m_oDb As SQLDMO.Database2
Private m_sObjName As String
Private m_sObjOwner As String
Private m_bIsView As Boolean
Private m_Elem() As am_colPriv

Private m_bExec As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = Not m_bExec
End Property

Private Sub chk_Click()
    LoadFlex
End Sub

Private Sub cmd_Click(Index As Integer)
    If cmd(Index).Enabled Then
        If Index = 0 Then
            Me.Hide
        Else
            Save
        End If
    End If
End Sub

Private Sub Flex_DblClick()
    
    With Flex
        If .MouseCol > 0 And (.MouseRow > 0) Then
            If .Row > 0 And .Col > 1 Then FlexEdit vbKeySpace
        End If
    End With

End Sub
Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With Flex
        If .MouseCol > 0 And (.MouseRow > 0) Then
            If .Row > 0 And .Col > 1 And KeyCode = vbKeyF2 Then FlexEdit 0
        End If
    End With

End Sub
Private Sub Flex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        FlexEdit vbKeySpace
    End If

End Sub
Private Sub Flex_RowColChange()
    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
    
    If Not m_bLoading Then
        If Not b Then
            b = True
            
            With Flex
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex.CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1

                    .Col = iCol
                    miRow = iRowP
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
        SettaLargFlex Flex, 1, 200
    End If
    
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Caption = MyLoadResString(k_Priv_Col_Autorization)
    chk.Caption = MyLoadResString(k_Priv_Col_OnlyAut_Col)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    Label1(0).Caption = MyLoadResString(k_Priv_User_Name)
    Label1(1).Caption = MyLoadResString(k_Priv_Object_Name)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fPrivCol = Nothing
End Sub

Public Sub Stuff(ByVal oDb As SQLDMO.Database2, ByVal sObjName As String, ByVal bIsView As Boolean, ByVal sUserName As String)
        
    Dim oObj As Object
    
    m_sObjName = sObjName
    m_bIsView = bIsView
    
    Screen.MousePointer = vbHourglass
    
    Set m_oDb = oDb
    m_oDb.Tables.Refresh True
    m_oDb.Views.Refresh True
    If bIsView Then
        Set oObj = GetViewByName(objServer, m_oDb.Name, sObjName)
    Else
        Set oObj = GetTbByName(objServer, m_oDb.Name, sObjName)
    End If
    m_sObjOwner = oObj.Owner
    txt(0).Text = sUserName
    txt(1).Text = oObj.Owner & "." & oObj.Name
    LoadObjectCol oObj
    myLoadPermission oObj
    LoadFlex
    
    Set oObj = Nothing
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub LoadObjectCol(ByVal oObj As Object)

    Dim i As Integer, iCount As Integer
    Dim oCol As SQLDMO.Column
        
    i = i - 1
    If m_bIsView Then
        Dim oList As SQLDMO.SQLObjectList
        Set oList = oObj.ListColumns
        iCount = oList.Count - 1
    Else
        iCount = oObj.Columns.Count - 1
    End If
    
    ReDim m_Elem(iCount) As am_colPriv
    
    If m_bIsView Then
        For Each oCol In oList
            i = i + 1
            m_Elem(i).colName = oCol.Name
        Next
    Else
        For Each oCol In oObj.Columns
            i = i + 1
            m_Elem(i).colName = oCol.Name
        Next
    End If
    Set oList = Nothing
    
End Sub
Private Sub myLoadPermission(ByVal oObj As Object)

    Dim oList As SQLDMO.SQLObjectList
    Dim oList2 As SQLDMO.SQLObjectList
    Dim oPerm As SQLDMO.Permission2
    Dim oCol As SQLDMO.Column
    Dim sUser As String
    
    Dim iElem As Integer
    Dim lPriv As Long
    Dim bGrant As Boolean
        
    Debug.Print oObj.Name
    sUser = txt(0).Text
    
    'Set oList = oObj.ListUserPermissions(txt(0).Text)
    'Set oList = oObj.ListPermissions(SQLDMOPriv_Select Or SQLDMOPriv_Update)
    'Set oList = oObj.ListPermissions(SQLDMOPriv_AllObjectPrivs)
    'Set oList = oObj.ListObjectPermissions
    '// had to switch to Database.ListObjectPermissions becouse the others did
    '// always return an empty OjectList
    Set oList = m_oDb.ListObjectPermissions
    
    For Each oPerm In oList
        Debug.Print oPerm.ObjectName, oPerm.Grantee
        If StrComp(oPerm.ObjectName, m_sObjName, vbTextCompare) = 0 And StrComp(oPerm.Grantee, sUser, vbTextCompare) = 0 Then
            Set oList2 = oPerm.ListPrivilegeColumns
            Debug.Print oList2.Count
            For Each oCol In oList2
                Debug.Print oCol.Name
                iElem = GetObjPos(oCol.Name)
                If iElem <> -1 Then
                    Debug.Print oCol.Name, oPerm.PrivilegeTypeName, oPerm.Granted
                    Debug.Print oPerm.GrantedGranted

                    lPriv = oPerm.PrivilegeType
                    bGrant = oPerm.Granted
                    
                    If (lPriv And SQLDMOPriv_Select) = SQLDMOPriv_Select Then
                        m_Elem(iElem).iPrivilOrig(0) = IIf(bGrant, am_Declared_Grant, am_Declared_Deny)
                    End If
                    If (lPriv And SQLDMOPriv_Update) = SQLDMOPriv_Update Then
                        m_Elem(iElem).iPrivilOrig(1) = IIf(bGrant, am_Declared_Grant, am_Declared_Deny)
                    End If
                End If
            Next
        End If
    Next
    For iElem = 0 To UBound(m_Elem)
        m_Elem(iElem).iPrivil(0) = m_Elem(iElem).iPrivilOrig(0)
        m_Elem(iElem).iPrivil(1) = m_Elem(iElem).iPrivilOrig(1)
    Next
    
    Set oList2 = Nothing
    Set oList = Nothing
    Set oPerm = Nothing

End Sub
Private Function GetObjPos(ByVal sObjName As String) As Integer
    Dim iRet As Integer
    Dim i As Integer
    
    iRet = -1
    For i = 0 To UBound(m_Elem)
        If StrComp(sObjName, m_Elem(i).colName, vbTextCompare) = 0 Then
            iRet = i
            Exit For
        End If
    Next
    GetObjPos = iRet
End Function
Private Sub LoadFlex()

    Dim i As Integer
    Dim iCounter As Integer
    Dim iImg As Integer
    Dim bAdd As Boolean

    m_bLoading = True
    With Flex
        .Redraw = False
        .Clear
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 4
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 1) = MyLoadResString(k_Column_Name)
        .TextMatrix(0, 2) = "SELECT"
        .TextMatrix(0, 3) = "UPDATE"
        
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        .ColWidth(1) = 500
        
        
        .FixedAlignment(1) = flexAlignLeftCenter
        
        iCounter = 0
        For i = 0 To UBound(m_Elem)
            m_Elem(i).iPtr = 0
        Next
        For i = 0 To UBound(m_Elem)
            If chk.Value Then
                bAdd = (m_Elem(i).iPrivil(0) + m_Elem(i).iPrivil(1)) <> 0
            Else
                bAdd = True
            End If
            If bAdd Then
                iCounter = iCounter + 1
                .AddItem vbTab & m_Elem(i).colName
                m_Elem(i).iPtr = iCounter           'Riga Flex
                                                
                .Row = .Rows - 1
                .Col = 2
                Select Case m_Elem(i).iPrivil(0)
                    Case am_Declared_Deny
                        iImg = k_ResChkDel
                    Case am_Declared_Grant
                        iImg = k_ResChkON
                    Case Else
                        iImg = k_ResChkOFF
                End Select
                Set Flex.CellPicture = LoadResPicture(iImg, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                .Col = 3
                Select Case m_Elem(i).iPrivil(1)
                    Case am_Declared_Deny
                        iImg = k_ResChkDel
                    Case am_Declared_Grant
                        iImg = k_ResChkON
                    Case Else
                        iImg = k_ResChkOFF
                End Select
                Set Flex.CellPicture = LoadResPicture(iImg, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            End If
        Next
                    
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth

        If .Rows > 1 Then
            .FixedRows = 1
        Else
            .Row = .Rows - 1
            For i = 0 To .Cols - 1
                .Col = i
                .CellBackColor = vbInactiveBorder
            Next
        End If
        .FixedCols = 2
        .Redraw = True
        SettaLargFlex Flex, 1, 200
        m_bLoading = False
        miRow = 0
        .TopRow = IIf(.Rows > 1, 1, 0)
        .Row = IIf(.Rows > 1, 1, 0)
        .Col = 1
    End With
    
End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim iImg As Integer, iCol As Integer
    Dim iRow As Integer
    Dim iElem As Integer
    
    bEdit = True
    iCol = Flex.Col
    iRow = Flex.Row
        
    If iCol < 2 Or iRow < 1 Or Flex.CellBackColor = vbInactiveBorder Then bEdit = False
    
    If KeyAscii <> vbKeySpace And KeyAscii <> 0 Then bEdit = False
            
    If bEdit Then
        iCol = iCol - 2
        iElem = GetElementByPtr(iRow)
        If iElem <> -1 Then
            iImg = 1 + (m_Elem(iElem).iPrivil(iCol))
            If iImg >= 3 Then iImg = 0
            m_Elem(iElem).iPrivil(iCol) = iImg
            m_Elem(iElem).IsChaged = True
            Debug.Print m_Elem(iElem).colName
            cmd(1).Enabled = True
        End If
        
        If iImg = Deny Then
            iImg = k_ResChkDel
        ElseIf iImg = Grant Then
            iImg = k_ResChkON
        Else
            iImg = k_ResChkOFF
        End If
        Set Flex.CellPicture = LoadResPicture(iImg, vbResBitmap)
    End If
    
End Sub

Private Function GetElementByPtr(ByVal iPtr As Integer) As Integer

    Dim iRet As Integer
    Dim i As Integer
    
    iRet = -1
    For i = 0 To UBound(m_Elem)
        If m_Elem(i).iPtr = iPtr Then
            iRet = i
            Exit For
        End If
    Next
    GetElementByPtr = iRet
End Function

Private Sub Save()

    Dim i As Integer
    Dim lErr As Long, sErr As String
    Dim sErrors As String
    Dim sCommand As String
    Const k_CommandGrant As String = "GRANT 1% ON [2%].[3%] ([4%]) TO [5%]"
    Const k_CommandDeny As String = "DENY 1% ON [2%].[3%] ([4%]) TO [5%] CASCADE"
    Const k_CommandRevoke As String = "REVOKE 1% ON [2%].[3%] ([4%]) TO [5%] CASCADE"
    Dim sPriv(1) As String
    Dim bDone As Boolean
    
    sPriv(0) = "SELECT"
    sPriv(1) = "UPDATE"
    
    Screen.MousePointer = vbHourglass
    
    For i = 0 To UBound(m_Elem)
        If m_Elem(i).IsChaged Then
            Debug.Print m_Elem(i).colName
            If m_Elem(i).iPrivilOrig(0) <> m_Elem(i).iPrivil(0) Then
                If m_Elem(i).iPrivil(0) = am_Declared_Grant Then
                    sCommand = ReplaceMsg(k_CommandGrant, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(0), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                ElseIf m_Elem(i).iPrivil(0) = am_Declared_Deny Then
                    sCommand = ReplaceMsg(k_CommandDeny, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(0), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                Else
                    sCommand = ReplaceMsg(k_CommandRevoke, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(0), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                End If
                lErr = ErrExecPrivil(sCommand, sErr)
                
                If lErr <> 0 Then sErrors = IIf(Len(sErrors) <> 0, sErrors & vbCrLf & String$(30, "-" & vbCrLf), "") & MyLoadResString(k_Priv_Error_Executing) & vbCrLf & sCommand & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
            End If
            
            If m_Elem(i).iPrivilOrig(1) <> m_Elem(i).iPrivil(1) Then
                If m_Elem(i).iPrivil(1) = am_Declared_Grant Then
                    sCommand = ReplaceMsg(k_CommandGrant, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(1), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                ElseIf m_Elem(i).iPrivil(1) = am_Declared_Deny Then
                    sCommand = ReplaceMsg(k_CommandDeny, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(1), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                Else
                    sCommand = ReplaceMsg(k_CommandRevoke, Array("1%", "2%", "3%", "4%", "5%"), Array(sPriv(1), m_sObjOwner, m_sObjName, m_Elem(i).colName, txt(0).Text))
                End If
                lErr = ErrExecPrivil(sCommand, sErr)
                
                If lErr <> 0 Then sErrors = IIf(Len(sErrors) <> 0, sErrors & vbCrLf & String$(30, "-" & vbCrLf), "") & sCommand & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
            End If
            bDone = True
        End If
    Next
    
    If Len(sErrors) <> 0 Then
        
        Dim fRes As fResult
        Set fRes = New fResult
        With fRes
            .Action() = act_Null
            '.WWrapVisible() = False
                                       
            .tRes.Text = sErrors
            .Caption = MyLoadResString(k_GenWarning)
            Beep
            Screen.MousePointer = vbDefault
            .Show vbModal, Me
        End With
        Set fRes = Nothing
        Screen.MousePointer = vbHourglass
    End If
    
    If bDone Then
        m_bExec = True
        Me.Stuff m_oDb, m_sObjName, m_bIsView, txt(0).Text
    End If
    cmd(1).Enabled = False
    Screen.MousePointer = vbDefault
    
End Sub
Private Function ErrExecPrivil(ByVal sCommand As String, ByRef sErr As String) As Long

    Dim lErr As Long
    Dim sResult As String
    Dim oQry As SQLDMO.QueryResults
    
    Debug.Print sCommand
    On Local Error Resume Next
    Set oQry = m_oDb.ExecuteWithResultsAndMessages(sCommand, Len(sCommand), sResult)
    lErr = Err.Number
    If lErr <> 0 Then
        sErr = Err.Description
        If Len(sResult) <> 0 Then sErr = sErr & vbCrLf & sResult
    End If
    On Local Error GoTo 0
    Set oQry = Nothing
    
    ErrExecPrivil = lErr
End Function
