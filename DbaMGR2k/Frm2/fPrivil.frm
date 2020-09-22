VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fPrivil 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7350
   ClipControls    =   0   'False
   Icon            =   "fPrivil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "col"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   2
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   2400
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":000C
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":05CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":0B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":0C60
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":0D54
            Key             =   "Sp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPrivil.frx":0E60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Regen"
      Height          =   360
      Index           =   1
      Left            =   4665
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.OptionButton opt 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   6135
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "fPrivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iObjType As optAction
Private m_sDbName As String
Private sOName As String
Private oObj As Object
Private oDb As SQLDMO.Database2
Private m_bLoading As Boolean
Private miRow As Integer
Private oColl As Collection
Private oCollBase As Collection

Public Sub SetObj(oPasDB As SQLDMO.Database2, oPasObj As Object, ByVal sName As String, iType As optAction)

    iObjType = iType
    Set oDb = oPasDB
    Set oObj = oPasObj
    sOName = sName
    Dim sBuf As String, sBuf2 As String
    
    m_sDbName = oPasDB.Name
    Select Case iObjType
        Case act_Roles
            sBuf = MyLoadResString(k_List_all_objects)
            sBuf2 = MyLoadResString(k_List_only_Objects_with_Permissions_for_this_Role)
        Case act_User
            sBuf = MyLoadResString(k_List_all_objects)
            sBuf2 = MyLoadResString(k_List_only_Objects_with_Permissions_for_this_User)
        Case act_DBOption
            opt(0).Visible = False
            opt(1).Visible = False
            With Flex
                .Move .Left, opt(0).Top, .Width, cmd(0).Top - .Top - k_FlexBevel
            End With
        
        Case Else
            sBuf = MyLoadResString(k_List_all_Users_user_defined_DB_Roles_public)
            sBuf2 = MyLoadResString(k_List_only_Users_user_defined_DB_Roles_public_with_permission_on_this_object)
        
    End Select
    opt(0).Caption = sBuf
    opt(1).Caption = sBuf2
    
    If iObjType = act_DBOption Then
        cmd(3).Visible = False
        LoadDBPermission
    Else
        LoadElem
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If cmd(Index).Enabled Then
        Select Case Index
            Case 0
                Unload Me
            Case 1
                If iObjType = act_DBOption Then
                    LoadDBPermission
                Else
                    ReconnectObj
                    LoadElem
                End If
                Flex.SetFocus
            Case 2
                If iObjType = act_DBOption Then
                    UpdateDB
                Else
                    Update
                End If
                Flex.SetFocus
            Case 3  'Columns
                If iObjType <> act_DBOption Then LoadColumns
        End Select
    End If
End Sub

Private Sub Flex_DblClick()
    Dim iMinCol As Integer
    With Flex
        If .MouseCol > 0 And (.MouseRow > 0) Then
            If iObjType = act_DBOption Then
                iMinCol = 1
            Else
                iMinCol = 2
            End If
            If .Row > 0 And .Col > iMinCol Then FlexEdit vbKeySpace
        End If
    End With

End Sub
Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iMinCol As Integer
    With Flex
        If .MouseCol > 0 And (.MouseRow > 0) Then
            If iObjType = act_DBOption Then
                iMinCol = 1
            Else
                iMinCol = 2
            End If
            
            If .Row > 0 And .Col > iMinCol And KeyCode = vbKeyF2 Then FlexEdit 0
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

    Dim iRow As Integer
    Dim bVal As Boolean
    If Not m_bLoading Then
        If iObjType <> act_DBOption Then
            iRow = Flex.Row
            If iRow <> 0 Then
                With oColl.Item(Flex.TextMatrix(iRow, 1))
                    If .oType = SQLDMOObj_UserTable Or .oType = SQLDMOObj_View Then bVal = True
                End With
            End If
            cmd(3).Enabled = bVal
            
        End If
    End If
       
End Sub

Private Sub Form_Load()
    CenterForm Me
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Refresh)
    cmd(2).Caption = MyLoadResString(k_Apply)
    cmd(3).Caption = MyLoadResString(k_PrivColumns)
    
End Sub
Private Sub LoadFlex()

    Dim i As Integer, l As Integer
    Dim x As Integer
    Dim iRet As iPrivType
    Dim idBitmap As Integer
    Dim bUnav(5) As Boolean
    Dim lOType As Long
    Dim iImg As Integer
    Dim bAdd As Boolean
    Dim lErr As Long, sErr As String

    m_bLoading = True
    With Flex
        .Redraw = False
        .Clear
        .FixedRows = 0
        .Cols = 9
        .Rows = 1
        .Clear
        If iObjType <> act_User And iObjType <> act_Roles Then
            .TextMatrix(0, 0) = MyLoadResString(k_Users_DB_Roles_public)
            .TextMatrix(0, 1) = MyLoadResString(k_Users_DB_Roles_public)
        Else
            .TextMatrix(0, 0) = MyLoadResString(k_Object)
            .TextMatrix(0, 1) = MyLoadResString(k_Object)
            .TextMatrix(0, 2) = MyLoadResString(k_Owner)
        End If
        .TextMatrix(0, 3) = "SELECT"
        .TextMatrix(0, 4) = "INSERT"
        .TextMatrix(0, 5) = "UPDATE"
        .TextMatrix(0, 6) = "DELETE"
        .TextMatrix(0, 7) = "EXEC"
        .TextMatrix(0, 8) = "DRI"
        
        .ColWidth(0) = 500
        
        If iObjType <> act_User And iObjType <> act_Roles Then
            '.ColWidth(1) = 1500
            .ColWidth(2) = 0
        Else
            '.ColWidth(1) = 1200
            .ColWidth(2) = 900
        End If
        
        .FixedAlignment(3) = flexAlignCenterCenter
        .FixedAlignment(4) = flexAlignCenterCenter
        .FixedAlignment(5) = flexAlignCenterCenter
        .FixedAlignment(6) = flexAlignCenterCenter
        .FixedAlignment(7) = flexAlignCenterCenter
        .FixedAlignment(8) = flexAlignCenterCenter
        
        If iObjType = act_Tbl Then
            bUnav(4) = True
        ElseIf iObjType = act_View Then
            bUnav(4) = True
            bUnav(5) = True
        ElseIf iObjType = act_SP Then
            bUnav(0) = True
            bUnav(1) = True
            bUnav(2) = True
            bUnav(3) = True
            bUnav(5) = True
        End If

        For l = 1 To oColl.Count
            If opt(1).Value Then
                bAdd = oColl.Item(l).IsSet()
            Else
                bAdd = True
            End If
            If bAdd Then
        
                .AddItem vbTab & oColl.Item(l).oName() & IIf(iObjType = act_User Or iObjType = act_Roles, vbTab & oColl.Item(l).oGroup(), "")
                lOType = oColl.Item(l).oType()
                
                'If iObjType = act_User Then
                If iObjType = act_User Or iObjType = act_Roles Then
                    Erase bUnav
                    If (lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable Then
                        bUnav(4) = True
                        iImg = 3
                    ElseIf (lOType And SQLDMOObj_View) = SQLDMOObj_View Then
                        bUnav(4) = True
                        bUnav(5) = True
                        iImg = 4
                    ElseIf (lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure Then
                        bUnav(0) = True
                        bUnav(1) = True
                        bUnav(2) = True
                        bUnav(3) = True
                        bUnav(5) = True
                        iImg = 5
                    ElseIf (lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction Then
                        Debug.Print oColl.Item(l).oName
                        bUnav(0) = Not oColl.Item(l).IsTableFunction()
                        bUnav(1) = bUnav(0)
                        bUnav(2) = bUnav(0)
                        bUnav(3) = bUnav(0)
                        bUnav(4) = Not bUnav(0)
                        bUnav(5) = False
                        iImg = 6
                    End If
                Else
                    If oColl.Item(l).isGroup() Then
                        iImg = 2
                    Else
                        iImg = 1
                    End If
                    If (lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction Then
                        Debug.Print oColl.Item(l).oName
                        bUnav(0) = Not oColl.Item(l).IsTableFunction()
                        bUnav(1) = bUnav(0)
                        bUnav(2) = bUnav(0)
                        bUnav(3) = bUnav(0)
                        bUnav(4) = Not bUnav(0)
                        bUnav(5) = False
                    End If
                End If
                .Row = .Rows - 1
                .Col = 0
                On Local Error Resume Next
                Set Flex.CellPicture = iList.ListImages(iImg).Picture
                .CellPictureAlignment = flexAlignCenterCenter
                For x = 3 To 8
                    .Col = x
                    iRet = oColl.Item(l).ItemVal(x - 3)
                    If iRet = Deny Then
                        idBitmap = k_ResChkDel
                    ElseIf iRet = Grant Then
                        idBitmap = k_ResChkON
                    ElseIf iRet = GrantCol Then
                        idBitmap = k_ResColGrant
                    ElseIf iRet = DenyCol Then
                        idBitmap = k_ResColDeny
                    ElseIf iRet = GrantDenyCol Then
                        idBitmap = k_ResColGrantDeny
                    Else
                        idBitmap = k_ResChkOFF
                    End If
                    
                    If bUnav(x - 3) = True Or iRet = Unavailable Then
                        .CellBackColor = vbInactiveBorder
                    Else
                        Set Flex.CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                    End If
                    lErr = Err.Number
                    sErr = Err.Description
                    If lErr <> 0 Then Exit For
                Next
                On Local Error GoTo 0
                If lErr <> 0 Then
                    .RemoveItem .Rows - 1
                    .RemoveItem .Rows - 1
                    sErr = sErr & vbCrLf & ReplaceMsg(MyLoadResString(k_Privil_OutOfMemory), Array("1%", "2%", "|"), Array(l - 1, oColl.Count - l - 2, vbCrLf))
                    Exit For
                End If
            End If
        Next
                    
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth

        If .Rows > 1 Then
            .FixedRows = 1
        Else
            For x = 0 To 8
                .Col = x
                .CellBackColor = vbInactiveBorder
                If x > 2 Then .CellAlignment = flexAlignCenterCenter
            Next
        End If
        .MergeRow(0) = True
        .MergeCells = flexMergeRestrictColumns
        .Redraw = True
        For i = 3 To 8
            .ColWidth(i) = 750
        Next
        SettaLargFlex Flex, 1, 0
        m_bLoading = False
        
        miRow = 0
        .TopRow = IIf(.Rows > 1, 1, 0)
        .Row = IIf(.Rows > 1, 1, 0)
        .Col = 1
    End With
    
    If lErr <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCollBase = Nothing
    Set oColl = Nothing
    Set oObj = Nothing
    Set oDb = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    LoadFlex
    Screen.MousePointer = vbDefault
End Sub
Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim iImg As Integer, iElem As Integer
    Dim iMinCol As Integer
    
    bEdit = True
    iElem = Flex.Col
    
    If iObjType = act_DBOption Then
        iMinCol = 2
    Else
        iMinCol = 3
    End If
        
    If iElem < iMinCol Or Flex.CellBackColor = vbInactiveBorder Then bEdit = False
        
    If bEdit Then
        If iObjType = act_DBOption Then
            iElem = iElem - 2
        Else
            iElem = iElem - 3
        End If

        With oColl.Item(Flex.TextMatrix(Flex.Row, 1))
            iImg = 1 + (.ItemVal(iElem))
            If iImg >= 3 Then iImg = 0
            .ItemVal(iElem) = iImg
            .IsChanged() = True
            cmd(2).Enabled = True
        End With
        
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

Private Sub Update()

    Dim i As Integer, l As Integer
    Dim oPriv As cPrivil
    Dim oOldPriv As cPrivil
    Dim lVal As iPrivType, lOldVal As iPrivType
    Dim lPriv As Long
    Dim oUserObj As Object
    Dim lErr As Long, sErr As String
    Dim bAreErrors As Boolean
    
    
    Screen.MousePointer = vbHourglass
    For i = 1 To oColl.Count
        Set oPriv = oColl.Item(i)
        Set oOldPriv = oCollBase.Item(i)
        Debug.Print oPriv.oName
        If oPriv.IsChanged() Then
            If iObjType = act_User Or iObjType = act_Roles Then
                Select Case oPriv.oType()
                    Case oTbl
                        'Set oUserObj = oDb.Tables(oPriv.oName(), oPriv.oGroup())
                        Set oUserObj = GetTbByName(objServer, oDb.Name, oPriv.oName())
                    Case oSP
                        'Set oUserObj = oDb.StoredProcedures(oPriv.oName(), oPriv.oGroup())
                        Set oUserObj = GetSpByName(objServer, oDb.Name, oPriv.oName())
                    Case oView
                        'Set oUserObj = oDb.Views(oPriv.oName(), oPriv.oGroup())
                        Set oUserObj = GetViewByName(objServer, oDb.Name, oPriv.oName())
                    Case oFunc
                        'Set oUserObj = oDb.UserDefinedFunctions(oPriv.oName(), oPriv.oGroup())
                        Set oUserObj = GetFuncByName(objServer, oDb.Name, oPriv.oName())
                End Select
            End If
        
            For l = 0 To 5
                lVal = oPriv.ItemVal(l)
                lOldVal = oOldPriv.ItemVal(l)
                If lVal <> lOldVal Then
                    Select Case l
                        Case 0
                            lPriv = SQLDMOPriv_Select
                        Case 1
                            lPriv = SQLDMOPriv_Insert
                        Case 2
                            lPriv = SQLDMOPriv_Update
                        Case 3
                            lPriv = SQLDMOPriv_Delete
                        Case 4
                            lPriv = SQLDMOPriv_Execute
                        Case 5
                            lPriv = SQLDMOPriv_References
                    End Select
                    'NotSet = 0
                    'Grant = 1
                    'Deny = 2

                    On Local Error Resume Next
                    If lVal = Grant Then
                        If iObjType <> act_User And iObjType <> act_Roles Then
                            oObj.Grant lPriv, oPriv.oName() ', , True
                        Else
                            oUserObj.Grant lPriv, sOName  ', , True
                        End If
                    ElseIf lVal = Deny Then
                        If iObjType <> act_User And iObjType <> act_Roles Then
                            oObj.Deny lPriv, oPriv.oName() ', , True
                        Else
                            oUserObj.Deny lPriv, sOName ', , True
                        End If
                    ElseIf lVal = NotSet Then
                        If iObjType <> act_User And iObjType <> act_Roles Then
                            oObj.Revoke lPriv, oPriv.oName()                            ',,true,true
                        Else
                            oUserObj.Revoke lPriv, sOName ',,true,true
                        End If
                    End If
                    lErr = Err.Number
                    sErr = sErr & IIf(Len(sErr) <> 0, vbCrLf, "") & lErr & " - " & Err.Description
                    If lErr <> 0 Then bAreErrors = True
                    On Local Error GoTo 0
                    
                End If
            Next
        End If
        
    Next
    Set oPriv = Nothing
    Set oOldPriv = Nothing
    
    ReconnectObj
    LoadElem
    Screen.MousePointer = vbDefault
    If bAreErrors Then MsgBox MyLoadResString(kMsgBoxError) & ": " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub
Private Sub LoadDBPermission()

    Dim oQry As SQLDMO.QueryResults
    Dim oList As SQLDMO.SQLObjectList
    Dim oPerm As SQLDMO.Permission

    Dim oRole As SQLDMO.DatabaseRole
    Dim oUser As SQLDMO.User
    Dim vVar As Variant
    Dim lPriv As Long, bVal As Boolean
    Dim i As Integer, l As Integer
    
    Screen.MousePointer = vbHourglass
    Set oColl = New Collection
    Set oCollBase = New Collection
    Dim oPriv As cPrivil
    Dim iColV As Integer, iRowV As Integer
    
    
    Set oQry = oDb.EnumUsers
    If Not oQry Is Nothing Then
        vVar = Qry2Var(oQry)
        If IsArray(vVar) Then
            iColV = UBound(vVar, 1)
            iRowV = UBound(vVar, 2)
                
    
    'UserName    nvarchar(29)    Database username.
    'GroupName   nvarchar(33)    Database role. One row is returned for each role of which the user is a member.
    'LoginName   nvarchar(45)    Login name.
    'DefDBName   nvarchar(13)    Database used by default when a connection is made using the listed login.
    'UserID  char(14)    System-generated database user identifier.
    'SUserID char(16)    System-generated login identifier.
                For i = 0 To iRowV
                    Debug.Print vVar(0, i); vVar(1, i); vVar(2, i)
                    'If Not vVar(1, i) = "db_owner" Then
                    If StrComp(vVar(0, i), "dbo", vbTextCompare) <> 0 Then
                        Debug.Print vVar(0, i)
                        Debug.Print vVar(1, i)
                        Set oPriv = New cPrivil
                        oPriv.oName() = vVar(0, i)
                        oPriv.oGroup() = vVar(1, i)
                        On Local Error Resume Next
                        oColl.Add oPriv, vVar(0, i)
                        On Local Error GoTo 0
                    End If
                 Next
            End If
        End If
        Set oQry = Nothing

        For Each oRole In oDb.DatabaseRoles
            If Not oRole.IsFixedRole Or oRole.Name = "public" Then
                Set oPriv = New cPrivil
                oPriv.oName() = oRole.Name
                oPriv.isGroup() = True
                oColl.Add oPriv, oRole.Name
            End If
        Next
        Set oRole = Nothing
    'End If
    'SQLDMOPriv_CreateTable
    'SQLDMOPriv_CreateView
    'SQLDMOPriv_CreateProcedure
    'SQLDMOPriv_CreateDefault
    'SQLDMOPriv_CreateRule
    'SQLDMOPriv_DumpDatabase
    'SQLDMOPriv_DumpTransaction
    
    'SQLDMOPriv_CreateDatabase
        
    Set oList = oDb.ListDatabasePermissions(SQLDMOPriv_AllDatabasePrivs)
    Debug.Print oList Is Nothing
    For Each oPerm In oList
        Debug.Print oPerm.Grantee
        On Local Error Resume Next
        Set oPriv = oColl.Item(oPerm.Grantee)
        If Err = 5 Then
            '// se User non caricato
            Set oUser = oDb.Users(oPerm.Grantee)
            Debug.Print oUser.Name
            Set oPriv = New cPrivil
            oPriv.oName() = oPerm.Grantee
            On Local Error Resume Next
            oColl.Add oPriv, oPerm.Grantee
            On Local Error GoTo 0
            Set oUser = Nothing
        
        End If
        On Local Error GoTo 0
        
       
       lPriv = oPerm.PrivilegeType
       bVal = oPerm.Granted
       
       If (lPriv And SQLDMOPriv_CreateTable) = SQLDMOPriv_CreateTable Then oPriv.ItemVal(0) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_CreateView) = SQLDMOPriv_CreateView Then oPriv.ItemVal(1) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_CreateProcedure) = SQLDMOPriv_CreateProcedure Then oPriv.ItemVal(2) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_CreateDefault) = SQLDMOPriv_CreateDefault Then oPriv.ItemVal(3) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_CreateRule) = SQLDMOPriv_CreateRule Then oPriv.ItemVal(4) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_CreateFunction) = SQLDMOPriv_CreateFunction Then oPriv.ItemVal(5) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_DumpDatabase) = SQLDMOPriv_DumpDatabase Then oPriv.ItemVal(6) = IIf(bVal, Grant, Deny)
       If (lPriv And SQLDMOPriv_DumpTransaction) = SQLDMOPriv_DumpTransaction Then oPriv.ItemVal(7) = IIf(bVal, Grant, Deny)
    Next
    Set oPerm = Nothing
    Set oList = Nothing
        
    For i = 1 To oColl.Count
        With oColl.Item(i)
            Set oPriv = New cPrivil
            oPriv.oName() = .oName()
            oPriv.oType() = odbType.oDb '.oType
            'oPriv.oGroup() = .oGroup()
            oPriv.isGroup() = .isGroup()
            For l = 0 To 7
                oPriv.ItemVal(l) = .ItemVal(l)
            Next
            oCollBase.Add oPriv, oPriv.oName()
        End With
    Next
    Set oPriv = Nothing
    cmd(2).Enabled = False

    LoadFlexDB
    Screen.MousePointer = vbDefault
End Sub

Private Sub LoadFlexDB()

    Dim i As Integer, l As Integer
    Dim x As Integer
    Dim iRet As iPrivType
    Dim idBitmap As Integer
    Dim iImg As Integer

    With Flex
        .Redraw = False
        .Clear
        .FixedRows = 0
        .Cols = 10
        .Rows = 1
        .Clear
        
        .TextMatrix(0, 0) = MyLoadResString(k_Users_DB_Roles_public)
        .TextMatrix(0, 1) = MyLoadResString(k_Users_DB_Roles_public)
                
        .TextMatrix(0, 2) = MyLoadResString(k_Priv_Create_Table)
        .TextMatrix(0, 3) = MyLoadResString(k_Priv_Create_View)
        .TextMatrix(0, 4) = MyLoadResString(k_Priv_Create_SP)
        .TextMatrix(0, 5) = MyLoadResString(k_Priv_Create_Default)
        .TextMatrix(0, 6) = MyLoadResString(k_Priv_Create_Rule)
        .TextMatrix(0, 7) = MyLoadResString(k_Priv_CreateFunction)
        .TextMatrix(0, 8) = MyLoadResString(k_Priv_Backup_DB)
        .TextMatrix(0, 9) = MyLoadResString(k_Priv_Backup_Log)
        
        .ColWidth(0) = 500
        .ColWidth(1) = 400

        For x = 2 To .Cols - 1
            .FixedAlignment(x) = flexAlignCenterCenter
        Next

        For l = 1 To oColl.Count
            .AddItem vbTab & oColl.Item(l).oName() '& IIf(iObjType = act_User Or iObjType = act_Roles, vbTab & oColl.Item(l).oGroup(), "")
                            
            .Row = .Rows - 1
            .Col = 0
            If oColl.Item(l).isGroup() Then
                iImg = 2
            Else
                iImg = 1
            End If

            Set Flex.CellPicture = iList.ListImages(iImg).Picture
            .CellPictureAlignment = flexAlignCenterCenter
            For x = 3 To .Cols
                .Col = x - 1
                iRet = oColl.Item(l).ItemVal(x - 3)
                If iRet = Deny Then
                    idBitmap = k_ResChkDel
                ElseIf iRet = Grant Then
                    idBitmap = k_ResChkON
                Else
                    idBitmap = k_ResChkOFF
                End If
                
                Set Flex.CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            Next
        Next
                    
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth

        If .Rows > 1 Then
            .FixedRows = 1
        Else
            For x = 0 To .Cols - 1
                .Col = x
                .CellBackColor = vbInactiveBorder
            Next
        End If
        .MergeRow(0) = True
        .MergeCells = flexMergeRestrictColumns
        .Redraw = True
        For i = 2 To .Cols - 1
            .ColWidth(i) = 1050
        Next
        SettaLargFlex Flex, 1, 1000
        
        miRow = 0
        .TopRow = IIf(.Rows > 1, 1, 0)
        .Row = IIf(.Rows > 1, 1, 0)
        .Col = 1
    End With
End Sub

Private Sub UpdateDB()

    Dim i As Integer, l As Integer
    Dim oPriv As cPrivil
    Dim oOldPriv As cPrivil
    Dim lVal As iPrivType, lOldVal As iPrivType
    Dim lPriv As SQLDMO.SQLDMO_PRIVILEGE_TYPE
    Dim lErr As Long, sErr As String
    Dim bAreErrors As Boolean
    
        
    Screen.MousePointer = vbHourglass
    For i = 1 To oColl.Count
        Set oPriv = oColl.Item(i)
        Set oOldPriv = oCollBase.Item(i)
        Debug.Print oPriv.oName
        If oPriv.IsChanged() Then
        
            For l = 0 To 7
                lVal = oPriv.ItemVal(l)
                lOldVal = oOldPriv.ItemVal(l)
                If lVal <> lOldVal Then
    'SQLDMOPriv_CreateTable
    'SQLDMOPriv_CreateView
    'SQLDMOPriv_CreateProcedure
    'SQLDMOPriv_CreateDefault
    'SQLDMOPriv_CreateRule
    'SQLDMOPriv_DumpDatabase
    'SQLDMOPriv_DumpTransaction
                
                    Select Case l
                        Case 0
                            lPriv = SQLDMOPriv_CreateTable
                        Case 1
                            lPriv = SQLDMOPriv_CreateView
                        Case 2
                            lPriv = SQLDMOPriv_CreateProcedure
                        Case 3
                            lPriv = SQLDMOPriv_CreateDefault
                        Case 4
                            lPriv = SQLDMOPriv_CreateRule
                        Case 5
                            lPriv = SQLDMOPriv_CreateFunction
                        Case 6
                            lPriv = SQLDMOPriv_DumpDatabase
                        Case 7
                            lPriv = SQLDMOPriv_DumpTransaction
                    End Select
                    'NotSet = 0
                    'Grant = 1
                    'Deny = 2

                    On Local Error Resume Next
                    If lVal = Grant Then
                        Call oDb.Grant(lPriv, oPriv.oName())
                    ElseIf lVal = Deny Then
                        Call oDb.Deny(lPriv, oPriv.oName())
                    ElseIf lVal = NotSet Then
                        Call oDb.Revoke(lPriv, oPriv.oName())
                    End If
                    lErr = Err.Number
                    sErr = sErr & IIf(Len(sErr) <> 0, vbCrLf, "") & lErr & " - " & Err.Description
                    If lErr <> 0 Then bAreErrors = True
                    On Local Error GoTo 0
                End If
            Next
        End If
        
    Next
    Set oPriv = Nothing
    Set oOldPriv = Nothing
    LoadDBPermission
    Screen.MousePointer = vbDefault
    
    If bAreErrors Then MsgBox MyLoadResString(kMsgBoxError) & ": " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub
Private Function GetTableFuncStatus(ByVal oDb As SQLDMO.Database2, ByVal sFunctionName As String) As Boolean

    Dim sCommand As String
    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    
    sCommand = Replace("select OBJECTPROPERTY (object_id ('1%'),'IsTableFunction'),OBJECTPROPERTY (object_id ('1%'),'IsInlineFunction')", "1%", sFunctionName)
    On Local Error Resume Next
    Set oQry = oDb.ExecuteWithResults(sCommand, Len(sCommand))
    On Local Error GoTo 0
    
    If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
    Set oQry = Nothing
    If IsArray(vVar) Then
        ' if (vVar(0,0)=1 or vVar(1,0)=1) then is table
        GetTableFuncStatus = CBool(vVar(0, 0)) Or CBool(vVar(1, 0))
    End If

End Function
Private Sub LoadColumns()

    Dim iRow As Integer
    Dim bVal As Boolean
    Dim bCancel As Boolean
    Dim sObjName As String
    Dim sUser As String
    
    iRow = Flex.Row
    If iRow <> 0 Then
        Debug.Print oColl.Item(iRow).oName
        With oColl.Item(Flex.TextMatrix(iRow, 1))
            If .oType = SQLDMOObj_UserTable Or .oType = SQLDMOObj_View Then
                If iObjType = act_Roles Or iObjType = act_User Then
                    sUser = sOName
                    sObjName = .oName
                Else
                    sUser = .oName
                    sObjName = sOName
                End If
            
                Dim oF_ColP As fPrivCol
                Set oF_ColP = New fPrivCol
                With oF_ColP
                    .Stuff oDb, sObjName, oColl.Item(iRow).oType = SQLDMOObj_View, sUser
                    .Show vbModal, Me
                    bCancel = oF_ColP.Cancelled()
                End With

                Unload oF_ColP
                Set oF_ColP = Nothing
                
                If Not bCancel Then
                    ReconnectObj
                    
                    If iObjType = act_User Or iObjType = act_Roles Then
                        sObjName = Flex.TextMatrix(iRow, 1)
                    Else
                        sObjName = Flex.TextMatrix(iRow, 1)
                    End If
                    ReloadPermission oColl.Item(sObjName)
                End If
            
                bVal = True
            End If
        End With
    End If
    cmd(3).Enabled = bVal
    
End Sub

Private Sub ReloadPermission(ByVal oPriv As cPrivil)

    Dim oTmpPriv As cPrivil
    Dim oList As SQLDMO.SQLObjectList
    Dim oList2 As SQLDMO.SQLObjectList
    Dim oPerm As SQLDMO.Permission2
    Dim sObjName As String
    Dim lPriv As Long, lOType As Long
    Dim bVal As Boolean
    Dim bTableViewColumnDone As Boolean
    Dim bUnav(5) As Boolean
    Dim i As Integer, iRet As Integer
    Dim idBitmap As Integer
    Dim sObj2TestName As String
    Dim bGO As Boolean
    Dim sUser As String
        
    Debug.Print oPriv.oName
    Screen.MousePointer = vbHourglass
    
    If iObjType = act_Roles Or iObjType = act_User Then
        sObj2TestName = oPriv.oName()
        sUser = sOName
    Else
        sUser = oPriv.oName()
        sObj2TestName = sOName
    End If
    
    sObjName = oPriv.oName()
    Set oTmpPriv = New cPrivil
    With oTmpPriv
        .oName() = sObjName
        .isGroup() = oPriv.isGroup()
        .IsTableFunction() = oPriv.IsTableFunction()
        .oType() = oPriv.oType()
    End With
    
    Set oList = oDb.ListObjectPermissions
        
    For Each oPerm In oList
        Debug.Print oPerm.Grantee, oPerm.ObjectName
        If StrComp(oPerm.ObjectName, sObj2TestName, vbTextCompare) = 0 And StrComp(oPerm.Grantee, sUser, vbTextCompare) = 0 Then
        
        'SQLDMOPriv_Select
        'SQLDMOPriv_Insert
        'SQLDMOPriv_Update
        'SQLDMOPriv_Delete
        'SQLDMOPriv_Execute
        'SQLDMOPriv_References
        
            Debug.Print oPerm.Grantee
            Debug.Print oPerm.ObjectName, oPerm.ObjectTypeName
            lOType = oPerm.ObjectType
            
            lPriv = oPerm.PrivilegeType
            bVal = oPerm.Granted
            Debug.Print oPerm.Grantee
            Debug.Print oPerm.ObjectName
            Debug.Print oPerm.PrivilegeTypeName
            If iObjType = act_User Or iObjType = act_Roles Then
                bGO = StrComp(oPerm.Grantee, sOName, vbTextCompare) = 0
            Else
                bGO = True
            End If
            
            If bGO Then
                bTableViewColumnDone = False
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Then
                    Set oList2 = oPerm.ListPrivilegeColumns
                    If oList2.Count Then
                        'Debug.Print oList2.Item(1)
                        If (lPriv And SQLDMOPriv_Select) = SQLDMOPriv_Select Then
                            oTmpPriv.ItemVal(0) = oTmpPriv.ItemVal(0) + IIf(bVal, GrantCol, DenyCol)
                        End If
                        If (lPriv And SQLDMOPriv_Update) = SQLDMOPriv_Update Then
                            oTmpPriv.ItemVal(2) = oTmpPriv.ItemVal(2) + IIf(bVal, GrantCol, DenyCol)
                        End If
                        bTableViewColumnDone = True
                    End If
                End If
                
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Then
                    If Not bTableViewColumnDone Then
                        If (lPriv And SQLDMOPriv_Select) = SQLDMOPriv_Select Then oTmpPriv.ItemVal(0) = IIf(bVal, Grant, Deny)
                    End If
                End If
    
                If (lPriv And SQLDMOPriv_Insert) = SQLDMOPriv_Insert Then oTmpPriv.ItemVal(1) = IIf(bVal, Grant, Deny)
    
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Then
                    If Not bTableViewColumnDone Then
                        If (lPriv And SQLDMOPriv_Update) = SQLDMOPriv_Update Then oTmpPriv.ItemVal(2) = IIf(bVal, Grant, Deny)
                    End If
                End If
    
                If (lPriv And SQLDMOPriv_Delete) = SQLDMOPriv_Delete Then oTmpPriv.ItemVal(3) = IIf(bVal, Grant, Deny)
                
                oTmpPriv.ItemVal(4) = Unavailable
                
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Then
                    oTmpPriv.ItemVal(5) = Unavailable
                Else
                    If (lPriv And SQLDMOPriv_References) = SQLDMOPriv_References Then oTmpPriv.ItemVal(5) = IIf(bVal, Grant, Deny)
                End If
            End If
        End If
    Next
    
    Set oPerm = Nothing
    Set oList = Nothing
    Set oList2 = Nothing
    
    With oCollBase.Item(sObjName)
        .ResetSet
        For i = 0 To 7
            .ItemVal(i) = oTmpPriv.ItemVal(i)
        Next
    End With
    With oColl.Item(sObjName)
        .ResetSet
        For i = 0 To 7
            .ItemVal(i) = oTmpPriv.ItemVal(i)
        Next
    End With
            
    Set oTmpPriv = Nothing
    
    With Flex
        m_bLoading = True
        .Redraw = False
        bUnav(4) = True
        If oColl.Item(sObjName).oType = oView Then bUnav(5) = True

        For i = 3 To 8
            .Col = i
            iRet = oColl.Item(sObjName).ItemVal(i - 3)
            If iRet = Deny Then
                idBitmap = k_ResChkDel
            ElseIf iRet = Grant Then
                idBitmap = k_ResChkON
            ElseIf iRet = GrantCol Then
                idBitmap = k_ResColGrant
            ElseIf iRet = DenyCol Then
                idBitmap = k_ResColDeny
            ElseIf iRet = GrantDenyCol Then
                idBitmap = k_ResColGrantDeny
            Else
                idBitmap = k_ResChkOFF
            End If
            
            If bUnav(i - 3) = True Or iRet = Unavailable Then
                .CellBackColor = vbInactiveBorder
            Else
                Set Flex.CellPicture = LoadResPicture(idBitmap, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            End If
        Next
        .Redraw = True
        m_bLoading = False
    End With
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub LoadElem()

    Dim oQry As SQLDMO.QueryResults
    Dim oList As SQLDMO.SQLObjectList
    Dim oList2 As SQLDMO.SQLObjectList
    Dim oPerm As SQLDMO.Permission2
    Dim oDBobj As SQLDMO.DBObject
    Dim oRole As SQLDMO.DatabaseRole
    Dim oUser As SQLDMO.User
    Dim vVar As Variant
    Dim lPriv As Long, bVal As Boolean
    Dim i As Integer, l As Integer
    Dim lOType As Long
    Dim bGO As Boolean
    Dim bIsTableFunction As Boolean
    Dim bTableViewColumnDone As Boolean
    
    Screen.MousePointer = vbHourglass
    Set oColl = New Collection
    Set oCollBase = New Collection
    Dim oPriv As cPrivil
    Dim iColV As Integer, iRowV As Integer
    
    
    If iObjType <> act_User And iObjType <> act_Roles Then
        Set oQry = oDb.EnumUsers
        If Not oQry Is Nothing Then
            vVar = Qry2Var(oQry)
            If IsArray(vVar) Then
                iColV = UBound(vVar, 1)
                iRowV = UBound(vVar, 2)
                
    
    'UserName    nvarchar(29)    Database username.
    'GroupName   nvarchar(33)    Database role. One row is returned for each role of which the user is a member.
    'LoginName   nvarchar(45)    Login name.
    'DefDBName   nvarchar(13)    Database used by default when a connection is made using the listed login.
    'UserID  char(14)    System-generated database user identifier.
    'SUserID char(16)    System-generated login identifier.
                For i = 0 To iRowV
                    Debug.Print vVar(0, i); vVar(1, i); vVar(2, i)
                    If StrComp(vVar(0, i), "dbo", vbTextCompare) <> 0 Then
                        Debug.Print vVar(0, i)
                        Debug.Print vVar(1, i)
                        Set oPriv = New cPrivil
                        oPriv.oName() = vVar(0, i)
                        oPriv.oGroup() = vVar(1, i)
                        
                        If iObjType = act_View Then
                            oPriv.oType() = oView
                        ElseIf iObjType = act_SP Then
                            oPriv.oType() = oSP
                        ElseIf iObjType = act_User_Function Then
                            oPriv.oType() = oFunc
                            oPriv.IsTableFunction() = GetTableFuncStatus(oDb, sOName)
                        ElseIf iObjType = act_Tbl Then
                            oPriv.oType() = oTbl
                        End If
                        
                        On Local Error Resume Next
                        oColl.Add oPriv, vVar(0, i)
                        On Local Error GoTo 0
                    End If
                 Next
            End If
        End If
        Set oQry = Nothing

        For Each oRole In oDb.DatabaseRoles
            If Not oRole.IsFixedRole Or oRole.Name = "public" Then
                Set oPriv = New cPrivil
                oPriv.oName() = oRole.Name
                oPriv.isGroup() = True
                If iObjType = act_View Then
                    oPriv.oType() = oView
                ElseIf iObjType = act_SP Then
                    oPriv.oType() = oSP
                ElseIf iObjType = act_User_Function Then
                    oPriv.oType() = oFunc
                    oPriv.IsTableFunction() = GetTableFuncStatus(oDb, sOName)
                    
                ElseIf iObjType = act_Tbl Then
                    oPriv.oType() = oTbl
                End If
                oColl.Add oPriv, oRole.Name
            End If
        Next
        Set oRole = Nothing
    Else
        Set oList = oDb.ListObjects(SQLDMOObj_StoredProcedure Or SQLDMOObj_UserTable Or SQLDMOObj_View Or SQLDMOObj_UserDefinedFunction)
        For Each oDBobj In oList
            'If Not oDBobj.SystemObject Then
                Debug.Print oDBobj.Name; oDBobj.Type 'SQLDMOObj_UserTable,SQLDMOObj_View  ,SQLDMOObj_StoredProcedure, SQLDMOObj_UserDefinedFunction
                Set oPriv = New cPrivil
                oPriv.oName() = oDBobj.Name()
                oPriv.oGroup() = oDBobj.Owner
                oPriv.oType() = oDBobj.Type
                If (oPriv.oType() And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction Then
                    bIsTableFunction = GetTableFuncStatus(oDb, oDBobj.Name())
                    oPriv.IsTableFunction() = bIsTableFunction
                End If
                
                oColl.Add oPriv, oDBobj.Name()
            'End If
        Next
        Set oDBobj = Nothing
        Set oList = Nothing
        
    End If
    
    Set oList = oDb.ListObjectPermissions
    
    For Each oPerm In oList
        Debug.Print oPerm.Grantee, oPerm.ObjectName
        
        On Local Error Resume Next
        If iObjType = act_Roles Then
            Set oPriv = oColl.Item(oPerm.ObjectName)
            Debug.Print oPerm.ObjectType
            bGO = Err.Number = 0
        ElseIf iObjType <> act_User Then
            If StrComp(oPerm.ObjectName, sOName, vbTextCompare) = 0 Then
            
                Set oPriv = oColl.Item(oPerm.Grantee)
                If Err = 5 Then
                    '// se User non caricato
                    Set oUser = oDb.Users(oPerm.Grantee)
                    Debug.Print oUser.Name
                    Set oPriv = New cPrivil
                    oPriv.oName() = oPerm.Grantee
                    
                    If iObjType = act_View Then
                        oPriv.oType() = oView
                    ElseIf iObjType = act_SP Then
                        oPriv.oType() = oSP
                    ElseIf iObjType = act_User_Function Then
                        oPriv.oType() = oFunc
                    ElseIf iObjType = act_Tbl Then
                        oPriv.oType() = oTbl
                    End If
                    
                    On Local Error Resume Next
                    oColl.Add oPriv, oPerm.Grantee
                    On Local Error GoTo 0
                    Set oUser = Nothing
                
                End If
            End If
        Else
            Set oPriv = oColl.Item(oPerm.ObjectName)
            bGO = Err.Number = 0
        End If
        On Local Error GoTo 0
        'SQLDMOPriv_Select
        'SQLDMOPriv_Insert
        'SQLDMOPriv_Update
        'SQLDMOPriv_Delete
        'SQLDMOPriv_Execute
        'SQLDMOPriv_References
        Debug.Print oPerm.Grantee, oPriv.oName, oPerm.PrivilegeType
        If iObjType = act_User Or iObjType = act_Roles Then
            If bGO Then bGO = oPerm.Grantee = oObj.Name
        Else
            bGO = StrComp(oPerm.ObjectName, sOName, vbTextCompare) = 0
        End If
        
        
        If bGO Then
            bTableViewColumnDone = False
            Debug.Print oPriv.oName
            Debug.Print oPerm.ObjectName, oPerm.ObjectTypeName
            lOType = oPerm.ObjectType
            If StrComp(oPerm.ObjectTypeName, "UserDefinedFunction", vbTextCompare) = 0 Then
                lOType = SQLDMOObj_UserDefinedFunction
            End If
            
            lPriv = oPerm.PrivilegeType
            bVal = oPerm.Granted
            
            If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Then
                Set oList2 = oPerm.ListPrivilegeColumns
                If oList2.Count Then
                    'Debug.Print oList2.Item(1)
                    If (lPriv And SQLDMOPriv_Select) = SQLDMOPriv_Select Then
                        oPriv.ItemVal(0) = oPriv.ItemVal(0) + IIf(bVal, GrantCol, DenyCol)
                    End If
                    If (lPriv And SQLDMOPriv_Update) = SQLDMOPriv_Update Then
                        oPriv.ItemVal(2) = oPriv.ItemVal(2) + IIf(bVal, GrantCol, DenyCol)
                    End If
                    bTableViewColumnDone = True
                End If
            End If
            
            
            If iObjType <> act_User Then
                oPriv.oType() = oPerm.ObjectType
            End If
            
            If (lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction Then
                bIsTableFunction = oPriv.IsTableFunction()
            End If
            

            If (lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure Or ((lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction And Not bIsTableFunction) Then
                oPriv.ItemVal(0) = Unavailable
            Else
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Or bIsTableFunction Then
                    If (lPriv And SQLDMOPriv_Select) = SQLDMOPriv_Select Then
                        If Not bTableViewColumnDone Then oPriv.ItemVal(0) = IIf(bVal, Grant, Deny)
                    End If
                End If
            End If
            If (lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure Or ((lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction And Not bIsTableFunction) Then
                oPriv.ItemVal(1) = Unavailable
            Else
                If (lPriv And SQLDMOPriv_Insert) = SQLDMOPriv_Insert Then oPriv.ItemVal(1) = IIf(bVal, Grant, Deny)
            End If
            If (lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure Or ((lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction And Not bIsTableFunction) Then
                oPriv.ItemVal(2) = Unavailable
            Else
                If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Or bIsTableFunction Then
                    If (lPriv And SQLDMOPriv_Update) = SQLDMOPriv_Update Then
                        If Not bTableViewColumnDone Then oPriv.ItemVal(2) = IIf(bVal, Grant, Deny)
                    End If
                End If
            End If
            If (lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure Or ((lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction And Not bIsTableFunction) Then
                oPriv.ItemVal(3) = Unavailable
            Else
                If (lPriv And SQLDMOPriv_Delete) = SQLDMOPriv_Delete Then oPriv.ItemVal(3) = IIf(bVal, Grant, Deny)
            End If
            
            If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_UserTable) = SQLDMOObj_UserTable) Or ((lOType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction And bIsTableFunction) Then
                oPriv.ItemVal(4) = Unavailable
            Else
                If (lPriv And SQLDMOPriv_Execute) = SQLDMOPriv_Execute Then oPriv.ItemVal(4) = IIf(bVal, Grant, Deny)
            End If
            
            If ((lOType And SQLDMOObj_View) = SQLDMOObj_View) Or ((lOType And SQLDMOObj_StoredProcedure) = SQLDMOObj_StoredProcedure) Then
                oPriv.ItemVal(5) = Unavailable
            Else
                If (lPriv And SQLDMOPriv_References) = SQLDMOPriv_References Then oPriv.ItemVal(5) = IIf(bVal, Grant, Deny)
            End If
        End If
    Next
    
    Set oPerm = Nothing
    Set oList2 = Nothing
    Set oList = Nothing
        
    For i = 1 To oColl.Count
        With oColl.Item(i)
            Set oPriv = New cPrivil
            oPriv.oName() = .oName()
            oPriv.oType() = .oType
            oPriv.oGroup() = .oGroup()
            oPriv.isGroup() = .isGroup()
            Debug.Print .oType
            If (.oType And SQLDMOObj_UserDefinedFunction) = SQLDMOObj_UserDefinedFunction Then
                If iObjType = act_User_Function Then .IsTableFunction() = GetTableFuncStatus(oDb, sOName)
                oPriv.IsTableFunction() = .IsTableFunction()
            End If
            
            For l = 0 To 5
                oPriv.ItemVal(l) = .ItemVal(l)
            Next
            oCollBase.Add oPriv, oPriv.oName()
        End With
    Next
    Set oPriv = Nothing
    cmd(2).Enabled = False

    LoadFlex
    Screen.MousePointer = vbDefault
End Sub

Private Sub ReconnectObj()

    objServer.Databases.Refresh True
    Set oDb = objServer.Databases(m_sDbName)
    If iObjType = act_SP Then
        Set oObj = GetSpByName(objServer, m_sDbName, sOName)
    ElseIf iObjType = act_Tbl Then
        Set oObj = GetTbByName(objServer, m_sDbName, sOName)
    ElseIf iObjType = act_View Then
        Set oObj = GetViewByName(objServer, m_sDbName, sOName)
    ElseIf iObjType = act_User_Function Then
        Set oObj = GetFuncByName(objServer, m_sDbName, sOName)
    ElseIf iObjType = act_User Then
        Set oObj = objServer.Databases(m_sDbName).Users(sOName)
    ElseIf iObjType = act_Roles Then
        Set oObj = objServer.Databases(m_sDbName).DatabaseRoles(sOName)
    End If

End Sub
