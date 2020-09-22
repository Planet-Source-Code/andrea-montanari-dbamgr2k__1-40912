VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fLista 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   ClipControls    =   0   'False
   Icon            =   "fLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Dipendenze"
      Height          =   360
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Esci"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Height          =   360
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2835
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5001
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "fLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type udtCol
    colName As String
    ColWidth As Long
End Type

Private m_bExitFrom As Boolean
Private bConfirm As Boolean
Private iActionMode As optLstAction
Private sDbName As String
Private sCurUser As String
Private vntGlobal() As Variant
Private vntObj As Variant
Private bLoading As Boolean
Private miRow As Integer

Public Property Get Confirm() As Boolean
    Confirm = bConfirm
End Property

Private Sub cmd_Click(Index As Integer)
    If cmd(Index).Enabled Then
        Select Case Index
            Case 0
                Me.Hide
            Case 1
                bConfirm = True
                Me.Hide
            Case 2
                If iActionMode = EnumUserObj Then
                    If Flex.Rows <> 1 Then LoadChangingObj
                Else
                    If Flex.RowSel = Flex.Row Then ShowDependencies
                End If
        End Select
    End If
    
End Sub
Private Sub Flex_DblClick()
        
    Dim iImg As Integer
    Dim iRow As Integer
    If Flex.Col = 1 And Flex.Rows <> 1 And Flex.Row <> 0 Then
        
        If Flex.CellBackColor = vbInactiveBorder Then Exit Sub
        
        iRow = Flex.Row - 1
        iImg = 1 - (-vntObj(iRow, 3))
        vntObj(iRow, 3) = CBool(iImg)
        Set Flex.CellPicture = LoadResPicture(IIf(Abs(iImg), k_ResChkON, k_ResChkOFF), vbResBitmap)
    End If

End Sub
Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Flex.Rows <> 1 And Flex.Col = 1 Then
        If KeyCode = vbKeyF2 Then Flex_DblClick
    End If

End Sub
Private Sub Flex_KeyPress(KeyAscii As Integer)

    If Flex.Rows <> 1 And Flex.Col = 1 Then
        If KeyAscii = vbKeySpace Then
            KeyAscii = 0
            Flex_DblClick
        End If
    End If
    
End Sub
Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
    
    If Not bLoading And iActionMode = EnumUserObj And Flex.Rows <> 1 Then
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

Private Sub Form_Load()
   
   cmd(0).Caption = MyLoadResString(k_Cancel)
   cmd(1).Caption = "Ok"
   CenterForm Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set fLista = Nothing
End Sub

Public Sub Stuff_Exec(ByVal iVal As optLstAction, ByVal sDB As String, vntObjects As Variant)
    
    Dim iElem As Integer
    Dim iEl As Integer
    Dim bExit As Boolean
    
    iActionMode = iVal
    sDbName = sDB
    
    Select Case iActionMode
        Case EnumUserObj
            sCurUser = vntObjects(0)
            Me.Caption = ReplaceMsg(MyLoadResString(k_ObjectOwnedby), Array("1%"), Array(sCurUser))
            cmd(1).Visible = False  'salva
            cmd(2).Caption = MyLoadResString(k_ModifyObjectOwner)
            cmd(2).Width = Me.ScaleWidth - k_MargineFRM - cmd(2).Left
            cmd(2).Visible = True
            bExit = ErrLoadOwnedObjs(sDB, sCurUser)
        Case spidDrop
            Me.Caption = MyLoadResString(k_DeleteSPID)
            ReDim vntGlobal(0, 0) As Variant
            vntGlobal(0, 0) = vntObjects(0)
        Case dbDrop, dbDetach
            Me.Caption = MyLoadResString(IIf(iActionMode = dbDrop, k_Drop_Objects, k_Detach_Databases))
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 1) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
                'not allowed for this element
                'If objServer.Databases("[" & CStr(vntObjects(iEl)) & "]").SystemObject Then
                If IsNumeric(vntObjects(iEl)) Then vntObjects(iEl) = "[" & vntObjects(iEl) & "]"
                If objServer.Databases(CStr(vntObjects(iEl))).SystemObject = True Then
                    vntGlobal(iEl, 1) = "SysObject"
                End If
            Next
            On Local Error GoTo 0
        Case tableDrop
            Dim oTb As SQLDMO.Table
            ShowCmdDepen
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Tables)
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 2) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
                Set oTb = GetTbByName(objServer, sDbName, vntObjects(iEl))
                vntGlobal(iEl, 1) = oTb.Owner
                
                'not allowed for this element
                'vntGlobal(iEl, 1) = objServer.Databases(sDbName).Tables(CStr(vntObjects(iEl))).Owner
                'If objServer.Databases(sDbName).Tables(CStr(vntObjects(iEl))).SystemObject Then
                vntGlobal(iEl, 1) = oTb.Owner
                If oTb.SystemObject Then
                    vntGlobal(iEl, 2) = "SysObject"
                End If
                Set oTb = Nothing
            Next
            On Local Error GoTo 0
            
        Case viewDrop
            Dim oView As SQLDMO.View
            ShowCmdDepen
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Views)
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 2) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
                Set oView = GetViewByName(objServer, sDbName, vntObjects(iEl))
                vntGlobal(iEl, 1) = oView.Owner
                'vntGlobal(iEl, 1) = objServer.Databases(sDbName).Views(CStr(vntObjects(iEl))).Owner
                'not allowed for this element
                'If objServer.Databases(sDbName).Views(CStr(vntObjects(iEl))).SystemObject Then
                If oView.SystemObject Then
                    vntGlobal(iEl, 2) = "SysObject"
                End If
                Set oView = Nothing
            Next
            On Local Error GoTo 0
        Case spDrop
            ShowCmdDepen
            Dim objSp As SQLDMO.StoredProcedure
            
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Stored_Procedure) '" Stored Procedures"
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 2) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                Set objSp = GetSpByName(objServer, sDbName, vntObjects(iEl))
                vntGlobal(iEl, 0) = vntObjects(iEl)
                vntGlobal(iEl, 1) = objSp.Owner
                'vntGlobal(iEl, 1) = objServer.Databases(sDbName).StoredProcedures(CStr(vntObjects(iEl))).Owner
                'not allowed for this element
                'If objServer.Databases(sDbName).StoredProcedures(CStr(vntObjects(iEl))).SystemObject Then
                If objSp.SystemObject Then
                    vntGlobal(iEl, 2) = "SysObject"
                End If
                Set objSp = Nothing
            Next
            On Local Error GoTo 0
        Case funcDrop
            ShowCmdDepen
            Dim objFunc As SQLDMO.UserDefinedFunction
            
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_objFunction)  '" User Function
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 2) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                Set objFunc = GetFuncByName(objServer, sDbName, vntObjects(iEl))
                vntGlobal(iEl, 0) = vntObjects(iEl)
                vntGlobal(iEl, 1) = objFunc.Owner
                'vntGlobal(iEl, 1) = objServer.Databases(sDbName).StoredProcedures(CStr(vntObjects(iEl))).Owner
                'not allowed for this element
                'If objServer.Databases(sDbName).StoredProcedures(CStr(vntObjects(iEl))).SystemObject Then
                If objFunc.SystemObject Then
                    vntGlobal(iEl, 2) = "SysObject"
                End If
                Set objFunc = Nothing
            Next
            On Local Error GoTo 0
        
        Case udtDrop
            ShowCmdDepen
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_User_Defined_Data_Type)
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 0) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
            Next
            On Local Error GoTo 0
        Case userDrop
            'ShowCmdDepen
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Users)
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 0) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
            Next
            On Local Error GoTo 0
        Case roleDrop
            'ShowCmdDepen
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Database_Role)
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 1) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)

                'not allowed for this element
                If objServer.Databases(sDbName).DatabaseRoles(CStr(vntObjects(iEl))).IsFixedRole Then
                    vntGlobal(iEl, 2) = "FixedRole"
                End If
            Next
            On Local Error GoTo 0
        Case loginDrop
            Me.Caption = MyLoadResString(k_Drop_Objects) & " " & MyLoadResString(k_Login) '" Login(s)"
            iElem = UBound(vntObjects)
            ReDim vntGlobal(iElem, 0) As Variant
            On Local Error Resume Next
            For iEl = 0 To iElem
                vntGlobal(iEl, 0) = vntObjects(iEl)
            Next
            On Local Error GoTo 0
        
    End Select
    If Not bExit Then LoadFlex
    
    If iActionMode = EnumUserObj Then cmd(2).Enabled = Flex.Rows <> 1
    m_bExitFrom = bExit
    
End Sub

Private Sub LoadFlex()

    Dim sBuf As String
    Dim i As Integer
    Dim iBool As Integer, bDisable As Boolean
    
    Dim udtC() As udtCol
    
    Dim iCols As Integer
    Dim iRows As Integer
    Dim iLoop As Integer, iPic As Integer
    
    bLoading = True
    Select Case iActionMode
        Case spidDrop
            ReDim udtC(0) As udtCol
            udtC(0).colName = "Spid"
            udtC(0).ColWidth = 2100
            iRows = 0
            iCols = 1
        
        Case dbDrop, dbDetach
            ReDim udtC(1) As udtCol
            udtC(0).colName = MyLoadResString(k_RES_Database) '"Databases"
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = "SysObjects"
            udtC(1).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 2
        Case tableDrop
            ReDim udtC(2) As udtCol
            udtC(0).colName = MyLoadResString(k_Tables)
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = MyLoadResString(k_Owner)
            udtC(1).ColWidth = 400
            
            udtC(2).colName = "SysObjects"
            udtC(2).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 3
        Case viewDrop
            ReDim udtC(2) As udtCol
            udtC(0).colName = MyLoadResString(k_Views)
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = MyLoadResString(k_Owner)
            udtC(1).ColWidth = 400
            
            udtC(2).colName = "SysObjects"
            udtC(2).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 3
        Case spDrop
            ReDim udtC(2) As udtCol
            udtC(0).colName = MyLoadResString(k_Stored_Procedure) '"Stored Procedures"
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = MyLoadResString(k_Owner)
            udtC(1).ColWidth = 400
            
            udtC(2).colName = "SysObjects"
            udtC(2).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 3
        Case funcDrop
            ReDim udtC(2) As udtCol
            udtC(0).colName = MyLoadResString(k_objFunction) '"User Function
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = MyLoadResString(k_Owner)
            udtC(1).ColWidth = 400
            
            udtC(2).colName = "SysObjects"
            udtC(2).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 3
            
        Case roleDrop
            ReDim udtC(1) As udtCol
            
            udtC(0).colName = MyLoadResString(kTviewDatabasesRoles) '"DB Roles"
            udtC(0).ColWidth = 2100
            
            udtC(1).colName = "FixedRole"
            udtC(1).ColWidth = 1100
            
            iRows = UBound(vntGlobal, 1)
            iCols = 2
        Case userDrop
            ReDim udtC(0) As udtCol
            
            udtC(0).colName = MyLoadResString(k_User)
            udtC(0).ColWidth = 2100
            
            iRows = UBound(vntGlobal)
            iCols = 1
        Case udtDrop
            ReDim udtC(0) As udtCol
            
            udtC(0).colName = MyLoadResString(k_User_Defined_Data_Type)
            udtC(0).ColWidth = 2100
                        
            iRows = UBound(vntGlobal)
            iCols = 1
        Case loginDrop
            ReDim udtC(0) As udtCol
            
            udtC(0).colName = MyLoadResString(k_Login) '"login"
            udtC(0).ColWidth = 2100
            
            iRows = UBound(vntGlobal)
            iCols = 1
        Case EnumUserObj
            ReDim udtC(2) As udtCol
            udtC(0).colName = "" 'MyLoadResString(k_Type)
            udtC(0).ColWidth = 500
            
            udtC(1).colName = MyLoadResString(k_Type)
            udtC(1).ColWidth = 500
            
            udtC(2).colName = MyLoadResString(k_Object)
            udtC(2).ColWidth = 400
            
            iRows = -1
            If Not IsEmpty(vntObj) Then
                If IsArray(vntObj) Then iRows = UBound(vntObj, 1)
            End If
            iCols = 3
        
    End Select
    
    With Flex
        .Redraw = False
        .Clear
        .Cols = iCols + 1
        .FixedCols = 0
        .Rows = 1
        .Clear
         
         For iLoop = 0 To UBound(udtC)
            .TextMatrix(0, iLoop + 1) = udtC(iLoop).colName
            .ColWidth(iLoop + 1) = udtC(iLoop).ColWidth
         Next
         
         .ColWidth(0) = 300
         For iLoop = 0 To iRows
            sBuf = ""
            If iActionMode <> EnumUserObj Then
                For i = 0 To iCols - 1
                    sBuf = sBuf & vbTab & vntGlobal(iLoop, i)
                Next
            Else
                sBuf = vbTab & vbTab & vbTab & vntObj(iLoop, 0)
                iPic = vntObj(iLoop, 2)
            End If
            Flex.AddItem sBuf
            
            bDisable = False
            If iActionMode = EnumUserObj Then
                If iPic = SQLDMOObj_SystemTable Or iPic = SQLDMOObj_UserTable Then
                    iPic = k_ResTB
                ElseIf iPic = SQLDMOObj_View Then
                    iPic = k_ResView
                ElseIf iPic = SQLDMOObj_StoredProcedure Then
                    iPic = k_ResSP
                ElseIf iPic = SQLDMOObj_Default Then
                    iPic = k_ResDef
                ElseIf iPic = SQLDMOObj_Rule Then
                    iPic = k_ResRule
                ElseIf iPic = SQLDMOObj_UserDefinedFunction Then
                    iPic = k_ResFUNCTION
                ElseIf iPic = SQLDMOObj_UserDefinedDatatype Then
                    iPic = k_ResUDT
                    bDisable = True
                ElseIf iPic = SQLDMOObj_Trigger Then
                   iPic = k_ResTRIGGER
                   bDisable = True
                Else
                    iPic = -1
                End If
                
                .Row = .Rows - 1
                .Col = 1
                iBool = IIf(CBool(vntObj(iLoop, 3)) = False, k_ResChkOFF, k_ResChkON)
                Set Flex.CellPicture = LoadResPicture(iBool, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                If iPic <> -1 Then
                    .Col = 2
                    Set Flex.CellPicture = LoadResPicture(iPic, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
                If bDisable Then
                    For i = 0 To .Cols - 1
                        .Col = i
                        .CellBackColor = vbInactiveBorder
                    Next
                
                End If
            End If
            
         Next
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        
        .FixedCols = 1
        If .Rows > 1 Then
            .FixedRows = 1
        Else
            .Row = .Rows - 1
            For iLoop = 0 To .Cols - 1
                .Col = iLoop
                .CellBackColor = vbInactiveBorder
            Next
        End If
        .Redraw = True
        SettaLargFlex Flex, .Cols - 1, 300
        If .Rows <> 1 Then .TopRow = 1
        .Row = IIf(.Rows > 1, 1, .Rows - 1)
    End With
    bLoading = False
    Flex_RowColChange

End Sub

Private Sub ShowCmdDepen()
    cmd(2).Caption = MyLoadResString(k_Show_Dependencies)
    cmd(2).Visible = True
End Sub
Private Sub ShowDependencies()
        
    Dim objDependent As clsDepen
    Set objDependent = New clsDepen
    
    objDependent.ShowDependent sDbName, Flex.TextMatrix(Flex.Row, 1), iActionMode
    Set objDependent = Nothing

End Sub

Private Function ErrLoadOwnedObjs(ByVal sDbName As String, ByVal sUser As String) As Boolean

    Dim i As Integer
    Dim sDbo As String
    
    Dim oList As SQLDMO.SQLObjectList
    Dim oObj As SQLDMO.DBObject
    
    On Local Error Resume Next
    Set oList = objServer.Databases(sDbName).Users(sUser).ListOwnedObjects(SQLDMOObj_AllDatabaseUserObjects, SQLDMOObjSort_Type Or SQLDMOObjSort_Name)
    On Local Error Resume Next
    
    If oList Is Nothing Then
        ErrLoadOwnedObjs = True
        Exit Function
    End If
    sDbo = WhoIsDBO(objServer.Databases(sDbName))
    
    vntObj = vbEmpty
    i = oList.Count - 1
    If i <> -1 Then ReDim vntObj(i, 3) As Variant
    
    i = 0
    For Each oObj In oList
        vntObj(i, 0) = oObj.Name
        vntObj(i, 1) = oObj.Owner
        If StrComp(vntObj(i, 1), "dbo", vbTextCompare) = 0 Then vntObj(i, 1) = sDbo
        
        vntObj(i, 2) = oObj.Type
        
        vntObj(i, 3) = 0
        i = i + 1
    Next
    
    Set oObj = Nothing
    Set oList = Nothing
    
End Function
Private Sub LoadChangingObj()

    Dim i As Integer
    Dim iLoop As Integer
    Dim vntObj2Change As Variant
    Dim sOwner As String
    
        
    iLoop = -1
    For i = 0 To UBound(vntObj, 1)
        If CBool(vntObj(i, 3)) Then iLoop = iLoop + 1
    Next
    If iLoop <> -1 Then
        cmd(0).Enabled = False
        cmd(2).Enabled = False
        fResult.Action = act_Null
        fResult.WWrapVisible() = False
        fResult.Caption = Me.Caption
        fResult.tRes.Text = ReplaceMsg(MyLoadResString(k_ChangingPreWarning), Array("|"), Array(vbCrLf))
        fResult.Show vbModal, Me
        
        DoEvents
        Me.Refresh
        
        Screen.MousePointer = vbHourglass
    
        ReDim vntObj2Change(iLoop, 2) As Variant
        iLoop = 0
        For i = 0 To UBound(vntObj, 1)
            If CBool(vntObj(i, 3)) Then
                sOwner = vntObj(i, 1)
                vntObj2Change(iLoop, 0) = vntObj(i, 0)
                vntObj2Change(iLoop, 1) = vntObj(i, 1)
                vntObj2Change(iLoop, 2) = vntObj(i, 2)
                iLoop = iLoop + 1
            End If
        Next
        
    
        Dim fChangeOwner As fDbOwner
        Set fChangeOwner = New fDbOwner
        With fChangeOwner
            .DatabaseObjsProp sDbName, vntObj2Change, sOwner 'sCurUser
            .Show vbModal, Me
        End With
        Unload fChangeOwner
        Set fChangeOwner = Nothing
        
        Screen.MousePointer = vbHourglass
        Stuff_Exec EnumUserObj, sDbName, Array(sCurUser)
        cmd(0).Enabled = True

        Screen.MousePointer = vbDefault
        If m_bExitFrom Then
            frmMain.CallBackReattach
            Unload Me
        End If
    End If
    
End Sub
