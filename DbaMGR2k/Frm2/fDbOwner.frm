VERSION 5.00
Begin VB.Form fDbOwner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change DB Owner"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "fDbOwner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox tOwner 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "fDbOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sDbName As String
Private m_iMode As amChangeOwner
Private m_vntObj2Change As Variant

Public Sub DatabaseProp(ByVal sVal As String)
    
    Dim lErr As Long, sErr As String
    Dim sOwner As String
    Dim sItem As String, sTmp As String
    
    Dim oLog As SQLDMO.Login
    
    Screen.MousePointer = vbHourglass
    m_iMode = amChangeDB
    
    Me.Caption = Replace(MyLoadResString(k_Change_DB_Owner_Frm), "1%", sVal)
    
    sDbName = sVal
    
    cmd(1).Enabled = False
    Dim oDb As SQLDMO.Database2
    
    objServer.Databases.Refresh
    On Local Error Resume Next
    Set oDb = objServer.Databases(sVal)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        cbo.Clear
        For Each oLog In objServer.Logins
            If oLog.IsMember("sysadmin") Or oLog.IsMember("dbcreator") Then
                sItem = oLog.Name
                
                sTmp = oLog.GetUserName(sDbName)
                If Len(sTmp) <> 0 Then
                    If StrComp(sTmp, sItem, vbTextCompare) <> 0 Then
                        If (StrComp("sa", sItem, vbTextCompare) <> 0) And (StrComp("dbo", sTmp, vbTextCompare) <> 0) Then sItem = sTmp
                    End If
                End If
                cbo.AddItem sItem
                cbo.ItemData(cbo.NewIndex) = -1 'is a login
            End If
        Next
        
        If cbo.ListCount <> 0 Then
            sOwner = oDb.Owner
            tOwner.Text = sOwner
            cbo.ListIndex = GetItem(sOwner, cbo)
            cmd(1).Enabled = True
        End If
    End If
    
    Set oLog = Nothing
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
   
    
End Sub

'Public Sub DatabaseObjsProp(ByVal sDB As String, vntObj As Variant, ByVal sOldOwner As String)
'
'    Dim iItem As Integer
'    Dim lErr As Long, sErr As String
'    Dim oUser As SQLDMO.User
'    Dim oLogin As SQLDMO.Login
'    Dim sItem As String, sTmp As String
'
'    Screen.MousePointer = vbHourglass
'    m_iMode = amChangeObj
'
'    Me.Caption = Replace(MyLoadResString(k_Change_Obj_Owner_Frm), "1%", sDB)
'
'    sDbName = sDB
'
'    cmd(1).Enabled = False
'    Dim oDB As SQLDMO.Database2
'
'    objServer.Databases.Refresh
'    On Local Error Resume Next
'    Set oDB = objServer.Databases(sDB)
'    lErr = Err.Number
'    sErr = Err.Description
'    On Local Error GoTo 0
'
'    If lErr = 0 Then
'        cbo.Clear
'        For Each oLogin In objServer.Logins
'            If oLogin.IsMember("sysadmin") Then
'                sItem = oLogin.Name
'
'                sTmp = oLogin.GetUserName(sDB)
'                If Len(sTmp) <> 0 Then
'                    If (StrComp(sTmp, sItem, vbTextCompare) <> 0) And Not (StrComp(sTmp, "dbo", vbTextCompare) = 0) Then sItem = sTmp
'                End If
'                cbo.AddItem sItem 'oLogin.Name
'                cbo.ItemData(cbo.NewIndex) = -1 'is a login
'            End If
'        Next
'        For Each oUser In oDB.Users
'            If oUser.IsMember("db_ddladmin") Or oUser.IsMember("db_owner") Then
'                sItem = oUser.Name
'                If StrComp(sItem, "dbo", vbTextCompare) <> 0 Then
'                    iItem = GetItem(sItem, cbo)
'                    If iItem = -1 Then cbo.AddItem sItem
'                End If
'            End If
'        Next
'
'        If cbo.ListCount <> 0 Then
'            tOwner.Text = sOldOwner
'            cbo.ListIndex = GetItem(sOldOwner, cbo)
'            cmd(1).Enabled = True
'        End If
'    End If
'    Set oUser = Nothing
'    Set oDB = Nothing
'    m_vntObj2Change = vntObj
'
'    Screen.MousePointer = vbDefault
'    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
'
'End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Unload Me
    Else
        If m_iMode = amChangeDB Then
            ChangeDBOwner
        Else
            ChangeObjectsOwner
        End If
    End If
End Sub

Private Sub Form_Load()
    lbl(0).Caption = MyLoadResString(k_Old_Obj_Owner)
    lbl(1).Caption = MyLoadResString(k_New_Obj_Owner)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fDbOwner = Nothing
End Sub

Private Sub ChangeDBOwner()

    Dim sNewDbOwner As String
    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim sMessage As String
    Dim oLogin As SQLDMO.Login
    Dim sTMPLogin As String
    Dim lErr2 As Long, sErr2 As String
    Dim sMsg As String
    
    
    sNewDbOwner = cbo.Text
    
    If StrComp(sNewDbOwner, tOwner.Text, vbTextCompare) = 0 Then
        If MsgBox(ReplaceMsg(MyLoadResString(k_Change_DbOwnerSame), Array("1%", "2%", "|"), Array(sNewDbOwner, sDbName, vbCrLf)), vbQuestion Or vbOKCancel, App.EXEName) = vbCancel Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    sTMPLogin = "am-" & Format$(Now, "YYYYMMDD")
'    Set oLogin = New SQLDMO.Login
'    oLogin.Name = sTMPLogin
'    oLogin.Type = SQLDMOLogin_Standard
'
'    On Local Error Resume Next
'    objServer.Logins.Add oLogin
'    lErr = Err.Number
'    sErr = Err.Description
'    On Local Error GoTo 0
'
'    If lErr = 0 Then
'        On Local Error Resume Next
'        objServer.ServerRoles("sysadmin").AddMember sTMPLogin
'        lErr = Err.Number
'        sErr = Err.Description
'        On Local Error GoTo 0
'    End If
    
    lErr = ErrExistThisLogin(sTMPLogin, oLogin, sErr)
    If lErr = 0 Then
        On Local Error Resume Next
        Set oDb = objServer.Databases(sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    
        If lErr = 0 Then
            If StrComp(sTMPLogin, oDb.Owner, vbTextCompare) <> 0 Then
                On Local Error Resume Next
                oDb.SetOwner sTMPLogin, OverrideIfAlreadyUser:=True
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
            End If
        End If
        
        If lErr = 0 Then
            On Local Error Resume Next
            oDb.SetOwner sNewDbOwner, OverrideIfAlreadyUser:=True
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            
            If lErr = 0 Then sMessage = ReplaceMsg(MyLoadResString(k_Modified_DB_Owner), Array("1%", "2%", "|"), Array(sDbName, sNewDbOwner, vbCrLf))
        End If
    
    End If

    If Not oLogin Is Nothing Then
        If StrComp(sTMPLogin, sNewDbOwner, vbTextCompare) <> 0 Then
            Set oLogin = Nothing
            On Local Error Resume Next
            objServer.Logins.Remove sTMPLogin
            lErr2 = Err.Number
            sErr2 = Err.Description
            On Local Error GoTo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
    Set oDb = Nothing
    If lErr <> 0 Or lErr2 <> 0 Then
        If lErr <> 0 Then sMsg = MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        If lErr2 <> 0 Then sMsg = sMsg & vbCrLf & MyLoadResString(kMsgBoxError) & ": " & lErr2 & " - " & sErr2
        MsgBox sMsg, vbInformation Or vbOKOnly, App.EXEName
        'MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        MsgBox sMessage, vbInformation Or vbOKOnly, App.EXEName
        DatabaseProp sDbName
    End If
        
End Sub

'Private Sub ChangeObjectsOwner()
'
'    Dim sNewOwner As String
'    Dim lErr As Long, sErr As String
'    Dim oDb As SQLDMO.Database2
'    Dim sMessage As String
'    Dim iLoop As Integer
'    Dim lType As Integer
'    Dim sType As String
'    Dim oObj As Object 'SQLDMO.Table
'    Dim sDbo As String
'    Dim bExit As Boolean
'
'    cmd(1).Enabled = False
'
'    sNewOwner = cbo.Text
'
'    If StrComp(sNewOwner, tOwner.Text, vbTextCompare) <> 0 Then
'        Screen.MousePointer = vbHourglass
'
'        On Local Error Resume Next
'        Set oDb = objServer.Databases(sDbName)
'        lErr = Err.Number
'        sErr = Err.Description
'        On Local Error GoTo 0
'
'        If lErr = 0 Then
'
'            If cbo.ItemData(cbo.ListIndex) = -1 Then    'is a login
'                bExit = ErrExistThisUser(oDb, sNewOwner, sMessage)
'            End If
'
'            If Not bExit Then
'                sDbo = oDb.Owner
'                If StrComp(sDbo, sNewOwner, vbTextCompare) = 0 Or StrComp("sa", sNewOwner, vbTextCompare) = 0 Then sNewOwner = "dbo"
'                oDb.Tables.Refresh True
'                oDb.Views.Refresh True
'                oDb.UserDefinedDatatypes.Refresh True
'                oDb.StoredProcedures.Refresh True
'                oDb.Rules.Refresh True
'                oDb.Defaults.Refresh True
'                oDb.UserDefinedFunctions.Refresh True
'
'                For iLoop = 0 To UBound(m_vntObj2Change, 1)
'                    lType = m_vntObj2Change(iLoop, 2)
'
'                    If lType = SQLDMOObj_SystemTable Or lType = SQLDMOObj_UserTable Then
'                        Set oObj = GetTbByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = MyLoadResString(k_Table)
'                    ElseIf lType = SQLDMOObj_View Then
'                        Set oObj = GetViewByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = MyLoadResString(k_View)
'                    ElseIf lType = SQLDMOObj_StoredProcedure Then
'                        Set oObj = GetSpByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = MyLoadResString(k_Stored_Procedure)
'                    ElseIf lType = SQLDMOObj_Default Then
'                        Set oObj = GetDefaultByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = MyLoadResString(k_ObjDefault)
'                    ElseIf lType = SQLDMOObj_Rule Then
'                        Set oObj = GetRuleByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = MyLoadResString(k_Rule)
'                    ElseIf lType = SQLDMOObj_UserDefinedDatatype Then
'                        sType = MyLoadResString(k_User_Defined_Data_Type)
'                    ElseIf lType = SQLDMOObj_Trigger Then
'                        'Set oObj = GetTrigByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                        sType = "Trigger"
'                    ElseIf lType = SQLDMOObj_UserDefinedFunction Then
'                        sType = MyLoadResString(k_objFunction)
'                        Set oObj = GetFuncByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
'                    Else
'                        sType = MyLoadResString(k_RES_Object_Not_Found_simple)
'                    End If
'
'                    If Len(sMessage) <> 0 Then sMessage = sMessage & vbCrLf & String$(20, "-") & vbCrLf
'
'                    sMessage = sMessage & ReplaceMsg(MyLoadResString(k_ChangingObjectOwner), Array("1%", "2%", "3%", "4%"), Array(sType, m_vntObj2Change(iLoop, 0), m_vntObj2Change(iLoop, 1), sNewOwner)) & vbCrLf
'                    Debug.Print sMessage
'
'                    If Not oObj Is Nothing Then
'                        On Local Error Resume Next
'                        oObj.Owner = sNewOwner
'                        lErr = Err.Number
'                        sErr = Err.Description
'                        On Local Error GoTo 0
'                    Else
'                        lErr = Err_Free
'                        sErr = MyLoadResString(k_ChangingImpossible)
'                    End If
'                    If lErr <> 0 Then
'                        sMessage = sMessage & MyLoadResString(kMsgBoxError) & ": " & IIf(lErr <> Err_Free, lErr & " - ", "") & sErr
'                    Else
'                        sMessage = sMessage & MyLoadResString(k_ChangingDone)
'                    End If
'                    Debug.Print sMessage
'                    If lErr <> 0 Then Exit For
'
'                    sType = ""
'                    Set oObj = Nothing
'                Next
'            End If
'        End If
'
'        Screen.MousePointer = vbDefault
'    Else
'        sMessage = MyLoadResString(k_DB_Owner_is_Same)
'    End If
'
'    Set oDb = Nothing
'    fResult.Action() = act_Null
'    fResult.WWrapVisible() = False
'    fResult.tRes.Text = sMessage
'    fResult.Caption = MyLoadResString(k_ChangingResult)
'    fResult.Show vbModal, Me
'    Unload Me
'
'End Sub
'
'Private Function ErrExistThisUser(ByRef oDb As SQLDMO.Database2, ByVal sUserName As String, ByRef sMessage As String) As Boolean
'
'    Dim oUser As SQLDMO.User
'    Dim oLogin As SQLDMO.Login
'    Dim bSkip As Boolean
'    Dim bAdd2DDL_Admin As Boolean
'    Dim lErr As Long, sErr As String
'
'    If StrComp(sUserName, "sa", vbTextCompare) = 0 Then Exit Function
'
'    bAdd2DDL_Admin = True
'    On Local Error Resume Next
'    Set oUser = oDb.Users(sUserName)
'
'    If oUser Is Nothing Then
'        Err.Clear
'        For Each oUser In oDb.Users
'            If StrComp(oUser.Login, sUserName, vbTextCompare) = 0 Then
'                Set oLogin = objServer.Logins(oUser.Login)
'
'                If oLogin.IsMember("sysadmin") Or oLogin.SystemObject Then
'                    bAdd2DDL_Admin = False
'                End If
'                bSkip = True
'                Exit For
'            End If
'        Next
'        If Not bSkip Then
'            sMessage = ReplaceMsg(MyLoadResString(k_AddingChangingUser), Array("1%", "2%"), Array(sUserName, sDbName)) & vbCrLf
'
'            Set oUser = New SQLDMO.User
'            oUser.Name = sUserName
'            oUser.Login = sUserName
'
'            oDb.Users.Add oUser
'
'            If Err.Number = 0 Then
'                Set oLogin = objServer.Logins(oUser.Login)
'
'                If oLogin.IsMember("sysadmin") Or oLogin.SystemObject Then bAdd2DDL_Admin = False
'            End If
'        End If
'        Debug.Print Err.Description
'        If Err.Number = 0 Then
'            If bAdd2DDL_Admin Then
'                Debug.Print oUser.Name, oUser.IsMember("db_ddladmin"), oUser.IsMember("db_owner"), oUser.SystemObject
'                If oUser.IsMember("db_ddladmin") Or oUser.IsMember("db_owner") Or oUser.SystemObject Then
'                    bAdd2DDL_Admin = False
'                End If
'            End If
'            If bAdd2DDL_Admin Then
'                oDb.DatabaseRoles("db_ddladmin").AddMember sUserName
'                sMessage = sMessage & ReplaceMsg(MyLoadResString(k_Adding_DDL_User), Array("1%", "2%"), Array(sUserName, "db_ddladmin"))
'            End If
'        End If
'        If Err.Number <> 0 Then
'            lErr = Err.Number
'            sErr = Err.Description
'
'            sMessage = sMessage & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
'            ErrExistThisUser = True
'        End If
'    End If
'
'End Function
Private Function ErrExistThisLogin(ByVal sLoginName As String, ByRef oLogin As SQLDMO.Login, ByRef sMessage As String) As Boolean
            
    Dim lErr As Long, sErr As String
    
    On Local Error Resume Next
    Set oLogin = objServer.Logins(sLoginName)
    
    If oLogin Is Nothing Then
        Err.Clear
        sMessage = ReplaceMsg(MyLoadResString(k_AddingChangingLogin), Array("1%"), Array(sLoginName)) & vbCrLf
        
        Set oLogin = New SQLDMO.Login
        oLogin.Name = sLoginName
        oLogin.Type = SQLDMOLogin_Standard
        
        objServer.Logins.Add oLogin
    End If
    
    If Err.Number = 0 And Not oLogin Is Nothing Then
        If Not oLogin.IsMember("sysadmin") Then
            objServer.ServerRoles("sysadmin").AddMember sLoginName
            sMessage = sMessage & ReplaceMsg(MyLoadResString(k_ChangingLoginRoles), Array("1%", "2%"), Array("sysadmin", sLoginName))
        End If
    End If
    If Err.Number <> 0 Then
        lErr = Err.Number
        sErr = Err.Description
        sMessage = sMessage & MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr
        ErrExistThisLogin = True
        Set oLogin = Nothing
    End If
        
End Function
Public Sub DatabaseObjsProp(ByVal sDB As String, vntObj As Variant, ByVal sOldOwner As String)
    
    Dim iItem As Integer
    Dim lErr As Long, sErr As String
    Dim oUser As SQLDMO.User
    Dim sItem As String
    
    Screen.MousePointer = vbHourglass
    m_iMode = amChangeObj
    
    Me.Caption = Replace(MyLoadResString(k_Change_Obj_Owner_Frm), "1%", sDB)
    
    sDbName = sDB
    
    cmd(1).Enabled = False
    Dim oDb As SQLDMO.Database2
    
    objServer.Databases.Refresh
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDB)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        cbo.Clear
        sItem = oDb.Owner
        cbo.AddItem sItem 'oLogin.Name
        cbo.ItemData(cbo.NewIndex) = -1 'is a login
        
        For Each oUser In oDb.Users
            Debug.Print oUser.Name
            If oUser.IsMember("db_ddladmin") Or oUser.IsMember("db_owner") Then
                sItem = oUser.Name
                If StrComp(sItem, "dbo", vbTextCompare) <> 0 Then
                    iItem = GetItem(sItem, cbo)
                    If iItem = -1 Then cbo.AddItem sItem
                End If
            End If
        Next
        
        If cbo.ListCount <> 0 Then
            tOwner.Text = sOldOwner
            cbo.ListIndex = GetItem(sOldOwner, cbo)
            cmd(1).Enabled = True
        End If
    End If
    Set oUser = Nothing
    Set oDb = Nothing
    m_vntObj2Change = vntObj
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
End Sub

Private Sub ChangeObjectsOwner()

    Dim sNewOwner As String
    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim sMessage As String
    Dim iLoop As Integer
    Dim lType As Integer
    Dim sType As String
    Dim oObj As Object 'SQLDMO.Table
    Dim sDbo As String
        
    cmd(1).Enabled = False
    
    sNewOwner = cbo.Text
    
    If StrComp(sNewOwner, tOwner.Text, vbTextCompare) <> 0 Then
        Screen.MousePointer = vbHourglass
        
        On Local Error Resume Next
        Set oDb = objServer.Databases(sDbName)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If lErr = 0 Then
        
            sDbo = oDb.Owner
            If StrComp(sDbo, sNewOwner, vbTextCompare) = 0 Or StrComp("sa", sNewOwner, vbTextCompare) = 0 Then sNewOwner = "dbo"
            oDb.Tables.Refresh True
            oDb.Views.Refresh True
            oDb.UserDefinedDatatypes.Refresh True
            oDb.StoredProcedures.Refresh True
            oDb.Rules.Refresh True
            oDb.Defaults.Refresh True
            oDb.UserDefinedFunctions.Refresh True
        
            For iLoop = 0 To UBound(m_vntObj2Change, 1)
                lType = m_vntObj2Change(iLoop, 2)
            
                If lType = SQLDMOObj_SystemTable Or lType = SQLDMOObj_UserTable Then
                    Set oObj = GetTbByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = MyLoadResString(k_Table)
                ElseIf lType = SQLDMOObj_View Then
                    Set oObj = GetViewByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = MyLoadResString(k_View)
                ElseIf lType = SQLDMOObj_StoredProcedure Then
                    Set oObj = GetSpByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = MyLoadResString(k_Stored_Procedure)
                ElseIf lType = SQLDMOObj_Default Then
                    Set oObj = GetDefaultByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = MyLoadResString(k_ObjDefault)
                ElseIf lType = SQLDMOObj_Rule Then
                    Set oObj = GetRuleByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = MyLoadResString(k_Rule)
                ElseIf lType = SQLDMOObj_UserDefinedDatatype Then
                    sType = MyLoadResString(k_User_Defined_Data_Type)
                ElseIf lType = SQLDMOObj_Trigger Then
                    'Set oObj = GetTrigByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                    sType = "Trigger"
                ElseIf lType = SQLDMOObj_UserDefinedFunction Then
                    sType = MyLoadResString(k_objFunction)
                    Set oObj = GetFuncByName(objServer, sDbName, CStr(m_vntObj2Change(iLoop, 0)))
                Else
                    sType = MyLoadResString(k_RES_Object_Not_Found_simple)
                End If
                                
                If Len(sMessage) <> 0 Then sMessage = sMessage & vbCrLf & String$(20, "-") & vbCrLf

                sMessage = sMessage & ReplaceMsg(MyLoadResString(k_ChangingObjectOwner), Array("1%", "2%", "3%", "4%"), Array(sType, m_vntObj2Change(iLoop, 0), m_vntObj2Change(iLoop, 1), sNewOwner)) & vbCrLf
                Debug.Print sMessage
                
                If Not oObj Is Nothing Then
                    On Local Error Resume Next
                    'sNewOwner = "dbo"
                    oObj.Owner = sNewOwner
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                Else
                    lErr = Err_Free
                    sErr = MyLoadResString(k_ChangingImpossible)
                End If
                If lErr <> 0 Then
                    sMessage = sMessage & MyLoadResString(kMsgBoxError) & ": " & IIf(lErr <> Err_Free, lErr & " - ", "") & sErr
                Else
                    sMessage = sMessage & MyLoadResString(k_ChangingDone)
                End If
                Debug.Print sMessage
                If lErr <> 0 Then Exit For
                
                sType = ""
                Set oObj = Nothing
            Next
        End If
        
        Screen.MousePointer = vbDefault
    Else
        sMessage = MyLoadResString(k_DB_Owner_is_Same)
    End If
    
    Set oDb = Nothing
    fResult.Action() = act_Null
    fResult.WWrapVisible() = False
    fResult.tRes.Text = sMessage
    fResult.Caption = MyLoadResString(k_ChangingResult)
    fResult.Show vbModal, Me
    Unload Me
        
End Sub

