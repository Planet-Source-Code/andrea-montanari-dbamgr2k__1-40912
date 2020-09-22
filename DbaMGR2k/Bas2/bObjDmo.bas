Attribute VB_Name = "bObjDmo"
Option Explicit
Public Function IsDbAvailable(ByVal sDbName As String) As Boolean

    Dim oDb As SQLDMO.Database
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    If Not oDb Is Nothing Then
        IsDbAvailable = Not oDb.DBOption.Offline
    End If
    Set oDb = Nothing
    
End Function
Public Function IsDbAvailableReadWrite(ByVal sDbName As String) As Boolean

    Dim oDb As SQLDMO.Database
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    If Not oDb Is Nothing Then
        If Not oDb.DBOption.Offline Then
            IsDbAvailableReadWrite = Not oDb.DBOption.ReadOnly
        End If
    End If
    Set oDb = Nothing
    
End Function

Public Function GetSpByName(ByVal oServer As SQLDMO.SQLServer2, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.StoredProcedure

    Dim i As Integer
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .StoredProcedures.Count
                If StrComp(.StoredProcedures(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetSpByName = .StoredProcedures(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function
Public Function GetDefaultByName(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.Default

    Dim i As Integer
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Defaults.Count
                If StrComp(.Defaults(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetDefaultByName = .Defaults(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function
Public Function GetRuleByName(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.Rule

    Dim i As Integer
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Rules.Count
                If StrComp(.Rules(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetRuleByName = .Rules(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function
'Public Function GetTrigByName(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.Trigger
'
'    Dim i As Integer, y As Integer
'
'    sObjName = Replace(sObjName, "[", "")
'    sObjName = Replace(sObjName, "]", "")
'    If Len(sObjName) <> 0 Then
'        With oServer.Databases(sDB)
'            For i = 1 To .Tables.Count
'                For y = 1 To .Tables(i).Triggers.Count
'                    If StrComp(.Tables(i).Triggers(y).Name, sObjName, vbTextCompare) = 0 Then
'                        Set GetTrigByName = .Tables(i).Triggers(y)
'                        Exit Function
'                    End If
'                Next
'            Next
'        End With
'    End If
'
'End Function
Public Function GetViewByName(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.View

    Dim i As Integer
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Views.Count
                If StrComp(.Views(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetViewByName = .Views(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function
Public Function GetTbByName(ByVal oServer As SQLDMO.SQLServer2, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.Table

    Dim i As Integer
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Tables.Count
                If StrComp(.Tables(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetTbByName = .Tables(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function

Public Function GetUdtByName(ByVal oServer As SQLDMO.SQLServer2, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.UserDefinedDatatype

    Dim i As Integer
    Dim oDb As SQLDMO.Database2
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        Set oDb = oServer.Databases(sDB)
        With oDb
            For i = 1 To .UserDefinedDatatypes.Count
                If StrComp(.UserDefinedDatatypes(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetUdtByName = .UserDefinedDatatypes(i)
                    Exit For
                End If
            Next
        End With
    End If

End Function

Public Function GetViewOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String

    Dim i As Integer
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Views.Count
                If StrComp(.Views(i).Name, sObjName, vbTextCompare) = 0 Then
                    GetViewOwner = .Views(i).Owner
                    Exit For
                End If
            Next
        End With
    End If

End Function

Public Function GetTableOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String

    Dim i As Integer
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Tables.Count
                If StrComp(.Tables(i).Name, sObjName, vbTextCompare) = 0 Then
                    GetTableOwner = .Tables(i).Owner
                    Exit For
                End If
            Next
        End With
    End If

End Function

Public Function GetSPOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String
    
    Dim i As Integer
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .StoredProcedures.Count
                If StrComp(.StoredProcedures(i).Name, sObjName, vbTextCompare) = 0 Then
                    GetSPOwner = .StoredProcedures(i).Owner
                    Exit For
                End If
            Next
        End With
    End If

End Function
'Public Function GetUDTOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String
'
'    Dim i As Integer
'    sObjName = Replace(sObjName, "[", "")
'    sObjName = Replace(sObjName, "]", "")
'    If Len(sObjName) <> 0 Then
'        With oServer.Databases(sDB)
'            For i = 1 To .UserDefinedDatatypes.Count
'                If StrComp(.UserDefinedDatatypes(i).Name, sObjName, vbTextCompare) = 0 Then
'                    GetUDTOwner = .UserDefinedDatatypes(i).Owner
'                    Exit For
'                End If
'            Next
'        End With
'    End If
'
'End Function
'Public Function GetTrigOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String
'
'    Dim i As Integer, y As Integer
'    sObjName = Replace(sObjName, "[", "")
'    sObjName = Replace(sObjName, "]", "")
'    If Len(sObjName) <> 0 Then
'        With oServer.Databases(sDB)
'            For i = 1 To .Tables.Count
'                For y = 1 To .Tables(i).Triggers.Count
'                    If StrComp(.Tables(i).Triggers(y).Name, sObjName, vbTextCompare) = 0 Then
'                        GetTrigOwner = .Tables(i).Triggers(y).Owner
'                        Exit Function
'                    End If
'                Next
'            Next
'        End With
'    End If
'
'End Function
Public Function GetCreateOwner(ByVal sDbName As String) As String

    Dim sOwner As String
    sOwner = objServer.Login
    If GetMembership(SysAdm) Or GetMembership(db_dbo, sDbName) Then
        sOwner = "dbo"
    End If
    GetCreateOwner = sOwner

End Function
Public Function GetDefaultOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String

    Dim i As Integer
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Defaults.Count
                If StrComp(.Defaults(i).Name, sObjName, vbTextCompare) = 0 Then
                    GetDefaultOwner = .Defaults(i).Owner
                    Exit For
                End If
            Next
        End With
    End If

End Function
Public Function GetRuleOwner(ByVal oServer As SQLDMO.SQLServer, ByVal sDB As String, ByVal sObjName As String) As String

    Dim i As Integer
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        With oServer.Databases(sDB)
            For i = 1 To .Rules.Count
                If StrComp(.Rules(i).Name, sObjName, vbTextCompare) = 0 Then
                    GetRuleOwner = .Rules(i).Owner
                    Exit For
                End If
            Next
        End With
    End If

End Function

Public Function errTestBackUpSet(ByVal oRest As SQLDMO.Restore) As Boolean
    Dim oQry As SQLDMO.QueryResults
    
    On Local Error Resume Next
    Set oQry = oRest.ReadBackupHeader(objServer)
    errTestBackUpSet = Err.Number <> 0
    On Local Error GoTo 0
    Set oQry = Nothing

End Function
Public Function GetFuncByName(ByVal oServer As SQLDMO.SQLServer2, ByVal sDB As String, ByVal sObjName As String) As SQLDMO.UserDefinedFunction

    Dim i As Integer
    Dim oDb As SQLDMO.Database2
    
    sObjName = Replace(sObjName, "[", "")
    sObjName = Replace(sObjName, "]", "")
    If Len(sObjName) <> 0 Then
        Set oDb = oServer.Databases(sDB)
        With oDb
            For i = 1 To .UserDefinedFunctions.Count
                If StrComp(.UserDefinedFunctions(i).Name, sObjName, vbTextCompare) = 0 Then
                    Set GetFuncByName = .UserDefinedFunctions(i)
                    Exit For
                End If
            Next
        End With
    End If
    Set oDb = Nothing

End Function
Public Function UDFtype(ByVal iType As SQLDMO.SQLDMO_UDF_TYPE) As String

    Dim sRet As String
    Select Case iType
        Case SQLDMOUDF_Unknown
            sRet = MyLoadResString(k_Unknown)
        Case SQLDMOUDF_Scalar
            sRet = MyLoadResString(k_Func_Scalar)
        Case SQLDMOUDF_Table
            sRet = MyLoadResString(k_Func_Table)
        Case SQLDMOUDF_Inline
            sRet = MyLoadResString(k_Func_InLine)
    End Select
    UDFtype = sRet
    
End Function

Public Function StoredProcType(ByVal iType As SQLDMO.SQLDMO_PROCEDURE_TYPE) As String

    Dim sRet As String
    Select Case iType
        Case SQLDMOProc_Unknown, SQLDMOProc_Macro, SQLDMOProc_ReplicationFilter
            sRet = MyLoadResString(k_Unknown)
        Case SQLDMOProc_Standard
            sRet = MyLoadResString(k_Sp_Type_Standard)
        Case SQLDMOProc_Extended
            sRet = MyLoadResString(k_Sp_Type_Extended)
    End Select
    
    StoredProcType = sRet
    
End Function
Public Function TriggerType(ByVal iType As SQLDMO_TRIGGER_TYPE) As String

    Dim sRet As String
    
    Select Case iType
        Case SQLDMOTrig_All
            sRet = MyLoadResString(k_TRI_Type_All)
        Case SQLDMOTrig_Delete
            sRet = MyLoadResString(k_TRI_Type_Delete)
        Case SQLDMOTrig_Insert
            sRet = MyLoadResString(k_TRI_Type_Insert)
        Case SQLDMOTrig_Update
            sRet = MyLoadResString(k_TRI_Type_Update)
        Case Else
            sRet = MyLoadResString(k_Unknown)
    End Select
    
    TriggerType = sRet
    
End Function

Public Function IsObjectOwner(ByVal sDbName As String, ByVal sObjName As String, ByVal iObjType As am_SqlPropTypeOwner) As Boolean

    Dim oObj As Object
    Dim oLogin As SQLDMO.Login
            
    If iObjType = am_OwnSP Then
        Set oObj = GetSpByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnView Then
        Set oObj = GetViewByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnFunction Then
        Set oObj = GetFuncByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnTable Then
        Set oObj = GetTbByName(objServer, sDbName, sObjName)
    End If
    
    If Not oObj Is Nothing Then
        If StrComp(oObj.Owner, objServer.Login, vbTextCompare) = 0 Then
            IsObjectOwner = True
        Else
            Set oLogin = objServer.Logins(objServer.Login)
            If StrComp(oObj.Owner, oLogin.GetUserName(sDbName), vbTextCompare) = 0 Then IsObjectOwner = True
            Set oLogin = Nothing
        End If
    End If
    Set oObj = Nothing
    
End Function
Public Function GetObjectOwner(ByVal sDbName As String, ByVal sObjName As String, ByVal iObjType As am_SqlPropTypeOwner) As String

    Dim oObj As Object
            
    If iObjType = am_OwnSP Then
        Set oObj = GetSpByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnView Then
        Set oObj = GetViewByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnFunction Then
        Set oObj = GetFuncByName(objServer, sDbName, sObjName)
    ElseIf iObjType = am_OwnTable Then
        Set oObj = GetTbByName(objServer, sDbName, sObjName)
    End If
    
    If Not oObj Is Nothing Then GetObjectOwner = oObj.Owner
    Set oObj = Nothing
    
End Function

Public Function ErrExexSqlDirect(ByVal oDb As SQLDMO.Database2, ByVal sSql As String, ByRef sErr As String) As Long

    Dim lErr As Long
    
    If Not oDb Is Nothing Then
        On Local Error Resume Next
        oDb.ExecuteImmediate sSql, SQLDMOExec_Default, Len(sSql)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        Debug.Print lErr; sErr
    End If
    ErrExexSqlDirect = lErr
End Function


