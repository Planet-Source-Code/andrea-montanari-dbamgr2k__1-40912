VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDepen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ErrGetDependent(ByVal sDbName As String, ByVal sObjName As String, ByVal objType As optLstAction, ByRef collChild As Collection, ByRef collParent As Collection, ByRef sErr As String, Optional ByVal oTB4Trig As Object) As Long 'SQLDMO.Table = Nothing) As Long

    Dim oQry As SQLDMO.QueryResults
    Dim oList As SQLObjectList
    Dim sBuf As String
    Dim obj As cDepObj
    Dim lErr As Long
    Dim oTb As SQLDMO.Table
    Dim oView As SQLDMO.View
    Dim oSP As SQLDMO.StoredProcedure
    Dim oUdt As SQLDMO.UserDefinedDatatype
    Dim oFunc As SQLDMO.UserDefinedFunction
        
    Set collChild = New Collection
    Set collParent = New Collection

    On Local Error Resume Next
    If objType = tableDrop Then
        Set oTb = GetTbByName(objServer, sDbName, sObjName)
        Set oQry = oTb.EnumDependencies(SQLDMODep_Children) 'SQLDMODep_Children Or SQLDMODep_Parents)
        Set collChild = LoadCollFromQRY(oQry)
        Set oQry = oTb.EnumDependencies(SQLDMODep_Parents)
        Set collParent = LoadCollFromQRY(oQry)
        Set oTb = Nothing
    ElseIf objType = funcDrop Then
        Set oFunc = GetFuncByName(objServer, sDbName, sObjName)
        Set oQry = oFunc.EnumDependencies(SQLDMODep_Children)
        Set collChild = LoadCollFromQRY(oQry)
        Set oQry = oFunc.EnumDependencies(SQLDMODep_Parents)
        Set collParent = LoadCollFromQRY(oQry)
        Set oFunc = Nothing
    
    ElseIf objType = triggerDROP Then
        If Not oTB4Trig Is Nothing Then
            Set oQry = oTB4Trig.Triggers(sObjName).EnumDependencies(SQLDMODep_Children) 'SQLDMODep_Children Or SQLDMODep_Parents)
            Set collChild = LoadCollFromQRY(oQry)
            Set oQry = oTB4Trig.Triggers(sObjName).EnumDependencies(SQLDMODep_Parents) 'SQLDMODep_Children Or SQLDMODep_Parents)
            Set collParent = LoadCollFromQRY(oQry)
        End If
    ElseIf objType = viewDrop Then
        Set oView = GetViewByName(objServer, sDbName, sObjName)
        Set oQry = oView.EnumDependencies(SQLDMODep_Children)  'SQLDMODep_Children Or SQLDMODep_Parents)
        Set collChild = LoadCollFromQRY(oQry)
        Set oQry = oView.EnumDependencies(SQLDMODep_Parents)  'SQLDMODep_Children Or SQLDMODep_Parents)
        Set collParent = LoadCollFromQRY(oQry)
        Set oView = Nothing
        
    ElseIf objType = spDrop Then
        Set oSP = GetSpByName(objServer, sDbName, sObjName)
        Set oQry = oSP.EnumDependencies(SQLDMODep_Children) 'SQLDMODep_Children Or SQLDMODep_Parents)
        Set collChild = LoadCollFromQRY(oQry)
        Set oQry = oSP.EnumDependencies(SQLDMODep_Parents)
        Set collParent = LoadCollFromQRY(oQry)
        Set oSP = Nothing
    ElseIf objType = udtDrop Then
        Set collChild = New Collection
        Set collParent = New Collection
        Set oUdt = GetUdtByName(objServer, sDbName, sObjName)
        Set oList = oUdt.ListBoundColumns
        Set collChild = LoadCollFromSQL_LIST(oList)
        sBuf = oUdt.Default
        If Len(sBuf) <> 0 Then
            Set obj = New cDepObj
            obj.oSequence = 1
            obj.oType = SQLDMOObj_Default
            obj.oName = sBuf
            obj.oOwner = oUdt.Owner
            collParent.Add obj
            Set obj = Nothing
        End If
        sBuf = oUdt.Rule
        If Len(sBuf) <> 0 Then
            
            Set obj = New cDepObj
            obj.oSequence = 1
            obj.oType = SQLDMOObj_Rule
            obj.oName = sBuf
            obj.oOwner = oUdt.RuleOwner
            collParent.Add obj
            Set obj = Nothing
        End If
        Set oUdt = Nothing
                
    End If
    Set oUdt = Nothing
    Set oTb = Nothing
    Set oView = Nothing
    Set oSP = Nothing
    Set oTB4Trig = Nothing
    
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    ErrGetDependent = lErr

End Function

Private Function LoadCollFromQRY(oQry As SQLDMO.QueryResults) As Collection

    Dim x As Integer
    Dim lRows As Long, l As Long
    Dim oObj As cDepObj
    
    Dim retColl As New Collection
    
    If oQry.ResultSets <> 0 Then
        For x = 1 To oQry.ResultSets
            oQry.CurrentResultSet = x
            lRows = oQry.Rows
            For l = 1 To lRows
                Set oObj = New cDepObj
                oObj.oSequence = oQry.GetColumnLong(l, 4)
                oObj.oType = oQry.GetColumnLong(l, 1)
                oObj.oName = oQry.GetColumnString(l, 2)
                oObj.oOwner = oQry.GetColumnString(l, 3)
                On Local Error Resume Next
                retColl.Add oObj, oObj.oName
                On Local Error GoTo 0
            Next
        Next
    End If
    Set oObj = Nothing
    
    Set LoadCollFromQRY = retColl
                         
End Function
Private Function LoadCollFromSQL_LIST(oList As SQLObjectList) As Collection

    Dim oObj As cDepObj
    Dim oColumn As SQLDMO.Column
    Dim sqlObj As Object
    
    Dim retColl As New Collection
    
    If Not oList Is Nothing Then
        For Each oColumn In oList
            Set oObj = New cDepObj
            Set sqlObj = oColumn.Parent
            oObj.oSequence = 1
            oObj.oType = SQLDMOObj_UserTable
            oObj.oName = sqlObj.Name
            oObj.oOwner = sqlObj.Owner
            On Local Error Resume Next
            retColl.Add oObj, sqlObj.Name
            On Local Error GoTo 0
        Next
    End If
    Set oObj = Nothing
    
    Set LoadCollFromSQL_LIST = retColl
                         
End Function

Public Sub ShowDependent(ByVal sDbName As String, ByVal sObjName As String, ByVal objType As optLstAction)

    Dim collChild As Collection
    Dim collParent As Collection

    Dim lErr As Long, sErr As String
    
    Screen.MousePointer = vbHourglass
    
    lErr = ErrGetDependent(sDbName, sObjName, objType, collChild, collParent, sErr)
    
    If lErr = 0 Then
        fDepend.Stuff_Exec sObjName, collChild, collParent
        Screen.MousePointer = vbDefault
        fDepend.Show vbModal
    Else
        Screen.MousePointer = vbDefault
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
    Set collChild = Nothing
    Set collParent = Nothing

End Sub

