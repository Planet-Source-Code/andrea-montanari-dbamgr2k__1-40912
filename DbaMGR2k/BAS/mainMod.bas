Attribute VB_Name = "mainMod"
Option Explicit
'Declarations for Windows versions
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'End Declarations for Windows versions

Public Type udtLANG
    sLang As String
    idLang As Integer
End Type

Public Enum TypeEnum
    iNoEdit = -1
    iNormalEdit = 0
    iCombo
    iPush
    iEllipsis
End Enum
Public Enum am_sqlAccessType
    amAccessAll = 0
    amAccessDBO
    amAccessSingle
End Enum
Public Enum amCbo2Read
    amNone = 0
    amRestricAccess
    amRecovery
End Enum

Public Enum logPrivil
    'Server
    SysAdm = 0
    DBCreat = 1
    SetupAdm = 2
    DiskAdm = 3
    SecurAdm = 4
    SvrAdm = 5
    ProcAdm = 6
    'DataBase
    db_SecurAdm = 10
    db_BckOper = 11
    db_DDLAdmin = 12
    db_AccessAdmin = 13
    db_DataReader = 14
    db_dbo = 15
End Enum

Public Enum am_SqlPropTypeOwner
    am_OwnTable = 0
    am_OwnFunction
    am_OwnSP
    am_OwnView
    am_OwnUDT
    am_OwnRule
    am_OwnDefault
End Enum

Public Enum am_SqlPropType
    am_DbProp = 0
    am_TbProp
    am_ColProp
    am_IdxProp
    am_KeyProp
    am_TrigPropTB
    am_TrigPropVIEW
    am_ViewProp
    am_ViewColProp
    am_ViewIdxProp
    am_SpProp
    am_ProcParam
    am_FuncProp
    am_FuncParam
    am_UdtProp
End Enum

'larghezza Oggetti
Public Const k_FlexBevel = 100
Public Const k_MargineFRM = 200

'larghezza Flex
Public Const k_FlexPARAM_cellaSelector = 300

'lunghezza massima del LocigalName e PhysicalName in edit
Public Const k_MaxColumnLen = 256
Public Const k_MaxColumnDefaultLen = 1000
Public Const k_MaxDBLen = 128
Public Const k_MaxPath_DBLen = 260

Public Const kNEXT = "{Tab}"        'lancia in SendKeys
Public Const kPREV = "+{Tab}"       'al campo Prima o Dopo
Public Const kYEAR = "DD/MM/YYYY"
Public Const kShortYEAR = "DD/MM/YY"
Public Const kInvertYEAR = "MM/DD/YYYY"
Public Const kLNum = "#,0"
Public Const kCurNum = "#,0.00"
Public Const kLng = "LANGUAGE"
Public Const kEglishOffest = 2000

Public Const kErrBCP_TypeFMT = 13   'Tipo non corrispondente
Public Const kErrBCP_NotValidFMT = 321    'Formato di file non valido
Public Const kErrBCP_InvRec = 63    'Numero di record non valido
Public Const kErrProperty = 380     'Valore Invalido


            'SetWindowPos
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const GWL_STYLE = (-16)

            'GetSystemMetrics
Public Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on
                                    '  vertical scroll bar

            'SENDMESSAGE
Public Const WM_USER = &H400
Public Const CB_LIMITTEXT = &H141
Public Const CB_SHOWDROPDOWN = &H14F
Public Const LB_SETTABSTOPS = &H192
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const WM_SETREDRAW = &HB

Public Const kErr_NotFound = -2147199728
Public Const kErr_NotValidUser = -2147220588
Public Const kErrDB_Offline As Long = -2147199230

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempFileName Lib "KERNEL32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
'Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function ReadFile Lib "KERNEL32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function GetFileSize Lib "KERNEL32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function lread Lib "KERNEL32" Alias "_lread" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
Public Declare Function lopen Lib "KERNEL32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Public Declare Function lclose Lib "KERNEL32" Alias "_lclose" (ByVal hFile As Long) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Const kSqlServerSupported As Integer = 8
Public Const k_DbaMgr_Welcome1 As String = "DbaMgr2K is an alternative administrative Console for MSDE 1.0||CONDITION OF USE||THIS PROGRAM IS FREEWARE.|YOU USE THE PROGRAM AT YOUR OWN RISK - THE AUTHOR ACCEPTS NO RESPONSIBILITY FOR ANY PROBLEMS OR LOSSES WHICH MIGHT ARISE FROM ITS USE, AND SUPPORT FOR THE SOFTWARE IS NOT GUARANTEED.|||ALL KIND OF INFORMATION ARE PROVIDED ON A ""AS IS"" BASIS WITHOUT WARRANTY OF ANY KIND.  THE PRODUCER DISCLAIMS ALL WARRANTIES, EITHER EXPRESS OR IMPLIED, INCLUDING THE WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.  IN NO EVENT SHALL THE PRODUCER OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER INCLUDING DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, LOSS OF BUSINESS PROFITS OR SPECIAL DAMAGES, EVEN IF THE PRODUCER OR ITS SUPPLIERS HAVE BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.||"
Public Const k_DbaMgr_Welcome2 As String = "THE PRODUCER DOES NOT WARRANT THAT THE FUNCTIONS FOR THE LICENSED SOFTWARE FOR CODE CONTAINED IN THE PACKAGE WILL MEET YOUR REQUIREMENTS, OR THAT THE OPERATION OF THE LICENSED SOFTWARE OR CODE WILL BE UNINTERRUPTED OR ERROR-FREE, OR THAT DEFECTS IN THE LICENSED SOFTWARE OR CODE CAN BE CORRECTED.  FURTHERMORE, THE PRODUCER DOES NOT WARRANT OR MAKE ANY REPRESENTATIONS REGARDING THE USE OR THE RESULTS OF THE USE OF THE LICENSED SOFTWARE, CODE OR RELATED DOCUMENTATION IN TERMS OF THEIR CORRECTNESS, ACCURACY, RELIABILITY, OR OTHERWISE.  NO ORAL OR WRITTEN INFORMATION OR ADVICE GIVEN BY THE PRODUCER OR ITS AUTHORIZED REPRESENTATIVES SHALL CREATE A WARRANTY OR IN ANY WAY INCREASE THE SCOPE OF THIS WARRANTY.  SHOULD THE LICENSED SOFTWARE OR CODE PROVE DEFECTIVE AFTER THE PRODUCER HAS DELIVERED THE SAME, YOU, AND YOU ALONE, SHALL ASSUME THE ENTIRE COST ASSOCIATED WITH ALL NECESSARY SERVICING, REPAIR OR CORRECTION.||"
Public Const k_DbaMgr_Welcome3 As String = "NO LIABILITY FOR CONSEQUENTIAL DAMAGES. IN NO EVENT SHALL THE AUTHOR OR ITS SUPPLIERS BE LIABLE TO YOU FOR ANY CONSEQUENTIAL, SPECIAL, INCIDENTAL OR INDIRECT DAMAGES OF ANY KIND ARISING OUT OF THE DELIVERY, PERFORMANCE OR USE OF THE SOFTWARE, EVEN IF THE AUTHOR HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. IN NO EVENT WILL THE AUTHOR'S LIABILITY FOR ANY CLAIM, WHETHER IN CONTRACT, TORT OR ANY OTHER THEORY OF LIABILITY, EXCEED THE LICENSE FEE PAID BY YOU, IF ANY."

Public g_strLNG As String
Public Const kMaxIniSize As Integer = 1024

Public Const Err_Free = vbObjectError + 512 + 1

Public Const ErrDetachedDBPosition = -2147216381

Public objServer As SQLDMO.SQLServer2    'SQL DMO Server object

Public Enum optAction
    act_Null = 0
    act_Response = 1
    act_Computed = 2
    act_Server = 95
    act_DBTasks = 99
    act_DBOption = 100
    act_User = 200
    act_Tbl = 300
    act_Bck = 400
    act_Login = 500
    act_SP = 600
    act_Tri = 700
    act_View = 800
    act_Roles = 900
    act_UDT = 1000
    act_ATTACH = 1100
    act_DETACH = 1200
    act_ProcInfo = 130
    act_Q_Options = 1300
    act_User_Function = 1400
    act_Extended_VIEW_Idx_Trig = 1500
    act_Extended_TABLE_Idx_Rel = 1501
    act_Extended_TABLE_Trig = 1502
    act_Modify_Idx_Sql = 1550
End Enum
Public Enum optScript
    opt_Tbl = 110
    opt_View = 111
    opt_SP = 112
    opt_Def = 113
    opt_Rule = 114
    opt_UDT = 115
    opt_TRI = 116
    opt_Func = 117
End Enum
Public Enum optLstAction
    noAct = 0
    dbDrop
    dbDetach
    roleDrop
    tableDrop
    spDrop
    viewDrop
    udtDrop
    userDrop
    loginDrop
    triggerDROP
    spidDrop
    funcDrop
    EnumUserObj
End Enum

Public Enum amChangeOwner
    amChangeDB = 0
    amChangeObj
End Enum
Public Enum amChangingOwnerType
    amObjTable = 0
    amObjStoredProc
    amObjView
    amObjFunction
    'amObjDefault
    'amObjRule
End Enum

Public Enum amDetachingMode
    amCheckDb = 0
    amNoCheck
    amAskEach
End Enum


' Returns True is executed under Windows NT,
' False if executed under Windows 95/98
'Public Function IsWindowsNT() As Boolean
'    IsWindowsNT = (GetVersion() And &H80000000) = 0
'End Function

Public Function IsWindowsNT() As Boolean
    'By Yusuf Incekara  -   yincekara@rocketmail.com
    
    Dim os              As OSVERSIONINFO  ' receives version information
    Dim RetVal          As Long  ' return value
    Dim lPlatform       As Long
    
    IsWindowsNT = False
    os.dwOSVersionInfoSize = Len(os)  ' set the size of the structure
    RetVal = GetVersionEx(os)  ' read Windows's version information
    'MsgBox "Windows version number:" & os.dwMajorVersion & Chr(Asc(".")) & os.dwMinorVersion
    'MsgBox "OS Version Info Size = " & os.dwOSVersionInfoSize
    'MsgBox "BuildNumber = " & os.dwBuildNumber
    'MsgBox "Platform ID = " & os.dwPlatformId 'Note If ID =0 win 3.x , ID=1 win9x and ID =2 WINNT
    'MsgBox "CSD Version = " & os.szCSDVersion
    lPlatform = os.dwPlatformId 'Note If ID =0 win 3.x , ID=1 win9x and ID =2 WINNT
    If lPlatform = 2 Then IsWindowsNT = True
    
End Function

Public Function NomeComputer() As String

    Dim sBuffer As String
    
    sBuffer = String$(255, vbNullChar)
    If GetComputerName(sBuffer, 255&) <> 0 Then
        NomeComputer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        NomeComputer = MyLoadResString(k_Unknown)
    End If

End Function
Public Function CenterForm(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
End Function

Public Function GetIni(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal IniFile As String) As String
    
    Dim lpDefault As String, lpReturnString As String, Size As Long, Valid As Long
    
    lpDefault = ""
    lpReturnString = Space$(kMaxIniSize)
    Size = Len(lpReturnString$)
    Valid = GetPrivateProfileString(ByVal lpAppName, ByVal lpKeyName, ByVal lpDefault, ByVal lpReturnString, ByVal Size, ByVal IniFile)
    GetIni = Left$(lpReturnString, Valid)
    
End Function
Public Function WriteIni(lpAppName As String, lpKeyName As String, lpString As String, ByVal IniFile As String) As Long
    
    Dim lValid As Long

    lValid = WritePrivateProfileString(lpAppName, ByVal lpKeyName, ByVal lpString, ByVal IniFile)
    WriteIni = lValid
    
End Function

Public Sub OverStrike(ctrl As TextBox)

    If ctrl.SelLength = 0 Then
        If ctrl.MaxLength = 1 Then
            ctrl.SelStart = 0
        End If
        ctrl.SelLength = 1
    End If

End Sub

Public Function ApiciSQL(ByVal sBuf As String) As String
    ApiciSQL = Replace(sBuf, "'", "''")
End Function
Public Function SqlFinalS(ByVal sBuf As String) As String
    
    If Len(sBuf) Then SqlFinalS = "'" & sBuf & "'"
    
End Function

Public Function GetItem(ByVal sVal As String, cbo As ComboBox) As Integer

    Dim i As Integer
    sVal = UCase$(sVal)
    For i = 0 To cbo.ListCount - 1
        If StrComp(sVal, cbo.List(i), vbTextCompare) = 0 Then
            GetItem = i
            Exit Function
        End If
    Next
    GetItem = -1
    
End Function

Public Sub ComboLimit(ctrl As ComboBox, ByVal iMaxChar As Integer)

    Dim lRet As Long
    lRet = SendMessage(ctrl.hWnd, CB_LIMITTEXT, iMaxChar, ByVal 0&)

End Sub

Public Function GetTempFile(ByVal sPrefix As String) As String

    Dim sTempPath As String, sTempFileName As String
    Dim rc As Long

    sTempPath = Space$(128)
    rc = GetTempPath(128, sTempPath)
    sTempFileName = Space$(260)
    rc = GetTempFileName(sTempPath, sPrefix, 0, sTempFileName)
    GetTempFile = Left$(sTempFileName, Len(RTrim$(sTempFileName)) - 1)

End Function

Public Function sPad(ByVal bRight As Integer, ByVal sString As String, ByVal iLen As Integer) As String

    Dim sBuff As String
    
    sString = Trim$(sString)
    sBuff = Space$(iLen)
    If bRight Then
        RSet sBuff = sString
    Else
        LSet sBuff = sString
    End If
    sPad = sBuff

End Function
Public Function hasScrollBars(ByVal hWnd As Long) As Boolean

    Dim lFlag As Long, iGotScroll As Long
    
    lFlag = GetWindowLong(hWnd, GWL_STYLE)
    If (&H200000 And lFlag) <> False Then
        iGotScroll = True
    Else
        iGotScroll = False
    End If
    hasScrollBars = iGotScroll
    
End Function

Public Function Qry2Var(oQry As SQLDMO.QueryResults2, Optional ByVal bRetColName As Boolean = False) As Variant
    
    Dim iCols As Integer, lRows As Long
    Dim i As Integer, l As Long
    Dim lCurR As Long
    Dim varArray As Variant
    
    Debug.Print oQry Is Nothing
    If Not oQry Is Nothing Then
        If oQry.ResultSets <> 0 Then
            oQry.CurrentResultSet = 1
            iCols = oQry.Columns
            lRows = oQry.Rows - 1
            lRows = IIf(bRetColName, lRows + 1, lRows)
            If lRows >= 0 Then
            'ReDim varArray(1 To iCols, IIf(Not bRetColName, lRows, lRows - 1)) As Variant
                ReDim varArray(iCols - 1, lRows) As Variant
                                                                                                                          
                If bRetColName Then
                    For i = 1 To iCols - 1
                        varArray(i - 1, lCurR) = oQry.ColumnName(i)
                    Next
                    lCurR = 1
                End If
                    
                On Local Error Resume Next
                For l = 1 To lRows + 1
                     For i = 1 To iCols
                        Select Case oQry.ColumnType(i)
                            Case SQLDMO_DTypeBigint
                                varArray(i - 1, lCurR) = CCur(oQry.GetColumnBigInt(l, i))
                            Case SQLDMO_DTypeSQLVariant
                                                            
                                Debug.Print oQry.GetColumnString(l, i)
                                'varArray(i - 1, lCurR) = oQry.GetColumnSQLVARIANT(l, i)
                                varArray(i - 1, lCurR) = "!! sqlVariant !!"
                                
                                
'                                Debug.Print CCur(oQry.GetColumnDouble(l, i))
'                                Debug.Print CStr(oQry.GetColumnString(l, i))
'                                Debug.Print CDate(oQry.GetColumnDate(l, i))
'                                Debug.Print CBool(oQry.GetColumnBool(l, i))
'                                Debug.Print CCur(oQry.GetColumnBigInt(l, i))
                                'varArray(i - 1, lCurR) = CCur(oQry.GetColumnDouble(l, i))
                            Case SQLDMO_DTypeUChar, SQLDMO_DTypeChar, SQLDMO_DTypeUVarchar, SQLDMO_DTypeVarchar
                                varArray(i - 1, lCurR) = CStr(oQry.GetColumnString(l, i)) 'oQry.ColumnName(i)
                            Case SQLDMO_DTypeMoney, SQLDMO_DTypeMoney
                                varArray(i - 1, lCurR) = CCur(oQry.GetColumnDouble(l, i))
                            Case SQLDMO_DTypeFloat4, SQLDMO_DTypeFloat8
                                varArray(i - 1, lCurR) = CDbl(oQry.GetColumnDouble(l, i))
                            Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
                                varArray(i - 1, lCurR) = CDate(oQry.GetColumnDate(l, i))
                            
                            'Aggiunto Case 2 perche' altrimenti non carica colonne Identity
                            'anche se NON e' un valore listato tra le costanti
                            Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                                varArray(i - 1, lCurR) = CLng(oQry.GetColumnLong(l, i))
                            Case SQLDMO_DTypeBit
                                varArray(i - 1, lCurR) = CBool(oQry.GetColumnBool(l, i))
                        End Select
                    Next
                    lCurR = lCurR + 1
                Next
            End If
        End If
        On Local Error GoTo 0
    End If
'    If lCurR = 0 And IsEmpty(varArray(1, 0)) Then
'        varArray = vbEmpty
'    End If
    Qry2Var = varArray
        
End Function

Public Sub LoadPermission(ByVal sDbName As String, ByVal sObjectName As String, ByVal oType As optAction)

    Dim lErr As Long, sErr As String
    Dim oDb As Database2
    Dim sObj As String
    Dim bCancel As Boolean
    Dim oObj As Object
        
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    lErr = Err.Number
    If lErr = 0 Then
        Select Case oType
            Case act_SP
                oDb.StoredProcedures.Refresh True
                'Set oObj = oDB.StoredProcedures(sObjectName)
                Set oObj = GetSpByName(objServer, oDb.Name, sObjectName)
                sObj = MyLoadResString(k_Stored_Procedure)
            Case act_View
                oDb.Views.Refresh True
                'Set oObj = oDB.Views(sObjectName)
                Set oObj = GetViewByName(objServer, oDb.Name, sObjectName)
                sObj = MyLoadResString(k_View)
            Case act_Tbl
                oDb.Tables.Refresh True
                'Set oObj = oDB.Tables(sObjectName)
                Set oObj = GetTbByName(objServer, oDb.Name, sObjectName)
                sObj = MyLoadResString(k_Table)
            Case act_User
                oDb.Users.Refresh True
                Set oObj = oDb.Users(sObjectName)
                sObj = MyLoadResString(k_User)
            Case act_Roles
                Set oObj = oDb.DatabaseRoles(sObjectName)
                sObj = MyLoadResString(k_Database_Role)
            Case act_User_Function
                oDb.UserDefinedFunctions.Refresh True
                'Set oObj = oDB.Tables(sObjectName)
                Set oObj = GetFuncByName(objServer, oDb.Name, sObjectName)
                sObj = MyLoadResString(k_objFunction)
        End Select
        lErr = Err.Number
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        If oType <> act_Roles Then
            bCancel = oObj.SystemObject = True
        End If
        'If oObj.SystemObject = True Then
        If bCancel Then
            On Local Error Resume Next
            
            If oType = act_User Then
                lErr = Err_Free
                sErr = MyLoadResString(k_Non_Valid_Option_for_DbOwner)
            Else
                lErr = 425
                Err.Raise lErr, Description:=MyLoadResString(k_Unavailable_for_System_Objects)
                sErr = Err.Description
            End If
            On Local Error GoTo 0
        End If
    End If
        
    If lErr = 0 Then
        fPrivil.Caption = MyLoadResString(k_Permissions) & " - DB: [" & sDbName & "] " & sObj & " [" & sObjectName & "]"
        fPrivil.SetObj oDb, oObj, sObjectName, oType
        Screen.MousePointer = vbDefault
        fPrivil.Show vbModal
    End If
    Set oObj = Nothing
    Set oDb = Nothing
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName

End Sub

Public Function GetMembership(ByVal IsMember As logPrivil, Optional ByVal sDbName As String = "") As Boolean

    Dim bRet As Boolean
    If Not objServer Is Nothing Then
        If Not objServer.SaLogin Then
            Select Case IsMember
                Case SysAdm
                    bRet = objServer.Issysadmin
                Case DBCreat
                    bRet = objServer.Isdbcreator
                Case SetupAdm
                    bRet = objServer.Issetupadmin
                Case DiskAdm
                    bRet = objServer.Isdiskadmin
                Case SecurAdm
                    bRet = objServer.Issecurityadmin
                Case SvrAdm
                    bRet = objServer.Isserveradmin
                Case ProcAdm
                    bRet = objServer.Isprocessadmin
                Case Else
                    If Len(sDbName) <> 0 Then
                        On Local Error Resume Next
                        Select Case IsMember
                            Case db_dbo
                                bRet = objServer.Databases(sDbName).Isdb_owner
                            Case db_SecurAdm
                                bRet = objServer.Databases(sDbName).Isdb_securityadmin
                            Case db_BckOper
                                bRet = objServer.Databases(sDbName).Isdb_backupoperator
                            Case db_DDLAdmin
                                bRet = objServer.Databases(sDbName).Isdb_ddladmin
                            Case db_AccessAdmin
                                bRet = objServer.Databases(sDbName).Isdb_accessadmin
                            Case db_DataReader
                                bRet = objServer.Databases(sDbName).Isdb_datareader
                        End Select
                        On Local Error GoTo 0
                    End If
            End Select
        Else
            bRet = True
        End If
    End If
    GetMembership = bRet

End Function

Public Function GetShortPath(strFileName As String) As String
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(256, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, Len(strPath))
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function ReplaceMsg(ByVal sMsg As String, arr2Replace As Variant, arrReplace As Variant)
    
    Dim iLoop As Integer
    If IsArray(arr2Replace) And IsArray(arrReplace) Then
        For iLoop = LBound(arr2Replace) To UBound(arr2Replace)
            sMsg = Replace(sMsg, arr2Replace(iLoop), arrReplace(iLoop), Compare:=vbTextCompare)
        Next
    End If
    ReplaceMsg = sMsg
End Function
Public Function GetPackage() As String
    
    Dim sBuf As String
    Dim sSP As String
    
    Select Case objServer.IsPackage
        Case SQLDMO_MSDE
            sBuf = "MSDE"
        Case SQLDMO_OFFICE
            sBuf = "Office 2000"
        Case SQLDMO_STANDARD
            sBuf = "Standard"
        Case SQLDMO_ENTERPRISE
            sBuf = "Enterprise"
        Case Else
            sBuf = MyLoadResString(k_Unknown)   '"Motore non riconosciuto"
    End Select
    sBuf = "Microsoft SqlServer " & objServer.VersionMajor & "." & objServer.VersionMinor & " # " & sBuf & " # "
    
    sSP = objServer.VersionString
    If InStr(sSP, "8.00.679") Then
        sSP = " S.Pack 2 PostFix 2002/10/02"
    ElseIf InStr(sSP, "8.00.665") Then
        sSP = " S.Pack 2 PostFix 2002/08/14"
    ElseIf InStr(sSP, "8.00.655") Then
        sSP = " S.Pack 2 PostFix 2002/07/24"
    ElseIf InStr(sSP, "8.00.650") Then
        sSP = " S.Pack 2 PostFix 2002/07/10"
    ElseIf InStr(sSP, "8.00.608") Then
        sSP = " S.Pack 2 PostFix 2002/04/17"
    ElseIf InStr(sSP, "8.00.578") Then
        sSP = " S.Pack 2 PostFix 2002/02/20"
    ElseIf InStr(sSP, "8.00.534") Then
        sSP = " Service Pack 2"
    ElseIf InStr(sSP, "8.00.384") Then
        sSP = " Service Pack 1"
    ElseIf InStr(sSP, "8.00.194") Then
        sSP = " Retail"
    Else
        sSP = ""
    End If
        
    GetPackage = sBuf & sSP
    
End Function

Public Sub Main()

    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    On Local Error GoTo 0
    
    AbSplash.SplashOn frmMain, 3000
    
    
    Screen.MousePointer = vbDefault
    frmMain.Show
        
    'AbSplash.SplashOff

    
    
End Sub
Public Function MyLoadResString(ByVal iKey As Integer, Optional ByVal strFile As String = "") As String

    Dim sRet As String
    
    strFile = GetProperName(strFile)
    If Len(strFile) = 0 Then strFile = LngFileName()
    
    sRet = GetIni(kLng, Format$(iKey), strFile)
    On Local Error Resume Next
    
    If Len(sRet) = 0 Then
        sRet = VB.LoadResString(iKey + kEglishOffest)
        WriteIni kLng, Format$(iKey), sRet, strFile
    End If
    MyLoadResString = sRet
    
End Function

Public Function LngFileName() As String
    If Len(g_strLNG) = 0 Then g_strLNG = "English"
    LngFileName = App.Path & "\" & g_strLNG & ".ln2"
End Function
Public Function GetProperName(ByVal strFile As String) As String

    If Len(strFile) <> 0 Then
        strFile = App.Path & "\" & strFile & ".ln2"
        If Len(Dir$(strFile)) <> 0 Then GetProperName = strFile
    End If

End Function
Public Function ErrGetUsersMapping(ByVal sDbName As String, ByRef colDbUsers As Collection, ByRef sErr As String) As Long

    Dim lErr As Long
    
    Dim oDb As SQLDMO.Database2
    Dim i As Integer
    
    Dim sSql As String
    Dim oQry As SQLDMO.QueryResults
    Dim oVar As Variant
    Dim oDBUs As cDBUser
        
    Set colDbUsers = New Collection
    
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        
        sSql = "SELECT ms.name AS LoginName, db.name AS UserName , CASE WHEN ms.name IS NULL THEN 0 ELSE 1 END AS Mapped FROM master..syslogins AS ms RIGHT JOIN [1%]..sysusers AS db ON ms.sid=db.sid WHERE db.status=2 and db.hasdbaccess=1"
        sSql = ReplaceMsg(sSql, Array("1%"), Array(sDbName))

        Debug.Print sSql

        On Local Error Resume Next
        Set oQry = oDb.ExecuteWithResults(sSql)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        If lErr = 0 Then
            If Not oQry Is Nothing Then oVar = Qry2Var(oQry)
        End If
        Set oQry = Nothing
    
        If lErr = 0 Then
            For i = LBound(oVar, 2) To UBound(oVar, 2)
                Set oDBUs = New cDBUser
                oDBUs.LogName() = oVar(0, i)
                oDBUs.Name() = oVar(1, i)
                oDBUs.Mapped() = oVar(2, i)
                colDbUsers.Add oDBUs
            Next
            Set oDBUs = Nothing
        End If
    End If
    Set oDb = Nothing
    ErrGetUsersMapping = lErr
    
End Function

Public Function IsDbUser(ByVal sDbName As String) As Boolean

    Dim bRet As Boolean
    Dim oUser As SQLDMO.User
    
    If Not objServer.SaLogin Then
        On Local Error Resume Next
        Set oUser = objServer.Databases(sDbName).Users(objServer.Login)
        bRet = Err = 0
        On Local Error GoTo 0
    Else
        bRet = True
    End If
    Set oUser = Nothing
    
    IsDbUser = bRet
    
End Function
Public Sub NoPopUpMenu(ByRef txtBox As TextBox)

    txtBox.Enabled = False
    Beep
    MsgBox MyLoadResString(k_POP_UP_UNAVAILABLE), vbInformation
    txtBox.Enabled = True

End Sub
Public Function WhoIsDBO(ByVal oDb As SQLDMO.Database) As String

    Dim oUser As SQLDMO.User

    For Each oUser In oDb.Users
        If StrComp(oUser.Name, "dbo", vbTextCompare) = 0 Then
            WhoIsDBO = oUser.Login
            Exit For
        End If
    Next
    Set oUser = Nothing
        
End Function

Public Sub LoadExtendedProp(ByVal iObjType As am_SqlPropType, ByVal sDbName As String, Optional ByVal sOwner As String = "", Optional ByVal s1LevelObj As String = "", Optional ByVal s2LevelObj As String = "")

    Dim lErr As Long, sErr As String
        
    Dim fProperty As fSqlProp
    Set fProperty = New fSqlProp
    With fProperty
        Screen.MousePointer = vbHourglass
        .Stuff iObjType, sDbName, sOwner, s1LevelObj, s2LevelObj
        lErr = .Errors(sErr)
        If lErr = 0 Then
            Screen.MousePointer = vbDefault
            .Show vbModal
        End If
    End With
    Set fProperty = Nothing
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
End Sub
