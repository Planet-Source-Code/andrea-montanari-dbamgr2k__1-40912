Attribute VB_Name = "adoBas"
Option Explicit
Private Type typeErrore
    lNumber As Long
    sDescription As String
    sSource As String
    lNative As Long
End Type

Public Function ErrMyOpenConn(tmpConn As ADODB.Connection, _
                        ByVal sConnString As String, _
                        ByRef oErrore As cErrMgr, _
                        Optional ByVal Cursor As ADODB.CursorLocationEnum = adUseServer, _
                        Optional ByVal iTimeOut As Integer = 30) As Boolean

    Dim lErr As Long
    Dim intConnect As Integer
    
    intConnect = Val(GetIni("CONNECT", "TIMEOUT", App.Path & "\" & App.EXEName & ".INI"))
    If intConnect = 0 Then intConnect = "60"
    

    If tmpConn Is Nothing Then Set tmpConn = New ADODB.Connection
    Err.Clear

    'sConnString = "SERVER=" & ServerName & ";" & sConnString
    'sConnString = "SERVER=" & NomeComputer() & ";" & sConnString
    
    sConnString = "SERVER=" & objServer.Name & ";" & sConnString
    
    On Local Error Resume Next
    With tmpConn
        .ConnectionString = sConnString
        .ConnectionTimeout = intConnect
        .CursorLocation = Cursor
        .CommandTimeout = iTimeOut
        .Provider = "SQLOLEDB"
        .Mode = adModeReadWrite
        .Open
    End With
    lErr = Err.Number
    Debug.Print Err.Description
    Debug.Print sConnString
    If lErr <> 0 Or tmpConn.State = adStateClosed Then
        If Not oErrore Is Nothing Then
            oErrore.PushError Err.Number, Err.Description, Err.Source, 0
            GetConnError tmpConn, oErrore
        End If
    End If
    
    ErrMyOpenConn = lErr <> 0

End Function
Public Sub GetConnError(tmpConn As ADODB.Connection, ByRef oErrore As cErrMgr)

    Dim objErr As ADODB.Error
            
    If Not tmpConn Is Nothing Then
        For Each objErr In tmpConn.Errors
            oErrore.PushError objErr.Number, objErr.Description, objErr.Source, objErr.NativeError
        Next
    End If
    tmpConn.Errors.Clear
    
End Sub

Public Function ExceptionHandler(ByRef objErr As cErrMgr, ByRef sRetString As String, Optional ByVal bDisplay As Boolean = True) As Integer

    Dim i As Integer, iRet As Integer
    Dim sErr As String, sMsg As String, bFirstLine As Boolean

    Dim oErr As typeErrore
    Dim vErr As Variant
        
    Dim szBuf As String
        
    bFirstLine = True
    
    On Local Error Resume Next
    For i = 1 To objErr.Count
        vErr = objErr.PopError(i)
        If IsArray(vErr) Then
            If i > 1 Then
                If vErr(errNum) = oErr.lNumber Then Exit For
            End If
            If bFirstLine Then
                bFirstLine = False
            Else
                sErr = sErr & vbLf
            End If
                
            oErr.lNumber = vErr(errNum)
            oErr.sDescription = vErr(errDescription)
            oErr.sSource = vErr(errSource)
            oErr.lNative = vErr(errNative)
        
            sMsg = oErr.sDescription
            szBuf = IIf(oErr.lNative <> 0, " - Native " & oErr.lNative, "")
            sErr = sErr & "Num. " & oErr.lNumber & szBuf & " - " & sMsg
        End If
    Next
    Set objErr = New cErrMgr

    sRetString = sErr
    If bDisplay Then
        Screen.MousePointer = vbDefault
        If Len(sErr) Then
            Beep
            iRet = MsgBox(sErr, vbInformation, App.EXEName & " - Errore")
            ExceptionHandler = iRet
        End If
    End If
End Function

