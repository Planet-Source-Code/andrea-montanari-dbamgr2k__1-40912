VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDetGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type am_DetachedFiles
    sFileName As String
    bPrimary As Boolean
End Type
Private m_bIsLocalHost As Boolean
Private sPrimaryMDFfile As String

Private m_lErr As Long
Private m_sErr As String
Private m_lWarning As Long
Private m_sWarning As String

Public m_FileColl As Collection

Public Property Get Warnings(ByRef sWarning As String) As Long
    Warnings = m_lWarning
    sWarning = m_sWarning
End Property
Public Property Get Errors(ByRef sError As String) As Long
    Errors = m_lErr
    sError = m_sErr
End Property
Public Property Get AreWarnings() As Boolean
    AreWarnings = m_lWarning <> 0
End Property

Public Property Get AreErrors() As Boolean
    AreErrors = m_lErr <> 0
End Property
Public Property Let PrimaryMdfFile(ByVal sFullPathFile As String)
    sPrimaryMDFfile = sFullPathFile
    ClearWarnings
    ClearError
    LoadCollection
End Property
Public Property Get PrimaryMdfFile() As String
    PrimaryMdfFile = sPrimaryMDFfile
End Property

Private Sub Class_Initialize()
    Set m_FileColl = New Collection
    m_bIsLocalHost = StrComp(objServer.NetName, NomeComputer(), vbTextCompare) = 0
End Sub
Private Sub Class_Terminate()
    Set m_FileColl = Nothing
End Sub
Private Sub LoadCollection()

    Dim oList As SQLDMO.NameList
    Dim i As Integer, iPrimaryFile As Integer, iItem As Integer
    Dim udtFiles() As am_DetachedFiles
    Dim lErr As Long, sErr As String
    Dim lWarning As Long, sTmpWarning As String
    Dim bIsPrimary As Boolean

    On Local Error Resume Next
    Set oList = objServer.ListDetachedDBFiles(sPrimaryMDFfile)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If Not oList Is Nothing And lErr = 0 Then
        iItem = -1
        iPrimaryFile = -1
        For i = 1 To oList.Count
            iItem = iItem + 1
            ReDim Preserve udtFiles(iItem) As am_DetachedFiles
            udtFiles(iItem).sFileName = oList.Item(i)
            
            On Local Error Resume Next
            bIsPrimary = objServer.IsDetachedPrimaryFile(udtFiles(iItem).sFileName)
            lWarning = Err.Number
            sTmpWarning = Err.Description
            On Local Error GoTo 0
            
            If lWarning <> 0 Then
                If lWarning = ErrDetachedDBPosition Then
                    sTmpWarning = sTmpWarning & vbCrLf & ReplaceMsg(MyLoadResString(k_DbDetached_PrimaryFileMoved), Array("1%", "|"), Array(udtFiles(iItem).sFileName, vbCrLf))
                End If
                m_lWarning = lWarning
                m_sWarning = IIf(Len(m_sWarning) <> 0, m_sWarning & vbCrLf, "") & sTmpWarning
                If i = 1 Then
                    udtFiles(iItem).sFileName = sPrimaryMDFfile
                End If
            End If
            If bIsPrimary Then
                udtFiles(iItem).bPrimary = True
                iPrimaryFile = iItem
            End If
        Next

        If lErr = 0 Then
            Set oList = objServer.ListDetachedLogFiles(sPrimaryMDFfile)
            For i = 1 To oList.Count
                iItem = iItem + 1
            ReDim Preserve udtFiles(iItem) As am_DetachedFiles
                udtFiles(iItem).sFileName = oList.Item(i)
            Next
            If lWarning <> 0 And iPrimaryFile = -1 Then iPrimaryFile = 0

            If iItem <> -1 And iPrimaryFile <> -1 Then
                Me.AddFile2Col udtFiles(iPrimaryFile).sFileName, True
                
                If m_lErr = 0 Then
                    For i = 0 To UBound(udtFiles)
                        If i <> iPrimaryFile Then Me.AddFile2Col udtFiles(i).sFileName, True
                    Next
                End If
            End If
        End If
    End If
    
    If lErr <> 0 Then
        m_lErr = IIf(lErr = 0, Err_Free, lErr)
        m_sErr = IIf(lErr = 0, MyLoadResString(k_DbDetached_NoDbFile), sErr)
    End If
    
End Sub
Public Sub AddFile2Col(ByVal sFileName As String, Optional ByVal bTestForError As Boolean = False)

    Dim sCommand As String
    Dim vArrFile As Variant
    Dim FI As CFileInfo
    Dim oQry As SQLDMO.QueryResults
    Dim lErr As Long, sErr As String
    
    Me.ClearError
    
    sCommand = Replace("EXEC xp_fileexist N'1%'", "1%", sFileName)
    
    On Local Error Resume Next
    Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
        
    If Not oQry Is Nothing Then vArrFile = Qry2Var(oQry)
    Set oQry = Nothing
        
    If IsArray(vArrFile) Then
        If CBool(vArrFile(1, 0)) = 0 Then
            Dim oFile As cDetacDB
            Set oFile = New cDetacDB
            
            oFile.PhysicalName() = sFileName
            oFile.FileExist() = CBool(vArrFile(0, 0))
            
            If m_bIsLocalHost And oFile.FileExist() Then
                Set FI = New CFileInfo
                FI.FullPathName() = sFileName
                Set oFile.oFileFinfo() = FI
                
                Set FI = Nothing
            End If
            
            On Local Error Resume Next
            m_FileColl.Add oFile, sFileName
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
        End If
    End If
    Set FI = Nothing
    Set oFile = Nothing

    If bTestForError And lErr <> 0 Then
        m_lErr = lErr
        m_sErr = sErr
    End If
    
End Sub
Public Sub ClearError()
    m_lErr = 0
    m_sErr = ""
End Sub
Public Sub ClearWarnings()
    m_lWarning = 0
    m_sWarning = ""
End Sub
