VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fBr4Fold 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "fBr4Fold.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tSelected 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4920
      Width           =   4575
   End
   Begin VB.ComboBox cboSelect 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox tFile 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   3960
      Width           =   4575
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   360
      Index           =   0
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   5760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBr4Fold.frx":000C
            Key             =   "fClose"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBr4Fold.frx":27C0
            Key             =   "fOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBr4Fold.frx":4F74
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tView 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6588
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "iList"
      Appearance      =   1
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected File:"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "cbo"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "&File Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "fBr4Fold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'---    Form da utilizzare al posto
'---    di SHBrowseForFolder per
'---    ricerche/selezioni sul
'---    Server Sql invece che
'---    sulla macchina locale
'------------------------------------
Option Explicit
Public Enum enB4Folder
    enB4isSelectDir = 0
    enB4isSelectFile
    enB4SelectFileFilters
End Enum
Private Enum enNodeType
    enDisk = 0
    enFolder
    enFile
End Enum

Private m_lErrInit As Long
Private m_sErrorMsg As String

Private m_sLikeExt As String

Private m_enType As enB4Folder
Private mbAccept As Boolean
Private m_sInitialPath As String
Private m_sInitialFile As String
Private m_bInitLoad As Boolean
Private m_vntComboFilter As Variant
Private m_bFileMustExist As Boolean
Private m_bCanNotChangeDir As Boolean

Public Property Let FileMustExist(ByVal bVal As Boolean)
    m_bFileMustExist = bVal
    tFile.Locked = m_bFileMustExist
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = Not mbAccept
End Property
Public Property Get ErrLoadingInit() As Boolean
    ErrLoadingInit = m_lErrInit <> 0
End Property
Public Sub GetError(ByRef lErrorCode As Long, ByRef sErrorMsg As String)
    lErrorCode = m_lErrInit
    sErrorMsg = m_sErrorMsg
End Sub


Private Sub cboSelect_Click()

    If Not m_bInitLoad Then
        If cboSelect.ListCount <> 0 Then
            m_sLikeExt = m_vntComboFilter(0, cboSelect.ListIndex)
            tView_NodeClick tView.SelectedItem
        End If
    End If
    
End Sub

Private Sub cmd_Click(index As Integer)

    Dim bStop As Boolean
    If cmd(index).Enabled Then
        If index = 1 Then
'            If m_enType <> enB4isSelectDir Then
                bStop = ErrTestExistFinalDir()
                If bStop Then Exit Sub
                mbAccept = True
 '           Else
 '               If Len(Trim$(tSelected.Text)) <> 0 Then
 '               mbAccept = True
'            End If
        End If
        Me.Hide
    End If
End Sub
Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fBr4Fold = Nothing
End Sub

Public Sub Init(ByVal iBrowseType As enB4Folder, ByVal sCaption As String, Optional ByVal sInitPath As String = "", Optional ByVal sInitFile As String = "", Optional ByVal vntFilters As Variant = vbEmpty)
    
    If objServer Is Nothing Then
        m_lErrInit = Err_Free
        m_sErrorMsg = MyLoadResString(k_B4F_SqlIsNothing)
        Exit Sub
    End If
    m_enType = iBrowseType
    m_sInitialPath = sInitPath
    m_sInitialFile = sInitFile
    
    
    Me.Caption = sCaption
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_B4F_Ok)
    
    If iBrowseType <> enB4isSelectDir Then
        If IsArray(vntFilters) And iBrowseType = enB4SelectFileFilters Then
            m_vntComboFilter = vntFilters
            lbl(1).Visible = True
            cboSelect.Visible = True
            lbl(1).Caption = MyLoadResString(k_B4F_FileName)
        Else
            m_enType = enB4isSelectFile
        End If
    End If
        
    lbl(0).Caption = MyLoadResString(k_B4F_FileName)
    If m_enType = enB4isSelectFile Then
        lbl(2).Caption = MyLoadResString(k_B4F_SelFile)
    Else
        lbl(2).Caption = MyLoadResString(k_B4F_SelPath)
    End If
    tView.ToolTipText = Replace(MyLoadResString(k_B4F_TviewToolTip), "1%", objServer.NetName)
    
    If m_enType = enB4isSelectDir Then
        lbl(0).Visible = False
        tFile.Visible = False
        lbl(1).Visible = False
        cboSelect.Visible = False
        tView.Height = lbl(2).Top - tView.Top - 120
    End If
    
    Screen.MousePointer = vbHourglass
    m_bInitLoad = True
    LoadTview
    If Len(sInitPath) <> 0 Then
        SearchPath sInitPath
    End If
    m_bInitLoad = False
    Screen.MousePointer = vbDefault
    
End Sub
Public Property Get FullPath() As String
    FullPath = tSelected.Text
End Property
Public Property Get FileName() As String
    FileName = Trim$(tFile.Text)
End Property
Public Property Get FolderPath() As String

    Dim iPos As Integer
    Dim sRet As String
    
    sRet = tSelected.Text
    iPos = InStrRev(sRet, "\")
    If iPos <> 0 Then sRet = Left$(sRet, iPos - 1)
    FolderPath = sRet

End Property

Private Sub LoadTview()

    Dim i As Integer
    Dim vArrMedia As Variant
    Dim vArrDir As Variant
    Dim vArrFiles As Variant
    Dim lErr As Long, sErr As String
    Dim sCurDir As String
    Dim sCommand As String
    
    Dim oQry As SQLDMO.QueryResults
    On Local Error Resume Next
    Set oQry = objServer.EnumAvailableMedia '(SQLDMOMedia_FixedDisk Or SQLDMOMedia_CDROM Or SQLDMOMedia_Tape)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If m_enType = enB4SelectFileFilters Then
        For i = 0 To UBound(m_vntComboFilter, 2)
            cboSelect.AddItem m_vntComboFilter(1, i)
        Next
        cboSelect.ListIndex = 0
        m_sLikeExt = m_vntComboFilter(0, 0)
    End If
    
    If lErr = 0 Then
        vArrMedia = Qry2Var(oQry)
        If IsArray(vArrMedia) Then
            BB_Sort vArrMedia
            LoadNode vArrMedia, Array(0), Array(0), enDisk
                
            For i = 0 To UBound(vArrMedia, 2)
                vArrDir = vbEmpty
                sCurDir = vArrMedia(0, i)
                If Right$(sCurDir, 1) = "\" Then sCurDir = Left$(sCurDir, Len(sCurDir) - 1)
                
                On Local Error Resume Next
                Set oQry = objServer.EnumDirectories(sCurDir)
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
                
                If lErr = 0 Then vArrDir = Qry2Var(oQry)
                If IsArray(vArrDir) Then
                    BB_Sort vArrDir
                    LoadNode vArrDir, Array(0), Array(0), enFolder, vArrMedia(0, i)
                    
                    If m_enType <> enB4isSelectDir Then
                        sCommand = Replace("EXECUTE dbo.xp_dirtree N'1%',1,1", "1%", vArrMedia(0, i))
                        On Local Error Resume Next
                        Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
                        lErr = Err.Number
                        sErr = Err.Description
                        On Local Error GoTo 0
                        If lErr = 0 Then vArrFiles = Qry2Var(oQry)
                        If IsArray(vArrFiles) Then
                            BB_Sort vArrFiles
                            PurgeFilesFromDirs vArrFiles, vArrDir
                            LoadNode vArrFiles, Array(0), Array(0), enFile, vArrMedia(0, i)
                        End If
                        
                    End If
                    On Local Error Resume Next
                    tView.Nodes(vArrMedia(0, i)).Tag = "+"
                    On Local Error GoTo 0
                End If
            Next
        End If
    Else
        m_lErrInit = lErr
        m_sErrorMsg = sErr
    End If
    Set oQry = Nothing
    
End Sub
Private Sub LoadNode(ByVal varArray As Variant, ByVal vntItem4Key As Variant, ByVal vntItem2Add As Variant, ByVal iType As enNodeType, Optional ByVal sRootKey As String = "")

    Dim iRows As Integer
    Dim iCols As Integer
    Dim iCurRow As Integer
    Dim iCurCol As Integer
    Dim sParent As String
    Dim nNode As Node
    Dim sItem As String
    Dim sText As String
    Dim sKey As String
    Dim sParKey As String
    Dim sImage As String
    Dim nNodePar As Node
    Dim bAdd As Boolean
    
    If Not IsArray(varArray) Then Exit Sub
    iRows = UBound(varArray, 2)
    iCols = UBound(varArray, 1)
    
    If iType = enDisk Then
        sImage = "fOpen"
    ElseIf iType = enFolder Then
        sImage = "fClose"
    Else
        sImage = "File"
    End If
    
    If Len(sRootKey) <> 0 Then
        sParent = UCase$(sRootKey)
        On Local Error Resume Next
        Set nNodePar = tView.Nodes.Item(sRootKey)
        On Local Error GoTo 0
        If Not nNodePar Is Nothing Then
            sParKey = nNodePar.Key
        Else
            Exit Sub
        End If
    End If
    
    For iCurRow = 0 To iRows
        If iType = enDisk Then
            
            'bAdd = Not (varArray(3, iCurRow) = SQLDMOMedia_CDROM )And Not (varArray(3, iCurRow) =SQLDMOMedia_Floppy)
            bAdd = Not (varArray(3, iCurRow) = SQLDMOMedia_Floppy)
            Debug.Print varArray(0, iCurRow), bAdd
        Else
            bAdd = True
        End If
        
        If bAdd Then
            sText = ""
            sKey = sParKey
            For iCurCol = 0 To UBound(vntItem4Key, 1)
                sItem = varArray(vntItem4Key(iCurCol), iCurRow)
                sKey = sKey & sItem
            Next
            For iCurCol = 0 To UBound(vntItem2Add, 1)
                sItem = varArray(vntItem2Add(iCurCol), iCurRow)
                sText = sText & sItem
            Next
            If iType <> enFile Then
                If Right$(sKey, 1) <> "\" Then sKey = sKey & "\"
            End If
    
            If Len(sRootKey) <> 0 Then
                Set nNode = tView.Nodes.Add(sParent, tvwChild, UCase$(sKey))
            Else
                Set nNode = tView.Nodes.Add(, , UCase$(sKey))
            End If

            With nNode
                .Text = sText
                .Image = sImage
                .Expanded = True
                
                .Tag = IIf(iType <> enFile, "-", "F")
            End With
        End If
    Next

End Sub
Private Sub SearchPath(ByVal sPath2Search As String)

    Dim sDirs() As String
    Dim i As Integer
    Dim nNode As Node
    Dim sKey As String
    Dim bErr As Boolean
    
    sKey = ""
    sPath2Search = UCase$(sPath2Search)
    sDirs() = Split(sPath2Search, "\")
    For i = 0 To UBound(sDirs)
        If Len(sDirs(i)) Then
            sKey = sKey & sDirs(i) & "\"
            On Local Error Resume Next
            Set nNode = tView.Nodes(sKey)
            On Local Error GoTo 0
            If Not nNode Is Nothing Then
                nNode.Selected = True
                tView_NodeClick nNode
            Else
                bErr = True
                Exit For
            End If
        End If
    Next
    If Not bErr Then
        On Local Error Resume Next
        Set nNode = tView.Nodes(UCase$(sPath2Search))
        On Local Error GoTo 0
        If Not nNode Is Nothing Then
            nNode.Selected = True
            nNode.EnsureVisible
        End If
    End If
    
End Sub

Private Sub tFile_Change()
    
    Dim sFile As String
    Dim nNode As Node
    
    Set nNode = tView.SelectedItem
    If nNode.Image = "File" Then
        sFile = nNode.Parent.Key & nNode.Text
    Else
        sFile = nNode.Key & Trim$(tFile.Text)
    End If
    tSelected.Text = sFile
End Sub

Private Sub tFile_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyEscape, 92   '"\"
            KeyAscii = 0
    End Select
End Sub

Private Sub tView_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim sParent As String
    Dim sCurrent As String
    Dim sKey As String
    Dim oQry As SQLDMO.QueryResults
    Dim lErr As Long, sErr As String
    Dim vArrDir As Variant
    Dim vArrFiles As Variant
    Dim iMousePointer As Integer
    Dim sCommand As String
    
    Const kNoMoreDepth = -2147199498
    
    sKey = Node.Key
    tFile.Text = ""
    tSelected.Text = Node.Key
    
    iMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    If Node.Image <> "File" Then
    
        If Node.Tag = "-" Then
            On Local Error Resume Next
            Set oQry = objServer.EnumDirectories(sKey)
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
        
            If lErr = 0 Then vArrDir = Qry2Var(oQry)
            If IsArray(vArrDir) Then
                BB_Sort vArrDir
                LoadNode vArrDir, Array(0), Array(0), enFolder, Node.Key
                If tView.Nodes(Node.Key).Image = "fClose" Then tView.Nodes(Node.Key).Image = "fOpen"
            Else
                If lErr = 0 Then lErr = kNoMoreDepth
            End If
            If m_enType <> enB4isSelectDir Then
                If tView.Nodes(Node.Key).Tag = "-" Then
                    sCommand = Replace("EXECUTE dbo.xp_dirtree N'1%',1,1", "1%", Node.Key)
                    On Local Error Resume Next
                    Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    If lErr = 0 Then vArrFiles = Qry2Var(oQry)
                    If IsArray(vArrFiles) Then
                        BB_Sort vArrFiles
                        PurgeFilesFromDirs vArrFiles, vArrDir
                        LoadNode vArrFiles, Array(0), Array(0), enFile, Node.Key
                    End If
                End If
            End If
            tView.Nodes(Node.Key).Tag = "+"
        Else
            ReloadCurrent sKey 'Padre
        End If
    Else
        sCurrent = Node.Key
        sParent = tView.Nodes(Node.Key).Parent.Key
        ReloadCurrent sParent, sCurrent
        
'        If Len(m_sLikeExt) <> 0 Then
'            sParent = tView.Nodes(Node.Key).Parent.Key
'            Set bNode = tView.Nodes(Node.Key).Parent.Child
'            While Not bNode Is Nothing
'                sKey = bNode.Key
'                tView.Nodes.Remove sKey
'                Set bNode = bNode.Next
'            Wend
'
'            On Local Error Resume Next
'            Set oQry = objServer.EnumDirectories(sParent)
'            lErr = Err.Number
'            sErr = Err.Description
'            On Local Error GoTo 0
'
'            If lErr = 0 Then vArrDir = Qry2Var(oQry)
'            If IsArray(vArrDir) Then
'                BB_Sort vArrDir
'                LoadNode vArrDir, Array(0), Array(0), enFolder, sParent
'                If tView.Nodes(Node.Key).Image = "fClose" Then tView.Nodes(Node.Key).Image = "fOpen"
'            Else
'                If lErr = 0 Then lErr = kNoMoreDepth
'            End If
'
'            sCommand = Replace("EXECUTE dbo.xp_dirtree N'1%',1,1", "1%", sParent)
'            On Local Error Resume Next
'            Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
'            lErr = Err.Number
'            sErr = Err.Description
'            On Local Error GoTo 0
'            If lErr = 0 Then vArrFiles = Qry2Var(oQry)
'            If IsArray(vArrFiles) Then
'                BB_Sort vArrFiles
'                PurgeFilesFromDirs vArrFiles, vArrDir
'                LoadNode vArrFiles, Array(0), Array(0), enFile, sParent
'            End If
'        End If
'        On Local Error Resume Next
'        Set bNode = tView.Nodes(sCurrent)
'        lErr = Err.Number
'        On Local Error GoTo 0
        
'        If lErr = 0 Then
'            tFile.Text = Node.Text
'        Else
'            tFile.Text = ""
'            lErr = 0
'        End If
        
        tFile.Text = Node.Text
    End If
    Set oQry = Nothing
    Screen.MousePointer = iMousePointer
    
    If Not m_bInitLoad And lErr <> 0 Then
        If lErr <> kNoMoreDepth Then
            Screen.MousePointer = vbDefault
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
        Else
            If tView.Nodes(Node.Key).Image = "fClose" Then tView.Nodes(Node.Key).Image = "fOpen"
        End If
    End If
    
End Sub
Private Sub PurgeFilesFromDirs(ByRef vArrFiles As Variant, ByVal vArrDir As Variant)
    
    Dim lLoop As Long
    Dim vRet As Variant
    Dim lPosInArray As Long
    Dim iType As Integer
    Dim bIsArrayDir As Boolean
    Dim bIsLike As Boolean
    Dim bAdd As Boolean
    Dim sLikeFilter As String
    

    sLikeFilter = UCase$(m_sLikeExt)
    bIsArrayDir = IsArray(vArrDir)
    bIsLike = Len(m_sLikeExt) <> 0
        
    For lLoop = 0 To UBound(vArrFiles, 2)
        Debug.Print vArrFiles(2, lLoop), vArrFiles(0, lLoop)
        If bIsArrayDir Then
            bAdd = False
            lPosInArray = BinarySearch(vArrFiles(0, lLoop), vArrDir, 0)
            If lPosInArray = -1 Then
                If InStr(vArrFiles(0, lLoop), ".") <> 0 Then bAdd = True
                If Not bAdd Then
                    iType = vArrFiles(2, lLoop)
                    If iType = 1 Then bAdd = True
                End If
            End If
        Else
            bAdd = True
        End If
        If bAdd Then
            If bIsLike Then
                bAdd = UCase$(vArrFiles(0, lLoop)) Like sLikeFilter
            End If
        End If
        If bAdd Then vRet = Add2VarArray(Array(vArrFiles(0, lLoop)), vRet)
    Next
    vArrFiles = vRet

End Sub
Private Function BinarySearch(ByVal sKey As String, ByVal vArrDir As Variant, Optional ByVal iItem2Test As Integer = 0) As Long

    Dim iComp As Integer
    'Out:
    'Index in the array of the string if found
    Dim rlIndex As Long
    'Local variables:
    'Index in the array
    Dim lnIdx As Long
    'Lower bound of the search interval
    Dim lnMin As Long
    'Upper bound of the search interval
    Dim lnMax As Long
    'Return an invalid index, if string is not found
    rlIndex = -1
    lnMax = UBound(vArrDir, 2)
    lnMin = LBound(vArrDir, 2)
    'lookup vsName in rasArray()
    Do While lnMin <= lnMax
        lnIdx = (lnMax + lnMin) \ 2
        
                
        iComp = StrComp(sKey, vArrDir(iItem2Test, lnIdx), vbTextCompare)
        If iComp = 0 Then
            rlIndex = lnIdx
            Exit Do
        ElseIf iComp < 0 Then
            lnMax = lnIdx - 1
        Else
            lnMin = lnIdx + 1
        End If
    Loop
        
    BinarySearch = rlIndex

End Function

Private Sub ReloadCurrent(ByVal sParentKey As String, Optional ByVal sCurrentKey As String = "")

    Dim bNode As Node
    Dim bErr As Boolean
    Dim oQry As SQLDMO.QueryResults
    Dim vArrDir As Variant
    Dim vArrFiles As Variant
    Dim bNoSubDir As Boolean
    Dim sCommand As String
    Dim lErr As Long, sErr As String
    Dim sKey As String
    Dim sDirKey As String
    
    If Len(m_sLikeExt) <> 0 Then
        If tView.Nodes(sParentKey).Children Then
            Set bNode = tView.Nodes(sParentKey).Child
            While Not bNode Is Nothing
                sKey = bNode.Key
                Set bNode = bNode.Next
                tView.Nodes.Remove sKey
            Wend
        End If
        
        If Len(sParentKey) = 3 Then
            sDirKey = Left$(sParentKey, 2)
        Else
            sDirKey = sParentKey
        End If
        
        On Local Error Resume Next
        Set oQry = objServer.EnumDirectories(sDirKey)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        If Not bErr Then vArrDir = Qry2Var(oQry)
        If IsArray(vArrDir) Then
            BB_Sort vArrDir
            LoadNode vArrDir, Array(0), Array(0), enFolder, sParentKey
            If tView.Nodes(sParentKey).Image = "fClose" Then tView.Nodes(sParentKey).Image = "fOpen"
        Else
            bNoSubDir = True
        End If
            
        sCommand = Replace("EXECUTE dbo.xp_dirtree N'1%',1,1", "1%", sParentKey)
        On Local Error Resume Next
        Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        If lErr = 0 Then vArrFiles = Qry2Var(oQry)
        If IsArray(vArrFiles) Then
            BB_Sort vArrFiles
            PurgeFilesFromDirs vArrFiles, vArrDir
            LoadNode vArrFiles, Array(0), Array(0), enFile, sParentKey
        End If
        If Len(sCurrentKey) <> 0 Then
            On Local Error Resume Next
            Set bNode = tView.Nodes(sCurrentKey)
            lErr = Err.Number
            On Local Error GoTo 0
            If lErr = 0 Then
                bNode.Selected = True
                bNode.EnsureVisible
            End If
            lErr = 0
        End If
    End If
    Set oQry = Nothing
    If lErr <> 0 Then MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)

End Sub
Private Function ErrTestExistFinalDir() As Boolean

    Dim sDir2Test As String
    Dim iPos As Integer
    Dim lErr As Long, sErr As String
    Dim bRet As Boolean
    Dim oQry As SQLDMO.QueryResults
    Dim sCommand As String
    Dim vVar As Variant
    Dim iItem2Test As Integer
    
    Screen.MousePointer = vbHourglass
    sDir2Test = tSelected.Text
    
    If m_enType <> enB4isSelectDir Then
        tFile.Text = Trim$(tFile.Text)
        If Len(tFile.Text) = 0 Then
            lErr = Err_Free
            sErr = MyLoadResString(k_B4F_Err_NoFileSelected)
        End If
    Else
        tSelected.Text = Trim$(tSelected.Text)
        If Len(tSelected.Text) = 0 Then
            lErr = Err_Free
            sErr = MyLoadResString(k_B4F_Err_NoDirSelected)
        End If
    End If
    
    If lErr = 0 Then
        If m_enType <> enB4isSelectDir Then
            If Not m_bFileMustExist Then
                iItem2Test = 1
                iPos = InStrRev(sDir2Test, "\")
                If iPos <> 0 Then sDir2Test = Left$(sDir2Test, iPos)
            Else
                iItem2Test = 0
            End If
            
            If m_bCanNotChangeDir Then
                If StrComp(sDir2Test, m_sInitialPath, vbTextCompare) <> 0 Then
                    lErr = Err_Free
                    sErr = Replace(MyLoadResString(k_B4F_Err_CantChangeDir), "1%", sDir2Test)
                End If
            End If
        End If
    End If
    
    If lErr = 0 Then
        If m_enType <> enB4isSelectDir Then
            sCommand = Replace("EXECUTE xp_fileexist N'1%'", "1%", sDir2Test)
            On Local Error Resume Next
            Set oQry = objServer.ExecuteWithResults(sCommand, Len(sCommand))
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0
            If lErr = 0 Then vVar = Qry2Var(oQry)
            If IsArray(vVar) Then
                bRet = vVar(iItem2Test, 0) <> 0
            End If
        End If
    End If
    If m_enType <> enB4isSelectDir Then
        If lErr = 0 And Not bRet Then
            lErr = Err_Free
            If Not m_bFileMustExist Then
                sErr = MyLoadResString(k_B4F_Err_DirMustExists)
            Else
                sErr = MyLoadResString(k_B4F_Err_FileMustExists)
            End If
        End If
    Else
        bRet = lErr = 0 And Len(tSelected.Text) <> 0
    End If
    Set oQry = Nothing
    
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        bRet = False
        MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End If
    
    ErrTestExistFinalDir = Not bRet
    
End Function
