VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fRoles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "-"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   ClipControls    =   0   'False
   Icon            =   "fRoles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   360
      Index           =   4
      Left            =   3600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Perm"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   3600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox tDati 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.OptionButton opt 
      Caption         =   "Application Role"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.OptionButton opt 
      Caption         =   "Standard Role"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox tDati 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2415
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Image img 
      Height          =   270
      Index           =   0
      Left            =   0
      Picture         =   "fRoles.frx":000C
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lbl 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "fRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bLoading As Boolean
Private sDbName As String
Private sStuff As String
Private bNew As Boolean
Private bIsPub As Boolean
Private colUsr As Collection
Private oDb As SQLDMO.Database2
Private mlErr As Long, msErr As String

Public Property Get ErrStat(ByRef sErr As String) As Long
    ErrStat = mlErr
    sErr = msErr
End Property
Private Sub cmd_Click(Index As Integer)

    Dim sBuf As String
    
    Select Case Index
        Case 0
            Unload Me
        Case 1  'Save
            SaveRole
            If Flex.Enabled Then Flex.SetFocus
            cmd(1).Enabled = False
                
        Case 2
            sBuf = Replace(sStuff, "[", "")
            sBuf = Replace(sBuf, "]", "")
        
            LoadPermission sDbName, sBuf, act_Roles
        Case 3  'Add
            ListUsers
            Flex.SetFocus
        Case 4  'Drop
            RemoveUser
            Flex.SetFocus
    End Select

End Sub

Private Sub Flex_RowColChange()
    If Not bLoading Then
        If Flex.Col = 0 Then Flex.Col = 1
    End If
End Sub

Private Sub Form_Load()
    
    Flex.RowHeightMin = img(0).Height
    Set colUsr = New Collection
    lbl(0).Caption = MyLoadResString(k_Role_Name)
    cmd(2).Caption = MyLoadResString(k_Permissions)
    cmd(3).Caption = MyLoadResString(k_Add)
    cmd(4).Caption = MyLoadResString(k_Drop)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Save)
    
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colUsr = Nothing
    Set oDb = Nothing
    Set fRoles = Nothing
End Sub

Private Sub opt_Click(Index As Integer)

    Flex.Enabled = Index = 0
    cmd(3).Enabled = Index = 0 And Not bNew
    If Index = 0 Then
        cmd(4).Enabled = Flex.Rows > 1 And Not bIsPub
    Else
        cmd(4).Enabled = False
    End If
    tDati(1).Enabled = Index = 1 And bNew
    lbl(1).Enabled = Index = 1 And bNew

End Sub
Public Sub LoadRole(ByVal sName As String, ByVal sRole As String)
    
    Dim i As Integer
    Dim oRole As SQLDMO.DatabaseRole
    Dim oQry As SQLDMO.QueryResults
    Dim sBuf As String
               
    bLoading = True
    sDbName = sName
    sStuff = sRole
    bNew = Len(sRole) = 0
    
    tDati(0).Enabled = bNew
    opt(0).Enabled = bNew
    opt(1).Enabled = bNew
    cmd(2).Enabled = Not bNew
    Screen.MousePointer = vbHourglass

    Set colUsr = New Collection
    
    If Not bNew Then
        Me.Caption = MyLoadResString(k_Database_Role_Properties) & " '" & sStuff & "'"
        Set oDb = objServer.Databases(sDbName)
        Set oRole = oDb.DatabaseRoles(sStuff)
        
        tDati(0).Text = oRole.Name
        bIsPub = UCase$(oRole.Name) = "PUBLIC"
        cmd(2).Enabled = Not oRole.IsFixedRole Or bIsPub
        
        opt(1).Value = oRole.AppRole
        If Not oRole.AppRole Then
            On Local Error Resume Next
            Set oQry = oRole.EnumDatabaseRoleMember
            mlErr = Err.Number
            msErr = Err.Description
            On Local Error GoTo 0
            
            If mlErr = 0 Then
                With oQry
                    If Not oQry Is Nothing Then
                        If .ResultSets <> 0 Then
                            .CurrentResultSet = 1
                            For i = 1 To .Rows
                                sBuf = .GetColumnString(i, 1)
                                colUsr.Add sBuf, sBuf
                            Next
                        End If
                    End If
                End With
            End If
        Else
            opt(1).Value = True
            'tDati(1).Text = oRole.Password
        End If
    Else
        Me.Caption = MyLoadResString(k_Database_Role_Properties) & MyLoadResString(k_New_DB_Role)
        opt(0).Value = True
    End If
              
    If mlErr = 0 Then
        LoadFlex
        opt_Click IIf(opt(0).Value, 0, 1)
    End If
    bLoading = False
    Flex_RowColChange
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadFlex()

    Dim i As Integer

    With Flex
        .Redraw = False
        .Clear
        .FixedRows = 0
        .Cols = 2
        .Rows = 1
        .Clear
        .TextMatrix(0, 0) = MyLoadResString(k_User)
        .TextMatrix(0, 1) = MyLoadResString(k_User)
        .ColWidth(0) = 500
        
        .FixedAlignment(1) = flexAlignLeftCenter

        For i = 1 To colUsr.Count
            .AddItem vbTab & colUsr.Item(i)
            .Row = .Rows - 1
            .Col = 0
            Set Flex.CellPicture = LoadResPicture(k_ResUser, vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
        Next
                    
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth

        If .Rows > 1 Then
            .FixedRows = 1
        Else
            .Col = 0
            .CellBackColor = vbInactiveBorder
            .Col = 1
            .CellBackColor = vbInactiveBorder
        End If
        .MergeRow(0) = True
        .MergeCells = flexMergeRestrictColumns
        
        .Redraw = True
        SettaLargFlex Flex, 1, 300
        .TopRow = IIf(.Rows > 1, 1, 0)
        .Row = IIf(.Rows > 1, 1, 0)
        .Col = 0
    End With
    
End Sub

Private Sub SaveRole()
    
    Dim sName As String, sPwd As String
    Dim lErr As Long, sErr As String
        
    sName = Trim$(tDati(0).Text)
    sPwd = Trim$(tDati(1).Text)
    If Len(sName) Then
        Screen.MousePointer = vbHourglass
        On Local Error Resume Next
        If oDb Is Nothing Then Set oDb = objServer.Databases(sDbName)
        lErr = Err.Number
        If lErr = 0 Then
            Dim oRole As SQLDMO.DatabaseRole
            
            Set oRole = New SQLDMO.DatabaseRole
            oRole.Name = Trim$(tDati(0).Text)
            oRole.AppRole = opt(1).Value
            If opt(1).Value Then
                If Len(Trim$(tDati(1).Text)) Then oRole.Password = sPwd
            End If
            oDb.DatabaseRoles.Add oRole
            lErr = Err.Number
            Set oRole = Nothing
        End If
        sErr = Err.Description
        
        If lErr = 0 Then
            LoadRole sDbName, sName
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbDefault
            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
        End If
    End If
    
End Sub

Private Sub tDati_Change(Index As Integer)
    If Not bLoading Then
        'If Index = 0 Then cmd(1).Enabled = Len(Trim$(tDati(Index).Text))
        cmd(1).Enabled = Len(Trim$(tDati(Index).Text)) Or Index = 1
    End If
End Sub
Private Sub ListUsers()

    Dim lErr As Long, sErr As String
    Dim i As Integer, iC As Integer
    Dim bOk As Boolean
    
    Dim oUser As cUser
    Dim colX As New Collection
    
    Dim oQry As SQLDMO.QueryResults
    Dim oRole As SQLDMO.DatabaseRole
    Dim sBuf As String
    
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    Set oQry = oDb.EnumUsers
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If Not oQry Is Nothing And lErr = 0 Then
        With oQry
            If .ResultSets <> 0 Then
                .CurrentResultSet = 1
                For i = 1 To .Rows
                    sBuf = .GetColumnString(i, 3)
                    If Not UCase$(.GetColumnString(i, 3)) = "SA" And Not UCase$(.GetColumnString(i, 1)) = "DBO" Then
                        Set oUser = New cUser
                        oUser.oName = .GetColumnString(i, 3)
                        oUser.oGroup = .GetColumnString(i, 2)
                        On Local Error Resume Next
                        colX.Add oUser, oUser.oName
                        On Local Error GoTo 0
                    End If
                Next
            End If
        End With
                    
        For Each oRole In oDb.DatabaseRoles
            If Not oRole.IsFixedRole And Not UCase$(oRole.Name) = UCase$(sStuff) Then
                Set oUser = New cUser
                oUser.oName = oRole.Name
                oUser.oGroup = oUser.oName
                On Local Error Resume Next
                colX.Add oUser, oUser.oName
                On Local Error GoTo 0
            End If
        Next
    End If
    
    For i = 1 To colUsr.Count
        sBuf = UCase$(colUsr.Item(i))
        For iC = 1 To colX.Count
            If sBuf = UCase$(colX.Item(iC).oName) Then
                colX.Item(iC).Skip = True
                Exit For
            End If
        Next
    Next

    
    With fAddR.Lst
        For iC = 1 To colX.Count
            If Not colX.Item(iC).Skip Then
                .AddItem colX.Item(iC).oName
                .ItemData(.NewIndex) = iC
            End If
        Next
        bOk = .ListCount
    End With
    
    If bOk Then
        Screen.MousePointer = vbDefault
        fAddR.Show vbModal, Me
        Screen.MousePointer = vbHourglass
        If Not fAddR.Cancelled() Then
            On Local Error Resume Next
            Set oRole = oDb.DatabaseRoles(sStuff)
            lErr = Err.Number
            If lErr = 0 Then
                With fAddR.Lst
                    For i = 0 To .ListCount - 1
                        If .Selected(i) Then
                            sBuf = colX.Item(.ItemData(i)).oName
                            oRole.AddMember sBuf
                            lErr = Err.Number
                            If lErr = 0 Then colUsr.Add sBuf, sBuf
                            
                        End If
                        If lErr <> 0 Then Exit For
                    Next
                End With
            End If
            sErr = Err.Description
            On Local Error GoTo 0
        End If
        Unload fAddR
    Else
        MsgBox MyLoadResString(k_No_more_Users_4_Role), vbInformation Or vbOKOnly, App.EXEName
    End If

    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        bLoading = True
        LoadFlex
        bLoading = False
        Flex_RowColChange
        opt_Click Abs(opt(1).Value)
        Screen.MousePointer = vbDefault
    End If
    
End Sub
Private Sub RemoveUser()

    Dim lErr As Long, sErr As String
    Dim oRole As SQLDMO.DatabaseRole
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    Set oRole = oDb.DatabaseRoles(sStuff)
    lErr = Err.Number
    If lErr = 0 Then
        oRole.DropMember Flex.TextMatrix(Flex.Row, 1)
        lErr = Err.Number
        If lErr = 0 Then colUsr.Remove Flex.Row
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        bLoading = True
        LoadFlex
        bLoading = False
        Flex_RowColChange
        opt_Click Abs(opt(1).Value)
        Screen.MousePointer = vbDefault
    End If

End Sub
