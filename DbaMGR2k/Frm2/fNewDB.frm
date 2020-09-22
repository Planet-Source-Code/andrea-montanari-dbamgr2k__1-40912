VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fNewDB 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proprietà Database"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   Icon            =   "fNewDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Caption         =   "Database File"
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   6735
      Begin VB.TextBox tDati 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DB Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox tFlex 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   200
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Proprietà File"
      Height          =   1695
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   6735
      Begin VB.Frame fra 
         Caption         =   "Max Dimensione File"
         Height          =   975
         Index           =   4
         Left            =   3480
         TabIndex        =   18
         Top             =   600
         Width           =   2415
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "fNewDB.frx":000C
            Top             =   600
            Width           =   375
         End
         Begin VB.OptionButton opt 
            Caption         =   "Max (MB)"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Caption         =   "Illimitata"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Crescita File"
         Height          =   975
         Index           =   3
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Width           =   2655
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2040
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "fNewDB.frx":0010
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2040
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "fNewDB.frx":0013
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton opt 
            Caption         =   "Percent."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Caption         =   "Megabytes"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crescita Automatica"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Transaction Log File"
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   6735
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   6360
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0015
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0469
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":05C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0725
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":0FF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fNewDB.frx":114D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Refresh"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "New"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Property"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Permission"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "Database File"
      Height          =   3975
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   6735
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2520
         Width           =   6375
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   6375
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Compatibility Level"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Sort"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   4575
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "data"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fNewDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sWorkPath As String
Private bOldDB As Boolean
Private bLoading As Boolean
Private bRefreshData As Boolean
Private sSqlPath As String
Private miCol As Integer
Private miRow As Integer
Private miFlex As Integer
Private m_blnAnnullaMod As Boolean
Private m_iTmpType As iPropG


Private sDbName As String
Private sPrimaryGroup As String
Private Enum fileStatus
    statusOnLine = 0
    statusDelete = 1
    statusPrimary = 2
End Enum

Private Type fileOpt
    iStatus As fileStatus
    bDeleted As Boolean
    bAdded As Boolean
    sFileN As String
    sFilePos As String
    bFilePosChanged As Boolean
    sFilePosChanging As String
    sFileGroup As String
    lInitSize As Long
    bAutoG As Boolean
    bPercent As Boolean
    lGrow As Long
    bMaxG As Boolean
    lREstrictG As Long
End Type
Private dataFile() As fileOpt
Private LogFile() As fileOpt

Public Sub LoadDB(ByVal sDB As String)
    
    Dim lErr As Long
    Dim iFileC As Integer, iGroupC As Integer
    Dim iCur As Integer
    
    bOldDB = True
    sDbName = sDB
    objServer.Databases.Refresh True
    Dim oDb As SQLDMO.Database2
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    On Local Error GoTo 0
    
    
    Screen.MousePointer = vbHourglass
    If Not oDb Is Nothing Then
        tDati(0).Text = sDbName
        tDati(0).Locked = True
        With oDb
            On Local Error Resume Next
            iFileC = .FileGroups.Count
            lErr = Err.Number
            On Local Error Resume Next
            
            iCur = 0
            ReDim dataFile(0) As fileOpt
            
            For iGroupC = 1 To .FileGroups.Count
                For iFileC = 1 To .FileGroups(iGroupC).DBFiles.Count
                    
                    ReDim Preserve dataFile(iCur) As fileOpt
                    dataFile(iCur).sFileN = .FileGroups(iGroupC).DBFiles(iFileC).Name
                    dataFile(iCur).sFilePos = .FileGroups(iGroupC).DBFiles(iFileC).PhysicalName
                    dataFile(iCur).lInitSize = .FileGroups(iGroupC).DBFiles(iFileC).Size
                    dataFile(iCur).sFileGroup = .FileGroups(iGroupC).Name
                    dataFile(iCur).bAutoG = .FileGroups(iGroupC).DBFiles(iFileC).FileGrowth > 0   '<> -1
                    dataFile(iCur).bPercent = .FileGroups(iGroupC).DBFiles(iFileC).FileGrowthType
                    dataFile(iCur).lGrow = .FileGroups(iGroupC).DBFiles(iFileC).FileGrowth
                    dataFile(iCur).bMaxG = .FileGroups(iGroupC).DBFiles(iFileC).MaximumSize > 0 '<> -1
                    dataFile(iCur).lREstrictG = .FileGroups(iGroupC).DBFiles(iFileC).MaximumSize
                    If dataFile(iCur).lREstrictG <= 0 Then
                        dataFile(iCur).lREstrictG = .FileGroups(iGroupC).DBFiles(iFileC).Size
                    End If
                    If iCur = 0 Then
                        dataFile(iCur).iStatus = statusPrimary
                    Else
                        dataFile(iCur).iStatus = statusOnLine
                    End If
                    iCur = iCur + 1
                Next
            Next
            
            iCur = 0
            ReDim LogFile(0) As fileOpt
            
            For iFileC = 1 To .TransactionLog.LogFiles.Count
                    
                ReDim Preserve LogFile(iCur) As fileOpt
                LogFile(iCur).sFileN = .TransactionLog.LogFiles(iFileC).Name
                LogFile(iCur).sFilePos = .TransactionLog.LogFiles(iFileC).PhysicalName
                LogFile(iCur).lInitSize = .TransactionLog.LogFiles(iFileC).Size
                LogFile(iCur).bAutoG = .TransactionLog.LogFiles(iFileC).FileGrowth > 0  '<> -1
                LogFile(iCur).bPercent = .TransactionLog.LogFiles(iFileC).FileGrowthType
                LogFile(iCur).lGrow = .TransactionLog.LogFiles(iFileC).FileGrowth
                LogFile(iCur).bMaxG = .TransactionLog.LogFiles(iFileC).MaximumSize > 0  '<> -1
                LogFile(iCur).lREstrictG = .TransactionLog.LogFiles(iFileC).MaximumSize
                If LogFile(iCur).lREstrictG <= 0 Then
                    LogFile(iCur).lREstrictG = .TransactionLog.LogFiles(iFileC).Size
                End If
                iCur = iCur + 1
            Next
        End With
        tBar.Buttons(keyRefresh).Enabled = True
        tBar.Buttons(keyProperty).Enabled = True
        tBar.Buttons(keyNew).Enabled = True
        tBar.Buttons(keyPermission).Enabled = GetMembership(db_SecurAdm, oDb.Name)
                
        cbo(0).ListIndex = GetItem(oDb.Collation, cbo(0))
        
        cbo(1).ListIndex = GetItem(oDb.CurrentCompatibility, cbo(1))
        
        Label1(1).Enabled = False
        cbo(0).Enabled = False
        Label1(2).Enabled = True
        cbo(1).Enabled = True
        
    End If
    
    Screen.MousePointer = vbDefault
    Set oDb = Nothing
    
End Sub
Private Sub LoadFlex(ByVal iIndex As Integer)
    
    Dim i As Integer
    Dim sBuf As String
    Dim iImg As Integer
    Dim iCol2Large As Integer
    
    bLoading = True
    With Flex(iIndex)
        '.Redraw = False
        .FixedCols = 0
        .FixedRows = 0
        
        .Clear
        .Rows = 1
        .Cols = 1
        If iIndex = 0 Then
            .Cols = 7
            .TextMatrix(0, 1) = MyLoadResString(k_File_Name)
            .TextMatrix(0, 2) = MyLoadResString(k_File_Name)
            .TextMatrix(0, 3) = MyLoadResString(k_Location)
            .TextMatrix(0, 4) = MyLoadResString(k_Location)
            .TextMatrix(0, 5) = MyLoadResString(IIf(bOldDB, k_Space_Allocated, k_Initial_Allocation)) & " (MB)"
            .TextMatrix(0, 6) = MyLoadResString(k_HtmFileGroup)
            .ColWidth(0) = k_FlexPARAM_cellaSelector
                        
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeCol(4) = True
            .MergeRow(0) = True
            iCol2Large = 2
        Else
            .Cols = 5
            .MergeCells = flexMergeNever
        
            .TextMatrix(0, 1) = MyLoadResString(k_File_Name)
            .TextMatrix(0, 2) = MyLoadResString(k_Location)
            .TextMatrix(0, 3) = MyLoadResString(k_Location)
            .TextMatrix(0, 4) = MyLoadResString(IIf(bOldDB, k_Space_Allocated, k_Initial_Allocation)) & " (MB)"
            
            .ColWidth(0) = k_FlexPARAM_cellaSelector
            '.ColWidth(2) = TextWidth(String$(22, "X"))
        
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(2) = True
            .MergeCol(3) = True
            .MergeRow(0) = True
        
            iCol2Large = 1
        End If
        If iIndex = 0 Then
            .ColWidth(1) = k_FlexPARAM_cellaSelector
            .ColWidth(2) = (.Width - .ColWidth(4)) / 3
            .ColWidth(3) = k_FlexPARAM_cellaSelector
            .ColWidth(4) = (.Width - .ColWidth(4) - .ColWidth(2)) / 2
            .ColWidth(5) = 600
        Else
            .ColWidth(0) = k_FlexPARAM_cellaSelector
            .ColWidth(2) = k_FlexPARAM_cellaSelector
            .ColWidth(3) = .Width / 3
        End If
        
        If iIndex = 0 Then
            For i = 0 To UBound(dataFile)
                sBuf = vbTab & vbTab & dataFile(i).sFileN & vbTab & vbTab & dataFile(i).sFilePos & vbTab & dataFile(i).lInitSize & vbTab & dataFile(i).sFileGroup
                .AddItem sBuf
                
                .Row = .Rows - 1
                .Col = 1
                If dataFile(i).iStatus = statusPrimary Or (Not bOldDB And i = 0) Then
                    iImg = k_ResPrimaryFile
                ElseIf dataFile(i).iStatus = statusDelete Then
                    iImg = k_ResChkDel
                Else
                    iImg = k_ResChkOFF
                End If
                Set Flex(iIndex).CellPicture = LoadResPicture(iImg, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                
                
                .Col = 3
                Set Flex(iIndex).CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                If bOldDB Then
                    If Not dataFile(i).bAdded Then .CellBackColor = vbInactiveBorder
                End If
                
                If i = 0 Then
                    .Row = .Rows - 1
                    .Col = .Cols - 1
                    .CellBackColor = vbInactiveBorder
                End If
            Next
        Else
            For i = 0 To UBound(LogFile)
            
                sBuf = vbTab & LogFile(i).sFileN & vbTab & vbTab & LogFile(i).sFilePos & vbTab & LogFile(i).lInitSize
                .AddItem sBuf
                
                .Row = .Rows - 1
                .Col = 2
                Set Flex(iIndex).CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
                '.CellBackColor = vbInactiveBorder
            Next
        End If
        .FixedCols = 1
        .FixedRows = 1
        '.Col = 1
        .Redraw = True
    End With
    SettaLargFlex Flex(iIndex), iCol2Large, 0
    bLoading = False
    miRow = 0
    Flex_RowColChange iIndex
    
End Sub

Private Sub chk_Click(Index As Integer)
    
    If Index = 0 Then
        If bRefreshData Then
            If opt(0).Value Then
                If Val(tDati(1).Text) = 0 Then tDati(1).Text = 1
            End If
            fra(3).Enabled = chk(Index).Value
            fra(4).Enabled = chk(Index).Value
        End If
    End If
    
End Sub

Private Sub SaveDB()
        
    Dim lErr As Long, sErr As String
    Dim oDb As SQLDMO.Database2
    Dim oDBFileData As SQLDMO.DBFile
    Dim oLogFile As SQLDMO.LogFile
    Dim oGroup As SQLDMO.FileGroup
    Dim sActualGroup As String, sBuf As String
    Dim iItem As Integer
    Dim collFileG As Collection
    Dim collFileGDB As Collection
    Dim bExist As Boolean
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If Len(sDbName) = 0 Then
        lErr = 380
        sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_DbName))) '"DB Name"))
    End If
    If lErr = 0 Then
        For i = 0 To UBound(dataFile)
            If Len(Trim$(dataFile(i).sFileN)) = 0 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_File_Name))) '"File Name"))
            ElseIf Len(Trim$(dataFile(i).sFilePos)) = 0 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_Location))) '"File Pos."))
            ElseIf Len(Trim$(dataFile(i).sFileGroup)) = 0 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_HtmFileGroup))) '"File Group"))
            ElseIf Val(dataFile(i).lInitSize) < 1 Or Val(dataFile(i).lInitSize) < 1 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_File_Size))) '"File Size"))
            End If
        Next
    End If
    
    If lErr = 0 Then
        For i = 0 To UBound(LogFile)
            If Len(Trim$(LogFile(i).sFileN)) = 0 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_File_Name))) '"File Name"))
            ElseIf Len(Trim$(LogFile(i).sFilePos)) = 0 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_Location))) '"File Pos."))
            ElseIf Val(LogFile(i).lInitSize) < 1 Or Val(LogFile(i).lInitSize) < 1 Then
                lErr = 380
                sErr = ReplaceMsg(MyLoadResString(k_RES_Invalid_Value), Array("1%"), Array(MyLoadResString(k_File_Size))) '"File Size"))
            End If
        Next
    End If
    
    If lErr = 0 Then
        If Not bOldDB Then
            Set oDb = New SQLDMO.Database2
            
            oDb.Name = sDbName
            For i = 0 To UBound(dataFile)
                sActualGroup = dataFile(i).sFileGroup
                If StrComp(sActualGroup, sPrimaryGroup, vbTextCompare) = 0 Then
                    On Local Error Resume Next
                    Set oGroup = oDb.FileGroups.Item(sActualGroup)
                    If Err <> 0 Then
                        Err.Clear
                        Set oGroup = New SQLDMO.FileGroup
                        oGroup.Name = sActualGroup
                        oDb.FileGroups.Add oGroup
                    End If
                    lErr = Err.Number
                    sErr = Err.Description
                    On Local Error GoTo 0
                    
                    If lErr = 0 And Not oGroup Is Nothing Then
                        AddDbFile2Group True, oGroup, Nothing, dataFile(i), i = 0, lErr, sErr
                    End If
                End If
                If lErr <> 0 Then Exit For
            Next
            
            If lErr = 0 Then
                For i = 0 To UBound(LogFile)
                    AddLog2TransLog True, oDb.TransactionLog, Nothing, LogFile(i), lErr, sErr
                    If lErr <> 0 Then Exit For
                Next
            End If
            
            sBuf = cbo(0).Text
            If StrComp(sBuf, cbo(0).List(0), vbTextCompare) <> 0 Then
                oDb.Collation = sBuf
            Else
                oDb.Collation = objServer.Collation
            End If

            If lErr = 0 Then
                On Local Error Resume Next
                objServer.Databases.Add oDb
                lErr = Err.Number
                sErr = Err.Description
                On Local Error GoTo 0
            End If
            If lErr = 0 Then bOldDB = True
        Else
            On Local Error Resume Next
            Set oDb = objServer.Databases(sDbName)
            lErr = Err.Number
            sErr = Err.Description
            On Local Error GoTo 0

            If lErr = 0 Then
                With oDb
                
                    'compatibility
                    sBuf = cbo(1).Text
                    If StrComp(sBuf, .CurrentCompatibility, vbTextCompare) <> 0 Then .CurrentCompatibility = sBuf
                
                    'Data
                    For i = 0 To UBound(dataFile)
                        sActualGroup = dataFile(i).sFileGroup
                        On Local Error Resume Next
                        Set oGroup = oDb.FileGroups.Item(sActualGroup)
                        If Err <> 0 Then
                            Err.Clear
                            Set oGroup = New SQLDMO.FileGroup
                            oGroup.Name = sActualGroup
                            oDb.FileGroups.Add oGroup
                        End If
                        lErr = Err.Number
                        sErr = Err.Description
                        On Local Error GoTo 0
                    
                        If dataFile(i).bAdded Then
                            If lErr = 0 Then
                                AddDbFile2Group True, oGroup, Nothing, dataFile(i), False, lErr, sErr
                            End If
                        ElseIf dataFile(i).iStatus = statusDelete Then
                        
                            If lErr = 0 Then Call GetDBFileDATA(oGroup, dataFile(i).sFileN, lErr, sErr, iItem)
                            On Local Error Resume Next
                            If lErr = 0 Then oGroup.DBFiles.Remove iItem
                            lErr = Err.Number
                            sErr = Err.Description
                            On Local Error GoTo 0
                        Else
                            If lErr = 0 Then Set oDBFileData = GetDBFileDATA(oGroup, dataFile(i).sFileN, lErr, sErr, iItem)
                                                        
                            If lErr = 0 Then
                                AddDbFile2Group False, oGroup, oDBFileData, dataFile(i), False, lErr, sErr
                            End If
                        End If
                        If lErr <> 0 Then Exit For
                    Next

                                        
                    'Log
                    If lErr = 0 Then
                        For i = 0 To UBound(LogFile)
                            If LogFile(i).bAdded Then
                                AddLog2TransLog True, oDb.TransactionLog, Nothing, LogFile(i), lErr, sErr
                            Else
                                Set oLogFile = GetDBFileLog(oDb.TransactionLog, LogFile(i).sFileN, lErr, sErr, iItem)
                                
                                If lErr = 0 Then
                                    AddLog2TransLog False, oDb.TransactionLog, oLogFile, LogFile(i), lErr, sErr
                                End If
                            End If
                            If lErr <> 0 Then Exit For
                        Next
                    End If
                    
                    If lErr = 0 Then
                        objServer.Databases.Refresh True
                        Set collFileG = New Collection
                        Set collFileGDB = New Collection
                        For i = 1 To objServer.Databases(sDbName).FileGroups.Count
                            sActualGroup = objServer.Databases(sDbName).FileGroups.Item(i).Name
                            collFileGDB.Add sActualGroup, sActualGroup
                            If objServer.Databases(sDbName).FileGroups.Item(i).DBFiles.Count > 0 Then
                                collFileG.Add sActualGroup, sActualGroup
                            End If
                        Next
                        
                        For i = 1 To collFileGDB.Count
                            sActualGroup = collFileGDB.Item(i)
                            On Local Error Resume Next
                            sBuf = collFileG.Item(sActualGroup)
                            bExist = Err.Number = 0
                            On Local Error GoTo 0
                            
                            If Not bExist Then
                                On Local Error Resume Next
                                objServer.Databases(sDbName).FileGroups.Remove sActualGroup
                                lErr = Err.Number
                                sErr = Err.Description
                                On Local Error GoTo 0
                            End If
                            If lErr <> 0 Then
                                Screen.MousePointer = vbDefault
                                MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
                                Screen.MousePointer = vbHourglass
                                lErr = 0
                                Exit For
                            End If
                        Next
                    End If
                    
                End With
            End If
        End If
    End If
    
   
    Screen.MousePointer = vbDefault
    If lErr <> 0 Then
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        objServer.Databases.Refresh True
        LoadDB sDbName
        bRefreshData = False
        Set tStrip.SelectedItem = tStrip.Tabs(1)
        bRefreshData = True
    End If
    
ExitHere:
    
    Set oDBFileData = Nothing
    Set oLogFile = Nothing
    Set oGroup = Nothing
    Set oDb = Nothing

End Sub

Private Sub Flex_DblClick(Index As Integer)
    FlexEdit Index, 0
End Sub

Private Sub Flex_GotFocus(Index As Integer)
    TestActiveControl
End Sub

Private Sub Flex_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim iRow As Integer
    Dim iImg As Integer
    Dim iCol As Integer
    
    If Index = 0 And Flex(Index).Col = 1 And (KeyCode = vbKeyF2 Or KeyCode = vbKeyDELETE) And bOldDB Then
        iRow = Flex(Index).Row
        If iRow - 1 <= UBound(dataFile) Then
            If dataFile(iRow - 1).bAdded Then
                DeleteRow
            ElseIf Not dataFile(iRow - 1).iStatus = statusPrimary Then
                If dataFile(iRow - 1).iStatus = statusDelete Then
                    iImg = k_ResChkOFF
                    dataFile(iRow - 1).iStatus = statusOnLine
                Else
                    iImg = k_ResChkDel
                    dataFile(iRow - 1).iStatus = statusDelete
                End If
                bLoading = True
                iCol = Flex(Index).Col
                Flex(Index).Col = 1
                Set Flex(Index).CellPicture = LoadResPicture(iImg, vbResBitmap)
                Flex(Index).CellPictureAlignment = flexAlignCenterCenter
                Flex(Index).Col = iCol
                bLoading = False
            End If
        End If
    Else
        If KeyCode = vbKeyF2 Then FlexEdit Index, 0
    End If
End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)
    FlexEdit Index, KeyAscii
End Sub

Private Sub Flex_RowColChange(Index As Integer)

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
       
    If Not bLoading Then
        If Not b Then
            b = True
            TestActiveControl
            SaveRecord
            
            With Flex(Index)
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                    
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex(Index).CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex(Index).CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1

                    .Col = iCol
                    miRow = iRowP
                End If
                LoadRecord Index

            End With
            b = False
        End If
    End If

End Sub

Private Sub Flex_Scroll(Index As Integer)
    tFlex.Visible = False
End Sub

Private Sub Form_Activate()
    
    Static b As Boolean
    If Not b Then
        b = True
        bRefreshData = False
        LoadFlex 0
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo
    'spostamento sulla scheda successiva.
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tStrip.SelectedItem.Index
        If i = tStrip.Tabs.Count Then
            'È stata raggiunta l'ultima scheda e quindi
            'torna alla scheda 1.
            Set tStrip.SelectedItem = tStrip.Tabs(1)
        Else
            'Incrementa l'indice della scheda
            Set tStrip.SelectedItem = tStrip.Tabs(i + 1)
        End If
    End If
    
End Sub

Private Sub Form_Initialize()
    
    sSqlPath = objServer.Registry.SQLDataRoot & "\"
    sWorkPath = sSqlPath

    
    ReDim dataFile(0) As fileOpt
    ReDim LogFile(0) As fileOpt
    sPrimaryGroup = "PRIMARY"
    
    With dataFile(0)
        .sFileN = ""
        .sFilePos = ""
        .sFileGroup = sPrimaryGroup
        .lInitSize = 2
        .bAutoG = True
        .bPercent = True
        .lGrow = 10
        .bMaxG = False 'True
        '.lREstrictG = 2
    End With
    With LogFile(0)
        .sFileN = ""
        .sFilePos = ""
        .lInitSize = 1
        .bAutoG = True
        .bPercent = True
        .lGrow = 10
        .bMaxG = False 'True
        '.lREstrictG = 2
    End With

    
End Sub

Private Sub Form_Load()
    
    tBar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    tBar.Buttons(keySave).ToolTipText = MyLoadResString(k_Save)
    tBar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_RES_FitGrid)
    tBar.Buttons(keyNew).ToolTipText = MyLoadResString(k_AddNewFile)
    tBar.Buttons(keyDelete).ToolTipText = MyLoadResString(k_RemoveAddedFile)
    tBar.Buttons(keyProperty).ToolTipText = MyLoadResString(k_Options)
    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource)
    tBar.Buttons(keyPermission).ToolTipText = MyLoadResString(k_Permissions)
    
    tStrip.Tabs(1).Caption = MyLoadResString(k_NewDbDataFile)
    tStrip.Tabs(2).Caption = MyLoadResString(k_NewDbTrLog)
    tStrip.Tabs(3).Caption = MyLoadResString(k_General)
    
    fra(5).Caption = MyLoadResString(k_Settings)
    Label1(1).Caption = MyLoadResString(k_NewDbSort)
    Label1(2).Caption = MyLoadResString(k_CompatLevel)
    
    Me.Caption = MyLoadResString(k_Database_Property)
    fra(2).Caption = MyLoadResString(k_File_Property)
    fra(3).Caption = MyLoadResString(k_File_Growth)
    fra(4).Caption = MyLoadResString(k_Max_File_Growth)
    chk(0).Caption = MyLoadResString(k_Automatic_Growth)
    opt(2).Caption = MyLoadResString(k_Unlimited)
    Flex(0).RowHeightMin = tFlex.Height
    Flex(1).RowHeightMin = tFlex.Height
    
    InitCombo
    Label1(2).Enabled = False
    cbo(1).Enabled = False
            
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fNewDB = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
    
    If bRefreshData Then
        If Index < 2 Then
            tDati(1).Enabled = Index = 0
            tDati(2).Enabled = Index = 1
        Else
            tDati(3).Enabled = Index = 3
        End If
    End If
    
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If Index < 2 Then
        tDati(1).Enabled = Index = 0
        tDati(2).Enabled = Index = 1
    Else
        tDati(3).Enabled = Index = 3
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    TestActiveControl
    Select Case Button.Key
        Case keyEnd
            Unload Me
        Case keySave
            SaveRecord
            SaveDB
        Case keyFitGrid
            LoadFlex (tStrip.SelectedItem.Index - 1)
        Case keyNew
            Add1Row
        Case keyDelete
            DeleteRow
        Case keyProperty
            LoadProperty
        Case keyPermission
            PrivateLoadPermission
        Case keyRefresh
            Screen.MousePointer = vbHourglass
            objServer.Databases.Refresh True
            LoadDB sDbName
            bRefreshData = False
            If tStrip.Enabled Then
                Set tStrip.SelectedItem = tStrip.Tabs(1)
            End If
            bRefreshData = True
            Screen.MousePointer = vbDefault
    End Select
    
End Sub

Private Sub tDati_Change(Index As Integer)
    
    Dim sBuf As String
    sBuf = Trim$(tDati(Index).Text)
    
    Select Case Index
        Case 0  'Nome DB
            If Not bOldDB Then
                sBuf = Trim$(tDati(Index).Text)

                If Len(sBuf) Then
                    dataFile(0).sFileN = sBuf & "_D"
                    dataFile(0).sFilePos = sWorkPath & sBuf & ".MDF"

                    LogFile(0).sFileN = sBuf & "_L"
                    LogFile(0).sFilePos = sWorkPath & sBuf & ".LDF"
                Else
                    dataFile(0).sFileN = ""
                    dataFile(0).sFilePos = ""
                    LogFile(0).sFileN = ""
                    LogFile(0).sFilePos = ""
                End If
                Flex(0).TextMatrix(1, 2) = dataFile(0).sFileN
                Flex(0).TextMatrix(1, 4) = dataFile(0).sFilePos
            End If
    End Select
   
End Sub

Private Sub tDati_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index > 0 Then Digit_KDown tDati(Index), KeyCode
End Sub

Private Sub tDati_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys kNEXT
        Case vbKeyEscape
            KeyAscii = 0
            SendKeys kPREV
        Case vbKeyBACK
            If Index > 0 Then Digit_KPress tDati(Index), KeyAscii
        Case 48 To 57, 65 To 90, 95, 97 To 122
            If Index > 0 Then Digit_KPress tDati(Index), KeyAscii
            
        Case Else
            KeyAscii = 0
    End Select
            
End Sub

Private Sub tDati_Validate(Index As Integer, Cancel As Boolean)
    
    Dim iVal As Integer
        
    Select Case Index
        Case 1  'Crescita in MB
            If Val(tDati(Index).Text) = 0 Then
                tDati(Index).Text = 1
                Cancel = True
            End If
        Case 2  'Crescita %
            iVal = Val(tDati(Index).Text)
            If iVal > 100 Then
                tDati(Index).Text = 100
                Cancel = True
            ElseIf iVal = 0 Then
                tDati(Index).Text = 1
                Cancel = True
            End If
        Case 3
'            If Val(tDati(Index).Text) < 2 Then
'                Beep
'                iVal = optFile(tStrip.SelectedItem.Index - 1).lREstrictG
'                If iVal < 2 Then iVal = 2
'                Cancel = True
'                tDati(Index).Text = iVal
'            End If
    End Select
    If Cancel Then Beep
    
End Sub

Private Sub tFlex_Change()
    
    Dim tmpData As fileOpt
    Dim bOk As Boolean
    Dim sExt As String
            
    If miFlex = 0 Then
        LSet tmpData = dataFile(miRow - 1)
        bOk = miCol = 2
        sExt = ReplaceMsg(".1%DF", Array("1%"), Array(IIf(miRow > 1, "N", "M")))
    Else
        LSet tmpData = LogFile(miRow - 1)
        bOk = miCol = 1
        sExt = ".LDF"
    End If
    If bOk Then bOk = Len(tmpData.sFilePos) = 0
    
    If bOk Then
        tmpData.bFilePosChanged = True
        tmpData.sFilePosChanging = sWorkPath & Trim$(tFlex.Text) & sExt
        tmpData.sFilePosChanging = Replace(tmpData.sFilePosChanging, "_D", "", Compare:=vbTextCompare)
        tmpData.sFilePosChanging = Replace(tmpData.sFilePosChanging, "_L", "", Compare:=vbTextCompare)
        Flex(miFlex).TextMatrix(miRow, IIf(miFlex = 0, 4, 3)) = tmpData.sFilePosChanging
    
        If miFlex = 0 Then
            LSet dataFile(miRow - 1) = tmpData
        Else
            LSet LogFile(miRow - 1) = tmpData
        End If
    End If
    
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub tStrip_BeforeClick(Cancel As Integer)

    If bRefreshData Then
        SaveRecord
    End If
End Sub
Private Sub tStrip_Click()

    Dim i As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    fra(0).Visible = i = 0
    fra(1).Visible = i = 1
    fra(2).Visible = i <> 2
    fra(5).Visible = i = 2
    
    fra(i).ZOrder
    If i < 2 Then LoadFlex i
    If i = 0 Then
        tDati(0).SetFocus
    ElseIf i = 1 Then
        Flex(1).SetFocus
    Else
        If cbo(0).Enabled Then
            i = 0
        Else
            i = 1
        End If
        cbo(i).SetFocus
    End If

End Sub

Private Sub LoadRecord(ByVal Index As Integer)

    Dim tmpData As fileOpt
    
    bRefreshData = True
    If Index = 0 Then
        LSet tmpData = dataFile(miRow - 1)
    Else
        LSet tmpData = LogFile(miRow - 1)
    End If

    chk(0).Value = Abs(tmpData.bAutoG)

    opt(0).Value = Abs(tmpData.bPercent) = 0
    opt(1).Value = Abs(tmpData.bPercent)

    If Not tmpData.bPercent Then
        tDati(1).Text = tmpData.lGrow
        tDati(2).Text = 10
    Else
        tDati(1).Text = 1
        tDati(2).Text = tmpData.lGrow
    End If

    opt(2).Value = Abs(tmpData.bMaxG) = 0
    opt(3).Value = Abs(tmpData.bMaxG) <> 0

    If Not tmpData.bMaxG Then
        tDati(3).Text = 2
    Else
        tDati(3).Text = tmpData.lREstrictG
    End If

    fra(3).Enabled = chk(0).Value
    fra(4).Enabled = chk(0).Value

    tDati(1).Enabled = opt(0).Value
    tDati(2).Enabled = opt(1).Value
    tDati(3).Enabled = opt(3).Value
    tBar.Buttons(keyDelete).Enabled = tmpData.bAdded

End Sub

Private Sub SaveRecord()

    Dim i As Integer
    Dim tmpData As fileOpt
    
    i = tStrip.SelectedItem.Index - 1
    If i = 0 Then sDbName = Trim$(tDati(0).Text)
    'If bRefreshData And miRow > 1 Then
    If bRefreshData And miRow > 0 Then
        
        If i = 0 Then
            LSet tmpData = dataFile(miRow - 1)
        Else
            LSet tmpData = LogFile(miRow - 1)
        End If
        Debug.Print tmpData.bAdded
        With tmpData
            .bAutoG = chk(0).Value
            .bPercent = opt(1).Value
            If Not .bPercent Then
                .lGrow = Val(tDati(1).Text)
            Else
                .lGrow = Val(tDati(2).Text)
            End If
            .bMaxG = opt(3).Value
            If .bMaxG Then
                .lREstrictG = Val(tDati(3).Text)
            Else
                .lREstrictG = -1
            End If
        End With
        If i = 0 Then
            LSet dataFile(miRow - 1) = tmpData
        Else
            LSet LogFile(miRow - 1) = tmpData
        End If
    End If
End Sub
Private Sub FlexEdit(ByVal iFlex As Integer, ByVal KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer
    Dim tmpData As fileOpt
    
        
    If Len(Trim$(tDati(0).Text)) = 0 Then Exit Sub
        
    miFlex = iFlex
    
    sBuf = Flex(miFlex).Text

    bEdit = True

    miCol = Flex(miFlex).Col
    
    
    If iFlex = 0 Then
        LSet tmpData = dataFile(Flex(iFlex).Row - 1)
        iLen = 5
    Else
        LSet tmpData = LogFile(Flex(iFlex).Row - 1)
        iLen = 4
    End If
    If Not tmpData.bAdded Then
        If bOldDB Then
            
            If miCol <> iLen Then Exit Sub
        Else
            If miCol = 1 And miFlex = 0 Then Exit Sub
        End If
    End If
    iLen = 0
    
    m_iTmpType = am_String
    If miFlex = 0 Then
        Select Case miCol
            Case 2, 4
                m_iTmpType = am_String
                If miCol = 2 Then
                    iLen = k_MaxDBLen
                Else
                    iLen = k_MaxPath_DBLen
                End If
                                
'            Case 3
'                m_iTmpType = am_String
'                iLen = 30
            Case 5
                m_iTmpType = am_Long
                iLen = 2
            Case 6
                m_iTmpType = am_String
                iLen = 25
        End Select
    Else
        Select Case miCol
            Case 1, 3
                m_iTmpType = am_String
                If miCol = 1 Then
                    iLen = k_MaxDBLen
                Else
                    iLen = k_MaxPath_DBLen
                End If
'            Case 2
'                m_iTmpType = am_String
'                iLen = 30
            Case 4
                m_iTmpType = am_Long
                iLen = 2
            Case 5
                bEdit = False
        End Select
    End If
    
    If Flex(miFlex).CellBackColor = vbInactiveBorder Then bEdit = False

    If miCol < 1 Then Exit Sub
    If bEdit Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
            Case 0 To 32
                If miFlex = 0 Then
                    If miCol = 1 Then
                        bEdit = False
                    ElseIf miCol = 3 Then
                        bEdit = False
                        OpenDialog
                    End If
                Else
                    If miCol = 2 Then
                        bEdit = False
                        OpenDialog
                    End If
                End If
            Case Else
                If miFlex = 0 Then
                    If miCol = 1 Then
                        bEdit = False
                    ElseIf miCol = 3 Then
                        bEdit = False
                    End If
                End If
            
        End Select
    End If

    If bEdit Then
        
        tFlex.MaxLength = 0
        tFlex.Text = ""
        tFlex.MaxLength = iLen
        tFlex.Text = sBuf
        Debug.Print m_iTmpType = am_Long
        tFlex.SelStart = 0
        With Flex(miFlex)
            tFlex.Move fra(miFlex).Left + .Left + .CellLeft, fra(miFlex).Top + .Top + .CellTop, .CellWidth, .CellHeight
        End With
        tFlex.Visible = True
        tFlex.ZOrder
        tFlex.SetFocus
        If (m_iTmpType And am_String) = am_String Then
            If KeyAscii = vbKeySpace Then KeyAscii = 0
        End If
        If KeyAscii <> 0 Then
            On Local Error Resume Next
            SendKeys Chr$(KeyAscii)
            On Local Error GoTo 0
        End If
    End If

End Sub

Private Sub tFlex_KeyPress(KeyAscii As Integer)

    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93    '<>|[]
            KeyAscii = 0
        Case 92
            If miFlex = 0 Then
                If miCol <> 4 Then KeyAscii = 0
            Else
                If miCol <> 3 Then KeyAscii = 0
            End If
            OverStrike tFlex
        Case vbKeyEscape
            KeyAscii = 0
            m_blnAnnullaMod = True
            tFlex = Flex(miFlex).Text
            Flex(miFlex).SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex(miFlex).SetFocus
            
        Case vbKeyBACK
            If (m_iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (m_iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (m_iTmpType And am_String) = am_String Then

            End If
        Case vbKeySpace
            KeyAscii = 0
        Case Else
            If (m_iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (m_iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (m_iTmpType And am_String) = am_String Then
                OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
    
    Dim bUpdate As Boolean
    Dim tmpData As fileOpt
    Dim sBuf As String
    Dim lVal As Long
        
    If Not m_blnAnnullaMod Then
        tFlex.Text = Trim$(tFlex.Text)
        If miFlex = 0 Then
            LSet tmpData = dataFile(miRow - 1)
            Select Case miCol
                Case 2, 4, 6
                    If Len(tFlex.Text) Then
                        bUpdate = True
                        sBuf = tFlex.Text
                    End If
                Case 5
                    bUpdate = ValF(tFlex.Text) >= Val(tmpData.lInitSize)
                    If bUpdate Then lVal = ValF(tFlex.Text)
            End Select
        Else
            LSet tmpData = LogFile(miRow - 1)
            Select Case miCol
                Case 1, 3
                    If Len(tFlex.Text) Then
                        bUpdate = True
                        sBuf = tFlex.Text
                    End If
                Case 4
                    bUpdate = ValF(tFlex.Text) >= Val(tmpData.lInitSize)
                    If bUpdate Then lVal = ValF(tFlex.Text)
            End Select
        End If
        If bUpdate Then
            With tmpData
                If miFlex = 0 Then
                    If miCol = 2 Then
                        .sFileN = sBuf
                        If tmpData.bFilePosChanged = True Then
                            tmpData.bFilePosChanged = False
                            tmpData.sFilePos = tmpData.sFilePosChanging
                            tmpData.sFilePosChanging = ""
                            Flex(miFlex).TextMatrix(miRow, 4) = tmpData.sFilePos
                        End If
                        
                    ElseIf miCol = 4 Then
                        .sFilePos = sBuf
                    ElseIf miCol = 6 Then
                        .sFileGroup = sBuf
                    ElseIf miCol = 5 Then
                        .lInitSize = lVal
                    End If
                    LSet dataFile(miRow - 1) = tmpData
                Else
                    If miCol = 1 Then
                        .sFileN = sBuf
                        If tmpData.bFilePosChanged = True Then
                            tmpData.bFilePosChanged = False
                            tmpData.sFilePos = tmpData.sFilePosChanging
                            tmpData.sFilePosChanging = ""
                            Flex(miFlex).TextMatrix(miRow, 3) = tmpData.sFilePos
                        End If
                    
                    ElseIf miCol = 3 Then
                        .sFilePos = sBuf
                    ElseIf miCol = 4 Then
                        .lInitSize = lVal
                    End If
                    LSet LogFile(miRow - 1) = tmpData
                End If
            End With
        End If
    End If
    If bUpdate Then
        Debug.Print Flex(1).Rows
        Debug.Print Flex(1).Cols
        
        Flex(miFlex).TextMatrix(miRow, miCol) = tFlex.Text
    Else
        If Not m_blnAnnullaMod Then Beep
    End If
        
    tFlex.Visible = False
    m_blnAnnullaMod = True
    
End Sub

Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUP Then
        If miRow > 1 Then
            Flex(miFlex).Row = Flex(miFlex).Row - 1
            Flex(miFlex).SetFocus
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyDOWN Then
        If miRow < Flex(miFlex).Rows - 1 Then
            Flex(miFlex).Row = Flex(miFlex).Row + 1
            Flex(miFlex).SetFocus
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyLEFT Or KeyCode = vbKeyRIGHT Then
        If KeyCode = vbKeyLEFT Then
            If tFlex.SelStart = 0 Then
                KeyCode = 0
                Flex(miFlex).Col = Flex(miFlex).Col - 1
                Flex(miFlex).SetFocus
            End If
        ElseIf KeyCode = vbKeyRIGHT Then
            If tFlex.SelStart = Len(tFlex.Text) And Flex(miFlex).Col < Flex(miFlex).Cols - 1 Then
                KeyCode = 0
                Flex(miFlex).Col = Flex(miFlex).Col + 1
                Flex(miFlex).SetFocus
            End If
        End If
    Else
        If (m_iTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (m_iTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (m_iTmpType And am_String) = am_String Then
        End If
    End If
End Sub

Private Sub TestActiveControl()

    If tFlex.Visible Then
        tFlex_LostFocus
    End If

   On Local Error Resume Next
   If TypeOf Me.ActiveControl Is TextBox Then
      If Me.ActiveControl Is tFlex Then
         tFlex_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub

Private Sub Add1Row()
    
    Dim i As Integer
    Dim iItem As Integer
    Dim tmpData As fileOpt
    Dim tmpTestData As fileOpt
    Dim sBuf As String
    Dim iCol As Integer, iRow As Integer
    
    i = tStrip.SelectedItem.Index - 1
    
    If i = 0 Then
        iItem = UBound(dataFile)
        LSet tmpTestData = dataFile(iItem)
    Else
        iItem = UBound(LogFile)
        LSet tmpTestData = LogFile(iItem)
    End If
    
    If tmpTestData.bAdded And (Len(tmpTestData.sFileN) = 0 Or Len(tmpTestData.sFilePos) = 0) Or (i = 0 And Len(tmpTestData.sFileGroup) = 0) Then
        Exit Sub
    End If
        
    With tmpData
        .bAdded = True
        .iStatus = statusOnLine
        .lInitSize = 1
        
        .bAutoG = True
        .bPercent = True
        .lGrow = 10
        .bMaxG = False
        .lREstrictG = 0
        
        If i = 0 Then
            .sFileGroup = sPrimaryGroup
        End If
    End With
    bLoading = True
    If i = 0 Then
        iItem = UBound(dataFile) + 1
        ReDim Preserve dataFile(iItem) As fileOpt
        LSet dataFile(iItem) = tmpData
        sBuf = vbTab & vbTab & dataFile(iItem).sFileN & vbTab & vbTab & dataFile(iItem).sFilePos & vbTab & dataFile(iItem).lInitSize & vbTab & dataFile(iItem).sFileGroup
    Else
        iItem = UBound(LogFile) + 1
        ReDim Preserve LogFile(iItem) As fileOpt
        LSet LogFile(iItem) = tmpData
        sBuf = vbTab & LogFile(iItem).sFileN & vbTab & vbTab & LogFile(iItem).sFilePos & vbTab & LogFile(iItem).lInitSize
    End If
    With Flex(i)
        iCol = .Col
        iRow = .Row
        .AddItem sBuf
        
        If i = 0 Then
            .Row = .Rows - 1
            .Col = 1
            Set Flex(i).CellPicture = LoadResPicture(k_ResChkOFF, vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            
        End If
        .Row = .Rows - 1
        
        .Col = IIf(i = 0, 3, 2)
        Set Flex(i).CellPicture = LoadResPicture(k_ResEllipsis, vbResBitmap)
        .CellPictureAlignment = flexAlignCenterCenter

        .Row = iRow
        .Col = iCol
    End With
    
    bLoading = False
    
End Sub
Private Sub DeleteRow()

    Dim iFlex As Integer
    Dim i As Integer
    Dim iItem As Integer
    Dim tmpTestData() As fileOpt
    
    Dim iCol As Integer, iRow As Integer
    
    i = tStrip.SelectedItem.Index - 1
    iFlex = i
    iCol = Flex(i).Col
    iRow = Flex(i).Row
    iItem = iRow - 1
    
    If i = 0 Then
        dataFile(iItem).bDeleted = True
    Else
        LogFile(iItem).bDeleted = True
    End If
    
    iItem = -1
    If i = 0 Then
        For i = 0 To UBound(dataFile)
            If Not dataFile(i).bDeleted Then
                iItem = iItem + 1
                ReDim Preserve tmpTestData(iItem) As fileOpt
                LSet tmpTestData(iItem) = dataFile(i)
            End If
        Next
        
        ReDim dataFile(iItem) As fileOpt
        For i = 0 To UBound(tmpTestData)
            LSet dataFile(i) = tmpTestData(i)
        Next
        
    Else
        For i = 0 To UBound(LogFile)
            If Not LogFile(i).bDeleted Then
                iItem = iItem + 1
                ReDim Preserve tmpTestData(iItem) As fileOpt
                LSet tmpTestData(iItem) = LogFile(i)
            End If
        Next
        ReDim LogFile(iItem) As fileOpt
        For i = 0 To UBound(tmpTestData)
            LSet LogFile(i) = tmpTestData(i)
        Next
    End If
    bLoading = True
    Flex(iFlex).RemoveItem iRow
    miRow = miRow - 1
    bLoading = False
    Flex_RowColChange iFlex
    
End Sub
Private Function GetDBFileDATA(ByVal objGroup As SQLDMO.FileGroup, ByVal sName As String, _
                                ByRef lErr As Long, ByRef sErr As String, ByRef iItem As Integer) As SQLDMO.DBFile

    Dim i As Integer
    Dim bDone As Boolean
    
    With objGroup
        For i = 1 To .DBFiles.Count
            If StrComp(.DBFiles.Item(i).Name, sName, vbTextCompare) = 0 Then
                Set GetDBFileDATA = .DBFiles.Item(i)
                iItem = i
                bDone = True
                Exit For
            End If
        Next
    End With
    If Not bDone Then
        lErr = 327
        sErr = ReplaceMsg(MyLoadResString(k_RES_Object_Not_Found), Array("1%", "2%"), Array("DBFile", sName))
    End If
    
End Function
Private Function GetDBFileLog(ByVal objTrans As SQLDMO.TransactionLog, ByVal sName As String, _
                                ByRef lErr As Long, ByRef sErr As String, ByRef iItem As Integer) As SQLDMO.LogFile

    Dim i As Integer
    Dim bDone As Boolean
    
    With objTrans
        For i = 1 To .LogFiles.Count
            If StrComp(.LogFiles.Item(i).Name, sName, vbTextCompare) = 0 Then
                Set GetDBFileLog = .LogFiles.Item(i)
                iItem = i
                bDone = True
                Exit For
            End If
        Next
    End With
    If Not bDone Then
        lErr = 327
        sErr = ReplaceMsg(MyLoadResString(k_RES_Object_Not_Found), Array("1%", "2%"), Array("LogFile", sName))
    End If
    
End Function

Private Sub AddDbFile2Group(ByVal bAdd As Boolean, _
                            ByRef objGroup As SQLDMO.FileGroup, _
                            ByRef oDBFileData As SQLDMO.DBFile, _
                            dataFile As fileOpt, ByVal bIsPrimary As Boolean, _
                            ByRef lErr As Long, ByRef sErr As String)


    If bAdd Or oDBFileData Is Nothing Then
        Set oDBFileData = New SQLDMO.DBFile
        bAdd = True
    End If
    lErr = 0
    On Local Error Resume Next

    If bAdd Then
        oDBFileData.Name = dataFile.sFileN
        oDBFileData.PhysicalName = dataFile.sFilePos
        oDBFileData.PrimaryFile = bIsPrimary
    End If
    oDBFileData.Size = dataFile.lInitSize

    If dataFile.bAutoG Then
        oDBFileData.FileGrowthType = IIf(dataFile.bPercent, SQLDMOGrowth_Percent, SQLDMOGrowth_MB)
        oDBFileData.FileGrowth = dataFile.lGrow
        If dataFile.bMaxG Then oDBFileData.MaximumSize = dataFile.lREstrictG
    Else
        oDBFileData.FileGrowth = 0
    End If
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0

    If lErr = 0 And bAdd Then
        On Local Error Resume Next
        objGroup.DBFiles.Add oDBFileData
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    End If
End Sub

Private Sub AddLog2TransLog(ByVal bAdd As Boolean, _
                            ByRef oTransactionLog As SQLDMO.TransactionLog, _
                            ByRef oLogFile As SQLDMO.LogFile, _
                            LogFile As fileOpt, ByRef lErr As Long, ByRef sErr As String)

    If oLogFile Is Nothing Or bAdd Then
        Set oLogFile = New SQLDMO.LogFile
        bAdd = True
    End If

    On Local Error Resume Next
    If bAdd Then
        oLogFile.Name = LogFile.sFileN
        oLogFile.PhysicalName = LogFile.sFilePos
    End If
    oLogFile.Size = LogFile.lInitSize

    If LogFile.bAutoG Then
        oLogFile.FileGrowthType = IIf(LogFile.bPercent, SQLDMOGrowth_Percent, SQLDMOGrowth_MB)
        oLogFile.FileGrowth = LogFile.lGrow
        If LogFile.bMaxG Then oLogFile.MaximumSize = LogFile.lREstrictG
    Else
        oLogFile.FileGrowth = 0
    End If
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 And bAdd Then
        On Local Error Resume Next
        oTransactionLog.LogFiles.Add oLogFile
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    
End Sub

Private Sub LoadProperty()
    
    Dim lErr As Long, sErr As String
    Dim bExit As Boolean
    Dim oDb As SQLDMO.Database2

    fGenOpt.Caption = "Database '" & sDbName & "'"
    fGenOpt.DBName() = sDbName
    fGenOpt.Action() = act_DBOption
    fGenOpt.Show vbModal, Me
    
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        bExit = oDb.DBOption.Offline = True
'        lErr = kErrDB_Offline
'        sErr = ReplaceMsg(MyLoadResString(k_DB_is_OffLine), Array("1%", "[#]"), Array(sDbName, vbCrLf))
'        MsgBox sErr, vbInformation Or vbOKOnly, App.EXEName
'        bExit = True
    Else
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
        bExit = True
    End If
    Set oDb = Nothing

    If bExit Then
        Unload Me
    Else
        LoadDB sDbName
    End If
    
End Sub
Private Sub OpenDialog()

    Dim bCancel As Boolean
    Dim bErr As Boolean
    Dim sTarget As String
    Dim sInit As String
    Dim lErr As Long, sErr As String
    Dim iPos As Integer
    Dim sDbName As String
    Dim iResourceStr As Integer
    

    If miFlex = 0 Then
        iResourceStr = k_LocateDB_file
        sInit = dataFile(miRow - 1).sFilePos
        If Len(sInit) = 0 Then
            If miRow - 1 > 0 Then
                sInit = dataFile(miRow - 2).sFilePos
            Else
                sInit = sSqlPath
            End If
            iPos = InStrRev(sInit, "\")
            If iPos <> 0 Then sInit = Left$(sInit, iPos)
        Else
            iPos = InStrRev(sInit, "\")
            If iPos <> 0 Then
                sInit = Left$(sInit, iPos)
                sDbName = Mid$(dataFile(miRow - 1).sFilePos, iPos + 1)
            End If
        End If
    Else
        iResourceStr = k_LocateLog_file
        sInit = LogFile(miRow - 1).sFilePos
        If Len(sInit) = 0 Then
            If miRow - 1 > 0 Then
                sInit = LogFile(miRow - 2).sFilePos
            Else
                sInit = sSqlPath
            End If
            iPos = InStrRev(sInit, "\")
            If iPos <> 0 Then sInit = Left$(sInit, iPos)
        Else
            iPos = InStrRev(sInit, "\")
            If iPos <> 0 Then
                sInit = Left$(sInit, iPos)
                sDbName = Mid$(LogFile(miRow - 1).sFilePos, iPos + 1)
            End If
        End If
    End If
    
    Dim fBrowse As fBr4Fold
    Set fBrowse = New fBr4Fold
    
    With fBrowse
        .Init enB4isSelectDir, MyLoadResString(iResourceStr), sInit
        bErr = .ErrLoadingInit()
        If Not bErr Then
            .Show vbModal, Me
            bCancel = .Cancelled()
            If Not bCancel Then sTarget = .FullPath()
        Else
            .GetError lErr, sErr
        End If
    End With
    
    Unload fBrowse
    Set fBrowse = Nothing
    
    If Not bErr And Not bCancel Then
        If Right$(sTarget, 1) <> "\" Then sTarget = sTarget & "\"
        If miFlex = 0 Then
            dataFile(miRow - 1).sFilePos = sTarget & sDbName
        Else
            LogFile(miRow - 1).sFilePos = sTarget & sDbName
        End If
        Flex(miFlex).TextMatrix(miRow, miCol + 1) = sTarget & sDbName
    End If
        
End Sub

Private Sub PrivateLoadPermission()

    Dim oDb As SQLDMO.Database2
    On Local Error Resume Next
    Set oDb = objServer.Databases(sDbName)
    On Local Error GoTo 0

    If Not oDb Is Nothing Then
        fPrivil.Caption = MyLoadResString(k_Permissions) & " - DB: [" & sDbName & "]"
        fPrivil.SetObj oDb, Nothing, sDbName, act_DBOption
        Screen.MousePointer = vbDefault
        fPrivil.Show vbModal
    End If
    Set oDb = Nothing
End Sub

Private Sub InitCombo()
    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    Dim oNList As SQLDMO.NameList
    Dim i As Integer
    
    
    Set oQry = objServer.EnumCollations()
    If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
    If IsArray(vVar) Then
        With cbo(0)
            ReDim SortRules(UBound(vVar, 2) + 1) As String
            SortRules(0) = "<Default>"
            .AddItem SortRules(0)
            For i = 0 To UBound(vVar, 2)
                SortRules(i + 1) = vVar(0, i)
                .AddItem SortRules(i + 1)
            Next
            .ListIndex = 0
        End With
    End If
    
    Set oNList = objServer.ListCompatibilityLevels
    For i = 1 To oNList.Count
        With cbo(1)
            .AddItem oNList.Item(i)
            .ListIndex = 0
        End With
    Next

End Sub
