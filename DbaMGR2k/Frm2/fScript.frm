VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ClipControls    =   0   'False
   Icon            =   "fScript.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   1440
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   360
      Index           =   1
      Left            =   6240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   7335
      Begin VB.CommandButton cmd 
         Caption         =   "<< Rimuovi"
         Height          =   360
         Index           =   4
         Left            =   6000
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Aggiungi >>"
         Height          =   360
         Index           =   3
         Left            =   6000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Preview"
         Height          =   360
         Index           =   2
         Left            =   6000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2020
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3572
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Index           =   2
         Left            =   3720
         TabIndex        =   7
         Top             =   2180
         Width           =   3500
         _ExtentX        =   6165
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   7335
      Begin VB.Frame Frame 
         Caption         =   "Frame1"
         Height          =   1335
         Index           =   4
         Left            =   3840
         TabIndex        =   15
         Top             =   2640
         Width           =   3255
         Begin VB.OptionButton opt 
            Caption         =   "Option1"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   2895
         End
         Begin VB.OptionButton opt 
            Caption         =   "Option1"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   2895
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Frame1"
         Height          =   1335
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   3255
         Begin VB.OptionButton opt 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton opt 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   2895
         End
         Begin VB.OptionButton opt 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2895
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2020
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3572
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame2"
      Height          =   4335
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   7335
      Begin RichTextLib.RichTextBox rtf 
         Height          =   2415
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"fScript.frx":000C
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   1290
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2275
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.TabStrip tStrip 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Generale"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formattazione"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Opzioni"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblScript 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "fScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sDbName As String
Private mAction As optAction
Private mlErr As Long
Private msErr As String

Private WithEvents objTrans As SQLDMO.Transfer2
Attribute objTrans.VB_VarHelpID = -1

Private scriptProp() As elemFlexG
Private formatProp() As elemFlexG
Private OtherProp() As elemFlexG

Private collNoScript As Collection
Private coll2Script As Collection

Public Property Let TypeOfScript(ByVal iVal As optAction)
    mAction = iVal
End Property

Public Property Let DB2Script(ByVal sVal As String)
    sDbName = sVal
End Property
Public Sub Init(Optional ByVal bAllTB As Boolean = True, _
                Optional ByVal bAllView As Boolean = True, _
                Optional ByVal bAllSP As Boolean = True, _
                Optional ByVal bAllDEF As Boolean = True, _
                Optional ByVal bAllRule As Boolean = True, _
                Optional ByVal bAllUDT As Boolean = True, _
                Optional ByVal bAllFunc As Boolean = True, _
                Optional ByVal bJustLoaded As Boolean = False)

    Dim oDb As SQLDMO.Database2
    Dim i As Integer, bAdd As Boolean
    Dim oObj As Object
    
    Screen.MousePointer = vbHourglass
        
    If Not bJustLoaded Then
        Set collNoScript = New Collection
        Set coll2Script = New Collection
    End If
    
    If Len(sDbName) Then
        On Local Error Resume Next
        Set oDb = objServer.Databases(sDbName)
        mlErr = Err.Number
        msErr = Err.Description
        On Local Error GoTo 0
    Else
        On Local Error Resume Next
        Err.Raise 380
        mlErr = Err.Number
        msErr = Err.Description
        On Local Error GoTo 0
    End If
    
    If mlErr = 0 Then
        On Local Error Resume Next
        If UBound(scriptProp) = 0 Or Err <> 0 Then LoadScriptProp True
        On Local Error GoTo 0
        
        If mAction = act_DBOption Then
            For Each oObj In oDb.Tables
                If oObj.SystemObject = 0 Then Add2Coll bAllTB, opt_Tbl, oObj
            Next
            For Each oObj In oDb.Views
                If oObj.SystemObject = 0 Then Add2Coll bAllView, opt_View, oObj
            Next
            For Each oObj In oDb.StoredProcedures
                If oObj.SystemObject = 0 Then Add2Coll bAllSP, opt_SP, oObj
            Next
            For Each oObj In oDb.Defaults
                Add2Coll bAllDEF, opt_Def, oObj
            Next
            For Each oObj In oDb.Rules
                Add2Coll bAllRule, opt_Rule, oObj
            Next
            For Each oObj In oDb.UserDefinedDatatypes
                Add2Coll bAllUDT, opt_UDT, oObj
            Next
            For Each oObj In oDb.UserDefinedFunctions
                Add2Coll bAllFunc, opt_Func, oObj
            Next

        Else
            For Each oObj In oDb.Tables
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_Tbl Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If oObj.SystemObject = 0 And bAdd Then Add2Coll bAllTB, opt_Tbl, oObj
            Next
            For Each oObj In oDb.Views
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_View Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If oObj.SystemObject = 0 And bAdd Then Add2Coll bAllView, opt_View, oObj
            Next
            For Each oObj In oDb.StoredProcedures
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_SP Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If oObj.SystemObject = 0 And bAdd Then Add2Coll bAllSP, opt_SP, oObj
            Next
            For Each oObj In oDb.Defaults
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_Def Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If bAdd Then Add2Coll bAllDEF, opt_Def, oObj
            Next
            
            For Each oObj In oDb.Rules
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_Rule Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If bAdd Then Add2Coll bAllRule, opt_Rule, oObj
            Next
            For Each oObj In oDb.UserDefinedDatatypes
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_UDT Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If bAdd Then Add2Coll bAllUDT, opt_UDT, oObj
            Next
            For Each oObj In oDb.UserDefinedFunctions
                bAdd = True
                For i = 1 To coll2Script.Count
                    If coll2Script.Item(i).oName = oObj.Name And coll2Script.Item(i).oType = opt_Func Then
                        bAdd = False
                        Exit For
                    End If
                Next
                If bAdd Then Add2Coll bAllFunc, opt_Func, oObj
            Next
            
        End If
    End If
    Set oObj = Nothing
    Set oDb = Nothing
    
    If mlErr = 0 Then
        SetPropertyValue
        LoadGrids 0
        LoadGrids 1
        LoadGrids 2
        LoadGrids 3
        LoadGrids 4
        cmd(3).Enabled = collNoScript.Count
        cmd(1).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
        cmd(4).Enabled = coll2Script.Count
        cmd(2).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
    End If
    Screen.MousePointer = vbDefault

End Sub
Private Sub Add2Coll(ByVal b2Script As Boolean, ByVal iType As optScript, obj2Load As Object)

    Dim oObj As cScrpO
    Set oObj = New cScrpO
        
    With obj2Load
        oObj.oName() = obj2Load.Name
        oObj.oOwner() = obj2Load.Owner
        oObj.oType() = iType
    End With
    
    On Local Error Resume Next
    If b2Script Then
        coll2Script.Add oObj, oObj.oName()
    Else
        collNoScript.Add oObj, oObj.oName()
    End If
    Debug.Print Err.Number
    Debug.Print Err.Description
    On Local Error GoTo 0
    Set oObj = Nothing
    
    Debug.Print coll2Script.Count
    Debug.Print collNoScript.Count
    
End Sub
Public Property Let ScriptThis(oObj As cScrpO)

    Set collNoScript = New Collection
    Set coll2Script = New Collection

    On Local Error Resume Next
    coll2Script.Add oObj, oObj.oName()
    mlErr = Err.Number
    msErr = Err.Description
    On Local Error GoTo 0

End Property
Private Sub LoadScriptProp(Optional ByVal bAll = False)
    
    Dim i As Integer
    ReDim scriptProp(7) As elemFlexG
    For i = 0 To UBound(scriptProp)
        scriptProp(i).sField = MyLoadResString(k_All_Objects + i)
        scriptProp(i).iType = am_Bool
        scriptProp(i).vVal = bAll
    Next
    ReDim formatProp(5) As elemFlexG
    For i = 0 To UBound(formatProp)
        formatProp(i).sField = MyLoadResString(k_Generate_CREATE_DATABASE_command + i)
        formatProp(i).iType = am_Bool
        formatProp(i).vVal = IIf(i < 2, True, False)
    Next
    ReDim OtherProp(6) As elemFlexG
    For i = 0 To UBound(OtherProp)
        OtherProp(i).sField = MyLoadResString(k_Script_DB_Users_Roles + i)
        OtherProp(i).iType = am_Bool
        OtherProp(i).vVal = False
    Next
        
End Sub
Private Sub cmd_Click(Index As Integer)
    
    If cmd(Index).Enabled Then
        Select Case Index
            Case 0  'Unload
                Unload Me
            Case 1, 2 '1=Ok, 2=Preview
                Execute Index = 2
            Case 3, 4
                ChangeFlex Index - 2
        End Select
    End If
End Sub
Private Sub ChangeFlex(ByVal Index As Integer)

    Dim oObj As cScrpO
    Dim iFlex As Integer
    
    If Index = 1 Then
        Set oObj = collNoScript.Item(Flex(Index).Row)
        iFlex = 2
    Else
        Set oObj = coll2Script.Item(Flex(Index).Row)
        iFlex = 1
    End If
    
    If Index = 1 Then
        coll2Script.Add oObj, oObj.oName()
    Else
        collNoScript.Add oObj, oObj.oName()
    End If
    If Err = 0 Then
        If Index = 1 Then
            collNoScript.Remove Flex(Index).Row
        Else
            coll2Script.Remove Flex(Index).Row
        End If
    End If
    If Err = 0 Then
        With Flex(iFlex)
            .AddItem vbTab & oObj.oName() & vbTab & oObj.oOwner()
            
            .Row = .Rows - 1
            .Col = 0
            .RowData(.Row) = Flex(Index).RowData(Flex(Index).Row)
            Set Flex(iFlex).CellPicture = LoadResPicture(.RowData(.Row), vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
            If .Rows > 1 Then .FixedRows = 1
            .RowSel = .Row
            .ColSel = .Col
        End With
        
        If Flex(Index).Rows > 2 Then
            Flex(Index).RemoveItem Flex(Index).Row
        Else
            Flex(Index).FixedRows = 0
            Flex(Index).RemoveItem 1
            Flex(Index).Row = 0
            Flex(Index).Col = 1
            Flex(Index).CellBackColor = vbInactiveBorder
            Flex(Index).Col = 2
            Flex(Index).CellBackColor = vbInactiveBorder
        End If
    End If
    
    On Local Error GoTo 0
            
    SetPropertyValue
    
    cmd(2).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
    cmd(1).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
    cmd(3).Enabled = collNoScript.Count
    
    cmd(4).Enabled = coll2Script.Count
    
    Set oObj = Nothing
    LoadGrids 0
    
End Sub
Private Sub SetPropertyValue()

    Dim i As Integer
    scriptProp(0).vVal = collNoScript.Count = 0
    
    scriptProp(1).vVal = True
    For i = 1 To collNoScript.Count
        If collNoScript.Item(i).oType() = opt_Tbl Then
            scriptProp(1).vVal = False
            Exit For
        End If
    Next
    
    scriptProp(2).vVal = True
    For i = 1 To collNoScript.Count
        If collNoScript.Item(i).oType() = opt_View Then
            scriptProp(2).vVal = False
            Exit For
        End If
    Next
    
    scriptProp(3).vVal = True
    For i = 1 To collNoScript.Count
        If collNoScript.Item(i).oType() = opt_SP Then
            scriptProp(3).vVal = False
            Exit For
        End If
    Next
    
    scriptProp(4).vVal = True
    For i = 1 To collNoScript.Count
        Debug.Print collNoScript.Item(i).oName
        If collNoScript.Item(i).oType() = opt_Def Then
            scriptProp(4).vVal = False
            Exit For
        End If
    Next
    scriptProp(5).vVal = True
    For i = 1 To collNoScript.Count
        If collNoScript.Item(i).oType() = opt_Rule Then
            scriptProp(5).vVal = False
            Exit For
        End If
    Next
    scriptProp(6).vVal = True
    For i = 1 To collNoScript.Count
        Debug.Print collNoScript.Item(i).oName
        If collNoScript.Item(i).oType() = opt_UDT Then
            scriptProp(6).vVal = False
            Exit For
        End If
    Next
    scriptProp(7).vVal = True
    For i = 1 To collNoScript.Count
        Debug.Print collNoScript.Item(i).oName
        If collNoScript.Item(i).oType() = opt_Func Then
            scriptProp(7).vVal = False
            Exit For
        End If
    Next
    
End Sub
Private Sub Flex_DblClick(Index As Integer)
    If Index = 0 Or Index > 2 Then
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then
            FlexEdit vbKeySpace, Index
        End If
    Else
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then ChangeFlex Index
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
Private Sub Form_Load()

    mAction = act_Null
   
    Set collNoScript = New Collection
    Set coll2Script = New Collection
    
    CenterForm Me
    cmd(0).Caption = MyLoadResString(k_Cancel)
    tStrip.Tabs(1).Caption = MyLoadResString(k_General)
    tStrip.Tabs(2).Caption = MyLoadResString(k_Formatting)
    tStrip.Tabs(3).Caption = MyLoadResString(k_Options)
    cmd(3).Caption = MyLoadResString(k_Add_All)
    cmd(4).Caption = MyLoadResString(k_Remove_All)
    Frame(1).Caption = MyLoadResString(k_Scripting_Opt_how_2_scrip)
    
    Frame(3).Caption = MyLoadResString(k_File_Format)
    opt(0).Caption = MyLoadResString(k_MS_DOS_Text_OEM)
    opt(1).Caption = MyLoadResString(k_Windows_Text_Ansi)
    opt(2).Caption = MyLoadResString(k_International_Text_Unicode)
    Frame(4).Caption = MyLoadResString(k_Files_to_Generate)
    opt(3).Caption = MyLoadResString(k_Create_One_File)
    opt(4).Caption = MyLoadResString(k_Create_One_File_per_Object)
           
End Sub
Public Property Get ErrCode() As Long
    ErrCode = mlErr
End Property
Public Property Get ErrMsg() As String
    ErrMsg = msErr
End Property

Public Property Get ErrOccured() As Boolean

    If mlErr = 0 And mAction = act_Null Then
        On Local Error Resume Next
        Err.Raise 380
        mlErr = Err.Number
        msErr = Err.Description
        On Local Error GoTo 0
    End If
    ErrOccured = mlErr <> 0
End Property
Private Sub Form_Unload(Cancel As Integer)
    
    Set collNoScript = Nothing
    Set coll2Script = Nothing
    Set objTrans = Nothing

    Set fScript = Nothing
End Sub

Private Sub LoadGrids(ByVal Index As Integer)
        
    Dim i As Integer, y As Integer, bVal As Boolean
    Dim sBuf As String
    Dim iCount As Integer
    Dim lErr As Long, sErr As String

    Select Case Index
        Case 0, 3, 4
            With Flex(Index)
                .Redraw = False
                .ScrollBars = flexScrollBarBoth
                .Clear
                .FixedCols = 0
                .FixedRows = 0
                
                .Cols = 2
                .Rows = 1
                
                .ScrollBars = flexScrollBarBoth
                .TextMatrix(0, 0) = MyLoadResString(k_Property)
                .TextMatrix(0, 1) = MyLoadResString(k_Value)
                
                Select Case Index
                    Case 0
                        iCount = UBound(scriptProp)
                    Case 3
                        iCount = UBound(formatProp)
                    Case 4
                        iCount = UBound(OtherProp)
                End Select
                
                For i = 0 To iCount
                    Select Case Index
                        Case 0
                            sBuf = scriptProp(i).sField
                            y = scriptProp(i).iType
                        Case 3
                            sBuf = formatProp(i).sField
                            y = formatProp(i).iType
                        Case 4
                            sBuf = OtherProp(i).sField
                            y = OtherProp(i).iType
                    End Select
                
                    Select Case y
                        Case am_Bool
                        Case am_String
                            Select Case Index
                                Case 0
                                    sBuf = sBuf & vbTab & scriptProp(i).vVal
                                Case 3
                                    sBuf = sBuf & vbTab & formatProp(i).vVal
                                Case 4
                                    sBuf = sBuf & vbTab & OtherProp(i).vVal
                            End Select
                    End Select
                    On Local Error Resume Next
                    .AddItem sBuf
                    If y = am_Bool Then
                        .Row = .Rows - 1
                        .Col = .Cols - 1
                        Select Case Index
                            Case 0
                                bVal = scriptProp(i).vVal
                            Case 3
                                bVal = formatProp(i).vVal
                            Case 4
                                bVal = OtherProp(i).vVal
                        End Select
                                                
                        Set Flex(Index).CellPicture = LoadResPicture(IIf(Abs(bVal), k_ResChkON, k_ResChkOFF), vbResBitmap)
                        .CellPictureAlignment = flexAlignCenterCenter
                        
                    End If
                    lErr = Err.Number
                    sErr = Err.Description
                    If lErr <> 0 Then Exit For
                    On Local Error GoTo 0
                Next
                If lErr <> 0 Then
                    .RemoveItem .Rows - 1
                    .RemoveItem .Rows - 1
                    sErr = sErr & vbCrLf & ReplaceMsg(MyLoadResString(k_Script_OutOfMemory), Array("1%", "2%", "|"), Array(i - 2, iCount - 2, vbCrLf))
                End If
                                
                .FixedCols = 1
                .AllowUserResizing = flexResizeColumns
                .ScrollBars = flexScrollBarBoth
                If .Rows > 1 Then .FixedRows = 1
                .Row = 1
                .Redraw = True
                Select Case Index
                    Case 0, 3
                        .ColWidth(0) = (.Width / 3) * 2
                        SettaLargFlex Flex(Index), 0, 300
                    Case 4
                        .ColWidth(0) = (.Width / 8) * 7
                        SettaLargFlex Flex(Index), 1, 300
                End Select
                
                .TopRow = 1

            End With
        Case 1, 2
            With Flex(Index)
                .Redraw = False
                .ScrollBars = flexScrollBarBoth
                .Clear
                .FixedCols = 0
                .FixedRows = 0
                
                .Cols = 3
                .Rows = 1
                
                .ScrollBars = flexScrollBarBoth
                
                .TextMatrix(0, 0) = MyLoadResString(k_Type)
                .TextMatrix(0, 1) = MyLoadResString(k_Name)
                .TextMatrix(0, 2) = MyLoadResString(k_Ownership)

                Select Case Index
                    Case 1
                        iCount = collNoScript.Count
                    Case 2
                        iCount = coll2Script.Count
                End Select
                
                For i = 1 To iCount
                    Select Case Index
                        Case 1
                            sBuf = vbTab & collNoScript.Item(i).oName() & vbTab & collNoScript.Item(i).oOwner()
                            y = collNoScript.Item(i).oType()
                        Case 2
                            sBuf = vbTab & coll2Script.Item(i).oName() & vbTab & coll2Script.Item(i).oOwner()
                            y = coll2Script.Item(i).oType()
                            
                    End Select
                
                    .AddItem sBuf
                    
                    .Row = .Rows - 1
                    .Col = 0
                                                            
                    .RowData(.Row) = y
                    On Local Error Resume Next
                    Set Flex(Index).CellPicture = LoadResPicture(.RowData(.Row), vbResBitmap)
                    lErr = Err.Number
                    sErr = Err.Description
                    If lErr <> 0 Then Exit For
                    On Local Error GoTo 0
                    .CellPictureAlignment = flexAlignCenterCenter
                Next
                If lErr <> 0 Then
                    .RemoveItem .Rows - 1
                    .RemoveItem .Rows - 1
                    sErr = sErr & vbCrLf & ReplaceMsg(MyLoadResString(k_Script_OutOfMemory), Array("1%", "2%", "|"), Array(i - 2, iCount - 2, vbCrLf))
                End If
                                    
                .FixedCols = 1
                .AllowUserResizing = flexResizeColumns
                .ScrollBars = flexScrollBarBoth
                If .Rows > 1 Then
                    .FixedRows = 1
                    .Row = 1
                    bVal = True
                Else
                    .Row = 0
                    .Col = 1
                    .CellBackColor = vbInactiveBorder
                    .Col = 2
                    .CellBackColor = vbInactiveBorder
                    bVal = False
                End If
                .ColWidth(0) = 500
                .ColWidth(1) = (.Width / 2)

                .Redraw = True
                SettaLargFlex Flex(Index), 2, 300
                If bVal Then
                    .TopRow = 1
                Else
                    .TopRow = 0
                End If
            End With
        
    End Select
    If lErr <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbCrLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
    
End Sub

Private Sub objTrans_PercentCompleteAtStep(ByVal Message As String, ByVal Percent As Long)
    pBar.Value = Percent
    lblScript.Caption = Message
    DoEvents
End Sub
Private Sub objTrans_ScriptTransferPercentComplete(ByVal Message As String, ByVal Percent As Long)
    pBar.Value = Percent
    lblScript.Caption = Message
    DoEvents
End Sub
Private Sub objTrans_StatusMessage(ByVal Message As String)
    lblScript.Caption = Message
    DoEvents
End Sub
Private Sub objTrans_TransferPercentComplete(ByVal Message As String, ByVal Percent As Long)
    pBar.Value = Percent
    lblScript.Caption = Message
    DoEvents
End Sub

Private Sub rtf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim P As POINTAPI
    Dim i As Integer
    
    If Button = vbRightButton Then
        Dim M As New cPUpMenu
        M.Add MyLoadResString(k_MnuCut), 1
        M.Add MyLoadResString(k_MnuCopy), 2
        M.Add MyLoadResString(k_MnuPaste), 3, , bDisable:=Len(Clipboard.GetText) = 0 Or rtf.Locked
        M.Add "", 4, bSeparator:=True
        M.Add MyLoadResString(k_MnuSelectAll), 5
        
        Call GetCursorPos(P)
        i = M.Popup(Me.hWnd, P.x, P.y)
        M.Clear
        Set M = Nothing
        
        Select Case i
            Case 1
                Clipboard.SetText rtf.SelText, vbCFText
                rtf.SelText = ""
            Case 2
                Clipboard.SetText rtf.SelText, vbCFText
            Case 3
                rtf.SelRTF = Clipboard.GetText(vbCFText)
            Case 5
                rtf.SelStart = 0
                rtf.SelLength = Len(rtf.Text)
        End Select
    End If

End Sub

Private Sub tStrip_Click()

    Dim i As Integer
            
    i = tStrip.SelectedItem.Index - 1
    Frame(i).Visible = True
    Frame(i).ZOrder 0
    Select Case i
        Case 0
            Flex(0).SetFocus
        Case 1
            rtf.SetFocus
            DoEvents
    End Select
        
End Sub
Private Sub FlexEdit(KeyAscii As Integer, ByVal Index As Integer)
    
    Dim iRow As Integer
    Dim iImg As Integer
    Dim i As Integer
    
    If Flex(Index).CellBackColor = vbInactiveBorder Then Exit Sub
    
    iRow = Flex(Index).Row - 1
    

    Select Case KeyAscii
        Case vbKeySpace, 0
        
            Select Case Index
                Case 0
                    iImg = 1 - (-CBool(scriptProp(iRow).vVal))
                    scriptProp(iRow).vVal = CBool(iImg)
                Case 3
                    iImg = 1 - (-CBool(formatProp(iRow).vVal))
                    formatProp(iRow).vVal = CBool(iImg)
                    If iRow = 0 Then
                        cmd(2).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
                        cmd(1).Enabled = coll2Script.Count Or CBool(formatProp(0).vVal)
                    End If
                Case 4
                    iImg = 1 - (-CBool(OtherProp(iRow).vVal))
                    OtherProp(iRow).vVal = CBool(iImg)
            End Select
            Set Flex(Index).CellPicture = LoadResPicture(IIf(iImg, k_ResChkON, k_ResChkOFF), vbResBitmap)
            Flex(Index).CellPictureAlignment = flexAlignCenterCenter
            
            Select Case Index
                Case 0
                    Select Case iRow
                        Case 0
                            For i = 1 To UBound(scriptProp)
                                scriptProp(i).vVal = CBool(iImg)
                            Next
                            Init CBool(iImg), CBool(iImg), CBool(iImg), CBool(iImg), CBool(iImg), CBool(iImg), CBool(iImg)
                        Case Else
                            If scriptProp(1).vVal And scriptProp(2).vVal And scriptProp(3).vVal _
                                And scriptProp(4).vVal And scriptProp(5).vVal And scriptProp(6).vVal Then scriptProp(0).vVal = True
                            Init scriptProp(1).vVal, scriptProp(2).vVal, scriptProp(3).vVal, scriptProp(4).vVal, scriptProp(5).vVal, scriptProp(6).vVal, scriptProp(7).vVal
                    End Select
            End Select
        
    End Select

End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim iCode As Integer
    
    If Index = 0 Or Index > 2 Then
        iCode = KeyAscii
        KeyAscii = 0
        FlexEdit iCode, Index
    Else
        KeyAscii = 0
    End If
    
End Sub
Private Sub ReadFromFile(sTempFile As String)

    Dim numOfBytes As Long
    Dim buf As String
    Dim hFile As Long
    On Local Error Resume Next
    hFile = lopen(sTempFile, 0)
    numOfBytes = GetFileSize(hFile, ByVal 0&)
    buf = Space(numOfBytes)
    If Err.Number <> 0 Then
        buf = MyLoadResString(k_Error_reading_System_Handle) & sTempFile _
                    & vbCrLf & Err.Number & "-" & Err.Description
    Else
        lread hFile, buf, numOfBytes
    End If
    rtf.Text = rtf.Text & IIf(Len(rtf.Text), vbCrLf, "") & buf
    Debug.Print buf

    lclose (hFile)
    DeleteFile (sTempFile)
    On Local Error GoTo 0

End Sub

Private Sub Execute(ByVal bPreview As Boolean)
        
    Dim oDb As SQLDMO.Database2
    Dim oTb As SQLDMO.Table
    Dim oView As SQLDMO.View
    Dim oSP As SQLDMO.StoredProcedure
    Dim oFunc As SQLDMO.UserDefinedFunction
    Dim x As SQLDMO_SCRIPT_TYPE
    Dim y As SQLDMO_SCRIPT2_TYPE
    
    Dim i As Integer
    Dim iFileType As Integer
    Dim b1File As Boolean
    Dim sFileName As String
    Dim sDBFileName As String
    Dim sFileOut As String
    Dim lErr As Long, sErr As String
    
    Screen.MousePointer = vbHourglass
    rtf.Text = ""
    
    If Not bPreview Then
        b1File = opt(3).Value
    Else
        b1File = True
    End If
    
    
    If bPreview Then
        sFileName = GetTempFile("tmp")
        If Len(sFileName) = 0 Then GoTo ExitHere
    Else
        Screen.MousePointer = vbDefault
                
        If b1File Then
            sFileName = GetTempFile("tmp")
            sFileOut = GetFileOutPut()
        Else
            sFileOut = GetDirOutPut()
            sFileName = sFileOut
        End If
        Screen.MousePointer = vbHourglass
        If Len(sFileName) = 0 Or Len(sFileOut) = 0 Then GoTo ExitHere
    End If
    
    'If mAction = act_DBOption Then
    If CBool(formatProp(0).vVal) Then
        If bPreview Then
            sDBFileName = GetTempFile("tmp")
        ElseIf Not b1File Then
            i = InStrRev(sFileName, "\")
            'If i <> 0 Then sDBFileName = Left$(sFileName, i) & "dbo." & sDbName & ".txt"
            If i <> 0 Then sDBFileName = Left$(sFileName, i) & GetCreateOwner(sDbName) & "." & sDbName & ".txt"
        End If
        If Len(sDBFileName) = 0 Then sDBFileName = GetTempFile("tmp")
    End If
   
    On Local Error GoTo ErrTrap
    If Len(sDbName) Then
        Set oDb = objServer.Databases(sDbName)
        Set objTrans = New SQLDMO.Transfer2
    Else
        Err.Raise 380
    End If
    
    If Not bPreview Then
        For i = 0 To 2
            If opt(i).Value Then
                iFileType = i
                Exit For
            End If
        Next
    
        Select Case iFileType
            Case 1
                y = SQLDMOScript2_AnsiFile
            Case 2
                y = SQLDMOScript2_UnicodeFile
        End Select
    End If
    x = SQLDMOScript_Default

    objTrans.CopyAllDefaults = CBool(scriptProp(4).vVal)
    objTrans.CopyAllRules = CBool(scriptProp(5).vVal)
    objTrans.CopyAllUserDefinedDatatypes = CBool(scriptProp(6).vVal)
    objTrans.CopyAllFunctions = CBool(scriptProp(7).vVal)
    
    
    objTrans.IncludeUsers = CBool(OtherProp(0).vVal)
    objTrans.IncludeLogins = CBool(OtherProp(1).vVal)
    
    
    If b1File Then x = x Or SQLDMOScript_AppendToFile
    If CBool(formatProp(1).vVal) Then x = x Or SQLDMOScript_Drops
    objTrans.IncludeDependencies = CBool(formatProp(2).vVal)
    If CBool(formatProp(3).vVal) Then x = x Or SQLDMOScript_IncludeHeaders
    
    If CBool(formatProp(4).vVal) Then y = y Or SQLDMOScript2_ExtendedProperty
    If CBool(formatProp(5).vVal) Then y = y Or SQLDMOScript2_70Only
    
    
    
    If CBool(OtherProp(2).vVal) Then x = x Or SQLDMOScript_ObjectPermissions Or SQLDMOScript_DatabasePermissions Or SQLDMOScript_Permissions
    If CBool(OtherProp(3).vVal) Then x = x Or SQLDMOScript_Indexes Or SQLDMOScript_ClusteredIndexes
    If CBool(OtherProp(4).vVal) Then y = y Or SQLDMOScript2_FullTextCat Or SQLDMOScript2_FullTextIndex
    If CBool(OtherProp(5).vVal) Then x = x Or SQLDMOScript_Triggers
    If CBool(OtherProp(6).vVal) Then x = x Or SQLDMOScript_DRI_All
    
'    objTrans.ScriptType = x
'    objTrans.Script2Type = Y
    
    If CBool(scriptProp(0).vVal) Then
        objTrans.CopyAllObjects = True
    Else
    
        If CBool(scriptProp(1).vVal) Then
            objTrans.CopyAllTables = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_Tbl Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_UserTable, coll2Script.Item(i).oOwner()
                    
'                    Set oTB = objServer.Databases(sDBName).Tables(coll2Script.Item(i).oName())
'                    Debug.Print oTB.Script(SQLDMOScript_OwnerQualify Or SQLDMOScript_Default Or SQLDMOScript_NoDRI)
'                    Debug.Print oTB.Script(SQLDMOScript_OwnerQualify Or SQLDMOScript_NoDRI Or SQLDMOScript_PrimaryObject)
                    
                End If
            Next
        End If
        If CBool(scriptProp(2).vVal) Then
            objTrans.CopyAllViews = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_View Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_View, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        If CBool(scriptProp(3).vVal) Then
            objTrans.CopyAllStoredProcedures = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_SP Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_StoredProcedure, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        If CBool(scriptProp(4).vVal) Then
            objTrans.CopyAllDefaults = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_Def Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_Default, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        If CBool(scriptProp(5).vVal) Then
            objTrans.CopyAllRules = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_Rule Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_Rule, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        If CBool(scriptProp(6).vVal) Then
            objTrans.CopyAllUserDefinedDatatypes = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_UDT Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_UserDefinedDatatype, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        If CBool(scriptProp(7).vVal) Then
            objTrans.CopyAllFunctions = True
        Else
            For i = 1 To coll2Script.Count
                If coll2Script.Item(i).oType() = opt_Func Then
                    objTrans.AddObjectByName coll2Script.Item(i).oName(), SQLDMOObj_UserDefinedFunction, coll2Script.Item(i).oOwner()
                End If
            Next
        End If
        
    End If

    objTrans.CopyData = SQLDMOCopyData_False
    objTrans.ScriptType = x
    objTrans.Script2Type = y
    'If mAction = act_DBOption Then
    
    
    If CBool(formatProp(0).vVal) Then
        If Len(sDBFileName) Then Call oDb.Script(x, sDBFileName, y)
    End If
            
    pBar.Visible = True
    lblScript.Visible = True
    
    Call oDb.ScriptTransfer(objTrans, IIf(b1File, SQLDMOXfrFile_SingleFile, SQLDMOXfrFile_SingleFilePerObject), sFileName)
    
    If bPreview Then
        If Len(sDBFileName) Then ReadFromFile sDBFileName
        ReadFromFile sFileName
        Set tStrip.SelectedItem = tStrip.Tabs(2)
    Else
        If b1File Then
            If Len(sDBFileName) Then TransferDBFile sDBFileName, sFileOut, True
            If Len(sFileName) Then TransferDBFile sFileName, sFileOut
        Else
            sFileName = ""
            sDBFileName = ""
        End If
    End If
    
ExitHere:
    If Len(sFileName) Then DeleteFile (sFileName)
    If Len(sDBFileName) Then DeleteFile (sDBFileName)
    
    Set oTb = Nothing
    Set oView = Nothing
    Set oSP = Nothing
    Set oFunc = Nothing
    Set oDb = Nothing
    Set objTrans = Nothing
    
    pBar.Visible = False
    lblScript.Visible = False


    Screen.MousePointer = vbDefault
    Exit Sub
    
    
ErrTrap:
    lErr = Err.Number
    sErr = Err.Description
    Screen.MousePointer = vbDefault
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
    Resume ExitHere

End Sub

Private Function GetFileOutPut() As String

    Dim sFile As String
    Dim lErr As Long, sErr As String
    Dim obj As cOpen
    Set obj = New cOpen
    
    obj.Title = MyLoadResString(k_Save_As)
    obj.Def_Dir() = App.Path
    obj.Filter() = "Query File (*.Sql)" & Chr$(0) & "*.sql" & Chr$(0) & "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    obj.f_Name = MyLoadResString(k_New) & ".sql"
    obj.formHwnd() = Me.hWnd

    obj.Do_Save
    lErr = obj.ErrCode()
    If lErr = 0 Then
        sFile = obj.FullPath()
        GetFileOutPut = sFile
    Else
        If lErr = 1 Then
            lErr = 0
        Else
            sErr = obj.ErrMsg()
        End If
    End If
    Set obj = Nothing
    
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If

End Function
Private Function GetDirOutPut() As String

    Dim sBuf As String
    Dim a As BROWSEINFO
    
    Dim pidlBrowse As Long
    
    a.lpszTitle = MyLoadResString(k_Save_Scripts_in_Directory)
        
    a.hOwner = Me.hWnd
    pidlBrowse = SHBrowseForFolder(a)
    sBuf = GetPathFromPIDL(pidlBrowse)
    If Len(sBuf) Then
        If Right$(sBuf, 1) <> "\" Then sBuf = sBuf & "\"
        GetDirOutPut = sBuf
    End If

End Function
Private Sub TransferDBFile(ByVal sFileIn As String, ByVal sFileOut As String, Optional ByVal bInit As Boolean = False)

    Dim iFileIn As Integer, iFileOut As Integer
    Dim lErr As Long, sErr As String
    Dim sBuf As String
    
    iFileIn = FreeFile
    On Local Error Resume Next
    Open sFileIn For Input Shared As #iFileIn
    lErr = Err.Number
    If lErr = 0 Then
        iFileOut = FreeFile
        If bInit Then
            Open sFileOut For Output As #iFileOut
        Else
            Open sFileOut For Append As #iFileOut
        End If
        lErr = Err.Number
    End If
    sErr = Err.Description
    On Local Error GoTo 0
    
    If lErr = 0 Then
        On Local Error Resume Next
        Do While Not EOF(iFileIn) And lErr = 0
            Line Input #iFileIn, sBuf
            Print #iFileOut, sBuf
            lErr = Err.Number
            sErr = Err.Description
        Loop
        On Local Error GoTo 0
    End If
    Close
    
    If lErr <> 0 Then MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & " - " & sErr, vbInformation Or vbOKOnly, App.EXEName
    
End Sub
