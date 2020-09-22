VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fDepend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "used"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   ClipControls    =   0   'False
   Icon            =   "fDepend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2715
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2715
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "fDepend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sObjName As String

Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    CenterForm Me
    cmd(0).Caption = MyLoadResString(k_Cancel)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fDepend = Nothing
End Sub

Public Sub Stuff_Exec(ByVal sVal As String, oCollChild As Collection, oCollParent As Collection)
        
    sObjName = sVal
    Me.Caption = MyLoadResString(k_Dependencies_for) & sObjName
    lbl(0).Caption = MyLoadResString(k_Objects_obj_THAT_Depends_on) & " " & sVal
    lbl(1).Caption = Replace(MyLoadResString(k_Objects_THAT_obj_Depends_on), "%", sVal)
            
    LoadGrid 0, oCollChild
    LoadGrid 1, oCollParent
End Sub

Private Sub LoadGrid(ByVal intFlex As Integer, oColl As Collection)
    
    Dim lRows As Long
    Dim sBuf As String
    Dim i As Integer
    Dim iValPic As Integer
    
    With Flex(intFlex)
        .Redraw = False
        .Cols = 3
        .TextMatrix(0, 0) = MyLoadResString(k_Object_Owner) '"Object (Owner)"
        .TextMatrix(0, 1) = MyLoadResString(k_Object_Owner) '"Object (Owner)"
        .TextMatrix(0, 2) = MyLoadResString(k_Object_Sequence)  '"Sequence"
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeRow(0) = True
    
        lRows = oColl.Count
        For i = 1 To oColl.Count
            
            iValPic = oColl.Item(i).oType()
            If iValPic = SQLDMOObj_SystemTable Or iValPic = SQLDMOObj_UserTable Then
                iValPic = k_ResTB
            ElseIf iValPic = SQLDMOObj_View Then
                iValPic = k_ResView
            ElseIf iValPic = SQLDMOObj_StoredProcedure Then
                iValPic = k_ResSP
            ElseIf iValPic = SQLDMOObj_Default Then
                iValPic = k_ResDef
            ElseIf iValPic = SQLDMOObj_Rule Then
                iValPic = k_ResRule
            ElseIf iValPic = SQLDMOObj_UserDefinedDatatype Then
                iValPic = k_ResUDT
            ElseIf iValPic = SQLDMOObj_Trigger Then
               iValPic = k_ResTRIGGER
            ElseIf iValPic = SQLDMOObj_UserDefinedFunction Then
                iValPic = k_ResFUNCTION
            Else
                iValPic = -1
            End If
            sBuf = oColl.Item(i).oName() & " (" & oColl.Item(i).oOwner & ")" & vbTab & oColl.Item(i).oSequence
            If Len(sBuf) Then
                .AddItem vbTab & sBuf
                If iValPic <> -1 Then
                    .Row = .Rows - 1
                    Set Flex(intFlex).CellPicture = LoadResPicture(iValPic, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
            End If
        Next
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColWidth(0) = 350
        .ColWidth(2) = 850
        'If .Cols = 1 Then .Cols = 2
        If .Rows = 1 Then .AddItem ""
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        
        .FixedRows = 1
        '.FixedCols = 1
        .Redraw = True
        SettaLargFlex Flex(intFlex), 1, 200
        .TopRow = 1
        .Row = 1
    End With
        
End Sub

