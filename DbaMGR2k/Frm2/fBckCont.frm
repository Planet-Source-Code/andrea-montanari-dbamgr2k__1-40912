VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fBckCont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BackUp Media Content"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "ok"
      Height          =   360
      Index           =   1
      Left            =   7080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   800
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1402
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1755
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "fBckCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i_intSelected As Integer
Private i_blnRestore As Boolean
Private i_blnReturn As Boolean
Private i_lErrReadingDevice As Long, i_sErrReadingDevice As String

Public Property Get ErrReadingDevice(ByRef sErr As String) As Long
    sErr = i_sErrReadingDevice
    ErrReadingDevice = i_lErrReadingDevice
End Property
Public Sub LoadFromDisk(oRest As SQLDMO.Restore2, ByVal intSet As Integer)

    Dim lErr As Long, sErr As String

    If Not oRest Is Nothing Then
        i_blnRestore = True
    
        Dim oQry As SQLDMO.QueryResults
        
        On Local Error Resume Next
        Set oQry = oRest.ReadBackupHeader(objServer)
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        
        
        i_intSelected = intSet
        
        If lErr = 0 Then
            If Not oQry Is Nothing Then LoadGrid oQry, 0, intSet
            Flex(1).Visible = False
            Flex(0).Height = 2 * Flex(0).Height
            cmd(0).Top = Flex(0).Top + Flex(0).Height + 100
            cmd(1).Top = cmd(0).Top
            If i_intSelected > (Flex(0).Rows - 1) Then i_intSelected = 0
            cmd(1).Visible = i_intSelected >= 0
            cmd(1).Enabled = (i_intSelected > 0) And (i_intSelected <= Flex(0).Rows - 1)
            Me.Height = cmd(0).Top + 2 * cmd(0).Height
        Else
            MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
        End If
    End If
    Set oQry = Nothing

End Sub
Public Sub LoadData(ByVal sVal As String)
        
    Dim oBac As SQLDMO.BackupDevice
    
    On Local Error Resume Next
    Set oBac = objServer.BackupDevices(sVal)
    Debug.Print Err.Description
    On Local Error GoTo 0
    
    If Not oBac Is Nothing Then
        Dim oQry As SQLDMO.QueryResults
        On Local Error Resume Next
        Set oQry = oBac.ReadMediaHeader
        i_lErrReadingDevice = Err.Number
        i_sErrReadingDevice = Err.Description
        On Local Error GoTo 0
        If Not oQry Is Nothing Then LoadGrid oQry, 0
        
        Set oQry = Nothing
        
        If i_lErrReadingDevice = 0 Then
            On Local Error Resume Next
            Set oQry = oBac.ReadBackupHeader
            i_lErrReadingDevice = Err.Number
            i_sErrReadingDevice = Err.Description
            On Local Error GoTo 0
                            
            If Not oQry Is Nothing Then LoadGrid oQry, 1
        End If
    End If
    Set oQry = Nothing
    Set oBac = Nothing
    
End Sub
Private Sub LoadGrid(oQry As SQLDMO.QueryResults, ByVal Index As Integer, Optional ByVal intSet As Integer = -1)

    Dim i As Integer, iCols As Integer, lRows As Long, l As Long
    Dim sBuf As String
    Dim iStartCol  As Integer, iThisCol As Integer

    If oQry.ResultSets <> 0 Then
   
        With Flex(Index)
            .Redraw = False
            .ColWidth(0) = 300

            oQry.CurrentResultSet = 1
            iCols = oQry.Columns + 1 + IIf(intSet >= 0, 1, 0)
            lRows = oQry.Rows
            .Cols = iCols
            
            If intSet >= 0 Then
                Flex(Index).TextMatrix(0, 1) = "Restore"
                iStartCol = 2
                iThisCol = 1
            Else
                iStartCol = 1
            End If
            
                
            For i = iStartCol To iCols - 1
                Debug.Print oQry.ColumnName(i - iThisCol)
                Flex(Index).TextMatrix(0, i) = oQry.ColumnName(i - iThisCol)
                .ColWidth(i) = Me.TextWidth(Flex(Index).TextMatrix(0, i)) + 200
            Next
                
                'sBuf = Qry.GetRangeString(2, rowdelim:=vbLf, coldelim:=vbTab)

            On Local Error Resume Next
            For l = 1 To lRows
                                  
                For i = iStartCol To iCols - 1
                    If i > 1 Then sBuf = sBuf & vbTab
                    Select Case oQry.ColumnType(i - iThisCol)
                        Case SQLDMO_DTypeUnknown
                            sBuf = sBuf & MyLoadResString(k_Unknown)
                        Case SQLDMO_DTypeBinary
                            sBuf = sBuf & "..Binary"
                        Case SQLDMO_DTypeImage
                            sBuf = sBuf & "..BLOB"
                        Case SQLDMO_DTypeText, SQLDMO_DTypeNText
                            sBuf = sBuf & "..MEMO"
                        Case SQLDMO_DTypeGUID
                            sBuf = sBuf & "..Guid"
                        Case SQLDMO_DTypeUChar, SQLDMO_DTypeChar, SQLDMO_DTypeUVarchar, SQLDMO_DTypeVarchar
                            sBuf = sBuf & oQry.GetColumnString(l, i - iThisCol)
                        Case SQLDMO_DTypeMoney
                            sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i - iThisCol), "#,0.0000")
                            .ColAlignment(i) = flexAlignRightCenter
                        Case SQLDMO_DTypeFloat4, SQLDMO_DTypeMoney
                            sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i - iThisCol), "#,0.0000")
                            .ColAlignment(i) = flexAlignRightCenter
                        Case SQLDMO_DTypeFloat8
                            sBuf = sBuf & Format$(oQry.GetColumnDouble(l, i - iThisCol), "#,0.00000000")
                            .ColAlignment(i) = flexAlignRightCenter
                        Case SQLDMO_DTypeDateTime, SQLDMO_DTypeDateTime4
                            sBuf = sBuf & Format$(oQry.GetColumnDate(l, i - iThisCol), "General Date")
                            .ColAlignment(i) = flexAlignLeftCenter
                        Case SQLDMO_DTypeInt1, SQLDMO_DTypeInt2, SQLDMO_DTypeInt4, 2
                            If i = 3 + iThisCol Then
                                Select Case oQry.GetColumnLong(l, i - iThisCol)
                                    Case 1
                                        sBuf = sBuf & "Database"
                                    Case 2
                                        sBuf = sBuf & "Transaction Log"
                                    Case 4
                                        sBuf = sBuf & "File"
                                    Case 5
                                        sBuf = sBuf & "Differential Database"
                                    Case 6
                                        sBuf = sBuf & "Differential File Group"
                                End Select
                            Else
                                sBuf = sBuf & Format$(oQry.GetColumnLong(l, i - iThisCol), "#,0")
                            End If
                        Case SQLDMO_DTypeBit
                            sBuf = sBuf & Format$(oQry.GetColumnBool(l, i - iThisCol))
                    End Select
                Next
                .AddItem vbTab & sBuf
                sBuf = ""
                    
                If intSet >= 0 Then
                    .Row = .Rows - 1
                    .Col = 1
                    Set Flex(Index).CellPicture = LoadResPicture(IIf(intSet = l, k_ResChkON, k_ResChkOFF), vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                End If

            Next
            
            On Local Error GoTo 0
            If .Rows = 1 Then .AddItem ""
            .Row = 1
            .ScrollBars = flexScrollBarBoth
            .FixedRows = 1
            .FixedCols = 1
            
            .Redraw = True
            .TopRow = 1
        End With
    End If

End Sub
Private Sub cmd_Click(Index As Integer)

    If cmd(Index).Enabled Then
        If Index = 0 Then
            Unload Me
        Else
            i_blnReturn = i_blnRestore
            Me.Hide
        End If
    End If
End Sub


Private Sub Form_Load()
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_Apply)
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fBckCont = Nothing
End Sub
Private Sub FlexEdit()
    
    Dim bEdit As Boolean
    Dim iElem As Integer
    Dim iRow As Integer, i As Integer
    
    bEdit = False
    iElem = Flex(0).Col
    If iElem = 1 Then bEdit = True
            
    If bEdit Then
        i_intSelected = 0
        With Flex(0)
            iRow = Flex(0).Row
            For i = 1 To .Rows - 1
                .Col = 1
                .Row = i
                If iRow = i Then i_intSelected = Val(.TextMatrix(i, 7))
                Set Flex(0).CellPicture = LoadResPicture(IIf(i = iRow, k_ResChkON, k_ResChkOFF), vbResBitmap)
                .CellPictureAlignment = flexAlignCenterCenter
            Next
            cmd(1).Enabled = (i_intSelected > 0) And (i_intSelected <= .Rows - 1)
            .Row = iRow
        End With
    End If
    
End Sub

Public Property Get SelectedRestoreFile() As Integer
    SelectedRestoreFile = i_intSelected
End Property
Private Sub Flex_DblClick(Index As Integer)

    If Index = 0 And i_blnRestore Then
        If Flex(Index).MouseCol > 0 And Flex(Index).MouseRow > 0 Then
            FlexEdit
        End If
    End If

End Sub

Private Sub Flex_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 0 And i_blnRestore Then
        If KeyAscii = vbKeySpace Then FlexEdit
    End If
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = Not i_blnReturn
End Property
