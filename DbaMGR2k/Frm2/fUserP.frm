VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fUserP 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "-"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5805
   ClipControls    =   0   'False
   Icon            =   "fUserP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList iList 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fUserP.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fUserP.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fUserP.frx":05BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fUserP.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fUserP.frx":0878
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FitGrid"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "fUserP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bLoading As Boolean
Private miRow As Integer
Private sDbName As String
Private WorkColl As Collection
Private mlWidth As Long
Private mlHeight As Long

Public Property Let UsersData(colDB As Collection, ByVal szDBName As String)

    Set WorkColl = New Collection

    Set WorkColl = colDB
    sDbName = szDBName
    Me.Caption = ReplaceMsg(MyLoadResString(k_Search4Orphaned), Array("1%"), Array(szDBName))
    
    LoadFlex
End Property

Private Sub Flex_RowColChange()

    Dim iRowP As Integer
    Dim iCol As Integer
    Static b As Boolean
    
    If Not bLoading Then
        If Not b Then
            b = True
            With Flex
                If .Row <> miRow Or (.RowData(.Row) = 0) Then
                
                    iCol = .Col
                    iRowP = .Row
                    If miRow > .Rows - 1 Then miRow = .Rows - 1
                    .Row = miRow
                    .Col = 0
                    Set Flex.CellPicture = Nothing
                    .RowData(.Row) = 0
                                        
                    .Row = iRowP
                    .Col = 0
                    Set Flex.CellPicture = LoadResPicture(k_ResSelector, vbResBitmap)
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowData(.Row) = 1
                    
                    .Col = iCol
                    miRow = iRowP
                    
                    Tbar.Buttons(keyDelete).Enabled = Not WorkColl.Item(iRowP).Mapped()
                
                End If
                
            End With
            b = False
        End If
    End If

End Sub

Private Sub Form_Activate()
    Static bFirst As Boolean
                
    If Not bFirst Then
        bFirst = True
    
        With fResult
            .WWrapVisible = False
            .tRes.Text = MyLoadResString(k_OrphanedWARNING)
            .Action = act_Null
            .Caption = Me.Caption
        End With
                
        fResult.Show vbModal, Me
        SettaLargFlex Flex, 2, Flex.ColWidth(3)
    End If

End Sub

Private Sub Form_Load()
    mlWidth = Me.Width
    mlHeight = Me.Height

    CenterForm Me
    Tbar.Buttons(keyEnd).ToolTipText = MyLoadResString(k_Cancel)
    Tbar.Buttons(keyFitGrid).ToolTipText = MyLoadResString(k_Fit_Grid)
    Tbar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource)
    
    Tbar.Buttons(keyDelete).ToolTipText = MyLoadResString(k_Drop_Orhaned_User)
    
End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            If Me.Width < mlWidth Then Me.Width = mlWidth
            If Me.Height < mlHeight Then Me.Height = mlHeight
            
            Flex.Width = Me.ScaleWidth - (2 * Flex.Left)
            Flex.Height = Me.ScaleHeight - Flex.Top - k_FlexBevel
            b = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set WorkColl = Nothing
    Set fUserP = Nothing

End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case keyEnd
            Unload Me
        Case keyFitGrid
            Screen.MousePointer = vbHourglass
            LoadFlex
            Screen.MousePointer = vbDefault
        Case keyRefresh
            Regen
        Case keyDelete
            DropOrphaned
            
    End Select
End Sub
Private Sub Regen()
    Dim lErr As Long, sErr As String
    Dim oCol As Collection
    
    Screen.MousePointer = vbHourglass
    lErr = ErrGetUsersMapping(sDbName, oCol, sErr)
    If lErr = 0 Then
        Me.UsersData(oCol) = sDbName
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbDefault
        MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End If
    
End Sub
Private Sub LoadFlex()

    Dim i As Integer, iPic As Integer
    Dim sBuf As String
    Dim bVal As Boolean
    
    With Flex
        bLoading = True
        
        .Redraw = False
        .Clear
        .FixedCols = 0
        .FixedRows = 0
        
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 1) = MyLoadResString(k_Mapped)
        .TextMatrix(0, 2) = MyLoadResString(k_AssociatedLogin)
        .TextMatrix(0, 3) = MyLoadResString(k_UserMap)
        .ColWidth(0) = k_FlexPARAM_cellaSelector
        .ColWidth(1) = 800
        .ColWidth(3) = .Width \ 3
        
        For i = 1 To WorkColl.Count
            sBuf = vbTab & vbTab & WorkColl.Item(i).LogName() & vbTab & WorkColl.Item(i).Name()
            .AddItem sBuf
            .Row = .Rows - 1
            .Col = 1
            bVal = WorkColl.Item(i).Mapped()
            If Not bVal Then
                iPic = k_ResChkDel
            Else
                iPic = k_ResChkON
                .CellBackColor = vbInactiveBorder
            End If
                
            Set Flex.CellPicture = LoadResPicture(iPic, vbResBitmap)
            .CellPictureAlignment = flexAlignCenterCenter
        Next
        
        .AllowBigSelection = False
        .ScrollBars = flexScrollBarBoth
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .FixedRows = 1
        .FixedCols = 1
        .Redraw = True
        
        SettaLargFlex Flex, 2, .ColWidth(3)
        miRow = 0
        .TopRow = 1
        .Row = 1
        miRow = 1
        bLoading = False
        Flex_RowColChange
    End With
    
End Sub
Private Sub DropOrphaned()

    Dim lErr As Long, sErr As String
    Dim sSql As String, sUser As String
    
    sUser = Flex.TextMatrix(miRow, 3)
    If MsgBox(ReplaceMsg(MyLoadResString(k_Ask4Drop), Array("1%"), Array(sUser)), vbOKCancel, App.EXEName) = vbOK Then
        sSql = "EXEC sp_dropuser '" & sUser & "'"
        Screen.MousePointer = vbHourglass
        On Local Error Resume Next
        objServer.Databases(sDbName).ExecuteImmediate sSql
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        Screen.MousePointer = vbDefault
        If lErr = 0 Then
            Regen
        Else
            MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
        End If
    End If

End Sub
