VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TableSet 
   Caption         =   "*"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   ClipControls    =   0   'False
   Icon            =   "TableSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5636
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iList 
      Left            =   5520
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TableSet.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TableSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private adoConn As ADODB.Connection
Private adoRS As ADODB.Recordset
Private mlErr As Long
Private mSqlString As String
Private msDbName As String
Private msTbName As String
Private objErr As New cErrMgr
Private bIsClient As Boolean
Private m_lWidth As Long
Private m_lHeigh As Long


Public Property Get ErrReturned() As Boolean
    ErrReturned = mlErr <> 0
End Property
Public Property Let SqlStatement(ByVal sSql As String)
    mSqlString = sSql
End Property
Public Property Let DBName(ByVal sVal As String)
    msDbName = sVal
End Property
Public Property Let TableName(ByVal sVal As String)
    msTbName = sVal
End Property

Private Sub DataGrid_AfterUpdate()

    If adoConn.CursorLocation = adUseClient Then adoRS.Update

End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        If DataGrid.EditActive = False Then DataGrid.EditActive = True
    End If
End Sub

Private Sub Form_Load()

    m_lWidth = Me.Width
    m_lHeigh = Me.Height

    tBar.Buttons(keyRefresh).ToolTipText = MyLoadResString(k_RefreshFromSource)
    CenterForm Me
End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height < m_lHeigh Then Me.Height = m_lHeigh
            SizeControls
            b = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objErr = Nothing
'    If Not adoRS Is Nothing Then
'        If (adoRS.State And adStateOpen) = adStateOpen Then adoRS.Close
 '   End If
    Set adoRS = Nothing
    
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then adoConn.Close
    End If
    Set adoConn = Nothing
    
    Set TableSet = Nothing
End Sub
Private Sub SizeControls()
    
    If bIsClient Then DataGrid.Top = tBar.Height + 10
    DataGrid.Height = Me.ScaleHeight - DataGrid.Top
    DataGrid.Width = Me.ScaleWidth - (2 * DataGrid.Left)
            
End Sub

Public Sub Execute(Optional ByVal sSql As String = "", Optional ByRef b2Time As Boolean = False)
    
    Dim lErr As Long
    Dim sConnString As String
    Dim iLockType As ADODB.LockTypeEnum
    Dim i As Integer
        
    If Not b2Time Then
        iLockType = Val(GetIni("CONNECT", "LOCK_TYPE", App.Path & "\" & App.EXEName & ".INI"))
        If iLockType = 0 Then iLockType = adLockPessimistic
    
        For i = 1 To DataGrid.Columns.Count - 1
            DataGrid.Columns.Remove 0
        Next
        Screen.MousePointer = vbHourglass
    
        Me.Caption = MyLoadResString(k_Data_in) & "[" & msDbName & ".." & msTbName & "]"
    Else
        iLockType = adLockOptimistic
    End If
    
    Set objErr = New cErrMgr
    'sConnString = "Initial Catalog=" & msDbName & ";User Id= " & objServer.Login & ";" & IIf(Len(objServer.Password), " Password= " & objServer.Password & ";", "")
    
        'Trusted
    If CBool(Val(GetIni("CONNECT", "TRUSTED", App.Path & "\" & App.EXEName & ".INI"))) Then
        sConnString = "Initial Catalog=" & msDbName & "; Integrated Security = SSPI;"
    Else
        'Standard
        sConnString = "Initial Catalog=" & msDbName & ";User Id= " & objServer.Login & ";" & IIf(Len(objServer.Password), " Password= " & objServer.Password & ";", "")
    End If
    
    
    If Not b2Time Then
        lErr = ErrMyOpenConn(adoConn, sConnString, objErr)
    Else
        Set adoRS = Nothing
        Set adoConn = Nothing
        lErr = ErrMyOpenConn(adoConn, sConnString, objErr, adUseClient)
    End If
    
    If lErr = 0 Then
        If Len(sSql) <> 0 Then mSqlString = sSql

        Set adoRS = New ADODB.Recordset
        On Local Error Resume Next
        adoRS.Open mSqlString, adoConn, adOpenKeyset, iLockType, adCmdText
        
        lErr = Err.Number
        If lErr <> 0 Then
            Set objErr = New cErrMgr
            objErr.PushError Err.Number, Err.Description, Err.Source, 0
            Call GetConnError(adoConn, objErr)
        End If
        If lErr = 0 Then
            If adoRS.State <> adStateOpen Then
                lErr = Err_Free
                objErr.PushError Err_Free, MyLoadResString(k_RsStateNotOpen), "OpenRecordSet", 0
            End If
        End If
        On Local Error GoTo 0
        
        
        If lErr = 0 Then
            If adoRS.Supports(adBookmark) Then
                On Local Error Resume Next
                Set DataGrid.DataSource = adoRS
                lErr = Err.Number
                If lErr <> 0 Then
                    If objErr Is Nothing Then Set objErr = New cErrMgr
                    objErr.PushError Err.Number, Err.Description, Err.Source, 0
                    Call GetConnError(adoConn, objErr)
                End If
                On Local Error GoTo 0
            Else
                If Not b2Time Then
                    lErr = 0
                    Execute mSqlString, True
                    bIsClient = True
                    tBar.Visible = True
                    DataGrid.ToolTipText = MyLoadResString(k_ClientCursorWarning)
                Else
                    lErr = Err_Free
                    objErr.PushError Err_Free, MyLoadResString(k_Error_Opening_Client_RS), "OpenRecordSet", 0
                End If
            End If
        
        End If
        
        If lErr <> 0 Then Call ExceptionHandler(objErr, "", True)
    End If
    If Not b2Time Then
        If bIsClient Then Me.Caption = Me.Caption & MyLoadResString(k_ClientConnection)
        Screen.MousePointer = vbDefault
    End If
    mlErr = lErr
    
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case keyRefresh
            adoRS.Requery
    End Select
End Sub
