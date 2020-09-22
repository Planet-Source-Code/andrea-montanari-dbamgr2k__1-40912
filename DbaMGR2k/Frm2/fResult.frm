VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fResult 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "x"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "..."
      Height          =   240
      Index           =   4
      Left            =   5160
      TabIndex        =   9
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox chkWWrap 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   320
      Width           =   2655
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
      Height          =   360
      Index           =   3
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      Height          =   360
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox tRes 
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fResult.frx":0000
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esci"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Nuovo"
      Height          =   360
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNewI 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Elemento Selezionato"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "fResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bIsObjectOwner As Boolean
Private m_IsTrig4View As Boolean
Private iActionMode As optAction

Private bConfirm As Boolean
Private bDirty As Boolean
Private bLoaded As Boolean
Private m_sDbName As String
Private sStuf As String

Public Property Let IsTrig4View(ByVal bVal As Boolean)
    m_IsTrig4View = bVal
End Property

Public Property Get Confirm() As Boolean
    Confirm = bConfirm
End Property
Private Sub cbo_Click()
    If iActionMode = act_Tri Then
        LoadTrigger cbo.List(cbo.ListIndex)
    End If
End Sub

Private Sub chkWWrap_Click()
    tRes.RightMargin = IIf(chkWWrap.Value, 40000, 0)
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            If iActionMode = act_Response Or iActionMode = act_Computed Or iActionMode = act_Modify_Idx_Sql Then
                Me.Hide
            Else
                Unload Me
            End If
        Case 1  'New
            tRes.SetFocus
            If iActionMode = act_Tri Then
                If Not lblNewI.Visible Then NewTrigger
            ElseIf iActionMode = act_SP Then
                LoadPermission m_sDbName, sStuf, act_SP
            End If
        Case 2  'Save
            tRes.SetFocus
            cmd(4).Visible = False
            If iActionMode = act_Tri Then
                If lblNewI.Visible Then SaveTrigger
            ElseIf iActionMode = act_SP Then
                If bDirty Then SaveStoredProc
            ElseIf iActionMode = act_Response Or iActionMode = act_Computed Or iActionMode = act_Modify_Idx_Sql Then
                bConfirm = True
                Me.Hide
            End If
        Case 3  'delete
            tRes.SetFocus
            cmd(4).Visible = False
            If iActionMode = act_Tri Then
                If lblNewI.Visible Then
                    lblNewI.Visible = False
                    tRes.Text = ""
                    tRes.Locked = True
                Else
                    If cbo.ListCount Then DelTrigger cbo.ListIndex
                End If
            End If
        Case 4  'Extended Properties
            ManageExtendedProp
    End Select

End Sub
Private Sub SaveStoredProc()

    Dim lErr As Long, sErr As String
    
    Dim oSP As SQLDMO.StoredProcedure
    Set oSP = New SQLDMO.StoredProcedure
    
    oSP.Name = sStuf
    oSP.Text = tRes.Text
    On Local Error Resume Next
    objServer.Databases(m_sDbName).StoredProcedures.Add oSP
    lErr = Err.Number
    sErr = Err.Description
    On Local Error GoTo 0
    Set oSP = Nothing
    
    If lErr <> 0 Then
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        cmd(2).Visible = False
        tRes.Locked = True
        bDirty = False
    End If
    
End Sub
Private Sub SaveTrigger()

    Dim lErr As Long, sErr As String
    Dim sOwner As String
    Dim oObj As Object
    
    If m_IsTrig4View Then
        Set oObj = GetViewByName(objServer, m_sDbName, sStuf)
    Else
        Set oObj = GetTbByName(objServer, m_sDbName, sStuf)
    End If
    
    If oObj Is Nothing Then
        lErr = Err_Free
        sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
    End If
        
    If lErr = 0 Then
        sOwner = oObj.Owner
        
        Dim oTr As SQLDMO.Trigger
        Set oTr = New SQLDMO.Trigger
        
        oTr.Name = lblNewI
        oTr.Text = tRes.Text
        'oTr.Owner = sOwner
        On Local Error Resume Next
        'objServer.Databases(sDbName).Tables(sTbName).Triggers.Add oTr
        oObj.Triggers.Add oTr
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        Set oTr = Nothing
    End If
    
    Set oObj = Nothing
    If lErr <> 0 Then
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    Else
        cbo.AddItem lblNewI
        lblNewI.Visible = False
        cbo.ListIndex = cbo.ListCount - 1
    End If
    
End Sub
Private Sub DelTrigger(ByVal iItem As Integer)

    Dim lErr As Long, sErr As String
    Dim oObj As Object
    
    If m_IsTrig4View Then
        Set oObj = GetViewByName(objServer, m_sDbName, sStuf)
    Else
        Set oObj = GetTbByName(objServer, m_sDbName, sStuf)
    End If
        
    If oObj Is Nothing Then
        lErr = Err_Free
        sErr = MyLoadResString(k_RES_Object_Not_Found_simple)
    End If
    
    If lErr = 0 Then
        On Local Error Resume Next
        oObj.Triggers.Remove cbo.Text
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
    End If
    Set oObj = Nothing
    
'    On Local Error Resume Next
'    objServer.Databases(sDbName).Tables(sStuf).Triggers.Remove cbo.Text
'    lErr = Err.Number
'    sErr = Err.Description
'    On Local Error GoTo 0
    
    If lErr = 0 Then
        cbo.RemoveItem iItem
        tRes.Text = ""
        If cbo.ListCount Then cbo.ListIndex = 0
    Else
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbInformation Or vbOKOnly, App.EXEName
    End If
    
End Sub
Private Sub NewTrigger()

    Dim bConfirm As Boolean
    Dim sBuf As String
    Dim sOwner As String
    Dim sTrigg As String
    
    Screen.MousePointer = vbDefault
    fPwd.Action() = act_Tbl
    
    fPwd.Caption = MyLoadResString(k_New) & " Trigger ..."
    fPwd.Stuff() = "Trigger"
    fPwd.Show vbModal, Me
    bConfirm = fPwd.Confirm()
    If bConfirm Then sBuf = fPwd.Stuff()
    Unload fPwd
        
    If bConfirm Then
        cmd(4).Visible = False
        If m_IsTrig4View Then
            sOwner = GetViewOwner(objServer, m_sDbName, sStuf)
        Else
            sOwner = GetTableOwner(objServer, m_sDbName, sStuf)
        End If
    
        lblNewI.Visible = True
        lblNewI = sBuf
        sTrigg = "CREATE TRIGGER [" & sBuf & "]" & vbLf & "ON " & sOwner & "." & sStuf & vbLf & String$(2, vbTab)
        
        If m_IsTrig4View Then
            sTrigg = sTrigg & "INSTEAD OF INSERT / UPDATE / DELETE"
        Else
            sTrigg = sTrigg & "FOR DELETE, INSERT, UPDATE"
        End If
        sTrigg = sTrigg & vbCrLf & "AS" & vbTab & "BEGIN" & vbCrLf & String$(2, vbTab) & String$(2, vbCrLf) & "END"
        
        tRes.Text = sTrigg
        tRes.Locked = False
    End If

End Sub
Private Sub Form_Load()
    lbl.Caption = MyLoadResString(k_Selected_Object)
    cmd(0).Caption = MyLoadResString(k_Cancel)
    cmd(1).Caption = MyLoadResString(k_New)
    cmd(2).Caption = MyLoadResString(k_Save)
    cmd(3).Caption = MyLoadResString(k_Remove)
    cmd(4).ToolTipText = MyLoadResString(k_Extended_Property)
    chkWWrap.Caption = MyLoadResString(k_WordWrap)
    chkWWrap.ToolTipText = MyLoadResString(k_WordWrapHelp)
    
    CenterForm Me
End Sub

Private Sub Form_Resize()
    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            On Error Resume Next
            If Me.Width < 6030 Then Me.Width = 6030
            If Me.Height < 4245 Then Me.Height = 4245
            SizeControls
            b = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fResult = Nothing
End Sub
Private Sub SizeControls()

    Dim i As Integer
    cmd(0).Top = Me.ScaleHeight - (cmd(0).Height + 50)

    For i = 1 To 4
        If i < 4 Then
            cmd(i).Top = cmd(0).Top
            cmd(i).Left = cmd(0).Left + i * (cmd(0).Width + 25)
        
        Else   'Extended Properties
            cmd(i).Top = cmd(0).Top + (cmd(0).Height - cmd(i).Height)
            cmd(i).Left = Me.ScaleWidth - cmd(i).Width - 100
        End If
    Next
        
        
    If chkWWrap.Visible Then
        If cbo.Visible Then
            chkWWrap.Top = cbo.Top + cbo.Height + 15
        Else
            chkWWrap.Top = 15
        End If
    End If
    If chkWWrap.Visible Then
        tRes.Top = chkWWrap.Top + chkWWrap.Height + 15
    Else
        If cbo.Visible Then
            tRes.Top = cbo.Top + cbo.Height + 15
        Else
            tRes.Top = 15
        End If
    End If
    tRes.Height = cmd(0).Top - 50 - tRes.Top
        
'    If iActionMode = act_Null Or iActionMode = act_SP Or iActionMode = act_DETACH Then
'        If chkWWrap.Visible Then
'            If cbo.Visible Then
'                chkWWrap.Top = cbo.Top + cbo.Height + 15
'            Else
'                chkWWrap.Top = 15
'            End If
'        End If
'        If chkWWrap.Visible Then
'            tRes.Top = chkWWrap.Top + chkWWrap.Height + 15
'        Else
'            tRes.Top = 15
'        End If
'        tRes.Height = cmd(0).Top - 50 - tRes.Top
'    Else
'        If chkWWrap.Visible Then
'            If cbo.Visible Then
'                chkWWrap.Top = cbo.Top + cbo.Height + 15
'            Else
'                chkWWrap.Top = 15
'            End If
'        End If
'        If chkWWrap.Visible Then
'            tRes.Top = chkWWrap.Top + chkWWrap.Height + 15
'        Else
'            tRes.Top = cbo.Top + cbo.Height + 25
'        End If
'        tRes.Height = cmd(0).Top - 50 - tRes.Top
'    End If
    tRes.Width = Me.ScaleWidth - (2 * tRes.Left)
        
End Sub
Public Property Let WWrapVisible(ByVal bVal As Boolean)
    chkWWrap.Visible = bVal
End Property
'Public Sub ObjectTextProperty(ByVal sDb As String, sObjName As String, ByVal itype As am_SqlPropType)
'    sDbName = sDb
'    sStuf = sObjName
'    m_iType = itype
'
'    if itype=am_ViewProp
'        Set oObj = GetViewByName(objServer, sDbName, sStuf)
'    Else
'        Set oObj = GetTbByName(objServer, sDbName, sStuf)
'    End If
'
'    m_bIsObjectOwner = IsObjectOwner(sDbName, sStuf, IIf(m_IsTrig4View, am_OwnView, am_OwnTable))
'
'End Sub
Public Property Let Action(ByVal i As optAction)
        
    iActionMode = i
    cmd(4).Visible = False
    Select Case iActionMode
        Case act_Modify_Idx_Sql
            lbl.Visible = False
            lblNewI.Visible = False
            cbo.Visible = False
            cmd(1).Visible = False
            cmd(2).Visible = True
            cmd(3).Visible = False
            tRes.Locked = False
            bDirty = True
        
        Case act_Null, act_Response  'Result
            lbl.Visible = False
            lblNewI.Visible = False
            cbo.Visible = False
            bLoaded = True
            tRes.Locked = True
            If i = act_Response Then
                cmd(2).Caption = "&Ok"
                cmd(2).Visible = True
                cmd(2).Default = True
            End If
        Case act_Computed
            lbl.Visible = False
            lblNewI.Visible = False
            cbo.Visible = False
            bLoaded = True
            cmd(2).Caption = "&Ok"
            cmd(2).Visible = True
        Case act_SP, act_User_Function
            lbl.Visible = False
            lblNewI.Visible = False
            cbo.Visible = False
            cmd(1).Visible = False
            cmd(2).Visible = True
            cmd(3).Visible = False
            tRes.Locked = False
            bDirty = True
        Case act_Tri
            Debug.Print cmd(1).Caption
            cmd(1).Visible = True
            cmd(2).Visible = True
            cmd(3).Visible = True
    End Select
    
End Property

Public Sub Stuff_SP(ByVal sDB As String, ByVal sSP As String)
    m_sDbName = sDB
    sStuf = sSP
End Sub

Public Sub Stuff_Trigger(ByVal sDB As String, ByVal sTbOrView As String)
    
    Dim i As Integer
        
    Dim oObj As Object
    m_sDbName = sDB
    sStuf = sTbOrView
    
    
    'Set oTbl = objServer.Databases(sDbName).Tables(sStuf)
    'Set oTbl = GetTbByName(objServer, sDbName, sStuf)
    If m_IsTrig4View Then
        Set oObj = GetViewByName(objServer, m_sDbName, sStuf)
    Else
        Set oObj = GetTbByName(objServer, m_sDbName, sStuf)
    End If
        
    m_bIsObjectOwner = IsObjectOwner(m_sDbName, sStuf, IIf(m_IsTrig4View, am_OwnView, am_OwnTable))
    
    For i = 1 To oObj.Triggers.Count
        cbo.AddItem oObj.Triggers(i).Name
    Next
    Set oObj = Nothing
    If cbo.ListCount Then
        cbo.ListIndex = 0
    Else
        tRes.Locked = True
    End If
    
End Sub
Private Sub LoadTrigger(ByVal sName As String)

    Dim oTr As SQLDMO.Trigger
    Dim iT As Integer
    
    Dim oObj As Object
    
    cmd(4).Visible = False
    
    If m_IsTrig4View Then
        Set oObj = GetViewByName(objServer, m_sDbName, sStuf)
    Else
        Set oObj = GetTbByName(objServer, m_sDbName, sStuf)
    End If
        
    If Not oObj Is Nothing Then
        For iT = 1 To oObj.Triggers.Count
            If StrComp(oObj.Triggers(iT).Name, sName, vbTextCompare) = 0 Then
                Set oTr = oObj.Triggers(iT)
                tRes.Text = oTr.Text
                Set oTr = Nothing
                tRes.Locked = True
                cmd(4).Visible = True
            End If
            Exit For
        Next
    End If
    Set oObj = Nothing
    
    
'    Set oTr = objServer.Databases(sDbName).Tables(sTbName).Triggers(sName)
'    tRes.Text = oTr.Text
'    Set oTr = Nothing
'    tRes.Locked = True
        
    
End Sub

Private Sub tRes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim P As POINTAPI
    Dim i As Integer
    
    If Button = vbRightButton Then
        Dim M As New cPUpMenu
        M.Add MyLoadResString(k_MnuCut), 1
        M.Add MyLoadResString(k_MnuCopy), 2
        M.Add MyLoadResString(k_MnuPaste), 3, , bDisable:=Len(Clipboard.GetText) = 0 Or tRes.Locked
        M.Add "", 4, bSeparator:=True
        M.Add MyLoadResString(k_MnuSelectAll), 5
        
        Call GetCursorPos(P)
        i = M.Popup(Me.hWnd, P.x, P.y)
        M.Clear
        Set M = Nothing
        
        Select Case i
            Case 1
                Clipboard.SetText tRes.SelText, vbCFText
                tRes.SelText = ""
            Case 2
                Clipboard.SetText tRes.SelText, vbCFText
            Case 3
                tRes.SelRTF = Clipboard.GetText(vbCFText)
            Case 5
                tRes.SelStart = 0
                tRes.SelLength = Len(tRes.Text)
        End Select
    End If

End Sub
Private Sub ManageExtendedProp()
    
    Dim iMode As am_SqlPropType
    Dim iOwner As am_SqlPropTypeOwner
    Dim sTrig As String
    Dim sObjName As String
    
    If iActionMode = act_Tri Then
        tRes.SetFocus
        sTrig = cbo.Text
        sObjName = sStuf
        sObjName = Replace(sStuf, "[", "")
        sObjName = Replace(sObjName, "]", "")
    
        If m_IsTrig4View Then
            iOwner = am_OwnView
            iMode = am_TrigPropVIEW
        Else
            iOwner = am_OwnTable
            iMode = am_TrigPropTB
        End If
        If Not m_bIsObjectOwner Then Exit Sub
    
        LoadExtendedProp iMode, m_sDbName, GetObjectOwner(m_sDbName, sObjName, iOwner), sObjName, sTrig
    End If
End Sub

