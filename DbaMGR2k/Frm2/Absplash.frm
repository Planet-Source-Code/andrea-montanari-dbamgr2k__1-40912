VERSION 5.00
Begin VB.Form AbSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "application title"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "Absplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   480
      Top             =   480
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "User Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "user information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "user name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Line linDivide 
      Index           =   1
      X1              =   120
      X2              =   6240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblFileDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "file description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label lblTrademark 
      BackStyle       =   0  'Transparent
      Caption         =   "trademark information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "copyright information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Line linDivide 
      Index           =   0
      X1              =   120
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "version information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "application title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image imgIcon 
      Height          =   1215
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "AbSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'      Copyright © 1997 Microsoft Corporation.  All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------

Option Explicit

' API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

' API Constants
Private Const GWL_STYLE         As Long = (-16)
Private Const WS_CAPTION        As Long = &HC00000
Private Const WS_CAPTION_NOT    As Long = &HFFFFFFFF - WS_CAPTION

Private Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGKEYSYSINFO    As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Private Const gREGVALSYSINFOLOC As String = "MSINFO"
Private Const gREGVALSYSINFO    As String = "PATH"

' NT location of user name and company
Private Const gNTREGKEYINFO     As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Private Const gNTREGVALUSER     As String = "RegisteredOwner"
Private Const gNTREGVALCOMPANY  As String = "RegisteredOrganization"

' Win95 locataion of user name and company
Private Const g95REGKEYINFO     As String = "Software\Microsoft\MS Setup (ACME)\User Info"
Private Const g95REGVALUSER     As String = "DefName"
Private Const g95REGVALCOMPANY  As String = "DefCompany"

' Change these to what you want the default name and user info to be
Private Const DEFAULT_USER_NAME As String = "USER INFORMATION NOT AVAILABLE"
Private Const DEFAULT_USER_INFO As String = vbNullString

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

' Information for warning information at bottom of form
'Private Const gWarningInfo      As String = "generic warning information"
Private gWarningInfo             As String

Private mBoxHeight              As Integer
Private mStyle                  As StyleType
Private mTitleBarHidden         As Boolean

' Type declarations
Private Type StyleType
    OldStyle As Long
    NewStyle As Long
End Type 'StyleType

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    
    'gWarningInfo = "generic warning information"

'   Fill in all of the information that comes from the App object
    With App
        Caption = "Project Information about " & App.Title
        lblTitle.Caption = .Title
        
        If .CompanyName <> "" Then
            lblCompanyName.Caption = "A product of " & .CompanyName & " " & k_EMail
        Else
            lblCompanyName.Caption = ""
        End If
        
        lblVersion.Caption = "Version " & .Major & "." & .Minor & "." & _
                             .Revision '& " (32-bit)"
        lblCopyright.Caption = "CopyRights :" & .LegalCopyright
        lblTrademark.Caption = "Trademarks :" & .LegalTrademarks
        lblFileDescription.Caption = "Descriptions :" & App.FileDescription
        lblUserName.Caption = GetUserName
        lblUserInfo.Caption = GetUserCompany
        
                
    End With 'App
    
'   Get "default" height of About Box
    mBoxHeight = Height
    
    Dim oLang As cLang
    Set oLang = New cLang
    oLang.WriteRes2Ini True
    oLang.WriteRes2Ini False
    Set oLang = Nothing
    
'    Dim oLang As cLang
'    Set oLang = New cLang
'    oLang.WriteRes2Ini True
'    oLang.WriteRes2Ini False
'    Set oLang = Nothing
    
    
End Sub

Private Sub cmdOK_Click()
    'Hide ' If you want to unload the form, change this to Unload Me
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Public Sub About(frmParent As Form, Optional lUserName As String, _
                 Optional lUserInfo As String)
    imgIcon.Picture = frmParent.Icon
    cmdOK.Enabled = True
    cmdSysInfo.Enabled = True
    Dim sBuf As String
    Dim bRet As Boolean

    
'   Modify the form style to show the title bar
    ShowTitleBar
    
'   A resize event is needed in order to apply the changes to the form style.  Setting
'   the height to 0 should do it.
    If Height = mBoxHeight Then
        Height = 0
    End If
    
'   Set height of About Box to "default" height
    Height = mBoxHeight
    
    lblContact(0).Caption = k_EMail
    lblContact(0).ToolTipText = MyLoadResString(k_RES_e_mail)
    lblContact(1).Caption = k_URLwww
    lblContact(1).ToolTipText = MyLoadResString(k_RES_www)
    
    Show vbModal, frmParent
    
End Sub

Public Sub SplashOn(frmParent As Form, Optional MinDisplay As Long)

    If Not Visible Then
        Dim lHeight As Integer
        
        imgIcon.Picture = frmParent.Icon
        cmdOK.Enabled = False
        cmdSysInfo.Enabled = False
    
'       If a delay is specified, set up the Timer
        If MinDisplay > 0 Then
            Timer1.Interval = MinDisplay
            Timer1.Enabled = True
        End If
        
        
'       Modify the form style to hide the title bar
        HideTitleBar
        
'       Need to cause a form resize in order to get updated ScaleHeight value
        lHeight = Height
        Height = 0
        Height = lHeight
        
'       Set height to hide the "About Box Only" information
        Height = linDivide(1).Y1 + (Height - ScaleHeight)
        
'       Show the form
        Show vbModeless, frmParent

'       For some reason, need a Refresh to make sure Splash Screen gets painted
        Refresh
    End If
End Sub

Public Sub SplashOff()
    If Visible Then
'       Wait until any minimum display time elapses
        Do While Timer1.Enabled
            DoEvents
        Loop
        
        'Hide ' If you want to unload the form, change this to Unload Me
        Unload Me

'       Modify the form style to show the title bar
        ShowTitleBar
        
'       Set height of About Box to "default" height
        Height = mBoxHeight
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AbSplash = Nothing
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub HideTitleBar()
'   Change the style of the form to not show a title bar
    If mTitleBarHidden Then Exit Sub
    
    mTitleBarHidden = True
    
    With mStyle
        .OldStyle = GetWindowLong(hWnd, GWL_STYLE)
        .NewStyle = .OldStyle And WS_CAPTION_NOT
        SetWindowLong hWnd, GWL_STYLE, .NewStyle
    End With 'mStyle
End Sub

Private Sub ShowTitleBar()
'   Change the style of the form to show a title bar
    If Not mTitleBarHidden Then Exit Sub
    mTitleBarHidden = False
    SetWindowLong hWnd, GWL_STYLE, mStyle.OldStyle
End Sub

Private Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existence Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox MyLoadResString(k_SYS_INFO_UNAVAILABE), vbOKOnly
    
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Temporary Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occurred...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function GetUserName() As String
    Dim KeyVal As String
            
'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   None of the above
    Else
        GetUserName = DEFAULT_USER_NAME
    End If
End Function

Private Function GetUserCompany() As String
    Dim KeyVal As String
    
'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   None of the above
    Else
        GetUserCompany = DEFAULT_USER_INFO
    End If
End Function
Private Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub


Private Sub lblContact_Click(Index As Integer)

    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant

    If Index = 0 Then
        sFile = "mailto:" & lblContact(Index).Caption
    Else
        sFile = lblContact(Index).Caption
    End If
    
    sTopic = "Open"
    sParams = 0&
    sDirectory = 0&
    Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    
End Sub


