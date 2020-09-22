Attribute VB_Name = "shellW"
Option Explicit

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SE_ERR_NOASSOC As Long = 31

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
 
'Public Sub EseguiShell(ByVal cmdLine As String, ByVal blnWait As Boolean)
'   Dim hHandle As Double
'   Dim lngRet As Long
'
'   cmdLine = Trim$(cmdLine)
'   If Len(cmdLine) Then
'      hHandle = Shell(cmdLine, vbMaximizedFocus)
'      lngRet = 1
'      If blnWait Then
''NON FUNZIONA - NON ASPETTA L?ANNULLAMENTO DELL'HANDLE
''
'         While lngRet <> 0
'            Sleep 5000
'            lngRet = OpenProcess(PROCESS_QUERY_INFORMATION, False, hHandle)
'            DoEvents
'            CloseHandle lngRet
'         Wend
'      End If
'   End If
'End Sub
 
Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long, _
                           Optional ByVal blnWait As Boolean = True)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
   
   If blnWait Then Sleep 5000

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub
Public Sub Wait4Delay()
   Sleep 1000
End Sub


