Attribute VB_Name = "shFolder"
Option Explicit
Public Const CSIDL_DESKTOP = &H0                 '{desktop}
Public Const CSIDL_INTERNET = &H1                'Internet Explorer (icon on desktop)
Public Const CSIDL_PROGRAMS = &H2                'Start Menu\Programs
Public Const CSIDL_CONTROLS = &H3                'My Computer\Control Panel
Public Const CSIDL_PRINTERS = &H4                'My Computer\Printers
'Public Const CSIDL_PERSONAL = &H5                'My Documents
Public Const CSIDL_FAVORITES = &H6               '{user name}\Favorites
Public Const CSIDL_STARTUP = &H7                 'Start Menu\Programs\Startup
Public Const CSIDL_RECENT = &H8                  '{user name}\Recent
Public Const CSIDL_SENDTO = &H9                  '{user name}\SendTo
Public Const CSIDL_BITBUCKET = &HA               '{desktop}\Recycle Bin
Public Const CSIDL_STARTMENU = &HB               '{user name}\Start Menu
Public Const CSIDL_DESKTOPDIRECTORY = &H10       '{user name}\Desktop
Public Const CSIDL_DRIVES = &H11                 'My Computer
'Public Const CSIDL_NETWORK = &H12                'Network Neighborhood
Public Const CSIDL_NETHOOD = &H13                '{user name}\nethood
Public Const CSIDL_FONTS = &H14                  'windows\fonts
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16       'All Users\Start Menu
Public Const CSIDL_COMMON_PROGRAMS = &H17        'All Users\Programs
Public Const CSIDL_COMMON_STARTUP = &H18         'All Users\Startup
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
Public Const CSIDL_APPDATA = &H1A                '{user name}\Application Data
Public Const CSIDL_PRINTHOOD = &H1B              '{user name}\PrintHood
Public Const CSIDL_LOCAL_APPDATA = &H1C          '{user name}\Local Settings\Application Data (non roaming)
Public Const CSIDL_ALTSTARTUP = &H1D             'non localized startup
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E      'non localized common startup
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23          'All Users\Application Data
Public Const CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()
Public Const CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
Public Const CSIDL_PROGRAM_FILES = &H26           'C:\Program Files
Public Const CSIDL_MYPICTURES = &H27              'C:\Program Files\My Pictures
Public Const CSIDL_PROFILE = &H28                 'USERPROFILE
Public Const CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
Public Const CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B    'C:\Program Files\Common
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
Public Const CSIDL_COMMON_TEMPLATES = &H2D        'All Users\Templates
Public Const CSIDL_COMMON_DOCUMENTS = &H2E        'All Users\Documents
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F       'All Users\Start Menu\Programs\Administrative Tools
Public Const CSIDL_ADMINTOOLS = &H30              '{user name}\Start Menu\Programs\Administrative Tools

Public Const CSIDL_FLAG_CREATE = &H8000           'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000      'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_MASK = &HFF00             'mask for all possible flag values
Public Const SHGFP_TYPE_CURRENT = &H0             'current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = &H1

Public Const MAX_PATH = 260
Public Const S_OK = 0

'Converts an item identifier list to a file system path.
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
   (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


Public Function GetSpecialFolderLocation(ByVal hWnd As Long, CSIDL As Long) As String
    Dim sPath As String
    Dim pidl As Long
    'fill the idl structure with the specified folder item
    If CSIDL = 0 Then CSIDL = CSIDL_PERSONAL
    If SHGetSpecialFolderLocation(hWnd, CSIDL, pidl) = S_OK Then
        'if the pidl is returned, initialize
        'and get the path from the id list
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
        'free the pidl and return the path
            Call CoTaskMemFree(pidl)
            GetSpecialFolderLocation = Left(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
    
End Function




