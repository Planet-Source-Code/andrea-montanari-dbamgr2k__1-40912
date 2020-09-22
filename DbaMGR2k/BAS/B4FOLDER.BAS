Attribute VB_Name = "b4folder"
Option Explicit
'Public Const MAX_PATH = 260
Public Enum SpecialShellFolderIDs
' IE3 defs
  CSIDL_DESKTOP = &H0                             ' <desktop>
  CSIDL_PROGRAMS = &H2                         ' Start Menu\Programs
  CSIDL_CONTROLS = &H3                          ' My Computer\Control Panel
  CSIDL_PRINTERS = &H4                           ' My Computer\Printers
  CSIDL_PERSONAL = &H5                          ' My Documents
  CSIDL_FAVORITES = &H6                          ' <user name>\Favorites
  CSIDL_STARTUP = &H7                             ' Start Menu\Programs\Startup
  CSIDL_RECENT = &H8                               ' <user name>\Recent
  CSIDL_SENDTO = &H9                               ' <user name>\SendTo
  CSIDL_BITBUCKET = &HA                          ' <desktop>\Recycle Bin
  CSIDL_STARTMENU = &HB                       ' <user name>\Start Menu
  CSIDL_DESKTOPDIRECTORY = &H10      ' <user name>\Desktop
  CSIDL_DRIVES = &H11                              ' My Computer
  CSIDL_NETWORK = &H12                         ' Network Neighborhood
  CSIDL_NETHOOD = &H13                          ' <user name>\nethood
  CSIDL_FONTS = &H14                                ' windows\fonts
  CSIDL_TEMPLATES = &H15

' IE4 defs
  CSIDL_INTERNET = &H1                           ' Internet Explorer (icon on desktop) (IE4)
  CSIDL_COMMON_STARTMENU = &H16                  ' All Users\Start Menu
  CSIDL_COMMON_PROGRAMS = &H17                    ' All Users\Programs
  CSIDL_COMMON_STARTUP = &H18                        ' All Users\Startup
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19   ' All Users\Desktop
  CSIDL_APPDATA = &H1A                           ' <user name>\Application Data
  CSIDL_PRINTHOOD = &H1B                      ' <user name>\PrintHood
  CSIDL_ALTSTARTUP = &H1D                     ' non localized startup
  CSIDL_COMMON_ALTSTARTUP = &H1E   ' non localized common startup
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22

' NT5 (IE5)
  CSIDL_LOCAL_APPDATA = &H1C               ' <user name>\Local Settings\Applicaiton Data (non roaming) (IE5)
  CSIDL_COMMON_APPDATA = &H23           ' All Users\Application Data
  CSIDL_WINDOWS = &H24                           ' GetWindowsDirectory()
  CSIDL_SYSTEM = &H25                               ' GetSystemDirectory()
  CSIDL_PROGRAM_FILES = &H26                ' C:\Program Files
  CSIDL_MYPICTURES = &H27                       ' C:\Program Files\My Pictures
  CSIDL_PROFILE = &H28                               ' USERPROFILE
  CSIDL_SYSTEMX86 = &H29                         ' x86 system directory on RISC
  CSIDL_PROGRAM_FILESX86 = &H2A                      ' x86 C:\Program Files on RISC
  CSIDL_PROGRAM_FILES_COMMON = &H2B          ' C:\Program Files\Common
  CSIDL_PROGRAM_FILES_COMMONX86 = &H2C    ' x86 Program Files\Common on RISC
  CSIDL_COMMON_TEMPLATES = &H2D                   ' All Users\Templates
  CSIDL_COMMON_DOCUMENTS = &H2E                  ' All Users\Documents
'
  CSIDL_FLAG_CREATE = &H8000&      ' combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
End Enum

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_PIDL = &H8
Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Type SHITEMID
    cb As Long     'Size of the ID (including cb itself)
    abID As Byte   'The item ID (variable length)
End Type
Type ITEMIDLIST
    mkid As SHITEMID
End Type


Declare Function SHGetPathFromIDList _
     Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
     (ByVal pidl As Long, ByVal pszPath As String) As Long
    
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, _
                              ByVal nFolder As SpecialShellFolderIDs, _
                              pidl As Long) As Long

Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

Public Const NOERROR = 0


Declare Function SHBrowseForFolder _
     Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Type BROWSEINFO
       'Handle of the owner window for the dialog box.
        hOwner As Long
        'Pointer to an item identifier list (an
       'ITEMIDLIST structure) specifying the location of
       'the "root" folder to browse from. Only the
       'specified folder and its subfolders appear in the dialog box.
       'This member can be NULL, and in that case, the namespace
       'root (the desktop folder) is used.
        pidlRoot As Long
       'Pointer to a buffer that receives the display
       'name of the folder selected by the user. The
       'size of this buffer is assumed to be MAX_PATH bytes.
        pszDisplayName As String
       'Pointer to a null-terminated string that is
       'displayed above the tree view control in the
       'dialog box.This string can be used to specify
       'instructions to the user.
       lpszTitle As String
       'Value specifying the types of folders to be
       'listed in the dialog box as well as other options.
       'This member can include zero or more of
       'the following values below.
       ulFlags As Long
       'Address an application-defined function that the
       'dialog box calls when events occur. For more information,
       'see the description of the BrowseCallbackProc function.
       'This member can be NULL. (note: VB4 does not support
       'callbacks, therefore this member is ignored.)
       lpfn As Long
       'Application-defined value that the dialog box
       'passes to the callback function (if one is specified).
        lParam As Long
        'Variable that receives the image associated with
       'the selected folder. The image is specified as an
       'index to the system image list.
       iImage As Long
End Type

'Only returns file system directories. If the
   'user selects folders that are not part of the
   'file system, the OK button is grayed.
    Public Const BIF_RETURNONLYFSDIRS = &H1
   'Does not include network folders below the
   'domain level in the tree view control.    'For starting the Find Computer.
    Public Const BIF_DONTGOBELOWDOMAIN = &H2
   'Includes a status area in the dialog box.
   'The callback function can set the status
   'text by sending messages to the dialog box.
    Public Const BIF_STATUSTEXT = &H4
   'Only returns file system ancestors. If the
   'user selects anything other than a file
   'system ancestor, the OK button is grayed.
    Public Const BIF_RETURNFSANCESTORS = &H8
   'Only returns computers. If the user selects
   'anything other than a computer, the OK    'button is grayed.
    Public Const BIF_BROWSEFORCOMPUTER = &H1000


Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function
Public Function GetPathFromPIDL(pidl As Long) As String
  Dim sPath As String * MAX_PATH   ' 260
  If pidl Then
    If SHGetPathFromIDList(pidl, sPath) Then
      GetPathFromPIDL = GetStrFromBufferA(sPath)
    End If
  End If
End Function

Public Function GetFileDisplayName(ByVal pszPath As Variant) As String
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_DISPLAYNAME, sfi) Then
    GetFileDisplayName = GetStrFromBufferA(sfi.szDisplayName)
  End If
End Function

Public Function GetFileInfo(ByVal pszPath As Variant, _
                                        uFlags As Long, _
                                        sfi As SHFILEINFO) As Long
  If (VarType(pszPath) = vbString) Then
    ' Must be an explicit path (not a display name).
    GetFileInfo = SHGetFileInfo(CStr(pszPath), 0, sfi, Len(sfi), uFlags)
  ElseIf (VarType(pszPath) = vbLong) Then ' assume good pidl
    GetFileInfo = SHGetFileInfo(CLng(pszPath), 0, sfi, Len(sfi), uFlags Or SHGFI_PIDL)
  End If
End Function

Public Function GetPIDLFromFolderID(hOwner As Long, nFolder As SpecialShellFolderIDs) As Long
  Dim pidl As Long
  If (SHGetSpecialFolderLocation(hOwner, nFolder, pidl) = NOERROR) Then
    GetPIDLFromFolderID = pidl
  End If
End Function



