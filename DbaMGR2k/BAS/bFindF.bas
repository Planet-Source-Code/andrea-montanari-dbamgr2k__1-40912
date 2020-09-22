Attribute VB_Name = "bFindF"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you many not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MAXDWORD As Long = &HFFFFFFFF
'Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_FLAGS = FILE_ATTRIBUTE_ARCHIVE Or _
                                    FILE_ATTRIBUTE_HIDDEN Or _
                                    FILE_ATTRIBUTE_NORMAL Or _
                                    FILE_ATTRIBUTE_READONLY

Public Const DRIVE_UNKNOWNTYPE As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'custom UDT for searching - add additional members if required
Public Type FILE_PARAMS
   bRecurse As Boolean     'set True to perform a recursive search
   bList As Boolean        'set True to add results to listbox
   bFound As Boolean       'set only with SearchTreeForFile methods
   sFileRoot As String     'search starting point, ie c:\, c:\winnt\
   sFileNameExt As String  'filename/filespec to locate, ie *.dll, notepad.exe
   sResult As String       'path to file. Set only with SearchTreeForFile methods
   nFileCount As Long      'total file count matching filespec. Set in FindXXX only
   nFileSize As Double     'total file size matching filespec. Set in FindXXX only
End Type

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function SearchTreeForFile Lib "imagehlp" _
  (ByVal sFileRoot As String, _
   ByVal InputPathName As String, _
   ByVal OutputPathBuffer As String) As Boolean

Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long

Public Declare Function GetDriveType Lib "kernel32" _
   Alias "GetDriveTypeA" _
  (ByVal nDrive As String) As Long



Public Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

'Public Function SearchPathForFile(FP As FILE_PARAMS) As Boolean
'
'    Dim sResult As String
'
'    'pad a return string and search the passed drive
'    sResult = Space(MAX_PATH)
'
'    'SearchTreeForFile returns True (1) if found,
'    'or False otherwise. If True, sResult holds
'    'the full path.
'    FP.bFound = SearchTreeForFile(FP.sFileRoot, FP.sFileNameExt, sResult)
'
'    'if found, strip the trailing nulls and exit
'    If FP.bFound Then
'        FP.sResult = LCase$(TrimNull(sResult))
'    End If
'
'    SearchPathForFile = FP.bFound
'
'End Function
Public Function SearchSystemForFile(FP As FILE_PARAMS) As Boolean

   Dim nSize As Long
   Dim sBuffer As String
   Dim currDrive As String
   Dim sResult As String
       
  'retrieve the available drives on the system
   sBuffer = Space$(64)
   nSize = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
   
  'nSize returns the size of the drive string
   If nSize Then
   
     'strip off trailing nulls
      sBuffer = Left$(sBuffer, nSize)
     
     'search each fixed disk drive for the file
      Do Until sBuffer = ""

        'strip off one drive item from sBuffer
         FP.sFileRoot = StripItem(sBuffer)

        'just search the local file system
         If GetDriveType(FP.sFileRoot) = DRIVE_FIXED Then
         
           'pad a return string and search the passed drive
            sResult = Space(MAX_PATH)
      
            FP.bFound = SearchTreeForFile(FP.sFileRoot, FP.sFileNameExt, sResult)
            
           'if found, strip the trailing nulls and exit
            If FP.bFound Then
               FP.sResult = LCase$(TrimNull(sResult))
               Exit Do
            End If
         
         End If
      
      Loop
      
   End If
      
   SearchSystemForFile = FP.bFound

End Function

Private Function StripItem(startStrg As String) As String

  'Take a string separated by Chr(0)'s,
  'and split off 1 item, and shorten the
  'string so that the next item is ready
  'for removal.
   Dim pos As Integer
   
   pos = InStr(startStrg, Chr$(0))
   
   If pos Then
      StripItem = Mid(startStrg, 1, pos - 1)
      startStrg = Mid(startStrg, pos + 1, Len(startStrg))
   End If
   
End Function


