Attribute VB_Name = "CopyMod"
Option Explicit

Public Const MAXDWORD As Long = &HFFFF
Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

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

Public Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
      
Public Declare Function CompareFileTime Lib "kernel32" _
  (lpFileTime1 As FILETIME, _
   lpFileTime2 As FILETIME) As Long

Public Declare Function CopyFile Lib "kernel32" _
   Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, _
   ByVal lpNewFileName As String, _
   ByVal bFailIfExists As Long) As Long
   
Public Declare Function CreateDirectory Lib "kernel32" _
    Alias "CreateDirectoryA" _
   (ByVal lpPathName As String, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long



