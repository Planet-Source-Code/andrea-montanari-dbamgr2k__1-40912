VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private sDlgTitle   As String
Private sDef_Dir    As String
Private lErrCode    As Long
Private sErrMsg     As String
Private lhWnd       As Long
Private sf_Name     As String
Private sFilter     As String
'Private sPattern    As String
Private sFullPath   As String
Private sPathRoot   As String

Const ErrInternal = 51
Const ErrPropInvalid = 380

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 128
'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not 'a standard Win32 type.
'Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
                    Or OFN_LONGNAMES _
                    Or OFN_FILEMUSTEXIST _
                    Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
                    Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST _
                    Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_READONLY

Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
                        Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
                        
'&H80000 Or &H4 Or OFN_OVERWRITEPROMPT

'Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

'Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
                    Or OFN_LONGNAMES _
                    Or OFN_OVERWRITEPROMPT _
                    Or OFN_HIDEREADONLY

'Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Private Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nCustFilterSize   As Long
  nFilterIndex      As Long
  sFile             As String
  nFileSize         As Long
  sFileTitle        As String
  nTitleSize        As Long
  sInitDir          As String
  sDlgTitle         As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExt          As Integer
  sDefFileExt       As String
  nCustDataSize     As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Private OFN As OPENFILENAME
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Property Let formHwnd(ByVal vData As String)
    lhWnd = vData
End Property
'Public Property Let Pattern(ByVal vData As String)
'    sPattern = vData
'End Property
Public Property Let Filter(ByVal vData As String)
    sFilter = vData
End Property
Public Property Get FullPath() As String
    FullPath = sFullPath
End Property
Public Property Get PathRoot() As String
    PathRoot = sPathRoot
End Property

Public Property Get ErrMsg() As String
    ErrMsg = sErrMsg
End Property
Public Property Let Title(ByVal vData As String)
    sDlgTitle = vData
End Property

Public Property Let f_Name(ByVal vData As String)
    sf_Name = vData
End Property
Public Property Get f_Name() As String
    f_Name = sf_Name
End Property
Public Property Get ErrCode() As Long
    ErrCode = lErrCode
End Property
Public Property Let Def_Dir(ByVal vData As String)
    sDef_Dir = vData
End Property
Public Sub Do_Save()
  Dim sp As Long
  Dim LongName As String
  Dim ShortName As String
  Dim ShortSize As Long
  Dim sBuf As String

 'to keep lines short, I've abbreviated
 'a Null$ to n, and the filter$ to f.
  Dim n As String, n2 As String, f As String
  n = Chr$(0)
  n2 = n & n

 '------------------------------------------------
 'INITIALIZATION
 '------------------------------------------------
 'fill in the size of the OFN structure
  OFN.nStructSize = Len(OFN)

 'assign the owner of the dialog; this can be null if no owner.
  OFN.hwndOwner = lhWnd

 '------------------------------------------------
 'FILTERS
 '------------------------------------------------
 'There are 2 methods of setting filters (patterns) for
 'use in the dropdown combo of the dialog.

 'The first, using OFN.sFilter, fills the combo with the
 'specifed filters, and works as the VB common dialog does.
 'These must be in the "Friendly Name"-null$-Extention format,
 'terminating with 2 null strings.
'      f = "Visual Basic Forms" & n & "*.frm" & n
'  f = f & "Visual Basic Modules" & n & "*.bas" & n
'  f = f & "Visual Basic Projects" & n & "*.vbp" & n
'  f = f & "Text Files" & n & "*.txt" & n
'  f = f & "All Files" & n & "*.*" & n2
  
    f = sFilter
    f = AddNullChar(f)
  
  OFN.sFilter = f '& n2

 'The second method, uses sCustomFilter and nCustFilterSize
 'to pass the filters to use and the size of the filter string.
 'The operating system copies the strings to the buffer when
 'the user closes the dialog box. The system uses the strings
 'to initialize the user-defined file filter the next time the
 'dialog box is created. If this parameter is NULL, the dialog
 'box lists but does not save user-defined filter strings.

 'To see the difference, comment out the OFN.sFilter = f above,
 'and uncomment the 2 lines below.
 'OFN.sCustomFilter = f
 'OFN.nCustFilterSize = Len(OFN.sCustomFilter)

 'nFilterIndex specifies an index into the buffer pointed to
 'by sFilter. The system uses the index value to obtain a pair
 'of strings to use as the initial filter description and
 'filter pattern for the dialog box. The first pair of strings
 'has an index value of 1. When the user closes the dialog box,
 'the system copies the index of the selected filter strings
 'into this location.
  
  'OFN.nFilterIndex = 2  '"Visual Basic Modules"

 '------------------------------------------------
 'FILENAME
 '------------------------------------------------
 'sFile points to a buffer that contains a filename used to
 'initialize the File Name edit control. The first character
 'of this buffer must be NULL if initialization is not necessary.
 'When the GetOpenFileName or GetSaveFileName function returns,
 'this buffer contains the drive designator, path, filename, and
 'extension of the selected file.

 'perform no filename initialization (Filename textbox is blank)
 'and initialize the sFile buffer for the return value
 ' OFN.sFile = Chr$(0)
 ' OFN.sFile =  Space$(1024)
 'OR
 'pass a default filename and initialize for return value
 sBuf = AddNullChar(sf_Name & Space$(1024))
  OFN.sFile = sBuf 'sf_Name & Space$(1024) & n2
  OFN.nFileSize = Len(OFN.sFile)

 'default extention applied to a selected file if it has no extention
  'OFN.sDefFileExt = "*.bas" & n2

 'sFileTitle points to a buffer that receives the title of the
 'selected file. The application should use this string
 'to display the file title. If this member is NULL, the
 'function does not copy the file title.
  OFN.sFileTitle = Space$(512) & n2
  OFN.nTitleSize = Len(OFN.sFileTitle) & n2

 'sInitDir is the string that specifies the initial file directory.
 'If this member is NULL, the system uses the current directory as
 'the initial directory.
 sBuf = AddNullChar(sDef_Dir)
  OFN.sInitDir = sBuf 'sDef_Dir & n2
  

 '------------------------------------------------
 'MISC
 '------------------------------------------------
 'sDlgTitle is the title to display in the dialog.  If null,
 'the default title for the dialog is used.
 sBuf = AddNullChar(sDlgTitle)
  OFN.sDlgTitle = sBuf 'sDlgTitle & n2

 'flags are the actions and options for the dialog.
  OFN.Flags = OFS_FILE_SAVE_FLAGS 'OFS_FILE_OPEN_FLAGS Or OFN_ALLOWMULTISELECT

 'Show the File Open Dialog
  If GetSaveFileName(OFN) Then
        Debug.Print OFN.sFile
        sFullPath = OFN.sFile
        If Len(sFullPath) Then
            If InStr(sFullPath, vbNullChar) Then sFullPath = Left$(sFullPath, InStr(sFullPath, vbNullChar) - 1)
        End If
        
        
        Debug.Print Left$(OFN.sFile, OFN.nFileOffset)
        Debug.Print Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
        Debug.Print Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
        Debug.Print OFN.sFileTitle

        
        'Short 8.3 File Name (using (OFN.sFileTitle):
        LongName = OFN.sFileTitle
        ShortName = Space$(128)
        ShortSize = Len(ShortName)
        sp = GetShortPathName(LongName, ShortName, ShortSize)
        sf_Name = Left$(ShortName, sp)
        If Len(sf_Name) Then
            If InStr(sf_Name, vbNullChar) Then sf_Name = Left$(sf_Name, InStr(sf_Name, vbNullChar) - 1)
        End If

    
'        'Short 8.3 File Name (using OFN.sFile):
'        LongName = OFN.sFile
'        shortName = Space$(128)
'        ShortSize = Len(shortName)
'        sp = GetShortPathName(LongName, shortName, ShortSize)
'        sFullPath = Left$(shortName, sp)
        
        'File Path (from OFN.nFileOffset):
        sPathRoot = Left$(OFN.sFile, OFN.nFileOffset)
        If Len(sPathRoot) Then
            If InStr(sPathRoot, vbNullChar) Then sPathRoot = Left$(sf_Name, InStr(sPathRoot, vbNullChar) - 1)
        End If


'    Dim startStr As String
'    Dim tmp As String
'
'    startStr = Trim(OFN.sFile)
'
'    Do
'
'      tmp = GetPart(startStr, Chr$(0))
'      List1.AddItem tmp
'
'    Loop While Len(startStr) > 0


'    Text1 = OFN.sFile
'    Text2 = Left$(OFN.sFile, OFN.nFileOffset)
'    Text3 = Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
'    Text4 = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
'    Text5 = OFN.sFileTitle
'
'    LongName = OFN.sFileTitle
'    shortName = Space$(128)
'    ShortSize = Len(shortName)
'
'    sp = GetShortPathName(LongName, shortName, ShortSize)
'    Text6 = LCase$(Left$(shortName, sp))
'
'    LongName = OFN.sFile
'    shortName = Space$(128)
'    ShortSize = Len(shortName)
'
'    sp = GetShortPathName(LongName, shortName, ShortSize)
'    Text7 = LCase$(Left$(shortName, sp))
'
'    chkReadOnly.Value = Abs((OFN.flags And OFN_READONLY) = OFN_READONLY)

  End If

End Sub
Public Sub Do_Open()

  Dim sp As Long
  Dim LongName As String
  Dim ShortName As String
  Dim ShortSize As Long
  Dim sBuf As String

 'to keep lines short, I've abbreviated
 'a Null$ to n, and the filter$ to f.
  Dim n As String, n2 As String, f As String
  n = Chr$(0)
  n2 = n & n

 '------------------------------------------------
 'INITIALIZATION
 '------------------------------------------------
 'fill in the size of the OFN structure
  OFN.nStructSize = Len(OFN)

 'assign the owner of the dialog; this can be null if no owner.
  OFN.hwndOwner = lhWnd

 '------------------------------------------------
 'FILTERS
 '------------------------------------------------
 'There are 2 methods of setting filters (patterns) for
 'use in the dropdown combo of the dialog.

 'The first, using OFN.sFilter, fills the combo with the
 'specifed filters, and works as the VB common dialog does.
 'These must be in the "Friendly Name"-null$-Extention format,
 'terminating with 2 null strings.
'      f = "Visual Basic Forms" & n & "*.frm" & n
'  f = f & "Visual Basic Modules" & n & "*.bas" & n
'  f = f & "Visual Basic Projects" & n & "*.vbp" & n
'  f = f & "Text Files" & n & "*.txt" & n
'  f = f & "All Files" & n & "*.*" & n2
  
    f = sFilter
    f = AddNullChar(f)
'    If Len(sFilter) Then
'        If Right$(sFilter, 2) <> n2 Then sFilter = sFilter & n2
'    Else
'        sFilter = n2
'    End If
  
  OFN.sFilter = f '& n2

 'The second method, uses sCustomFilter and nCustFilterSize
 'to pass the filters to use and the size of the filter string.
 'The operating system copies the strings to the buffer when
 'the user closes the dialog box. The system uses the strings
 'to initialize the user-defined file filter the next time the
 'dialog box is created. If this parameter is NULL, the dialog
 'box lists but does not save user-defined filter strings.

 'To see the difference, comment out the OFN.sFilter = f above,
 'and uncomment the 2 lines below.
 'OFN.sCustomFilter = f
 'OFN.nCustFilterSize = Len(OFN.sCustomFilter)

 'nFilterIndex specifies an index into the buffer pointed to
 'by sFilter. The system uses the index value to obtain a pair
 'of strings to use as the initial filter description and
 'filter pattern for the dialog box. The first pair of strings
 'has an index value of 1. When the user closes the dialog box,
 'the system copies the index of the selected filter strings
 'into this location.
  'OFN.nFilterIndex = 2  '"Visual Basic Modules"

 '------------------------------------------------
 'FILENAME
 '------------------------------------------------
 'sFile points to a buffer that contains a filename used to
 'initialize the File Name edit control. The first character
 'of this buffer must be NULL if initialization is not necessary.
 'When the GetOpenFileName or GetSaveFileName function returns,
 'this buffer contains the drive designator, path, filename, and
 'extension of the selected file.

 'perform no filename initialization (Filename textbox is blank)
 'and initialize the sFile buffer for the return value
 ' OFN.sFile = Chr$(0)
 ' OFN.sFile =  Space$(1024)
 'OR
 'pass a default filename and initialize for return value
  
  sBuf = sf_Name & Space$(1024)
  sBuf = AddNullChar(sBuf)
  
  OFN.sFile = sBuf 'sf_Name & Space$(1024) & n2
  OFN.nFileSize = Len(OFN.sFile)

 'default extention applied to a selected file if it has no extention
  'OFN.sDefFileExt = "*.bas" & n2

 'sFileTitle points to a buffer that receives the title of the
 'selected file. The application should use this string
 'to display the file title. If this member is NULL, the
 'function does not copy the file title.

 
  OFN.sFileTitle = Space$(512) & n2
  OFN.nTitleSize = Len(OFN.sFileTitle) & n2

 'sInitDir is the string that specifies the initial file directory.
 'If this member is NULL, the system uses the current directory as
 'the initial directory.

  sBuf = AddNullChar(sDef_Dir)
  OFN.sInitDir = sBuf 'sDef_Dir & n2

 '------------------------------------------------
 'MISC
 '------------------------------------------------
 'sDlgTitle is the title to display in the dialog.  If null,
 'the default title for the dialog is used.
 sBuf = AddNullChar(sDlgTitle)
  OFN.sDlgTitle = sBuf 'sDlgTitle & n2

 'flags are the actions and options for the dialog.
  OFN.Flags = OFS_FILE_OPEN_FLAGS ' Or OFN_ALLOWMULTISELECT

 'Show the File Open Dialog
  If GetOpenFileName(OFN) Then

        Debug.Print OFN.sFile
        sFullPath = OFN.sFile
        If Len(sFullPath) Then
            If InStr(sFullPath, vbNullChar) Then sFullPath = Left$(sFullPath, InStr(sFullPath, vbNullChar) - 1)
        End If
        
        
        Debug.Print Left$(OFN.sFile, OFN.nFileOffset)
        Debug.Print Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
        Debug.Print Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
        Debug.Print OFN.sFileTitle

        
        'Short 8.3 File Name (using (OFN.sFileTitle):
        LongName = OFN.sFileTitle
        ShortName = Space$(128)
        ShortSize = Len(ShortName)
        sp = GetShortPathName(LongName, ShortName, ShortSize)
        sf_Name = Left$(ShortName, sp)
        If Len(sf_Name) Then
            If InStr(sf_Name, vbNullChar) Then sf_Name = Left$(sf_Name, InStr(sf_Name, vbNullChar) - 1)
        End If

    
'        'Short 8.3 File Name (using OFN.sFile):
'        LongName = OFN.sFile
'        shortName = Space$(128)
'        ShortSize = Len(shortName)
'        sp = GetShortPathName(LongName, shortName, ShortSize)
'        sFullPath = Left$(shortName, sp)
        
        'File Path (from OFN.nFileOffset):
        sPathRoot = Left$(OFN.sFile, OFN.nFileOffset)
        If Len(sPathRoot) Then
            If InStr(sPathRoot, vbNullChar) Then sPathRoot = Left$(sf_Name, InStr(sPathRoot, vbNullChar) - 1)
        End If

'    Dim startStr As String
'    Dim tmp As String
'
'    startStr = Trim(OFN.sFile)
'
'    Do
'
'      tmp = GetPart(startStr, Chr$(0))
'      List1.AddItem tmp
'
'    Loop While Len(startStr) > 0
'
'
'    Text1 = OFN.sFile
'    Text2 = Left$(OFN.sFile, OFN.nFileOffset)
'    Text3 = Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
'    Text4 = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
'    Text5 = OFN.sFileTitle
'
'    LongName = OFN.sFileTitle
'    shortName = Space$(128)
'    ShortSize = Len(shortName)
'
'    sp = GetShortPathName(LongName, shortName, ShortSize)
'    Text6 = LCase$(Left$(shortName, sp))
'
'    LongName = OFN.sFile
'    shortName = Space$(128)
'    ShortSize = Len(shortName)
'
'    sp = GetShortPathName(LongName, shortName, ShortSize)
'    Text7 = LCase$(Left$(shortName, sp))
'
'    chkReadOnly.Value = Abs((OFN.flags And OFN_READONLY) = OFN_READONLY)

  End If

End Sub
Private Function AddNullChar(ByVal sVal As String) As String

    Do While Right$(sVal, 2) <> Chr$(0) & Chr$(0)
        sVal = sVal & Chr$(0)
    Loop
    AddNullChar = sVal

End Function

