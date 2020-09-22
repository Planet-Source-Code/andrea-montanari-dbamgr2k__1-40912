VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'************************************
'
'   classe di Gestione PopUp Menu
'
'   versione        :
'   ultima revisione:
'************************************

Option Explicit

Public MenuItems As New Collection
Public hMenu As Long

Public Sub Clear()
    Dim o As Object
    
    For Each o In MenuItems
        Call o.Clear
    Next
    
    Set MenuItems = Nothing
    
End Sub

'Public Function Add_Ret(ByVal Caption As String, ByVal ID As Long, Optional Flags) As cMnuItem
'    Dim o As New cMnuItem
'
'    o.Caption = Caption
'    o.ID = ID
'    If Not IsMissing(Flags) Then
'        o.Flags = Flags
'    End If
'    Set o.Parent = Me
'    MenuItems.Add o, "M" & o.ID
'
'    Set Add_Ret = o
'
'End Function

Public Sub Add(ByVal Caption As String, ByVal ID As Long, _
                    Optional bSeparator As Boolean = False, _
                    Optional bDisable As Boolean = False)
            
    Dim o As New cMnuItem
    Dim lFlags As Long
    
    If bSeparator Then lFlags = MF_SEPARATOR
    If bDisable Then lFlags = MF_DISABLED Or MF_GRAYED
    
    o.Caption = Caption
    o.ID = ID
    'If Not IsMissing(Flags) Then
        o.Flags = lFlags
    'End If
    Set o.Parent = Me
    MenuItems.Add o, "M" & o.ID
    
End Sub

Public Sub BuildMenu()
    Dim o As cMnuItem
    Dim l As Long
    
    Call FreeMenu
    
    '// allocate handle here
    hMenu = CreatePopupMenu()
    
    For Each o In MenuItems
        o.BuildMenu
        
        l = o.Flags And (Not MF_DEFAULT)
        '// add item to menu here
        If o.hMenu <> 0 Then
            Call AppendMenu(hMenu, l Or MF_POPUP, o.hMenu, o.Caption)
        Else
            Call AppendMenu(hMenu, l, o.ID, o.Caption)
        End If
        If o.Flags And MF_DEFAULT Then
            Call SetMenuDefaultItem(hMenu, o.ID, MF_BYCOMMAND)
        End If
    Next
    
End Sub

Public Sub FreeMenu()
    Dim o As cMnuItem
    
    If hMenu <> 0 Then
        For Each o In MenuItems
            If o.hMenu <> 0 Then o.FreeMenu
        Next
        '// delete menu here
        Call DestroyMenu(hMenu)
    End If
    hMenu = 0
    
End Sub


Public Function Popup(ByVal hWnd As Long, ByVal x As Single, ByVal y As Single) As Long

    Dim xx As Long, yy As Long
    Dim r As RECT
    
    Call BuildMenu
    
    '// convert x, y to coordinates
    xx = x
    yy = y
    r.Bottom = 100
    r.Right = 100
    
    '// popup menu here
    Popup = TrackPopupMenu(hMenu, &H100, xx, yy, 0, hWnd, r)
    
End Function


Private Sub Class_Terminate()
    
    FreeMenu
    Clear
    Set MenuItems = Nothing
    
End Sub

