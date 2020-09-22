VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMnuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'************************************
'
'   classe di PopUp Menu ITEM
'
'   versione        :
'   ultima revisione:
'************************************

Option Explicit

Public Caption As String
Public ID As Long
Public Flags As Long
Public MenuItems As New Collection
Public Parent As Object

Public hMenu As Long

'Public Function Add(ByVal Caption As String, ByVal ID As Long, Optional Flags) As cMnuItem
'    Dim o As New cMnuItem
'
'    o.Caption = Caption
'    o.ID = ID
'    If Not IsMissing(Flags) Then
'        o.Flags = Flags
'    End If
'    Set o.Parent = Me
'
'    MenuItems.Add o, "M" & o.ID
'
'    Set Add = o
'
'End Function

Public Sub BuildMenu()
    Dim o As cMnuItem
    
    Call FreeMenu
    
    If MenuItems.Count > 0 Then
        hMenu = CreatePopupMenu()
        For Each o In MenuItems
            o.BuildMenu
            
            '// add item to menu here
            If o.hMenu <> 0 Then
                Call AppendMenu(hMenu, o.Flags Or MF_POPUP, o.hMenu, o.Caption)
            Else
                Call AppendMenu(hMenu, o.Flags, o.ID, o.Caption)
            End If
        Next
    End If
    
End Sub

Public Sub Clear()
    Dim o As Object
    
    For Each o In MenuItems
        Call o.Clear
    Next
    
    'Set o.Parent = Nothing
    Set MenuItems = Nothing
    
End Sub

'Public Sub EnableItem(ByVal Enabled As Boolean)
'
'    '// update flags
'    If Not Enabled Then
'        Flags = Flags Or (MF_ENABLED Or MF_GRAYED)
'    Else
'        Flags = Flags And (Not (MF_ENABLED Or MF_GRAYED))
'    End If
'
'    '// if the menu exists, set the status
'    If Parent.hMenu <> 0 Then
'        If Enabled Then
'            Call EnableMenuItem(Parent.hMenu, Me.ID, MF_BYCOMMAND Or (MF_ENABLED Or MF_GRAYED))
'        Else
'            Call EnableMenuItem(Parent.hMenu, Me.ID, MF_BYCOMMAND Or MF_UNCHECKED)
'        End If
'    End If
'
'End Sub
'
'Public Sub CheckItem(ByVal Checked As Boolean)
'
'    '// update flags
'    If Checked Then
'        Flags = Flags Or MF_CHECKED
'    Else
'        Flags = Flags And (Not MF_CHECKED)
'    End If
'
'    '// if the menu exists, set the status
'    If Parent.hMenu <> 0 Then
'        If Checked Then
'            Call CheckMenuItem(Parent.hMenu, Me.ID, MF_BYCOMMAND Or MF_CHECKED)
'        Else
'            Call CheckMenuItem(Parent.hMenu, Me.ID, MF_BYCOMMAND Or MF_UNCHECKED)
'        End If
'    End If
'
'End Sub

Public Sub FreeMenu()
    Dim o As cMnuItem
    
    If hMenu <> 0 Then
        For Each o In MenuItems
            If o.hMenu <> 0 Then o.FreeMenu
        Next
        '// delete menu here
        'Call DestroyMenu(hMenu)
    End If
    hMenu = 0
    
End Sub

Private Sub Class_Terminate()
        
    FreeMenu
    Clear
    
End Sub

