Attribute VB_Name = "FlexEdit"
Option Explicit
Public Const kFG_L = 1
Public Const kFG_H = 2
Public Const kFG_Sx = 3
Public Const kFG_Up = 4
Public Const kFG_Rx = 5
Public Const kFG_Down = 6
Public Const kFG_Row = 7
Public Const kFG_Col = 8
Public Const kEti_mRow = 9
Public Const kEti_mCol = 10
Public Const kEti_L = 11
Public Const kEti_H = 12
Public Const kOff_X = 13
Public Const kOff_Y = 14

Public Enum iPropG
    am_Bool = 1
    am_Long = 2
    am_Dec = 4
    am_Date = 8
    am_String = 16
End Enum
Public Type elemFlexG
    sField As String
    vVal As Variant
    iType As iPropG
    iLen As Integer
End Type
Public Type LayOutG
    dDB As Boolean
    sField As String
    bVisib As Boolean
    bComp As Boolean
    bStrong As Boolean
    xPos As Integer
    yPos As Integer
    iLen As Integer
    iType As iPropG
    bForce As Boolean
End Type
Public Type ColProperty
    Key As Boolean
    Id As Long
    Name As String
    DataType As String
    Size As Integer
    Nulls As Boolean
    Default As String
    Collation As String
End Type
Public Type rptPage
    sName As String
    iHeight As Long
    iWidth As Long
    iPageBreak As Long
End Type
Public Type PageElement
    sName As String
    vVal As Variant
    iType As iPropG
    iLen As Integer
End Type

Public Sub SettaLargFlex(ByRef flexObj As MSFlexGrid, ByVal l_intCol2Large As Integer, ByVal lngWdef As Long)

   Dim l_lngLarg As Long
   Dim l_lngSumLarg As Long
   Dim l_intLoop As Integer
   Dim l_lngWork As Long
   Dim lngVSBarW As Long
   Dim lngWCol2Large As Long
   
   lngVSBarW = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CYVTHUMB)
              
    With flexObj

        l_lngLarg = .Width
        For l_intLoop = 0 To .Cols - 1
            l_lngSumLarg = l_lngSumLarg + .ColWidth(l_intLoop) + .GridLineWidth
            If l_intLoop <> l_intCol2Large Then
                l_lngWork = l_lngWork + .ColWidth(l_intLoop)
            Else
                lngWCol2Large = .ColWidth(l_intLoop)
            End If
            l_lngWork = l_lngWork + .GridLineWidth
        Next
      

        If hasScrollBars(.hWnd) Then
            lngWCol2Large = l_lngLarg - l_lngWork - lngVSBarW - k_FlexBevel
        Else
            lngWCol2Large = l_lngLarg - l_lngWork - k_FlexBevel
        End If
        If lngWCol2Large <= 0 Then lngWCol2Large = lngWdef
        .ColWidth(l_intCol2Large) = lngWCol2Large
    End With
   
End Sub

