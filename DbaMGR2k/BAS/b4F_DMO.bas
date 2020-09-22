Attribute VB_Name = "b4F_DMO"
Option Explicit

Public Sub BB_Sort(ByRef vArr2Sort As Variant, Optional ByVal iItem2Test As Integer = 0)

    Dim Value As String
    Dim Index As Long
    Dim fItem As Long
    Dim indLimit As Long, lastSwap As Long
    Dim numEl As Long
    Dim iCol As Integer
    Dim iCols As Integer
    Dim iTest As Integer
        
    Dim vTmp As Variant
    
    ReDim vTmp(UBound(vArr2Sort, 1)) As Variant
        
    numEl = UBound(vArr2Sort, 2)
            
    fItem = 0
    lastSwap = UBound(vArr2Sort, 2)
    Do
        indLimit = lastSwap - 1
        lastSwap = 0
        For Index = fItem To indLimit
            For iCol = 0 To iCols
                vTmp(iCol) = vArr2Sort(iCol, Index)
            Next
            Value = vArr2Sort(iItem2Test, Index)
            iTest = StrComp(Value, vArr2Sort(iItem2Test, Index + 1), vbTextCompare)
            If iTest > 0 Then
                For iCol = 0 To iCols
                    vArr2Sort(iCol, Index) = vArr2Sort(iCol, Index + 1)
                    vArr2Sort(iCol, Index + 1) = vTmp(iCol)
                Next
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
        
End Sub

Public Sub PurgeFilesDetached(ByRef vArrFiles As Variant, ByVal sFilter As String)
    
    Dim lLoop As Long
    Dim vRet As Variant
    Dim bIsLike As Boolean
    Dim bAdd As Boolean
    
    sFilter = UCase$(sFilter)
    
    bIsLike = Len(sFilter) <> 0
        
    For lLoop = 0 To UBound(vArrFiles, 2)
        bAdd = True
        Debug.Print vArrFiles(2, lLoop), vArrFiles(0, lLoop)
        
        If bIsLike Then
            bAdd = UCase$(vArrFiles(0, lLoop)) Like UCase$(sFilter)
        End If
        If bAdd Then
            vRet = Add2VarArray(Array(vArrFiles(0, lLoop)), vRet)
        End If
    Next
    vArrFiles = vRet

End Sub
Public Function Add2VarArray(ByVal varArray1 As Variant, ByVal varArray2 As Variant) As Variant
        
    Dim iCol As Integer
    Dim iLoopC As Integer
    Dim lRow2 As Long
        
    iCol = UBound(varArray1)
    If IsArray(varArray2) Then
        lRow2 = UBound(varArray2, 2) + 1
        ReDim Preserve varArray2(iCol, lRow2) As Variant
    Else
        ReDim varArray2(iCol, lRow2) As Variant
    End If
        
    For iLoopC = 0 To iCol
        varArray2(iLoopC, lRow2) = varArray1(iLoopC)
    Next
    Add2VarArray = varArray2
    
End Function

