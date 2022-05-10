Attribute VB_Name = "ftRange"
Option Explicit
Option Compare Text
Option Base 1

Public Const ERR_RANGE_SRC_TARGET_MISMATCH = vbObjectError + 522
Public Const ERR_INVALID_FT_ARGUMENTS = vbObjectError + 523
Public Const ERR_NO_ERROR As String = "(INFO)"
Public Const ERR_ERROR As String = "(ERROR)"
Public Const ERR_INVALID_RANGE_SIZE = vbObjectError + 529
Public Const ERR_RANGE_AREA_COUNT = vbObjectError + 531
Public Const ERR_INVALID_ARRAY_SIZE = vbObjectError + 535

Public Type ArrInfo
    Rows As Long
    Columns As Long
    Dimensions As Long
    Ubound_first As Long
    LBound_first As Long
    UBound_second As Long
    LBound_second As Long
End Type
Public Type RangeInfo
    Rows As Long
    Columns As Long
    AreasSameRows As Boolean
    AreasSameColumns As Boolean
    Areas As Long
End Type
Public Type AreaStruct
    RowStart As Long
    RowEnd As Long
    ColStart As Long
    ColEnd As Long
    rowCount As Long
    ColumnCount As Long
End Type

Public Function TestMe(srcRng As Range, find As Variant, replace As Variant)

        Dim arr() As Variant
        Dim aInfo As ArrInfo
        
        arr = BuildRC1(srcRng)
        aInfo = ArrayInfo(arr)
        
        'Get dimensions
        Dim arrD As Long
        arrD = aInfo.Dimensions
        
        Dim d1 As Long, d2 As Long
        For d1 = aInfo.LBound_first To aInfo.Ubound_first
            For d2 = aInfo.LBound_second To aInfo.UBound_second
                If arr(d1, d2) = find Then arr(d1, d2) = replace
            Next d2
        Next d1
        
        srcRng.Value2 = arr
        
        If aInfo.Dimensions > 0 Then
            Erase arr
        End If
        

End Function

Private Function ConvertArrToRCArr(ByVal arr As Variant) As Variant
'DONE
    Dim retV() As Variant, rwCount As Long, isBase0 As Boolean, arrIDX As Long

    If IsArrInit(arr) = False Then
        ReDim retV(1 To 1, 1 To 1)
        retV(1, 1) = arr
        ConvertArrToRCArr = retV
        Exit Function
    End If

    If ArrayDimensions(arr) = 1 Then
        isBase0 = LBound(arr) = 0
        rwCount = UBound(arr) - LBound(arr) + 1
        If isBase0 Then
            ReDim retV(1 To UBound(arr) + 1, 1 To 1)
        Else
            ReDim retV(1 To UBound(arr), 1 To 1)
        End If
        
        For arrIDX = LBound(arr) To UBound(arr)
            If isBase0 Then
                retV(arrIDX + 1, 1) = arr(arrIDX)
            Else
                retV(arrIDX, 1) = arr(arrIDX)
            End If
        Next arrIDX
        ConvertArrToRCArr = retV
    Else
        ConvertArrToRCArr = arr
    End If

End Function

Public Function IsArrInit(inpt As Variant) As Boolean
    IsArrInit = ArrayDimensions(inpt) > 0
End Function

'Return Type: ArrInfo for any Array
'('Row' Count, 'Column' Count, Array Dimensions)
Public Function ArrayInfo(arr As Variant) As ArrInfo
On Error Resume Next
    Dim tmp As ArrInfo
    tmp.Dimensions = ArrayDimensions(arr)
    If tmp.Dimensions > 0 Then
        tmp.LBound_first = LBound(arr, 1)
        tmp.Ubound_first = UBound(arr, 1)
        tmp.Rows = (tmp.Ubound_first - tmp.LBound_first) + 1
    End If
    If tmp.Dimensions = 1 Then
        tmp.Columns = 1
    Else
        If tmp.Dimensions = 2 Then
            tmp.Columns = (UBound(arr, 2) - LBound(arr, 2)) + 1
        End If
    End If
    If tmp.Dimensions >= 2 Then
        tmp.LBound_second = LBound(arr, 2)
        tmp.UBound_second = UBound(arr, 2)
    End If
    ArrayInfo = tmp
End Function

Public Function BuildRC1(rng As Range) As Variant
On Error GoTo E:
    Dim failed As Boolean
    Dim retArray As Variant
    Dim rgInfo As RangeInfo
    
    If rng.Count = 1 Then
        ReDim retArray(1 To 1, 1 To 1)
        retArray(1, 1) = rng.value
        GoTo Finalize:
    End If
    
    rgInfo = RngInfo(rng)
    
 
    
    If rgInfo.Areas = 1 Then
        retArray = rng.value
        If ArrayDimensions(retArray) = 1 Then
            retArray = ConvertArrToRCArr(retArray)
        End If
        GoTo Finalize:
    End If
    
    'Areas > 1
    If rgInfo.AreasSameRows = False And rgInfo.AreasSameColumns = False Then
        Err.Raise ERR_INVALID_RANGE_SIZE, Description:="All areas in Range must have matching RowCount or ColumnCount (ftRangeArray.BuildRC1)"
    End If
    
    ReDim retArray(1 To rgInfo.Rows, 1 To rgInfo.Columns)
    
    ' ***** ***** ***** ***** ***** ***** ***** ***** *****
    Dim areaInfo As AreaStruct
    Dim idxAREA As Long, rngArea As Range, idxAreaRow As Long, idxAreaCol As Long
    Dim idxArrayRow As Long, idxArrayCol As Long
    Dim arrayRowOffset As Long, arrayColOffset As Long
    ' ***** ***** ***** ***** ***** ***** ***** ***** *****
    
    arrayRowOffset = 0
    arrayColOffset = 0
    
    If rgInfo.AreasSameRows Then
        ' *** *** *** *** *** ***
        ' *** SAME ROWS *** *
        ' *** *** *** *** *** ***
        For idxAREA = 1 To rgInfo.Areas
            areaInfo = RangeArea(rng.Areas(idxAREA))
            For idxAreaRow = 1 To areaInfo.rowCount
                For idxAreaCol = 1 To areaInfo.ColumnCount
                    retArray(idxAreaRow, idxAreaCol + arrayColOffset) = rng.Areas(idxAREA)(idxAreaRow, idxAreaCol)
                Next idxAreaCol
            Next idxAreaRow
            arrayColOffset = arrayColOffset + areaInfo.ColumnCount
        Next idxAREA
    
    Else
        ' *** *** *** *** *** ***
        ' *** SAME COLS *** *
        ' *** *** *** *** *** ***
        For idxAREA = 1 To rgInfo.Areas
            areaInfo = RangeArea(rng.Areas(idxAREA))
            For idxAreaRow = 1 To areaInfo.rowCount
                For idxAreaCol = 1 To areaInfo.ColumnCount
                    retArray(idxAreaRow + arrayRowOffset, idxAreaCol) = rng.Areas(idxAREA)(idxAreaRow, idxAreaCol)
                Next idxAreaCol
            Next idxAreaRow
            arrayRowOffset = arrayRowOffset + areaInfo.rowCount
        Next idxAREA
    End If

Finalize:
    On Error Resume Next
    
    If Not failed Then
        BuildRC1 = retArray
    End If
    
    If ArrayDimensions(retArray) > 0 Then
        Erase retArray
    End If
    
    Set rngArea = Nothing
    
    Exit Function
E:
    failed = True
    'my custom error handler
    'ErrorCheck
    Resume Finalize:
    
End Function

'Calculates Number of Dimensions of Array
'Non-array types and will return 0, so you can use following to determin if 'arry' is an array
'[is array if true]   If ArrayDimensions([some variable]) > 0 Then ...
Public Function ArrayDimensions(ByRef arry As Variant) As Long
On Error Resume Next
    Dim dimCount As Long
    Do While Err.Number = 0
        Dim tmp As Variant
        tmp = UBound(arry, dimCount + 1)
        If Err.Number = 0 Then
            dimCount = dimCount + 1
        End If
    Loop
    If dimCount > 0 Then
        If UBound(arry) < LBound(arry) Then
            dimCount = 0
        End If
    End If
    ArrayDimensions = dimCount
End Function
Public Function RngInfo(rg As Range) As RangeInfo
    Dim retV As RangeInfo
    If rg Is Nothing Then
        retV.Rows = 0
        retV.Columns = 0
        retV.AreasSameRows = False
        retV.AreasSameColumns = False
        retV.Areas = 0
    Else
        retV.Rows = RangeRowCount(rg)
        retV.Columns = RangeColCount(rg)
        retV.AreasSameRows = ContiguousRows(rg)
        retV.AreasSameColumns = ContiguousColumns(rg)
        retV.Areas = rg.Areas.Count
    End If
    
    RngInfo = retV
End Function

Public Function RangeArea(rg As Range) As AreaStruct
   
   If rg.Areas.Count <> 1 Then
        'You need to have your own error raised here
        'RaiseError ERR_RANGE_AREA_COUNT, "Range Area Count <> 1 (ftRangeArray.RangeArea)"
        Err.Raise vbObjectError + 513, "your error desc"
    End If
    
    Dim retV As AreaStruct
    retV.RowStart = rg.Row
    retV.RowEnd = rg.Row + rg.Rows.Count - 1
    retV.ColStart = rg.Column
    retV.ColEnd = rg.Column + rg.Columns.Count - 1
    retV.rowCount = rg.Rows.Count
    retV.ColumnCount = rg.Columns.Count
    
    RangeArea = retV

End Function



Private Function RangeRowCount(ByVal rng As Range) As Long

    Dim tmpCount As Long
    Dim rowDict As Dictionary
    Dim rCount As Long, areaIDX As Long, rwIDX As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    'Check first if all First/Count are the same, if they are, no need to loop through everything
    If AreasMatchRows(rng) Then
        tmpCount = rng.Areas(1).Rows.Count
    Else
        Set rowDict = New Dictionary
        For areaIDX = 1 To rng.Areas.Count
            For rwIDX = 1 To rng.Areas(areaIDX).Rows.Count
                Dim realRow As Long
                realRow = rng.Areas(areaIDX).Rows(rwIDX).Row
                rowDict(realRow) = realRow
            Next rwIDX
        Next areaIDX
        tmpCount = rowDict.Count
    End If

Finalize:
    RangeRowCount = tmpCount
    Set rowDict = Nothing

End Function




'returns 0 if any area has different numbers of columns than another
Private Function RangeColCount(ByVal rng As Range) As Long

    Dim tmpCount As Long
    Dim colDict As Dictionary
    Dim firstCol As Long, areaIDX As Long, colIDX As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    
    If AreasMatchCols(rng) Then
        tmpCount = rng.Areas(1).Columns.Count
    Else
        Set colDict = New Dictionary
        For areaIDX = 1 To rng.Areas.Count
            firstCol = rng.Areas(areaIDX).Column
            colDict(firstCol) = firstCol
            For colIDX = 1 To rng.Areas(areaIDX).Columns.Count
                If colIDX > 1 Then colDict(firstCol + (colIDX - 1)) = firstCol + (colIDX - 1)
            Next colIDX
        Next areaIDX
        tmpCount = colDict.Count
    End If
    
Finalize:
    RangeColCount = tmpCount
    Set colDict = Nothing

End Function
Private Function ContiguousRows(rng As Range) As Boolean
'RETURNS TRUE IF HAS 1 AREA OR ALL AREAS SHARE SAME FIRST/LAST ROW
    Dim retV As Boolean

    If rng Is Nothing Then
        ContiguousRows = True
        Exit Function
    End If

    If rng.Areas.Count = 1 Then
        retV = True
    Else
        'If any Area is outside the min/max row of any other area then return false
        Dim loop1 As Long, loop2 As Long, isDiffRange As Boolean
        Dim l1Start As Long, l1End As Long, l2Start As Long, l2End As Long
        
        For loop1 = 1 To rng.Areas.Count
            l1Start = rng.Areas(loop1).Row
            l1End = l1Start + rng.Areas(loop1).Rows.Count - 1
            
            For loop2 = 1 To rng.Areas.Count
                l2Start = rng.Areas(loop2).Row
                l2End = l1Start + rng.Areas(loop2).Rows.Count - 1
                If l1Start < l2Start Or l1End > l2End Then
                    isDiffRange = True
                End If
                If isDiffRange Then Exit For
            Next loop2
            If isDiffRange Then Exit For
        Next loop1
    End If

    retV = Not isDiffRange

    ContiguousRows = retV

End Function

Private Function ContiguousColumns(rng As Range) As Boolean

'RETURNS TRUE IF HAS 1 AREA OR ALL AREAS SHARE SAME FIRST/LAST COLUMN
    Dim retV As Boolean

    If rng Is Nothing Then
        ContiguousColumns = True
        Exit Function
    End If


    If rng.Areas.Count = 1 Then
        retV = True
    Else
        'If any Area is outside the min/max row of any other area then return false
        Dim loop1 As Long, loop2 As Long, isDiffRange As Boolean
        Dim l1Start As Long, l1End As Long, l2Start As Long, l2End As Long
        
        For loop1 = 1 To rng.Areas.Count
            l1Start = rng.Areas(loop1).Column
            l1End = l1Start + rng.Areas(loop1).Columns.Count - 1
            
            For loop2 = 1 To rng.Areas.Count
                l2Start = rng.Areas(loop2).Column
                l2End = l1Start + rng.Areas(loop2).Columns.Count - 1
                If l1Start < l2Start Or l1End > l2End Then
                    isDiffRange = True
                End If
                If isDiffRange Then Exit For
            Next loop2
            If isDiffRange Then Exit For
        Next loop1
    End If

    retV = Not isDiffRange

    ContiguousColumns = retV

End Function

Private Function AreasMatchRows(rng As Range) As Boolean

    If rng Is Nothing Then
        AreasMatchRows = True
        Exit Function
    End If

    Dim retV As Boolean
    If rng.Areas.Count = 1 Then
        retV = True
    Else
        Dim firstRow As Long, firstCount As Long, noMatch As Boolean, aIDX As Long
        firstRow = rng.Areas(1).Row
        firstCount = rng.Areas(1).Rows.Count
        For aIDX = 2 To rng.Areas.Count
            With rng.Areas(aIDX)
                If .Row <> firstRow Or .Rows.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIDX
        retV = Not noMatch
    End If

    AreasMatchRows = retV

End Function

Private Function AreasMatchCols(rng As Range) As Boolean

    Dim retV As Boolean
    If rng.Areas.Count = 1 Then
        retV = True
    Else
        Dim firstCol As Long, firstCount As Long, noMatch As Boolean, aIDX As Long
        firstCol = rng.Areas(1).Column
        firstCount = rng.Areas(1).Columns.Count
        For aIDX = 2 To rng.Areas.Count
            With rng.Areas(aIDX)
                If .Column <> firstCol Or .Columns.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIDX
        retV = Not noMatch
    End If

    AreasMatchCols = retV

End Function
