Attribute VB_Name = "mdlRangeInfo"
Option Explicit

Public Type RngInfo
    Rows As Long
    Columns As Long
    AreasSameRows As Boolean
    AreasSameColumns As Boolean
    Areas As Long
End Type

'Public Function ResetIt()
'    wsRangeInfo.Range("D18:D22").ClearContents
'End Function
'
'Public Function TestIt()
'    If Not Selection Is Nothing Then
'        Dim ri As RngInfo
'        ri = RangeInfo(Selection)
'        wsRangeInfo.Range("D18").value = ri.Rows
'        wsRangeInfo.Range("D19").value = ri.Columns
'        wsRangeInfo.Range("D20").value = ri.Areas
'        wsRangeInfo.Range("D21").value = ri.AreasSameRows
'        wsRangeInfo.Range("D22").value = ri.AreasSameColumns
'
'    End If
'End Function

Public Function RangeInfo(rg As Range) As RngInfo
    Dim retV As RngInfo
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
    
    RangeInfo = retV
End Function

Public Function RangeRowCount(ByVal rng As Range) As Long

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
Public Function RangeColCount(ByVal rng As Range) As Long

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

Public Function ContiguousRows(rng As Range) As Boolean
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

Public Function ContiguousColumns(rng As Range) As Boolean

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
