Attribute VB_Name = "pbTmpFindRangeDups"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbLRange v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA
'
' Find Duplicates in Range with multiple areas
'
' @module pbTmpFindRangeDups
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1



Public Function FindDuplicateRows(rng As Range, ParamArray checkRangeCols() As Variant) As Dictionary
    'EXAMPLE CALL:   set  [myDictionary] = FindDuplicateRows(Worksheets(1).Range("B5:C100"))
    'EXAMPLE CALL:   set  [myDictionary] = FindDuplicateRows(Worksheets(1).Range("B5:H100"), 1,3,4)
        'Since Range Start on Column B, the columns that would be used to check duplicates would be B, D, E (the 1st, 3rd, and 4th columns in the range)

    'RETURNS DICTIONARY WHERE KEY=WORKSHEET ROW AND VALUE = NUMBER OF DUPLICATES
    'If No Value is passed in for 'checkRangeCols', then the entire row in the ranges will be compared to find duplicates
    'If  'rng' contains multiple areas (for example you passed in something like [range].SpecialCells(xlCellTypeVisible),
        'Then All Areas will be checked for Column Consistency (i.e. All areas in Range Must have identical Column Number and total Columns)
        'If all areas in Range to match column structure, error is raised
On Error GoTo E:
    Dim failed As Boolean
      
    ' ~~~ ~~~ check for mismatched range area columns ~~~ ~~~
    Dim firstCol As Long, totCols As Long
    Dim areaIDX As Long
    If rng.Areas.Count >= 1 Then
        firstCol = rng.Areas(1).column
        totCols = rng.Areas(1).Columns.Count
        For areaIDX = 1 To rng.Areas.Count
            If Not rng.Areas(areaIDX).column = firstCol _
                Or Not rng.Areas(areaIDX).Columns.Count = totCols Then
                Err.Raise 17, Description:="FindDuplicateRows can not support mismatched columns for multiple Range Areas"
            End If
        Next areaIDX
    End If
    
    Dim retDict As New Dictionary, tmpDict As New Dictionary, compareColCount As Long, tmpIdx As Long
    Dim checkCols() As Long
    retDict.CompareMode = TextCompare
    tmpDict.CompareMode = TextCompare
    
    If rng.Areas.Count = 1 And rng.Rows.Count = 1 Then
        GoTo Finalize:
    End If
    ' ~~~ ~~~ Determine Number of columns being compared for each row  ~~~ ~~~
    If UBound(checkRangeCols) = -1 Then
        compareColCount = rng.Areas(1).Columns.Count
        ReDim checkCols(1 To compareColCount)
        For tmpIdx = 1 To compareColCount
            checkCols(tmpIdx) = tmpIdx
        Next tmpIdx
    Else
        compareColCount = (UBound(checkRangeCols) - LBound(checkRangeCols)) + 1
        ReDim checkCols(1 To compareColCount)
        For tmpIdx = LBound(checkRangeCols) To UBound(checkRangeCols)
            checkCols(tmpIdx + 1) = checkRangeCols(tmpIdx)
        Next tmpIdx
    End If
    
    For areaIDX = 1 To rng.Areas.Count
        Dim rowIDX As Long, checkCol As Long, compareArr As Variant, curKey As String
        For rowIDX = 1 To rng.Areas(areaIDX).Rows.Count
            compareArr = GetCompareValues(rng.Areas(areaIDX), rowIDX, checkCols)
            curKey = Join(compareArr, ", ")
            If Not tmpDict.Exists(curKey) Then
                tmpDict(curKey) = rng.Rows(rowIDX).Row
            Else
                Dim keyFirstRow As Long
                keyFirstRow = CLng(tmpDict(curKey))
                'if it exists, then it's a duplicate
                If Not retDict.Exists(keyFirstRow) Then
                    'the first worksheet row with this values is Value from tmpDict
                    retDict(keyFirstRow) = 2
                Else
                    retDict(keyFirstRow) = CLng(retDict(keyFirstRow)) + 1
                End If
            End If
        Next rowIDX
    Next areaIDX
    
Finalize:
    If Not failed Then
        Set FindDuplicateRows = retDict
        
        'For Fun, List the Rows and How Many Duplicates Exist
       Dim dKey As Variant
       For Each dKey In retDict.Keys
            Debug.Print "Worksheet Row: " & dKey & ", has " & retDict(dKey) & " duplicates"
       Next dKey
        
    End If

    Exit Function
E:
    failed = True
    MsgBox "FindDuplicateRows failed. (Error: " & Err.Number & ", " & Err.Description & ")"
    Err.Clear
    Resume Finalize:

End Function

Private Function GetCompareValues(rngArea As Range, rngRow As Long, compCols() As Long) As Variant
    Dim valsArr As Variant
    Dim colcount As Long
    Dim idx As Long, curCol As Long, valCount As Long
    colcount = UBound(compCols) - LBound(compCols) + 1
    ReDim valsArr(1 To colcount)
    For idx = LBound(compCols) To UBound(compCols)
        valCount = valCount + 1
        curCol = compCols(idx)
        valsArr(valCount) = CStr(rngArea(rngRow, curCol).Value2)
    Next idx
    GetCompareValues = valsArr
End Function


