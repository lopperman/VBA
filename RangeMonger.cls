VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeMonger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' RangeMonger v0.1.3
' (c) Paul Brower - https://github.com/lopperman/VBA
'
' Utilities for working with Ranges and ListObjects
'
' @class RangeMonger
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Private l_sheetProtection As ProtectionBase

'If a class implementing ProtectionBase is passed in, then if a sort/filter would fail because a sheet is protected,
'then it will automatically be unprotected
'the caller is required to make sure sheet is reprotected, if necessary
Public Property Set Unprotector(protector As ProtectionBase)
    Set l_sheetProtection = protector
End Property

Private Function Protect(wksht As Worksheet) As Boolean
    If wksht.ProtectionMode Then
        Protect = True
        Exit Function
    End If
    
    If Not l_sheetProtection Is Nothing Then
        Protect = l_sheetProtection.Protect(wksht)
    Else
        Protect = False
    End If
End Function

Private Function Unprotect(wksht As Worksheet) As Boolean
    If wksht.ProtectionMode = False Then
        Unprotect = True
        Exit Function
    End If
    If Not l_sheetProtection Is Nothing Then
        Unprotect = l_sheetProtection.Unprotect(wksht)
    Else
        Unprotect = False
    End If
End Function

'Remove All List Rows Where [field] matches [criteria]
'can specify full or partial match
'Throws error if Worksheet is Protected Against Deleting ListRows
'Returns Number of Deleted ListRows
Public Function DeleteFoundListRows(listObj As ListObject, field As Variant, criteria As Variant, MatchType As XlLookAt, Optional MatchCase As Boolean = False) As Long
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowDeletingRows = False Then
        If Unprotect(listObj.Parent) = False Then
            Err.Raise 419, listObj, "Protected sheet does not allow deleting rows"
        End If
    End If
    
    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    Dim sanityCount As Long, rCount As Long, nextRow As Long, deletedCount As Long
        
    With listObj
        rCount = .ListRows.Count
        Do While True
            nextRow = 0
            If .ListRows.Count = 0 Then
                Exit Do
            End If
            Dim fnd As Range
            Set fnd = .ListColumns(fieldIdx).DataBodyRange.find(criteria, LookAt:=MatchType, MatchCase:=MatchCase)
            
            If Not fnd Is Nothing Then
                nextRow = fnd.row - .HeaderRowRange.row
                If nextRow > 0 Then
                    .ListRows(nextRow).Delete
                    deletedCount = deletedCount + 1
                End If
            Else
                Exit Do
            End If
            
            sanityCount = sanityCount + 1
            If sanityCount > rCount Then
                Exit Do
            End If
        Loop
    End With

    DeleteFoundListRows = deletedCount

End Function

Private Function GetFieldIndex(ByRef listObj As ListObject, field As Variant) As Long
    Dim fieldIdx As Integer
    If IsNumeric(field) Then
        fieldIdx = CLng(field)
    Else
        fieldIdx = listObj.ListColumns(field).Index
    End If
    GetFieldIndex = fieldIdx
End Function

'Return the range in a ListObject of all the matching values in a single ListColumn
Public Function GetFoundRange(listObj As ListObject, field As Variant, criteria As Variant) As Range

    Dim foundRange As Range

    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)

    Dim firstRow As Long, lastrow As Long
    firstRow = GetFirstRowIndex(listObj, fieldIdx, criteria, True)
    If firstRow > 0 Then
        lastrow = GetLastRowIndex(listObj, fieldIdx, criteria)
        Set foundRange = listObj.ListColumns(fieldIdx).DataBodyRange(RowIndex:=firstRow)
        Set foundRange = foundRange.Resize(RowSize:=(lastrow - firstRow) + 1)
    End If

    Set GetFoundRange = foundRange

End Function

Public Function GetFirstRowIndex(listObj As ListObject, field As Variant, criteria As Variant, Optional sortIfNeeded As Boolean = True) As Long
    GetFirstRowIndex = GetRowIndex(listObj, field, criteria, True, sortIfNeeded)
End Function
Public Function GetLastRowIndex(listObj As ListObject, field As Variant, criteria As Variant, Optional sortIfNeeded As Boolean = True) As Long
    GetLastRowIndex = GetRowIndex(listObj, field, criteria, False, sortIfNeeded)
End Function

'Get ListRow Index of first or last occurence of [criteria] in ListColumn [field]
'Returns 0 if not found,, returns -1 if [field] not sorted
Private Function GetRowIndex(listObj As ListObject, field As Variant, criteria As Variant, firstRowIndex As Boolean, Optional sortIfNeeded As Boolean = True) As Long
    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    With listObj
        If .ListRows.Count = 0 Then
            GetRowIndex = 0
            GoTo Finalize:
        End If
        
        ClearFilter listObj
     
        'Make sure [field] is sorted
        If .Sort.SortFields.Count = 0 Then
            If sortIfNeeded Then
                AddSort listObj, field, clearPreviousSorts:=True
            End If
        Else
            Dim srtField As SortField, searchMode As Variant
            Set srtField = .Sort.SortFields(1)
            If srtField.key.column <> .ListColumns(fieldIdx).Range.column Then
                If sortIfNeeded Then
                    AddSort listObj, field, clearPreviousSorts:=True
                End If
            End If
        End If
        
        
        If firstRowIndex Then
            searchMode = 1
        Else
            searchMode = -1
        End If
        Dim crit As Variant
        Dim handled As Boolean
        If handled = False And TypeName(criteria) = "Date" Then
            handled = True
            crit = CLng(CDate(criteria))
        End If
        If handled = False Then
            crit = criteria
        End If
        Dim srchRng As Range
        Set srchRng = listObj.ListColumns(fieldIdx).DataBodyRange
        On Error Resume Next
        Dim fR As Variant: fR = WorksheetFunction.IfError(WorksheetFunction.XMatch(crit, srchRng, 0, searchMode), 0)
        If IsNumeric(fR) Then
            foundRow = CLng(fR)
        End If
    End With
    
Finalize:
    GetRowIndex = foundRow

End Function

'Add Sort to ListObject, optionally clearing previous sorts
Public Function AddSort(listObj As ListObject, field As Variant, Optional Order As XlSortOrder = xlAscending, Optional clearPreviousSorts As Boolean = False, Optional clearFilters As Boolean = True) As Boolean
On Error GoTo e:

    If listObj.Range.Worksheet.Protection.AllowSorting = False Then
        If Unprotect(listObj.Range.Worksheet) = False Then
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If


    Dim fieldIdx As Integer, rngInfo As String
    fieldIdx = GetFieldIndex(listObj, field)
    
    rngInfo = listObj.Name & "[" & listObj.ListColumns(fieldIdx).Name & "]"
    If listObj.ListRows.Count > 0 Then
        If clearFilters Then
            ClearFilter listObj
        End If
        Dim sortAlreadyValid As Boolean
        sortAlreadyValid = True
        With listObj.Sort
            If clearPreviousSorts = True Then
                .SortFields.Clear
                .SortFields.Add2 key:=Range(rngInfo), SortOn:=xlSortOnValues, Order:=Order, DataOption:=xlSortNormal
                .header = xlYes
                .Apply
            Else
                .SortFields.Add2 key:=Range(rngInfo), SortOn:=xlSortOnValues, Order:=Order, DataOption:=xlSortNormal
                .header = xlYes
                .Apply
            End If
            
        End With
    End If
   AddSort = True
   Exit Function
e:
   AddSort = False
        
End Function

'Filters for exact match, Returns count of filtered rows
Public Function AddFilterSimple(listObj As ListObject, field As Variant, crit1 As Variant, Optional clearExistFilters As Boolean = False) As Long
    Dim cnt As Long
    Dim fieldIdx As Integer
    fieldIdx = GetFieldIndex(listObj, field)
    
    If listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If Unprotect(listObj.Range.Worksheet) = False Then
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If
 
    
    If Not listObj Is Nothing And listObj.ListRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        End If
        Dim handled As Boolean
        With listObj
            If TypeName(crit1) = "Date" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & Int(crit1), Operator:=xlAnd, Criteria2:="<" & Int(crit1) + 1
            End If
            If handled = False And TypeName(crit1) = "String" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=crit1
            End If
            If handled = False And IsNumeric(crit1) Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & crit1, Operator:=xlAnd, Criteria2:="<=" & crit1
            End If
            cnt = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
        End With
    End If


    AddFilterSimple = cnt
End Function

Public Function AddFilterBetween(listObj As ListObject, field As Variant, crit1 As Variant, crit2 As Variant, Optional clearExistFilters As Boolean = False) As Long
    Dim cnt As Long
    Dim fieldIdx As Integer
    fieldIdx = GetFieldIndex(listObj, field)
    
    If listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If Unprotect(listObj.Range.Worksheet) = False Then
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If
 
    
    If Not listObj Is Nothing And listObj.ListRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        End If
        Dim handled As Boolean
        With listObj
            If TypeName(crit1) = "Date" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & Int(crit1), Operator:=xlAnd, Criteria2:="<=" & Int(crit2)
            End If
            If handled = False And TypeName(crit1) = "String" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & crit1, Operator:=xlAnd, Criteria2:="<=" & crit2
            End If
            If handled = False And IsNumeric(crit1) Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & crit1, Operator:=xlAnd, Criteria2:="<=" & crit2
            End If
            cnt = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
        End With
    End If


    AddFilterBetween = cnt
End Function

Public Function ClearFilter(listObj As ListObject) As Boolean
On Error GoTo e:
    Dim failed As Boolean
    
    If listObj Is Nothing Then
        GoTo Finalize:
    End If
    If listObj.ListRows.Count = 0 Then
        GoTo Finalize:
    End If
    
    If listObj.Range.Worksheet.ProtectionMode And listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If Unprotect(listObj.Range.Worksheet) = False Then
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If
    
    If FilterCount(listObj) > 0 Then
        listObj.AutoFilter.ShowAllData
    End If
    
Finalize:
    ClearFilter = Not failed
    
    Exit Function
e:
   failed = True
   Err.Raise Err.Number, Err.Source, Err.Description
   Resume Finalize:
    
End Function



'Get count of enabled filters
Public Function FilterCount(lstobj As ListObject) As Long
    Dim fltrCount As Long
    
    If Not lstobj Is Nothing Then
        Dim fltrIdx As Long
        If lstobj.AutoFilter.Filters.Count > 0 Then
            Dim f As Variant
            For fltrIdx = 1 To lstobj.AutoFilter.Filters.Count
                If lstobj.AutoFilter.Filters(fltrIdx).On Then
                    fltrCount = fltrCount + 1
                End If
            Next fltrIdx
        End If
    End If
    
    FilterCount = fltrCount
End Function

'Returns true if all items in [rng] are in the listObject in validcolIdx
Public Function RangeIsInsideListColumn(rng As Range, lstobj As ListObject, validColIdx As Long) As Boolean

    If rng Is Nothing Or lstobj Is Nothing Or validColIdx = 0 Then
        RangeIsInsideListColumn = False
        Exit Function
    End If
    If rng.Count = 1 Then
        If Intersect(rng, lstobj.ListColumns(validColIdx).DataBodyRange) Is Nothing Then
            RangeIsInsideListColumn = False
            Exit Function
        Else
            RangeIsInsideListColumn = True
            Exit Function
        End If
    End If
    
    If rng.Areas.Count = 1 Then
        If Intersect(rng, lstobj.ListColumns(validColIdx).DataBodyRange) Is Nothing Then
            RangeIsInsideListColumn = False
            Exit Function
        Else
            'check columns
            If rng.Columns.Count <> 1 Then
                RangeIsInsideListColumn = False
                Exit Function
            End If
            RangeIsInsideListColumn = True
            Exit Function
        End If
    Else
        Dim rngArea As Range
        For Each rngArea In rng.Areas
            If RangeIsInsideListColumn(rngArea, lstobj, validColIdx) = False Then
                RangeIsInsideListColumn = False
                Exit Function
            End If
        Next rngArea
    End If
    
    RangeIsInsideListColumn = True

End Function
