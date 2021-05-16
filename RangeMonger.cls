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
' RangeMonger v0.0.1
' (c) Paul Brower - https://github.com/vba-tools/vba-test
'
' Utilities for working with Ranges and ListObjects
'
' @class RangeMonger
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

'Add Sort to ListObject, optionally clearing previous sorts
Public Function AddSort(listObj As ListObject, field As Variant, Optional Order As XlSortOrder = xlAscending, Optional clearPreviousSorts As Boolean = False) As Boolean
    Dim fieldIdx As Integer, rngInfo As String
    If IsNumeric(field) Then
        fieldIdx = CLng(field)
    Else
        fieldIdx = listObj.ListColumns(field).Index
    End If
    rngInfo = listObj.Name & "[" & listObj.ListColumns(fieldIdx).Name & "]"
    With listObj.Sort
        If clearPreviousSorts = True Then
            .SortFields.Clear
        End If
        .SortFields.Add2 key:=Range(rngInfo), SortOn:=xlSortOnValues, Order:=Order, DataOption:=xlSortNormal
        .header = xlYes
        .Apply
    End With
End Function

'Filters for exact match, Returns count of filtered rows
Public Function AddFilterSimple(listObj As ListObject, field As Variant, crit1 As Variant, Optional clearExistFilters As Boolean = False) As Long
    Dim cnt As Long
    Dim fieldIdx As Integer
    If IsNumeric(field) Then
        fieldIdx = CLng(field)
    Else
        fieldIdx = listObj.ListColumns(field).Index
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

Public Function ClearFilter(listObj As ListObject) As Boolean
On Error GoTo E:
    Dim failed As Boolean
    
    If listObj Is Nothing Then
        GoTo Finalize:
    End If
    If listObj.ListRows.Count = 0 Then
        GoTo Finalize:
    End If
    
    If listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        Err.Raise 419, listObj, "Protected sheet does not allow filtering"
    End If
    
    If FilterCount(listObj) > 0 Then
        listObj.AutoFilter.ShowAllData
    End If
    
Finalize:
    ClearFilter = Not failed
    
    Exit Function
E:
   failed = True
   Err.Raise Err.Number, Err.Source, Err.Description
   Resume Finalize:
    
End Function



'Get count of enabled filters
Public Function FilterCount(lstObj As ListObject) As Long
    Dim fltrCount As Long
    
    If Not lstObj Is Nothing Then
        Dim fltrIdx As Long
        If lstObj.AutoFilter.Filters.Count > 0 Then
            Dim f As Variant
            For fltrIdx = 1 To lstObj.AutoFilter.Filters.Count
                If lstObj.AutoFilter.Filters(fltrIdx).On Then
                    fltrCount = fltrCount + 1
                End If
            Next fltrIdx
        End If
    End If
    
    FilterCount = fltrCount
End Function
