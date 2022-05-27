Attribute VB_Name = "mdlPivotTable"
Option Explicit
Option Base 1
Option Compare Text

    
Public Function BuildPivot(srcListObj As ListObject, _
    destRng As Range, pvtName As String, _
    rowFields() As String, dataFields() As String) As Boolean
On Error GoTo E:
        Dim failed As Boolean

        Dim pvt As PivotTable
        Set pvt = CreateTempPivot(srcListObj, destRng, pvtName)
        
        'There are 4 types of Pivot Fields (RowField and DataField are the most common - There's also ColumnField and PageField)
        'Add Row Fields
        CreatePivotRowField pvt, "Project"
        CreatePivotRowField pvt, "ProjDesc"

        CreatePivotRowField pvt, "Employee"
        CreatePivotRowField pvt, "Role"
        
        'Optional if you use this -- past in the fields you don't want additional subotals on
        RemoveSubtotals pvt, True, Array("Employee", "Role")
        
        CreatePivotRowField pvt, "lkBillRate"
        CreatePivotDataField pvt, "Hours", "BillHours", xlSum, "#,##0.00"
        CreatePivotDataField pvt, "TotRev", "ToBill", xlSum, "$#,##0.00"
        
Finalize:
    On Error Resume Next

    BuildPivot = Not failed

    Exit Function
E:
    failed = True
    'ErrorCheck (My Custom Error Handler -- you'll have to implement your own
    Resume Finalize:
End Function


Private Function CreateTempPivot(tmpLO As ListObject, destRng As Range, pvtName As String) As PivotTable
    
    'Create a new empty PivotTable
    ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=tmpLO.Name, Version:=6) _
        .CreatePivotTable TableDestination:=destRng, _
        tableName:=pvtName, DefaultVersion:=6
        
    Dim pvt As PivotTable
    Set pvt = destRng.Worksheet.PivotTables(pvtName)
        
        'set the various pivottable options you want
        With pvt
            .AllowMultipleFilters = True
            .ShowTableStyleColumnHeaders = False
            .ShowTableStyleRowHeaders = False
            .LayoutRowDefault = xlTabularRow
            .EnableDrilldown = False
            .EnableFieldList = True
            .EnableWizard = True
            .EnableWriteback = False
            .RepeatItemsOnEachPrintedPage = True
            .ShowPageMultipleItemLabel = False
            .RowAxisLayout xlTabularRow
            .SubtotalHiddenPageItems = False
            .SaveData = False
    
            .TableStyle2 = "PivotStyleLight2"
            .ShowTableStyleColumnHeaders = True
            .ShowTableStyleRowHeaders = True
            .ShowTableStyleColumnStripes = False
            .ShowTableStyleRowStripes = True
            .InGridDropZones = False
            .NullString = "Blank"
            .DisplayFieldCaptions = True
            .ShowDrillIndicators = False
            .RepeatAllLabels xlRepeatLabels
            .RowAxisLayout xlTabularRow
            .ShowTableStyleRowStripes = True
        
        End With
        
        Set tmpLO = Nothing
        
        Set CreateTempPivot = pvt
        Set pvt = Nothing
        
End Function


Private Function CreatePivotDataField(pvt As PivotTable, colName As String, aliasName As String, cFunction As XlConsolidationFunction, Optional fmt As String = vbNullString)
'   Create a DataField (summarized)
'   Example CreatePivotDataField myPivotTable, "Revenue", "Total-Revenue", xlSum, "_($* #,##0_);_($* (#,##0);_(* "" - ""??_);_(@_)"
    With pvt
        .AddDataField .PivotFields(colName), aliasName, cFunction
        If fmt <> vbNullString Then
            .PivotFields(aliasName).numberFormat = fmt
        End If
    End With
End Function

Private Function CreatePivotRowField(pvt As PivotTable, colName As String)
'   Example:  CreatePivotField myPivotTabl, "Project"

    With pvt.PivotFields(colName)
        .orientation = xlRowField
    End With
End Function

Private Function RemoveSubtotals(pvt As PivotTable, showColumnGrandTotal As Boolean, removeSubT As Variant)
'   Example: RemoveSubTotals myPivotTable, True, Array("Col1","Col2","Col3')

    'Technically, all your listobject fields are in the pivot table (they just might be hidden)
    'So you could run this on all your list object column names, even if their not visible

    Dim i As Long
    For i = LBound(removeSubT) To UBound(removeSubT)
        With pvt.PivotFields(CStr(removeSubT(i)))
            .Subtotals(1) = False
        End With
    Next i
    
    pvt.ColumnGrand = showColumnGrandTotal

End Function
