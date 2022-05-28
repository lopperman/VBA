Attribute VB_Name = "mdlPivotTable"
Option Explicit
Option Base 1
Option Compare Text

Public Function BuildDemoPivot()
    Beep
    Dim lo As ListObject, rng As Range
    
    'Completely Clear the Sheet to put the Pivot Table
    wsPivot.Cells.Clear
    
    Set lo = wsDemoData.ListObjects("tblProject")
    Set rng = wsPivot.Range("A10")
    
    If BuildPivot(lo, rng, "coolPivotTable") Then
        wsPivot.Activate
    End If
    
End Function
    
Public Function BuildPivot(Optional srcListObj As ListObject, Optional destRng As Range, Optional pvtName As String = vbNullString) As Boolean
On Error GoTo E:
        
        Dim failed As Boolean
        Dim pvt As PivotTable
        
        If srcListObj Is Nothing Then
            'Set srcLstObj = [A REFERENCE TO A LIST OBJECT]
            'ex: Set srcLstOb =ThisWorkbook.Worksheets("PivotDemo").ListObjects("tblProject")
        End If
        If destRng Is Nothing Then
            'set destRng =[SINGLE CELL RANGE WHERE PVT TABLE WILL BE CREATED
            'ex: Set destRng = ThisWorkbook.Worksheets("tmpSheet").Range("A10")
        End If
        
        'Set Reference to the ListObject (datasource for Pivot Table)
        If Len(pvtName) = 0 Then
            pvtName = "tmpPivot1"
        End If

        'Build the Pivot Table Object - this will create an 'empty' PivotTable, ]
        '  similar to what you see when you've done this manually
        '  before add pivot fields
        Set pvt = CreateTempPivot(srcListObj, destRng, pvtName)
        
        'There are 4 types of Pivot Fields (RowField and DataField are the most common - There's also ColumnField and PageField)
        'I've provided examples to add Row and Data fields. Those could also be called outside this method
        
        'this example assumes your list object has the following columns:
        '   Project, ProjectDesc, ItemDate, Employee, Role, BillRate, Hours, TotRev

        'Add Row Fields
        'NOTE: The order you add these fields, is the order default grouping/subtotaling, etc, will occur
        CreatePivotRowField pvt, "Employee"
        CreatePivotRowField pvt, "Project"
        CreatePivotRowField pvt, "ProjDesc"
        CreatePivotRowField pvt, "Role"
        
        'Optional if you use this -- past in the fields you don't want additional subotals on certain Row Fields
        RemoveSubtotals pvt, True, Array("Project", "ProjDesc", "Role")
        
        'Add 'Data' Fields (to sum,min, max, average, etc
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
