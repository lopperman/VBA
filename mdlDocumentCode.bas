Attribute VB_Name = "mdlDocumentCode"
Option Explicit
Option Base 1

Public Function GetFormulas(wb As Workbook)
    
    Dim c As New Collection
    'Format to store in collection
    '0=SheetName, 1=CellAddress or TableName.ColumnName (if list object), 2=Formula, 3=Formula2R1C1
    
    Dim ws As Worksheet, lo As ListObject, tRng As Range, fRng As Range, cl As Range, idx As Long, lstCol As Long
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set fRng = Nothing
        Set fRng = ws.usedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number <> 0 Then
            'no formulas on the sheet
            Err.Clear
        Else
            If Not fRng Is Nothing Then
                For idx = 1 To fRng.Areas.Count
                    Set tRng = fRng.Areas(idx)
                    For Each cl In tRng
                        'we'll get the ListObject formulas on the next pass
                        If cl.ListObject Is Nothing Then
                            If cl.HasFormula Then
                                c.add Array(ws.Name, cl.Address, "'" & cl.formula, "'" & cl.Formula2R1C1)
                            End If
                        End If
                    Next cl
                Next idx
                For Each lo In ws.ListObjects
                    If lo.listRows.Count = 0 Then
                        Beep
                        Debug.Print "Could not get formulas for table: " & lo.Name & " because it has 0 ListRows"
                    Else
                        For lstCol = 1 To lo.ListColumns.Count
                            If lo.ListColumns(lstCol).DataBodyRange(1, 1).HasFormula Then
                                c.add Array(ws.Name, lo.Name & "[" & lo.ListColumns(lstCol).Name & "]", "'" & lo.ListColumns(lstCol).DataBodyRange(1, 1).formula, "'" & lo.ListColumns(lstCol).DataBodyRange(1, 1).Formula2R1C1)
                            End If
                        Next lstCol
                    End If
                Next lo
            End If
        End If
    Next ws

    Dim arr() As Variant, col As Long
    ReDim arr(1 To c.Count, 1 To 4)
    For idx = 1 To c.Count
        For col = 1 To 4
            arr(idx, col) = c(idx)(col)
        Next col
    Next idx
    
    With Workbooks.add
        .Worksheets(1).Range("A1").Resize(rowSize:=c.Count, ColumnSize:=4).value = arr
        .Activate
    End With

End Function
