Attribute VB_Name = "knjname_MyXlsFunc_xls"
Option Explicit

'* Returns a cell based on r but having different column(specified by colCell) or row(specified by rowCell).
'*
'* Typical usage:
'* <pre><code>
'* For Each eachCell in targetedCells
'*   Debug.Print "The value of A&C is " & iC(eachCell, "A") & iC(eachCell, "C") & "."
'* Next
'* </code></pre>
'*
'* @param[in] r base range
'* @param[in] colCell
'*   Either "column name"(e.g. "A", "XA"), column_number(e.g. 1 for "A", 2 for "B" ...) or a cell (its column address is used.)
'* @param[in] rowCell
'*   Either number of row (e.g. 1, 2), or a cell (its row address is used.)
'* @returns A cell pointed with arguments.
Function iC( _
    ByVal r As Range, _
    Optional ByVal colCell As Variant, _
    Optional ByVal rowCell As Variant _
  ) As Range
    
    Dim colAt As Variant, rowAt As Variant

    If Not IsMissing(colCell) And "Nothing" <> TypeName(colCell) Then
        If "Range" = TypeName(colCell) Then
            colAt = colCell.Column
        Else
            colAt = colCell
        End If
    Else
        colAt = r.Column
    End If
    
    If Not IsMissing(rowCell) And "Nothing" <> TypeName(rowCell) Then
        If "Range" = TypeName(rowCell) Then
            rowAt = rowCell.row
        Else
            rowAt = rowCell
        End If
    Else
        rowAt = r.row
    End If
    
    Set iC = r.Worksheet.Cells(rowAt, colAt)
End Function

Function iCByName(ByVal r As Range, Optional ByVal colName$ = "", Optional ByVal rowName$ = "") As Range
    Dim colCell As Range, rowCell As Range

    With r.Worksheet
        If Len(colName) > 0 Then
            Set colCell = .Range(colName)
        End If
        If Len(rowName) > 0 Then
            Set rowCell = .Range(rowName)
        End If
        
        Set iCByName = iC(r, colCell, rowCell)
    End With
End Function

Function copySheetToLast(ByVal copied As Worksheet, Optional ByRef copyTo As Workbook) As Worksheet
    If Not copyTo Is Nothing Then
        copied.Copy After:=copyTo.Sheets(copyTo.Sheets.Count)
    Else
        copied.Copy
        Set copyTo = Workbooks(Workbooks.Count)
    End If
    Set copySheetToLast = copyTo.Sheets(copyTo.Sheets.Count)
End Function

Function rangeAsIterable(ByVal r As Range) As Object
    If r Is Nothing Then
        Set rangeAsIterable = New Collection
    Else
        Set rangeAsIterable = r
    End If
End Function

Function getNonConflictSheetName$(ByVal sheetNameCandidate$, ByVal within As Workbook)
    ' TODO implement
End Function

Function putCellValuesH(ByVal r As Range, ParamArray values() As Variant)
    If r Is Nothing Then
        
    End If

End Function

Sub test_iC()
    Debug.Print iC(ActiveCell, ActiveCell, 3)
End Sub
