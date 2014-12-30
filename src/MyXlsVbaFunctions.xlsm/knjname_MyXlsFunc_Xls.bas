Attribute VB_Name = "knjname_MyXlsFunc_xls"
Option Explicit

' Depends on "Micsoroft VBScript Regular Expressions 5.5"

Private xlsVbaFuncFSO As New FileSystemObject


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

Function hasWorksheet(ByVal wb As Workbook, ByVal sheetName$) As Boolean
    hasWorksheet = Not getWorksheet(wb, sheetName) Is Nothing
End Function

Function getWorkbook(ByVal bookName$) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = bookName Then
            Set getWorkbook = wb
            Exit Function
        End If
    Next
End Function

Function getWorksheet(ByVal wb As Workbook, ByVal sheetName$) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            Set getWorksheet = ws
            Exit Function
        End If
    Next
End Function

Function findFirstWorksheet(ByVal wb As Workbook, ByVal sheetRegexp As regexp) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If sheetRegexp.Test(ws.Name) Then
            Set findFirstWorksheet = ws
            Exit Function
        End If
    Next
End Function

Function findAllSheets(ByVal wb As Workbook, ByVal sheetRegexp As regexp) As Collection
    Set findAllSheets = New Collection
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If sheetRegexp.Test(ws.Name) Then
            findAllSheets.Add ws
        End If
    Next
End Function

Function getNonConflictSheetName$(ByVal sheetNameCandidate$, ByVal within As Workbook)
    ' TODO implement
End Function

Function putCellValues(ByRef r As Range, ParamArray values() As Variant) As Range
    If r Is Nothing Then
        Set r = Workbooks.Add.Worksheets(1).[A1]
    End If
    
    r.Resize(UBound(values) - LBound(values)) = values
    
    Set putCellValues = r
    moveToNextRow r
End Function

Function moveToNextRow(ByRef r As Range, Optional ByVal moveOffsetRow = 1, Optional ByVal moveOffsetColumn = 0) As Range
    Set r = r.Offset(moveOffsetRow, moveOffsetColumn)
    Set moveToNextRow = r
End Function

Function shortenTooLongDocumentPath$(ByVal documentPath$)
    If documentPath$ > 260 Then
        shortenTooLongDocumentPath$ = xlsVbaFuncFSO.GetFile(documentPath).ShortPath
    Else
        shortenTooLongDocumentPath$ = documentPath
    End If
End Function

Function openWorkbookQuietly(ByVal wbPath$, _
    Optional ByVal readonly = True, _
    Optional ByVal makeEventsOff As Boolean = False, _
    Optional ByVal findFromOpenedBooksAtFirst As Boolean = True, _
    Optional ByVal updateLinks As Boolean = False) As Workbook
    
On Error GoTo finalizingProc
    
    wbPath = shortenTooLongDocumentPath(wbPath)
    
    If makeEventsOff Then
        Dim currentEventsEnabled
        currentEventsEnabled = Application.EnableEvents
        Application.EnableEvents = False
    End If
    
    If findFromOpenedBooksAtFirst Then
        Set openWorkbookQuietly = getWorkbook(xlsVbaFuncFSO.GetFile(wbPath).Name)
    End If
    
    If openWorkbookQuietly Is Nothing Then
        Set openWorkbookQuietly = Workbooks.Open(wbPath, updateLinks:=updateLinks, readonly:=readonly, ignorereadonlyrecommended:=True, notify:=False)
    End If
    
    If makeEventsOff Then
        Application.EnableEvents = currentEventsEnabled
    End If
    
    Exit Function

finalizingProc:

    If makeEventsOff Then
        Application.EnableEvents = currentEventsEnabled
    End If
    
    ' Rethrow
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Function limitToUsedRange(ByVal r As Range) As Range
    Set limitToUsedRange = Application.Intersect(r.Worksheet.UsedRange, r)
End Function

Function betterEndXlDown(ByVal r As Range) As Object
    If Len(r) = 0 Then
        Set betterEndXlDown = New Collection
    ElseIf Len(r.Offset(1)) = 0 Then
        Set betterEndXlDown = r.Offset(1)
    Else
        Set betterEndXlDown = r.End(xlDown)
    End If
End Function
