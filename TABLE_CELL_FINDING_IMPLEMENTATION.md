# Table Cell Finding Implementation Guide

## Overview
This document describes how to implement the `tableCell` structure in the VBA code to reliably find and edit text within Word table cells.

## Problem Statement
The current approach uses comma-separated text or pipe-separated text in the `context` field (e.g., `"Distance loss | -20 dB (30m)"`). This fails because:
1. Word doesn't store table text with separators - each cell is independent
2. Text normalization doesn't preserve table structure
3. The Find operation searches across the entire document text, not table structure

## Solution: tableCell Structure

The LLM now outputs a `tableCell` object alongside (or instead of) the context field:

```json
{
  "action": "replace",
  "tableCell": {
    "rowHeader": "Distance loss",
    "columnHeader": "Predicted Value",
    "adjacentCells": {
      "above": "Criterion",
      "left": "Distance loss",
      "cellContent": "-20 dB (30m)"
    }
  },
  "target": "-20 dB (30m)",
  "replace": "-28 dB (78m)",
  "explanation": "Correcting distance calculation"
}
```

## VBA Implementation Pseudocode

### Main Function: FindTableCell

```vba
Private Function FindTableCell(ByVal suggestion As Object, ByVal searchRange As Range) As Range
    ' Returns the Range of the specific table cell, or Nothing if not found

    On Error GoTo ErrorHandler

    ' 1. Check if suggestion has tableCell structure
    If Not HasDictionaryKey(suggestion, "tableCell") Then
        Set FindTableCell = Nothing
        Exit Function
    End If

    Dim tableCellInfo As Object
    Set tableCellInfo = suggestion("tableCell")

    ' 2. Extract search parameters
    Dim rowHeader As String
    Dim columnHeader As String
    Dim adjacentCells As Object

    rowHeader = GetSuggestionText(tableCellInfo, "rowHeader", "")
    columnHeader = GetSuggestionText(tableCellInfo, "columnHeader", "")

    If HasDictionaryKey(tableCellInfo, "adjacentCells") Then
        Set adjacentCells = tableCellInfo("adjacentCells")
    End If

    ' 3. Find the table (search all tables in the search range)
    Dim tbl As Table
    Dim targetCell As Cell
    Dim foundCell As Range

    For Each tbl In searchRange.Tables
        Set targetCell = FindCellInTable(tbl, rowHeader, columnHeader, adjacentCells)

        If Not targetCell Is Nothing Then
            ' Found the cell! Return its range
            Set foundCell = targetCell.Range
            ' Remove the trailing cell marker (Chr 13 & Chr 7)
            foundCell.End = foundCell.End - 1
            Set FindTableCell = foundCell
            Exit Function
        End If
    Next tbl

    ' Not found in any table
    Set FindTableCell = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "Error in FindTableCell: " & Err.Description
    Set FindTableCell = Nothing
End Function
```

### Helper Function: FindCellInTable

```vba
Private Function FindCellInTable(ByVal tbl As Table, _
                                 ByVal rowHeader As String, _
                                 ByVal columnHeader As String, _
                                 ByVal adjacentCells As Object) As Cell
    ' Searches within a specific table for the target cell

    On Error Resume Next

    Dim targetRow As Long
    Dim targetCol As Long
    Dim foundRow As Boolean
    Dim foundCol As Boolean

    foundRow = False
    foundCol = False

    ' Strategy 1: Find by row header (first cell in row)
    If Len(rowHeader) > 0 Then
        Dim r As Long
        For r = 1 To tbl.Rows.Count
            Dim firstCellText As String
            firstCellText = GetCellText(tbl.Cell(r, 1))

            If InStr(1, firstCellText, rowHeader, vbTextCompare) > 0 Then
                targetRow = r
                foundRow = True
                Exit For
            End If
        Next r
    End If

    ' Strategy 2: Find by column header (first row)
    If Len(columnHeader) > 0 Then
        Dim c As Long
        For c = 1 To tbl.Columns.Count
            Dim headerText As String
            headerText = GetCellText(tbl.Cell(1, c))

            If InStr(1, headerText, columnHeader, vbTextCompare) > 0 Then
                targetCol = c
                foundCol = True
                Exit For
            End If
        Next c
    End If

    ' Strategy 3: Use adjacentCells to find the exact cell
    If Not adjacentCells Is Nothing Then
        Set FindCellInTable = FindCellByAdjacentContent(tbl, adjacentCells, targetRow, targetCol, foundRow, foundCol)
        Exit Function
    End If

    ' Strategy 4: Return cell at row/column intersection
    If foundRow And foundCol Then
        On Error Resume Next
        Set FindCellInTable = tbl.Cell(targetRow, targetCol)
        If Err.Number <> 0 Then Set FindCellInTable = Nothing
        Exit Function
    End If

    ' Strategy 5: If only row found, search across that row
    If foundRow And Not foundCol Then
        ' The LLM should specify which column, but we can search the entire row
        ' This is handled by adjacentCells or cellContent
        Set FindCellInTable = Nothing ' Need more info
        Exit Function
    End If

    ' Not found
    Set FindCellInTable = Nothing
End Function
```

### Helper Function: FindCellByAdjacentContent

```vba
Private Function FindCellByAdjacentContent(ByVal tbl As Table, _
                                           ByVal adjacentCells As Object, _
                                           ByVal hintRow As Long, _
                                           ByVal hintCol As Long, _
                                           ByVal hasHintRow As Boolean, _
                                           ByVal hasHintCol As Boolean) As Cell
    ' Uses adjacent cell content to pinpoint the exact cell

    On Error Resume Next

    Dim cellContent As String
    Dim aboveText As String
    Dim belowText As String
    Dim leftText As String
    Dim rightText As String

    ' Extract what we're looking for
    cellContent = GetSuggestionText(adjacentCells, "cellContent", "")
    aboveText = GetSuggestionText(adjacentCells, "above", "")
    belowText = GetSuggestionText(adjacentCells, "below", "")
    leftText = GetSuggestionText(adjacentCells, "left", "")
    rightText = GetSuggestionText(adjacentCells, "right", "")

    ' If we have row/col hints, start there
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long

    If hasHintRow Then
        startRow = hintRow
        endRow = hintRow
    Else
        startRow = 1
        endRow = tbl.Rows.Count
    End If

    If hasHintCol Then
        startCol = hintCol
        endCol = hintCol
    Else
        startCol = 1
        endCol = tbl.Columns.Count
    End If

    ' Search through the candidate cells
    Dim r As Long
    Dim c As Long

    For r = startRow To endRow
        For c = startCol To endCol
            Dim currentCell As Cell
            On Error Resume Next
            Set currentCell = tbl.Cell(r, c)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextCell
            End If
            On Error GoTo 0

            ' Check if this cell matches all criteria
            Dim matches As Boolean
            matches = True

            ' Check cell content
            If Len(cellContent) > 0 Then
                If InStr(1, GetCellText(currentCell), cellContent, vbTextCompare) = 0 Then
                    matches = False
                End If
            End If

            ' Check above
            If matches And Len(aboveText) > 0 And r > 1 Then
                On Error Resume Next
                Dim aboveCell As Cell
                Set aboveCell = tbl.Cell(r - 1, c)
                If Err.Number = 0 Then
                    If InStr(1, GetCellText(aboveCell), aboveText, vbTextCompare) = 0 Then
                        matches = False
                    End If
                Else
                    matches = False
                End If
                Err.Clear
                On Error GoTo 0
            End If

            ' Check below
            If matches And Len(belowText) > 0 And r < tbl.Rows.Count Then
                On Error Resume Next
                Dim belowCell As Cell
                Set belowCell = tbl.Cell(r + 1, c)
                If Err.Number = 0 Then
                    If InStr(1, GetCellText(belowCell), belowText, vbTextCompare) = 0 Then
                        matches = False
                    End If
                Else
                    matches = False
                End If
                Err.Clear
                On Error GoTo 0
            End If

            ' Check left
            If matches And Len(leftText) > 0 And c > 1 Then
                On Error Resume Next
                Dim leftCell As Cell
                Set leftCell = tbl.Cell(r, c - 1)
                If Err.Number = 0 Then
                    If InStr(1, GetCellText(leftCell), leftText, vbTextCompare) = 0 Then
                        matches = False
                    End If
                Else
                    matches = False
                End If
                Err.Clear
                On Error GoTo 0
            End If

            ' Check right
            If matches And Len(rightText) > 0 And c < tbl.Columns.Count Then
                On Error Resume Next
                Dim rightCell As Cell
                Set rightCell = tbl.Cell(r, c + 1)
                If Err.Number = 0 Then
                    If InStr(1, GetCellText(rightCell), rightText, vbTextCompare) = 0 Then
                        matches = False
                    End If
                Else
                    matches = False
                End If
                Err.Clear
                On Error GoTo 0
            End If

            ' If all checks passed, we found it!
            If matches Then
                Set FindCellByAdjacentContent = currentCell
                Exit Function
            End If

NextCell:
        Next c
    Next r

    ' Not found
    Set FindCellByAdjacentContent = Nothing
End Function
```

### Helper Function: GetCellText

```vba
Private Function GetCellText(ByVal c As Cell) As String
    ' Extracts normalized text from a cell
    On Error Resume Next

    Dim cellRange As Range
    Set cellRange = c.Range

    ' Remove the cell end markers (Chr 13 & Chr 7)
    Dim txt As String
    txt = cellRange.Text

    ' Remove trailing cell marker
    If Right$(txt, 1) = Chr$(7) Then
        txt = Left$(txt, Len(txt) - 1)
    End If
    If Right$(txt, 1) = Chr$(13) Then
        txt = Left$(txt, Len(txt) - 1)
    End If

    ' Normalize whitespace
    txt = NormalizeForDocument(txt)

    GetCellText = txt
End Function
```

## Integration into Existing Code

### Modify FindWithProgressiveFallback

Add tableCell check as Strategy 0 (before existing strategies):

```vba
Private Function FindWithProgressiveFallback(ByVal searchString As String, _
                                             ByVal searchRange As Range, _
                                             Optional ByVal matchCase As Boolean = False, _
                                             Optional ByVal suggestion As Object = Nothing) As Range
    On Error GoTo ErrorHandler

    Dim result As Range

    ' Strategy 0: Check for tableCell structure (NEW!)
    If Not suggestion Is Nothing Then
        If HasDictionaryKey(suggestion, "tableCell") Then
            Debug.Print "  - Strategy 0: Using tableCell structure..."
            Set result = FindTableCell(suggestion, searchRange)
            If Not result Is Nothing Then
                Debug.Print "    -> SUCCESS: Found via tableCell structure"
                Set FindWithProgressiveFallback = result
                Exit Function
            End If
            Debug.Print "    -> tableCell search failed, falling back to text search"
        End If
    End If

    ' Strategy 1: Try exact match with normalization
    ' ... (existing code continues)
```

### Update ProcessSuggestion Call

When calling `FindWithProgressiveFallback`, pass the suggestion object:

```vba
' Old code:
Set contextRange = FindWithProgressiveFallback(contextForSearch, searchRange, effectiveMatchCase)

' New code:
Set contextRange = FindWithProgressiveFallback(contextForSearch, searchRange, effectiveMatchCase, suggestion)
```

## Benefits

1. **Precise Location**: Uses table structure (rows/columns) instead of text patterns
2. **Adjacent Cell Context**: Can disambiguate cells with same content
3. **Robust**: Works even when cells contain special characters, formatting, or similar text
4. **Backward Compatible**: Falls back to regular text search if tableCell not provided
5. **LLM Friendly**: LLM can "see" adjacent cells and provide this context

## Testing Strategy

1. Test with simple 2-column tables
2. Test with complex multi-column tables
3. Test with merged cells (may need special handling)
4. Test with nested tables
5. Test with tables containing identical values
6. Test fallback when tableCell info is incomplete

## Example Use Cases

### Case 1: Simple Value Update
```json
{
  "action": "replace",
  "tableCell": {
    "rowHeader": "Distance loss",
    "adjacentCells": {
      "left": "Distance loss",
      "cellContent": "-20 dB"
    }
  },
  "target": "-20 dB",
  "replace": "-28 dB"
}
```

### Case 2: Multi-occurrence Disambiguation
```json
{
  "action": "replace",
  "tableCell": {
    "columnHeader": "Predicted Level",
    "adjacentCells": {
      "left": "Facade A",
      "above": "Predicted Level"
    }
  },
  "target": "42",
  "replace": "44"
}
```

### Case 3: Complex Table Navigation
```json
{
  "action": "replace",
  "tableCell": {
    "rowHeader": "Noise Source 1",
    "columnHeader": "LAeq",
    "adjacentCells": {
      "above": "LAeq",
      "left": "Noise Source 1",
      "cellContent": "65 dB"
    }
  },
  "target": "65 dB",
  "replace": "68 dB"
}
```
