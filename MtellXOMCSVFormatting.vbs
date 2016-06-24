Option Explicit
Public SegmentCSV, Segment, PHDTagName, PHDTagDescription, Asset, Path As String

Sub ImportFile()
Dim sPath As String
Dim i, j As Integer

Sheets("CSVList").Select
Range("A1").Select
Selection.End(xlDown).Select
j = ActiveCell.Row - 1

Asset = ThisWorkbook.Worksheets("CSVList").Range("B2")
    
For i = 1 To j
    Sheets("CSVList").Select
    Range("A" & i + 1).Select
    SegmentCSV = Range("A" & i + 1).Value
    
    
    Segment = Replace(SegmentCSV, ".csv", "")
    
    'Below we assume that the file, csvtest.csv,
    'is in the same folder as the workbook. If
    'you want something more flexible, you can
    'use Application.GetOpenFilename to get a
    'file open dialogue that returns the name
    'of the selected file.
    'On the page Fast text file import
    'I show how to do that - just replace the
    'file pattern "txt" with "csv".
    Path = ThisWorkbook.Path
    sPath = ThisWorkbook.Path & "\" & SegmentCSV
        
    'Procedure call. Semicolon is defined as separator,
    'and data is to be inserted on "Sheet2".
    'Of course you could also read the separator
    'and sheet name from the worksheet or an input
    'box. There are several options.
    copyDataFromCsvFileToSheet sPath, ";", "CSVData"
    
Next i

'This will help to save memory when runing the macro
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
'**************************************************************
Private Sub copyDataFromCsvFileToSheet(parFileName As String, _
parDelimiter As String, parSheetName As String)

Dim Data As Variant  'Array for the file values
Dim TargetRow, TargetCol As Integer
Dim Arr() As Variant ' declare an unallocated array

'Deleting/Creating CSVData Sheet will ensure empty cells that have not data
'and add bad data to the CSV
Sheets("CSVData").Select
ActiveWindow.SelectedSheets.Delete
CreateSheet ("CSVData")

Sheets("CSVData").Select
Sheets("CSVData").Cells.Select
Selection.ClearContents
'Function call - the file is read into the array
Data = getDataFromFile(parFileName, parDelimiter)

    'If the array isn't empty it is inserted into
    'the sheet in one swift operation.
    If Not isArrayEmpty(Data) Then
      'If you want to operate directly on the array,
      'you can leave out the following lines.
      With Sheets(parSheetName)
        'Delete any old content
        .Cells.ClearContents
        'A range gets the same dimensions as the array
        'and the array values are inserted in one operation.
        .Cells(1, 1).Resize(UBound(Data, 1), UBound(Data, 2)) = Data
      End With
    End If
    
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
    TrailingMinusNumbers:=True
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Timestamp"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Tag Name"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Tag Description"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Tag Value"
    Range("B2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B2").Select
    
    'Get Column and row number
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    TargetRow = ActiveCell.Row
    TargetCol = ActiveCell.Column
    
    Sheets("TagCrossReference").Select
    
    'Get Tag Name and Tag Description
    Cells.Find(What:=Segment, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False).Activate
    
    ActiveCell.Offset(0, 1).Select
    PHDTagName = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    PHDTagDescription = ActiveCell.Value
    
    Sheets("CSVData").Select
    
    'Fill Cells with PHD Tag and Description
    ThisWorkbook.Worksheets("CSVData").Range("B2:B" & TargetRow) = PHDTagName
    ThisWorkbook.Worksheets("CSVData").Range("C2:C" & TargetRow) = PHDTagDescription
    
    'Format date
    Arr = Range("A2:A" & TargetRow) ' Arr is now an allocated array
    Dim R As Long
    Dim C As Long
    Dim ArrDate As String
    
    'Modify this section to redo the date format
    For R = 1 To UBound(Arr, 1) ' First array dimension is rows.
        ArrDate = Arr(R, 1)
        ArrDate = Left(ArrDate, Len(ArrDate) - 4)
        Arr(R, 1) = ArrDate
    Next R
    
    Dim Destination As Range
    Set Destination = Range("A2")
    Destination.Resize(UBound(Arr), 1) = Arr
    
    'Format the date to be "m/d/yyyy h:mm:ss"
    Range("A2:A" & TargetRow).Select
    Selection.NumberFormat = "m/d/yyyy h:mm:ss"
    
    'Copy CSV to a new workbook, save as CSV with correct name, close, clear CSVData and repeat
    Sheets("CSVData").Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        Path & "\" & Asset & "_" & Segment & "_" & PHDTagName & ".csv" _
        , FileFormat:=xlCSVMSDOS, CreateBackup:=False
    ActiveWindow.Close savechanges:=False
    Sheets("CSVData").Cells.Select
    Selection.ClearContents
    Range("A1").Select
    
    'Clear Arr Data
    Erase Arr
    
End Sub
'**************************************************************
Public Function isArrayEmpty(parArray As Variant) As Boolean
'Returns False if not an array or a dynamic array
'that hasn't been initialised (ReDim) or
'deleted (Erase).

If IsArray(parArray) = False Then isArrayEmpty = True
On Error Resume Next
If UBound(parArray) < LBound(parArray) Then
   isArrayEmpty = True
   Exit Function
Else
   isArrayEmpty = False
End If

End Function
'**************************************************************
Private Function getDataFromFile(parFileName As String, _
parDelimiter As String, _
Optional parExcludeCharacter As String = "") As Variant
'parFileName is the delimited file (csv, txt ...)
'parDelimiter is the separator, e.g. semicolon.
'The function returns an empty array, if the file
'is empty or cannot be opened.
'Number of columns is based on the line with most
'columns and not the first line.
'parExcludeCharacter: Some csv files have strings in
'quotations marks ("ABC"), and if parExcludeCharacter = """"
'quotation marks are removed.

Dim locLinesList() As Variant 'Array
Dim locData As Variant        'Array
Dim i As Long                 'Counter
Dim j As Long                 'Counter
Dim locNumRows As Long        'Nb of rows
Dim locNumCols As Long        'Nb of columns
Dim fso As Variant            'File system object
Dim ts As Variant             'File variable
Const REDIM_STEP = 10000      'Constant

'If this fails you need to reference Microsoft Scripting Runtime.
'You select this in "Tools" (VBA editor menu).
Set fso = CreateObject("Scripting.FileSystemObject")

On Error GoTo error_open_file
'Sets ts = the file
Set ts = fso.OpenTextFile(parFileName)
On Error GoTo unhandled_error

'Initialise the array
ReDim locLinesList(1 To 1) As Variant
i = 0
'Loops through the file, counts the number of lines (rows)
'and finds the highest number of columns.
Do While Not ts.AtEndOfStream
  'If the row number Mod 10000 = 0
  'we redimension the array.
  If i Mod REDIM_STEP = 0 Then
    ReDim Preserve locLinesList _
    (1 To UBound(locLinesList, 1) + REDIM_STEP) As Variant
  End If
  locLinesList(i + 1) = Split(ts.ReadLine, parDelimiter)
  j = UBound(locLinesList(i + 1), 1) 'Nb of columns in present row
  'If the number of columns is then highest so far.
  'the new number is saved.
  If locNumCols < j Then locNumCols = j
  i = i + 1
Loop

ts.Close 'Close file

locNumRows = i

'If number of rows is zero
If locNumRows = 0 Then Exit Function

ReDim locData(1 To locNumRows, 1 To locNumCols + 1) As Variant

'Copies the file values into an array.
'If parExcludeCharacter has a value,
'the characters are removed.
If parExcludeCharacter <> "" Then
  For i = 1 To locNumRows
    For j = 0 To UBound(locLinesList(i), 1)
      If Left(locLinesList(i)(j), 1) = parExcludeCharacter Then
        If Right(locLinesList(i)(j), 1) = parExcludeCharacter Then
          locLinesList(i)(j) = _
          Mid(locLinesList(i)(j), 2, Len(locLinesList(i)(j)) - 2)
        Else
          locLinesList(i)(j) = _
          Right(locLinesList(i)(j), Len(locLinesList(i)(j)) - 1)
        End If
      ElseIf Right(locLinesList(i)(j), 1) = parExcludeCharacter Then
        locLinesList(i)(j) = _
        Left(locLinesList(i)(j), Len(locLinesList(i)(j)) - 1)
      End If
      locData(i, j + 1) = locLinesList(i)(j)
    Next j
  Next i
Else
  For i = 1 To locNumRows
    For j = 0 To UBound(locLinesList(i), 1)
      locData(i, j + 1) = locLinesList(i)(j)
    Next j
  Next i
End If

getDataFromFile = locData

'ActiveWorkbook.Save

Erase locLinesList
Erase locData

ts.Close 'Close file
fso.Close 'Close file

'help memory to be freed
Set fso = Nothing
Set ts = Nothing

Exit Function

error_open_file:  'Returns empty Variant
unhandled_error:  'Returns empty Variant

End Function
Private Sub CreateSheet(SheetName As String)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "CSVData"
    End With
End Sub
