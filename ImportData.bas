Attribute VB_Name = "ImportData"
Sub Import_Data()
'Imports Excel Spreadsheets and populates master forecast tables with data

On Error GoTo ErrHandler

'Disable screen updating
Application.ScreenUpdating = False

Dim TargetWorkbook As Workbook              'New workbooks opened here for importing
Dim TargetWorksheet As Worksheet            'Data is imported to this worksheet
Dim WorksheetNames() As String              'Array of worksheet names to import
Dim RegionNames() As String                 'Used to find correct table to import to
Dim DataArray(1 To 12)                      'Filled and passed to create new row
Dim Iter As Integer                         'Iterator
Dim TargetPath As String                    'Filename we're trying to open
Dim NotedDeals As String                    'Notable deals managers wish to call out by name
Dim DontAdd As Boolean                      'Bool value to determine if we need to add the data
Dim DuplicateCount As Integer               'How many duplicate / redundant records we skip

'Initialization
DuplicateCount = 0
DontAdd = False
Set TargetWorksheet = ThisWorkbook.Worksheets("Example History")
RegionNames = Split("Example,Central,East,West,Inside,EMEA,Renewal,Federal", ",")
WorksheetNames = Split("C:\Weekly Forecast\Example.xlsx,C:\Weekly Forecast\Central.xlsx,C:\Weekly Forecast\East.xlsx,C:\Weekly Forecast\West.xlsx," & _
"C:\Weekly Forecast\Inside.xlsx,C:\Weekly Forecast\EMEA.xlsx,C:\Weekly Forecast\Renewal.xlsx,C:\Weekly Forecast\Federal.xlsx", ",")

'Make sure we have a directory to look in
If Dir("C:\Weekly Forecast", vbDirectory) = "" Then
    MsgBox "Please ensure forecasts are in the directory 'C:\Weekly Forecast\'"
    GoTo ExitSub
End If

'For each sheet in sheets, import the data
For i = 0 To 7

    'Assign the Workbook File Name along with its Path
    TargetPath = WorksheetNames(i)
    Set TargetWorkbook = Workbooks.Open(TargetPath)
    Iter = TargetWorksheet.ListObjects.Item(RegionNames(i)).Range.Rows.Count
        
            'Check if table is already empty or if we've already added that date
    Set Result = TargetWorksheet.ListObjects.Item(RegionNames(i)).DataBodyRange
    If Result Is Nothing Then
        DontAdd = False
    ElseIf TargetWorkbook.Sheets(1).Cells(3, 1) = TargetWorksheet.ListObjects.Item(RegionNames(i)).DataBodyRange(Iter - 1, 1) Then
        DontAdd = True
        DuplicateCount = DuplicateCount + 1
    Else
        DontAdd = False
    End If

    'Get the data out of our import workbook
    If DontAdd = False Then
        'Date
        DataArray(1) = TargetWorkbook.Sheets(1).Cells(3, 1)
        
        'Forecast numbers
        DataArray(2) = TargetWorkbook.Sheets(1).Cells(2, 4)
        DataArray(3) = TargetWorkbook.Sheets(1).Cells(3, 4)
        DataArray(4) = TargetWorkbook.Sheets(1).Cells(4, 4)
        DataArray(5) = TargetWorkbook.Sheets(1).Cells(2, 5)
        DataArray(6) = TargetWorkbook.Sheets(1).Cells(3, 5)
        DataArray(7) = TargetWorkbook.Sheets(1).Cells(4, 5)
        DataArray(8) = TargetWorkbook.Sheets(1).Cells(2, 6)
        DataArray(9) = TargetWorkbook.Sheets(1).Cells(3, 6)
        DataArray(10) = TargetWorkbook.Sheets(1).Cells(4, 6)
        
        'Next Quarter
        DataArray(11) = TargetWorkbook.Sheets(1).Cells(6, 4)
        
        'Major Deals
        Iter = 8
        Do Until IsEmpty(TargetWorkbook.Sheets(1).Cells(Iter, 4).Value)
            If Iter = 8 Then
                NotedDeals = TargetWorkbook.Sheets(1).Cells(Iter, 4)
            Else
                NotedDeals = NotedDeals & ", " & TargetWorkbook.Sheets(1).Cells(Iter, 4)
            End If
            Iter = Iter + 1
        Loop
        DataArray(12) = NotedDeals
        
        'Pass this to add data row to table
        AddDataRow RegionNames(i), DataArray, TargetWorksheet
    End If
    
    'Close workbook
    TargetWorkbook.Save
    TargetWorkbook.Close False
    
Next i

'Process Completed
MsgBox "Import Complete" & vbCrLf & "Skipped " & DuplicateCount & " duplicate records"

'Exit and "error handling"
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "Something broke, sorry!" & vbCrLf & Err.Description
    Resume ExitSub
    Resume

End Sub

Sub AddDataRow(tableName As String, Values() As Variant, ThisSheet As Worksheet)
    'A a new data row to a table
    Dim sh As Worksheet
    Dim table As ListObject
    Dim col As Integer
    Dim lastRow As Range
    
    Set sh = ThisSheet
    Set table = sh.ListObjects.Item(tableName)
    
    'First check if the last row is empty; if not, add a row
    If table.ListRows.Count > 0 Then
        Set lastRow = table.ListRows(table.ListRows.Count).Range
        For col = 1 To lastRow.Columns.Count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                table.ListRows.Add
                Exit For
            End If
        Next col
    Else
        table.ListRows.Add
    End If

    'Iterate through the rows and populate it with the entries from values()
    Set lastRow = table.ListRows(table.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count - 2
        If col <= UBound(Values) + 1 Then lastRow.Cells(1, col) = Values(col)
    Next col
    
    'Add formulas for already won and remaining
    'TODO - dynamic column names
    lastRow.Cells(1, lastRow.Columns.Count - 1).Formula = "=[@[July Won]]+[@[August Won]]+[@[September Won]]"
    lastRow.Cells(1, lastRow.Columns.Count).Formula = "=([@[July Most Likely]]+[@[July Upside]]+[@[August Most Likely]]+[@[August Upside]]+[@[September Most Likely]]+[@[September Upside]])-([@[July Won]]+[@[August Won]]+[@[September Won]])"
    
End Sub




















