Attribute VB_Name = "ImportData"
Sub Import_Data()
'On Error GoTo ErrHandler

'Disable screen updating
Application.ScreenUpdating = False

Dim Target_Workbook As Workbook
Dim Target_Path As String
Dim Empty_Row As Integer
Dim target_data As String
Dim iterator As Integer
Dim deals As String
Dim dont_add As Boolean: dont_add = False
Dim Duplicate_Count As Integer: Duplicate_Count = 0

Dim data_arr(1 To 12)
Dim table_arr() As String
table_arr = Split("Example,Central,East,West,Inside,EMEA,Renewal,Federal", ",")

Dim target_worksheet As Worksheet
Set target_worksheet = ThisWorkbook.Worksheets("Example History")

Dim Forecast_Sheets() As String
Forecast_Sheets = Split("C:\Weekly Forecast\Example.xlsx,C:\Weekly Forecast\Central.xlsx,C:\Weekly Forecast\East.xlsx,C:\Weekly Forecast\West.xlsx," & _
"C:\Weekly Forecast\Inside.xlsx,C:\Weekly Forecast\EMEA.xlsx,C:\Weekly Forecast\Renewal.xlsx,C:\Weekly Forecast\Federal.xlsx", ",")

'make sure we have a directory to look in
If Dir("C:\Weekly Forecast", vbDirectory) = "" Then
    MsgBox "Please ensure forecasts are in the directory 'C:\Weekly Forecast\'"
    Exit Sub
End If




'For each sheet in sheets, import the data
For i = 0 To 7

    'Assign the Workbook File Name along with its Path
    Target_Path = Forecast_Sheets(i)

    Set Target_Workbook = Workbooks.Open(Target_Path)

    iterator = target_worksheet.ListObjects.Item(table_arr(i)).Range.Rows.Count
        
    'check if table is already empty or we already added that date
    Set result = target_worksheet.ListObjects.Item(table_arr(i)).DataBodyRange
    If result Is Nothing Then
        dont_add = False
    ElseIf Target_Workbook.Sheets(1).Cells(3, 1) = target_worksheet.ListObjects.Item(table_arr(i)).DataBodyRange(iterator - 1, 1) Then
        dont_add = True
        Duplicate_Count = Duplicate_Count + 1
    Else
        dont_add = False
    End If

    If dont_add = False Then
        'grab date
        data_arr(1) = Target_Workbook.Sheets(1).Cells(3, 1)
        
        'forecast numbers
        data_arr(2) = Target_Workbook.Sheets(1).Cells(2, 4)
        data_arr(3) = Target_Workbook.Sheets(1).Cells(3, 4)
        data_arr(4) = Target_Workbook.Sheets(1).Cells(4, 4)
        data_arr(5) = Target_Workbook.Sheets(1).Cells(2, 5)
        data_arr(6) = Target_Workbook.Sheets(1).Cells(3, 5)
        data_arr(7) = Target_Workbook.Sheets(1).Cells(4, 5)
        data_arr(8) = Target_Workbook.Sheets(1).Cells(2, 6)
        data_arr(9) = Target_Workbook.Sheets(1).Cells(3, 6)
        data_arr(10) = Target_Workbook.Sheets(1).Cells(4, 6)
        
        'Next Quarter
        data_arr(11) = Target_Workbook.Sheets(1).Cells(6, 4)
        
        'Major Deals
        iterator = 8
        Do Until IsEmpty(Target_Workbook.Sheets(1).Cells(iterator, 4).Value)
            If iterator = 8 Then
                deals = Target_Workbook.Sheets(1).Cells(iterator, 4)
            Else
                deals = deals & ", " & Target_Workbook.Sheets(1).Cells(iterator, 4)
            End If
            iterator = iterator + 1
        Loop
        data_arr(12) = deals
        
        Dim temp_str As String
        temp_str = table_arr(i)
        
        'Pass this to add data row to table
        AddDataRow temp_str, data_arr, target_worksheet
    End If
    
    'Close workbook
    Target_Workbook.Save
    Target_Workbook.Close False
    
Next i

'Enable screen updating
Application.ScreenUpdating = True

'Process Completed
MsgBox "Import Complete" & vbCrLf & "Skipped " & Duplicate_Count & " duplicate records"

ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Something broke, sorry!" & vbCrLf & Err.Description
    Resume ExitSub
    Resume

End Sub

Sub AddDataRow(tableName As String, values() As Variant, this_sheet As Worksheet)
    'add a new data row to a table
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim col As Integer
    Dim lastRow As Range

    Set sheet = this_sheet
    Set table = sheet.ListObjects.Item(tableName)
    

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

    'Iterate through the last row and populate it with the entries from values()
    Set lastRow = table.ListRows(table.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count - 2
        If col <= UBound(values) + 1 Then lastRow.Cells(1, col) = values(col)
    Next col
    
End Sub

