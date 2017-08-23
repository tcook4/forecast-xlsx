Attribute VB_Name = "CreateTemplates"
Sub Create_Forms()
'Create template forms based on forecast spreadsheet for each regional manager.
'These templates can be imported with the ImportData module contained in this workbook

On Error GoTo ErrHandler

'We don't need to update the screen
Application.ScreenUpdating = False

'We don't care about overwriting old saves
Application.DisplayAlerts = False

'Variable declarations
Dim SourceWorksheet As Worksheet                'Source worksheet from which we create our templates
Dim Month1, Month2, Month3, Quarter As String   'Month and quarter storage
Dim DataArray As Variant                        'Array to store our values passed to template worksheet
Dim RegionNames() As String                     'List of region names
Dim TemplateNames() As String                   'Array of all the templates to create
Dim FirstTime As Boolean: FirstTime = False     'Inform user of template directory location

'Initialization
Set SourceWorksheet = Sheets("data")
RegionNames = Split("Example,Central,East,West,Inside Sales,EMEA,Renewal,Fed", ",")
TemplateNames = Split("C:\Weekly Forecast Templates\Example.xlsx,C:\Weekly Forecast Templates\Central.xlsx,C:\Weekly Forecast Templates\East.xlsx," & _
"C:\Weekly Forecast Templates\West.xlsx,C:\Weekly Forecast Templates\Inside.xlsx,C:\Weekly Forecast Templates\EMEA.xlsx,C:\Weekly Forecast Templates\Renewal.xlsx," & _
"C:\Weekly Forecast Templates\Federal.xlsx", ",")

'Make sure our directory exists. If not, create it and tell user where it is
If Dir("C:\Weekly Forecast Templates", vbDirectory) = "" Then
    MkDir "C:\Weekly Forecast Templates"
    FirstTime = True
End If

'Fill our months and quarter based on the date
If Month(Date) >= 4 And Month(Date) <= 6 Then
    Quarter = "Q1"
    Month1 = "April"
    Month2 = "May"
    Month3 = "June"
ElseIf Month(Date) >= 7 And Month(Date) <= 9 Then
    Quarter = "Q2"
    Month1 = "July"
    Month2 = "August"
    Month3 = "September"
ElseIf Month(Date) >= 10 And Month(Date) <= 12 Then
    Quarter = "Q3"
    Month1 = "October"
    Month2 = "November"
    Month3 = "December"
ElseIf Month(Date) >= 1 And Month(Date) <= 3 Then
    Quarter = "Q4"
    Month1 = "January"
    Month2 = "February"
    Month3 = "March"
End If

'Create our 8 new workbooks
For i = 0 To 7

    'Grab the values out of our data sheet
    DataArray = Get_Data(RegionNames(i))

    Set NewBook = Workbooks.Add
    With NewBook
    
        'Write informaiton
        With .Sheets(1)
            'Write categories
            .Cells(1, 1) = RegionNames(i)
            .Cells(2, 1) = Quarter
            .Cells(3, 1) = Date
            .Cells(2, 3) = "Won"
            .Cells(3, 3) = "Most Likely"
            .Cells(4, 3) = "Upside"
            .Cells(1, 4) = Month1
            .Cells(1, 5) = Month2
            .Cells(1, 6) = Month3
            .Cells(6, 3) = "Next Quarter"
            .Cells(8, 3) = "Major Deals"
            
            'Write values
            .Cells(2, 4) = DataArray(0)
            .Cells(3, 4) = DataArray(1)
            .Cells(4, 4) = DataArray(2)
            .Cells(2, 5) = DataArray(3)
            .Cells(3, 5) = DataArray(4)
            .Cells(4, 5) = DataArray(5)
            .Cells(2, 6) = DataArray(6)
            .Cells(3, 6) = DataArray(7)
            .Cells(4, 6) = DataArray(8)
            
            'Change some formatting
            .Range("A1:A3").Font.Bold = True
            .Range("C2:C8").Font.Bold = True
            .Range("D1:F1").Font.Bold = True
            .Range("D2:F4").NumberFormat = "$#,##0_);($#,##0)"
            .Columns("A:F").AutoFit
            
        End With
        'Save, exit, and move on to the next
        .SaveAs Filename:=TemplateNames(i)
        .Close
    End With

Next

'All done - inform user
If FirstTime Then
    MsgBox ("Templates created!" & vbCrLf & "Located in C:\Weekly Forecast Templates")
Else
    MsgBox ("Templates created!")
End If

'Exit and "error handling"
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "Something broke, sorry!" & vbCrLf & Err.Description
    Resume ExitSub
    Resume

End Sub

Function Get_Data(Region As String) As Variant
'Extracts data from the 'data' worksheet

    Dim ws As Worksheet             'Worksheet 'data' we'll pull from
    Dim DataStorage(9) As Variant   'Data array to store opportunity amounts
    Dim MaxRow As Long              'Maximum row to look until

    'Initialization
    Set ws = Worksheets("data")
    MaxRow = Worksheets("data").Cells(Rows.Count, "A").End(xlUp).Row

    'Parse our data worksheet and grab opportunities related to this account
    For i = 1 To MaxRow
        With Worksheets("data")
            If .Cells(i, 1) = Region Then
                If Month(.Cells(i, 6)) = 7 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            DataStorage(0) = DataStorage(0) + .Cells(i, 5)
                        Case "2. Commit"
                            DataStorage(1) = DataStorage(1) + .Cells(i, 5)
                        Case "3. Most Likely"
                            DataStorage(1) = DataStorage(1) + .Cells(i, 5)
                        Case "4. Upside"
                            DataStorage(2) = DataStorage(2) + .Cells(i, 5)
                    End Select
                ElseIf Month(.Cells(i, 6)) = 8 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            DataStorage(3) = DataStorage(3) + .Cells(i, 5)
                        Case "2. Commit"
                            DataStorage(4) = DataStorage(4) + .Cells(i, 5)
                        Case "3. Most Likely"
                            DataStorage(4) = DataStorage(4) + .Cells(i, 5)
                        Case "4. Upside"
                            DataStorage(5) = DataStorage(5) + .Cells(i, 5)
                    End Select
                ElseIf Month(.Cells(i, 6)) = 9 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            DataStorage(6) = DataStorage(6) + .Cells(i, 5)
                        Case "2. Commit"
                            DataStorage(7) = DataStorage(7) + .Cells(i, 5)
                        Case "3. Most Likely"
                            DataStorage(7) = DataStorage(7) + .Cells(i, 5)
                        Case "4. Upside"
                            DataStorage(8) = DataStorage(8) + .Cells(i, 5)
                    End Select
                End If
            End If
        End With
    Next
    
    'Return our data
    Get_Data = DataStorage

End Function










