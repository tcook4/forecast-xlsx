Attribute VB_Name = "CreateTemplates"
Sub Create_Forms()

'We don't need to update the screen
Application.ScreenUpdating = False

Dim number As Double
Dim data_ws As Worksheet
Set data_ws = Sheets("data")
Dim data_array As Variant

Dim month1, month2, month3, quarter As String

'List of filenames
Dim Names() As String
Names = Split("Example,Central,East,West,Inside Sales,EMEA,Renewal,Fed", ",")

'Our list of filenames to create
Dim Forecast_Sheets() As String
Forecast_Sheets = Split("C:\Weekly Forecast Templates\Example.xlsx,C:\Weekly Forecast Templates\Central.xlsx,C:\Weekly Forecast Templates\East.xlsx,C:\Weekly Forecast Templates\West.xlsx,C:\Weekly Forecast Templates\Inside.xlsx,C:\Weekly Forecast Templates\EMEA.xlsx,C:\Weekly Forecast Templates\Renewal.xlsx,C:\Weekly Forecast Templates\Federal.xlsx", ",")

'We don't care about overwriting old saves
Application.DisplayAlerts = False

'make sure our directory exists
If Dir("C:\Weekly Forecast Templates", vbDirectory) = "" Then
    MkDir "C:\Weekly Forecast Templates"
End If

'Fill our months and quarter based on the date
If Month(Date) >= 4 And Month(Date) <= 6 Then
    quarter = "Q1"
    month1 = "April"
    month2 = "May"
    month3 = "June"
ElseIf Month(Date) >= 7 And Month(Date) <= 9 Then
    quarter = "Q2"
    month1 = "July"
    month2 = "August"
    month3 = "September"
ElseIf Month(Date) >= 10 And Month(Date) <= 12 Then
    quarter = "Q3"
    month1 = "October"
    month2 = "November"
    month3 = "December"
ElseIf Month(Date) >= 1 And Month(Date) <= 3 Then
    quarter = "Q4"
    month1 = "January"
    month2 = "February"
    month3 = "March"
End If

'Create 8 files based on names above
For i = 0 To 7

    'Grab the values out of our data sheet
    data_array = Get_Data(Names(i))

    Set newbook = Workbooks.Add
    With newbook
    
        'Write informaiton
        With .Sheets(1)
            'write formatting
            .Cells(1, 1) = Names(i)
            .Cells(2, 1) = quarter
            .Cells(3, 1) = Date
            .Cells(2, 3) = "Won"
            .Cells(3, 3) = "Most Likely"
            .Cells(4, 3) = "Upside"
            .Cells(1, 4) = month1
            .Cells(1, 5) = month2
            .Cells(1, 6) = month3
            .Cells(6, 3) = "Next Quarter"
            .Cells(8, 3) = "Major Deals"
            
            'write values
            .Cells(2, 4) = data_array(0)
            .Cells(3, 4) = data_array(1)
            .Cells(4, 4) = data_array(2)
            .Cells(2, 5) = data_array(3)
            .Cells(3, 5) = data_array(4)
            .Cells(4, 5) = data_array(5)
            .Cells(2, 6) = data_array(6)
            .Cells(3, 6) = data_array(7)
            .Cells(4, 6) = data_array(8)
            
            'do some formatting
            .Range("A1:A3").Font.Bold = True
            .Range("C2:C8").Font.Bold = True
            .Range("D1:F1").Font.Bold = True
            .Range("D2:F4").NumberFormat = "$#,##0_);($#,##0)"
            .Columns("A:F").AutoFit
            
        End With
        'save, exit, and move on to the next
        .SaveAs Filename:=Forecast_Sheets(i)
        .Close
    End With

Next

'Let screen start updating again
Application.ScreenUpdating = True

MsgBox ("Template creation complete!")


End Sub

Function Get_Data(act As String) As Variant

    Dim ws As Worksheet
    Set ws = Worksheets("data")
    
    Dim acct As String
    acct = act
        
    'Data storage
    Dim data_arr(9) As Variant
    
    Dim max_row As Long
    max_row = Worksheets("data").Cells(Rows.Count, "A").End(xlUp).Row

    For i = 1 To max_row
        With Worksheets("data")
            If .Cells(i, 1) = acct Then
                If Month(.Cells(i, 6)) = 7 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            data_arr(0) = data_arr(0) + .Cells(i, 5)
                        Case "2. Commit"
                            data_arr(1) = data_arr(1) + .Cells(i, 5)
                        Case "3. Most Likely"
                            data_arr(1) = data_arr(1) + .Cells(i, 5)
                        Case "4. Upside"
                            data_arr(2) = data_arr(2) + .Cells(i, 5)
                    End Select
                ElseIf Month(.Cells(i, 6)) = 8 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            data_arr(3) = data_arr(3) + .Cells(i, 5)
                        Case "2. Commit"
                            data_arr(4) = data_arr(4) + .Cells(i, 5)
                        Case "3. Most Likely"
                            data_arr(4) = data_arr(4) + .Cells(i, 5)
                        Case "4. Upside"
                            data_arr(5) = data_arr(5) + .Cells(i, 5)
                    End Select
                ElseIf Month(.Cells(i, 6)) = 9 Then
                    Select Case .Cells(i, 2)
                        Case "1. Won"
                            data_arr(6) = data_arr(6) + .Cells(i, 5)
                        Case "2. Commit"
                            data_arr(7) = data_arr(7) + .Cells(i, 5)
                        Case "3. Most Likely"
                            data_arr(7) = data_arr(7) + .Cells(i, 5)
                        Case "4. Upside"
                            data_arr(8) = data_arr(8) + .Cells(i, 5)
                    End Select
                End If
            End If
        End With
    Next
    
    
    
    '0, 1, 2 are won, most likely and upside for current month
    Get_Data = data_arr

End Function

