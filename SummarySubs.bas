Attribute VB_Name = "SummarySubs"
Sub UpdateTables()
 
 
Dim pt As PivotTable
Dim Field As PivotField
Dim NewCat As String

If Sheets("Summary").Range("b1").Value = "ALL" Then

    Set pt = Worksheets("Regional").PivotTables("Regionaltable")
    Set Field = pt.PivotFields("Sub-Region")
    NewCat = Worksheets("Regional").Range("b1").Value
    
    With pt
        Field.ClearAllFilters
    End With
    
    With Sheets("Summary")
        .Range("A16:f37").AutoFilter Field:=3
        .Range("A40:f48").AutoFilter Field:=3
        .Range("A52:f59").AutoFilter Field:=3
        .Range("A66:f71").AutoFilter Field:=3
    End With

Else

    Set pt = Worksheets("Regional").PivotTables("Regionaltable")
    Set Field = pt.PivotFields("Sub-Region")
    NewCat = Worksheets("Regional").Range("b1").Value
    
    With pt
        Field.ClearAllFilters
        Field.CurrentPage = NewCat
        pt.RefreshTable
    End With
    
    With Sheets("Summary")
        .Range("A16:f37").AutoFilter Field:=3, Criteria1:=.Range("b1").Value
        .Range("A40:f48").AutoFilter Field:=3, Criteria1:=.Range("b1").Value
        .Range("A52:f59").AutoFilter Field:=3, Criteria1:=.Range("b1").Value
        .Range("A66:f71").AutoFilter Field:=3, Criteria1:=.Range("b1").Value
    End With

End If

End Sub

