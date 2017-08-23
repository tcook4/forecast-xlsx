Attribute VB_Name = "PublicFunctions"
Sub Replace_Names()
'Replace "CSE" and "UK" with EMEA for region identification

Worksheets("data").Columns("A").Replace _
What:="CSE", Replacement:="EMEA", _
SearchOrder:=xlByColumns, MatchCase:=True
 
Worksheets("data").Columns("A").Replace _
What:="UK", Replacement:="EMEA", _
SearchOrder:=xlByColumns, MatchCase:=True

MsgBox ("Replacement complete!")

End Sub
