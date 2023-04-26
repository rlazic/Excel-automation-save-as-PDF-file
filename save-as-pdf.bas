Attribute VB_Name = "Module1"
Sub Hiding_blank_lines()
Attribute Hiding_blank_lines.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hiding_blank_lines Macro
'

'
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    ActiveSheet.Range("$l$1:$l$53").AutoFilter Field:=1, Criteria1:="<>"
    Range("A7").Select
End Sub
Sub Show_All_Lines()
Attribute Show_All_Lines.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Show_All_Lines Macro
'

'
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    ActiveSheet.Range("$l$1:$l$53").AutoFilter Field:=1
    Range("A7").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("B14").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
End Sub
Sub trash()
Attribute trash.VB_ProcData.VB_Invoke_Func = " \n14"
'
' trash Macro
'

'
    ActiveCell.FormulaR1C1 = "2"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("H8").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-3]C[-2]+R[-1]C[-5]"
    Range("C1:H13").Select
    Selection.AutoFilter
    Range("C16").Select
    ActiveSheet.Range("$C$3:$H$8").AutoFilter Field:=4, Criteria1:="<>"
    Range("A1").Select
End Sub


Sub publish()
Dim cell_1 As Range
Set cell_1 = Range("B9")
Dim name As String
    name = "C:\Users\Korisnik\Desktop\Mare Charter\Fakture\" & cell_1

Dim PT As PivotTable
Dim rng1 As Range
Dim rng2 As Range

Set PT = ActiveSheet.PivotTables("PivotTable2")
Set rng1 = ActiveSheet.Range("A1:K19", "A58:K73")

    Application.Union(rng1, PT.TableRange2).Select
    ActiveSheet.DisplayPageBreaks = False
    Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        name, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        

End Sub



