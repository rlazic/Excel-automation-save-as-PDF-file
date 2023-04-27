Sub publish()
    
'Define folder location of the saved file and which cell will be used to name the saved file
Dim cell_1 As Range
Set cell_1 = Range("B9")
Dim name As String
    name = "C:\Users\Korisnik\Desktop\Mare Charter\Fakture\" & cell_1

'Define range of cells which will be published in the file
'Code includes a pivot table, as well as a range of cells before and after this pivot table
Dim PT As PivotTable
Dim rng1 As Range
Dim rng2 As Range

Set PT = ActiveSheet.PivotTables("PivotTable2")
Set rng1 = ActiveSheet.Range("A1:K19", "A58:K73")

'Join the selected ranges and the pivot table
'Define properties of the saved file
    Application.Union(rng1, PT.TableRange2).Select
    ActiveSheet.DisplayPageBreaks = False
    Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        name, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        

End Sub



