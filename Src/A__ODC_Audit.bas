Attribute VB_Name = "A__ODC_Audit"
Function odcaudit()
Application.ScreenUpdating = False
Dim lookupws As Worksheet
Dim dataws As Worksheet
Dim lastrow As Long

Set lookupws = Sheets("Lookups")
Set dataws = Sheets("Compiled")
lastrow = dataws.Cells(dataws.Rows.Count, "A").End(xlUp).Row

Sheets("Exiles").Cells.Delete
Sheets("Quick Checks").Cells.Delete

'Delete Interfaces except ODC

dataws.Range("A:AD").Sort key1:=dataws.Range("E1"), Header:=xlYes

dataws.Range("A:AD").AdvancedFilter _
    Action:=xlFilterInPlace, _
    CriteriaRange:=lookupws.Range("A1:A2")

If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
    dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Copy
    dataws.Paste _
        Destination:=Sheets("Exiles").Range("A1")
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End If

dataws.ShowAllData

'Delete R & Y Void Types

dataws.Range("A:AD").Sort key1:=dataws.Range("AD1"), Header:=xlYes

dataws.Range("A:AD").AdvancedFilter _
    Action:=xlFilterInPlace, _
    CriteriaRange:=lookupws.Range("C1:D3")

If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Copy
    Sheets("Exiles").Range("A2").EntireRow.Insert shift:=xlDown
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End If

dataws.ShowAllData

'Delete all E#'s in Payee column
dataws.Range("A:AD").Sort key1:=dataws.Range("Q1"), Header:=xlYes

dataws.Range("A:AD").AdvancedFilter _
    Action:=xlFilterInPlace, _
    CriteriaRange:=lookupws.Range("F1:F2")

If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Copy
    Sheets("Exiles").Range("A2").EntireRow.Insert shift:=xlDown
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End If

dataws.ShowAllData

'Exile Erin Wendt, Tawny Cool, Christian Venables, Colleen Seus to Quick Check sheet
dataws.Range("A:AD").Sort key1:=dataws.Range("AC1"), Header:=xlYes

dataws.Range("A:AD").AdvancedFilter _
    Action:=xlFilterInPlace, _
    CriteriaRange:=lookupws.Range("A5:E10")

If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
    dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Copy
    dataws.Paste _
        Destination:=Sheets("Quick Checks").Range("A1")
    
    dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End If

dataws.ShowAllData

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
dataws.Range("O:O").Cells.Clear
dataws.Range("O1").Value = "Vendor #"
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Call categoryexiler
Call vendormatch

Application.ScreenUpdating = True
MsgBox ("All done")
End Function

Function lookupbuilder()
lastrow = Sheets("Lookup Builder").Cells(Sheets("Lookup Builder").Rows.Count, "A").End(xlUp).Row

For Each cell In Sheets("lookup builder").Range("A2:A" & lastrow)
    lubuild = lubuild & cell.Value & ","
Next cell

Sheets("lookup builder").Range("C3").Value = lubuild
End Function

