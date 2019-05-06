Attribute VB_Name = "Module2"
Sub CombinedAGTI()

    'Import instruction macro from CDR NAS export
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("MACRO")
    myFile = Application.GetOpenFilename(, , "Browse for Workbook")
    ThisWorkbook.Sheets("MACRO").Range("a1") = myFile
    Set wbO = Workbooks.Open(myFile)
    wbO.Sheets("AGTI Instructions").Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False


    'Clean up data to remove unnecessary columns and spaces from account numbers
    Worksheets("MACRO").Select
    Rows("1:12").Select
    Selection.Delete Shift:=xlUp
    Columns("H:O").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Range("H1").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    Range("A1:G6000").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$G$6000").AutoFilter Field:=3, Criteria1:="<>*nor*", _
        Operator:=xlAnd
    Range("A2:G6000").Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$G$6000").AutoFilter Field:=3
    Columns("A:A").Select
    Selection.Replace What:=":*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select
    Worksheets("WELCOME").Activate
    MsgBox "AGTI INSTRUCTION MACRO IMPORT SUCCESSFUL"
    
End Sub
Sub CombinedTXT()
 
    'Import .txt file from CDR WEB
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("WEB")
    myFile = Application.GetOpenFilename(, , "Browse for Workbook")
    ThisWorkbook.Sheets("WEB").Range("a1") = myFile
    Set wbO = Workbooks.Open(myFile)
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
     
    
    'Clean up data to remove unnecessary columns and spaces from account numbers
    Worksheets("WEB").Activate
    Columns("A:W").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Columns.AutoFit
    Columns("M:AB").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1:L6000").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$L$6000").AutoFilter Field:=8, Criteria1:="=0", _
        Operator:=xlOr, Criteria2:="="
    ActiveWindow.SmallScroll Down:=-9
    Range("A2:L6000").Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$L$858").AutoFilter Field:=8
    Range("A2").Select
    Worksheets("WELCOME").Activate
    Sheets("MACRO").Select
    Columns("D:D").Select
    Selection.Replace What:=" *", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Sheets("WEB").Select
    Columns("A:A").Select
    Selection.Replace What:=" *", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'Runs vlookup to cross reference MACRO against WEB for .txt specific files
    Worksheets("HIDDEN").Activate
    Range("A1:D2").Select
    Selection.Copy
    Sheets("MACRO").Select
    Range("H1").Select
    ActiveSheet.Paste
    Columns("H:K").Select
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],WEB!C[-7]:C[4],12,FALSE)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],WEB!C[-8]:C[1],8,FALSE)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-3]"
    Range("H2:J2").Select
    Selection.AutoFill Destination:=Range("H2:J6000")
    Range("A1").Select
    
    
    'Runs vlookup to cross reference WEB against MACRO for .txt specific files
    Worksheets("HIDDEN").Activate
    Range("A13:D14").Select
    Selection.Copy
    Sheets("WEB").Select
    Range("N1").Select
    ActiveSheet.Paste
    Range("N2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],MACRO!C[-10]:C[-8],3,FALSE)"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-14],MACRO!C[-11]:C[-10],2,FALSE)"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-8]"
    Range("N2:P2").Select
    Selection.AutoFill Destination:=Range("N2:P6000"), Type:=xlFillDefault
    Range("N2:P6000").Select
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    
    
    'This portion will clean up the #N/A values from WEB
    Sheets("WEB").Activate
    Range("A1:O6000").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Sheets("WELCOME").Select
    
    'This portion will clean up the #N/A values from MACRO
    Sheets("MACRO").Activate
    Range("A1:J6000").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Sheets("WELCOME").Select
    
    'Apply filter to all data at the end
    Sheets("MACRO").Select
    Range("A1:K1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Sheets("WEB").Select
    Range("A1:P1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Sheets("HIDDEN").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("WELCOME").Activate
    
    
    'Close the userform
    MsgBox "COMPLETE! NOW RESEARCH THE DIFFERENCES AND BALANCE THE TOTALS"
    Unload AGTI
    
End Sub
Sub CombinedXLS()
    
    'Import .xls file from CDR WEB
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("WEB")
    myFile = Application.GetOpenFilename(, , "Browse for Workbook")
    ThisWorkbook.Sheets("WEB").Range("a1") = myFile
    Set wbO = Workbooks.Open(myFile)
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    Worksheets("WELCOME").Activate
    
    
    'Clean up data to remove unnecessary columns and spaces from account numbers
    Worksheets("WEB").Activate
    Columns("A:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("MACRO").Select
    Columns("D:D").Select
    Selection.Replace What:=" *", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Sheets("WEB").Select
    Columns("A:A").Select
    Selection.Replace What:=" *", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        
    'Runs vlookup to cross reference WEB against MACRO for .xls specific files
    Sheets("HIDDEN").Select
    Range("A7:D8").Select
    Selection.Copy
    Sheets("MACRO").Select
    Range("H1").Select
    ActiveSheet.Paste
    Columns("H:K").Select
    Columns("H:K").EntireColumn.AutoFit
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],WEB!C[-7]:C[-2],6,FALSE)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],WEB!C[-8]:C[-6],3,FALSE)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-3]"
    Range("H2:J2").Select
    Selection.AutoFill Destination:=Range("H2:J6000"), Type:=xlFillDefault
    Range("H2:J6000").Select
    Range("A1:J6000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    ActiveWindow.SmallScroll Down:=-27
    Sheets("HIDDEN").Select
    Range("A19:D20").Select
    Selection.Copy
    Sheets("WEB").Select
    Range("G1").Select
    ActiveSheet.Paste
    Columns("G:K").Select
    Columns("G:K").EntireColumn.AutoFit
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],MACRO!C[-3]:C[-1],3,FALSE)"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],MACRO!C[-4]:C[-3],2,FALSE)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-6]"
    Range("G2:I2").Select
    Selection.AutoFill Destination:=Range("G2:I6000"), Type:=xlFillDefault
    Range("G2:I6000").Select
    ActiveWindow.SmallScroll Down:=-57
    Range("A1:I6000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Sheets("MACRO").Select
    
    
    'Final filter
    Range("A1:K1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Sheets("WEB").Select
    Range("A1:J1").Select
    Selection.AutoFilter
    Sheets("HIDDEN").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("WELCOME").Activate
    
    'Close the userform
    Unload AGTI
    MsgBox "COMPLETE! NOW RESEARCH THE DIFFERENCES AND BALANCE THE TOTALS"
    
    
 End Sub


