Sub TotalValues()
'
' TotalValues Macro
'

'
    Range("U1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Total Profit"
    Range("V2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-3]+RC[-2]"
    Range("U3").Select
    Selection.End(xlDown).Select
    Range("V9995").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Selection.End(xlUp).Select
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Total Cost"
    Range("W2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-5]*RC[-4]+RC[-3]+RC[-2]"
    Range("V3").Select
    Selection.End(xlDown).Select
    Range("W9995").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "Total Sales"
    Range("X2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]+RC[-2]"
    Range("W3").Select
    Selection.End(xlDown).Select
    Range("X9995").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Range("V9998").Select
    ActiveCell.FormulaR1C1 = "Total Cost Sum"
    Range("X9998").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-9996]C[-2]:R[-3]C[-2])"
    Range("V10000").Select
    ActiveCell.FormulaR1C1 = "Total Profit Sum"
    Range("X10000").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-9998]C[-1]:R[-5]C[-1])"
    Range("V10002").Select
    ActiveCell.FormulaR1C1 = "Total Sales Sum"
    Range("X10002").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-10000]C:R[-7]C)"
    Range("V10003").Select
End Sub
