Attribute VB_Name = "Module1"
Sub 금액()
Attribute 금액.VB_Description = "부교재 p144"
Attribute 금액.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' 금액 매크로
' 부교재 p144
'
' 바로 가기 키: Ctrl+k
'
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Selection.AutoFill Destination:=Range("F4:F11"), Type:=xlFillDefault
    Range("F4:F11").Select
End Sub
Sub 테두리()
Attribute 테두리.VB_Description = "부교재 p144"
Attribute 테두리.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 테두리 매크로
' 부교재 p144
'

'
    Range("B3:F11").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
