Attribute VB_Name = "Module1"
Sub �ݾ�()
Attribute �ݾ�.VB_Description = "�α��� p144"
Attribute �ݾ�.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' �ݾ� ��ũ��
' �α��� p144
'
' �ٷ� ���� Ű: Ctrl+k
'
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Selection.AutoFill Destination:=Range("F4:F11"), Type:=xlFillDefault
    Range("F4:F11").Select
End Sub
Sub �׵θ�()
Attribute �׵θ�.VB_Description = "�α��� p144"
Attribute �׵θ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �׵θ� ��ũ��
' �α��� p144
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
