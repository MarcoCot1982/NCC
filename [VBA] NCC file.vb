'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'*
'* ACCOUNT: GLOBAL
'* Standardize data imput for NCC tickets
'* 
'*
Sub Agent(Optional HiddenMacro as String)
'
' Agent Macro
'

'
'Desprotege hoja 2 y hoja 1
    Application.ScreenUpdating = False
    Sheets("LIST").Select
    ActiveSheet.Unprotect
    Sheets("Agent").Select
    ActiveSheet.Unprotect
'Verifica input de Account, Channel y Action (check empty)
        If Range("C2") = "" Then
        Range("C2").Select
        Range("C2").Interior.COLOR = vbYellow
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("LIST").Select
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("Agent").Select
        Range("C2").Select
    ElseIf Range("D2") = "" Then
        Range("C2").Select
        Range("C2").Interior.COLOR = xlNone
        Range("D2").Select
        Range("D2").Interior.COLOR = vbYellow
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("LIST").Select
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("Agent").Select
        Range("C2").Select
    ElseIf Range("E2") = "" Then
        Range("D2").Select
        Range("D2").Interior.COLOR = xlNone
        Range("E2").Select
        Range("E2").Interior.COLOR = vbYellow
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("LIST").Select
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("Agent").Select
        Range("C2").Select
    ElseIf Range("F2") = "" Then
        Range("E2").Select
        Range("E2").Interior.COLOR = xlNone
        Range("F2").Select
        Range("F2").Interior.COLOR = vbYellow
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("LIST").Select
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("Agent").Select
        Range("C2").Select
'Verifica input de Account, Channel y Action (check allowed text)
    ElseIf Range("D3") = 0 Then
        Range("C2:F2").Select
        Range("C2:F2").Interior.COLOR = vbRed
        Range("G2").Select
        ActiveCell.FormulaR1C1 = _
        "DATOS NO PERMITIDOS"
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("LIST").Select
        ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
        Sheets("Agent").Select
        Range("C2").Select
    Else
    Range("C2").Select
    Range("C2:E2").Interior.COLOR = xlNone
'Agrega Fecha
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("F2:G2").Select
    Selection.Replace What:="", Replacement:="N/A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DATOS NO PERMITIDOS", Replacement:="N/A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
'Copia/Pega en el listado
    Range("2:2").Select
    Selection.Copy
    Sheets("LIST").Select
    ActiveSheet.Unprotect
    Range("2:2").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'Quita bordes/restaura formato fecha
    Range("R1").Select
    Selection.Copy
    Range("2:2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Range("A2").NumberFormat = "dd/mm/yyyy;@"
    Cells.Select
'Volver a proteger y restaurar formato tabla
    Selection.Locked = True
    Selection.FormulaHidden = False
    ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("A1").Select
    Sheets("Agent").Select
    Range("A2:G2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("C2:G2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("C2:G2").Select
    Selection.Locked = False
    Range("B3").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.PROTECT DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("C2").Select
    Application.ScreenUpdating = True
    End If
End Sub
