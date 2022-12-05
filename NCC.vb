'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'*
'* ACCOUNT: GLOBAL
'* Standardize data imput for NCC tickets
'*
'*
Sub Agent(Optional HideMe As String)
'
' Agent Macro
'

'

'Desprotege hoja 2 y hoja 1
    Application.ScreenUpdating = False
    Sheets("LIST").Select
    ActiveSheet.Unprotect "NoEdit"
    Sheets("Agent").Select
    ActiveSheet.Unprotect "NoEdit"
    Range("2:2").Font.Color = vbBlack
'Verifica input de Account, Channel y Action (check empty)
    If Range("SERVICE") = "" Then
            Range("SERVICE").Interior.Color = vbYellow
            ActiveSheet.Protect "NoEdit"
            Sheets("LIST").Select
            ActiveSheet.Protect "NoEdit"
            Sheets("Agent").Select
            Range("SERVICE").Select
            MsgBox "Enter service name", vbExclamation + vbOKOnly, "MISSING DATA"
        ElseIf Range("TICKET") = "" Then
            Range("SERVICE").Interior.Color = xlNone
            Range("TICKET").Interior.Color = vbYellow
            ActiveSheet.Protect "NoEdit"
            Sheets("LIST").Select
            ActiveSheet.Protect "NoEdit"
            Sheets("Agent").Select
            Range("TICKET").Select
            MsgBox "Enter ticket number", vbExclamation + vbOKOnly, "MISSING DATA"
        ElseIf Range("CONTACT") = "" Then
            Range("SERVICE").Interior.Color = xlNone
            Range("TICKET").Interior.Color = xlNone
            Range("CONTACT").Interior.Color = vbYellow
            ActiveSheet.Protect "NoEdit"
            Sheets("LIST").Select
            ActiveSheet.Protect "NoEdit"
            Sheets("Agent").Select
            Range("CONTACT").Select
            MsgBox "Enter contact type", vbExclamation + vbOKOnly, "MISSING DATA"
        ElseIf Range("ACTION") = "" Then
            Range("SERVICE").Interior.Color = xlNone
            Range("TICKET").Interior.Color = xlNone
            Range("CONTACT").Interior.Color = xlNone
            Range("ACTION").Interior.Color = vbYellow
            ActiveSheet.Protect "NoEdit"
            Sheets("LIST").Select
            ActiveSheet.Protect "NoEdit"
            Sheets("Agent").Select
            Range("ACTION").Select
            MsgBox "Enter action performed", vbExclamation + vbOKOnly, "MISSING DATA"
'Verifica input de Account, Channel y Action (check allowed text)
    ElseIf Range("D3") = 0 Then
        ActiveSheet.Protect "NoEdit"
        Sheets("LIST").Select
        ActiveSheet.Protect "NoEdit"
        Sheets("Agent").Select
        Range("SERVICE").Select
        MsgBox "Datos no permitidos", vbCritical + vbOKOnly, "ERROR"
    Else
        Range("SERVICE").Select
        Range("SERVICE:CONTACT").Interior.Color = xlNone
'Agrega Fecha
    Range("A2") = Date
    Range("ACTION:G2").Select
    'Cells.Replace "", "N/A", xlWhole [DEPRECATED]
'Copia/Pega en el listado
    Range("2:2").Select
    Selection.Copy
    Sheets("LIST").Select
    ActiveSheet.Unprotect "NoEdit"
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
    ActiveSheet.Protect "NoEdit"
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
    Range("STOREDNAME").Select
    Selection.Copy
    Range("NAME").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("SERVICE").Select
    ActiveCell.Formula2R1C1 = _
        "=INDEX(LIST!RC:R[98]C,MODE(MATCH(LIST!RC:R[98]C,LIST!RC:R[98]C,0)))"
    Range("SERVICE").Copy
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    ActiveSheet.Protect "NoEdit"
        
    End If
    

End Sub

--------------------------------------------

WORKBOOK EVENT

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Sheets("LIST").Protect "NoEdit"
    Sheets("AGENT").Protect "NoEdit"

End Sub

Private Sub Workbook_Open()

    Dim Services As Integer
    Dim Contacts As Integer
    Dim Actions As Integer
    
    On Error GoTo ErrorShortCut
    
    Application.ScreenUpdating = False
    Worksheets("AGENT").Unprotect "NoEdit"
    Worksheets("IMPORT").Visible = xlSheetVisible
    Sheets("IMPORT").Select
    Range("F25").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("H6:J55").Copy
    Sheets("Agent").Select
    Range("H6:J55").Select
    ActiveSheet.Paste
        
    Services = Worksheets("Agent").Range("H7:H57").Cells.SpecialCells(xlCellTypeConstants).Count + 6
    Contacts = Worksheets("Agent").Range("I7:I57").Cells.SpecialCells(xlCellTypeConstants).Count + 6
    Actions = Worksheets("Agent").Range("J7:J57").Cells.SpecialCells(xlCellTypeConstants).Count + 6
    Range("Tabla1[[#Headers],[SERVICE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="SERVICELIST", RefersToR1C1:= _
        "=Agent!R6C8:R" & Services & "C8"
    Range("Tabla1[[#Headers],[TYPE OF CONTACT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="CONTACTLIST", RefersToR1C1:= _
        "=Agent!R6C9:R" & Contacts & "C9"
    Range("Tabla1[[#Headers],[ACTION]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="ACTIONLIST", RefersToR1C1:= _
        "=Agent!R6C10:R" & Actions & "C10"
    
    Worksheets("RESTORE").Visible = xlSheetVisible
    Sheets("RESTORE").Select
    Range("C2:G3").Copy
    Sheets("AGENT").Select
    Range("C2:G3").Select
    ActiveSheet.Paste
    Range("SERVICE").Select
    ActiveCell.Formula2R1C1 = _
        "=INDEX(LIST!RC:R[98]C,MODE(MATCH(LIST!RC:R[98]C,LIST!RC:R[98]C,0)))"
    Range("SERVICE").Copy
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ActiveWorkbook.Names.Add Name:="SERVICE", RefersToR1C1:= _
        "=Agent!R2C3"
    ActiveWorkbook.Names.Add Name:="TICKET", RefersToR1C1:= _
        "=Agent!R2C4"
    ActiveWorkbook.Names.Add Name:="CONTACT", RefersToR1C1:= _
        "=Agent!R2C5"
    ActiveWorkbook.Names.Add Name:="ACTION", RefersToR1C1:= _
        "=Agent!R2C6"
        
    Worksheets("AGENT").Protect "NoEdit"
    Worksheets("IMPORT").Visible = xlSheetHidden
    Worksheets("RESTORE").Visible = xlSheetHidden
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

ErrorShortCut:
    Worksheets("AGENT").Protect "NoEdit"
    Worksheets("IMPORT").Visible = xlSheetHidden
    Worksheets("RESTORE").Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub


