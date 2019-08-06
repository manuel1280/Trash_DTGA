Attribute VB_Name = "Base_Ope"
Option Explicit
Sub Base_operaciones()

    Application.ScreenUpdating = False

    Dim i As Integer
    Dim j As Integer
    Dim Y As Integer
    Dim x As Integer
    Dim Decreto As String

    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_BO"
    
         Columns("C:C").ColumnWidth = 29.5
         Columns("D:D").ColumnWidth = 47.5
    
    x = Control_Trash.Cont_Registros("Base_Op", 4)
    
For i = 2 To Control_Trash.cant_filas("Base_Op") + 1
    If Sheets("Base_Op").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("Base_Op").Cells(i, 4).Text = Trash.ListBox1.List(Y) Then

    Sheets("Tabla_BO").Range("C2") = Sheets("Base_Op").Cells(i, 2)
    Sheets("Tabla_BO").Range("C2").Font.Bold = True
    Sheets("Tabla_BO").Range("C2").HorizontalAlignment = xlCenter
    Range("C2").Value = UCase(Range("C2")) 'convertir a mayuscula
        
           Range("C2:D2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    'FORMATO DE LETRA*****************************************************
    'UBICACIÓN
    Sheets("Tabla_BO").Range("C3") = Sheets("Base_Op").Cells(1, 5)
    Sheets("Tabla_BO").Range("C3").Font.Bold = True
    Sheets("Tabla_BO").Range("C3").HorizontalAlignment = xlLeft
    Sheets("Tabla_BO").Range("C3").VerticalAlignment = xlCenter
    
        Sheets("Tabla_BO").Range("D3") = Sheets("Base_Op").Cells(i, 5) + " y de acuerdo a lo definido en el ordenamiento territorial, " + Sheets("Base_Op").Cells(i, 6)
        Sheets("Tabla_BO").Range("D3").Font.Bold = False
        Sheets("Tabla_BO").Range("D3").HorizontalAlignment = xlLeft
        Sheets("Tabla_BO").Range("D3").VerticalAlignment = xlTop
        Sheets("Tabla_BO").Range("D3").WrapText = True
        
    Sheets("Tabla_BO").Range("C4") = "Características"
    Sheets("Tabla_BO").Range("C4").Font.Bold = True
    Sheets("Tabla_BO").Range("C4").HorizontalAlignment = xlCenter
    Sheets("Tabla_BO").Range("C4:D4").MergeCells = True
    
   'CARACTERISTICAS
    Sheets("Base_Op").Select
    Range("G1:R1").Select
    Selection.Copy
    Sheets("Tabla_BO").Select
    Range("C5:C16").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlCenter
    Selection.WrapText = True

       Range("C5:C16").Select
    Selection.Interior.Pattern = xlNone 'sin color de fondo
        With Selection.Font
        .ColorIndex = xlAutomatic 'letra negra
        .TintAndShade = 0
    End With
    
    
    'OBSERVACIONES
    Sheets("Base_Op").Select
    Range(Cells(i, 7), Cells(i, 18)).Select
    Selection.Copy
    Sheets("Tabla_BO").Select
    Range("D5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlTop
    Selection.WrapText = True
        
    'complemento selñalizaciones
    Sheets("Tabla_BO").Range("D6") = Sheets("Tabla_BO").Range("D6") + " en la base de operaciones; referente a los sentidos de circulación, " + Sheets("Base_Op").Cells(i, 20)
    'complemento frecuencia de lavado
    Sheets("Tabla_BO").Range("D15") = Sheets("Tabla_BO").Range("D15") + ", con frecuencia de lavado " + Sheets("Base_Op").Cells(i, 19)
    
    'COMPARACIÓN CON LA NORMATIVA**********************************************
        Decreto = "artículo 2.3.2.2.2.3.50"
    
         If Sheets("base_op").Cells(i, 21) = 2 Then
    Sheets("Tabla_BO").Range("D3") = Sheets("Tabla_BO").Range("D3") + ". Presuntamente incumpliendo con el " + Decreto
    End If
        If Sheets("base_op").Cells(i, 22) = 2 Then
    Sheets("Tabla_BO").Range("D5") = Sheets("Tabla_BO").Range("D5") + ". Presuntamente incumpliendo con el númeral 1 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 23) = 2 Then
    Sheets("Tabla_BO").Range("D6") = Sheets("Tabla_BO").Range("D6") + ". Presuntamente incumpliendo con el númeral 3 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 24) = 2 Then
    Sheets("Tabla_BO").Range("D7") = Sheets("Tabla_BO").Range("D7") + ". Presuntamente incumpliendo con el númeral 1 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 25) = 2 Then
    Sheets("Tabla_BO").Range("D8") = Sheets("Tabla_BO").Range("D8") + ". Presuntamente incumpliendo con el númeral 1 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 26) = 2 Then
    Sheets("Tabla_BO").Range("D9") = Sheets("Tabla_BO").Range("D9") + ". Presuntamente incumpliendo con el númeral 1 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 27) = 2 Then
    Sheets("Tabla_BO").Range("D10") = Sheets("Tabla_BO").Range("D10") + ". Presuntamente incumpliendo con el númeral 2 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 28) = 2 Then
    Sheets("Tabla_BO").Range("D11") = Sheets("Tabla_BO").Range("D11") + ". Presuntamente incumpliendo con el númeral 1 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 29) = 2 Then
    Sheets("Tabla_BO").Range("D12") = Sheets("Tabla_BO").Range("D12") + ". Presuntamente incumpliendo con el númeral 4 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 30) = 2 Then
    Sheets("Tabla_BO").Range("D13") = Sheets("Tabla_BO").Range("D13") + ". Presuntamente incumpliendo con el númeral 5 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 31) = 2 Then
    Sheets("Tabla_BO").Range("D14") = Sheets("Tabla_BO").Range("D14") + ". Presuntamente incumpliendo con el númeral 6 del " + Decreto
    End If
        If Sheets("base_op").Cells(i, 33) = 1 Then
    Sheets("Tabla_BO").Range("D16") = Sheets("Tabla_BO").Range("D16") + ". Presuntamente incumpliendo con el parágrafo 1 del " + Decreto
    End If
    
    Sheets("Tabla_BO").Range("D5:D16").Select
    Selection.Interior.Pattern = xlNone 'sin color de fondo
    
    'FORMATO DE TABLA**********************************************
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_BO").Range("C2:D16").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Rows("2:16").EntireRow.AutoFit

        Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 15
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
    End If
    
    If x > 1 Then
    Sheets("Tabla_BO").Range("C17:D17").MergeCells = True
    End If
    
    End If
    Next Y
    End If
Next i
    
    Sheets("Tabla_BO").Range("C2:D2").MergeCells = True
    Range("A1").Select
    
    
End Sub
