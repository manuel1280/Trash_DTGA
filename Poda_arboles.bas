Attribute VB_Name = "Poda_arboles"
Option Explicit
Sub Poda()

' Poda
'
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim q As Integer
    Dim j As Integer
    Dim p As Integer
    Dim s As Integer
    Dim Y As Integer
    Dim x As Integer
        
    Dim verdad As Integer
    Dim falsos As Integer
    
    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_Poda"
         
    Sheets("Tabla_Poda").Range("C1") = "Hora"
    Sheets("Tabla_Poda").Range("C1").Font.Bold = True
    Sheets("Tabla_Poda").Range("C1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Poda").Range("C1").VerticalAlignment = xlCenter
    
    Sheets("Tabla_Poda").Range("D1") = "Fecha"
    Sheets("Tabla_Poda").Range("D1").Font.Bold = True
    Sheets("Tabla_Poda").Range("D1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Poda").Range("D1").VerticalAlignment = xlCenter
    Sheets("Tabla_Poda").Range("D1").WrapText = True
    
    Sheets("Tabla_Poda").Range("E1") = "Dirección del individuo arbóreo"
    Sheets("Tabla_Poda").Range("E1").Font.Bold = True
    Sheets("Tabla_Poda").Range("E1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Poda").Range("E1").VerticalAlignment = xlCenter
    
    Columns("D:D").ColumnWidth = 15.86
    Columns("C:C").ColumnWidth = 17.43
    Columns("E:E").ColumnWidth = 29.86
    
    
    x = Control_Trash.Cont_Registros("Poda_arboles", 4)

For i = 2 To Control_Trash.cant_filas("Poda_arboles") + 1
    If Sheets("Poda_arboles").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("Poda_arboles").Cells(i, 4).Text = Trash.ListBox1.List(Y) Then
         
        'HORA
        Sheets("Tabla_Poda").Range("C2") = Sheets("Poda_arboles").Cells(i, 3)
        Sheets("Tabla_Poda").Range("C2").Font.Bold = False
        Sheets("Tabla_Poda").Range("C2").HorizontalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("C2").VerticalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("C2").WrapText = True
        Range("C2").NumberFormat = "[$-x-systime]h:mm AM/PM"
        'FECHA
        Sheets("Tabla_Poda").Range("D2") = Sheets("Poda_arboles").Cells(i, 4)
        Sheets("Tabla_Poda").Range("D2").Font.Bold = False
        Sheets("Tabla_Poda").Range("D2").HorizontalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("D2").VerticalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("D2").WrapText = True
        Range("D2").NumberFormat = "m/d/yyyy"
        'DIRECCIÓN
        Sheets("Tabla_Poda").Range("E2") = Sheets("Poda_arboles").Cells(i, 5)
        Sheets("Tabla_Poda").Range("E2").Font.Bold = False
        Sheets("Tabla_Poda").Range("E2").HorizontalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("E2").VerticalAlignment = xlCenter
        Sheets("Tabla_Poda").Range("E2").WrapText = True
        
         Range("C2:E2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
        
    Sheets("Tabla_Poda").Range("C3") = "Observaciones"
    Sheets("Tabla_Poda").Range("C3").Font.Bold = True
    Sheets("Tabla_Poda").Range("C3").HorizontalAlignment = xlCenter
    Sheets("Tabla_Poda").Range("C3").VerticalAlignment = xlCenter
        
        'OBSERVACIONES
        Sheets("Tabla_Poda").Range("C4") = Sheets("Poda_arboles").Cells(i, 6)
        Sheets("Tabla_Poda").Range("C4").Font.Bold = False
        Sheets("Tabla_Poda").Range("C4").HorizontalAlignment = xlLeft
        Sheets("Tabla_Poda").Range("C4").VerticalAlignment = xlTop
        Sheets("Tabla_Poda").Range("C4").WrapText = True

        'VERIFICACIONES**********************
    Sheets("Poda_arboles").Select
    Range(Cells(1, 7), Cells(1, 11)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Tabla_Poda").Select
    Range("C5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlTop
    Selection.WrapText = True
    Selection.Interior.Pattern = xlNone 'sin color
    
    Range("C5:C9").Select
    Selection.Interior.Pattern = xlNone 'sin color de fondo
        With Selection.Font
        .ColorIndex = xlAutomatic 'letra negra
        .TintAndShade = 0
    End With '*******************************
    
    
    'OBSERVACIONES DE VERIFICACIONES*************************
    Sheets("Poda_arboles").Select
    Range(Cells(i, 12), Cells(i, 16)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Tabla_Poda").Select
    Range("D5").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Selection.WrapText = True
    Selection.HorizontalAlignment = xlLeft
    Selection.VerticalAlignment = xlTop
     
    Range("D5:D9").Select
    Selection.Interior.Pattern = xlNone 'sin color de fondo***
    
        'COMPROBACIÓN DE NORMATIVA
    If Sheets("Poda_arboles").Cells(i, 7) = 2 Then
    Sheets("Tabla_Poda").Range("D5") = Sheets("Tabla_Poda").Range("D5") + ", presuntamente incumpliendo el artículo 2.3.2.2.2.6.71."
    End If
   
    If Sheets("Poda_arboles").Cells(i, 8) = 2 Then
    Sheets("Tabla_Poda").Range("D6") = Sheets("Tabla_Poda").Range("D6") + ", presuntamente incumpliendo el artículo 2.3.2.2.2.6.71."
    End If
    
    If Sheets("Poda_arboles").Cells(i, 9) = 2 Then
    Sheets("Tabla_Poda").Range("D7") = Sheets("Tabla_Poda").Range("D7") + ", presuntamente incumpliendo el artículo 2.3.2.2.2.6.71."
    End If
    
    If Sheets("Poda_arboles").Cells(i, 10) = 2 Then
    Sheets("Tabla_Poda").Range("D8") = Sheets("Tabla_Poda").Range("D8") + ", presuntamente incumpliendo el artículo 2.3.2.2.2.6.72."
    End If
    
    Rows("2:9").EntireRow.AutoFit
    
     Sheets("Tabla_Poda").Range("C4:E4").MergeCells = True
     Sheets("Tabla_Poda").Range("C3:E3").MergeCells = True
     Sheets("Tabla_Poda").Range("D5:E5").MergeCells = True
     Sheets("Tabla_Poda").Range("D6:E6").MergeCells = True
     Sheets("Tabla_Poda").Range("D7:E7").MergeCells = True
     Sheets("Tabla_Poda").Range("D8:E8").MergeCells = True
     Sheets("Tabla_Poda").Range("D9:E9").MergeCells = True

    
    'FORMATO TABLA
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_Poda").Range("C1:E9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 8
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
    End If
    
    End If
    Next Y
    End If
Next i
    Sheets("Tabla_Poda").Range("A1").Select
    Application.ScreenUpdating = True
    
End Sub
