Attribute VB_Name = "Corte_cesped"
Option Explicit
Sub Corte_cesped()
Attribute Corte_cesped.VB_ProcData.VB_Invoke_Func = " \n14"
'
' corte_de_cesped
'
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim j As Integer
    Dim q As Integer
    Dim p As Integer
    Dim s As Integer
    Dim Y As Integer
    Dim x As Integer
    
    Dim verdad As Integer
    Dim falsos As Integer
           
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_C�sped"
        
    Sheets("Tabla_c�sped").Range("C1") = "Direcci�n del �rea intervenida"
    Sheets("Tabla_c�sped").Range("C1").Font.Bold = True
    Sheets("Tabla_c�sped").Range("C1").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("C1").VerticalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("C1").WrapText = True
    
    Sheets("Tabla_c�sped").Range("D1") = "Hora"
    Sheets("Tabla_c�sped").Range("D1").Font.Bold = True
    Sheets("Tabla_c�sped").Range("D1").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("D1").VerticalAlignment = xlCenter
    
    Sheets("Tabla_c�sped").Range("E1") = "Fecha"
    Sheets("Tabla_c�sped").Range("E1").Font.Bold = True
    Sheets("Tabla_c�sped").Range("E1").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("E1").VerticalAlignment = xlCenter
    
    Sheets("Tabla_c�sped").Range("F1") = "�rea verde intervenida de la zona verificada"
    Sheets("Tabla_c�sped").Range("F1").Font.Bold = True
    Sheets("Tabla_c�sped").Range("F1").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("F1").VerticalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("C1").WrapText = True
    
    
    Sheets("Tabla_c�sped").Range("G1") = "N�mero de operarios en cuadrilla"
    Sheets("Tabla_c�sped").Range("G1").Font.Bold = True
    Sheets("Tabla_c�sped").Range("G1").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("G1").VerticalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("C1").WrapText = True
    
    
    
    x = Control_Trash.Cont_Registros("Corte_c�sped", 5)

For i = 2 To Control_Trash.cant_filas("Corte_c�sped") + 1
    If Sheets("Corte_c�sped").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("Corte_c�sped").Cells(i, 5).Text = Trash.ListBox1.List(Y) Then

    Sheets("Corte_c�sped").Select
    Range(Cells(i, 3), Cells(i, 7)).Select
    Selection.Copy
    Sheets("Tabla_C�sped").Select
    Cells(2, 3).Select
    ActiveSheet.Paste
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    
    Sheets("Tabla_C�sped").Range("C2").WrapText = True
    Sheets("Tabla_C�sped").Range("C3") = "Verificaci�n"
    Sheets("Tabla_c�sped").Range("C3").Font.Bold = True
    Sheets("Tabla_c�sped").Range("C3").HorizontalAlignment = xlCenter
    Sheets("Tabla_c�sped").Range("C3").VerticalAlignment = xlCenter
    
    Sheets("Tabla_c�sped").Range("F2") = Sheets("Tabla_c�sped").Range("F2").Text + "m2"
    
    Sheets("Tabla_C�sped").Range("D3") = "Observaci�n"
    Sheets("Tabla_c�sped").Range("D3").Font.Bold = True
    Sheets("Tabla_c�sped").Range("D3").VerticalAlignment = xlCenter
    
        Sheets("Tabla_c�sped").Range("D10") = Sheets("Corte_c�sped").Cells(i, 8)
        Sheets("Tabla_c�sped").Range("D10").Font.Bold = False
        Sheets("Tabla_c�sped").Range("D10").HorizontalAlignment = xlCenter
        Sheets("Tabla_c�sped").Range("D10").VerticalAlignment = xlCenter
        Sheets("Tabla_c�sped").Range("D10").WrapText = True
    
    
    'VERIFICACIONES
    Sheets("Corte_c�sped").Select
    Range("N1:R1").Select
    Selection.Copy
    Sheets("Tabla_C�sped").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    'OBSERVACIONES DE VERIFICACIONES
    Sheets("Corte_c�sped").Select
    Range(Cells(i, 14), Cells(i, 18)).Select
    Selection.Copy
    Sheets("Tabla_C�sped").Select
    Range("D4:D8").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    'COMPROBACI�N  DE INCUMPLIMIENTO NORMATIVO
    If Sheets("corte_c�sped").Cells(i, 9) = 2 Then
    Sheets("Tabla_c�sped").Range("D4") = Sheets("Tabla_c�sped").Range("D4") + " Presuntamente incumpliendo con el art�culo 2.3.2.2.2.6.66."
    End If
    
    If Sheets("corte_c�sped").Cells(i, 10) = 2 Then
    Sheets("Tabla_c�sped").Range("D5") = Sheets("Tabla_c�sped").Range("D5") + " Presuntamente incumpliendo con el art�culo 2.3.2.2.2.6.66."
    End If
    
    If Sheets("corte_c�sped").Cells(i, 11) = 2 Then
    Sheets("Tabla_c�sped").Range("D6") = Sheets("Tabla_c�sped").Range("D6") + " Presuntamente incumpliendo con el art�culo 2.3.2.2.2.6.68."
    End If
    
    If Sheets("corte_c�sped").Cells(i, 12) = 2 Then
    Sheets("Tabla_c�sped").Range("D7") = Sheets("Tabla_c�sped").Range("D7") + " Presuntamente incumpliendo con el art�culo 2.3.2.2.2.6.68."
    End If
        
    If Sheets("corte_c�sped").Cells(i, 13) = 2 Then
    Sheets("Tabla_c�sped").Range("D8") = Sheets("Tabla_c�sped").Range("D8") + " Presuntamente incumpliendo con el art�culo 2.3.2.2.2.6.68."
    End If
    
    Sheets("Tabla_C�sped").Range("C10") = "Observaciones generales"
    Sheets("Tabla_C�sped").Range("C10").WrapText = True
    
    
    Sheets("Tabla_C�sped").Range("C9") = "Dotaci�n para operarios"
    Sheets("Tabla_C�sped").Range("C9").WrapText = True
       
    'DOTACION DE OPERARIO DE GUADA�A

    For s = 19 To 27
          If Sheets("Corte_c�sped").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
        Next s
        
        If falsos > 0 Then
        Sheets("Tabla_c�sped").Range("D9") = Sheets("Tabla_c�sped").Range("D9") + "El operario de guada�a no contaba con "
        End If
    For q = 19 To 27
         If Sheets("Corte_c�sped").Cells(i, q) = False Then
          Sheets("Tabla_c�sped").Range("D9") = Sheets("Tabla_c�sped").Range("D9") + Sheets("Corte_c�sped").Cells(1, q) + ", "
             Else
             If Sheets("Corte_c�sped").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 9 Then
                Sheets("Tabla_c�sped").Range("D9") = Sheets("Tabla_c�sped").Range("D9") + " El operario de guada�a contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If

    Next q
    verdad = 0
    falsos = 0

        
   'DOTACI�N DE OPERARIOS AUXILIARE ***********************************
                   
    If Not IsEmpty(Sheets("Corte_c�sped").Cells(i, 28)) Then
    Sheets("Tabla_C�sped").Range("D9") = Sheets("Tabla_C�sped").Range("D9") + ". Los auxiliares " + Sheets("Corte_c�sped").Cells(i, 28).Value
    End If
    
    'FORMATO DE TABLA *************************************************
    Sheets("Tabla_C�sped").Range("C4:C10").Font.Bold = False
    Sheets("Tabla_C�sped").Range("D4:D10").Font.Bold = False

        Sheets("Tabla_C�sped").Range("D4:D8").Select
    Selection.Interior.Pattern = xlNone 'sin color de fondo
    Sheets("Tabla_C�sped").Range("C4:C8").Select
        With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

    Range("C2:G2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

    For p = 3 To 10
        Sheets("Tabla_C�sped").Range(Cells(p, 4), Cells(p, 7)).Select
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        End With
    Next p
        Columns("D:D").EntireColumn.AutoFit
    Range("F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").ColumnWidth = 18.57
    Columns("G:G").ColumnWidth = 21
    Columns("F:F").ColumnWidth = 21.43
    Rows("1:1").Select
    
    Sheets("Tabla_c�sped").Range("D3").HorizontalAlignment = xlCenter
    
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_C�sped").Range("C1:G10").Borders(xlInsideHorizontal).LineStyle = xlContinuous
   
       Range("C4:C10").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheets("Tabla_C�sped").Range("C10").Font.Bold = True
   
    Rows("1:10").EntireRow.AutoFit
    
    Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 9
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
     End If
    
    End If
    Next Y
    End If
Next i
     Range("A1").Select
     Application.ScreenUpdating = True
End Sub
Sub prueba()
Attribute prueba.VB_ProcData.VB_Invoke_Func = " \n14"
'
' prueba Macro
'
Dim i As Integer
Dim q As Integer


Sheets("Tabla_C�sped").Range("A1").Select



End Sub
