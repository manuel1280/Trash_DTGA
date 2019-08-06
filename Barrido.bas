Attribute VB_Name = "Barrido"
Option Explicit

Sub Barrido()
Attribute Barrido.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Barrido Macro
'
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim j As Integer
    Dim s As Integer
    Dim q As Integer
    Dim Y As Integer
    Dim x As Integer
    
    Dim verdad As Integer
    Dim falsos As Integer

    
    'CREAR NUEVA PAGINA****************************************
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_Barrido"
    
    'FORMATO DE LETRA******************************************
    Sheets("Tabla_barrido").Range("C1") = "Macrorruta"
    Sheets("Tabla_barrido").Range("C1").Font.Bold = True
    Sheets("Tabla_barrido").Range("C1").HorizontalAlignment = xlGeneral
    Sheets("Tabla_barrido").Range("C1").VerticalAlignment = xlCenter
    
    Sheets("Tabla_barrido").Range("D1") = "Microrruta"
    Sheets("Tabla_barrido").Range("D1").Font.Bold = True
    Sheets("Tabla_barrido").Range("D1").HorizontalAlignment = xlGeneral
    Sheets("Tabla_barrido").Range("D1").VerticalAlignment = xlCenter
    
    Sheets("Tabla_barrido").Range("E1") = "Hora verificada"
    Sheets("Tabla_barrido").Range("E1").Font.Bold = True
    Sheets("Tabla_barrido").Range("E1").HorizontalAlignment = xlGeneral
    Sheets("Tabla_barrido").Range("E1").VerticalAlignment = xlCenter
        
    Sheets("Tabla_barrido").Range("F1") = "Dirección"
    Sheets("Tabla_barrido").Range("F1").Font.Bold = True
    Sheets("Tabla_barrido").Range("F1").HorizontalAlignment = xlGeneral
    Sheets("Tabla_barrido").Range("F1").VerticalAlignment = xlCenter
        
    Sheets("Tabla_barrido").Range("G1") = "Observaciones"
    Sheets("Tabla_barrido").Range("G1").Font.Bold = True
    Sheets("Tabla_barrido").Range("G1").HorizontalAlignment = xlGeneral
    Sheets("Tabla_barrido").Range("G1").VerticalAlignment = xlCenter
        
        Columns("G:G").Select
    Selection.ColumnWidth = 39.29
    Selection.ColumnWidth = 46.86

    x = Control_Trash.Cont_Registros("barrido", 10)

For i = 2 To Control_Trash.cant_filas("barrido") + 1
    If Sheets("barrido").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("barrido").Cells(i, 10).Text = Trash.ListBox1.List(Y) Then


    Sheets("Barrido").Select
    Range(Cells(i, 3), Cells(i, 7)).Select
    Selection.Copy
    Sheets("Tabla_Barrido").Select
    Cells(2, 3).Select
    ActiveSheet.Paste
    Selection.WrapText = True
    Selection.Interior.Pattern = xlNone 'sin color de fondo
        

'OBSERVACIOENS PARA OBSERVACION GENERAL******************************

    For s = 14 To 25
          If Sheets("Barrido").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
        Next s

        If falsos > 0 Then
        Sheets("Tabla_Barrido").Range("G2") = Sheets("Tabla_Barrido").Range("G2") + ". El operario no contaba con "
        End If
        
    For q = 14 To 25
         If Sheets("Barrido").Cells(i, q) = False Then
          Sheets("Tabla_Barrido").Range("G2") = Sheets("Tabla_Barrido").Range("G2") + Sheets("Barrido").Cells(1, q) + ", "
             Else
             If Sheets("Barrido").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 12 Then
                Sheets("Tabla_Barrido").Range("G2") = Sheets("Tabla_Barrido").Range("G2") + ". El operario contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If

    Next q
    verdad = 0
    falsos = 0
    
    If Not IsEmpty(Sheets("Barrido").Cells(i, 13)) Then
    Sheets("Tabla_Barrido").Range("G2") = Sheets("Tabla_Barrido").Range("G2") + " además contaba con " + Sheets("Barrido").Cells(i, 13)
    End If
        
   'FORMATO DE LETRA ***********************************
    Sheets("Tabla_barrido").Range("D2").HorizontalAlignment = xlCenter
    
    Sheets("Tabla_barrido").Range("D3") = "Horario"
    Sheets("Tabla_barrido").Range("D3").Font.Bold = True
    Sheets("Tabla_barrido").Range("D3").HorizontalAlignment = xlLeft
    Sheets("Tabla_barrido").Range("D3").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("D3").WrapText = True
    
        Sheets("Tabla_barrido").Range("D4") = Sheets("Barrido").Cells(i, 8)
        Sheets("Tabla_barrido").Range("D4").Font.Bold = False
        Sheets("Tabla_barrido").Range("D4").HorizontalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("D4").VerticalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("D4").WrapText = True
    
    Sheets("Tabla_barrido").Range("E3") = "Fecha verificada"
    Sheets("Tabla_barrido").Range("E3").Font.Bold = True
    Sheets("Tabla_barrido").Range("E3").HorizontalAlignment = xlLeft
    Sheets("Tabla_barrido").Range("E3").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("E3").WrapText = True
        
        Sheets("Tabla_barrido").Range("E4") = Sheets("Barrido").Cells(i, 10)
        Sheets("Tabla_barrido").Range("E4").Font.Bold = False
        Sheets("Tabla_barrido").Range("E4").HorizontalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("E4").VerticalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("E4").WrapText = True
        
    
    Sheets("Tabla_barrido").Range("D5") = "Frecuencia"
    Sheets("Tabla_barrido").Range("D5").Font.Bold = True
    Sheets("Tabla_barrido").Range("D5").HorizontalAlignment = xlLeft
    Sheets("Tabla_barrido").Range("D5").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("D5").WrapText = True

        Sheets("Tabla_barrido").Range("D6") = Sheets("Barrido").Cells(i, 10)
        Sheets("Tabla_barrido").Range("D6").Font.Bold = False
        Sheets("Tabla_barrido").Range("D6").HorizontalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("D6").VerticalAlignment = xlCenter
        Sheets("Tabla_barrido").Range("D6").WrapText = True
    
    Sheets("Tabla_barrido").Range("G5") = "Recolección de bolsas de barrido"
    Sheets("Tabla_barrido").Range("G5").Font.Bold = True
    Sheets("Tabla_barrido").Range("G5").HorizontalAlignment = xlLeft
    Sheets("Tabla_barrido").Range("G5").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("G5").WrapText = True
    
    Sheets("Tabla_barrido").Range("D6") = Sheets("Barrido").Cells(i, 9).Value
    Sheets("Tabla_barrido").Range("D6").Font.Bold = False
    Sheets("Tabla_barrido").Range("D6").HorizontalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("D6").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("D6").WrapText = True
    
    Sheets("Tabla_barrido").Range("G6") = "La microrruta de recolección es " + Sheets("Barrido").Cells(i, 12) + " con horario de recolección de " + Sheets("Barrido").Cells(i, 11)
    Sheets("Tabla_barrido").Range("D6").Font.Bold = False
    Sheets("Tabla_barrido").Range("G6").HorizontalAlignment = xlLeft
    Sheets("Tabla_barrido").Range("G6").VerticalAlignment = xlCenter
    Sheets("Tabla_barrido").Range("G6").WrapText = True
    
        
        Sheets("Tabla_barrido").Range(Cells(6, 4), Cells(7, 4)).MergeCells = True
        Sheets("Tabla_barrido").Range(Cells(4, 5), Cells(7, 5)).MergeCells = True
        Sheets("Tabla_barrido").Range(Cells(6, 7), Cells(7, 7)).MergeCells = True
    
    'FORMATO TABLA************************************************
    Sheets("tabla_barrido").Range("C1:G7").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("tabla_barrido").Range("C1:G7").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("tabla_barrido").Range("C1:G7").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_barrido").Range("C1:G7").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_barrido").Range("C1:G7").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_barrido").Range("C1:G7").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Range("G6:G7").Select
    Selection.Font.Bold = False
    Sheets("Tabla_barrido").Range("G2:G4").VerticalAlignment = xlTop
    
    Range("C2:C7").HorizontalAlignment = xlCenter
    Range("C2:C7").VerticalAlignment = xlCenter
    Range("F2:F7").VerticalAlignment = xlCenter
    Range("E4:E7").Select
    Selection.NumberFormat = "m/d/yyyy"
    

    Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 6
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
    End If
    
    If x > 1 Then
    Sheets("Tabla_barrido").Range("C8:C13").MergeCells = True
    Sheets("Tabla_barrido").Range("F8:F13").MergeCells = True
    Sheets("Tabla_barrido").Range("G8:G10").MergeCells = True
    End If
    
    Rows("2:7").EntireRow.AutoFit
    
    End If
    Next Y
    End If
Next i
    
    
    Sheets("Tabla_barrido").Range("C2:C7").MergeCells = True
    Sheets("Tabla_barrido").Range("F2:F7").MergeCells = True
    Sheets("Tabla_barrido").Range("G2:G4").MergeCells = True

    
    Columns("E:E").EntireColumn.AutoFit
    Range("A1").Select
    Application.ScreenUpdating = True

End Sub
