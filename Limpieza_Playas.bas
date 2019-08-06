Attribute VB_Name = "Limpieza_Playas"
Option Explicit

Sub Playa()

'
' Playa Macro
'
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim q As Integer
    Dim j As Integer
    Dim p As Integer
    Dim s As Integer
    Dim Y As Integer
    Dim x As Integer
    
    Dim falsos As Integer
    Dim verdad As Integer

    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_Playa"
    
    
    For i = 3 To 6
        Sheets("Tabla_Playa").Cells(1, i) = Sheets("Limpieza_playas").Cells(1, i) '.Value
        Sheets("Tabla_Playa").Cells(1, i).Font.Bold = True
        Sheets("Tabla_Playa").Cells(1, i).HorizontalAlignment = xlCenter
        Sheets("Tabla_Playa").Cells(1, i).VerticalAlignment = xlCenter
        Sheets("Tabla_Playa").Cells(1, i).WrapText = True
    Next i
    
    Columns("F:F").ColumnWidth = 46.86
    Columns("E:E").ColumnWidth = 18.14
    Columns("C:C").ColumnWidth = 15.86
    
'For i = 2 To 4

    x = Control_Trash.Cont_Registros("Limpieza_playas", 8)

For i = 2 To Control_Trash.cant_filas("Limpieza_playas") + 1
    If Sheets("Limpieza_playas").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("Limpieza_playas").Cells(i, 8).Text = Trash.ListBox1.List(Y) Then

    For p = 3 To 5
        Sheets("Tabla_Playa").Cells(2, p) = Sheets("Limpieza_playas").Cells(i, p).Value
        Sheets("Tabla_Playa").Cells(2, p).Font.Bold = False
        Sheets("Tabla_Playa").Cells(2, p).HorizontalAlignment = xlCenter
        Sheets("Tabla_Playa").Cells(2, p).VerticalAlignment = xlCenter
        Sheets("Tabla_Playa").Cells(2, p).WrapText = True
    Next p
    Range("D2").Select 'FORAMTO HORA
    Selection.NumberFormat = "[$-x-systime]h:mm AM/PM"
    
    Sheets("Tabla_Playa").Select
    Range("C2").Value = UCase(Range("C2")) 'convertir a mayuscula
    
    Sheets("Tabla_Playa").Range("D3") = "Fecha"
    Sheets("Tabla_Playa").Range("D3").Font.Bold = True
    Sheets("Tabla_Playa").Range("D3").HorizontalAlignment = xlLeft
    Sheets("Tabla_Playa").Range("D3").VerticalAlignment = xlCenter
    
    
        Sheets("Tabla_Playa").Range("D4") = Sheets("Limpieza_playas").Cells(i, 8)
        Sheets("Tabla_Playa").Range("D4").Font.Bold = False
        Sheets("Tabla_Playa").Range("D4").HorizontalAlignment = xlCenter
        Sheets("Tabla_Playa").Range("D4").VerticalAlignment = xlCenter
        Sheets("Tabla_Playa").Range("D4").WrapText = True
        Range("D4").NumberFormat = "m/d/yyyy"
    
    Sheets("Tabla_Playa").Range("E3") = "Área a intervenir"
    Sheets("Tabla_Playa").Range("E3").Font.Bold = True
    Sheets("Tabla_Playa").Range("E3").HorizontalAlignment = xlCenter
    Sheets("Tabla_Playa").Range("E3").VerticalAlignment = xlCenter
    Sheets("Tabla_Playa").Range("E3").WrapText = True
        
        Sheets("Tabla_Playa").Range("E4") = Sheets("Limpieza_playas").Cells(i, 7).Text + "m2"
        Sheets("Tabla_Playa").Range("E4").Font.Bold = False
        Sheets("Tabla_Playa").Range("E4").HorizontalAlignment = xlCenter
        Sheets("Tabla_Playa").Range("E4").VerticalAlignment = xlCenter
        Sheets("Tabla_Playa").Range("E4").WrapText = True
        
        Sheets("Tabla_Playa").Range("F2") = Sheets("Limpieza_playas").Cells(i, 6)
        Sheets("Tabla_Playa").Range("F2").Font.Bold = False
        Sheets("Tabla_Playa").Range("F2").HorizontalAlignment = xlLeft
        Sheets("Tabla_Playa").Range("F2").VerticalAlignment = xlTop
        Sheets("Tabla_Playa").Range("F2").WrapText = True
        
    
   'OBSERVACIOENS PARA OBSERVACION GENERAL

    For s = 11 To 14
          If Sheets("Limpieza_Playas").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
        Next s
        
        If falsos > 0 Then
        Sheets("Tabla_Playa").Range("F2") = Sheets("Tabla_Playa").Range("F2") + ". El operario no contaba con "
        End If
    For q = 11 To 14
         If Sheets("Limpieza_Playas").Cells(i, q) = False Then
          Sheets("Tabla_Playa").Range("F2") = Sheets("Tabla_Playa").Range("F2") + Sheets("Limpieza_Playas").Cells(1, q) + ", "
             Else
             If Sheets("Limpieza_Playas").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 4 Then
                Sheets("Tabla_Playa").Range("F2") = Sheets("Tabla_Playa").Range("F2") + ". El operario contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If

    Next q
    verdad = 0
    falsos = 0
    If Not IsEmpty(Sheets("Limpieza_Playas").Cells(i, 9)) Then
    Sheets("Tabla_Playa").Range("F2") = Sheets("Tabla_Playa").Range("F2") + " además " + Sheets("Limpieza_Playas").Cells(i, 9)
    End If
        
   '***********************************
    Sheets("Tabla_Playa").Range("F2").WrapText = True
    Sheets("Tabla_Playa").Range("F2").HorizontalAlignment = xlLeft
    Sheets("Tabla_Playa").Range("F2").VerticalAlignment = xlTop
    Sheets("Tabla_Playa").Range("F2").Font.Bold = False
    
    Rows("2:4").EntireRow.AutoFit

         'FORMATO TABLA

    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_Playa").Range("C1:F4").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 3
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
     End If
     
    If x > 1 Then
    Sheets("Tabla_Playa").Range("C5:C7").MergeCells = True
    Sheets("Tabla_Playa").Range("F5:F7").MergeCells = True
    End If
    
    End If
    Next Y
    End If
Next i
    Sheets("Tabla_Playa").Range("C2:C4").MergeCells = True
    Sheets("Tabla_Playa").Range("F2:F4").MergeCells = True
    Sheets("Tabla_Playa").Range("A1").Select
    Application.ScreenUpdating = True

End Sub



Sub prueba()

    'If IsEmpty(Sheets("Limpieza_playas").Range("I2")) Then
    'Hoja6.Range("F2") = "no esta vacia"
    If Sheets("Limpieza_playas").Cells(2, 9) <> " " Then
    Hoja6.Range("F2") = ", ademas contaba con " + Sheets("Limpieza_playas").Cells(2, 9)
    End If





End Sub


