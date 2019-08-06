Attribute VB_Name = "Lavado_areaspublicas"
Option Explicit

Sub Lavado()

'
' Lavado Macro
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
    ActiveSheet.Name = "Tabla_Lavado"
  
    
    Sheets("Tabla_Lavado").Range("C1") = "Zona objeto de lavado"
    Sheets("Tabla_Lavado").Range("C1").Font.Bold = True
    Sheets("Tabla_Lavado").Range("C1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Lavado").Range("C1").VerticalAlignment = xlCenter
    Sheets("Tabla_Lavado").Range("C1").WrapText = True
    
    Sheets("Tabla_Lavado").Range("D1") = "Hora"
    Sheets("Tabla_Lavado").Range("D1").Font.Bold = True
    Sheets("Tabla_Lavado").Range("D1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Lavado").Range("D1").VerticalAlignment = xlCenter
    
    
    Sheets("Tabla_Lavado").Range("E1") = "Observaciones"
    Sheets("Tabla_Lavado").Range("E1").Font.Bold = True
    Sheets("Tabla_Lavado").Range("E1").HorizontalAlignment = xlCenter
    Sheets("Tabla_Lavado").Range("E1").VerticalAlignment = xlCenter
    
    Columns("D:D").ColumnWidth = 15.86
    Columns("C:C").ColumnWidth = 17.43
    Columns("E:E").ColumnWidth = 50.2
    

    x = Control_Trash.Cont_Registros("Lavado_áreas", 5)

For i = 2 To Control_Trash.cant_filas("Lavado_áreas") + 1
    If Sheets("Lavado_áreas").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("Lavado_áreas").Cells(i, 5).Text = Trash.ListBox1.List(Y) Then
    
        Sheets("Tabla_Lavado").Range("C2") = Sheets("Lavado_áreas").Cells(i, 3)
        Sheets("Tabla_Lavado").Range("C2").Font.Bold = False
        Sheets("Tabla_Lavado").Range("C2").HorizontalAlignment = xlCenter
        Sheets("Tabla_Lavado").Range("C2").VerticalAlignment = xlCenter
        Sheets("Tabla_Lavado").Range("C2").WrapText = True
        
        Sheets("Tabla_lavado").Range("C2").Select
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
            End With
        
        Sheets("Tabla_Lavado").Range("D2") = Sheets("Lavado_áreas").Cells(i, 4)
        Sheets("Tabla_Lavado").Range("D2").Font.Bold = False
        Sheets("Tabla_Lavado").Range("D2").HorizontalAlignment = xlCenter
        Sheets("Tabla_Lavado").Range("D2").VerticalAlignment = xlCenter
        Sheets("Tabla_Lavado").Range("D2").WrapText = True
        Range("D2").NumberFormat = "[$-x-systime]h:mm: AM/PM"
        
        'OBSERVACIONES GENERALES
        Sheets("Tabla_Lavado").Range("E2") = Sheets("Lavado_áreas").Cells(i, 8)
        Sheets("Tabla_Lavado").Range("E2").Font.Bold = False
        Sheets("Tabla_Lavado").Range("E2").HorizontalAlignment = xlLeft
        Sheets("Tabla_Lavado").Range("E2").VerticalAlignment = xlTop
        Sheets("Tabla_Lavado").Range("E2").WrapText = True
        
        Sheets("Tabla_Lavado").Range("E5") = "Dotación de operarios "
        Sheets("Tabla_Lavado").Range("E5").Font.Bold = True
        Sheets("Tabla_Lavado").Range("E5").HorizontalAlignment = xlLeft
        Sheets("Tabla_Lavado").Range("E5").VerticalAlignment = xlCenter
            
            Sheets("Tabla_Lavado").Range("E6").Font.Bold = False
            Sheets("Tabla_Lavado").Range("E6").HorizontalAlignment = xlLeft
            Sheets("Tabla_Lavado").Range("E6").VerticalAlignment = xlTop
            Sheets("Tabla_Lavado").Range("E6").WrapText = True
            
        Sheets("Tabla_Lavado").Range("D3") = "Fecha"
        Sheets("Tabla_Lavado").Range("D3").Font.Bold = True
        Sheets("Tabla_Lavado").Range("D3").HorizontalAlignment = xlLeft
        Sheets("Tabla_Lavado").Range("D3").VerticalAlignment = xlCenter
            
            Sheets("Tabla_Lavado").Range("D4") = Sheets("Lavado_áreas").Cells(i, 5)
            Sheets("Tabla_Lavado").Range("D4").Font.Bold = False
            Sheets("Tabla_Lavado").Range("D4").HorizontalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("D4").VerticalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("D4").WrapText = True
            Range("D4").NumberFormat = "m/d/yyyy"
                      
        Sheets("Tabla_Lavado").Range("C5") = "Dirección"
        Sheets("Tabla_Lavado").Range("C5").Font.Bold = True
        Sheets("Tabla_Lavado").Range("C5").HorizontalAlignment = xlLeft
        Sheets("Tabla_Lavado").Range("C5").VerticalAlignment = xlCenter
        
            Sheets("Tabla_Lavado").Range("C6") = Sheets("Lavado_áreas").Cells(i, 6)
            Sheets("Tabla_Lavado").Range("C6").Font.Bold = False
            Sheets("Tabla_Lavado").Range("C6").HorizontalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("C6").VerticalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("C6").WrapText = True
        
        Sheets("Tabla_Lavado").Range("D5") = "Área lavada (m2)"
        Sheets("Tabla_Lavado").Range("D5").Font.Bold = True
        Sheets("Tabla_Lavado").Range("D5").HorizontalAlignment = xlLeft
        Sheets("Tabla_Lavado").Range("D5").VerticalAlignment = xlCenter
       
            Sheets("Tabla_Lavado").Range("D6") = Sheets("Lavado_áreas").Cells(i, 7)
            Sheets("Tabla_Lavado").Range("D6").Font.Bold = False
            Sheets("Tabla_Lavado").Range("D6").HorizontalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("D6").VerticalAlignment = xlCenter
            Sheets("Tabla_Lavado").Range("D6").WrapText = True
       
        
    'OBSERVACIONES COMPLEMENTARIAS****************************************
    
    'OPERARIO 1
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 10)) Then
        Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + "El operario de " + Sheets("Lavado_áreas").Cells(i, 10)
        
        For s = 13 To 21
          If Sheets("Lavado_áreas").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
        Next s
        
        If falsos > 0 Then
        Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " no contaba con "
        End If
    For q = 13 To 21
         If Sheets("Lavado_áreas").Cells(i, q) = False Then
          Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + Sheets("Lavado_áreas").Cells(1, q) + ", "
             Else
             If Sheets("Lavado_áreas").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 9 Then
                Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If

    Next q
    verdad = 0
    falsos = 0
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 22)) Then
    Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + ", ademas contaba con " + Sheets("Lavado_áreas").Cells(i, 22)
    End If
    End If
    
    'OPERARIO 2
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 11)) Then
    Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + ". El operario de " + Sheets("Lavado_áreas").Cells(i, 11)
    
    For s = 23 To 31
          If Sheets("Lavado_áreas").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
    Next s
        If falsos > 0 Then
        Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " no contaba con "
        End If
    For q = 23 To 31
        If Sheets("Lavado_áreas").Cells(i, q) = False Then
          Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + Sheets("Lavado_áreas").Cells(1, q) + ", "
         Else
             If Sheets("Lavado_áreas").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 9 Then
                Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If

    Next q
    verdad = 0
    falsos = 0
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 32)) Then
    Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + ", ademas contaba con " + Sheets("Lavado_áreas").Cells(i, 32)
    End If
    End If
    
    'OPERARIO 3
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 12)) Then
    Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + ". El operario de " + Sheets("Lavado_áreas").Cells(i, 12)
        
    For s = 33 To 41
          If Sheets("Lavado_áreas").Cells(i, s) = False Then
          falsos = falsos + 1
          End If
    Next s
        If falsos > 0 Then
        Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " no contaba con "
        End If
    For q = 33 To 41
        If Sheets("Lavado_áreas").Cells(i, q) = False Then
          Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + Sheets("Lavado_áreas").Cells(1, q) + ", "
         Else
             If Sheets("Lavado_áreas").Cells(i, q) = True Then
            verdad = verdad + 1
                If verdad = 9 Then
                Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + " contaba con los elementos de seguridad y elementos de trabajo"
         End If
         End If
         End If
    Next q
    verdad = 0
    falsos = 0
    If Not IsEmpty(Sheets("Lavado_áreas").Cells(i, 42)) Then
    Sheets("Tabla_lavado").Range("E6") = Sheets("Tabla_lavado").Range("E6") + ", ademas contaba con " + Sheets("Lavado_áreas").Cells(i, 42)
    End If
    End If

'***********************

    Rows("2:7").EntireRow.AutoFit
    
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_Lavado").Range("C1:E6").Borders(xlInsideHorizontal).LineStyle = xlContinuous
   
    Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 5
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
     End If
    
    If x > 1 Then
    Sheets("Tabla_lavado").Range("C7:C9").MergeCells = True
    Sheets("Tabla_lavado").Range("E7:E9").MergeCells = True
    End If
    
    End If
    Next Y
    End If
Next i
    Sheets("Tabla_lavado").Range("C2:C4").MergeCells = True
    Sheets("Tabla_lavado").Range("E2:E4").MergeCells = True
    Range("A1").Select
    
    Application.ScreenUpdating = True
    


End Sub

Sub prueba()





End Sub
