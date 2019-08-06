Attribute VB_Name = "RYT"
Option Explicit
Sub R_y_T()

    Application.ScreenUpdating = False

    Dim i As Integer
    Dim h As Integer
    Dim j As Integer
    Dim s As Integer
    Dim q As Integer
    Dim Y As Integer
    Dim x As Integer
    Dim m As Integer

    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_RyT"
    
     Columns("C:C").ColumnWidth = 56.25
     Columns("D:D").ColumnWidth = 15.86
    
    Sheets("Tabla_RyT").Range("C2") = "Placa del veh�culo"
    Sheets("Tabla_RyT").Range("C2").Font.Bold = False
    Sheets("Tabla_RyT").Range("C2").HorizontalAlignment = xlCenter
    
    Sheets("Tabla_RyT").Range("C3") = "Tipo de veh�culo"
    Sheets("Tabla_RyT").Range("C3").Font.Bold = False
    Sheets("Tabla_RyT").Range("C3").HorizontalAlignment = xlCenter

    Sheets("Tabla_RyT").Range("C4") = "Modelo"
    Sheets("Tabla_RyT").Range("C4").Font.Bold = False
    Sheets("Tabla_RyT").Range("C4").HorizontalAlignment = xlCenter

    Sheets("Tabla_RyT").Range("C5") = "Marca"
    Sheets("Tabla_RyT").Range("C5").Font.Bold = False
    Sheets("Tabla_RyT").Range("C5").HorizontalAlignment = xlCenter

    Sheets("Tabla_RyT").Range("C6") = "Capacidad"
    Sheets("Tabla_RyT").Range("C6").Font.Bold = False
    Sheets("Tabla_RyT").Range("C6").HorizontalAlignment = xlCenter
    
    Sheets("Tabla_RyT").Range("C7") = "Verificaciones"
    Sheets("Tabla_RyT").Range("C7").Font.Bold = True
    Sheets("Tabla_RyT").Range("C7").HorizontalAlignment = xlCenter
    
       Range("C2:D7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    'VERIFIACCIONES
    Sheets("Tabla_RyT").Range("C8") = "�El conductor porta el plano de la microrruta?"
    Sheets("Tabla_RyT").Range("C8").Font.Bold = False
    Sheets("Tabla_RyT").Range("C8").HorizontalAlignment = xlLeft
    'Sheets("Tabla_RyT").Range("C8").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C8").WrapText = True
    
        Sheets("Tabla_RyT").Range("C9") = "�Los documentos de transito se encuentran al d�a? (licencia de conducci�n, SOAT, revisi�n t�cnico mec�nica, tarjeta de propiedad)?"
    Sheets("Tabla_RyT").Range("C9").Font.Bold = False
    Sheets("Tabla_RyT").Range("C9").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C9").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C9").WrapText = True
    
        Sheets("Tabla_RyT").Range("C10") = "�El veh�culo se encuentra claramente identificado? (color, logotipos, placa de identificaci�n)?"
    Sheets("Tabla_RyT").Range("C10").Font.Bold = False
    Sheets("Tabla_RyT").Range("C10").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C10").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C10").WrapText = True
    
        Sheets("Tabla_RyT").Range("C11") = "�Posee equipo de comunicaciones? �Cual? (>5000 suscript.)?"
    Sheets("Tabla_RyT").Range("C11").Font.Bold = False
    Sheets("Tabla_RyT").Range("C11").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C11").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C11").WrapText = True
    
        Sheets("Tabla_RyT").Range("C12") = "�Transporta residuos de construcci�n, demolici�n u otros residuos que no sean susceptibles de ser compactados?"
    Sheets("Tabla_RyT").Range("C12").Font.Bold = False
    Sheets("Tabla_RyT").Range("C12").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C12").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C12").WrapText = True
    
        Sheets("Tabla_RyT").Range("C13") = "�Si cuenta con equipo de compactaci�n, ��ste puede ser detenido en caso de emergencia?"
    Sheets("Tabla_RyT").Range("C13").Font.Bold = False
    Sheets("Tabla_RyT").Range("C13").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C13").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C13").WrapText = True
    
        Sheets("Tabla_RyT").Range("C14") = "�Las cajas compactadoras son de tipo cerrada, de manera que impidan la p�rdida del l�quido (lixiviado)?"
    Sheets("Tabla_RyT").Range("C14").Font.Bold = False
    Sheets("Tabla_RyT").Range("C14").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C14").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C14").WrapText = True
    
        Sheets("Tabla_RyT").Range("C15") = "�Cuenta con mecanismo autom�tico que permita una r�pida acci�n de descarga de lixiviado?"
    Sheets("Tabla_RyT").Range("C15").Font.Bold = False
    Sheets("Tabla_RyT").Range("C15").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C15").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C15").WrapText = True
    
        Sheets("Tabla_RyT").Range("C16") = "�Posee balizas o luces de estrobosc�picas, ubicadas sobre la cabina?"
    Sheets("Tabla_RyT").Range("C16").Font.Bold = False
    Sheets("Tabla_RyT").Range("C16").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C16").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C16").WrapText = True
    
        Sheets("Tabla_RyT").Range("C17") = "�Posee balizas o luces de estrobosc�picas, ubicadas en la parte posterior de la caja de compactaci�n?"
    Sheets("Tabla_RyT").Range("C17").Font.Bold = False
    Sheets("Tabla_RyT").Range("C17").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C17").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C17").WrapText = True
    
        Sheets("Tabla_RyT").Range("C18") = "�Posee luces en la zona de la tolva?"
    Sheets("Tabla_RyT").Range("C18").Font.Bold = False
    Sheets("Tabla_RyT").Range("C18").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C18").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C18").WrapText = True
    
        Sheets("Tabla_RyT").Range("C19") = "�El tubo de escape se encuentra ubicado hacia arriba y por encima de su altura m�xima?"
    Sheets("Tabla_RyT").Range("C19").Font.Bold = False
    Sheets("Tabla_RyT").Range("C19").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C19").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C19").WrapText = True
    
    Sheets("Tabla_RyT").Range("C20") = "�Cuenta con estribos con superficies antideslizantes, y manijas adecuadas para sujetarse? "
    Sheets("Tabla_RyT").Range("C20").Font.Bold = False
    Sheets("Tabla_RyT").Range("C20").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C20").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C20").WrapText = True
    
            Sheets("Tabla_RyT").Range("C21") = "�Se encuentra dotado de elementos complementarios tales como cepillos, escobas y palas y estos se encuentran en buenas condiciones?"
    Sheets("Tabla_RyT").Range("C21").Font.Bold = False
    Sheets("Tabla_RyT").Range("C21").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C21").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C21").WrapText = True
    
    Sheets("Tabla_RyT").Range("C22") = "�Posee equipos de carretera y de atenci�n de incendios? "
    Sheets("Tabla_RyT").Range("C22").Font.Bold = False
    Sheets("Tabla_RyT").Range("C22").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C22").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C22").WrapText = True
    
    Sheets("Tabla_RyT").Range("C23") = "�Hay presencia de fuga de l�quido (lixiviado)? "
    Sheets("Tabla_RyT").Range("C23").Font.Bold = False
    Sheets("Tabla_RyT").Range("C23").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C23").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C23").WrapText = True
    
    Sheets("Tabla_RyT").Range("C24") = "�Los operarios cuentan con elementos de seguridad industrial?"
    Sheets("Tabla_RyT").Range("C24").Font.Bold = False
    Sheets("Tabla_RyT").Range("C24").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C24").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C24").WrapText = True
    
    Sheets("Tabla_RyT").Range("C25") = "Si no cuenta con caja compactadora, �se encuentran cubiertos los residuos s�lidos, de forma que no permita el esparcimiento de residuos durante el recorrido?  "
    Sheets("Tabla_RyT").Range("C25").Font.Bold = False
    Sheets("Tabla_RyT").Range("C25").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C25").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C25").WrapText = True
    
    Sheets("Tabla_RyT").Range("C26") = "Si no cuenta con equipo de compactaci�n, �posee mecanismos que eviten la p�rdida del l�quido (lixiviado)?  "
    Sheets("Tabla_RyT").Range("C26").Font.Bold = False
    Sheets("Tabla_RyT").Range("C26").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C25").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C26").WrapText = True
    
    Sheets("Tabla_RyT").Range("C27") = "Si no cuenta con equipo de compactaci�n, �posee luces ubicadas en la cabina?  "
    Sheets("Tabla_RyT").Range("C27").Font.Bold = False
    Sheets("Tabla_RyT").Range("C27").HorizontalAlignment = xlLeft
    Sheets("Tabla_RyT").Range("C25").VerticalAlignment = xlTop
    Sheets("Tabla_RyT").Range("C27").WrapText = True
    
    Sheets("Tabla_RyT").Range("D2") = "Decreto 1077 del 2015"
    Sheets("Tabla_RyT").Range("D2").Font.Bold = False

    Range("D2").WrapText = True
    
    Range("D2:D7").MergeCells = True
    Range("D2:D7").HorizontalAlignment = xlCenter
    Range("D2:D7").VerticalAlignment = xlCenter
    Range("D2:D7").WrapText = True
    
 
    'Formato de texto
    For m = 8 To 27
    Sheets("Tabla_RyT").Cells(m, 4).HorizontalAlignment = xlCenter
    Sheets("Tabla_RyT").Cells(m, 4).VerticalAlignment = xlCenter
    
    'INSERCI�N DE NORMATIVA A COMPARAR*******************************
    Select Case m
    Case 8
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.30"
    Case 9, 24
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (13)"
    Case 10
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (1)"
    Case 11
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (2)"
    Case 12
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (3)"
    Case 13
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (5)"
    Case 14, 23, 25, 26
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (10)"
    Case 15
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (6)"
    Case 16, 17, 18, 27
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (17)"
    Case 19
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (4)"
    Case 20
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (7)"
    Case 21
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (16)"
    Case 22
    Sheets("Tabla_RyT").Cells(m, 4) = "2.3.2.2.2.3.36 (14)"
    End Select
    Next m
    
    q = 5 'COLUMNA INICIAL DE REGISTRO
  
    x = Control_Trash.Cont_Registros("R&T", 4)
    
    'VERIFICACI�N NORMATIVA*******************************************
    'Busca si la empresa seleccionada esta en la base de datos y ejecuta
For s = 2 To Control_Trash.cant_filas("R&T") + 1
    If Sheets("R&T").Cells(s, 2) = Trash.ComboBox2.Text Then 'buscar nombre de empresa
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("R&T").Cells(s, 4).Text = Trash.ListBox1.List(Y) Then 'buscar fecha selecionada
    
    i = 15 'columna inicial de parametros a verificar
    
    For j = 8 To 27 'recorrido de filas (tabla)
    If Sheets("R&T").Cells(s, i) = 1 Then 'conversion de convenciones
    Sheets("Tabla_RyT").Cells(j, q) = "SI" ' el numero 1 es convertido a un "SI"
    Else
    If Sheets("R&T").Cells(s, i) = 2 Then
    Sheets("Tabla_RyT").Cells(j, q) = "NO" ' el numero 2 es convertido a un "NO"
    Else
    Sheets("Tabla_RyT").Cells(j, q) = "---" ' campo vacio es convertido a un "---"
    End If
    End If
    'Formato celda*******
    Sheets("Tabla_RyT").Cells(j, q).Font.Bold = False
    Sheets("Tabla_RyT").Cells(j, q).HorizontalAlignment = xlCenter
    Sheets("Tabla_RyT").Cells(j, q).VerticalAlignment = xlCenter
    Sheets("Tabla_RyT").Cells(j, q).Select
        With Selection.Font
        .Name = "Calibri"
        .Size = 10
        End With '*******
        
 'comparaci�n logica de la normativa
 'colorea de color azul los incumplimientos
    If Sheets("Tabla_RyT").Cells(j, q) = "SI" Then
      If j = 12 Or j = 23 Then
        Cells(j, q).Select
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
        End With
      End If
    Else
    If Sheets("Tabla_RyT").Cells(j, q) = "NO" Then
     If j <> 12 And j <> 23 Then
         Cells(j, q).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    End If
    End If
    End If
    i = i + 1
    Next j

' ENCABEZADO
    h = 9
        For j = 2 To 6 'fila tabla
    Sheets("Tabla_RyT").Cells(j, q) = Sheets("R&T").Cells(s, h)
    Sheets("Tabla_RyT").Cells(j, q).Font.Bold = False
    Sheets("Tabla_RyT").Cells(j, q).HorizontalAlignment = xlCenter
    Sheets("Tabla_RyT").Cells(j, q).VerticalAlignment = xlCenter
    h = h + 1
        Next j
    q = q + 1
    
    End If
    Next Y
    End If
Next s
        'FORMATO TABLA*****************************************
    Sheets("Tabla_RyT").Range(Cells(2, 3), Cells(27, q - 1)).Select
    With Selection
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    Sheets("Tabla_RyT").Range("C2:D27").Select
        With Selection.Font
        .Name = "Calibri"
        .Size = 10
        End With

    Rows("2:277").EntireRow.AutoFit
    
    
    Sheets("Tabla_RyT").Range("A1").Select
    Application.ScreenUpdating = True
End Sub
