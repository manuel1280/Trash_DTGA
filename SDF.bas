Attribute VB_Name = "SDF"
Option Explicit

Sub SDF()

    Application.ScreenUpdating = False

    Dim i As Integer
    Dim j As Integer
    Dim Y As Integer
    Dim x As Integer

    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Tabla_SDF"
    
    x = Control_Trash.Cont_Registros("SDF", 45)

For i = 2 To Control_Trash.cant_filas("SDF") + 1
    If Sheets("SDF").Cells(i, 2) = Trash.ComboBox2.Text Then
    For Y = 0 To Trash.ListBox1.ListCount - 1
    If Sheets("SDF").Cells(i, 45).Text = Trash.ListBox1.List(Y) Then
    
    
  'For i = 2 To 3
    'NOMBRE DEL RS
    Sheets("Tabla_SDF").Range("C2") = Sheets("SDF").Cells(i, 2)
    Sheets("Tabla_SDF").Range("C2").Font.Bold = True
    Sheets("Tabla_SDF").Range("C2").HorizontalAlignment = xlCenter
    Range("C2").Value = UCase(Range("C2")) 'convertir a mayuscula
        
           Range("C2:K2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    'SERVICIOS PUB
    Sheets("Tabla_SDF").Range("C9") = "SERVICIOS PÚBLICOS"
    Sheets("Tabla_SDF").Range("C9").Font.Bold = False
    Sheets("Tabla_SDF").Range("C9").HorizontalAlignment = xlCenter
         
           Range("C9:K9").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
        'vías de acceso
    Sheets("Tabla_SDF").Range("C11") = "VÍAS DE ACCESO"
    Sheets("Tabla_SDF").Range("C11").Font.Bold = False
    Sheets("Tabla_SDF").Range("C11").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C11").VerticalAlignment = xlTop

            'CERRAMIENTO PERIMETRAL
    Sheets("Tabla_SDF").Range("F11") = "CERRAMIENTO PERIMETRAL"
    Sheets("Tabla_SDF").Range("F11").Font.Bold = False
    Sheets("Tabla_SDF").Range("F11").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("F11").VerticalAlignment = xlTop
         
              'SISTEMA DE PESAJE
    Sheets("Tabla_SDF").Range("H11") = "SISTEMA DE PESAJE"
    Sheets("Tabla_SDF").Range("H11").Font.Bold = False
    Sheets("Tabla_SDF").Range("H11").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("H11").VerticalAlignment = xlTop
    
                'COMPACTACION DE RESIDUOS
    Sheets("Tabla_SDF").Range("C15") = "COMPACTACIÓN DE RESIDUOS"
    Sheets("Tabla_SDF").Range("C15").Font.Bold = False
    Sheets("Tabla_SDF").Range("C15").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C15").VerticalAlignment = xlTop
    
                    'COBERTURA DE RESIDUOS
    Sheets("Tabla_SDF").Range("H15") = "COBERTURA DE RESIDUOS"
    Sheets("Tabla_SDF").Range("H15").Font.Bold = False
    Sheets("Tabla_SDF").Range("H15").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("H15").VerticalAlignment = xlTop
    
                    'CONTROL DE AGUAS LLUVIAS - ESCORRENTÍA
    Sheets("Tabla_SDF").Range("C19") = "CONTROL DE AGUAS LLUVIAS - ESCORRENTÍA"
    Sheets("Tabla_SDF").Range("C19").Font.Bold = False
    Sheets("Tabla_SDF").Range("C19").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C19").VerticalAlignment = xlTop
    
                    'CONTROL DE VECTORES
    Sheets("Tabla_SDF").Range("H19") = "CONTROL DE VECTORES"
    Sheets("Tabla_SDF").Range("H19").Font.Bold = False
    Sheets("Tabla_SDF").Range("H19").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("H19").VerticalAlignment = xlTop
    
                    'CONTROL DE GASES
    Sheets("Tabla_SDF").Range("C23") = "CONTROL DE GASES"
    Sheets("Tabla_SDF").Range("C23").Font.Bold = False
    Sheets("Tabla_SDF").Range("C23").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C23").VerticalAlignment = xlTop
    
    
                    'MANEJO DE LIXIVIADOS
    Sheets("Tabla_SDF").Range("C26") = "MANEJO DE LIXIVIADOS"
    Sheets("Tabla_SDF").Range("C26").Font.Bold = False
    Sheets("Tabla_SDF").Range("C26").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C26").VerticalAlignment = xlTop
    
                'DISPOSICIÓN DE RESIDUOS DE DEMOLICIÓN Y CONSTRUCCIÓN
    Sheets("Tabla_SDF").Range("C30") = "DISPOSICIÓN DE RESIDUOS DE DEMOLICIÓN Y CONSTRUCCIÓN"
    Sheets("Tabla_SDF").Range("C30").Font.Bold = False
    Sheets("Tabla_SDF").Range("C30").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("C30").VerticalAlignment = xlTop
    
        'ALMACENAMIENTO DE RESIDUOS PELIGROSOS
    Sheets("Tabla_SDF").Range("H30") = "ALMACENAMIENTO DE RESIDUOS PELIGROSOS"
    Sheets("Tabla_SDF").Range("H30").Font.Bold = False
    Sheets("Tabla_SDF").Range("H30").HorizontalAlignment = xlCenter
    Sheets("Tabla_SDF").Range("H30").VerticalAlignment = xlTop
       
             Range("C11:K11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
            Range("C15:K15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
            Range("C19:K19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
            Range("C23:K23").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
                Range("C26:K26").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
                Range("C30:K30").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    

    'UBICACIÓN
    Sheets("Tabla_SDF").Range("C3") = "Ubicación: " + Sheets("SDF").Cells(i, 3)
    Sheets("Tabla_SDF").Range("C3").Font.Bold = False
    Sheets("Tabla_SDF").Range("C3").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C3").VerticalAlignment = xlTop
            
        Sheets("Tabla_SDF").Range("C3:F4").Select
        Selection.MergeCells = True
        Selection.WrapText = True
    
    'municipios atendidos
    Sheets("Tabla_SDF").Range("G3") = "Municipios atendidos: " + Sheets("SDF").Cells(i, 4)
    Sheets("Tabla_SDF").Range("G3").Font.Bold = False
    Sheets("Tabla_SDF").Range("G3").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("G3").VerticalAlignment = xlTop
        
        Sheets("Tabla_SDF").Range("G3:K4").Select
        Selection.MergeCells = True
        Selection.WrapText = True
        
    'autorización ambiental
    Sheets("Tabla_SDF").Range("C5") = "Ultima autorización ambiental: " + Sheets("SDF").Cells(i, 5)
    Sheets("Tabla_SDF").Range("C5").Font.Bold = False
    Sheets("Tabla_SDF").Range("C5").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C5").VerticalAlignment = xlTop
    
        Sheets("Tabla_SDF").Range("C5:E7").MergeCells = True
        Sheets("Tabla_SDF").Range("C5:E7").Select
        Selection.WrapText = True
    
    'INICIO DE OPERACIONES
    Sheets("Tabla_SDF").Range("C8") = "Año incio de operaciones:  " + Sheets("SDF").Cells(i, 44).Text
    Sheets("Tabla_SDF").Range("C8").Font.Bold = False
    Sheets("Tabla_SDF").Range("C8").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C8").VerticalAlignment = xlTop
    
        Sheets("Tabla_SDF").Range("C8:E8").MergeCells = True
        Sheets("Tabla_SDF").Range("C8:E8").Select
        Selection.WrapText = True
    
    'Área total del predio:
    Sheets("Tabla_SDF").Range("F5") = "Área total del predio(m2):   " + Sheets("SDF").Cells(i, 6).Text
    Sheets("Tabla_SDF").Range("F5").Font.Bold = False
    Sheets("Tabla_SDF").Range("F5").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F5").VerticalAlignment = xlTop
        
        Sheets("Tabla_SDF").Range("F5:H5").MergeCells = True
        Sheets("Tabla_SDF").Range("F5:H7").Select
        Selection.WrapText = True
        
    'Área del frente de trabajo:
    Sheets("Tabla_SDF").Range("F6") = "Área del frente de trabajo (m2):   " + Sheets("SDF").Cells(i, 7).Text
    Sheets("Tabla_SDF").Range("F6").Font.Bold = False
    Sheets("Tabla_SDF").Range("F6").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F6").VerticalAlignment = xlTop
    
         Sheets("Tabla_SDF").Range("F6:H6").MergeCells = True
        Sheets("Tabla_SDF").Range("F6:H6").Select
        Selection.WrapText = True
    
      'Toneladas promedio día:
    Sheets("Tabla_SDF").Range("F7") = "Toneladas promedio día:  " + Sheets("SDF").Cells(i, 8).Text + "Ton"
    Sheets("Tabla_SDF").Range("F7").Font.Bold = False
    Sheets("Tabla_SDF").Range("F7").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F7").VerticalAlignment = xlTop
    
        Sheets("Tabla_SDF").Range("F7:H7").MergeCells = True
        Sheets("Tabla_SDF").Range("F7:H7").Select
        Selection.WrapText = True
    
       'Tiempo vida útil:
    Sheets("Tabla_SDF").Range("I5") = "Tiempo vida útil:   " + Sheets("SDF").Cells(i, 9).Text + "años"
    Sheets("Tabla_SDF").Range("I5").Font.Bold = False
    Sheets("Tabla_SDF").Range("I5").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("I5").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("I5:K5").MergeCells = True
        Sheets("Tabla_SDF").Range("I5:K5").Select
        Selection.WrapText = True
    
       'Capacidad total:
    Sheets("Tabla_SDF").Range("I6") = "Capacidad total:  " + Sheets("SDF").Cells(i, 10).Text + "Ton"
    Sheets("Tabla_SDF").Range("I6").Font.Bold = False
    Sheets("Tabla_SDF").Range("I6").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("I6").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("I6:K6").MergeCells = True
        Sheets("Tabla_SDF").Range("I6:K6").Select
        Selection.WrapText = True

       'Capacidad remanente:
    Sheets("Tabla_SDF").Range("I7") = "Capacidad remanente:  " + Sheets("SDF").Cells(i, 11).Text + "Ton"
    Sheets("Tabla_SDF").Range("I7").Font.Bold = False
    Sheets("Tabla_SDF").Range("I7").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("I7").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("I7:K7").MergeCells = True
        Sheets("Tabla_SDF").Range("I7:K7").Select
        Selection.WrapText = True
       'Tipo de SDF:
    Sheets("Tabla_SDF").Range("F8") = "Tipo de SDF: " + Sheets("SDF").Cells(i, 12)
    Sheets("Tabla_SDF").Range("F8").Font.Bold = False
    Sheets("Tabla_SDF").Range("F8").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F8").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("F8:H8").MergeCells = True
        Sheets("Tabla_SDF").Range("F8:H8").Select
        Selection.WrapText = True
    
       'Número de celdas activas:
    Sheets("Tabla_SDF").Range("I8") = "Número de celdas activas:   " + Sheets("SDF").Cells(i, 13).Text
    Sheets("Tabla_SDF").Range("I8").Font.Bold = False
    Sheets("Tabla_SDF").Range("I8").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("I8").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("I8:K8").MergeCells = True
    Sheets("Tabla_SDF").Range("I8:K8").Select
        Selection.WrapText = True
    
        'Servicios publicos:
    Sheets("Tabla_SDF").Range("C10") = Sheets("SDF").Cells(i, 14)
    Sheets("Tabla_SDF").Range("C10").Font.Bold = False
    Sheets("Tabla_SDF").Range("C10").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C10").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C10:K10").MergeCells = True
    Sheets("Tabla_SDF").Range("C10:K10").Select
        Selection.WrapText = True
    
        'Ancho vía de acceso (m):
    Sheets("Tabla_SDF").Range("C12") = "Ancho vía de acceso:  " + Sheets("SDF").Cells(i, 15).Text + "m"
    Sheets("Tabla_SDF").Range("C12").Font.Bold = False
    Sheets("Tabla_SDF").Range("C12").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C12").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C12:E12").MergeCells = True
    Sheets("Tabla_SDF").Range("C12:E12").Select
        Selection.WrapText = True
    
        'Estado:
    Sheets("Tabla_SDF").Range("C13") = "Estado: " + Sheets("SDF").Cells(i, 16)
    Sheets("Tabla_SDF").Range("C13").Font.Bold = False
    Sheets("Tabla_SDF").Range("C13").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C13").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C13:E13").MergeCells = True
    Sheets("Tabla_SDF").Range("C13:E13").Select
        Selection.WrapText = True
        'Material:
    Sheets("Tabla_SDF").Range("C14") = "Material: " + Sheets("SDF").Cells(i, 17)
    Sheets("Tabla_SDF").Range("C14").Font.Bold = False
    Sheets("Tabla_SDF").Range("C14").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C14").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("C14:E14").MergeCells = True
    Sheets("Tabla_SDF").Range("C14:E14").Select
        Selection.WrapText = True
        'Puertas de acceso:
    If Sheets("SDF").Cells(i, 18) = True Then
        Sheets("Tabla_SDF").Range("F12") = "Puertas de acceso: SI "
        Sheets("Tabla_SDF").Range("F12").Font.Bold = False
        Sheets("Tabla_SDF").Range("F12").HorizontalAlignment = xlLeft
        Sheets("Tabla_SDF").Range("F12").VerticalAlignment = xlTop
    Else
    Sheets("Tabla_SDF").Range("F12") = "Puertas de acceso: NO "
    End If
    Sheets("Tabla_SDF").Range("F12:G12").MergeCells = True
    Sheets("Tabla_SDF").Range("F12:G12").Select
        Selection.WrapText = True
    
        'Estado del cerramiento:
    Sheets("Tabla_SDF").Range("F13") = "Estado del cerramiento: " + Sheets("SDF").Cells(i, 19)
    Sheets("Tabla_SDF").Range("F13").Font.Bold = False
    Sheets("Tabla_SDF").Range("F13").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F13").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("F13:G13").MergeCells = True
    Sheets("Tabla_SDF").Range("F13:G13").Select
        Selection.WrapText = True
    
       'Material del cerramiento:
    Sheets("Tabla_SDF").Range("F14") = "Material del cerramiento: " + Sheets("SDF").Cells(i, 20)
    Sheets("Tabla_SDF").Range("F14").Font.Bold = False
    Sheets("Tabla_SDF").Range("F14").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("F14").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("F14:G14").MergeCells = True
    Sheets("Tabla_SDF").Range("F14:G14").Select
        Selection.WrapText = True
       'Posee bascula:
    If Sheets("SDF").Cells(i, 21) = True Then
        Sheets("Tabla_SDF").Range("H12") = "Posee bascula: SI"
        Sheets("Tabla_SDF").Range("H12").Font.Bold = False
        Sheets("Tabla_SDF").Range("H12").HorizontalAlignment = xlLeft
        Sheets("Tabla_SDF").Range("H12").VerticalAlignment = xlTop
    Else
        Sheets("Tabla_SDF").Range("H12") = "Posee bascula: NO"
    End If
    
    Sheets("Tabla_SDF").Range("H12:I12").MergeCells = True
    Sheets("Tabla_SDF").Range("H12:I12").Select
        Selection.WrapText = True
        
        'Capacidad:
    Sheets("Tabla_SDF").Range("H13") = "Capacidad: " + Sheets("SDF").Cells(i, 22).Text + "Ton"
    Sheets("Tabla_SDF").Range("H13").Font.Bold = False
    Sheets("Tabla_SDF").Range("H13").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H13").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H13:I13").MergeCells = True
    Sheets("Tabla_SDF").Range("H13:I13").Select
        Selection.WrapText = True

        'Fecha última calibración:
    Sheets("Tabla_SDF").Range("H14") = "Fecha última calibración: " + Sheets("SDF").Cells(i, 23).Text
    Sheets("Tabla_SDF").Range("H14").Font.Bold = False
    Sheets("Tabla_SDF").Range("H14").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H14").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H14:I14").MergeCells = True
    Sheets("Tabla_SDF").Range("H14:I14").Select
        Selection.WrapText = True
        'Sistema de registro:
    Sheets("Tabla_SDF").Range("J12") = "Sistema de registro: " + Sheets("SDF").Cells(i, 24)
    Sheets("Tabla_SDF").Range("J12").Font.Bold = False
    Sheets("Tabla_SDF").Range("J12").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("J12").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("J12:K14").MergeCells = True
    Sheets("Tabla_SDF").Range("J12:K14").Select
        Selection.WrapText = True
        'Densidad:
    Sheets("Tabla_SDF").Range("C16") = "Densidad (Ton/m3): " + Sheets("SDF").Cells(i, 25).Text
    Sheets("Tabla_SDF").Range("C16").Font.Bold = False
    Sheets("Tabla_SDF").Range("C16").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C16").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("C16:D16").MergeCells = True
    Sheets("Tabla_SDF").Range("C16:D16").Select
        Selection.WrapText = True
        'Sistema de medición:
    Sheets("Tabla_SDF").Range("C17") = "Sistema de medición: " + Sheets("SDF").Cells(i, 26)
    Sheets("Tabla_SDF").Range("C17").Font.Bold = False
    Sheets("Tabla_SDF").Range("C17").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C17").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C17:D18").MergeCells = True
    Sheets("Tabla_SDF").Range("C17:D17").Select
        Selection.WrapText = True

        'Sistema para la estabilidad de taludes:
    Sheets("Tabla_SDF").Range("E16") = "Sistema para la estabilidad de taludes: " + Sheets("SDF").Cells(i, 27)
    Sheets("Tabla_SDF").Range("E16").Font.Bold = False
    Sheets("Tabla_SDF").Range("E16").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("E16").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("E16:G18").MergeCells = True
    Sheets("Tabla_SDF").Range("E16:G18").Select
        Selection.WrapText = True
    
        'Material
    Sheets("Tabla_SDF").Range("H16") = "Material: " + Sheets("SDF").Cells(i, 28)
    Sheets("Tabla_SDF").Range("H16").Font.Bold = False
    Sheets("Tabla_SDF").Range("H16").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H16").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H16:I16").MergeCells = True
    Sheets("Tabla_SDF").Range("H16:I16").Select
        Selection.WrapText = True

        'Horario
    Sheets("Tabla_SDF").Range("H17") = "Horario: " + Sheets("SDF").Cells(i, 29)
    Sheets("Tabla_SDF").Range("H17").Font.Bold = False
    Sheets("Tabla_SDF").Range("H17").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H17").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H17:I17").MergeCells = True
    Sheets("Tabla_SDF").Range("H17:I17").Select
        Selection.WrapText = True

        'Frecuencia
    Sheets("Tabla_SDF").Range("H18") = "Frecuencia: " + Sheets("SDF").Cells(i, 30)
    Sheets("Tabla_SDF").Range("H18").Font.Bold = False
    Sheets("Tabla_SDF").Range("H18").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H18").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("H18:I18").MergeCells = True
    Sheets("Tabla_SDF").Range("H18:I18").Select
        Selection.WrapText = True
        'Maquinaria utilizada
    Sheets("Tabla_SDF").Range("J16") = "Maquinaria utilizada: " + Sheets("SDF").Cells(i, 31)
    Sheets("Tabla_SDF").Range("J16").Font.Bold = False
    Sheets("Tabla_SDF").Range("J16").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("J16").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("J16:K18").MergeCells = True
    Sheets("Tabla_SDF").Range("J16:K18").Select
        Selection.WrapText = True
        'Sistema
    Sheets("Tabla_SDF").Range("C20") = "Sistema: " + Sheets("SDF").Cells(i, 32)
    Sheets("Tabla_SDF").Range("C20").Font.Bold = False
    Sheets("Tabla_SDF").Range("C20").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C20").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C20:G21").MergeCells = True
    Sheets("Tabla_SDF").Range("c20:G21").Select
        Selection.WrapText = True
    
         'Estado
    Sheets("Tabla_SDF").Range("C22") = "Estado: " + Sheets("SDF").Cells(i, 34)
    Sheets("Tabla_SDF").Range("C22").Font.Bold = False
    Sheets("Tabla_SDF").Range("C22").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C22").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C22:G22").MergeCells = True
    Sheets("Tabla_SDF").Range("c22:G22").Select
        Selection.WrapText = True
  
        'Sistema
    Sheets("Tabla_SDF").Range("H20") = "Sistema: " + Sheets("SDF").Cells(i, 33)
    Sheets("Tabla_SDF").Range("H20").Font.Bold = False
    Sheets("Tabla_SDF").Range("H20").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H20").VerticalAlignment = xlTop
   
   Sheets("Tabla_SDF").Range("H20:K21").MergeCells = True
    Sheets("Tabla_SDF").Range("H20:K21").Select
        Selection.WrapText = True
   
        'fRECUENCIA
    Sheets("Tabla_SDF").Range("H22") = "Frecuencia: " + Sheets("SDF").Cells(i, 35)
    Sheets("Tabla_SDF").Range("H22").Font.Bold = False
    Sheets("Tabla_SDF").Range("H22").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H22").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H22:K22").MergeCells = True
    Sheets("Tabla_SDF").Range("H22:K22").Select
        Selection.WrapText = True
    
        'Sistema de extracción y/o aprovechamiento:
    Sheets("Tabla_SDF").Range("C24") = "Sistema de extracción y/o aprovechamiento: " + Sheets("SDF").Cells(i, 36)
    Sheets("Tabla_SDF").Range("C24").Font.Bold = False
    Sheets("Tabla_SDF").Range("C24").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C24").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C24:K25").MergeCells = True
    Sheets("Tabla_SDF").Range("C24:K25").Select
        Selection.WrapText = True
    
        'Tren de tratamiento:
    Sheets("Tabla_SDF").Range("C27") = "Tren de tratamiento: " + Sheets("SDF").Cells(i, 37)
    Sheets("Tabla_SDF").Range("C27").Font.Bold = False
    Sheets("Tabla_SDF").Range("C27").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C27").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C27:F29").MergeCells = True
    Sheets("Tabla_SDF").Range("C27:F29").Select
        Selection.WrapText = True

        'Caudal manejado
    Sheets("Tabla_SDF").Range("G27") = "Caudal tratado: " + Sheets("SDF").Cells(i, 46)
    Sheets("Tabla_SDF").Range("G27").Font.Bold = False
    Sheets("Tabla_SDF").Range("G27").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("G27").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("G27:G29").MergeCells = True
    Sheets("Tabla_SDF").Range("G27:G29").Select
        Selection.WrapText = True


        'Permiso de vertimientos:
    Sheets("Tabla_SDF").Range("H27") = "Permiso de vertimientos: " + Sheets("SDF").Cells(i, 38)
    Sheets("Tabla_SDF").Range("H27").Font.Bold = False
    Sheets("Tabla_SDF").Range("H27").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H27").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H27:I29").MergeCells = True
    Sheets("Tabla_SDF").Range("H27:I29").Select
        Selection.WrapText = True

        'Tratamiento de biosólidos:
    Sheets("Tabla_SDF").Range("J27") = "Tratamiento de biosólidos: " + Sheets("SDF").Cells(i, 39)
    Sheets("Tabla_SDF").Range("J27").Font.Bold = False
    Sheets("Tabla_SDF").Range("J27").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("J27").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("J27:K29").Select
        Selection.MergeCells = True
        Selection.WrapText = True

        'Área:
    Sheets("Tabla_SDF").Range("C31") = "Área: " + Sheets("SDF").Cells(i, 40).Text
    Sheets("Tabla_SDF").Range("C31").Font.Bold = False
    Sheets("Tabla_SDF").Range("C31").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("C31").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("C31:C32").MergeCells = True
    Sheets("Tabla_SDF").Range("C31:C32").Select
        Selection.WrapText = True

        'Sistema de disposición:
    Sheets("Tabla_SDF").Range("D31") = "Sistema de disposición: " + Sheets("SDF").Cells(i, 41)
    Sheets("Tabla_SDF").Range("D31").Font.Bold = False
    Sheets("Tabla_SDF").Range("D31").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("D31").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("D31:G32").MergeCells = True
    Sheets("Tabla_SDF").Range("D31:G32").Select
        Selection.WrapText = True

        'Tipo de residuos:
    Sheets("Tabla_SDF").Range("H31") = "Tipo de residuos: " + Sheets("SDF").Cells(i, 42)
    Sheets("Tabla_SDF").Range("H31").Font.Bold = False
    Sheets("Tabla_SDF").Range("H31").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H31").VerticalAlignment = xlTop
    
    Sheets("Tabla_SDF").Range("H31:K31").MergeCells = True
    Sheets("Tabla_SDF").Range("H31:K31").Select
        Selection.WrapText = True
    
        'Método de disposición:
    Sheets("Tabla_SDF").Range("H32") = "Método de disposición: " + Sheets("SDF").Cells(i, 43)
    Sheets("Tabla_SDF").Range("H32").Font.Bold = False
    Sheets("Tabla_SDF").Range("H32").HorizontalAlignment = xlLeft
    Sheets("Tabla_SDF").Range("H32").VerticalAlignment = xlTop

    Sheets("Tabla_SDF").Range("H32:K32").MergeCells = True
    Sheets("Tabla_SDF").Range("H32:K32").Select
        Selection.WrapText = True
    
    Sheets("Tabla_SDF").Select
    Range("C2:K2").MergeCells = True
    Range("C9:K9").MergeCells = True
    Range("C11:E11").MergeCells = True
    Range("F11:G11").MergeCells = True
    Range("H11:K11").MergeCells = True
    Range("C15:G15").MergeCells = True
    Range("H15:K15").MergeCells = True
    Range("C19:G19").MergeCells = True
    Range("H19:K19").MergeCells = True
    Range("C23:K23").MergeCells = True
    Range("C26:K26").MergeCells = True
    Range("C30:G30").MergeCells = True
    Range("H30:K30").MergeCells = True

    'FORMATO DE TABLA
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlEdgeRight).LineStyle = xlContinuous
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlInsideVertical).LineStyle = xlContinuous
    Sheets("Tabla_SDF").Range("C2:K32").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Rows("2:32").EntireRow.AutoFit

        Rows("2:2").Select 'INSERTAR FILAS
     If i <= x Then
        For j = 1 To 32
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next j
    End If
    'Sheets("Tabla_SDF").Range("C33:K33").MergeCells = True

    End If
    Next Y
    End If
Next i
    
    Sheets("Tabla_SDF").Range("C2:k2").MergeCells = True
    Range("A1").Select
    
End Sub
