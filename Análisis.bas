Attribute VB_Name = "Análisis"
Option Explicit

Sub DF()
Attribute DF.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
     Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb3"

    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gráfica_DF"
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb3", Version:=6).CreatePivotTable TableDestination _
        :="Gráfica_DF!R1C1", TableName:="TablaDinámica4", DefaultVersion:=6
    Sheets("Gráfica_DF").Select
    Cells(1, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gráfica_DF!$A$1:$C$18")
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre del operador")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre del SDF")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Hora")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.PivotFields("Hora").Orientation = xlHidden
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha de verificacion")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con valla de información"), _
        "Suma de Cuenta con valla de información", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con cerramiento perimetral"), _
        "Suma de Cuenta con cerramiento perimetral", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con puertas de acceso"), _
        "Suma de Cuenta con puertas de acceso", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con bascula de pesaje"), _
        "Suma de Cuenta con bascula de pesaje", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se realiza registro de pesaje"), _
        "Suma de Se realiza registro de pesaje", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Control de acceso al público"), _
        "Suma de Control de acceso al público", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ingreso de residuos peligrosos no permitidos"), _
        "Suma de Ingreso de residuos peligrosos no permitidos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ingreso de lodos contaminados y/o cenizas"), _
        "Suma de Ingreso de lodos contaminados y/o cenizas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con vías externas"), _
        "Suma de Cuenta con vías externas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con vías internas"), _
        "Suma de Cuenta con vías internas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con áreas administrativas"), _
        "Suma de Cuenta con áreas administrativas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Posee control de incendios"), _
        "Suma de Posee control de incendios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Operarios con equipos de seguridad"), _
        "Suma de Operarios con equipos de seguridad", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cubrimiento diario de residuos"), _
        "Suma de Cubrimiento diario de residuos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Área del frente de trabajo"), _
        "Suma de Área del frente de trabajo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Densidad de compactación"), _
        "Suma de Densidad de compactación", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se efectúa control de gases"), _
        "Suma de Se efectúa control de gases", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se efectúa control de vectores"), _
        "Suma de Se efectúa control de vectores", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con manejo y control de lixiviados"), _
        "Suma de Cuenta con manejo y control de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Existe registro actualizado de las operaciones realizadas"), _
        "Suma de Existe registro actualizado de las operaciones realizadas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Medición de estabilidad del terreno"), _
        "Suma de Medición de estabilidad del terreno", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se realiza reciclaje en el frente de trabajo"), _
        "Suma de Se realiza reciclaje en el frente de trabajo", xlSum
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -773.25
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 77.25
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 0.7387218045, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -261
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 4.5
    Range("A10").Select
    
    Application.ScreenUpdating = True
        
End Sub
Sub BOP()
Attribute BOP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BASE DE OPERACIONES

        Application.ScreenUpdating = False
    
    Application.Goto Reference:="Tabla_Formularios_.accdb_15"

    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gráfica_BOP"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_15", Version:=6).CreatePivotTable TableDestination _
        :="Gráfica_BOP!R1C1", TableName:="TablaDinámica6", DefaultVersion:=6
    Sheets("Gráfica_BOP").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gráfica_BOP!$A$1:$C$18")
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ubicación de acuerdo con el ordenamiento territorial") _
        , "Suma de Ubicación de acuerdo con el ordenamiento territorial", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Áreas adecuadas para parqueo y maniobra de vehículos") _
        , "Suma de Áreas adecuadas para parqueo y maniobra de vehículos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Adecuada señalización de las áreas de la base operacional"), _
        "Suma de Adecuada señalización de las áreas de la base operacional", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Adecuada señalización de los sentidos de circulación") _
        , "Suma de Adecuada señalización de los sentidos de circulación", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con oficinas administrativas"), _
        "Suma de Cuenta con oficinas administrativas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta zona de control de operaciones"), _
        "Suma de Cuenta zona de control de operaciones", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Vestidores e instalaciones sanitarias para el personal" _
        ), "Suma de Vestidores e instalaciones sanitarias para el personal", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con servicios públicos"), _
        "Suma de Cuenta con servicios públicos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta zona de insumos para la prestación del servicio" _
        ), "Suma de Cuenta zona de insumos para la prestación del servicio", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con equipos para la atención de emergencias"), _
        "Suma de Cuenta con equipos para la atención de emergencias", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con equipos de control contra incendios"), _
        "Suma de Cuenta con equipos de control contra incendios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Cuenta con equipos de comunicación entre la base y los vehículos"), _
        "Suma de Cuenta con equipos de comunicación entre la base y los vehículos", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Hay vehículos de recolección con residuos de la operación"), _
        "Suma de Hay vehículos de recolección con residuos de la operación", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Cuenta con zona de lavado o utilizan en otras instalaciones"), _
        "Suma de Cuenta con zona de lavado o utilizan en otras instalaciones", xlSum
        '*************
    'ActiveSheet.PivotTables("TablaDinámica6").PivotFields("Nombre de la empresa"). _
     '   CurrentPage = Trash.ComboBox2.Text
      '  ActiveSheet.PivotTables("TablaDinámica6").PivotFields("Fecha"). _
       ' CurrentPage = CDate(Trash.ComboBox1.Text)
        '***************
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -471.75
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 78.75

    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -702.75
    ActiveSheet.Shapes("Gráfico 1").IncrementTop -13.5
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 0.6612824278, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -567.75
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 1.5
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 0.7202072539, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -252
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 6.75
    Columns("A:N").Select
    Selection.ColumnWidth = 30.43
    Selection.ColumnWidth = 17.71
    Selection.ColumnWidth = 12.43

    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 4.2043478261, msoFalse, _
        msoScaleFromTopLeft
    Range("A7").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    Range("A6").Select
    
        Application.ScreenUpdating = True
End Sub
Sub Corte_cesped()
Attribute Corte_cesped.VB_ProcData.VB_Invoke_Func = " \n14"
'

    Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb_17"
    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gráfica_CCésped"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_17", Version:=6).CreatePivotTable TableDestination _
        :="Gráfica_CCésped!R1C1", TableName:="TablaDinámica8", DefaultVersion:=6
    Sheets("Gráfica_CCésped").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gráfica_CCésped!$A$1:$C$18")
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Dirección del área intervenida")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Corte de césped de forma programada"), _
        "Suma de Corte de césped de forma programada", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("La altura de césped supera los diez centímetros"), _
        "Suma de La altura de césped supera los diez centímetros", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con valla informativa"), _
        "Suma de Cuenta con valla informativa", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se realiza demarcación de la zona de trabajo"), _
        "Suma de Se realiza demarcación de la zona de trabajo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con malla de protección"), _
        "Suma de Cuenta con malla de protección", xlSum

    Columns("B:F").Select
    Range("F1").Activate
    Selection.ColumnWidth = 29.86
    Selection.ColumnWidth = 26.57
    Selection.ColumnWidth = 24
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -363
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 107.25
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 1.3198198198, msoFalse, _
        msoScaleFromTopLeft
    ActiveWindow.ScrollColumn = 1
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -157.5
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 3
    Range("A1").Select

    Application.ScreenUpdating = True
End Sub
Sub Poda_arboles()
Attribute Poda_arboles.VB_ProcData.VB_Invoke_Func = " \n14"
'
'PODA ARBOLES
    Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb8"

    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gráfica_PArboles"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb8", Version:=6).CreatePivotTable TableDestination _
        :="Gráfica_PArboles!R1C1", TableName:="TablaDinámica10", DefaultVersion:=6
    Sheets("Gráfica_PArboles").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gráfica_PArboles!$A$1:$C$18")
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "individuo arboreo o dirección")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Malla de protección"), "Suma de Malla de protección", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Demarcación de la zona"), _
        "Suma de Demarcación de la zona", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Valla informativa"), "Suma de Valla informativa", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Capacitación de operarios"), _
        "Suma de Capacitación de operarios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Permiso de poda"), "Suma de Permiso de poda", xlSum
    Range("E1").Select
    ActiveWindow.ScrollColumn = 2
    Columns("A:F").Select
    Range("F1").Activate
    Selection.ColumnWidth = 16.14
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -313.5
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 96.75
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 1.309352518, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 5.25
    ActiveSheet.Shapes("Gráfico 1").IncrementTop -0.75
    ActiveSheet.PivotTables("TablaDinámica10").PivotSelect "Nombre de la empresa", _
        xlButton, True
        Application.ScreenUpdating = True
End Sub

Sub RYT()
'
'VEHÍCULOS
'
    Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb_1"
    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gráfica_vehículos"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_1", Version:=6).CreatePivotTable TableDestination _
        :="Gráfica_vehículos!R1C1", TableName:="TablaDinámica1", DefaultVersion:=6
    Sheets("Gráfica_vehículos").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gráfica_vehículos!$A$1:$C$18")
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Microrruta")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Placa del vehículo")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Plano de la micorruta"), _
        "Suma de Plano de la micorruta", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Documentos de transito"), _
        "Suma de Documentos de transito", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Vehículo claramente identificado"), _
        "Suma de Vehículo claramente identificado", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de comunicaciones"), _
        "Suma de Equipo de comunicaciones", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Transporta residuos de construcción"), _
        "Suma de Transporta residuos de construcción", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de compactación con detenido de emergencia"), _
        "Suma de Equipo de compactación con detenido de emergencia", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Caja compactadora cerrada que impide fuga de lixiviados"), _
        "Suma de Caja compactadora cerrada que impide fuga de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Mecanismo automático para liberación de lixiviados"), _
        "Suma de Mecanismo automático para liberación de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Balizas y luces estroboscópicas sobre la cabina"), _
        "Suma de Balizas y luces estroboscópicas sobre la cabina", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Balizas y luces estroboscópicas en caja de compactación"), _
        "Suma de Balizas y luces estroboscópicas en caja de compactación", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Luces en la zona de tolva"), _
        "Suma de Luces en la zona de tolva", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Tubo de escape por encima de altura máxima del vehículo"), _
        "Suma de Tubo de escape por encima de altura máxima del vehículo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Estribos antideslizantes y manijas en buen estado"), _
        "Suma de Estribos antideslizantes y manijas en buen estado", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Elementos complementarios para la operación"), _
        "Suma de Elementos complementarios para la operación", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de carreteras y atención de incendios"), _
        "Suma de Equipo de carreteras y atención de incendios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Hay fuga de lixiviados"), _
        "Suma de Hay fuga de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Operarios con elementos de seguridad"), _
        "Suma de Operarios con elementos de seguridad", xlSum
    Range("D1").Select

    Columns("C:Q").Select
    Range("Q1").Activate
    Selection.ColumnWidth = 12.57
    Selection.ColumnWidth = 8
    Columns("B:B").Select
    Selection.ColumnWidth = 16.71
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Legend.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft -242.25
    ActiveSheet.Shapes("Gráfico 1").IncrementTop 100.5
    ActiveSheet.Shapes("Gráfico 1").ScaleWidth 2.3224043716, msoFalse, _
        msoScaleFromTopLeft
    Range("A1").Select
        Application.ScreenUpdating = True
End Sub

