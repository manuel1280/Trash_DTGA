Attribute VB_Name = "An�lisis"
Option Explicit

Sub DF()
Attribute DF.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
     Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb3"

    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gr�fica_DF"
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb3", Version:=6).CreatePivotTable TableDestination _
        :="Gr�fica_DF!R1C1", TableName:="TablaDin�mica4", DefaultVersion:=6
    Sheets("Gr�fica_DF").Select
    Cells(1, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gr�fica_DF!$A$1:$C$18")
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 15
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
        PivotTable.PivotFields("Cuenta con valla de informaci�n"), _
        "Suma de Cuenta con valla de informaci�n", xlSum
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
        PivotTable.PivotFields("Control de acceso al p�blico"), _
        "Suma de Control de acceso al p�blico", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ingreso de residuos peligrosos no permitidos"), _
        "Suma de Ingreso de residuos peligrosos no permitidos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ingreso de lodos contaminados y/o cenizas"), _
        "Suma de Ingreso de lodos contaminados y/o cenizas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con v�as externas"), _
        "Suma de Cuenta con v�as externas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con v�as internas"), _
        "Suma de Cuenta con v�as internas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con �reas administrativas"), _
        "Suma de Cuenta con �reas administrativas", xlSum
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
        PivotTable.PivotFields("�rea del frente de trabajo"), _
        "Suma de �rea del frente de trabajo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Densidad de compactaci�n"), _
        "Suma de Densidad de compactaci�n", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se efect�a control de gases"), _
        "Suma de Se efect�a control de gases", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se efect�a control de vectores"), _
        "Suma de Se efect�a control de vectores", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con manejo y control de lixiviados"), _
        "Suma de Cuenta con manejo y control de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Existe registro actualizado de las operaciones realizadas"), _
        "Suma de Existe registro actualizado de las operaciones realizadas", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Medici�n de estabilidad del terreno"), _
        "Suma de Medici�n de estabilidad del terreno", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se realiza reciclaje en el frente de trabajo"), _
        "Suma de Se realiza reciclaje en el frente de trabajo", xlSum
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -773.25
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 77.25
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 0.7387218045, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -261
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 4.5
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
    ActiveSheet.Name = "Gr�fica_BOP"

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_15", Version:=6).CreatePivotTable TableDestination _
        :="Gr�fica_BOP!R1C1", TableName:="TablaDin�mica6", DefaultVersion:=6
    Sheets("Gr�fica_BOP").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gr�fica_BOP!$A$1:$C$18")
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Ubicaci�n de acuerdo con el ordenamiento territorial") _
        , "Suma de Ubicaci�n de acuerdo con el ordenamiento territorial", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("�reas adecuadas para parqueo y maniobra de veh�culos") _
        , "Suma de �reas adecuadas para parqueo y maniobra de veh�culos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Adecuada se�alizaci�n de las �reas de la base operacional"), _
        "Suma de Adecuada se�alizaci�n de las �reas de la base operacional", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Adecuada se�alizaci�n de los sentidos de circulaci�n") _
        , "Suma de Adecuada se�alizaci�n de los sentidos de circulaci�n", xlSum
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
        PivotTable.PivotFields("Cuenta con servicios p�blicos"), _
        "Suma de Cuenta con servicios p�blicos", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta zona de insumos para la prestaci�n del servicio" _
        ), "Suma de Cuenta zona de insumos para la prestaci�n del servicio", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con equipos para la atenci�n de emergencias"), _
        "Suma de Cuenta con equipos para la atenci�n de emergencias", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con equipos de control contra incendios"), _
        "Suma de Cuenta con equipos de control contra incendios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Cuenta con equipos de comunicaci�n entre la base y los veh�culos"), _
        "Suma de Cuenta con equipos de comunicaci�n entre la base y los veh�culos", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Hay veh�culos de recolecci�n con residuos de la operaci�n"), _
        "Suma de Hay veh�culos de recolecci�n con residuos de la operaci�n", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Cuenta con zona de lavado o utilizan en otras instalaciones"), _
        "Suma de Cuenta con zona de lavado o utilizan en otras instalaciones", xlSum
        '*************
    'ActiveSheet.PivotTables("TablaDin�mica6").PivotFields("Nombre de la empresa"). _
     '   CurrentPage = Trash.ComboBox2.Text
      '  ActiveSheet.PivotTables("TablaDin�mica6").PivotFields("Fecha"). _
       ' CurrentPage = CDate(Trash.ComboBox1.Text)
        '***************
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -471.75
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 78.75

    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -702.75
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop -13.5
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 0.6612824278, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -567.75
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 1.5
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 0.7202072539, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -252
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 6.75
    Columns("A:N").Select
    Selection.ColumnWidth = 30.43
    Selection.ColumnWidth = 17.71
    Selection.ColumnWidth = 12.43

    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 4.2043478261, msoFalse, _
        msoScaleFromTopLeft
    Range("A7").Select
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    Range("A6").Select
    
        Application.ScreenUpdating = True
End Sub
Sub Corte_cesped()
Attribute Corte_cesped.VB_ProcData.VB_Invoke_Func = " \n14"
'

    Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb_17"
    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gr�fica_CC�sped"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_17", Version:=6).CreatePivotTable TableDestination _
        :="Gr�fica_CC�sped!R1C1", TableName:="TablaDin�mica8", DefaultVersion:=6
    Sheets("Gr�fica_CC�sped").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gr�fica_CC�sped!$A$1:$C$18")
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Direcci�n del �rea intervenida")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Corte de c�sped de forma programada"), _
        "Suma de Corte de c�sped de forma programada", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("La altura de c�sped supera los diez cent�metros"), _
        "Suma de La altura de c�sped supera los diez cent�metros", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con valla informativa"), _
        "Suma de Cuenta con valla informativa", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Se realiza demarcaci�n de la zona de trabajo"), _
        "Suma de Se realiza demarcaci�n de la zona de trabajo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Cuenta con malla de protecci�n"), _
        "Suma de Cuenta con malla de protecci�n", xlSum

    Columns("B:F").Select
    Range("F1").Activate
    Selection.ColumnWidth = 29.86
    Selection.ColumnWidth = 26.57
    Selection.ColumnWidth = 24
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -363
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 107.25
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 1.3198198198, msoFalse, _
        msoScaleFromTopLeft
    ActiveWindow.ScrollColumn = 1
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -157.5
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 3
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
    ActiveSheet.Name = "Gr�fica_PArboles"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb8", Version:=6).CreatePivotTable TableDestination _
        :="Gr�fica_PArboles!R1C1", TableName:="TablaDin�mica10", DefaultVersion:=6
    Sheets("Gr�fica_PArboles").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gr�fica_PArboles!$A$1:$C$18")
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Nombre de la empresa")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Fecha")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "individuo arboreo o direcci�n")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Malla de protecci�n"), "Suma de Malla de protecci�n", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Demarcaci�n de la zona"), _
        "Suma de Demarcaci�n de la zona", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Valla informativa"), "Suma de Valla informativa", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Capacitaci�n de operarios"), _
        "Suma de Capacitaci�n de operarios", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Permiso de poda"), "Suma de Permiso de poda", xlSum
    Range("E1").Select
    ActiveWindow.ScrollColumn = 2
    Columns("A:F").Select
    Range("F1").Activate
    Selection.ColumnWidth = 16.14
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -313.5
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 96.75
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 1.309352518, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 5.25
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop -0.75
    ActiveSheet.PivotTables("TablaDin�mica10").PivotSelect "Nombre de la empresa", _
        xlButton, True
        Application.ScreenUpdating = True
End Sub

Sub RYT()
'
'VEH�CULOS
'
    Application.ScreenUpdating = False
    Application.Goto Reference:="Tabla_Formularios_.accdb_1"
    
    Sheets.Add After:=Sheets("R&T")
    ActiveSheet.Name = "Gr�fica_veh�culos"
  
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla_Formularios_.accdb_1", Version:=6).CreatePivotTable TableDestination _
        :="Gr�fica_veh�culos!R1C1", TableName:="TablaDin�mica1", DefaultVersion:=6
    Sheets("Gr�fica_veh�culos").Select
    Cells(1, 1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Gr�fica_veh�culos!$A$1:$C$18")
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft 240
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 15
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
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Placa del veh�culo")
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
        PivotTable.PivotFields("Veh�culo claramente identificado"), _
        "Suma de Veh�culo claramente identificado", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de comunicaciones"), _
        "Suma de Equipo de comunicaciones", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Transporta residuos de construcci�n"), _
        "Suma de Transporta residuos de construcci�n", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de compactaci�n con detenido de emergencia"), _
        "Suma de Equipo de compactaci�n con detenido de emergencia", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Caja compactadora cerrada que impide fuga de lixiviados"), _
        "Suma de Caja compactadora cerrada que impide fuga de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Mecanismo autom�tico para liberaci�n de lixiviados"), _
        "Suma de Mecanismo autom�tico para liberaci�n de lixiviados", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Balizas y luces estrobosc�picas sobre la cabina"), _
        "Suma de Balizas y luces estrobosc�picas sobre la cabina", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Balizas y luces estrobosc�picas en caja de compactaci�n"), _
        "Suma de Balizas y luces estrobosc�picas en caja de compactaci�n", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Luces en la zona de tolva"), _
        "Suma de Luces en la zona de tolva", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields( _
        "Tubo de escape por encima de altura m�xima del veh�culo"), _
        "Suma de Tubo de escape por encima de altura m�xima del veh�culo", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Estribos antideslizantes y manijas en buen estado"), _
        "Suma de Estribos antideslizantes y manijas en buen estado", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Elementos complementarios para la operaci�n"), _
        "Suma de Elementos complementarios para la operaci�n", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Equipo de carreteras y atenci�n de incendios"), _
        "Suma de Equipo de carreteras y atenci�n de incendios", xlSum
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
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveChart.Legend.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Gr�fico 1").IncrementLeft -242.25
    ActiveSheet.Shapes("Gr�fico 1").IncrementTop 100.5
    ActiveSheet.Shapes("Gr�fico 1").ScaleWidth 2.3224043716, msoFalse, _
        msoScaleFromTopLeft
    Range("A1").Select
        Application.ScreenUpdating = True
End Sub

