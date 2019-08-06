Attribute VB_Name = "Control_Trash"
Option Explicit

Sub Generar()

Trash.Show

    
End Sub
Sub filtrar_empresa()
    Dim i As Integer
    Dim Visibles As Integer
    
    Application.ScreenUpdating = False

'BARRIDO
    Visibles = 0
    Trash.ComboBox1.Clear
    
    For i = 1 To cant_filas("barrido") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Barrido").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("barrido").Cells(i + 1, 10).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox4.Enabled = True
        Else
        Trash.CheckBox4.Enabled = False
    End If

'BASE OPERACIONES

    Visibles = 0
    For i = 1 To cant_filas("Base_OP") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Base_OP").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("Base_OP").Cells(i + 1, 4).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox7.Enabled = True
        Else
        Trash.CheckBox7.Enabled = False
    End If

'CORTE CÉSPED

    Visibles = 0
    'Trash.ListBox1.Clear
    For i = 1 To cant_filas("Corte_césped") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Corte_césped").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("corte_césped").Cells(i + 1, 5).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox2.Enabled = True
        Else
        Trash.CheckBox2.Enabled = False
    End If

'LAVADO ÁREAS
    Visibles = 0
    For i = 1 To cant_filas("Lavado_áreas") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Lavado_áreas").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("lavado_áreas").Cells(i + 1, 5).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox5.Enabled = True
        Else
        Trash.CheckBox5.Enabled = False
    End If

'LIMPIEZA PLAYAS
    Visibles = 0
    For i = 1 To cant_filas("Limpieza_playas") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Limpieza_playas").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("limpieza_playas").Cells(i + 1, 8).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox8.Enabled = True
        Else
        Trash.CheckBox8.Enabled = False
    End If

'PODA ARBOLES
    Visibles = 0
    For i = 1 To cant_filas("Poda_arboles") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("Poda_arboles").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("poda_arboles").Cells(i + 1, 4).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox3.Enabled = True
        Else
        Trash.CheckBox3.Enabled = False
    End If

'SDF
    Visibles = 0
    For i = 1 To cant_filas("SDF") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("SDF").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("SDF").Cells(i + 1, 45).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox6.Enabled = True
        Else
        Trash.CheckBox6.Enabled = False
    End If

'VEHÍCULOS R&T
    Visibles = 0
    For i = 1 To cant_filas("R&T") 'buscar si la hoja tiene el nombre de la empresa
        If Sheets("R&T").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        Trash.ComboBox1.AddItem (Sheets("R&T").Cells(i + 1, 4).Text)
        Visibles = Visibles + 1
        End If
    Next i
    
    If Visibles > 0 Then
        Trash.CheckBox1.Enabled = True
        Else
        Trash.CheckBox1.Enabled = False
    End If

End Sub

Sub ordenar_fechas()
     'Array
    Dim Arreglo() As String
    'Array temporal
    Dim TempArray() As String
    Dim x As Integer, x2 As Integer, Y As Integer
    Dim z As Integer, elemento As Variant
    Dim i As Integer, q As Integer
    Dim p As Integer, k As Integer
    Dim j As Integer
    
 q = 0 'limpiar contador

'CALCULO DIMENSION DEL ARRAY AUXILIAR'****************
 
 q = Trash.ComboBox1.ListCount

'CARGAR EL ARRAY AUXILIAR'-***************************
    ReDim Arreglo(q)
    For i = 0 To q - 1
    Arreglo(i) = Trash.ComboBox1.List(i)
    Next i

'BORRAR DATOS DUPLICADOS DEL COMBOBOX**************************************
    Trash.ComboBox1.Clear
    For i = LBound(Arreglo) To UBound(Arreglo)
          'Redimensionamos el Array temporal y preservamos el valor
          ReDim Preserve TempArray(i)
          'Asignamos al array temporal el valor del otro array
          TempArray(i) = Arreglo(i)
          
        Next
      
    For x = 0 To UBound(Arreglo)
        z = 0
         For Y = 0 To UBound(Arreglo)
            'Si el elemento del array es igual al array temporal
            If Arreglo(x) = TempArray(z) And Y <> x Then
                'Entonces Eliminamos el valor duplicado
                Arreglo(Y) = ""
            End If
            z = z + 1
        Next Y
    Next x
   For i = 0 To q - 1
   Next i
   
    'Recorremos el array. Para recorrer un array con ForEach
    'la variable de la colección debe ser de tipo Variant
    For Each elemento In Arreglo
        'Si el elemento es distinto de una cadena vacia
        'lo agegamos al combobox
        If elemento <> "" Then Trash.ComboBox1.AddItem (elemento)
    Next
'***********************************************************


End Sub

Public Function cant_filas(Nombre As String)
    Dim Contador As Boolean
    Dim D As Integer
    Dim i As Integer
    D = 0
    i = 2
    Contador = True
    
    While Contador = True
    If IsEmpty(Sheets(Nombre).Cells(i, 2)) Then
    Contador = False
    Else
    D = D + 1
    i = i + 1
    End If
    Wend
    cant_filas = D
End Function

Sub filtros_fechas()

    Dim i As Integer, j As Integer, Presentes As Integer
    'BARRIDO **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("barrido")
        If Sheets("barrido").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("barrido").Cells(i + 1, 10).Text 'valor en la base de datos
                If Sheets("barrido").Cells(i + 1, 10).Text = Trash.ListBox1.List(j) Then
                    'Trash.Label2.Caption = j 'indice de lista
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox4.Enabled = True
                    Else
                    Trash.CheckBox4.Enabled = False
                End If
                 
    'BASE OPERACIONES **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("Base_OP")
        'Trash.Label2.Caption = i
        If Sheets("Base_OP").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("Base_OP").Cells(i + 1, 4).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox7.Enabled = True
                    Else
                    Trash.CheckBox7.Enabled = False
                End If
                
    'Corte_césped **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("Corte_césped")
        'Trash.Label2.Caption = i
        If Sheets("Corte_césped").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("Corte_césped").Cells(i + 1, 5).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox2.Enabled = True
                    Else
                    Trash.CheckBox2.Enabled = False
                End If
                
        'Lavado_áreas **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("Lavado_áreas")
        'Trash.Label2.Caption = i
        If Sheets("Lavado_áreas").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("Lavado_áreas").Cells(i + 1, 5).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox5.Enabled = True
                    Else
                    Trash.CheckBox5.Enabled = False
                End If
                           
        'Limpieza_playas **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("Limpieza_playas")
        'Trash.Label2.Caption = i
        If Sheets("Limpieza_playas").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("Limpieza_playas").Cells(i + 1, 8).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox8.Enabled = True
                    Else
                    Trash.CheckBox8.Enabled = False
                End If
            
        'Poda_arboles **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("Poda_arboles")
        'Trash.Label2.Caption = i
        If Sheets("Poda_arboles").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("Poda_arboles").Cells(i + 1, 4).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox3.Enabled = True
                    Else
                    Trash.CheckBox3.Enabled = False
                End If
                
    'SDF **************************
        Presentes = 0
    For i = 1 To Control_Trash.cant_filas("SDF")
        'Trash.Label2.Caption = i
        If Sheets("SDF").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("SDF").Cells(i + 1, 45).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox6.Enabled = True
                    Else
                    Trash.CheckBox6.Enabled = False
                End If
         
    
    'R&T **************************
    Presentes = 0
    For i = 1 To Control_Trash.cant_filas("R&T")
        'Trash.Label2.Caption = i
        If Sheets("R&T").Cells(i + 1, 2) = Trash.ComboBox2.Text Then
            'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For j = 0 To Trash.ListBox1.ListCount - 1
                'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
                'Trash.Label4.Caption = Sheets("R&T").Cells(i + 1, 4).Text 'valor en la base de datos
                If Sheets("R&T").Cells(i + 1, 4).Text = Trash.ListBox1.List(j) Then
                   'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                    Presentes = Presentes + 1
                End If
                    'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
            Next j
        End If
    Next i
                If Presentes > 0 Then
                    Trash.CheckBox1.Enabled = True
                    Else
                    Trash.CheckBox1.Enabled = False
                End If
         
                
End Sub

Public Function Cont_Registros(Nombre As String, iFecha As Integer)

    Dim i As Integer, Y As Integer
    Dim Contador As Integer

    'CALCULAR REGISTROS
    For i = 1 To Control_Trash.cant_filas(Nombre)
        If Sheets(Nombre).Cells(i + 1, 2) = Trash.ComboBox2.Text Then
        'Trash.Label4.Caption = Trash.ComboBox2.Text '1er filtro
            For Y = 0 To Trash.ListBox1.ListCount - 1
        'Trash.Label1.Caption = Trash.ListBox1.ListCount 'tamaño de la lista
        'Trash.Label4.Caption = Sheets("barrido").Cells(i + 1, 10).Text 'valor en la base de datos
                If Sheets(Nombre).Cells(i + 1, iFecha).Text = Trash.ListBox1.List(Y) Then
        'Trash.Label2.Caption = j 'indice de lista
        'Trash.Label3.Caption = Trash.ListBox1.List(j) 'dato de la lista
                Contador = Contador + 1
                End If
            Next Y
        End If
    Next i
    
    Cont_Registros = Contador

End Function

Sub agregar_empresas()
 'Array
    Dim Arreglo() As String
    'Array temporal
    Dim TempArray() As String
    Dim x As Integer, x2 As Integer, Y As Integer
    Dim z As Integer, elemento As Variant
    Dim i As Integer, q As Integer
    Dim p As Integer, k As Integer
    Dim j As Integer
    
'************************ EMPRESAS ************************************************************
'CALCULO DIMENSION DEL ARRAY AUXILIAR'****************
 For j = 2 To Sheets.Count
 i = 2
 While Not IsEmpty(Sheets(j).Cells(i, 2))
   q = q + 1 'dimension del array
   i = i + 1 'contador de filas
  Wend
 Next j

'CARGAR EL ARRAY AUXILIAR'-***************************
    ReDim Arreglo(q)
    i = 0
    For p = 2 To Sheets.Count
        j = 2
        While Not IsEmpty(Sheets(p).Cells(j, 2))
        Arreglo(i) = Sheets(p).Cells(j, 2).Text
        j = j + 1
        i = i + 1
        Wend
    Next p
'BORRAR DATOS DUPLICADOS DEL COMBOBOX*****************
    For i = LBound(Arreglo) To UBound(Arreglo)
          'Redimensionamos el Array temporal y preservamos el valor
          ReDim Preserve TempArray(i)
          'Asignamos al array temporal el valor del otro array
          TempArray(i) = Arreglo(i)
    Next
      
    For x = 0 To UBound(Arreglo)
        z = 0
         For Y = 0 To UBound(Arreglo)
            'Si el elemento del array es igual al array temporal
            If Arreglo(x) = TempArray(z) And Y <> x Then
                'Entonces Eliminamos el valor duplicado
                Arreglo(Y) = ""
            End If
            z = z + 1
        Next Y
    Next x
   For i = 0 To q - 1
   Next i
   
    'Recorremos el array. Para recorrer un array con ForEach
    'la variable de la coleccióndebe ser de tipo Variant
    For Each elemento In Arreglo
        'Si el elemento es distinto de una cadena vacia
        'lo agegamos al combobox
        If elemento <> "" Then Trash.ComboBox2.AddItem (elemento)
    Next
End Sub
