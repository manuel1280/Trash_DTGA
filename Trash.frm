VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Trash 
   Caption         =   "Trash_DTGA"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   OleObjectBlob   =   "Trash.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Trash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const WS_MINIMIZEBOX As Long = &H20000
'Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16)

Option Explicit

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CheckBox5_Click()

End Sub

Private Sub ComboBox1_AfterUpdate()
  
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox1_Click()
    Me.ListBox1.AddItem (Trash.ComboBox1.Text) 'añadir al hacer click
    Control_Trash.filtros_fechas
End Sub

Private Sub ComboBox1_DropButtonClick()

End Sub
Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ComboBox2_Change()
Me.ComboBox1.Clear
Me.ListBox1.Clear
End Sub

Private Sub ComboBox2_Click()
    Control_Trash.filtrar_empresa
    Control_Trash.ordenar_fechas
    
End Sub

Private Sub CommandButton1_Click()
    Dim numHojas As Integer
    numHojas = Sheets.Count
     
If Not IsEmpty(Me.ComboBox2.Text) And Me.ListBox1.ListCount > 0 Then
 If numHojas <= 10 Then
 
    Application.ScreenUpdating = False
    If Me.CheckBox4.Enabled = True And Me.CheckBox4.Value = True Then
    Barrido.Barrido
    End If
    
    If Me.CheckBox3.Enabled = True And Me.CheckBox3.Value = True Then
    Poda_arboles.Poda
    End If
    
    If Me.CheckBox2.Enabled = True And Me.CheckBox2.Value = True Then
    Corte_cesped.Corte_cesped
    End If
    
    If Me.CheckBox1.Enabled = True And Me.CheckBox1.Value = True Then
    RYT.R_y_T
    End If
    
    If Me.CheckBox8.Enabled = True And Me.CheckBox8.Value = True Then
    Limpieza_Playas.Playa
    End If
    
    If Me.CheckBox7.Enabled = True And Me.CheckBox7.Value = True Then
    Base_Ope.Base_operaciones
    End If
  
    If Me.CheckBox6.Enabled = True And Me.CheckBox6.Value = True Then
    SDF.SDF
    End If
    
    If Me.CheckBox5.Enabled = True And Me.CheckBox5.Value = True Then
    Lavado_areaspublicas.Lavado
    End If

    
  Application.ScreenUpdating = True
  
 
   Else
   MsgBox "Hojas adicionales generadas, oprima Borrar", vbOK
   End If
 Else
 MsgBox "Seleccione nombre de la empresa y fecha a generar", vbOK
 End If

    
End Sub

Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False
    ActiveWorkbook.RefreshAll
    Trash.ComboBox2.Clear
    Control_Trash.agregar_empresas
    Control_Trash.filtrar_empresa
    Control_Trash.filtros_fechas
    Application.ScreenUpdating = False
End Sub

Private Sub CommandButton3_Click()
    Dim Eliminar As Integer
    Eliminar = Sheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    While Eliminar > 10
        Sheets(Eliminar).Select
        ActiveWindow.SelectedSheets.Delete
        Eliminar = Eliminar - 1
    Wend
    
    Sheets("INICIO").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub CommandButton4_Click()
    Dim numHojas As Integer
    numHojas = Sheets.Count
    
 'If numHojas <= 10 Then
 
    Application.ScreenUpdating = False
    
    If Me.CheckBox1.Enabled = True And Me.CheckBox1.Value = True Then
    Análisis.RYT
    End If
    
    If Me.CheckBox3.Enabled = True And Me.CheckBox3.Value = True Then
    Análisis.Poda_arboles
    End If
    
    If Me.CheckBox6.Enabled = True And Me.CheckBox6.Value = True Then
    Análisis.DF
    End If
    
    If Me.CheckBox2.Enabled = True And Me.CheckBox2.Value = True Then
    Análisis.Corte_cesped
    End If
    
    If Me.CheckBox7.Enabled = True And Me.CheckBox7.Value = True Then
    Análisis.BOP
    End If
    
    Application.ScreenUpdating = True
  'Else
  'MsgBox "Hojas adicionales generadas, oprima Borrar", vbOK
  'End If

End Sub

Private Sub ListBox1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ListBox1.RemoveItem (ListBox1.ListIndex) 'borrar al hacer doble click
    Control_Trash.filtros_fechas
    If Me.ListBox1.ListCount = 0 Then
        Control_Trash.filtrar_empresa
        Control_Trash.ordenar_fechas
    End If
End Sub
Private Sub ToggleButton1_Click()
    
If Me.ToggleButton1.Value = True Then
    
    Me.CommandButton1.Enabled = False
    Me.CommandButton4.Enabled = True
    Me.ListBox1.Visible = False
    Me.ComboBox1.Enabled = False
    Control_Trash.filtrar_empresa
    'If Me.CheckBox4.Enabled = True And Me.CheckBox4.Value = True Then
    'Me.CheckBox4.Enabled = False
    'End If
    
    'If Me.CheckBox5.Enabled = True And Me.CheckBox5.Value = True Then
    'Me.CheckBox5.Enabled = False
    'End If
    
    'If Me.CheckBox8.Enabled = True And Me.CheckBox8.Value = True Then
    'Me.CheckBox8.Enabled = False
    'End If
    Else
    Me.CommandButton1.Enabled = True
    Me.CommandButton4.Enabled = False
    Me.ListBox1.Visible = True
    Me.ComboBox1.Enabled = True
    Control_Trash.filtros_fechas

End If
    
    
End Sub

Private Sub UserForm_Activate()
    Control_Trash.agregar_empresas
         
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()

    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    If Application.Version < 9 Then
        lngMyHandle = FindWindow("THUNDERXFRAME", Me.Caption)
    Else
        lngMyHandle = FindWindow("THUNDERDFRAME", Me.Caption)
    End If
    lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
    lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX 'Or WS_MAXIMIZEBOX
    SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle

End Sub
