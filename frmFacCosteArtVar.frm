VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacCosteArtVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar precio coste art. varios en pedido"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17580
   ClipControls    =   0   'False
   Icon            =   "frmFacCosteArtVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   17580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   12840
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   2
      Text            =   "dto2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   1
      Text            =   "dto"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   9120
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "neto"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   5835
      Width           =   3255
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   15120
      TabIndex        =   3
      Top             =   6000
      Width           =   1135
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   16320
      TabIndex        =   4
      Top             =   6000
      Width           =   1135
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   16320
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1320
      Top             =   6120
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacCosteArtVar.frx":000C
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIndicador 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   13680
      TabIndex        =   11
      Top             =   480
      Width           =   2355
   End
   Begin VB.Label Label22 
      Caption         =   "Leyendo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacCosteArtVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean
Dim cad As String

Dim N As Integer


Private Sub cmdAceptar_Click()
  'TotalLineas llevo
        
    '     importeb1 PrVenta,importeb2 PrCompra
    cad = " importeb1>0 and importeb2=0 and codusu"
    cad = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", cad, CStr(vUsu.Codigo))
    N = Val(cad)
    cad = ""
    If N > 0 Then cad = "Existen " & N & " linea" & IIf(N > 1, "s", "") & " pendientes de asignar coste" & vbCrLf & vbCrLf
    cad = cad & "¿Desea continuar asignado los costes indicados ?"
        
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
    Set miRsAux = New ADODB.Recordset
    
        
    
    CadenaDesdeOtroForm = "OK"
    Me.Tag = 0
    Unload Me
    
    
    
Error1:
    
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
   
    
    If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    CadenaDesdeOtroForm = ""
    Me.Tag = 0
    Unload Me
  
   
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
       'txtAux.SelStart = Len(Me.txtAux.Text)
       
       txtAux(0).SetFocus
       
       txtAux(0).SelStart = 0
       txtAux(0).SelLength = Len(Me.txtAux(0).Text)
       txtAux(0).Refresh
       
    End If
End Sub

Private Sub Form_Activate()


    
    If PrimeraVez Then
        Label22.Caption = "Actualizar costes artículos varios"
         PrimeraVez = False
         Iniciar
             
             
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
 
    
    Me.Tag = 1 'NO se puede cerrar mas que de boton

    DataGrid1.Width = Me.Width - 400
    Me.cmdCancelar.Left = Me.Width - 1365
    Me.cmdAceptar.Left = Me.Width - 2565
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
     gridCargado = False
    
    

   
    
    
    PrimeraVez = True
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()


On Error GoTo ECarga

   
    
    
                    ' codpr   nompr   art     nomart
    cad = "SELECT campo1,codigo1,nombre1,nombre2,nombre3, importe1 cantidad, importe2 prventa,"
                  '  precioar        dto1          dto2
    cad = cad & " importeb3 precio,importe4 Dto_1 ,importe5 Dto_2 "
    '              precio vtta(aplia dtos)       dtcomp2   precoste   marge
    cad = cad & " , importeb2 PrCompra ,porcen1 dto1 ,porcen2 dto2 ,importeb5 Margen  , importeb4 Coste"
    cad = cad & " FROM tmpinformes WHERE codusu=" & vUsu.Codigo & " ORDER BY campo1"
    
    gridCargado = False
    data1.ConnectionString = conn
    data1.RecordSource = cad
    data1.CursorType = adOpenDynamic
    data1.LockType = adLockPessimistic
    data1.Refresh
    
    
    
        
    DataGrid1.RowHeight = 360
    DataGrid1.Columns(0).visible = False
    
    DataGrid1.Columns(1).Caption = "Prov."
    DataGrid1.Columns(1).Width = 800
    

        
    DataGrid1.Columns(2).Caption = "Proveedor"
    DataGrid1.Columns(2).Width = 2800
    
    DataGrid1.Columns(3).Caption = "Art"
    DataGrid1.Columns(3).Width = 1600
    'DataGrid1.Columns(3).NumberFormat = FormatoImporte
    'DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(4).Caption = "Descripcion"
    DataGrid1.Columns(4).Width = 4200
    DataGrid1.Columns(5).visible = False  'cantidad
        DataGrid1.Columns(6).visible = False  'precio venta (ya aplicados dtos)
            
    For NumRegElim = 7 To 13
        'canti     prepvp   precompr  dtoco1 dtcomp2   precoste   marge
        Debug.Print DataGrid1.Columns(NumRegElim).Caption
        If NumRegElim = 7 Or NumRegElim = 10 Then
            'Precio
            
            DataGrid1.Columns(NumRegElim).Width = 1500
            DataGrid1.Columns(NumRegElim).NumberFormat = FormatoPrecio
        
        Else
            DataGrid1.Columns(NumRegElim).Width = IIf(NumRegElim = 5, 1000, 800)
            DataGrid1.Columns(NumRegElim).NumberFormat = FormatoImporte
        End If
        DataGrid1.Columns(NumRegElim).Alignment = dbgRight
        
        
    Next
    DataGrid1.Columns(NumRegElim).visible = False  'coste
    
    
    
    DataGrid1.ScrollBars = dbgAutomatic
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For N = 0 To 2
            txtAux(N).visible = visible
            Me.txtAux(N).Locked = True
        Next
    Else
        DeseleccionaGrid Me.DataGrid1
        
        alto = DataGrid1.Top + 220
        If DataGrid1.Row >= 0 Then alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        
        
        'Fijamos altura y posición Top
        '-------------------------------
        For N = 0 To 2
            txtAux(N).Top = alto

            
            txtAux(N).Left = DataGrid1.Columns(N + 10).Left + 130 'codalmac
            txtAux(N).Width = DataGrid1.Columns(N + 10).Width - 10
            Me.txtAux(N).Locked = False
            txtAux(N).visible = visible
            txtAux(N).Text = DataGrid1.Columns(N + 10).Text
        Next
            PonerFoco txtAux(0)
    End If

    
    
End Sub







Private Sub Iniciar()
Dim C1 As Currency
    
    CargaGrid
'    data1.Recordset.MoveFirst
'    Do
'        C1 = C1 + data1.Recordset!cantidad
'        data1.Recordset.MoveNext
'    Loop Until data1.Recordset.EOF
'    data1.Recordset.MoveFirst
'    Text1.Text = Format(C1, FormatoCantidad)
'    Text1.Tag = C1
'
    
    BotonModificar
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    txtAux(Index).SelStart = 0
    txtAux(Index).SelLength = Len(txtAux(Index).Text)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        'ModificarExistencia

            Me.txtAux(0).SelStart = 0
            Me.txtAux(0).SelLength = Len(Me.txtAux(0).Text)

    End If
    Screen.MousePointer = vbHourglass
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If

        Case 40 'Desplazamiento Flecha Hacia Abajo
                Screen.MousePointer = vbHourglass
                PasarSigReg
                
                
    End Select
    
     If KeyCode = 38 Or KeyCode = 40 Then
        'ModificarExistencia
            PonerFoco txtAux(0)
            Me.txtAux(0).SelStart = 0
            Me.txtAux(0).SelLength = Len(Me.txtAux(0).Text)

    End If
    
EKeyD:
    If Err.Number <> 0 Then Err.Clear
    Screen.MousePointer = vbDefault
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        If Index = 2 Then
            KeyAscii = 0
            ModificarExistencia
        
            PasarSigReg
        Else
            PonerFoco txtAux(Index + 1)
        End If
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Importe As Currency
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            
            If Not PonerFormatoDecimal(txtAux(Index), IIf(Index = 0, 2, 4)) Then .Text = ""
        End If
        
        If .Text = "" Then .Text = DataGrid1.Columns(7 + Index)
    End With

End Sub






Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
       
    Modo = Kmodo
  
    

    b = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera b
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
      
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub







Private Sub BotonModificar()
    If data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    'txtAux.Text = Trim(txtAux.Text)

    DatosOk = True
        
    For N = 0 To 2
        If txtAux(N).Text <> "" Then
            If Not EsNumerico(txtAux(N).Text) Then
                DatosOk = False
            Else
                If N > 0 Then
                    
                    'Si lanza el evento antes del losfocus no hace el formateo
                    If N = 2 Then If InStr(1, txtAux(N).Text, ".") > 0 Then txtAux(N).Text = Replace(txtAux(N).Text, ".", ",")
                    
                    If ImporteFormateado(txtAux(N).Text) >= 100 Then DatosOk = False
                End If
            End If
        Else
            DatosOk = False
        End If
    Next
End Function


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    'PonerOpcionesMenuGeneral Me
End Sub




Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        PonerFoco Me.txtAux(0)
    ElseIf DataGrid1.Bookmark = data1.Recordset.RecordCount Then
       PonerFocoBtn cmdAceptar
    End If
    

End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
    Screen.MousePointer = vbHourglass
    lblIndicador.ForeColor = vbRed
    lblIndicador.Caption = "Modificando"
    lblIndicador.Refresh
    If DatosOk Then
         Screen.MousePointer = vbHourglass
        If ActualizarExistencia() Then
            
            lblIndicador.ForeColor = vbBlack
            lblIndicador.Refresh
            
            NumReg = data1.Recordset.AbsolutePosition
            CargaGrid
            
                    
            If NumReg < data1.Recordset.RecordCount Then
                data1.Recordset.Move NumReg - 1
            Else
                data1.Recordset.MoveLast
            End If
        End If

            
            
            ModificarExistencia = True
    Else
            ModificarExistencia = False
  
    End If
    lblIndicador.Caption = ""
    lblIndicador.Refresh
    Screen.MousePointer = vbDefault
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim Importe As Currency
Dim Impo2 As Currency
    On Error GoTo EActualizar
        'importeb2 PrCompra ,porcen1 dto1 ,porcen2 dto2 , importeb4 Coste,importeb5 Margen"
        Screen.MousePointer = vbHourglass
        
        cad = "importeb2 |porcen1 |porcen2 |"
        SQL = ""
        For N = 0 To 2
            SQL = SQL & ", " & RecuperaValor(cad, N + 1) & " = "
            SQL = SQL & DBSet(Me.txtAux(N), "N")
        Next
        
        
        Importe = ImporteFormateado(txtAux(1).Text) 'descuento
        Importe = (100 - Importe) / 100
        'Van sobre resto los descuentos
        Importe = Round(ImporteFormateado(txtAux(0).Text) * Importe, 4)
        Impo2 = ImporteFormateado(txtAux(2).Text) 'descuento
        Impo2 = (100 - Impo2) / 100
        Importe = Round(Impo2 * Importe, 4)
        Impo2 = 1  'data1.Recordset!cantidad
        Importe = Round(Importe / Impo2, 2)  'unitaro REDONDEADP a 2
        
        Impo2 = data1.Recordset!prventa   'venta unitario
        Impo2 = Impo2 - Importe ' beneficio  unitario vta  -compr
        If Importe = 0 Then
            Impo2 = 0
        Else
            Impo2 = Round((Impo2 / Importe) * 100, 2) 'margen% sobre compra
        End If
        SQL = SQL & ", importeb4 = " & DBSet(Importe, "N")  'prcoste ud
        SQL = SQL & ", importeb5 = " & DBSet(Impo2, "N")  'prcoste ud
        'quitamos la primera coma
        SQL = Mid(SQL, 2)
        cad = "UPDATE tmpinformes SET " & SQL & " WHERE codusu=" & vUsu.Codigo & " AND campo1= " & DataGrid1.Columns(0)
        conn.Execute cad
        
        
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function



