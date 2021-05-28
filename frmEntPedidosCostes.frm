VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEntPedidosCostes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar costes articulos varios"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   ClipControls    =   0   'False
   Icon            =   "frmEntPedidosCostes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
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
      Height          =   320
      Left            =   10560
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   1
      Text            =   "existencia"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5355
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12000
      TabIndex        =   0
      Top             =   5400
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1320
      Top             =   5640
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
      Bindings        =   "frmEntPedidosCostes.frx":000C
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.Label Label22 
      Caption         =   "Leyendo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   5520
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
      TabIndex        =   2
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmEntPedidosCostes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Albaran As String

Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean
Dim cad As String


Private Sub cmdAceptar_Click()
  'TotalLineas llevo
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select slialb.numlinea,slialb.codartic,slialb.nomartic,substring(ampliaci,1,25),precioar,precoste from"
    cad = cad & " slialb,sartic where slialb.codartic=sartic.codartic and slialb.codtipom=" & DBSet(Trim(Mid(Albaran, 1, 3)), "T")
    cad = cad & " AND slialb.numalbar=" & Trim(Mid(Albaran, 4)) & " and sartic.artvario=1 and coalesce(precoste,0)=0"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "": CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux!precoste) Then
            cad = cad & miRsAux!codArtic & "  " & miRsAux!NomArtic & vbCrLf
        Else
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux!codArtic & "  " & miRsAux!NomArtic & vbCrLf
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If cad <> "" Then
        cad = "Falta asignar coste: " & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
        PonerFoco txtAux
        Exit Sub
    End If
    
    If CadenaDesdeOtroForm <> "" Then
        CadenaDesdeOtroForm = "Coste asignado a CERO: " & vbCrLf & vbCrLf & CadenaDesdeOtroForm
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        CadenaDesdeOtroForm = ""
    End If
        
    
        
    CadenaDesdeOtroForm = "OK"
    Me.Tag = 0
    Unload Me
    
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        CadenaDesdeOtroForm = ""
    End If
        
    
End Sub

Private Sub cmdCancelar_Click()
   
    'Me.Tag = 0
    Unload Me
  
   
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
       'txtAux.SelStart = Len(Me.txtAux.Text)
       
       txtAux.SetFocus
       
       txtAux.SelStart = 0
       txtAux.SelLength = Len(Me.txtAux.Text)
       txtAux.Refresh
       
    End If
End Sub

Private Sub Form_Activate()


    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        
         PrimeraVez = False
         Iniciar
         
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
 
    txtAux.Left = 0
    Me.Tag = 1 'NO se puede cerrar mas que de boton

    DataGrid1.Width = Me.Width - 400
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    gridCargado = False
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()


On Error GoTo ECarga

   
    
    cad = "Select slialb.numlinea,slialb.codartic,slialb.nomartic,substring(ampliaci,1,25),precioar,precoste from"
    cad = cad & " slialb,sartic where slialb.codartic=sartic.codartic and slialb.codtipom=" & DBSet(Trim(Mid(Albaran, 1, 3)), "T")
    cad = cad & " AND slialb.numalbar=" & Trim(Mid(Albaran, 4)) & " and sartic.artvario=1"
    

    data1.ConnectionString = conn
    data1.RecordSource = cad
    data1.CursorType = adOpenDynamic
    data1.LockType = adLockPessimistic
    data1.Refresh
   
    
    
        

    DataGrid1.Columns(0).visible = False
    
    DataGrid1.Columns(1).Caption = "Codigo"
    DataGrid1.Columns(1).Width = 1700
    
    
        
    DataGrid1.Columns(2).Caption = "Articulo"
    DataGrid1.Columns(2).Width = 4800
    
    DataGrid1.Columns(3).Caption = "Ampliacion"
    DataGrid1.Columns(3).Width = 2800
    
    
    
    DataGrid1.Columns(4).Caption = "Precio vta"
    DataGrid1.Columns(4).Width = 1350
    DataGrid1.Columns(4).NumberFormat = FormatoImporte
    DataGrid1.Columns(4).Alignment = dbgRight
            
    DataGrid1.Columns(5).Caption = "COSTE"
    DataGrid1.Columns(5).Width = 1350
    DataGrid1.Columns(5).NumberFormat = FormatoCantidad
    DataGrid1.Columns(5).Alignment = dbgRight
            
    
    
    
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
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = DBLet(data1.Recordset!precoste, "N")
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        If txtAux.Left = 0 Then
            txtAux.Left = DataGrid1.Columns(5).Left + 130
            txtAux.Width = DataGrid1.Columns(5).Width - 10
        End If
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
      '  txtAux.SelStart = 0
       ' txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub






Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = 1 Then Cancel = 1        'o aceptar o cancelar
    
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub Iniciar()
   CargaGrid
    
    BotonModificar
    
End Sub

Private Sub txtAux_GotFocus()
    txtAux.SelStart = 0
    txtAux.SelLength = Len(txtAux.Text)
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                     Me.txtAux.SelStart = 0
                Me.txtAux.SelLength = Len(Me.txtAux.Text)
                'txtaux.Refresh
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        ModificarExistencia
        
        PasarSigReg
        
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
Dim Importe As Currency
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then
            .Text = "0,00"
        Else
                If Not EsNumerico(.Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    On Error Resume Next
                    .Text = "0,00"
                    PonerFoco txtAux
                    Exit Sub
                End If
                
                
                'Es numerico
                cad = TransformaPuntosComas(.Text)
                If CadenaCurrency(cad, Importe) Then .Text = Format(Importe, "0.00")
                    
                
        
        End If
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
    PonerFoco txtAux
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux.Text = Trim(txtAux.Text)
    DatosOk = False
    If txtAux.Text <> "" Then
        If EsNumerico(txtAux.Text) Then DatosOk = True
    End If
End Function


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    
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
        PonerFoco Me.txtAux
    ElseIf DataGrid1.Bookmark = data1.Recordset.RecordCount Then
       PonerFocoBtn cmdAceptar
    End If
    

End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long


    If DatosOk Then
        
        If ActualizarExistencia() Then
            
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
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim cantidad As Currency


    On Error GoTo EActualizar

    
        


    cantidad = TransformaPuntosComas(txtAux.Text)
    
    'If Cantidad < 0 Then Err.Raise 513, , "No se permiten negativos"
    
    'If cantidad > data1.Recordset!disponible Then Err.Raise 513, , "Cantidad disponible:" & data1.Recordset!disponible
    SQL = "N"
    If IsNull(data1.Recordset!precoste) Then
        SQL = ""
    Else
        If cantidad <> data1.Recordset!precoste Then SQL = ""
    End If
        'Actualizar la Tabla: sinven con la cantidad introducida
        '-------------------------------------------------------
'
     SQL = "UPDATE slialb Set precoste = " & TransformaComasPuntos(CStr(cantidad))
     SQL = SQL & " WHERE numlinea = " & data1.Recordset!numlinea
     SQL = SQL & " AND slialb.codtipom=" & DBSet(Trim(Mid(Albaran, 1, 3)), "T")
     SQL = SQL & " AND slialb.numalbar=" & Trim(Mid(Albaran, 4))
     conn.Execute SQL
        
    
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function



