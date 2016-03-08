VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepCargarNSerie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Introducir Nº Series"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8520
   ClipControls    =   0   'False
   Icon            =   "frmRepCargarNSerie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtauz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkConsecutivos 
      Caption         =   "Consecutivos"
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   5355
      Width           =   2775
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7035
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7035
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cargar Series"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRepCargarNSerie.frx":000C
      Height          =   4485
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7911
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   8
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnCargar 
         Caption         =   "&Cargar Series"
         HelpContextID   =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmRepCargarNSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================== PUBLICAS =======================
Public Event CargarNumSeries()
Public DeVentas As Boolean 'Si llamo a cargar Nº Series desde VEntas, sino se llamo desde compras
Public NumAlb As String

'=========================== LOCALES =======================
Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer
Dim NombreTabla As String
Dim Ordenacion As String


Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form


Dim NumeroMantenimiento As String


'=========================== PROCEDIMIENTOS =======================

Private Sub chkConsecutivos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Integer
On Error GoTo Error1

    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 4 'Modificar Num Serie (Introducir Valores a Cargar)
        If Me.chkConsecutivos.Value = 1 Then
        'Cargar la tabla temporal con los Nº de Serie consecutivos
            If Trim(txtAux.Text) <> "" Then
                CargarNSeriesConsecutivos txtAux.Text
            Else
                MsgBox "Introduzca el primer Nº de Serie para obtener los sucesivos.", vbInformation
                PonerFoco txtAux
            End If
        Else 'NO consecutivos, solo carga 1 nº serie
            If DatosOk(txtAux) Then
                If ActualizarNumSerie(txtAux, txtauz.Text) Then
                    'Enero 2015
                    'Si ha puesto mantenimiento, lo guarda
                    If DeVentas Then NumeroMantenimiento = Trim(txtauz.Text)
                    
    '                TerminaBloquear
                    CargaTxtAux False, False
                    NumReg = Data1.Recordset.AbsolutePosition
                    If NumReg <= 0 Then NumReg = 1
                    CargaGrid True
                    
                    
                    
                    If Data1.Recordset.RecordCount >= NumReg Then Data1.Recordset.Move NumReg - 1
                    
                        If Not Data1.Recordset.EOF Then
                            Data1.Recordset.MoveNext
                            If Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
                            CargaTxtAux True, True
                            PonerFoco txtAux
                        Else
                            PonerModo 0
                        End If
                    
                End If
            Else
                txtAux.Text = ""
                PonerFoco txtAux
            End If
            
        End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    If Modo = 4 Then 'Modificar
          CargaTxtAux False, False
          PonerModo 0
    End If
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then 'Modo4: Modificar
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    PonerFoco txtAux
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 4 'Modificar
        .Buttons(2).Image = 21 'Cargar Nº Series
        .Buttons(4).Image = 15 'Salir
    End With
    
    PulsadoSalir = False
    PrimeraVez = True
    DataGrid1.ClearFields
    
    NombreTabla = "tmpnseries"
    Ordenacion = " ORDER BY codusu, codartic,numlinealb, numlinea"
    
'    PonerModo 4
    CargaGrid True
    BotonModificar
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim i As Byte
Dim SQL As String
On Error GoTo ECarga

    gridCargado = False
    
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    PrimeraVez = False
        
    'Cod. Artic
    DataGrid1.Columns(0).visible = False
    
    DataGrid1.Columns(1).Caption = "Articulo"
    DataGrid1.Columns(1).Width = 1500

    DataGrid1.Columns(2).Caption = "Desc. Articulo"
    DataGrid1.Columns(2).Width = 3200
       
    'Nº Serie
    DataGrid1.Columns(3).Caption = "Nº Serie"
    DataGrid1.Columns(3).Width = 1500
        
        'Nº Serie
    DataGrid1.Columns(4).Caption = "Mantenimiento"
    DataGrid1.Columns(4).Width = 1290
    DataGrid1.Columns(4).visible = DeVentas
    
    DataGrid1.Columns(5).visible = False
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    DataGrid1.ScrollBars = dbgVertical
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    On Error GoTo ErrCarga
    
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtauz.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            If Data1.Recordset.EOF Then
                'Nuevo 22 Junio 2011
                txtAux.Text = ""
                txtauz.Text = ""
            Else
                txtAux.Text = DBLet(Data1.Recordset!numSerie, "T")
                txtAux.Locked = False
                txtauz.Text = DBLet(Data1.Recordset!NumMante, "T")
                txtauz.Locked = False
                
                If DeVentas Then
                    If Trim(txtauz.Text) = "" Then txtauz.Text = NumeroMantenimiento
                End If
            End If
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        txtauz.Top = alto
        txtauz.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(3).Left + 130 'Nº Serie
        txtAux.Width = DataGrid1.Columns(3).Width - 10
        
        txtauz.Left = DataGrid1.Columns(4).Left + 130 'Nº Serie
        txtauz.Width = DataGrid1.Columns(4).Width - 10
        
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux.visible = visible
    txtauz.visible = visible And DeVentas
    
    ConseguirFoco txtAux, 3
    'PonerFoco txtAux
    Exit Sub
    
ErrCarga:
    MuestraError Err.Number, "Cargar txtAux.", Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub


Private Sub mnCargar_Click()
Dim SQL As String
Dim RStmp As ADODB.Recordset
Dim cadExisten As String
Dim cadExisten2 As String
Dim devuelve As String
Dim b As Boolean

    b = True
    'Antes de salir de ventana de carga de Nº series comprobar que en la tabla TMP
    'se han cargado todos los Nº de Serie
    SQL = "select numserie from tmpnseries where (numserie='' or numserie=' ') AND codusu=" & vUsu.Codigo
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RStmp.EOF Then
        MsgBox " Debe Introducir todos los Nº de Serie antes de Salir.", vbExclamation
        b = False
    End If
    RStmp.Close
    Set RStmp = Nothing
    If Not b Then Exit Sub
    
    'Comprobar que los Nº de Serie cargados en la temporal no esten asignados ya
    cadExisten = ""
    cadExisten2 = ""
    SQL = "SELECT numserie, codartic FROM tmpnseries "
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo
    SQL = SQL & " ORDER BY codartic "
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If DeVentas Then 'Se llamo desde VENTAS
        'Comprobar que los Nº de Serie cargados en la temporal no esten asignados ya
        'a un albaran de VENTA antes de cargarlos
        While Not RStmp.EOF 'para cada Nº Serie en la TMP
            devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numalbar", "numserie", RStmp!numSerie, "T", , "codartic", RStmp!codArtic, "T")
            If (devuelve <> "") And (Val(devuelve) <> Val(NumAlb)) Then 'Existe ya
                If cadExisten = "" Then
                    cadExisten = RStmp!numSerie
                Else
                    cadExisten = cadExisten & ", " & RStmp!numSerie
                End If
            End If
            
            '==== Laura 17/01/2007
            'Comprobar que los Nº de Serie cargados en la temporal no esten asignados ya
            'a una factura de VENTA antes de cargarlos
            devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numfactu", "numserie", RStmp!numSerie, "T", , "codartic", RStmp!codArtic, "T")
            If (devuelve <> "") Then  'Existe ya
                If cadExisten2 = "" Then
                    cadExisten2 = RStmp!numSerie
                Else
                    cadExisten2 = cadExisten2 & ", " & RStmp!numSerie
                End If
            End If
            '====
            
            RStmp.MoveNext
        Wend
        If cadExisten <> "" Then 'Se han encontrado Nº Series ya asignados a ALBaran
            cadExisten = "Los siguientes Nº Serie ya estan asignados a un Albaran de VENTA:" & vbCrLf & vbCrLf & cadExisten
            cadExisten = cadExisten & vbCrLf & vbCrLf & "Introduzca otro Nº."
            MsgBox cadExisten, vbExclamation
            b = False
        End If
        
        If cadExisten2 <> "" Then 'Se han encontrado Nº Series ya asignados a Facturas
            cadExisten2 = "Los siguientes Nº Serie ya estan asignados a una Factura de VENTA:" & vbCrLf & cadExisten2
            cadExisten2 = cadExisten2 & vbCrLf & vbCrLf & "Introduzca otro Nº."
            MsgBox cadExisten2, vbExclamation
            b = False
        End If
        
        
    Else 'Se llamo desde COMPRAS
        'Comprobar que los Nº de Serie cargados en la temporal no esten asignados ya
        'a un albaran de COMPRA antes de cargarlos en la tabla "sserie"
        While Not RStmp.EOF
            devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numalbpr", "numserie", RStmp!numSerie, "T", , "codartic", RStmp!codArtic, "T")
            If devuelve <> "" And devuelve <> NumAlb Then 'Existe ya
                If cadExisten = "" Then
                    cadExisten = RStmp!numSerie
                Else
                    cadExisten = cadExisten & ", " & RStmp!numSerie
                End If
            End If
            RStmp.MoveNext
        Wend
        If cadExisten <> "" Then 'Se han encontrado Nº Series ya asignados a otros ALBaranes
            cadExisten = "Los siguientes Nº Serie ya estan asignados a un Albaran de COMPRA:" & vbCrLf & vbCrLf & cadExisten
            cadExisten = cadExisten & vbCrLf & vbCrLf & "Introduzca otro Nº."
            MsgBox cadExisten, vbExclamation
           b = False
        End If
        
        'Comprobar que los Nº de Serie introducidos no esten repetidos en la tabla temporal
        
    End If
    RStmp.Close
    Set RStmp = Nothing
    If Not b Then Exit Sub
    
    'Cargar las lineas de las plantilla como lineas de la Oferta y Salir (Volver a Mto Ofertas)
    PulsadoSalir = True
    Unload Me
    RaiseEvent CargarNumSeries
End Sub


Private Sub mnModificar_Click()
    BotonModificar
End Sub




Private Sub mnSalir_Click()
    PulsadoSalir = True
    If MsgBox("Va a salir sin Introducir los Nº de Serie." & vbCrLf & "¿Desea continuar?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload Me
End Sub


Private Sub txtAux_GotFocus()
    ConseguirFoco txtAux, 3
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
    gridCargado = False
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                'If DataGrid1.Row > 0 Then
                '    DataGrid1.Row = DataGrid1.Row - 1
                '    CargaTxtAux True, True
                'End If
                Data1.Recordset.MovePrevious
                If Data1.Recordset.BOF Then Data1.Recordset.MoveNext
                gridCargado = True
                CargaTxtAux True, True
                
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                'If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                '    DataGrid1.Row = DataGrid1.Row + 1
                '    CargaTxtAux True, True
                'End If
                
                'ENERO 2009
                Data1.Recordset.MoveNext
                If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
                gridCargado = True
                CargaTxtAux True, True
                
                
    End Select
     gridCargado = True
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    If KeyAscii = 27 Then cmdCancelar_Click 'ESC
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Modificar
            mnModificar_Click
        Case 2 'Cargar Plantillas y Salir
            mnCargar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'MODIFICAR
    b = (Modo = 4)
    Me.cmdAceptar.visible = b
    Me.cmdCancelar.visible = b
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnModificar.Enabled = Not b
    
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

     SQL = "select numlinea," & NombreTabla & ".codartic, nomartic," & NombreTabla & ".numserie , " & NombreTabla & ".nummante  "
     'JUNIO 2010
     'TEINSA. Puede ser que haya dos lineas con el mismo articulo y distintos num serie
     SQL = SQL & " ,numlinealb"
     SQL = SQL & " FROM " & NombreTabla & " INNER JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic "
     SQL = SQL & " WHERE codusu=" & vUsu.Codigo
     SQL = SQL & Ordenacion
    
     MontaSQLCarga = SQL
End Function


Private Sub BotonModificar()
    PonerModo 4
    'CargaGrid True
    CargaTxtAux True, True
    PonerFoco txtAux
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarNumSerie(nSerie As String, NumMante As String) As Boolean
'Actualizar en la tabla temporal tmpnseries el Nº de Serie
Dim SQL As String
On Error GoTo EActualizar
           
    'Insertar en la Tabla Temporal
    SQL = "UPDATE " & NombreTabla & " SET numserie="
    'Ha puesto algo.... TRIM
    If nSerie <> " " Then nSerie = Trim(nSerie)
    SQL = SQL & DBSet(nSerie, "T")
    
    'Modificacion
    SQL = SQL & " , nummante = " & DBSet(NumMante, "T")
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    SQL = SQL & " AND numlinea=" & Data1.Recordset!numlinea
    'Junio2010
    SQL = SQL & " AND numlinealb = " & Data1.Recordset!numlinealb
    
    
    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        SQL = "Actualizando Nº de Serie."
        MuestraError Err.Number, SQL, Err.Description
        ActualizarNumSerie = False
    Else
        ActualizarNumSerie = True
    End If
End Function


Private Function DatosOk(nSerie As String) As Boolean
Dim devuelve As String
Dim b As Boolean
Dim NumAlbar As String

    DatosOk = False
    b = True
    If Data1.Recordset.EOF Then Exit Function
    
    'Comprobar si existe en la tabla sserie, y ya esta comprado
    If DeVentas Then
        NumAlbar = "numalbar"
    Else
        NumAlbar = "numalbpr"
    End If
    
    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", nSerie, "T", NumAlbar, "codartic", Data1.Recordset!codArtic, "T")
    If devuelve <> "" Then  'Existe en tabla sserie
        'Comprobar si ya ha sido comprado en otro albaran
        If NumAlbar <> "" Then
            If DeVentas Then
                If Val(NumAlbar) <> Val(NumAlb) Then
                    devuelve = "El Nº de Serie ya ha sido vendido en el Albaran Nº: " & Format(NumAlbar, "0000000")
                    MsgBox devuelve, vbExclamation
                    b = False
                End If
            ElseIf NumAlbar <> NumAlb Then
                    devuelve = "El Nº de Serie ya ha sido comprado en el Albaran Nº: " & NumAlbar
                    MsgBox devuelve, vbExclamation
                    b = False
            End If
        End If
    End If
    
        'comprobar si existe en la tabla temporal con lo cual quiere decir que ya hemos
        'introducido ese nº de serie
        devuelve = "select count(*) from tmpnseries WHERE codusu=" & vUsu.Codigo
        devuelve = devuelve & " and codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " and numserie=" & DBSet(nSerie, "T")
        devuelve = devuelve & " and numlinea <>" & Data1.Recordset!numlinea
        If RegistrosAListar(devuelve) > 0 Then
            devuelve = "Ya ha introducido una linea con el nº de serie: " & nSerie
            MsgBox devuelve, vbInformation
            b = False
        End If
        
        
'        NumAlbar = "numlinea"
'        devuelve = DevuelveDesdeBDNew(conAri, "tmpnseries", "numserie", "numserie", nSerie, "T", NumAlbar, "codartic", Data1.Recordset!codArtic, "T")
'        If devuelve <> "" Then
'            If CInt(NumAlbar) <> CInt(Data1.Recordset!numlinea) Then
'                devuelve = "Ya ha introducido una linea con el nº de serie: " & nSerie
'                MsgBox devuelve, vbInformation
'                b = False
'            End If
'        End If
'    End If
    DatosOk = b
End Function


Private Sub CargarNSeriesConsecutivos(cadNSerie As String)
'Obtiene los Nº de Serie sucesivos para un mismo articulo a partir de un
'Nº de Serie dado
'Ej: A partid de la entrada: artic=0000016  numserie=0016-001
'   genera para cada linea de ese artic que encuentre: numserie=0016-02, 0016-03,...
Dim codArtic As String
Dim NumMante As String

    If Data1.Recordset.EOF Then Exit Sub
    
    CargaTxtAux False, False
    gridCargado = False
    
    'Insertar el primer registro
    If DatosOk(cadNSerie) Then
        NumMante = ""
        If DeVentas Then
            If Trim(txtauz.Text) <> "" Then
                    If MsgBox("¿Desea poner a los numeros de serie el mantenimiento: " & txtauz.Text & " ?", vbQuestion + vbYesNo) = vbYes Then NumMante = txtauz.Text
            End If
        End If
        codArtic = Data1.Recordset!codArtic
        ActualizarNumSerie cadNSerie, NumMante  'pongo un NULL en nummante
        Data1.Recordset.MoveNext
    
        'Insertar sucesivos registros
        While Not Data1.Recordset.EOF
          If Data1.Recordset!codArtic = codArtic Then
            'Obtener el Nº De Serie Siguiente
            cadNSerie = ObtenerNSerieSiguiente(cadNSerie)
            If DatosOk(cadNSerie) Then ActualizarNumSerie cadNSerie, NumMante
          End If
          Data1.Recordset.MoveNext
        Wend
    End If
    
    gridCargado = True
    CargaGrid True
    CargaTxtAux True, True
    PonerFoco txtAux
End Sub

Private Sub txtAux_LostFocus()
    'PonerFocoBtn Me.cmdAceptar
End Sub







'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
'
'
'  Txtaus  ES EL NUMERO DE MANTENIMIENTO
Private Sub txtauz_LostFocus()
    PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub txtauz_GotFocus()
    ConseguirFoco txtauz, 3
End Sub

Private Sub txtauz_KeyDown(KeyCode As Integer, Shift As Integer)
    gridCargado = False
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                'If DataGrid1.Row > 0 Then
                '    DataGrid1.Row = DataGrid1.Row - 1
                '    CargaTxtAux True, True
                'End If
                Data1.Recordset.MovePrevious
                If Data1.Recordset.BOF Then Data1.Recordset.MoveNext
                gridCargado = True
                CargaTxtAux True, True
                
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                'If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                '    DataGrid1.Row = DataGrid1.Row + 1
                '    CargaTxtAux True, True
                'End If
                
                'ENERO 2009
                Data1.Recordset.MoveNext
                If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
                gridCargado = True
                CargaTxtAux True, True
                PonerFoco txtauz
                
    End Select
    
End Sub


Private Sub txtauz_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    If KeyAscii = 27 Then cmdCancelar_Click 'ESC
End Sub


