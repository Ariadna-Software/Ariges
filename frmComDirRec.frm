VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComDirRecogida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dpto"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   Icon            =   "frmComDirRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   10575
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7920
         TabIndex        =   10
         Top             =   120
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   9240
         TabIndex        =   11
         Top             =   120
         Width           =   1035
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   9240
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Height          =   540
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2655
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2355
         End
      End
   End
   Begin VB.Frame FrameDirecciones 
      Caption         =   "Direcciones rec. proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   3915
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   10695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   1560
         TabIndex        =   29
         Tag             =   "CODCLIEN|N|N|0||sdirec|codclien|000|S|"
         Text            =   "Text3"
         Top             =   2760
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   1875
         Index           =   9
         Left            =   6840
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "Obs|T|S|||sdirenvio|observa||N|"
         Text            =   "frmComDirRec.frx":000C
         Top             =   1800
         Width           =   3765
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Código Direc./Dpto|N|N|0|999|sdirec|coddirec|000|S|"
         Text            =   "Text3"
         Top             =   360
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Domicilio|T|N|||sdirec|domdirec||N|"
         Text            =   "Text3"
         Top             =   1080
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Población|T|N|||sdirec|pobdirec||N|"
         Text            =   "Text3"
         Top             =   1785
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Provincia|T|N|||sdirec|prodirec||N|"
         Text            =   "Text3"
         Top             =   2145
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Teléfono|T|S|||sdirec|teldirec||N|"
         Text            =   "Text3"
         Top             =   1080
         Width           =   1605
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Persona Contacto|T|S|||sdirec|perdirec||N|"
         Text            =   "Text3"
         Top             =   720
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Nombre Direc./Dpto|T|N|||sdirec|nomdirec||N|"
         Text            =   "Text3"
         Top             =   720
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Fax|T|S|||sdirec|faxdirec||N|"
         Text            =   "Text3"
         Top             =   1425
         Width           =   1605
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "C.Postal|T|N|||sdirec|codpobla||N|"
         Text            =   "Text3"
         Top             =   1425
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   58
         Left            =   5520
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   23
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
         Height          =   255
         Index           =   24
         Left            =   360
         TabIndex        =   19
         Top             =   1425
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   18
         Top             =   1785
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   26
         Left            =   360
         TabIndex        =   17
         Top             =   2145
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   28
         Left            =   5520
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Pers. Contacto"
         Height          =   255
         Index           =   27
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         Height          =   255
         Index           =   30
         Left            =   5520
         TabIndex        =   14
         Top             =   1425
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   3240
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas descuento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7200
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmComDirRecogida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DireccionesEnvio As Boolean
Public codprove As Long
Public nomprove As String
Public VerDatoDpto As Integer  'Si trae valor es que situaremos en el registro


Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

Private CadenaConsulta2 As String
Private Ordenacion2 As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim PrimVez As Boolean






Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim B As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                'If Me.DireccionesEnvio Then
                    B = InsertarModificarLineaRecog
                'Else
                '    B = InsertarModificarLineaDpto
                'End If
                
                    
                If B Then
                    If Data1.Recordset.EOF Then
                        CadenaConsulta2 = " where codprove = " & codprove & " AND " & Ordenacion2 & " = " & Text1(0).Text
                        CadenaConsulta2 = "Select * from " & NombreTabla & CadenaConsulta2
                        Data1.RecordSource = CadenaConsulta2
                    End If
                    PosicionarData
                End If
                
            End If
        
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
      
        
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
Dim Hay As Boolean
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
        
        Case 3 'Insertar
            Hay = False
            If Not Data1.Recordset Is Nothing Then
                If Not Data1.Recordset.EOF Then Hay = True
            End If
            If Not Hay Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            

        
    End Select
  
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, Ordenacion2, "codprove = " & Me.codprove)
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia "codprove = " & Me.codprove
    Else
        CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE codprove = " & Me.codprove & " ORDER BY " & Ordenacion2
        PonerCadenaBusqueda
    End If

End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
   
    
    PonerModo 4
    PonerFoco Text1(1)
End Sub





Private Sub cmdRegresar_Click()
Dim Cad As String

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
    
        PonerModo 2
        If Not Data1.Recordset.EOF Then Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    Else
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
    
        Cad = Data1.Recordset.Fields(1) & "|"
        Cad = Cad & Data1.Recordset.Fields(2) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub



Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        BotonVerTodos
        
        If Me.VerDatoDpto >= 0 Then
            If SituarData(Data1, Ordenacion2 & "=" & Me.VerDatoDpto, Me.lblIndicador) Then PonerCampos
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    
    

End Sub

Private Sub Form_Load()
    'Icono del formulario
    PrimVez = True
    Me.Icon = frmPpal.Icon
    Me.imgBuscar(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
    ' ICONITOS DE LA BARRA
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 10
        .Buttons(10).Image = 16  ' Imprimir
        .Buttons(11).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    LimpiarCampos   'Limpia los campos TextBox
    
    
    Text1(10).Top = 8000
    Text1(10).Locked = True
    
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    
  
    '## A mano
    'If Not DireccionesEnvio Then
    '    NombreTabla = DevuelveTextoDepto(False)
    '    Caption = NombreTabla
    '    FrameDirecciones.Caption = NombreTabla
    '    NombreTabla = "sdirec"
    '    Ordenacion2 = " coddirec"
    '    NumRegElim = 3195
    'Else
        Caption = "Dir. recogida"
        NombreTabla = "sdirRecog"
        Ordenacion2 = " coddirre"
        NumRegElim = 3915
    'End If
    Caption = Caption & "  Proveedor: " & UCase(nomprove) & "(" & codprove & ")"
    FrameDirecciones.Height = NumRegElim
    Frame2.Top = Me.FrameDirecciones.Top + FrameDirecciones.Height + 120
    Me.Height = Frame2.Top + Frame2.Height + 720
    NumRegElim = 0
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codprove=-1"
    'Data1.Refresh
    PonerTags
    
    'BotonBuscar
    Modo = 0


End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub






Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaConsulta2 = CadenaDevuelta
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        'Poblacion
        Text1(4).Text = ObtenerPoblacion(Text1(3).Text, CadenaSeleccion)
        'provincia
        Text1(5).Text = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Set frmCP = New frmCPostal
    frmCP.DatosADevolverBusqueda = "0"
    frmCP.Show vbModal
    Set frmCP = Nothing
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub





Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    'If BLOQUEADesdeFormulario(Me) Then BotonModificar
   
    BotonModificar


    
End Sub

Private Sub mnNuevo_Click()
  
    
         BotonAnyadir

End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub





Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then cadB = cadB & " AND "
    cadB = cadB & " codclien = " & Me.codprove

    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then

        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " ORDER BY  " & Ordenacion2
        PonerCadenaBusqueda
    End If
End Sub




Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo > 0 Then
            If Modo = 1 Then 'Busqueda
                 MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda." & vbCrLf & Caption, vbInformation
                 PonerFoco Text1(0)
            Else
                MsgBox "No hay ningún registro en la tabla " & NombreTabla & vbCrLf & Caption, vbInformation
                Me.lblIndicador.Caption = ""
                If Modo = 0 Then PonerModo 0
            End If
        Else
            PonerModo 0
        End If
        
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte

 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 5 Then lblIndicador.Caption = "Lineas dto"
    
    
   
    B = Modo < 5
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)

    cmdRegresar.visible = B
  
      
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera B Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    Text1(10).Locked = True  'Siempre bloqueado
    
    
    imgBuscar(0).visible = Modo = 1 Or Modo > 2
    BloquearChecks Me, Modo
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
                        
    If Modo <= 2 Then PonerFocoChk chkVistaPrevia
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    B = Modo < 3

    'Añadir
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = Modo = 2

    
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    

     '---------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim I As Integer

    If Modo = 3 Then Text1(10).Text = codprove
    
    


    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    DatosOk = B
End Function






Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim cto As Byte
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo
'            If Text1(Index).Text <> "" Then
             If PonerFormatoEntero(Text1(Index)) Then

            End If

        Case 3
            If Text1(Index).Text <> "" Then
                If IsNumeric(Text1(Index)) Then
                    Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, CadenaConsulta2)
                    Text1(Index + 2).Text = CadenaConsulta2
                    CadenaConsulta2 = ""
                End If
            End If
'        Case 10 To 13
'            If Text1(Index).Text <> "" Then
'                If Not PonerFormatoEntero(Text1(Index)) Then
'                    Text1(Index).Text = ""
'                    PonerFoco Text1(Index)
'                Else
'
'                    If Index = 13 Then
'                        cto = 10
'                    Else
'                        If Index = 12 Then
'                            cto = 2
'                        Else
'                            cto = 4
'                        End If
'                    End If
'
'                    Text1(Index).Text = Right(String("0", cto) & Text1(Index).Text, cto)
'
'                End If
'            End If
                    
    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5  'Nuevo
                mnNuevo_Click
        Case 6  'Modificar
                mnModificar_Click

          Case 7
            mnEliminar_Click

        Case 11: mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then PonerFocoBtn Me.cmdRegresar
    
    cmdCancelar.Cancel = True
    
    
   
    Me.cmdRegresar.visible = B
    
    
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear

    
End Sub


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(" & Ordenacion2 & "=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub





Private Sub PonerTags()

    'If Not DireccionesEnvio Then
        
'        Text1(0).Tag = "Código Direc./Dpto|N|N|0|999|sdirRecog|coddirec|000|S|"
'        Text1(1).Tag = "Nombre Direc./Dpto|T|N|||sdirRecog|nomdirec|||"
'        Text1(2).Tag = "Domicilio|T|N|||sdirRecog|domdirec|||"
'        Text1(3).Tag = "C.Postal|T|N|||sdirRecog|codpobla|||"
'        Text1(4).Tag = "Población|T|N|||sdirRecog|pobdirec|||"
'        Text1(5).Tag = "Provincia|T|N|||sdirRecog|prodirec|||"
'        Text1(6).Tag = "Persona Contacto|T|S|||sdirRecog|perdirec|||"
'        Text1(7).Tag = "Teléfono|T|S|||sdirRecog|teldirec|||"
'        Text1(8).Tag = "Fax|T|S|||sdirRecog|faxdirec|||"
'        Text1(9).Tag = "e-mail|T|S|||sdirec|maidirec|||"
'        Text1(10).Tag = "Código Banco|N|S|0|9999|sdirec|codbanco|0000||"
'        Text1(11).Tag = "Sucursal|N|S|0|9999|sdirec|codsucur|0000||"
'        Text1(12).Tag = "Dígito Control|T|S|||sdirec|digcontr|00||"
'        Text1(13).Tag = "Cuenta Bancaria|T|S|||sdirec|cuentaba|0000000000||"
'
'
'        'codclien
'        Text1(15).Tag = "CODCLIEN|N|N|0||sdirec|codclien|000|S|"
   ' Else
   '
        Text1(0).Tag = "Código|N|N|0|999|sdirRecog|coddirre|000|S|"
        Text1(1).Tag = "Nombre Direc|T|N|||sdirRecog|nomdirre|||"
        Text1(2).Tag = "Domicilio|T|N|||sdirRecog|domdirre|||"
        Text1(3).Tag = "C.Postal|T|N|||sdirRecog|codpobla|||"
        Text1(4).Tag = "Población|T|N|||sdirRecog|pobdirre|||"
        Text1(5).Tag = "Provincia|T|N|||sdirRecog|prodirre|||"
        Text1(6).Tag = "Persona Contacto|T|S|||sdirRecog|perdirre|||"
        Text1(7).Tag = "Teléfono|T|S|||sdirRecog|teldirre|||"
        Text1(8).Tag = "Fax|T|S|||sdirRecog|faxdirre|||"
        Text1(9).Tag = "Obs|T|S|||sdirRecog|observa|||"
        'codclien
        Text1(10).Tag = "CODCLIEN|N|N|0||sdirRecog|codprove|000|S|"
   ' End If
        
    'Cuales estaran visibles
   
'    For NumRegElim = 9 To 13
'        Text1(NumRegElim).visible = Not Me.DireccionesEnvio
'    Next
    
End Sub




Private Sub MandaBusquedaPrevia(cadB As String)

        CadenaConsulta2 = ParaGrid(Text1(0), 10, "Código")
        CadenaConsulta2 = CadenaConsulta2 & ParaGrid(Text1(1), 75, "Nombre")
        
        
            
            
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = CadenaConsulta2
        CadenaConsulta2 = ""
        frmB.vTabla = NombreTabla
        frmB.vSQL = cadB
        
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Direcccion envio proveedor " & Me.nomprove
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
        frmB.vCargaFrame = False
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        
        If CadenaConsulta2 <> "" Then
            CadenaConsulta2 = RecuperaValor(CadenaConsulta2, 1)
            CadenaConsulta2 = " codprove = " & Me.codprove & " AND " & Ordenacion2 & " = " & CadenaConsulta2
            CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE " & CadenaConsulta2
            PonerCadenaBusqueda
        
        Else   'es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
        CadenaConsulta2 = ""
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    Cad = "codprove = " & Me.codprove & " AND coddirre "
    Cad = DevuelveDesdeBD(conAri, "numpedpr", "scappr", Cad, CStr(Data1.Recordset!coddirre))
    If Cad = "" Then
        Cad = "codprove = " & Me.codprove & " AND coddirre "
        Cad = DevuelveDesdeBD(conAri, "numpedpr", "schppr", Cad, CStr(Data1.Recordset!coddirre))
    End If
    If Cad <> "" Then
        MsgBox "Direccion de recogida esta en pedidos: (" & Cad & ")", vbExclamation
        Exit Sub
    End If
    

  
   
        
        '### a mano
        Cad = "¿Seguro que desea eliminar?"
        Cad = Cad & vbCrLf & "Codigo : " & Data1.Recordset.Fields(1)
        Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(2)

        'Borramos
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            On Error GoTo Error2
            Screen.MousePointer = vbHourglass
            NumRegElim = Data1.Recordset.AbsolutePosition
            Data1.Recordset.Delete
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else 'solo habia un registro
                LimpiarCampos
                PonerModo 0
            End If
        End If
  
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Cliente", Err.Description
    End If
End Sub




Private Function InsertarModificarLineaRecog() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLineaRecog = False
    SQL = ""
   If Modo = 3 Then
        
            SQL = "INSERT INTO sdirRecog (codprove,coddirre,nomdirre,domdirre,codpobla,pobdirre,prodirre,perdirre,teldirre,faxdirre,observa) VALUES ("
            SQL = SQL & codprove & ", "
            SQL = SQL & Text1(0).Text
            For I = 1 To 5
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T")
            Next I
                    
            For I = 6 To 8 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next I
                        
            SQL = SQL & "," & DBSet(Text1(I).Text, "T", "S") & ")"
 
        
    Else
            SQL = "UPDATE sdirRecog Set nomdirre = " & DBSet(Text1(1).Text, "T")
            SQL = SQL & ", domdirre = " & DBSet(Text1(2).Text, "T")
            SQL = SQL & ", codpobla = " & DBSet(Text1(3).Text, "T")
            SQL = SQL & ", pobdirre = " & DBSet(Text1(4).Text, "T")
            SQL = SQL & ", prodirre = " & DBSet(Text1(5).Text, "T")
            SQL = SQL & ", perdirre = " & DBSet(Text1(6).Text, "T")
            SQL = SQL & ", teldirre = " & DBSet(Text1(7).Text, "T")
            SQL = SQL & ", faxdirre = " & DBSet(Text1(8).Text, "T")
            SQL = SQL & ", observa = " & DBSet(Text1(9).Text, "T")
            SQL = SQL & " WHERE codprove =" & codprove & " AND "
            SQL = SQL & " coddirre =" & (Text1(0).Text)
    End If
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaRecog = True
    Else
        PonerFoco Text1(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Direcciones de recogida" & vbCrLf & Err.Description
End Function





