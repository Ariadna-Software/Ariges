VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMailPromo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mailing promociones"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8775
   Icon            =   "frmMailPromo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSitua 
      Height          =   315
      ItemData        =   "frmMailPromo.frx":000C
      Left            =   4800
      List            =   "frmMailPromo.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "Situa|N|N|||smailpromoca|situacion|||"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha fin|F|N|||smailpromoca|fechafin|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Desc|T|N|||smailpromoca|descripcion|||"
      Text            =   "Tex"
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      ToolTipText     =   "Buscar artículo"
      Top             =   5280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   4920
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "1|N|S|0||t|c|##,##0.000|S|"
      Text            =   "cantidad"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "nombre artic"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "codartic"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   7755
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   7755
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7515
      TabIndex        =   19
      Top             =   7755
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   7710
      Width           =   3000
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
         TabIndex        =   18
         Top             =   180
         Width           =   2595
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Index           =   2
      Left            =   240
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "O|T|S|||smailpromoca|observa||N|"
      Text            =   "frmMailPromo.frx":0038
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha inicio|F|N|||smailpromoca|fechaini|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
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
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Realizar proceso promocion"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3840
      Top             =   7680
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Codigo|N|N|0||smailpromoca|codigo||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   3000
      Top             =   8040
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMailPromo.frx":003E
      Height          =   4200
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7408
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   24
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   23
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Artículos"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   6960
      Picture         =   "frmMailPromo.frx":0053
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Fin"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   21
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   5760
      Picture         =   "frmMailPromo.frx":00DE
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   550
      Width           =   1095
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
      Left            =   360
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmMailPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As Boolean
Public Event DatoSeleccionado(CadenaSeleccion As String)


'--------------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1

Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Private HaDevueltoDatos As Boolean



Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        Text1(kCampo).BackColor = vbWhite
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk(True) Then
            If InsertarDesdeForm(Me) Then
                Espera 0.5
                Data1.Refresh
                
                 PosicionarData True
                 If Not Me.Data1.Recordset.EOF Then
                    BotonLineas
                    mnNuevo_Click
                 End If
            End If
        End If
    Case 4 'MODIFICAR
        If DatosOk(True) Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData False
                 
             End If
         End If
            
    Case 5
        If InsertarModificarLinea Then
            'Reestablecemos los campos
            'y ponemos el grid
            If ModificaLineas = 2 Then TerminaBloquear
           
                DataGrid1.AllowAddNew = False
          
            CargaGridArticulos True
            If ModificaLineas = 1 Then 'Insertar
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                    Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & DBSet(txtAux(0).Text, "T"))
                ModificaLineas = 0
                PonerBotonCabecera True
                Me.lblIndicador.Caption = ""
                    CargaTxtAux False, False
                    DataGrid1.Enabled = True
                    PonerFocoGrid Me.DataGrid1
                
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    Set FrmArt = New frmAlmArticu2
    'frmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo busqueda
    FrmArt.DesdeTPV = False
    FrmArt.Show vbModal
    Set FrmArt = Nothing
    txtAux_LostFocus 0
End Sub



Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
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
            PonerFoco Text1(0)
            
        Case 5 'Mantenimiento Lineas traspasos
            CargaTxtAux False, False
            DataGrid1.Enabled = True
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then 'Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
             PonerBotonCabecera True
            DataGrid1.Refresh
            PonerFocoBtn Me.cmdRegresar
        
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo >= 5 Then 'modo 5: Mantenimiento Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdRegresar.visible = Me.DatosADevolverBusqueda
        Me.cmdCancelar.Cancel = True
        Me.cmdRegresar.Caption = "Regresar"
    Else
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        RaiseEvent DatoSeleccionado(Data1.Recordset.Fields(0) & "|" & Data1.Recordset.Fields(1) & "|")
        Unload Me
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton Añadir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 10 'Mantenimiento Líneas
        .Buttons(10).Image = 20 '
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(13).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
       
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)


    cadSeleccion = ""
    NombreTabla = "smailpromoca"
    NomTablaLineas = "smailpromoli" '
    Ordenacion = " ORDER BY codigo"
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " WHERE false"

    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGridArticulos False
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGridArticulos(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
      
    DataGrid1.Columns(0).visible = False 'Cod. trasp
   
    I = 1
    'Cod. Artículo
    DataGrid1.Columns(I).Caption = "Cod. Articulo"
    DataGrid1.Columns(I).Width = 1900
    
    'Nombre Artículo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Nombre Articulo"
    DataGrid1.Columns(I).Width = 4200
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Precio"
    DataGrid1.Columns(I).Width = 1500
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoPrecio
    
   
       
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
       
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = 290
        Next I
        Me.cmdAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
            Next I
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For I = 0 To txtAux.Count - 1
'                If i <> 1 Then txtAux(i).Locked = False
                'LAURA 19/10/2006
                If I <> 1 Then BloquearTxt txtAux(I), False
            Next I
            cmdAux.Enabled = True
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I + 1).Text
            Next I
            BloquearTxt txtAux(0), True
            cmdAux.Enabled = False
            BloquearTxt txtAux(2), False

        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        
        'Fijamos altura y posición Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(1).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width + 30 'Nom artic
        txtAux(1).Width = DataGrid1.Columns(2).Width + 10
        txtAux(2).Left = DataGrid1.Columns(3).Left + 180
        txtAux(2).Width = DataGrid1.Columns(3).Width + 60
      '  For I = 1 To txtAux.Count - 1 'Cantidad y Observacion
      '      txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 25
      '      txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 35
      '  Next I
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux.visible = visible
End Sub







''Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'''Almacenes Propios
''Dim indice As Byte
''    indice = CByte(Me.imgBuscar(0).Tag)
''    Text1(indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
''    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2)
''End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
       
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(CInt(imgFecha(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub






Private Sub imgBuscar_Click(index As Integer)
Dim Aux As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
'    imgBuscar(2).Tag = Index
    
    Select Case index
'        Case 2
'            Aux = CadenaConsulta
'            CadenaConsulta = ""
'            Set frmPl = New frmAlmPlagas
'            frmPl.DatosADevolverBusqueda = "0|1|"
'            frmPl.Show vbModal
'            Set frmPl = Nothing
'            If CadenaConsulta <> "" Then
'                Text1(2).Text = RecuperaValor(CadenaConsulta, 1)
'                Text2(0).Text = RecuperaValor(CadenaConsulta, 2)
'            End If
'            CadenaConsulta = Aux
    End Select
    
    PonerFoco Text1(index)
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(index As Integer)
Dim Indice As Byte

   Screen.MousePointer = vbHourglass
   imgFecha(1).Tag = index
   Set frmF = New frmCal
   frmF.Fecha = Now
   

   
   PonerFormatoFecha Text1(index)
   If Text1(index).Text <> "" Then frmF.Fecha = CDate(Text1(index).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(index)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo >= 5 Then   'Eliminar lineas Traspaso Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Traspaso Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo >= 5 Then  'Modificar lineas Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificarLinea
    Else 'Modificar Cabecera Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo >= 5 Then  'Añadir lineas Traspaso Almacenes
        BotonAnyadirLineas
    
    Else 'Añadir Cabecera Traspaso Almacenes
        BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    If index <> 5 Then ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And index = 5 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(index As Integer)
    
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
    
    Select Case index
        Case 0 'Codigo
            
        Case 1, 3 'Fecha
            If Text1(index).Text <> "" And Modo <> 1 Then PonerFormatoFecha Text1(index)
        Case 2
            'Observaciones
            If Text1(index).Text <> "" Then Text1(index).Text = QuitarCaracterEnter(Text1(index).Text)
    End Select
End Sub


Private Sub txtAux_GotFocus(index As Integer)
    ConseguirFocoLin txtAux(index)
End Sub

Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 3 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(index As Integer)
Dim devuelve As String

    'Quitar espacios en blanco por los lados
    txtAux(index).Text = Trim(txtAux(index).Text)
    
    Select Case index
        Case 0 'Cod. Articulo
            If txtAux(index).Text = "" Then
                txtAux(index + 1).Text = ""
            ElseIf ModificaLineas = 1 Then 'Insertando linea
                'Comprobamos si ya existe una linea con el artículo, solo si estamos insertando (ModificaLineas=1)
                'conAri: conexion a BD Ariges
                devuelve = "codigo=" & DBSet(Text1(0).Text, "T") & " AND codartic "
                devuelve = DevuelveDesdeBD(conAri, "codartic", NomTablaLineas, devuelve, txtAux(0).Text, "T")
                If devuelve <> "" Then
                    devuelve = "Ya hay una línea con ese Artículo: " & vbCrLf
                    devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
                    MsgBox devuelve, vbExclamation
                     txtAux(index).Text = ""
                    PonerFoco txtAux(index)
                Else
                    cadSeleccion = "precioac"
                    devuelve = "sartic.codartic"
                    devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic left join spromo on sartic.codartic=spromo.codartic", devuelve, txtAux(0).Text, "T", cadSeleccion)
                    If devuelve = "" Then
                        MsgBox "No existe el artículo", vbExclamation
                    Else
                        If cadSeleccion = "" Then
                            MsgBox "El articulo no esta en promocion", vbExclamation
                            devuelve = ""
                        End If
                    End If
                    txtAux(1).Text = devuelve
                    If txtAux(1).Text = "" Then
                        txtAux(index).Text = ""
                        PonerFoco txtAux(index)
                    End If
                    cadSeleccion = ""
                End If
            End If
            
        Case 2, 3 'Cantidad (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            If txtAux(index) <> "" Then
                If Not PonerFormatoDecimal(txtAux(index), 2) Then txtAux(index).Text = ""
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    Select Case Button.index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
            
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 9, 10
                If Data1.Recordset.EOF Then Exit Sub
                C = ""
                If Button.index = 9 Then
                    If Val(Data1.Recordset!Situacion) > 0 Then C = "Imposible modificar lineas en esta situacion"
                Else
                    If Val(Data1.Recordset!Situacion) = 2 Then C = "Situacion: cerrada"
                End If
                If C <> "" Then
                    MsgBox C, vbExclamation
                    Exit Sub
                End If
        
                If Button.index = 9 Then
                    BotonLineas
                Else
                    ProcesoMailingPromo
                End If
        Case 12 'Imprimir
            BotonImprimir
            
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo

    'Modo 2. Hay datos y estamos visualizandolos
    '-------------------------------------------
    b = (Kmodo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
              
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), (Modo <> 1 And Modo <> 3)
    
    'Modo 1:Busqueda / Modo 3: Insertar / Modo 4: Modificar
    '-------------------------------------------------------
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 1 To 3
        If I <> 2 Then Me.imgFecha(I).Enabled = b
    Next I
    
    BloquearCmb cboSitua, Modo <> 1
    
    
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    cmdRegresar.visible = False
    If Modo = 2 Then
        '
        If Me.DatosADevolverBusqueda Then
            If Not Me.Data1.Recordset.EOF Then cmdRegresar.visible = True

        End If
    End If
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
                        
    Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    
    
        Toolbar1.Buttons(10).visible = True
  
    
    
         b = (Modo = 2) Or (Modo >= 5)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
        'Lineas Traspaso Almacenes
        Toolbar1.Buttons(9).Enabled = b
        'Actualizar
        Toolbar1.Buttons(10).Enabled = b
        'Imprimir
        Toolbar1.Buttons(12).Enabled = b
            
        '-------------------------------
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'VerTodos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox

    'Aqui va el especifico de cada form es
    '### a mano
    cboSitua.ListIndex = -1
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones(Flechas) de Desplazamiento de Registros de la Toolbar

    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, index
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, index
            PonerCampos
    End Select
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
Dim tabla As String
On Error GoTo EMontaSQL
 
    
        tabla = NomTablaLineas
    
    
        SQL = "SELECT " & tabla & ".codigo, "
        SQL = SQL & tabla & ".codartic, Articulos.nomartic, " & tabla & ".precioMail "
        SQL = SQL & " FROM ((" & tabla & " LEFT JOIN sartic AS Articulos ON " & tabla & ".codartic ="
        SQL = SQL & " Articulos.codartic))"
    
    
    If enlaza Then
        SQL = SQL & ObtenerWhereCP(True)  '" WHERE codtrasp = " & Data1.Recordset!codtrasp
    Else
        SQL = SQL & " WHERE false"
    End If
    SQL = SQL & " ORDER BY " & tabla & ".codartic"
    MontaSQLCarga = SQL
    
EMontaSQL:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGridArticulos False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        PonerFoco Text1(0)
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
      PonerModo 5
    ModificaLineas = 0
    PonerBotonCabecera True
    CargaGridArticulos True
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
        
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGridArticulos False
    
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codigo", "")

    Me.cboSitua.ListIndex = 0
    PonerFoco Text1(4)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
'    vWhere = ObtenerWhereCP(False)
'    If Modo = 5 Then
'        cmdAceptar.Tag = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
'    Else
'        cmdAceptar.Tag = SugerirCodigoSiguienteStr("advtrataPlagas", "numlinea", vWhere)
'    End If
    
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    
        AnyadirLinea DataGrid1, Data2
    
        CargaTxtAux True, True
        PonerFoco txtAux(0)
   
    

End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True
    
    
    
    If Val(Data1.Recordset!Situacion) > 0 Then
        BloquearTxt Text1(1), True
        BloquearTxt Text1(3), True
    End If
    
    PonerFoco Text1(1)
End Sub




Private Sub BotonModificarLinea()
Dim I As Integer

    
        If Data2.Recordset.EOF Then Exit Sub
        If Data2.Recordset.RecordCount < 1 Then Exit Sub
   
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar

    Screen.MousePointer = vbHourglass
    
    PonerBotonCabecera False
    Me.lblIndicador.Caption = "MODIFICAR"
    
    
        If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
            I = DataGrid1.Bookmark - DataGrid1.FirstRow
            DataGrid1.Scroll 0, I
            DataGrid1.Refresh
        End If
        
     
    
        CargaTxtAux True, False
        PonerFoco txtAux(2) 'Poner el foco
        Me.DataGrid1.Enabled = False
   
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Val(Data1.Recordset!Situacion) > 0 Then
        MsgBox "Imposible eliminar los registros en esta situacion", vbExclamation
        Exit Sub
    End If
    
    
    
    SQL = "Va a eliminar " & vbCrLf
    SQL = SQL & "------------------------------------------" & vbCrLf & vbCrLf
    SQL = SQL & vbCrLf & "Codigo   : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Descripcion  : " & CStr(Data1.Recordset.Fields(1))
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
'
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        NumRegElim = Data1.Recordset.Fields(0)
'        vTipoMov.DevolverContador CodTipoMov, NumRegElim
'        Set vTipoMov = Nothing
    
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGridArticulos False
            PonerModo 0
        End If
    End If
     
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
    
    conn.BeginTrans
    SQL = ObtenerWhereCP(True)  '" WHERE  codtrasp=" & Data1.Recordset!codtrasp
    
    'Lineas
    conn.Execute "Delete  from " & NomTablaLineas & SQL
  
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub BotonEliminarLinea()
Dim SQL As String
On Error GoTo Error2
    'Ciertas comprobaciones
    
        If Data2.Recordset.EOF Then Exit Sub
  

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    SQL = "Seguro que desea eliminar la línea :"
    SQL = SQL & "del Artículo" & vbCrLf & "Código: @2"
    SQL = SQL & vbCrLf & "Descripción: @3"
    
        SQL = Replace(SQL, "@1", "del Artículo")
        SQL = Replace(SQL, "@2", Data2.Recordset!codArtic)
        SQL = Replace(SQL, "@3", Data2.Recordset.Fields(3))
  
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    
       
            SQL = NomTablaLineas
            
      
        
        SQL = "Delete from " & SQL & ObtenerWhereCP(True)
        SQL = SQL & " and codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        
        conn.Execute SQL
        
        
            CancelaADODC Me.Data2
            CargaGridArticulos True
            CancelaADODC Me.Data2
       
    End If
    ModificaLineas = 0
Error2:
        Screen.MousePointer = vbDefault
        ModificaLineas = 0
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Línea ", Err.Description
End Sub


Private Function DatosOk(Optional cabecera As Boolean) As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function

   
    
    DatosOk = b
End Function


   

Private Function DatosOkLinea() As Boolean
Dim b As Boolean

Dim cad As String

    DatosOkLinea = False
    b = True
    
    cad = ""
    kCampo = 0
    If txtAux(0).Text = "" Then
        cad = "El campo Cod. Artículo no puede ser nulo"
    Else
        If txtAux(1).Text = "" Then cad = "El campo Cod. Artículo incorrecto"
    End If
    If txtAux(2).Text = "" Then cad = vbCrLf & "Precio no puede estar vacio": kCampo = 2
    If cad <> "" Then
        b = False
        PonerFoco txtAux(kCampo)
    End If
    
    
    
    DatosOkLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.cmdRegresar.Cancel = True
        If Modo = 5 Then
            Me.lblIndicador.Caption = "Lin. artículos"
        Else
            Me.lblIndicador.Caption = "Lin. plagas"
        End If
    Else
        Me.lblIndicador.Caption = ""
        Me.cmdCancelar.Cancel = True
    End If
     'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Function InsertarModificarLinea() As Boolean

On Error GoTo EInsertarModificarLinea

   
        InsertarModificarLinea = InsertarModificarLineaArt
   
    CadenaConsulta = ""
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas " & vbCrLf & Err.Description
    InsertarModificarLinea = False
End Function

Private Function InsertarModificarLineaArt() As Boolean
Dim SQL As String

    
    
    SQL = ""
    

    
    InsertarModificarLineaArt = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea() Then 'INSERTAR
            SQL = "INSERT INTO smailpromoli(codigo,codartic,precioMail)"
            SQL = SQL & " VALUES (" & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(2).Text, "N", "N") & ") "
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLinea() Then
            SQL = "UPDATE smailpromoli Set precioMail = " & DBSet(txtAux(2).Text, "N")
            SQL = SQL & ObtenerWhereCP(True) & " AND " '" WHERE codtrasp =" & Val(Text1(0).Text) & " AND "
            SQL = SQL & " codartic =" & DBSet(Data2.Recordset!codArtic, "T")
        End If
    End Select
    
    If ejecutar(SQL, False) Then InsertarModificarLineaArt = True
    

End Function




Private Sub MandaBusquedaPrevia(CadB As String)


    
       

    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Código|advtrata|codtrata|T||30·Denominacion|advtrata|nomtrata|T||70·"
    frmB.vTabla = NombreTabla
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Tratamientos"
    frmB.vselElem = 0
    frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
    '#
    frmB.Show vbModal
    Set frmB = Nothing

    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerBusqueda()
Dim CadB As String
    
    CadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & Ordenacion
            PonerCadenaBusqueda
        Else
'            CadenaConsulta = "select * from " & NombreTabla & Ordenacion
'            PonerCadenaBusqueda
            MsgBox "Introducir criterios de búsqueda", vbExclamation
            PonerFoco Text1(0)
        End If
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        Me.DataGrid1.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    If Err.Number <> 0 Then MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
   ' Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "splagas", "nombrepl")
    CargaGridArticulos True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub






Private Function InsertarCabeceraHistorico() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarCab

    SQL = "SELECT codtrasp,fechatra,almaorig,almadest,codtraba,observa1 from advtrata "
    SQL = SQL & ObtenerWhereCP(True)
    SQL = SQL & " AND fechatra='" & Format(Data1.Recordset!fechatra, "yyyy-mm-dd") & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        SQL = "INSERT INTO schtra (codtrasp, fechatra,hormovim,almaorig,almadest,codtraba,observa1) "
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(RS.Fields(1).Value, "yyyy-mm-dd") & "', '"
        SQL = SQL & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields(2).Value & ", " & RS.Fields(3).Value & ", "
        SQL = SQL & RS.Fields(4).Value & ", " & DBSet(RS.Fields(5).Value, "T") & ")"
    End If
    RS.Close
    Set RS = Nothing
    
    conn.Execute SQL
    
EInsertarCab:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(MenError As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarLineas

    SQL = "SELECT codtrasp, numlinea, codartic, cantidad, observa2 from slitra "
    SQL = SQL & ObtenerWhereCP(True)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    RS.MoveFirst
    While Not RS.EOF
        SQL = "INSERT INTO slhtra (codtrasp, fechamov, numlinea, codartic, cantidad, observa2)"
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(Data1.Recordset!fechatra, FormatoFecha) & "', "
        SQL = SQL & RS.Fields(1).Value & ", " & DBSet(RS.Fields(2).Value, "T") & ", "
        SQL = SQL & DBSet(RS.Fields(3).Value, "N") & ", " & DBSet(RS.Fields(4).Value, "T") & ")"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
        RS.Close
        Set RS = Nothing
        MenError = Err.Number & ": " & Err.Description
    Else
        MenError = ""
        InsertarLineasHistorico = True
    End If
End Function




Private Function BorrarTraspaso(MenError As String) As Boolean
Dim SQL As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    SQL = "Delete from "
    SQL = SQL & "slitra"
    SQL = SQL & " WHERE codtrasp = " & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    'La cabecera
    SQL = "Delete from "
    SQL = SQL & "advtrata"
    SQL = SQL & " WHERE codtrasp =" & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
        MenError = Err.Number & ": " & Err.Description
    Else
        BorrarTraspaso = True
        MenError = ""
    End If
End Function





Private Sub BotonImprimir()

    
  
    '    AbrirListado (7) '7: Informe Traspaso de Almacen
        

End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim cad As String
Dim numParam As Byte
Dim nomDocu As String

    cadParam = "|"
    numParam = 0
    If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
    
    indRPT = 2 '2: Historico Traspaso de Almacen
    If PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .NombreRPT = nomDocu
            .NombrePDF = pPdfRpt
            .EnvioEMail = False
            .Opcion = 7
            .Titulo = "Hist. Traspaso Alm."
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                'o estamos en Historico
                cad = "{schtra.codtrasp}= " & Data1.Recordset!codtrasp
                cad = cad & " and {schtra.fechatra}= Date(" & Year(Data1.Recordset!fechatra) & "," & Month(Data1.Recordset!fechatra) & "," & Day(Data1.Recordset!fechatra) & ")" & ""
                .FormulaSeleccion = cad
            End If
            .Show vbModal
        End With
    End If
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function ObtenerWhereCP(conWhere As Boolean) As String
On Error Resume Next
    ObtenerWhereCP = " codigo= " & DBSet(Text1(0).Text, "T")
    If conWhere Then ObtenerWhereCP = " WHERE " & ObtenerWhereCP
    
End Function


Private Sub PosicionarData(VieneDeInsertar As Boolean)
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String
Dim b As Boolean
    
    b = Data1.Recordset.EOF
    If Not b And VieneDeInsertar Then b = True

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    End If
End Sub




Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGridArticulos False
    If Err.Number <> 0 Then Err.Clear
End Sub

'-------------------------------------------------
' El proceso mete en spromo
Private Sub ProcesoMailingPromo()
Dim Reestablecer As Boolean
Dim Aux As String
Dim b As Boolean
    'Comprobaciones
    Reestablecer = Val(Data1.Recordset!Situacion) = 1
    
    
    If Not Reestablecer Then
        'Va a coger el precioac,fechaini,fechafin de spromo, y ponerlas en las lineas de mailpromo  y pmv en
        'Veamos que no hay ninguna abierta con estos articulos
        
        Aux = "codartic in (select codartic from smailpromoli,smailpromoca where smailpromoli.codigo=smailpromoca.codigo "
        Aux = Aux & " AND smailpromoli.codigo<>" & Data1.Recordset!Codigo & " and situacion = 1 ) AND 1"

        Aux = DevuelveDesdeBD(conAri, "min(codigo)", "smailpromoli", Aux, "1")
        If Val(Aux) > 0 Then
            MsgBox "El articulo esta en otro proceso de mailing/promociones", vbExclamation
            Exit Sub
        End If
        
    Else
        'Cogera de mailpromo y establecerá en spromo precioac,fechaini,fechafin
        
    End If
    If Reestablecer Then
        Aux = "Volver a poner los precios originales y cerrar el mailing"
    Else
        Aux = "Va a actualizar las promociones con estos precios y realizar el proceso de mailing"
    End If
    Aux = Aux & vbCrLf & vbCrLf & "¿continuar con el proceso?"
    If MsgBox(Aux, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    'Si se puede , adelante
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    Set miRsAux = New ADODB.Recordset
    If Reestablecer Then
        'metemos en spormo los valores antiugos
        b = ReestablecerEnSPromo
    Else
        
        'grabamos en spromo con los valores actuales
        b = GuardarEnSPromo
    End If
    Set miRsAux = Nothing
    If b Then
        conn.CommitTrans
        PosicionarData False
    Else
        conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
End Sub





Private Function GuardarEnSPromo() As Boolean
Dim Aux As String
    On Error GoTo eGuardarEnSPromo
    
    GuardarEnSPromo = True
    'Cogeremos de spromo u grabaremos los valores en las columnas
    Aux = "select s.codartic,spromo.fechaini ,spromo.fechafin ,spromo.precioac ,preciomail,s.pmv,preciominvta "
    Aux = Aux & " from smailpromoli s ,spromo, sartic where s.codartic=spromo.codartic and s.codartic=Sartic.codartic and spromo.codlista=1"
    Aux = Aux & " AND codigo=" & Data1.Recordset!Codigo
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic
    While Not miRsAux.EOF
        'GRabo los valores de spromo y pmv en lineas
        Aux = "update smailpromoli set precio_Sprees=" & DBSet(miRsAux!precioac, "N")
        Aux = Aux & " , PMV=" & DBSet(miRsAux!preciominvta, "N")
        Aux = Aux & " , fechainiart= " & DBSet(miRsAux!FechaIni, "F")
        Aux = Aux & " , fechafinart=" & DBSet(miRsAux!FechaFin, "F")
        Aux = Aux & " WHERE codigo = " & Data1.Recordset!Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
        conn.Execute Aux
        
        Aux = "UPDATE spromo set precioac=" & DBSet(miRsAux!preciomail, "N")
        Aux = Aux & " , FechaIni =" & DBSet(Data1.Recordset!FechaIni, "F")
        Aux = Aux & " , FechaFin=" & DBSet(Data1.Recordset!FechaFin, "F")
        Aux = Aux & " WHERE codlista = 1 AND codartic=" & DBSet(miRsAux!codArtic, "T")
        conn.Execute Aux
        
        Aux = "UPDATE Sartic set preciominvta=" & DBSet(miRsAux!preciomail, "N")
        Aux = Aux & " WHERE  codartic=" & DBSet(miRsAux!codArtic, "T")
        conn.Execute Aux
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    Aux = "UPDATE smailpromoca set situacion=1 ,usuacepta =" & DBSet(vUsu.Login, "F") & ", fechaAcepta = " & DBSet(Now, "FH")
    Aux = Aux & " WHERE codigo = " & Data1.Recordset!Codigo
    conn.Execute Aux
    Espera 0.5
    GuardarEnSPromo = True
    Exit Function
eGuardarEnSPromo:
    MuestraError Err.Number, , Err.Description
End Function

Private Function ReestablecerEnSPromo() As Boolean
Dim Aux As String
    On Error GoTo eGuardarEnSPromo
    ReestablecerEnSPromo = False
    'Cogeremos de spromo u grabaremos los valores en las columnas
    Aux = "select s.codartic,precio_Sprees, fechainiart,fechafinart"
    Aux = Aux & " from smailpromoli s ,spromo, sartic where s.codartic=spromo.codartic and s.codartic=Sartic.codartic and spromo.codlista=1"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic
    While Not miRsAux.EOF
        'GRao spromo y sartic pmv
        
        Aux = "UPDATE spromo set precioac=" & DBSet(miRsAux!precio_Sprees, "N")
        Aux = Aux & " , FechaIni =" & DBSet(miRsAux!fechainiart, "F")
        Aux = Aux & " , FechaFin=" & DBSet(miRsAux!fechafinart, "F")
        Aux = Aux & " WHERE codlista = 1 AND codartic=" & DBSet(miRsAux!codArtic, "T")
        conn.Execute Aux
        
        Aux = "UPDATE Sartic set preciominvta=" & DBSet(miRsAux!precio_Sprees, "N")
        Aux = Aux & " WHERE  codartic=" & DBSet(miRsAux!codArtic, "T")
        conn.Execute Aux
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    Aux = "UPDATE smailpromoca set situacion=2 ,usurestaura  =" & DBSet(vUsu.Login, "F") & ", fecharestaura = " & DBSet(Now, "FH")
    Aux = Aux & " WHERE codigo = " & Data1.Recordset!Codigo
    conn.Execute Aux
    ReestablecerEnSPromo = True
    Exit Function
eGuardarEnSPromo:
    MuestraError Err.Number, , Err.Description
End Function
