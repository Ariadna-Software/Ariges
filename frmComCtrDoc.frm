VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComCtrDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de albaranes compras"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16425
   ClipControls    =   0   'False
   Icon            =   "frmComCtrDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   16425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   25
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   26
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14625
      TabIndex        =   24
      Top             =   225
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   12840
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Observaciones|T|S|||sctrcompr|observa||N|"
      Text            =   "obas"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   11160
      TabIndex        =   23
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   8520
      TabIndex        =   22
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   10080
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "Fecha arhivo|F|S|||sctrcompr|fechaarch|dd/mm/yyyy|N|"
      Text            =   "archivo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   11520
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Trab. archivo|N|S|||sctrcompr|codtrab1|0|N|"
      Text            =   "trab"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7965
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   9990
      Width           =   4740
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   9990
      Width           =   4860
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Trabajador recepcion|N|N|||sctrcompr|codtraba|0|N|"
      Text            =   "trrecep"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmComCtrDoc.frx":000C
      Left            =   4320
      List            =   "frmComCtrDoc.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Tipo|N|N|||sctrcompr|tipodoc|||"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   17
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   7440
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Fecha recepcion|F|N|||sctrcompr|fechalleg|dd/mm/yyyy|N|"
      Text            =   "recepcion"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Documento|T|N|||sctrcompr|documento||S|"
      Text            =   "documento"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4770
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "Fecha documento|F|N|||sctrcompr|fechadoc|dd/mm/yyyy|S|"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   9720
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   195
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   14
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Proveedor|N|N|||sctrcompr|codprove|0|S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
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
      Height          =   375
      Left            =   13950
      TabIndex        =   9
      Top             =   9870
      Width           =   1065
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
      Height          =   375
      Left            =   15225
      TabIndex        =   10
      Top             =   9870
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   12720
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Bindings        =   "frmComCtrDoc.frx":0024
      Height          =   8865
      Left            =   225
      TabIndex        =   11
      Top             =   810
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   15637
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   5130
      Picture         =   "frmComCtrDoc.frx":0039
      Tag             =   "-1"
      ToolTipText     =   "Buscar trabajador"
      Top             =   9720
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador archivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   7965
      TabIndex        =   21
      Top             =   9750
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador recepción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3000
      TabIndex        =   20
      Top             =   9750
      Width           =   2100
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   9945
      Tag             =   "-1"
      ToolTipText     =   "Buscar trabajador"
      Top             =   9720
      Width           =   240
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
      TabIndex        =   12
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Attribute VB_Name = "frmComCtrDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 '%=%=frmComProveedores
Attribute frmProv.VB_VarHelpID = -1

Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As String



Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    EsBusqueda = True
                    
                    CadenaBusqueda = " AND sctrcompr.codprove= " & txtAux(0).Text & " AND fechadoc=" & DBSet(txtAux(1).Text, "F") & " AND documento = " & DBSet(txtAux(2).Text, "T")
                    
                    CargaGrid True
                    mnNuevo_Click
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion Data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
             End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
Dim J As Integer
    HaDevueltoDatos = ""
    Select Case Index
        Case 0
            
            Set frmProv = New frmBasico2
            AyudaProveedores frmProv, txtAux(0)
            Set frmProv = Nothing
            If HaDevueltoDatos <> "" Then
                txtAux(0).Text = RecuperaValor(HaDevueltoDatos, 1)
                txtAux2(0).Text = RecuperaValor(HaDevueltoDatos, 2)
            End If
            
        Case 1, 2, 3 'fecha
            Set frmF = New frmCal
            frmF.Fecha = Now
            If Index = 3 Then
                J = 5
            ElseIf Index = 2 Then
                J = 3
            Else
                J = 1
            End If
            
            If txtAux(J).Text <> "" Then frmF.Fecha = CDate(txtAux(J).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
            If HaDevueltoDatos <> "" Then
                
                txtAux(J).Text = HaDevueltoDatos
                PonerFoco txtAux(J)
            End If

    End Select
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
            PonerFocoGrid DataGrid1
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub






Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub Data1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            txtAux2(1).Text = DBLet(Data1.Recordset!NomTraba, "T")
            txtAux2(2).Text = DBLet(Data1.Recordset!NomTraba2, "T")
        End If
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    imgBuscar(1).Picture = imgBuscar(0).Picture

    'ICONOS de La toolbar
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    

    Ordenacion = " ORDER BY codprove,tipodoc,fechadoc,documento "
    CadenaConsulta = MontaSQLCarga(False)
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
   
        PonerModo 0
 
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    'numalbar,fechaalb,codprove,nomprove,fecenvio,if(docarchiv=1,"SI","")
    tots = "S|txtAux(0)|T|Código|800|;S|cmdAux(0)|B||0|;S|txtAux2(0)|T|Proveedor|3500|;S|Combo1|C|Tipo|700|;"
    tots = tots & "S|txtAux(1)|T|F.Documen|1350|;S|cmdAux(1)|B||0|;S|txtAux(2)|T|Documento|1350|;"
    'trab 1
    tots = tots & "S|txtAux(3)|T|F.Recepción|1350|;S|cmdAux(2)|B||0|;S|txtAux(4)|T|Trab.Rec|1000|;"
    tots = tots & "S|txtAux(5)|T|F.Archivo|1350|;S|cmdAux(3)|B||0|;S|txtAux(6)|T|Trab.Arc|1000|;"
    tots = tots & "S|txtAux(7)|T|Obs|3050|;N|||||;N|||||;"
    arregla tots, DataGrid1, Me, 350


'    'dtos alineados a la dcha
'    DataGrid1.Columns(6).Alignment = dbgCenter

    DataGrid1.ScrollBars = dbgNone
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim JJ As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For JJ = 0 To txtAux.Count - 1
        If JJ = 0 Then
            txtAux2(JJ).Height = Me.DataGrid1.RowHeight
            txtAux2(JJ).Top = alto
            txtAux2(JJ).visible = b
        End If
        txtAux(JJ).Height = DataGrid1.RowHeight
        txtAux(JJ).Top = alto
        txtAux(JJ).visible = b
    Next JJ

    Me.Combo1.visible = b
    Me.Combo1.Top = alto
    
    For JJ = 0 To Me.cmdaux.Count - 1
        Me.cmdaux(JJ).Height = Me.DataGrid1.RowHeight
        Me.cmdaux(JJ).Top = alto
        Me.cmdaux(JJ).visible = b
    Next JJ
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    HaDevueltoDatos = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = CadenaSeleccion
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indCodigo As Integer
'        Set frmT = New frmAdmTrabajadores
'        frmT.DatosADevolverBusqueda = "0|1|"
'        frmT.Show vbModal
        If Index = 0 Then
            indCodigo = 4
        Else
            indCodigo = 6
        End If
        
        
        Set frmT = New frmBasico2
        AyudaTrabajadores frmT, txtAux(indCodigo)
        Set frmT = Nothing
        
        
        If HaDevueltoDatos <> "" Then
        
            txtAux2(Index + 1).Text = RecuperaValor(HaDevueltoDatos, 2)
            HaDevueltoDatos = RecuperaValor(HaDevueltoDatos, 1)
            If Index = 0 Then
                txtAux(4).Text = HaDevueltoDatos
                HaDevueltoDatos = 4
            Else
                txtAux(6).Text = HaDevueltoDatos
                HaDevueltoDatos = 6
            End If
            PonerFoco txtAux(CInt(HaDevueltoDatos))
        End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub



Private Sub mnNuevo_Click()

    LimpiarCampos
    DataGrid1.AllowAddNew = True
    PonerModo 3
    AnyadirLinea DataGrid1, Me.Data1
    LLamaLineas ObtenerAlto(Me.DataGrid1, 10)
    
    
    'Abril 2015
    'Pongo el trabajador conectado y la fecha de hoy
    txtAux(3).Text = Format(Now, "dd/mm/yyyy")
    txtAux(4).Text = PonerTrabajadorConectado(CadenaConsulta)
    Me.txtAux2(1).Text = CadenaConsulta
    PonerFoco txtAux(0)
    
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1 'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            BotonEliminar
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 8 'Imprimir
'            BotonImprimir
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
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)

                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(2), (Modo = 4)
    Me.cmdaux(0).Enabled = (Modo <> 4)
    Me.cmdaux(1).Enabled = (Modo <> 4)
    Combo1.Enabled = Modo <> 4
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    imgBuscar(0).visible = b
    imgBuscar(1).visible = b


    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    
    'Insertar
    b = (Modo = 2) Or Modo = 0
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    
    
    b = (Modo = 1 Or (Modo >= 3))
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    ' imprimir
    Toolbar1.Buttons(8).Enabled = False
    
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
    PonerCampos
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
    
    SQL = "select sctrcompr.codprove,nomprove,if(tipodoc=1,""ALB"",""PED"") tipdoc,fechadoc,documento,"
    SQL = SQL & " fechalleg,sctrcompr.codtraba,fechaarch,sctrcompr.codtrab1,observa,straba.nomtraba,straba_1.nomtraba nomtraba2"
    SQL = SQL & " FROM   ((`sctrcompr` `sctrcompr` INNER JOIN `sprove` `sprove` ON `sctrcompr`.`codprove`=`sprove`.`codprove`)"
    SQL = SQL & " LEFT OUTER JOIN `straba` `straba` ON `sctrcompr`.`codtraba`=`straba`.`codtraba`)"
    SQL = SQL & " LEFT OUTER JOIN `straba` `straba_1` ON `sctrcompr`.`codtrab1`=`straba_1`.`codtraba`"
    
    SQL = SQL & " where  1=1 "
    SQL = SQL & ""
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND  sctrcompr.codprove = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 25)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
 

    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")

    txtAux2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    If UCase(DBLet(DataGrid1.Columns(2).Value, "T")) = "ALB" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    
    
    
    txtAux(1).Text = DBLet(Data1.Recordset!FechaDoc, "F")
    
    txtAux(2).Text = DBLet(Data1.Recordset!Documento, "T")
    txtAux(3).Text = DBLet(Data1.Recordset!fechalleg, "F")
    If IsNull(Data1.Recordset!CodTraba) Then
        txtAux(4).Text = ""
    Else
        txtAux(4).Text = Data1.Recordset!CodTraba
    End If
    txtAux2(1).Text = DBLet(Data1.Recordset!NomTraba, "T")
    
    If IsNull(Data1.Recordset!CodTrab1) Then
        txtAux(6).Text = ""
    Else
        txtAux(6).Text = Data1.Recordset!CodTrab1
    End If
    txtAux2(2).Text = DBLet(Data1.Recordset!NomTraba2, "T")

    If IsNull(Data1.Recordset!fechaarch) Then
        txtAux(5).Text = ""
    Else
        txtAux(5).Text = Data1.Recordset!fechaarch
    End If


    DataGrid1.Enabled = False
    PonerFoco txtAux(4)
End Sub





Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cad As String

    On Error GoTo ErrDatosOK

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
'    If Me.Combo1.ListIndex = 1 Then
'        If txtAux(3).Text = "" Then
'            if msgbox(2
    
    If Modo = 3 Then
        'Insertar. Veamos si existe un documento en pedido o albaran
        If Me.Combo1.ListIndex = 0 Then
            If Not IsNumeric(txtAux(2).Text) Then
                MsgBox "Nº documento(pedido) debe ser numerico", vbExclamation
                Exit Function
            End If
            
            cad = "numpedpr = " & txtAux(2).Text & " AND codprove "
            cad = DevuelveDesdeBD(conAri, "numpedpr", "scappr", cad, txtAux(0).Text)
            
        Else
            cad = " numalbar = " & DBSet(txtAux(2).Text, "T") & " AND codprove"
            cad = DevuelveDesdeBD(conAri, "numalbar", "scaalp", cad, txtAux(0).Text)
        End If
        If cad = "" Then
            cad = "No existe el documento(" & Combo1.Text & "): " & txtAux(2).Text & "        ¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then b = False
        End If
    End If
    DatosOk = b
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function




Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
 '   If chkVistaPrevia = 1 Then
 '       MandaBusquedaPrevia cadB
 '   ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaBusqueda = " AND " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        'CadenaBusqueda = " AND " & cadB
        PonerCadenaBusqueda
 '   End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
        
        If Data1.Recordset.RecordCount > 14 Then DataGrid1.ScrollBars = dbgVertical
        
    End If
    
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    'PonerCamposForma Me, Data1
    CargaGrid True
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If Index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
                           
        'Caption = KeyCode
        Case 65
            If (Shift And vbCtrlMask) > 0 Then LanzaImg Index
'--
'        Case vbKeyAdd, 43, 187
'            LanzaImg Index
    End Select
End Sub

Private Sub LanzaImg(txtIndex As Integer)
Dim i As Integer
    i = -1
    If txtIndex = 2 Then i = 0
    If txtIndex = 4 Then i = 2
    If txtIndex = 3 Then i = 1
    If i >= 0 Then cmdAux_Click i
    
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'   KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'proveedor
            Case 1: KEYBusqueda KeyAscii, 1 'fec documento
            Case 3: KEYBusqueda KeyAscii, 2 'fec recepcion
            Case 5: KEYBusqueda KeyAscii, 3 'fec archivo
            
            
            Case 4: KEYBusqueda2 KeyAscii, 0 'trabajador
            Case 6: KEYBusqueda2 KeyAscii, 1 'trabajador
            
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cad As String

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    
    If txtAux(Index).Text = "+" Then txtAux(Index).Text = ""
    Select Case Index
        Case 0
            cad = ""
            If txtAux(Index).Text <> "" Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                    txtAux(Index).Text = ""
                Else
                    'If Index = 3 Then
                        cad = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text)
                    'Else
                    '    Cad = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", txtAux(Index).Text)
                    'End If
                    If cad = "" Then MsgBox "No existe el valor en la BD: " & txtAux(Index).Text, vbExclamation
                End If
                If cad = "" And txtAux(Index).Text <> "" Then
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
                      
            End If
            txtAux2(Index).Text = cad
        Case 1, 3, 5 'fecha
              PonerFormatoFecha txtAux(Index)
              
        Case 4, 6
            cad = ""
            If PonerFormatoEntero(txtAux(Index)) Then
                cad = PonerNombreDeCod(txtAux(Index), conAri, "straba", "nomtraba", "codtraba", "Trabajadores", "N")
                If cad = "" And Modo <> 1 Then
                    PonerFoco txtAux(Index)
                    txtAux(Index).Text = ""
                End If
            Else
                txtAux(Index).Text = ""
            End If
            If Index = 4 Then
                txtAux2(1).Text = cad
            Else
                txtAux2(2).Text = cad
            End If
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub






Private Sub BotonEliminar()
Dim cad As String
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    cad = "Va a eliminar el registro: " & vbCrLf
    cad = cad & "------------------------------" & vbCrLf & vbCrLf
    
    cad = cad & vbCrLf & "Proveedor:   " & Format(Data1.Recordset.Fields(0), "000000") & " " & Data1.Recordset.Fields(1)
    cad = cad & vbCrLf & "Documento:   " & Data1.Recordset.Fields(2) & " " & Data1.Recordset.Fields(4)
    cad = cad & vbCrLf & vbCrLf & " ¿Continuar? "
    
    
    
    
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        
        cad = "DELETE FROM `sctrcompr` WHERE `codprove`=" & Data1.Recordset!Codprove & " and `tipodoc`="
        If Data1.Recordset.Fields(2) = "ALB" Then
            cad = cad & "1"
        Else
            cad = cad & "0"
        End If
        cad = cad & " AND `fechadoc`=" & DBSet(Data1.Recordset!FechaDoc, "F") & " and `documento`=" & DBSet(Data1.Recordset!Documento, "T")
        
        
    
        If ejecutar(cad, False) Then
            CargaGrid True
            SituarDataTrasEliminar Data1, NumRegElim
        End If
    
        
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Trabajador", Err.Description
End Sub


