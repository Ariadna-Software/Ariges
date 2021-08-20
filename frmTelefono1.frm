VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelefono1 
   Caption         =   "Utilidades de telefonía"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   Icon            =   "frmTelefono1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVerdatos 
      Height          =   9780
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10440
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   2025
         TabIndex        =   41
         Top             =   9000
         Width           =   840
         Begin MSComctlLib.Toolbar Toolbar5 
            Height          =   330
            Left            =   135
            TabIndex        =   42
            Top             =   180
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Facturar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   90
         TabIndex        =   39
         Top             =   9000
         Width           =   1875
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   90
            TabIndex        =   40
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
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
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
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
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   840
         Picture         =   "frmTelefono1.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Imprimir"
         Top             =   9030
         Width           =   495
      End
      Begin VB.CommandButton cmdBus 
         Height          =   495
         Left            =   1440
         Picture         =   "frmTelefono1.frx":7254
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Buscar teléfono"
         Top             =   9030
         Width           =   495
      End
      Begin VB.CommandButton cmdAriadna 
         Caption         =   "Ariadna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7650
         TabIndex        =   35
         Top             =   9075
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdEliminarDatosFracion 
         Height          =   495
         Left            =   120
         Picture         =   "frmTelefono1.frx":DAA6
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Eliminar Datos Facturacion"
         Top             =   9030
         Width           =   495
      End
      Begin VB.CheckBox chkMostrarBase 
         Alignment       =   1  'Right Justify
         Caption         =   "Base imponible"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3645
         TabIndex        =   33
         Top             =   9255
         Width           =   1755
      End
      Begin VB.ComboBox cboFichero 
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
         Index           =   0
         Left            =   6030
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   315
         Width           =   4110
      End
      Begin VB.CommandButton cmdFacturar 
         Caption         =   "Facturar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   14
         Top             =   9120
         Width           =   540
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   8865
         TabIndex        =   11
         Top             =   9075
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwTF 
         Height          =   7845
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   13838
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   3599
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   9242
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Plazos"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "OrdenTotal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Agr"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TfnosImplicados"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ImporteAlbaranes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ImporteVtaPlazos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "BaseImpo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "CodClien"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblI 
         Caption         =   "Datos fichero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   36
         Top             =   8760
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "Ficheros disponibles:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   8760
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Ficheros disponibles:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3780
         TabIndex        =   15
         Top             =   360
         Width           =   2145
      End
      Begin VB.Label Label1 
         Caption         =   "Datos pre-factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   2385
      End
   End
   Begin VB.Frame FrameImportacion 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox cboCompanyia2 
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
         ItemData        =   "frmTelefono1.frx":E4A8
         Left            =   120
         List            =   "frmTelefono1.frx":E4AF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
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
         Index           =   1
         Left            =   6600
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
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
         Left            =   5160
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   6060
      End
      Begin VB.TextBox Text2 
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
         Index           =   0
         Left            =   6360
         TabIndex        =   4
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   7830
         Picture         =   "frmTelefono1.frx":E4BD
         ToolTipText     =   "Buscar fecha"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   4725
         Picture         =   "frmTelefono1.frx":E548
         Tag             =   "-1"
         ToolTipText     =   "Buscar ruta"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblinf 
         Alignment       =   2  'Center
         Caption         =   "Información de proceso"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Importación ficheros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   2
         Left            =   2070
         TabIndex        =   10
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha emisión facturas importadas:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre cualificado del fichero de importación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   4875
      End
   End
   Begin VB.Frame FrameDtosTelefonia 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   7230
      Begin VB.CommandButton cmdListadoDto 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3990
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboFichero 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Width           =   2760
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   5550
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   6
         Left            =   405
         TabIndex        =   20
         Top             =   360
         Width           =   6090
      End
   End
   Begin MSComDlg.CommonDialog cmmDia 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCatadau 
      Height          =   2895
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox txtCatadau2 
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
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1440
         Width           =   345
      End
      Begin VB.CommandButton cmdCSV 
         Caption         =   "Aceptar"
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
         Left            =   4260
         TabIndex        =   26
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
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
         Index           =   5
         Left            =   5580
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtCatadau2 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   960
         Width           =   5715
      End
      Begin VB.Label lblInf2_ANTIGUO 
         Caption         =   "Información de proceso"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Dígito factura"
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
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Importación CSV"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   8
         Left            =   1080
         TabIndex        =   29
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   930
         Picture         =   "frmTelefono1.frx":EF4A
         ToolTipText     =   "Buscar ruta"
         Top             =   960
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTelefono1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.- Importar=FALS
    ' 1.- Importar
    
    ' 2.- Listado descuentos comprataiivo copera
    ' 3.- Rsumen fracion
    ' 4.- Datos face

    
    ' 5.- Importacion CATADAU

    ' 6.- Datos importados

Dim Cad As String
Dim i As Integer

Dim ImporteAuxiliar As Currency

Dim IVA_standard As Currency

Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1



Private Sub cboCompanyia2_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboFichero_Click(Index As Integer)
    If Index = 0 Then
        If Me.cboFichero(Index).ListIndex < 0 Then Exit Sub
        Screen.MousePointer = vbHourglass
        CargarListView cboFichero(Index).List(cboFichero(Index).ListIndex)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdAriadna_Click()
Dim Normales As Boolean
Dim idBanco As Integer
Dim SQL As String

        Normales = True
        If MsgBox("Normales", vbQuestion + vbYesNo) <> vbNo Then Normales = False

        SQL = InputBox("banco", "", "1")
        If SQL = "" Then Exit Sub
        idBanco = Val(SQL)
        
        SQL = InputBox("Fichero", "", "")
        If SQL = "" Then Exit Sub
        
        EstableceValoresFacturaTelefoniaROOT SQL
        
        'cboFichero(0).List(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
        '(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
         GenerarFacturasTelefonia idBanco, Label1(5), Normales, False

End Sub

Private Sub cmdBus_Click()
    If lwTF.ListItems.Count = 0 Then Exit Sub
    
   
    
    Cad = InputBox("Telefono a buscar", "Telefonia", "")
    Cad = Trim(Cad)
    If Cad = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    'select t.serie,t.ano,t.numfact from tel_cab_factura t, tel_cab_factura_agr a where t.Serie=a.serie and t.Ano=a.Ano and t.NumFact=a.NumFact and fichero='CI0915850635' and a.telefono='644056126'
    lwTF.Tag = " AND fichero = '" & cboFichero(0).Text & "' AND a.telefono"
    lwTF.Tag = "t.Serie=a.serie and t.Ano=a.Ano and t.NumFact=a.NumFact  " & lwTF.Tag
    
    lwTF.Tag = DevuelveDesdeBD(conAri, "concat(t.Serie,'|',t.Ano,'|' ,t.NumFact,'|')", " tel_cab_factura t, tel_cab_factura_agr a", lwTF.Tag, Cad, "T")
    
    If lwTF.Tag = "" Then
        Cad = "No se ha encontrado al telefono " & Cad & " en este fichero"
        MsgBox Cad, vbExclamation
        i = lwTF.ListItems.Count + 1
    Else
    
        Cad = "Serie = '" & RecuperaValor(lwTF.Tag, 1) & "' AND Ano =" & RecuperaValor(lwTF.Tag, 2) & " AND NumFact =" & RecuperaValor(lwTF.Tag, 3)
        For i = 1 To lwTF.ListItems.Count
            'IT.Tag = "Serie = '" & miRsAux!Serie & "' AND Ano =" & miRsAux!Ano & " AND NumFact =" & miRsAux!NumFact
            If lwTF.ListItems(i).Tag = Cad Then Exit For
            
        Next
    
    End If
    Screen.MousePointer = vbDefault
    If i <= lwTF.ListItems.Count Then
        lwTF.ListItems(i).EnsureVisible
        lwTF.ListItems(i).Selected = True
        Set lwTF.SelectedItem = lwTF.ListItems(i)
        PonerFocoOBj lwTF
    End If
    
End Sub

Private Sub cmdCSV_Click()

    MsgBox "Aqui no deberia entrar"
End Sub

Private Sub HacerCoarval()

    
    
    'ANTES
    'cad = ""
    'If Trim(Me.txtCatadau(0).Text) = "" Or Me.txtCatadau(1).Text = "" Then
    '    cad = "Campos obligatorios"
    'Else
    '    If Not IsNumeric(Me.txtCatadau(1).Text) Then
    '        cad = "Dígito incorrecto"
    '    Else
    '        If Dir(Me.txtCatadau(0).Text, vbArchive) = "" Then cad = "No existe el archivo"
    '    End If
    'End If
    
    'AHORA
    
    i = -1
    Cad = DevuelveDesdeBD(conAri, "DigitoCoarval", "spara2", "1", "1")
    If Cad <> "" Then i = Val(Cad)
    
    If i < 0 Then
        Cad = "-No esta establecido el digito de coarval"
    Else
        Cad = ""
    End If
    
    If Text1.Text = "" Then
        Cad = "-Falta fichero"
    Else
        If Dir(Text1.Text, vbArchive) = "" Then Cad = "-No existe el archivo:" & Text1.Text
    End If
    
    If Text2(0).Text <> "" Then
        Cad = "-La fecha de las facturas la lleva el fichero. NO indique ninguna" & vbCrLf & Cad
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
            
    Cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", "codtipom", "ALT", "T")
    If Cad = "" Then Cad = "0"
    
    If Val(Cad) > 0 Then
        MsgBox "Albaranes telefonia pendientes de facturar. Avise soporte técnico", vbExclamation
        Exit Sub
    End If
    
    Cad = PonerTrabajadorConectado("")
    If Cad = "" Then
        MsgBox "Error obteniendo datos trabajador conectado", vbExclamation
        Exit Sub
    End If
            
            
    Cad = "Continuar con la generacion de facturas de telefonía con el digito " & i & "?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If GenerarImportacionCatadau(i) Then InsertarFacturasTelefonoCoarval
    
    Me.lblInf.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEliminarDatosFracion_Click()
Dim mGen2 As TelGenerador
Dim cT As CTiposMov
Dim NumOld As Long

    If cboFichero(0).ListCount = 0 Then Exit Sub
    If Me.lwTF.ListItems.Count = 0 Then Exit Sub
    If vUsu.Nivel > 1 Then Exit Sub
    
    Cad = "Desea eliminar el fichero de facturacion: " & cboFichero(0).Text & "?"
    If MsgBox(Cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    If MsgBox("Seguro que desea eliminar el fichero?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    If MsgBox("Proceso irreversibe. ¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Set mGen2 = New TelGenerador

'mGen2.EliminarTodoElFichero "CI0915168016", Label1(5)


    mGen2.EliminarTodoElFichero cboFichero(0).Text, Label1(5)
    
    Set mGen2 = Nothing
    lwTF.ListItems.Clear
    
    Set cT = New CTiposMov
    cT.Leer "FAT"
    
    'cad = "fecfactu <= " & Format(vEmpresa.FechaFin, FormatoFecha) & " AND
    Cad = ""
    Cad = Cad & "  fecfactu >= " & Format(vEmpresa.FechaIni, FormatoFecha) & " AND codtipom "
    Cad = DevuelveDesdeBD(conAri, "concat(max(numfactu),'|',max(fecfactu),'|')", "scafac", Cad, "FAT", "T")
     If Cad = "" Then Cad = "||"
    If Cad = "||" Then
        Cad = ""
        NumOld = -1
    Else
        
        NumOld = CLng(RecuperaValor(Cad, 1))
        If vParamAplic.NumeroInstalacion <> vbTaxco Then
            If DevuelveDesdeBD(conAri, "DigitoCoarval", "spara2", "1", "1") <> "" Then
                
                NumOld = Val(Mid(CStr(NumOld), 3)) 'quitamos los dos primeros digitos
            End If
        End If
        Cad = RecuperaValor(Cad, 1) & "    max fecha: " & Format(RecuperaValor(Cad, 2), "dd/mm/yyyy")
    End If
    Cad = vbCrLf & vbCrLf & "Contadores:  " & cT.Contador & vbCrLf & "Facturas:  " & Cad & vbCrLf
    Cad = "REVISE LOS CONTADORES DE FACTURA    " & Cad & vbCrLf
    If NumOld > 0 Then Cad = Cad & "  ---- ACTUALIZAR contador FAT: " & NumOld
    
    Cad = String(45, "*") & vbCrLf & vbCrLf & Cad & vbCrLf & vbCrLf & String(45, "*")
    
    
    
    
    
   
    If NumOld = -1 Then
        MsgBox Cad, vbCritical
    Else
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
            cT.Contador = NumOld - 1
            cT.IncrementarContador cT.TipoMovimiento
        End If
    End If
     Set cT = Nothing
    Screen.MousePointer = vbDefault
    

    
    Unload Me
End Sub

Private Sub cmdFacturar_Click()
    'Algun dato a traspasar
    If Me.lwTF.ListItems.Count = 0 Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Label1(5).Caption = "Comprobar"
    Label1(5).Refresh

    HacerFacturacionTelefonia
    
    Label1(5).Caption = ""
    Screen.MousePointer = vbDefault


End Sub

Private Sub HacerFacturacionTelefonia()
Dim B As Boolean
Dim J As Byte
Dim Col As Collection
Dim F As Date
Dim CambiaArticuloLineasFactura As Boolean

    'Primera comprobacion
    'No puede haber ningun albaran "ALT" pendiente de facturar
    
    'Ocubre 2013
    'Se facturan desde tel_cabfactura y los albaranes asociados al numero de telefono, SIN MIRAR fechas
    'Solo comprobare que estan marcados para facturar
    'Ademas comprobaremos que los albaranes tienen el numero correcto de socio/telefono/departamento
    On Error GoTo eHacerFacturacionTelefonia
    
    

    
    Cad = ""
    For NumRegElim = 1 To Me.lwTF.ListItems.Count
        If lwTF.ListItems(NumRegElim).Bold Then
            If lwTF.ListItems(NumRegElim).ForeColor = vbRed Then Cad = Cad & "-" & Me.lwTF.ListItems(NumRegElim).Text & " " & lwTF.ListItems(NumRegElim).SubItems(1) & vbCrLf
        End If
    Next NumRegElim
    If Cad <> "" Then
        Cad = "Estos telefonos tienen albaranes pero no estan marcados para facturar: " & vbCrLf & Cad
        Cad = Cad & vbCrLf & "*** ¿Seguro que desea continuar?"
    Else
        'NUEVO oCT 2013
        Cad = "factursn=0 AND codtipom"
        Cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", Cad, "ALT", "T")
        If Cad = "" Then Cad = "0"
        i = 0
        If Val(Cad) > 0 Then
            Cad = "Existen albaranes sin marca de facturar " & vbCrLf
            Cad = Cad & vbCrLf & "***    ¿Continuar?      ****"
        Else
            Cad = ""
        End If
    End If
       
    If Cad <> "" Then
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    'Si todos los telefonos esta asociados a un telefono/cliente ARIGES
    'Veremos si todas las facturas tiene telefono
    Cad = "tel_cab_factura c left join tel_cab_factura_agr l on"
    Cad = Cad & " C.Serie = L.Serie And C.Ano = L.Ano And C.NumFact = L.NumFact"
    Cad = DevuelveDesdeBD(conAri, "count(*)", Cad, "fichero = '" & cboFichero(0).List(cboFichero(0).ListIndex) & "' AND l.numfact is null AND 1", "1")
    If Val(Cad) > 0 Then
        MsgBox "Facturas sin telefono asignado", vbExclamation
        Exit Sub
    End If
    
    Cad = "tel_cab_factura c inner join tel_cab_factura_agr l on"
    Cad = Cad & " C.Serie = L.Serie And C.Ano = L.Ano And C.NumFact = L.NumFact"
    Cad = Cad & " left join sclientfno on IdTelefono=l.Telefono"  'l.telefobno son de tel_cab_factura_agr
    Cad = DevuelveDesdeBD(conAri, "count(*)", Cad, "IdTelefono is null AND fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Telefonos sin asignar a clientes ARIGES", vbExclamation
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    'Comprobaremos que todos los albaranes que YA estan en telefonia, tienen correctos los clientes/departame
    Label1(5).Caption = "Comprobacion alb telefonia"
    Label1(5).Refresh
    
    Cad = "Select * from scaalb where codtipom='ALT' AND factursn=1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
            
            '1º  Que el numero de telefono lo tengo
           
                'select  concat(codclien,'|',if(coddirec is null,'',coddirec),'|') from sclientfno where IdTelefono ='625780666'
            Cad = "concat(codclien,'|',if(coddirec is null,'',coddirec),'|')"
            Cad = DevuelveDesdeBD(conAri, Cad, "sclientfno", "IdTelefono", miRsAux!referenc, "T")
            If Cad = "" Then
                Err.Raise 513, , "No se encuentra referencia: " & DBLet(miRsAux!referenc, "T")
            Else
                'Mismo cliente
                If Val(RecuperaValor(Cad, 1)) = miRsAux!codClien Then
                    Cad = RecuperaValor(Cad, 2)
                   
                    
                    If DBLet(miRsAux!CodDirec, "T") <> Cad Then Err.Raise 513, , "Coddirec incorrectas el nº de telefono: " & DBLet(miRsAux!referenc, "T")
                        
                    
                Else
                    Err.Raise 513, , "Distinto cliente: " & DBLet(miRsAux!referenc, "T")
                End If
            End If
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    

    Cad = DevuelveDesdeBD(conAri, "fecha", "tel_cab_factura", "fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
    If Cad = "" Then
        MsgBox "Error obteniendo fecha factura", vbExclamation
        Exit Sub
    End If
    F = CDate(Cad)
    
    Cad = PonerTrabajadorConectado("")
    If Cad = "" Then
        MsgBox "Error obteniendo datos trabajador conectado", vbExclamation
        Exit Sub
    End If
    
    'Veremos si se solapan las facturas. Obtenemos la fecha
    Label1(5).Caption = "Solapar nº factura"
    Label1(5).Refresh
    
    
    'Veamos series y (min) n1factura
    
    
    Cad = "select serie,min(numfact) minim ,max(numfact) maxi from tel_cab_factura WHERE fichero=" & DBSet(cboFichero(0).List(cboFichero(0).ListIndex), "T") & " group by 1"
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    
    'SERIE|minimo|max]
    While Not miRsAux.EOF
        Cad = DevuelveDesdeBD(conAri, "codtipom", "stipom", "letraser", miRsAux!Serie, "T")
        If Cad = "" Then Err.Raise 513, "No se encuentra letraser=" & miRsAux!Serie
        
        Cad = Cad & "|" & miRsAux!Serie & "|"
        Cad = Cad & DBLet(miRsAux!minim, "N") & "|"
        Cad = Cad & DBLet(miRsAux!maxi, "N") & "|"
        Col.Add Cad
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Col.Count = 0 Then Err.Raise 513, , "Ningun valor devuelto"
    
    
    For J = 1 To Col.Count
            
        CadenaDesdeOtroForm = RecuperaValor(Col.Item(J), 1)
        
        If CadenaDesdeOtroForm = "ALT" Then 'o es FAT o es FAI
        
            'VEremos la fecha de importacion
            NumRegElim = Year(F)
            Cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            Cad = Cad & " AND numfactu between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            Cad = DevuelveDesdeBD(conAri, "max(fecfactu)", "scafac", Cad, CadenaDesdeOtroForm, "T")
            If Cad <> "" Then
                If CDate(Cad) > F Then Err.Raise 513, , "Fecha facturada mayor que fecha factura telefonia"
            End If
            
            
            'Veamos si se solapan numeros de factura
            Cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            Cad = Cad & " AND numfactu between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            Cad = DevuelveDesdeBD(conAri, "count(*)", "scafac", Cad, CadenaDesdeOtroForm, "T")
            If Cad = "" Then Cad = "0"
            If Val(Cad) > 0 Then Err.Raise 513, , "Se solapan " & Cad & " factura(s)"
            
    
    
    
            'Salto de factura. Veremos cual es la ultima fra trasapsada
            Cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            Cad = DevuelveDesdeBD(conAri, "max(numfactu)", "scafac", Cad, CadenaDesdeOtroForm, "T")
            If Cad <> "" Then
                NumRegElim = Val(RecuperaValor(Col.Item(J), 4)) - Val(RecuperaValor(Col.Item(J), 4))
                If NumRegElim > 1 Then Err.Raise 513, , "Salto factura"
            End If
        
    
        Else
            'FARA FAI comprobaremos que no existe NINGUN albaran en scaalb
            'con un numero igaul al de la factura. NO deberia ya que en su momento cogio de scaalb
            
            
            Cad = " numalbar between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            Cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", Cad, "ALI", "T")
            If Cad = "" Then Cad = "0"
            'If Val(Cad) > 0 Then Err.Raise 513, , "Se solapan albaranes internos"
    
    
        End If
    Next J
    
    'Ok , pues adelante
    '
    
    
    Screen.MousePointer = vbDefault
    Label1(5).Caption = ""
    CadenaDesdeOtroForm = ""
    frmListado3.Opcion = 36
    frmListado3.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        
        'Obtenemos la compañia que vamos a facturar
        CambiaArticuloLineasFactura = False
        If vParamAplic.TieneTelefonia2 = 3 Then
            Cad = DevuelveDesdeBD(conAri, "distinct(companyia)", "tel_cab_factura", "Fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
            If Cad = "ORA" Then
                'ORANGE
                Cad = DevuelveDesdeBD(conAri, "artiTelefNorORAN", "spara2", "1", "1")
                If Cad <> "" Then
                    If Cad <> vParamAplic.ArtiTelefonia Then
                        CambiaArticuloLineasFactura = True
                        vParamAplic.ArtiTelefonia = Cad
                    End If
                End If
            Else
                If Cad = "VOD" Then
                    'VODAFONE
                    Cad = DevuelveDesdeBD(conAri, "artiTelefNorVOD", "spara2", "1", "1")
                    If Cad <> "" Then
                        If Cad <> vParamAplic.ArtiTelefonia Then
                            CambiaArticuloLineasFactura = True
                            vParamAplic.ArtiTelefonia = Cad
                        End If
                    End If
                End If
            End If
        End If
            
    'Reestablecemos el articulo de telefonia
    

        
     B = traspasofacturasTelefonia(cboFichero(0).List(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
        
        
        If CambiaArticuloLineasFactura Then
            'Sea como sea, dejo el articulo de telefonia como estaba
            Cad = DevuelveDesdeBD(conAri, "codartictel", "spara1", "1", "1")
            vParamAplic.ArtiTelefonia = Cad
        End If
        
        If B Then
            ACtualizarPuntosTelefonia
            Unload Me
        End If
        
    End If
    
    
eHacerFacturacionTelefonia:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    CadenaDesdeOtroForm = ""
End Sub





Private Sub ACtualizarPuntosTelefonia()
    DoEvents
    Label1(5).Caption = "Ajuste puntos"
    Label1(5).Refresh
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    Cad = "select Telefono,BaseImponible,base_exenta from tel_cab_factura WHERE fichero= '" & cboFichero(0).List(cboFichero(0).ListIndex) & "' ORDER BY Telefono"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF

            'FALTA### deberiamos parametrizar
            'Los puntos
            ' Email Martin: 15/05/13   - 1 punto por cada 0,20 € de Base Imponible.
            
            
                            
            i = CInt((miRsAux!BaseImponible + DBLet(miRsAux!base_exenta, "N")) / 0.2)
            Cad = "UPDATE sclientfno SET puntos = puntos + " & CStr(i)
            Cad = Cad & " WHERE IdTelefono = " & DBSet(miRsAux!Telefono, "T")
            conn.Execute Cad
            miRsAux.MoveNext
    Wend
    miRsAux.Close

End Sub

Private Sub cmdImprimir_Click()
Dim B
    If lwTF.ListItems.Count = 0 Then Exit Sub
    If cboFichero(0).ListCount = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    B = GeneraDatosImpresion
    Label1(5).Caption = ""
    Screen.MousePointer = vbDefault
    If B Then
    
        With frmImprimir
            .FormulaSeleccion = "{tmpinformes.codusu}=" & vUsu.Codigo
            .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|fichero=""" & Me.cboFichero(0).Text & """|"
            .NumeroParametros = 2
    
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = Me.Caption
            .Opcion = 3000   'VAN TODOS EN ESTE SACO
            .NombrePDF = "rprevioFratele.rpt"
            .NombreRPT = .NombrePDF
            .ConSubInforme = False
            .MostrarTreeDesdeFuera = False
            .Show vbModal
        End With
    
    End If
End Sub

Private Sub cmdListadoDto_Click()

    If Me.cboFichero(1).ListIndex < 0 Then Exit Sub
    '
    i = 65
    If Opcion = 3 Then i = 69
    If Opcion = 4 Then
        i = 70
        HacerAccionesDelJOinDeRafa
    
    End If
    
    
    If Opcion = 6 Then
        frmListado4.vCadena = cboFichero(1).Text
        frmListado4.Opcion = 8
        frmListado4.Show vbModal
        Exit Sub
    End If
    
    Cad = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", CStr(i), "N")
    If Cad = "" Then
        MsgBox "Error obtener informe: " & i, vbExclamation
    Else
        
        With frmImprimir
            'Comun
            .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
            .NumeroParametros = 1
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 5
            .ConSubInforme = True
            .NombreRPT = Cad
            Select Case Opcion
            Case 3
                .Titulo = "Resumen facturacion soporte"
                'frmVisReport.FicheroInforme = "C:\Telefonia" & "\Informes\" & "detalle_facturas_soporte.rpt"
                .FormulaSeleccion = "{tel_cab_factura.Fichero} = '" & cboFichero(1).Text & "'"
            Case 4
                .Titulo = "Factura resumen"
                '.FormulaSeleccion = "{tel_cab_factura.Fichero} = '" & cboFichero(1).Text & "'"
                .FormulaSeleccion = "{tmpcrmcobros.codusu} = " & vUsu.Codigo
        
            Case Else
                'DOS
                .FormulaSeleccion = "{tmp_inf_descuentos.Fichero} = '" & cboFichero(1).Text & "'"
                .Titulo = "Estudio descuentos telefonía"
            
            End Select
            .Show vbModal
        End With
    End If
End Sub

Private Sub HacerAccionesDelJOinDeRafa()

    'RAfa tenia un subreporte con un command y dentro un UNION
    ' Como los commands no se pueden enlazar tenemos que cargar
    'en una tmp
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    
    'Cojo el JOIN que habia en el rpt y lo meto aqui
    Cad = ""
    Cad = Cad & "select " & vUsu.Codigo & ",0,0, a.CodCuota as Codigo, a.DescCuota  as Nombre, sum(a.importe), b.PorcentajeOperador as Porc, (sum(a.importe) * b.PorcentajeOperador)/100"
    Cad = Cad & " , b.Porcentaje, (sum(a.importe) * b.Porcentaje)/100 from tel_lin_factura_cuotas as a,"
    Cad = Cad & " tel_desc_cuotas as b, tel_cab_factura As C where A.serie = C.serie and a.NumFact = c.Numfact"
    Cad = Cad & " and a.Ano = c.Ano and a.CodCuota = b.CodCuota and fichero='" & cboFichero(1).Text & "'"
    ''CI0544330498'
    Cad = Cad & " group by c.Fichero, a.CodCuota UNION "
    Cad = Cad & " select " & vUsu.Codigo & ",0,0,a.CodTipoTrafico as Codigo, a.DescTipoTrafico as Nombre, sum(a.importe)"
    Cad = Cad & " , b.PorcentajeOperador as Porc, (sum(a.importe) * b.PorcentajeOperador)/100 "
    Cad = Cad & " , b.Porcentaje, (sum(a.importe) * b.Porcentaje)/100 from tel_lin_factura_consumos as a,"
    Cad = Cad & " tel_desc_consumos as b,tel_cab_factura As C where A.serie = C.serie and a.NumFact = c.Numfact"
    Cad = Cad & " and a.Ano = c.Ano and a.CodTipoTrafico = b.CodTipoTrafico and fichero='" & cboFichero(1).Text & "'"
    Cad = Cad & " group by c.Fichero, a.CodTipoTrafico"
    
    'Lo metemos en tmp
    Cad = "INSERT INTO tmpinformes(codusu,campo1,campo2,nombre1,nombre2,importe1,porcen1,importe2,porcen2,importe3) " & Cad
    conn.Execute Cad
    
    'Para que solo coja un registro
    conn.Execute "DELETE FROM tmpcrmcobros WHERE codusu = " & vUsu.Codigo
    Cad = "INSERT INTO tmpcrmcobros(codusu,secuencial,forpa) VALUES (" & vUsu.Codigo
    Cad = Cad & ",1,'" & cboFichero(1).Text & "')"
    conn.Execute Cad
    
'ORDER BY Porc, Nombre;
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Command1_Click()

            
    If cboCompanyia2.ItemData(cboCompanyia2.ListIndex) > 4 Then
        MsgBox "Proceso no desarrollado", vbExclamation
        Exit Sub
    End If

    If cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 4 Then
        'COARVAL COARVAL
        HacerCoarval
    Else
        'Resto operadores
        
        Screen.MousePointer = vbHourglass
        Me.Command1.Enabled = False
        HacerImportacion
        lblInf.Caption = ""
        Me.Command1.Enabled = True
        Screen.MousePointer = vbDefault

    End If
End Sub






Private Sub HacerImportacion()
    Dim mGen2 As TelGenerador
    Dim resultado As Boolean
    Dim Mens As String
    Dim FicheroOrange As String
    
    
    
    
    Mens = ""
    If Text1.Text = "" Then
        Mens = "Fichero vacio"
    Else
        If Dir(Text1.Text) = "" Then
            Mens = "No existe fichero"
        Else
            'Comprobaremos si el fichero NO ha sido PROCESADO del todo, es decir, metido tb en scafac,slifac...
            'FALTA###
                    
                    
            Mens = Text1.Text
            
            
            If Len(Mens) > 12 Then Mens = Right(Mens, 12)
            

            '1era cosa a tener en cuenta. El fichero no puede estar procesado
            Mens = DevuelveDesdeBD(conAri, "fecha", "tel_fichtraspasados", "Fichero", Mens, "T")
            If Mens <> "" Then Mens = "El fichero se traspaso el dia " & Mens & vbCrLf
                
            
        End If
    End If
        
    
    '-- Controlamos que la fecha de emisión de facturas sea mas o menos correcta
    If Not IsDate(Text2(0)) Then
        Mens = Mens & vbCrLf & "Ha de introducir una fecha de emisión de facturas correcta"
    Else
        If vEmpresa.FechaIni > CDate(Text2(0).Text) Or DateAdd("yyyy", 1, vEmpresa.FechaFin) < CDate(Text2(0).Text) Then _
            Mens = Mens & vbCrLf & "Fuera de ejercicios contables"
            
    End If
    
    If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 3 Then
        If Text1.Text <> "" Then
            If InStr(1, Text1.Text, ".") > 0 Then
                If LCase(Right(Text1.Text, 3)) <> "csv" Then Mens = Mens & vbCrLf & "El fichero de VODAFONE no ebe llevar extension(solo CSV) "
            End If
        End If
    End If
            
    
    If Mens <> "" Then
        Mens = "Campos obligados" & vbCrLf & vbCrLf & Mens
        MsgBox Mens, vbExclamation
        Exit Sub
    End If
    
    '***********************************************************************
    '***********************************************************************
    '***********************************************************************
    '
    ' En referencia de las FAT grabaremos el NUMERO de telefono
    ' Una vez este en tal_cab_Factura entonces cuando vayamos a factura
    ' de tel_cab veremos su hay algun ALT nuevo que se ha generado desde
    ' albaranes de telefonia, con lo cual los contadores tienen que estar
    ' separados (FAT y ALT) ya que el proceso de facturacion desde el num
    ' de factura lo mete en albaranes y de ahi a scafac, entoces si
    ' hubiera algun albaran ALT asociado a ese numereo de telefono(refern)
    ' entonces lo tendria que meter como facturacion colectiva e irian
    ' los dos juntos(o 3 o cuatro...)
    '
    '
    '***********************************************************************
    '***********************************************************************
    '***********************************************************************
    'Solo dejamos UN fichero en 'proceso' de facturacion
    resultado = False
    Mens = Text1.Text
    If Len(Text1) > 12 Then Mens = Right(Mens, 12)
    
    
    
    'Para el proceso de ORANGE, como el nombre del fichero puede variar "demasiado"
    'preprocesaremos el fichero para obtener el numero de factura que
    'esta en la segunda linea. Si el numero de factura ya ha sido procesado entonces
    'daremos el mensaje
    'número de factura:;A10020017274-0813  --> A1002017274
    If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 2 Then
        'ES ORANGE
        Set mGen2 = New TelGenerador
        lblInf.Caption = "obtener nombre unico fichero"
        lblInf.Refresh
        FicheroOrange = mGen2.DevuelveNombreFicheroOrange(Text1.Text)
        Set mGen2 = Nothing
        If FicheroOrange = "" Then
            MsgBox "Imposible localizar datos factura en fichero Orange", vbExclamation
            Exit Sub
        End If
        
        
        
        'EN ORANGE tenemos que comprobar que el fichero NO ha siado traspasado
        Mens = DevuelveDesdeBD(conAri, "fecha", "tel_fichtraspasados", "Fichero", FicheroOrange, "T")
        If Mens <> "" Then
            MsgBox "El fichero se traspaso el dia " & Mens & vbCrLf, vbExclamation
            Exit Sub
        End If
        
        
        'Para la utilizacion posterior
        Mens = FicheroOrange
        
        
        lblInf.Caption = ""
        lblInf.Refresh
        
        
    End If
    
    
    
    
    
    
    
    'Mens = "select distinct(Fichero) from tel_cab_factura where not Fichero in (select Fichero from tel_fichtraspasados)"
    Cad = " not Fichero in (select Fichero from tel_fichtraspasados) AND 1"
    Cad = DevuelveDesdeBD(conAri, "distinct(Fichero)", "tel_cab_factura", Cad, "1")
    If Cad <> "" Then
        'Si el fichero que falta NO es el que estamos intentando pasar
        If Cad <> Mens Then
            MsgBox "Falta procesar el archivo: " & Cad, vbExclamation
            
            Exit Sub
            
        Else
            Cad = "Volver a cargar los datos del fichero: " & Mens & "?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            resultado = True
        End If
    End If
    
    
    
    
    
    
    
    
    'DAVID###
    'EL proceso se divide en:
    '   -proceso antiguo(todo lo que hacia Rafa)
    '   -y luego metemos en la slialb para pasarles el proceso de facturacion NORMAL de
    ' ariges
    'Un fichero puede ser importado muchas veces. Siempre borra los datos etc.
    'Al final, cuando pulse el boton de llevar a scafac, una vez haga esto, YA no puede volver a importar
    ' el fichero
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    
    Cad = String(40, "*") & vbCrLf
    Cad = Cad & Cad & vbCrLf & vbCrLf
    Mens = Cad & "Va a importar el fichero de telefonía:"
    Mens = Mens & vbCrLf & vbCrLf & "Compañia: " & Me.cboCompanyia2.Text
    Mens = Mens & vbCrLf & vbCrLf & "FECHA: " & Text2(0).Text & vbCrLf & vbCrLf & Cad
    
    
    'En resultado tenemos si ya ha hecho la pregunta de procesar, para que no la vuelva a hacer
    If Not resultado Then
        If MsgBox(Mens, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    resultado = False
    
    
    Set mGen2 = New TelGenerador
    If mGen2.LeerParametrosFacturacionTelefonica(cboCompanyia2.ItemData(cboCompanyia2.ListIndex)) Then
    
        Screen.MousePointer = vbHourglass
         resultado = False
        If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 1 Then
            resultado = mGen2.cargarBaseDatosMOVISTAR(Text1, lblInf)
            'resultado = True
        ElseIf Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 2 Then
           
            
            'Le paso el fichero fisico, el nombre estara en: FicheroOrange
            resultado = mGen2.cargarBaseDatosOrange(Text1.Text, lblInf)
            
        Else
            'VODAFONE
            resultado = mGen2.cargarBaseDatosVODAFONE(Text1.Text, lblInf)
        End If
        lblInf.Caption = ""
        Screen.MousePointer = vbDefault
        If Not resultado Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        
        'LLamadas entre coooperativistas
        lblInf.Caption = "Prepara datos proceso(II)"
        lblInf.Refresh
        
        
        'Si hay conceptos o cuotas nuevas las mete en tmpinformes para listarlas luego
        'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2)
        conn.Execute "DELETE from tmpinformes WHERE codusu = " & vUsu.Codigo
        
        'LLamadas entre coooperativistas
        lblInf.Caption = "Acciones cooperativa"
        lblInf.Refresh
         
        
        If vParamAplic.TieneTelefonia2 = 3 Then
                mGen2.RecalcularImporteLlamadasCoperativa Me.lblInf, cboCompanyia2.ItemData(cboCompanyia2.ListIndex)
        End If
        
        
         'VEmos cuotas
         resultado = mGen2.AjusteCuotasNuevas2(cboCompanyia2.ItemData(cboCompanyia2.ListIndex), Right(Text1, 12), Me.lblInf)
         
         'refacturamos
         If resultado Then mGen2.ComprobarConceptosFacturacion (cboCompanyia2.ItemData(cboCompanyia2.ListIndex))
        
                
        
        
        If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) <> 2 Then FicheroOrange = Text1.Text 'Para movistar dejo el nombre del fichero
        
        
        ' en vodafo refacturado, el fichero será CSV...
        If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 3 Then
            Mens = DevuelveDesdeBD(conAri, "vodafoneRefacturado", "spara2", "1", "1")
            If Mens = "1" Then
                FicheroOrange = mGen2.DevuelveElNombreFicheroTraspasado
            End If
            Mens = ""
        End If
        
        
        
        CadenaDesdeOtroForm = ""
        Screen.MousePointer = vbDefault
        If Not resultado Then
            Set mGen2 = Nothing
            Exit Sub
        End If
        
        Screen.MousePointer = vbDefault
        DoEvents
        
        resultado = mGen2.EmitirFacturas_(FicheroOrange, Text2(0), Me.lblInf, CByte(Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex)))
        'TRUE=Error
        If resultado Then
            Mens = "Se ha producido incidencias durante el proceso de generación. " & _
                    "Estas incidencias se han guardado en el fichero " & App.Path & "\emitefac.log" & vbCrLf & _
                    "¿Desea ver el contenido de este fichero?"
            If MsgBox(Mens, vbYesNo + vbQuestion) = vbYes Then
                Shell "notepad " & App.Path & "\emitefac.log", vbMaximizedFocus
            End If
        End If

        mGen2.calculaInformeDescuentos Mens
    
    
    
    End If
    
    
    
    Set mGen2 = Nothing
    lblInf.Caption = ""
     
    
    If Not resultado Then
        If vParamAplic.TieneTelefonia2 = 3 Then NuevasCuotasConceptos
    
        If Not resultado Then CadenaDesdeOtroForm = "SI"
    End If
    Unload Me
  
    
    
End Sub



Private Sub NuevasCuotasConceptos()
    
    On Error GoTo eNuevasCuotasConceptos
    Cad = "Select * from tmpinformes where codusu=" & vUsu.Codigo & " ORDER BY campo2,nombre1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not miRsAux.EOF Then
        i = FreeFile
        Cad = App.Path & "\NC" & Format(Now, "yymmddhhnn") & ".txt"
        CadenaDesdeOtroForm = Cad
        Open Cad For Output As #i
        While Not miRsAux.EOF
            
            If miRsAux!Codigo1 = 2 Then
                Cad = "Varios "
            Else
                Cad = IIf(miRsAux!campo1 = 0, "Cuota ", "Conce ")
            End If
            Cad = "    " & Cad & miRsAux!nombre1 & " :  " & miRsAux!nombre2 & vbCrLf
            Print #i, Cad
            miRsAux.MoveNext
        Wend
        Close #i
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If i > 0 Then
        LanzaVisorMimeDocumento Me.hwnd, CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
    End If
    
eNuevasCuotasConceptos:
   If Err.Number <> 0 Then MsgBox "Error leyendo cuotas nuevas creadas", vbExclamation
End Sub



Private Sub Form_Load()

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
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 40 ' facturar
    End With

    
    Me.FrameImportacion.visible = False
    Me.FrameVerdatos.visible = False
    FrameCatadau.visible = False
    i = Opcion
    Select Case Opcion
    Case 0
        lblI.Caption = ""
        Me.lwTF.ColumnHeaders(3).Width = IIf(vParamAplic.TelefoniaVtaPlazos, 900, 0)
        If Not vParamAplic.TelefoniaVtaPlazos Then Me.lwTF.ColumnHeaders(3).Width = Me.lwTF.ColumnHeaders(3).Width + 900
        PonerFrameVisible Me.FrameVerdatos
        CargaCombo Me.cboFichero(0), True
    Case 1
        PonerFrameVisible Me.FrameImportacion
        
        
        CargaComboCompanyia
        
        
    Case 2, 3, 4, 6
        i = 2 'cancelar(2)
        If Opcion = 3 Then
            Label1(6).Caption = "Facturación por soporte"
        ElseIf Opcion = 4 Then
            Label1(6).Caption = "Resumen por soporte"
            
        ElseIf Opcion = 6 Then
            Label1(6).Caption = "Datos importados fichero"
            
        Else
            '2
            Label1(6).Caption = "Estudio descuentos telefonía"
        End If
        
        PonerFrameVisible Me.FrameDtosTelefonia
        CargaCombo Me.cboFichero(1), False
        
'    Case 5
'        PonerFrameVisible FrameCatadau
'        lblInf.Caption = ""
    End Select
    
    cmdSalir(i).Cancel = True
    lblInf.Caption = ""
    
    Label1(5).Caption = ""
    Screen.MousePointer = vbDefault

End Sub

Private Sub PonerFrameVisible(ByRef Fr As Frame)
    Fr.Left = 120
    Fr.Top = 0
    Me.Height = Fr.Height + 510
    Me.Width = Fr.Width + 360
    Fr.visible = True
End Sub

Private Sub CargarListView(Fich As String)
Dim IT As ListItem
Dim Rc As ADODB.Recordset
Dim ImpoAux As Currency
Dim Aux As String
Dim J As Integer
Dim TelefonosFacturar As Integer

    Set miRsAux = New ADODB.Recordset
    Set Rc = New ADODB.Recordset
   
    IVA_standard = -1
    If vParamAplic.TelefoniaVtaPlazos Then
    
        Cad = "Select IdTelefono,PlazosMeses,ArtPlazos,ImportePlazo from sclientfno where PlazosMeses > 0 "
        Rc.Open Cad, conn, adOpenKeyset, adCmdText
        If Not Rc.EOF Then
            Cad = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", DBLet(Rc!artplazos, "T"), "T")
            If Cad <> "" Then
                Cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", Cad, "N")
                If Cad <> "" Then IVA_standard = CCur(Cad)
            End If
        End If
    Else
        If vParamAplic.ArtiTelefonia <> "" Then
            IVA_standard = 0
           Cad = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", DBLet(vParamAplic.ArtiTelefonia, "T"), "T")
            If Cad <> "" Then
                Cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", Cad, "N")
                If Cad <> "" Then IVA_standard = CCur(Cad)
            End If
        End If
    End If
    
    Cad = "select telefono,apellido1, apellido2,nombre,"
    Cad = Cad & " BaseImponible,Cuota,total,Serie ,Ano ,NumFact,EsAgupacion"
    Cad = Cad & "  from tel_cab_factura where fichero='" & Fich & "' order by telefono"
   

    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    lwTF.ListItems.Clear
    
    If Me.chkMostrarBase.Value = 1 Then
        Me.lwTF.ColumnHeaders(4).Text = "B.Imp."
    Else
        Me.lwTF.ColumnHeaders(4).Text = "Total"
    End If
    
    TelefonosFacturar = 0
    
    '++
'    lwTF.SmallIcons = frmPpal.ImgListPpal
    
    
    While Not miRsAux.EOF
    
        
        Cad = ""
        If Not IsNull(miRsAux!apellido1) Then Cad = miRsAux!apellido1
        If Not IsNull(miRsAux!apellido2) Then Cad = Trim(Cad & " " & miRsAux!apellido2)
        If Not IsNull(miRsAux!Nombre) Then
            If Cad <> "" Then Cad = Cad & ","
            Cad = Trim(Cad & " " & miRsAux!Nombre)
        End If
   
        Set IT = lwTF.ListItems.Add()
        
        
        
        
        IT.Text = miRsAux!Telefono
        IT.SubItems(1) = Cad
        
        If Me.chkMostrarBase.Value = 1 Then
            IT.SubItems(3) = Format(miRsAux!BaseImponible, "#,##0.00")
        Else
            IT.SubItems(3) = Format(miRsAux!total, "#,##0.00")
        End If
        IT.SubItems(4) = Format(miRsAux!total * 100, "0000000")
        IT.SubItems(5) = Val(miRsAux!EsAgupacion)
        IT.SubItems(2) = " "
        
        
        IT.Tag = "Serie = '" & miRsAux!Serie & "' AND Ano =" & miRsAux!Ano & " AND NumFact =" & miRsAux!NumFact
        'Abril2020
        If miRsAux!EsAgupacion = 1 Then
            Aux = "count(*)"
            Cad = IT.Tag & " AND 1 "
            Cad = DevuelveDesdeBD(conAri, "GROUP_CONCAT( telefono separator '|')", "tel_cab_factura_agr", Cad, "1", "N", Aux)
            If Cad = "" Then
                Err.Raise 513, , "Error encontrando telefonos para factura: " & IT.Tag
                Me.cmdFacturar.Enabled = False
                Aux = 1
            End If
            TelefonosFacturar = TelefonosFacturar + CInt(Aux)
            Aux = ""
            Cad = Cad & "|"
        Else
            TelefonosFacturar = TelefonosFacturar + 1
            Cad = miRsAux!Telefono
        End If
        IT.SubItems(6) = Cad
        IT.SubItems(7) = 0 'importe albaranes
        IT.SubItems(8) = 0 'importe ventas plazo
        IT.SubItems(9) = 0 'base imponible
        IT.SubItems(10) = miRsAux!BaseImponible
        
        If vParamAplic.TelefoniaVtaPlazos Then
            i = 0
            ImpoAux = 0
            
            While Cad <> ""
                J = InStr(1, Cad, "|")
                If J = 0 Then
                    If Cad = "" Then Err.Raise 513, , "Error leendo telefonos agrupados" & Cad
                    
                    Aux = Cad
                    Cad = ""
                    
                Else
                    Aux = Mid(Cad, 1, J - 1)
                    Cad = Mid(Cad, J + 1)
                End If
                
                Aux = "idtelefono =  '" & Aux & "'"
                Rc.Find Aux, , adSearchForward, 1
                If Not Rc.EOF Then
                    i = 1
                    ImpoAux = ImpoAux + DBLet(Rc!ImportePlazo, "N")
                End If
            Wend
            'La linea tiene albaranes pendientes
            If i > 0 Then
                                    
                    IT.SubItems(8) = ImpoAux
                                   
                    If Me.chkMostrarBase.Value = 0 Then ImpoAux = Round2(ImpoAux * ((100 + IVA_standard) / 100), 2)
                    
                    ImpoAux = ImporteFormateado(IT.SubItems(3)) + ImpoAux
                                        
                    IT.ListSubItems(3).ForeColor = vbBlue
                    IT.ListSubItems(3).Bold = True
                    IT.ListSubItems(3).ToolTipText = "Vta plazo"
                    IT.SubItems(2) = "S"
                    IT.SubItems(3) = Format(ImpoAux, FormatoImporte)
                    IT.SubItems(4) = Format(ImpoAux * 100, "0000000")
            End If
            
        End If
        
        
        
        
        
        
        
        
        'Para el WHERE. Si cambiamos ver boton buscar
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    ComprobarAlbaranesPendientes
    
    
    'Por si se queda la facturacion a medias
    'if False Then
    '   For i = Me.lwT.ListItems.Count To 1 Step -1
    '        If InStr(1, "617716340v629378165v636153242v646603031v661078672v674845196v", lwT.ListItems(i).Text) = 0 Then
    '            lwT.ListItems.Remove i
    '        End If
    '    Next
    'End If
    
    
    lblI.Caption = "Tfnos: " & TelefonosFacturar & "     Facturas : " & lwTF.ListItems.Count
    
    
    
     Set miRsAux = Nothing
    Set Rc = Nothing
End Sub

Private Sub ComprobarAlbaranesPendientes()
Dim Impaux As Currency

    On Error GoTo eComprobarAlbaranesPendientes
    
    
    Cad = "select scaalb.numalbar,referenc,codclien,nomclien,factursn,sum(importel) base from scaalb left join slialb "
    Cad = Cad & " on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    Cad = Cad & " Where scaalb.codtipom='ALT' group by scaalb.numalbar"
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    ImporteAuxiliar = 0
    While Not miRsAux.EOF        'QUe hago....
    
        'If miRsAux!referenc = "687751251" Then St op
    
    
        'Buscamos por el LW e telefono
        For i = 1 To Me.lwTF.ListItems.Count
            If InStr(lwTF.ListItems(i).SubItems(6), miRsAux!referenc) > 0 Then
                'Este es el numero de telefono
                    
                Exit For
            End If
        Next
        
        
        'Si NO lo ha encotrado, lo añado a CAD
        If i > lwTF.ListItems.Count Then
            Cad = Cad & vbCrLf & miRsAux!codClien & " " & miRsAux!NomClien & " -> " & miRsAux!referenc
        Else
            Me.lwTF.ListItems(i).Bold = True
            If miRsAux!factursn = 0 Then
                Me.lwTF.ListItems(i).ForeColor = vbRed
                lwTF.ListItems(i).ToolTipText = "Sin marca de facturar"
            Else
                Me.lwTF.ListItems(i).ForeColor = vbBlue
                lwTF.ListItems(i).ToolTipText = "Albaranes"
                'El total
                
            
                Impaux = 0
                If Me.chkMostrarBase.Value = 0 Then Impaux = IVA_standard
                 
                Impaux = Round2(DBLet(miRsAux!Base, "N") * ((100 + Impaux) / 100), 2)
                
                Impaux = ImporteFormateado(lwTF.ListItems(i).SubItems(3)) + Impaux
                ImporteAuxiliar = CCur(lwTF.ListItems(i).SubItems(7)) + DBLet(miRsAux!Base, "N")
                lwTF.ListItems(i).SubItems(7) = ImporteAuxiliar
                lwTF.ListItems(i).SubItems(3) = Format(Impaux, FormatoImporte)
                lwTF.ListItems(i).SubItems(4) = Format(Impaux * 100, "0000000")
            
                lwTF.ListItems(i).ListSubItems(4).Bold = True
                lwTF.ListItems(i).ListSubItems(4).ForeColor = vbBlue
            End If
        End If
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close


    

    If Cad <> "" Then
        Cad = "Albaranes que no se facturarán: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
    End If
        
    Exit Sub
eComprobarAlbaranesPendientes:
    MsgBox "Avise soporte tecnico. Albaran pdte   : " & miRsAux!Numalbar & "-" & i & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text2(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    cmmDia.ShowOpen
    If Index = 0 Then
        Text1.Text = cmmDia.FileName
    Else
       ' Me.txtCatadau(0).Text = cmmDia.FileName
    End If
End Sub

Private Sub ListView1_DblClick()

End Sub




Private Sub CargaCombo(ByRef CBO As ComboBox, FaltaProcesar As Boolean)

    Set miRsAux = New ADODB.Recordset
    
    CBO.Clear
    Cad = "Select distinct(fichero) from tel_cab_factura"
    If FaltaProcesar Then
        Cad = Cad & " WHERE not fichero in (select fichero from tel_FichTraspasados) ORDER BY fecha desc"
    Else
        
        Cad = Cad & " WHERE fecha > " & DBSet(DateAdd("yyyy", -2, Now), "F") & " ORDER BY fecha desc"
    End If
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & "1"
        CBO.AddItem miRsAux!Fichero
        miRsAux.MoveNext
    Wend
    miRsAux.Close
 
    
    
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text2(Indice)
   If Text2(0).Text <> "" Then frmF.Fecha = CDate(Text2(0).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text2(Indice)
End Sub

Private Sub Label1_Click(Index As Integer)
    If vUsu.Login = "root" Then
            cmdAriadna.visible = True
    End If
End Sub

Private Sub lwTf_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    i = ColumnHeader.Index - 1
    If ColumnHeader.Index = 4 Then i = 4
    
    If i = lwTF.SortKey Then
        If lwTF.SortOrder = lvwAscending Then
            lwTF.SortOrder = lvwDescending
        Else
            lwTF.SortOrder = lvwAscending
        End If
    Else
        If ColumnHeader.Index = 4 Then
            lwTF.SortKey = 4
        Else
            lwTF.SortKey = ColumnHeader.Index - 1
        End If
        lwTF.SortOrder = lvwAscending
    End If
End Sub

Private Sub lwTf_DblClick()
    If lwTF.ListItems.Count = 0 Then Exit Sub
    If lwTF.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    frmTelefonoVerFra.esAgrupacion = lwTF.SelectedItem.SubItems(5) = 1
    frmTelefonoVerFra.TieneAlbaranes = lwTF.SelectedItem.Bold
    frmTelefonoVerFra.Where2 = cboFichero(0).Text & "|" & lwTF.SelectedItem.Tag & "|"
    frmTelefonoVerFra.Show vbModal
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    ConseguirFoco Text2(Index), 3
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim T As String
Dim BorrarCampo As Boolean
    Text2(Index).Text = Trim(Text2(Index).Text)
    BorrarCampo = False
    If Text2(Index).Text <> "" Then
        T = Text2(Index).Text
        If EsFechaOK(T) Then
            If Index = 0 Then
                If CDate(T) < vEmpresa.FechaIni Or CDate(T) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then
                    MsgBox "Fechas fuera de ejercicio", vbExclamation
                    BorrarCampo = True
                End If
            End If
            Text2(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & Text2(Index).Text, vbExclamation
            BorrarCampo = True
        End If
    End If
    If BorrarCampo Then
        Text2(Index).Text = ""
        PonerFoco Text2(Index)
    End If
End Sub




Private Function GenerarImportacionCatadau(ByVal DigitoCoarval As Integer) As Boolean
Dim Campos() As String
Dim NF As Integer
Dim fin As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo eGenerarImportacionCatadau
    
    GenerarImportacionCatadau = False
    
    'tmpinformes codigo1 campo1 campo2  nombre1 fecha1 importe1 importe2 importe3
    conn.Execute "Delete from tmpinformes  WHERE codusu = " & vUsu.Codigo
    
    
    NF = FreeFile
    'Open Me.txtCatadau(0).Text For Input As #NF
    Open Text1.Text For Input As #NF
    'En toeria tiene que haber datos
    Cad = ""
    NumRegElim = -1 'Indciador de situacion de proceso
    Do
        
        Line Input #NF, Cad
        Cad = Trim(Cad)

        If Len(Cad) >= 5 Then
            If Mid(Cad, 1, 5) = ";;;;;" Then Cad = ""
        End If

        If Cad <> "" Then
            If NumRegElim >= 0 Then 'la primera NO vale y en la primera es -1
                If Mid(Cad, 1, 1) = ";" Then MsgBox "Empieza con ;", vbCritical
                Campos = Split(Cad, ";")
                NumRegElim = 0  'De momento BIEN
                Cad = Cad & vbCrLf & vbCrLf
                'Comprobaciones
                If UBound(Campos) < 11 Then
                    Cad = Cad & "Numero columnas incorrecto. Debian haber 11 columnas"
                    NumRegElim = 1
                Else
                    'OK. Columnas correctas
                    If Not IsNumeric(Campos(5)) Then
                        Cad = Cad & "Codigo socio incorrecto " & vbCrLf
                        NumRegElim = 1
                    Else
                        'Es el codigo de socio.
                        'Pos si acaso sa ha vuelto loco el de las factuas y lo envia "decimal"
                        Campos(5) = Replace(Campos(5), ",00", "")
                        Campos(5) = Replace(Campos(5), ".", "")
                    End If
                    Campos(0) = Campos(6)
                    If Len(Campos(0)) < 6 Then
                        Cad = Cad & "Longitud fra incorrecta " & vbCrLf
                        NumRegElim = 1
                    Else
                        If Not IsNumeric(Right(Campos(0), 6)) Then
                            Cad = Cad & "Numero fra incorrecta " & vbCrLf
                            NumRegElim = 1
                        End If
                    End If
                    
                    For i = 9 To 11
                        If Not IsNumeric(Campos(i)) Then
                            Cad = Cad & "Importes incorrectos " & Campos(i) & vbCrLf
                            NumRegElim = 1
                        End If
                    Next i
                    'Fecha factura
                    If Not IsDate(Campos(8)) Then
                        Cad = Cad & "Fecha incorrecta " & vbCrLf
                        NumRegElim = 1
                    End If
                    
                    
                    'SI llega a aqui, y ha ido bien, INSERTARA en tmp
                    If NumRegElim = 0 Then
                        '
                        
                        'Cad = Me.txtCatadau(1).Text & Right(Campos(6), 6)
                        Cad = DigitoCoarval & Right(Campos(6), 6)
                        
                        'codusu codigo1, campo1, campo2,  nombre1,
                        Cad = vUsu.Codigo & "," & Cad & "," & Campos(5) & "," & DBSet(Campos(4), "T")
                        ' fecha1, importe1, importe2, importe3
                        Cad = Cad & "," & DBSet(Campos(8), "F") & "," & DBSet(Campos(9), "N")
                        Cad = Cad & "," & DBSet(Campos(10), "N") & "," & DBSet(Campos(11), "N") & ")"
                        
                        Cad = "INSERT INTO tmpinformes(codusu,codigo1, campo1,   nombre1, fecha1, importe1, importe2, importe3) VALUES (" & Cad
                        conn.Execute Cad
                    Else
                        MsgBox Cad, vbExclamation
                        fin = True
                    End If
                End If 'numcols
            Else
                NumRegElim = 0  'para que empieze
                Cad = ""
            End If
        
            
        Else
            Cad = "OK"
        End If
        If NumRegElim = 1 Then
            MsgBox Cad, vbExclamation
            fin = True
        End If
        If EOF(NF) Then fin = True
    Loop Until fin
    Close #NF
    
    If Cad = "" Then
        MsgBox "NUmero registros incorrecto", vbExclamation
        NumRegElim = 1
    End If
    
    'OK. Ahora un par de comprobaciones mas
    If NumRegElim > 0 Then Exit Function
    
    
    
    
    'De momento va bien. Varias comporbaciones
    'Primer asunto. Codclien=0 NO lo procesamos. Serían internas
    Cad = "DELETE from tmpinformes where codusu=" & vUsu.Codigo & " AND campo1=0"
    conn.Execute Cad
    
    
    '1ª comprobacion
    Cad = "Select distinct(fecha1) from tmpinformes where codusu =" & vUsu.Codigo
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        If i = 0 Then
            Cad = miRsAux!fecha1
            Campos(0) = "01/01/" & Year(CDate(Cad))
            Campos(1) = "31/12/" & Year(CDate(Cad))
        End If
        i = i + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If i <> 1 Then
        MsgBox "No es fecha única", vbExclamation
        Exit Function
    End If
    
    Set cControlFra = New CControlFacturaContab
    
    Cad = cControlFra.FechaCorrectaContabilizazion(ConnConta, CDate(Cad))
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        NumRegElim = 1
    End If
    Set cControlFra = Nothing
    
    
    
    
    
    If NumRegElim = 1 Then Exit Function
    
    
    
    
    'UN par de comprobaciones mas
    Cad = "select codusu,max(codigo1) elmaximo ,min(codigo1) elminimo from tmpinformes where codusu=" & vUsu.Codigo & " group by 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    'Comprobemos si NO existe el numero factura
    Cad = " numfactu >= " & miRsAux!elminimo & " AND numfactu<=" & miRsAux!elmaximo
    Cad = Cad & " AND fecfactu>=" & DBSet(Campos(0), "F") & " AND fecfactu<=" & DBSet(Campos(1), "F")
    miRsAux.Close
    
    
    Cad = "Select count(*) FROM scafac where codtipom='FAT' AND " & Cad
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "0"
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then Cad = miRsAux.Fields(0)
    End If
    If Val(Cad) > 0 Then
        Cad = "Se van a solapar " & Cad & " registro(s) que se solaparán numeros de factura"
        Cad = Cad & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Function
        
    End If
        
    
    
    GenerarImportacionCatadau = True
    
    
    
    Exit Function
eGenerarImportacionCatadau:
    MuestraError Err.Number, Err.Description
    
End Function

Private Sub txtCatadau_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub



Private Sub InsertarFacturasTelefonoCoarval()
    CadenaDesdeOtroForm = ""
    frmListado3.Opcion = 36
    frmListado3.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        traspasofacturasTelefoniaCOARVAL Me.lblInf, CInt(CadenaDesdeOtroForm)
    End If
End Sub













Private Sub CargaComboCompanyia()

'    cboCompanyia2.Clear
'    cboCompanyia2.AddItem "Movistar"
'    cboCompanyia2.ItemData(cboCompanyia2.NewIndex) = 1
'
'    cboCompanyia2.AddItem "Orange"
'    cboCompanyia2.ItemData(cboCompanyia2.NewIndex) = 2
'

    CargarCombo_Tabla cboCompanyia2, "stfnoOperador", "codoperador", "nombre"
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        cboCompanyia2.ListIndex = 3
    ElseIf vParamAplic.NumeroInstalacion <> vbAlzira Then
        cboCompanyia2.ListIndex = 1
    Else
        cboCompanyia2.ListIndex = 0
    End If
    
    
End Sub



Private Function GeneraDatosImpresion() As Boolean
        
        
    On Error GoTo eGeneraDatosImpresion
    GeneraDatosImpresion = False
    Label1(5).Caption = "Preparando datos"
    Label1(5).Refresh
    conn.Execute "DELETE FROM tmpinformes where  codusu =" & vUsu.Codigo
    
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To Me.lwTF.ListItems.Count
    
        'tmpinformes( codusu,codigo1,campo1,nombre1,campo2,obser,importe1,importe2,importe3)
    
        Cad = ", (" & vUsu.Codigo & "," & NumRegElim & "," & NumRegElim
        Cad = Cad & "," & DBSet(lwTF.ListItems(NumRegElim).SubItems(1), "T")
        davidCodtipom = lwTF.ListItems(NumRegElim).SubItems(6)
        If InStr(1, davidCodtipom, "|") > 0 Then
            ImporteAuxiliar = 0
            i = 0
            Do
                i = InStr(i + 1, davidCodtipom, "|")
                If i > 0 Then ImporteAuxiliar = ImporteAuxiliar + 1
            Loop Until i = 0
            Cad = Cad & ",1"
            davidCodtipom = Replace(davidCodtipom, "|", vbCrLf)
            davidCodtipom = Left(davidCodtipom, Len(davidCodtipom) - 1)  'quitamos el ulimio pipe
        Else
            Cad = Cad & ",0"
            ImporteAuxiliar = 1
        End If
        Cad = Cad & "," & DBSet(davidCodtipom, "T")
        'Bases de fichero      Bases albaranes     Bases vtaplazos
        Cad = Cad & "," & DBSet(lwTF.ListItems(NumRegElim).SubItems(10), "N")
        Cad = Cad & "," & DBSet(lwTF.ListItems(NumRegElim).SubItems(7), "N")
        Cad = Cad & "," & DBSet(lwTF.ListItems(NumRegElim).SubItems(8), "N") & "," & CInt(ImporteAuxiliar) & ")"
        
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad
        If Len(CadenaDesdeOtroForm) > 5000 Then
            
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
            Cad = "INSERT INTO tmpinformes( codusu,codigo1,campo1,nombre1,campo2,obser,importe1,importe2,importe3,importe5) VALUES " & CadenaDesdeOtroForm
            conn.Execute Cad
            CadenaDesdeOtroForm = ""
        End If
    Next
    
    If Len(CadenaDesdeOtroForm) > 0 Then
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        Cad = "INSERT INTO tmpinformes( codusu,codigo1,campo1,nombre1,campo2,obser,importe1,importe2,importe3,importe5) VALUES " & CadenaDesdeOtroForm
        conn.Execute Cad
    End If
    
    Label1(5).Caption = "Calculando "
    Label1(5).Refresh
    Espera 0.5
    
    Cad = "UPDATE tmpinformes SET importe4=importe1+importe2+importe3 WHERE codusu =" & vUsu.Codigo
    conn.Execute Cad
    
    GeneraDatosImpresion = True
eGeneraDatosImpresion:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
    davidCodtipom = ""
    davidNumalbar = 0
    
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Nuevo
        Case 2  'Modificar
        Case 3 'Eliminar
            cmdEliminarDatosFracion_Click
        Case 5 'Busqueda
            cmdBus_Click
        Case 6 'Ver Todos
        Case 8 'Imprimir
            cmdImprimir_Click
    End Select
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            'facturar
            cmdFacturar_Click
    End Select
End Sub
