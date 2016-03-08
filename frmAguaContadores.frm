VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAguaContadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contadores "
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12900
   Icon            =   "frmAguaContadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Tag             =   "Tipo de uso|N|N|||aguacontadores|TipoFacturacion|||"
      Top             =   2880
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   1200
      TabIndex        =   79
      Top             =   6600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Facturar"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Envio factura por email"
      Height          =   255
      Index           =   1
      Left            =   10440
      TabIndex        =   14
      Tag             =   "Calibre|N|N|||aguacontadores|EnvioFacPorEmail|||"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Tipo de uso|N|N|||aguacontadores|TipoUso|||"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   27
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Fec. anterior|F|S|||aguacontadores|fechabaja|||"
      Text            =   "Text1"
      Top             =   2850
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   26
      Left            =   240
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Fec. alta|F|S|||aguacontadores|fechalta|||"
      Text            =   "Text1"
      Top             =   2850
      Width           =   1125
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   7920
      TabIndex        =   73
      Top             =   6600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   300
      MaxLength       =   6
      TabIndex        =   17
      Tag             =   "C. Postal|T|N|||aguacontadores|cpconta||N|"
      Text            =   "Text1"
      Top             =   5360
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   18
      Tag             =   "Población|T|N|||aguacontadores|pobconta||N|"
      Text            =   "Text1"
      Top             =   5360
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   24
      Tag             =   "Provincia|T|S|||aguacontadores|proenvio||N|"
      Text            =   "Text1"
      Top             =   6015
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   4680
      MaxLength       =   35
      TabIndex        =   21
      Tag             =   "Domicilio|T|S|||aguacontadores|domenvio||N|"
      Text            =   "Text1"
      Top             =   4705
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   22
      Left            =   4680
      MaxLength       =   35
      TabIndex        =   20
      Tag             =   "Nombre|T|S|||aguacontadores|nomenvio||N|"
      Text            =   "Text1"
      Top             =   4050
      Width           =   4125
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Facturar"
      Height          =   255
      Index           =   0
      Left            =   9000
      TabIndex        =   13
      Tag             =   "Calibre|N|N|||aguacontadores|facturar|||"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   2160
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Text            =   "Combo2"
      Top             =   720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   9120
      MaxLength       =   40
      TabIndex        =   26
      Tag             =   "Provincia|T|S|||aguacontadores|TitularBanco||N|"
      Text            =   "Text1"
      Top             =   4800
      Width           =   3645
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   9840
      MaxLength       =   4
      TabIndex        =   28
      Tag             =   "Banco|N|S|0|9999|aguacontadores|codbanco|0000||"
      Text            =   "Text1"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   10440
      MaxLength       =   4
      TabIndex        =   29
      Tag             =   "Sucursal|N|S|0|9999|aguacontadores|codsucur|0000||"
      Text            =   "Text1"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   11040
      MaxLength       =   2
      TabIndex        =   30
      Tag             =   "Digito Control|T|S|||aguacontadores|digcontr|00||"
      Text            =   "Text1"
      Top             =   5520
      Width           =   360
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   20
      Left            =   11520
      MaxLength       =   10
      TabIndex        =   31
      Tag             =   "Cuenta Bancaria|T|S|||aguacontadores|cuentaba|0000000000||"
      Text            =   "9999999999"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   9120
      MaxLength       =   4
      TabIndex        =   27
      Tag             =   "IBAN|T|S|||aguacontadores|iban|||"
      Text            =   "Text1"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   15
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Consumo|N|S|0||aguacontadores|consumo|0||"
      Text            =   "Text1"
      Top             =   1930
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   14
      Left            =   9120
      MaxLength       =   5
      TabIndex        =   25
      Tag             =   "Forpa|N|N|0||aguacontadores|codforpa|0||"
      Text            =   "Text1"
      Top             =   4080
      Width           =   645
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   14
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "Text2"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "lectura anterior|N|S|0||aguacontadores|lec_actual|0||"
      Text            =   "Text1"
      Top             =   1930
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Fec. anterior|F|S|||aguacontadores|fecha_actual||N|"
      Text            =   "Text1"
      Top             =   1930
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   22
      Tag             =   "C. Postal|T|S|||aguacontadores|cpenvio||N|"
      Text            =   "Text1"
      Top             =   5360
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   5640
      MaxLength       =   30
      TabIndex        =   23
      Tag             =   "Población|T|S|||aguacontadores|pobenvio||N|"
      Text            =   "Text1"
      Top             =   5360
      Width           =   3165
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   7
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Text2"
      Top             =   4050
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   300
      TabIndex        =   15
      Tag             =   "Cliente|N|N|||aguacontadores|codclien|||"
      Text            =   "Text1"
      Top             =   4050
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Calibre|N|N|||aguacontadores|codcalibre||N|"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   240
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fec. anterior|F|S|||aguacontadores|fecha_anterior|||"
      Text            =   "Text1"
      Top             =   1930
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   300
      MaxLength       =   30
      TabIndex        =   19
      Tag             =   "Provincia|T|N|||aguacontadores|proconta||N|"
      Text            =   "Text1"
      Top             =   6015
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Numero contador|T|N|||aguacontadores|contador||S|"
      Text            =   "Text"
      Top             =   720
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   300
      MaxLength       =   35
      TabIndex        =   16
      Tag             =   "Domicilio|T|N|||aguacontadores|domconta||N|"
      Text            =   "Text1"
      Top             =   4705
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "lectura anterior|N|S|0||aguacontadores|lec_anterior|0||"
      Text            =   "Text1"
      Top             =   1930
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   420
      Left            =   120
      TabIndex        =   35
      Top             =   9240
      Width           =   2535
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11520
      TabIndex        =   34
      Top             =   9240
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10200
      TabIndex        =   32
      Top             =   9240
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   450
      Left            =   6960
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      TabIndex        =   39
      Top             =   0
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas conceptos facturacion"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   9480
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   9240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   23
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   67
      Tag             =   "1|T|S|||aguacontadores|tarbq1|||"
      Text            =   "Text1"
      Top             =   720
      Width           =   6405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   24
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   68
      Tag             =   "e|T|S|||aguacontadores|tarbq2|||"
      Text            =   "Text1"
      Top             =   1440
      Width           =   6405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   25
      Left            =   6240
      MaxLength       =   35
      TabIndex        =   69
      Tag             =   "3|T|S|||aguacontadores|tarbq3|||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   6405
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   810
      Left            =   6960
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1429
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
   Begin VB.Label Label1 
      Caption         =   "Tipo facturación"
      Height          =   255
      Index           =   34
      Left            =   6240
      TabIndex        =   81
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Conceptos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   33
      Left            =   240
      TabIndex        =   78
      Top             =   6600
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de uso"
      Height          =   255
      Index           =   32
      Left            =   3240
      TabIndex        =   77
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "F.Baja"
      Height          =   195
      Index           =   31
      Left            =   1680
      TabIndex        =   76
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   27
      Left            =   2400
      Picture         =   "frmAguaContadores.frx":000C
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "F.Alta"
      Height          =   195
      Index           =   30
      Left            =   240
      TabIndex        =   75
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   26
      Left            =   960
      Picture         =   "frmAguaContadores.frx":010E
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Datos pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   29
      Left            =   9120
      TabIndex        =   74
      Top             =   3480
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1215
      Left            =   120
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   7080
      TabIndex        =   72
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   71
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "m3"
      Height          =   255
      Index           =   28
      Left            =   5280
      TabIndex        =   70
      Top             =   1650
      Width           =   375
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   12
      Left            =   3120
      Picture         =   "frmAguaContadores.frx":0210
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1650
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   840
      Picture         =   "frmAguaContadores.frx":0312
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Tarifa BQ3"
      Height          =   255
      Index           =   27
      Left            =   6240
      TabIndex        =   66
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tarifa BQ2"
      Height          =   255
      Index           =   26
      Left            =   6240
      TabIndex        =   65
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   12720
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Tarifa BQ1"
      Height          =   255
      Index           =   25
      Left            =   6240
      TabIndex        =   64
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   63
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Titular cuenta banco"
      Height          =   255
      Index           =   23
      Left            =   9120
      TabIndex        =   62
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IBAN"
      Height          =   195
      Index           =   22
      Left            =   9120
      TabIndex        =   61
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   20
      Left            =   4920
      TabIndex        =   60
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   19
      Left            =   2520
      TabIndex        =   59
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   58
      Top             =   1320
      Width           =   675
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   12720
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   7
      Left            =   960
      Picture         =   "frmAguaContadores.frx":0414
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   14
      Left            =   10320
      Picture         =   "frmAguaContadores.frx":0516
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Forma de pago"
      Height          =   195
      Index           =   17
      Left            =   9120
      TabIndex        =   57
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   56
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lectura  m3"
      Height          =   195
      Index           =   15
      Left            =   3600
      TabIndex        =   54
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   53
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   13
      Left            =   4680
      TabIndex        =   52
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   51
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   50
      Top             =   5820
      Width           =   975
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   5325
      Picture         =   "frmAguaContadores.frx":0618
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   5160
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Direccion envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   48
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Dirección suministro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   47
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Calibre"
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   46
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   945
      Picture         =   "frmAguaContadores.frx":071A
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   5130
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   45
      Top             =   1650
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   44
      Top             =   5820
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   4
      Left            =   1260
      TabIndex        =   43
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   42
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   41
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Lectura   m3"
      Height          =   255
      Index           =   1
      Left            =   1260
      TabIndex        =   38
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   37
      Top             =   480
      Width           =   975
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Lineas conecptos facturas"
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
Attribute VB_Name = "frmAguaContadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes3
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin ningun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'Para no hacer miles de consultas, cargo los importes
Private rsPrecios As ADODB.Recordset


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    InsertarEnConceptosContador
                    PosicionarData
                    If Modo = 2 Then CargaConceptosContador
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        Case 1  'BUSCAR
            HacerBusqueda
            
        Case 5
            
            
            For NumRegElim = 1 To ListView1.ListItems.Count
                If Abs(ListView1.ListItems(NumRegElim).Checked) <> IIf(ListView1.ListItems(NumRegElim).SubItems(2) = "", 0, 1) Then
                    'Hay que updatrar
                    CadenaConsulta = "UPDATE aguacontadoresconce SET facturar =" & Abs(ListView1.ListItems(NumRegElim).Checked)
                    If ListView1.ListItems(NumRegElim).Key = "K7" Then
                        If ListView1.ListItems(NumRegElim).Checked Then
                            CadenaConsulta = CadenaConsulta & ", descripcion = " & DBSet(ListView1.ListItems(NumRegElim).Text, "T")
                            CadenaConsulta = CadenaConsulta & ", importeconcepto =" & DBSet(ListView1.ListItems(NumRegElim).SubItems(1), "N")
                        End If
                    End If
                    CadenaConsulta = CadenaConsulta & " WHERE  contador =" & DBSet(Text1(0).Text, "T")
                    CadenaConsulta = CadenaConsulta & " AND codconceAg = " & Mid(ListView1.ListItems(NumRegElim).Key, 2)
                    If Ejecutar(CadenaConsulta, False) Then ListView1.ListItems(NumRegElim).SubItems(2) = IIf(Abs(ListView1.ListItems(NumRegElim).Checked) = 1, "Si", "")
                    
                End If
            Next
            CadenaConsulta = ""
            PonerModo 2
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3 'Insertar
        LimpiarCampos2
        PonerModo 0
        PonerOpcionesMenu
    Case 4, 5 'Modificar
        lblIndicador.Caption = ""
        TerminaBloquear
        PonerModo 2
        PonerCampos
    
    End Select
End Sub


Private Sub BotonAnyadir()

    
    LimpiarCampos2
    PonerModo 3
        
    

    PonerFoco Text1(0)
    Text1_GotFocus 0
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos2
        PonerModo 1
        
        
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
    'Ver todos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos2
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    
    Me.Combo2(0).Text = Me.Text1(23).Text
    Me.Combo2(1).Text = Me.Text1(24).Text
    Me.Combo2(2).Text = Me.Text1(25).Text
    
    PonerModo 4
    PonerFoco Text1(22)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    

    
    
    'Copmpruebo si esta vinculado a algun trabajador
    Cad = DevuelveDesdeBD(conAri, "count(*)", "aguahcolecturas", "contador", CStr(Data1.Recordset!Contador), "T")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Historico lectura. No se puede eliminar", vbExclamation
        Exit Sub
    End If
    Cad = DevuelveDesdeBD(conAri, "count(*)", "scafac1", "codtipom='FAG' AND referenc", CStr(Data1.Recordset!Contador), "T")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "El contador tiene facturas", vbExclamation
        Exit Sub
    End If
    
    
    
'


        Cad = "¿Seguro que desea eliminar el contador? " & vbCrLf
        Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), "0000")
        Cad = Cad & vbCrLf & "Cliente: " & Data1.Recordset!codclien & " " & Me.Text2(7).Text
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub


        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        
        Cad = "Delete from aguacontadoresconce where contador=" & DBSet(Data1.Recordset!Contador, "T")
        conn.Execute Cad
        
        Cad = "Delete from aguacontadores where contador=" & DBSet(Data1.Recordset!Contador, "T")
        conn.Execute Cad
        
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos2
            PonerModo 0
        End If

    Screen.MousePointer = vbDefault
    
Error2:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar el contador", Err.Description
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Me.Combo2(0).ListCount > 0 Then Exit Sub
    
    CargaGrid False
    
    lblIndicador.Caption = "Leyendo precios"
    lblIndicador.Refresh
    Set rsPrecios = New ADODB.Recordset
    CadenaConsulta = "Select * from aguacalibre"
    rsPrecios.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    lblIndicador.Caption = "Leyendo tarifas"
    lblIndicador.Refresh
    
    Set miRsAux = New ADODB.Recordset
    For kCampo = 0 To 2
        Combo2(kCampo).Clear
        lblIndicador.Caption = "Leyendo " & kCampo + 1 & " / 3"
        lblIndicador.Refresh
        CadenaConsulta = "tarbq" & kCampo + 1
        CadenaConsulta = "Select " & CadenaConsulta & " from aguacontadores  WHERE " & CadenaConsulta & " <> '' GROUP BY 1"
        miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Combo2(kCampo).AddItem miRsAux.Fields(0)
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next
    Set miRsAux = Nothing
    
    
    
    lblIndicador.Caption = ""
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 16
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas
        .Buttons(12).Image = 16  'Imprimir
        .Buttons(13).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    LimpiarCampos2
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "aguacontadores"
    Ordenacion = " ORDER BY contador"
        
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where contador='@D@'"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    
    
    'Los combos
    Set miRsAux = New ADODB.Recordset
    Me.Combo1(0).Clear
    miRsAux.Open "Select * from aguacalibre order by nomcalibre", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CadenaConsulta = miRsAux!nomcalibre & " (" & miRsAux!calibre
        If DBLet(miRsAux!Caudal, "N") > 0 Then CadenaConsulta = CadenaConsulta & " / " & miRsAux!Caudal
        Combo1(0).AddItem CadenaConsulta & ")"
        Combo1(0).ItemData(Combo1(0).NewIndex) = miRsAux!codcalibre
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Combo1(1).Clear
    Combo1(1).AddItem "Doméstico"
    Combo1(1).ItemData(0) = 0
    Combo1(1).AddItem "Industrial"
    Combo1(1).ItemData(1) = 1
    
    
    Combo1(2).Clear
    Combo1(2).AddItem "Contador"
    Combo1(2).ItemData(0) = 0
    Combo1(2).AddItem "Aforo"
    Combo1(2).ItemData(1) = 1
    Combo1(2).AddItem "Anexo"
    Combo1(2).ItemData(2) = 2
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    'codcalibre nomcalibre calibre caudal
End Sub


Private Sub LimpiarCampos2()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    Combo2(0).Text = ""
    Combo2(1).Text = ""
    Combo2(2).Text = ""
    Me.Check1(0).Value = 0: Me.Check1(1).Value = 0
    
    
    Me.ListView1.ListItems.Clear
    CargaGrid False
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Not rsPrecios Is Nothing Then rsPrecios.Close
    Set rsPrecios = Nothing
    
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub
    
Private Sub frmC_Selec(vFecha As Date)
    CadenaConsulta = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(indice + 2).Text = devuelve
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass


    If Index < 2 Then
        kCampo = 3
        If Index = 2 Then kCampo = 10
        'Codigo Postal
        Set frmCP = New frmCPostal
        frmCP.DatosADevolverBusqueda = "0"
        frmCP.Show vbModal
        Set frmCP = Nothing

        
        PonerFoco Text1(kCampo)
        VieneDeBuscar = True
        
        
    Else
        CadenaConsulta = ""
        Select Case Index
        Case 7
            
            Set frmCli = New frmFacClientes3
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            
            If CadenaConsulta <> "" Then
                Me.Text1(7).Text = RecuperaValor(CadenaConsulta, 1)
                Text1_LostFocus 7
            End If
        Case 14
             'FORPA
            
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
        
            If CadenaConsulta <> "" Then
                Text1(Index).Text = RecuperaValor(CadenaConsulta, 1)
                Text2(Index).Text = RecuperaValor(CadenaConsulta, 2)
            End If
            
        Case 6, 12, 26, 27
            Set frmC = New frmCal
            If Me.Text1(Index).Text <> "" Then
                frmC.Fecha = CDate(Text1(Index).Text)
            Else
                frmC.Fecha = Now
            End If
            frmC.Show vbModal
            Set frmC = Nothing
            If CadenaConsulta <> "" Then Text1(Index).Text = CadenaConsulta
        End Select
        CadenaConsulta = ""
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_DblClick()
    If Modo <> 5 Then Exit Sub
    
    'Solo la cuota de varios se puede modificar
    If ListView1.SelectedItem.Key <> "K7" Then Exit Sub
    
    CadenaDesdeOtroForm = ""
    frmListado3.opcion = 54
    frmListado3.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        Me.ListView1.SelectedItem.Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.ListView1.SelectedItem.SubItems(1) = RecuperaValor(CadenaDesdeOtroForm, 2)
        Me.ListView1.SelectedItem.SubItems(2) = "" 'Para que haga el UPDATE
        ListView1.SelectedItem.Checked = True
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnLineas_Click()
    BotonLineas
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
Dim devuelve As String
Dim vCli As CCliente

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0
                'comprobar si ya existe ese codigo de Agente en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)

                End If


        Case 1, 13
            If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
            If Text1(13).Text <> "" And Text1(1).Text <> "" Then
                If Me.Text1(3).Text <> "" Then Text1(15).Text = Val(Me.Text1(13).Text) - Val(Text1(1).Text)
            End If
        Case 6, 12, 26, 27
            
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 3, 10 'CPostal
            If Index = 3 Then
                If Text1(Index).Text = "" Then
                    Text1(Index + 1).Text = ""
                    Text1(Index + 2).Text = ""
                ElseIf Not VieneDeBuscar Then
                     Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                     Text1(Index + 2).Text = devuelve
                End If
            Else
                'Direccion envio factura
                If Text1(Index).Text = "" Then
                    Text1(Index - 1).Text = ""
                    Text1(Index - 2).Text = ""
                ElseIf Not VieneDeBuscar Then
                     Text1(Index - 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                     Text1(Index - 2).Text = devuelve
                End If
            End If
            VieneDeBuscar = False
            
            
        Case 7
            'Cliente
            devuelve = ""
            If PonerFormatoEntero(Text1(Index)) Then
                devuelve = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text)
                If devuelve <> "" Then
                    If Modo = 3 Then
                        Set vCli = New CCliente
                        vCli.LeerDatos Text1(Index).Text
                        If vCli.ClienteBloqueado Then
                            devuelve = ""
                        Else
                            Me.Text1(2).Text = vCli.Domicilio
                            Me.Text1(3).Text = vCli.CPostal
                            Me.Text1(4).Text = vCli.Poblacion
                            Me.Text1(5).Text = vCli.Provincia
                            Me.Text1(21).Text = vCli.Nombre
                            Me.Text1(16).Text = vCli.IBAN
                            Me.Text1(17).Text = Right("0000" & vCli.Banco, 4)
                            Me.Text1(18).Text = Right("0000" & vCli.Sucursal, 4)
                            Me.Text1(19).Text = vCli.DigControl
                            Me.Text1(20).Text = vCli.CuentaBan
                            Me.Text1(14).Text = vCli.ForPago
                            Me.Text2(14).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(14).Text)
                            
                        End If
                        Set vCli = Nothing
                    End If
                Else
                    MsgBox "No existe el cliente: " & Me.Text1(Index).Text
                End If
                
                If devuelve = "" Then
                    
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
                    
                                                
                    
                
            End If
            Me.Text2(Index).Text = devuelve
            
        Case 14
            'Forma de pago
            devuelve = ""
            If PonerFormatoEntero(Text1(Index)) Then
                devuelve = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(Index).Text)
                If devuelve = "" Then
                    MsgBox "No existe la forma de pago: " & Me.Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
                
            End If
            Me.Text2(Index).Text = devuelve
            
            
        Case 17, 18
            Text1(Index).Text = Format(Text1(Index).Text, "0000")
            
        Case 20
            'CuentaBan
            
            'Si hay valor en la cuenta le calculamos el IBAN
            If Me.Text1(Index).Text <> "" Then
                Me.Text1(Index).Text = Right(String(10, "0") & Text1(Index).Text, 10)
                devuelve = Text1(17).Text & Me.Text1(18).Text & Me.Text1(19).Text & Me.Text1(20).Text
            
                If Len(devuelve) = 20 Then
                    DevuelveIBAN2 "ES", devuelve, CadenaConsulta
                    If Len(CadenaConsulta) = 2 Then
                        CadenaConsulta = "ES" & CadenaConsulta
                        If Me.Text1(16).Text = "" Then
                            Text1(16).Text = CadenaConsulta
                        Else
                            If Me.Text1(16).Text <> CadenaConsulta Then MsgBox "Codigo IBAN distinto del calculado [" & CadenaConsulta & "]", vbExclamation
                        End If
                    End If
                End If
                CadenaConsulta = ""
            End If

        
    
            

    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then     'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 30, "Código")
        Cad = Cad & ParaGrid(Text1(7), 15, "Cod. Cliente")
        Cad = Cad & "Nombre|sclien|nomclien|T||45·"

        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = " aguacontadores left join sclien on aguacontadores.codclien=sclien.codclien"
            frmB.vSQL = cadB
            HaDevueltoDatos = False
            '###A mano

            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Contadores de agua"
            frmB.vselElem = 0
            frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                PonerFocoBtn Me.cmdRegresar
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerModo Modo
                PonerFoco Text1(kCampo)
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
         'PonerModo 0
         Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
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
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1

    Text2(7).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(7).Text)
    Text2(14).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(14).Text)

    CargaConceptosContador

    CargaGrid True
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

    Modo = Kmodo

    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    
    
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    
    
    
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    B = Modo = 1 Or Modo = 3 Or Modo = 4
    
    Me.Check1(0).Enabled = B: Me.Check1(1).Enabled = B
    BloquearCmb Combo1(0), Not B
    BloquearCmb Combo1(1), Not B
    BloquearCmb Combo1(2), Not B
    
    
    '---------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    B = B Or Modo = 1
    Combo2(0).visible = B
    Combo2(1).visible = B
    Combo2(2).visible = B
    
    
    'Modos: 1-3-4-5
    B = B Or Modo = 5
    cmdAceptar.visible = B
    cmdCancelar.visible = B
    
    
    
    B = Modo = 5
    ListView1.Checkboxes = B
    If Modo = 5 Then
        ListView1.ColumnHeaders(3).Width = 0
        For NumReg = 1 To ListView1.ListItems.Count
            ListView1.ListItems(NumReg).Checked = ListView1.ListItems(NumReg).SubItems(2) <> ""
        Next
       ' ListView1.Refresh
       ' Me.Refresh
       ' DoEvents
    Else
        ListView1.ColumnHeaders(3).Width = 900
    End If
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    
    B = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    Toolbar1.Buttons(10).Enabled = B
    Me.mnLineas.Enabled = B
    
    
    '----------------------------------------
    B = (Modo >= 3) 'Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    
    
    
        
    
    
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim BuscaChekc As String

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar datos OK
    If Not B Then Exit Function
        
    BuscaChekc = ""
    
    If Me.Text1(6).Text <> "" Xor Text1(1).Text <> "" Then
        BuscaChekc = BuscaChekc & "Falta datos  anterior " & vbCrLf
        kCampo = 6
    End If
    If Me.Text1(12).Text <> "" Xor Text1(13).Text <> "" Then
        BuscaChekc = BuscaChekc & "Falta datos  actual " & vbCrLf
        kCampo = 12
    End If
    
    If Me.Text1(6).Text <> "" And Text1(12).Text <> "" Then
        If CDate(Text1(12).Text) < CDate(Text1(6).Text) Then
            BuscaChekc = BuscaChekc & "Fecha anterior menor que actual" & vbCrLf
            kCampo = 6
        End If
    End If
    If Me.Text1(1).Text <> "" And Text1(13).Text <> "" Then
        If Val(Text1(13).Text) < Val(Text1(1).Text) Then
            BuscaChekc = BuscaChekc & "Lectura anterior menor que actual" & vbCrLf
            kCampo = 13
        End If
    End If
    
    
    
    If BuscaChekc <> "" Then
        MsgBox BuscaChekc, vbExclamation
        PonerFoco Text1(kCampo)
        Exit Function
    End If
        
        
    '- Validar que la cuenta bancaria es correcta
    If Comprueba_CuentaBan2(Text1(17).Text & Text1(18).Text & Text1(19).Text & Text1(20).Text, False) Then
        CadenaConsulta = Text1(17).Text & Text1(18).Text & Text1(19).Text & Text1(20).Text
        If Len(CadenaConsulta) = 20 Then
            
            BuscaChekc = ""
            If Me.Text1(16).Text <> "" Then BuscaChekc = Mid(Text1(16).Text, 1, 2)
            
                
            If DevuelveIBAN2(BuscaChekc, CadenaConsulta, CadenaConsulta) Then
                If Me.Text1(16).Text = "" Then
                    Me.Text1(16).Text = BuscaChekc & CadenaConsulta
                Else
                    If Mid(Text1(16).Text, 3) <> CadenaConsulta Then
                        CadenaConsulta = "Calculado : " & BuscaChekc & CadenaConsulta
                        CadenaConsulta = "Introducido: " & Me.Text1(16).Text & vbCrLf & CadenaConsulta & vbCrLf
                        CadenaConsulta = "Error en codigo IBAN" & vbCrLf & CadenaConsulta & "Continuar?"
                        If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then Exit Function
                    End If
                End If
            End If
                    
        End If
        CadenaConsulta = ""
        BuscaChekc = ""
    End If
        
        
            
    If Modo = 3 Or Modo = 4 Then
        Me.Text1(23).Text = Me.Combo2(0).Text
        Me.Text1(24).Text = Me.Combo2(1).Text
        Me.Text1(25).Text = Me.Combo2(2).Text
    End If
            
            
        
    DatosOk = B
End Function




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 10
            mnLineas_Click
        Case 12
            'Imprimir
            frmListado5.OpcionListado = 0
            frmListado5.Show vbModal
        Case 13  'Salir
            mnSalir_Click
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


Private Sub PosicionarData()
Dim Cad As String
Dim Indicador As String

    Cad = "(contador=" & DBSet(Text1(0).Text, "T") & ")"
    
    If Modo = 3 Then Data1.RecordSource = "select * from aguacontadores WHERE " & Cad
    
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos2
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
Dim SQL As String
On Error Resume Next
    SQL = " WHERE contador= " & DBSet(Text1(0).Text, "T")
    ObtenerWhereCP = SQL
End Function




Private Sub CargaGrid(enlaza As Boolean)
Dim C As String

    C = "select fecha_anterior,lec_anterior,fecha_actual,lec_actual,fecha_factura from aguahcolecturas where contador = "
    If enlaza Then
        C = C & DBSet(Data1.Recordset!Contador, "T")
    Else
        C = C & "'D@BYZ'"
    End If
    C = C & " order by fecha_factura desc"
    
    
    CargaGridGnral Me.DataGrid1, Adodc1, C, True

    
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    DataGrid1.Columns(0).Caption = "Inicio"
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(0).Alignment = dbgCenter
    
    DataGrid1.Columns(1).Caption = "Ante."
    DataGrid1.Columns(1).Width = 600
    DataGrid1.Columns(1).Alignment = dbgRight
    
    DataGrid1.Columns(2).Caption = "Fin"
    DataGrid1.Columns(2).Width = 1100
     DataGrid1.Columns(2).Alignment = dbgCenter
    
    DataGrid1.Columns(3).Caption = "Actual"
    DataGrid1.Columns(3).Width = 600
    DataGrid1.Columns(3).Alignment = dbgRight
    
    DataGrid1.Columns(4).Caption = "Facturacion"
    DataGrid1.Columns(4).Width = 1000
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).Alignment = dbgCenter

    
    
End Sub


Private Sub CargaConceptosContador()
Dim It As ListItem

    Me.ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "select * from aguacontadoresConce  where contador=" & DBSet(Data1.Recordset!Contador, "T") & " order by codconceag", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'codconceAg descripcion importeconcepto
        Set It = ListView1.ListItems.Add(, "K" & miRsAux!codconceAg)
        It.Text = miRsAux!Descripcion
        If DBLet(miRsAux!importeconcepto, "N") > 0 Then
            It.SubItems(1) = Format(miRsAux!importeconcepto, FormatoPrecio)
        Else
            It.SubItems(1) = " "
        End If
        It.SubItems(2) = IIf(miRsAux!facturar = 1, "Si", "")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub



Private Sub BotonLineas()
    PonerModo 5
    
End Sub


'Private Function DevuelvePrecio(Concepto As Integer) As String
'
'
'
'    DevuelvePrecio = ""
'
'    rsPrecios.Find "codcalibre =" & Me.Combo1(0).ItemData(Combo1(0).ListIndex), , adSearchForward
'    If Not rsPrecios.EOF Then
'        '"Doméstico"        0
'        '"Industrial"       1
'        If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Then
'            'domestico
'
'        Else
'            'industrial
'
'        End If
'
'
'
'    End If
'
'End Function


Private Sub InsertarEnConceptosContador()
    'Insertamos todos los conceptos facturables, menos el de varios
    CadenaConsulta = "INSERT INTO  aguacontadoresconce  "
    CadenaConsulta = CadenaConsulta & " select " & DBSet(Text1(0).Text, "T") & ",codconceAg "
    CadenaConsulta = CadenaConsulta & ",descconceAg, 0, if(codconceAg=7,0,1)  from aguaconceptos"
    If Not Ejecutar(CadenaConsulta, True) Then MsgBox "Error insertando conceptos. Avise soporte técnico", vbExclamation
    
End Sub
