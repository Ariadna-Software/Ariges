VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmMovimientosGr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Almacen"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   Icon            =   "frmAlmMovimientosGr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameToolAux0 
      Height          =   645
      Left            =   135
      TabIndex        =   36
      Top             =   2565
      Width           =   1500
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   37
         Top             =   180
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
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
      Height          =   300
      Left            =   13050
      TabIndex        =   35
      Top             =   315
      Width           =   1515
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   33
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   34
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3870
      TabIndex        =   31
      Top             =   135
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   32
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4995
      TabIndex        =   29
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   30
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   6360
      MaxLength       =   8
      TabIndex        =   27
      Tag             =   "Hora|H|N|||scamov|hormovim|hh:mm:ss|N|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1125
   End
   Begin VB.CheckBox chkImpresion 
      Caption         =   "Impreso"
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
      Height          =   255
      Left            =   7830
      TabIndex        =   26
      Tag             =   "Situaci�n Impresi�n|N|N|||scamov|situacio||N|"
      Top             =   1125
      Width           =   1260
   End
   Begin VB.ComboBox cboAux 
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
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Situaci�n Impresi�n|N|N|||scamov|situacio||N|"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   315
      Left            =   1200
      TabIndex        =   24
      ToolTipText     =   "Buscar art�culo"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
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
      Height          =   320
      Index           =   3
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   8
      Text            =   "observac"
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
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
      Height          =   320
      Index           =   2
      Left            =   3960
      MaxLength       =   16
      TabIndex        =   6
      Text            =   "cantidad"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   320
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   14
      Text            =   "nombre artic"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
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
      Height          =   290
      Index           =   0
      Left            =   240
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "codartic"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
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
      Left            =   12330
      TabIndex        =   9
      Top             =   8775
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
      Left            =   13590
      TabIndex        =   10
      Top             =   8775
      Width           =   1065
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
      Height          =   375
      Left            =   13590
      TabIndex        =   23
      Top             =   8775
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   8640
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   375
         TabIndex        =   22
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text2 
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
      Index           =   0
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1620
      Width           =   6555
   End
   Begin VB.TextBox Text2 
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
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   2025
      Width           =   6555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Index           =   4
      Left            =   9330
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "Observaciones|T|S|||scamov|observa1||N|"
      Text            =   "frmAlmMovimientosGr.frx":000C
      Top             =   1110
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Cod. Trabajador|N|N|0|9999|scamov|codtraba|0000|N|"
      Text            =   "Text1"
      Top             =   2025
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Cod. Almacen|N|N|0|999|scamov|codalmac|000|N|"
      Text            =   "Text1"
      Top             =   1620
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   4380
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||scamov|fecmovim|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmMovimientosGr.frx":0012
      Height          =   5280
      Left            =   120
      TabIndex        =   11
      Top             =   3300
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   9313
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   300
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
      BackColor       =   &H00FEF7E4&
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
      Left            =   1815
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "N� Movimiento|N|S|0||scamov|codmovim|0000000|S|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9720
      Top             =   300
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1530
      Picture         =   "frmAlmMovimientosGr.frx":0027
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   1665
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Hora"
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
      Left            =   5835
      TabIndex        =   28
      Top             =   1125
      Width           =   780
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1530
      ToolTipText     =   "Buscar trabajador"
      Top             =   2055
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   4065
      Picture         =   "frmAlmMovimientosGr.frx":0A29
      ToolTipText     =   "Buscar fecha"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
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
      Left            =   9330
      TabIndex        =   18
      Top             =   855
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Trabajador"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Almac�n"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
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
      Left            =   3285
      TabIndex        =   15
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "N� Movimiento"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   1500
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
Attribute VB_Name = "frmAlmMovimientosGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'historico schmov, y solo en modo de consulta

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del hist�rico de movimiento seleccionado (solo consulta)
Public hcoCodMovim As Long 'cod. movim del historico
Public hcoFechaMovim As Date 'Fecha del historico


'-----------------------------------------------------------------------

Private WithEvents frmB As frmBasico2 'frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'Mto de Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2 'AlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1


Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe


Private HaDevueltoDatos As Boolean



Private Sub cboAux_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboImpresion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
    Case 1 'BUSQUEDA
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk Then InsertarCabecera

    Case 4 'MODIFICAR
        If DatosOk Then
            If ModificaDesdeFormulario(Me, 1) Then
                TerminaBloquear
                Cad = "(" & ObtenerWhereCP(False) & ")"
                If SituarData(Data1, Cad, Indicador) Then
                    PonerModo 2
                    lblIndicador.Caption = Indicador
                Else
                    PonerModo 0
                End If
            End If
        End If
            
    Case 5 'Lineas Movimientos Almacenes
        If InsertarModificarLinea Then
            'Reestablecemos los campos y ponemos el grid
            DataGrid1.AllowAddNew = False
'            CargaGrid True
            If ModificaLineas = 1 Then 'Insertar
                CargaGrid True
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                TerminaBloquear
                CargaGrid True
                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                ModificaLineas = 0
'                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
                PonerModo 2
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click()
'    Set frmArt = New frmAlmArticu2
'    'frmArt.DatosADevolverBusqueda3 = "@1@" 'Abre en Modo busqueda
'    frmArt.DesdeTPV = False
'    frmArt.Show vbModal
'    Set frmArt = Nothing
'    PonerFoco txtAux(0)

    Set FrmArt = New frmBasico2
    AyudaArticulos FrmArt, txtAux(0)
    Set FrmArt = Nothing
    
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
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then '2 = Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
           ' PonerBotonCabecera True
            DataGrid1.Refresh
            DataGrid1.Enabled = True
            PonerModo 2
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
        'PonerBotonCabecera False
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
    End If
End Sub


Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton A�adir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'A�adir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(9).Image = 10 'Mantenimiento L�neas
'        .Buttons(10).Image = 39 'Actualizar
'        .Buttons(12).Image = 16 'Imprimir
'        .Buttons(13).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
'    End With
    
    
    For I = 0 To imgBuscar.Count - 1
        imgBuscar(I).Picture = imgBuscar(0).Picture
    Next

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
        .Buttons(1).Image = 13 '39 ' actualizar dto/familia
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "REG"
    
    'campo situacio solo en tabla scamov
    Me.chkImpresion.visible = Not EsHistorico
    'Campo Hora solo en el Historico
    Me.Label4.visible = EsHistorico
    Me.Text1(5).visible = EsHistorico
    
    cadSeleccion = ""
   
    If Not EsHistorico Then
        NombreTabla = "scamov"
        NomTablaLineas = "slimov" 'Tabla lineas de Movimientos
        Me.Caption = "Movimientos de Almacen"
    Else
        NombreTabla = "schmov"
        NomTablaLineas = "slhmov"
        CargarTagsHco Me, "scamov", NombreTabla
        Me.Caption = "Hist�rico Movimientos de Almacen"
    End If
    Ordenacion = " ORDER BY codmovim"
    
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> -1 Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " where codmovim=" & hcoCodMovim & " and fecmovim= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
    Else
         CadenaConsulta = CadenaConsulta & " WHERE codmovim = -1"
    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then 'Se llama desde DblClick frmAlmMovimArticulos
                                    'Se carga con el valor del registro del DblClick
        Data1.Recordset.MoveFirst
        Me.Text1(0).Text = Format(Data1.Recordset!codMovim, "0000000")
        Me.Text1(1).Text = Data1.Recordset!fecmovim
        Me.Text1(5).Text = Format(Data1.Recordset!hormovim, "hh:mm:ss")
        'Cod. Almacen
        Me.Text1(2).Text = Format(Data1.Recordset!codAlmac, "000")
        Me.Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac", "codalmac")
        'Cod. Trabajador
        Me.Text1(3).Text = Format(Data1.Recordset!CodTraba, "0000")
        Me.Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
        'Observaciones
        Text1(4).Text = DBLet(Data1.Recordset!observa1, "T")
        CargaGrid True
    Else
        CargaGrid False '(Modo = 2) 'False
    End If
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
    
    DataGrid1.Columns(0).visible = False 'Cod. Movim
    DataGrid1.Columns(1).visible = False 'Numlinea
    I = 2
    
    'Cod. Art�culo
    DataGrid1.Columns(I).Caption = "Art�culo"
    DataGrid1.Columns(I).Width = 2000
    
    'Nombre Art�culo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Nombre Art�culo"
    DataGrid1.Columns(I).Width = 3500
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cantidad"
    DataGrid1.Columns(I).Width = 1600
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    
    'tipo Movimiento
    I = I + 1
    DataGrid1.Columns(I).Caption = "T.Mov."
    DataGrid1.Columns(I).Width = 700
    DataGrid1.Columns(I).Alignment = dbgCenter
    
    'Observaciones
    I = I + 1
    DataGrid1.Columns(I).Caption = "Observaciones"
    DataGrid1.Columns(I).Width = 6050
       
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
    DataGrid1.RowHeight = 350
    
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = 290
        Next I
        Me.cmdAux.Top = 290
        Me.cboAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        CargarComboAux
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                If I <> 1 Then txtAux(I).Locked = False
            Next I
            cmdAux.Enabled = True
            cboAux.Enabled = True
            cboAux.ListIndex = -1
        Else  'Poner valor a los txtAux
            For I = 0 To txtAux.Count - 2
                txtAux(I).Text = DataGrid1.Columns(I + 2).Text
            Next I
            Select Case DataGrid1.Columns(5).Value
                Case "S"
                    Me.cboAux.ListIndex = 0
                Case "E"
                    Me.cboAux.ListIndex = 1
            End Select
            txtAux(3).Text = DataGrid1.Columns(6).Text
            txtAux(0).Locked = True
            cmdAux.Enabled = False
            cboAux.Enabled = True
            txtAux(2).Locked = False
            txtAux(3).Locked = False
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 240
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        'Fijamos altura y posici�n Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        cboAux.Top = alto - 5
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(2).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width  'Nombre Artic
        txtAux(1).Width = DataGrid1.Columns(3).Width - 35
        I = 2 'Cantidad
        txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 25
        txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 20
        'Tipo Movimiento
        cboAux.Left = txtAux(2).Left + txtAux(2).Width + 20
        cboAux.Width = DataGrid1.Columns(5).Width + 10
        I = 3 'Observac
        txtAux(I).Left = cboAux.Left + cboAux.Width + 30
        txtAux(I).Width = DataGrid1.Columns(6).Width - 60
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux.visible = visible
    cboAux.visible = visible
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
Dim Indice As Byte
    Indice = CByte(Me.imgBuscar(0).Tag)
    Text1(Indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Modo <> 5 Then 'Estamos en Cabecera
            'Recupera todo el registro de Traspaso Almacenes
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Estamos en Lineas
            'Llamamos desde el boton auxiliar de Art�culos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim Aux As String
Dim cadB As String

    HaDevueltoDatos = True
    Screen.MousePointer = vbHourglass
    cadB = ""
    Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
    cadB = Aux
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim Indice As Byte
    Indice = 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(Indice - 2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0 'Codigo Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
        Case 1  'Cod. Trabajador
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(3)
            Set frmT = Nothing

    End Select
    PonerFoco Text1(Index + 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Indice = 1
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(1)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then   'Eliminar lineas Movimiento Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Movimiento Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
Dim vWhere As String

    If Modo = 5 Then  'Modificar LINEAS
        vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
        If BloqueaRegistro(NomTablaLineas, vWhere) Then BotonModificarLinea
    Else 'Modificar Cabecera
       If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
     If Modo = 5 Then  'A�adir lineas Movimiento Almacenes
        BotonAnyadirLineas
    Else 'A�adir Cabecera Movimiento Almacenes
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

Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 4 Then ConseguirFoco Text1(Index), Modo
End Sub



Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 3 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1: KEYFecha2 KeyAscii, 0 ' fecha
                Case 2: KEYBusqueda KeyAscii, 0 'almacen
                Case 3: KEYBusqueda KeyAscii, 1 'trabajador
            End Select
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub KEYFecha2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Bloquear el contador si no estamos en busquedas
    If (Modo <> 1) And (Index = 0) Then BloquearTxt Text1(0), True, True

    Select Case Index
        Case 0 'Codigo Movimiento Almacen
            Text1(Index).Text = Format(Text1(Index).Text, "0000000")
        Case 1 'Fecha
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 2 'Codigo Almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 3  'Codigo Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 4 'Observaciones
            If Text1(Index).Text <> "" Then Text1(Index).Text = QuitarCaracterEnter(Text1(Index).Text)
    End Select
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadirLineas
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case Else
    End Select
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Actualizar
           BotonActualizar
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 0: KEYBusqueda2 KeyAscii, 0 'articulo
            End Select
        Else
            KEYpress KeyAscii
        End If
    
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 0 'Cod ARTICULO
            If txtAux(Index).Text = "" Then
                txtAux(Index + 1).Text = ""
            Else
                 PonerArticulo txtAux(0), txtAux(1), Text1(2).Text, CodTipoMov, ModificaLineas
            End If
            
        Case 2 'CANTIDAD (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            PonerFormatoDecimal txtAux(Index), 1
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
           mnModificar_Click
        Case 3 'Eliminar
           mnEliminar_Click
        Case 5 'Busqueda
           mnBuscar_Click
        Case 6 'Ver Todos
           mnVerTodos_Click
        Case 8 'Imprimir
           BotonImprimir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    lblIndicador.Caption = ""
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    b = (Kmodo = 2)
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Como el campo 0 es clave primaria, NO se puede modificar, es contador
    BloquearTxt Text1(0), (Modo <> 1), True
    cmdRegresar.visible = False
    
'    Me.cmdRegresar.visible = (Not b) And Not EsHistorico
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = b
'    Else
'        cmdRegresar.visible = False
'    End If
    
    '=================================================
    b = Modo <> 0 And Modo <> 2 'And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I

    Me.chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
Dim I As Byte
Dim bAux As Boolean


    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For I = 1 To 3
        Toolbar1.Buttons(I).Enabled = Not EsHistorico
    Next I
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    If Not EsHistorico Then
        'Modo 2. Hay datos y estamos visualizandolos
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
        'Lineas Movimientos Almacenes
'        Toolbar1.Buttons(9).Enabled = b
        'Actualizar
        Toolbar5.Buttons(1).Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
    Else
        'Actualizar
        FrameBotonGnral2.Enabled = False
        FrameBotonGnral2.visible = False
        FrameDesplazamiento.Left = FrameBotonGnral2.Left
    End If
    
    b = (Modo = 2) And Not EsHistorico
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkImpresion.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
''Botones de Desplazamiento de la Toolbar
'
'    Select Case Modo
'        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
'            If Data2.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data2, Index
'        Case Else 'Datos de Cabecera
'            If Data1.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data1, Index
'            PonerCampos
'    End Select
'Para desplazarse por los registros de control Data
    
    DesplazamientoData Data1, Index, True
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    tabla = NomTablaLineas

    SQL = "SELECT " & tabla & ".codmovim, "
    SQL = SQL & tabla & ".numlinea, " & tabla & ".codartic, Articulos.nomartic, "
    SQL = SQL & tabla & ".cantidad, if(" & tabla & ".tipomovi=0,""S"",""E"") as tipomovi, "
    SQL = SQL & tabla & ".motimovi "
    SQL = SQL & " FROM ((" & tabla & " LEFT JOIN sartic AS Articulos ON " & tabla & ".codartic ="
    SQL = SQL & " Articulos.codartic))"
    If enlaza Then
        SQL = SQL & " WHERE codmovim = " & Data1.Recordset!codMovim
        If EsHistorico Then SQL = SQL & " AND fecmovim = " & DBSet(Data1.Recordset!fecmovim, "F")
    Else
        SQL = SQL & " WHERE codmovim = -1"
    End If
    SQL = SQL & " ORDER BY " & tabla & ".numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    PonerModo (5)
    ModificaLineas = 0
    PonerBotonCabecera True
    CargaGrid True
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
           
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrid False
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(1).Text = NomTraba
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    
    PonerModo 5
    
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    vWhere = ObtenerWhereCP(False)
    cmdAceptar.Tag = SugerirCodigoSiguienteStr("slimov", "numlinea", vWhere)
    
'    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data2

    DataGrid1.Enabled = False
    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
Dim I As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub

    PonerModo 5


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar

    Screen.MousePointer = vbHourglass
    
'    PonerBotonCabecera False
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    cmdAceptar.Tag = Data2.Recordset!numlinea
    
    CargaTxtAux True, False
    DataGrid1.Enabled = False
    PonerFoco txtAux(2) 'Poner el foco
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Cabecera de Movimiento Almacen." & vbCrLf
    SQL = SQL & "----------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a eliminar el Movimiento:"
    SQL = SQL & vbCrLf & " N� Movim. : " & Text1(0).Text
    SQL = SQL & vbCrLf & " Fecha Mov.: " & CStr(Data1.Recordset.Fields(1))
    SQL = SQL & vbCrLf & " Almacen   : " & Text1(2).Text
    SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
    
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        NumRegElim = Data1.Recordset.Fields(0)
        vTipoMov.DevolverContador CodTipoMov, NumRegElim
        Set vTipoMov = Nothing
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Movimiento", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        SQL = " WHERE  codmovim=" & Data1.Recordset!codMovim
        
        'Lineas
        conn.Execute "Delete  from slimov " & SQL
        
        'Cabeceras
        conn.Execute "Delete  from scamov " & SQL
                      
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
    SQL = "Seguro que desea eliminar la l�nea del Art�culo:"
    SQL = SQL & vbCrLf & "C�digo: " & Data2.Recordset!codArtic
    SQL = SQL & vbCrLf & "Descripci�n: " & Data2.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from slimov where codmovim=" & Data2.Recordset!codMovim
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        SQL = SQL & " and codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        conn.Execute SQL
        CancelaADODC Me.Data2
        CargaGrid True
        CancelaADODC Me.Data2
    End If
    ModificaLineas = 0
    
Error2:
    Screen.MousePointer = vbDefault
    ModificaLineas = 0
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar L�nea de Art�culo de Movimiento Almacen", Err.Description
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim vStock As String
'Dim vstockOrig As Single  'Stock en el almacen Origen
'Dim SQL As String, devuelve As String

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    

    'Comprobar que todos los Art�culos estan en el nuevo almacen
    If Modo = 4 Then 'Modificando
        b = ComprobarStocksLineas
    End If

    DatosOk = True
End Function



Private Function ComprobarStocksLineas() As Boolean
'Comprobar para todas las lineas del traspaso que:
' - todos los Art�culos entan en el almacen origen
' - Comprobar que hay suficiente stock en el Almacen Origen de ese Articulo
Dim b As Boolean

    If Not Data2.Recordset.EOF Then  'Si hay lineas
        Data2.Recordset.MoveFirst
        b = True
        
        While Not Data2.Recordset.EOF And b
            If Data2.Recordset!tipomovi = "S" Then 'Mov. de salida
                b = ComprobarStock(Data2.Recordset!codArtic, Text1(2).Text, Data2.Recordset!cantidad, CodTipoMov)
            End If
            Data2.Recordset.MoveNext
        Wend
        Data2.Recordset.MoveFirst
    End If
    ComprobarStocksLineas = b
End Function




Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkLinea = False
    b = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Art�culo no puede ser nulo", vbExclamation
        b = False
        Exit Function
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Art�culos"
         b = False
    ElseIf Not IsNumeric(txtAux(2).Text) Then
        MsgBox "El campo Cantidad debe ser num�rico", vbExclamation
        b = False
    End If
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
     
    'Comprobamos si ya existe una linea con el art�culo, solo si estamos insertando (ModificaLineas=1)
    'BD 1: conexion a BD Ariges
    If ModificaLineas = 1 Then
        devuelve = DevuelveDesdeBDNew(conAri, "slimov", "codmovim", "codmovim", Text1(0).Text, "N", , "codartic", txtAux(0).Text, "T")
        If devuelve <> "" Then
            b = False
            devuelve = "Ya hay una l�nea con ese Art�culo: " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripci�n: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
        
        'Comprobamos si existe el art�culo, solo si estamos insertando (ModificaLineas=1)
        If Trim(txtAux(1).Text) = "" Then
            b = False
            devuelve = "No existe el Art�culo " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripci�n: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
    End If
    If Not b Then Exit Function
    
    
    'Entrada o salida marcado
    If cboAux.ListIndex = -1 Then
        MsgBox "Seleccione tipo de movimiento", vbExclamation
        PonerFocoCbo cboAux
        b = False
        Exit Function
    End If
    
    
    'Comprobar que hay suficiente stock en el Almacen
    'Si es movimiento de Salida
    If Me.cboAux.ListIndex = 0 Then
        b = ComprobarStock(txtAux(0).Text, Text1(2).Text, txtAux(2).Text, CodTipoMov)
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
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim SQL As String, Cad As String
On Error GoTo EInsertarModificarLinea
    
    SQL = ""
    InsertarModificarLinea = False
    
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea Then 'INSERTAR
            SQL = "INSERT INTO slimov (codmovim,numlinea,codartic,cantidad,tipomovi,motimovi) "
            SQL = SQL & " VALUES (" & Val(Text1(0).Text) & ", "
            SQL = SQL & cmdAceptar.Tag & ", "
            SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(2).Text, "N") & ", "
            If cboAux.ListIndex = -1 Then
                Cad = ValorNulo
            Else
                 Cad = cboAux.ItemData(cboAux.ListIndex)
            End If
            SQL = SQL & CSng(Cad) & ","
            SQL = SQL & DBSet(txtAux(3).Text, "T") & ") "
        End If
    Case 2 'Modificar
        If DatosOkLinea Then
            SQL = "UPDATE slimov Set cantidad = " & DBSet(txtAux(2).Text, "N")
            SQL = SQL & ", tipomovi = " & cboAux.ItemData(cboAux.ListIndex)
            SQL = SQL & ", motimovi = " & DBSet(txtAux(3).Text, "T")
            SQL = SQL & " WHERE codmovim =" & Val(Text1(0).Text) & " AND "
            SQL = SQL & " numlinea =" & Val(cmdAceptar.Tag)
        End If
    End Select
            
    If SQL <> "" Then
        conn.Execute SQL
        
        
        If ModificaLineas = 1 Then
            'Si tiene componentes preguntamos si queire insertar las lineas
            SQL = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codartic", txtAux(0).Text, "T")
            If Val(SQL) >= 1 Then
                If MsgBox("El articulo tiene componentes" & vbCrLf & "�Desea insertarlos?", vbQuestion + vbYesNoCancel) = vbYes Then
                
                    SQL = "select " & Val(Text1(0).Text) & "," & cmdAceptar.Tag & "+numlinea,codarti1,"
                    SQL = SQL & " (cantidad*" & DBSet(txtAux(2).Text, "N") & ")" & "," & cboAux.ItemData(cboAux.ListIndex)
                    SQL = SQL & ",concat('COMPONENTES ',codartic) from sarti1 where codartic=" & DBSet(txtAux(0).Text, "T")
                    SQL = "INSERT INTO slimov (codmovim,numlinea,codartic,cantidad,tipomovi,motimovi) " & SQL
                    ejecutar SQL, False
                    Espera 0.2
                End If
            End If
        End If
        
        
        
        
        
        
        InsertarModificarLinea = True
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas Traspaso Almacenes" & vbCrLf & Err.Description
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

'    'Llamamos a al form
'    Cad = ""
'    'Registro de la tabla de cabeceras: scamov
'    Cad = Cad & ParaGrid(Text1(0), 15, "N� Mov.")
'    Cad = Cad & ParaGrid(Text1(1), 20, "Fecha")
'    Cad = Cad & ParaGrid(Text1(2), 10, "Alm.")
'    Cad = Cad & "Desc. Alm. Orig|salmpr|nomalmac|T||40�"
'    tabla = "(" & NombreTabla & " LEFT JOIN salmpr ON " & NombreTabla & ".codalmac=salmpr.codalmac" & ") "
'    Titulo = Me.Caption
'
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco Text1(kCampo)
''            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmB = New frmBasico2
    AyudaAlmMovimientosPrev frmB, EsHistorico
    Set frmB = Nothing


End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        MsgBox "Introducir criterios de b�squeda", vbExclamation
        PonerFoco Text1(0)
    End If
    
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
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
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac")
    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel

    
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim SQL As String, EnAlmDest As String
Dim cantidad As Single, vStock As Single
Dim devuelve As String
Dim vCantidad As String
    On Error GoTo EActualizarStock

    ActualizarStocks = False
    While Not Data2.Recordset.EOF
        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", Data2.Recordset!codArtic, "T")
        If Val(devuelve) = 1 Then 'Hay control de stock

            cantidad = Data2.Recordset!cantidad 'Cant a traspasar
            vCantidad = TransformaComasPuntos(CStr(CCur(cantidad)))
            If Data2.Recordset!tipomovi = "E" Then 'Mov. de Entrada
                '==== Aumentar el stock en el Almacen
                'Comprobar que existe el articulo en Almacen Destino
                EnAlmDest = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Data2.Recordset!codArtic, "T", , "codalmac", Text1(2).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese art�culo en Almacen
                    SQL = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                    SQL = SQL & " VALUES (" & DBSet(Data2.Recordset!codArtic, "T") & "," & Val(Text1(2).Text) & ",''," & DBSet(cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
                Else 'Existe el artic en almac. Dest -> Aumentar stock
                    SQL = "UPDATE salmac Set canstock = canstock + " & vCantidad
                    SQL = SQL & " WHERE codartic =" & DBSet(Data2.Recordset!codArtic, "T") & " AND "
                    SQL = SQL & " codalmac =" & Data1.Recordset!codAlmac
                End If
                
            Else 'Mov. de Salida
                '==== Disminuir Stock en Almacen Origen
                EnAlmDest = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", Data2.Recordset!codArtic, "T", , "codalmac", Text1(2).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese art�culo en Almacen
                    devuelve = "No existe en el Almacen: " & Data1.Recordset!codAlmac & vbCrLf
                    devuelve = devuelve & "El Art�culo: " & Data2.Recordset!codArtic
                    MsgBox devuelve, vbExclamation
                Else 'Existe el artic en almac. Dest -> Disminuir stock
                    vStock = CSng(EnAlmDest)
                    If ComprobarHayStock(vStock, cantidad, Data2.Recordset!codArtic, Data2.Recordset!NomArtic, CodTipoMov) Then
                        SQL = "UPDATE salmac Set canstock = canstock - " & vCantidad
                        SQL = SQL & " WHERE codartic =" & DBSet(Data2.Recordset!codArtic, "T") & " AND "
                        SQL = SQL & " codalmac =" & Data1.Recordset!codAlmac
                    End If
                End If
            End If
            
            conn.Execute SQL
        End If
        Data2.Recordset.MoveNext
    Wend
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ActualizarStocks = False
    Else
        ActualizarStocks = True
    End If
EActualizarStock:
End Function


Private Sub BotonActualizar()
'Actualizar Traspaso Almacen
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n Movimiento para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then
        MsgBox "No hay lineas insertadas para este N� de Movimiento", vbExclamation
        Exit Sub
    End If
    
    SQL = "Actualizaci�n Movimientos Almacen." & vbCrLf
    SQL = SQL & "-------------------------------------------" & vbCrLf & vbCrLf

    If Not CBool(Data1.Recordset.Fields(5).Value) Then 'Informe No Impreso
        SQL = SQL & "NO ESTA IMPRESO EL MOVIMIENTO:" & vbCrLf
    End If
    SQL = SQL & vbCrLf & "N� Movim. : " & Format(Data1.Recordset.Fields(0), "0000000")
    SQL = SQL & vbCrLf & "Fecha        : " & CStr(Data1.Recordset.Fields(2))
    SQL = SQL & vbCrLf & "Almacen    : " & Format(Data1.Recordset.Fields(1), "000") & " - " & Text2(0).Text
    SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
'    Else 'Informe Impreso
'        SQL = "Actualizaci�n Movimientos Almacen." & vbCrLf
'        SQL = SQL & "--------------------------------------------" & vbCrLf & vbCrLf
'
'        SQL = SQL & "Va a Actualizar el Movimiento:"
'        SQL = SQL & vbCrLf & " N� Movim.  :  " & Format(Data1.Recordset.Fields(0), "0000000")
'        SQL = SQL & vbCrLf & " Fecha Mov.:  " & CStr(Data1.Recordset.Fields(2))
'        SQL = SQL & vbCrLf & " Almacen     :  " & CStr(Format(Data1.Recordset.Fields(1), "000"))
'        SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
'        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
'            Exit Sub
'        End If
'    End If
    
    Me.ProgressBar1.visible = True
    Me.ProgressBar1.Value = 0
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarTraspaso Then
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
            PonerModo 2
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            Espera 0.3
            Me.Refresh
        End If
    
    End If
    Me.ProgressBar1.visible = False
End Sub


Private Function ActualizarTraspaso() As Boolean
Dim Donde As String
Dim devuelve As String
Dim bol As Boolean
On Error GoTo EActualizarTraspaso
    
    'Comprobamos que no existe en historico
    devuelve = DevuelveDesdeBDNew(conAri, "schmov", "codmovim", "codmovim", Data1.Recordset!codMovim, "N", , "fecmovim", Data1.Recordset!fecmovim, "F")
    If Trim(devuelve) <> "" Then
        devuelve = "Ya existe en el hist�rico el movimiento:" & vbCrLf
        devuelve = devuelve & " N�: " & Data1.Recordset!codMovim & vbCrLf
        devuelve = devuelve & " Fecha: " & Data1.Recordset!fecmovim
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    If Not ComprobarStocksLineas Then Exit Function
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    Donde = ""
    bol = ActualizarElTraspasoAqui(Donde)

EActualizarTraspaso:
    If Err.Number <> 0 Or Donde <> "" Then
        devuelve = "Actualizar Movimiento." & vbCrLf & "----------------------------" & vbCrLf
        devuelve = devuelve & Donde
        MuestraError Err.Number, devuelve, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ActualizarTraspaso = True
    Else
        conn.RollbackTrans
        MuestraError Err.Number, devuelve, Err.Description
    End If
End Function


Private Function ActualizarElTraspasoAqui(ByRef ADonde As String) As Boolean

    ActualizarElTraspasoAqui = False
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en historico cabeceras movimientos almacen"
    If Not InsertarCabeceraHistorico Then Exit Function
    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Movimientos Almacen"
    If Not InsertarLineasHistorico Then Exit Function
    IncrementarProgres 2
    
    
     'Modificar stock
    ADonde = "Actualizando Stocks Almacenes"
    If Not ActualizarStocks() Then Exit Function
    IncrementarProgres 2
    
    
    'Insertamos en Movimientos Art�culos
    ADonde = "Insertando datos en Movimientos de Articulos"
    If Not InsertarMovimArticulos Then Exit Function
    IncrementarProgres 2
   
    
    'Borramos cabeceras y lineas del asiento
    ADonde = "Borrar cabeceras y lineas en Movimientos Almacen"
    If Not BorrarTraspaso(False) Then Exit Function
    IncrementarProgres 2
    
    ActualizarElTraspasoAqui = True
    ADonde = ""
End Function


Private Function InsertarCabeceraHistorico() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarCab

    SQL = "SELECT codmovim,codalmac,fecmovim,codtraba,observa1 from scamov where "
    SQL = SQL & " codmovim =" & Data1.Recordset!codMovim
    SQL = SQL & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        SQL = "INSERT INTO schmov (codmovim, fecmovim,hormovim,codalmac,codtraba,observa1) "
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(RS.Fields(2).Value, "yyyy-mm-dd") & "','"
        SQL = SQL & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields(1).Value & ", " & RS.Fields(3).Value
        SQL = SQL & ", " & DBSet(RS.Fields(4).Value, "T") & ")"
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


Private Function InsertarLineasHistorico() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarLineas

    SQL = "SELECT codmovim, numlinea, codartic, cantidad, tipomovi, motimovi from slimov where "
    SQL = SQL & " codmovim =" & Data1.Recordset!codMovim
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    RS.MoveFirst
    While Not RS.EOF
        SQL = "INSERT INTO slhmov (codmovim, fecmovim, numlinea, codartic, cantidad, tipomovi, motimovi)"
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "', "
        SQL = SQL & RS.Fields(1).Value & ", " & DBSet(RS.Fields(2).Value, "T") & ", "
        SQL = SQL & DBSet(RS.Fields(3).Value, "N") & ", " & RS.Fields(4).Value
        SQL = SQL & ", '" & RS.Fields(5).Value & "')"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        RS.Close
        Set RS = Nothing
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Function InsertarMovimArticulos() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vImporte As Single, vPrecioVenta As String
Dim vTipoMov As CTiposMov
Dim bol As Boolean
Dim Cad As String
On Error GoTo EInsertar

    bol = True
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        'Se han cargado correctamente los valores de la clase
        SQL = "SELECT scamov.codmovim, codalmac, fecmovim, codtraba, numlinea, codartic, cantidad, tipomovi "
        SQL = SQL & " from scamov LEFT JOIN slimov on scamov.codmovim=slimov.codmovim "
        SQL = SQL & " WHERE scamov.codmovim =" & Data1.Recordset!codMovim
        SQL = SQL & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
            'Obtener el precio de venta del articulo, si tiene control de stock
            Cad = "ctrstock"
            vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", RS.Fields!codArtic, "T", Cad)
            If vPrecioVenta <> "" Then
                vImporte = RS.Fields!cantidad * CSng(vPrecioVenta)
            Else
                vImporte = 0
            End If
            If Val(Cad) = 1 Then
                SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                SQL = SQL & " VALUES (" & DBSet(RS.Fields!codArtic, "T") & ", " & RS.Fields!codAlmac & ", '" & Format(RS.Fields!fecmovim, "yyyy-mm-dd") & "', '"
                SQL = SQL & Format(RS.Fields!fecmovim & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields!tipomovi & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(RS.Fields!cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & RS.Fields!CodTraba & ", '"
                SQL = SQL & vTipoMov.LetraSerie & "', " & RS.Fields!codMovim & ", " & RS.Fields!numlinea & ")"
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
    Else
        bol = False
    End If
    Set vTipoMov = Nothing
    RS.Close
    Set RS = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        Set vTipoMov = Nothing
        RS.Close
        Set RS = Nothing
    End If
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function



Private Sub IncrementarProgres(Veces As Integer)
On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * 10)
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub


Private Function BorrarTraspaso(EnHistorico As Boolean) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim SQL As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    SQL = "Delete from "
    If EnHistorico Then
        SQL = SQL & "slhmov"
        SQL = SQL & " WHERE codmovim = " & Data1.Recordset!codMovim
        SQL = SQL & " AND fecmovim = '" & Data1.Recordset!fecmovim & "'"
    Else
        SQL = SQL & "slimov"
        SQL = SQL & " WHERE codmovim = " & Data1.Recordset!codMovim
    End If
    conn.Execute SQL
    
    'La cabecera
    SQL = "Delete from "
    If EnHistorico Then
        SQL = SQL & "schmov"
        SQL = SQL & " WHERE codmovim =" & Data1.Recordset!codMovim
        SQL = SQL & " AND fecmovim='" & Data1.Recordset!fecmovim & "'"
    Else
        SQL = SQL & "scamov"
        SQL = SQL & " WHERE codmovim =" & Data1.Recordset!codMovim
    End If
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function


Private Sub CargarComboAux()
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Entrada, 1-Salida

    cboAux.Clear
    cboAux.AddItem "S"
    cboAux.ItemData(cboAux.NewIndex) = 0
    
    cboAux.AddItem "E"
    cboAux.ItemData(cboAux.NewIndex) = 1
        
End Sub


Public Sub ActualizarSituacionImpresion()
Dim Cad As String, Indicador As String
On Error GoTo EImpresion
   
    Cad = "(" & ObtenerWhereCP(False) & ")"
    If SituarData(Data1, Cad, Indicador) Then
        If Modo <> 5 Then
            PonerModo 2
        Else
            PonerModo 5
        End If
        PonerCampos
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
EImpresion:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonImprimir()
        If Text1(0).Text = "" Then Exit Sub
        frmListado.NumCod = Text1(0).Text
'        If Not EsHistorico Then
            'AbrirListado (8) '8: Informe Movimientos Almacen
            frmInformesNew.NumCod = Text1(0).Text
            frmInformesNew.EsHco = EsHistorico
            frmInformesNew.OpcionListado = 8
            frmInformesNew.Show vbModal
            If Not EsHistorico Then ActualizarSituacionImpresion
'        Else
'            BotonImprimirHco
'        End If
End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim Cad As String
Dim numParam As Byte
Dim nomDocu As String


    cadParam = "|"
    numParam = 0
    pRptvMultiInforme = 0
    If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub

    indRPT = 4 '4: Historico Movimientos de Almacen
    If PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .NumeroParametros = numParam
            .NombreRPT = nomDocu
            .NombrePDF = pPdfRpt
            .EnvioEMail = False
            .Opcion = 8
            .Titulo = "Hist. Movimientos Alm."
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                Cad = "{schmov.codmovim}= " & Data1.Recordset!codMovim
                Cad = Cad & " and {schmov.fecmovim}= Date(" & Year(Data1.Recordset!fecmovim) & "," & Month(Data1.Recordset!fecmovim) & "," & Day(Data1.Recordset!fecmovim) & ")" & ""
                
                .FormulaSeleccion = Cad
            End If
            .Show vbModal
        End With
    End If
End Sub



Private Function InsertarMovimiento(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean
On Error GoTo EInsertarMovim
    
    bol = True
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    MenError = "Error al insertar en la tabla de Movimientos(smovim)."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del recibo."
    bol = vTipoMov.IncrementarContador(CodTipoMov)

EInsertarMovim:
        If Err.Number <> 0 Then
            MenError = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarMovimiento = True
        Else
            conn.RollbackTrans
            InsertarMovimiento = False
        End If
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
'Obtiene la sentencia WHERE para seleccionar registros de la tabla por Clave Primaria
On Error Resume Next
    If conWhere Then
        ObtenerWhereCP = " WHERE codmovim= " & Val(Text1(0).Text)
    Else
        ObtenerWhereCP = " codmovim= " & Val(Text1(0).Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        SQL = CadenaInsertarDesdeForm(Me)
        
        If SQL <> "" Then
            If InsertarMovimiento(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                 'Ponerse en Modo Insertar Lineas
                BotonLineas
                BotonAnyadirLineas
            End If
        End If
    End If
    Set vTipoMov = Nothing
End Sub