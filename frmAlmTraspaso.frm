VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmTraspaso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Almacenes"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
   Icon            =   "frmAlmTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12960
      TabIndex        =   39
      Top             =   270
      Width           =   1515
   End
   Begin VB.Frame FrameToolAux0 
      Height          =   645
      Left            =   225
      TabIndex        =   37
      Top             =   2565
      Width           =   1815
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   38
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copiar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5040
      TabIndex        =   35
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   36
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
               Object.ToolTipText     =   "ï¿½ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   33
      Top             =   90
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   34
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   31
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   32
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
      Index           =   6
      Left            =   6465
      TabIndex        =   29
      Tag             =   "Hora|H|N|||scatra|hormovim|hh:mm:ss|N|"
      Text            =   "Text1"
      Top             =   930
      Width           =   990
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
      Left            =   8010
      TabIndex        =   28
      Tag             =   "Situación Impresión|N|N|||scatra|situacio||N|"
      Top             =   945
      Width           =   1125
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
      Left            =   2160
      TabIndex        =   26
      ToolTipText     =   "Buscar artículo"
      Top             =   5040
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   14
      Text            =   "observac"
      Top             =   5040
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
      Left            =   4920
      MaxLength       =   16
      TabIndex        =   13
      Text            =   "cantidad"
      Top             =   5040
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
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "nombre artic"
      Top             =   5040
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
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   11
      Text            =   "codartic"
      Top             =   5040
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
      Left            =   12195
      TabIndex        =   6
      Top             =   8790
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
      Left            =   13470
      TabIndex        =   7
      Top             =   8790
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
      Left            =   13455
      TabIndex        =   25
      Top             =   8775
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   8670
      Width           =   3000
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
         TabIndex        =   24
         Top             =   180
         Width           =   2595
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
      Index           =   1
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   1755
      Width           =   6195
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
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1365
      Width           =   6195
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
      Index           =   2
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   2145
      Width           =   6195
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
      Height          =   1500
      Index           =   5
      Left            =   9315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "Observaciones|T|S|||scatra|observa1||N|"
      Text            =   "frmAlmTraspaso.frx":000C
      Top             =   990
      Width           =   5175
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
      Index           =   4
      Left            =   2205
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Cod. Trabajador|N|N|0|9999|scatra|codtraba|0000|N|"
      Text            =   "Text1"
      Top             =   2145
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
      Index           =   3
      Left            =   2205
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "Almacen Destino|N|N|0|999|scatra|almadest|000|N|"
      Text            =   "Text1"
      Top             =   1755
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
      Left            =   2205
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Almacen Origen|N|N|0|999|scatra|almaorig|000|N|"
      Text            =   "Text1"
      Top             =   1365
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
      Left            =   4470
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||scatra|fechatra|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   930
      Width           =   1350
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   480
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
      Alignment       =   2  'Center
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
      Left            =   2205
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Nº Traspaso|N|S|0||scatra|codtrasp|0000000|S|"
      Text            =   "0000000"
      Top             =   930
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9720
      Top             =   480
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
      Left            =   3480
      TabIndex        =   27
      Top             =   8835
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmTraspaso.frx":0012
      Height          =   5280
      Left            =   225
      TabIndex        =   8
      Top             =   3330
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9313
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
      Index           =   3
      Left            =   10800
      Tag             =   "-1"
      ToolTipText     =   "Buscar actividad"
      Top             =   720
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1905
      Picture         =   "frmAlmTraspaso.frx":0027
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   1395
      Width           =   240
   End
   Begin VB.Label Label7 
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
      Left            =   5985
      TabIndex        =   30
      Top             =   960
      Width           =   465
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1905
      ToolTipText     =   "Buscar trabajador"
      Top             =   2145
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1905
      ToolTipText     =   "Buscar almacen"
      Top             =   1755
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   4155
      Picture         =   "frmAlmTraspaso.frx":0A29
      ToolTipText     =   "Buscar fecha"
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Left            =   9315
      TabIndex        =   19
      Top             =   720
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
      TabIndex        =   18
      Top             =   2145
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Almacén Destino"
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
      Top             =   1755
      Width           =   1710
   End
   Begin VB.Label Label3 
      Caption         =   "Almacén Origen"
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
      Top             =   1365
      Width           =   1620
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
      Left            =   3420
      TabIndex        =   15
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Traspaso"
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
      TabIndex        =   10
      Top             =   960
      Width           =   1320
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
      Left            =   315
      TabIndex        =   9
      Top             =   8820
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
Attribute VB_Name = "frmAlmTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'historico schmov, y solo en modo de consulta

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del histórico del traspaso seleccionado (solo consulta)
Public hcoCodMovim As Long 'cod. traspaso del historico
Public hcoFechaMovim As Date 'Fecha del historico

'--------------------------------------------------------------------------

Private WithEvents frmB As frmBasico2 'frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores 'Mto de Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2 'frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmVarN As frmVariosNew
Attribute frmVarN.VB_VarHelpID = -1

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

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Private HaDevueltoDatos As Boolean

Dim Movimiento As String




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
        If DatosOk(True) Then InsertarCabecera
        
    Case 4 'MODIFICAR
        If DatosOk(True) Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                                                            'Borramos el pedido vinculado al antiguo traspaso
                 If vParamAplic.NumeroInstalacion = vbHerbelca Then BorrarPedidoVinculado_
                    
                
                 
                 PosicionarData
                 
                 'En herbelca volvemos a generar
                 If vParamAplic.NumeroInstalacion = vbHerbelca Then
                    Espera 0.5
                    CrearPedidoVinculado
                    LineaPedidoVinculado 4, 0
                 End If
             End If
         End If
            
    Case 5 'LINEAS Traspaso Almacenes
        If InsertarModificarLinea Then
        
            
            'Reestablecemos los campos
            'y ponemos el grid
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 2 Then TerminaBloquear
            CargaGrid True
            
            If ModificaLineas = 1 Then 'Insertar
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                ModificaLineas = 0
'                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
                DataGrid1.Enabled = True
'                DataGrid1.SetFocus
                PonerModo 2

                PonerFocoGrid Me.DataGrid1
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
'    Set FrmArt = New frmAlmArticu2
    'frmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo busqueda
'    FrmArt.DesdeTPV = False
'    FrmArt.Show vbModal
'    Set FrmArt = Nothing
    Set FrmArt = New frmBasico2
    AyudaArticulos FrmArt, txtAux(0)
    Set FrmArt = Nothing

    PonerFoco txtAux(0)
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
'             PonerBotonCabecera True
            DataGrid1.Refresh
            'PonerFocoBtn Me.cmdRegresar
            DataGrid1.Enabled = True
            PonerModo 2
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdRegresar.visible = False
    End If
End Sub


Private Sub DataGrid1_DblClick()

    
    If Modo <> 2 Then Exit Sub
    
    
    
    ' premp preve preuc
    If DBLet(Data2.Recordset!premp, "N") <> 0 Then davidCodtipom = davidCodtipom & " Precio medio ponderado = " & Format(Data2.Recordset!premp, FormatoPrecio) & vbCrLf
    If DBLet(Data2.Recordset!preUC, "N") <> 0 Then davidCodtipom = davidCodtipom & " Precio última compra   = " & Format(Data2.Recordset!preUC, FormatoPrecio) & vbCrLf
    If DBLet(Data2.Recordset!preve, "N") <> 0 Then davidCodtipom = davidCodtipom & " Precio venta(artículo) = " & Format(Data2.Recordset!preve, FormatoPrecio) & vbCrLf
    
    If davidCodtipom = "" Then davidCodtipom = "***   Ningun precio guardado **"
    
    If EsHistorico Then
        davidCodtipom = "Precio hco: " & vbCrLf & vbCrLf & davidCodtipom
    Else
        davidCodtipom = "Precio actual: " & vbCrLf & vbCrLf & davidCodtipom
    End If
    
    MsgBox davidCodtipom, vbExclamation
    
    davidCodtipom = ""
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
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
'    btnAnyadir = 5 'Posicion del boton Añadir en la toolbar1
'    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(9).Image = 10 'Mantenimiento Líneas
'        .Buttons(10).Image = 39 'Actualizar
'        .Buttons(12).Image = 16 'Imprimir
'        .Buttons(13).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
'    End With
    
    For i = 0 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
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
    
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
            .Buttons(5).Image = 32  'Copiar
        End With
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox
       
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "TRA"
    'campo situacio solo en tabla scatra
    Me.chkImpresion.visible = Not EsHistorico
    'campo Hora solo en tabla hist. schtra
    Me.Label7.visible = EsHistorico
    Me.Text1(6).visible = EsHistorico
    
    cadSeleccion = ""
    
    If Not EsHistorico Then
        NombreTabla = "scatra"
        NomTablaLineas = "slitra" 'Tabla lineas de Traspasos Almacen
        Me.Caption = "Traspaso de Almacen"
        Label6.Caption = "Observaciones"
    Else
        NombreTabla = "schtra"
        NomTablaLineas = "slhtra"
        CargarTagsHco Me, "scatra", NombreTabla
        Me.Caption = "Histórico Traspaso de Almacen"
        'Label6.Caption = "Obs."
    End If
    
    Ordenacion = " ORDER BY codtrasp"
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> -1 Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " where codtrasp=" & hcoCodMovim & " and fechatra= '" & Format(hcoFechaMovim, "yyyy-mm-dd") & "'"
    Else
         CadenaConsulta = CadenaConsulta & " WHERE codtrasp = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then 'Se llama desde DblClick frmAlmMovimArticulos
                                    'Se carga con el valor del registro del DblClick
        Data1.Recordset.MoveFirst
        Me.Text1(0).Text = Format(Data1.Recordset!codtrasp, "0000000")
        Me.Text1(1).Text = Data1.Recordset!fechatra
        Me.Text1(6).Text = Format(Data1.Recordset!hormovim, "hh:mm:ss")
        'Almacen Origen
        Me.Text1(2).Text = Format(Data1.Recordset!almaorig, "000")
        Me.Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac", "codalmac")
        'Almacen Destino
        Me.Text1(3).Text = Format(Data1.Recordset!almadest, "000")
        Me.Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "salmpr", "nomalmac", "codalmac")
        'Cod. Trabajador
        Me.Text1(4).Text = Format(Data1.Recordset!CodTraba, "0000")
        Me.Text2(2).Text = PonerNombreDeCod(Text1(4), conAri, "straba", "nomtraba")
        Text1(5).Text = DBLet(Data1.Recordset!observa1, "T")
        CargaGrid True
        Toolbar1.Buttons(5).Enabled = True 'Imprimir
    Else
        CargaGrid False '(Modo = 2) 'False
    End If
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim Sql As String

    On Error GoTo ECarga

    B = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, Sql, False
      
    DataGrid1.RowHeight = 350
      
      
      
    DataGrid1.Columns(0).visible = False 'Cod. trasp
    DataGrid1.Columns(1).visible = False 'Numlinea
    
    'Para ocultarlas rapido
    DataGrid1.Columns(6).visible = False 'Numlinea
    DataGrid1.Columns(7).visible = False 'Numlinea
    DataGrid1.Columns(8).visible = False 'Numlinea
    
    i = 2
    'Cod. Artículo
    DataGrid1.Columns(i).Caption = "Artículo"
    DataGrid1.Columns(i).Width = 2050
    
    'Nombre Artículo
    i = i + 1
    DataGrid1.Columns(i).Caption = "Nombre Artículo"
    DataGrid1.Columns(i).Width = 4700
    
    'Cantidad
    i = i + 1
    DataGrid1.Columns(i).Caption = "Cantidad"
    DataGrid1.Columns(i).Width = 1650
    DataGrid1.Columns(i).Alignment = dbgRight
    DataGrid1.Columns(i).NumberFormat = FormatoImporte & " "
    
    'Observaciones
    i = i + 1
    DataGrid1.Columns(i).Caption = "Observaciones"
    DataGrid1.Columns(i).Width = 5300
       
       
       
       
       
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
       
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
    
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim i As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = 290
        Next i
        Me.cmdAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
            Next i
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For i = 0 To txtAux.Count - 1
'                If i <> 1 Then txtAux(i).Locked = False
                'LAURA 19/10/2006
                If i <> 1 Then BloquearTxt txtAux(i), False
            Next i
            cmdAux.Enabled = True
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = DataGrid1.Columns(i + 2).Text
            Next i
            BloquearTxt txtAux(0), True
            cmdAux.Enabled = False
            BloquearTxt txtAux(2), False
            BloquearTxt txtAux(3), False
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 240
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        
        'Fijamos altura y posición Top
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(2).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width + 10 'Nom artic
        txtAux(1).Width = DataGrid1.Columns(3).Width - 25
        For i = 2 To txtAux.Count - 1 'Cantidad y Observacion
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 25
            txtAux(i).Width = DataGrid1.Columns(i + 2).Width - 35
        Next i
    End If

    'Los ponemos Visibles o No
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = visible
    Next i
    cmdAux.visible = visible
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacenes Propios
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
            'Llamamos desde el boton auxiliar de Artículos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 4
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(Indice - 2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVarN_DatoSeleccionado(CadenaSeleccion As String)
    Movimiento = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If (Modo = 2 Or Modo = 0) And Index <> 3 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0, 1 'Codigo Almacen Origen/Destino
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
        Case 2  'Cod. Trabajador
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(4)
            Set frmT = Nothing
        Case 3 ' observaciones
            If Modo = 5 Or Modo = 0 Then
            
            Else
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(5).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then
                        CadenaDesdeOtroForm = DBLet(Data1.Recordset!observa1, "T")
                    End If
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then
                       Text1(5).Text = Mid(CadenaDesdeOtroForm, 3)
                    End If
                End If
                CadenaDesdeOtroForm = ""
            End If
    End Select
    
    If Index = 3 Then
        PonerFoco Text1(5)
    Else
        PonerFoco Text1(Index + 2)
    End If
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
    If Modo = 5 Then   'Eliminar lineas Traspaso Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Traspaso Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then  'Modificar lineas Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificarLinea
    Else 'Modificar Cabecera Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then  'Añadir lineas Traspaso Almacenes
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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index <> 5 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 5 Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 5 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1: KEYFecha2 KeyAscii, 0 ' fecha
                Case 2: KEYBusqueda KeyAscii, 0 'almacen
                Case 3: KEYBusqueda KeyAscii, 1 'almacen
                Case 4: KEYBusqueda KeyAscii, 2 'trabajador
            End Select
        Else
            If Index <> 5 Then KEYpress KeyAscii
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
        Case 0 'Codigo Traspaso Almacen
            If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000000")
        Case 1 'Fecha
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoFecha Text1(Index)
        Case 2, 3 'Codigo Almacen Origen/Destino
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
                'no existe el almacen
                If Text2(Index - 2).Text = "" Then PonerFoco Text1(Index)
            Else
                Text2(Index - 2).Text = ""
                
            End If
        Case 4  'Codigo Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index - 2).Text = ""
            End If
        Case 5 'Observaciones
            'If Text1(Index).Text <> "" Then Text1(Index).Text = QuitarCaracterEnter(Text1(Index).Text)
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
        Case 5
            BotonCopiarLineas
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
Dim devuelve As String
    
    'Quitar espacios en blanco por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    
    Select Case Index
        Case 0 'Cod. Articulo
            If txtAux(Index).Text = "" Then
                txtAux(Index + 1).Text = ""
            ElseIf ModificaLineas = 1 Then 'Insertando linea
                'Comprobamos si ya existe una linea con el artículo, solo si estamos insertando (ModificaLineas=1)
                'conAri: conexion a BD Ariges
                devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "codtrasp", "codtrasp", Text1(0).Text, "N", , "codartic", txtAux(0).Text, "T")
                If devuelve <> "" Then
                    devuelve = "Ya hay una línea con ese Artículo: " & vbCrLf
                    devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
                    MsgBox devuelve, vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    PonerArticulo txtAux(0), txtAux(1), Text1(2).Text, CodTipoMov, ModificaLineas
                End If
            End If
            
        Case 2 'Cantidad (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            If txtAux(Index) <> "" Then PonerFormatoDecimal txtAux(Index), 1
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
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    lblIndicador.Caption = ""
    PonerIndicador lblIndicador, Modo

    'Modo 2. Hay datos y estamos visualizandolos
    '-------------------------------------------
    B = (Kmodo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    'DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    cmdRegresar.visible = False
              
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), (Modo <> 1), True
    
    'Modo 1:Busqueda / Modo 3: Insertar / Modo 4: Modificar
    '-------------------------------------------------------
    'b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        If i <> 3 Then Me.imgBuscar(i).Enabled = B
    Next i
    
    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        imgBuscar(2).Enabled = Modo = 1
        BloquearTxt Text1(4), Modo <> 1
    End If
    
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
Dim i As Byte
Dim bAux As Boolean

    
    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For i = 1 To 3
        Toolbar1.Buttons(i).Enabled = Not EsHistorico
    Next i
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    If Not EsHistorico Then
         B = (Modo = 2) Or (Modo = 5)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (B Or Modo = 0)
        Me.mnNuevo.Enabled = (B Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
        
        '--------------------------------
        B = (Modo = 2)
        'Actualizar
        Toolbar5.Buttons(1).Enabled = B
        'Imprimir
        Toolbar1.Buttons(8).Enabled = B
            
        '-------------------------------
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'VerTodos
        Toolbar1.Buttons(6).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
    Else
        'Actualizar
        FrameBotonGnral2.Enabled = False
        FrameBotonGnral2.visible = False
        FrameDesplazamiento.Left = FrameBotonGnral2.Left
    End If
    
    B = (Modo = 2) And Not EsHistorico
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        bAux = (B And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    
        ToolAux(i).Buttons(5).Enabled = B
    Next i
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    Me.chkImpresion.Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones(Flechas) de Desplazamiento de Registros de la Toolbar
'    Select Case Modo
'        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
'            If Data2.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data2, Index
'        Case Else 'Datos de Cabecera
'            If Data1.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data1, Index
'            PonerCampos
'    End Select
    
    DesplazamientoData Data1, Index, True
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

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
Dim Sql As String
Dim tabla As String
On Error GoTo EMontaSQL
 
    tabla = NomTablaLineas

    Sql = "SELECT " & tabla & ".codtrasp, "
    Sql = Sql & tabla & ".numlinea, " & tabla & ".codartic, Articulos.nomartic, "
    Sql = Sql & tabla & ".cantidad, " & tabla & ".observa2 , "
    
    
    If EsHistorico Then
        Sql = Sql & " preciomphco premp, preciopvphco preve ,preciouchco preuc"
    Else
        Sql = Sql & " preciomp premp, preciove preve ,preciouc preuc"
    End If
    
    Sql = Sql & " FROM ((" & tabla & " LEFT JOIN sartic AS Articulos ON " & tabla & ".codartic ="
    Sql = Sql & " Articulos.codartic))"
    If enlaza Then
        Sql = Sql & ObtenerWhereCP(True)  '" WHERE codtrasp = " & Data1.Recordset!codtrasp
    Else
        Sql = Sql & " WHERE false "
    End If
    Sql = Sql & " ORDER BY " & tabla & ".numlinea"
    MontaSQLCarga = Sql
    
EMontaSQL:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonBuscar()
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
    CargaGrid True
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
    CargaGrid False
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado
    Text1(4).Text = PonerTrabajadorConectado(NomTraba)
    Text2(2).Text = NomTraba
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    PonerModo 5
    
    vWhere = ObtenerWhereCP(False)
    cmdAceptar.Tag = SugerirCodigoSiguienteStr("slitra", "numlinea", vWhere)
    
'    PonerBotonCabecera False
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
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub

Private Sub BotonModificarLinea()
Dim i As Integer

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
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    cmdAceptar.Tag = Data2.Recordset!numlinea

    CargaTxtAux True, False
    PonerFoco txtAux(2) 'Poner el foco
    Screen.MousePointer = vbDefault
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Sql = "Cabecera de Traspaso Almacenes." & vbCrLf
    Sql = Sql & "------------------------------------------" & vbCrLf & vbCrLf
    
    Sql = Sql & "Va a eliminar el Traspaso:" & vbCrLf
    Sql = Sql & vbCrLf & "Nº Traspaso   : " & Text1(0).Text
    Sql = Sql & vbCrLf & "Fecha Trasp.  : " & CStr(Data1.Recordset.Fields(1))
    Sql = Sql & vbCrLf & "Almac. Origen : " & Text1(2).Text
    Sql = Sql & vbCrLf & "Almac. Destino: " & Text1(3).Text
    Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
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
            CargaGrid False
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
Dim Sql As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
On Error GoTo FinEliminar
    
    conn.BeginTrans
    Sql = ObtenerWhereCP(True)  '" WHERE  codtrasp=" & Data1.Recordset!codtrasp
    
    'Lineas
    conn.Execute "Delete  from " & NomTablaLineas & Sql
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & Sql
                      
                      
                      
                      
                            'Borramos el pedido vinculado al antiguo traspaso
    If vParamAplic.NumeroInstalacion = vbHerbelca Then BorrarPedidoVinculado_
                      
                      
                      
                      
                      
  'Devolvemos contador, si no estamos actualizando
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador CodTipoMov, Data1.Recordset.Fields(0)
    Set vTipoMov = Nothing
                      
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
Dim Sql As String
On Error GoTo Error2
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    Sql = "Seguro que desea eliminar la línea del Artículo:"
    Sql = Sql & vbCrLf & "Código: " & Data2.Recordset!codArtic
    Sql = Sql & vbCrLf & "Descripción: " & Data2.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        
        
        
        NumRegElim = Me.Data2.Recordset.AbsolutePosition
        Sql = "Delete from slitra where codtrasp=" & Data2.Recordset!codtrasp
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        Sql = Sql & " and codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        
        If vParamAplic.NumeroInstalacion = vbHerbelca Then LineaPedidoVinculado ModificaLineas, Data2.Recordset!numlinea
        
        conn.Execute Sql
        CancelaADODC Me.Data2
        CargaGrid True
        CancelaADODC Me.Data2
        
        
        
        
        
        SituarDataPosicion Me.Data2, NumRegElim, Sql
        lblIndicador.Caption = Sql
    End If
    ModificaLineas = 0
Error2:
        Screen.MousePointer = vbDefault
        ModificaLineas = 0
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Línea de Artículo de Traspaso Almacenes", Err.Description
End Sub

Private Sub BotonCopiarLineas()
Dim Sql As String
Dim Sql2 As String
Dim vCadena As String


    Sql = "select codtrasp, fechatra, hormovim, schtra.almaorig, salmpr.nomalmac nomalmori, schtra.almadest, salmpr.nomalmac nomalmdes, schtra.codtraba, straba.nomtraba "
    Sql = Sql & " from ((schtra inner join straba on schtra.codtraba = straba.codtraba) "
    Sql = Sql & " inner join salmpr on schtra.almaorig = salmpr.codalmac) "
    Sql = Sql & " inner join salmpr dest on schtra.almadest = dest.codalmac "
    
    Sql2 = " where fechatra >= " & DBSet(DateAdd("m", -1, Now), "F")

    If TotalRegistros(Sql & Sql2) = 0 Then
        If TotalRegistros(Sql) = 0 Then
            MsgBox "No hay traspasos de almacén en el histórico"
            Exit Sub
        Else
            vCadena = Sql & "||0|"
        End If
    Else
        vCadena = Sql & "|" & Sql2 & "|1|"
    End If

    Movimiento = ""

    Set frmVarN = New frmVariosNew
    frmVarN.CADENA = vCadena
    frmVarN.Opcion = 100
    frmVarN.Show vbModal
    
    Set frmVarN = Nothing
    
    If Movimiento <> "" Then
        If CopiarMovimientos(Movimiento) Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            CargaGrid True
        Else
            MsgBox "No se ha realizado es proceso", vbExclamation
        End If
    End If
End Sub

Private Function CopiarMovimientos(movim As String) As Boolean
Dim Sql As String
Dim vResult As String
Dim vResult2 As String
Dim RS As ADODB.Recordset
Dim numlin As String

    On Error GoTo eCopiarMovimientos

    CopiarMovimientos = False

    Sql = "select slitra.codartic from slhtra inner join slitra on slhtra.codartic = slitra.codartic where slhtra.codtrasp = " & DBSet(movim, "N")
    Sql = Sql & " and slitra.codtrasp = " & DBSet(Text1(0).Text, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vResult = ""
    vResult2 = ""
    While Not RS.EOF
        vResult = vResult & ", " & DBLet(RS!codArtic, "T")
        vResult2 = vResult2 & "," & DBSet(RS!codArtic, "T")
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If vResult <> "" Then
        If MsgBox("Los siguientes artículos se encuentran en este traspaso: " & vbCrLf & Mid(vResult, 3) & vbCrLf & " ¿ Desea continuar ? ", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    numlin = DevuelveDesdeBDNew(conAri, "slitra", "max(numlinea)", "codtrasp", Text1(0).Text, "N")
    If ComprobarCero(numlin) = "0" Then numlin = "0"
    Sql = "insert ignore into slitra (codtrasp,numlinea,codartic,cantidad,observa2)"
    Sql = Sql & " select " & DBSet(Text1(0), "N") & "," & "@Lin:=@Lin + 1 ,codartic,cantidad,observa2"
    Sql = Sql & " from slhtra, (select @Lin:= " & numlin & ") aa where codtrasp = " & DBSet(movim, "N")
    
    If vResult <> "" Then Sql = Sql & " and not slhtra.codartic in (" & Mid(vResult2, 2) & ")"
    
    conn.Execute Sql
    
    
    'En herbelca, volvemos a generar las lineas del pedido vinculado
    If vParamAplic.NumeroInstalacion = vbHerbelca Then LineaPedidoVinculado 4, 0
       
    
    
    
    CopiarMovimientos = True
    Exit Function
    
eCopiarMovimientos:
    MuestraError Err.Number, "Copiar Movimientos", Err.Description
End Function


Private Function DatosOk(Optional cabecera As Boolean) As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function

    'Comprobar que almacen origen y destino son distintos
    If Trim(Text1(2).Text) = Trim(Text1(3).Text) Then
        MsgBox "Almacen Origen y Destino no pueden ser el mismo.", vbExclamation
        B = False
        Exit Function
    End If
    
    If Not cabecera Then B = ComprobarStocksLineas
    
    DatosOk = B
End Function



Private Function ComprobarStocksLineas() As Boolean
'Comprobar para todas las lineas del traspaso que:
' - todos los Artículos entan en el almacen origen
' - Comprobar que hay suficiente stock en el Almacen Origen de ese Articulo
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String

    On Error GoTo ErrStock
    
    
    '---- Laura: 27/09/2006
    If Data2 Is Nothing Then Exit Function
    
    Sql = Data2.RecordSource
    If Sql = "" Then Exit Function
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    'para cada linea comprabar stock del articulo en almacen
    B = True
    While Not RS.EOF And B
        B = ComprobarStock(RS!codArtic, Data1.Recordset!almaorig, RS!cantidad, CodTipoMov)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    ComprobarStocksLineas = B
    Exit Function
    '----

    '## ANTES
'    If Not Data2.Recordset.EOF Then  'Si hay lineas
'        Data2.Recordset.MoveFirst
'        b = True
'
'        While Not Data2.Recordset.EOF And b
'            b = ComprobarStock(Data2.Recordset!codArtic, Data1.Recordset!almaorig, Data2.Recordset!Cantidad, CodTipoMov)
'            Data2.Recordset.MoveNext
'        Wend
'    End If
'    ComprobarStocksLineas = b
    '##
    
ErrStock:
    ComprobarStocksLineas = False
    MuestraError Err.Number, "Comprobar stock.", Err.Description
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim Aux As String

    DatosOkLinea = False
    B = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Artículo no puede ser nulo", vbExclamation
        B = False
        Exit Function
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Artículos"
         B = False
    ElseIf Not IsNumeric(txtAux(2).Text) Then
        MsgBox "El campo Cantidad debe ser numérico", vbExclamation
        B = False
    End If
    If Not B Then
        PonerFoco txtAux(2)
        Exit Function
    End If
    
    'b = ComprobarStock(txtAux(0).Text, txtAux(1).Text, txtAux(2).Text, CodTipoMov)
    B = ComprobarStock(txtAux(0).Text, Text1(2).Text, txtAux(2).Text, CodTipoMov)
         
         
     If vParamAplic.ManipuladorFitosanitarios2 Then
        Aux = DevuelveDesdeBD(conAri, "codalmac", "advparametros", "1", "1")
        If Aux = "" Then Aux = "1"
        If Val(Data1.Recordset!almadest) = Val(Aux) Then
            
            'Lleva almacen para fitosinatiarios y es el DESTINO
            
            Aux = " sartic.codcateg=scateg.codcateg and ctrlotes =1 and codartic"
            Aux = DevuelveDesdeBD(conAri, "numserie", "sartic,scateg", Aux, txtAux(0).Text, "T")
            If Aux <> "" Then
                'LLEVA CONTROL DE LOTES
                'La observacion debe ser un LOTE valido
                txtAux(3).Text = Trim(txtAux(3).Text)
                If txtAux(3).Text = "" Then
                    Aux = "Debe indicar el lote para el traspaso de fitosantiarios"
                Else
                    Aux = "numlotes=" & DBSet(txtAux(3).Text, "T") & " AND codartic =" & DBSet(txtAux(0).Text, "T") & " AND 1"
                    
                    Aux = DevuelveDesdeBD(conAri, "canentra - vendida ", "slotes", Aux, "1 ORDER BY (canentra - vendida)>0 desc ,fecentra")
                    If Aux = "" Then
                        Aux = "NO existe el lote " & txtAux(3).Text & " para el articulo: " & txtAux(2).Text
                    Else
                        If CCur(Aux) <= 0 Then
                            Aux = "NO existe cantidad disponible para el lote: " & txtAux(3).Text
                        Else
                            Aux = ""
                        End If
                    End If
                End If
                If Aux <> "" Then
                    MsgBox Aux, vbExclamation
                    B = False
                End If
                      
            End If
        End If
    End If
         
         
    DatosOkLinea = B
End Function


Private Sub PonerBotonCabecera(B As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "LINEAS DETALLE"
    Else
        Me.lblIndicador.Caption = ""
    End If
     'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim Sql As String
Dim TodasLasLineaas As Boolean
On Error GoTo EInsertarModificarLinea
    
    Sql = ""
    InsertarModificarLinea = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea() Then 'INSERTAR
            Sql = "INSERT INTO slitra (codtrasp,numlinea,codartic,cantidad,observa2) "
            Sql = Sql & " VALUES (" & Val(Text1(0).Text) & ", "
            Sql = Sql & cmdAceptar.Tag & ", "
            Sql = Sql & DBSet(txtAux(0).Text, "T") & ", "
            Sql = Sql & DBSet(txtAux(2).Text, "N") & ","
            Sql = Sql & DBSet(txtAux(3).Text, "T") & ") "
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLinea() Then
            Sql = "UPDATE slitra Set cantidad = " & DBSet(txtAux(2).Text, "N")
            Sql = Sql & ", observa2 = " & DBSet(txtAux(3).Text, "T")
            Sql = Sql & ObtenerWhereCP(True) & " AND " '" WHERE codtrasp =" & Val(Text1(0).Text) & " AND "
            Sql = Sql & " numlinea =" & cmdAceptar.Tag
        End If
    End Select
            
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarLinea = True
        
        'Si tiene componentes preguntamos si queire insertar las lineas
        TodasLasLineaas = False
        Sql = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codartic", txtAux(0).Text, "T")
        If Val(Sql) >= 1 Then
            If MsgBox("El articulo tiene componentes" & vbCrLf & "¿Desea insertarlos?", vbQuestion + vbYesNoCancel) = vbYes Then
                Sql = "select " & Val(Text1(0).Text) & "," & cmdAceptar.Tag & "+numlinea,codarti1,"
                Sql = Sql & " (cantidad*" & DBSet(txtAux(2).Text, "N") & "),concat('COMPONENTES ',codartic) from sarti1 where codartic=" & DBSet(txtAux(0).Text, "T")
                Sql = "INSERT INTO slitra (codtrasp,numlinea,codartic,cantidad,observa2) " & Sql
                ejecutar Sql, False
                Espera 0.2
                TodasLasLineaas = True
            End If
        End If
            
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            If TodasLasLineaas Then
                'Rehacer pedido
                LineaPedidoVinculado 4, 0
            Else
                 LineaPedidoVinculado ModificaLineas, Me.cmdAceptar.Tag
            End If
         End If
    
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
'    If Modo <> 5 Then 'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: scatra
'        Cad = Cad & ParaGrid(Text1(0), 12, "Nº Trasp.")
'        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha")
'        Cad = Cad & ParaGrid(Text1(2), 7, "Orig.")
'        Cad = Cad & "Desc. Alm. Orig|salmpr|nomalmac|T||30·"
'        Cad = Cad & ParaGrid(Text1(3), 7, "Dest.")
'        Cad = Cad & "Alm. Dest|AlmDestino|nomalmac as almdest|T||29·"
'
'        Tabla = "(" & NombreTabla & " LEFT JOIN salmpr ON " & NombreTabla & ".almaorig=salmpr.codalmac" & ")"
'        Tabla = Tabla & " LEFT JOIN salmpr AS AlmDestino ON " & NombreTabla & ".almadest=AlmDestino.codalmac "
'        'tabla = tabla & NombreTabla & ".coddirec=sdirec.coddirec"
'
'        ' tabla = "scatra"
'        Titulo = Me.Caption
'    Else 'Estamos en modo Lineas
'        Cad = Cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
'        Tabla = "sartic"
'        Titulo = "Articulos"
'    End If
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = Tabla
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
''        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            PonerFoco Text1(kCampo)
''        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmB = New frmBasico2
    AyudaAlmMovTraspasoPrev frmB, EsHistorico
    Set frmB = Nothing

End Sub

Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
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
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac", "codalmac")
    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "salmpr", "nomalmac", "codalmac")
    Text2(2).Text = PonerNombreDeCod(Text1(4), conAri, "straba", "nomtraba")
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim Sql As String
Dim cantidad As Single
Dim devuelve As String
Dim RS As ADODB.Recordset

    On Error GoTo EActualizarStock

    ActualizarStocks = False
    
    '---- Laura: 27/09/2006
    'sustituir el data2 por el RecordSEt
    Set RS = New ADODB.Recordset
    RS.Open Data2.RecordSource, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'    While Not Data2.Recordset.EOF

        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codArtic, "T")
        If Val(devuelve) = 1 Then

            cantidad = CSng(RS!cantidad) 'Cant a traspasar
            
            '==== Almacen Origen
            Sql = "UPDATE salmac Set canstock = canstock - " & DBSet(cantidad, "N")
            Sql = Sql & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
            Sql = Sql & " codalmac =" & Data1.Recordset!almaorig
            conn.Execute Sql
        
            '==== Almacen Destino
            'Comprobar que existe el articulo en Almacen Destino
            devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codalmac", "codartic", RS!codArtic, "T", , "codalmac", Text1(3).Text, "N")
            If devuelve = "" Then 'No hay de ese artículo en Destino
                Sql = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                Sql = Sql & " VALUES (" & DBSet(RS!codArtic, "T") & "," & Val(Text1(3).Text) & ",''," & DBSet(cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
            Else 'Existe el artic en almac. Dest -> Aumentar stock
                Sql = "UPDATE salmac Set canstock = canstock + " & DBSet(cantidad, "N")
                Sql = Sql & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
                Sql = Sql & " codalmac =" & Data1.Recordset!almadest
            End If
            
            conn.Execute Sql
        End If
        'Data2.Recordset.MoveNext
        RS.MoveNext
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
Dim Sql As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún Traspaso para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
    
    If Data2.Recordset.EOF Then
        MsgBox "No hay lineas insertadas para este Nº de Traspaso", vbExclamation
        Exit Sub
    End If
    
    If Data1.Recordset!situacio = 0 Then 'Informe No Impreso
        Sql = "Actualización Traspaso Almacenes." & vbCrLf
        Sql = Sql & "------------------------------------------" & vbCrLf & vbCrLf
        Sql = Sql & "NO ESTA IMPRESO EL TRASPASO:" & vbCrLf
        Sql = Sql & vbCrLf & "Nº Trasp.     :  " & Format(Data1.Recordset.Fields(0), "0000000")
        Sql = Sql & vbCrLf & "Fecha Trasp.  :  " & CStr(Data1.Recordset.Fields(1))
        Sql = Sql & vbCrLf & "Almac. Origen :  " & Format(Data1.Recordset.Fields(2), "000") & " - " & Text2(0).Text & "     "
        Sql = Sql & vbCrLf & "Almac. Destino:  " & Format(Data1.Recordset.Fields(3), "000") & " - " & Text2(1).Text & "     "
        Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    Else 'Informe Impreso
        Sql = "Actualización Traspaso Almacenes." & vbCrLf
        Sql = Sql & "-----------------------------------------" & vbCrLf & vbCrLf
        Sql = Sql & "Va a Actualizar el Traspaso:" & vbCrLf
        Sql = Sql & vbCrLf & "   Nº Trasp.   : " & Format(Data1.Recordset.Fields(0), "0000000")
        Sql = Sql & vbCrLf & "   Fecha Trasp.: " & CStr(Data1.Recordset.Fields(1))
        Sql = Sql & vbCrLf & "   Almac. Orig.: " & Format(Data1.Recordset.Fields(2), "000") & " - " & Text2(0).Text & "     "
        Sql = Sql & vbCrLf & "   Almac. Dest.: " & Format(Data1.Recordset.Fields(3), "000") & " - " & Text2(1).Text & "     "
        Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If
    
    
    'bloqueamos el registro que vamos a traspasar
    If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
    
    
    'realizamos el traspaso de almacenes
    Me.ProgressBar1.visible = True
    NumRegElim = Data1.Recordset.AbsolutePosition
    
    If ActualizarTraspaso Then
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            Me.Refresh
        End If
    End If
    Me.ProgressBar1.visible = False
    TerminaBloquear
End Sub


Private Function ActualizarTraspaso() As Boolean
Dim Donde As String
Dim devuelve As String
Dim bol As Boolean

    On Error GoTo EActualizarTraspaso
    
    'Comprobamos que no existe en historico de traspasos
    devuelve = DevuelveDesdeBDNew(conAri, "schtra", "codtrasp", "codtrasp", Data1.Recordset!codtrasp, "N", , "fechatra", Data1.Recordset!fechatra, "F")
    If Trim(devuelve <> "") Then
        devuelve = "Ya existe en el histórico el traspaso:" & vbCrLf
        devuelve = devuelve & " Nº: " & Data1.Recordset!codtrasp & vbCrLf
        devuelve = devuelve & " Fecha: " & Data1.Recordset!fechatra
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    'Comprobar que en almacen origen existe la cantidad a traspasar
    'de cada linea de articulo
    If Not ComprobarStocksLineas Then Exit Function
    
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarElTraspaso(Donde)

EActualizarTraspaso:
        If Err.Number <> 0 Then
'            devuelve = "Actualizar Traspaso." & vbCrLf & "----------------------------" & vbCrLf
'            devuelve = devuelve & Donde
'            MuestraError Err.Number, devuelve, Err.Description
            devuelve = Err.Description & ": " & Err.Number
            bol = False
        Else
            devuelve = ""
        End If
        If bol Then
            conn.CommitTrans
            ActualizarTraspaso = True
        Else
            conn.RollbackTrans
            devuelve = "Actualizar Traspaso." & vbCrLf & "----------------------------" & vbCrLf & "ERROR: " & Donde & vbCrLf & devuelve
            MsgBox devuelve, vbExclamation
        End If
End Function


Private Function ActualizarElTraspaso(ByRef ADonde As String) As Boolean
Dim cadError As String

    ActualizarElTraspaso = False
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en historico cabeceras traspaso almacenes"
    If Not InsertarCabeceraHistorico Then Exit Function
    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Traspaso Almacenes"
    If Not InsertarLineasHistorico(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2
    
    'Modificar stock (Tabla: salmac)
    ADonde = "Actualizando Stocks Almacenes"
    If Not ActualizarStocks() Then Exit Function
    IncrementarProgres 2
    
    'Insertamos en Movimientos Artículos (Tabla: smoval)
    ADonde = "Insertando datos en Movimientos de Articulos"
    If Not InsertarMovimArticulos(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2

    
    'Borramos cabeceras y lineas del TRaspaso
    ADonde = "Borrar cabeceras y lineas en Traspaso Almacenes"
    If Not BorrarTraspaso(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2
    
    
    If vParamAplic.NumeroInstalacion = vbHerbelca Then BorrarPedidoVinculado_
    
    
    
    ActualizarElTraspaso = True
End Function


Private Function InsertarCabeceraHistorico() As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarCab

    Sql = "SELECT codtrasp,fechatra,almaorig,almadest,codtraba,observa1,codclienvinculado,codpedidovinuclado from scatra "
    Sql = Sql & ObtenerWhereCP(True)
    Sql = Sql & " AND fechatra='" & Format(Data1.Recordset!fechatra, "yyyy-mm-dd") & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        Sql = "INSERT INTO schtra (codtrasp, fechatra,hormovim,almaorig,almadest,codtraba,observa1,codclienvinculado,codpedidovinuclado) "
        Sql = Sql & " VALUES (" & RS.Fields(0).Value & ", '" & Format(RS.Fields(1).Value, "yyyy-mm-dd") & "', '"
        Sql = Sql & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields(2).Value & ", " & RS.Fields(3).Value & ", "
        Sql = Sql & RS.Fields(4).Value & ", " & DBSet(RS.Fields(5).Value, "T") & ","
        Sql = Sql & DBSet(RS.Fields(6).Value, "N", "S") & ", " & DBSet(RS.Fields(7).Value, "N", "S") & ")"
    End If
    RS.Close
    Set RS = Nothing
    
    conn.Execute Sql
    
EInsertarCab:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(MenError As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarLineas

    'SQL = "SELECT codtrasp, numlinea, codartic, cantidad, observa2 from slitra "
    'Diciembre 2021
    Sql = "SELECT codtrasp, numlinea, slitra.codartic, cantidad, observa2,preciomp,preciouc,preciove "
    Sql = Sql & " FROM slitra LEFT JOIN  sartic ON slitra.codartic=sartic.codartic"
    
    Sql = Sql & ObtenerWhereCP(True)
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    RS.MoveFirst
    While Not RS.EOF
        Sql = "INSERT INTO slhtra (codtrasp, fechamov, numlinea, codartic, cantidad, observa2, preciomphco ,preciopvphco ,preciouchco)"
        Sql = Sql & " VALUES (" & RS.Fields(0).Value & ", '" & Format(Data1.Recordset!fechatra, FormatoFecha) & "', "
        Sql = Sql & RS.Fields(1).Value & ", " & DBSet(RS.Fields(2).Value, "T") & ", "
        Sql = Sql & DBSet(RS.Fields(3).Value, "N") & ", " & DBSet(RS.Fields(4).Value, "T") & ","
        'Dicimebre 2021
        Sql = Sql & DBSet(RS!PrecioMP, "N") & ", " & DBSet(RS!PrecioVe, "N") & ", " & DBSet(RS!precioUC, "N") & ")"
        conn.Execute Sql
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


Private Sub IncrementarProgres(Veces As Integer)
On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * 10)
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub


Private Function BorrarTraspaso(MenError As String) As Boolean
Dim Sql As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    Sql = "Delete from "
    Sql = Sql & "slitra"
    Sql = Sql & " WHERE codtrasp = " & Data1.Recordset!codtrasp
    conn.Execute Sql
    
    'La cabecera
    Sql = "Delete from "
    Sql = Sql & "scatra"
    Sql = Sql & " WHERE codtrasp =" & Data1.Recordset!codtrasp
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
        MenError = Err.Number & ": " & Err.Description
    Else
        BorrarTraspaso = True
        MenError = ""
    End If
End Function

Public Sub ActualizarSituacionImpresion()
Dim Cad As String
Dim Indicador As String
On Error GoTo EImpresion
   
        Cad = "(codtrasp=" & Val(Text1(0).Text) & ")"
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


Private Function InsertarMovimArticulos(MenError As String) As Boolean
Dim Sql As String, Cad As String
Dim RS As ADODB.Recordset
Dim vImporte As Single, vPrecioVenta As String
Dim vTipoMov As CTiposMov
Dim bol As Boolean
    
    On Error GoTo EInsertar

    bol = True
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        'Se han cargado correctamente los valores de la clase
        Sql = "SELECT scatra.codtrasp, fechatra, almaorig, almadest, codtraba, numlinea, codartic, cantidad "
        Sql = Sql & " FROM scatra LEFT JOIN slitra ON scatra.codtrasp=slitra.codtrasp "
        Sql = Sql & " WHERE scatra.codtrasp =" & Data1.Recordset!codtrasp
    
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
             'Obtener el precio de venta del articulo, si tiene control de stock
            Cad = "ctrstock"
            vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", RS.Fields!codArtic, "T", Cad)
            If vPrecioVenta <> "" Then
                vImporte = Round2(RS.Fields!cantidad * CSng(vPrecioVenta), 2)
            Else
                vImporte = 0
            End If
            If Val(Cad) = 1 Then
                'Insertar Movimiento de Salida en Almacen Origen
                Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                Sql = Sql & " VALUES (" & DBSet(RS.Fields!codArtic, "T") & ", " & RS.Fields!almaorig & ", '" & Format(RS.Fields!fechatra, "yyyy-mm-dd") & "', '"
                Sql = Sql & Format(RS.Fields!fechatra & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', 0" & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(RS.Fields!cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & RS.Fields!CodTraba & ", "
                Sql = Sql & DBSet(vTipoMov.LetraSerie, "T") & ", " & RS.Fields!codtrasp & ", " & RS.Fields!numlinea & ")"
                conn.Execute Sql
                
                'Insertar Movimiento de Entrada en Almacen Destino
                Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                Sql = Sql & " VALUES (" & DBSet(RS.Fields!codArtic, "T") & ", " & RS.Fields!almadest & ", '" & Format(RS.Fields!fechatra, "yyyy-mm-dd") & "', '"
                Sql = Sql & Format(RS.Fields!fechatra & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', 1" & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(RS.Fields!cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & RS.Fields!CodTraba & ", "
                Sql = Sql & DBSet(vTipoMov.LetraSerie, "T") & ", " & RS.Fields!codtrasp & ", " & RS.Fields!numlinea & ")"
                conn.Execute Sql
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
        MenError = Err.Number & ": " & Err.Description
    End If
    
    If Err.Number <> 0 Or Not bol Then
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
        MenError = ""
    End If
End Function


Private Sub BotonImprimir()
    If Text1(0).Text = "" Then Exit Sub
    frmListado.NumCod = Text1(0).Text
    
'    If Not EsHistorico Then
'        AbrirListado (7) '7: Informe Traspaso de Almacen
        frmInformesNew.NumCod = Text1(0).Text
        frmInformesNew.EsHco = EsHistorico
        frmInformesNew.OpcionListado = 7
        frmInformesNew.Show vbModal
        If Not EsHistorico Then ActualizarSituacionImpresion
'
'        ActualizarSituacionImpresion
'    Else
'        BotonImprimirHco
'    End If
End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim Cad As String
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
                Cad = "{schtra.codtrasp}= " & Data1.Recordset!codtrasp
                Cad = Cad & " and {schtra.fechatra}= Date(" & Year(Data1.Recordset!fechatra) & "," & Month(Data1.Recordset!fechatra) & "," & Day(Data1.Recordset!fechatra) & ")" & ""
                .FormulaSeleccion = Cad
            End If
            .Show vbModal
        End With
    End If
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function InsertarTraspaso(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean
On Error GoTo EInsertarMovimiento
    
    bol = True
    InsertarTraspaso = False
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla de Traspasos(scatra)."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del recibo."
    vTipoMov.IncrementarContador (CodTipoMov)
    

EInsertarMovimiento:
        If Err.Number <> 0 Then
            MenError = "Insertando Traspaso." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarTraspaso = True
        Else
            conn.RollbackTrans
            InsertarTraspaso = False
        End If
End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
On Error Resume Next
    If conWhere Then
        ObtenerWhereCP = " WHERE codtrasp= " & Val(Text1(0).Text)
    Else
        ObtenerWhereCP = " codtrasp= " & Val(Text1(0).Text)
    End If
End Function


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

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


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarTraspaso(Sql, vTipoMov) Then
                
                
                'Si es herbelca, creará el pedido vinculado
                CrearPedidoVinculado
            
            
            
            
            
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


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid False
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub CrearPedidoVinculado()
Dim Aux As String
Dim CC As CTiposMov

    
    Aux = DevuelveDesdeBD(conAri, "clientevinculadoherbelca", "salmpr", "codalmac", Text1(3).Text)
    If Aux = "" Then Exit Sub 'El almacen NO lleva cliente vinculado para la generacion de pedido
    
    
    'SI que lleva,
    NumRegElim = Val(Aux)
    Set CC = New CTiposMov
    If CC.Leer("PEV") Then
        Aux = ""
        If Modo = 2 Then
            If DBLet(Data1.Recordset!codpedidovinuclado, "N") > 0 Then Aux = Data1.Recordset!codpedidovinuclado
        End If
        If Aux = "" Then
            CC.ConseguirContador CC.TipoMovimiento
            CC.IncrementarContador CC.TipoMovimiento
        Else
            CC.Contador = Val(Aux)
        End If
        
        Aux = "INSERT INTO scaped(numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,"
        Aux = Aux & " nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Aux = Aux & " referenc, codtraba,codagent,codforpa,dtoppago,dtognral,tipofact,"
        Aux = Aux & " observa01,servcomp,restoped,recogecl,mailconfir,observaciones,cerrado)"
        Aux = Aux & " select  " & CC.Contador & " as numpedcl, fechatra,fechatra,week(fechatra,3),1, "
        Aux = Aux & " sclien.codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclie1"
        Aux = Aux & " ,concat('TRASPASO ', right(concat('0000000',codtrasp),7)) as referenc , codtraba,codagent,codforpa"
        Aux = Aux & " ,0 dtoppago,0 dtognral,0 tipofact,concat('TRASPASO ', right(concat('0000000',codtrasp),7)) observa01,"
        Aux = Aux & " 1 servcomp,0 restoped, 0 recogecl"
        Aux = Aux & " , concat('TRASPASO ', right(concat('0000000',codtrasp),7)) mailconfir ,observa1 observaciones, 0 cerrado"
        'Aux = Aux & " From scatra inner join salmprTrasCli on scatra.almadest=salmprTrasCli.codalmac"
        'Aux = Aux & " inner join sclien on sclien.codclien=" & NumRegElim & " WHERE scatra.codtrasp = " & Text1(0).Text
        Aux = Aux & " From scatra inner join sclien on sclien.codclien=" & NumRegElim & " WHERE scatra.codtrasp = " & Text1(0).Text
        
        
        
        If ejecutar(Aux, False) Then
            Aux = "UPDATE scatra SET codclienvinculado =" & NumRegElim
            Aux = Aux & " , codpedidovinuclado = " & CC.Contador & " WHERE codtrasp =" & Text1(0).Text
            ejecutar Aux, False
            Espera 0 - 25
        End If
        
        
    End If
    Set CC = Nothing
End Sub



Private Sub BorrarPedidoVinculado_()
Dim Aux As String

    
    Aux = DBLet(Data1.Recordset!codpedidovinuclado, "T")
    If Aux = "" Then Exit Sub
    ejecutar "DElete from sliped where numpedcl=" & Aux, False
    ejecutar "DElete from scaped where numpedcl=" & Aux, False
    
End Sub


'ACCION: 1 insertar line=linea
'        2 insline =linea
'       3 borrar
'       LineaPedidoVinculado
'       4. Borrar todos las lineas del pedido Volverlas a insertar
Private Sub LineaPedidoVinculado(Accion As Byte, linea As Integer)
Dim NumPedcl As String
Dim Aux As String
    
    
    If DBLet(Data1.Recordset!codpedidovinuclado, "N") = 0 Then Exit Sub
    
    
    'NumPedcl = "(" & data1.Recordset!almadest & " * 1000000) + " & data1.Recordset!Codtrasp
    NumPedcl = Data1.Recordset!codpedidovinuclado
        
    
    
    Aux = "DELETE FROM sliped WHERE numpedcl= " & Data1.Recordset!codpedidovinuclado
    If Accion < 4 Then Aux = Aux & " AND numlinea = " & linea
    ejecutar Aux, False
    
    If Accion = 3 Then Exit Sub 'Solo queria Borrar
    Aux = "insert into sliped(numpedcl,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,servidas,numbultos,bultosser,precioar,dtoline1,dtoline2,importel,origpre)"
    Aux = Aux & " select  " & NumPedcl & " as numpedcl,numlinea,scatra.almadest ,slitra.codartic,nomartic,observa2,cantidad,0 servidas,0 numbulto,0 butlosser,preciouc,0 dot1,0 dto2,"
    Aux = Aux & " round(preciouc* cantidad,2),'M'  From slitra inner join sartic on slitra.codartic=sartic.codartic"
    Aux = Aux & " inner join scatra on slitra.codtrasp=scatra.codtrasp "
    Aux = Aux & " WHERE slitra.Codtrasp= " & Data1.Recordset!codtrasp
    If Accion < 4 Then Aux = Aux & " AND numlinea = " & linea
    ejecutar Aux, False




End Sub
