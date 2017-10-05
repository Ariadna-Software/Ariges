VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProdEnvas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de envasado"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data4 
      Height          =   495
      Left            =   7560
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
   Begin VB.Frame FrameCalidad 
      Height          =   5655
      Left            =   120
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdAux2 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         ToolTipText     =   "Buscar artículo"
         Top             =   4080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtCalidad 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   6480
         TabIndex        =   18
         Text            =   "result"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboCalidad 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtCalidad 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   17
         Text            =   "espec"
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCalidad 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   48
         Text            =   "nomar"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkCalidad 
         BackColor       =   &H8000000E&
         Caption         =   "Si"
         Height          =   255
         Left            =   9120
         TabIndex        =   19
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCalidad 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Text            =   "codartic"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   5025
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8864
         _Version        =   393216
         AllowUpdate     =   0   'False
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
               ColumnAllowSizing=   -1  'True
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   45
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   44
      Top             =   1800
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdAux2 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   8880
      TabIndex        =   41
      ToolTipText     =   "Buscar artículo"
      Top             =   5880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux2 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   42
      ToolTipText     =   "Buscar artículo"
      Top             =   5880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtComponentes 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   5880
      Width           =   3615
   End
   Begin VB.TextBox txtComponentes 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtComponentes 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txtComponentes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   10560
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txtComponentes 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   9000
      TabIndex        =   40
      Text            =   "Text2"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   8640
      MaxLength       =   16
      TabIndex        =   8
      Tag             =   "Cantidad"
      Text            =   "1,234,567,891.25"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtComponentes 
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   5880
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   5520
      Top             =   8040
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2295
      Left            =   120
      TabIndex        =   37
      Top             =   5400
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "Código Almacen"
      Text            =   "codalmac"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1080
      MaxLength       =   18
      TabIndex        =   6
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   6120
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Cantidad"
      Text            =   "1,234,567,891.25"
      Top             =   3180
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   34
      Tag             =   "Nombre Artículo"
      Text            =   "nomArtic"
      Top             =   3180
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   33
      ToolTipText     =   "Buscar almacen"
      Top             =   3180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   32
      ToolTipText     =   "Buscar artículo"
      Top             =   3180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   410
      Width           =   11295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Nº Pedido|N|S|||senvprod|numpedcl|00000000|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   1035
         Index           =   3
         Left            =   7320
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Obs|T|S|||senvprod|descripcion|||"
         Top             =   165
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha creación|F|N|||senvprod|feccreacion|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº ord produccion|N|S|0||senvprod|codigo|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha producción|F|S|||senvprod|fecproduccion|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5280
         ToolTipText     =   "Buscar Nº Serie"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   36
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   35
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "F. creacion"
         Height          =   255
         Index           =   14
         Left            =   1590
         TabIndex        =   30
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2520
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   29
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. produccion"
         Height          =   255
         Index           =   51
         Left            =   3120
         TabIndex        =   28
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   7935
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10170
      TabIndex        =   21
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Top             =   8040
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5520
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
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
            Object.ToolTipText     =   "Calidad"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas produccion"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar sublineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar orden produccion"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Orden Instal."
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6480
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   5040
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10170
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2280
      Left            =   1320
      TabIndex        =   31
      Top             =   2520
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4022
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
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      Height          =   1095
      Index           =   1
      Left            =   4560
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   2040
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   120
      X2              =   11280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Envasado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   2400
      TabIndex        =   43
      Top             =   7920
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Articulos envasado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Componentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   120
      TabIndex        =   38
      Top             =   5040
      Width           =   1530
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
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnGeneraFactura 
         Caption         =   "Generar factura"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpPedido 
            Caption         =   "&Pedido"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnImpOrde 
            Caption         =   "&Orden Instalación"
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmProdEnvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1


Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmPe As frmFacEntPedidos
Attribute frmPe.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'   NUEVO 03/06/2009
'          6.-  Modificar cantidad en componentes
'   6.-  Modo SUBLINEAS
'-------------------------------------------------------------------------


Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

'Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec



'SQL de la tabla principal del formulario
Private CadenaConsulta As String


Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1






Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim OpcionConElPedido As Byte
    ' 0. NADA
    ' >1 traer los datos del pedido
    '   =1 AÑAIDR LOS DATOS
    '   =2 borrar los anteriores


Private Sub cboCalidad_Click()
    cboCalidad_LostFocus
    If cboCalidad.ListIndex >= 0 Then PonerFoco txtAux(9)
End Sub

Private Sub cboCalidad_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cboCalidad_LostFocus()

    If Modo = 7 Then
        If ModificaLineas = 1 Then
            If cboCalidad.ListIndex > 0 Then txtCalidad(2).Text = DevuelveDesdeBD(conAri, "especificaciones", "scalidad", "codigo", cboCalidad.ItemData(cboCalidad.ListIndex))
        End If
        
    End If
        
End Sub

Private Sub chkCalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'ACtualizamos y movmemos al siguiente
        If ModificaLineas = 1 Then
            KEYpressGnral KeyAscii, 3, False
        Else
            If ModificarExistencia Then PasarSigReg
        End If
    End If
End Sub

Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid3.Bookmark < data4.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid3.Bookmark = DataGrid3.Bookmark + 1
        ModificarDatosCalidad True
    ElseIf DataGrid3.Bookmark = data4.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub



'================================================================================






Private Sub cmdAceptar_Click()
'Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR Cabecera Pedido
            
            If DatosOk Then InsertarCabecera
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                
                If ModificaDesdeFormulario(Me, 1) Then
                    ActualizarLineasPedido
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    CargaGridCalidad True
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True

                End If
                Me.DataGrid1.Enabled = True
            End If
            
            
        Case 6 'Modif cantidad componentes
            
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarSubLinea Then
                    If PrimeraLin Then
                        'CargaGrid3
                    Else
                        '
                    End If
                    CargaGrid3 True
                    BotonAnyadirSubLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If UpdateaCantidadComponentes Then
                    TerminaBloquear
                    ModificarCantidadComponentes False
                    CargaGrid3 True
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    DataGrid2.Enabled = True
                End If
            End If
            
        Case 7
            
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If data3.Recordset.EOF = True Then PrimeraLin = True
                If InsertarSubLineaCalidad Then
                    CargaGridCalidad True
                    BotonAnyadirSubLineaCalidad
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If UpdateaDatosCalidad Then
                    TerminaBloquear
                    ModificarDatosCalidad False
                    CargaGridCalidad True
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    DataGrid3.Enabled = True
                End If
            End If
            
            
            
            
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            Set FrmArt = New frmAlmArticu2
            'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
            FrmArt.DesdeTPV = False
            FrmArt.Show vbModal
            Set FrmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdAux2_Click(Index As Integer)
    If Index = 0 Then
        'Porvceedor
        EsCabecera = False
        MandaBusquedaPrevia ""
        
    Else
        'sartic
        
            Set FrmArt = New frmAlmArticu2
            'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
            FrmArt.DesdeTPV = False
            FrmArt.Show vbModal
            Set FrmArt = Nothing
            If Index = 1 Then
                PonerFoco txtComponentes(4)
            Else
                PonerFoco txtCalidad(0)
            End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            CargaTxtAux False, False
           
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            CargaGrid3 True
        Case 6
            TerminaBloquear
            ModificarCantidadComponentes False
           
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid2.AllowAddNew = False
                If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid2.Enabled = True
            CargaGrid3 True
            
            'TerminaBloquear
           '
           ' PonerModo 2
           ' CargaGrid3 True
           ' HabilitarModifCantidad False
            
        Case 7
            TerminaBloquear
            ModificarDatosCalidad False
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid3.AllowAddNew = False
                If Not data4.Recordset.EOF Then data4.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid3.Enabled = True
            CargaGridCalidad True
            
    End Select
End Sub

Private Sub OcultarMostrarFramaCalid(Ocultar As Boolean)
    
    Me.Option1(0).Value = Ocultar
    Me.Option1(1).Value = Not Ocultar
    Me.FrameCalidad.visible = Not Ocultar

End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
    
    
    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    'Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    'Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    CargaGrid3 False
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")

    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
End Sub




Private Sub BotonAnyadirSubLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid2, data3
    
    ModificarCantidadComponentes True

    DoEvents
    PonerFoco txtComponentes(4)
End Sub

Private Sub BotonAnyadirSubLineaCalidad()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid3, data4
    Me.chkCalidad.Value = 0
    Me.cboCalidad.ListIndex = -1
    ModificarDatosCalidad True

    DoEvents
    PonerFoco txtCalidad(0)
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
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
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
    OcultarMostrarFramaCalid True
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        

End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If Data2.Recordset.EOF Then Exit Sub
    
  '  vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
  '  vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
  '  If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False

    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub BotonModificarSubLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If data3.Recordset.EOF Then Exit Sub
    
  
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False

    ModificarCantidadComponentes True
    
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonModificarSubLineaCalidad()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If data4.Recordset.EOF Then Exit Sub
    
  
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    ModificarDatosCalidad True

    
    Me.DataGrid3.Enabled = False
    PonerFoco txtCalidad(3)
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub





Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If Not IsNull(Data1.Recordset!fecproduccion) Then
        MsgBox "Orden cerrada. No se puede eliminar", vbExclamation
        Exit Sub
    End If

    cad = "Envasado." & vbCrLf
    cad = cad & "----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la orden de envasado:"
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")
    cad = cad & vbCrLf & vbCrLf & "¿Desea Eliminarla? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de envasado?     "
    SQL = SQL & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codArtic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!codigo
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        
        'Las sublineas
        conn.Execute "DELETE FROM slienvpr2 " & SQL
        conn.Execute "DELETE FROM slienvprcalidad " & SQL
        'Las lineas
        conn.Execute "DELETE FROM slienvpr " & SQL
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
'        SituarDataTrasEliminar Data2, NumRegElim
        SituarDataPosicion Me.Data2, NumRegElim, SQL
        
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub



Private Sub BotonEliminarSubLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If data3.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la sublínea de envasado?     "
    SQL = SQL & vbCrLf
    SQL = SQL & vbCrLf & "Artículo:  " & data3.Recordset!codarti2 & " - " & data3.Recordset!NomArtic
    SQL = SQL & vbCrLf & "Lote:  " & DBLet(data3.Recordset!numLote)
    SQL = SQL & vbCrLf & "Cantidad:  " & Format(DBLet(data3.Recordset!cantidad, "N"), FormatoCantidad2)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = data3.Recordset.AbsolutePosition
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codArtic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!codigo
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        SQL = SQL & " AND codarti2 = " & DBSet(data3.Recordset!codarti2, "T")
        'Las sublineas
        conn.Execute "DELETE FROM slienvpr2 " & SQL
 
        ModificaLineas = 0
        CargaGrid3 True

        SituarDataPosicion Me.data3, NumRegElim, SQL
        

    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub




Private Sub BotonEliminarLineaCalidad()
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If data4.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar el dato de  calidad?     "
    SQL = SQL & vbCrLf
    SQL = SQL & vbCrLf & "Artículo:  " & data4.Recordset!codarti2 & " - " & data4.Recordset!NomArtic
    SQL = SQL & vbCrLf & "Ensayo:  " & DBLet(data4.Recordset!ensayo, "T")
    SQL = SQL & vbCrLf & "Resultado:  " & DBLet(data4.Recordset!resultado, "T")
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = data4.Recordset.AbsolutePosition
        SQL = " WHERE codartic = " & DBSet(data4.Recordset!codArtic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!codigo
        SQL = SQL & " and codalmac=" & data4.Recordset!codAlmac
        SQL = SQL & " AND codarti2 = " & DBSet(data4.Recordset!codarti2, "T")
        SQL = SQL & " AND codigoensayo = " & DBSet(data4.Recordset!codigoensayo, "T")
        
        
        'Las sublineas
        conn.Execute "DELETE FROM slienvprcalidad " & SQL
 
        ModificaLineas = 0
        CargaGridCalidad True

        SituarDataPosicion Me.data4, NumRegElim, SQL
        

    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Or Modo = 6 Or Modo = 7 Then 'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'cad = Data1.Recordset.Fields(0) & "|"
        'cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = Data1.Recordset.Fields(0)
        RaiseEvent DatoSeleccionado2(cad)
        Unload Me
    End If
End Sub



Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)


    On Error GoTo Error1

'    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
'
'    End If
'
    If Modo = 2 Or Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            'Poner descripcion de ampliacion lineas
            CargaGrid3 True
        Else
            
        End If
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid3_DblClick()
    If Modo = 7 Then
        If ModificaLineas = 1 Then Exit Sub
    Else
        If Modo <> 2 Then Exit Sub
    End If
    
    If data4.Recordset.EOF Then Exit Sub
    
    If Modo = 2 Then BotonCalidad
    BotonModificarSubLineaCalidad
    'PonerFoco txtCalidad(3)
End Sub

Private Sub Form_Activate()
    If Me.Tag <> "" Then
        Me.Tag = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 20
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 33
        
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 37 'Cambiar cantidad componentes
        
        'Enero08
        .Buttons(12).Image = 21 'Cerrar orden produccion
        
        
        .Buttons(14).Image = 16 'Imprimir Pedido
      '  .Buttons(15).Image = 27 'Imprimir Orden Instalacion
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

      
    LimpiarCampos   'Limpia los campos TextBox
   

    NombreTabla = "senvprod"
    Ordenacion = " ORDER BY codigo "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargarCombo_Tabla cboCalidad, "scalidad", "codigo", "ensayo", , False, "ensayo"

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codigo=-1"
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
    Else
        If Data1.Recordset.EOF Then
            PonerModo 1
            Text1(0).BackColor = vbYellow
        Else
            Me.Tag = "P" 'Para que en el activate ponga los campos
            PonerModo 2
        End If
    End If

    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub




Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    If Modo = 5 Then
        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    ElseIf Modo = 7 Then
        txtCalidad(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
        txtCalidad(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Cod Artic
    Else
        txtComponentes(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    End If
        
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        Else
            txtComponentes(1).Text = RecuperaValor(CadenaDevuelta, 1)
            txtComponentes(2).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub









Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub







Private Sub frmPe_DatoSeleccionado2(CadenaSeleccion As String)
    Text1(4).Text = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmPe = New frmFacEntPedidos
    frmPe.DatosADevolverBusqueda2 = "0"
    frmPe.Show vbModal
    Set frmPe = Nothing

    
    
    Screen.MousePointer = vbDefault
    
    
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub mnBuscar_Click()

    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    ElseIf Modo = 6 Then
        BotonEliminarSubLinea
    ElseIf Modo = 7 Then
        BotonEliminarLineaCalidad
    Else 'Eliminar Pedido
         BotonEliminar
    End If
End Sub






Private Sub mnImpOrde_Click()
'Impreme la Orden de Instalacion de un pedido
Dim cadFormula As String, cadParam As String
Dim devuelve As String, nomDocu As String
Dim numParam As Byte

    'Comprobar que hay un pedido seleccionado
    If Text1(0).Text = "" Then
        MsgBox "No hay ningún Pedido seleccionado.", vbInformation
        Exit Sub
    End If

    'Comprobar que algun Articulo pertenece a la familia de Instalaciones
    If Not PedidoConInstalaciones Then
        MsgBox "El Pedido no tiene ningún Artículo que sea Instalación.", vbInformation
        Exit Sub
    End If

    '=======================================================================
    '=============== FORMULA    ============================================
    cadFormula = ""
    cadParam = ""
    numParam = 0
    
    If Text1(0).Text <> "" Then 'Seleccionar el Pedido
        devuelve = "{" & NombreTabla & ".numpedcl}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    End If
    
    'Seleccionar solo las lineas de Articulos que son de una familia que es Instalacion
    devuelve = "{sfamia.instalac}=1"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    If Not PonerParamRPT2(9, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub

    With frmImprimir
        .NombreRPT = nomDocu
        .NombrePDF = pPdfRpt
        .FormulaSeleccion = cadFormula
        .SeleccionaRPTCodigo = pRptvMultiInforme
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 39
        .Titulo = ""
        .Show vbModal
    End With
End Sub




Private Sub mnLineas_Click()
    BotonMtoLineas
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    ElseIf Modo = 6 Then 'Sublineas
        BotonModificarSubLinea
    ElseIf Modo = 7 Then 'Sublineas
        BotonModificarSubLineaCalidad
    
    Else  'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    ElseIf Modo = 6 Then
        BotonAnyadirSubLinea
    ElseIf Modo = 7 Then
        BotonAnyadirSubLineaCalidad
    Else 'Añadir Cabecera de Pedidos
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




Private Sub Option1_Click(Index As Integer)
       Option1(0).FontBold = Index = 0
       Option1(1).FontBold = Index = 1
       
       Me.FrameCalidad.visible = Index = 1
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
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)
            
            If Index = 2 And Text1(Index).Text <> "" Then 'Fecha Entrega
                'Comprobar que es posterior a la del pedido
                If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
               
            End If
            
    
        Case 4 '
            If PonerFormatoEntero(Text1(Index)) Then

            Else
               
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(6).Text = Data1.Recordset!nifClien Then
'                    Exit Sub
'                End If
'            End If
'            PonerDatosClienteVario (Text1(Index).Text)
             
        Case 9 'Cod. Postal

            
 
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 20, "Nº Orden")
        cad = cad & ParaGrid(Text1(1), 20, "Fecha creación")
        cad = cad & ParaGrid(Text1(2), 20, "Fecha producción")
        tabla = NombreTabla
      
        Titulo = "Ordenes producción"
        devuelve = "0|"

    Else
        
        Titulo = "Proveedores"
        Desc = "Prov"
        
        'Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        cad = cad & "Cod. " & Desc & "|sprove|codprove|N||15·"
        cad = cad & "Desc. " & Desc & "|sprove|nomprove|T||35·"
        tabla = "sprove"
        devuelve = "0|1|"
    End If
    
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
          
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
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


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True


    'Calidad
    CargaGridCalidad True


    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    

       
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    

    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    Option1(0).Enabled = Modo < 3
    Option1(1).Enabled = Modo < 3
    

    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(2), b
    b = Modo = 0 Or Modo = 2 Or Modo >= 5
    BloquearTxt Text1(1), b
    BloquearTxt Text1(3), b
    BloquearTxt Text1(4), b

  
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
  
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5 And Modo <> 7)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    'Las imagenes añadimos el modo 6
    b = b And Modo <> 6
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    imgBuscar(0).visible = b


    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    'Solo en modificamos cantidad en modo6
    b = Modo = 6
    For i = 0 To txtComponentes.Count - 1
        txtComponentes(i).visible = False
    Next i
    Me.cmdAux2(0).visible = False 'b FALTA###
    
    If Modo = 2 Then
        DataGrid1.Enabled = True
        DataGrid2.Enabled = True
    End If
    
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim b As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then Exit Function
    
    
          
     'Si ha puesto el numero de pedido entonces
     'deberemos traer los datos
     OpcionConElPedido = 0
     If Text1(4).Text <> "" Then
     
        devuelve = DevuelveDesdeBD(conAri, "numpedcl", "scaped", "numpedcl", Text1(4).Text)
        If devuelve = "" Then
            MsgBox "No existe el pedido: " & Text1(4).Text, vbExclamation
            Exit Function
        End If
        If Modo = 3 Then
            OpcionConElPedido = 1 'INSERTAMOS Y A CORRER
        Else
            'Modificar. Si ya tenia datos entonces puede ser que quiera eliminar los datos anteriores
            'Si tenia pedido o no
            If Val(Text1(4).Text) <> DBLet(Data1.Recordset!numpedcl, "N") Then
                If Not Data2.Recordset.EOF Then
                    devuelve = "Se van a insertar las lineas del pedido: " & Text1(4).Text
                    devuelve = devuelve & vbCrLf & "¿Desea eliminar las lineas anteriores?"
                    NumRegElim = Val(MsgBox(devuelve, vbQuestion + vbYesNoCancel))
                    If CByte(NumRegElim) = vbCancel Then Exit Function
                    If CByte(NumRegElim) = vbYes Then
                        OpcionConElPedido = 2
                    Else
                        OpcionConElPedido = 1
                    End If
                Else
                    'EOF. insertamos
                    OpcionConElPedido = 1
                End If
            End If
        End If
    End If
    b = True
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim i As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" And i <> 3 Then
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
        
    
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLineaCompo() As Boolean
    DatosOkLineaCompo = False
    
    If Me.txtComponentes(3).Text = "" Then
        MsgBox "Cantidad obligatoria", vbExclamation
        Exit Function
    Else
        If Not IsNumeric(txtComponentes(3).Text) Then
            MsgBox "Campo numerico", vbExclamation
            Exit Function
        End If
    End If
    
    If Me.txtComponentes(1).Text <> "" Then
        If Not IsNumeric(txtComponentes(1).Text) Then
            MsgBox "Error en proveedor", vbExclamation
            Exit Function
        End If
    End If
    
    DatosOkLineaCompo = True
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
            
        
        Case 9
            
            BotonCalidad
            OcultarMostrarFramaCalid False
        
        Case 10  'Lineas
            mnLineas_Click
            
            
        Case 11
            'Modificar cantidad de componentes
            'ModificarCantidadComponentes
            If Data2.Recordset Is Nothing Then Exit Sub
            If Data2.Recordset.EOF Then Exit Sub
            
            ModificaLineas = 0
            OcultarMostrarFramaCalid True
            PonerModo 6
            PonerBotonCabecera True
    
            
        Case 12, 14
            'IMPRIMIR (14)    y cerrar(12) orden produccion
            '--------------------------------------------------------------------
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una orden de envasado", vbExclamation
                Exit Sub
            End If
            If Button.Index = 12 Then
                If Not IsNull(Data1.Recordset!fecproduccion) Then
                    MsgBox "La orden de envasado ya esta cerrada", vbExclamation
                    Exit Sub
                End If
                
                
                If BLOQUEADesdeFormulario(Me) Then
                    
                    frmProduVarios.Intercambio = Data1.Recordset!codigo & "|" & Data1.Recordset!feccreacion & "|"
                    frmProduVarios.Opcion = 1  'Envasado
                    frmProduVarios.Show vbModal
                
                    'TErminamos de bloquear
                    TerminaBloquear
                    
                    'Refrescamos
                    CadenaConsulta = DevuelveDesdeBD(conAri, "fecproduccion", "senvprod", "codigo", CStr(Data1.Recordset!codigo))
                    If CadenaConsulta <> "" Then
                        'Ok YA tiene fecha produccion
                        CadenaConsulta = Data1.RecordSource
                        Data1.Refresh
                        'Y ponemos los campos
                        PosicionarData
                  
                    Else
                        'Ha cancelado
                        CadenaConsulta = Data1.RecordSource
                    End If
                    
                
                End If
                
            Else
                'Imprimir orden prod
                With frmImprimir
                    .ConSubInforme = True
                    .FormulaSeleccion = "{senvprod.codigo} = " & Data1.Recordset!codigo
                    .NombreRPT = "rOrdenEnvasado.rpt"
                    .OtrosParametros = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
                    .NumeroParametros = 1
                    .Titulo = "Orden de produccion"
                    .Opcion = 2003 'Esta libre
                    .Show vbModal
                End With
            End If

        Case 15 'Imprimir Orden Instalacion
            mnImpOrde_Click
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
      
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String
Dim vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        SQL = "INSERT INTO slienvpr"
        SQL = SQL & "( codigo, codalmac, codartic ,cantidad,numlote ) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(4).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(3).Text, "T") & ")"
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        'Insertamos en lineas2
        ActualizarComponentes True
        
        
        
        'Insertaremos en calidad
        
        
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Produccion" & vbCrLf & Err.Description
End Function




Private Function InsertarSubLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String


    On Error GoTo EInsertarLinea

    InsertarSubLinea = False
    SQL = ""
    If txtComponentes(4).Text = "" Then
        MsgBox "Campo articulo obligado", vbExclamation
        Exit Function
    End If
    If DatosOkLineaCompo() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        SQL = "INSERT INTO slienvpr2"
        SQL = SQL & "(`codigo`,`codalmac`,`codartic`,`codarti2`,`cantidad`,`numlote`,`codprove`)"
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & Val(Data2.Recordset!codAlmac) & ","
        SQL = SQL & DBSet(Data2.Recordset!codArtic, "T") & ","
        SQL = SQL & DBSet(txtComponentes(4).Text, "T") & "," & DBSet(txtComponentes(3).Text, "N") & ","
        SQL = SQL & DBSet(txtComponentes(0).Text, "T") & "," & txtComponentes(1).Text & ")"
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        InsertarSubLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Produccion" & vbCrLf & Err.Description
End Function


Private Function InsertarSubLineaCalidad() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String


    On Error GoTo EInsertarLinea
    

    InsertarSubLineaCalidad = False
    SQL = ""
    If txtCalidad(0).Text = "" Then SQL = "- Campo articulo obligado" & vbCrLf
    If txtCalidad(1).Text = "" Then SQL = "- Articulo incorrecto" & vbCrLf
    If cboCalidad.ListIndex < 0 Then SQL = SQL & "- Ensayo obligado" & vbCrLf
    If SQL <> "" Then
        MsgBox "Campos erroneos: " & vbCrLf & SQL, vbExclamation
        Exit Function
    End If
    
   
        'Conseguir el siguiente numero de linea
        SQL = "INSERT INTO slienvprcalidad("
        SQL = SQL & "codigo,codalmac,codartic,codarti2,codigoensayo,especificaciones,resultado,conforme) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & Val(Data2.Recordset!codAlmac) & ","
        SQL = SQL & DBSet(Data2.Recordset!codArtic, "T") & ","
        SQL = SQL & DBSet(txtCalidad(0), "T") & "," & cboCalidad.ItemData(cboCalidad.ListIndex) & ","
        SQL = SQL & DBSet(txtCalidad(2).Text, "T") & "," & DBSet(txtCalidad(3).Text, "T") & "," & Abs(Me.chkCalidad.Value) & ")"
        
    
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        InsertarSubLineaCalidad = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas calidad" & vbCrLf & Err.Description
End Function





Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
Dim HaModificadoCantidad As Boolean
    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    HaModificadoCantidad = False
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE slienvpr set codalmac=" & txtAux(0).Text & " , codartic =" & DBSet(txtAux(1).Text, "T")
        SQL = SQL & ", numlote = " & DBSet(txtAux(3).Text, "T", "S")
        SQL = SQL & ", cantidad = " & DBSet(txtAux(4).Text, "N")
        SQL = SQL & " WHERE codigo =" & Data1.Recordset!codigo & " AND codalmac = " & Data2.Recordset!codAlmac
        SQL = SQL & " AND codartic =" & DBSet(Data2.Recordset!codArtic, "T")
        If ImporteFormateado(txtAux(4).Text) <> Data2.Recordset!cantidad Then HaModificadoCantidad = True
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        
        
        ActualizarComponentes True
        
        ModificarLinea = True
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " '& TituloLinea
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
        
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    CargaGrid3 enlaza
    
    
    
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub





Private Sub CargaGrid3(enlaza As Boolean)
Dim SQL As String

    SQL = "codigo = -1"


    If enlaza Then
       If Not Data2.Recordset.EOF Then
            SQL = " codigo = " & Data1.Recordset!codigo
            SQL = SQL & " AND codalmac = " & Data2.Recordset!codAlmac
            SQL = SQL & " AND slienvpr2.codartic = " & DBSet(Data2.Recordset!codArtic, "T")
            
       End If
    End If

    
    
    SQL = " Where slienvpr2.codarti2 = sartic.codArtic And " & SQL
    SQL = " sartic,slienvpr2 left join sprove on slienvpr2.codprove = sprove.codprove" & SQL
    SQL = " Select codarti2,nomartic,numlote,slienvpr2.codprove,nomprove,cantidad  from " & SQL

    data3.ConnectionString = conn
    data3.RecordSource = SQL
    data3.Refresh
    If DataGrid2.DataSource Is Nothing Then DataGrid2.ClearFields
        
    Set DataGrid2.DataSource = data3
    DataGrid2.RowHeight = 290
    DataGrid2.Columns(0).Caption = "Codigo"
    DataGrid2.Columns(0).Width = 1700
    
    
    DataGrid2.Columns(1).Caption = "Articulo"
    DataGrid2.Columns(1).Width = 3600

    DataGrid2.Columns(2).Caption = "Lote"
    DataGrid2.Columns(2).Width = 1500

    DataGrid2.Columns(3).Caption = "Prov."
    DataGrid2.Columns(3).Width = 750

    DataGrid2.Columns(4).Caption = "Nom. proveedor"
    DataGrid2.Columns(4).Width = 1600

    DataGrid2.Columns(5).Caption = "Cantidad"
    DataGrid2.Columns(5).Width = 1400
    DataGrid2.Columns(5).NumberFormat = FormatoPrecio
    DataGrid2.Columns(5).Alignment = dbgRight
End Sub



Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte

    On Error GoTo ECargaGrid

    vData.Refresh

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(0).Caption = "Alm."
                vDataGrid.Columns(0).Width = 500
                vDataGrid.Columns(0).NumberFormat = "000"
                
                vDataGrid.Columns(1).Caption = "Articulo"
                vDataGrid.Columns(1).Width = 1700

                
                vDataGrid.Columns(2).Caption = "Desc. Artículo"
                vDataGrid.Columns(2).Width = 3800

                vDataGrid.Columns(3).Caption = "Lote"
                vDataGrid.Columns(3).Width = 1800

      
                vDataGrid.Columns(4).Caption = "Cantidad"
                vDataGrid.Columns(4).Width = 1550
                vDataGrid.Columns(4).Alignment = dbgRight
                vDataGrid.Columns(4).NumberFormat = FormatoPrecio
             
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub








Private Sub CargaGridCalidad(enlaza As Boolean)
Dim SQL As String

    
    'SQL = "select codarti2,nomartic,ensayo,sliordprcalidad.especificaciones,resultado,if(conforme=1,'Si','') ok"
    'SQL = SQL & "  from sliordprcalidad,sartic,scalidad where  sartic.codArtic = sliordprcalidad.codarti2"
    'SQL = SQL & " and codigoensayo=scalidad.codigo  and  sliordprcalidad.codigo= "
    'SQL = SQL & IIf(enlaza, Data1.Recordset!codigo, -1)
    'SQL = SQL & " order by codarti2,ensayo"
    
    
    
    SQL = "select codarti2,nomartic,ensayo,slienvprcalidad.especificaciones,resultado,if(conforme=1,'Si','') ok"
    SQL = SQL & " ,slienvprcalidad.codArtic ,codAlmac ,codigoensayo" 'No se ven
    SQL = SQL & "  from slienvprcalidad,sartic,scalidad where  sartic.codArtic = slienvprcalidad.codarti2"
    SQL = SQL & " and codigoensayo=scalidad.codigo  and  slienvprcalidad.codigo= "
    SQL = SQL & IIf(enlaza, Data1.Recordset!codigo, -1)
    SQL = SQL & " order by codarti2,ensayo"
    
    
    data4.ConnectionString = conn
    data4.RecordSource = SQL
    data4.Refresh
    If DataGrid3.DataSource Is Nothing Then DataGrid3.ClearFields
        
    Set DataGrid3.DataSource = data4
    DataGrid3.RowHeight = 290
    DataGrid3.Columns(0).Caption = "Codigo"
    DataGrid3.Columns(0).Width = 1500
    
    
    DataGrid3.Columns(1).Caption = "Articulo"
    DataGrid3.Columns(1).Width = 2600

    DataGrid3.Columns(2).Caption = "Ensayo"
    DataGrid3.Columns(2).Width = 1400

    DataGrid3.Columns(3).Caption = "Especificación"
    DataGrid3.Columns(3).Width = 1950

    DataGrid3.Columns(4).Caption = "Resultado"
    DataGrid3.Columns(4).Width = 1800

    DataGrid3.Columns(5).Caption = "OK"
    DataGrid3.Columns(5).Width = 700
    
    DataGrid3.Columns(6).visible = False
    DataGrid3.Columns(7).visible = False
    DataGrid3.Columns(8).visible = False
    'DataGrid3.Columns(5).NumberFormat = FormatoPrecio
    'DataGrid3.Columns(5).Alignment = dbgRight
End Sub













'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = DataGrid1.Columns(i).Text
                txtAux(i).Locked = i < 2
            Next i
        End If
               
'        'El Campo de Origen del precio se actualiza por programa al modificar el precio
'        BloquearTxt txtAux(5), True
'        'El campo Importe es calculado y lo bloqueamos.
'        BloquearTxt txtAux(8), True
'

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(0).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(1).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(2).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(3).Width - 10
        'Cantidad
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 10
        txtAux(4).Width = DataGrid1.Columns(4).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(0).Enabled = ModificaLineas = 1
        cmdAux(1).Enabled = ModificaLineas = 1
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Modo6: Pasar de Pedido a Albaran
        If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    End If
End Sub




Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Modo 6: Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
'            PonerServidas
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As CStock
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            'If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then 'Cod Artic
                txtAux(2).Text = "" 'Nom Artic
                Exit Sub
            End If
            If txtAux(0).Text = "" Then 'Cod Almacen
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If

            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, "", ModificaLineas, devuelve) Then
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            End If
            
        Case 2 'desc Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 4 'CANTIDAD
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 2) Then   'Tipo 2: 4 decimales
    
                Else
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
            
        
    End Select
    

End Sub


Private Sub BotonMtoLineas()
       
        ModificaLineas = 0
        OcultarMostrarFramaCalid True
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean



    On Error GoTo FinEliminar

        conn.BeginTrans

        conn.Execute "Delete from slienvpr where codigo =" & Text1(0).Text
        conn.Execute "Delete from senvprod where codigo =" & Text1(0).Text
        b = True
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido" & vbCrLf, Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    CargaGridCalidad False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = Replace(ObtenerWhereCP, NombreTabla & ".", "")
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
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String

    On Error Resume Next
    
    SQL = NombreTabla & ".codigo= " & Val(Text1(0).Text)
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT codalmac,slienvpr.codartic,nomartic,numlote,cantidad "
    SQL = SQL & " FROM slienvpr,sartic WHERE slienvpr.codartic=sartic.codartic AND "
    If enlaza Then
        SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, "slienvpr")
    Else
        SQL = SQL & " codigo = -1"
    End If
    SQL = SQL & " Order by codigo"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo >= 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
            
        b = (Modo = 2)
        
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = b
        'Me.mnLineas.Enabled = B
        
        
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        'Generar Albaran desde Pedido
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenAlbaran.Enabled = b
        
        Toolbar1.Buttons(12).Enabled = b
        Me.mnGeneraFactura.Enabled = b
        Toolbar1.Buttons(13).Enabled = b
        
        
        
      
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub







    

Private Function PedidoConInstalaciones() As Boolean
'Comprobar si en las lineas del Pedido hay algun articulo que sea Instalacion
'Si no hay niguna linea que sea instalacion no se imprimira la Orden de Instalacion
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EInstalac

    PedidoConInstalaciones = False
    SQL = "SELECT sliped.codartic, sliped.numlinea,scaped.numpedcl, sfamia.instalac "
    SQL = SQL & " FROM ((sliped INNER JOIN scaped ON sliped.numpedcl=scaped.numpedcl) "
    SQL = SQL & " INNER JOIN sartic ON sliped.codartic=sartic.codartic) INNER JOIN "
    SQL = SQL & " sfamia ON sartic.codfamia=sfamia.codfamia "
    SQL = SQL & " WHERE scaped.numpedcl = " & Val(Text1(0).Text) & " And sfamia.instalac = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        PedidoConInstalaciones = False
    Else
        PedidoConInstalaciones = True
    End If
    RS.Close
    Set RS = Nothing
    
EInstalac:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar si hay Articulos que son Instalaciones.", Err.Description
End Function






Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String

    On Error GoTo EEliminarPed

     SQL = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
   ' conn.Execute "Delete from " & NomTablaLineas & sql

    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function








Private Sub InsertarCabecera()
Dim cT As CTiposMov
    
    'Ahora lo insertaremos por tipo de movimiento
    'Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codigo")
    Set cT = New CTiposMov
    If cT.Leer("PRE") Then   'PRE:  Produccion envasado
        Text1(0).Text = cT.ConseguirContador("PRE")
        cT.IncrementarContador "PRE"
        If InsertarDesdeForm(Me) Then
                
                ActualizarLineasPedido
        
                'Si tiene pedido traeremos las lineas del pedido
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE codigo = " & Text1(0).Text & Ordenacion
                PonerCadenaBusqueda
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas
                BotonAnyadirLinea
        Else
            cT.DevolverContador "PRE", cT.Contador
        End If
    Else
        'Error leyendo tipo MOVIMIENTO
        
    End If
    Set cT = Nothing
End Sub

Private Sub ActualizarLineasPedido()
Dim SQL As String
    If OpcionConElPedido = 0 Then Exit Sub
    
    'Si tiene que coger pero no tiene pedido (NO DEBERIA PASAR)
    If Text1(4).Text = "" Then Exit Sub
    
    If OpcionConElPedido = 2 Then
        'Eliminamos los que hubieren
        SQL = "DELETE FROM slienvpr where codigo = " & Text1(0).Text
        conn.Execute SQL
    End If
    SQL = "INSERT IGNORE INTO slienvpr(codigo,codalmac,codartic,cantidad)"
    SQL = SQL & "select " & Text1(0).Text & ",codalmac,codartic,sum(cantidad) from sliped"
    SQL = SQL & " Where numpedcl = " & Text1(4).Text
    SQL = SQL & " group by 1,2,3"
    conn.Execute SQL
    
End Sub

Private Function UpdateaCantidadComponentes() As Boolean
Dim SQL As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    UpdateaCantidadComponentes = False
    
    SQL = "UPDATE slienvpr2 SET cantidad = " & DBSet(txtComponentes(3).Text, "N")
    'LOTE
    SQL = SQL & ", numlote = " & DBSet(txtComponentes(0).Text, "T", "S")
    SQL = SQL & ", codprove = " & txtComponentes(1).Text
    
    SQL = SQL & " WHERE codartic = " & DBSet(Data2.Recordset!codArtic, "T")
    SQL = SQL & " and codigo=" & Data1.Recordset!codigo
    SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
    SQL = SQL & " and codarti2=" & DBSet(data3.Recordset!codarti2, "T")
    conn.Execute SQL
    
    
    ActualizarComponentes False
    
    
    Espera 0.5
    CargaGridCalidad True
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        UpdateaCantidadComponentes = True
    End If
    Screen.MousePointer = vbDefault
End Function




'ACutalizaremos las sublineas(componentes)
'Es decir. Si insertamos o modificamos un elemento que tiene componentes
'insertaremos en sliorpd
Private Sub ActualizarComponentes(DesdeArticuloProduccion As Boolean)
Dim SQL As String
Dim LoteEncontrado As Boolean
Dim RN As ADODB.Recordset

    
    If ModificaLineas = 2 Then
        'BORRAMOS los datos que hubieren
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codArtic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!codigo
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        
        If DesdeArticuloProduccion Then conn.Execute "DELETE FROM slienvpr2 " & SQL

        
        conn.Execute "DELETE FROM slienvprcalidad " & SQL
    End If
        
        
    If ModificaLineas <> 2 Or DesdeArticuloProduccion Then
        Espera 0.2
        
        SQL = "INSERT INTO slienvpr2"
        SQL = SQL & "( codigo, codalmac, codartic ,codarti2,cantidad,numlote,codprove ) "
        If vParamAplic.ComponentePorcentaje Then
            'Los componentes en materias primas entran como porcentajes. Tipo fontentas
            SQL = SQL & "select " & Val(Text1(0).Text) & ", " & Val(txtAux(0).Text) & ","
            SQL = SQL & DBSet(txtAux(1).Text, "T") & ",sarti1.codarti1,"
            SQL = SQL & " if (mateprima=0,cantidad * " & DBSet(txtAux(4).Text, "N")
            SQL = SQL & ", (Cantidad / 100) * " & DBSet(txtAux(4).Text, "N") & "),NULL,codprove"
            SQL = SQL & " FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic"
            SQL = SQL & " WHERE sarti1.codartic = " & DBSet(txtAux(1).Text, "T")
        Else
            'Por cantidad. Lo de siempre vamos
            SQL = SQL & "select " & Val(Text1(0).Text) & ", " & Val(txtAux(0).Text) & ","
            SQL = SQL & DBSet(txtAux(1).Text, "T") & ",sarti1.codarti1,cantidad * " & DBSet(txtAux(4).Text, "N") & ",NULL,codprove "
            SQL = SQL & " FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic"
            SQL = SQL & " WHERE sarti1.codartic = " & DBSet(txtAux(1).Text, "T")
            
        End If
        conn.Execute SQL
        
    End If
    'CALIDAD
    LoteEncontrado = False
    If ModificaLineas = 2 Then
        Set RN = New ADODB.Recordset
        'Vere si existe el lote en envasado
        SQL = "Select * from sliordpr where codartic = " & DBSet(data3.Recordset!codarti2, "T")
        SQL = SQL & " AND numlote = " & DBSet(txtComponentes(0).Text, "T", "N")
        SQL = SQL & " and codalmac=" & Data2.Recordset!codAlmac
        RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RN.EOF Then
            MsgBox "No se encuentra el lote en produccion", vbExclamation
        
        Else
            LoteEncontrado = True
            
            SQL = "select " & Data1.Recordset!codigo & ",codalmac," & DBSet(Data2.Recordset!codArtic, "T") & ",codartic,codigoensayo,resultado,conforme"
            SQL = SQL & " from sliordprcalidad where "
            SQL = SQL & " codartic=" & DBSet(RN!codArtic, "T")
            SQL = SQL & " and codalmac=" & RN!codAlmac
            SQL = SQL & " and codigo=" & RN!codigo
            SQL = "INSERT INTO slienvprcalidad (codigo,codalmac,codartic,codarti2,codigoensayo,resultado,conforme) " & SQL
            conn.Execute SQL
        End If
        RN.Close
        Set RN = Nothing
    End If
    
    
    
    
    If Not LoteEncontrado Then
        SQL = "INSERT INTO slienvprcalidad"
        SQL = SQL & "(codigo,codalmac,codartic,codarti2,codigoensayo,especificaciones,resultado,conforme) "
        SQL = SQL & "select " & Val(Text1(0).Text) & ", " & Val(Data2.Recordset!codAlmac) & ","
        SQL = SQL & " sarti1.codartic,codarti1,codigoensayo,especificaciones,'',0"
        SQL = SQL & " FROM   sarti1 INNER JOIN sarti7 ON sarti1.codarti1 = sarti7.codArtic "
        SQL = SQL & " WHERE sarti1.codartic = " & DBSet(Data2.Recordset!codArtic, "T")
        SQL = SQL & " AND sarti1.codarti1 = " & DBSet(data3.Recordset!codarti2, "T")
        conn.Execute SQL

    End If
End Sub






'Praparamos para modificar la cantidad de los compoenntes
Private Sub ModificarCantidadComponentes(visible As Boolean)
Dim i As Integer
    
    
    If visible Then
        For i = 0 To 3
            If data3.Recordset.EOF Then
                Me.txtComponentes(i).Top = DataGrid2.Top + DataGrid2.RowTop(0) + 10
            Else
                Me.txtComponentes(i).Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + 10
            End If
            Me.txtComponentes(i).Left = DataGrid2.Left + DataGrid2.Columns(2 + i).Left
            Me.txtComponentes(i).Width = DataGrid2.Columns(2 + i).Width
            If ModificaLineas = 2 Then
                txtComponentes(i).Text = DataGrid2.Columns(2 + i).Text
            Else
                txtComponentes(i).Text = ""
            End If
        Next
        For i = 0 To 1
            Me.txtComponentes(4 + i).Left = DataGrid2.Left + DataGrid2.Columns(i).Left
            Me.txtComponentes(4 + i).Top = Me.txtComponentes(1).Top
            If ModificaLineas = 2 Then
                txtComponentes(4 + i).Text = DataGrid2.Columns(i).Text
            Else
                txtComponentes(4 + i).Text = ""
            End If
            Me.txtComponentes(4 + i).Width = DataGrid2.Columns(i).Width
        Next i
        cmdAux2(0).Left = txtComponentes(2).Left - 90
        cmdAux2(0).Top = Me.txtComponentes(0).Top
        cmdAux2(1).Top = Me.txtComponentes(0).Top
        cmdAux2(1).Left = Me.txtComponentes(5).Left - 30
        HabilitarModifCantidad True
        
    Else
        
    
    End If
    
    cmdAux2(0).visible = visible
    cmdAux2(1).visible = visible And ModificaLineas = 1
    For i = 0 To Me.txtComponentes.Count - 1
        
        txtComponentes(i).visible = visible
        If i >= 4 And ModificaLineas <> 1 Then txtComponentes(i).visible = False

    Next
    
End Sub



Private Sub HabilitarModifCantidad(Habilitar As Boolean)
    If Habilitar Then
        DeseleccionaGrid DataGrid1
        DeseleccionaGrid DataGrid2
    End If
    DataGrid1.Enabled = Not Habilitar
    DataGrid2.Enabled = Not Habilitar
End Sub



Private Sub txtCalidad_GotFocus(Index As Integer)
     ConseguirFoco txtCalidad(Index), 3
End Sub

Private Sub txtCalidad_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCalidad_LostFocus(Index As Integer)
Dim SQL As String

    txtCalidad(Index).Text = Trim(txtCalidad(Index).Text)
    If Index = 0 Then
        
        If Me.txtCalidad(Index).Text = "" Then
            txtCalidad(1).Text = ""
        Else
            SQL = "codartic"
            txtCalidad(1).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtCalidad(Index).Text, "T", SQL)
            If txtCalidad(1).Text = "" Then
                MsgBox "No existe el articulo: " & txtCalidad(Index).Text, vbExclamation
                txtCalidad(Index).Text = ""
                PonerFoco txtCalidad(Index)
            Else
                txtCalidad(Index).Text = SQL
            End If
        End If
    End If
    
End Sub


Private Sub txtComponentes_GotFocus(Index As Integer)
    ConseguirFoco txtComponentes(Index), 3
End Sub

Private Sub txtComponentes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtComponentes_LostFocus(Index As Integer)
Dim C As String
    txtComponentes(Index).Text = Trim(txtComponentes(Index).Text)
    Select Case Index
    Case 1
        CadenaConsulta = ""
        If txtComponentes(Index).Text <> "" Then
            If PonerFormatoEntero(txtComponentes(Index)) Then
                CadenaConsulta = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtComponentes(Index).Text)
                If CadenaConsulta = "" Then
                    MsgBox "No existe el proveedor: " & txtComponentes(1).Text, vbExclamation
                    txtComponentes(Index).Text = ""
                    PonerFoco txtComponentes(Index)
                End If
            Else
                txtComponentes(Index).Text = ""
                PonerFoco txtComponentes(Index)
            End If
        End If
        txtComponentes(2).Text = CadenaConsulta
        CadenaConsulta = ""
    
    
    Case 3
        If txtComponentes(Index).Text <> "" Then PonerFormatoDecimal txtComponentes(Index), 2      '4 decimales
    Case 4
        'codartic
        CadenaConsulta = ""
        C = ""
        If txtComponentes(Index).Text <> "" Then
           C = "codprove"
           CadenaConsulta = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtComponentes(Index).Text, "T", C)
           If CadenaConsulta = "" Then
                    MsgBox "No existe el articulo: " & txtComponentes(Index).Text, vbExclamation
                    txtComponentes(Index).Text = ""
                    PonerFoco txtComponentes(Index)
                    C = ""
               
            Else
                txtComponentes(1).Text = C
            End If
        End If
        
        txtComponentes(5).Text = CadenaConsulta
        
        CadenaConsulta = ""
    End Select
End Sub


Private Sub BotonCalidad()
    If Modo <> 2 Then Exit Sub
        
    PonerModo 7
    cmdRegresar.visible = True
    
End Sub



'Praparamos para modificar la cantidad de los compoenntes
Private Sub ModificarDatosCalidad(visible As Boolean)
Dim i As Integer
    
    cmdAux2(2).visible = False
    
    If visible Then
        cboCalidad.visible = False
        If data4.Recordset.EOF Then
            Me.txtCalidad(3).Top = DataGrid3.Top + DataGrid3.RowTop(0) + 10
        Else
            Me.txtCalidad(3).Top = DataGrid3.Top + DataGrid3.RowTop(DataGrid3.Row) + 10
        End If
        If ModificaLineas = 2 Then
            'Solo modifica resultado y conforme
            txtCalidad(3).Width = DataGrid3.Columns(4).Width
            txtCalidad(3).Text = DataGrid3.Columns(4).Text
            Me.txtCalidad(3).Left = DataGrid3.Left + DataGrid3.Columns(4).Left + 15
            txtCalidad(3).visible = True
            Me.chkCalidad.Value = Abs(DataGrid3.Columns(5).Text = "Si")
        Else
            cboCalidad.visible = True
            cmdAux2(2).visible = True
            For i = 0 To 3
                If i < 2 Then
                    Me.txtCalidad(i).Left = DataGrid3.Left + DataGrid3.Columns(i).Left
                    Me.txtCalidad(i).Width = DataGrid3.Columns(i).Width
                Else
                    Me.txtCalidad(i).Left = DataGrid3.Left + DataGrid3.Columns(i + 1).Left + 30
                    Me.txtCalidad(i).Width = DataGrid3.Columns(i + 1).Width - 30
                End If
                cboCalidad.Top = Me.txtCalidad(3).Top
                cboCalidad.Left = DataGrid3.Left + DataGrid3.Columns(2).Left + 45
                cboCalidad.Width = DataGrid3.Columns(2).Width - 15
                cmdAux2(2).Top = Me.txtCalidad(3).Top
                cmdAux2(2).Left = txtCalidad(1).Left - 60
                Me.txtCalidad(i).Top = Me.txtCalidad(3).Top
                txtCalidad(i).visible = True
                txtCalidad(i).Text = ""
            Next i
            BloquearTxt txtCalidad(1), True
            BloquearTxt txtCalidad(2), True
            
        End If
        
        chkCalidad.Left = DataGrid3.Columns(5).Left + 240
        Me.chkCalidad.visible = True
        chkCalidad.Top = txtCalidad(3).Top
    Else
        For i = 0 To Me.txtCalidad.Count - 1
            txtCalidad(i).visible = visible
        Next i
        Me.chkCalidad.visible = False
        cboCalidad.visible = False
        
    End If
    

End Sub




Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String
        If ModificaLineas = 1 Then Exit Function
        
        If UpdateaDatosCalidad() Then
            TerminaBloquear
            NumReg = data4.Recordset.AbsolutePosition
            CargaGridCalidad True
            If SituarDataPosicion(data4, NumReg, Indicador) Then
                  BotonModificarSubLineaCalidad
            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    
End Function



Private Function UpdateaDatosCalidad() As Boolean
Dim SQL As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    UpdateaDatosCalidad = False
    
  
    
    SQL = "UPDATE slienvprcalidad  "
    SQL = SQL & " SET  resultado  = " & DBSet(txtCalidad(3).Text, "T", "S")
    SQL = SQL & ", conforme = " & Abs(Me.chkCalidad.Value)
    
    SQL = SQL & " WHERE codartic = " & DBSet(data4.Recordset!codArtic, "T")
    SQL = SQL & " and codigo=" & Data1.Recordset!codigo
    SQL = SQL & " and codalmac=" & data4.Recordset!codAlmac
    SQL = SQL & " and codarti2=" & DBSet(data4.Recordset!codarti2, "T")
    SQL = SQL & " and codigoensayo=" & data4.Recordset!codigoensayo
    conn.Execute SQL
    Espera 0.1
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        UpdateaDatosCalidad = True
    End If
    Screen.MousePointer = vbDefault
End Function

