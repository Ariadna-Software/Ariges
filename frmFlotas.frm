VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFlotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flotas-Maquinaria"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11415
   ClipControls    =   0   'False
   Icon            =   "frmFlotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   15
      Left            =   9720
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "Coste|N|S|||sflotas|precostehora|||"
      Top             =   3720
      Width           =   1455
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
      Height          =   360
      Index           =   2
      Left            =   6240
      MaxLength       =   100
      TabIndex        =   40
      Tag             =   "Codigo|T|S|||sflotas|ampliacion|||"
      Top             =   3120
      Width           =   4935
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
      Index           =   9
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   39
      Tag             =   "Tara|N|S|0||sflotas|tara|0||"
      Top             =   3120
      Width           =   855
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
      Index           =   10
      Left            =   3720
      MaxLength       =   15
      TabIndex        =   38
      Tag             =   "R|N|S|0||sflotas|PMA|0||"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Index           =   14
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   3720
      Width           =   4965
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
      Height          =   360
      Index           =   14
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "Proveedor|N|S|||sflotas|codprove|||"
      Text            =   "Tex"
      Top             =   3720
      Width           =   975
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
      Height          =   360
      Index           =   13
      Left            =   5880
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "T|T|S|||sflotas|NumRoma|||"
      Top             =   1920
      Width           =   1095
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
      Height          =   360
      Index           =   12
      Left            =   5880
      MaxLength       =   15
      TabIndex        =   9
      Tag             =   "Fecha ult. inspeccion|F|S|||sflotas|fecultinsp|||"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboMarca 
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
      ItemData        =   "frmFlotas.frx":000C
      Left            =   1680
      List            =   "frmFlotas.frx":000E
      TabIndex        =   2
      Text            =   "cboMarca"
      Top             =   1320
      Width           =   2535
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
      Height          =   960
      Index           =   11
      Left            =   1680
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   13
      Tag             =   "Codigo|T|S|||sflotas|Observa|||"
      Top             =   4320
      Width           =   9495
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
      Height          =   360
      Index           =   8
      Left            =   8520
      MaxLength       =   15
      TabIndex        =   10
      Tag             =   "Fecha adqu|F|S|||sflotas|Fecbaja|||"
      Top             =   2520
      Width           =   1455
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
      Height          =   360
      Index           =   7
      Left            =   8520
      MaxLength       =   60
      TabIndex        =   7
      Tag             =   "R|T|S|||sflotas|conductor|||"
      Top             =   1920
      Width           =   2655
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
      Height          =   360
      Index           =   6
      Left            =   5880
      MaxLength       =   15
      TabIndex        =   3
      Tag             =   "R|T|S|||sflotas|modelo|||"
      Top             =   1320
      Width           =   2295
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
      Height          =   360
      Index           =   5
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   29
      Tag             =   "R|T|S|||sflotas|marca|||"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10200
      TabIndex        =   16
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox cboTipo 
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
      ItemData        =   "frmFlotas.frx":0010
      Left            =   9000
      List            =   "frmFlotas.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo vehiculo|N|N|||sflotas|Tipo|||"
      Top             =   1320
      Width           =   2175
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
      Height          =   360
      Index           =   4
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "Fecha adqu|F|S|||sflotas|Fecadq|||"
      Top             =   2520
      Width           =   1695
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
      Height          =   360
      Index           =   3
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "R|T|S|||sflotas|referencia|||"
      Top             =   1920
      Width           =   1455
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
      Height          =   360
      Index           =   1
      Left            =   5880
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Codigo|T|N|||sflotas|nomflota|||"
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   14
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10200
      TabIndex        =   15
      Top             =   7800
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   7680
      Width           =   2655
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
         TabIndex        =   23
         Top             =   180
         Width           =   2115
      End
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
      Height          =   360
      Index           =   0
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "Codigo|T|N|||sflotas|codflota||S|"
      Text            =   "Tex"
      Top             =   720
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Intercalar"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6720
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3360
      Top             =   7920
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   4680
      Top             =   7800
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
      Bindings        =   "frmFlotas.frx":0014
      Height          =   2010
      Left            =   1680
      TabIndex        =   17
      Top             =   5400
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3545
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
   Begin VB.Label Label1 
      Caption         =   "Coste/hora"
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
      Index           =   14
      Left            =   8520
      TabIndex        =   44
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliacion"
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
      Left            =   5040
      TabIndex        =   43
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Tara"
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
      Left            =   150
      TabIndex        =   42
      Top             =   3120
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "P.M.A."
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
      Index           =   10
      Left            =   3000
      TabIndex        =   41
      Top             =   3120
      Width           =   645
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   14
      Left            =   1320
      Picture         =   "frmFlotas.frx":0029
      ToolTipText     =   "Buscar cliente varios"
      Top             =   3720
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      Left            =   150
      TabIndex        =   36
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Roma"
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
      Index           =   13
      Left            =   4440
      TabIndex        =   35
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Ult. I.T.V."
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
      Index           =   12
      Left            =   4440
      TabIndex        =   34
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   5640
      Picture         =   "frmFlotas.frx":012B
      ToolTipText     =   "Buscar fecha"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observa."
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
      Index           =   11
      Left            =   150
      TabIndex        =   33
      Top             =   4320
      Width           =   885
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   8280
      Picture         =   "frmFlotas.frx":06B5
      ToolTipText     =   "Buscar fecha"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "F. baja"
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
      Index           =   8
      Left            =   7320
      TabIndex        =   32
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Conductor"
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
      Index           =   7
      Left            =   7320
      TabIndex        =   31
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo"
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
      Index           =   6
      Left            =   4440
      TabIndex        =   30
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Marca"
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
      Index           =   5
      Left            =   150
      TabIndex        =   28
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
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
      Index           =   4
      Left            =   8520
      TabIndex        =   27
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "F.adquisicion"
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
      Index           =   3
      Left            =   150
      TabIndex        =   26
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "frmFlotas.frx":0C3F
      ToolTipText     =   "Buscar fecha"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Referencia"
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
      Index           =   2
      Left            =   150
      TabIndex        =   25
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
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
      Left            =   4440
      TabIndex        =   24
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo"
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
      Left            =   150
      TabIndex        =   20
      Top             =   720
      Width           =   750
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
      TabIndex        =   19
      Top             =   8160
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
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
Attribute VB_Name = "frmFlotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1

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

Dim CadenaConsulta2 As String
Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean



Private Sub cboMarca_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    ComprobarCombo
                    PosicionarData
                    'BotonMtoLineas
                    'BotonAnyadirLinea False
                End If
            End If
        Case 4 'MODIFICAR
               If DatosOk Then
                    If ModificaDesdeFormulario(Me, 1) Then
                        'Si me cambia el proveedor, entonces guardo el LOG
                        If Val(DBLet(Data1.Recordset!Codprove, "N")) <> Val(Text1(14).Text) Then
                                'Ha cambiado el proveedor
                                Set LOG = New cLOG
                                
                                CadenaConsulta2 = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", CStr(DBLet(Data1.Recordset!Codprove, "N")))
                                
                                CadenaConsulta2 = "Anterior: " & DBLet(Data1.Recordset!Codprove, "T") & " - " & CadenaConsulta2
                                CadenaConsulta2 = "Actual: " & Text1(14).Text & " - " & Text2(14).Text & vbCrLf & CadenaConsulta2
                                CadenaConsulta2 = "[FLOTAS]" & vbCrLf & CadenaConsulta2
                                LOG.Insertar 29, vUsu, CadenaConsulta2
                                Set LOG = Nothing
                        End If
                        ComprobarCombo
                        TerminaBloquear
                        PosicionarData
                    End If
                End If
        Case 5 'InsertarModificar linea
                'Actualizar el registro en la tabla de lineas 'slipla' (Plantillas)
'                If ModificaLineas = 1 Then 'INSERTAR lineas
'                    If InsertarLinea Then
'                        CargaGrid True
'
'                        If LineaIntercalar2 > 0 Then
'                            'HA intercalado la linea. Ponemos luego en normal
'                            Me.DataGrid1.Enabled = True
'                            DataGrid1.AllowAddNew = False
'                            NumRegElim = LineaIntercalar2
'                            'CargaTxtAux False, False
'                            'CargaGrid2 DataGrid1, Data2
'
'                            ModificaLineas = 0
'                            PonerBotonCabecera True
'                            LLamaLineas 9
'
'                        Else
'                            BotonAnyadirLinea False
'                        End If
'
'
'
'
'                    End If
'                ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
'                    If ModificarLinea Then
'                        TerminaBloquear
'                        ModificaLineas = 0
'                        PonerBotonCabecera True
'                        CargaGrid True
'                        LLamaLineas 9
'                    End If
'                End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




'Private Sub cmdAux_Click(Index As Integer)
''    Set frmA = New frmAlmArticu2
''    'frmA.DatosADevolverBusqueda3 = "@1@" 'Poner en modo Busqueda
''    frmA.DesdeTPV = False
''    frmA.Show vbModal
''    Set frmA = Nothing
''    PonerFoco txtAux(0)
'End Sub

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
        Case 5 'Lineas
            TerminaBloquear
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            DataGrid1.Enabled = True
            ModificaLineas = 0
            PonerBotonCabecera True
            'LLamaLineas 9
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Lineas Ofertas
        DataGrid1.ClearFields
        PonerModo 2
        Me.lblIndicador.Caption = ""
    ElseIf Modo = 2 Then
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
 
        CadenaConsulta2 = Data1.Recordset.Fields(0) & "|"
        CadenaConsulta2 = CadenaConsulta2 & Data1.Recordset.Fields(2) & "|"
        RaiseEvent DatoSeleccionado(CadenaConsulta2)
        Unload Me
    End If
End Sub




Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    btnAnyadir = 5
    btnPrimero = 16 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        
        .Buttons(9).Image = 10 'Mto Lineas
        .Buttons(11).Image = 34 'Intercalar
        .Buttons(13).Image = 16 'Imprimir
        .Buttons(14).Image = 15 'Salir
        .Buttons(16).Image = 6 'Primero
        .Buttons(17).Image = 7 'Ante  rior
        .Buttons(18).Image = 8 'Siguiente
        .Buttons(19).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    Me.cmdRegresar.visible = DatosADevolverBusqueda <> ""    'Si es regresar
    
    CargarCombo_Tabla Me.cboTipo, "sflotatipo", "tipflota", "desctipflota"
    
    NombreTabla = "sflotas" 'Tabla Cabecera Plantillas
    NomTablaLineas = "slipla" 'Tabla Lineas Plantillas
    Ordenacion = " ORDER BY codflota"
    CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE codflota = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    CargaCboMarcas
    
    PonerModo 0
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim i As Byte
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    
    '### a mano
  
    DataGrid1.ScrollBars = dbgAutomatic
    
    
      
    i = 0 'Cod. Tipo Unidad
        DataGrid1.Columns(i).Caption = "Fecha"
        DataGrid1.Columns(i).Width = 1100
        DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy"
    
    i = 1 'Desc. Tipo Unidad
        DataGrid1.Columns(i).Caption = "Concepto"
        DataGrid1.Columns(i).Width = 2900
'
    i = 2 'Abrev.
        DataGrid1.Columns(i).Caption = "Ampliacion"
        DataGrid1.Columns(i).Width = 2750
    i = 3 'base im.
        DataGrid1.Columns(i).Caption = "Base Imp"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).Alignment = dbgRight
        DataGrid1.Columns(i).NumberFormat = FormatoImporte
    
    DataGrid1.Enabled = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



'Private Sub LLamaLineas(alto As Single)
''Pone posicion TOP y LEFT de los controles en el form
'Dim jj As Integer
'Dim b As Boolean
'
'
'    DeseleccionaGrid Me.DataGrid1
'
'    'Fijamos el ancho
'    b = (Modo = 5 And ModificaLineas = 1 Or ModificaLineas = 2)
'
'    For jj = 0 To txtAux.Count - 1
'        txtAux(jj).Height = DataGrid1.RowHeight
'        txtAux(jj).Top = alto
'        txtAux(jj).visible = b
'        If b Then txtAux(jj).Text = ""
'    Next jj
'
'    jj = 0
'    Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
'    Me.cmdAux(jj).Top = alto
'    Me.cmdAux(jj).visible = b
'End Sub



'Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento Articulos
'    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
'    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String

    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    CadenaConsulta2 = ""
    Select Case Index
        Case 14  '
'                Set frmP = New frmComProveedores
'                frmP.DatosADevolverBusqueda = "0"
'                frmP.Show vbModal
                Set frmP = New frmBasico2
                AyudaProveedores frmP, Text1(2)
                Set frmP = Nothing
    End Select
    PonerFoco Text1(2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_Selec(vFecha As Date)
    CadenaConsulta2 = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)

    CadenaConsulta2 = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Index = 0 Then
        NumRegElim = 4
    ElseIf Index = 1 Then
        NumRegElim = 8
    Else
        NumRegElim = 12
    End If
    If Me.Text1(NumRegElim).Text <> "" Then frmC.Fecha = CDate(Text1(NumRegElim).Text)
    frmC.Show vbModal
    If CadenaConsulta2 <> "" Then
    
        Text1(NumRegElim).Text = CadenaConsulta2
        CadenaConsulta2 = ""
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Plantilla
        ' BotonEliminarLinea
    Else   'Eliminar Plantilla
         BotonEliminar
    End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
       '  BotonModificarLinea
    Else   'Modificar Cabecera Oferta
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea False
    Else 'Añadir Cabecera de Ofertas
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
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    With Text1(Index)
        Select Case Index
            Case 0 'Codigo Plantilla
             
            Case 2 'Codigo Grupo Plantilla
                
            Case 4, 8
                PonerFormatoFecha Me.Text1(Index)
            Case 9, 10
                If Not PonerFormatoEntero(Me.Text1(Index)) Then Me.Text1(Index).Text = ""
            
            Case 14
                CadenaConsulta2 = ""
                If PonerFormatoEntero(Me.Text1(Index)) Then
                    CadenaConsulta2 = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Text1(Index).Text)
                    If CadenaConsulta2 = "" Then MsgBox "No existe el proveedor: " & Text1(Index).Text, vbExclamation
                End If
                Text2(Index).Text = CadenaConsulta2
                If CadenaConsulta2 = "" Then Text1(Index).Text = ""
           Case 15
                If Not PonerFormatoDecimal(Me.Text1(Index), 1) Then Me.Text1(Index).Text = ""
         End Select
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
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
        Case 9 'Mantenimiento Lineas
            mnLineas_Click
            
        Case 11
            'Insertar intecalando
            If Modo <> 5 Then Exit Sub
            If ModificaLineas <> 0 Then Exit Sub
            BotonAnyadirLinea True
            
            
        Case 13 'Mantenimiento Lineas
            Imprimir
        
        Case 14  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = 13 Then 'ENTER
'        KeyAscii = 0
'        SendKeys "{tab}"
'    ElseIf KeyAscii = 27 Then 'ESC
'        Select Case Modo
'            Case 0, 2: Unload Me
'            Case 1: cmdCancelar_Click 'Buscar
'            Case 5 'Lineas
'                If ModificaLineas = 0 Then PonerModo 2
'        End Select
'    End If
'    If Err.Number <> 0 Then Err.Clear
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
           
    '==============================
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    cboMarca.visible = B
    cmdRegresar.visible = (Modo = 5)
    
        'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        
        cmdRegresar.visible = Modo = 2
    End If
    
    
    
    BloquearCmb Me.cboTipo, Not B
    
'    For I = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(I).Enabled = b
'    Next I
'

    Me.imgFecha(0).Enabled = B
    Me.imgFecha(1).Enabled = B
    chkVistaPrevia.Enabled = (Modo <= 2)
     
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '===============================
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

    B = (Modo = 2) Or (Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    B = (Modo = 2)
    'Lineas
    Toolbar1.Buttons(9).Enabled = B
    Me.mnLineas.Enabled = B

    'intecalar
    B = Modo = 5
    Toolbar1.Buttons(11).Enabled = B
    Me.mnLineas.Enabled = B


    B = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not B Or (Modo = 5)
    Me.mnNuevo.Enabled = Not B Or (Modo = 5)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    Me.cboTipo.ListIndex = -1
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
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

    
    
    SQL = "select fecha, nomconcef,Ampliacion,baseimp from sflotasregistro ,sflotasconce"
    SQL = SQL & " WHERE sflotasregistro.CODCONCEF =sflotasconce.CODCONCEF AND codflota = "
    If enlaza Then
        SQL = SQL & DBSet(Text1(0).Text, "T") 'Data1.Recordset!codPlant
    Else
        SQL = SQL & "'#@DABZ'"
    End If
    SQL = SQL & " ORDER BY fecha "
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

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
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta2 = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    Me.cboTipo.ListIndex = 1
    cboMarca.Text = ""
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea(Intercalar As Boolean)
'Dim anc As Single
'
'    'Si no estaba modificando lineas salimos
'    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 2 Then Exit Sub
'
'    ModificaLineas = 1 'Ponemos Modo Añadir Linea
'
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerBotonCabecera False
'
'    If Intercalar Then
'        lblIndicador.Caption = "** INTERCALAR **"
'        If Not Data2.Recordset.EOF Then
'            LineaIntercalar2 = Data2.Recordset!numlinea
'        End If
'        txtAux(0).BackColor = vbRed
'    Else
'        lblIndicador.Caption = "INSERTAR"
'        txtAux(0).BackColor = vbWhite
'        LineaIntercalar2 = 0
'    End If
'    lblIndicador.Refresh
'
'
'
'    AnyadirLinea DataGrid1, Data2
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc
'    PonerFoco txtAux(0)
End Sub


Private Sub BotonMtoLineas()
On Error GoTo ErrorLineas
    
    
    Exit Sub  'No hay mantenimiento de lineas
    
    
    Screen.MousePointer = vbHourglass
    PonerModo (5)
    ModificaLineas = 0
   
    PonerBotonCabecera True
    CargaGrid True
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(1)
    Me.cboMarca.Text = Text1(5).Text
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not Data2.Recordset.EOF Then
        MsgBox "Tiene registros relacionados.", vbExclamation
        Exit Sub
    End If
    
    'Comprobar
    SQL = DevuelveDesdeBD(conAri, "count(*)", "advpartes", "codflota", CStr(Data1.Recordset!codflota), "T")
    If SQL = "" Then SQL = "0"
    If Val(SQL) > 0 Then
        MsgBox "Existen registros relacionados(" & SQL & ") en Partes de trabajo", vbExclamation
        Exit Sub
    End If
    
    'En euler asiganmos a la tarea la matricula
    SQL = DevuelveDesdeBD(conAri, "count(*)", "sreloj", "codflota", CStr(Data1.Recordset!codflota), "T")
    If SQL = "" Then SQL = "0"
    If Val(SQL) > 0 Then
        MsgBox "Existen registros relacionados(" & SQL & ") en trabajos realizados", vbExclamation
        Exit Sub
    End If
    
   
    
    
    SQL = "VEHICULO                 " & vbCrLf
    SQL = SQL & "----------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar el vehiculo:"
    SQL = SQL & vbCrLf & "Código : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Descripcion : " & Text1(1).Text
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Plantilla", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        If Data1.Recordset.EOF Then
            Eliminar = False
            Exit Function
        End If
        conn.BeginTrans

        
        'Lineas
        'conn.Execute "Delete  from slipla " & SQL
        
        
        'Cabeceras
        SQL = " WHERE codflota=" & DBSet(Data1.Recordset!codflota, "T")
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


Private Function DatosOk() As Boolean
Dim B As Boolean
On Error Resume Next

    DatosOk = False
    Text1(5).Text = Me.cboMarca.Text
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    If B And Modo = 3 Then
        'No dejo que piongan ni coma ni "|"
        If InStr(1, Text1(0).Text, ",") > 0 Then B = False
        If InStr(1, Text1(0).Text, "|") > 0 Then B = False
        If Not B Then MsgBox "Caracteres incorrectos( ,  | )", vbExclamation
    End If
    DatosOk = B
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scapla
    Cad = Cad & ParaGrid(Text1(0), 12, "Código")
    Cad = Cad & ParaGrid(Text1(1), 45, "Descripcion")
    
    Cad = Cad & "Tipo|sflotatipo|desctipflota|T||18·"
    
    tabla = NombreTabla & " LEFT JOIN sflotatipo ON " & NombreTabla & ".tipo=sflotatipo.tipflota"
    Titulo = "Flotas"
 
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim cadMen As String
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        cadMen = "No hay ningún registro en la tabla " & NombreTabla
        If Modo = 1 Then
            MsgBox cadMen & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox cadMen, vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
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
    If Text1(14).Text <> "" Then
        Text2(14).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Text1(14).Text)
    Else
        Text2(14).Text = ""
    End If
    CargaGrid True
      
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub






Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = Mid(ObtenerWhereCP, 7)
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        Indicador = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ObtenerWhereCP() As String
Dim SQL As String
    
    SQL = " WHERE codflota= " & DBSet(Text1(0).Text, "T")
    ObtenerWhereCP = SQL
End Function


'Private Sub txtAux_GotFocus(Index As Integer)
'    ConseguirFocoLin txtAux(Index)
'End Sub
'
'Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'        KEYdown KeyCode
'End Sub
'
'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'        KEYpress KeyAscii
'End Sub
'
'
'Private Sub txtAux_LostFocus(Index As Integer)
'
'    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    Select Case Index
'        Case 0 'Cod Articulo
'           txtAux(1).Text = PonerNombreDeCod(txtAux(0), 1, "sartic", "nomartic", "codartic", " Artículo ", "T")
'           If txtAux(1).Text = "" And txtAux(0).Text <> "" Then PonerFoco txtAux(0)
'        Case 2 'Cantidad
'            If txtAux(Index).Text <> "" Then
'                PonerFormatoDecimal txtAux(Index), 1 'Tipo 1: Decimal(12,2)
'                PonerFocoBtn Me.cmdAceptar
'            End If
'    End Select
'End Sub


'Private Function InsertarLinea() As Boolean
''Inserta un registro en la tabla de lineas de Plantilla: slipla
'Dim SQL As String
'Dim numlinea As String, vWhere As String
'
'    On Error GoTo EInsertarLinea
'
'    InsertarLinea = False
'    SQL = ""
'    If DatosOkLinea Then 'Lineas de Ofertas
'
'         If LineaIntercalar2 = 0 Then
'            'INSERCION NORMAL
'            vWhere = Mid(ObtenerWhereCP, 7)
'            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
'
'        Else
'            SQL = ObtenerWhereCP
'            SQL = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 1  " & SQL & " and numlinea >= " & LineaIntercalar2
'            SQL = SQL & " order by numlinea desc " 'Para que empieza por las ultimas
'            conn.Execute SQL
'            numlinea = LineaIntercalar2
'        End If
'
'
'
'        'Conseguir el siguiente numero de linea
'
'        SQL = "INSERT INTO " & NomTablaLineas
'        SQL = SQL & " (codplant, numlinea, codartic, cantidad) "
'        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & DBSet(txtAux(0).Text, "T") & ","
'        SQL = SQL & DBSet(txtAux(2).Text, "N") & ") "
'    End If
'
'    If SQL <> "" Then
'        conn.Execute SQL
'        InsertarLinea = True
'    End If
'    Exit Function
'EInsertarLinea:
'    MuestraError Err.Number, "Insertar Lineas Plantillas" & vbCrLf & Err.Description
'End Function
'
'
'Private Function DatosOkLinea() As Boolean
'Dim b As Boolean
'Dim vArtic As CArticulo
'Dim SQL As String
'
'    On Error GoTo EDatosOkLinea
'
'    DatosOkLinea = False
'    b = True
'
'    If txtAux(0).Text = "" Then
'        MsgBox "El campo Cod. Articulo no puede ser nulo.", vbExclamation
'        b = False
'        PonerFoco txtAux(0)
'        Exit Function
'    End If
'    'If Not b Then Exit Function
'
'    'Comprobar que existe el articulo seleccionado
'    Set vArtic = New CArticulo
'    If Not vArtic.Existe(txtAux(0).Text) Then
'        b = False
'        PonerFoco txtAux(0)
'    ElseIf ModificaLineas = 1 Then
'        'si existe miramos si ya hay una linea con ese artículo antes de insertar
'        SQL = "SELECT COUNT(*) FROM " & NomTablaLineas & ObtenerWhereCP & " AND codartic=" & DBSet(txtAux(0).Text, "T")
'        If RegistrosAListar(SQL) > 0 Then
'            SQL = "Ya existe una línea en la plantilla con el Artículo: " & txtAux(0).Text & vbCrLf & "¿Desea añadir la linea?"
'            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then b = False
'        End If
'    End If
'    Set vArtic = Nothing
'
'    DatosOkLinea = b
'EDatosOkLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Function
'
'
'Private Sub BotonModificarLinea()
''Modificar una linea
'Dim vWhere As String
'Dim anc As Single
'Dim i As Byte
'
'    On Error GoTo EModificarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Then Exit Sub '1= Insertar
'
'    If Data2.Recordset.EOF Then Exit Sub
'
'    'Si BLOQUEA REGISTRO
'    vWhere = Mid(ObtenerWhereCP, 7) & " and numlinea=" & Data2.Recordset!numlinea
'    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
'
'    DataGrid1.Enabled = False
'
'    ModificaLineas = 2 'Modificar
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc
'
'    'cargamos los datos
'    For i = 0 To txtAux.Count - 1
'        txtAux(i).Text = DataGrid1.Columns(i + 2).Text
'    Next i
'
'    PonerFoco txtAux(0)
'
'EModificarLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub
'
'
'Private Function ModificarLinea() As Boolean
''Modifica un registro en la tabla de Lineas Plantillas: slipla
'Dim SQL As String
'
'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    SQL = ""
'    If DatosOkLinea Then
'        SQL = "UPDATE " & NomTablaLineas & " Set codartic = " & DBSet(txtAux(0).Text, "T") & ", "
'        SQL = SQL & " cantidad = " & DBSet(txtAux(2).Text, "N")
'        SQL = SQL & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea
'    End If
'
'    If SQL <> "" Then
'        conn.Execute SQL
'        ModificarLinea = True
'    End If
'    Exit Function
'
'EModificarLinea:
'    MuestraError Err.Number, "Modificar Lineas Plantilla" & vbCrLf & Err.Description
'End Function
'
'
'
'Private Sub BotonEliminarLinea()
''Eliminar una linea De Mantenimiento. Tabla: slima1
'Dim SQL As String
'
'    On Error GoTo EEliminarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'    If Data2.Recordset.EOF Then Exit Sub
'
'    ModificaLineas = 3 'Eliminar
'    SQL = "¿Seguro que desea eliminar la línea de Plantilla?     " & vbCrLf
'    SQL = SQL & vbCrLf & "Plantilla: " & Text1(0).Text & " - " & Text1(1).Text
'    SQL = SQL & vbCrLf & "NumLinea: " & Data2.Recordset!numlinea
'    SQL = SQL & vbCrLf & "Articulo: " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
'
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'        'Hay que eliminar
'        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP
'        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
'        conn.Execute SQL
'        ModificaLineas = 0
'        CargaGrid True
'
'        CancelaADODC Me.Data2
'    End If
'    PonerFocoBtn Me.cmdRegresar
'
'EEliminarLinea:
'        Screen.MousePointer = vbDefault
'        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
'End Sub

Private Sub Imprimir()
    With frmImprimir
        .FormulaSeleccion = ""
        .OtrosParametros = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        .NumeroParametros = 2
    
    

        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = "Flotas"
        .Opcion = 3000   'VAN TODOS EN ESTE SACO
        .NombrePDF = "rFlota.rpt"
        .NombreRPT = .NombrePDF
        .ConSubInforme = False
        .MostrarTreeDesdeFuera = False
        .Show vbModal
    End With
End Sub


Private Sub CargaCboMarcas()
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select marca from sflotas group by 1 order by 1", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cboMarca.Clear
    While Not miRsAux.EOF
        If Not IsNull(miRsAux!marca) Then cboMarca.AddItem miRsAux!marca
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub ComprobarCombo()
    'Para que si el dato que se ha insetado de marca, NO existia, vuelva a cargarlo
    For NumRegElim = 0 To Me.cboMarca.ListCount - 1
        If Me.Text1(5).Text = cboMarca.List(NumRegElim) Then
            Exit For
        End If
    Next
    'NO estaba
    If NumRegElim >= Me.cboMarca.ListCount Then CargaCboMarcas
    
    
    
    
End Sub
