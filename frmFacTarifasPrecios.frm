VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTarifasPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifas de Artículos"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   ClipControls    =   0   'False
   Icon            =   "frmFacTarifasPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePro 
      Enabled         =   0   'False
      Height          =   705
      Left            =   3195
      TabIndex        =   39
      Top             =   900
      Visible         =   0   'False
      Width           =   7260
      Begin VB.Label lblPro 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   135
         TabIndex        =   40
         Top             =   225
         Width           =   7020
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   37
      Top             =   135
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   165
         TabIndex        =   38
         Top             =   180
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cambio Precios"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   135
      TabIndex        =   35
      Top             =   900
      Width           =   2985
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmFacTarifasPrecios.frx":000C
         Left            =   120
         List            =   "frmFacTarifasPrecios.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5175
      TabIndex        =   33
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   225
         TabIndex        =   34
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   31
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
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
      Left            =   8955
      TabIndex        =   30
      Top             =   270
      Width           =   1530
   End
   Begin VB.Frame FrameNuevo 
      Caption         =   "Valores Nuevos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1230
      Left            =   3915
      TabIndex        =   26
      Top             =   2745
      Width           =   6540
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
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Fecha Cambio|F|S|||slista|fechanue|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1350
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
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "Precio Caja Nuevo|N|S|0|999999.0000|slista|precion1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   720
         Width           =   1665
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
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   5
         Tag             =   "Precio Nuevo|N|S|0|999999.0000|slista|precionu|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   290
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Cambio"
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
         Left            =   3225
         TabIndex        =   29
         Top             =   720
         Width           =   1440
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4665
         Picture         =   "frmFacTarifasPrecios.frx":0050
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Caja"
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
         Left            =   165
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
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
         Left            =   165
         TabIndex        =   27
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Valores Actuales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1230
      Left            =   120
      TabIndex        =   23
      Top             =   2745
      Width           =   3705
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
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Precio Caja Actual|N|S|0|999999.0000|slista|precioa1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   720
         Width           =   1665
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
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "Precio Actual|N|N|0|999999.0000|slista|precioac|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   290
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Caja"
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
         Left            =   165
         TabIndex        =   25
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Precio"
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
         Left            =   165
         TabIndex        =   24
         Top             =   285
         Width           =   735
      End
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
      Index           =   1
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "Fecha Cambio"
      Top             =   5160
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
      Height          =   320
      Index           =   0
      Left            =   840
      MaxLength       =   4
      TabIndex        =   19
      Text            =   "NumLinea"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkPermiteDto 
      Caption         =   "Permite Descuento"
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
      Left            =   8235
      TabIndex        =   2
      Tag             =   "Permite Descuento|N|N|||slista|dtopermi||N|"
      Top             =   2250
      Width           =   2190
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
      Left            =   4920
      MaxLength       =   12
      TabIndex        =   22
      Text            =   "Precio Caja"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   3240
      MaxLength       =   12
      TabIndex        =   21
      Text            =   "Precio Unidad"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   8280
      TabIndex        =   8
      Top             =   7860
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
      Left            =   9435
      TabIndex        =   9
      Top             =   7860
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
      Left            =   9450
      TabIndex        =   10
      Top             =   7875
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   135
      TabIndex        =   17
      Top             =   7785
      Width           =   2175
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
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1995
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
      Left            =   3675
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   1680
      Width           =   6705
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   2190
      Width           =   6075
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
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "Cod. Tarifa|N|N|0|999|slista|codlista|000|S|"
      Text            =   "Text1"
      Top             =   2190
      Width           =   585
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
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Articulo|T1|N|||slista|codartic||S|"
      Text            =   "Text1"
      Top             =   1680
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1320
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   1200
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacTarifasPrecios.frx":00DB
      Height          =   3465
      Left            =   135
      TabIndex        =   11
      Top             =   4185
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6112
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
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
      Left            =   1125
      Picture         =   "frmFacTarifasPrecios.frx":00F0
      Tag             =   "-1"
      ToolTipText     =   "Buscar artículo"
      Top             =   1710
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1140
      ToolTipText     =   "Buscar tarifa"
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Tarifa"
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
      Left            =   210
      TabIndex        =   14
      Top             =   2190
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "Artículo"
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
      Left            =   210
      TabIndex        =   13
      Top             =   1710
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
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
      Left            =   2520
      TabIndex        =   12
      Top             =   7965
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
   Begin VB.Menu mnFiltro 
      Caption         =   "&Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "Sin filtro"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnSeleProv 
         Caption         =   "Seleccionar proveedor"
      End
   End
End
Attribute VB_Name = "frmFacTarifasPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public codArtic As String

Private WithEvents frmB As frmBasico2 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas  'Form Mantenimiento Tarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmPro As frmBasico2 '%=%=frmComProveedores
Attribute frmPro.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer


Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean



Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    ClikFiltro (cboFiltro.ListIndex = 0)
End Sub

Private Sub cboFiltro_Validate(Cancel As Boolean)
    If PrimeraVez Then Exit Sub
    ClikFiltro (cboFiltro.ListIndex = 0)
End Sub



Private Sub chkPermiteDto_GotFocus()
'    chkPermiteDto.TabIndex = 2
    ConseguirfocoChk Modo
End Sub

Private Sub chkPermiteDto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPermiteDto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PonerFoco Text1(2) 'ENTER
    KEYpress KeyAscii
End Sub

Private Sub chkPermiteDto_LostFocus()
'    PonerFoco Text1(5)
'    chkPermiteDto.TabIndex = 20
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'Busqueda
            HacerBusqueda
        Case 3 'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4 'Modificar
           If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    
                    'si se ha modificado el precio actual
                    If CCur(Me.Data1.Recordset!precioac) <> CCur(Me.Text1(2).Text) Then
                        'si la tarifa modificada coincide con la de parametros aplicacion
                        If CInt(Me.Data1.Recordset!codlista) = vParamAplic.CodTarifa Then
                            'actualizar el PVP del articulo a partir del precio tarifa nuevo
                            'y quitandole el margen de la tarifa correspondiente
                            '---------------------------------------------------------
                            '- bloquear el articulo a modificar
                            If BloquearArticulo Then
                                Screen.MousePointer = vbDefault
                               ' MsgBox "Se va a actualizar el PVP del artículo.", vbInformation
                                '- actualizar su PVP
                                Screen.MousePointer = vbHourglass
                                ActualizarPVPArticulo
                                
                                
                                'Abril 2014
                                'Parametro para actualizar el precio especial
                                If vParamAplic.ActualizaPrecioEspecial Then ActualizarPrecioEspecialGenerico Data1.Recordset!codArtic, CCur(Me.Text1(2).Text), False, ""
        

                                
                                
                                
                                TerminaBloquear
                                
                                
                                
                                
                            End If
                        End If
                    End If
                    If codArtic <> "" Then
                        Unload Me
                    Else
                        PosicionarData
                    End If
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
                'PonerOpcionesMenu
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

'Private Sub cmdRegresar_Click()
''Este es el boton Cabecera
'Dim cad As String
'Dim Indicador As String
'
'    'Quitar lineas y volver a la cabecera
'    If Modo = 5 Then 'modo 5: Lineas Articulos x Almacen
'        DataGrid1.ClearFields
'        cad = "(codmovim=" & Val(Text1(0).Text) & ")"
'        If SituarData(Data1, cad, Indicador) Then
'            PonerModo 2
'            lblIndicador.Caption = Indicador
'            Me.Toolbar1.Buttons(9).Enabled = True
'            Me.Toolbar1.Buttons(10).Enabled = True
'        End If
'    End If
'End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
Dim C As String

    If kCampo < 0 Then
        
        'Es la primera vez
        kCampo = 0
        limpiar Me
        Text1(0).Text = codArtic
        C = ""
        If codArtic <> "" Then
            C = DevuelveDesdeBD(conAri, " distinct codartic", "slista", "codartic", codArtic, "T")
        End If
        If C <> "" Then
            HacerBusqueda
            If codArtic <> "" And Modo = 2 Then
                If BLOQUEADesdeFormulario(Me) Then BotonModificar
            
            End If
        Else
            If codArtic <> "" Then
                BotonAnyadir
                Text1(0).Text = codArtic
                C = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", codArtic, "T")
                Me.Text2(0).Text = C
                PonerFoco Text1(1)
            End If
        End If
    End If
    
    If PrimeraVez Then
        Me.cboFiltro.ListIndex = 0
        PrimeraVez = False
    End If
    
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    
    'Icono de busqueda
    Me.imgBuscar(1).Picture = Me.imgBuscar(0).Picture
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        .Buttons(8).Image = 16  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 21 'Para cambiar precios
        .Buttons(2).Image = 42 'incre % precio
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
    
    CargaFiltros
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    mnSeleProv.Tag = "" 'bus por proveedor
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "slista" 'Tabla Tarifas Precios de Articulos
    Ordenacion = " ORDER BY slista.codartic"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE false " 'No recupera datos"

    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    CargaGrid (Modo = 2)
    kCampo = -1
    
    
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim SQL As String

    On Error GoTo ECarga

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False

    DataGrid1.RowHeight = 350

    DataGrid1.Columns(0).visible = False 'Cod. Articulo
    DataGrid1.Columns(1).visible = False 'Cod. Lista
    i = 2
    
    'Numero Linea
    DataGrid1.Columns(i).Caption = "Línea"
    DataGrid1.Columns(i).Width = 1750
    
    'Fecha Cambio
    DataGrid1.Columns(i + 1).Caption = "Fecha Cambio"
    DataGrid1.Columns(i + 1).Width = 2000
    
    'Precio Unidad
    DataGrid1.Columns(i + 2).Caption = "Precio Unidad"
    DataGrid1.Columns(i + 2).Width = 3000
    DataGrid1.Columns(i + 2).Alignment = dbgRight
    DataGrid1.Columns(i + 2).NumberFormat = FormatoPrecio & " "
    
    'Precio Caja
    DataGrid1.Columns(i + 3).Caption = "Precio Caja"
    DataGrid1.Columns(i + 3).Width = 3000
    DataGrid1.Columns(i + 3).Alignment = dbgRight
    DataGrid1.Columns(i + 3).NumberFormat = FormatoPrecio & " "
    
    DataGrid1.ScrollBars = dbgAutomatic
       
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    codArtic = ""
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

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
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(4).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    NombreTabla = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Tarifas
    Text1(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0 'Codigo Articulo
            Set frmA = New frmBasico2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abre en modo Busqueda
            AyudaArticulos frmA, Text1(Index)
            Set frmA = Nothing
        Case 1  'Cod. Tarifa
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Indice = 4
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub


Private Sub mnFiltro1_Click()
    ClikFiltro True
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub ClikFiltro(Quitar As Boolean)
Dim Aux As String
    If Modo > 2 Then Exit Sub

    If Quitar Then
        
        mnFiltro1.Checked = True
        Me.mnSeleProv.Checked = False
        Me.mnSeleProv.Caption = "Seleccione proveedor"
        mnSeleProv.Tag = ""
        lblPro.Caption = ""
        FramePro.visible = False
    Else

        Aux = NombreTabla
        NombreTabla = ""
'        Set frmPro = New frmComProveedores
'        frmPro.DatosADevolverBusqueda = "0|1|"
'        frmPro.Show vbModal
        Set frmPro = New frmBasico2
        AyudaProveedores frmPro
        Set frmPro = Nothing
        
        If NombreTabla <> "" Then
            
            
            
            mnFiltro1.Checked = False
            Me.mnSeleProv.Checked = True
            Me.mnSeleProv.Caption = RecuperaValor(NombreTabla, 1) & " - " & RecuperaValor(NombreTabla, 2)
            Me.mnSeleProv.Tag = RecuperaValor(NombreTabla, 1)
            lblPro.Caption = RecuperaValor(NombreTabla, 2)
            FramePro.visible = True
        Else
            NombreTabla = Aux
            Exit Sub   'ha cancelado
        End If
        NombreTabla = Aux
    End If
    Modo = 2
    BotonBuscar
End Sub


Private Sub mnSeleProv_Click()
    ClikFiltro False
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'articulo
            Case 1: KEYBusqueda KeyAscii, 1 'tarifa
            Case 4: KEYFecha KeyAscii, 0 'fecha cambio
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim SQL As String
Dim devuelve As String
Dim campo As String, TipoCampo As String
Dim Codigo As String
Dim tabla As String
Dim Titulo As String
Dim index2 As Integer 'Indice del control Text2
Dim indImgP As Integer 'Indice del control ImgBuscar (Prismaticos)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    Select Case Index
        Case 0 'Codigo Articulo
            campo = "nomartic"
            TipoCampo = "T"
            tabla = "sartic"
            Codigo = "codartic"
            Titulo = "Articulos"
            indImgP = Index 'Imagen Prismaticos
            index2 = Index  'Indice para Text2
            
        Case 1 'Codigo Tarifa
            campo = "nomlista"
            TipoCampo = "N"
            tabla = "starif"
            Codigo = "codlista"
            Titulo = "Tarifas"
            indImgP = Index
            index2 = Index
            Text1(Index).Text = Format(Text1(Index).Text, "000")
                
        Case 2, 3, 5, 6 'Precios Actuales y Nuevos
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" And Modo <> 1 Then
                If PonerFormatoDecimal(Text1(Index), 2) And (Index = 5 Or Index = 6) Then
                    'fecha de cambio
                    If Text1(4).Text = "" Then Text1(4).Text = Format(Now, "dd/mm/yyyy")
                End If
            End If
            
        Case 4 'Fecha
            If Modo <> 1 And Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            PonerFocoBtn Me.cmdAceptar
    End Select

    If (Index = 0 Or Index = 1) And Modo <> 1 Then
    'Indices 0 al 1 del textbox Text1 son busquedas a otras tablas
    'Comprobar los cambios para actualizar la descripcion de Text2
        If Text1(Index).Text = "" Then
             Text2(index2).Text = ""
             Exit Sub
        Else
            'BD 1: conexion a BD Ariges
            If TipoCampo = "N" Then
                If Not PonerFormatoEntero(Text1(Index)) Then
                    Text2(index2).Text = ""
                    Exit Sub
                End If
            End If
            SQL = DevuelveDesdeBD(conAri, campo, tabla, Codigo, Text1(Index).Text, TipoCampo)
            Text2(index2).Text = SQL
            If SQL = "" Then 'No existe
                devuelve = "No existe el código de " & Titulo & vbCrLf
                devuelve = devuelve & "Código: " & Text1(Index).Text
                MsgBox devuelve, vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Nuevo
                BotonAnyadir
        Case 2  'Modificar
                If BLOQUEADesdeFormulario(Me) Then BotonModificar
        Case 3 'Eliminar
                BotonEliminar
        Case 5 'Busqueda
            BotonBuscar
        Case 6 'Ver Todos
            BotonVerTodos
        Case 8 'Imprimir
'            BotonImprimir
            AbrirListado (28) '28: Informe Tarifas de Articulos
    End Select
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            frmListado2.Opcion = 34
            frmListado2.Show vbModal
            BotonBuscar
        Case 2
            'Nuevo aumento % precio
            If vUsu.Nivel > 1 Then Exit Sub
            frmListado3.Opcion = 62
            frmListado3.Show vbModal
            BotonBuscar
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    
     'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1


    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
           
    '--------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B And Modo <> 4 'Si modificar no activado pq son claves ajenas
    Next i
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B 'And Modo <> 3 'Si es insertar se modifica
    Next i
    
'    Me.chkPermiteDto.Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    
    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Modo", Err.Description
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerModoOpcionesMenu()
Dim B As Boolean

    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    '-------------------------------------
    'b = (Modo >= 3)
    B = Not (Modo = 0 Or Modo = 2)
    'Insertar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnNuevo.Enabled = Not B
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub

Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPermiteDto.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
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
Dim tabla As String
    
    tabla = "slist1"

    SQL = "SELECT * FROM " & tabla
    If enlaza Then
        SQL = SQL & " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND codlista=" & Data1.Recordset!codlista
    Else
        SQL = SQL & " WHERE false "
    End If
    SQL = SQL & " ORDER BY " & tabla & ".numlinea desc"
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
Dim Aux As String

'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    Aux = ""
    If Me.mnSeleProv.Tag <> "" Then
        Aux = "sartic.codprove = " & Me.mnSeleProv.Tag
    End If
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia Aux
    Else
        If Aux <> "" Then Aux = " AND " & Aux
        CadenaConsulta = "select * from " & NombreTabla & ",sartic WHERE " & NombreTabla & ".codartic=sartic.codartic  " & Aux & " " & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    
    PonerFoco Text1(0)
    
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1 y campo2 es clave primaria, NO se puede modificar
'    If Text1(4).Text = "" Then Text1(4).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Tarifas de Precios." & vbCrLf
    SQL = SQL & "---------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar la Tarifa del Articulo:"
    SQL = SQL & vbCrLf & "Cod. Artic :  " & Text1(0).Text
    SQL = SQL & vbCrLf & "Cod. Tarif :  " & Text1(1).Text
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
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
            MuestraError Err.Number, "Eliminar Tarifa de Articulo", Err.Description
            'MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        SQL = " WHERE  codartic=" & DBSet(Data1.Recordset!codArtic, "T") & ""
        SQL = SQL & " AND codlista=" & Val(Data1.Recordset!codlista)
        
        'Lineas
        conn.Execute "Delete  from slist1 " & SQL
        
        'Cabeceras
        conn.Execute "Delete  from slista " & SQL
                      
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

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    'Comprobar que si hay valores nuevos, la fecha de cambio no es nulo
    If Not EsVacio(Text1(5)) Or Not EsVacio(Text1(6)) Then
        B = (Not EsVacio(Text1(4)))
    End If
    If Not B Then
        MsgBox "La Fecha de Cambio debe tener valor.", vbInformation
        Exit Function
    End If
    
    'Comprobar que si no hay valores nuevos no haya fecha de Cambio
    If EsVacio(Text1(5)) And EsVacio(Text1(6)) Then
        B = (EsVacio(Text1(4)))
    End If
    If Not B Then
        MsgBox "No hay valores nuevos para la fecha de cambio", vbInformation
        Exit Function
    End If
    
    'si se modifica el precio actual no hay fecha de cambio ni precio nuevo
    If Modo = 4 Then
        If CCur(Me.Data1.Recordset!precioac) <> CCur(Me.Text1(2).Text) Then
            B = EsVacio(Text1(5)) And EsVacio(Text1(6)) And EsVacio(Text1(4))
            If Not B Then
                MsgBox "Si se modifican precios actuales, los precios nuevos y fecha cambio no deben tener valor.", vbInformation
                Exit Function
            End If
        End If
    End If
    
    
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
'    cad = cad & ParaGrid(Text1(0), 18, "Cod Artic.")
'    cad = cad & "Desc. Artic|sartic|nomartic|T||48·"
'    cad = cad & ParaGrid(Text1(1), 7, "Tar.")
'    cad = cad & "Tarifa|starif|nomlista|T||12·"
'    cad = cad & ParaGrid(Text1(2), 14, "Precio")
'    tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
'    tabla = tabla & " LEFT JOIN starif ON " & NombreTabla & ".codlista=starif.codlista"
'
''    tabla = "slista"
'    Titulo = "Tarifas de Artículos"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|2|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri
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
    
    AyudaTarifaPrecios frmB, Text1(0), cadB

    Set frmB = Nothing
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If Me.mnSeleProv.Tag <> "" Then
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & " sartic.codprove = " & mnSeleProv.Tag
    End If
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & ",sartic WHERE " & NombreTabla & ".codartic=sartic.codartic AND " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
    'Poner el nombre del cod. Articulo
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sartic", "nomartic")
    'Poner el nombre del cod. tarifa
    Text2(1).Text = PonerNombreDeCod(Text1(1), 1, "starif", "nomlista")
    
    'Si los campos de precios nuevos son cero mostrar cadena vacia
    If Text1(5).Text <> "" Then
        If Text1(5).Text = 0 Then Text1(5).Text = ""
    End If
    If Text1(6).Text <> "" Then
        If Text1(6).Text = 0 Then Text1(6).Text = ""
    End If
    
    BloquearChecks Me, Modo
    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub BotonActualizar()
'Actualizar Tarifas de Articulos
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningúna Tarifa para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
   
    SQL = "Actualización Tarifas de Artículos." & vbCrLf
    SQL = SQL & "-----------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Actualizar la Tarifa del Articulo:"
    SQL = SQL & vbCrLf & " Cod. Artic. :  " & Data1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & " Cod. Tarifa :  " & CStr(Format(Data1.Recordset.Fields(1), "000"))
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
        Exit Sub
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarTarifa Then
        SituarDataTrasEliminar Data1, NumRegElim
    End If
End Sub


Private Function ActualizarTarifa() As Boolean
Dim Donde As String
Dim SQL As String
Dim bol As Boolean
On Error GoTo EActualizarTarifa
    
   
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarLaTarifa(Donde)

EActualizarTarifa:
        If Err.Number <> 0 Then
            SQL = "Actualizar Tarifa." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
'            If OpcionActualizar = 1 Then
                MuestraError Err.Number, SQL, Err.Description
'            Else
'                SQL = Donde & " -> " & Err.Description
'                SQL = Mid(SQL, 1, 200)
'                InsertaError SQL
'            End If
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            ActualizarTarifa = True
        Else
            conn.RollbackTrans
        End If
End Function


Private Function ActualizarLaTarifa(ByRef ADonde As String) As Boolean

    ActualizarLaTarifa = False
    
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Tarifas de Almacen"
    If Not InsertarLineasHistorico Then Exit Function
    
    'Modificamos en cabeceras de Tarifas
    ADonde = "Modificando datos en cabecera de Tarifas Precios"
    If Not ModificarCabecera Then Exit Function
    
    ActualizarLaTarifa = True
End Function


Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim SQL As String
On Error Resume Next

    SQL = "UPDATE slista SET precioac=precionu, precioa1=precion1, fechanue=null, precionu=0, precion1=0"
    SQL = SQL & " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND codlista=" & Data1.Recordset!codlista
   
    conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
    Else
        ModificarCabecera = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim SQL As String
Dim NumF As String
On Error Resume Next

    'Obtenemos la siguiente numero de linea de tarifa
    SQL = "codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND codlista=" & Data1.Recordset!codlista
    NumF = SugerirCodigoSiguienteStr("slist1", "numlinea", SQL)

    SQL = "INSERT INTO slist1 (codartic, codlista, numlinea, fechacam, precioac, precioa1)"
    SQL = SQL & " VALUES (" & DBSet(Data1.Recordset.Fields(0).Value, "T") & ", " & Data1.Recordset.Fields(1).Value & ", "
    SQL = SQL & NumF & ", " & DBSet(Text1(4).Text, "F") & ", "
    SQL = SQL & DBSet(Data1.Recordset!precioac, "N") & ", " & DBSet(Data1.Recordset!precioa1, "N") & ") "
    conn.Execute SQL
        
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Sub BotonImprimir()

        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(codartic=" & DBSet(Text1(0).Text, "T") & " AND codlista=" & Text1(1).Text & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
'    If SituarData2(Data1, Text1(0).Text, Text1(1).Text, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'                        LimpiarCampos
        PonerModo 0
    End If
End Sub



Private Function ActualizarPVPArticulo() As Boolean
Dim cTar As CTarifaArt
Dim margen As Currency
Dim newPrecio As Currency
Dim SQL As String
    
    On Error GoTo ErrActPVP
    
    'DEBERIAMOS ACTUALIZAR LA FECHA de actualizacion del PVP
    
    Set cTar = New CTarifaArt
    If cTar.LeerDatos(Text1(0).Text, CInt(Text1(1).Text)) Then
        margen = Round2(cTar.MargenComercial / 100, 4)
        newPrecio = Round2((cTar.PrecioActual / (margen + 1)), 4)
        If newPrecio > 0 Then
            SQL = "UPDATE sartic SET preciove=" & DBSet(newPrecio, "N")
            SQL = SQL & " WHERE codartic=" & DBSet(Text1(0).Text, "T")
            conn.Execute SQL
        End If
    End If
    Set cTar = Nothing
    Exit Function
    
ErrActPVP:
    MuestraError Err.Number, "Actualizar precio PVP del articulo.", Err.Description
End Function


Private Function BloquearArticulo() As Boolean
Dim cadWhere As String
    cadWhere = "codartic=" & DBSet(Text1(0), "T")
    BloquearArticulo = BloqueaRegistro("sartic", cadWhere)
End Function

Private Sub CargaFiltros()
Dim Aux As String
    
    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Seleccionar Proveedor "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
