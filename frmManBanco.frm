VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos Propios"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15000
   Icon            =   "frmManBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   40
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   41
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
      Height          =   300
      Left            =   12285
      TabIndex        =   39
      Top             =   240
      Width           =   1605
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
      Index           =   18
      Left            =   4230
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "IBAN|T|N|||banpropi|iban|||"
      Top             =   4920
      Width           =   555
   End
   Begin VB.TextBox txtAux 
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
      Index           =   17
      Left            =   13980
      MaxLength       =   3
      TabIndex        =   13
      Tag             =   "Código Ordenante|T|S|||banpropi|codorden34|||"
      Text            =   "1234567890"
      Top             =   3090
      Width           =   735
   End
   Begin VB.TextBox text2 
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
      Index           =   12
      Left            =   11775
      TabIndex        =   37
      Top             =   3540
      Width           =   2940
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Contacto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2415
      Left            =   8895
      TabIndex        =   32
      Top             =   3975
      Width           =   5865
      Begin VB.TextBox txtAux 
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
         Index           =   16
         Left            =   120
         MaxLength       =   40
         TabIndex        =   18
         Tag             =   "Mail|T|S|||banpropi|maibanpr|||"
         Top             =   1815
         Width           =   5625
      End
      Begin VB.TextBox txtAux 
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "Web|T|S|||banpropi|wwwbanpr|||"
         Top             =   1215
         Width           =   5610
      End
      Begin VB.TextBox txtAux 
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Teléfono|T|S|||banpropi|telbanpr|||"
         Top             =   615
         Width           =   2010
      End
      Begin VB.TextBox txtAux 
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
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Fax|T|S|||banpropi|faxbanpr|||"
         Top             =   615
         Width           =   2250
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   0
         Left            =   600
         Top             =   1575
         Width           =   240
      End
      Begin VB.Image imgWeb 
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmManBanco.frx":000C
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Mail"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Web"
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
         Left            =   120
         TabIndex        =   35
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
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
         Left            =   120
         TabIndex        =   34
         Top             =   375
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "Fax"
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
         Left            =   3495
         TabIndex        =   33
         Top             =   375
         Width           =   810
      End
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Cta.Contable|T|S|||banpropi|codmacta|||"
      Top             =   3540
      Width           =   1260
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
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Cuenta|T|N|||banpropi|cuentaba|||"
      Top             =   4920
      Width           =   555
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
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "D.C.|T|N|||banpropi|digcontr|||"
      Top             =   4920
      Width           =   555
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "Persona Contacto|T|S|||banpropi|perbanpr|||"
      Top             =   2640
      Width           =   4230
   End
   Begin VB.TextBox txtAux 
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
      Index           =   11
      Left            =   10485
      MaxLength       =   3
      TabIndex        =   12
      Tag             =   "Sufijo|T|S|||banpropi|sufijoem|||"
      Top             =   3090
      Width           =   780
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "Código Postal|T|S|||banpropi|codpopr|||"
      Top             =   1290
      Width           =   1140
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Entidad|N|N|||banpropi|codbanco|0000||"
      Top             =   4920
      Width           =   555
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
      Index           =   0
      Left            =   240
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "Código Propio|N|N|1|99|banpropi|codbanpr|00|S|"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Provincia|T|S|||banpropi|probanpr|||"
      Top             =   2190
      Width           =   4230
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "Población|T|S|||banpropi|pobbanpr|||"
      Top             =   1740
      Width           =   4230
   End
   Begin VB.TextBox txtAux 
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
      Left            =   10485
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Domicilio|T|S|||banpropi|dombanpr|||"
      Top             =   840
      Width           =   4200
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
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Sucursal|N|N|||banpropi|codsucur|0000||"
      Top             =   4920
      Width           =   555
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
      Left            =   12450
      TabIndex        =   19
      Top             =   6585
      Visible         =   0   'False
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
      Left            =   13725
      TabIndex        =   20
      Top             =   6570
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre Banco|T|N|||banpropi|nombanpr|||"
      Top             =   4920
      Width           =   795
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManBanco.frx":0596
      Height          =   5625
      Left            =   150
      TabIndex        =   23
      Top             =   840
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   9922
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      Left            =   13680
      TabIndex        =   30
      Top             =   6585
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   180
      TabIndex        =   21
      Top             =   6660
      Width           =   2385
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
         Height          =   255
         Left            =   45
         TabIndex        =   22
         Top             =   180
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   14325
      TabIndex        =   42
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Sufijo Transferencia"
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
      Left            =   11835
      TabIndex        =   38
      Top             =   3105
      Width           =   2160
   End
   Begin VB.Label Label4 
      Caption         =   "Cta.Contable"
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
      Left            =   8895
      TabIndex        =   31
      Top             =   3540
      Width           =   1305
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   9660
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3570
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Contacto"
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
      Index           =   0
      Left            =   8895
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Sufijo"
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
      Left            =   8895
      TabIndex        =   28
      Top             =   3090
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
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
      Index           =   7
      Left            =   8895
      TabIndex        =   27
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Provincia"
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
      Left            =   8895
      TabIndex        =   26
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
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
      Index           =   9
      Left            =   8895
      TabIndex        =   25
      Top             =   1740
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
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
      Index           =   8
      Left            =   8895
      TabIndex        =   24
      Top             =   1290
      Width           =   1020
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmManBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO (Se lo copia)                          +-+-
' +-+- Menú: Bancos Propios (con un par)                    +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String

' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String

' ### [Monica] 08/09/2006
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos


Private kCampo As Integer

Dim BuscaChekc As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim b As Boolean
Dim i As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(18).visible = Not b
    txtAux(2).visible = Not b
    txtAux(3).visible = Not b
    txtAux(4).visible = Not b
    txtAux(5).visible = Not b
    ' **************************************************
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    For i = 6 To 17
        BloquearTxt txtAux(i), b
    Next i
    
    ' ********************************************************

    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'codbanpr
    
    BloquearImgBuscar Me, Modo

End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim b As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    b = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b

    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim i As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("banpropi", "codbanpr")
        'NumF = SugerirCodigoSiguienteStr("sdexpgrp", "codsupdt", "codempre=" & vSesion.Empresa)
        'NumF = ""
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
    For i = 1 To 17
        txtAux(i).Text = ""
    Next i
    ' cuenta CCC
    txtAux(18).Text = "ES00"
    txtAux(2).Text = "0"
    txtAux(3).Text = "0"
    txtAux(4).Text = ""
    txtAux(5).Text = "0"
    ' **************************************************

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim i As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    CargaGrid "codbanpr = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    For i = 0 To 17
        txtAux(i).Text = ""
    Next i

    LLamaLineas DataGrid1.Top + 240, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(0)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top  'DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(18).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = ComprobarCero(Trim(DataGrid1.Columns(3).Text))
    txtAux(3).Text = ComprobarCero(Trim(DataGrid1.Columns(4).Text))
    txtAux(4).Text = DataGrid1.Columns(5).Text
    txtAux(5).Text = DataGrid1.Columns(6).Text
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For i = 0 To 5
        txtAux(i).Top = alto
    Next i
    txtAux(18).Top = alto
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    SQL = "¿Seguro que desea eliminar el Banco Propio?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from banpropi where codbanpr = " & adodc1.Recordset!codBanpr
        
        conn.Execute SQL
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "RESULTADO BUSQUEDA"
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cmdAceptar_Click()
Dim i As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codbanpr = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    i = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Move i - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'    Else
'        lblIndicador.Caption = ""
'    End If
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    ' *** adrede: per a tornar el TipoSuplem ***
    ' cad = cad & TipoSuplem & "|"
    ' ******************************************
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer

'    '******* repasar si n'hi ha botó d'imprimir o no******
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Tots
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        'el 9 i el 10 son separadors
'        .Buttons(11).Image = 10  'Imprimir
'        .Buttons(12).Image = 11  'Eixir
'    End With
    '*****************************************************
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With


' ### [Monica] 08/09/2006
    'IMAGES para busqueda
   For i = 0 To Me.imgBuscar.Count - 1
       Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
   Next i

    'carga IMAGES de mail
    For i = 0 To Me.imgMail.Count - 1
        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i

    chkVistaPrevia.Value = CheckValueLeer(Name)
    ' *** SI N'HI HAN COMBOS ***
    ' CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT codbanpr, nombanpr, iban, codbanco, codsucur, "
    CadenaConsulta = CadenaConsulta & "digcontr, cuentaba, dombanpr, codpopr, "
    CadenaConsulta = CadenaConsulta & "pobbanpr, probanpr, perbanpr, sufijoem, "
    CadenaConsulta = CadenaConsulta & "codmacta, telbanpr, faxbanpr, wwwbanpr, maibanpr, codorden34 "
    CadenaConsulta = CadenaConsulta & " FROM banpropi "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    'CargaForaGrid
    ' *********************************************
    
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

' ### [Monica] 08/09/2006
Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 12
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
          
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1

End Sub

Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If txtAux(16).Text <> "" Then
            LanzaMailGnral txtAux(16).Text
        End If
    End If
End Sub

Private Sub imgWeb_Click(Index As Integer)
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If LanzaHomeGnral(txtAux(15).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    
    'Prepara para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim i As Integer
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    ' *** si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
    ' `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
    If vSQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSQL  ' ### [Monica] 08/09/2006: antes habia AND
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    SQL = SQL & " ORDER BY codbanpr"
    '**************************************************************++
    
'    adodc1.RecordSource = SQL
'    adodc1.CursorType = adOpenDynamic
'    adodc1.LockType = adLockOptimistic
'    DataGrid1.ScrollBars = dbgNone
'    adodc1.Refresh
'    Set DataGrid1.DataSource = adodc1 ' per a que no ixca l'error de "la fila actual no está disponible"
       
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
       
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Código|750|;S|txtAux(1)|T|Denominación|2657|;S|txtAux(18)|T|IBAN|750|;S|txtAux(2)|T|Entidad|900|;"
    tots = tots & "S|txtAux(3)|T|Sucursal|1000|;S|txtAux(4)|T|D.C.|550|;"
    tots = tots & "S|txtAux(5)|T|Cuenta|1400|;"
    
    For i = 1 To 12
        tots = tots & "N||||0|;"
    Next i
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    'DataGrid1.Columns(1).Alignment = dbgRight
    'DataGrid1.Columns(2).Alignment = dbgRight
    DataGrid1.Columns(5).Alignment = dbgRight
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0, 2, 3
            PonerFormatoEntero txtAux(Index)
        
        Case 1
            If txtAux(Index).Text <> "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 12 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            text2(Index).Text = PonerNombreCuenta(txtAux(Index), Modo)

'       Case 6, 7 'dates
'           If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
            
'       Case 8 'Importe
'           If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1 'tipo1: Decimal(12,2)
            
'       Case 9
'           PonerFormatoDecimal txtAux(Index), 4 'tipo 4: Decimal(5,2)
        Case 18 ' codigo de iban
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Then
        Dim cta As String
        Dim CC As String
        If txtAux(2).Text <> "" And txtAux(3).Text <> "" And txtAux(4).Text <> "" And txtAux(5).Text <> "" Then
            
            cta = Format(txtAux(2).Text, "0000") & Format(txtAux(3).Text, "0000") & Format(txtAux(4).Text, "00") & Format(txtAux(5).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If txtAux(18).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then txtAux(18).Text = "ES" & cta
                Else
                    CC = CStr(Mid(txtAux(18).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(txtAux(18).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
                
            End If
        End If
    End If

    '**************************************************************************
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
' *** només per ad este manteniment ***
Dim Rs As Recordset
Dim Cad As String
Dim cta As String
Dim cadMen As String

'Dim exped As String
' *************************************

    b = CompForm(Me)
    If Not b Then Exit Function


    If b And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
        Datos = DevuelveDesdeBD("codbanpr", "banpropi", "codbanpr", txtAux(0).Text, "N")
'       Datos = DevuelveDesdeBDNew(1, "sdexpgrp", "codsupdt", "codsupdt", txtAux(1).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Banco: " & txtAux(0).Text, vbExclamation
            DatosOk = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer atres comprovacions ***
        
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If txtAux(2).Text = "" Or txtAux(3).Text = "" Or txtAux(4).Text = "" Or txtAux(5).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            txtAux(18).Text = ""
            txtAux(2).Text = ""
            txtAux(3).Text = ""
            txtAux(4).Text = ""
            txtAux(5).Text = ""
        Else
            cta = Format(txtAux(2).Text, "0000") & Format(txtAux(3).Text, "0000") & Format(txtAux(4).Text, "00") & Format(txtAux(5).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El banco no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del banco no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco txtAux(2)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.txtAux(18).Text <> "" Then BuscaChekc = Mid(txtAux(18).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.txtAux(18).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.txtAux(18).Text = BuscaChekc & cta
                    Else
                        If Mid(txtAux(18).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.txtAux(18).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco txtAux(18)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
        
        
    ' *********************************************

    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
    PonerContRegIndicador
    
End Sub

Private Sub CargaForaGrid()
        If DataGrid1.Columns.Count <= 2 Then Exit Sub
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        txtAux(6) = DataGrid1.Columns(7).Text
        txtAux(7) = DataGrid1.Columns(8).Text
        txtAux(8) = DataGrid1.Columns(9).Text
        txtAux(9) = DataGrid1.Columns(10).Text
        txtAux(10) = DataGrid1.Columns(11).Text
        txtAux(11) = DataGrid1.Columns(12).Text
        txtAux(12) = DataGrid1.Columns(13).Text
        txtAux(13) = DataGrid1.Columns(14).Text
        txtAux(14) = DataGrid1.Columns(15).Text
        txtAux(15) = DataGrid1.Columns(16).Text
        txtAux(16) = DataGrid1.Columns(17).Text
        txtAux(17) = DataGrid1.Columns(18).Text

        ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
        text2(12).Text = PonerNombreCuenta(txtAux(12), Modo)
        
        'txtAux2(4).Text = PonerNombreDeCod(txtAux(4), "poblacio", "despobla", "codpobla", "N")
'       If txtAux(4).Text <> "" Then _
'           txtAux2(4).Text = DevuelveDesdeBDNew(1, "supdtogr", "nomsuple", "codsuple", txtAux(4).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
        ' **********************************************************************
 End Sub

Private Sub LimpiarCampos()
Dim i As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For i = 6 To 18
        txtAux(i).Text = ""
    Next i
    ' ****************************************************
    text2(12).Text = "" ' el nombre de la cuenta contable la ponemos a cero

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "banpropi"
        .Informe2 = "rManBanco.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{sbanco.codbanpr} = " & vSesion.Empresa
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={banpropi.codbanpr}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_Lostfocus()
  WheelUnHook
End Sub

