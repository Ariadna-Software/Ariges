VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacDtosFamMarca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dtos. Familia/Marca Cliente"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14745
   ClipControls    =   0   'False
   Icon            =   "frmFacDtosFamMarca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   33
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
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
      Enabled         =   0   'False
      Height          =   705
      Left            =   3870
      TabIndex        =   31
      Top             =   135
      Visible         =   0   'False
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   165
         TabIndex        =   32
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
      Left            =   12825
      TabIndex        =   30
      Top             =   270
      Visible         =   0   'False
      Width           =   1665
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
      Index           =   9
      Left            =   10560
      MaxLength       =   5
      TabIndex        =   10
      Tag             =   "Comision|N|S|0|99.90|sdtofm|comision|#0.00|N|"
      Text            =   "comi"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      ItemData        =   "frmFacDtosFamMarca.frx":000C
      Left            =   9720
      List            =   "frmFacDtosFamMarca.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Cod. Tarifa|N|N|||sdtofm|dtoesp||S|"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
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
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. Cliente|N|S|0|999999|sdtofm|codclien|000000|S|"
      Text            =   "cod. clien"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
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
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha Aplicación|F|N|||sdtofm|fechadto|dd/mm/yyyy|N|"
      Text            =   "fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   8190
      Width           =   2535
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
         Left            =   240
         TabIndex        =   29
         Top             =   180
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
      Left            =   1080
      TabIndex        =   27
      ToolTipText     =   "Buscar cliente"
      Top             =   3600
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
      Index           =   1
      Left            =   1920
      TabIndex        =   26
      ToolTipText     =   "Buscar familia artículo"
      Top             =   3600
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
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "Buscar marca"
      Top             =   3600
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
      Index           =   4
      Left            =   7320
      TabIndex        =   24
      ToolTipText     =   "Buscar tarifa"
      Top             =   3600
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
      Index           =   3
      Left            =   3960
      TabIndex        =   23
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
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
      Index           =   8
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "Cod. Tarifa|N|S|0|999|sdtofm|codactiv|000|S|"
      Text            =   "codt"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2445
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
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Descuento 1|N|N|0|99.90|sdtofm|dtoline1|#0.00|N|"
      Text            =   "Dto 1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Descuento 2|N|N|0|99.90|sdtofm|dtoline2|#0.00|N|"
      Text            =   "Dto 2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "Dto. 1 Cajas|N|S|0|99.90|sdtofm|dtocaja1|#0.00|N|"
      Text            =   "dto1 Cajas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Index           =   7
      Left            =   6000
      MaxLength       =   5
      TabIndex        =   7
      Tag             =   "Dto. 2 Cajas|N|S|0|99.90|sdtofm|dtocaja2|#0.00|N|"
      Text            =   "dto2 Cajas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Cod. Marca|N|S|0|9999|sdtofm|codmarca|0000|S|"
      Text            =   "codmarca"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   10635
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   7785
      Width           =   3885
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
      Left            =   12240
      TabIndex        =   11
      Top             =   8310
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
      Left            =   13395
      TabIndex        =   12
      Top             =   8310
      Width           =   1065
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   7785
      Width           =   5715
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
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   7785
      Width           =   4245
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
      Index           =   1
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Cod. Familia|N|S|0|9999|sdtofm|codfamia|0000|S|"
      Text            =   "cod.famia"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9240
      Top             =   4920
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
      Bindings        =   "frmFacDtosFamMarca.frx":0022
      Height          =   6360
      Left            =   240
      TabIndex        =   14
      Top             =   1035
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   11218
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
      Left            =   13395
      TabIndex        =   13
      Top             =   8265
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   10635
      TabIndex        =   21
      Top             =   7500
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Familia"
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
      Left            =   6180
      TabIndex        =   17
      Top             =   7500
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      Top             =   7500
      Width           =   2280
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
      TabIndex        =   15
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
Attribute VB_Name = "frmFacDtosFamMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmBasico2 'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFam As frmBasico2 'frmAlmFamiliaArticulo  'Form Mantenimiento Familias Articulos
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas  'Form Mantenimiento Marcas
Attribute frmM.VB_VarHelpID = -1
'Private WithEvents frmT As frmFacTarifas  'Form Mantenimiento Tarifas
Private WithEvents frmAc As frmFacActividades
Attribute frmAc.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim WhereConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

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
                CargaGrid True
                BotonAnyadir
            End If
        End If
    Case 4 'MODIFICAR
           If DatosOk And BLOQUEADesdeFormulario(Me) Then
                'Marzo 2010
                'antes cuando habia clvae ppal sin valores nul
                'If ModificaDesdeFormulario(Me, 3) Then
                If ModificaRegistro() Then
                    TerminaBloquear
                    NumReg = Data1.Recordset.AbsolutePosition
                    PonerModo 2
                    CancelaADODC Me.Data1
                    CargaGrid True
'                    CargaTxtAux False, False
                    LLamaLineas 10
                    SituarDataPosicion Data1, NumReg, Indicador
                End If
                lblIndicador.Caption = Indicador
                DataGrid1.SetFocus
            End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Cod Cliente
            Set frmC = New frmBasico2
            AyudaClientes frmC, txtAux(0)
            Set frmC = Nothing
        Case 1 'Cod Familia
'            Set frmFam = New frmAlmFamiliaArticulo
'            frmFam.DatosADevolverBusqueda = "0"
'            frmFam.Show vbModal
            Set frmFam = New frmBasico2
            AyudaFamilias frmFam, txtAux(Index)
            Set frmFam = Nothing
        Case 2 'Cod Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3 'Fecha aplicacion
            Screen.MousePointer = vbHourglass
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index).Text <> "" Then frmF.Fecha = CDate(txtAux(Index).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
        Case 4 'Tarifas
            Set frmAc = New frmFacActividades
            frmAc.DatosADevolverBusqueda = "0"
            frmAc.Show vbModal
            Set frmAc = Nothing
    End Select
    
    If Index = 4 Then
        PonerFoco txtAux(8)
    Else
        PonerFoco txtAux(Index)
    End If
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else 'No hay Registros en la Tabla
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            DeseleccionaGrid Me.DataGrid1
            PonerModo 2
            LLamaLineas 10
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.AbsolutePosition > 0 Then
            'Poner descripcion del Cliente
            If IsNull(Data1.Recordset!codClien) Then
                Text2(0).Text = ""
            Else
                Text2(0).Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", CStr(Data1.Recordset!codClien), "N")
            End If
            'Poner descripcion de Familia
            If IsNull(Data1.Recordset!Codfamia) Then
                Text2(1).Text = ""
            Else
                Text2(1).Text = DevuelveDesdeBDNew(conAri, "sfamia", "nomfamia", "codfamia", CStr(Data1.Recordset!Codfamia), "N")
            End If
            
            'Poner descripcion de Familia
            If IsNull(Data1.Recordset!codmarca) Then
                Text2(2).Text = ""
            Else
                Text2(2).Text = DevuelveDesdeBDNew(conAri, "smarca", "nommarca", "codmarca", CStr(Data1.Recordset!codmarca), "N")
            End If
            
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        End If
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(9).Image = 21  'Generar dto
'        .Buttons(10).Image = 42 'generar dtos prove
'        .Buttons(11).Image = 16 'Imprimir
'        .Buttons(12).Image = 15 'Salir
'    End With
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
        .Buttons(1).Image = 21 'Generar dto
        .Buttons(2).Image = 42 'generar dtos prove
    End With
    
    
    'JUNIO 2014
    'Vamos a trabajar con dto actividad
    'En el toolbar lo comentaremos tb
    Toolbar1.Buttons(9).visible = False
    Toolbar1.Buttons(10).visible = False
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    NombreTabla = "sdtofm" 'Tabla Descuentos Familia/Marca
    Ordenacion = " ORDER BY codclien, codfamia, codmarca "
    WhereConsulta = " WHERE codclien = -1"
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
On Error GoTo ECarga

    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    CargaGrid2
    DataGrid1.Enabled = (Modo = 2)
    Me.DataGrid1.ScrollBars = dbgAutomatic
    PrimeraVez = False
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    'SQL = "SELECT codclien, codfamia, codmarca, fechadto, dtoline1, dtoline2, dtocaja1, dtocaja2, " & tabla & ".codactiv, " & " Tarifas.nomlista "
    tots = "S|txtAux(0)|T|Cliente|1050|;S|cmdAux(0)|B||0|;S|txtAux(1)|T|Familia|950|;S|cmdAux(1)|B||0|;S|txtAux(2)|T|Marca|900|;S|cmdAux(2)|B||0|;"
    tots = tots & "S|txtAux(3)|T|Fecha Dto.|1450|;S|cmdAux(3)|B||0|;S|txtAux(4)|T|Dto 1|800|;S|txtAux(5)|T|Dto 2|800|;S|txtAux(6)|T|Dto Caja1|1200|;S|txtAux(7)|T|Dto Caja2|1200|;"
    tots = tots & "S|txtAux(8)|T|Actividad|1100|;S|cmdAux(4)|B||0|;S|txtAux2|T|Desc.Actividad|2450|;S|Combo1|C|Esp|650|;"
    tots = tots & "S|txtAux(9)|T|Comision|1150|;"
    arregla tots, DataGrid1, Me, 350

    'dtos alineados a la dcha
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(8).Alignment = dbgCenter
    DataGrid1.Columns(11).Alignment = dbgRight
ECarga2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

        DeseleccionaGrid Me.DataGrid1
        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

        For jj = 0 To txtAux.Count - 1
            txtAux(jj).Height = DataGrid1.RowHeight
            txtAux(jj).Top = alto
            txtAux(jj).visible = b
        Next jj
        txtAux2.Height = Me.DataGrid1.RowHeight
        txtAux2.Top = alto
        Combo1.Top = alto
        txtAux2.visible = b
        Combo1.visible = b
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = b
        Next jj
End Sub


Private Sub Form_Unload(Cancel As Integer)
     CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento actividades
    txtAux(8).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtAux2.Text = RecuperaValor(CadenaSeleccion, 2)
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
            Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtAux(3).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Clientes
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Familias
    txtAux(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento MARCAS
    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
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
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
     Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    'JUNIO 2014
    'Los puntos 10 y 11 NO se hacen. Quitar  mas adelante el codigo fuente
    If Button.Index = 10 Or Button.Index = 9 Then Exit Sub
    
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
            AbrirListado (54) '54: Listado Descuentos Familia/Marca
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
    
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(0)
        Case 2    'Preparamos para que pueda Modificar
            Me.cmdRegresar.visible = False
    End Select
                            
    BloquearClavesP (Modo = 4) ' si modificar
           
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    Me.DataGrid1.Enabled = (Modo = 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    'Modo 2. Hay datos y estamos visualizandolos
    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(1).Enabled = (b Or (Modo = 0))
    Me.mnNuevo.Enabled = (b Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    Toolbar5.Buttons(1).Enabled = Toolbar1.Buttons(5).Enabled
    Toolbar5.Buttons(2).Enabled = Toolbar1.Buttons(5).Enabled
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
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
Dim tabla As String
    
    tabla = "sdtofm"
    SQL = "SELECT codclien, codfamia, codmarca, fechadto, dtoline1, dtoline2, dtocaja1, dtocaja2, " & tabla & ".codactiv, " & " sactiv.nomactiv,if(dtoesp=0,"""",""Si"") dtoesp,comision"
    SQL = SQL & " FROM " & tabla & " LEFT JOIN sactiv ON " & tabla & ".codactiv ="
    SQL = SQL & " sactiv.codactiv"

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        ElseIf CadenaConsulta = "" Then
            If CadenaBusqueda <> "" Then
                CadenaBusqueda = CadenaBusqueda & " OR (" & MontaWHERE(True, True) & ")"
            Else
                'CadenaBusqueda = " WHERE (codclien=" & txtAux(0).Text & " and codfamia=" & txtAux(1).Text & " and codmarca=" & txtAux(2).Text & ")"
                CadenaBusqueda = " WHERE (" & MontaWHERE(True, True) & ")"
            End If
            SQL = SQL & CadenaBusqueda
        End If
    Else
        SQL = SQL & " WHERE codclien = -1"
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
        anc = ObtenerAlto(Me.DataGrid1) + 20
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
'Ver todos


    If vParamAplic.AlmacenB > 90 Then
        'HERBELCA.
        'Tienen ma de un millon de registros. le cuesta mucho
        If MsgBox("Seguro que desea ""Ver todos""?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    EsBusqueda = False
    LimpiarCampos
    
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
'        CadenaConsulta = ""
'    End If
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    Combo1.ListIndex = 0
    PonerFoco txtAux(0)
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
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    'poner valores grabados
    For i = 0 To 2
        If Not IsNull(Data1.Recordset.Fields(i)) Then
            txtAux(i).Text = DBLet(Data1.Recordset.Fields(i), "N")
            FormateaCampo txtAux(i)
        Else
            txtAux(i).Text = ""
        End If
    Next i
    txtAux(3).Text = DBLet(Data1.Recordset.Fields(3).Value, "F")
    For i = 4 To 7
        txtAux(i).Text = DBLet(Data1.Recordset.Fields(i), "N")
        FormateaCampo txtAux(i)
    Next i
    'i=8
    txtAux(i).Text = DBLet(Data1.Recordset.Fields(i), "T")
    txtAux2.Text = DBLet(Data1.Recordset.Fields(9), "T")
    If DBLet(Data1.Recordset!dtoesp, "T") = "Si" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    txtAux(9).Text = DBLet(Data1.Recordset.Fields(11), "T")
    PonerFoco txtAux(3)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        SQL = "¿Seguro que desea eliminar el Descuento para?" & vbCrLf
        SQL = SQL & vbCrLf & "Cliente: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Text2(0).Text
        SQL = SQL & vbCrLf & "Familia: " & Format(Data1.Recordset.Fields(1).Value, "0000") & " - " & Text2(1).Text
        SQL = SQL & vbCrLf & "Marca : " & Format(Data1.Recordset.Fields(2).Value, "0000") & " - " & Text2(2).Text
        SQL = SQL & vbCrLf & "Fecha : " & Format(Data1.Recordset!fechadto, "dd/mm/yyyy")
        SQL = SQL & vbCrLf & "Actividad : " & Format(Data1.Recordset!codactiv, "0000") & " - " & DBLet(Data1.Recordset!nomactiv, "T")
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            SQL = "Delete from sdtofm where codclien" & vDBSET(Data1.Recordset!codClien, True, True, False)
            SQL = SQL & " and codfamia " & vDBSET(Data1.Recordset!Codfamia, True, True, False) & " and codmarca " & vDBSET(Data1.Recordset!codmarca, True, True, False)
            SQL = SQL & " and codactiv " & vDBSET(Data1.Recordset!codactiv, True, True, False) & " and fechadto=" & DBSet(Data1.Recordset!fechadto, "F", "S")
            conn.Execute SQL
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataTrasEliminar Data1, NumRegElim
            CargaGrid2
        End If
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim C As String
Dim C2 As String

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'es obligado O EL Cliente o la actividad
    If txtAux(0).Text <> "" And txtAux(8).Text <> "" Then
        MsgBox "Ponga cliente o actividad, o ninguno", vbExclamation
        Exit Function
    End If

    
    
    'Como NO hay clave primaria tengo que comprobar que NO exista un valor
    Set RS = New ADODB.Recordset
    
    If Modo = 3 Then
        'Esta INSERTAND
        C = "Select * from sdtofm"
        C = C & " WHERE " & MontaWHERE(True, True)
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            MsgBox "Ya existe un registro con esos datos!", vbExclamation
        Else
            C = ""
        End If
        RS.Close
        If C <> "" Then Exit Function
    Else
        'Compruebo si ha cambiado de la clave primaria
        C = MontaWHERE(True, True)
        C2 = MontaWHERE(False, True)
        If C2 <> C Then
               'HA CAMBIADO VALORES DE LA CLAVE PRIMARIA)o de los identificativos)
               Debug.Print "FALTA###"
        Else
            C = "" 'NO HA CAMBIADO NADA
        End If
        If C <> "" Then
            'Compruebo si ya existe un valor para esos valores
            
        End If
                
    End If
    

    'Solo lleva comision si el descuento es especial
    If txtAux(9).Text <> "" Then
        If Combo1.ListIndex = 0 Then
            MsgBox "Solo llevan comision los descuentos especiales", vbExclamation
            PonerFocoCbo Combo1
            Exit Function
        End If
    End If
    
    
    DatosOk = True
End Function




Private Function MontaWHERE(ConLosTxt As Boolean, ComprobarConFecha As Boolean) As String
Dim s As String
    
    If ConLosTxt Then
        s = " codclien " & vDBSET(txtAux(0).Text, True, True, ConLosTxt)
        s = s & " and codfamia " & vDBSET(txtAux(1).Text, True, True, ConLosTxt)
        s = s & " and codmarca " & vDBSET(txtAux(2).Text, True, True, ConLosTxt)
        If ComprobarConFecha Then s = s & " and fechadto " & vDBSET(txtAux(3).Text, False, False, ConLosTxt)
        s = s & " and " & NombreTabla & ".codactiv " & vDBSET(txtAux(8).Text, True, True, ConLosTxt)
        
    Else
        
        'Contra el DATA1
        s = " codclien " & vDBSET(Data1.Recordset!codClien, True, True, ConLosTxt)
        s = s & " and codfamia " & vDBSET(Data1.Recordset!Codfamia, True, True, ConLosTxt)
        s = s & " and codmarca " & vDBSET(Data1.Recordset!codmarca, True, True, ConLosTxt)
        If ComprobarConFecha Then s = s & " and fechadto " & vDBSET(Data1.Recordset!fechadto, False, False, ConLosTxt)
        s = s & " and " & NombreTabla & ".codactiv " & vDBSET(Data1.Recordset!codactiv, True, True, ConLosTxt)
    End If
    MontaWHERE = s
End Function


Private Function vDBSET(Valor As Variant, EsNumerico As Boolean, esNULO As Boolean, DesdeTextos As Boolean) As Variant
Dim eNulo As Boolean
    If DesdeTextos Then
        eNulo = Valor = ""
    Else
        eNulo = IsNull(Valor)
    End If
    
    If eNulo Then
        vDBSET = " is null"
    Else
        If EsNumerico Then
            vDBSET = " = " & Val(Valor)
        Else
            vDBSET = " = '" & Format(Valor, FormatoFecha) & "'"
        End If
    End If
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(txtAux(0), 40, "Cod. Clien.")
    cad = cad & ParaGrid(txtAux(1), 20, "Cod. Artic")
    tabla = NombreTabla
    Titulo = "Precios Especiales"

    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
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
                PonerFoco txtAux(kCampo)
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
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        PonerModo Modo
        CargaGrid False
         MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        If EsBusqueda Then CadenaConsulta = ""
        PonerCampos
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
    PonerCamposForma Me, Data1
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)

    'JUNIO 2014
    'Los puntos 10 y 11 NO se hacen. Quitar  mas adelante el codigo fuente
    If Button.Index = 10 Or Button.Index = 9 Then Exit Sub
    
    Select Case Button.Index
        Case 1
            'Generar desde sfamiadtos
            CadenaDesdeOtroForm = ""
            frmVarios.Opcion = 7
            frmVarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                'Ha generado datos
                CadenaConsulta = "codclien = " & RecuperaValor(CadenaDesdeOtroForm, 1)
                CadenaConsulta = CadenaConsulta & " AND fechadto = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F")
                CadenaDesdeOtroForm = ""
                CadenaBusqueda = " WHERE " & CadenaConsulta
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaConsulta
                EsBusqueda = True
                PonerCadenaBusqueda
                EsBusqueda = False
            End If
            
        
        Case 2
            CadenaDesdeOtroForm = ""
            frmListado3.Opcion = 17
            frmListado3.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                'Ha generado datos
                CadenaBusqueda = RecuperaValor(CadenaDesdeOtroForm, 1)
                If CadenaBusqueda <> "" Then CadenaBusqueda = CadenaBusqueda & " AND "
                CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
                
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
                CadenaConsulta = CadenaBusqueda & "codfamia IN (" & CadenaDesdeOtroForm & ")"
                CadenaDesdeOtroForm = ""
                CadenaBusqueda = " WHERE " & CadenaConsulta
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaConsulta
                EsBusqueda = True
                PonerCadenaBusqueda
                EsBusqueda = False
            End If
        
    End Select

End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente
            Case 1: KEYBusqueda KeyAscii, 1 'familia
            Case 2: KEYBusqueda KeyAscii, 2 'marca
            Case 3: KEYBusqueda KeyAscii, 3 'fecha
            Case 8: KEYBusqueda KeyAscii, 4 'actividad
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub



Private Sub BloquearClavesP(bol As Boolean)
'Si BloquearClavesPrimarias=true deshablilita los textbox de codigos y lo pone amarillo
'y habilita el resto de campos para introducir nuevos valores
'Si BloquearClavesPrimarias=false habilita los textbox de codigos para introducir
Dim i As Byte

    For i = 0 To 2 'Codigos
        BloquearTxt txtAux(i), bol
        Me.cmdAux(i).Enabled = Not bol
    Next i
    Me.cmdAux(4).Enabled = Not bol
    BloquearTxt txtAux(8), bol
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
'    If txtAux(Index).Text = "" Then Exit Sub
    If Modo = 1 Then Exit Sub
    Select Case Index
        Case 0 'cod. Cliente
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "sclien", "nomclien")
            Else
                Text2(Index).Text = ""
            End If
        Case 1 'Cod. Familia
            If txtAux(Index).Text <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "sfamia", "nomfamia")
                Else
                    Text2(Index).Text = ""
                End If
                If Text2(Index).Text = "" Then txtAux(Index).Text = ""
            Else
                Text2(Index).Text = ""
            End If
        Case 2 'Cod. Marca
            If txtAux(Index).Text <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "smarca", "nommarca")
                Else
                    Text2(Index).Text = ""
                End If
                If Text2(Index).Text = "" Then txtAux(Index).Text = ""
            Else
                Text2(Index).Text = ""
            End If
        Case 3 'Fecha Descuento
            PonerFormatoFecha txtAux(Index)
        Case 4 To 7, 9 'Descuentos
            'Formato tipo 4: Decimal(4,2)
            PonerFormatoDecimal txtAux(Index), 4
        Case 8 'Cod. actividad
            If txtAux(Index).Text <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then
                    txtAux2.Text = PonerNombreDeCod(txtAux(Index), conAri, "sactiv", "nomactiv")
                Else
                    txtAux2.Text = ""
                End If
                If txtAux2.Text = "" Then txtAux(Index).Text = ""
            Else
                 txtAux2.Text = ""
            End If
    End Select
End Sub



Private Function ModificaRegistro() As Boolean
Dim C2 As String
    On Error GoTo EModificaRegistro
    ModificaRegistro = False
    
    C2 = "UPDATE sdtofm SET "
    C2 = C2 & " fechadto = " & DBSet(txtAux(3), "F", "N")
    'dto
    C2 = C2 & ", dtoline1 = " & DBSet(txtAux(4), "N", "N")
    C2 = C2 & ", dtoline2 = " & DBSet(txtAux(5), "N", "N")
    'dto caja
    C2 = C2 & ", dtocaja1 = " & DBSet(txtAux(6), "N", "S")
    C2 = C2 & ", dtocaja2 = " & DBSet(txtAux(7), "N", "S")
    
    C2 = C2 & ", dtoesp = " & Combo1.ListIndex
    
    C2 = C2 & ", comision = " & DBSet(txtAux(9), "N", "S")
    
    C2 = C2 & " WHERE " & MontaWHERE(True, False)
    conn.Execute C2
    
    
    
    
    
    
    ModificaRegistro = True
    Exit Function
EModificaRegistro:
    MuestraError Err.Number, "Modifica Registro"
End Function

