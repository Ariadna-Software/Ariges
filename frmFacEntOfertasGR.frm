VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntOfertasGR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   17895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacEntOfertasGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   17895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFiltro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8640
      TabIndex        =   106
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox cbFiltro 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkMisOfertas 
         Caption         =   "Mis ofertas"
         Height          =   240
         Left            =   2760
         TabIndex        =   107
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6120
      TabIndex        =   102
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   103
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
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3480
      TabIndex        =   100
      Top             =   0
      Width           =   2535
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   101
         Top             =   150
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar pedido"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recordatorio"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Valoracion"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir proforma"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Genera factura"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   98
      Top             =   0
      Width           =   3315
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   60
         TabIndex        =   99
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
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   0
      Left            =   13080
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   96
      Text            =   "BASE IMP."
      Top             =   240
      Width           =   1125
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   360
      Index           =   56
      Left            =   14280
      MaxLength       =   15
      TabIndex        =   95
      Text            =   "Text1 7"
      Top             =   240
      Width           =   1650
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   80
      Top             =   720
      Width           =   17655
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   12960
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   480
         Width           =   4425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   12120
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scapre|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   5
         Left            =   6600
         MaxLength       =   60
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scapre|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   480
         Width           =   5265
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   4
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scapre|codclien|000000|N|"
         Text            =   "99999"
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   1220
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Oferta|F|N|||scapre|fecofert|dd/mm/yyyy|N|"
         Text            =   "10/12/2199"
         Top             =   480
         Width           =   1305
      End
      Begin VB.CheckBox chkAceptado 
         Caption         =   "Aceptada"
         Height          =   255
         Index           =   0
         Left            =   4230
         TabIndex        =   3
         Tag             =   "Aceptada|N|N|||scapre|aceptado||N|"
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   360
         Index           =   0
         Left            =   75
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Oferta|N|S|0||scapre|numofert|0000000|S|"
         Text            =   "Text1 7"
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrega|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
         Text            =   "10/12/2199"
         Top             =   480
         Width           =   1305
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   13920
         ToolTipText     =   "Buscar trabajador"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada por"
         Height          =   240
         Index           =   21
         Left            =   12120
         TabIndex        =   86
         Top             =   180
         Width           =   1500
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6480
         ToolTipText     =   "Buscar cliente"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Index           =   0
         Left            =   5640
         TabIndex        =   84
         Top             =   180
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmFacEntOfertasGR.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   50
         Left            =   75
         TabIndex        =   82
         Top             =   165
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3840
         Picture         =   "frmFacEntOfertasGR.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Entrega"
         Height          =   255
         Index           =   51
         Left            =   2760
         TabIndex        =   81
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "F. Oferta"
         Height          =   255
         Index           =   14
         Left            =   1230
         TabIndex        =   83
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   16
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Text            =   "frmFacEntOfertasGR.frx":0122
      Top             =   9240
      Width           =   8805
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   9060
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   16560
      TabIndex        =   39
      Top             =   9120
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   15360
      TabIndex        =   38
      Top             =   9120
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8400
      Top             =   9120
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
      Left            =   8400
      Top             =   9240
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7200
      Left            =   120
      TabIndex        =   43
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1800
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   12700
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFacEntOfertasGR.frx":015F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtAux(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAux(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboOpcion"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameToolAux(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Textos de la Carta"
      TabPicture(1)   =   "frmFacEntOfertasGR.frx":017B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(19)"
      Tab(1).Control(1)=   "Text1(20)"
      Tab(1).Control(2)=   "Text1(18)"
      Tab(1).Control(3)=   "Text1(30)"
      Tab(1).Control(4)=   "Text1(29)"
      Tab(1).Control(5)=   "Text1(28)"
      Tab(1).Control(6)=   "Text1(27)"
      Tab(1).Control(7)=   "Text1(26)"
      Tab(1).Control(8)=   "Text1(25)"
      Tab(1).Control(9)=   "Text1(24)"
      Tab(1).Control(10)=   "Text1(23)"
      Tab(1).Control(11)=   "Text1(22)"
      Tab(1).Control(12)=   "Text1(21)"
      Tab(1).Control(13)=   "Label1(2)"
      Tab(1).Control(14)=   "Label1(45)"
      Tab(1).Control(15)=   "Label1(5)"
      Tab(1).Control(16)=   "Label1(3)"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Concepto y Gestión Oferta"
      TabPicture(2)   =   "frmFacEntOfertasGR.frx":0197
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(36)"
      Tab(2).Control(1)=   "Text1(35)"
      Tab(2).Control(2)=   "Text1(32)"
      Tab(2).Control(3)=   "Text1(31)"
      Tab(2).Control(4)=   "Label1(28)"
      Tab(2).Control(5)=   "Label1(18)"
      Tab(2).Control(6)=   "Label1(37)"
      Tab(2).Control(7)=   "Label1(38)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Documentos"
      TabPicture(3)   =   "frmFacEntOfertasGR.frx":01B3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameToolAux(0)"
      Tab(3).Control(1)=   "ListView1"
      Tab(3).Control(2)=   "AcroPDF1"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Totales"
      TabPicture(4)   =   "frmFacEntOfertasGR.frx":01CF
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text2(40)"
      Tab(4).Control(1)=   "Text1(40)"
      Tab(4).Control(2)=   "Text1(39)"
      Tab(4).Control(3)=   "Text1(38)"
      Tab(4).Control(4)=   "FrameFactura"
      Tab(4).Control(5)=   "cboMotTra"
      Tab(4).Control(6)=   "Text1(37)"
      Tab(4).Control(7)=   "Label1(44)"
      Tab(4).Control(8)=   "Label1(43)"
      Tab(4).Control(9)=   "Label1(36)"
      Tab(4).Control(10)=   "Label1(27)"
      Tab(4).Control(11)=   "Label1(52)"
      Tab(4).Control(12)=   "Label1(29)"
      Tab(4).Control(13)=   "Label1(40)"
      Tab(4).ControlCount=   14
      Begin VB.Frame FrameToolAux 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   -74640
         TabIndex        =   165
         Top             =   840
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   166
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   9
         Left            =   12120
         MaxLength       =   12
         TabIndex        =   164
         Tag             =   "Importe"
         Text            =   "codigva"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   40
         Left            =   -64440
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   163
         Text            =   "Text2"
         Top             =   720
         Width           =   4305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   40
         Left            =   -65400
         MaxLength       =   10
         TabIndex        =   162
         Tag             =   "Trab. pedido|N|S|||scapre|codtrabped|0000||"
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   39
         Left            =   -69000
         MaxLength       =   10
         TabIndex        =   159
         Tag             =   "Traspasoa|F|S|||scapre|fecpedcl|dd/mm/yyyy|N|"
         Top             =   720
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   38
         Left            =   -71520
         MaxLength       =   10
         TabIndex        =   157
         Tag             =   "NºPed|N|S|||scapre|numpedcl||N|"
         Top             =   720
         Width           =   1305
      End
      Begin VB.Frame FrameFactura 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   -74520
         TabIndex        =   114
         Top             =   2760
         Width           =   16455
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   55
            Left            =   13920
            MaxLength       =   15
            TabIndex        =   138
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   137
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   136
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   135
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   134
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   49
            Left            =   13920
            MaxLength       =   5
            TabIndex        =   133
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   52
            Left            =   14520
            MaxLength       =   15
            TabIndex        =   132
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   50
            Left            =   13920
            MaxLength       =   5
            TabIndex        =   131
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   53
            Left            =   14520
            MaxLength       =   15
            TabIndex        =   130
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   51
            Left            =   13920
            MaxLength       =   5
            TabIndex        =   129
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   54
            Left            =   14520
            MaxLength       =   15
            TabIndex        =   128
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   48
            Left            =   12240
            MaxLength       =   15
            TabIndex        =   127
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   42
            Left            =   11520
            MaxLength       =   5
            TabIndex        =   126
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   645
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   39
            Left            =   9240
            MaxLength       =   4
            TabIndex        =   125
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   45
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   124
            Text            =   "Text1 7"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   47
            Left            =   12240
            MaxLength       =   15
            TabIndex        =   123
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   41
            Left            =   11520
            MaxLength       =   5
            TabIndex        =   122
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   645
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   38
            Left            =   9240
            MaxLength       =   4
            TabIndex        =   121
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   44
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   120
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   46
            Left            =   12240
            MaxLength       =   15
            TabIndex        =   119
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   40
            Left            =   11520
            MaxLength       =   5
            TabIndex        =   118
            Text            =   "Text1 7"
            Top             =   960
            Width           =   645
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   37
            Left            =   9240
            MaxLength       =   4
            TabIndex        =   117
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   43
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   116
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   360
            Index           =   1
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   115
            Text            =   "Text1 7"
            Top             =   2400
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL OFERTA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   39
            Left            =   11520
            TabIndex        =   153
            Top             =   2640
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "="
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
            Index           =   32
            Left            =   5520
            TabIndex        =   152
            Top             =   1013
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "-"
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
            Index           =   31
            Left            =   3720
            TabIndex        =   151
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "-"
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
            Index           =   30
            Left            =   1920
            TabIndex        =   150
            Top             =   1013
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   240
            Index           =   8
            Left            =   5760
            TabIndex        =   149
            Top             =   367
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto.Gnral"
            Height          =   240
            Index           =   12
            Left            =   3960
            TabIndex        =   148
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto. PP"
            Height          =   240
            Index           =   11
            Left            =   2160
            TabIndex        =   147
            Top             =   367
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   240
            Index           =   10
            Left            =   240
            TabIndex        =   146
            Top             =   367
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   22
            Left            =   15000
            TabIndex        =   145
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   14280
            TabIndex        =   144
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   9360
            TabIndex        =   143
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   12000
            TabIndex        =   142
            Top             =   360
            Width           =   615
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   15840
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   12840
            TabIndex        =   141
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   240
            Index           =   9
            Left            =   10440
            TabIndex        =   140
            Top             =   367
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Total opciones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   139
            Top             =   2400
            Width           =   1770
         End
      End
      Begin VB.ComboBox cboMotTra 
         Height          =   360
         Left            =   -69000
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Tag             =   "Motivo|N|S|||scapre|motivoTraspaso||N|"
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   37
         Left            =   -71520
         MaxLength       =   10
         TabIndex        =   109
         Tag             =   "Traspasoa|F|S|||scapre|fechamov|dd/mm/yyyy|N|"
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Frame FrameToolAux 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   240
         TabIndex        =   104
         Top             =   3300
         Width           =   2805
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Intercalar"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Plantilla"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "traer lineas ofertas"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox Text1 
         Height          =   2595
         Index           =   36
         Left            =   -66120
         MultiLine       =   -1  'True
         TabIndex        =   37
         Tag             =   "ObCom|T|S|||scapre|obscompra|||"
         Text            =   "frmFacEntOfertasGR.frx":01EB
         Top             =   4320
         Width           =   8205
      End
      Begin VB.ComboBox cboOpcion 
         Height          =   360
         ItemData        =   "frmFacEntOfertasGR.frx":01F3
         Left            =   11160
         List            =   "frmFacEntOfertasGR.frx":01FD
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   2595
         Index           =   35
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   36
         Tag             =   "Obs CRM|T|S|||scapre|observacrm|||"
         Text            =   "frmFacEntOfertasGR.frx":0209
         Top             =   4320
         Width           =   7365
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   5
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   57
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2835
         Left            =   240
         TabIndex        =   64
         Top             =   450
         Width           =   17175
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   34
            Left            =   10125
            MaxLength       =   40
            TabIndex        =   91
            Tag             =   "E-mail confirmación|T|S|||scapre|mailconfir||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aqteter"
            Top             =   2400
            Width           =   6315
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   33
            Left            =   11160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   89
            Text            =   "Text2"
            Top             =   1560
            Width           =   5355
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   33
            Left            =   10125
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Dir Envio|N|S|0|9999|scapre|coddiren|0000|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   12
            Left            =   11160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   77
            Tag             =   "Direccion/Dpto.|T|S|||scapre|nomdirec||N|"
            Text            =   "Text2"
            Top             =   210
            Width           =   5355
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   12
            Left            =   10125
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Direccion/Dpto.|N|S|0|9999|scapre|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   210
            Width           =   900
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   11
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scapre|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1524
            Width           =   3045
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   9
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scapre|codpobla||N|"
            Text            =   "Text15"
            Top             =   1086
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   10
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scapre|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1086
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   7
            Left            =   4875
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scapre|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   210
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   6
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scapre|nifclien||N|"
            Text            =   "123456789"
            Top             =   210
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   13
            Left            =   1560
            MaxLength       =   60
            TabIndex        =   13
            Tag             =   "Referencia Cliente|T|S|||scapre|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1962
            Width           =   4845
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   17
            Left            =   10125
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Cod. Agente|N|N|0|9999|scapre|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   648
            Width           =   900
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   17
            Left            =   11160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   71
            Text            =   "Text2"
            Top             =   648
            Width           =   5355
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   14
            Left            =   10125
            MaxLength       =   3
            TabIndex        =   19
            Tag             =   "Forma de Pago|N|N|0|999|scapre|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1110
            Width           =   900
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   14
            Left            =   11160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   66
            Text            =   "Text2"
            Top             =   1110
            Width           =   5355
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   15
            Left            =   10125
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scapre|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1995
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   16
            Left            =   12360
            MaxLength       =   5
            TabIndex        =   16
            Tag             =   "Descuento General|N|N|0|99.90|scapre|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1995
            Width           =   735
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   360
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "Tipo Facturación|N|N|||scapre|tipofact||N|"
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   8
            Left            =   1560
            MaxLength       =   60
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scapre|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   648
            Width           =   6465
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail confirmación"
            Height          =   255
            Index           =   23
            Left            =   8580
            TabIndex        =   92
            Top             =   2453
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   9840
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Dir. Envio"
            Height          =   240
            Index           =   6
            Left            =   8580
            TabIndex        =   90
            Top             =   1530
            Width           =   930
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   1320
            ToolTipText     =   "Buscar población"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   240
            Index           =   1
            Left            =   8580
            TabIndex        =   79
            Top             =   270
            Width           =   570
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   9840
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   78
            Top             =   1524
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   76
            Top             =   1086
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   3660
            TabIndex        =   75
            Top             =   210
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "N.I.F."
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   74
            Top             =   210
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1320
            ToolTipText     =   "Buscar cliente varios"
            Top             =   225
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   73
            Top             =   1962
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   240
            Index           =   34
            Left            =   8580
            TabIndex        =   72
            Top             =   708
            Width           =   705
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   9840
            ToolTipText     =   "Buscar agente"
            Top             =   708
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   240
            Index           =   15
            Left            =   8580
            TabIndex        =   70
            Top             =   1110
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   8580
            TabIndex        =   69
            Top             =   1995
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   11400
            TabIndex        =   68
            Top             =   1995
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Factura"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   67
            Top             =   2453
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   9840
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   65
            Top             =   648
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   19
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Plazo Entrega 2|T|S|||scapre|plazos02||N|"
         Top             =   900
         Width           =   11805
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Validez de la oferta|T|S|||scapre|plazos03||N|"
         Top             =   1560
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   21
         Tag             =   "Plazo Entrega 1|T|S|||scapre|plazos01||N|"
         Top             =   480
         Width           =   11805
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   63
         ToolTipText     =   "Buscar artículo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   62
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   8
         Left            =   10080
         MaxLength       =   12
         TabIndex        =   58
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   7
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   54
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   6
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   53
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   4
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   52
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   2235
         Index           =   32
         Left            =   -66120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Tag             =   "Gestión Oferta|T|S|||scapre|seguiofe||N|"
         Text            =   "frmFacEntOfertasGR.frx":0211
         Top             =   1080
         Width           =   8325
      End
      Begin VB.TextBox Text1 
         Height          =   2235
         Index           =   31
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   34
         Tag             =   "Concepto Oferta|T|S|||scapre|concepto||N|"
         Text            =   "frmFacEntOfertasGR.frx":0219
         Top             =   1080
         Width           =   7485
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   3
         Left            =   6240
         MaxLength       =   16
         TabIndex        =   51
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   49
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   48
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   33
         Tag             =   "Observación 5|T|S|||scapre|observa05||N|"
         Text            =   "1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA1AAAWAWAWA"
         Top             =   6360
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   29
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación 4|T|S|||scapre|observa04||N|"
         Top             =   5940
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   28
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observación 3|T|S|||scapre|observa03||N|"
         Top             =   5520
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   27
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 2|T|S|||scapre|observa02||N|"
         Top             =   5100
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   26
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 1|T|S|||scapre|observa01||N|"
         Top             =   4680
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   25
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Asunto Carta 5|T|S|||scapre|asunto05||N|"
         Top             =   3840
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Asunto Carta 4|T|S|||scapre|asunto04||N|"
         Top             =   3420
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Asunto Carta 3|T|S|||scapre|asunto03||N|"
         Top             =   3000
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Asunto Carta|T|S|||scapre|asunto02||N|"
         Top             =   2580
         Width           =   11925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   21
         Left            =   -72600
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Asunto Carta 1|T|S|||scapre|asunto01||N|"
         Top             =   2160
         Width           =   11925
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntOfertasGR.frx":0221
         Height          =   2895
         Left            =   240
         TabIndex        =   59
         Top             =   4080
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   -72840
         TabIndex        =   154
         Top             =   960
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fichero"
            Object.Width           =   0
         EndProperty
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   4935
         Left            =   -64440
         TabIndex        =   155
         Top             =   960
         Width           =   6375
         _cx             =   5080
         _cy             =   5080
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador pedido"
         Height          =   240
         Index           =   44
         Left            =   -67320
         TabIndex        =   161
         Top             =   780
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   43
         Left            =   -69960
         TabIndex        =   160
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
         Height          =   240
         Index           =   36
         Left            =   -72360
         TabIndex        =   158
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido vinculado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   27
         Left            =   -74520
         TabIndex        =   156
         Top             =   773
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Index           =   52
         Left            =   -69960
         TabIndex        =   113
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   29
         Left            =   -72360
         TabIndex        =   112
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Oferta archivada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   40
         Left            =   -74520
         TabIndex        =   111
         Top             =   1380
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones compra"
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
         Index           =   28
         Left            =   -66120
         TabIndex        =   94
         Top             =   3960
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "Obse"
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   93
         Top             =   3960
         Width           =   4995
      End
      Begin VB.Label Label1 
         Caption         =   "Validez Oferta"
         Height          =   240
         Index           =   2
         Left            =   -74760
         TabIndex        =   87
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Oferta"
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
         Index           =   37
         Left            =   -74760
         TabIndex        =   61
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Gestión Oferta"
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
         Index           =   38
         Left            =   -66240
         TabIndex        =   60
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   45
         Left            =   -74640
         TabIndex        =   47
         Top             =   4680
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto Carta"
         Height          =   240
         Index           =   5
         Left            =   -74760
         TabIndex        =   45
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo Entrega"
         Height          =   240
         Index           =   3
         Left            =   -74760
         TabIndex        =   44
         Top             =   450
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   420
      Left            =   16560
      TabIndex        =   40
      Top             =   9120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   240
      Left            =   16200
      TabIndex        =   97
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   8
      Left            =   4200
      ToolTipText     =   "Buscar forma de pago"
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label lblF 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   88
      Top             =   9240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2880
      TabIndex        =   46
      Top             =   9000
      Width           =   1335
   End
End
Attribute VB_Name = "frmFacEntOfertasGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)



Public DatosOferta As String   'Para situarla

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmC As frmBasico2 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 '%=%=frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmOfe As frmBasico2
Attribute frmOfe.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer 'Listados para Ofertas
Attribute frmList.VB_VarHelpID = -1

'Carga de Plantillas en la linea de la Oferta
Private WithEvents frmPlant As frmFacCargaPlantilla  'Form para cargar plantillas
Attribute frmPlant.VB_VarHelpID = -1
'Carga las lineas de otra Oferta
Private WithEvents frmTOferta As frmFacTraerOferta
Attribute frmTOferta.VB_VarHelpID = -1

Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1




Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private Modo As Byte

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera2 As Byte  '0-Cabecera    1.-Coddirec   2.- direnvio
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No


Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private CadenaSQL As String 'Para crear consulta de Generar Pedido a partir de la Oferta

Private Ordenacion As String   'ORDER BY de la cadena consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal


Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim txtAnterior As String

Dim LineaIntercalar As Integer 'NO reutilizar

Dim PulsadoMas2 As Boolean
Dim PulsaF2 As Boolean
' ---- [15/09/2009] (LAURA)
'Dim ElArticulo As String   'para la sdesca
' ----

Dim GrabaLogCambioPrecioDto As Boolean

Dim PorDebajoPrecioMinimo As Boolean
Dim CarpetaDestinoOferta As String


'Para buscar por los chks
Private BuscaChekc As String

'=====================================================================================


Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboMotTra_Click()
    If Modo <> 4 Then Exit Sub
    'Solo modificando
    If cboMotTra.ListIndex >= 0 Then
        If Text1(37).Text = "" Then Text1(37).Text = Format(Now, "dd/mm/yyyy")
    End If
End Sub

Private Sub cboOpcion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAceptado_Click(Index As Integer)
    If Modo = 1 Then CheckCadenaBusqueda chkAceptado(Index), BuscaChekc
End Sub

Private Sub chkAceptado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkAceptado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean
Dim CambiaDpto As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
                Me.SSTab1.Tab = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    CadenaSQL = ""
                    
                    CambiaDpto = False
                    If Text1(12).Text = "" Then
                        If Not IsNull(Data1.Recordset!CodDirec) Then CambiaDpto = True
                    Else
                        If Text1(12).Text <> Val(DBLet(Data1.Recordset!CodDirec, "N")) Then CambiaDpto = True
                    End If
                    
                    If CambiaDpto Then
                        If Text2(12).Text = "" Then
                            CadenaSQL = "NULL"
                        Else
                            CadenaSQL = "'" & DevNombreSQL(Text2(12).Text) & "'"
                        End If
                        CadenaSQL = "UPDATE scapre SET nomdirec=" & CadenaSQL & " WHERE"
                        CadenaSQL = CadenaSQL & " numofert= " & Text1(0).Text
                    
                        
                    End If
                    'Actualizar los datos del cliente si es de varios
                    EsDeVarios = EsClienteVarios(Text1(4).Text)
                    If EsDeVarios Then ActualizarClienteVarios Text1(4).Text, Text1(6).Text
                    TerminaBloquear
                    
                    If CadenaSQL <> "" Then
                        Espera 0.2
                        ejecutar CadenaSQL, False
                        CadenaSQL = ""
                    End If
                    PosicionarData
                    
                    
                End If
            End If
            
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'slima1' (Revisiones)
            If ModificaLineas = 1 Then 'INSERTAR lineas Ofertas
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    If LineaIntercalar > 0 Then
                        'HA intercalado la linea. Ponemos luego en normal
                        Me.DataGrid1.Enabled = True
                        DataGrid1.AllowAddNew = False
                        NumRegElim = LineaIntercalar
                        CargaTxtAux False, False
                        CargaGrid2 DataGrid1, Data2
                        PosicionarData2
                        ModificaLineas = 0
                        PonerBotonCabecera True
                        BloquearTxt Text2(16), True
                        cmdRegresar_Click
                    Else
                        'Que meta otra
                        BotonAnyadirLinea False
                    End If
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    cmdRegresar_Click
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
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
            Set FrmArt = New frmBasico2
            'FrmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo Busqueda
            'FrmArt.DeConsulta = True
'            FrmArt.DesdeTPV = False
'            FrmArt.Show vbModal
            AyudaArticulos FrmArt, txtAux(Index)
            Set FrmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
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
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                DataGrid1.Enabled = True
            End If
            LineaIntercalar = 0
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            cmdRegresar_Click
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    BloquearTxt Text1(0), True, True
    
    
    PonerObservacionesPordefecto
    
    
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    'Si fuera agente debe poner el codigo de agente
    If vUsu.CodigoAgente > 0 Then
        Text1(17).Text = vUsu.CodigoAgente
        Text1_LostFocus 17
    End If
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        Me.chkAceptado(0).Value = 1
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
    End If
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea(Intercalando As Boolean)


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    Precio = ""
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    Me.SSTab1.Tab = 0
    If Intercalando Then
        lblIndicador.Caption = "** INTERCALAR **"
        If Not Data2.Recordset.EOF Then
            LineaIntercalar = Data2.Recordset!numlinea
        End If
    Else
        LineaIntercalar = 0
        lblIndicador.Caption = "INSERTAR"
    End If
    lblIndicador.Refresh
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(1, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    Me.cboOpcion.ListIndex = -1
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    'Text2(18).Text = ""
    
   ' BloquearTxt txtAux(6), True
   ' BloquearTxt txtAux(7), True
    
    'Para recordar que estamos intercalando
    If Intercalando Then
        txtAux(0).BackColor = vbRed
    Else
        txtAux(0).BackColor = vbWhite
    End If
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
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
        BuscaChekc = ""
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
Dim Cad As String
'Ver todos
    
    'Agosto 2011
    'Si el usuario es AGENTE solo puede ver las suyas
    
    Cad = " true "  '1=1
    If vUsu.CodigoAgente > 0 Then
        Cad = "(codagent = " & vUsu.CodigoAgente
        If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then Cad = Cad & " OR  codclien IN (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
        Cad = Cad & ")"
    End If
    PonerFiltro Cad  'Añadira al SQL el filtro establecido
    
    
      
      
    
    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera2 = 0
        MandaBusquedaPrevia Cad
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & Cad & Ordenacion
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


    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        
    'Si es Cliente de Varios no se pueden modificar sus datos
    EsDeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente EsDeVarios
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
    
    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    Me.SSTab1.Tab = 0
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    DeseleccionaGrid DataGrid1
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    txtAux(0).BackColor = vbWhite
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    
    
    'Abril 2015
    'Para ver si permite descuento
    Dim vPreFact As CPreciosFact
    Set vPreFact = New CPreciosFact
    vPreFact.CodigoArtic = CStr(Data2.Recordset!codArtic)
    vPreFact.CodigoClien = Text1(4).Text
    vPreFact.FijarTarifaActividad
    'para ver si bloqueamos el TXT de descuentos
    vPreFact.ObtenerPrecio True, Text1(1).Text, "", ""
    txtAux(6).Enabled = vPreFact.DtoPermitido
    txtAux(7).Enabled = vPreFact.DtoPermitido
    Set vPreFact = Nothing
    
    
    
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim Cad As String
Dim vTipoMov As CTiposMov
Dim NumOferElim As Long
Dim Borra As Boolean
    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "Cabecera de Ofertas." & vbCrLf
    Cad = Cad & "-----------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Oferta:            "
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    Borra = False
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then Borra = True
    
    If Borra And vParamAplic.NumeroInstalacion = vbEuler Then
        Cad = String(50, "*") & vbCrLf
        Cad = Cad & Cad & Cad
        Cad = Cad & vbCrLf & vbCrLf & "¿SEGURO QUE QUIERE BORRAR LA OFERTA ? " & vbCrLf & vbCrLf & Cad
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Borra = False
    End If
    If Borra Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumOferElim = Data1.Recordset.Fields(0).Value
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumOferElim
        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Oferta", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim Sql As String
Dim ImpReciclado As Single
Dim pos As Integer

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    Me.SSTab1.Tab = 0
    ModificaLineas = 3 'Eliminar
    
    Sql = "¿Seguro que desea eliminar la línea de Oferta?     "
    Sql = Sql & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    Sql = Sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    Sql = Sql & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    

    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        Sql = "Delete from " & NomTablaLineas & " WHERE " & ObtenerWhereCP
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute Sql
        
        
           'Llegado aqui, si tiene Punto verde(tasa ecologica)
        'Y el cliente tiene tasa recliclado
        If ClienteConTasaReciclado Then
            Sql = CStr(Data2.Recordset!codArtic)
            If ArticuloConTasaReciclado(Sql, ImpReciclado) Then
                
               'Si el articulo siguiente es PV entoces lo updatearemos
               Sql = Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
               Sql = Sql & " and numlinea"
            
               pos = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
               Sql = DevuelveDesdeBD(conAri, "codartic", NomTablaLineas, Sql, CStr(pos))
               'En SQL tengo el codarti de la linea SIGUIENTE
               'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
               If Sql = vParamAplic.ArtReciclado Then
               
                    Sql = "DELETE FROM " & NomTablaLineas
                    'WHERE
                    'Si el articulo siguiente es PV entoces lo updatearemos
                    Sql = Sql & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
                    Sql = Sql & " and numlinea=" & pos
                    conn.Execute Sql
              End If  'linea siguiente con codarti=puntoverde
            End If  'articulo con reciclado
        End If ' de cliente con tasa reciclado
            
        
        
        
        
        Text2(16).Text = ""
        'Text2(18).Text = ""
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
        
        Modo = 5
        cmdRegresar_Click
        
'        CancelaADODC
    End If
    
    

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        
        '25 Enero 2010
        'probando. Falta conmprobar que ha ido bien desde entonces
        'If DataGrid1.Row >= 0 Then
        '    DeseleccionaGrid DataGrid1
        '    DataGrid1.Bookmark = 1
        'End If
        
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
        cmdCancelar.Cancel = True
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
'        If Data1.Recordset.EOF Then
'            MsgBox "Ningún registro devuelto.", vbExclamation
'            Exit Sub
'        End If
'        Cad = Data1.Recordset.Fields(0) & "|"
'        Cad = Cad & Data1.Recordset.Fields(1) & "|"
'        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo EKeyPress
    
    If KeyAscii = 27 Then 'ESC
        If Modo = 5 Then 'Modo Lineas
            cmdRegresar_Click
        ElseIf Modo = 0 Or Modo = 2 Then 'Estamos en Cabecera
            Unload Me
        End If
    End If
    
EKeyPress:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Aux As String
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If (Modo <> 2 And Modo <> 5) Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        If X > 1750 And X < 8000 Then
            Aux = DBLet(DataGrid1.Columns(4).Value, "T") & "  "
            Select Case DataGrid1.Columns(8).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
            Me.DataGrid1.ToolTipText = Trim(Aux & "  " & Me.DataGrid1.ToolTipText)
        Else
            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        'Poner descripcion de ampliacion lineas
        Text2(16).Text = DevuelveDesdeBDNew(1, NomTablaLineas, "ampliaci", "numofert", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
        'Text2(18).Text = C2
    Else
        Text2(16).Text = ""
        'Text2(18).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If DatosOferta <> "" Then
            PonerModo 1
            Text1(0).Text = DatosOferta
            HacerBusqueda
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim Im As Image
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    
    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(19).Picture
    For Each Im In imgBuscar
        Im.Picture = frmPpal.imgListComun.ListImages(1).Picture
    Next
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 25
    
    
    

    ' ICONITOS DE LA BARRA
    
    
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
       
    'Lineas
    With Me.ToolbarAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        '3 4 5
        
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 32
        .Buttons(6).Image = 39
        .Buttons(7).Image = 38
        
    End With
    With Me.ToolbarAux(1)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        
        
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        
        .Buttons(1).Image = 21 '
        .Buttons(2).Image = 30 '
        .Buttons(3).Image = 22  '
        .Buttons(4).Image = 11 '
        
        'Herbelca. Facturar a FAZ
        .Buttons(5).Image = 20 '
        
        
    End With



    
    Me.SSTab1.Tab = 0
      
    kCampo = 0
    If vParamAplic.QueEmpresaEs = vbHerbelca Then
        If vUsu.Nivel < 2 Then
            If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then kCampo = 1
        End If
    End If
    'Factura
    Toolbar2.Buttons(5).visible = kCampo = 1
    
      
    LimpiarCampos   'Limpia los campos TextBox
    CargarComboFacturacion
    CodTipoMov = "OFE"
    VieneDeBuscar = False 'Para el CPostal
   
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = DevuelveTextoDepto(True)
        
    If vParamAplic.TieneCRM Then
        Label1(18).Caption = "Observaciones CRM"
    Else
        Label1(18).Caption = "Observaciones internas"
    End If
    
    
    CargarCombo_Tabla cboMotTra, "smotitrasofer", "codmotiv", "descmotivo", , False, "descmotivo"


    NombreTabla = "scapre"
    NomTablaLineas = "slipre" 'Tabla lineas de Ofertas
    Me.Caption = "Ofertas Clientes"
    

    Ordenacion = " ORDER BY numofert "

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'Direcion envio SOLO si esta en parametros
    Label1(6).visible = vParamAplic.DireccionesEnvio
    imgBuscar(7).visible = vParamAplic.DireccionesEnvio
    Text1(33).visible = vParamAplic.DireccionesEnvio
    Text2(33).visible = vParamAplic.DireccionesEnvio
    
    
    '
    SSTab1.TabVisible(3) = EulerParam <> ""
    txtAux(9).Left = 30000
        
    'Filtro
    CargarFiltro
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where false"
    Data1.Refresh
'    If DatosADevolverBusqueda = "" Then
    If Me.DatosOferta = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True

    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    PrimeraVez = True
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkAceptado(0).Value = False
    Me.cboFacturacion.ListIndex = -1
    Me.cboMotTra.ListIndex = -1
    Text3(0).Text = "BASE IMP."
    
    CargaArchivo ""
    ListView1.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera2 = 0 Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
        
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        Else
            If EsCabecera2 = 1 Then 'Llama desde VerTodos del Form
                'Llama desde Prismatico Direcciones/Departamentos
                Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                'DESDE ENVIO
                Text1(33).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(33).Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
    HaDevueltoDatos = True
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmCV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(Indice).Text)
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
        If EsCabecera2 = 1 Then 'Llama desde VerTodos del Form
            Text1(12).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
        Else
            'DESDE ENVIO
            Text1(33).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
            Text2(33).Text = RecuperaValor(CadenaSeleccion, 2)
        End If
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen el Listado de Oferta para generar el Pedido
Dim vSQL As String

    BuscaChekc = CadenaSeleccion


    'Construimos parte de la SQL para insertar en Pedidos
    vSQL = ""
    vSQL = " '" & Format(RecuperaValor(CadenaSeleccion, 2), FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    vSQL = vSQL & Format(RecuperaValor(CadenaSeleccion, 3), FormatoFecha) & "' as fecentre, " 'Fecha entrega
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 4) & " as sementre, " 'Sem entrega
    vSQL = vSQL & "0 as visadore, " & "codclien, nomclien, domclien, codpobla, pobclien, proclien, nifclien, "
    vSQL = vSQL & "telclien, coddirec, nomdirec, referenc, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Operador de Pedido
    vSQL = vSQL & "codagent, codforpa, dtoppago, dtognral, tipofact, observa01, observa02, observa03, "
    vSQL = vSQL & "observa04, observa05, 0 as servcomp,0 as restoped, " & Text1(0).Text & " as numofert, '" 'Nº Oferta
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecofert " 'Fecha Oferta
    '09/12/2010 mailconfir
    vSQL = vSQL & ",mailconfir,observacrm " 'Fecha Oferta
    '30/12/2013
    vSQL = vSQL & ",coddiren"
    CadenaSQL = vSQL
End Sub


Private Sub frmOfe_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmPlant_CargarPlantillas()
Dim Rs As ADODB.Recordset
Dim RSLineas As ADODB.Recordset
Dim Sql As String, devuelve As String
Dim codAlmac As String
'codTarif As String
Dim cantidad As Currency
Dim NumCajas As Integer, RestoUnid As Currency
Dim Precio As String, Dto1 As String, Dto2 As String
Dim OrigP As String 'De donde viene el precio: promocion, precio especial,º.
Dim CPrecioFact As CPreciosFact
Dim COntadorLInea As Integer

    Screen.MousePointer = vbHourglass
    
    'Si se ha seleccionado alguna plantilla para añadir sus lineas a la Oferta
    '(cantidad de alguna linea de tmpscapla > 0), entonces añadimos todas las
    'lineas de esa oferta poniendo en cantidad de slipre de lineas de oferta
    'el resultado de multiplicar la cantidad de tmpscapla * cantidad de slipla
    Sql = "SELECT * FROM tmpscapla WHERE codusu=" & vUsu.Codigo & " AND cantidad>0 ORDER BY codplant"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Obtener el almacen por defecto del trabajador
    'Poner el Almacen por defecto del Trabajador
    codAlmac = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    
    
    'Obtener la tarifa del cliente. LO Fijare en la funcion
    'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")



    'Consigo el contador
    COntadorLInea = Val(SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP))
    
    While Not Rs.EOF  'Para cada plantilla
        'Añadimos todas las lineas de esa plantilla en la cantidad correcta en las
        'lineas de la Oferta
        Sql = "SELECT * FROM slipla WHERE codplant=" & Rs!codplant & " order by numlinea"
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RSLineas.EOF
            'Comprobar si el articulo se vende por cajas antes de entrar a la función
            devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RSLineas!codArtic, "T")
            If devuelve <> "" Then
            'Si se puede vender por cajas(devuelve>1) poner numero de unidades/caja en una linea con el
            'precio de caja, y otra linea con el resto unidades un precio unidad
                cantidad = (Rs!cantidad * RSLineas!cantidad)
                NumCajas = ObtenerNumCajas(CStr(cantidad), devuelve)
                RestoUnid = cantidad - (NumCajas * CInt(devuelve))
                'Obtener el precio a aplicar
                Set CPrecioFact = New CPreciosFact
                'CPrecioFact.CodigoLista = codTarif
                CPrecioFact.CodigoArtic = RSLineas!codArtic
                CPrecioFact.CodigoClien = Text1(4).Text
                CPrecioFact.FijarTarifaActividad
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, "")
                
                'Si PorCaja vuelve de ObtenerPrecio a False se aplica precio
                'de Unidad aunque se venda por cajas, ya que ha regresado con pvp de articulo
                Dto1 = CPrecioFact.Descuento1
                Dto2 = CPrecioFact.Descuento2
                Set CPrecioFact = Nothing
                    
                If PorCaja And NumCajas > 0 Then 'El Articulo se Vende Por Cajas y Cantidad supera la cant en 1 caja
                    'Obtener el precio y los descuentos adecuados
                    'Insertar 2 lineas: 1 linea con la cantidad que se puede vender en cajas y al precio de caja
                    InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, NumCajas * CInt(devuelve), Precio, Dto1, Dto2, OrigP, COntadorLInea
                    '2 linea con el resto de la cantidad que no llega a una caja a precio de unidad
                    If RestoUnid > 0 Then InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, RestoUnid, Precio, Dto1, Dto2, OrigP, COntadorLInea
'                    Else
'                        InsertarLineaDePlantilla rsLineas!codArtic, codAlmac, codTarif, Cantidad, 0
'                    End If
                Else 'No llega a una caja
                    InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, cantidad, Precio, Dto1, Dto2, OrigP, COntadorLInea
                End If
            End If
            RSLineas.MoveNext
        Wend
        RSLineas.Close
        Set RSLineas = Nothing
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Borrar de la Tabla Temporal (tmpscapla) los registros insertados tras añadir
    'las lineas de las plantillas seleccionadas
    DescargarDatosTMP
    'Actualizar el Grid con las lineas Añadidas
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub frmTOferta_CargarOferta2(NumOfe As String)
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim numlinea As String, vWhere As String
Dim i  As Integer
Dim CopiaDesdeHco As Boolean
    On Error GoTo ECargarOferta
    
    'Si se ha seleccionado alguna oferta para añadir sus lineas a la Oferta
    If NumOfe = "" Then Exit Sub
    CopiaDesdeHco = False
    If Mid(NumOfe, 1, 1) = "H" Then
        'Desde hco
        CopiaDesdeHco = True
        NumOfe = Mid(NumOfe, 2)
    End If
    
    Screen.MousePointer = vbHourglass
    If CopiaDesdeHco Then
        Sql = DevuelveDesdeBD(conAri, "distinct(fecofert)", "slhpre", "numofert", RecuperaValor(NumOfe, 1))
        'por si hubiera mas de una fecha. Solo cojo una
        
        Sql = "Select * from slhpre where numofert=" & RecuperaValor(NumOfe, 1) & " AND fecofert='" & Format(Sql, FormatoFecha) & "'"
    Else
        Sql = "Select * from " & NomTablaLineas & " where numofert=" & RecuperaValor(NumOfe, 1)
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF  'Para cada linea de oferta
        'Obtener el siguiente numero de linea
        vWhere = ObtenerWhereCP
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        Sql = "INSERT INTO " & NomTablaLineas & " (numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,origpre,esopcion) "
        Sql = Sql & " VALUES(" & Text1(0).Text & ", " & numlinea & ", " & Rs!codAlmac & ", " & DBSet(Rs!codArtic, "T") & ", " & DBSet(Rs!NomArtic, "T") & ", "
        Sql = Sql & DBSet(Rs!Ampliaci, "T", "S")
        Sql = Sql & ", " & DBSet(Rs!cantidad, "N") & ", " & DBSet(Rs!precioar, "N") & ", " & DBSet(Rs!dtoline1, "N") & ", " & DBSet(Rs!dtoline2, "N") & ", " & DBSet(Rs!ImporteL, "N") & ", "
        Sql = Sql & DBSet(CStr(Rs!origpre), "T", "S") & ","
        'SQL = SQL & Rs!esopcion & "," & DBSet(DBLet((Rs!observa), "T"), "T", "S") & ")"
        Sql = Sql & Rs!esopcion & ")"
        conn.Execute Sql
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    Sql = RecuperaValor(NumOfe, 2)  'Copio observaciones
    vWhere = RecuperaValor(NumOfe, 3)  'Copio datos carta
    i = Val(Sql) + Val(vWhere)
    If i > 0 Then
        'Cargo en RS la oferta
        If CopiaDesdeHco Then
            Sql = DevuelveDesdeBD(conAri, "distinct(fecofert)", "slhpre", "numofert", RecuperaValor(NumOfe, 1))
            'por si hubiera mas de una fecha. Solo cojo una
            Sql = "Select * from schpre where numofert=" & RecuperaValor(NumOfe, 1) & " AND fecofert='" & Format(Sql, FormatoFecha) & "'"
        Else
            Sql = "Select * from " & NombreTabla & " where numofert=" & RecuperaValor(NumOfe, 1)
        End If
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'UPDATEAMOS los campos de la oferta de observaciones
            Sql = ""
            If RecuperaValor(NumOfe, 2) = "1" Then 'Copio observaciones
                
                
                For i = 1 To 5
                    vWhere = "observa0" & i
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(Rs.Fields(vWhere), "T"), "T", "S")
                    Sql = Sql & numlinea
                Next i
                
                '15 Marzo 2010. Cunado pone copiar observacinoes TB tiene que copiar el campo concepto
                Sql = Sql & ", concepto = " & DBSet(DBLet(Rs!concepto, "T"), "T")
                '12 Sept 2012  Copiara tb el plazo de entrega
                Sql = Sql & ", plazos01 = " & DBSet(DBLet(Rs!plazos01, "T"), "T", "S")
                Sql = Sql & ", plazos02 = " & DBSet(DBLet(Rs!plazos02, "T"), "T", "S")
                
                
            End If
            
            If RecuperaValor(NumOfe, 3) = "1" Then 'Copio carta
                For i = 1 To 5
                    vWhere = "asunto0" & i
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(Rs.Fields(vWhere), "T"), "T", "S")
                    Sql = Sql & numlinea
                Next i
                
            End If
       
            Sql = Mid(Sql, 2)  'quito la primera coma
            Sql = Sql & " WHERE numofert = " & Text1(0).Text
            Sql = "UPDATE " & NombreTabla & " SET " & Sql
            Rs.Close
        conn.Execute Sql
        PosicionarData  'vuelvo a cargar los datos
        PonerCampos
        Else
            MsgBox "Error buscando oferta destino: " & Text1(0).Text & ".  EOF", vbExclamation
        End If
    End If
    
    
    Set Rs = Nothing

    'Actualizar el Grid con las lineas Añadidas
    If i = 0 Then CalcularDatosFactura   'Si no mete obser y carta que carge los totales
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
    
ECargarOferta:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Traer de otra Oferta.", Err.Description
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte


    Indice = 0
    If Index <> 8 Then
        
        If Modo = 2 Or Modo = 0 Then Indice = 1
    Else
        'observaciones
        If Modo = 0 Then Indice = 1
    End If
    
    If Indice = 1 Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    TerminaBloquear

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            
            Set frmC = New frmBasico2
            AyudaClientes frmC, Text1(4).Text
            Set frmC = Nothing
           
            Indice = 5
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 4
                txtAnterior = Text1(4).Text
            End If
        Case 1 'NIF para cliente de Varios
'            Set frmCV = New frmFacClientesV
'            frmCV.DatosADevolverBusqueda = "0"
'            frmCV.Show vbModal
'            Set frmCV = Nothing
            Indice = 6
            Set frmCV = New frmBasico2
            AyudaClientesV frmCV, Text1(Indice)
            Set frmCV = Nothing
            
            
        Case 2 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else

                EsCabecera2 = 1
                'ANTES
                '01/DICIEMBRE/2010   DAVID
                'MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
                LanzaBusquedaDpto True, Indice
                
             End If
             
        Case 3 'Realizada Por Trabajador
            Indice = 3
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            Indice = 14
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
            
            PonerFoco Text1(Indice)
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
'            Set frmFP = Nothing
            
        Case 5 'Agente
            Indice = 17
            PonerFoco Text1(Indice)
'            Set frmA = New frmFacAgentesCom
'            frmA.DatosADevolverBusqueda = "0"
'            frmA.Show vbModal
'            Set frmA = Nothing
            Set frmA = New frmBasico2
            AyudaAgentesComerciales frmA, Text1(Indice), , True
            Set frmA = Nothing

            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            
        Case 7
            If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera2 = 2
                'MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 33
                LanzaBusquedaDpto False, Indice
                
             End If
             
        Case 8
                
                CadenaDesdeOtroForm = Text2(16).Text
                frmFacClienteObser.Modificar = Modo = 5 And ModificaLineas > 0
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 5 And ModificaLineas > 0 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text2(16).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
        If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub


Private Sub mnGenPedido_Click()
'Pasar una Oferta a Pedido
Dim devuelve As String
Dim CA As Collection
Dim LineasQueSepasan As String
Dim TextoPrecioMinimo As String

    'Comprobar que hay una Oferta seleccionada
    If Text1(0).Text = "" Then Exit Sub
    
    
    '23 Abril 2010
    'NO puede tener la fecha entraga a NULO
    If Trim(Text1(2).Text) = "" Then
        MsgBox "Fecha entrega no puede ser nula", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    '17 Diciembre 2010
    If EsClienteBloqueado2(Text1(4).Text, 2, False, False) Then Exit Sub
    
    
    
    'Si esta YA en pedid, o archivada
    devuelve = ""
    If Text1(37).Text <> "" Then
        devuelve = "La oferta esta archivada: " & Text1(37).Text & " - " & Me.cboMotTra.Text
    Else
        If Text1(38).Text <> "" Then devuelve = "La oferta esta en el pedido: " & Text1(38).Text & " - " & Text1(39).Text
    End If
    If devuelve <> "" Then
        MsgBox devuelve, vbExclamation
        Exit Sub
    End If
    
    
    'Comprobar que la Oferta seleccionada esta aceptada
    devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "aceptado", "numofert", Text1(0).Text, "N")
    If devuelve = "0" Then
        devuelve = "La Oferta debería estar aceptada para pasarla a Pedido. " & vbCrLf & " ¿Continuar?"
        If MsgBox(devuelve, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        devuelve = "1"
    End If
    
    TextoPrecioMinimo = ""
    
    
    
    If devuelve = "1" Then
        
        If vParamAplic.LogCambioPrecDto Or vParamAplic.NumeroInstalacion = vbHerbelca Then
            'Comprobaremos si ha cambiado el precio de cada linea
            CadenaSQL = ""
            
            PorDebajoPrecioMinimo = False
            ComprobarPrecioDtoArticulo CA
            devuelve = CadenaSQL
            CadenaSQL = ""
            If PorDebajoPrecioMinimo Then
                MsgBox "Precio articulo inferior al mínimo permitido" & devuelve, vbExclamation
                If vUsu.Nivel > 0 Then Exit Sub
            
            End If
            
            If Not HaCambiadoprecioDto(CA) Then Exit Sub
            
            
            
        End If
    
        devuelve = Trim(DBLet(Data1.Recordset!obscompra, "T"))
        If devuelve <> "" Then
            MsgBox devuelve, vbInformation, "Observaciones compra"
            devuelve = ""
        End If
        
        
        'Si pasa la oferta completa o no
        Precio = ""   'Para que luego NO vuelva a hacer la pregunta
        BuscaChekc = "" 'para los datos de pedido, fecha y trabajador y updatear l'oferta
        TieneOpciones (False)
    
        
        'Pedir: Operador de Pedido, fecha pedido, y fecha entrega (calcular semana)
'        AbrirListadoOfer (37) '37: Pedir datos para Pedido (NO IMPRIME LISTADO)
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 37
        frmList.codClien = Text1(4).Text
        frmList.FecEntre = Text1(2).Text
        frmList.Show vbModal
        Set frmList = Nothing
        If CadenaSQL = "" Then Exit Sub
        
        'Agosto2011
        '---------------------------------------------
        'Las ofertas pueden llevar opciones.
        'Si llevan opciones mostraremos un lw con las opciones
        'LineasQueSepasan: Llevara las lineas que SI se van a pasar
        '                   Si esta vacio son todas
        LineasQueSepasan = ""
        If TieneOpciones(False) Then
        
            '------------------------------------------------------------
            '
            CadenaDesdeOtroForm = ""
            frmFacTrasOfertaOpciones.Caption = Text1(0).Text
            frmFacTrasOfertaOpciones.Label1.Caption = Text1(5).Text
            frmFacTrasOfertaOpciones.Show vbModal
            If CadenaDesdeOtroForm = "NO" Then Exit Sub
            If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)                  'QUITO LA PRIMERA COMA
            LineasQueSepasan = CadenaDesdeOtroForm
            
            
        End If
        Precio = ""
        'Tenemos en CadenaSQL parte de la SELECT para insertar el Pedido
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        PasarOfertaAPedido CadenaSQL, CA, LineasQueSepasan

        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            LimpiarDataGrids
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnImpFactProF_Click()
'Imprime factura pro forma
    BotonImprimirProForma (59) '59: Informe Factura ProForma
End Sub

Private Sub mnImpOferta_Click()
'Imprime una Oferta
       frmListadoOfer.NumCod = Text1(0).Text   'Nº de Oferta
       frmListadoOfer.FecEntre = Text1(1).Text 'Fecha de Oferta
       'FALTA###
       'If EsHistorico Then
        '    AbrirListadoOfer (35) '35: Informe Historico de Ofertas
       'Else
            AbrirListadoOfer (31) '31: Informe de Ofertas
       'End If
End Sub

Private Sub mnImpRecordatorio_Click()
    frmListadoOfer.NumCod = Text1(0).Text
    frmListadoOfer.codClien = Text1(4).Text
    AbrirListadoOfer (32) '32: Recordatorio de Ofertas
End Sub

Private Sub mnImpValoracion_Click()
    frmListadoOfer.codClien = Text1(4).Text
    frmListadoOfer.NumCod = Text1(0).Text 'Nº de Oferta
    AbrirListadoOfer (33) '33: Valoracion de Ofertas
End Sub

''Private Sub mnLineas_Click()
''
''    BotonMtoLineas 0, "Ofertas"
''End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea False
    Else 'Añadir Cabecera de Ofertas
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub


Private Sub mnOferta_Click()
    'Añadir las lineas de otra oferta a la Oferta
    Set frmTOferta = New frmFacTraerOferta
    frmTOferta.Antiguo = False
    frmTOferta.Show vbModal
    Set frmTOferta = Nothing
    
    
    
End Sub

Private Sub mnPlantillas_Click()
'Añadir Plantilla de Oferta
    Set frmPlant = New frmFacCargaPlantilla
    frmPlant.Show vbModal
    Set frmPlant = Nothing
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



Private Sub HacerDocumentosPDF(Index As Integer)



    Debug.Assert False
    'If EsHistorico Then Exit Sub
    
    
    If Modo <> 2 Then Exit Sub
    
    If Index > 0 Then
        CadenaDesdeOtroForm = ""
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then CadenaDesdeOtroForm = "OK"
        End If
        If CadenaDesdeOtroForm = "" Then
            MsgBox "Seleccione un documento", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If Index = 2 Then
        If MsgBox("Seguro que desea eliminar el documento seleccionado?", vbQuestion + vbYesNo) = vbYes Then
            Precio = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
            If EliminarArhivoPDFOferta(Precio) Then cargaDocumentos
        End If
    Else
        If Index = 1 Then Exit Sub
        If CarpetaDestinoOferta = "" Then
          '  If Year(Data1.Recordset!fecofert) < 2017 Then
          '      MsgBox "Carpeta destino no esta establecida", vbExclamation
          '      Exit Sub
          '  End If
            
            'Creamos la estructura (Deberia haber sido creada con anterioridad, al dar nuevo. Pero para las de 2016 dejo pasar
            CrearCarpeta
            
        End If
        CadenaDesdeOtroForm = ""
        If Index = 1 Then CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
        frmEuler.Opcion = Index
        frmEuler.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 0 Then
                'Insertamos documento
                Precio = RecuperaValor(CadenaDesdeOtroForm, 1)
                If Dir(Precio, vbArchive) = "" Then
                    MsgBox "No existe fichero", vbExclamation
                    
                Else
                    
                    If CarpetaDestinoOferta = "" Then
                        'Crear estructura
                        
                        MsgBox "falta Crear estructura", vbExclamation
                        Exit Sub
                    End If
                    Precio = RecuperaValor(CadenaDesdeOtroForm, 2)
                    txtAnterior = NombreArchivoEULER(Precio)
                
                    If CopiaArhivoPDFOfertaEuler(CarpetaDestinoOferta, RecuperaValor(CadenaDesdeOtroForm, 1)) Then cargaDocumentos
                    
                
                End If
            Else
                'Modificar
                Precio = "UPDATE sliprepdfs SET ficheroDesc = " & DBSet(CadenaDesdeOtroForm, "T")
                Precio = Precio & " WHERE numofert = " & Text1(0).Text & " AND numlinea = " & Mid(ListView1.SelectedItem.Key, 2)
                ejecutar Precio, False
                ListView1.SelectedItem.Text = CadenaDesdeOtroForm
            End If
        End If
    End If







End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
'Documentos
Private Sub ListView1_Click()
Dim Cad As String
Dim C As String
    Cad = ""
    If CarpetaDestinoOferta <> "" Then
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then
                If LCase(Right(ListView1.SelectedItem.SubItems(1), 3)) = "pdf" Then Cad = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
            End If
        End If
    End If
    
    CargaArchivo Cad
    'Debug.Assert False
    'cmdPDF.visible = cad <> ""
  
        
End Sub

Private Sub ListView1_DblClick()
Dim C As String
    
    C = ""
    If CarpetaDestinoOferta <> "" Then
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then C = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
        
        End If
    End If
    
    
    If C <> "" Then LanzaVisorMimeDocumento Me.hwnd, C
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub SSTab1_Click(PreviousTab As Integer)
    Label1(35).visible = SSTab1.Tab = 0
    Text2(16).visible = Modo > 3 Or SSTab1.Tab = 0
    imgBuscar(8).visible = SSTab1.Tab = 0
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer
    txtAnterior = Text1(Index).Text
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False
    If (Index <> 31 And Index <> 32) Then ConseguirFoco Text1(Index), Modo, cadkey
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 30 And KeyCode = 40 Then
        Me.SSTab1.Tab = 2
        PonerFoco Text1(31)
    Else
        If Not Text1(Index).MultiLine Then KEYdown KeyCode
        
         If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        
            If Text1(Index).Text = "" Then
                Ind = -1
                Select Case Index
                Case 3
                    Ind = 3
                Case 4
                    Ind = 0
                Case 6
                    Ind = 1
                Case 9
                    Ind = 6
                Case 12
                    Ind = 2
                Case 17
                    Ind = 5
                Case 14
                    Ind = 4
                Case 33
                    Ind = 7
                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
        
        
        
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then
        If KeyAscii = 13 And (Index = 30 Or Index = 32) Then 'ENTER
            If Index = 32 Then
    '            PonerFocoBtn Me.cmdAceptar
            ElseIf Index = 30 Then
                Me.SSTab1.Tab = 2
                PonerFoco Text1(31)
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
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
       
             
    'Han pulsado el mas
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(Index).Text = ""
        Exit Sub
    End If
       
       
     If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
 
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Por si no ha cambiado nada
    If txtAnterior = Text1(Index).Text Then Exit Sub
    
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
            If vParamAplic.NumeroInstalacion = vbFenollar And Text1(Index).Text <> "" Then Text1(2).Text = Text1(1).Text
            
        Case 3 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else
                    PonerDatosCliente (Text1(Index).Text)
                End If
            Else
                LimpiarDatosCliente
            End If
            If Modo <> 1 Then
                'Si no estamos en busqueda, y se ha equivado poniendo cliente... vuelve a cliente
                If Text1(Index).Text = "" Then PonerFoco Text1(Index)
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
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 12 'Cod. Direc
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionada tiene esa direccion
           If PonerDptoEnCliente Then
                'Comprobar que el cliente tiene mantenimientos en esa dired/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then
                ComprobarRefObligatoria
            End If
            
        Case 14 'Forma de Pago
            If Me.SSTab1.Tab = 0 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then  'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
        Case 33
        
        
            devuelve = ""
            If Text1(Index).Text <> "" Then
                
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation

                    PonerFoco Text1(Index)
                Else
                    'Comprobar codenvio
                    devuelve = DevuelveDesdeBDNew(1, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", "", "coddiren", Text1(33).Text, "N")
                
                    If devuelve = "" Then
                        
                        MsgBox "No existe la dirección de envio:" & Text1(Index).Text, vbInformation
                        Text1(33).Text = ""
                        PonerFoco Text1(33)
                    End If
                End If
                
            Else
                PonerFoco Text1(Index)
            End If
            Text2(Index).Text = devuelve
        Case 37
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If vUsu.CodigoAgente > 0 Then
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & "( codagent =" & vUsu.CodigoAgente
        If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then cadB = cadB & " OR codclien IN (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
        cadB = cadB & ")"
    End If
    PonerFiltro cadB  'Añadira al SQL el filtro establecido
    
    
    If chkVistaPrevia = 1 Then
        EsCabecera2 = 0
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera2 = 0 Then
            'cad = cad & ParaGrid(Text1(0), 15, "Nº Oferta")
            'cad = cad & ParaGrid(Text1(1), 20, "Fecha Ofer.")
            'cad = cad & ParaGrid(Text1(4), 15, "Cliente")
            'cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
            'tabla = NombreTabla
            'If EsHistorico Then
            '    Titulo = "Histórico de Ofertas"
            '    devuelve = "0|1|"
            'Else
            '    Titulo = "Ofertas"
            '    devuelve = "0|"
            'End If
            Set frmOfe = New frmBasico2
            CadenaConsulta = ""
            AyudaOfertas frmOfe, , cadB, True
            Set frmOfe = Nothing
            
            If CadenaConsulta <> "" Then
                'FALTA#####
                Cad = DBSet(RecuperaValor(CadenaConsulta, 2), "F")
                
                CadenaConsulta = "numofert =" & RecuperaValor(CadenaConsulta, 1) & " AND fecofert = " & Cad
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaConsulta & " " & Ordenacion
                Cad = "" 'para que no haga nada bbajo
                PonerCadenaBusqueda
                
            Else
                CadenaConsulta = Data1.Recordset.Source
            End If
            
'            devuelve = "0|"
    Else 'Llama desde lineas, para cargar solo los depart/direc. del cliente seleccionado
        If EsCabecera2 = 1 Then
            'DEPARTAMENTOS DIRECCIONES
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            Else
                Titulo = "Obras Cliente: "
                Desc = "Obras"
            End If
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
            Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
            tabla = "sdirec"
            devuelve = "0|1|"
        
        Else
            'direcciones de envio
            'Tipo herbelca
            Titulo = "Dir. envio de: " & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Cod. envio|sdirenvio|coddiren|N||15·"
            Cad = Cad & "Descripcion envio|sdirenvio|nomdiren|T||35·"
            tabla = "sdirenvio"
            devuelve = "0|1|"
        End If 'de cabecera=1
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        If EsCabecera2 > 0 Then frmB.Label1.FontSize = 11
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
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(kCampo)
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If
    Text3(0).Text = "BASE IMP."
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
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    
    'Trabajador pedido
    Text2(40).Text = IIf(Text1(40).Text = "", "", PonerNombreDeCod(Text1(40), conAri, "straba", "nomtraba", "codtraba"))
    
    
           
    If vParamAplic.DireccionesEnvio Then Text2(33).Text = DevuelveDesdeBDNew(conAri, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", , "coddiren", Text1(33).Text, "N")
      
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    CalcularDatosFactura
    
    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        If Text1(32).Text <> "" Then MsgBox Text1(32).Text, vbInformation
    End If
    
    cargaDocumentos
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    'toolbar aux
    BotonesToolBarAux
    
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
    lblF.Caption = ""
    
 
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    

    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    cmdRegresar.visible = Modo = 5 And ModificaLineas = 0
    
    
        
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Poner Flechas de desplazamiento visibles
    B = (Modo = 2)
    NumReg = 1
    If B And Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DespalzamientoVisible NumReg = 2
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Poner siempre el campo numOferta (contador) bloqueado, excepto cuando
    'estamos en modo de Busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True
    
    'Siempre bloqueado tambien los datos del pedido. Solo en busqueda
    BloquearTxt Text1(38), B, False
    BloquearTxt Text1(39), B, False
    BloquearTxt Text1(40), B, False
    
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 56
        BloquearTxt Text3(i), True
    Next i
    
    
    
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    For i = 46 To 48
        Text3(i).BackColor = &HFFFFC0
        Text3(i + 6).BackColor = &HFFFFC0
    Next i
    'Campos total Factura en verde
    Text3(55).BackColor = &HC0FFC0
    Text3(56).BackColor = &HC0FFC0    'Tatal factura
    '---------------------------------------------------
    
    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Me.cboFacturacion.Enabled = B
    Me.chkAceptado(0).Enabled = B
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    
    
    
    
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    BloquearCmb Me.cboMotTra, Not B
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(1).visible = False
    Me.imgBuscar(8).Enabled = Modo <> 0
       
    'Modo Linea de Ofertas
    Me.Label1(35).visible = SSTab1.Tab = 0
    Me.Text2(16).visible = SSTab1.Tab = 0
    BloquearTxt Text2(16), True
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    'Realizado por
    If vParamAplic.NumeroInstalacion = vbHerbelca Then BloquearCampoTrabajador
        
    
    
    
    
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
Dim B As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then Exit Function
   
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
     If Trim(Text1(4).Text) <> "" Then

        devuelve = DevuelveDesdeBDNew(1, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            B = False
        End If
    End If
    
    If B Then
        'Lleva direcciones de envio. Comprobamos que la que ha puesto existe...
        If vParamAplic.DireccionesEnvio Then
            If Text1(33).Text = "" Xor Text2(33).Text = "" Then
                MsgBox "Dirección de envio INCORRECTA", vbExclamation
                B = False
            End If
            'Ha puesto un codenvio y parece ser que existe... LO COMPURBEO que no hay referenciales
            If B And Text1(33).Text <> "" Then
                devuelve = DevuelveDesdeBDNew(1, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", "", "coddiren", Text1(33).Text, "N")
                If devuelve = "" Then
                    MsgBox "NO existe la dirección de envio: " & Text1(33).Text, vbExclamation
                    PonerFoco Text1(33)
                    B = False
                End If
            End If
         End If 'de direnvii
    End If 'de b=true
            
    If B Then
        If EsDeVarios Then
            If vParamAplic.FrasMostradorSerieDistinta Then
                'Tiene contadores distintos.... FORMA DE PAGO deberia ser efec o tartje
                devuelve = DevuelveDesdeBDNew(1, " sforpa", "tipforpa", "codforpa", Text1(14).Text)
                If devuelve <> "0" And devuelve <> "6" Then
                    If MsgBox("La forma pago deberia ser efectivo o tarjeta.   ¿Continuar? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
                    If Not B Then PonerFoco Text1(14)
                End If
                devuelve = ""
            End If
        End If
    End If

        
    
        
        
    If B Then
        'Si el usuario conectado es agente SOLO puede ponerse a el como agente
        
        If vUsu.CodigoAgente > 0 And Val(Text1(17).Text) <> vUsu.CodigoAgente Then
            'MAL
            MsgBox "Agente distinto del conectado", vbExclamation
            B = False
        End If
    End If
            
    If Not B Then Exit Function
          
          
    'Marzo 2012
    'Si pone motivo traspaso tiene que poner fecha
    If Me.Text1(37).Text <> "" Xor cboMotTra.ListIndex >= 0 Then
        
        devuelve = IIf(cboMotTra.ListIndex >= 0, "Motivo: " & cboMotTra.Text, "")
        MsgBox "Si indica fecha archivo debe indicar motivo archivo" & vbCrLf & devuelve, vbExclamation
        devuelve = ""
        If cboMotTra.ListIndex >= 0 Then cboMotTra.ListIndex = -1
        B = False
    End If
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea(ByRef DescuentoSuperiorPermitido As Boolean) As Boolean
Dim B As Boolean
Dim i As Byte
Dim vArtic As CArticulo
Dim Aux As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    'Si el articulo es SEPARADOR lo completo a ceros
    If txtAux(1).Text <> "" And txtAux(1).Text = vParamAplic.ArtSeparador Then ValoresRestoCamposSeparador True
        
    
    
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(6), txtAux(7), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(8).Text Then txtAux(8).Text = Aux
    
    
    
    
    
    
    B = True
    For i = 0 To txtAux.Count - 2  'el ultmi campo es codigiva
        If txtAux(i).Text = "" Then
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            B = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
        
    'Comprobar que existe de el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        B = False
        PonerFoco txtAux(1)

    End If
    
    
    
    
    
    If B Then
        '--------------------------------
        'Comprobaremos que esta vendiendo por encima del dto permitido
        'Preguntaremos si sigue o para
               
        GrabaLogCambioPrecioDto = False
        PorDebajoPrecioMinimo = False
        If B Then
            'Si todo ha ido bien..
            'Y lleva el parametro
            If vParamAplic.LogCambioPrecDto Then ComprobarCambioPrecioDto vArtic
            
            'En herbelca, si el precio es inferior al precio minimo
            If vParamAplic.NumeroInstalacion = 2 Then
                If PorDebajoPrecioMinimo Then
                    Aux = "Precio inferior al mínimo permitido"
                    
                    MsgBox Aux, vbExclamation
                    If vUsu.Nivel > 0 Then B = False
                        
                End If
            End If
        End If
        
    End If
    vArtic.LeerDatos vArtic.Codigo
    txtAux(9).Text = vArtic.TipoIVA
    
    Set vArtic = Nothing
    If vParamAplic.PtosAsignar > 0 Then
        If Me.txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
            MsgBox "No puede utilizar articulo de canje", vbExclamation
            B = False
        End If
        
        
        
        
    End If
    
    
    
    DatosOkLinea = B
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_GotFocus(Index As Integer)
     lblF.Caption = "" 'para que no ponga nada
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        PonerFocoBtn Me.cmdAceptar
        'KEYpress KeyAscii
    End If
    
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


'Private Sub hacerToolbarAntiguo(Indice As Integer)
'
'    If vParamAplic.QueEmpresaEs = 1 Then
'        'WHOSE. NO permitimos ni
'        If Indice = 11 Then Exit Sub
'        If Indice = 19 Then Exit Sub
'    End If
'
'
'
'
'    If Indice = 17 Or Indice = 19 Then
'        'Valoracion , fra proforma
'        'If TieneOpciones Then Exit Sub
'        TieneOpciones True   'Siempre dejo seguir
'    End If
'
'
'    Select Case Indice
'
'
'
'        Case 10: mnLineas_Click  'Lineas
'        Case 11:
'                If Modo = 5 Then
'                    'Insertar intercalando
'                    BotonAnyadirLinea True
'                Else
'                    mnGenPedido_Click 'Generar Pedido
'                End If
'        Case 12: mnPlantillas_Click ' Plantillas. Solo visible en Mantenimiento Lineas.
'        Case 13: mnOferta_Click 'Traer Lineas de Otra Oferta
'
'
'        Case 16 'Recordatorio
'            mnImpRecordatorio_Click
'        Case 17 'Valoracion
'            mnImpValoracion_Click
'        Case 18 'Imprimir
'            mnImpOferta_Click
'        Case 19 'Imprimir factura por forma
'            mnImpFactProF_Click
'
'
'        Case 20
'
'            If vUsu.Nivel > 1 Then Exit Sub
'            If Modo <> 2 Then Exit Sub
'            If Data1.Recordset.EOF Then Exit Sub
'
'
'            'Comprobacion
'
'            If Not PuedePasarFacuraFAZ Then Exit Sub
'
'            'Lanzaremos la pregunta del banco
'            'Vamos a generar un ALBARAN ALZ y desde ahi, una factura FAZ
'            CadenaDesdeOtroForm = Text1(1).Text
'            frmListado3.Opcion = 50
'            frmListado3.Show vbModal
'            If Mid(CadenaDesdeOtroForm, 1, 3) = "OK#" Then
'                'OK. Ha seleccionado el banco
'                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 4)
'                Screen.MousePointer = vbHourglass
'                lblIndicador.Caption = "Pasando ALZ"
'                lblIndicador.Refresh
'
'                'Cambiamos a conta de B
'                AbrirConexionConta True
'
'                PasarOfertaFacturaFAZ
'
'                'Reestablecemos CONTA normal
'                lblIndicador.Caption = ""
'                Screen.MousePointer = vbDefault
'            End If
'
'
'
'
'
'        Case 22: mnSalir_Click    'Salir
'
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento Indice  ' (Button.Index - btnPrimero)
'    End Select
'End Sub







Private Sub PonerOpcionesMenu()
Dim J As Byte

    On Error Resume Next
    
    PonerOpcionesMenuGeneral Me
        
    
    'If J < vUsu.Nivel Then Me.mnGenPedido.Enabled = False
    
    
    'toolbar aux
    BotonesToolBarAux
    
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub













Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean
    
    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim Sql As String
Dim numlinea As String
Dim VtaDtoSup As Boolean
Dim ImpReciclado As Single


    On Error GoTo EInsertarLinea

    InsertarLinea = False
    Sql = ""
    If DatosOkLinea(VtaDtoSup) Then 'Lineas de Ofertas
        'Conseguir el siguiente numero de linea
        
        If LineaIntercalar = 0 Then
            'INSERCION NORMAL
            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP)
        
        Else
            Sql = ObtenerWhereCP
            Sql = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 1 WHERE " & Sql & " and numlinea >= " & LineaIntercalar
            Sql = Sql & " order by numlinea desc " 'Para que empieza por las ultimas
            conn.Execute Sql
            numlinea = LineaIntercalar
        End If
        Sql = "INSERT INTO " & NomTablaLineas
        Sql = Sql & "(numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,esopcion,codigiva) "
        Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(3).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ", " 'Dto 2
        Sql = Sql & DBSet(txtAux(8).Text, "N") & ", " 'Importe
        Sql = Sql & DBSet(txtAux(5).Text, "T") & ", "
        Sql = Sql & Abs(Me.cboOpcion.ListIndex = 1) & ","
        'SQL = SQL & DBSet(Text2(18).Text, "T") & ") "   'sept 2011
        Sql = Sql & txtAux(9).Text & ")"
     End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLinea = True
        
        
        ' ---- [28/10/2010] (DAVID)
        'Esta linea lleva dto superior al permitido.
        'Lo saco fuera del trans
        If VtaDtoSup Then GrabarLogDtoSuperior "OFE", Text1(0), Text1(1).Text, numlinea, True
        
        ' ---- [13/01/2011] (DAVID)
        'Si ha cambiado, si tiene el parametro... todo esta ahi
        TrataCambioPrecioDto

       
       ' ---- [05/11/2010] (DAVID)
       'Tasa reciclado en ofertas
        If ClienteConTasaReciclado Then
            If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                'Insertamos la linea del reciclado
             
                
                
                Sql = "INSERT INTO " & NomTablaLineas
                Sql = Sql & "(numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre) "
                Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                Sql = Sql & DBSet(vParamAplic.ArtReciclado, "T") & ","
                Sql = Sql & DBSet(DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T"), "T") & ", Null, "
                Sql = Sql & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
                Sql = Sql & DBSet(ImpReciclado, "N") & ",0,0,"
                'Importe linea
                ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                Sql = Sql & DBSet(ImpReciclado, "N") & ", 'A')"
                conn.Execute Sql
                    
                
            End If 'articulo con sunida reciclado
        End If  'Cliente con tasa reciclado
        
        
       
        
    End If
    Exit Function

EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta" & vbCrLf & Err.Description
End Function


Private Function InsertarLineaDePlantilla(codArtic As String, codAlmac As String, cantidad As Currency, Precio As String, Dto1 As String, Dto2 As String, OrigP, ByRef numlinea As Integer) As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim Sql As String
'Dim NumLInea As String
Dim NomArtic As String
Dim Importe As String

    On Error GoTo EInsertarLinea

    InsertarLineaDePlantilla = False
    Sql = ""
    
    'Conseguir el siguiente numero de linea
    
    
    
    Sql = "INSERT INTO " & NomTablaLineas
    Sql = Sql & " (numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre) "
    Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", "
    Sql = Sql & codAlmac & ", " & DBSet(codArtic, "T") & ", "
    NomArtic = DevuelveDesdeBDNew(1, "sartic", "nomartic", "codartic", codArtic, "T")
    Sql = Sql & DBSet(NomArtic, "T") & ", " & ValorNulo & ", " & DBSet(cantidad, "N") & ", "
                   
    Importe = CalcularImporte(CStr(cantidad), Precio, Dto1, Dto2, vParamAplic.TipoDtos)
    Sql = Sql & DBSet(Precio, "N") & ", "
    Sql = Sql & DBSet(Dto1, "N") & ", "
    Sql = Sql & DBSet(Dto2, "N") & ", "
    Sql = Sql & DBSet(Importe, "N") & ", '"
    Sql = Sql & OrigP & "')"
     
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLineaDePlantilla = True
        numlinea = numlinea + 1
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta." & vbCrLf & Err.Description
End Function



Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de Revisiones: slima1
Dim Sql As String
Dim VtaDtoSup   As Boolean
    On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    If DatosOkLinea(VtaDtoSup) Then
        Sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        Sql = Sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & " cantidad = " & DBSet(txtAux(3).Text, "N", "N") & ", precioar = " & DBSet(txtAux(4).Text, "N", "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(6).Text, "N", "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N", "N") & ", "
        Sql = Sql & " importel = " & DBSet(txtAux(8).Text, "N") & ", origpre=" & DBSet(txtAux(5).Text, "T")
        'Agosto 2011
        Sql = Sql & " , esopcion = " & Abs(Me.cboOpcion.ListIndex = 1)
        'septiembre 2011
        'SQL = SQL & " , observa = " & DBSet(Text2(18).Text, "T", "S")
        'Marzo 2021
        Sql = Sql & " , codigiva = " & txtAux(9).Text
        
        Sql = Sql & " WHERE " & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea
    End If

    If Sql <> "" Then
        conn.Execute Sql
        ModificarLinea = True
        
        ' ---- [28/10/2010] (DAVID)
        'Esta linea lleva dto superior al permitido.
        'Lo saco fuera del trans
        If VtaDtoSup Then GrabarLogDtoSuperior "OFE", Text1(0), Text1(1).Text, CStr(Data2.Recordset!numlinea), False
        ' ---- [13/01/2011] (DAVID)
        'Si ha cambiado, si tiene el parametro... todo esta ahi
        TrataCambioPrecioDto

        
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Oferta" & vbCrLf & Err.Description
End Function



Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Sql As String

    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez, 360
 
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B

    PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(2).Caption = "Alm."
                vDataGrid.Columns(2).Width = 700
                vDataGrid.Columns(2).NumberFormat = "000"
                
                vDataGrid.Columns(3).Caption = "Artículo"
                vDataGrid.Columns(3).Width = 2200
                
                vDataGrid.Columns(4).Caption = "Descripción Artículo"
                vDataGrid.Columns(4).Width = 5800
                
'                vDataGrid.Columns(5).Caption = "Ampl. Línea"
'                vDataGrid.Columns(5).Width = 7980
                vDataGrid.Columns(5).visible = False
                
                vDataGrid.Columns(6).Caption = "Cantidad"
                vDataGrid.Columns(6).Width = 1200
                vDataGrid.Columns(6).Alignment = dbgRight
                vDataGrid.Columns(6).NumberFormat = FormatoImporte
                
                vDataGrid.Columns(7).Caption = "Precio"
                vDataGrid.Columns(7).Width = 1450
                vDataGrid.Columns(7).Alignment = dbgRight
                vDataGrid.Columns(7).NumberFormat = FormatoPrecio
                
                vDataGrid.Columns(8).Caption = "OP"
                vDataGrid.Columns(8).Width = 450
                vDataGrid.Columns(8).Alignment = dbgCenter
                
                
                vDataGrid.Columns(9).Caption = "Dto. 1"
                vDataGrid.Columns(9).Width = 800
                vDataGrid.Columns(9).Alignment = dbgRight
                vDataGrid.Columns(9).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(10).Caption = "Dto. 2"
                vDataGrid.Columns(10).Width = 800
                vDataGrid.Columns(10).Alignment = dbgRight
                vDataGrid.Columns(10).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(11).Caption = "Importe"
                vDataGrid.Columns(11).Width = 1500
                vDataGrid.Columns(11).Alignment = dbgRight
                vDataGrid.Columns(11).NumberFormat = FormatoImporte
                
                
                vDataGrid.Columns(12).Caption = "Opcion"
                vDataGrid.Columns(12).Width = 800
                vDataGrid.Columns(12).Alignment = dbgLeft
                
                vDataGrid.Columns(13).visible = False 'CODIGIVA
                
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


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
'        txtAux2.visible = visible
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
            cboOpcion.ListIndex = -1
        Else 'Vamos a modificar
        
            
            For i = 0 To txtAux.Count - 2
                If i < 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 2).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                End If
                txtAux(i).Locked = False
            Next i
            If Me.DataGrid1.Columns(12).Text = "Si" Then
                cboOpcion.ListIndex = 1
            Else
                cboOpcion.ListIndex = -1
            End If
             txtAux(9).Text = DataGrid1.Columns(13).Text
        End If
        
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 30)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cboOpcion.Top = alto
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio  !!!! -2 que el ultimo(9) es el IVA
        For i = 4 To txtAux.Count - 2
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 10
        Next i
        
        
        cboOpcion.Left = txtAux(8).Left + txtAux(8).Width + 30
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
    cboOpcion.visible = visible
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


        Select Case Button.Index
            
        Case 1: mnNuevo_Click  'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar

        Case 5: mnBuscar_Click 'Buscar
        Case 6: BotonVerTodos  'Todos
        Case 8: mnImpOferta_Click  'Todos
        End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
    If Modo <> 2 Then Exit Sub
    
    
    Select Case Button.Index
    Case 1
        mnGenPedido_Click
    Case 2
        mnImpRecordatorio_Click
    Case 3
        mnImpValoracion_Click
    Case 4
            mnImpFactProF_Click
        
        
    Case 5
        
            If vUsu.Nivel > 1 Then Exit Sub
            If Modo <> 2 Then Exit Sub
            If Data1.Recordset.EOF Then Exit Sub
            
            
            'Comprobacion
            
            If Not PuedePasarFacuraFAZ Then Exit Sub
            
            'Lanzaremos la pregunta del banco
            'Vamos a generar un ALBARAN ALZ y desde ahi, una factura FAZ
            CadenaDesdeOtroForm = Text1(1).Text
            frmListado3.Opcion = 50
            frmListado3.Show vbModal
            If Mid(CadenaDesdeOtroForm, 1, 3) = "OK#" Then
                'OK. Ha seleccionado el banco
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 4)
                Screen.MousePointer = vbHourglass
                lblIndicador.Caption = "Pasando ALZ"
                lblIndicador.Refresh
                
                'Cambiamos a conta de B
                AbrirConexionConta True
                
                PasarOfertaFacturaFAZ
                
                'Reestablecemos CONTA normal
                lblIndicador.Caption = ""
                Screen.MousePointer = vbDefault
            End If
            
    End Select
    
End Sub

Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
        
        
        If Modo <> 2 Then Exit Sub
        
        If vParamAplic.QueEmpresaEs = 1 Then
            'WHOSE. NO permitimos ni
            If Button.Index = 11 Then Exit Sub
        End If
    



    If Index = 1 Then
        '----------------------------------------------------------------------
        'Documentos
        
        HacerDocumentosPDF Button.Index - 1
    
    Else
        '----------------------------------------------------------------------
        'Lineas oferta
        Select Case Button.Index
        Case 3
            'Eliminar linea
            ModificaLineas = 0
            BotonEliminarLinea
        Case 6, 7
            If Button.Index = 6 Then
                'Plantilla
                mnPlantillas_Click
                
            Else
                'traer lineas
                mnOferta_Click
                
            End If
            
        Case 1, 2, 5
            
            PonerModo 5
            If Modo = 5 Then
                If Button.Index = 2 Then
                    'modificar linea
                    ModificaLineas = 2
                    BotonModificarLinea
                Else
                    '1 y 5
                    'Insertar e Insertar intercalando
                    ModificaLineas = 1
                    BotonAnyadirLinea Button.Index = 5
                End If
            End If
        End Select
            
    End If
        
        
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento Button.Index - 1
End Sub

Private Sub TxtAux_Change(Index As Integer)
    'Precio y Modo Borrar Lineas
    If Index = 4 And ModificaLineas = 2 Then
        txtAux(5).Text = "M"
        BloquearTxt txtAux(6), False
        BloquearTxt txtAux(7), False
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer
    
    

    
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
'    ConseguirFoco txtAux(Index), Modo, cadkey
    ConseguirFocoLin txtAux(Index), cadkey
    
    LabelAyudatxtAux Index, lblF
    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    
    
    '   F2   F2   F2    F2  F2
    If KeyCode = 113 Then
        
        If Index = 3 Then
            PulsaF2 = True
            AbrirForm_Articulos txtAux(1).Text
            
        ElseIf Index = 4 Then
            If vUsu.CodigoAgente = 0 Then
                'Los usuarios/agente no pueden ver esto
                PulsaF2 = True
                AbrirConsultaPrecio2 Text1(4).Text, txtAux(1).Text, Text1(1).Text, Text1(13)
            End If
            
        Else
            If Index = 6 Or Index = 7 Then
                PulsaF2 = True
                AbrirFormularioDtos txtAux(1).Text
            End If
        End If
        
        
    Else
          If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
            If Index < 2 Or Index = 9 Then  'Para los que tienen busqueda
                If Modo = 5 And ModificaLineas = 1 Then
                    If txtAux(Index).Text = "" Then
                        PulsadoMas2 = True
                        KeyCode = 0
                
                        PulsarTeclaMas False, Index
                    End If
                End If
             End If
           End If
    End If
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
'Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As CStock
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim cantidad As String
Dim B As Boolean

Dim StatusArticMayorCero As Boolean

 
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If
    
    If PulsaF2 Then
        PulsaF2 = False
        Exit Sub
    End If

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    'txtAux(Index).Text = Trim(txtAux(Index))
    If txtAux(Index).Text = "" And (Index <> 1 And Index <> 4 And Index <> 3) Then Exit Sub
    
    
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)
            
        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , , StatusArticMayorCero) Then
                
                If devuelve <> txtAux(1).Text Then
                    'ha cambiado el articulo
                    Me.txtAux(3).Text = ""
                    Me.txtAux(4).Text = ""
                    Me.txtAux(5).Text = ""
                    Me.txtAux(6).Text = ""
                    Me.txtAux(7).Text = ""
                End If
                B = (Me.ActiveControl.Name = "txtAux")
                If B Then B = (Me.ActiveControl.Index = 2)
                
                If B Then
                    If txtAux(2).Locked Then
                        If StatusArticMayorCero Then PonerFoco txtAux(3)
                    Else
                        PonerFoco txtAux(2)
                    End If
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
'
'             'Si es articulo de varios podemos modificar la descripción del articulo, sino bloqueamos.
'            If Not EsArticuloVarios(txtAux(Index).Text) Then
'                BloquearTxt txtAux(2), True
'                PonerFoco txtAux(3)
'            Else
'                BloquearTxt txtAux(2), False
'                PonerFoco txtAux(2)
'            End If
        
        Case 2 'desc articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            If txtAux(1).Text = vParamAplic.ArtSeparador Then
                'Articulo separador, pongo CEROS cantidad....dto2 e importe, la opcion la pongo a NO
                'y paso el foco a ampliacion
                ValoresRestoCamposSeparador False  'que los ponga vacios
                PonerFoco Text2(16)
            End If
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                B = True
                Set vCStock = New CStock
                If Not InicializarCStock2(vCStock, "S", 0, False) Then B = False
                If vCStock.MueveStock Then
                    If Not vCStock.MoverStock(False, False) Then B = False
                End If
                If Not B Then
                    Set vCStock = Nothing
                    Exit Sub
                End If
                
                B = False
                If Modo = 5 Then 'Modo lineas
                    If ModificaLineas = 1 Then 'insertar linea
                        B = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then B = True
                    End If
                End If
                
                If B Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If devuelve <> "" Then
                        cantidad = txtAux(Index).Text
                        
                        'Mayo 2009
                        'Si este parametro esta a FALSE, siempre cojera precio ud
                        If vParamAplic.CajasCompletas Then
                            NumCajas = CPrecioFact.ObtenerNumCajas(cantidad, devuelve)
                            RestoUnid = CInt(ComprobarCero(cantidad)) - NumCajas * CInt(devuelve)
                        Else
                            NumCajas = 0
                            If Val(devuelve) > 1 Then
                                If CCur(txtAux(3).Text) >= CCur(devuelve) Then NumCajas = 1
                            End If
                            RestoUnid = 0
                        End If
                    
                    
                 ''       'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                 ''       'precio de caja, y otra linea con el resto unidades un precio unidad
                 ''
                 ''       NumCajas = ObtenerNumCajas(Cantidad, devuelve)
                 ''       RestoUnid = CInt(Cantidad) - NumCajas * CInt(devuelve)
            
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        Set CPrecioFact = New CPreciosFact
                        'CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        CPrecioFact.CodigoClien = Text1(4).Text
                        CPrecioFact.FijarTarifaActividad
                        PorCaja = (NumCajas > 0)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, "")
                        
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Artículo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
    '                        TxtAux(3).Text = NumCajas * CInt(devuelve)
                            PonerFoco txtAux(Index)
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And B) Then
                                txtAux(4).Text = Precio
                                txtAux(5).Text = OrigP 'De donde viene el precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                            PonerFoco txtAux(4)
                            ConseguirFocoLin txtAux(4)
    '                            PonerFoco Text2(16)
    
                            'Si tiene dto permitido
                            If Not CPrecioFact.DtoPermitido Then
                                txtAux(6).Text = "0"
                                txtAux(7).Text = "0"
                                txtAux(6).Enabled = False
                                txtAux(7).Enabled = False
                            End If
    
    
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If 'modo 5
                Set vCStock = Nothing
            End If 'formato decimal
            
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then
                        txtAux(5).Text = "M"
                        BloquearTxt txtAux(6), False
                        BloquearTxt txtAux(7), False
                        PonerFoco txtAux(6)
                    End If
                End If
            End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
        
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        If vParamAplic.ArtReciclado <> "" Then
            ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
        Else
            ClienteConTasaReciclado = False
        End If



        LineaIntercalar = 0
        Me.SSTab1.Tab = numTab
        TituloLinea = Cad
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    Sql = " WHERE  numofert=" & Data1.Recordset!NumOfert

    'Lineas de Ofertas
    conn.Execute "Delete from " & NomTablaLineas & Sql
    
    'Cabecera
    conn.Execute "Delete from " & NombreTabla & Sql

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


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
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


Private Function ObtenerWhereCP() As String
On Error Resume Next

    ObtenerWhereCP = " numofert= " & Text1(0).Text
End Function


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
    
    Sql = "SELECT numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2,importel "
    'Agosto 2011
    Sql = Sql & ",if( EsOpcion=1,""Si"","""")"
    'Marzo 2021
    Sql = Sql & ", codigiva"
    Sql = Sql & " FROM " & NomTablaLineas
    If enlaza Then
        Sql = Sql & " WHERE " & ObtenerWhereCP
        'FALTA###
        'SQL = SQL & " and fecofert='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        Sql = Sql & " WHERE false"
    End If
    Sql = Sql & " Order by numofert, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean, bol As Boolean
Dim i As Byte
      

        B = (Modo = 2 Or Modo = 0)
        Toolbar1.Buttons(5).Enabled = B
        Toolbar1.Buttons(6).Enabled = B
        

        'Insertar
        Toolbar1.Buttons(1).Enabled = B
        
        Toolbar1.Buttons(8).Enabled = True   'IMprimir
        
        B = (Modo = 2)
        Toolbar1.Buttons(2).Enabled = B  'modificar
        Toolbar1.Buttons(3).Enabled = B     'eliminar
        
        
        For i = 1 To 5
            Toolbar2.Buttons(i).Enabled = B
        Next
    
    
End Sub



Private Sub CargarComboFacturacion()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Factura Colectiva, 1-Factura x Albaran

    cboFacturacion.Clear
    cboFacturacion.AddItem "Factura Colectiva"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0

    cboFacturacion.AddItem "Factura x Albaran"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1

End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Ofertas
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numofert", "numofert", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Ofertas (scapre)."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    EsDeVarios = EsClienteVarios(Text1(4).Text)
    If EsDeVarios Then
'        MenError = "Error al actualizar el Cliente de Varios (sclvar)."
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador de la Oferta."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Oferta." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim i As Byte

    For i = 4 To 13
        Text1(i).Text = ""
    Next i
    If Modo = 3 Then
        For i = 14 To 17
            Text1(i).Text = ""
        Next i
        Text2(12).Text = ""
        Text2(14).Text = ""
        Text2(17).Text = ""
        Text1(33).Text = ""
        Text2(33).Text = ""
        Me.cboFacturacion.ListIndex = -1
    End If
End Sub
    

Private Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Integer
Dim NumCajas As Integer
Dim cantidad As Integer, UniPorCaja As Integer

    On Error Resume Next

    cantidad = CInt(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function


Private Function DescargarDatosTMP()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim Sql As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    Sql = "DELETE from tmpscapla" & " where codusu= " & vUsu.Codigo
    conn.Execute Sql
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal", Err.Description
End Function



Private Function InsertarPedido(cadSQL As String, MenError As String, numPed As String, QueLineasPasanPedido As String, Articulos_NO_Rotacion As String) As Boolean
'Devuelve el mensane de error si se produce
'OUT -> numPed: Nº Pedido que inserta
'Dim cadError As String
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String
Dim vSQL As String

    On Error GoTo EInsertarPedido
    
    bol = False
    InsertarPedido = bol
    
    'Obtener el Contador de PEDIDO
    codtipom = "PEV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            numPed = vTipoMov.ConseguirContador(codtipom)
            devuelve = DevuelveDesdeBDNew(1, "scaped", "numpedcl", "numpedcl", numPed, "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codtipom)
                numPed = vTipoMov.ConseguirContador(codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    

    'Acabar la sql con el contador seleccionado
    vSQL = "INSERT INTO scaped (numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec, nomdirec, referenc,codtraba,codagent, codforpa, dtoppago, dtognral, tipofact,"
    vSQL = vSQL & "observa01, observa02, observa03, observa04, observa05,servcomp,restoped,numofert,fecofert,mailconfir,observacrm,coddiren)"
    vSQL = vSQL & " SELECT " & numPed & " as numpedcl, " & cadSQL
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numofert=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (scaped )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Pedido (sliped)."
    If Not InsertarLineasPedido(numPed, QueLineasPasanPedido, Articulos_NO_Rotacion) Then Exit Function
    
    MenError = "Error al actualizar el contador del Pedido."
'    bol = vTipoMov.IncrementarContador("REG")
    bol = vTipoMov.IncrementarContador(codtipom)
    Set vTipoMov = Nothing
    'bol = True
    
EInsertarPedido:
        If Err.Number <> 0 Then bol = False
        InsertarPedido = bol
End Function

'QueLineaspasamos
'       Si esta vacio son todas. Si no indicara cuales son
Private Sub PasarOfertaAPedido(vSQL As String, ByRef CambiosArt As Collection, QueLineaspasamos As String)
Dim bol As Boolean
Dim MenError As String
Dim numPed As String
Dim LineasNoRotacion As String
Dim TextoGestionOferta As String

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    'Insertar en tablas de Pedido la Oferta
    LineasNoRotacion = ""
    TextoGestionOferta = Trim(Text1(32).Text)   'seguiofe
    bol = InsertarPedido(vSQL, MenError, numPed, QueLineaspasamos, LineasNoRotacion)
    If bol Then 'Si se inserta Pedido
       'Pasar la Oferta al Historico de Oferta y Borrarla de Ofertas
       vSQL = " scapre.numofert= " & Text1(0).Text
       
       
      
        
        Debug.Assert False
                
        'If EliminarOferta Then
        '    bol = ActualizarElTraspaso(MenError, vSQL, "OFE", "0")
        'Else
            'numpedcl  fecpedcl   codtrabped
            MenError = "Actualizar oferta"
            vSQL = "UPDATE   scapre SET  numpedcl = " & numPed
            vSQL = vSQL & ",aceptado =1, fecpedcl =" & DBSet(RecuperaValor(BuscaChekc, 2), "F")
            vSQL = vSQL & ", codtrabped =" & RecuperaValor(BuscaChekc, 1)
            vSQL = vSQL & " WHERE numofert =" & Text1(0).Text
            conn.Execute vSQL
        'End If
        
    Else
        MsgBox MenError, vbExclamation
    End If
    
EGenPedido:
    If Err.Number <> 0 Then
        MenError = "Pasando Oferta a Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        Screen.MousePointer = vbDefault
        
        
        
        If LineasNoRotacion = "" Then
            MenError = "La Oferta de Venta Nº: " & Text1(0).Text & vbCrLf & vbCrLf & "ha generado el Pedido Nº: " & Format(numPed, "0000000")
        
        
        Else
            MenError = "Se ha generado el pedido Nº: " & Format(numPed, "0000000") & vbCrLf
            MenError = MenError & "Artículos de NO rotación:" & vbCrLf & String(30, "=") & vbCrLf & LineasNoRotacion
            
                     
            If vParamAplic.NumeroInstalacion = vbHerbelca Then
                MenError = MenError & "¿Enviar email confirmación?"
                If MsgBox(MenError, vbQuestion + vbYesNoCancel) = vbYes Then
                    'Imprime una confirmacion entrega Pedido
                    frmListadoOfer.NumCod = numPed  'Nº de Pedido
                    frmListadoOfer.FecEntre = Text1(2).Text  'fecha del pedido
                
                    AbrirListadoOfer (238) '38: Informe confirmacion entrega de Pedidos

                End If
                MenError = ""
            End If
            
            
        End If
        
        
        If MenError <> "" Then MsgBox MenError, vbInformation
        MenError = ""
        
        'seguiofe
        If TextoGestionOferta <> "" Then MsgBox TextoGestionOferta, vbInformation
        
        
        
        
        If vParamAplic.LogCambioPrecDto Then GrabaLog CambiosArt
        
        'Si tiene
    Else
        conn.RollbackTrans
    End If
    BuscaChekc = ""
End Sub


Private Function InsertarLineasPedido(NumPedido As String, QueLineasPasanAlPedido As String, vArticulos_NO_Rotacion As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Long
Dim N As Long
    On Error GoTo Elin

    'David SEPT2009
    'Falta bultos i bultos servidos.  Con lo cual
    'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
    Sql = ""
    Sql = "SELECT " & NumPedido & " as numpedcl, numlinea, codalmac, lp.codartic, lp.nomartic, lp.ampliaci, "
    Sql = Sql & "cantidad, " & "0 as servidas,"
    'HERBELCA.
    'Numbultos servira para saber las servidas
    If vParamAplic.AlmacenB < 90 Then  'para todos menos para herbelca
        Sql = Sql & "cantidad"  'bultos
    Else
        Sql = Sql & "0"  'preparadas 'HERBELCA
    End If
    '            bultos servidos
    Sql = Sql & ",0, lp.precioar, dtoline1, dtoline2, lp.importel, origpre , NULL as numlote," 'Null de numlote
    ' ---- [21/10/2009] [LAURA] : se añade el centro de coste a pedidos
    Sql = Sql & "NULL as codccost" 'centro de coste
    
     'SAIL
    Sql = Sql & ",NULL,NULL" 'centro de coste
    
    'ENERO 2019
    'Solicitadas
    Sql = Sql & ",cantidad , 0 idL "
    'Precio coste
    Sql = Sql & ", if(artvario=1,0, coalesce(preciouc,0)) precoste"
    
    Sql = Sql & " FROM " & NomTablaLineas & " as lp ,sartic WHERE  numofert=" & Text1(0).Text
    Sql = Sql & " AND lp.codartic=sartic.codartic "
    If vParamAplic.ArtSeparador <> "" Then Sql = Sql & " AND lp.codartic <> " & DBSet(vParamAplic.ArtSeparador, "T")
    'Si llevaba opciones
    If QueLineasPasanAlPedido <> "" Then Sql = Sql & " AND numlinea IN (" & QueLineasPasanAlPedido & ")"
    
    
    '11 Enero 09
    'Ordenado por numlinea
    
    Sql = "INSERT INTO sliped " & Sql
    conn.Execute Sql
    

    'Ahora actualizo los bultos
    'Marzo 2014.
    'HERBELCA.
    'Numbultos servira para saber las servidas
    Set Rs = New ADODB.Recordset
    If vParamAplic.NumeroInstalacion <> 2 Then 'para todos menos para herbelca
        Sql = "Select cantidad , unicajas,numlinea from sliped,sartic where sliped.codartic = sartic.codartic and unicajas >1 and  sliped.numpedcl = " & NumPedido
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            i = Rs!cantidad \ CLng(Rs!unicajas)
            If (Rs!cantidad Mod CLng(Rs!unicajas)) > 0 Then i = i + 1
            Sql = "UPDATE sliped Set numbultos=" & i & " WHERE sliped.numpedcl = " & NumPedido & " AND numlinea = " & Rs!numlinea
            conn.Execute Sql
            Rs.MoveNext
        Wend
        Rs.Close
        
    Else
        
        Sql = ""
        If vParamAplic.ArtReciclado <> "" Then Sql = Sql & ", " & DBSet(vParamAplic.ArtReciclado, "T")
        If vParamAplic.ArtSeparador <> "" Then Sql = Sql & ", " & DBSet(vParamAplic.ArtSeparador, "T")
        If Sql <> "" Then
           Sql = Mid(Sql, 2)
           Sql = " AND NOT sliped.codartic IN (" & Sql & ")"
        End If
        
        Sql = "Select sliped.codartic,sliped.nomartic from sliped,sartic where sliped.codartic = sartic.codartic  and  sliped.numpedcl = " & NumPedido & Sql
        Sql = Sql & " AND rotacion=0"
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
           vArticulos_NO_Rotacion = vArticulos_NO_Rotacion & Rs!codArtic & " - " & Rs!NomArtic & vbCrLf
            
           Rs.MoveNext
        Wend
        Rs.Close
        
    End If
    
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then 'para todos menos para herbelca
        Sql = "Select numpedcl,numlinea from sliped where sliped.numpedcl = " & NumPedido
        Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
        i = 0
        While Not Rs.EOF
            i = i + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        
        
        Sql = DevuelveDesdeBD(conAri, "contador", "stipom", "codtipom", "LPD", "T")
         N = Val(Sql) + 1
        Sql = "UPDATE stipom SET contador=" & N + i & " WHERE codtipom = 'LPD'"
        conn.Execute Sql
        
        i = N
        While Not Rs.EOF
            
            
            Sql = "UPDATE sliped Set idl=" & i & " WHERE sliped.numpedcl = " & NumPedido & " AND numlinea = " & Rs!numlinea
            conn.Execute Sql
            i = i + 1
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Set Rs = Nothing
    
    
    InsertarLineasPedido = True
    Exit Function
Elin:
    Set Rs = Nothing
     'Hay error , almacenamos y salimos
    InsertarLineasPedido = False

        
End Function


Private Function InicializarCStock2(ByRef vCStock As CStock, TipoM As String, numlinea As String, ForzarDetaMov As Boolean) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM
    If ForzarDetaMov Then
        vCStock.DetaMov = "ALV"
    Else
        vCStock.DetaMov = CodTipoMov
    End If
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente de la oferta
    vCStock.Documento = Text1(0).Text 'Nº de oferta
    vCStock.FechaMov = Text1(1).Text 'Fecha oferta
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then '1=Insertar Linea
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
    End If
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock2 = False
    Else
        InicializarCStock2 = True
    End If
End Function



Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim limpiar As Boolean

    On Error GoTo EPonerDatos
    
    If codClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(codClien) Then
        If vCliente.LeerDatos(codClien) Then
            'si el cliente esta bloqueado salimos
            limpiar = vCliente.ClienteBloqueado(1, False)
            If Not limpiar Then
                'Si va por agente
                If vUsu.CodigoAgente > 0 Then
                    limpiar = vCliente.Agente <> vUsu.CodigoAgente
                    If limpiar Then MsgBox "Cliente incorrecto", vbExclamation
                End If
            End If
            If limpiar Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(vCliente.Codigo, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
                
                vCliente.PonDatosDireccionEnvio Text1(33), Text2(33)
                
            End If
            
            If Modo = 3 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
                
                
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
            
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(1).Text
            
  
      
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub



Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim B As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    B = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
'    Text1(6).Text = vCliente.NIF
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not B Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente

    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
        vCliente.NIF = NIF
        vCliente.Nombre = Text1(5).Text
        vCliente.Domicilio = Text1(8).Text
        vCliente.CPostal = Text1(9).Text
        vCliente.Poblacion = Text1(10).Text
        vCliente.Provincia = Text1(11).Text
        vCliente.TfnoClien = Text1(7).Text
        vCliente.ActualizarClienteV (NIF)
    End If
    Set vCliente = Nothing
    
    ActualizarClienteVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarClienteVarios = False
    Else
        ActualizarClienteVarios = True
    End If
End Function


Private Sub BotonImprimirProForma(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vTipoM As CTiposMov

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Oferta para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 17 'Facturas Proforma Clientes
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If
    
    'Pasar la letra serie de la factura como parámetro
    Set vTipoM = New CTiposMov
    If vTipoM.Leer("FAV") Then
        
    End If
    Set vTipoM = Nothing
    
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Oferta
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº Oferta
        devuelve = "{" & NombreTabla & ".numofert}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
    End If
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
    
        'Tb lo vamos a tratar
        .outTipoDocumento = 5
        .outClaveNombreArchiv = Text1(0).Text
        .outCodigoCliProv = CLng(Val(Text1(4).Text))
        
    
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = "Factura ProForma"
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub CalcularDatosFactura()
Dim T
Dim cadWhere As String
Dim Sql As String
Dim vFactu As CFactura

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For Each T In Text3
        T.Text = ""
    Next
    
    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP
    
    'Agosto 2011. Solo sumaran aquellas que en esopcion NO sea 1
    cadWhere = cadWhere & " AND esopcion <> 1"
    If Me.Data2.Recordset.RecordCount < 1 Then Exit Sub
    
    'SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    'If RegistrosAListar(SQL) = 0 Then Exit Sub
    Sql = DevuelveDesdeBD(conAri, "sum(importel)", NomTablaLineas, "esopcion=1 AND numofert", Text1(0).Text)
    Text3(1).Text = Sql
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, False) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(55).Text = vFactu.TotalFac
        Text3(56).Text = vFactu.BaseImp
        
        
        'Recargos de equivalencia
        Text3(49).Text = vFactu.PorceIVA1RE
        Text3(50).Text = vFactu.PorceIVA2RE
        Text3(51).Text = vFactu.PorceIVA3RE
        Text3(52).Text = vFactu.ImpIVA1RE
        Text3(53).Text = vFactu.ImpIVA2RE
        Text3(54).Text = vFactu.ImpIVA3RE
        
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    
    
    Set vFactu = Nothing
End Sub



Private Function FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    For i = 49 To 54
        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
     'TOTALES
    Text3(55).Text = Format(Text3(55).Text, FormatoImporte)
    Text3(56).Text = Format(Text3(56).Text, FormatoImporte)
    Text3(1).Text = Format(Text3(1).Text, FormatoImporte)  'las opciones
End Function




Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text2(12).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function


Private Sub ComprobarRefObligatoria()
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    If vClien.TieneRefObligatoria(Text1(13).Text) Then
        If Text1(13).Text = "" Then PonerFoco Text1(13)
    End If
    Set vClien = Nothing
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                'El Data esta vacio, desde el modo de inicio se pulsa Insertar
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 1, "Oferta"
                BotonAnyadirLinea False
            End If
        End If
        FormateaCampo Text1(0)
    End If
    Set vTipoMov = Nothing
End Sub






'Para obtener los dtos por cantidad lo que hace es a partir de un
'subtring del articulo(poscion 3 a 6) va a sdesca con la suma de la cantidad
'si en sdesca y dentro de los desde /hasta cantidad encuentra un dto lo aplica


Private Sub DescuentosCantidad(Articulo As String)
Dim Cad As String
Dim R As ADODB.Recordset
Dim NuevoDto As Currency
Dim Importe As Currency
Dim bAct As Boolean
Dim Refrescar As Boolean
    On Error GoTo EDescuentosCantidad
    
    If Not vParamAplic.DtoxCantidad Then Exit Sub ' ---- [14/09/2009] (LAURA)
     
    If MsgBox("¿Desea recalcular los descuentos por cantidad?", vbQuestion + vbYesNo) = vbYes Then    'masl 140909
    
        'Si no  tenemos portes, ni nos pasamos
        If vParamAplic.TipoPortes <> 1 Then Exit Sub
        
        Refrescar = False
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        Cad = " WHERE numofert = " & Text1(0).Text
        'cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from " & NomTablaLineas & cad
        'If Articulo <> "" Then cad = cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        
        Cad = "select codartic ,sum(cantidad) suma from " & NomTablaLineas & Cad
        If Articulo <> "" Then Cad = Cad & " AND codartic= '" & Articulo & "'"
        
        
        
        'Y origen PRECIO no es precio especial
        Cad = Cad & " AND origpre <> 'E'"
        Cad = Cad & " group by 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                Cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                Cad = "select * from sdesca where desdecan <=" & Cad & " and " & Cad & " <= hastacan and envagran = '"
                Cad = Cad & miRsAux!codArtic & "'"
                R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Cad = ""
                If Not R.EOF Then Cad = R!dtolinea
                R.Close
                
                
                If Cad <> "" Then
                    'OK tiene nuevo descuento
                    NuevoDto = CCur(Cad)
                    
                    'Cojo los articulos del albaran y le meto el dto
                    Cad = " WHERE numofert = " & Text1(0).Text
                    Cad = "select * from " & NomTablaLineas & Cad
                    '                                 a partir de la 3era posicion
                    'cad = cad & " AND codartic like '__" & miRsAux!raiz & "%'"
                    Cad = Cad & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
                    R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not R.EOF
                        '-- comprobar si admite descuento
                        If R!origpre = "T" Then
                            Cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            Cad = DevuelveDesdeBDNew(conAri, "slista", "dtopermi", "codartic", R!codArtic, "T", , "codlista", Cad, "N")
                            bAct = (Cad = "1")
                        ElseIf R!origpre = "A" Or R!origpre = "M" Then
                            bAct = True
                        Else
                            bAct = False
                        End If
                        
                        If bAct Then
                            Cad = CalcularImporte(CStr(R!cantidad), CStr(R!precioar), CStr(NuevoDto), CStr(R!dtoline2), vParamAplic.TipoDtos)
                            Importe = CCur(Cad)
                            Cad = "UPDATE " & NomTablaLineas & " set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            Cad = Cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            Cad = Cad & " WHERE numofert = " & Text1(0).Text
                            Cad = Cad & " and numlinea = " & R!numlinea
                            conn.Execute Cad
                            Refrescar = True
                        End If
                        'Siguiente
                        R.MoveNext
                    Wend
                    R.Close
                    
                End If
                'sig
                miRsAux.MoveNext
        Wend
        miRsAux.Close
   End If 'masl 140909
   
   
   If Refrescar Then
        CargaGrid DataGrid1, Data2, True
        CalcularDatosFactura
   End If
   
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub


Private Sub PosicionarData2()
    On Error GoTo EPosicionarData2
    
    Data2.Recordset.Find "numlinea = " & NumRegElim
    If Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub

Private Sub LanzaBusquedaDpto(Departamento As Boolean, Indice As Byte)

    Set frmDptoEnvio = New frmFacCliEnvDpto
    frmDptoEnvio.DireccionesEnvio = Not Departamento
    If Text1(Indice).Text <> "" Then
        frmDptoEnvio.VerDatoDpto = CInt(Text1(Indice).Text)
    Else
        frmDptoEnvio.VerDatoDpto = -1
    End If
    frmDptoEnvio.codClien = CLng(Text1(4).Text)
    frmDptoEnvio.NomClien = Text1(5).Text
    frmDptoEnvio.Show vbModal
    Set frmDptoEnvio = Nothing
End Sub




'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera2 = 0
        If imgBuscar(Index).visible Then imgBuscar_Click Index
        
    Else
        'Lineas
        EsCabecera2 = 1
        cmdAux_Click Index
        
        
    End If
        
End Sub





Private Sub ComprobarPrecioDtoArticulo(ByRef Carts As Collection)
Dim Aux As String
Dim RN As ADODB.Recordset
Dim Cambiado As String
Dim CPrecioFact As CPreciosFact
Dim Importe As Currency
Dim OrigP As String
Dim vAr As CArticulo
Dim PrecioAticuloVenta As Currency
Dim CambiaDto As Boolean
Dim Caj2 As Integer


    On Error GoTo eComprobarPrecioDtoArticulo
    Set Carts = New Collection
    Aux = "Select * from " & NomTablaLineas & " WHERE numofert= " & Text1(0).Text & " ORDER BY numlinea"
    Set RN = New ADODB.Recordset
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set CPrecioFact = New CPreciosFact
    CPrecioFact.CodigoClien = Text1(4).Text  'Cliente
    CPrecioFact.FijarTarifaActividad
    While Not RN.EOF
        'Para cada linea vere que precio tiene
        CambiaDto = False
        If RN!origpre <> "M" Then
            Cambiado = ""
                        
            PorCaja = False
            Aux = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RN!codArtic, "T")
            If Val(Aux) > 1 Then
                
                Caj2 = ObtenerNumCajas(CStr(RN!cantidad), Aux)
                If (Caj2 > 0) Then PorCaja = True
            End If
            CPrecioFact.CodigoArtic = RN!codArtic
            Aux = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, "")
            
            If RN!origpre <> OrigP Then Cambiado = "Orig: " & RN!origpre & "-" & OrigP
            
            Importe = CCur(Aux)
            If Importe <> RN!precioar Then
                Cambiado = Cambiado & "  Prec: " & RN!precioar & "-" & Importe
                CambiaDto = True
            End If
                
            'Dtos
            Importe = CCur(CPrecioFact.Descuento1)
            If RN!dtoline1 > Importe Then
                Cambiado = Cambiado & "  Dto1: " & RN!dtoline1 & "-" & Importe
                CambiaDto = True
            End If
            'dto2
            Importe = CCur(CPrecioFact.Descuento2)
            If RN!dtoline2 > Importe Then
                Cambiado = Cambiado & "  Dto2: " & RN!dtoline2 & "/" & Importe
                CambiaDto = True
            End If
            
            
                    
            
            'Si ha cambiado:
            If Cambiado <> "" Then
                Cambiado = Mid("L" & RN!numlinea & "      ", 1, 6) & " " & Mid(RN!codArtic & "          ", 1, 12) & " ->" & Cambiado
                Carts.Add Cambiado
            End If
        
        End If
        
        
        If vParamAplic.NumeroInstalacion = 2 Then
            
            If RN!cantidad = 0 Then
                PrecioAticuloVenta = 0
            Else
               PrecioAticuloVenta = Round(RN!ImporteL / RN!cantidad, 4)
            End If
                            
            'PorDebajoPrecioMinimo
            Set vAr = New CArticulo
            If vAr.LeerDatos(RN!codArtic) Then
                Aux = ""
                If RN!origpre = "P" Then Aux = "1"
                
                If RN!origpre = "E" Then If Not CambiaDto Then Aux = "1"
                
                
                If vAr.EsDeVarios = 1 Then Aux = "1"
                
                If Aux = "" Then
                    'If Not vAr.EstablecidoPrecioMinimo Then vAr.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
                    vAr.FijarprecioMinimo_ CDate(Text1(1).Text), Val(Text1(4).Text)
                    
                    
                    If vAr.EstablecidoPrecioMinimo Then
                            'If PrecioAticuloVenta < vAr.PrecioMinimo Then
                            If vAr.PrecioMinimo - PrecioAticuloVenta > 0.01 Then
                                PorDebajoPrecioMinimo = True
                                CadenaSQL = CadenaSQL & vbCrLf & "- " & vAr.Codigo & "   " & vAr.Nombre & ": " & RN!ImporteL & " (" & vAr.PrecioMinimo & ")"
                            End If
                    End If
                
                End If
            End If
            
        End If
    
        'siguiente
        RN.MoveNext
    Wend
    RN.Close
    
    
    
     
    
    
eComprobarPrecioDtoArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobando cambio precios para LOG"
    Set RN = Nothing
    Set CPrecioFact = Nothing
    Set vAr = Nothing
End Sub



Private Function HaCambiadoprecioDto(ByRef cart As Collection) As Boolean
Dim Aux As String
Dim i As Integer

    HaCambiadoprecioDto = True
    If cart.Count = 0 Then Exit Function

    'Ha cambiado algo
    '-----------------
    For i = 1 To cart.Count
        Aux = Aux & vbCrLf & cart(i)
    Next
    Aux = vbCrLf & Aux
    Aux = "Existen nuevos precios para algunos articulos de la oferta.        ¿Continuar?" & vbCrLf & String(60, "*") & Aux
    If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then HaCambiadoprecioDto = False


End Function


Private Sub GrabaLog(ByRef CA As Collection)
Dim Aux As String
Dim i  As Integer

    If CA.Count = 0 Then Exit Sub


    'Ahora ire grabando el LOG
    
    Set LOG = New cLOG
    i = 1
    While i <= CA.Count
        If Aux = "" Then Aux = "Oferta: " & Text1(0).Text & " " & Text1(1).Text
        
        If Len(Aux) + Len(CA(i)) > 253 Then
        
            'Añado el LOG tal y como esta. No muevo el I
            LOG.Insertar 15, vUsu, Aux
            Aux = ""
            Espera 1
        Else
            'Meto en la cadena
            Aux = Aux & vbCrLf & CA(i)
            i = i + 1
        End If
    Wend
    LOG.Insertar 15, vUsu, Aux
    
    Set LOG = Nothing

End Sub




Private Sub ComprobarCambioPrecioDto(ByRef El_Articulo As CArticulo)
Dim CPrecioFact As CPreciosFact
Dim Impo As Currency
Dim SQ As String
Dim Particular As Boolean
Dim Cajas As String
Dim PrMinimo As Currency
Dim PrecioArtFinal As Currency
Dim ComprobarPrecioMinimo As Boolean
    On Error GoTo EComprobarCambioPrecioDto
    
        

    'Al modificar puede ser que no haya pasado por codartic
    Cajas = "unicajas"
    SQ = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", txtAux(1).Text, "T", Cajas)
    If SQ = "1" Then Exit Sub


    SQ = DevuelveDesdeBD(conAri, "particular", "sclien", "codclien", Text1(4).Text)
    Particular = SQ = "1"
    
        
       
    'El resto
    ComprobarPrecioMinimo = True
    If ModificaLineas = 1 Then
        'ESTAMOS INSERTANDO
        
            If Particular Then
                SQ = DevuelveDesdeBD(conAri, "maxdtopar", "sfamia,sartic", "sartic.codfamia=sfamia.codfamia  and codartic", txtAux(1).Text, "T")
                If SQ <> "" Then
                    Impo = ImporteFormateado(txtAux(6).Text)
                    Impo = Impo + ImporteFormateado(txtAux(7).Text)
                    If Impo > CCur(SQ) Then GrabaLogCambioPrecioDto = True
                End If
            
            
            Else
                'Los dtos
                '------------------------------------------
                Set CPrecioFact = New CPreciosFact
                            
                CPrecioFact.CodigoClien = Text1(4).Text
                
                'Obtenemos la Tarifa del Cliente
                'AHORA ESTA DENTRO DE LA CLASE
                'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                'CPrecioFact.CodigoLista = codTarif
                CPrecioFact.FijarTarifaActividad
                CPrecioFact.CodigoArtic = txtAux(1).Text
                
                If Val(Cajas) > 1 Then
                    Impo = Val(CCur(txtAux(3).Text)) - Val(Cajas)
                    If Impo >= 0 Then Cajas = ""
                End If
                
                
                SQ = CPrecioFact.ObtenerPrecio(Cajas = "", Text1(1).Text, "", "")
                'SQ = CPrecioFact.ObtenerPrecioDtoFamilia(Cajas = "", Text1(1).Text, "")
                SQ = CalcularImporte(txtAux(3).Text, SQ, CPrecioFact.Descuento1, CPrecioFact.Descuento2, vParamAplic.TipoDtos)
                    
    
                Impo = ImporteFormateado(txtAux(6).Text)
                If Impo <> CCur(CPrecioFact.Descuento1) Then
                    GrabaLogCambioPrecioDto = True
                Else
                    Impo = ImporteFormateado(txtAux(7).Text)
                    If Impo <> CCur(CPrecioFact.Descuento2) Then GrabaLogCambioPrecioDto = True
                End If
    
    
                Set CPrecioFact = Nothing
            End If

    Else
    
        'MODIFICANDO
        'Si ha cambiado el precio,dto1 o dto
        Impo = ImporteFormateado(txtAux(4).Text)
        If Impo <> CCur(Data2.Recordset!precioar) Then
            GrabaLogCambioPrecioDto = True
        Else
            Impo = ImporteFormateado(txtAux(6).Text)
            If Impo <> CCur(Data2.Recordset!dtoline1) Then
                GrabaLogCambioPrecioDto = True
            Else
                Impo = ImporteFormateado(txtAux(7).Text)
                If Impo <> CCur(Data2.Recordset!dtoline2) Then GrabaLogCambioPrecioDto = True
            End If
        End If
        ComprobarPrecioMinimo = GrabaLogCambioPrecioDto
    End If  'Modificar-añadir
    

    'OCtubre2018     En noviembre  YA NO COMBROBAMOS EN LA EDCION. SOLO en el pase
    'If vParamAplic.NumeroInstalacion = 2 And ComprobarPrecioMinimo Then
    If False Then
        'En herbelca, si ha cambiado el precio, tenemos que comprobar si es menor que el precio minimo
        El_Articulo.LeerDatos El_Articulo.Codigo
        SQ = ""
        If txtAux(5).Text = "P" Then SQ = "1"
        If El_Articulo.EsDeVarios = 1 Then SQ = "1"
        
        If SQ = "" Then
            PrecioArtFinal = 0
            If CCur(txtAux(3).Text) <> 0 Then PrecioArtFinal = Round2(CCur(txtAux(8).Text) / CCur(txtAux(3).Text), 4)
       
            
            El_Articulo.FijarprecioMinimo_ CDate(Text1(1).Text), Val(Text1(4).Text)
            
        
            If El_Articulo.EstablecidoPrecioMinimo Then
                If PrecioArtFinal < El_Articulo.PrecioMinimo Then
                    If Abs(PrecioArtFinal - El_Articulo.PrecioMinimo) > 0.009 Then PorDebajoPrecioMinimo = True
                End If
            End If
        
        
            Set CPrecioFact = Nothing
        End If  'de varios
    End If

    Exit Sub
EComprobarCambioPrecioDto:
    MuestraError Err.Number, "Comprobando cambio precio descuento.  El programa CONTINUARA"
End Sub


Private Sub TrataCambioPrecioDto()
Dim Rc

    If Not GrabaLogCambioPrecioDto Then Exit Sub
    Rc = Screen.MousePointer
    frmListado3.Opcion = 0
    If ModificaLineas = 1 Then
        frmListado3.OtrosDatos = "N."
    Else
        frmListado3.OtrosDatos = "M."
    End If
    frmListado3.OtrosDatos = frmListado3.OtrosDatos & " OFe " & Text1(0).Text & " " & Text1(1).Text & " Articulo " & txtAux(1).Text
    
    
    frmListado3.Show vbModal
    
    
    Screen.MousePointer = Rc
    
    
End Sub



Private Function TieneOpciones(MostrarMensaje As Boolean) As Boolean
Dim Sql As String
    TieneOpciones = False
    
    'Si tiene opciones NO depodremos listar algunas cosas: fra provorma
    If Text1(0).Text <> "" Then
        Sql = DevuelveDesdeBD(conAri, "count(*)", NomTablaLineas, "esopcion=1 AND numofert", Text1(0).Text)
        If Val(Sql) > 0 Then
            If MostrarMensaje Then MsgBox "Oferta con opciones", vbExclamation
            TieneOpciones = True
        Else
            'Va a pasar de oferta a pedidod
            If Not MostrarMensaje Then
                
                    If Precio = "" Then
                        Precio = "NO"
                        If MsgBox("¿Pasar todas las lineas de la oferta al pedido?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            TieneOpciones = True
                            Precio = "SI"
                        End If
                    Else
                        If Precio = "SI" Then TieneOpciones = True
                    End If
                
            End If
        End If
    End If
End Function

'Aceros-->pone un cero
'Vacios-->pone ""
Private Sub ValoresRestoCamposSeparador(Aceros As Boolean)
Dim C As String
    'Cuando el articulo SEA separador....
    txtAux(5).Text = ""
    If Aceros Then
        C = "0"
        Me.cboOpcion.ListIndex = 0
        
    Else
        Me.cboOpcion.ListIndex = -1
        C = ""
    End If
    txtAux(3).Text = C: txtAux(4).Text = C:
    txtAux(6).Text = C: txtAux(7).Text = C:
    txtAux(8).Text = C  ': txtAux(8).Text = "0":
    
    
    If Aceros Then txtAux(5).Text = "M"
    
    
End Sub


Private Sub PonerObservacionesPordefecto()

    Set miRsAux = New ADODB.Recordset
    txtAnterior = " Select plazos01,plazos02,plazos03,asunto01,asunto02,asunto03,asunto04,asunto05,"
    txtAnterior = txtAnterior & "observa01,observa02,observa03,observa04,observa05 FROM spara2"
    miRsAux.Open txtAnterior, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        For NumRegElim = 0 To miRsAux.Fields.Count - 1
            'Van seguidos desde el 18
            Text1(NumRegElim + 18).Text = DBLet(miRsAux.Fields(NumRegElim), "T")
        Next
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    txtAnterior = ""
End Sub



Private Function PuedePasarFacuraFAZ() As Boolean
Dim vCStock As CStock


    On Error GoTo ePuedePasarFacuraFAZ
    
    Screen.MousePointer = vbHourglass
    PuedePasarFacuraFAZ = False
    Set miRsAux = New ADODB.Recordset
    
    'MAYO
    'QUitamos esta primera comprobacion pq edire fecha factura
    
    'Primera comprobacion.
    'Las fechas dentro de ejercicio actual y siguiente
    'If CDate(Text1(2).Text) < vEmpresa.FechaIni Or CDate(Text1(2).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then Err.Raise 513, , "Fechas fuera de ejercicio"
    
    
    'Segunda comprobacion.
    'QUe no exista la factura numofer con fecha fecofer en FAZ
    CadenaSQL = "Select nomclien from scafac where codtipom='FAZ' and numfactu=" & Me.Text1(0).Text
    CadenaSQL = CadenaSQL & " AND fecfactu =" & DBSet(Text1(2).Text, "F")
    miRsAux.Open CadenaSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaSQL = ""
    If Not miRsAux.EOF Then CadenaSQL = miRsAux!NomClien
    miRsAux.Close
    
    If CadenaSQL <> "" Then Err.Raise 513, , "Ya existe la factura: " & CadenaSQL
    
    
    'Que no exista el albaran
    CadenaSQL = "Select nomclien from scaalb where codtipom='ALZ' and numalbar=" & Me.Text1(0).Text
    CadenaSQL = CadenaSQL & " AND fechaalb =" & DBSet(Text1(2).Text, "F")
    miRsAux.Open CadenaSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaSQL = ""
    If Not miRsAux.EOF Then CadenaSQL = miRsAux!NomClien
    miRsAux.Close
    
    If CadenaSQL <> "" Then Err.Raise 513, , "Ya existe el albaran: " & CadenaSQL
        
    
    
    
    'Vemos stocks
    Set vCStock = New CStock
    NumRegElim = 0
    Data2.Recordset.MoveFirst
    CadenaSQL = ""
    ModificaLineas = 0 'para que lea de la BD
    While Not Data2.Recordset.EOF
        If Data2.Recordset.Fields(12) = "" Then   'ES OPCION   O el valor es "" o el valor es Si
            If Data2.Recordset!codArtic <> vParamAplic.ArtSeparador Then
                NumRegElim = 1
                'PARA LA COMPROBACION PONGO COMO SI FUERA UNA ALBARAN
                If InicializarCStock2(vCStock, "S", 0, True) Then
                    If vCStock.MueveStock Then
                        If Not vCStock.MoverStock(False, False, True) Then CadenaSQL = CadenaSQL & "No hay stock: " & Data2.Recordset!codArtic & vbCrLf
                    End If
                Else
                    CadenaSQL = CadenaSQL & "cStock: " & Data2.Recordset!codArtic & vbCrLf
                End If
            End If
        End If
        Data2.Recordset.MoveNext
    Wend
    Data2.Recordset.MoveFirst
    
    
    If NumRegElim = 0 Then CadenaSQL = "No hay lineas para realizar la factura"
    
    
    If CadenaSQL <> "" Then Err.Raise 513, , "Comprobar datos: " & vbCrLf & CadenaSQL
    
    
    'Comprobare que existe la cuenta en la contabilidad B
    CadenaSQL = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", Text1(4).Text)
    If CadenaSQL <> "" Then
        'Voy a la tabla de cuentas
        CadenaSQL = DevuelveDesdeBD(conAri, "nommacta", "conta" & vParamAplic.ContabilidadB & ".cuentas", "codmacta", CadenaSQL, "T")
    End If
    
    
    If CadenaSQL = "" Then
         Err.Raise 513, , "Comprobar datos: " & vbCrLf & "Error en cuenta contable cliente " & Me.Text1(4).Text
    Else
        'TODO OK. Podemos generar la factura
        PuedePasarFacuraFAZ = True
    End If
    
    
    
    
    
ePuedePasarFacuraFAZ:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set vCStock = Nothing
    CadenaSQL = ""
    Screen.MousePointer = vbDefault
End Function





Private Sub PasarOfertaFacturaFAZ()
Dim BancoPr As Integer
Dim FechaFac As Date
    'El proceso se divide en
    '   1.- Crear albaran y los movimientos
    '   2.- pasar albaran a la factura
    '   3.- Eliminar oferta pasandola al HCO con una marca en un campo
    '   4.- Grabar LOG
    '   5.- Situar DATA


   
    BancoPr = Val(RecuperaValor(CadenaDesdeOtroForm, 1))
    FechaFac = RecuperaValor(CadenaDesdeOtroForm, 2)
    
    If Not BloqueaRegistro(NombreTabla, "numofert=" & Text1(0).Text) Then Exit Sub
    
    lblIndicador.Caption = "Inserta  ALZ"
    lblIndicador.Refresh
    conn.BeginTrans
    HaDevueltoDatos = GenerarAlbaransmovalFAZ
    If HaDevueltoDatos Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    
    CadenaDesdeOtroForm = ""
    
    If HaDevueltoDatos Then
        'OK. Ha generado el albaran
        
        
        lblIndicador.Caption = "Pasando factura"
        lblIndicador.Refresh
        Espera 0.8
        
        CadenaConsulta = "SELECT scaalb.* FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien  WHERE scaalb.codtipom = 'ALZ' AND scaalb.numalbar = " & Text1(0).Text
        CadenaSQL = "scaalb.codtipom = 'ALZ' AND scaalb.numalbar = " & Text1(0).Text
        
        HaDevueltoDatos = TraspasoAlbaranesFacturas(CadenaConsulta, CadenaSQL, Format(FechaFac, "dd/mm/yyyy"), CStr(BancoPr), Nothing, lblIndicador, True, "ALZ", "", 1, False, True, False, False)
        CadenaConsulta = ""
        
        CadenaDesdeOtroForm = "OFE: " & Text1(0).Text & vbCrLf & "-Generado albaran    OK"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "-Generacion factura     "
        If Not HaDevueltoDatos Then
            'Error generando factura. Soporte tecnico
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " ERROR"
        Else
            'OK
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " OK"
                
            'Updateamos el campo
            lblIndicador.Caption = "Pasando oferta hco"
            lblIndicador.Refresh
            CadenaSQL = Trim(Text1(32).Text)
            If CadenaSQL <> "" Then CadenaSQL = vbCrLf & vbCrLf & CadenaSQL
            CadenaSQL = "[FAZ]  " & CStr(Now) & " " & vUsu.Nombre & "[#]" & CadenaSQL
            CadenaSQL = "UPDATE scapre set seguiofe =" & DBSet(CadenaSQL, "T") & " WHERE numofert=" & Text1(0).Text
            conn.Execute CadenaSQL
            Espera 0.6
            
            'La traspasamos al hco
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "-Hco:     "
            CadenaConsulta = "scapre.numofert=" & Text1(0).Text
            CadenaSQL = ""
            If ActualizarElTraspaso(CadenaSQL, CadenaConsulta, "OFE") Then
                HaDevueltoDatos = True
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " OK"
                
            Else
                HaDevueltoDatos = False
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " ERROR"
            End If
            
            If HaDevueltoDatos Then
                Espera 0.5
                 NumRegElim = Data1.Recordset.AbsolutePosition
                'Situamos data tras actualizar, ya que se ha eliminado
                If SituarDataTrasEliminar(Data1, NumRegElim) Then
                    PonerCampos
                Else
                    LimpiarCampos
                    'Poner los grid sin apuntar a nada
                    LimpiarDataGrids
                    PonerModo 0
                End If
            End If
            
            
            
        End If
        
        
        Set LOG = New cLOG
        LOG.Insertar 28, vUsu, CadenaDesdeOtroForm
        Set LOG = Nothing
        
    End If
    
End Sub




Private Function GenerarAlbaransmovalFAZ() As Boolean
Dim vCStock As CStock
Dim B As Boolean
Dim vSQL As String

    On Error GoTo eGenerarAlbaransmovalFAZ

    GenerarAlbaransmovalFAZ = False
    
    CadenaDesdeOtroForm = "codenvio"
    CadenaSQL = DevuelveDesdeBD(conAri, "codzonas", "sclien", "codclien", Text1(4).Text, "N", CadenaDesdeOtroForm)
    If CadenaSQL = "" Then
        CadenaSQL = DevuelveDesdeBD(conAri, "codzonas", "szonas", "1", "1")
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "codenvio", "senvio", "1", "1")
    End If
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    vSQL = vSQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    vSQL = vSQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,coddiren,tipAlbaran,codzonas) "
    vSQL = vSQL & " SELECT 'ALZ',numofert,fecofert,1,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    vSQL = vSQL & " coddirec,nomdirec,referenc,codtraba,codtraba,codtraba,codagent,codforpa," & CadenaDesdeOtroForm
    vSQL = vSQL & " ,dtoppago,dtognral,tipofact , observa01, observa02, observa03, observa04, observa05, NumOfert, fecofert, Null, Null, Null,NULL, coddiren, 0," & CadenaSQL
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numofert=" & Text1(0).Text
    conn.Execute vSQL
    
    'Insertar lineas
    vSQL = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos"
    vSQL = vSQL & ",precioar,dtoline1,dtoline2,importel,origpre,codproveX,numlote,codccost,codtipor,codcapit,"
    vSQL = vSQL & "precoste,codtraba,pvpInferior)"
    
    vSQL = vSQL & " SELECT  'ALZ',numofert,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,0 ,"
    vSQL = vSQL & "precioar,dtoline1,dtoline2,importel,origpre,codproveX,NULL,NULL,NULL,codcapit,NULL,NULL,0 "
    CadenaSQL = "numofert=" & Text1(0).Text & " AND esopcion=0 and codartic <> " & DBSet(vParamAplic.ArtSeparador, "T")
    vSQL = vSQL & " FROM slipre WHERE " & CadenaSQL
    
    conn.Execute vSQL

    
    'LOS STOCKS
    vSQL = "Select * FROM slipre WHERE " & CadenaSQL
    Set vCStock = New CStock
    B = True
    
    vSQL = "select * from slipre WHERE numofert = " & Text1(0).Text & " AND " & CadenaSQL
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'NO PUEDE SER EOF
    vCStock.tipoMov = "S"
    vCStock.DetaMov = "ALZ"
    vCStock.Trabajador = CLng(Text1(3).Text) 'En codigope ponemos el Cliente
    vCStock.Documento = Text1(0).Text
    vCStock.FechaMov = CDate(Text1(1).Text)
    
    While Not miRsAux.EOF
        vCStock.cantidad = miRsAux!cantidad
        vCStock.Importe = miRsAux!ImporteL
        vCStock.LineaDocu = miRsAux!numlinea
        vCStock.codAlmac = miRsAux!codAlmac
        vCStock.codArtic = miRsAux!codArtic
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If vCStock.cantidad <> 0 Then
            B = vCStock.ActualizarStock(False, False)
            If Not B Then miRsAux.MoveLast
        End If
        If B Then miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    If B Then GenerarAlbaransmovalFAZ = True
    
    
eGenerarAlbaransmovalFAZ:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
End Function







Private Sub BloquearCampoTrabajador()
Dim B As Boolean
    imgBuscar(3).Enabled = Modo = 1
    BloquearTxt Text1(3), Modo <> 1
'    b = False
'    If Modo = 1 Then
'        b = True
'    Else
'        If Modo = 3 Or Modo = 4 Then b = vUsu.Nivel = 0
'    End If
    
 
    
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub



'Antiguo toolbar
'With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'
'        .Buttons(10).Image = 10 'Mto Lineas Ofertas
'        .Buttons(11).Image = 26 'Generar Pedido
'        .Buttons(12).Image = 32 'Cargar Plantilla
'        .Buttons(13).Image = 24 'Traer Lineas de Otra Oferta
'
'        .Buttons(16).Image = 30 'Recordatorio
'        .Buttons(17).Image = 31 'Valoracion
'        .Buttons(18).Image = 16 'Imprimir
'        .Buttons(19).Image = 40 'Imprimir factura pro forma
'        .Buttons(20).Image = 32  'GENERAR FAZ directamente con el mismo
'
'        .Buttons(22).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With





Private Sub BotonesToolBarAux()
Dim B As Boolean


    '   5.-  Mantenimiento Lineas
    B = False
    
    If Modo = 2 Then
        B = True
    Else
        If Modo = 5 Then
            'If ModificaLineas > 0 Then b = T
        End If
    End If

        
        
    
    ToolbarAux(0).Buttons(1).Enabled = B
    ToolbarAux(0).Buttons(6).Enabled = B
    ToolbarAux(0).Buttons(7).Enabled = B

    
    If B Then
        If Data2.Recordset Is Nothing Then
            B = False
        Else
            B = Me.Data2.Recordset.RecordCount > 0
        End If
    End If

    
    ToolbarAux(0).Buttons(2).Enabled = B
    ToolbarAux(0).Buttons(3).Enabled = B
    ToolbarAux(0).Buttons(5).Enabled = B



    'Documentos
    B = Modo = 2
    ToolbarAux(1).Buttons(1).Enabled = B
    ToolbarAux(1).Buttons(2).Enabled = False
    ToolbarAux(1).Buttons(3).Enabled = B And ListView1.ListItems.Count > 0



End Sub



Private Sub CargarFiltro()
    cbFiltro.Clear
    cbFiltro.AddItem ("Sin filtro")
    cbFiltro.AddItem ("Pdtes aceptadas")
    cbFiltro.AddItem ("Pdtes NO aceptadas")
    cbFiltro.AddItem ("Pdtes ult. año")
    cbFiltro.AddItem ("En pedido")
    cbFiltro.AddItem ("Archivadas")
    
    If True Then cbFiltro.ListIndex = 1
    
End Sub

Private Sub PonerFiltro(ByRef CadenaSQL As String)

    If CadenaSQL <> "" Then CadenaSQL = CadenaSQL & " AND "
    Select Case cbFiltro.ListIndex
    Case 1
            'Pdtes aceptadas
            CadenaSQL = CadenaSQL & " aceptado =1 "
    Case 2
            'Pdtes NO aceptadas
            CadenaSQL = CadenaSQL & " aceptado = 0"
    Case 3
            
            'Pdtes ult. año
            CadenaSQL = CadenaSQL & " fecofert>= " & DBSet(DateAdd("yyyy", -1, Now()), "F")
    Case 4
            'En pedido
            CadenaSQL = CadenaSQL & " numpedcl >0  "
    Case 5
            CadenaSQL = CadenaSQL & " motivoTraspaso >0 "
    Case Else
            CadenaSQL = CadenaSQL & " TRUE " 'por si acaso
    End Select
                                            'Ni en pedido ni arhcivada
    If cbFiltro.ListIndex > 0 And cbFiltro.ListIndex < 4 Then CadenaSQL = CadenaSQL & " AND numpedcl is null AND motivoTraspaso is null"
    
    
    If Me.chkMisOfertas.Value = 1 Then
        If FrameFiltro.Tag = "" Then
            FrameFiltro.Tag = PonerTrabajadorConectado("")
            If FrameFiltro.Tag = "" Then
                FrameFiltro.Tag = " false "
            Else
                FrameFiltro.Tag = NombreTabla & ".codtraba = " & FrameFiltro.Tag
            End If
        End If
        CadenaSQL = CadenaSQL & " AND " & FrameFiltro.Tag
                
    End If
    
    
End Sub











'******************************************************
'******************************************************
'***** Documentos adjuntos                  ***********
'******************************************************
'******************************************************
'******************************************************
Private Sub PulsaBotonPdf(Index As Integer)
    
    
    If Modo <> 2 Then Exit Sub
    
    If Index > 0 Then
        CadenaDesdeOtroForm = ""
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then CadenaDesdeOtroForm = "OK"
        End If
        If CadenaDesdeOtroForm = "" Then
            MsgBox "Seleccione un documento", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If Index = 2 Then
        If MsgBox("Seguro que desea eliminar el documento seleccionado?", vbQuestion + vbYesNo) = vbYes Then
            Precio = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
            If EliminarArhivoPDFOferta(Precio) Then
                'Precio = "DELETE FROM `sliprepdfs` WHERE numofert =" & Text1(0).Text
                'Precio = Precio & " AND numlinea = " & Mid(ListView1.SelectedItem.Key, 2)
                'ejecutar Precio, False
                
                cargaDocumentos
            End If
        
        
        End If
    Else
        If Index = 1 Then Exit Sub
        If CarpetaDestinoOferta = "" Then
          '  If Year(Data1.Recordset!fecofert) < 2017 Then
          '      MsgBox "Carpeta destino no esta establecida", vbExclamation
          '      Exit Sub
          '  End If
            
            'Creamos la estructura (Deberia haber sido creada con anterioridad, al dar nuevo. Pero para las de 2016 dejo pasar
            CrearCarpeta
            
        End If
        CadenaDesdeOtroForm = ""
        If Index = 1 Then CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
        frmEuler.Opcion = Index
        frmEuler.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 0 Then
                'Insertamos documento
                Precio = RecuperaValor(CadenaDesdeOtroForm, 1)
                If Dir(Precio, vbArchive) = "" Then
                    MsgBox "No existe fichero", vbExclamation
                    
                Else
                    
                    If CarpetaDestinoOferta = "" Then
                        'Crear estructura
                        
                        MsgBox "falta Crear estructura", vbExclamation
                        Exit Sub
                    End If
                    Precio = RecuperaValor(CadenaDesdeOtroForm, 2)
                    txtAnterior = NombreArchivoEULER(Precio)
                
                    If CopiaArhivoPDFOfertaEuler(CarpetaDestinoOferta, RecuperaValor(CadenaDesdeOtroForm, 1)) Then
                        Precio = DevuelveDesdeBD(conAri, "max(numlinea)", "sliprePdfs", "numofert", Text1(0).Text)
                        If Precio = "" Then Precio = "0"
                        Precio = CStr(Val(Precio) + 1)
                
                        'INSERT INTO BD
                        Precio = "insert into `sliprepdfs` (`numofert`,`numlinea`,`ficheroDesc`,`ficheroNombre`) values ( " & Text1(0).Text & "," & Precio & ","
                        Precio = Precio & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & "," & DBSet(txtAnterior & ".pdf", "T") & ")"
                        'ejecutar Precio, False
                        
                        cargaDocumentos
                    End If
                
                End If
            Else
                'Modificar
                Precio = "UPDATE sliprepdfs SET ficheroDesc = " & DBSet(CadenaDesdeOtroForm, "T")
                Precio = Precio & " WHERE numofert = " & Text1(0).Text & " AND numlinea = " & Mid(ListView1.SelectedItem.Key, 2)
                ejecutar Precio, False
                ListView1.SelectedItem.Text = CadenaDesdeOtroForm
            End If
        End If
    End If
End Sub

Private Sub cargaDocumentos()
Dim i As Integer
Dim Archvi As String
    Me.ListView1.ListItems.Clear
    
    
    Archvi = "SI"
    
    i = Year(CDate(Text1(2).Text))
    'If i < 2017 Then Archvi = ""
    
    CarpetaDestinoOferta = ComprobarExisteCarpetaPDFOferta(i, CLng(Text1(4).Text), CLng(Text1(0).Text))
    If CarpetaDestinoOferta = "" Then Archvi = ""
    
    If Archvi = "" Then
        ListView1_Click
        Exit Sub
    End If
    

    i = 0
    Archvi = Dir(CarpetaDestinoOferta, vbDirectory)   ' Recupera la primera entrada.
    Do While Archvi <> ""   ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If Archvi <> "." And Archvi <> ".." Then
           ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
           If (GetAttr(CarpetaDestinoOferta & Archvi) And vbDirectory) <> vbDirectory Then
                'Debug.Print MiNombre   ' Muestra la entrada
                
                 i = i + 1
                 Me.ListView1.ListItems.Add , "C" & i, Archvi
                 
                 Me.ListView1.ListItems(i).SubItems(1) = Archvi
                 
           End If   ' solamente si representa un directorio.
        End If
        Archvi = Dir   ' Obtiene siguiente entrada.
    Loop

    
    
    ListView1_Click
    
    
End Sub




Private Function CargaArchivo(Archivo As String) As Boolean
    
    On Error GoTo eCargaArchivo
    CargaArchivo = False
    
    If Archivo = "" Then
        AcroPDF1.visible = False
    Else
        AcroPDF1.LoadFile (Archivo)
        AcroPDF1.Tag = Archivo
        AcroPDF1.visible = True
    End If
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function

Private Function CrearCarpeta() As Boolean
Dim Aux As String
 On Error Resume Next
 
    CrearCarpeta = False
    Aux = EulerParam & "\Ofertas\" & Year(Data1.Recordset!fecofert)
    If Dir(Aux, vbDirectory) = "" Then
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Aux
            Exit Function
        End If
    End If
    
    Aux = Aux & "\" & Format(Data1.Recordset!codClien, "000000") & " " & Data1.Recordset!NomClien
    If Dir(Aux, vbDirectory) = "" Then
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Aux
            Exit Function
        End If
    End If
    
    Aux = Aux & "\" & Trim(Format(Data1.Recordset!NumOfert, "0000000") & " " & DBLet(Data1.Recordset!referenc, "T"))
    
    If Dir(Aux, vbDirectory) <> "" Then
        'YA existe
    
    Else
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
            Aux = ""
        Else
            'Intentamos crear documentacion interna
             MkDir Aux & "\Documentacion interna"
             If Err.Number <> 0 Then Err.Clear
        End If
        
    End If
    CarpetaDestinoOferta = Aux & "\"
End Function

