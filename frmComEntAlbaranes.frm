VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComEntAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11985
   Icon            =   "frmComEntAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   3380
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   116
      Text            =   "nom ccoste"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Frame Frame2 
      Height          =   930
      Left            =   120
      TabIndex        =   60
      Top             =   420
      Width           =   11760
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7050
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Cod. Proveedor|N|N|0|999999|scaalp|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   540
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7880
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Proveedor|T|N|||scaalp|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   540
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albaran|F|N|||scaalp|fechaalb|dd/mm/yyyy|S|"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Albaran|T|N|0||scaalp|numalbar||S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   7050
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scaalp|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   7880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
         Text            =   "Text2"
         Top             =   180
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3000
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Pedido|N|S|0||scaalp|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   3975
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Pedido|F|S|||scaalp|fecpedpr|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   6770
         Picture         =   "frmComEntAlbaranes.frx":000C
         ToolTipText     =   "Buscar trabajador"
         Top             =   190
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6770
         Picture         =   "frmComEntAlbaranes.frx":010E
         ToolTipText     =   "Buscar proveedor"
         Top             =   580
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   5700
         TabIndex        =   67
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alb."
         Height          =   255
         Index           =   14
         Left            =   1440
         TabIndex        =   66
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2355
         Picture         =   "frmComEntAlbaranes.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   65
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada Por"
         Height          =   255
         Index           =   21
         Left            =   5700
         TabIndex        =   64
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   63
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   10
         Left            =   3975
         TabIndex        =   62
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   5800
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   31
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10530
      TabIndex        =   28
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   5880
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1080
      Top             =   4800
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
      TabIndex        =   32
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
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
            Object.ToolTipText     =   "Cambiar proveedor"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mover a historico"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Series"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir etiquetas estanteria"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impirmir"
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
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   50
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   103
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   9000
         MaxLength       =   15
         TabIndex        =   102
         Text            =   "TOTAL"
         Top             =   100
         Width           =   885
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7560
         TabIndex        =   33
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   960
      Top             =   5160
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
      Height          =   4380
      Left            =   120
      TabIndex        =   34
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1395
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmComEntAlbaranes.frx":029B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(35)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdAux(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAux(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "FrameCliente"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2(17)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2(16)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAux(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdAux(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmComEntAlbaranes.frx":02B7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(29)"
      Tab(1).Control(1)=   "Text1(28)"
      Tab(1).Control(2)=   "Text1(27)"
      Tab(1).Control(3)=   "Text1(26)"
      Tab(1).Control(4)=   "Text2(26)"
      Tab(1).Control(5)=   "chkDocArchi"
      Tab(1).Control(6)=   "Text1(25)"
      Tab(1).Control(7)=   "FrameHco"
      Tab(1).Control(8)=   "Text2(21)"
      Tab(1).Control(9)=   "Text1(21)"
      Tab(1).Control(10)=   "Text1(19)"
      Tab(1).Control(11)=   "Text1(18)"
      Tab(1).Control(12)=   "Text1(17)"
      Tab(1).Control(13)=   "Text1(16)"
      Tab(1).Control(14)=   "Text1(15)"
      Tab(1).Control(15)=   "Label1(48)"
      Tab(1).Control(16)=   "Label1(47)"
      Tab(1).Control(17)=   "imgFecha(2)"
      Tab(1).Control(18)=   "Label1(4)"
      Tab(1).Control(19)=   "imgBuscar(8)"
      Tab(1).Control(20)=   "imgFecha(1)"
      Tab(1).Control(21)=   "Label1(3)"
      Tab(1).Control(22)=   "imgBuscar(4)"
      Tab(1).Control(23)=   "Label1(1)"
      Tab(1).Control(24)=   "Label1(45)"
      Tab(1).Control(25)=   "Label1(44)"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmComEntAlbaranes.frx":02D3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -69120
         MaxLength       =   80
         TabIndex        =   21
         Tag             =   "T|T|S|||scaalp|SReferencia||N|"
         Top             =   2160
         Width           =   5205
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   28
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   20
         Tag             =   "O|T|S|||scaalp|NReferencia||N|"
         Top             =   2160
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -74880
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Fecha entraga|F|S|||scaalp|fecentrega|dd/mm/yyyy|N|"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   122
         Tag             =   "Envio|N|S|0|9999|scaalp|codenvio|0000|N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   26
         Left            =   -72600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   121
         Text            =   "Text2"
         Top             =   960
         Width           =   3405
      End
      Begin VB.CheckBox chkDocArchi 
         Caption         =   "Documento archivado"
         Height          =   330
         Left            =   -73200
         TabIndex        =   18
         Tag             =   "Ar|N|S|||scaalp|docarchiv|||"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -74880
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Fecha recepcion|F|S|||scaalp|fecenvio|dd/mm/yyyy||"
         Top             =   1545
         Width           =   1185
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   10920
         TabIndex        =   118
         ToolTipText     =   "Buscar centro coste"
         Top             =   3480
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   44
         Tag             =   "centro coste"
         Text            =   "cc"
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   10320
         MaxLength       =   3
         TabIndex        =   115
         Tag             =   "IVA"
         Text            =   "IVA"
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   16
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   45
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   3960
         Visible         =   0   'False
         Width           =   6765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   17
         Left            =   9460
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   46
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1440
         Left            =   -69120
         TabIndex        =   104
         Top             =   360
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   109
            Top             =   260
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   1455
            MaxLength       =   30
            TabIndex        =   108
            Text            =   "Text1"
            Top             =   630
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   23
            Left            =   2115
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   107
            Text            =   "Text2"
            Top             =   630
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   1455
            MaxLength       =   30
            TabIndex        =   106
            Text            =   "Text1"
            Top             =   1000
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   2115
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   105
            Text            =   "Text2"
            Top             =   1000
            Width           =   3525
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   112
            Top             =   260
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   111
            Top             =   630
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1080
            Picture         =   "frmComEntAlbaranes.frx":02EF
            ToolTipText     =   "Buscar trabajador"
            Top             =   630
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   110
            Top             =   1000
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmComEntAlbaranes.frx":03F1
            ToolTipText     =   "Buscar incidencia"
            Top             =   1000
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74520
         TabIndex        =   70
         Top             =   600
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   87
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   86
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   85
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   84
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   83
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   82
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   81
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   80
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   79
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   78
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   77
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   76
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   75
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   74
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   73
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   72
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   71
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   101
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   100
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL ALBARAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   39
            Left            =   5640
            TabIndex        =   99
            Top             =   2660
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "+"
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
            Index           =   36
            Left            =   11880
            TabIndex        =   98
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "+"
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
            Index           =   8
            Left            =   7320
            TabIndex        =   97
            Top             =   1320
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7560
            TabIndex        =   96
            Top             =   1230
            Width           =   1335
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
            TabIndex        =   95
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
            Index           =   31
            Left            =   3720
            TabIndex        =   94
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
            TabIndex        =   93
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   18
            Left            =   5760
            TabIndex        =   92
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   22
            Left            =   3960
            TabIndex        =   91
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   23
            Left            =   2160
            TabIndex        =   90
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   27
            Left            =   5760
            TabIndex        =   88
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   21
         Left            =   -72660
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   480
         Width           =   3405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "Trab. Pedido|N|S|0|9999|scaalp|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   660
      End
      Begin VB.Frame FrameCliente 
         Height          =   1400
         Left            =   220
         TabIndex        =   50
         Top             =   315
         Width           =   11400
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scaalp|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   960
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1245
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scaalp|codpobla||N|"
            Text            =   "Text15"
            Top             =   916
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1875
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scaalp|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   916
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3435
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Proveedor|T|S|||scaalp|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   190
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Proveedor|T|N|||scaalp|nifprove||N|"
            Text            =   "123456789"
            Top             =   190
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Forma de Pago|N|N|0|999|scaalp|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   190
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   52
            Text            =   "Text2"
            Top             =   190
            Width           =   3660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   14
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalp|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   553
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   8445
            MaxLength       =   7
            TabIndex        =   15
            Tag             =   "Descuento General|N|N|0|99.90|scaalp|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   553
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scaalp|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   553
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   960
            Picture         =   "frmComEntAlbaranes.frx":04F3
            ToolTipText     =   "Buscar proveedor vario"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   975
            Picture         =   "frmComEntAlbaranes.frx":05F5
            ToolTipText     =   "Buscar población"
            Top             =   916
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   5700
            TabIndex        =   59
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   58
            Top             =   916
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2745
            TabIndex        =   57
            Top             =   190
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   56
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   55
            Top             =   190
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P. Pago"
            Height          =   255
            Index           =   25
            Left            =   5700
            TabIndex        =   54
            Top             =   555
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   7740
            TabIndex        =   53
            Top             =   553
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   6600
            Picture         =   "frmComEntAlbaranes.frx":06F7
            ToolTipText     =   "Buscar forma de pago"
            Top             =   190
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   51
            Top             =   553
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   49
         ToolTipText     =   "Buscar artículo"
         Top             =   3540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   48
         ToolTipText     =   "Buscar almacen"
         Top             =   3540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   38
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   16
         TabIndex        =   43
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   42
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   41
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   40
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   39
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   37
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3540
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   -72240
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observación 5|T|S|||scaalp|observa5||N|"
         Top             =   3960
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   18
         Left            =   -72240
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Observación 4|T|S|||scaalp|observa4||N|"
         Top             =   3660
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   -72240
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Observación 3|T|S|||scaalp|observa3||N|"
         Top             =   3360
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   -72240
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Observación 2|T|S|||scaalp|observa2||N|"
         Top             =   3060
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   -72240
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Observación 1|T|S|||scaalp|observa1||N|"
         Top             =   2760
         Width           =   8445
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntAlbaranes.frx":07F9
         Height          =   2025
         Left            =   225
         TabIndex        =   47
         Top             =   1860
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   3572
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Su referencia"
         Height          =   255
         Index           =   48
         Left            =   -69120
         TabIndex        =   126
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nuestra referencia"
         Height          =   255
         Index           =   47
         Left            =   -73200
         TabIndex        =   125
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   -73920
         Picture         =   "frmComEntAlbaranes.frx":080E
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma de envio"
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   123
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -73680
         Picture         =   "frmComEntAlbaranes.frx":0899
         ToolTipText     =   "Buscar trabajador"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   -73920
         Picture         =   "frmComEntAlbaranes.frx":099B
         ToolTipText     =   "Buscar fecha"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. archiv"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   120
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   480
         TabIndex        =   114
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Lote"
         Height          =   255
         Index           =   2
         Left            =   8880
         TabIndex        =   113
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   -73800
         Picture         =   "frmComEntAlbaranes.frx":0A26
         ToolTipText     =   "Buscar trabajador"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trab. Pedido"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74280
         TabIndex        =   35
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recogida"
         Height          =   255
         Index           =   44
         Left            =   -74880
         TabIndex        =   124
         Top             =   1965
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10530
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblF 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   119
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Centro coste"
      Height          =   255
      Index           =   46
      Left            =   2400
      TabIndex        =   117
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
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
Attribute VB_Name = "frmComEntAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Long 'Codigo de Proveedor

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
'cadena que selecciona los albaranes de un proveedor para mostrar
'antes de facturarlos
Public cadSelAlbaranes As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProve As frmComProveedores  'Form Mto Proveedores
Attribute frmProve.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV   'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1

Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1


'-------------------------------------------------------------------------
Private Modo As Byte
'-----------------------------
'Se distinguen varios Modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
'4.- Mantenimiento de Nº de Serie

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim primeravez As Boolean
Dim PrimeraVezForm As Boolean

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
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

Dim cadList As String 'cadena para pasar al historico

'Febrero 2010
'Despues de meter la primera linea, continua ofertando el almacen de la linea anterior
Dim AlmacenLineas As Integer  '

Dim PulsadoMas2 As Boolean


'Diciembre 2010.
'Para ver si al volver a cabecera muestra el listado o no
Dim HaModifEnLineas As Boolean






Private Sub chkDocArchi_Click()
   '  If Modo = 1 Then CheckCadenaBusqueda chkDocArchi, BuscaChekc
End Sub
Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub










Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String
Dim precioUC As Currency
Dim FechaUltCompra As Date
Dim EnPromocionOPrecioEspecial As String
Dim ParaElMargen As Currency

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR CABECERA
            If DatosOk Then InsertarCabecera
            
        Case 4  'MODIFICAR CABECERA
            If DatosOk Then
                If ModificarCabAlbaran Then
                    If cadSelAlbaranes = "" Then TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Antes de insertar la linea guardamos el sartic.preciouc actual
            'para aplicar margen despues, pq en Insertar linea se actualiza ya el preciouc
            If txtAux(1).Text = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "Campos obligatorios", vbExclamation
                PonerFoco txtAux(1)
                Exit Sub
            End If
            numlinea = "ultfecco"
            precioUC = ComprobarCero(DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T", numlinea))
            'Fecha ultima compra
            FechaUltCompra = "01/01/1900"
            If numlinea <> "" Then
                FechaUltCompra = CDate(numlinea)
                numlinea = ""
            End If
                
         
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                          
                If InsertarLinea(numlinea) Then
                    'Comprobar si Hay Nº SERIE en compras y Mostrar
                    'ventana para pedir los Nº Serie de la cantidad introducida
                    If vParamAplic.NumSeries Then
                        ComprobarNumSeries (numlinea)
                    End If
                    
                    'Comprobar si se ha modificado el precio desde la ultima compra
                    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
                    'y el precio de las TArifas aplicandole el margen
                    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
                    'If precioUC <> CCur(txtAux(4).Text) Then
                    If CCur(txtAux(3).Text) <> 0 Then
                            If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
                                If CDate(Text1(1).Text) >= FechaUltCompra Then
                                
                                    'Marzo. Para saber si esta en pormociones/ofertas
                                    precioUC = 0
                                    
                                    EnPromocionOPrecioEspecial = "fechaini<=" & DBSet(Text1(1).Text, "F")
                                    EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & " AND fechafin>=" & DBSet(Text1(1).Text, "F") & " and codartic"
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", EnPromocionOPrecioEspecial, txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial <> "" Then precioUC = 1
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "sprees", "codartic", txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial <> "" Then precioUC = precioUC + 2
                                    If precioUC = 0 Then
                                        EnPromocionOPrecioEspecial = ""
                                    Else
                                        EnPromocionOPrecioEspecial = "ATENCION. Artículo en:"
                                        If precioUC = 1 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                                        If precioUC = 2 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                                        EnPromocionOPrecioEspecial = vbCrLf & String(20, "*") & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                                        EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial
                                    End If
    
                                    EnPromocionOPrecioEspecial = "Se ha modificado el precio última compra." & vbCrLf & "¿Desea actualizar los precios de venta?" & EnPromocionOPrecioEspecial
                                    If MsgBox(EnPromocionOPrecioEspecial, vbQuestion + vbYesNo) = vbYes Then
                                        'Comprobar que el artículo tiene margen comercial
                                        If ArticuloTieneMargen(txtAux(1).Text) Then
          
                                                frmComActPrecios.parCodArtic = txtAux(1).Text
                                                frmComActPrecios.parNomArtic = txtAux(2).Text
                                                frmComActPrecios.Show vbModal
            
                                        End If
                                    Else
                                    
                                        If vParamAplic.RecalculoMargen Then ActualizacionAutomaticaMargen txtAux(1).Text
                                              
                                    End If
                                    
                                    'Veremos si es un articulo "escandallo"
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial = "" Then EnPromocionOPrecioEspecial = "0"
                                    If Val(EnPromocionOPrecioEspecial) > 0 Then
                                        frmListado4.Opcion = 1
                                        frmListado4.vCadena = txtAux(1).Text & "|"
                                        frmListado4.Show vbModal
                                    End If
                                    EnPromocionOPrecioEspecial = ""
                                    
                                End If   'Fecha ultima compra
                            End If  'Precio ultima compra
                    End If  'Por si han puesto cantidad=0
                    
                    
                    'Abril 2013
                    'INsertar linea de componentes
                    InsertarSubComponentes
                     
                     
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If DatosOkLinea() Then
                    If ModificarLinea Then
                        
                        
                        
                        If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
                        
                               'Marzo. Para saber si esta en pormociones/ofertas
                                precioUC = 0
                                EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", txtAux(1).Text, "T")
                                If EnPromocionOPrecioEspecial <> "" Then precioUC = 1
                                EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", txtAux(1).Text, "T")
                                If EnPromocionOPrecioEspecial <> "" Then precioUC = precioUC + 2
                                If precioUC = 0 Then
                                    EnPromocionOPrecioEspecial = ""
                                Else
                                    EnPromocionOPrecioEspecial = "ATENCION. Artículo en:"
                                    If precioUC = 1 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                                    If precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                                    EnPromocionOPrecioEspecial = vbCrLf & String(20, "*") & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                                    EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial
                                End If

                                EnPromocionOPrecioEspecial = "Se ha modificado el precio última compra." & vbCrLf & "¿Desea actualizar los precios de venta?" & EnPromocionOPrecioEspecial
                    
                        
                        
                        
                            If MsgBox(EnPromocionOPrecioEspecial, vbQuestion + vbYesNo) = vbYes Then
                                'Comprobar que el artículo tiene margen comercial
                                If ArticuloTieneMargen(txtAux(1).Text) Then
                                    'Aplicar margen comercial a los precios
                                    'Modificar precios de venta en articulo y tarifas
                                    frmComActPrecios.parCodArtic = txtAux(1).Text
                                    frmComActPrecios.parNomArtic = txtAux(2).Text
        '                            frmcomactprecios.parPrecioUC =
                                    frmComActPrecios.Show vbModal
                                End If
                            Else
                                If vParamAplic.RecalculoMargen Then ActualizacionAutomaticaMargen txtAux(1).Text
                            End If
                            
                            'Veremos si es un articulo "escandallo"
                            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", txtAux(1).Text, "T")
                            If EnPromocionOPrecioEspecial = "" Then EnPromocionOPrecioEspecial = "0"
                            If Val(EnPromocionOPrecioEspecial) > 0 Then
                                frmListado4.Opcion = 1
                                frmListado4.vCadena = txtAux(1).Text & "|"
                                frmListado4.Show vbModal
                            End If
                            EnPromocionOPrecioEspecial = ""
                            
                            
                        End If
                        
                        NumRegElim = Data2.Recordset!numlinea
                        TerminaBloquear
                        CargaTxtAux False, False
                        CargaGrid2 DataGrid1, Data2
                        ModificaLineas = 0
                        PonerBotonCabecera True
                        BloquearTxt Text2(16), True
                        BloquearTxt Text2(17), True
                        
                        'AQUI
                        PosicionarData2
                        
                    End If
                    Me.DataGrid1.Enabled = True
                End If
            End If
            CalcularDatosFactura 'rellenar campos pestaña de totales
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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



Private Function ComprobarCambioFecha() As Boolean
'Comprueba si se ha modificado el campo fecha de la cabecera.
'Ya que es clave primaria y se deberan cambiar tambien la fecha
'en tablas sliap y smoval
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Izquierda As String, Derecha As String
Dim llis As Collection
Dim I As Integer
Dim b As Boolean


    If Data1.Recordset.EOF Then Exit Function

    
    If (CDate(Text1(1).Text) <> CDate(Data1.Recordset!FechaAlb)) Then
    'si ha modificado la fecha de albaran
        On Error GoTo EComprobar
        
        'seleccionar todas las lineas de ese albaran para actualizar la fecha (slialp)
        SQL = "SELECT * FROM " & NomTablaLineas & " WHERE numalbar=" & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Set llis = New Collection
            
        'Nos guardamos todas las lineas con la modificacion de la fecha para
        'volverlas a insertar
        BACKUP_TablaIzquierda RS, Izquierda
        
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha, "fechaalb", CStr(Text1(1).Text)
            llis.Add Derecha
            RS.MoveNext
        Wend
        
        RS.Close
       
        
        
        'Numeros de lotes fitosantiarios
        'Cambiamos en slotes
        If vParamAplic.ManipuladorFitosanitarios2 Then
            SQL = " Select * from " & NomTablaLineas & " WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
            SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
            SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
            
            
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                
                SQL = "UPDATE slotes SET fecentra = " & DBSet(Text1(1).Text, "F") & " WHERE "
                SQL = SQL & " codartic=" & DBSet(RS!codArtic, "T") & " AND numlotes=" & DBSet(RS!numlotes, "T") & " AND fecentra=" & DBSet(RS!FechaAlb, "F")
                conn.Execute SQL
                RS.MoveNext
            Wend
            RS.Close
            
        End If


        
        Set RS = Nothing
        
        
        
        
        
        
        'Eliminamos las lineas que tenia ese albaran (slialp) para volverlas a insertar con la fecha nueva
        SQL = "DELETE from slialp WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
        
        'Actualizamos la fecha en la cabecera (scaalp)
        SQL = "UPDATE scaalp SET fechaalb = " & DBSet(Text1(1).Text, "F")
        SQL = SQL & " WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
        
        'Actualizamos la fecha en la tabla smoval
        SQL = "UPDATE smoval SET fechamov=" & DBSet(Text1(1).Text, "F")
        SQL = SQL & " WHERE document = " & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codigope=" & Data1.Recordset!Codprove
        SQL = SQL & " AND detamovi='" & CodTipoMov & "'"
        conn.Execute SQL
        
        
        'Actualizar la fecha compra en los numeros de serie del albaran (si tiene articulos con num. serie)
        SQL = "UPDATE sserie SET fechacom=" & DBSet(Text1(1).Text, "F")
        SQL = SQL & " WHERE fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND "
        SQL = SQL & " numalbpr=" & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
            
        
                
        'Volvemos a insertar las lineas con la fecha correcta (slialp)
        SQL = ""
        For I = 1 To llis.Count
            If (I Mod 10) = 0 Then
                SQL = SQL & CStr(llis(I)) & ","
                SQL = Mid(SQL, 1, Len(SQL) - 1) 'quitamos ultima coma
                SQL = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & SQL & ";"
                conn.Execute SQL
                SQL = ""
            Else
                SQL = SQL & CStr(llis(I)) & ","
            End If
        Next I
        
        If SQL <> "" Then
            SQL = Mid(SQL, 1, Len(SQL) - 1) 'quitamos ultima coma
            SQL = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & SQL & ";"
            conn.Execute SQL
            SQL = ""
        End If
        Set llis = Nothing
    End If
    b = True
    
EComprobar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "El campo fecha no se ha podido modificar", Err.Description
        b = False
    End If
    If b Then
        ComprobarCambioFecha = True
    Else
        ComprobarCambioFecha = False
    End If
End Function



Private Sub cmdAux_Click(index As Integer)
    Select Case index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco txtAux(index)
            
        Case 1 'Busqueda de Cod. Artic
            Set FrmArt = New frmAlmArticu2
            'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo búsqueda
            
            FrmArt.Show vbModal
            Set FrmArt = Nothing
            PonerFoco txtAux(index)
            
        Case 2 'COD. CENTRO COSTE
            If vEmpresa.TieneAnalitica Then
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtAux(9)
            End If
    End Select
    
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
            BloquearTxt Text2(17), True
            DataGrid1.Columns(5).Caption = "Articulo"
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            lblF.Caption = ""
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Albaranes: scaalp (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    Text1(2).Text = PonerTrabajadorConectado(NomTraba)
    Text2(2).Text = NomTraba
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    HaModifEnLineas = True
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = Format(AlmacenLineas, "000")
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    Text2(17).Text = ""
     txtAux2(9).Text = ""
    BloquearTxt Text2(16), False
    BloquearTxt Text2(17), True
    
    
    ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(2).Text, "N")
        Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
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
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia cadSelAlbaranes
    Else
        LimpiarCampos
        LimpiarDataGrids
        If cadSelAlbaranes = "" Then
            CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        Else
            CadenaConsulta = "Select * from " & NombreTabla & " " & " WHERE " & cadSelAlbaranes & Ordenacion
        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim SQL As String
Dim DeVarios As Boolean

    On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(2)
    
    If EsDeVarios Then
        If Data1.Recordset.EOF Then Exit Sub
        SQL = " SELECT * FROM sprvar WHERE nifprove='" & Data1.Recordset!nifProve & "' FOR UPDATE "
        conn.Execute SQL
    End If
    
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(4).Text)
    BloquearDatosProve (DeVarios)
    
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
'Dim cArt As CArticulo  'Si bloquearamos por el articulo

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    HaModifEnLineas = True
    
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    ModificaLineas = 2 'Modificar
    CargaTxtAux True, False
    'ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    
    'bloquear el num_lote si el articulo es de una categoria q no lleva control
    'de nº de lote
'    BloquearTxt Text2(17), (DBLet(Data2.Recordset!numlotes, "T") = "")
'    Set cArt = New CArticulo
'    If cArt.LeerDatos(Data2.Recordset!codArtic) Then
'        BloquearTxt Text2(17), Not cArt.TieneNumLote
'    End If
'    Set cArt = Nothing
    BloquearTxt Text2(17), False
    
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de albaranes compras (scaalp)
' y los registros correspondientes de las tablas de lineas (slialp)
'al eliminar un albaran ademas habrá que restaurar valores:
' - actualizar stock en (salmac)
' - eliminar los movimientos que inserto el albaran en (smoval)
' - actualizar el ultprecio compra y ultima fecha compra en funcion del ult. movimiento ALC en smoval
' - reestablecer el precio medio ponderado
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    
    If Not PuedeEliminar(False) Then Exit Sub
    
    cad = "Cabecera de Albaranes Compras" & vbCrLf
    cad = cad & "-------------------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Albaran:            "
    cad = cad & vbCrLf & "Nº:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha: " & Text1(1).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? " & vbCrLf & vbCrLf
    cad = cad & "-------------------------------------------------"
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
    
        NumRegElim = Data1.Recordset.AbsolutePosition

        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        If Not Eliminar() Then
            CargaGrid Me.DataGrid1, Me.Data2, True
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
    
    If Not PuedeEliminar(True) Then Exit Sub
    
    
    
    HaModifEnLineas = True
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de Albaran?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            SituarDataTrasEliminar Data2, NumRegElim
            CargaGrid2 DataGrid1, Data2
            CalcularDatosFactura
        End If
'        CancelaADODC
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
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
        If HaModifEnLineas Then ComprobarPedidosClientesDesdeAlbProveedor Text1(0).Text, Text1(1).Text, Text1(4).Text
        
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub




Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim devuelve As String
Dim CadLote As String

    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        'cadLote = "numlotes"
        'devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numalbar", Text1(0).Text, "T", cadLote, "numlinea", Data2.Recordset!numlinea, "N")
        'Poner descripcion de ampliacion lineas
        Text2(16).Text = DBLet(Data2.Recordset!Ampliaci, "T")
        'poner el numero de lote
        Text2(17).Text = DBLet(Data2.Recordset!numlotes, "T")
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
        If vEmpresa.TieneAnalitica Then
            Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            txtAux2(9).Text = ""
        End If
        
    End If
    
    If ModificaLineas = 1 Then Text2(16).Text = ""
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_activate()
    Screen.MousePointer = vbDefault
    If PrimeraVezForm Then
        PrimeraVezForm = False
        
        'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
        If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
        
        'Viene de click en VerAlbaranes en formulario de "Recepcion de Facturas compra"
        If cadSelAlbaranes <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 20
    Modo = 0
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        
        .Buttons(9).Image = 45 'Mto Lineas Albaran
        .Buttons(10).Image = 10 'Mto Lineas Albaran
        .Buttons(12).Image = 32 'Pasar a hco pero sin mover la smoval ni precios ni "leches"
        .Buttons(14).Image = 33 'Nº Serie
        .Buttons(15).Image = 40 'Imprimir etiquetas estanteria
        .Buttons(16).Image = 16 'Imprimir Albaran proveedor (REA)
        
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
    cmdAux(2).Tag = "-1"
   
    'Sept 2010
    'Todos podran imprimirse un albaran
    'La imprimir solo es posible para albaranes a socios(REA)
    'Toolbar1.Buttons(12).visible = vParamAplic.IVA_REA > 0
    
    CodTipoMov = "ALC"
    VieneDeBuscar = False

    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    If Not EsHistorico Then
        NombreTabla = "scaalp"
        NomTablaLineas = "slialp" 'Tabla lineas de Albaranes
        Me.Caption = "Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar, fechaalb,codprove "
    Else
        NombreTabla = "schalp"
        NomTablaLineas = "slhalp"
        CargarTagsHco Me, "scaalp", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(22).Tag = "Fecha Eliminación|F|N|||schalp|fechelim|dd/mm/yyyy|N|"
        Text1(23).Tag = "Trabajador Eliminación|N|N|0|9999|schalp|trabelim|0000|N|"
        Text1(24).Tag = "Incidencia elim.|T|N|||schalp|codincid||N|"
        Me.Caption = "Histórico Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar,fechaalb,codprove "
    End If
    
         
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
    ElseIf cadSelAlbaranes <> "" Then
        CadenaConsulta = CadenaConsulta & " WHERE " & cadSelAlbaranes
    Else
        CadenaConsulta = CadenaConsulta & " WHERE numalbar = -1"
    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
       
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    primeravez = True
    PrimeraVezForm = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
            Text1(0).BackColor = vbYellow
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    'Aqui va el especifico de cada form es
    '### a mano
    chkDocArchi.Value = 0
    Text3(0).Text = "BASE IMP."
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
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Val(cmdAux(2).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 3)
            CadB = CadB & " and " & Aux
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

        Indice = 9
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
        'provincia
        Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
    Text1(CByte(Me.imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod envio
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom envio

End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 12
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico al eliminar albaran
    cadList = ""
    cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
    cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
    cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
End Sub



Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
Dim Cant As Currency
Dim I As Byte
Dim cadSerie As String
Dim nSerie As CNumSerie

'si llegamos aqui hemos hecho un abono y vamos a eliminar el
'nº de serie de la tabla sserie del articulo que hemos devuelto.

    Cant = CCur(txtAux(3).Text)
    Cant = Abs(Cant)

    'Para cada valor empipado actualizar la tabla sserie
    On Error GoTo ErrorNSerie

    For I = 1 To Cant
        cadSerie = RecuperaValor(CadenaSeleccion, I + 1) 'Cod Forma Pago
        If cadSerie <> "" Then
            Set nSerie = New CNumSerie
            nSerie.numSerie = cadSerie
            nSerie.Articulo = RecuperaValor(CadenaSeleccion, 1)
            
            'como vamos a devolver esos nº serie de ese articulo
            'los eliminamos de la tabla sserie, ya no tenemos esos artículos
            nSerie.EliminarNumSerie
            Set nSerie = Nothing
        End If
    Next I

ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
    End If
End Sub



Private Sub frmNSerie_CargarNumSeries()
'Cuando vuelve del formulario donde se han introducido los nº de Serie a cargar
'Insertar un registro en la tabla "sserie" para cada articulo

    'Estamos en COMPRAS
    If ModificaLineas = 4 Then
        'Viene de boton VErNumSeries de la toolbar, abre la ventana de cargar numSeries
        'y muestra los que tenga asignados el albaran
        CargarNumSeries
    Else
       'Viene de insertar Nº de series al insertar una linea y pasa
       'los valores almacenados en la temporal a la sserie
       InsertarNumSeriesDeTMP
    End If
End Sub


Private Sub frmProve_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
    FormateaCampo Text1(4)
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom prove
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = 2
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case index
        Case 0 'Cod. Prove
            PonerFoco Text1(4)
            Set frmProve = New frmComProveedores
            frmProve.DatosADevolverBusqueda = "0"
            frmProve.Show vbModal
            Set frmProve = Nothing
            Indice = 4
            
        Case 1 'Realizada Por Trabajador
            Indice = 2
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 2 'Cod. Postal
            Indice = 9
            If Not Text1(Indice).Locked Then
                Set frmCP = New frmCPostal
                frmCP.DatosADevolverBusqueda = "0"
                frmCP.Show vbModal
                Set frmCP = Nothing
                
                VieneDeBuscar = True
            End If
            
        Case 3 'Forma de Pago
            Indice = 12
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'NIF proveedor varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0"
            frmPV.Show vbModal
            Set frmPV = Nothing
            Indice = 6
        Case 8
            Indice = 26
            PonerFoco Text1(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmF = New frmCal
    frmF.Fecha = Now
    If index = 0 Then
        Indice = 1 'fecalb
    ElseIf index = 1 Then
        Indice = 25
    Else
        Indice = 27
    End If
    Me.imgFecha(0).Tag = Indice
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
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
         Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Cabecera Albaran
        If cadSelAlbaranes = "" Then
            If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        End If
        BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera de Albaran
         Me.SSTab1.Tab = 0
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

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    If index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
Dim b As Boolean
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        b = False
        If Text1(index).Text = "" Then
            b = True
        Else
            If Text1(index).SelLength = Len(Text1(index).Text) Then b = True
        End If
        If b Then
                Ind = -1
                Select Case index
                Case 2
                    Ind = 1

                Case 4
                    Ind = 0
                Case 6
                    Ind = 5
                Case 9
                    Ind = 2
                Case 12
                    Ind = 3
                Case 21
                    Ind = 4

                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
End Sub


Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
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
Private Sub Text1_LostFocus(index As Integer)
Dim devuelve As String
        
        
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(index).Text = ""
        Exit Sub
    End If
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case index
        Case 1, 25, 27 'Fecha Albaran y fecha arhivo
                If Text1(index).Text <> "" Then PonerFormatoFecha Text1(index)
                
        Case 2 'Cod Trabajador
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(index).Text = ""
            End If
            
        Case 4 'Cod. Proveedor
            If PonerFormatoEntero(Text1(index)) Then
                If Modo = 1 Then 'Busqueda
                    'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                Else 'Si Insertar, recuperar datos de Tabla sprove
                    PonerDatosProveedor (Text1(index).Text)
                End If
            Else
                LimpiarDatosProve
            End If
            
         Case 6 'NIF
            If Not EsDeVarios Or Modo <> 3 Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifProve Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(index).Text)
             
        Case 9 'Cod. Postal
            If Text1(index).Locked Then Exit Sub
            If Text1(index).Text = "" Then
                Text1(index + 1).Text = ""
                Text1(index + 2).Text = ""
                Exit Sub
            End If
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(index + 1).Text = ObtenerPoblacion(Text1(index).Text, devuelve)
                Text1(index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sforpa", "nomforpa")
            Else
                Text2(index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(index), 4) Then 'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
                If index = 14 Then
                    Me.SSTab1.Tab = 1
                    PonerFoco Text1(15)
                End If
            Else
                If index = 14 And Text1(index).Text = "" Then
                    Me.SSTab1.Tab = 1
                    PonerFoco Text1(15)
                End If
            End If
        Case 19
            PonerFocoBtn Me.cmdAceptar
                
                    
        Case 26 'Codenvio
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "senvio", "nomenvio")
            Else
                Text2(index).Text = ""
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If cadSelAlbaranes <> "" Then CadB = CadB & " AND " & cadSelAlbaranes
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "Nº Albaran")
    cad = cad & ParaGrid(Text1(1), 15, "Fecha Alb.")
    cad = cad & ParaGrid(Text1(4), 15, "Provedor")
    cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
    tabla = NombreTabla
    Titulo = "Albaranes"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges

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
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
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
    
    'Trabajador Albaran
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "straba", "nomtraba", "codtraba")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
    
    'Trabajador del Pedido
    Text2(21).Text = PonerNombreDeCod(Text1(21), conAri, "straba", "nomtraba", "codtraba")
    'Cod envio
    Text2(26).Text = PonerNombreDeCod(Text1(26), conAri, "senvio", "nomenvio")
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(23).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
        Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "sincid", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pestaña de totales
     
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Vuelvo a poner el lbl en la columna
    If Modo = 5 Then DataGrid1.Columns(5).Caption = "Articulo"
    lblF.Caption = ""
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    '22 Sept 2010
    'El albaran lo puede imprimir en cualquier empresa
    'If vParamAplic.IVA_REA > 0 Then Toolbar1.Buttons(12).Enabled = b
    Toolbar1.Buttons(12).Enabled = b
    
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Campo Nº Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    BloquearTxt Text1(0), (Modo <> 1) And (Modo <> 3), True
    
    'La fecha de albaran es clave primaria pero dejamos modificarla
    BloquearTxt Text1(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    b = (Modo <> 1)
    'Bloquear los campos de Pedido, excepto en Busqueda
    BloquearTxt Text1(3), b
    BloquearTxt Text1(20), b
    BloquearTxt Text1(21), b
    
    
   
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
     
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 50
        BloquearTxt Text3(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
          
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    If cmdCancelar.visible Then cmdCancelar.Cancel = True
    chkDocArchi.Enabled = b
        
    
    
    
    For I = 0 To Me.imgFecha.Count - 1
'        Me.imgFecha(i).Enabled = b
        BloquearImg imgFecha(I), Not b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(4).Enabled = (Modo = 1)
    Me.imgBuscar(0).Enabled = (Modo = 3 Or Modo = 1)
              
    'Modo Linea de Albaranes. Campo Ampliacion Linea
    Me.Label1(35).visible = (Modo = 5)
    Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    'Modo Linea de Albaranes. Campo num_lote
    Me.Label1(2).visible = (Modo = 5)
    Me.Text2(17).visible = (Modo = 5)
    BloquearTxt Text2(17), True
    
    ' ---- [20/10/2009] [LAURA] : añadir del centro de coste
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    Me.txtAux2(9).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    BloquearTxt txtAux2(9), True
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim cad As String

    On Error GoTo EDatosOK

    DatosOk = False
       
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    If Modo = 3 Then
        'Comprobare que este albaran/proveedor NO existe
        'a que es campo clave en scafpc1
        cad = " year(fechaalb)=" & Year(CDate(Text1(1).Text)) & " AND numalbar=" & DBSet(Text1(0).Text, "T") & " AND codprove"
        cad = DevuelveDesdeBD(conAri, "fechaalb", "scaalp", cad, Text1(4).Text, "N")
        
        If cad = "" Then
            cad = " year(fechaalb)=" & Year(CDate(Text1(1).Text)) & " AND numalbar=" & DBSet(Text1(0).Text, "T") & " AND codprove"
            cad = DevuelveDesdeBD(conAri, "fechaalb", "scafpa", cad, Text1(4).Text, "N")
            If cad <> "" Then cad = cad & " (Facturado)"
        End If
        
        If cad <> "" Then
            cad = Text1(5).Text & vbCrLf & "con fecha " & cad
            cad = "Ya existe el albaran: " & Text1(0).Text & " del proveedor " & vbCrLf & Text1(4).Text & " " & cad
            
            MsgBox cad, vbExclamation
            b = False
        End If
    End If
    
    
    'Febrero 208
    'LLegado aqui veremos las fechas son "razonables"
    If Modo = 3 And b Then
        cad = ""
        If CDate(Text1(1).Text) < vEmpresa.FechaIni Then
            cad = "menor "
        Else
            If CDate(Text1(1).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then cad = "mayor"
        End If
        If cad <> "" Then
            cad = "Fecha " & cad & " que ejercicios.          " & Text1(1).Text & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then b = False
        End If
    End If
        
        
    If Modo = 4 And b Then
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If CDate(Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                'Si tienen lotes asignados NO podremos seguir
                If Not PuedeEliminar(False) Then b = False
            End If
        End If
    End If
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte
Dim cart As CArticulo
Dim Aux As String
Dim DiferenciaCantidad As Currency

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(5), txtAux(6), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(7).Text Then
        Aux = "Importe linea distinto calculado: " & Aux & "  <>  " & txtAux(7).Text & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    

    
    
    
    b = True
    'Comprobar que los campos requeridos tengan valor
    For I = 0 To txtAux.Count - 1
        If txtAux(I).Text = "" Then
            If I = 9 And vEmpresa.TieneAnalitica = False Then
                'no hace nada pq puede ser nulo
            Else
                Screen.MousePointer = vbDefault
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
    
    
    'si el articulo tiene control de numero de lotes, el campo del lote será requerido
    Set cart = New CArticulo
    If cart.LeerDatos(txtAux(1).Text) Then
        If cart.TieneNumLote Then
            If Trim(Text2(17).Text) = "" Then
                b = False
                MsgBox "El nº de lote no puede ser nulo." & vbCrLf & vbCrLf & "El artículo tiene control de lotes.", vbExclamation
                PonerFoco Text2(17)
            End If
            
            
            'Cuando lleve Registro fitosanitario entonces el modificar articulo "cantidad" sera SI no se ha vendido NADA
            '------------------------------------------------------------------------------
            If b And vParamAplic.ManipuladorFitosanitarios2 Then
            
            
                'Modificando
                If ModificaLineas = 2 Then
                    
                    cadList = ""
                    'Si cambia cantidad vendida o nº lote entonces comprobaremos
                    If Data2.Recordset!numlotes <> Me.Text2(17).Text Then
                        cadList = "LO"
                    ElseIf Data2.Recordset!cantidad <> ImporteFormateado(Me.txtAux(3).Text) Then
                        cadList = "OK"
                    End If
            
                    If cadList <> "" Then
                        
                        If cadList = "LO" Then
                            
                            cadList = "numlote =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                            cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                            cadList = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", cadList, "1")
                            If Val(cadList) > 0 Then
                                MsgBox "El lote ya ha sido vendido", vbExclamation
                                b = False
                            Else
                                cadList = "numlote =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                                cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                                cadList = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", cadList, "1")
                                If Val(cadList) > 0 Then
                                    MsgBox "El lote esta siendo vendido", vbExclamation
                                    b = False
                                End If
                            End If
                        Else
                            'Es el mismo lote de venta. Vemos si cambiando la cantidad tiene bastante
                            DiferenciaCantidad = Data2.Recordset!cantidad - ImporteFormateado(Me.txtAux(3).Text)
                            If DiferenciaCantidad > 0 Then
                                'Hemos puesto menos cantidad de la que habiamos puesto
                                cadList = "numlotes =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                                cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                                cadList = DevuelveDesdeBD(conAri, "(canentra-vendida)", "slotes", cadList, "1")
                                If cadList = "" Then cadList = "0"
                                If DiferenciaCantidad > CCur(cadList) Then
                                      MsgBox "Cantidad disponible insuficiente", vbExclamation
                                      b = False
                                End If
                            End If
                        End If
                    End If
                
                
                Else
                    'Insertando linea
                    'No puede poner el mismo articulo, lote en el mismo albaran

                    cadList = "numlotes =" & DBSet(Text2(17).Text, "T") & " AND fechaalb = " & DBSet(Data1.Recordset!FechaAlb, "F")
                    cadList = cadList & " AND codartic = " & DBSet(txtAux(1).Text, "T") & " AND 1"
                    cadList = DevuelveDesdeBD(conAri, "count(*)", "slialp", cadList, "1")
                    If Val(cadList) > 0 Then
                        'NO PUEDE
                        MsgBox "Ya existe el mismo articulo-lote-albaran ", vbExclamation
                        b = False
                    End If
                End If ' manniulador
            End If 'ModificaLineas=2
            
            
            
        End If
    End If
    Set cart = Nothing
    
'    If Me.Text2(17).Locked = False Then
'        If Trim(Text2(17).Text) = "" Then
'            b = False
'            MsgBox "El nº de lote no puede ser nulo." & vbCrLf & vbCrLf & "El artículo tiene control de lotes.", vbExclamation
'        End If
'    End If
    
        
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'campo num_lote y Flecha hacia abajo
        If index = 16 And Text2(17).Locked Then PonerFocoBtn Me.cmdAceptar
        If index = 17 Then PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
       If index = 16 And Text2(17).Locked Then
            PonerFocoBtn Me.cmdAceptar
       ElseIf index = 17 Then
            PonerFocoBtn Me.cmdAceptar
            
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text2_LostFocus(index As Integer)
    'quitamos los espacios en blanco
    Text2(index).Text = Trim(Text2(index).Text)
    
    If index = 16 And (Text2(index).Locked = False) Then Text2(index).Text = UCase(Text2(index).Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
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
            'Enero 2010
            ModificarProveedor
            
        Case 10  'Lineas
            mnLineas_Click
            
        Case 12
            'A hco sin tocar stocks ni smoval ni precios ni leches en vinagre
            EliminarSinStocks
            
        Case 14 'Nº Series
            BotonNSeries
            
        Case 15
            ImpirmirEtiqEsta
        
        Case 16
            'Imprimir SOLO si lleva REA
            Imprimir
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub

Private Sub Imprimir()

  
        'Es la impresion de los albaranes de socios

        CadenaDesdeOtroForm = "|||"
        If Text1(0).Text <> "" Then
            
            'vOY A CARGAR LOS DATOS
            CadenaDesdeOtroForm = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(4).Text & "|" & Text1(5).Text & "|"
        End If
        frmListado2.Opcion = 10
        frmListado2.Show vbModal
        
    
End Sub


Private Sub ImpirmirEtiqEsta()
    If Modo <> 2 Then Exit Sub
    If Me.Data2.Recordset.EOF Then Exit Sub
    
    frmListado.OpcionListado = 513
    frmListado.CadTag = "Albaran: " & Text1(0).Text & " de " & Text1(1).Text & ": " & Text1(4).Text & "-" & Text1(5).Text
    frmListado.NumCod = Data2.RecordSource
    frmListado.Show vbModal
    
    
    
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub

    
Private Function InsertarLinea(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
'OUT -> NumLinea: devuelve el Nº de linea que acaba de insertar
Dim SQL As String
Dim b As Boolean
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim MenError As String
Dim DentroTRANS As Boolean
Dim ImpReciclado As Single
    
    InsertarLinea = False
    SQL = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    SQL = ObtenerWhereCP(False)
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", SQL)
    Me.cmdAux(0).Tag = numlinea
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E", numlinea) Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""
    
    
    If DatosOkLinea() Then 'Lineas de Albaranes Proveedor
         
        
        'Inserta en tabla "slialp"
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
        SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "S") & ", " & DBSet(txtAux(5).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " & DBSet(Text2(17).Text, "T") & ","
        SQL = SQL & DBSet(txtAux(9).Text, "T", "S") 'centro coste
        SQL = SQL & ");"
     Else
        Set vCStock = Nothing
        Exit Function
     End If
    
    If SQL <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        MenError = "Insertando lineas Albaran Compras"
        conn.Execute SQL
        
        
        '==== LAURA 20/09/2006
        'Realizar estas actualizaciones antes de modificar el stock del almacen
        MenError = "Actualizar ult. fecha compra"
        '-- Actualizar en la tabla sartic el ult precio de compra y la ult. fecha compra
        Set vArtic = New CArticulo
        vArtic.Codigo = txtAux(1).Text
        
        
        'Si es negativo NO actualizo PUC
        If CCur(txtAux(3).Text) > 0 Then
            'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
            'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
            b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, CStr(Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4)))
        Else
            b = True
        End If
                
        'Actualizar en la tabla sartic el precio medio ponderado
        If CCur(txtAux(3).Text) <> 0 Then
            MenError = "Actualizar precio medio ponderado"
            'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
            'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text))
            If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4))
            Set vArtic = Nothing
            '====
        End If
        
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If b Then
            MenError = "Actualizando Stocks"
            b = vCStock.ActualizarStock
            
            
            vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
            
        End If
        
        
        If b Then
            'si el articulo tiene control de numero de lotes, insertar en la tabla slotes
            If Me.Text2(17).Locked = False Then
                'si ya existe la linea aumentamos la cantidad entrada
                SQL = " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra"
                SQL = DevuelveDesdeBD(conAri, "canentra", "slotes", SQL, DBSet(Text1(1).Text, "F"), "")
                If SQL = "" Then SQL = "0"
                If CCur(SQL) <> 0 Then
                    
                    If DBSet(txtAux(3).Text, "N") < 0 Then
                        SQL = ""
                    Else
                        SQL = "+"
                    End If
                    SQL = "UPDATE slotes SET canentra=canentra  " & SQL & DBSet(txtAux(3).Text, "N")
                    
                    SQL = SQL & " WHERE " & " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra=" & DBSet(Text1(1).Text, "F")
                    conn.Execute SQL
                Else
                    SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
                    SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(Text2(17).Text, "T") & ", "
                    'fecha entrada, cantidad entrada y cantidad asignada
                    SQL = SQL & DBSet(Text1(1).Text, "F") & "," & DBSet(txtAux(3).Text, "N") & ","
                    
                    If DBSet(txtAux(3).Text, "N") < 0 Then
                        'Es un abono. La cantidad "vendida" es la misma que la "cantidad"
                        ' ya que de este material NO se vende
                        SQL = SQL & "0)"     'SQL = SQL & DBSet(txtAux(3).Text, "N") & ")"
                    Else
                        SQL = SQL & "0)"
                    End If
                    conn.Execute SQL
                End If
            End If
        End If
        
        
        
        
        'Articulo reciclado
        If b Then
          If vParamAplic.ArtReciclado <> "" Then
                
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                     
                    MenError = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                  
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
                    SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & ", " & DBSet(MenError, "T") & ",null, "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
                    SQL = SQL & DBSet(ImpReciclado, "S") & ", 0,0,"
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    ImpReciclado = Round2(ImpReciclado, 2)
                    SQL = SQL & DBSet(ImpReciclado, "N") & ",null,null);"
                    MenError = "Art. reciclado"
                    conn.Execute SQL
                End If
            End If
        End If
        
        
        
    End If
    
    Set vCStock = Nothing
    
EInsertarLinea:
    If Err.Number <> 0 Then b = False
    
    If b Then
        If DentroTRANS Then conn.CommitTrans
        AlmacenLineas = CInt(txtAux(0).Text)
        InsertarLinea = True
    Else
        If DentroTRANS Then conn.RollbackTrans
        InsertarLinea = False
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & MenError & vbCrLf, Err.Description
    End If
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String, vWhere As String
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim b As Boolean
Dim MenError As String
Dim dentroTRANSAC As Boolean
Dim cadNumLote As String
Dim cLote As CNumLote
Dim ImpReciclado As Single


    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    dentroTRANSAC = False
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    
    vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
    Set vArtic = New CArticulo
    If Not vArtic.LeerDatos(txtAux(1).Text) Then Exit Function
    

'    If DatosOkLinea() Then
    'sql para actualizar la linea del albaran compras
    SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
    SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T", "S") & ", "
    SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
    SQL = SQL & "precioar=" & DBSet(txtAux(4).Text, "S") & ", " 'precio
    SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
    SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N") & ", "
    SQL = SQL & "numlotes=" & DBSet(Text2(17).Text, "T", "S") & ","
    SQL = SQL & "codccost=" & DBSet(txtAux(9).Text, "T", "S")
    
    vWhere = ObtenerWhereCP(True) & " AND numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    SQL = SQL & vWhere

    If SQL <> "" Then
        dentroTRANSAC = True
        conn.BeginTrans
            
        MenError = "Actualizando Lineas Albaran Compras"
        conn.Execute SQL
            
            
        '==== Laura 20/09/2006, antes de actualizar el stock
        ' deshacer el precio medio ponderado y luego calcularlo otra vez con los nuevos valores
        MenError = "Recalcular precio medio ponderado del articulo."
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!Cantidad), CCur(Data2.Recordset!precioar))
        b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!cantidad), CCur(Data2.Recordset!ImporteL) / CCur(Data2.Recordset!cantidad))
        
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text), CCur(Data2.Recordset!Cantidad))
        If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4), CCur(Data2.Recordset!cantidad))
        
        'Actualizar ultima fecha de compra del articulo
        If b Then
            'Noacutalizamos si cantidad negativa
            If CCur(txtAux(3).Text) > 0 Then
                MenError = "Actualizando ult. fecha compra"
                '-- Laura 18/12/2006: actualizar precio_ult_compra con el precio aplicandole el descuento, ed. importe/cantidad.
                'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
                b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4))
            End If
        End If
        '====
            
            
        'Actualizar Stocks de los articulos y movimientos
        '===================================================
        If b Then
            MenError = "Actualizando stocks y movimientos almacen"
            'si no se ha modificado el almacen reestablecemos cantidad y precio
            If CInt(Data2.Recordset!codAlmac) = CInt(txtAux(0).Text) Then
'                MenError = "Actualizando Stocks"
                b = vCStock.ModificarStock(Data2.Recordset!cantidad)
            Else
                'deshacer el movimiento para el almacen anterior y devolver stock
                b = InicializarCStock(vCStock, "S") 'movim. de salida
                If b Then b = vCStock.DevolverStock2
                            
                'Insertar el movimiento para el nuevo almacen y actualizar stock
                b = InicializarCStock(vCStock, "E") 'mov. de entrada
                If b Then b = vCStock.ActualizarStock
            End If
        End If
                

        
        '=== CONTROL Nº DE LOTES DEL ARTICULO
        '===============================================
        If b Then
            'comprobar si el artículo que modificamos tiene control de lotes
            MenError = "Actualizando Nº Lote."
            If vArtic.TieneNumLote Then
                    'si no existe en la tabla slotes lo añadimos sino lo modificamos
                    SQL = "SELECT COUNT(*) FROM slotes "
                    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T")
                    SQL = SQL & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                    If RegistrosAListar(SQL) > 0 Then
                        'actualizar la cantidad de entrada de la tabla slotes
                        SQL = "UPDATE slotes SET canentra=" & DBSet(txtAux(3).Text, "N")  '- CSng(Me.Data2.Recordset!cantidad), "N")
                        SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                        conn.Execute SQL
                    ElseIf Text2(17).Text <> "" Then
                        'SI NO EXISTE LO INSERTAMOS
                        SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
                        SQL = SQL & DBSet(Data2.Recordset!codArtic, "T") & "," & DBSet(Text2(17).Text, "T") & "," & DBSet(Data2.Recordset!FechaAlb, "F") & ","
                        SQL = SQL & DBSet(txtAux(3).Text, "N") & ",0)"
                        conn.Execute SQL
                    End If
                                
                    'SI HEMOS MODIFICADO EL Nº DE LOTE
                    'DESCONTAMOS LA CANTIDAD DE LA LINEA DE LA VIEJA
                    'Y SI ES CERO LA BORRAMOS
                    If Text2(17).Text <> CStr(DBLet(Data2.Recordset!numlotes, "T")) Then
                        If Not IsNull(Data2.Recordset!numlotes) Then
                            If DBLet(Data2.Recordset!numlotes, "T") <> "" Then
                                'actualizar la cantidad de entrada de la tabla slotes
                                SQL = "UPDATE slotes SET canentra=canentra - " & DBSet(txtAux(3).Text, "N")
                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                                conn.Execute SQL
                                'borrar si
                                SQL = "DELETE FROM slotes "
                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                                SQL = SQL & " AND canentra=0"
                                conn.Execute SQL
                            End If
                        End If
                    End If
            End If
        End If
                
            
        
        'Articulo reciclado
        If b Then
            If vParamAplic.ArtReciclado <> "" Then
                
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                                           
                     'Si el articulo siguiente es PV entoces lo updatearemos
                     SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " AND numlinea"
                     NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                     SQL = DevuelveDesdeBD(conAri, "codartic", "slialp", SQL, CStr(NumRegElim))
                     'En SQL tengo el codarti de la linea SIGUIENTE
                     'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                     If SQL = vParamAplic.ArtReciclado Then
                     
                          SQL = "UPDATE " & NomTablaLineas & " SET "
                          SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
                          SQL = SQL & "precioar= " & DBSet(ImpReciclado, "N") & ", " 'precio
                          ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                          SQL = SQL & "importel= " & DBSet(ImpReciclado, "N")  'Importe
                          'WHERE
                          SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                          conn.Execute SQL
                    End If  'linea siguiente con codarti=puntoverde
                End If
            End If
        End If
            
            
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        ModificarLinea = b
    End If
        
    
    Set vCStock = Nothing
    Set vArtic = Nothing
    Exit Function
    
EModificarLinea:
    If dentroTRANSAC Then conn.RollbackTrans
    If Not vArtic Is Nothing Then Set vArtic = Nothing
    If Not vCStock Is Nothing Then Set vCStock = Nothing
    ModificarLinea = False
    MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & MenError & vbCrLf & Err.Description
End Function





Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If cmdRegresar.visible Then
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, primeravez
    
    CargaGrid2 vDataGrid, vData
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    If Modo = 2 Then vDataGrid.Enabled = True
    primeravez = False
    
    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte
On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    vDataGrid.Columns(2).visible = False
    vDataGrid.Columns(3).visible = False
    
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            I = 4
            vDataGrid.Columns(I).Caption = "Alm."
            vDataGrid.Columns(I).Width = 500
            vDataGrid.Columns(I).NumberFormat = "000"
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Articulo"
            vDataGrid.Columns(I).Width = 1700
            I = I + 1
            vDataGrid.Columns(I).Caption = "Desc. Artículo"
            vDataGrid.Columns(I).Width = 3400
            
            I = I + 1
            vDataGrid.Columns(I).visible = False
            I = I + 1
            vDataGrid.Columns(I).Caption = "Cantidad"
            vDataGrid.Columns(I).Width = 850
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoImporte
            
            I = I + 1
            vDataGrid.Columns(I).Caption = "Precio"
            vDataGrid.Columns(I).Width = 1140
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoPrecio2
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Dto.1"
            vDataGrid.Columns(I).Width = 550
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoDescuento
            
            I = I + 1
            vDataGrid.Columns(I).Caption = "Dto.2"
            vDataGrid.Columns(I).Width = 550
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Importe"
            vDataGrid.Columns(I).Width = 1080
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoImporte
            
            I = I + 1
            vDataGrid.Columns(I).visible = False 'numlote
            
            I = I + 1
            vDataGrid.Columns(I).Caption = "IVA"
            vDataGrid.Columns(I).Width = 390
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = "# "
            
            I = I + 1
            If vEmpresa.TieneAnalitica Then
                vDataGrid.Columns(I).Caption = "CCoste"
                vDataGrid.Columns(I).Width = 660
            Else
                vDataGrid.Columns(I).visible = False 'codccost
            End If
            vDataGrid.Columns(I + 1).visible = False 'ampliaci
            vDataGrid.Columns(I + 2).visible = False 'numlote
            
            
    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                If I < 3 Then 'campos anteriores a ampliacion linea (ampliaci)
                    txtAux(I).Text = DataGrid1.Columns(I + 4).Text
                '## LAURA 19/06/2008
                ElseIf I < 8 Then
                    txtAux(I).Text = DataGrid1.Columns(I + 5).Text
                Else
                    txtAux(I).Text = DataGrid1.Columns(I + 6).Text
                End If
                '##
                txtAux(I).Locked = False
            Next I
        End If
               
        'El campo Importe es calculado y lo bloqueamos.
        'David. Febrero 2009.  NO bloqueamos el importe. Para que puedan ajustar los valores
        'BloquearTxt txtAux(7), True
    
    
        '#Laura 15/11/2006
        'no se puede modificar el almacen y el articulo pq no elimina bien de smoval
        'y no reestablece stock si se cambia el articulo (REVISAR!!!)
'        BloquearTxt txtAux(0), (ModificaLineas = 2) 'codalmac
        BloquearTxt txtAux(1), (ModificaLineas = 2) 'codartic
'        Me.cmdAux(0).Enabled = (ModificaLineas <> 2)
        Me.cmdAux(1).Enabled = (ModificaLineas <> 2)
        '#
    
    
        '## LAURA 19/06/2008
        '   Añadimos columna de IVA siempre bloqueada
        BloquearTxt txtAux(8), True
        '##
    
        ' ---- [20/10/2009] [LAURA] : añadir centro de coste
        BloquearTxt txtAux(9), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        Me.cmdAux(2).Enabled = Not txtAux(9).Locked
        Me.cmdAux(2).visible = Me.cmdAux(2).Enabled
        ' ----
    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(2).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(2).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(5).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(6).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(8).Width - 10
        'Precio, Dto1, Dto2, Precio
        For I = 4 To 7
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 5).Width - 10
        Next I
        
        '## LAURA 19/06/2008
        txtAux(8).Left = txtAux(7).Left + txtAux(7).Width + 10
        txtAux(8).Width = DataGrid1.Columns(14).Width - 10
        '##
        
        ' ---- [20/10/2009] [LAURA] : añadir el centro de coste
        txtAux(9).Left = txtAux(8).Left + txtAux(8).Width + 10
        txtAux(9).Width = DataGrid1.Columns(15).Width - 10
        cmdAux(2).Left = txtAux(9).Left + txtAux(9).Width - cmdAux(2).Width
        ' ----
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 2
            txtAux(I).visible = visible
        Next I
        txtAux(9).visible = visible And vEmpresa.TieneAnalitica
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub txtAux_GotFocus(index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, index)
    kCampo = index
    ConseguirFocoLin txtAux(index), cadkey
    
    If index = 3 Or index = 4 Or index = 1 Then
        
        If index = 3 Then
            lblF.Caption = "Ver articulo"
        ElseIf index = 4 Then
            lblF.Caption = "Ver precio"
        Else
            lblF.Caption = "EAN"
        End If
    Else
        lblF.Caption = ""
    End If
    
    
End Sub

Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    
    ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
    If KeyCode = 113 Then
        If index = 3 Then AbrirForm_Articulos
    
        If index = 1 Then Me.DataGrid1.Columns(5).Caption = "EAN"
        If index = 4 And txtAux(1).Text <> "" Then
                frmListadoPrecios.Opcion = 0
                frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                frmListadoPrecios.Show vbModal
        End If
    ' ----
    ElseIf KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                If index < 2 Or index = 9 Then  'Para los que tienen busqueda
                    If Modo = 5 And ModificaLineas = 1 Then
                        If txtAux(index).Text = "" Then
                            PulsadoMas2 = True
                            KeyCode = 0
                
                            PulsarTeclaMas False, index
                        End If
                    End If
                End If
           
        
    
    
    ElseIf Not (index = 0 And KeyCode = 38) Then
        KEYdown KeyCode
    End If
    
End Sub


Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(index As Integer)
Dim devuelve As String
Dim TipoDto As Byte
Dim b As Boolean
Dim bLote As Boolean
Dim okArticulo As Boolean

    
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(index).Text = ""
        Exit Sub
    End If
    
    If Not PerderFocoGnralLineas(txtAux(index), ModificaLineas) Then Exit Sub
    
    
    Select Case index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(index).Text)
            txtAux(index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(index)
    
        Case 1 'Cod. ARTICULO
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            
            If Me.DataGrid1.Columns(5).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEan(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , bLote, devuelve)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , bLote, devuelve)
            End If
            If okArticulo Then
                '---- [20/10/2009] [LAURA] : añadir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtAux(9).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then
                    txtAux(9).Text = devuelve
                    Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
                End If
                '----
            
            
                BloquearTxt Text2(17), Not bLote
                
                '## LAURA 19/06/2008
                'obtener el cod. iva del articulo
                txtAux(8).Text = DevuelveDesdeBDNew(conAri, "sartic", "codigiva", "codartic", txtAux(1).Text, "T")
                
                '##
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(index)
            End If
            
        Case 2 'Desc. Articulo
            If txtAux(index).Locked = False Then txtAux(index).Text = UCase(txtAux(index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(index), 1) Then  'Tipo 1: Decimal(12,2)
                If (Modo = 5 And ModificaLineas = 1) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                End If
            End If

        Case 4 'Precio
            PonerFormatoDecimal_Single txtAux(index), 9 'Tipo 9: COnstante
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            If txtAux(index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(index), 1) Then  'Tipo 3: Decimal(12,2)
                    If ModificaLineas = 2 Then
                        'Ponemos el importe que tenia
                        txtAux(index).Text = DataGrid1.Columns(12).Text
                    Else
                        txtAux(index).Text = "0.00"
                    End If
                End If
            End If
            
        Case 9 'COD. CENTRO COSTE
            ' ---- [20/10/2009] [LAURA]: añadir centro de coste a la linea
            If txtAux(index).Text = "" Then
                 txtAux2(index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux2(index).Text = PonerNombreCCoste(Me.txtAux(index))
            End If
    End Select
    
    
     If (index = 3 Or index = 4 Or index = 5 Or index = 6) Then
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
        txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub



Private Sub ObtenerPrecioCompra()
Dim vPrecio As CPreciosCom
Dim cad As String
Dim Aux2 As String
    On Error GoTo EPrecios
    
    Set vPrecio = New CPreciosCom
    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
            txtAux(5).Text = vPrecio.Descuento1
            txtAux(6).Text = vPrecio.Descuento2
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
       
        cad = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If cad <> "" Then txtAux(4).Text = cad
        
            vPrecio.CodigoArtic = txtAux(1).Text
            vPrecio.CodigoProve = Text1(4).Text
            cad = vPrecio.ObtenerDescuentos2(Text1(1).Text, Aux2)
            If cad = "" Then cad = "0"
            txtAux(5).Text = cad
            If Aux2 = "" Then Aux2 = "0"
            txtAux(6).Text = Aux2
            
            'txtAux(5).Text = "0"
            'txtAux(6).Text = "0"

    End If
    PonerFormatoDecimal_Single txtAux(4), 9
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
    
    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        If Data2.Recordset.EOF Then Text2(16).Text = ""
        PonerModo 5
        PonerBotonCabecera True
        AlmacenLineas = -1
        HaModifEnLineas = False
        PonerUltAlmacen
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String
Dim b As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

        conn.BeginTrans
        vWhere = " " & ObtenerWhereCP(False)
                
        
        'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
        'Eliminar los movimientos de smoval
        b = ReestablecerStock(vWhere)
        

        If b Then
                        'Hay que eliminar los lotes
            If vParamAplic.ManipuladorFitosanitarios2 Then EliminarEnSlotes
                    
        End If
        
        If b Then
            'Pasar los datos al historico de albaranes de compra y borrarlos de albaranes
            'scaalp --> schalp
            'slialp --> slhalp
            b = ActualizarElTraspaso("", vWhere, CodTipoMov, cadList)
            
            'Eliminar los numeros de serie del albaran sino estan vendidos a ningun cliente
            If b Then
                SQL = "DELETE FROM sserie WHERE numalbpr=" & DBSet(Data1.Recordset!Numalbar, "T")
                SQL = SQL & " AND fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F")
                SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
                SQL = SQL & " AND (isnull(numfactu) and isnull(numalbar))"
                conn.Execute SQL
            End If
            
        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Albaran Compras", Err.Description
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
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Dim Indicador As String
'Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         'vWhere = "(" & ObtenerWhereCP(False) & ")"
         
         'SEPT 2010
         'NO parece ir muy fino
         'Voy a cambiar
         Data1.Refresh
         
         'If SituarDataMULTI(Data1, vWhere, Indicador) Then
         If SituarData Then
             PonerModo 2
            
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
Dim SQL As String
On Error Resume Next
    
    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & NombreTabla & ".numalbar= " & DBSet(Text1(0).Text, "T") & " and " & NombreTabla & ".fechaalb='" & Format(Text1(1).Text, FormatoFecha)
    SQL = SQL & "' and " & NombreTabla & ".codprove=" & Val(Text1(4).Text)
    
    ObtenerWhereCP = SQL
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
Dim SQL As String
    
    SQL = "SELECT numalbar,fechaalb," & NomTablaLineas & ".codprove, numlinea, codalmac, " & NomTablaLineas & ".codartic,"
    SQL = SQL & NomTablaLineas & ".nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, numlotes,codigiva,codccost "
    
    'Octubre 2015
    'Hay que cargar todos los datos del "foragrid" de abajo. Si no esta estamos leyendo DOS veces
    'peeeero, no se muestran en el grid
    SQL = SQL & ",ampliaci, numlotes "
    
    SQL = SQL & " FROM " & NomTablaLineas
    'Para que tanto el hco como el nomal apunte a slialp
    
    SQL = SQL & " inner join sartic on " & NomTablaLineas & ".codartic = sartic.codartic"
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
    Else
        SQL = SQL & " WHERE numalbar = -1"
    End If
    SQL = SQL & " Order by numalbar, fechaalb, codprove, numlinea"
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
   
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        'Toolbar1.Buttons(7).Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
        'Me.mnEliminar.Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
        
        'No permito borrar
        If b Then
            'Si modo=2 NO dejare que borre
            If Modo = 2 And cadSelAlbaranes <> "" Then b = False
        End If
        Toolbar1.Buttons(7).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = Toolbar1.Buttons(7).Enabled
            
        b = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        
        Toolbar1.Buttons(10).Enabled = (Modo = 2)
        Me.mnLineas.Enabled = (Modo = 2)
        Toolbar1.Buttons(9).Enabled = b
        Toolbar1.Buttons(14).Enabled = b
        
        Toolbar1.Buttons(12).Enabled = b
        
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = (Not b)
        Me.mnBuscar.Enabled = (Not b)
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = (Not b)
        Me.mnVerTodos.Enabled = (Not b)
End Sub



Private Function InsertarAlbaran(vSQL As String) As Boolean
Dim MenError As String
Dim devuelve As String
Dim bol As Boolean

    On Error GoTo EInsertarOferta
    
    bol = False
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    
    'Actualizar el campo fecha ult.compra(fechamov) en la tabla proveedores (sprove)
    devuelve = DevuelveDesdeBDNew(conAri, "sprove", "fechamov", "codprove", Text1(4).Text, "N")
    If (devuelve = "") Then devuelve = "01/01/1900"
    If CDate(Text1(1).Text) > CDate(devuelve) Then
        vSQL = "UPDATE sprove SET fechamov=" & DBSet(Text1(1).Text, "F")
        vSQL = vSQL & " WHERE codprove=" & Text1(4).Text
        conn.Execute vSQL, , adCmdText
    End If
    bol = True
    
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarAlbaran = True
    Else
        conn.RollbackTrans
        InsertarAlbaran = False
    End If
End Function


Private Sub LimpiarDatosProve()
Dim I As Byte

    For I = 4 To 14
        Text1(I).Text = ""
    Next I
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim cLote As CNumLote
Dim cart As CArticulo
Dim SQL As String
Dim b As Boolean
Dim Aux As String
Dim ImpReciclado As Single

    EliminarLinea = False
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
    
    '==== Laura: 20/09/2006
    'Inicializar la clase para actualizar precio medio ponderado del Articulo
    Set cart = New CArticulo
    If Not cart.LeerDatos(vCStock.codArtic) Then Exit Function
    '====
    
    On Error GoTo EEliminarLinea
    conn.BeginTrans
    
    'Eliminar las lineas de la tabla "sserie", insertadas para la linea del albaran a eliminar
    SQL = " WHERE  numalbpr= " & DBSet(Text1(0).Text, "T") & " and fechacom='" & Format(Text1(1).Text, FormatoFecha)
    SQL = SQL & "' and codprove=" & Val(Text1(4).Text) & " AND numline2=" & Data2.Recordset!numlinea
    conn.Execute "Delete from sserie " & SQL
    
        
    'Si tiene tasa recilcado y la siguiente linea es la de reciclado me la cargo
        'Tasa recilcaje
    If vParamAplic.ArtReciclado <> "" Then
        If ArticuloConTasaReciclado(CStr(Data2.Recordset!codArtic), ImpReciclado) Then

                SQL = ObtenerWhereCP(False) & " and numlinea>" & Data2.Recordset!numlinea & " AND 1 "
                SQL = Replace(SQL, NombreTabla, NomTablaLineas)
                'Vere si la siguiente linea es de tasa de reciclado me la cargo tb
                Aux = "numlinea"
                SQL = DevuelveDesdeBD(conAri, "codartic", NomTablaLineas, SQL, "1", "N", Aux)

                If SQL = vParamAplic.ArtReciclado Then
                    'la borro tb
                    SQL = ObtenerWhereCP(True) & " and numlinea=" & Aux
                    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
                    SQL = "DELETE FROM " & NomTablaLineas & SQL
                    conn.Execute SQL
                End If



        End If
    End If
    
    
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = ObtenerWhereCP(True) & " and numlinea=" & Data2.Recordset!numlinea
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    SQL = "Delete from " & NomTablaLineas & SQL
    conn.Execute SQL 'Eliminar linea
    
    '==== Laura: 20/09/2006
    'reestablecer el precio medio ponderado,
    'debe calcularse antes de reestablecer el stock
    '-- Laura 19/12/2006: calcular el precio medio ponderado con precio con los descuentos ( importe/cantidad)
    'cArt.ReestablecerPrecioMedPon vCStock.Cantidad, CCur(Data2.Recordset!precioar)
    If CCur(Data2.Recordset!cantidad) <> 0 Then cart.ReestablecerPrecioMedPon vCStock.cantidad, Round2(CCur(Data2.Recordset!ImporteL) / CCur(Data2.Recordset!cantidad), 4)
    Set cart = Nothing
    '====
    
    b = vCStock.DevolverStock2
    Set vCStock = Nothing
    
    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
    'si la linea se queda con cero borrarla.
    If b Then
        If Not IsNull(Data2.Recordset!numlotes) Then
            Set cLote = New CNumLote
            If cLote.LeerDatos(CStr(Data2.Recordset!codArtic), CStr(Data2.Recordset!numlotes), CStr(Data2.Recordset!FechaAlb)) Then
                b = cLote.Eliminar(CSng(Data2.Recordset!cantidad))
            
            End If
            Set cLote = Nothing
        End If
    End If
    

EEliminarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        EliminarLinea = True
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = CodTipoMov '"ALC=Albaran de Compra"
    vCStock.FechaMov = Text1(1).Text
    vCStock.Trabajador = CLng(Text1(4).Text) 'En smoval guardamos el Proveedor
    vCStock.Documento = Text1(0).Text
    
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "E") Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(7).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ReestablecerStock(cadSel As String) As Boolean
Dim vCStock As CStock
Dim cart As CArticulo
Dim cLote As CNumLote
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ERestablecer
    
    SQL = "SELECT * FROM " & NomTablaLineas & " WHERE " & Replace(cadSel, NombreTabla, NomTablaLineas)
    SQL = SQL & " ORDER BY numlinea desc "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    While (Not RS.EOF) And b
        'para cada linea de albaran deshacemos movimientos y precios medios ponderados
        Set vCStock = New CStock
           If InicializarCStock(vCStock, "S", RS!numlinea) Then
                'estos valores hay q leerlos del RS y no del data2
                 vCStock.codArtic = RS!codArtic
                 vCStock.codAlmac = CInt(RS!codAlmac)
                 vCStock.cantidad = CSng(RS!cantidad)
                 vCStock.Importe = CCur(RS!ImporteL)
                 vCStock.LineaDocu = RS!numlinea
           
                '==== Laura 20/09/2006
                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
                Set cart = New CArticulo
                If cart.LeerDatos(vCStock.codArtic) Then
                    'Laura 19/12/2006: Calcular precio medio pond. con precio con los descuentos (importe/cantidad)
                    'If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                    If Not cart.ReestablecerPrecioMedPon(CCur(vCStock.cantidad), Round2(vCStock.Importe / vCStock.cantidad, 4)) Then b = False
                    
                    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
                    'si la linea se queda con cero borrarla.
                    If b Then
                        If cart.TieneNumLote Then
                            Set cLote = New CNumLote
                            If cLote.LeerDatos(cart.Codigo, CStr(DBLet(RS!numlotes, "T")), CStr(RS!FechaAlb)) Then
                                b = cLote.Eliminar(vCStock.cantidad)
                            
                            End If
                            Set cLote = Nothing
                        End If
                    End If
                End If
                Set cart = Nothing
                '====
                
                
                'Actualiza el stock en salmac y borra de smoval
                'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
                'en Almacen ahora lo tiene que borrar(S).
                If b Then
                    If Not vCStock.DevolverStock2() Then b = False
                End If
           Else
               b = False
           End If
'           Data2.Recordset.MoveNext
           Set vCStock = Nothing
    
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    '#### ANTES DEL 20/09/2006
    
'    b = True
    
'    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
'       'en Almacen ahora lo tiene que borrar(S).
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'                '==== Laura 20/09/2006
'                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
'                Set cArt = New CArticulo
'                If cArt.LeerDatos(vCStock.codArtic) Then
'                    If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(Data2.Recordset!precioar)) Then b = False
'                End If
'                Set cArt = Nothing
'
'               'Actualiza el stock en salmac y borra de smoval
'                If b Then
'                    If Not vCStock.DevolverStock() Then b = False
'                End If
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If

ERestablecer:
    If Err.Number <> 0 Then b = False
    If Not b Then
        ReestablecerStock = False
        MuestraError Err.Number, "Reestablecer stock.", Err.Description
    Else
        ReestablecerStock = True
    End If
End Function




Private Function ReestablecerUltFecCompra() As Boolean
Dim cart As CArticulo
Dim SQL As String
Dim b As Boolean

    On Error GoTo ERestCompra
    
    b = True
    
'    select distinct codartic from slialp
'where numalbar=2100045 and fechaalb='2006-09-15' and codprove=21
    
    
'    If Not Data2.Recordset.EOF Then
'       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada articulo del albaran reestablecer la fecha ultima compra
'       'y el precio ultima compra
'
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'               'Actualiza el stock en salmac y borra de smoval
'               If Not vCStock.DevolverStock() Then b = False
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If
    
    
    
    
    ReestablecerUltFecCompra = b
    
    
ERestCompra:
    
'    If Not b Then
        ReestablecerUltFecCompra = False
'    Else
'        ReestablecerUltFecCompra = True
'    End If
End Function





'Private Function ReestablecerPrecioMedPon() As Boolean
''reestablecer el valor del precio medio ponderado
''Dim vCStock As CStock
'Dim b As Boolean
'
'    On Error GoTo EResPMP
'
'    b = True
''    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
''       Data2.Recordset.MoveFirst
''
''       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
''       'en Almacen ahora lo tiene que borrar(S).
''       While (Not Data2.Recordset.EOF) And b
''           Set vCStock = New CStock
''           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
''               'Actualiza el stock en salmac y borra de smoval
''               If Not vCStock.DevolverStock() Then b = False
''           Else
''               b = False
''           End If
''           Data2.Recordset.MoveNext
''           Set vCStock = Nothing
''       Wend
''    End If
'    ReestablecerPrecioMedPon = b
'    Exit Function
'
'EResPMP:
''    If Not b Then
'        ReestablecerPrecioMedPon = False
''    Else
''        ReestablecerPrecioMedPon = True
''    End If
'End Function



Private Sub InsertarCabecera()
Dim SQL As String

    SQL = CadenaInsertarDesdeForm(Me)
    If SQL <> "" Then
        If InsertarAlbaran(SQL) Then
'                            PosicionarData
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas 1, "Albaranes"
            BotonAnyadirLinea
        End If
    End If
    Me.SSTab1.Tab = 0
End Sub


Private Sub BotonNSeries()
Dim cadWhere As String, SQL As String
Dim RSLineas As ADODB.Recordset

    ModificaLineas = 4

    cadWhere = " WHERE numalbar=" & DBSet(Text1(0).Text, "T")
    cadWhere = cadWhere & " and fechaalb=" & DBSet(Text1(1).Text, "F")
    cadWhere = cadWhere & " and slialp.codprove=" & Text1(4).Text
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    SQL = "SELECT numlinea, slialp.codartic, sum(cantidad) as cantidad "
    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    SQL = SQL & " GROUP BY numlinea,codartic ORDER BY Codartic "

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Abre el formulario de pedir nº serie al comprarlos
        'pero mostrando los nº de serie ya asignados para poder modificarlos
        PedirNSeries RSLineas
    Else
        MsgBox "No hay ninguna linea de Articulo con Control de Nº Serie", vbInformation
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    ModificaLineas = 0
    DescargarDatosTMPNumSeries ("tmpnseries")
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim SQL As String
Dim linea As Integer

    On Error GoTo EPedirNSeries

        'Inicializo la tabla temporal de los num.serie
        PedirNSeriesGnral RS, False
        
        RS.MoveFirst
        While Not RS.EOF
            linea = 0
            'Cargar los Nº de serie asignados al albaran
            SQL = "SELECT numserie, codartic FROM sserie "
            SQL = SQL & " WHERE numalbpr=" & DBSet(Text1(0).Text, "T")
            SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and codprove=" & Text1(4).Text & " and numline2=" & RS!numlinea
            SQL = SQL & " ORDER BY codartic "
            
            Set RSseries = New ADODB.Recordset
            RSseries.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " and codartic=" & DBSet(RS!codArtic, "T")
                SQL = SQL & " and numlinealb=" & RS!numlinea
                SQL = SQL & " and numlinea=" & linea
                conn.Execute SQL
                RSseries.MoveNext
            Wend
            RS.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
        
        SQL = "select count(*) from tmpnseries Where codusu=" & vUsu.Codigo
        If RegistrosAListar(SQL) > 0 Then
            Set frmNSerie = New frmRepCargarNSerie
            frmNSerie.DeVentas = False 'Se llama desde Alb. de compras
            frmNSerie.NumAlb = Text1(0).Text
            frmNSerie.Show vbModal
            Set frmNSerie = Nothing
            Espera 0.2
        Else
            MsgBox "No hay nº de serie asignados a ese albaran", vbInformation
        End If
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal o actualizalo
Dim SQL As String
Dim b As Boolean

    On Error GoTo ECargar
    conn.BeginTrans
    
    'Borrar todos los Nº de Serie asignados a ese albaran de compra
    'y que no tienen asignado ya un albaran de venta
    SQL = "DELETE FROM sserie "
    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
    SQL = SQL & " and (isnull(numalbar) and isnull(numfactu))"
    conn.Execute SQL
    
    'Si algun Nº serie tenia asignado albaran venta y no lo pude borrar entonces limpiamos
    'los campos del albaran de compra
    SQL = "UPDATE sserie SET codprove=" & ValorNulo & ", numalbpr=" & ValorNulo & ", fechacom="
    SQL = SQL & ValorNulo & ", numline2=" & ValorNulo
    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
    conn.Execute SQL
    
    b = InsertarNumSeriesDeTMP
    
   
ECargar:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    
End Sub


Private Function InsertarNumSeriesDeTMP() As Boolean
'Inserta en la tabla sserie todos los nº de serie q se han cargado en la temporal
Dim SQL As String
Dim Numalbar As String
Dim b As Boolean
Dim RStmp As ADODB.Recordset
Dim nSerie As CNumSerie

    On Error GoTo EInsertarNSeries

    'Inicializamos el objeto nº de serie con los valores comunes a todos
    Set nSerie = New CNumSerie
    nSerie.Proveedor = CInt(Text1(4).Text)
    nSerie.NumAlbProve = Text1(0).Text
    nSerie.fechacom = Text1(1).Text
    
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    SQL = SQL & " ORDER BY codartic, numlinealb, numlinea "
    
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
    b = True
    While Not RStmp.EOF And b
        nSerie.numSerie = RStmp!numSerie
        nSerie.Articulo = RStmp!codArtic
        nSerie.NumLinAlbPr = RStmp!numlinealb
        
        'obtenemos los dias de garantia del articulo
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", RStmp!codArtic, "T")
        'fin garantia= fecha albaran + dias de garantia
        nSerie.FinGarantia = CStr(CDate(Text1(1).Text) + CInt(ComprobarCero(SQL)))
    
        'Comprobar si existe en la tabla sserie ese nº de serie
        Numalbar = "numalbpr" 'Nº albaran de Venta prove
        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", RStmp!numSerie, "T", Numalbar, "codartic", RStmp!codArtic, "T")
        If SQL <> "" Then
            If Numalbar = "" Then 'ya existe el nº serie y actualizamos ya que no esta asignado a ningun albaran
                b = nSerie.ActualizarNumSerie(False)
            End If
        Else
            b = nSerie.InsertarNumSerie
        End If
        
'        b = InsertarNSerie(RStmp!NumSerie, RStmp!codArtic, RStmp!NumLinealb)
        RStmp.MoveNext
    Wend
    RStmp.Close
    Set RStmp = Nothing
    
    Set nSerie = Nothing
    
EInsertarNSeries:
    If Err.Number <> 0 Then b = False
    If Not b Then
        InsertarNumSeriesDeTMP = False
    Else
        InsertarNumSeriesDeTMP = True
    End If
End Function




Private Sub PonerDatosProveedor(Codprove As String, Optional nifProve As String)
Dim vProve As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If Codprove = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(Codprove) Then
        If vProve.LeerDatos(Codprove) Then
            'NUEVO. Situacion proveedor
            If vProve.ProveedorBloqueado Then
                LimpiarDatosProve
                Set vProve = Nothing
                PonerFoco Text1(4)
                Exit Sub
            End If
            
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!Codprove) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(4).Text = vProve.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vProve.Nombre  'Nom prove
                Text1(8).Text = vProve.Domicilio
                Text1(9).Text = vProve.CPostal
                Text1(10).Text = vProve.Poblacion
                Text1(11).Text = vProve.Provincia
                Text1(6).Text = vProve.NIF
                Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(12).Text = vProve.ForPago
                Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
                Text1(13).Text = Format(vProve.DtoPPago, FormatoDescuento)
                Text1(14).Text = Format(vProve.DtoGnral, FormatoDescuento)
            End If


            'Solo insertando
            If Modo = 3 Then vProve.PonerObservaciones Text1(15), Text1(16), Text1(17), Text1(18), Text1(19)


            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
        If Modo = 3 Then PonerFoco Text1(4)
        
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    If b Then
        Text1(5).Text = vProve.Nombre   'Nom proveedor
        Text1(8).Text = vProve.Domicilio
        Text1(9).Text = vProve.CPostal
        Text1(10).Text = vProve.Poblacion
        Text1(11).Text = vProve.Provincia
        Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub



Private Sub BloquearDatosProve(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(5).visible = bol 'NIF
        Me.imgBuscar(5).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
Dim vProve As CProveedor

    On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(5).Text
        vProve.Domicilio = Text1(8).Text
        vProve.CPostal = Text1(9).Text
        vProve.Poblacion = Text1(10).Text
        vProve.Provincia = Text1(11).Text
        vProve.TfnoAdmon = Text1(7).Text
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function



Private Sub CalcularDatosFactura()
Dim I As Byte
Dim cadWhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 33 To 50
         Text3(I).Text = ""
    Next I
    
    cadWhere = ObtenerWhereCP(False)
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(13).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(14).Text))
    vFactu.FijarTipoIvaProveedor Val(Text1(4).Text)
    If vFactu.CalcularDatosFactura2(cadWhere, NombreTabla, NomTablaLineas, CDate(Text1(1).Text), False) Then
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
        Text3(49).Text = vFactu.TotalFac
        Text3(50).Text = vFactu.BaseImp
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 33 To 36
        If I = 34 Or I = 35 Then Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    
    'Desglose B.Imponible por IVA
    For I = 43 To 45
        If Text3(I).Text <> "" Then
             If CSng(Text3(I).Text) = 0 And Text3(I - 6).Text = "" Then
                Text3(I).Text = QuitarCero(Text3(I).Text)
                Text3(I - 3).Text = QuitarCero(Text3(I - 3).Text)
                Text3(I - 6).Text = QuitarCero(Text3(I - 6).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I + 3).Text)
            Else
                Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
                Text3(I - 3) = Format(Text3(I - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(I + 3).Text = Format(Text3(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
    
    'Formatear el total de Factura
    Text3(49).Text = Format(Text3(49).Text, FormatoImporte)
    Text3(50).Text = Format(Text3(50).Text, FormatoImporte)
End Sub




Private Sub ComprobarNumSeries(numlinea As String)
'Comprobamos para una linea de Albaran si el articulo tiene control de nº de serie
'y procedemos
Dim SQL As String
Dim cadW As String
Dim RSLineas As ADODB.Recordset
'Dim Mostrar As Boolean 'Indica si vamos a pedir num series o a mostrarlos
'Dim cant As Integer 'cantidad que vamos a insertar


    'si la cantidad es >0 pedimos nº serie articulos comprados
    'si la cantidad es <0 mostramos los nº serie para devolver (ABONOS)
                    
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If SQL = "1" Then 'Si el Articulo tiene control de nº de serie
'        If Modo = 5 Then
'            If ModificaLineas = 1 Then 'INSERTAR linea
'                If CCur(txtAux(3).Text) > 0 Then 'cantidad linea
'                    Mostrar = False
'                    cant = CSng(txtAux(3).Text)
'                ElseIf CCur(txtAux(3).Text) < 0 Then 'cantidad linea
'                    'Es un ABONO
'                    'cantidad es < 0 (es un abono, devolvemos el articulo comprado)
'                    Mostrar = True
'                End If
'
'            ElseIf ModificaLineas = 2 Then 'MODIFICAR linea
'                'comprobar que la cantidad introducida se ha modificado
'                If CSng(txtAux(3).Text) <> CSng(Data2.Recordset!Cantidad) Then
'                    cant = CSng(txtAux(3).Text) - CSng(Data2.Recordset!Cantidad)
'                    If cant > 0 Then 'añadir nuevos num serie
'                        Mostrar = False
'                    ElseIf cant < 0 Then 'mostrar num serie y quitar el que toca
'                        Mostrar = True
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End If
'        End If
        
        
        
        If CCur(txtAux(3).Text) > 0 Then 'cantidad
'        If Mostrar = False Then
                SQL = "El Articulo tiene control de Nº de Serie." & vbCrLf & vbCrLf
                SQL = SQL & "Introduzca los Nº de Serie"
                If ModificaLineas = 2 Then
                    SQL = SQL & " que se han añadido"
                End If
                MsgBox SQL & "." & vbCrLf, vbInformation
                'Cargar la tabla temporal con tantas filas como cantidad de Articulo
                'Para introducir el Nº de Serie
                DescargarDatosTMPNumSeries "tmpnseries"
                CargarDatosTMPNumSeries "tmpnseries", txtAux(1).Text, CInt(txtAux(3).Text), numlinea
                'Visualizar en pantalla el Grid, y rellenar los Nº Serie
                ModificaLineas = 0
                Set frmNSerie = New frmRepCargarNSerie
                frmNSerie.DeVentas = False
                frmNSerie.NumAlb = ""
                frmNSerie.Show vbModal
                Set frmNSerie = Nothing
                
        Else   'cantidad es < 0 (es un ABONO, devolvemos el articulo comprado)
           
            'Comprobar que efectivamente estan en tabla sserie los NºSerie del Articulo
            ' y que no esten asignados ya a otro albaran de venta
            SQL = " select distinct count(numserie) from sserie "
            cadW = " WHERE codartic=" & DBSet(txtAux(1).Text, "T")
            cadW = cadW & " and codprove=" & Text1(4).Text
            cadW = cadW & " and (numalbar='' or isnull(numalbar))"
            SQL = SQL & cadW
            
            If RegistrosAListar(SQL) > 0 Then 'Hay Nº de Serie para elegir
                'mostrar los nº de serie de ese proveedor que no esten vendidos y selecccionar
                'el que vamos a devolver
                'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
                SQL = "SELECT codartic, sum(cantidad) as cantidad, numlinea "
                SQL = SQL & " FROM " & NomTablaLineas
                
                cadW = " WHERE numalbar=" & DBSet(Text1(0).Text, "T") & " and "
                cadW = cadW & " fechaalb=" & DBSet(Text1(1).Text, "F")
                cadW = cadW & " and codprove= " & Text1(4).Text & " and numlinea=" & numlinea

                SQL = SQL & cadW
                SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

                Set RSLineas = New ADODB.Recordset
                RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

                MostrarNSeries RSLineas
                RSLineas.Close
                Set RSLineas = Nothing
            End If
        End If
    End If
End Sub



Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionamos
'los que vamos a devolver (Para ABONOS)
Dim SQL As String
Dim Campos As String
On Error GoTo EMostrarNSeries

    SQL = MostrarNSeriesGnral(RSLineas, Campos)
    SQL = SQL & " and sserie.codprove=" & Text1(4).Text
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = SQL
    frmMen.OpcionMensaje = 4 'Nº Series Articulo
    frmMen.vCampos = Campos
    frmMen.Show vbModal
    Set frmMen = Nothing
    
EMostrarNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificaAlb

    conn.BeginTrans

    MenError = "Modificando fecha de albaran en tablas relacionadas."
    b = ComprobarCambioFecha
                
    If b Then
        MenError = "Modificando el albaran (scaalb)."
        b = ModificaDesdeFormulario(Me, 1)
        
        If b Then
            'Actualizar los datos del Proveedor si es de varios
            MenError = "Actualizando proveedor de varios."
            b = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
        End If
    End If

EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MsgBox "Error Modificando el albaran." & vbCrLf & MenError, vbExclamation
    End If
    ModificarCabAlbaran = b
    Espera 0.2
End Function




'Private Function ArticuloTieneMargen() As Boolean
'Dim cad As String
'
'    'Comprobar que el artículo tiene margen comercial
'    cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", txtAux(1).Text, "T")
'    If cad = "" Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El artículo no tiene margen comercial para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'
'    'comprobar que las tarifas tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(txtAux(1).Text, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El artículo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'    ArticuloTieneMargen = True
'
'End Function


Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    cmdAux(2).Tag = "2"

    Set frmB = New frmBuscaGrid
    If vParamAplic.ContabilidadNueva Then
        frmB.vCampos = "Codigo|ccoste|codccost|T||20·Descripción|ccoste|nomccost|T||70·"
        frmB.vTabla = "ccoste"
    Else
        frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
        frmB.vTabla = "cabccost"
    End If
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Centros de coste"
    frmB.vselElem = 0
    frmB.vConexionGrid = conConta
    
    frmB.Show vbModal
    Set frmB = Nothing
    cmdAux(2).Tag = "-1"
End Sub


' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
Private Sub AbrirForm_Articulos()
    If Trim(txtAux(1).Text) = "" Then Exit Sub
    
'    Set frmArt = New frmAlmArticulos
'    frmArt.DeConsulta = True
'    frmArt.DatosADevolverBusqueda3 = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
'    frmArt.parNumTAb = 6
'    frmArt.Show vbModal
'    Set frmArt = Nothing
    
    
    frmAlmArticulos.DeConsulta = True
    frmAlmArticulos.DatosADevolverBusqueda = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
    frmAlmArticulos.parNumTAb = 6
    frmAlmArticulos.Show vbModal
    Set frmAlmArticulos = Nothing
    
    
End Sub
' -----


Private Sub ModificarProveedor()
Dim OK As Boolean
    OK = True
    If EsHistorico Then
        OK = False
    Else
            If Modo = 2 Then
                If Data1.Recordset Is Nothing Then
                    OK = False
                Else
                    If Data1.Recordset.EOF Then
                       OK = False
                    Else
                        'data1.Recordset!esdevarios
                        If EsDeVarios Then
                            MsgBox "Proveedor de VARIOS", vbExclamation
                            OK = False
                        End If
                    End If
                End If
            Else
                OK = False
            End If
    End If
    If OK Then
        If vUsu.Nivel > 1 Then
            MsgBox "usuario sin permiso", vbExclamation
            OK = False
        End If
    End If
    If Not OK Then Exit Sub
    

    
    CadenaDesdeOtroForm = ""
    frmListado2.Opcion = 18
    frmListado2.Show vbModal
    If CadenaDesdeOtroForm = "" Then Exit Sub
    'Si es el mismo no hago nada
    If (CLng(CadenaDesdeOtroForm)) = CLng(Text1(4).Text) Then
        MsgBox "Mismo proveedor", vbExclamation
        Exit Sub
    End If
    

        
        
     Screen.MousePointer = vbHourglass
    'pedimos el nuevo proveedor
    Set miRsAux = New ADODB.Recordset
    conn.BeginTrans
    OK = HacerUpdatesCodProve(CLng(CadenaDesdeOtroForm))
    If OK Then
        'Situamos
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    If OK Then
        'Sitauremos
        CadenaDesdeOtroForm = " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & CadenaDesdeOtroForm
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaDesdeOtroForm & " " & Ordenacion
        PonerCadenaBusqueda
        
    End If
    Screen.MousePointer = vbDefault
    
    
    
End Sub




Private Function HacerUpdatesCodProve(NuevoProve As Long) As Boolean
Dim CadenaLineas As String
Dim J As Integer
Dim SQL As String
Dim vPr As CProveedor
        
        On Error GoTo EHacerUpdatesCodProve
        HacerUpdatesCodProve = False
        
        Set vPr = New CProveedor
        If Not vPr.LeerDatos(CStr(NuevoProve)) Then
            Set vPr = Nothing
            Exit Function
        End If
        
            
        
        
        SQL = "Select "
        SQL = SQL & "fechaalb,numlotes,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
        SQL = SQL & "dtoline1,dtoline2,importel,codprove,codccost"
        SQL = SQL & " FROM slialp"
        SQL = SQL & " WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CadenaLineas = ""
        While Not miRsAux.EOF
            SQL = ", ('" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
            'Texto
            For J = 1 To 5
                If IsNull(miRsAux.Fields(J)) Then
                    SQL = SQL & ",NULL"
                Else
                    SQL = SQL & ",'" & DevNombreSQL(miRsAux.Fields(J)) & "'"
                End If
            Next J
            'Numero
                        'Texto
            For J = 6 To 12
                If IsNull(miRsAux.Fields(J)) Then
                    SQL = SQL & ",NULL"
                Else
                    SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux.Fields(J)))
                End If
            Next J
            'Nuevo proveedor
            SQL = SQL & "," & NuevoProve
            
            'Abril 2011.  codccost
            SQL = SQL & "," & DBSet(miRsAux!CodCCost, "T", "S")
            
            CadenaLineas = CadenaLineas & SQL & ")"

            'Sig
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las lineas
        If CadenaLineas <> "" Then
                
            SQL = "DELETE FROM slialp WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
            conn.Execute SQL
                    
            SQL = "INSERT INTO slialp ("
            SQL = SQL & "fechaalb,numlotes,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
            SQL = SQL & "dtoline1,dtoline2,importel,codprove,codccost"
            'Quito la primara coma
            CadenaLineas = Mid(CadenaLineas, 2)
            SQL = SQL & ") VALUES " & CadenaLineas
            CadenaLineas = SQL
        
        End If
        
        'ACtualizamos
        'Busco los datos del proveedor
        SQL = "UPDATE scaalp SET codprove = " & NuevoProve
        'Resto de datos del proveedor: nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove
        SQL = SQL & ", nomprove=" & DBSet(vPr.Nombre, "T")
        SQL = SQL & ",domprove=" & DBSet(vPr.Domicilio, "T")
        SQL = SQL & ",codpobla=" & DBSet(vPr.CPostal, "T")
        SQL = SQL & ",pobprove=" & DBSet(vPr.Poblacion, "T")
        SQL = SQL & ",proprove=" & DBSet(vPr.Provincia, "T")
        SQL = SQL & ",nifprove=" & DBSet(vPr.NIF, "T")
        SQL = SQL & ",telprove=" & DBSet(vPr.TfnoAdmon, "T", "S")
        SQL = SQL & " WHERE"
        SQL = SQL & " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        conn.Execute SQL
        Set vPr = Nothing
        If CadenaLineas <> "" Then
                'meto las lineas con el nuevo proveedor
                conn.Execute CadenaLineas
                
                'UPDATEO las tablas de smoval
                SQL = "UPDATE smoval SET codigope = " & NuevoProve
                SQL = SQL & " WHERE detamovi='ALC' AND "
                SQL = SQL & " document = " & DBSet(Text1(0).Text, "T") & " AND fechamov = " & DBSet(Text1(1).Text, "F") & " AND codigope = " & Text1(4).Text
                conn.Execute SQL
                
                
        End If
        
        
        HacerUpdatesCodProve = True
        Exit Function
EHacerUpdatesCodProve:
    MuestraError Err.Number, Err.Description
End Function


Private Sub PonerUltAlmacen()
Dim C As String
    If AlmacenLineas < 0 Then
       If Not Data2.Recordset.EOF Then
            C = ObtenerWhereCP(True)
            C = Replace(C, NombreTabla, NomTablaLineas)
            AlmacenLineas = DevuelveUltimoAlmacen(NomTablaLineas, C)
       End If
            
       If AlmacenLineas < 0 Then
            'No hay datos todavia
            '                                                                trabajador
            C = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(2).Text, "N")
            If C <> "" Then AlmacenLineas = Val(C)
        End If
    End If
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, index As Integer)

    If InsertandoCabecera Then
        If imgBuscar(index).visible Then imgBuscar_Click index
        
    Else
        'Lineas
        If index = 9 Then index = 2
        cmdAux_Click index
        
        
    End If
        
End Sub


Private Function SituarData() As Boolean
    On Error GoTo ES
    SituarData = False
    '
     'DBSet(Text1(0).Text, "T") & " and " & NombreTabla & ".fechaalb='" & Format(Text1(1).Text, FormatoFecha)
    'SQL = SQL & "' and " & NombreTabla & ".codprove=" & Val(Text1(4).Text)
    
    While Not Data1.Recordset.EOF
        If Val(Data1.Recordset!Codprove) = Val(Text1(4).Text) Then
            If Data1.Recordset!Numalbar = Text1(0).Text Then
                If Format(Data1.Recordset!FechaAlb, "dd/mm/yyyy") = Text1(1).Text Then
                    Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                    SituarData = True
                    Exit Function
                End If
            End If
        End If
        Data1.Recordset.MoveNext
    Wend
    
    'Si llega aqui no lo ha encontrado
    Exit Function
ES:
    MuestraError Err.Number
End Function


Private Sub EliminarSinStocks()
Dim vWhere As String

    cadList = String(50, "*") & vbCrLf
    cadList = cadList & "Desea llevar a histórico el albarán....."
    cadList = cadList & vbCrLf & "Nº:  " & Text1(0).Text
    cadList = cadList & vbCrLf & "Fecha: " & Text1(1).Text
    cadList = cadList & vbCrLf & vbCrLf & " ¿Continuar? "
    cadList = cadList & vbCrLf & vbCrLf & String(50, "*") & vbCrLf

    If MsgBox(cadList, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Screen.MousePointer = vbHourglass
    
        NumRegElim = Data1.Recordset.AbsolutePosition

        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        
        vWhere = ObtenerWhereCP(False)
        
        If ActualizarElTraspaso("", vWhere, CodTipoMov, cadList) Then
            CargaGrid Me.DataGrid1, Me.Data2, True
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If

    End If
End Sub

Private Sub InsertarSubComponentes()
Dim SQL As String
Dim TipoDto As Byte
Dim J As Integer
Dim cantidad As Currency
Dim numlinea As Integer
Dim vCStock As CStock
'Si el articulo es de conjuntos, preguntara si quiere insertar la lineas de los conjuntos
   
        SQL = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", txtAux(1).Text, "T")
        If SQL = "1" Then
        
        
            'SI!!!!!!, es de conjuntos
            If MsgBox("Articulo con componentes. Desea insertar las lineas?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
            

            SQL = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            TipoDto = CByte(SQL)
            cantidad = ImporteFormateado(txtAux(3).Text)
            
            SQL = ObtenerWhereCP(False)
            SQL = Replace(SQL, NombreTabla, NomTablaLineas)
            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", SQL)
                 
   
   
            SQL = "Select sarti1.*,nomartic from sarti1,sartic where sarti1.codarti1=sartic.codartic and sarti1.codartic=" & DBSet(txtAux(1).Text, "T")
            Set miRsAux = New ADODB.Recordset
            'miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                'Limpiamos todo menos el almacen y el CC si lo tuviera
                For J = 1 To 8
                    txtAux(J).Text = ""
                Next
                
                
                Text2(16).Text = ""
                Text2(17).Text = ""
                txtAux(1).Text = miRsAux!codarti1
                txtAux(2).Text = miRsAux!NomArtic
                'Cantidad es la cantidad de la linea ppal * la del escandallo
                txtAux(3).Text = cantidad * miRsAux!cantidad
            
                ObtenerPrecioCompra
            
                
                txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            
            
                Set vCStock = New CStock
                If InicializarCStock(vCStock, "E", CStr(numlinea)) Then
                
            
            
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
                    SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(4).Text, "S") & ", " & DBSet(txtAux(5).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(6).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " & DBSet(Text2(17).Text, "T") & ","
                    SQL = SQL & DBSet(txtAux(9).Text, "T", "S") 'centro coste
                    SQL = SQL & ");"
                
                    If Not ejecutar(SQL, True) Then
                        MsgBox "Error añadiendo el componente: " & miRsAux!NomArtic, vbExclamation
                    
                    Else
                        numlinea = numlinea + 1
                        If Not vCStock.ActualizarStock Then MsgBox "Error actualizando stock componentes: " & miRsAux!NomArtic, vbExclamation
                    End If
                Else
                    MsgBox "Error stock: " & miRsAux!NomArtic, vbExclamation
                End If
                    
                    
                Set vCStock = Nothing
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
        
 
    
End Sub



Private Function PuedeEliminar(EsLinea As Boolean) As Boolean
Dim RN As ADODB.Recordset
Dim Aux As String

    PuedeEliminar = True
    
    
    If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Function
    
    If EsLinea Then
        If DBLet(Data2.Recordset!numlotes, "T") = "" Then Exit Function
    End If
    
    cadList = ObtenerWhereCP(True)
    cadList = Replace(cadList, NombreTabla, NomTablaLineas)
    cadList = "Select codartic,numlotes,fechaalb from " & NomTablaLineas & cadList & " AND numlotes "
    If EsLinea Then
        cadList = cadList & " = " & DBSet(Data2.Recordset!numlotes, "T")
    Else
        cadList = cadList & " <> ''"
    End If
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    miRsAux.Open cadList, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadList = ""
    While Not miRsAux.EOF
        
        Aux = "SELECT * FROM slotes WHERE "
        Aux = Aux & " codartic=" & DBSet(miRsAux!codArtic, "T") & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!FechaAlb, "F")
        RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RN.EOF Then
            MsgBox "Lote no encontrado: " & miRsAux!codArtic & "    " & miRsAux!numlotes & "     " & miRsAux!FechaAlb, vbExclamation
        Else
            If RN!vendida > 0 Then
                cadList = cadList & "Ya vendida  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
            Else
                'Vemos si existe LOTES realacionados en ventas, aunque NO deberia pasar
                Aux = "numlote =" & DBSet(miRsAux!numlotes, "T") & " AND fecentra = " & DBSet(miRsAux!FechaAlb, "F")
                Aux = Aux & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND 1"
                Aux = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", Aux, "1")
                If Val(Aux) > 0 Then
                    cadList = cadList & "Lotes asignados  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
                Else
                    Aux = "numlote =" & DBSet(miRsAux!numlotes, "T") & " AND fecentra = " & DBSet(miRsAux!FechaAlb, "F")
                    Aux = Aux & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND 1"
                    Aux = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", Aux, "1")
                    If Val(Aux) > 0 Then cadList = cadList & "En venta(I)  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
                End If
            End If
        End If
        RN.Close
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
        
    If cadList <> "" Then
        cadList = "Error en lotes. " & vbCrLf & vbCrLf & cadList
        MsgBox cadList, vbExclamation
        PuedeEliminar = False
    End If
    
    Set miRsAux = Nothing
    Set RN = Nothing
End Function


Private Sub EliminarEnSlotes()
Dim cad As String
    Set miRsAux = New ADODB.Recordset
    cad = ObtenerWhereCP(True)
    cad = Replace(cad, NombreTabla, NomTablaLineas)
    cad = "Select codartic,numlotes,fechaalb from " & NomTablaLineas & cad & " AND numlotes  <> ''"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    While Not miRsAux.EOF
                
        cad = "DELETE FROM slotes WHERE "
        cad = cad & " codartic=" & DBSet(miRsAux!codArtic, "T") & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!FechaAlb, "F")
        conn.Execute cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub
