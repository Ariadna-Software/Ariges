VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComEntPedidos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Proveedor"
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   15390
   Icon            =   "frmComEntPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10920
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Index           =   50
      Left            =   10485
      MaxLength       =   15
      TabIndex        =   137
      Text            =   "Text1 7"
      Top             =   270
      Width           =   2115
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   133
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   134
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
      Left            =   3825
      TabIndex        =   131
      Top             =   90
      Width           =   1875
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   132
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Albar�n"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Descuento"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Simular con otro proveedor"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5805
      TabIndex        =   129
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   130
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
      Height          =   195
      Left            =   13185
      TabIndex        =   128
      Top             =   315
      Width           =   1575
   End
   Begin VB.TextBox txtAux2 
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
      Index           =   8
      Left            =   4155
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   120
      Text            =   "nom ccoste"
      Top             =   10470
      Visible         =   0   'False
      Width           =   7785
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   120
      TabIndex        =   103
      Top             =   870
      Width           =   15175
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
         Left            =   8220
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scappr|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   570
         Width           =   6750
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
         Left            =   7290
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scappr|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   570
         Width           =   915
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
         Left            =   7290
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Realizado Por|N|N|0|9999|scappr|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   160
         Width           =   915
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
         Index           =   3
         Left            =   8220
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   106
         Text            =   "Text2"
         Top             =   160
         Width           =   6750
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scappr|fecpedpr|dd/mm/yyyy|N|"
         Top             =   420
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N� Pedido|N|S|0||scappr|numpedpr|0000000|S|"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1125
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
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
         Left            =   3405
         TabIndex        =   2
         Tag             =   "Resto de Pedido|N|N|||scappr|restoped||N|"
         Top             =   375
         Width           =   1980
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   7020
         Picture         =   "frmComEntPedidos.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   585
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   108
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Realizado Por"
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
         Index           =   21
         Left            =   5640
         TabIndex        =   107
         Top             =   180
         Width           =   1365
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   7005
         ToolTipText     =   "Buscar trabajador"
         Top             =   165
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ped."
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
         Index           =   14
         Left            =   1575
         TabIndex        =   105
         Top             =   180
         Width           =   1170
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2745
         Picture         =   "frmComEntPedidos.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
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
         Index           =   50
         Left            =   240
         TabIndex        =   104
         Top             =   180
         Width           =   1005
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
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   67
      Text            =   "frmComEntPedidos.frx":0A99
      Top             =   10095
      Width           =   9540
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   10020
      Width           =   2175
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
         TabIndex        =   55
         Top             =   180
         Width           =   1755
      End
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
      Left            =   14265
      TabIndex        =   23
      Top             =   10035
      Width           =   1035
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
      Left            =   13095
      TabIndex        =   22
      Top             =   10035
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6600
      Top             =   1320
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
      Top             =   1440
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
      Height          =   7785
      Left            =   120
      TabIndex        =   56
      Top             =   1920
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmComEntPedidos.frx":0AD6
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
      Tab(0).Control(5)=   "txtAux(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAux(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameToolAux0"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Direcciones/Observaciones/Totales"
      TabPicture(1)   =   "frmComEntPedidos.frx":0AF2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(30)"
      Tab(1).Control(1)=   "Text1(29)"
      Tab(1).Control(2)=   "FrameFactura"
      Tab(1).Control(3)=   "FrameHco"
      Tab(1).Control(4)=   "FrameDirFactura"
      Tab(1).Control(5)=   "FrameDirMercancia"
      Tab(1).Control(6)=   "Text1(21)"
      Tab(1).Control(7)=   "Text1(20)"
      Tab(1).Control(8)=   "Text1(19)"
      Tab(1).Control(9)=   "Text1(18)"
      Tab(1).Control(10)=   "Text1(17)"
      Tab(1).Control(11)=   "Label1(48)"
      Tab(1).Control(12)=   "Label1(47)"
      Tab(1).Control(13)=   "Label1(45)"
      Tab(1).ControlCount=   14
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
         Index           =   30
         Left            =   -65880
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "T|T|S|||scappr|SReferencia||N|"
         Top             =   2880
         Width           =   5505
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
         Index           =   29
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "O|T|S|||scappr|NReferencia||N|"
         Top             =   2880
         Width           =   5505
      End
      Begin VB.Frame FrameFactura 
         Height          =   2265
         Left            =   -74880
         TabIndex        =   139
         Top             =   5400
         Width           =   14580
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
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
            Index           =   49
            Left            =   12150
            MaxLength       =   15
            TabIndex        =   49
            Text            =   "Text1 7"
            Top             =   1740
            Width           =   2115
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Index           =   48
            Left            =   12150
            MaxLength       =   15
            TabIndex        =   48
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   2115
         End
         Begin VB.TextBox Text3 
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
            Index           =   42
            Left            =   9060
            MaxLength       =   5
            TabIndex        =   46
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   750
         End
         Begin VB.TextBox Text3 
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
            Index           =   39
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   45
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   885
         End
         Begin VB.TextBox Text3 
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
            Index           =   45
            Left            =   9945
            MaxLength       =   15
            TabIndex        =   47
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1890
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Index           =   47
            Left            =   12150
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "Text1 7"
            Top             =   900
            Width           =   2115
         End
         Begin VB.TextBox Text3 
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
            Index           =   41
            Left            =   9060
            MaxLength       =   5
            TabIndex        =   42
            Text            =   "Text1 7"
            Top             =   900
            Width           =   750
         End
         Begin VB.TextBox Text3 
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
            Index           =   38
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   41
            Text            =   "Text1 7"
            Top             =   900
            Width           =   885
         End
         Begin VB.TextBox Text3 
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
            Index           =   44
            Left            =   9945
            MaxLength       =   15
            TabIndex        =   43
            Text            =   "Text1 7"
            Top             =   900
            Width           =   1890
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Index           =   46
            Left            =   12150
            MaxLength       =   15
            TabIndex        =   40
            Text            =   "Text1 7"
            Top             =   540
            Width           =   2115
         End
         Begin VB.TextBox Text3 
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
            Index           =   40
            Left            =   9060
            MaxLength       =   5
            TabIndex        =   38
            Text            =   "Text1 7"
            Top             =   540
            Width           =   750
         End
         Begin VB.TextBox Text3 
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
            Index           =   37
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   37
            Text            =   "Text1 7"
            Top             =   540
            Width           =   885
         End
         Begin VB.TextBox Text3 
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
            Index           =   43
            Left            =   9945
            MaxLength       =   15
            TabIndex        =   39
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1890
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Index           =   36
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   36
            Text            =   "Text1 7"
            Top             =   465
            Width           =   1530
         End
         Begin VB.TextBox Text3 
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
            Index           =   35
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "Text1 7"
            Top             =   465
            Width           =   1335
         End
         Begin VB.TextBox Text3 
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
            Index           =   34
            Left            =   2010
            MaxLength       =   15
            TabIndex        =   34
            Text            =   "Text1 7"
            Top             =   465
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   33
            Text            =   "Text1 7"
            Top             =   465
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.IVA"
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
            Index           =   42
            Left            =   7920
            TabIndex        =   153
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
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
            Index           =   41
            Left            =   9135
            TabIndex        =   152
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL PEDIDO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   39
            Left            =   9960
            TabIndex        =   151
            Top             =   1755
            Width           =   1890
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
            TabIndex        =   150
            Top             =   960
            Width           =   135
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
            Left            =   11880
            TabIndex        =   149
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
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
            Index           =   33
            Left            =   12150
            TabIndex        =   148
            Top             =   285
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
            Left            =   5160
            TabIndex        =   147
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
            Left            =   3480
            TabIndex        =   146
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
            Left            =   1800
            TabIndex        =   145
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
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
            Index           =   18
            Left            =   5520
            TabIndex        =   144
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto.Gnral"
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
            Index           =   22
            Left            =   3720
            TabIndex        =   143
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto. PP"
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
            Index           =   23
            Left            =   2010
            TabIndex        =   142
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
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
            Index           =   24
            Left            =   240
            TabIndex        =   141
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
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
            Index           =   27
            Left            =   9945
            TabIndex        =   140
            Top             =   270
            Width           =   1665
         End
      End
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   225
         TabIndex        =   135
         Top             =   3000
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   136
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
         Index           =   2
         Left            =   11040
         TabIndex        =   119
         ToolTipText     =   "Buscar centro coste"
         Top             =   6120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   315
         Index           =   8
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   66
         Tag             =   "centro coste"
         Text            =   "cc"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminaci�n"
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
         Height          =   2160
         Left            =   -67395
         TabIndex        =   113
         Top             =   495
         Width           =   7175
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
            Index           =   24
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   50
            Top             =   345
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
            Index           =   25
            Left            =   1665
            MaxLength       =   30
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   840
            Width           =   780
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
            Index           =   25
            Left            =   2505
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   840
            Width           =   4500
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
            Index           =   26
            Left            =   1665
            MaxLength       =   30
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   1320
            Width           =   540
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
            Index           =   26
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   114
            Text            =   "Text2"
            Top             =   1320
            Width           =   4725
         End
         Begin VB.Label Label1 
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
            Index           =   37
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label1 
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
            Index           =   38
            Left            =   75
            TabIndex        =   117
            Top             =   915
            Width           =   1140
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1305
            ToolTipText     =   "Buscar trabajador"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
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
            Index           =   40
            Left            =   120
            TabIndex        =   116
            Top             =   1365
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1305
            ToolTipText     =   "Buscar incidencia"
            Top             =   1365
            Width           =   240
         End
      End
      Begin VB.Frame FrameDirFactura 
         Caption         =   "Direcci�n Factura"
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
         Height          =   2145
         Left            =   -67410
         TabIndex        =   93
         Top             =   510
         Width           =   7175
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
            Left            =   1230
            MaxLength       =   30
            TabIndex        =   32
            Tag             =   "Direc. Factura|N|S|0|999|scappr|coddiref|000|N|"
            Text            =   "Text1"
            Top             =   375
            Width           =   540
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
            Left            =   1785
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   102
            Text            =   "Text2"
            Top             =   375
            Width           =   5220
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
            Index           =   24
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   97
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1590
            Width           =   5745
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
            Index           =   22
            Left            =   1245
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   96
            Text            =   "Text15"
            Top             =   1185
            Width           =   900
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
            Index           =   23
            Left            =   2190
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   95
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1185
            Width           =   4815
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
            Index           =   21
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   94
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   780
            Width           =   5760
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   945
            ToolTipText     =   "Buscar direcci�n"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   13
            Left            =   135
            TabIndex        =   101
            Top             =   1590
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Index           =   12
            Left            =   120
            TabIndex        =   100
            Top             =   1185
            Width           =   1050
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
            Index           =   11
            Left            =   120
            TabIndex        =   99
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "C�digo"
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
            Index           =   10
            Left            =   120
            TabIndex        =   98
            Top             =   375
            Width           =   795
         End
      End
      Begin VB.Frame FrameDirMercancia 
         Caption         =   "Direcci�n Mercancia"
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
         Height          =   2145
         Left            =   -74805
         TabIndex        =   83
         Top             =   510
         Width           =   7175
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
            Index           =   17
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   88
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   780
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
            Index           =   19
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   87
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1185
            Width           =   4755
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
            Index           =   18
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   86
            Text            =   "Text15"
            Top             =   1185
            Width           =   900
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
            Index           =   20
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   85
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1590
            Width           =   5715
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
            Height          =   315
            Index           =   15
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   24
            Tag             =   "Direc. Mercancia|N|S|0|999|scappr|coddirea|000|N|"
            Text            =   "Text1"
            Top             =   375
            Width           =   540
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
            Index           =   15
            Left            =   1875
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   84
            Text            =   "Text2"
            Top             =   375
            Width           =   5175
         End
         Begin VB.Label Label1 
            Caption         =   "C�digo"
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
            Left            =   120
            TabIndex        =   92
            Top             =   375
            Width           =   870
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
            Index           =   5
            Left            =   120
            TabIndex        =   91
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Index           =   2
            Left            =   120
            TabIndex        =   90
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   120
            TabIndex        =   89
            Top             =   1590
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   1035
            ToolTipText     =   "Buscar direcci�n"
            Top             =   375
            Width           =   240
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   2505
         Left            =   240
         TabIndex        =   72
         Top             =   495
         Width           =   14775
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
            Index           =   27
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Direc. recogida|N|S|0|999|scappr|coddirre|000|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   630
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
            Index           =   27
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   1920
            Width           =   3960
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
            Index           =   4
            Left            =   10800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   123
            Text            =   "Text2"
            Top             =   1920
            Width           =   3915
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
            Index           =   31
            Left            =   10080
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "Envio|N|S|0|999|scappr|codenvio|0000|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   660
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
            Index           =   28
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Fecha entraga|F|S|||scappr|fecentrega|dd/mm/yyyy|N|"
            Top             =   1920
            Width           =   1350
         End
         Begin VB.CheckBox chkObra 
            Caption         =   "Obra"
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
            Left            =   6150
            TabIndex        =   12
            Tag             =   "Obra|N|N|||scappr|obra||N|"
            Top             =   1455
            Width           =   855
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
            Left            =   9795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   111
            Text            =   "Text2"
            Top             =   600
            Width           =   4875
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
            Index           =   23
            Left            =   8775
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Solicitado Por|N|S|0|9999|scappr|codtrab1|0000|N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   975
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
            Left            =   9795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   109
            Text            =   "Text2"
            Top             =   190
            Width           =   4875
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
            Index           =   22
            Left            =   8775
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Cliente|N|S|0|999999|scappr|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   190
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
            Index           =   16
            Left            =   8775
            MaxLength       =   25
            TabIndex        =   16
            Tag             =   "Tipo Portes|T|S|||scappr|tipoporte||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwww"
            Top             =   1440
            Width           =   2130
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
            Index           =   11
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scappr|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1440
            Width           =   4545
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
            Index           =   9
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scappr|codpobla||N|"
            Text            =   "Text15"
            Top             =   1005
            Width           =   855
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
            Index           =   10
            Left            =   2355
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Poblaci�n|T|N|||scappr|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1005
            Width           =   4605
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
            Left            =   4365
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "tel�fono Proveedor|T|S|||scappr|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   190
            Width           =   2595
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
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF Proveedor|T|N|||scappr|nifprove||N|"
            Text            =   "123456789"
            Top             =   190
            Width           =   1905
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
            Index           =   12
            Left            =   8775
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Forma de Pago|N|N|0|999|scappr|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1005
            Width           =   975
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
            Index           =   12
            Left            =   9795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   74
            Text            =   "Text2"
            Top             =   1005
            Width           =   4875
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
            Index           =   13
            Left            =   11880
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   795
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
            Index           =   14
            Left            =   13860
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "12"
            Top             =   1440
            Width           =   795
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
            Left            =   1440
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scappr|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   600
            Width           =   5520
         End
         Begin VB.Label Label1 
            Caption         =   "F.Recogida"
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
            Index           =   44
            Left            =   6240
            TabIndex        =   127
            Top             =   1973
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Recogida"
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
            Index           =   28
            Left            =   120
            TabIndex        =   126
            Top             =   1973
            Width           =   990
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   11
            Left            =   1200
            ToolTipText     =   "Buscar direcci�n"
            Top             =   1980
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Envio"
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
            Index           =   49
            Left            =   9240
            TabIndex        =   125
            Top             =   1973
            Width           =   600
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   7440
            Picture         =   "frmComEntPedidos.frx":0B0E
            ToolTipText     =   "Buscar fecha"
            Top             =   1980
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   12
            Left            =   9840
            ToolTipText     =   "Buscar direcci�n"
            Top             =   1980
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   8490
            ToolTipText     =   "Buscar trabajador"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Solicitado por"
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
            Index           =   6
            Left            =   7230
            TabIndex        =   112
            Top             =   600
            Width           =   1230
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   8490
            ToolTipText     =   "Buscar cliente"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Para Cliente"
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
            Index           =   4
            Left            =   7230
            TabIndex        =   110
            Top             =   195
            Width           =   1305
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1215
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   210
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Portes"
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
            Index           =   3
            Left            =   7215
            TabIndex        =   82
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   1170
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Index           =   17
            Left            =   120
            TabIndex        =   81
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Index           =   16
            Left            =   120
            TabIndex        =   80
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
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
            Index           =   19
            Left            =   3465
            TabIndex        =   79
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
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
            Index           =   20
            Left            =   120
            TabIndex        =   78
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
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
            Index           =   15
            Left            =   7230
            TabIndex        =   77
            Top             =   1005
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dto.P.P"
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
            Index           =   25
            Left            =   11070
            TabIndex        =   76
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Dto.Gral"
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
            Index           =   26
            Left            =   13005
            TabIndex        =   75
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   8460
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1020
            Width           =   240
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
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   960
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
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   71
         ToolTipText     =   "Buscar art�culo"
         Top             =   6120
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
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   70
         ToolTipText     =   "Buscar almacen"
         Top             =   6120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
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
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
         Tag             =   "Nombre Art�culo"
         Text            =   "nomArtic"
         Top             =   6120
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   12
         TabIndex        =   68
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   6
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   65
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   6120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   315
         Index           =   5
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   64
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   4
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   63
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   6120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   315
         Index           =   3
         Left            =   6000
         MaxLength       =   16
         TabIndex        =   62
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   60
         Tag             =   "C�digo Art�culo"
         Text            =   "Artic Artic Artic5"
         Top             =   6060
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
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
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   59
         Tag             =   "C�digo Almacen"
         Text            =   "codalmac"
         Top             =   6060
         Visible         =   0   'False
         Width           =   615
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
         Index           =   21
         Left            =   -73185
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observaci�n 5|T|S|||scappr|observa5||N|"
         Top             =   5010
         Width           =   12885
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
         Index           =   20
         Left            =   -73185
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observaci�n 4|T|S|||scappr|observa4||N|"
         Top             =   4650
         Width           =   12885
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
         Index           =   19
         Left            =   -73185
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observaci�n 3|T|S|||scappr|observa3||N|"
         Top             =   4290
         Width           =   12885
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
         Index           =   18
         Left            =   -73185
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observaci�n 2|T|S|||scappr|observa2||N|"
         Top             =   3930
         Width           =   12885
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
         Index           =   17
         Left            =   -73185
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observaci�n 1|T|S|||scappr|observa1||N|"
         Top             =   3570
         Width           =   12885
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntPedidos.frx":0B99
         Height          =   3930
         Left            =   240
         TabIndex        =   69
         Top             =   3690
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   6932
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
      Begin VB.Label Label1 
         Caption         =   "S/Referencia"
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
         Index           =   48
         Left            =   -67320
         TabIndex        =   155
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "N/Referencia"
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
         Index           =   47
         Left            =   -74760
         TabIndex        =   154
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         Index           =   45
         Left            =   -74760
         TabIndex        =   58
         Top             =   3570
         Width           =   1545
      End
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
      Left            =   14220
      TabIndex        =   53
      Top             =   10035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "BASE IMPONIBLE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   29
      Left            =   8505
      TabIndex        =   138
      Top             =   285
      Width           =   1890
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   13
      Left            =   4095
      ToolTipText     =   "Ampliacion"
      Top             =   9810
      Width           =   240
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
      Left            =   12195
      TabIndex        =   122
      Top             =   10530
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Centro coste"
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
      Index           =   46
      Left            =   2400
      TabIndex        =   121
      Top             =   10515
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
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
      Index           =   35
      Left            =   2400
      TabIndex        =   57
      Top             =   9825
      Width           =   1695
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
      Begin VB.Menu mnGeneraDtos 
         Caption         =   "Modificar &descuentos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnImpPedido 
         Caption         =   "&Imprimir Pedido"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnSimularProv 
         Caption         =   "Simular proveedor"
         Shortcut        =   ^P
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
Attribute VB_Name = "frmComEntPedidos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)


Public MostrarDatos As String  'Para ver un dato enconcreto
Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schppr, y solo en modo de consulta
                              
                              
Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmBasico2  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmProveV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmProveV.VB_VarHelpID = -1
Private WithEvents frmDir As frmBasico2 'frmComDirecciones
Attribute frmDir.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'frmFacClientesGr 'form mantenimiento clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmInc As frmIncidencias  'form mantenimiento incidencias eliminacion
Attribute frmInc.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar n� Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmNLote As frmAlmCargarNLote   'Form Cargar n� lote
Attribute frmNLote.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer 'Listados
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1

Private WithEvents frmDiren As frmComDirecciones
Attribute frmDiren.VB_VarHelpID = -1


Private FrmArt2 As frmAlmArticulosGr   'Form Articulos

Private WithEvents frmRecoge As frmComDirRecogida
Attribute frmRecoge.VB_VarHelpID = -1


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
'   6.- Cargar cantidad servidas al Generar Albaran no completo (Pedido --> Albaran)
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean


'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom
Dim CodTipoMov As String


Dim EsDeVarios As Boolean 'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnAnyadir As Byte

'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)

Dim PulsadoMas2 As Boolean

Private Sub chkObra_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'================================================================================

Private Sub cmdAceptar_Click()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
                    SQL = CadenaInsertarDesdeForm(Me)
                    If SQL <> "" Then
                        If InsertarPedido(SQL, vTipoMov) Then
                            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                            PonerCadenaBusqueda
                            PonerModo 2
                            'Ponerse en Modo Insertar Lineas
                            BotonMtoLineas 1, "Pedidos"
                            BotonAnyadirLinea
                        End If
                    End If
                    FormateaCampo Text1(0)
                End If
                Set vTipoMov = Nothing
            End If
            Me.SSTab1.Tab = 0
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    'Actualizar los datos del Proveedor si es de varios
                    ActualizarProveVarios Text1(4).Text, Text1(6).Text
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
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                     NumRegElim = Data2.Recordset!numlinea
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PosicionarData2
                    
                    BloquearTxt Text2(16), True
                     PonerModo 2
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
            
        Case 6 'Pasar Pedido a Albaran
            If BLOQUEADesdeFormulario(Me) Then GenerarAlbaran
            TerminaBloquear
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
            PonerFoco txtAux(Index)
            
        Case 1 'Busqueda de Cod. Artic
            Set FrmArt = New frmBasico2
            'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
'            FrmArt.DesdeTPV = False
'            FrmArt.Show vbModal
            AyudaArticulos FrmArt, txtAux(Index)
            Set FrmArt = Nothing
            PonerFoco txtAux(Index)
            
        Case 2 'COD. CENTRO DE COSTE
            If vEmpresa.TieneAnalitica Then
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtAux(8)
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
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            Else
                ModificaLineas = 0
            End If
'            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            PonerModo 2
           ' PonerCampos lo quita david
            
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            If MsgBox("Desea cancelar la introducci�n de unidades del pedido?", vbQuestion + vbYesNo) = vbYes Then
                TerminaBloquear
                InicializarServidas
                PonerModo 2
                CargaTxtAuxServidas False, False
                CargaGrid DataGrid1, Data2, True, False
                
            Else
                PonerFoco Me.txtAux(3)
            End If
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
Dim alma As Integer

    PonerModo 5


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
       
       
       
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    
'    PonerBotonCabecera False
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso


    lblIndicador.Caption = "INSERTAR"
    
    alma = 0
    If vParamAplic.NumeroInstalacion = 2 Then
        If Not Data2.Recordset.EOF Then
            Data2.Recordset.MoveFirst
            alma = Data2.Recordset!codAlmac
        End If
    End If
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    'HERBELCA 2015. Si hay mas de una linea que coja el primer almacen
    If alma = 0 Then
        txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    Else
        txtAux(0).Text = alma
    End If
    
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(8).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
    End If
    ' ----
    
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
        Text1(0).BackColor = vbLightBlue  'vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
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
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim SQL As String
On Error GoTo EModificar

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
            
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
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo EModificarLinea

    PonerModo 5

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    DeseleccionaGrid DataGrid1
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim Cad As String
Dim vTipoMov As CTiposMov
Dim NumPedElim As Long 'Numero del Pedido que se ha Eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "Cabecera de Pedidos Compras." & vbCrLf
    Cad = Cad & "--------------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Pedido:            "
    Cad = Cad & vbCrLf & "N�:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Proveedor:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
       
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(0).Value
        
        CadenaSQL = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 81
        frmList.Show vbModal
        Set frmList = Nothing
    
        If CadenaSQL = "" Then Exit Sub
        Cad = ""
        Cad = DBSet(RecuperaValor(CadenaSQL, 1), "F") & " as fechelim,"
        Cad = Cad & RecuperaValor(CadenaSQL, 2) & " as trabelim,"
        Cad = Cad & DBSet(RecuperaValor(CadenaSQL, 3), "T") & " as codincid"
        CadenaSQL = Cad
        
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumPedElim
        Set vTipoMov = Nothing
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
    SQL = "�Seguro que desea eliminar la l�nea del Pedido?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Art�culo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
'        CancelaADODC
        PonerModo 2
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub BotonGenerarAlbaran()
    'Pasar una Pedido a Albaran
Dim Resp As Byte

    'Comprobar que hay un Pedido seleccionado
    If Text1(0).Text = "" Then Exit Sub
        
    'Comprobar que hay lineas
    Resp = 1
    If Not (Data2.Recordset Is Nothing) Then
        If Not Data2.Recordset.EOF Then
            If Not IsNull(Data2.Recordset!numlinea) Then Resp = 0
        End If
    End If
    If Resp = 1 Then
        MsgBox "Pedido sin lineas", vbExclamation
        Exit Sub
    End If
    'Preguntar si se Recibe el pedido completo o no
    Resp = MsgBox("�Recibir el pedido completo?", vbYesNoCancel + vbQuestion)
    If Resp = vbCancel Then Exit Sub
    
    
    'Agosto 2013
    If vParamAplic.AlmacenB > 90 Then  'EMPRESA HEREBELCA
        IMprimirNormalPedidoProv True   'imprime directament el pedido
    End If
    
    If Resp = vbYes Then 'RECIBIR EL PEDIDO COMPLETO
        AlbCompleto = True
        Screen.MousePointer = vbHourglass

        GenerarAlbaran
        TerminaBloquear
        
    ElseIf Resp = vbNo Then 'RECIBIR PEDIDO INCOMPLETO
        AlbCompleto = False
        Me.SSTab1.Tab = 0
        TerminaBloquear
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a recibir para cada l�nea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        CargaTxtAuxServidas True, True
        PrimeraVez = True
    Else
        TerminaBloquear
    End If

End Sub





Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        'DAVID. Pongo a pi�on el numero de pedido. YA NO SE UTILIZA
        'cad = Data1.Recordset.Fields(0)
        'RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1
    
    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    End If
    
    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedpr", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
        Else
            Text2(16).Text = ""
        End If
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
        If Not Data2.Recordset.EOF And vEmpresa.TieneAnalitica Then
            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            txtAux2(8).Text = ""
        End If
    Else
        If Modo = 2 Then
            If Not Data2.Recordset.EOF Then Text2(16).Text = DBLet(Data2.Recordset!Ampliaci, "T")
        End If
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    If MostrarDatos <> "" Then
        MostrarDatos = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim SelectInicial As String
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next
    
    
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 24
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 10 'Mto Lineas Ofertas
'        .Buttons(11).Image = 26 'Generar Albaran
'
'        'OCtubre 2011
'        .Buttons(12).Image = 43 'Modificar descuentos
'
'        .Buttons(14).Image = 16 'Imprimir Pedido
'        .Buttons(16).Image = 45  'simular
'        .Buttons(21).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With


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
        .Buttons(1).Image = 32 'Generar Albaran
        .Buttons(2).Image = 43 'Modificar descuentos
        .Buttons(3).Image = 45  'simular
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
        End With
    Next i


    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    cmdAux(2).Tag = -1
          
    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    
    'Si no lleva datosdvolverbusqueda
    
    If Not EsHistorico Then
        NombreTabla = "scappr"
        NomTablaLineas = "slippr" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr "

    Else
        NombreTabla = "schppr"
        NomTablaLineas = "slhppr"
        CargarTagsHco Me, "scappr", NombreTabla
        'Estos campos solo estan en la tabla del hist�rico
        Text1(24).Tag = "Fecha Eliminaci�n|F|N|||" & NombreTabla & "|fechelim|dd/mm/yyyy|N|"
        Text1(25).Tag = "Trabajador Eliminaci�n|N|N|0|9999|" & NombreTabla & "|trabelim|0000|N|"
        Text1(26).Tag = "Incidencia elim.|T|N|||" & NombreTabla & "|codincid||N|"
        Me.Caption = "Hist�rico Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr,fecpedpr "
    End If
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    If MostrarDatos = "" Then
        CodTipoMov = "-1"
    Else
        CodTipoMov = MostrarDatos
    End If
    Data1.RecordSource = "Select * from " & NombreTabla & "  WHERE numpedpr= " & CodTipoMov
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
 
    If MostrarDatos = "" Then
        PonerModo 0
    Else
        PonerModo 2
    End If
    
    
    CodTipoMov = "PEC"
    VieneDeBuscar = False
    
    
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True

    'Poner los grid sin apuntar a nada
    If MostrarDatos = "" Then LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkRestoPed.Value = 0
    Me.chkObra.Value = 0
    Text3(0).Text = "TOTAL"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    conn.Execute "DELETE FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    'DatosADevolverBusqueda2 = "
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Val(cmdAux(2).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
            Me.txtAux(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Val(cmdAux(2).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
            Me.txtAux(8).Text = RecuperaValor(CadenaSeleccion, 1)
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
            cadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000000")
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(22).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
    FormateaCampo Text1(22)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmDir_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Direcciones
Dim Indice As Byte
    Indice = CByte(Me.imgBuscar(0).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Direccion
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Direc

    CargarDatosDirec Text1(Indice).Text, Indice
End Sub

Private Sub frmDiren_DatoSeleccionado(CadenaSeleccion As String)
    TituloLinea = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
    CadenaSQL = CadenaSeleccion
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte

    Indice = 12
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de incidencias
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod incidencia
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nom incidencia
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen en el Listado
'para pasar de Pedido a Albaran, o para pasar al historico
    
    CadenaSQL = CadenaSeleccion
End Sub



Private Sub frmNSerie_CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'N� de Serie introducidos en la Tabla Temporal
Dim RStmp As ADODB.Recordset
Dim RSalb As ADODB.Recordset
Dim SQL As String
Dim i As Byte
Dim B As Boolean
    
    On Error GoTo EInsertar

    
    SQL = "SELECT slialp.codartic, numlinea, cantidad "
    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
    SQL = SQL & " WHERE numalbar=" & DBSet(Me.cmdAux(1).Tag, "T") & " and fechaalb=" & DBSet(Me.cmdAux(0).Tag, "F") & " and "
    SQL = SQL & "slialp.codprove=" & Text1(4).Text
    SQL = SQL & " And nseriesn = 1 "
    SQL = SQL & " ORDER BY codartic, numlinea "

    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RSalb.EOF 'Para cada linea del ALbaran
        'Recuperar los N� Serie de ese articulo cargados en la Temporal
        'Seleccionar los n� de serie cargados en la temporal: tmpnseries
        SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
        SQL = SQL & " AND codartic=" & DBSet(RSalb!codArtic, "T")
        SQL = SQL & " ORDER BY codartic, numlinea "
        Set RStmp = New ADODB.Recordset
        RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'If Not RStmp.EOF Then RStmp.MoveFirst
        'Intentar asignar un N� serie al total de cantidad del articulo
        
        B = True
        For i = 1 To RSalb!cantidad
            If Not RStmp.EOF Then
                InsertarNSerie RStmp!numSerie, RStmp!codArtic, RSalb!numlinea
                RStmp.MoveNext
            End If
        Next i
        RStmp.Close
        Set RStmp = Nothing
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando N� Serie", Err.Description
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Prove
    FormateaCampo Text1(Indice)
End Sub

Private Sub frmProveV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento Proveedores varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'nif Prove
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'nom Prove
    PonerDatosProveVario Text1(Indice).Text
End Sub

Private Sub frmRecoge_DatoSeleccionado(CadenaSeleccion As String)
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1) 'coddirre
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2) 'nomdirre
    FormateaCampo Text1(27)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(0).Tag)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    If Indice = 23 Then Indice = 1
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Index = 13 Then
        If Modo = 0 Then Exit Sub
        If Not (Modo = 2 Or Modo = 5) Then Exit Sub
                
    Else
        If Modo = 2 Or Modo = 0 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Indice = 0
    Select Case Index
        Case 0 'Cod. Proveedor
            Indice = 4
            Set frmProv = New frmBasico2
'            frmProv.DatosADevolverBusqueda = "0"
'            frmProv.Show vbModal
            AyudaProveedores frmProv, Text1(Indice)
            Set frmProv = Nothing
            
            
        Case 1 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            
        Case 2, 8 'Realizada Por Trabajador
            If Index = 2 Then
                Indice = 3
            Else
                Indice = 23
            End If
            Me.imgBuscar(0).Tag = Indice
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
            
        Case 3 'Forma de Pago
            Indice = 12
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
        
        
        Case 4
            TituloLinea = ""
            Set frmDiren = New frmComDirecciones
            frmDiren.DatosADevolverBusqueda = "0"
            frmDiren.Show vbModal
            Set frmDiren = Nothing
            If TituloLinea <> "" Then
                Text1(15).Text = Format(RecuperaValor(TituloLinea, 1), "000") 'Cod Direccion
                Text2(15).Text = RecuperaValor(TituloLinea, 2) 'Nom Direc
            
                CargarDatosDirec Text1(15).Text, 15
        
            
                TituloLinea = ""
            End If
        Case 5  'dpto
            Dim vAux As String
            If Index = 4 Then
                'YA NO ESTA.
                Indice = 15
                vAux = "sdirpr.tipodire = 0"
            End If
            If Index = 5 Then
                Indice = 2
                vAux = "sdirpr.tipodire = 1"
            End If
            Me.imgBuscar(0).Tag = Indice
'            Set frmDir = New frmComDirecciones
'            frmDir.DatosADevolverBusqueda = "0"
'            frmDir.Show vbModal
'            Set frmDir = Nothing
            Set frmDir = New frmBasico2
            AyudaDireccionesCompra frmDir, Text1(Indice), vAux
            Set frmDir = Nothing
            
        Case 6 'NIF de Proveedores VARIOS
            Indice = 6
            Set frmProveV = New frmComProveV
            frmProveV.DatosADevolverBusqueda = "0"
            frmProveV.Show vbModal
            Set frmProveV = Nothing
            
        Case 7 'Cliente
            Indice = 22
'            Set frmCli = New frmFacClientesGr
'            frmCli.DatosADevolverBusqueda = "0"
'            frmCli.Show vbModal
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(Indice).Text
            Set frmCli = Nothing
            
        Case 10 'Incidencias
            Indice = 26
            Set frmInc = New frmIncidencias
            frmInc.DatosADevolverBusqueda = "0"
            frmInc.Show vbModal
            Set frmInc = Nothing
            
        Case 11
            If Text1(4).Text = "" Then
                MsgBox "Ponga primero el proveedor", vbExclamation
                PonerFoco Text1(4)
            Else
                Indice = 27
                Set frmRecoge = New frmComDirRecogida
                frmRecoge.Codprove = CLng(Text1(4).Text)
                frmRecoge.nomprove = Text1(5).Text
                If Text1(Indice).Text <> "" Then
                    frmRecoge.VerDatoDpto = Text1(Indice).Text
                Else
                    frmRecoge.VerDatoDpto = -1
                End If
                frmRecoge.Show vbModal
                Set frmRecoge = Nothing
            
            End If
            
            
        
        Case 12
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0|1|"
            CadenaSQL = ""
            frmFE.Show vbModal
            Set frmFE = Nothing
            If CadenaSQL <> "" Then
                Text1(31).Text = RecuperaValor(CadenaSQL, 1)
                Text2(4).Text = RecuperaValor(CadenaSQL, 2)
                CadenaSQL = ""
            End If


        Case 13
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
    If Indice > 0 Then PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   If Index = 0 Then
        Indice = 1 'Index + 1
   Else
        Indice = 28
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


Private Sub mnGenAlbaran_Click()
    'bloqueamos el pedido y lo pasamos a Albaran
    If BLOQUEADesdeFormulario(Me) Then BotonGenerarAlbaran
End Sub


Private Sub mnGeneraDtos_Click()
Dim B As Boolean
    If Text1(0).Text = "" Then Exit Sub 'por si las moscas
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    
    'Modifica los descuentos de este albaran y recalcula los importes de las lineas(y por ende el total)
     If BLOQUEADesdeFormulario(Me) Then
        
        B = False
        
        CadenaDesdeOtroForm = Text1(5).Text & "(" & Text1(4).Text & ")|" & Text1(0).Text & " de " & Text1(1).Text & "|"
        'en el load pone a "" la variable
        frmVarios.Opcion = 9
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            conn.BeginTrans
            B = ActualizarDtos
            If B Then
                conn.CommitTrans
            Else
                conn.RollbackTrans
            End If
        End If
        'Termina bloquear
        TerminaBloquear
        If B Then PonerCampos
     End If
End Sub

Private Sub mnImpPedido_Click()
    IMprimirNormalPedidoProv False
End Sub

Private Sub IMprimirNormalPedidoProv(EsImpresionDirecta As Boolean)
'Imprime un Pedido
       frmListadoOfer.NumCod = Text1(0).Text    'N� de Pedido
       frmListadoOfer.codClien = Text1(4).Text 'Cod.Proveedor
       If EsHistorico Then
            AbrirListadoOfer (56) '59: Informe de Pedidos Compras (Historico)
            frmListadoOfer.FecEntre = Text1(1).Text
       Else
            If EsImpresionDirecta Then
                AbrirListadoOfer (407) '55: Informe de Pedidos Compras
            Else
                AbrirListadoOfer (55) '55: Informe de Pedidos Compras
            End If
       End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Pedidos"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Cabecera de Pedidos
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

Private Sub mnSimularProv_Click()
    'Aqui
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    
    
    
    CadenaDesdeOtroForm = ""
    frmListado3.OtrosDatos = Text1(1).Text & "|" & Text1(13).Text & "|" & Text1(14).Text & "|" & Text1(0).Text & "|"
    frmListado3.Opcion = 57
    frmListado3.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE numpedpr IN (" & CadenaDesdeOtroForm & ") ORDER BY numpedpr"
        PonerCadenaBusqueda
        CadenaDesdeOtroForm = ""
    End If
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
Dim B As Boolean
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        B = False
        If Text1(Index).Text = "" Then
            B = True
        Else
            If Text1(Index).SelLength = Len(Text1(Index).Text) Then B = True
        End If
        If B Then
                Ind = -1
                Select Case Index
                Case 2
                    Ind = 5
                Case 3
                    Ind = 2
                Case 4
                    Ind = 0
                Case 6
                    Ind = 6
                Case 9
                    Ind = 1
                Case 12
                    Ind = 3
                Case 15
                    Ind = 4
                Case 22, 23
                    Ind = Index - 15
                Case 27
                    Ind = 11
                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 2 'trabajador
            Case 4: KEYBusqueda KeyAscii, 0 'proveedor
            Case 6: KEYBusqueda KeyAscii, 6 'nif
            Case 9: KEYBusqueda KeyAscii, 1 'cpostal
            Case 22: KEYBusqueda KeyAscii, 7 'cliente
            Case 23: KEYBusqueda KeyAscii, 8 'solicitado
            Case 12: KEYBusqueda KeyAscii, 3 'forma de pago
            Case 27: KEYBusqueda KeyAscii, 11 'recogida
            Case 31: KEYBusqueda KeyAscii, 12 'envio
            
            Case 15: KEYBusqueda KeyAscii, 4 'direccion mercancia
            Case 25: KEYBusqueda KeyAscii, 9 'trabajador
            Case 26: KEYBusqueda KeyAscii, 10 'incidencia
            Case 2: KEYBusqueda KeyAscii, 5 'direccion fra
            
            Case 28: KEYFecha KeyAscii, 1 'fecha de recogida
            Case 1: KEYFecha KeyAscii, 0 'fecha de pedido
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
    imgFec_Click (Indice)
End Sub

Private Sub imgFec_Click(i As Integer)
    MsgBox "FALTA MONI !!!!     Lo he creado para que no me de error"
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
Dim i As Byte
        
        
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
    
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 28 'Fecha Oferta, Fecha Entrega
            PonerFormatoFecha Text1(Index)
        
        Case 3, 23, 31 'Cod Trabajador
            i = Index
            If Index = 23 Then i = 1
            If Index = 31 Then i = 4
            If PonerFormatoEntero(Text1(Index)) Then
                If Index = 31 Then
                    Text2(i).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio", "codenvio", "el envio")
                Else
                    Text2(i).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba", "el Trabajador")
                End If
                
            Else
                Text2(i).Text = ""
            End If
            
        Case 4 'Cod. Prove
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                Else ' cargar datos de Tabla sprove
                    PonerDatosProveedor (Text1(Index).Text)
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
            PonerDatosProveVario (Text1(Index).Text)
             
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
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                 If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 15, 2 'Cod. Direccion
            If PonerFormatoEntero(Text1(Index)) Then
                Me.imgBuscar(0).Tag = Index
                If Not CargarDatosDirec(Text1(Index).Text, CByte(Index)) Then
                    PonerFoco Text1(Index)
                End If
            Else
                LimpiarDatosDirec CByte(Index)
            End If
            
        Case 22 'cod.cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(0).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            Else
                Text2(0).Text = ""
            End If
            
        Case 21
            If Me.ActiveControl.Name = "SSTab1" Then PonerFocoBtn Me.cmdAceptar
            
        Case 26 'cod Incidencia de eliminacion
            If EsHistorico Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sincid", "nomincid")
                If Not (Text2(Index).Text = "" And Text1(Index).Text <> "") Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    PonerFoco Text1(Index)
                End If
            End If
            
        Case 27
            devuelve = ""
            If Text1(4).Text <> "" Then
                If Text1(Index).Text <> "" Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        devuelve = PonerNombreDeCod(Text1(Index), conAri, "sdirRecog", "nomdirre", "codprove=" & Text1(4).Text & " AND coddirre")
                        If devuelve = "" Then PonerFoco Text1(Index)
                    End If
                End If
            Else
                If Modo > 2 Then
                    MsgBox "Debe poner proveedor", vbExclamation
                    PonerFoco Text1(4)  'que ponga el proveedor
                End If
            End If
            Text2(27).Text = devuelve
            If devuelve = "" And Text1(Index).Text <> "" Then Text1(Index).Text = ""
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
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
Dim devuelve As String
    'Llamamos a al form
    '##A mano
'    cad = ""
''    If EsCabecera Then
'        cad = cad & ParaGrid(Text1(0), 15, "N� Pedido")
'        cad = cad & ParaGrid(Text1(1), 20, "Fecha Ped.")
'        cad = cad & ParaGrid(Text1(4), 15, "Proveedor")
'        cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
'        tabla = NombreTabla
'        Titulo = "Pedidos Compras"
'        If EsHistorico Then
'            Titulo = "Hist�rico de Pedidos"
'            devuelve = "0|1|"
'        Else
'            Titulo = "Pedidos"
'            devuelve = "0|"
'        End If
''    End If
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
''        frmB.vDevuelve = "0|1|"
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri 'Conexi�n a BD: Ariges
''        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
''        If HaDevueltoDatos Then
''''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            PonerFoco Text1(kCampo)
'        'End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmB = New frmBasico2
    AyudaPedidosCompra frmB, NombreTabla, Text1(0)
    Set frmB = Nothing


End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
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
'Carga las Pesta�as con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slippr
    CargaGrid DataGrid1, Data2, True


    PonerModo 2

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
    
    'Realizado por
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
    'Cliente para
    Text2(0).Text = PonerNombreDeCod(Text1(22), conAri, "sclien", "nomclien")
    'Solicitado por
    Text2(1).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
    
    'Direccion de recogida
    If Text1(27).Text <> "" Then
        Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "sdirRecog", "nomdirre", "codprove=" & Text1(4).Text & " AND coddirre")
    Else
        Text2(27).Text = ""
    End If
     'Envio
    Text2(4).Text = PonerNombreDeCod(Text1(31), conAri, "senvio", "nomenvio", "codenvio", "el envio")
    'Poner las direcciones
    CargarDatosDirec Text1(15).Text, 15
    CargarDatosDirec Text1(2).Text, 2
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Pedidos
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(25).Text = PonerNombreDeCod(Text1(25), conAri, "straba", "nomtraba", "codtraba")
        Text2(26).Text = PonerNombreDeCod(Text1(26), conAri, "sincid", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pesta�a de totales
    
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
Dim B As Boolean
On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    lblF.Caption = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos

        cmdRegresar.visible = False
 
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True
       
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
       
       
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 50
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
       
       
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    chkObra.Enabled = B
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(13).Enabled = Modo = 2 Or Modo = 5
    Me.imgBuscar(1).visible = False
           
    'Modo Linea de Ofertas. Poner el campo ampliacion linea
  '  Me.Label1(35).visible = (Modo = 5)
  '  Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    
    ' ---- [20/10/2009] [LAURA] : a�adir del centro de coste
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    Me.txtAux2(8).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    BloquearTxt txtAux2(8), True
    
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim B As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
            
            
    If B Then
        'El trabajador debe existir
        CadenaSQL = ""
        If Text2(3).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Trabajador pedido"
        'Recogida
        If Text1(27).Text <> "" Then
            If Text2(27).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Direccion recogida de mercancia"
        End If
        'Solicitado por
        If Text1(23).Text <> "" Then
            If Text2(1).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Trabajador que solicita pedido"
        End If


        If CadenaSQL <> "" Then
            CadenaSQL = "Error en campos: " & vbCrLf & CadenaSQL
            MsgBox CadenaSQL, vbExclamation
            B = False
        End If
    End If
    CadenaSQL = ""
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim B As Boolean
'Dim devuelve As String
Dim i As Byte
Dim vArtic As CArticulo
Dim Aux As String
Dim TipoDto As Byte
    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(5), txtAux(6), TipoDto)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(7).Text Then txtAux(7).Text = Aux
    

    
    
    
    B = True
    'Comprobar que los campos NOT NULL tienen valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            If i = 8 And vEmpresa.TieneAnalitica = False Then
                'no hace nada pq puede ser nulo
            Else
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
        
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        B = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
'    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", txtAux(1).Text, "T", , "codalmac", txtAux(0).Text, "N")
'    If devuelve = "" Then
'        MsgBox "No existen unidades del Art�culo: " & txtAux(1).Text & "  en el Almacen: " & txtAux(0).Text, vbExclamation
'        b = False
'        PonerFoco txtAux(1)
'    End If
    
    DatosOkLinea = B
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Ampliaci�n linea y ENTER
        KeyAscii = 0
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case Else
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8 'Imprimir Pedido
             mnImpPedido_Click
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
Dim numlinea As String, vWhere As String
Dim cantidad As Currency
Dim J As Integer
Dim TipoDto  As Byte

On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        cantidad = ImporteFormateado(txtAux(3).Text)
        
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel,codccost) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0,"
        SQL = SQL & DBSet(txtAux(4).Text, "S") & "," & DBSet(txtAux(5).Text, "N") & ", "   'Mayo 2009   La "N" es ahora una "S"
        SQL = SQL & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
        SQL = SQL & DBSet(txtAux(7).Text, "N") & "," 'importe
        SQL = SQL & DBSet(txtAux(8).Text, "T", "S") 'centro coste
        SQL = SQL & ")"
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    
    
    'Si el articulo es de conjuntos, preguntara si quiere insertar la lineas de los conjuntos
    If InsertarLinea = True Then
        SQL = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", txtAux(1).Text, "T")
        If SQL = "1" Then
        
            'SI!!!!!!, es de conjuntos
            If MsgBox("Articulo con componentes. Desea insertar las lineas?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
            
            
            
            SQL = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            TipoDto = CByte(SQL)
            
            
            SQL = "Select sarti1.*,nomartic from sarti1,sartic where sarti1.codarti1=sartic.codartic and sarti1.codartic=" & DBSet(txtAux(1).Text, "T")
            Set miRsAux = New ADODB.Recordset
            'miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                'Limpiamos todo menos el almacen y el CC si lo tuviera
                For J = 1 To 7
                    txtAux(J).Text = ""
                Next
            
                txtAux(1).Text = miRsAux!codarti1
                txtAux(2).Text = miRsAux!NomArtic
                'Cantidad es la cantidad de la linea ppal * la del escandallo
                txtAux(3).Text = cantidad * miRsAux!cantidad
            
                ObtenerPrecioCompra
            
                
                txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            
            
            
                numlinea = Val(numlinea) + 1
                SQL = "INSERT INTO " & NomTablaLineas
                SQL = SQL & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel,codccost) "
                SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
                SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
                SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0,"
                SQL = SQL & DBSet(txtAux(4).Text, "S") & "," & DBSet(txtAux(5).Text, "N") & ", "   'Mayo 2009   La "N" es ahora una "S"
                SQL = SQL & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
                SQL = SQL & DBSet(txtAux(7).Text, "N") & "," 'importe
                SQL = SQL & DBSet(txtAux(8).Text, "T", "S") 'centro coste
                SQL = SQL & ")"
            
            
                If Not ejecutar(SQL, True) Then MsgBox "Error insertando articulo componente: " & miRsAux!codArtic & " " & miRsAux!NomArtic, vbExclamation
            
            
            
            
                miRsAux.MoveNext
            Wend
            
        End If
        
    End If 'insertar =OK
    
    
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Pedido" & vbCrLf & Err.Description
End Function








Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "S") & ", "   'MAYO 2009.  La "N" es ahora una "S"
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N") & ", "
        SQL = SQL & "codccost= " & DBSet(txtAux(8).Text, "T", "S")
        SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
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
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim B As Boolean
Dim SQL As String

On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez

    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not B
    PrimeraVez = False
    gridCargado = True
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Optional conServidas As Boolean)
Dim i As Byte
On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    i = 1
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                i = i + 1
                vDataGrid.Columns(i).Caption = "Alm."
                If conServidas Then
                    vDataGrid.Columns(i).Width = 450 + 100
                Else
                    vDataGrid.Columns(i).Width = 500 + 100
                End If
                vDataGrid.Columns(i).NumberFormat = "000"
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Art�culo"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1600 + 600
                Else
                    vDataGrid.Columns(i).Width = 1700 + 600
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Descripci�n"
                If conServidas Then
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(i).Width = 3100 + 800
                    Else
                        vDataGrid.Columns(i).Width = 3200 + 1200
                    End If
                Else
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(i).Width = 3400 + 1200
                    Else
                        vDataGrid.Columns(i).Width = 3600 + 1200
                    End If
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Ampl. L�nea"
                vDataGrid.Columns(i).Width = 7980
                vDataGrid.Columns(i).visible = False
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Cantidad"
                vDataGrid.Columns(i).Width = 900 + 300
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                i = i + 1
                If conServidas Then
                    'Cargar el grid con la columna de cantidad servida
                    vDataGrid.Columns(i).Caption = "Recibidas"
                    vDataGrid.Columns(i).Width = 1200
                    vDataGrid.Columns(i).Alignment = dbgRight
                    vDataGrid.Columns(i).NumberFormat = FormatoImporte
                    i = i + 1
                End If
                vDataGrid.Columns(i).Caption = "Precio"
                If vEmpresa.TieneAnalitica Then
                    vDataGrid.Columns(i).Width = 1100 + 300
                Else
                    vDataGrid.Columns(i).Width = 1200 + 300
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoPrecio2   'Mayo 2009
                
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.1"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 550 + 200
                Else
                    vDataGrid.Columns(i).Width = 600 + 200
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.2"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 550 + 200
                Else
                    vDataGrid.Columns(i).Width = 600 + 200
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
                i = i + 1
                vDataGrid.Columns(i).Caption = "Importe"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1100 + 500
                Else
                    vDataGrid.Columns(i).Width = 1300 + 500
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                
                '---- [19/10/2009] [LAURA] : se a�ade el centro de coste
                i = i + 1
                If vEmpresa.TieneAnalitica Then
                    vDataGrid.Columns(i).Caption = "CCost"
                    If conServidas Then
                        vDataGrid.Columns(i).Width = 650
                    Else
                        vDataGrid.Columns(i).Width = 700
                    End If
                Else
                    vDataGrid.Columns(i).visible = False
                End If
                
                i = i + 1
                vDataGrid.Columns(i).visible = False  'ampliacion
                
'                vDataGrid.Columns(i).Alignment = dbgRight
'                vDataGrid.Columns(i).NumberFormat = FormatoImporte
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    vDataGrid.RowHeight = 350
    
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                If i < 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 2).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                End If
                txtAux(i).Locked = False
            Next i
        End If
        
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True
        
        
        ' ---- [20/10/2009] [LAURA] : a�adir centro de coste
        BloquearTxt txtAux(8), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        Me.cmdAux(2).Enabled = Not txtAux(8).Locked
        Me.cmdAux(2).visible = Me.cmdAux(2).Enabled
        ' ----
        
        

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 30)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
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
        'Precio, Dto1, Dto2, Precio
        For i = 4 To txtAux.Count - 1
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 10
        Next i
        
        cmdAux(2).Left = txtAux(i - 1).Left + txtAux(i - 1).Width - cmdAux(2).Width
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 2
            txtAux(i).visible = visible
        Next i
        txtAux(8).visible = visible And vEmpresa.TieneAnalitica
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo RECIBIDAS de la tabla slippr
Dim alto As Single
Dim i As Byte

    i = 3
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(i).Top = 290
        txtAux(i).visible = visible
        txtAux(i).BackColor = vbWhite
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtAux(i).Text = ""
            BloquearTxt txtAux(i), False
'            txtAux(i).BackColor = &H80000013
        End If
      
        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        txtAux(i).Top = alto
        txtAux(i).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
        txtAux(i).Left = alto
        txtAux(i).Width = DataGrid1.Columns(7).Width - 15
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux(i).visible = visible
        PonerFoco txtAux(i)
    End If
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Generar Albaran
            mnGenAlbaran_Click
        Case 2
            mnGeneraDtos_Click
        Case 3
            'Simulacion en difierido ;)
            mnSimularProv_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    If Index = 3 Or Index = 4 Then
        If Modo <> 6 Then
            If Index = 3 Then
                lblF.Caption = "Ver articulo"
            Else
                lblF.Caption = "Ver precio"
            End If
        End If
    Else
        lblF.Caption = ""
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
            PonerServidas True
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub




Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Pasar de Pedido a Albaran
        ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
        If KeyCode = 113 Then
            If Index = 3 Then AbrirForm_Articulos
            If Index = 4 And txtAux(1).Text <> "" Then
                frmListadoPrecios.Opcion = 0
                frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                frmListadoPrecios.Show vbModal
            End If
        
        
         
        
        Else
          If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                If Index < 2 Or Index = 8 Then  'Para los que tienen busqueda
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
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAuxServidas True, True
                Else
                    If Data2.Recordset.BOF Then
                        PonerFoco txtAux(3)
                    Else
                        gridCargado = False
                        Data2.Recordset.MovePrevious
                        gridCargado = True
                        If Data2.Recordset.BOF Then Data2.Recordset.MoveFirst
                         If DataGrid1.Row > 0 Then
                            DataGrid1.Row = DataGrid1.Row - 1
                            CargaTxtAuxServidas True, True

                        End If
                    End If
                End If
                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
                ConseguirFoco txtAux(3), Modo
                
            Case 40 'Desplazamiento Flecha Hacia Abajo
'                If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
'                    DataGrid1.Row = DataGrid1.Row + 1
'                    CargaTxtAuxServidas True, True
'                Else
'                    PonerFocoBtn Me.cmdAceptar
'                End If
'                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
'                ConseguirFoco txtAux(3), Modo
                
                PonerServidas True
        End Select
    End If
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
'Dim vPrecio As CPreciosCom
Dim TipoDto As Byte
Dim B As Boolean


    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = ""
        Exit Sub
    End If

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

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
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , , devuelve) Then
                '---- [20/10/2009] [LAURA] : a�adir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtAux(8).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'por familia
                    txtAux(8).Text = devuelve
                    Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
                End If
                '----
                
                B = (Me.ActiveControl.Name = "txtAux")
                If B Then B = (Me.ActiveControl.Index = 0)
                
                If Not B Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
                
                
                If Modo = 5 And ModificaLineas = 2 Then
                    'Modificando. Ha cambiado el articulo
                    If txtAux(1).Text <> Data2.Recordset!codArtic Then
                         'Limpiamos
                         txtAux(4) = "": txtAux(5) = "": txtAux(6) = "": txtAux(7).Text = ""
                               
                    End If
                End If
                
                
            Else
            
                
                PonerFoco txtAux(Index)
            End If
            
            
'            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov) Then
'                If txtAux(2).Locked Then PonerFoco txtAux(3)
'                'Si es articulo de varios podemos modificar la descripci�n del articulo, sino bloqueamos.
''                If Not EsArticuloVarios(txtAux(Index).Text) Then
''                    BloquearTxt txtAux(2), True
''                Else
''                    BloquearTxt txtAux(2), False
''                    PonerFoco txtAux(2)
''                End If
'            Else
'                PonerFoco txtAux(Index)
'            End If
            
        Case 2 'Desc. Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                If (Modo = 5) And (ModificaLineas = 1 Or (ModificaLineas = 2 And txtAux(4).Text = "")) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                    
'                    Set vPrecio = New CPreciosCom
'                    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
'                        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
'                            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
'                            PonerFormatoDecimal txtAux(4), 2
'                            txtAux(5).Text = vPrecio.Descuento1
'                            PonerFormatoDecimal txtAux(5), 4
'                            txtAux(6).Text = vPrecio.Descuento2
'                            PonerFormatoDecimal txtAux(6), 4
'                        Else
'                            PonerFoco txtAux(Index)
'                        End If
'                    End If
'                    Set vPrecio = Nothing
                End If
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal_Single txtAux(Index), 9 'Tipo 9: Decimal(10,5)parametros
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
        Case 8 'COD. CENTRO DE COSTE
            ' ---- [20/10/2009] [LAURA]: a�adir centro de coste a la linea
            If txtAux(Index).Text = "" Then
                 txtAux2(Index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
                Me.txtAux2(Index).Text = PonerNombreCCoste(Me.txtAux(Index))
            End If
    End Select
    
    If Modo = 5 Then
         If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then 'Cant., Precio, Dto1, Dto2
'            If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'            If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
            If txtAux(1).Text = "" Then Exit Sub
            TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            PonerFormatoDecimal txtAux(7), 1
        End If
    End If
End Sub



Private Sub ObtenerPrecioCompra()
Dim vPrecio As CPreciosCom
Dim Cad As String
Dim aux2 As String

    On Error GoTo EPrecios
    
    Set vPrecio = New CPreciosCom
    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)    'FALTARA QUE DEVUELVE 5 decimales
'            PonerFormatoDecimal txtAux(4), 2
            txtAux(5).Text = vPrecio.Descuento1
'            PonerFormatoDecimal txtAux(5), 4
            txtAux(6).Text = vPrecio.Descuento2
'            PonerFormatoDecimal txtAux(6), 4
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
        Cad = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If Cad <> "" Then txtAux(4).Text = Cad
        
        'Septiembre 2010   'Descuentos
        vPrecio.CodigoArtic = txtAux(1).Text
        vPrecio.CodigoProve = Text1(4).Text
        Cad = vPrecio.ObtenerDescuentos2(Text1(1).Text, aux2)
        If Cad = "" Then Cad = "0"
        If aux2 = "" Then aux2 = "0"
        txtAux(5).Text = Cad
        txtAux(6).Text = aux2
    
    End If
    PonerFormatoDecimal_Single txtAux(4), 9   '10,5
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
    
    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = Cad
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
        DataGrid1_RowColChange 1, 1
End Sub


Private Function Eliminar() As Boolean
Dim B As Boolean
Dim vWhere As String
On Error GoTo FinEliminar

        conn.BeginTrans
         vWhere = ObtenerWhereCP(False)

'        If opt = 1 Then 'ELIMINAR
'            b = EliminarPedido(Data1.Recordset!numpedpr)
'        Else 'Pasar al HISTORICO
            B = ActualizarElTraspaso("", vWhere, CodTipoMov, CadenaSQL)
'        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido"
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = ObtenerWhereCP(False)
         vWhere = Replace(vWhere, NombreTabla & ".", "")
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
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
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


Private Function ObtenerWhereCP(conW As Boolean) As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String
On Error Resume Next
    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & NombreTabla & ".numpedpr= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NomTablaLineas & ".fecpedpr=" & DBSet(Text1(1).Text, "F")
    ObtenerWhereCP = SQL
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Optional conServidas As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numpedpr, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    If conServidas Then SQL = SQL & "recibida, "
'    SQL = SQL & "precioar, origpre, dtoline1, dtoline2,importel "
    SQL = SQL & "precioar, dtoline1, dtoline2,importel,codccost,ampliaci "
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fecpedpr='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numpedpr = -1"
    End If
    SQL = SQL & " Order by numpedpr, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim B As Boolean
Dim bAux As Boolean
Dim i As Integer

    B = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    'Insertar
    Toolbar1.Buttons(1).Enabled = (B Or Modo = 0) And Not EsHistorico
    Me.mnNuevo.Enabled = (B Or Modo = 0) And Not EsHistorico
    'Modificar
    Toolbar1.Buttons(2).Enabled = B And Not EsHistorico
    Me.mnModificar.Enabled = B And Not EsHistorico
    'eliminar
    Toolbar1.Buttons(3).Enabled = B And Not EsHistorico
    Me.mnEliminar.Enabled = B And Not EsHistorico
        
    B = (Modo = 2) And Not EsHistorico
    'Mantenimiento lineas
'        Toolbar1.Buttons(10).Enabled = (Modo = 2)
'        Me.mnLineas.Enabled = (Modo = 2)
    'Generar Albaran desde Pedido
    Toolbar5.Buttons(1).Enabled = B
    Me.mnGenAlbaran.Enabled = B
    
    'Octubre 2011
    'Modifica descuentos
    Toolbar5.Buttons(2).Enabled = B
    Me.mnGeneraDtos.Enabled = B
    
    'Simular
    Toolbar5.Buttons(3).Enabled = B
    Me.mnSimularProv.Enabled = B
    
    
    B = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B

    B = (Modo = 2) And Not EsHistorico
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i



End Sub


Private Function CargarDatosDirec(CodDirec As String, Indice As Byte) As Boolean
'Direcciones Propias
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim B As Boolean
On Error GoTo ECargarProve

    B = False
    If CodDirec <> "" Then
        devuelve = "Select nomdirec, domdirec, codpobla, pobdirec, prodirec "
        devuelve = devuelve & " FROM sdirpr Where coddirec=" & Val(CodDirec)
        
        Set RS = New ADODB.Recordset
        RS.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            Text1(Indice).Text = Format(CodDirec, "000")
            Text2(Indice).Text = RS.Fields!nomdirec 'Nom Direccion
            If Indice = 2 Then
                Indice = 21
            Else
                Indice = 17
            End If
            Text2(Indice).Text = RS.Fields!domdirec 'Domicilio
            Text2(Indice + 1).Text = RS.Fields!codpobla
            Text2(Indice + 2).Text = RS.Fields!pobdirec
            Text2(Indice + 3).Text = RS.Fields!prodirec
            B = True
        Else
            MsgBox "No existe la direcci�n: " & Text1(Indice).Text, vbInformation
            LimpiarDatosDirec (Indice)
        End If
        RS.Close
        Set RS = Nothing
    Else
        LimpiarDatosDirec (Indice)
        B = True
    End If
    
    CargarDatosDirec = B
    
ECargarProve:
    If Err.Number <> 0 Then CargarDatosDirec = False
End Function


Private Sub LimpiarDatosDirec(Indice As Byte)
    Text2(Indice).Text = ""
    If Indice = 2 Then
        Indice = 21
    Else
        Indice = 17
    End If
    Text2(Indice).Text = "" 'Domicilio
    Text2(Indice + 1).Text = "" 'cpostal
    Text2(Indice + 2).Text = "" 'poblacion
    Text2(Indice + 3).Text = "" 'provincia
End Sub


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numpedpr", "numpedpr", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador del Pedido."
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarPedido = True
        Else
            conn.RollbackTrans
            InsertarPedido = False
        End If
End Function


Private Sub LimpiarDatosProve()
'Limpia los campos del Form con datos del Proveedor
Dim i As Byte

    For i = 4 To 14
        Text1(i).Text = ""
    Next i
End Sub
    





Private Function PasarPedidoAAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'OUT -> numalb: Devuelve el N� de albaran asignado al pedido
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vWhere As String
Dim cProve As CProveedor

    On Error GoTo EGenPedido

    bol = False
            
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes Proveedor el Pedido  (scaalp, slialp)
    bol = InsertarAlbaran(MenError, NumAlb)
    
    
    
    
    
    
    'Para cada linea del pedido:
    ' Actualizar precio medio ponderado del articulo
    ' Actualizar precio y fecha ultima compra del articulo
    
    
    '17 Febrero 2011
    'LO QUITAMOS DE AQUI
'    If bol Then
'        MenError = "Actualizando Stocks"
'        bol = InsertarMovStock2(NumAlb, FechaAlb)
'    End If

    If bol Then
        'Actualizar la ult.fecha de compra del Proveedor
        MenError = "Actualizando ultima fecha compra en Proveedor."
        Set cProve = New CProveedor
        bol = cProve.ActualizaFechaUltCompra(Text1(4).Text, FechaAlb)
        Set cProve = Nothing
        
'        If bol Then
'            'Actualizar ult. fecha de compra y el precio ult compra de los articulos del Albaran
'            MenError = "Actualizando ultima fecha compra en Art�culos."
'            SQL = "numalbar=" & DBSet(NumAlb, "T") & " and fechaalb=" & DBSet(FechaAlb, "F") & " and slialp.codprove=" & Text1(4).Text
'            bol = ActualizarUltFechaCom(SQL)
'        End If
    End If
    
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminando cabecera y lineas del Pedido."
            bol = EliminarPedido(CLng(Text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-recibida y recibida= 0 en slippr
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
    End If
    
    
    
    If bol Then
        'si se ha generado correctamente el ALBARAN ver si hay alguna l�nea que tiene
        'el art�culo con control de n� de lote y pedir los n� de lotes.
        ComprobarNumLotesLineas NumAlb, FechaAlb
        
    End If
    
    
    
    
    If bol Then
        'Se ha generado correctamente el ALBARAN y vemos si tiene N� Series
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
        'Comprobar si Hay N� SERIE en compras y Mostrar
        'ventana para pedir los N� Serie de la cantidad introducida si lo requiere algun articulo
        ComprobarNSeriesLineas NumAlb, FechaAlb
        
        
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (slippr)
            MenError = "Eliminando lineas pedidido servidas completas."
            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            MenError = "Eliminando cabecera del pedido."
            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & Text1(0).Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text
                conn.Execute SQL
            End If
            RS.Close
            Set RS = Nothing
        End If
        bol = True
    End If
    
    
EGenPedido:
    If Err.Number <> 0 Then
'        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
'        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        PasarPedidoAAlbaran = True
    Else
        conn.RollbackTrans
        PasarPedidoAAlbaran = False
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function InsertarAlbaran(MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean
Dim vSQL As String
Dim FechaAlb As String
Dim FechaEntradaMercancia As String
Dim TrabAlb As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    NumAlb = RecuperaValor(CadenaSQL, 2)
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    TrabAlb = RecuperaValor(CadenaSQL, 1)
    FechaEntradaMercancia = RecuperaValor(CadenaSQL, 5)
    
    vSQL = "INSERT INTO scaalp (numalbar, fechaalb, codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, codtraba, codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, "
    vSQL = vSQL & " numpedpr, fecpedpr,codenvio,NReferencia,SReferencia,fecentrega,fentrada,codclien)"
    vSQL = vSQL & " SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, "
    vSQL = vSQL & "codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, "
    vSQL = vSQL & TrabAlb & " as codtraba,codtraba as codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5"
    vSQL = vSQL & " ,numpedpr, fecpedpr,codenvio,NReferencia,SReferencia,fecentrega"
    vSQL = vSQL & " , " & DBSet(FechaEntradaMercancia, "F") & " as fentrada   , codclien "
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedpr=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes Proveedor (scaalp)."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Albaran desde Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialp)."
    If Not InsertarLineasAlbaran(NumAlb, FechaAlb, FechaEntradaMercancia) Then Exit Function
    
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then
            bol = False
            MenError = MenError & vbCrLf & Err.Description
        End If
        If bol Then
            InsertarAlbaran = True
        Else
            InsertarAlbaran = False
        End If
End Function


Private Function InsertarLineasAlbaran(NumAlb As String, FechaAlb As String, FechaEntradaMercancia As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
'IN -> TipoM, numAlb
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte
'Dim InsertDirecto As Boolean
Dim cantidad As Currency
Dim ImpReciclado As Single
Dim numlinea As Integer
Dim ErrorFechaInventario As String
On Error GoTo EInsertarLinAlb

    
    'InsertDirecto = False
    'If AlbCompleto And vParamAplic.ArtReciclado = "" Then InsertDirecto = True
    'NUNCA PUEDE ENTRAR POR INSERT DIRECTO. Ya que los movimientos de almacen los hace aqui
    'If InsertDirecto Then
    '    'Insertar en la tabla de Albaran, los registros seleccionados de la tabla de Pedidos
    '    SQL2 = "SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, " & Val(Text1(4).Text) & " as codprove, numlinea, codartic, codalmac, nomartic, ampliaci, "
    '    SQL2 = SQL2 & "cantidad, precioar, dtoline1, dtoline2, importel,codccost "
    '    SQL2 = SQL2 & " FROM " & NomTablaLineas & " WHERE numpedpr=" & Val(Text1(0).Text)
    '    SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) " & SQL2
    '    conn.Execute SQL2, , adCmdText
    'Else
    
        'NO insert directo.
        'Es o bien pq no es completio o pq tiene tasa reciclado
        SQL2 = "select * from " & NomTablaLineas
        SQL2 = SQL2 & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        SQL2 = SQL2 & " ORDER BY numlinea"
        Set RS = New ADODB.Recordset
        RS.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        numlinea = 1
        While Not RS.EOF 'Para cada linea de pedido insertar una de albaran si recibidas >0
            SQL2 = ""
            If AlbCompleto Then
                'Va la linea entera
                SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) "
                SQL2 = SQL2 & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", "
                SQL2 = SQL2 & DBSet(RS!codArtic, "T") & "," & RS!codAlmac & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                SQL2 = SQL2 & DBSet(RS!cantidad, "N") & ", " & DBSet(RS!precioar, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                SQL2 = SQL2 & DBSet(RS!ImporteL, "N") & ","
                SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ")"
                cantidad = RS!cantidad
                ImpLinea = RS!ImporteL
            Else
                If RS!recibida > 0 Then
                    TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
                    ImpLinea = CalcularImporte(RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
                    SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) "
                    SQL2 = SQL2 & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", "
                    SQL2 = SQL2 & DBSet(RS!codArtic, "T") & "," & RS!codAlmac & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                    SQL2 = SQL2 & DBSet(RS!recibida, "N") & ", " & DBSet(RS!precioar, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                    SQL2 = SQL2 & DBSet(ImpLinea, "N") & ","
                    SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ")"
                    cantidad = RS!recibida
                End If
            End If
            
            
            
            If SQL2 <> "" Then
                
                conn.Execute SQL2, , adCmdText
                
                
                'AQui habria que hacer lo del stock
                'InsertarMovStock3 NumAlb, FechaAlb, numlinea, cantidad, CCur(ImpLinea), RS!codAlmac, RS!codArtic
                InsertarMovStock3 NumAlb, FechaEntradaMercancia, numlinea, cantidad, CCur(ImpLinea), RS!codAlmac, RS!codArtic
                
                
                
                
                numlinea = numlinea + 1
                'TASA RECILCADO
                If vParamAplic.ArtReciclado <> "" Then
                    If ArticuloConTasaReciclado(CStr(RS!codArtic), ImpReciclado) Then
                        ImpLinea = Round2(cantidad * ImpReciclado, 2)
                        SQL2 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                        'OCTUBRE 2011
                        'Error. Ponia rs!codartic en lugar de artrecicla: SQL2 = numlinea & ", " & DBSet(RS!codArtic, "T") .....
                        SQL2 = numlinea & ", " & DBSet(vParamAplic.ArtReciclado, "T") & "," & RS!codAlmac & ", " & DBSet(SQL2, "T") & ", " & DBSet("", "T") & ", "
                        SQL2 = " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & SQL2
                        'SQL2 = SQL2 & DBSet(Cantidad, "N") & ", " & DBSet(ImpReciclado, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                        SQL2 = SQL2 & DBSet(cantidad, "N") & ", " & DBSet(ImpReciclado, "S") & ",0,0,"
                        SQL2 = SQL2 & DBSet(ImpLinea, "N") & ","
                        SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ")"
                        SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, " & _
                            "cantidad, precioar, dtoline1, dtoline2, importel,codccost) " & SQL2
                        
                        conn.Execute SQL2
                        numlinea = numlinea + 1
            
                    End If
                End If
                
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    'End If
    
EInsertarLinAlb:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
        MuestraError Err.Number, "Insertar lineas albaran.", Err.Description
    Else
        InsertarLineasAlbaran = True
    End If
End Function



Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String
On Error GoTo EEliminarPed

     SQL = " WHERE  numpedpr=" & numPed

    'Lineas de Pedido
    conn.Execute "Delete from " & NomTablaLineas & SQL
        
    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function


Private Function ActualizarPedido() As Boolean
'Actualiza la tabla de lineas de pedido (sliped)
'cantidad=cantidad-servidas y servidas=0
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte

    On Error GoTo EActPedido

    SQL = "select numlinea, codalmac, codartic, cantidad, recibida, precioar, dtoline1, dtoline2 from " & NomTablaLineas
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF 'Para cada linea
        TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
        ImpLinea = CalcularImporte(RS!cantidad - RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
        SQL = "UPDATE " & NomTablaLineas & " SET cantidad=cantidad-recibida, recibida=0, importel=" & DBSet(ImpLinea, "N")
'        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic='" & RS!codArtic & "'"
        SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        SQL = SQL & " AND numlinea=" & RS!numlinea
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EActPedido:
    If Err.Number <> 0 Then
        ActualizarPedido = False
    Else
        ActualizarPedido = True
    End If
End Function


Private Function ActualizarCabPedido() As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "UPDATE " & NombreTabla & " SET restoped=1 " & ObtenerWhereCP(True)
    conn.Execute SQL
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock3(NumAlb As String, FechaAlb As String, NLin As Integer, cantidad As Currency, Importe As Currency, codAlmac As Integer, codArtic As String) As Boolean
Dim vCStock As CStock
Dim B As Boolean
Dim SQL As String
Dim cart As CArticulo


    'No lleva error, que salte en la rutina ppal
    On Error Resume Next

    InsertarMovStock3 = False
    
    Set vCStock = New CStock
    B = True
   
    vCStock.FechaMov = FechaAlb
    vCStock.tipoMov = "E"
    vCStock.DetaMov = "ALC"
    vCStock.Trabajador = CLng(Text1(4).Text) 'En codigope ponemos el Proveedor
    vCStock.codArtic = codArtic
    vCStock.codAlmac = CInt(codAlmac)
    
    
    vCStock.cantidad = CSng(cantidad)
    vCStock.Importe = CCur(Importe)
    
    
    vCStock.LineaDocu = NLin
    vCStock.Documento = NumAlb
    If vCStock.cantidad <> 0 Then
        '==== Laura 22/09/2006
        '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
        Set cart = New CArticulo
        If cart.LeerDatos(vCStock.codArtic) Then
            '17 Junio 2009
            'Si la cantidad es negativa no actualiza ni precio medio ponderado NI fecha ult compra
            If vCStock.cantidad >= 0 Then
            
                'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
                'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                If Not cart.ActualizarPrecioMedPond(CCur(vCStock.cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.cantidad), 4)) Then B = False
                
                '--actualizar fecha y precio ultima compra del articulo
                'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
                'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
                If Not cart.ActualizarUltFechaCompra(vCStock.FechaMov, Round2(CCur(vCStock.Importe) / CCur(vCStock.cantidad), 4)) Then B = False


                

            End If 'De cantidad >=0
        End If
        Set cart = Nothing
        '====
    
    
        'en actualizar stock comprobamos si el articulo tiene control de stock
        B = vCStock.ActualizarStock
    
    Else
        B = True  'Si no inserta pq la cantidad es cero n pasa nada
    End If
    InsertarMovStock3 = B
    
End Function

Private Sub ImprimirAlbaran(Numalbar As String, FechaAlb As String, Codprove As Long)
Dim cadNomRPT As String
Dim SQL As String
Dim numP As Byte
Dim param As String

    
    
    'Albaran socio
    If Not PonerParamRPT2(27, param, numP, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    
    
    



    
    
    SQL = CadenaDesdeHasta(CStr(FechaAlb), CStr(FechaAlb), "{scaalp.fechaalb}", "F")
    SQL = SQL & " AND  {scaalp.codprove} = " & Codprove
    SQL = SQL & " AND  {scaalp.numalbar} = """ & DevNombreSQL(Numalbar) & """"
    



    
     With frmImprimir
        .FormulaSeleccion = SQL
        .OtrosParametros = param
        .NumeroParametros = numP
        .SeleccionaRPTCodigo = pRptvMultiInforme
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 + 10   '2000 mas la opcion de entrada
        .NombrePDF = ""
        '.NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Function ActualizarServidas() As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "UPDATE " & NomTablaLineas & " SET recibida= " & DBSet(txtAux(3).Text, "N")
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas(HaciaAlante As Boolean)
Dim NumFila As Integer
Dim cadMen As String

'    NumFila = DataGrid1.Row
    NumFila = Data2.Recordset.AbsolutePosition
    If PonerFormatoDecimal(txtAux(3), 1) Then  'Tipo 1: Decimal(12,2)
        If CCur(txtAux(3).Text) > Data2.Recordset!cantidad Then
            cadMen = "La cantidad a Recibir no puede ser superior a la del Pedido."
            MsgBox cadMen, vbExclamation
            PonerFoco txtAux(3)
            Exit Sub
        End If
    End If
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
'    DataGrid1.Row = NumFila
    SituarDataPosicion Data2, CLng(NumFila), ""
    If HaciaAlante Then MoverSigRegistro
End Sub




Private Sub MoverSigRegistro()
    On Error GoTo EMover
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
        DataGrid1.Row = DataGrid1.Row + 1
        CargaTxtAuxServidas True, True
    Else
        PonerFocoBtn Me.cmdAceptar
    End If
    txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    PonerFoco txtAux(3)
    ConseguirFocoLin txtAux(3)
    txtAux(3).Refresh
EMover:
    If Err.Number <> 0 Then MuestraError Err.Description, "Mover registro.", Err.Description
End Sub





Private Sub GenerarAlbaran()
Dim numPed As Long 'N� Pedido
Dim NumAlb As String 'N� Albaran
Dim FechaAlb As String 'Fecha del Albaran
Dim FEntradaMercancia As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim ImprimeAlb As Long   'Si queremos imprimir guardare el codprove
Dim ArticuloEsEscandallo As String



    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedpr
    
    'pedir por pantalla: el operador, N� albaran y fecha albaran
    Set frmList = New frmListadoOfer
    frmList.OpcionListado = 57
    frmList.codClien = Me.Text1(4).Text
    CadenaSQL = ""
    frmList.Show vbModal
    Set frmList = Nothing
    
    If CadenaSQL = "" Then Exit Sub
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    SQL = RecuperaValor(CadenaSQL, 4)
    ImprimeAlb = -1
    If SQL = "1" Then ImprimeAlb = CLng(Text1(4).Text)
    FEntradaMercancia = RecuperaValor(CadenaSQL, 5) 'fec
    
    


    
    'Mostraremos un msg si algunos de los articulos tienen fecha inventario posterior
    SQL = "SELECT  codalmac,salmac.codartic,nomartic,fechainv FROM salmac,sartic where salmac.codartic=sartic.codartic and artvario=0 and "
    SQL = SQL & " fechainv > " & DBSet(FEntradaMercancia, "F")    ' DBSet(FechaAlb, "F")
    SQL = SQL & " and (codalmac,salmac.codartic) in ("
    SQL = SQL & " select codalmac,codartic from slippr WHERE numpedpr=" & numPed
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then SQL = SQL & " and slippr.recibida>0 "
    SQL = SQL & ")"
    B = ObtenerRSprecios(RS, SQL)
    SQL = ""
    If Not B Then
        MsgBox "Error obteniendo datos cruzados con inventarios", vbExclamation
    Else
        If Not RS.EOF Then
            
            While Not RS.EOF
                SQL = SQL & "   -" & RS!codArtic & "  " & RS!NomArtic & "   inventariado el " & RS!FechaINV & vbCrLf
                RS.MoveNext
            Wend
            
            
            If SQL <> "" Then
                SQL = "Las siguientes referencias tiene fecha inventario posterior al del albaran:" & vbCrLf & vbCrLf & SQL
                SQL = SQL & vbCrLf & "�Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then SQL = ""
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    If SQL <> "" Then Exit Sub
    
    
    'Antes de pasar el pedido al albaran nos guardamos los articulos cuyo precio_compra
    'se han modificado para preguntar despues si se quiere actualizar precios_venta
    'hay q guardarlo antes de pasar pedido a albaran ya q aqui se actualiza el precio_ult_compra
    '-- Laura 19/12/2006: calcular precio_med_pond con el precio aplicados los descuentos, ed. importe/cantidad
    ' Iremos cambiando el numero de decimales poc a poco ANTES era un 4
    SQL = "SELECT slippr.codartic,sartic.nomartic,round(slippr.importel/slippr.cantidad," & PrecioDecimales & ")"
    SQL = SQL & " as precioar,sartic.preciouc,sum(cantidad) "
    SQL = SQL & " FROM slippr INNER JOIN sartic ON slippr.codartic=sartic.codartic "
    'SQL = SQL & " WHERE numpedpr=" & numPed & " and (slippr.precioar<>sartic.preciouc)"
    SQL = SQL & " WHERE numpedpr=" & numPed & " and (round(slippr.importel/slippr.cantidad,4)<>sartic.preciouc)"
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then SQL = SQL & " and slippr.recibida>0 "
    SQL = SQL & " group by slippr.codartic,slippr.precioar,sartic.preciouc "
    SQL = SQL & " Having Sum(Cantidad) > 0"
    B = ObtenerRSprecios(RS, SQL)
    
    
    
    If PasarPedidoAAlbaran(NumAlb, FechaAlb) Then
        'Imprime los pedidos de cliente vinculados con los articulos del albaran de proveedor generado
        If Not ComprobarPedidosClientesDesdeAlbProveedor(NumAlb, CDate(FechaAlb), Text1(4).Text) Then MsgBox "Se ha generado correctamente el Albaran: " & NumAlb, vbInformation
                

        PonerModo 2
        
        
        'comprobar si hay lineas de art�culos cuyo precio_ultima_compra
        'se ha modificado y preguntar si que quieren actualizar los precio_venta
        '--------------------------------------------------------
        If B Then
            ArticuloEsEscandallo = ""
            While Not RS.EOF
            
                'Primero compruebo si es escandallo de otro select count(*) from ariges3.sarti1 where codarti1='0020080939'
                SQL = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", RS!codArtic, "T")
                If SQL <> "" Then
                    If Val(SQL) > 0 Then
                        ArticuloEsEscandallo = ArticuloEsEscandallo & RS!codArtic & "|"
                    End If
                End If
                SQL = "Se ha modificado el precio �ltima compra del art�culo:" & vbCrLf
                SQL = SQL & RS!codArtic & ":  " & RS!NomArtic & vbCrLf
                SQL = SQL & vbCrLf & "�Desea actualizar los precios de venta?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                
                
                    
                
                    'Comprobar que el art�culo tiene margen comercial
                    If ArticuloTieneMargen(RS!codArtic) Then
                        'Aplicar margen comercial a los precios
                        'Modificar precios de venta en articulo y tarifas
                        frmComActPrecios.parCodArtic = RS!codArtic
                        frmComActPrecios.parNomArtic = RS!NomArtic
                        frmComActPrecios.Show vbModal
                    End If
                Else
                     
                    If vParamAplic.RecalculoMargen Then ActualizacionAutomaticaMargen RS!codArtic
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
            
            
            If ArticuloEsEscandallo <> "" Then
                frmListado4.Opcion = 1
                frmListado4.vCadena = ArticuloEsEscandallo
                frmListado4.Show vbModal
            End If
            
            
            
        End If
        
       
        
        
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            SQL = DevuelveDesdeBDNew(conAri, "scappr", "numpedpr", "numpedpr", Text1(0).Text, "N")
            If SQL = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                PonerCampos
                CargaGrid DataGrid1, Data2, True, False
            End If
            CargaTxtAuxServidas False, False
        
            'Eliminar las filas del pedido que se servieron completas (slippr)
'            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
'            Conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
'            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & numPed
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
'                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & numPed
'                Conn.Execute SQL
'                PosicionarDataTrasEliminar
'            Else 'Quedan lineas en el pedido --> Actualizar las lineas
'                PosicionarData
'                PonerCampos
'                CargaGrid DataGrid1, Data2, True, False
'            End If
'            RS.Close
'            Set RS = Nothing
'            CargaTxtAuxServidas False, False
        End If
       
        
'        Imprimer albaran si se solicit�
        If ImprimeAlb >= 0 Then ImprimirAlbaran NumAlb, FechaAlb, ImprimeAlb
        Screen.MousePointer = vbDefault
    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim SQL As String
    On Error Resume Next
    SQL = "UPDATE " & NomTablaLineas & " SET recibida= 0 "
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    conn.Execute SQL
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub ComprobarNumLotesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de N� Lotes si hay algun articulo en las lineas de pedido que
'requiere N� de lote en compras pedirlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String

    On Error GoTo ErrLotes

    cadWhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
    cadWhere = cadWhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
    cadWhere = cadWhere & " slialp.codprove=" & Text1(4).Text

    'seleccionamos aquellas lineas del albaran insertado que tengan control de lote
    SQL = "SELECT slialp.* "
    SQL = SQL & " FROM (slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic) "
    SQL = SQL & " LEFT OUTER JOIN scateg ON sartic.codcateg=scateg.codcateg "
    SQL = SQL & cadWhere
    SQL = SQL & " AND scateg.ctrlotes = 1"


    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSLineas.EOF Then
        'Comprobar si NO Hay N� SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los N� Serie de la cantidad introducida
'        Me.cmdAux(1).Tag = NumAlb
'        Me.cmdAux(0).Tag = FechaAlb
        PedirNLotes RSLineas
    
'        Set frmNLote = New frmAlmCargarNLote
'        frmNLote.parSQL = SQL
'        frmNLote.Show vbModal
'        Set frmNLote = Nothing

    End If
    
    RSLineas.Close
    Set RSLineas = Nothing
    Exit Sub

ErrLotes:
    MuestraError Err.Number, "Pedir N� de lote.", Err.Description
End Sub




Private Sub ComprobarNSeriesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de N� Series si hay algun articulo en las lineas de pedido que requiere N� de serie
'y hay control de N� de serie en compras pedirlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
        
    If vParamAplic.NumSeries Then 'So control de N� Series en COMPRAS
        cadWhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
        cadWhere = cadWhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
        cadWhere = cadWhere & " slialp.codprove=" & Text1(4).Text
        
        'Seleccionamos aquellas lineas de albaran que tienen N� de Serie
        SQL = "SELECT slialp.codartic, sum(cantidad) as cantidad, slialp.numlinea "
        SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " And nseriesn = 1 "
        SQL = SQL & " GROUP BY codartic ORDER BY Codartic "
    
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RSLineas.EOF Then
            'Comprobar si NO Hay N� SERIE en Compras y si no se realizo alli
            'Mostrar ahora ventana para pedir los N� Serie de la cantidad introducida
            Me.cmdAux(1).Tag = NumAlb
            Me.cmdAux(0).Tag = FechaAlb
            PedirNSeries RSLineas
        End If
        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
On Error GoTo EPedirNSeries
        
        'Visualizar en pantalla el Grid, y rellenar los N� Serie
        PedirNSeriesGnral RS, True

        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = False 'Se llama desde Alb. de Venta
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub PedirNLotes(ByRef RS As ADODB.Recordset)
Dim cadSel As String

    On Error GoTo EPedirNLotes
        
    cadSel = "numalbar=" & DBSet(RS!Numalbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND codprove=" & DBSet(RS!Codprove, "N")
    
    'Visualizar en pantalla el Grid, y rellenar los N� Serie
    If Not PedirNLotesGnral(RS, True) Then
'             Visualizar en pantalla el Grid, y rellenar los N� Serie
        MsgBox "No se han podido mostrar todos los Art�culos con N� de Lote.", vbInformation
    End If

        Set frmNLote = New frmAlmCargarNLote
        frmNLote.Desde2 = "" 'Desde proveedores
        frmNLote.parSelSQL = cadSel
        frmNLote.Show vbModal
        Set frmNLote = Nothing
        
        
     'Eliminar de la tabla temporal tmpnlotes los lotes introducidos
    DescargarDatosTMPNumLotes "tmpnlotes", cadSel
        
EPedirNLotes:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de N� Serie
'Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim Numalbar As String
Dim nSerie As CNumSerie
Dim B As Boolean

    On Error GoTo EInsertarNS

    Set nSerie = New CNumSerie
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    nSerie.Proveedor = CInt(Text1(4).Text)
    nSerie.NumAlbProve = Me.cmdAux(1).Tag
    nSerie.fechacom = Me.cmdAux(0).Tag
    nSerie.NumLinAlbPr = numlinea
    'calculamos la fecha de fin garantia para el articulo comprado
    nSerie.ObtenFechaFinGarantia codArtic, Me.cmdAux(0).Tag
    
    'Comprobar si existe en la tabla sserie
    Numalbar = "numalbpr" 'N� albaran de Compra
    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", Numalbar, "codartic", codArtic, "T")
    If devuelve <> "" Then 'EXISTE en tabla sserie
        If Numalbar = "" Then
            B = nSerie.ActualizarNumSerie(False)
        End If
    Else
        B = nSerie.InsertarNumSerie
    End If
    Set nSerie = Nothing
    
EInsertarNS:
    If Err.Number <> 0 Then B = False
    If Not B Then
        InsertarNSerie = False
    Else
        InsertarNSerie = True
    End If
End Function



Private Sub PonerDatosProveedor(Codprove As String, Optional nifProve As String)
'lee de la tabla de proveedores y pone los valores
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
            If Modo = 3 Then
                Observaciones = ""
            Else
                'Modificando y no ha puesto nada
                Observaciones = Trim(Text1(17) & Text1(18) & Text1(19) & Text1(20) & Text1(21))
            End If
            If Observaciones = "" Then vProve.PonerObservaciones Text1(17), Text1(18), Text1(19), Text1(20), Text1(21)


            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
            
            If Modo = 3 Then
                'Insertando
                If Not EsDeVarios Then
                    PonerFocoChk Me.chkObra
                Else
                    PonerFoco Text1(5)
                End If
            End If
        End If
    Else
        LimpiarDatosProve
        PonerFoco Text1(4)
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim B As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    B = vProve.LeerDatosProveVario(nifProve)
    
    If B Then
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
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea proveedor de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(6).visible = bol 'NIF
        Me.imgBuscar(6).Enabled = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'poblacion
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
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
        'Actualiza la tabla de proveedores varios con los datos que tenemos
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
Dim i As Byte
Dim cadWhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 50
         Text3(i).Text = ""
    Next i
    
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
Dim i As Byte

    For i = 33 To 36
        If i = 34 Or i = 35 Then Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 And Text3(i - 6).Text = "" Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'Los Totales
    For i = 49 To 50
'        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
End Sub




Private Function ActualizarUltFechaCom(cadW As String) As Boolean
''Actualiza la ultima fecha de compra y el ult. precio de compra
''en el articulo, poniendo los valores del albaran de compra
'Dim SQL As String
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EActualizaFecha
'
'    SQL = "select distinct numalbar,fechaalb,slialp.codartic,max(slialp.precioar) as precioar , sartic.ultfecco "
'    SQL = SQL & " from slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic "
''    SQL = SQL & " where numalbar='K2500088' and fechaalb='2005-10-06' and slialp.codprove=21"
'    SQL = SQL & " WHERE " & cadW
'    SQL = SQL & " and (fechaalb>ultfecco or isnull(ultfecco))"
'    SQL = SQL & " group by numalbar,fechaalb,slialp.codartic "
'    SQL = SQL & " order by codartic "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!FechaAlb, "F") & ", preciouc=" & DBSet(RS!precioar, "N")
'        SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
'        Conn.Execute SQL
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'EActualizaFecha:
'    If Err.Number <> 0 Then
'        ActualizarUltFechaCom = False
'    Else
'        ActualizarUltFechaCom = True
'    End If
End Function



Private Function ObtenerRSprecios(ByRef RS As ADODB.Recordset, cadSQL As String) As Boolean
    On Error GoTo ErrRS
    Set RS = New ADODB.Recordset
    RS.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ObtenerRSprecios = True
    Exit Function
    
ErrRS:
    ObtenerRSprecios = False
    If Not RS Is Nothing Then Set RS = Nothing
    MuestraError Err.Number, "Cargando RS precios ultima compra.", Err.Description
End Function




Private Sub AbrirForm_CentroCoste()
'    Screen.MousePointer = vbHourglass
'    cmdAux(2).Tag = "2"
'
'    Set frmB = New frmBuscaGrid
'    If vParamAplic.ContabilidadNueva Then
'        frmB.vCampos = "Codigo|ccoste|codccost|T||20�Descripci�n|ccoste|nomccost|T||70�"
'        frmB.vTabla = "ccoste"
'    Else
'        frmB.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
'        frmB.vTabla = "cabccost"
'    End If
'    frmB.vSQL = ""
'    HaDevueltoDatos = False
'    '###A mano
'    frmB.vDevuelve = "0|1|"
'    frmB.vTitulo = "Centros de coste"
'    frmB.vselElem = 0
'    frmB.vConexionGrid = conConta
'
'    frmB.Show vbModal
'    Set frmB = Nothing
'    cmdAux(2).Tag = "-1"
    cmdAux(2).Tag = "2"

    Set frmB = New frmBasico2
    AyudaCentroCoste frmB, txtAux(8)
    Set frmB = Nothing

    cmdAux(2).Tag = "-1"
End Sub



' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
Private Sub AbrirForm_Articulos()
    If Trim(txtAux(1).Text) = "" Then Exit Sub
    
    Set FrmArt2 = New frmAlmArticulosGr
    FrmArt2.DatosADevolverBusqueda = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
    FrmArt2.parNumTAb = 6
    FrmArt2.Show vbModal
    Set FrmArt2 = Nothing
End Sub
' -----


'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        If imgBuscar(Index).visible Then imgBuscar_Click Index
        
    Else
        'Lineas
        If Index = 8 Then Index = 2
        cmdAux_Click Index
        
        
    End If
        
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



'Private Function InsertarMovStock3(NumAlb As String, FechaAlb As String) As Boolean
'Dim vCStock As CStock
'Dim b As Boolean
'Dim RS As ADODB.Recordset
'Dim SQL As String
'Dim cart As CArticulo
'
'    On Error Resume Next
'
'    InsertarMovStock2 = False
'
'    Set vCStock = New CStock
'    b = True
'
'    SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
'    SQL = "select * from " & NomTablaLineas & SQL
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    vCStock.Fechamov = FechaAlb
'
'    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
'    While (Not RS.EOF) And b
'        If InicializarCStockAlbar(vCStock, "E", CStr(RS!numlinea), RS) Then
'            vCStock.Documento = NumAlb
'            If vCStock.Cantidad <> 0 Then
'                '==== Laura 22/09/2006
'                '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
'                Set cart = New CArticulo
'                If cart.LeerDatos(vCStock.codArtic) Then
'                    '17 Junio 2009
'                    'Si la cantidad es negativa no actualiza ni precio medio ponderado NI fecha ult compra
'                    If vCStock.Cantidad >= 0 Then
'
'                        'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
'                        'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
'                        If Not cart.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
'
'                        '--actualizar fecha y precio ultima compra del articulo
'                        'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
'                        'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
'                        If Not cart.ActualizarUltFechaCompra(vCStock.Fechamov, Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
'
'                    End If 'De cantidad >=0
'                End If
'                Set cart = Nothing
'                '====
'
'
'                'en actualizar stock comprobamos si el articulo tiene control de stock
'                b = vCStock.ActualizarStock
'            End If
'        Else
'            b = False
'        End If
'        RS.MoveNext
'    Wend
'    Set vCStock = Nothing
'    RS.Close
'    Set RS = Nothing
'
'    InsertarMovStock2 = b
'
'End Function





Private Function ActualizarDtos() As Boolean
Dim SQL As String
Dim TipoDto As Byte
Dim Dto1 As Currency
Dim Dto2 As Currency

On Error GoTo eActualizarDtos

         TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
         SQL = "SELECT numlinea,cantidad, precioar FROM " & NomTablaLineas
         SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " ORDER BY numlinea"
         Set miRsAux = New ADODB.Recordset
         miRsAux.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
         
         While Not miRsAux.EOF
         
            SQL = CalcularImporteSng(CStr(miRsAux!cantidad), CStr(miRsAux!precioar), RecuperaValor(CadenaDesdeOtroForm, 1), RecuperaValor(CadenaDesdeOtroForm, 2), TipoDto)
            'Ya tengo el importe
            SQL = "UPDATE " & NomTablaLineas & " SET importel = " & DBSet(SQL, "N")
            SQL = SQL & ", dtoline1=" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "N")
            SQL = SQL & ", dtoline2=" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "N")
            SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            SQL = SQL & " AND numlinea = " & miRsAux!numlinea
            conn.Execute SQL
            
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        ActualizarDtos = True 'Ok
eActualizarDtos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Function



