VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComHcoFacturas2GR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Proveedores"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   15315
   Icon            =   "frmComHcoFacturasGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   125
      Top             =   45
      Width           =   3165
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   126
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
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
      Left            =   3360
      TabIndex        =   123
      Top             =   45
      Width           =   1920
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   124
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Vincular albaranes"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Fras Pendientes de contabilizar"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5400
      TabIndex        =   121
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   122
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
      Left            =   13590
      TabIndex        =   120
      Top             =   270
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   102
      Top             =   795
      Width           =   15035
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
         Left            =   4470
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   495
         Width           =   1340
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
         Left            =   9075
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scafpc|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   495
         Width           =   5745
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
         Left            =   8010
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   495
         Width           =   1005
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
         Left            =   2925
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   495
         Width           =   1340
      End
      Begin VB.TextBox Text1 
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
         MaxLength       =   20
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "00000000000000000000"
         Top             =   495
         Width           =   2565
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizada"
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
         Left            =   6015
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|scafpc|intconta||N|"
         Top             =   540
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   9090
         Picture         =   "frmComHcoFacturasGR.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepción"
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
         Left            =   4470
         TabIndex        =   110
         Top             =   210
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
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
         Index           =   29
         Left            =   2925
         TabIndex        =   106
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Left            =   240
         TabIndex        =   105
         Top             =   210
         Width           =   1395
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
         Left            =   8010
         TabIndex        =   107
         Top             =   210
         Width           =   1050
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   120
      Top             =   4080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   13440
      Top             =   3480
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
      Height          =   8355
      Left            =   135
      TabIndex        =   44
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1845
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   14737
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmComHcoFacturasGR.frx":0A0E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmComHcoFacturasGR.frx":0A2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2(3)"
      Tab(1).Control(1)=   "Text3(12)"
      Tab(1).Control(2)=   "Text3(1)"
      Tab(1).Control(3)=   "Text3(0)"
      Tab(1).Control(4)=   "Text3(9)"
      Tab(1).Control(5)=   "Text3(10)"
      Tab(1).Control(6)=   "FrameObserva"
      Tab(1).Control(7)=   "Text2(16)"
      Tab(1).Control(8)=   "Text2(17)"
      Tab(1).Control(9)=   "txtAux2(8)"
      Tab(1).Control(10)=   "FrameToolAux0"
      Tab(1).Control(11)=   "txtAux3(2)"
      Tab(1).Control(12)=   "txtAux3(4)"
      Tab(1).Control(13)=   "txtAux3(3)"
      Tab(1).Control(14)=   "Text2(2)"
      Tab(1).Control(15)=   "chkDocArchi"
      Tab(1).Control(16)=   "txtAux(8)"
      Tab(1).Control(17)=   "txtAux3(1)"
      Tab(1).Control(18)=   "txtAux3(0)"
      Tab(1).Control(19)=   "txtAux(3)"
      Tab(1).Control(20)=   "txtAux(2)"
      Tab(1).Control(21)=   "txtAux(1)"
      Tab(1).Control(22)=   "txtAux(0)"
      Tab(1).Control(23)=   "Text2(0)"
      Tab(1).Control(24)=   "Text2(1)"
      Tab(1).Control(25)=   "txtAux(4)"
      Tab(1).Control(26)=   "txtAux(5)"
      Tab(1).Control(27)=   "txtAux(6)"
      Tab(1).Control(28)=   "txtAux(7)"
      Tab(1).Control(29)=   "DataGrid1"
      Tab(1).Control(30)=   "DataGrid2"
      Tab(1).Control(31)=   "cmdObserva"
      Tab(1).Control(32)=   "Text3(3)"
      Tab(1).Control(33)=   "Text3(2)"
      Tab(1).Control(34)=   "Text3(11)"
      Tab(1).Control(35)=   "Label1(24)"
      Tab(1).Control(36)=   "Label1(35)"
      Tab(1).Control(37)=   "Label1(3)"
      Tab(1).Control(38)=   "Label1(46)"
      Tab(1).Control(39)=   "Label1(22)"
      Tab(1).Control(40)=   "imgBuscar(7)"
      Tab(1).Control(41)=   "Label1(13)"
      Tab(1).Control(42)=   "Label1(8)"
      Tab(1).Control(43)=   "Label1(21)"
      Tab(1).Control(44)=   "imgBuscar(5)"
      Tab(1).Control(45)=   "Label1(9)"
      Tab(1).Control(46)=   "imgBuscar(6)"
      Tab(1).Control(47)=   "Label1(18)"
      Tab(1).Control(48)=   "Label1(6)"
      Tab(1).ControlCount=   49
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
         Left            =   -64320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   135
         Text            =   "Text2"
         Top             =   1704
         Width           =   4095
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
         Index           =   12
         Left            =   -65160
         MaxLength       =   30
         TabIndex        =   101
         Tag             =   "Forma de envio|N|S|0||scafpa|codclien|0000|N|"
         Text            =   "Text1"
         Top             =   1704
         Width           =   780
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
         Index           =   1
         Left            =   -65160
         MaxLength       =   30
         TabIndex        =   99
         Tag             =   "Trabajador pedido|N|S|0|9999|scafpa|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   888
         Width           =   780
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
         Index           =   0
         Left            =   -65160
         MaxLength       =   30
         TabIndex        =   98
         Tag             =   "Trabajador Albaran|N|S|0|9999|scafpa|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   780
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
         Index           =   9
         Left            =   -65160
         MaxLength       =   10
         TabIndex        =   134
         Tag             =   "Fec. archiv|F|S|||scafpa|fecenvio|dd/mm/yyyy||"
         Top             =   2115
         Width           =   1350
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
         Index           =   10
         Left            =   -65160
         MaxLength       =   30
         TabIndex        =   100
         Tag             =   "Forma de envio|N|S|0|9999|scafpa|codenvio|0000|N|"
         Text            =   "Text1"
         Top             =   1296
         Width           =   780
      End
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
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
         Height          =   1260
         Left            =   -74760
         TabIndex        =   77
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   2475
         Width           =   14535
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
            Height          =   840
            Index           =   18
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Text            =   "frmComHcoFacturasGR.frx":0A46
            Top             =   240
            Width           =   12645
         End
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   82
            Tag             =   "Observación 5|T|S|||scafpa|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2880
            Width           =   12675
         End
         Begin VB.TextBox Text3 
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
            Left            =   1560
            MaxLength       =   80
            TabIndex        =   81
            Tag             =   "Observación 4|T|S|||scafpa|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2760
            Width           =   12675
         End
         Begin VB.TextBox Text3 
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
            MaxLength       =   80
            TabIndex        =   80
            Tag             =   "Observación 3|T|S|||scafpa|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2280
            Width           =   12675
         End
         Begin VB.TextBox Text3 
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
            Left            =   1755
            MaxLength       =   80
            TabIndex        =   79
            Tag             =   "Observación 2|T|S|||scafpa|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2160
            Width           =   12675
         End
         Begin VB.TextBox Text3 
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
            Left            =   1800
            MaxLength       =   80
            TabIndex        =   78
            Tag             =   "Observación 1|T|S|||scafpa|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1680
            Width           =   12675
         End
         Begin VB.Label Label1 
            Caption         =   "las lineas estan bajo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   23
            Left            =   120
            TabIndex        =   133
            Top             =   360
            Visible         =   0   'False
            Width           =   1650
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
         Left            =   -74730
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   37
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   7830
         Visible         =   0   'False
         Width           =   7245
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
         Index           =   17
         Left            =   -67425
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   38
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   7830
         Visible         =   0   'False
         Width           =   2490
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
         Left            =   -64830
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   129
         Text            =   "nom ccoste"
         Top             =   7830
         Visible         =   0   'False
         Width           =   4590
      End
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   -74760
         TabIndex        =   127
         Top             =   3720
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   128
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
      Begin VB.TextBox txtAux3 
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
         Index           =   2
         Left            =   -71850
         MaxLength       =   30
         TabIndex        =   95
         Tag             =   "Fec. archiv|F|N|||scafpa|fentrada|dd/mm/yyyy||"
         Text            =   "fecentrada"
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
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
         Left            =   -69870
         MaxLength       =   30
         TabIndex        =   97
         Tag             =   "Fecha Pedido|F|S|||scafpa|fecpedpr|dd/mm/yyyy|N|"
         Text            =   "fechaped"
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
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
         Index           =   3
         Left            =   -70815
         MaxLength       =   30
         TabIndex        =   96
         Tag             =   "Nº Pedido|N|S|||scafpa|numpedpr|0000000|N|"
         Text            =   "numpedid"
         Top             =   1890
         Visible         =   0   'False
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
         Index           =   2
         Left            =   -64320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   116
         Text            =   "Text2"
         Top             =   1296
         Width           =   4095
      End
      Begin VB.CheckBox chkDocArchi 
         Caption         =   "Doc.Archivado"
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
         Left            =   -63600
         TabIndex        =   103
         Top             =   2160
         Width           =   3090
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
         Left            =   -66720
         MaxLength       =   4
         TabIndex        =   114
         Tag             =   "Centro coste|T|S|||slifac|codccost||N|"
         Text            =   "cc"
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux3 
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
         Index           =   1
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   94
         Tag             =   "Fecha Albaran|F|N|||scafpa|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
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
         Index           =   0
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   93
         Tag             =   "Nº Albaran|T|N|||scafpa|numalbar||N|"
         Text            =   "numalbar"
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
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
         Index           =   3
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   87
         Tag             =   "Cantidad|N|N|0||slifac|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   6480
         Visible         =   0   'False
         Width           =   735
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
         Left            =   -72840
         MaxLength       =   12
         TabIndex        =   86
         Tag             =   "Nombre Art.|T|N|||slifac|nomartic||N|"
         Text            =   "nomartic"
         Top             =   6480
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
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   85
         Tag             =   "Art.|T|N|||slifac|codartic||N|"
         Text            =   "codartic"
         Top             =   6480
         Visible         =   0   'False
         Width           =   735
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
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   84
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   6480
         Visible         =   0   'False
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
         Index           =   0
         Left            =   -64320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   480
         Width           =   4095
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
         Left            =   -64320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   888
         Width           =   4095
      End
      Begin VB.Frame FrameFactura 
         Height          =   3795
         Left            =   240
         TabIndex        =   58
         Top             =   3105
         Width           =   14555
         Begin VB.Frame FrmRetencionSocios 
            Height          =   855
            Left            =   360
            TabIndex        =   111
            Top             =   1170
            Width           =   3930
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
               Index           =   33
               Left            =   2235
               MaxLength       =   15
               TabIndex        =   34
               Tag             =   "Imp Ret|N|S|||scafpc|impret|#,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   1575
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
               Index           =   32
               Left            =   1140
               MaxLength       =   15
               TabIndex        =   33
               Tag             =   "% Ret|N|S|||scafpc|porret|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Left            =   1785
               TabIndex        =   113
               Top             =   360
               Width           =   195
            End
            Begin VB.Label Label1 
               Caption         =   "Retención"
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
               Left            =   120
               TabIndex        =   112
               Top             =   360
               Width           =   990
            End
         End
         Begin VB.TextBox Text1 
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
            Index           =   30
            Left            =   11340
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Total Factura|N|N|||scafpc|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2355
            Width           =   1875
         End
         Begin VB.TextBox Text1 
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
            Index           =   29
            Left            =   8550
            MaxLength       =   15
            TabIndex        =   32
            Tag             =   "Importe IVA 3|N|S|||scafpc|impoiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2355
            Width           =   2025
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
            Left            =   5490
            MaxLength       =   5
            TabIndex        =   30
            Tag             =   "% IVA 3|N|S|0|99.90|scafpc|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2355
            Width           =   750
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
            Index           =   20
            Left            =   4545
            MaxLength       =   3
            TabIndex        =   29
            Tag             =   "Cod. IVA 3|N|S|0|999|scafpc|tipoiva3|000|N|"
            Text            =   "Text1 7"
            Top             =   2355
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
            Index           =   26
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Base Imponible 3|N|S|||scafpc|baseiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2355
            Width           =   1845
         End
         Begin VB.TextBox Text1 
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
            Index           =   28
            Left            =   8550
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Importe IVA 2|N|S|||scafpc|impoiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1935
            Width           =   2025
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
            Left            =   5490
            MaxLength       =   5
            TabIndex        =   26
            Tag             =   "& IVA 2|N|S|0|99.90|scafpc|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1935
            Width           =   750
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
            Index           =   19
            Left            =   4545
            MaxLength       =   3
            TabIndex        =   25
            Tag             =   "Cod. IVA 2|N|S|0|999|scafpc|tipoiva2|000|N|"
            Text            =   "Text1 7"
            Top             =   1935
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
            Index           =   25
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   27
            Tag             =   "Base Imponible 2 |N|S|||scafpc|baseiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1935
            Width           =   1845
         End
         Begin VB.TextBox Text1 
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
            Index           =   27
            Left            =   8550
            MaxLength       =   15
            TabIndex        =   24
            Tag             =   "Importe IVA 1|N|N|||scafpc|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   2025
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
            Index           =   21
            Left            =   5490
            MaxLength       =   5
            TabIndex        =   22
            Tag             =   "% IVA 1|N|S|0|99.90|scafpc|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   750
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
            Index           =   18
            Left            =   4545
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Cod. IVA 1|N|S|0|999|scafpc|tipoiva1|000|N|"
            Text            =   "Text1 7"
            Top             =   1530
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
            Index           =   24
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   23
            Tag             =   "Base Imponible 1|N|N|||scafpc|baseiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   1845
         End
         Begin VB.TextBox Text1 
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
            Index           =   17
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   20
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1845
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
            Index           =   16
            Left            =   4545
            MaxLength       =   15
            TabIndex        =   19
            Tag             =   "Imp. Dto Gn|N|N|||scafpc|impgnral|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1410
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
            Index           =   15
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Imp. Dto PP|N|N|||scafpc|impppago|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1590
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
            Left            =   360
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Imp.Bruto|N|N|||scafpc|brutofac|#,###,###,##0.00|N|"
            Top             =   615
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
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
            Left            =   4545
            TabIndex        =   92
            Top             =   1245
            Width           =   960
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
            Left            =   5490
            TabIndex        =   91
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
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
            Height          =   255
            Index           =   39
            Left            =   11355
            TabIndex        =   69
            Top             =   2070
            Width           =   1890
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
            Index           =   38
            Left            =   10920
            TabIndex        =   68
            Top             =   2310
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4545
            X2              =   8145
            Y1              =   1185
            Y2              =   1185
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
            Index           =   37
            Left            =   8265
            TabIndex        =   67
            Top             =   1410
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
            Left            =   8535
            TabIndex        =   66
            Top             =   1245
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
            Left            =   5700
            TabIndex        =   65
            Top             =   630
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
            Left            =   4350
            TabIndex        =   64
            Top             =   630
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
            Left            =   2235
            TabIndex        =   63
            Top             =   630
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
            Index           =   14
            Left            =   6285
            TabIndex        =   62
            Top             =   330
            Width           =   1530
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
            Index           =   12
            Left            =   4560
            TabIndex        =   61
            Top             =   330
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto PP"
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
            Left            =   2550
            TabIndex        =   60
            Top             =   330
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Importe  Bruto"
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
            Left            =   375
            TabIndex        =   59
            Top             =   330
            Width           =   1620
         End
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
         Left            =   -70920
         MaxLength       =   12
         TabIndex        =   36
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   6480
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
         Index           =   5
         Left            =   -69240
         MaxLength       =   5
         TabIndex        =   89
         Tag             =   "Dto 1|N|N|0|99.90|slifac|dtoline1|#0.00|N|"
         Text            =   "Dto1"
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
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
         Left            =   -68520
         MaxLength       =   30
         TabIndex        =   90
         Tag             =   "Dto 2|N|N|0|99.90|slifac|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
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
         Left            =   -67920
         MaxLength       =   12
         TabIndex        =   88
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Proveedor"
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
         Height          =   2370
         Left            =   240
         TabIndex        =   45
         Top             =   495
         Width           =   14555
         Begin VB.CheckBox chkISP 
            Caption         =   "Inversión sujeto pasivo"
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
            Left            =   11670
            TabIndex        =   16
            Tag             =   "ISP|N|N|||scafpc|InvSujPas|||"
            Top             =   1620
            Width           =   2700
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
            Index           =   13
            Left            =   10050
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   108
            Text            =   "Text2"
            Top             =   330
            Width           =   4230
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
            Left            =   9330
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "Trabajador|N|N|0|9999|scafpc|codtraba|0000|N|"
            Text            =   "Text1"
            Top             =   330
            Width           =   675
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
            Left            =   1395
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scafpc|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1740
            Width           =   5820
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
            Left            =   1395
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scafpc|codpobla||N|"
            Text            =   "Text15"
            Top             =   1290
            Width           =   810
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
            Left            =   2250
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||scafpc|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1290
            Width           =   4980
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
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "teléfono proveedor|T|S|||scafpc|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   330
            Width           =   2550
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
            Left            =   1395
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF proveedor|T|N|||scafpc|nifprove||N|"
            Text            =   "123456789012345"
            Top             =   330
            Width           =   1920
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
            Left            =   9330
            MaxLength       =   3
            TabIndex        =   13
            Tag             =   "Forma de Pago|N|N|0|999|scafpc|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   825
            Width           =   675
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
            Index           =   10
            Left            =   10050
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   47
            Text            =   "Text2"
            Top             =   825
            Width           =   4230
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
            Index           =   11
            Left            =   8655
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafpc|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   840
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
            Left            =   10695
            MaxLength       =   5
            TabIndex        =   15
            Tag             =   "Descuento General|N|N|0|99.90|scafpc|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   840
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
            Left            =   1395
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scafpc|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   825
            Width           =   5835
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   9045
            ToolTipText     =   "Buscar trabajador"
            Top             =   360
            Width           =   240
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
            Index           =   1
            Left            =   7755
            TabIndex        =   109
            Top             =   330
            Width           =   1260
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1125
            ToolTipText     =   "Buscar población"
            Top             =   1320
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
            TabIndex        =   54
            Top             =   1740
            Width           =   1095
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
            Index           =   16
            Left            =   120
            TabIndex        =   53
            Top             =   1285
            Width           =   1005
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
            Height          =   255
            Index           =   19
            Left            =   3750
            TabIndex        =   52
            Top             =   375
            Width           =   915
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
            TabIndex        =   51
            Top             =   375
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1125
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   390
            Visible         =   0   'False
            Width           =   240
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
            Left            =   7755
            TabIndex        =   50
            Top             =   825
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
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
            Left            =   7755
            TabIndex        =   49
            Top             =   1605
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
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
            Left            =   9690
            TabIndex        =   48
            Top             =   1605
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   9045
            ToolTipText     =   "Buscar forma de pago"
            Top             =   855
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
            TabIndex        =   46
            Top             =   830
            Width           =   1005
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComHcoFacturasGR.frx":0A83
         Height          =   3105
         Left            =   -74760
         TabIndex        =   57
         Top             =   4380
         Width           =   14550
         _ExtentX        =   25665
         _ExtentY        =   5477
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmComHcoFacturasGR.frx":0A98
         Height          =   1905
         Left            =   -74760
         TabIndex        =   70
         Top             =   525
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   3360
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   -68025
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   540
         Width           =   375
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
         Index           =   3
         Left            =   -70785
         TabIndex        =   74
         Tag             =   "Fecha Pedido|F|S|||scafpa|fecpedpr|dd/mm/yyyy|N|"
         Top             =   1620
         Width           =   1350
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
         Index           =   2
         Left            =   -69330
         MaxLength       =   7
         TabIndex        =   73
         Tag             =   "Nº Pedido|N|S|||scafpa|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1620
         Width           =   1000
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
         Index           =   11
         Left            =   -69240
         MaxLength       =   10
         TabIndex        =   118
         Tag             =   "Fec. archiv|F|N|||scafpa|fentrada|dd/mm/yyyy||"
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Para cliente"
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
         Left            =   -67395
         TabIndex        =   136
         Top             =   1716
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
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
         Left            =   -74730
         TabIndex        =   132
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Lote"
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
         Left            =   -67425
         TabIndex        =   131
         Top             =   7560
         Visible         =   0   'False
         Width           =   1155
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
         Left            =   -64830
         TabIndex        =   130
         Top             =   7515
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Entrada"
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
         Left            =   -69375
         TabIndex        =   119
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -65400
         ToolTipText     =   "Buscar trabajador"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de envio"
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
         Left            =   -67395
         TabIndex        =   117
         Top             =   1319
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha archivo"
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
         Left            =   -67395
         TabIndex        =   115
         Top             =   2115
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
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
         Left            =   -67395
         TabIndex        =   56
         Top             =   525
         Width           =   1950
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   -65415
         ToolTipText     =   "Buscar trabajador"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
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
         Left            =   -67395
         TabIndex        =   55
         Top             =   922
         Width           =   1875
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -65430
         ToolTipText     =   "Buscar trabajador"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
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
         Left            =   -70785
         TabIndex        =   76
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
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
         Left            =   -69345
         TabIndex        =   75
         Top             =   1380
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   10215
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
         TabIndex        =   43
         Top             =   135
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
      Left            =   14130
      TabIndex        =   40
      Top             =   10320
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
      Left            =   12960
      TabIndex        =   39
      Top             =   10320
      Width           =   1035
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
      Left            =   14130
      TabIndex        =   41
      Top             =   10305
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
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
Attribute VB_Name = "frmComHcoFacturas2GR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Long 'Codigo de Proveedor    'DAVID.  Estaba integer

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmBasico2  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1

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
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean


'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Private BuscaChekc As String

Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Check2_Click()

End Sub

Private Sub chkDocArchi_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkDocArchi, BuscaChekc
End Sub
Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    If FrameObserva.visible Then
        PonerFoco Text3(4)
    Else
        PonerFocoBtn cmdAceptar
    End If
End Sub

'Private Sub chkISP_Click()
'    If Modo = 1 Then
'        CheckCadenaBusqueda chkISP, BuscaChekc
'    ElseIf Modo = 4 Then
'        ActualizarDatosFactura
'    End If
'End Sub

'Private Sub chkISP_KeyDown(KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub



Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                    TerminaBloquear
'                    PosicionarData
               Else
                    '---- Laura 24/10/2006
                    'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                    Text1(1).Text = Me.Data1.Recordset!FecFactu
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea(NumLinea) Then
'                    'Comprobar si el Articulo tiene control de Nº de Serie
'                    ComprobarNSeriesLineas NumLinea
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
'                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(17), True
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    
                    PonerModo 2
                    
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            BloquearTxt Text2(17), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
'            PonerBotonCabecera True
            PonerModo 2
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        BuscaChekc = ""
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
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
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select scafpc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & Ordenacion
        

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
Dim DeVarios As Boolean

    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
        
    'Si es proveedor de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(2).Text)
    BloquearDatosProve (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo EModificarLinea


    If Modo <> 2 Then Exit Sub
    
    If Val(Data1.Recordset!intconta) = 1 Then
        MsgBox "Factura contabilizada", vbExclamation
        Exit Sub
    End If


    PonerModo 5

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!Numalbar & "'"
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 8
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(17).Text = DataGrid1.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
'    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt Text2(17), False 'Campo Ampliacion Linea
'    PonerFoco txtAux(4)
    PonerFoco Text2(16)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
'
'    Select Case grid
'        Case "DataGrid1"
    'Dic 2012.   Lo vuelvo a poner
    If grid = "DataGrid1" Then
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1

            B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas

            For jj = 0 To txtAux.Count - 1
                If jj = 4 Or jj = 5 Or jj = 6 Or jj = 7 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = B
                    txtAux(jj).Enabled = jj = 4
                    
                Else
                    txtAux(jj).Enabled = False
                End If
                
            Next jj
    End If
        If grid = "DataGrid2" Then
            DeseleccionaGrid Me.DataGrid2
            B = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto + 20
                txtAux3(jj).visible = B
            Next jj
        End If
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Proveedor:  " & Text1(2).Text & " - " & Text1(3).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Text1(0).Text
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
       ' NumPedElim = Data1.Recordset.Fields(1).Value   MAAAAl. Ya que el nºfac prov es ALFANUMERICO
        
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
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub cmdObserva_Click()
    If Modo <> 2 And Modo <> 4 Then Exit Sub
    If Me.FrameObserva.visible = False Then
        Me.DataGrid1.visible = False
        Me.FrameObserva.visible = True
        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "volver.ico"
        Me.cmdObserva.ToolTipText = "volver lineas albaran"
        BloqueaText3
    Else
        Me.DataGrid1.visible = True
        Me.FrameObserva.visible = False
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(41).Picture
        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    End If
End Sub


Private Sub BloqueaText3()
Dim i As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For i = 0 To 1
        BloquearTxt Text3(i), (Modo <> 4) And Modo <> 1
    Next i
    If Me.FrameObserva.visible Then
        For i = 4 To 8
            BloquearTxt Text3(i), (Modo <> 4)
        Next i
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For i = 2 To 3
        'Feb 2011. Dejamos que busque por ellos
        BloquearTxt Text3(i), Modo <> 1
    Next i
    Me.chkDocArchi.Enabled = Modo = 1 Or Modo = 4
    BloquearTxt Text3(9), Not chkDocArchi.Enabled
    BloquearTxt Text3(10), Not chkDocArchi.Enabled
     BloquearTxt Text3(11), Modo <> 1
     BloquearTxt Text3(12), Modo <> 1
     
     
     
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci)
            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
        End If
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
        If vEmpresa.TieneAnalitica Then
            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            txtAux2(8).Text = ""
        End If
        
        
    Else
        Text2(16).Text = ""
        Text2(17).Text = ""
        txtAux2(8).Text = ""
    End If
    Exit Sub

Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Not Data3.Recordset.EOF Then
        Text3(0).Text = DBLet(Data3.Recordset.Fields!codtrab2, "T")
        Text3_LostFocus (0)
        Text3(1).Text = DBLet(Data3.Recordset.Fields!CodTrab1, "T")
        Text3_LostFocus (1)

'        Text3(2).Text = DBLet(data3.Recordset.Fields!numpedpr, "N")
'        If Text3(2).Text <> "0" Then
'            FormateaCampo Text3(2)
'        Else
'            Text3(2).Text = ""
'        End If
'        Text3(3).Text = DBLet(data3.Recordset.Fields!fecpedpr, "F")
        
        'Observaciones
        Text3(4).Text = DBLet(Data3.Recordset.Fields!observa1, "T")
        Text3(5).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
        Text3(6).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
        Text3(7).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
        Text3(8).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        Text3(9).Text = DBLet(Data3.Recordset.Fields!FecEnvio, "T")
        Text3(10).Text = DBLet(Data3.Recordset.Fields!CodEnvio, "T")
        Text3(11).Text = DBLet(Data3.Recordset.Fields!fentrada, "T")
        Text3_LostFocus (10)
        Me.chkDocArchi.Value = DBLet(Data3.Recordset!docarchiv, "N")
        Text3(12).Text = DBLet(Data3.Recordset.Fields!codClien, "T")
        Text3_LostFocus (12)
        
        Text2(18).Text = ""
        For i = 4 To 8
            If Text3(i).Text <> "" Then Text2(18).Text = Text2(18).Text & Text3(i).Text & vbCrLf
        Next i
        
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        Me.chkDocArchi.Value = 0
        Text2(0).Text = ""
        Text2(1).Text = ""
        Text2(18).Text = ""
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next



    ' ICONITOS DE LA BARRA
'    btnPrimero = 21
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(9).Image = 10 'Mto Lineas Ofertas
'        .Buttons(10).Image = 16 'Imprimir
'
'        .Buttons(12).Image = 33 'vicnular (alzira)
'
'        .Buttons(18).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        .Buttons(5).Image = 1   'Botón Buscar
        .Buttons(6).Image = 2   'Botón Todos
        .Buttons(8).Image = 16
    End With
    
    With Me.Toolbar5
'--
'        .ImageList = frmPpal.imgListComun
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 33 'vicnular (alzira)
        .Buttons(2).Image = 42 'fras pdtes de contabilizar
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
      
      
'-- cambiado por lo de abajo
'    Toolbar1.Buttons(12).visible = vParamAplic.NumeroInstalacion = 1
'--
'    FrameBotonGnral2.visible = vParamAplic.NumeroInstalacion = vbAlzira
'    FrameBotonGnral2.Enabled = vParamAplic.NumeroInstalacion = vbAlzira
'    If Not FrameBotonGnral2.visible Then
'        Me.FrameDesplazamiento.Left = FrameBotonGnral2.Left
'    End If
'[Monica]11/05/2021: añadimos el punto de facturas pendientes de contabilizar
    Me.FrameBotonGnral2.visible = True
    Me.Toolbar5.Buttons(1).visible = vParamAplic.NumeroInstalacion = vbAlzira
    Me.Toolbar5.Buttons(1).Enabled = vParamAplic.NumeroInstalacion = vbAlzira
    If Not vParamAplic.NumeroInstalacion = vbAlzira Then
        Me.FrameBotonGnral2.Width = 930
        Me.Toolbar5.Width = Me.FrameBotonGnral2.Width - 300
        Me.FrameDesplazamiento.Left = 4380
    End If
      
      
    'El frame FrmRetencionSocios es el que tendra el % retencion aplicable a los socios
    'Solo sera visible SI (y solo si) en paretros.ctaretnecion <>""
    FrmRetencionSocios.visible = vParamAplic.CtaReten <> ""
    Me.chkISP.visible = vParamAplic.InvSujetoPasivo

      
    LimpiarCampos   'Limpia los campos TextBox
     
    
    
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "scafpc"
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.fecrecep desc ,scafpc.codprove, scafpc.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
         PonerModo 0
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
    Me.chkDocArchi.Value = 0

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
        cadB = cadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaSeleccion, 3)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

        Indice = 7
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
        'provincia
        Text1(Indice + 2).Text = devuelve
End Sub

Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
        Text3(10).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod env
        Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom env
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(4).Tag)
    If Indice = 4 Then
        Indice = Indice + 9
        Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    Else
        Text3(Indice - 5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    End If
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            PonerFoco Text1(2)
            Set frmProv = New frmBasico2
'            frmProv.DatosADevolverBusqueda = "0"
'            frmProv.Show vbModal
            AyudaProveedores frmProv, Text1(2)
            Set frmProv = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
            
        Case 1 'NIF para proveedor de Varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0"
            frmPV.Show vbModal
            Set frmPV = Nothing
            Indice = 7
            PonerFoco Text1(Indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 7
            VieneDeBuscar = True
            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
'            PonerFoco Text1(Indice)
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice).Text
            Set frmFP = Nothing
            
            PonerFoco Text1(Indice)
            
        Case 4, 5, 6 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Me.imgBuscar(4).Tag = Index
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
            If Index = 4 Then
                Indice = 13
                Set frmT = New frmBasico2
                AyudaTrabajadores frmT, Text1(Indice)
                Set frmT = Nothing
            Else
                Indice = Index - 5
                Set frmT = New frmBasico2
                AyudaTrabajadores frmT, Text3(Indice)
                Set frmT = Nothing
            End If

            If Index = 4 Then
                PonerFoco Text1(13)
            Else
                PonerFoco Text3(Index - 5)
            End If
        Case 7
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0|1|"
            frmFE.Show vbModal
            Set frmFE = Nothing
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
    'Imprimir Factura PROVEEDOR
    
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Not PonerParamRPT2(97, BuscaChekc, Modo, TituloLinea, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    
    With frmImprimir
        .NombreRPT = CStr(TituloLinea)
        .NombrePDF = .NombreRPT
        .SeleccionaRPTCodigo = pRptvMultiInforme
        TituloLinea = "{scafpc.codprove} = " & Data1.Recordset!Codprove & " AND {scafpc.numfactu} =" & DBSet(Data1.Recordset!Numfactu, "T")
        TituloLinea = TituloLinea & " AND {scafpc.fecfactu} = DATE(" & Format(Data1.Recordset!FecFactu, "yyyy,mm,dd") & ")"
        .FormulaSeleccion = TituloLinea
        BuscaChekc = BuscaChekc & "pEmpresa=""" & vEmpresa.nomempre & """|"
        .OtrosParametros = BuscaChekc
        .NumeroParametros = Modo
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2054
        .Titulo = "Factura proveedor"
        .Show vbModal
    End With

    
    Modo = 2
    BuscaChekc = ""
    TituloLinea = ""
    

    
End Sub


Private Sub mnLineas_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
    If Me.FrmRetencionSocios.visible Then
        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
            MsgBox "Factura liquidación socios. No puede modificarse", vbExclamation
            Exit Sub
        End If
    End If

    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
'    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
'    If Me.FrmRetencionSocios.visible Then
'        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
'            MsgBox "Factura liquidación socios. No puede modificarse", vbExclamation
'            Exit Sub
'        End If
'    End If


    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM scafpa "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim SQL As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifpc "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


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


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 0 'proveedor
            Case 4: KEYBusqueda KeyAscii, 1 'Nif
            Case 7: KEYBusqueda KeyAscii, 2 'CP
            Case 13: KEYBusqueda KeyAscii, 4 'trabajador
            Case 10: KEYBusqueda KeyAscii, 3 'forma de pago
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 31 'Fecha factura,fecha recepcion,doc archivado
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 13 'Cod trabajador
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            If Modo > 2 Then
                If Text2(Index).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
        Case 2 'Cod. prove
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                PonerDatosProveedor (Text1(Index).Text)
            End If
        
        Case 4 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(4).Text = DBLet(Data1.Recordset!nifProve, "T") Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
        
        Case 7 'Cod. Postal
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
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
                If Modo > 2 Then
                    If Text2(Index).Text = "" Then
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
        Case 11, 12 'Descuentos
            If Modo = 4 Then 'comprobar que el dato a cambiado
                
                If Text1(Index).Text = "" Then
                    Text1(Index).Text = "0"
                Else
                    If Not PonerFormatoDecimal(Text1(Index), 4) Then Text1(Index).Text = "0"
                End If
            
                If Index = 11 Then
                    
                    'If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
                ElseIf Index = 12 Then
                    'If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
                End If
            End If
            
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4 'Tipo 4: Decimal(4,2)
                If Not ActualizarDatosFactura Then
                   If Index = 11 Then Text1(Index).Text = Data1.Recordset!DtoPPago
                   If Index = 12 Then Text1(Index).Text = Data1.Recordset!DtoGnral
                   FormateaCampo Text1(Index)
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    'De momento a mano
    If InStr(1, BuscaChekc, "chkDocArchi") > 0 Then
        'Ha clicado sobnre doarchiv
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & " scafpa.docarchiv = " & Me.chkDocArchi.Value
    End If
    
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String
'
'    'Llamamos a al form
'    '##A mano
'    cad = ""
''        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
'        cad = cad & ParaGrid(Text1(0), 18, "Nº Factura")
'        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
'        cad = cad & ParaGrid(Text1(2), 12, "Prov.")
'        cad = cad & ParaGrid(Text1(3), 55, "Nombre Prov")
'        tabla = NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
'        Titulo = "Facturas"
'        devuelve = "0|1|2|"
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
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
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
    AyudaFacturasCompra frmB, Text1(0), cadB
    Set frmB = Nothing
    
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then PonerFoco Text1(0)
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, Data3, True
    
    BotonesToolBarAux
   
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    'Para que no de error el chkarch
    PonerCamposForma Me, Data1
    
    '++
    If Text1(15).Text = "0,00" Then Text1(15).Text = ""
    If Text1(16).Text = "0,00" Then Text1(16).Text = ""
    If Text1(32).Text = "0,00" Then Text1(32).Text = ""
    If Text1(32).Text = "" Then Text1(33).Text = ""
    If CInt(ComprobarCero(Text1(19).Text)) = 0 Then
        Text1(19).Text = ""
        Text1(22).Text = ""
        Text1(25).Text = ""
        Text1(28).Text = ""
    End If
    If CInt(ComprobarCero(Text1(20).Text)) = 0 Then
        Text1(20).Text = ""
        Text1(23).Text = ""
        Text1(26).Text = ""
        Text1(29).Text = ""
    End If
    '++hasta aqui
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = 0
    If Text1(15).Text <> "" Then BrutoFac = CSng(Text1(15).Text)
    If Text1(16).Text <> "" Then BrutoFac = BrutoFac + CSng(Text1(16).Text)
    
    BrutoFac = CSng(Text1(14).Text) - BrutoFac
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "sforpa", "nomforpa")
    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    
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
Dim B As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'--    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    
    Me.Check1.Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    If Me.chkISP.visible Then Me.chkISP.Enabled = Check1.Enabled
    
    B = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(3), B 'referencia
    
    FrmRetencionSocios.Enabled = Not B
    
    'Importes siempre bloqueados
    For i = 14 To 30
        BloquearTxt Text1(i), (Modo <> 1)
    Next i

    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For i = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), (Modo <> 1)
    Next i
    
    'ampliacion linea
    B = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    Me.Label1(35).visible = True 'b
    Me.Label1(3).visible = True 'b
    Me.Text2(16).visible = True 'b
    Me.Text2(17).visible = True 'b
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    BloquearTxt Text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)


    ' ---- [20/10/2009] [LAURA] : añadir del centro de coste
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica) '-- And (Modo = 5)
    Me.txtAux2(8).visible = (vEmpresa.TieneAnalitica) '-- And (Modo = 5)
    BloquearTxt txtAux2(8), True



    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2) '-- And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    Me.chkDocArchi.Enabled = B
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    
    'Para que no den errores los 0's de los importes de dtos
    ComprobarDatosTotales
        
    'comprobamos datos OK de la tabla scafac
    B = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
       
    BuscaChekc = ""
    If Text3(0).Text = "" Or Text2(0).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(21).Caption & vbCrLf
    
    If Text3(1).Text = "" Xor Text2(1).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(9).Caption & vbCrLf
    If Text3(10).Text = "" Xor Text2(2).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(13).Caption & vbCrLf
    If BuscaChekc <> "" Then
        BuscaChekc = "Error en campos: " & vbCrLf & BuscaChekc
        MsgBox BuscaChekc, vbExclamation
        BuscaChekc = ""
        B = False
    End If
       
       
       
       
       
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim i As Byte
On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 5 Or i = 6 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = B
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 17 And KeyAscii = 13 Then 'campo nº de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 5 'trabajador albaran
            Case 1: KEYBusqueda KeyAscii, 6 'trabajador pedido
            Case 10: KEYBusqueda KeyAscii, 7 'forma de envio
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub Text3_LostFocus(Index As Integer)

    If Modo = 1 Then
        If Not PerderFocoGnral(Text3(Index), Modo) Then Exit Sub
    End If
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    If Modo = 4 Then
        'Modificar
        If Index = 0 Or Index = 1 Or Index = 10 Then
            If Not PonerFormatoEntero(Text3(Index)) Then Text3(Index).Text = ""
        End If
    End If
    
    Select Case Index
        Case 0, 1 'trabajador
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
            
        Case 9
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)
        Case 10
                    
            Text2(2).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Forma de envio", "N")
        Case 12
            Text2(3).Text = PonerNombreDeCod(Text3(Index), conAri, "sclien", "nomclien", "codclien", "Cliente", "N")
    End Select
End Sub


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    'Estamo presuponiendo que Solo hay un index
    If Button.Index = 2 Then BotonModificarLinea
    
'    Select Case Button.Index
'        Case 1
'        Case 2
'            BotonModificarLinea
'        Case 3
'        Case Else
'    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos

        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 9  'Lineas
            mnLineas_Click
        Case 8 'Imprimir Albaran
            mnImprimir_Click
        Case 12
            LanazarVincularAlbaranes
            
        Case 18    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vWhere As String
Dim B As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!Numalbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        SQL = "UPDATE " & NomTablaLineas & " SET "
        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N")
        SQL = SQL & ", numlotes=" & DBSet(Text2(17).Text, "T")
        SQL = SQL & vWhere
    End If
    
    If SQL <> "" Then
        'actualizar la factura y vencimientos
        B = ModificarFactura(SQL)
        ModificarLinea = B
    End If
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        B = False
    End If
    ModificarLinea = B
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
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled

    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
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
Dim tots As String

    On Error GoTo ECargaGrid
    
    vData.Refresh
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Lineas de Albaran
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            If vEmpresa.TieneAnalitica Then
                tots = tots & "S|txtAux(0)|T|Alm.|640|;S|txtAux(1)|T|Artículo|2250|;S|txtAux(2)|T|Descripción|4560|;"
                tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|1250|;S|txtAux(4)|T|Precio|1400|;S|txtAux(5)|T|Dto.1|860|;S|txtAux(6)|T|Dto.2|860|;"
                tots = tots & "S|txtAux(7)|T|Importe|1550|;N||||0|;"
                tots = tots & "S|txtAux(8)|T|CCos|620|;"
            Else
                tots = tots & "S|txtAux(0)|T|Alm.|640|;S|txtAux(1)|T|Artículo|2250|;S|txtAux(2)|T|Descripción|4980|;"
                tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|1250|;S|txtAux(4)|T|Precio|1400|;S|txtAux(5)|T|Dto.1|860|;S|txtAux(6)|T|Dto.2|860|;"
                tots = tots & "S|txtAux(7)|T|Importe|1750|;N||||0|;"
                tots = tots & "N||||0|;"
            End If
            
            arregla tots, DataGrid1, Me, 350
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgRight
            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,fecnvio,docarchiv, codclien "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albaran|1620|;S|txtAux3(1)|T|Fecha|1300|;"
            tots = tots & "S|txtAux3(2)|T|F.Entrada|1300|;S|txtAux3(3)|T|Pedido|1100|;S|txtAux3(4)|T|F.Pedido|1300|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me, 350
        
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanazarVincularAlbaranes
        Case 2
            'Facturas pendientes de contabilizar (PROVEEDORES)
            Screen.MousePointer = vbHourglass
            frmUtilidades.Opcion = 7
            frmUtilidades.Show vbModal
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), IIf(Modo = 5, 3, 5)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
            End If
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 6 Then PonerFoco Me.Text2(16)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim Cta As String
Dim B As Boolean

    On Error GoTo FinEliminar

        B = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        
        If vParamAplic.ContabilidadNueva Then
            SQL = "codmacta"
        Else
            SQL = "ctaprove"
        End If
        SQL = SQL & " ='" & Cta & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        If vParamAplic.ContabilidadNueva Then SQL = SQL & " AND numserie='" & SerieFraPro & "'"
        
        ConnConta.Execute "Delete from " & IIf(vParamAplic.ContabilidadNueva, "pagos", "spagop") & " WHERE " & SQL
        B = True
        
        
        'Eliminar en tablas de factura de Ariges: scafpc, scafpa, slifpc
        '---------------------------------------------------------------
        If B Then
            SQL = " " & ObtenerWhereCP(True)
        
            'Lineas de facturas (slifpc)
            conn.Execute "Delete from " & NomTablaLineas & SQL
        
            'Lineas de cabeceras de albaranes de la factura
            conn.Execute "Delete from scafpa " & SQL
            
            'Cabecera de facturas (scafpc)
            conn.Execute "Delete from " & NombreTabla & SQL
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If B Then
        
        End If
        
'        b = True
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    Eliminar = B
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String
On Error Resume Next
    SQL = "codprove= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    If Opcion = 1 Then
        SQL = "SELECT codprove, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost "
        SQL = SQL & " FROM slifpc " 'lineas de factura
    ElseIf Opcion = 2 Then
        SQL = "SELECT codprove,numfactu,fecfactu,numalbar, fechaalb, fentrada, numpedpr,fecpedpr,codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,fecenvio,docarchiv,codenvio,codclien  "
        SQL = SQL & " FROM scafpa " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then SQL = SQL & " AND numalbar=" & DBSet(Data3.Recordset.Fields!Numalbar, "T")
    Else
        SQL = SQL & " WHERE false " 'numfactu = -1"
    End If
    SQL = SQL & " ORDER BY codprove, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then SQL = SQL & ", numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean
Dim i As Integer
Dim bAux As Boolean

        B = (Modo = 2) '--Or (Modo = 5 And ModificaLineas = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        B = (Modo = 2)
        
        'Vicular
        Toolbar5.Buttons(1).Enabled = B
        Toolbar1.Buttons(8).Enabled = B
        
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
        
        '++
        Toolbar5.Buttons(2).Enabled = (Modo = 0 Or Modo = 2) And vUsu.Nivel = 0
        
        
        BotonesToolBarAux
        
        
End Sub

Private Sub BotonesToolBarAux()
Dim B As Boolean

        B = (Modo = 2)
        
        
        ToolAux(0).Buttons(1).Enabled = False
        If B Then
            If Data2.Recordset Is Nothing Then
                B = False
            Else
                B = Data2.Recordset.RecordCount > 0
            End If
        End If
        ToolAux(0).Buttons(2).Enabled = B
        ToolAux(0).Buttons(3).Enabled = False


End Sub


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
           
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!Codprove) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vProve.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vProve.Nombre  'Nom prove
                Text1(6).Text = vProve.Domicilio
                Text1(7).Text = vProve.CPostal
                Text1(8).Text = vProve.Poblacion
                Text1(9).Text = vProve.Provincia
                Text1(4).Text = vProve.NIF
                Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del proveedor
Dim vProve As CProveedor
Dim B As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    B = vProve.LeerDatosProveVario(nifProve)
    If B Then
        Text1(3).Text = vProve.Nombre   'Nom proveedor
        Text1(6).Text = vProve.Domicilio
        Text1(7).Text = vProve.CPostal
        Text1(8).Text = vProve.Poblacion
        Text1(9).Text = vProve.Provincia
        Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub LimpiarDatosProve()
Dim i As Byte

    For i = 3 To 9
        Text1(i).Text = ""
    Next i
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim SQL As String
Dim B As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If Data3.Recordset.EOF Then Exit Function
    
    'comprobar datos OK de la scafac1
    B = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not B Then Exit Function
    
    'Comprobaremos de albaranes factura
    BuscaChekc = ""
    If Text3(0).Text = "" Or Text2(0).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(21).Caption & vbCrLf
    If Text3(1).Text = "" Xor Text2(1).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(9).Caption & vbCrLf
    If Text3(10).Text = "" Xor Text2(2).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(13).Caption & vbCrLf
    
    If BuscaChekc <> "" Then
        BuscaChekc = "Error en campos: " & vbCrLf & BuscaChekc
        Err.Raise 513, , BuscaChekc
        
        'MsgBox BuscaChekc, vbExclamation
        'BuscaChekc = ""
        'Exit Function
    End If
    
    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        SQL = SQL & ", observa1=" & DBSet(Text3(4).Text, "T")
        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
    End If
    SQL = SQL & ", fecenvio=" & DBSet(Text3(9).Text, "F", "S")
    SQL = SQL & ", docarchiv=" & DBSet(Me.chkDocArchi.Value, "N")
    SQL = SQL & ", codenvio=" & DBSet(Text3(10).Text, "N", "S")
    SQL = SQL & ObtenerWhereCP(True)
    'Antes oct 2011
    'SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    SQL = SQL & " AND numalbar=" & DBSet(Data3.Recordset.Fields!Numalbar, "T")
    conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim vFactu As CFacturaCom
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
            MenError = "Modificando datos proveedor varios"
            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            
            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura
                
                'antes de Eliminar en las tablas de la Contabilidad
                Set vFactu = New CFacturaCom
                bol = vFactu.LeerDatos3(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                
                If bol Then
                    'Eliminar de la spagop
                    
                    SQL = "Delete from spagop WHERE  ctaprove='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
                    SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                    If vParamAplic.ContabilidadNueva Then
                        'FALTA numserie en el WHERE
                        
                        SQL = Replace(SQL, "ctaprove=", "codmacta=")
                        SQL = Replace(SQL, "spagop", "pagos")
                    End If
                    ConnConta.Execute SQL
                    
                    'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                    If bol Then
                        bol = vFactu.InsertarEnTesoreria(MenError, "")
                    End If
                End If
                Set vFactu = Nothing
            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError & vbCrLf & Err.Description
        MsgBox MenError, vbExclamation
    End If
End Function



Private Function FactContabilizada() As Boolean
Dim Cta As String, numasien As String
On Error GoTo EContab


    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        MsgBox "Opcion NO disponible. ", vbCritical
        FactContabilizada = True
        Exit Function
    End If


    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        If Cta <> "" Then
            If vParamAplic.ContabilidadNueva Then
                numasien = DevuelveDesdeBDNew(conConta, "factpro", "numasien", "codmacta", Cta, "T", , "numfactu", Text1(0).Text, "T", "fecfactu", Text1(1).Text, "F")
            Else
                numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", Cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
            End If
            If numasien <> "" Then
                FactContabilizada = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
            Else
                FactContabilizada = False
            End If
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
    End If
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosProve(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For i = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
'Modifica los datos de la tabla de Proveedores Varios
Dim vProve As CProveedor
On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(3).Text
        vProve.Domicilio = Text1(6).Text
        vProve.CPostal = Text1(7).Text
        vProve.Poblacion = Text1(8).Text
        vProve.Provincia = Text1(9).Text
        vProve.TfnoAdmon = Text1(5).Text
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


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim Cad As String
Dim RS As ADODB.Recordset
On Error Resume Next

    Cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    Cad = Cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    Cad = Cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then 'where para la factura
        Cad = " WHERE codprove=" & RS!Codprove & " AND numfactu= '" & RS!Numfactu & "' AND fecfactu=" & DBSet(RS!FecFactu, "F")
    Else
        Cad = " where numfactu=-1"
    End If
    RS.Close
    Set RS = Nothing

    ObtenerSelFactura = Cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaCom
Dim cadSel As String

    Set vFactu = New CFacturaCom
    cadSel = ObtenerWhereCP(False)
    cadSel = "slifpc." & cadSel
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    
    'Si tiene RETENCION
    If Me.FrmRetencionSocios.visible Then
        vFactu.PorRet = ImporteFormateado(Text1(32).Text)
        vFactu.ImpRet2 = ImporteFormateado(Text1(33).Text)
    End If
    vFactu.FijarTipoIvaProveedor CLng(Val(Text1(2).Text))
    
    
    If vFactu.CalcularDatosFactura2(cadSel, "scafpa", "slifpc", CDate(Text1(1).Text), Val(Data1.Recordset!InvSujPas) = 1) Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
        If Me.FrmRetencionSocios.visible Then
            Text1(32).Text = vFactu.PorRet
            Text1(33).Text = vFactu.ImpRet2
        End If
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 14 To 17
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(i)
    Next i
    
    For i = 24 To 26
        If Text1(i).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(i).Text) = 0 Then
                Text1(i).Text = QuitarCero(Text1(i).Text)
                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
                Text1(i + 3).Text = QuitarCero(Text1(i + 3).Text)
            Else
                FormateaCampo Text1(i)
                FormateaCampo Text1(i - 3)
                FormateaCampo Text1(i - 6)
                FormateaCampo Text1(i + 3)
            End If
        Else 'No hay Base Imponible
            Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
            Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
            Text1(i + 3).Text = ""
        End If
    Next i
    
    If Me.FrmRetencionSocios.visible Then
        FormateaCampo Text1(32)
        FormateaCampo Text1(33)
    End If
End Sub



Private Sub ComprobarDatosTotales()
Dim i As Byte

    For i = 14 To 17
        Text1(i).Text = ComprobarCero(Text1(i).Text)
    Next i
End Sub


Private Sub LanazarVincularAlbaranes()
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    'Es para Alzira
    'El proveedor debe estar como codprove en la ficha de las flotas de maquinaria
    'ya que realiza trabajos externos para la cooperativa

    BuscaChekc = DevuelveDesdeBD(conAri, "count(*)", "sflotas", "codprove", Text1(2).Text)
    If BuscaChekc = "" Then BuscaChekc = "0"
    If Val(BuscaChekc) = 0 Then
        MsgBox "Proveedor no esta asginado a los partes de trabajo", vbExclamation
        Exit Sub
    End If
    
    
    frmComCasarAlbaranes.Codprove = Data2.Recordset!Codprove
    frmComCasarAlbaranes.Show vbModal
End Sub

