VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntAlbaranesGR 
   BackColor       =   &H00AEAEAE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Clientes"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   17985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacEntAlbaranesGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3840
      TabIndex        =   144
      Top             =   0
      Width           =   2235
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   240
         TabIndex        =   145
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
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nº serie"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar factura"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Marcar para facturar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir portes"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
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
      TabIndex        =   142
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   143
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
      TabIndex        =   140
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   141
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
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
      Index           =   56
      Left            =   14520
      MaxLength       =   15
      TabIndex        =   139
      Text            =   "Text1 7"
      Top             =   240
      Width           =   1530
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Height          =   360
      Index           =   0
      Left            =   12600
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   138
      Text            =   "BASE IMPONIBLE"
      Top             =   240
      Width           =   1845
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
      Left            =   16320
      TabIndex        =   137
      Top             =   360
      Width           =   1575
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
      Index           =   9
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   110
      Text            =   "nom ccoste"
      Top             =   9000
      Visible         =   0   'False
      Width           =   3885
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
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   70
      Text            =   "frmFacEntAlbaranesGR.frx":000C
      Top             =   8640
      Visible         =   0   'False
      Width           =   7005
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   54
      Top             =   8880
      Width           =   3615
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   3540
      End
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   55
         Top             =   120
         Width           =   3075
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
      Left            =   16560
      TabIndex        =   50
      Top             =   8880
      Width           =   1335
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
      Left            =   15000
      TabIndex        =   49
      Top             =   8880
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   0
      Top             =   9480
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
      Left            =   0
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
      Height          =   6780
      Left            =   120
      TabIndex        =   56
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1680
      Width           =   17835
      _ExtentX        =   31459
      _ExtentY        =   11959
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmFacEntAlbaranesGR.frx":0049
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
      Tab(0).Control(13)=   "txtAux(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAux(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(11)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAux(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "FrameToolAux(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntAlbaranesGR.frx":0065
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(54)"
      Tab(1).Control(1)=   "Text2(54)"
      Tab(1).Control(2)=   "cmdAux2(0)"
      Tab(1).Control(3)=   "txtAux2(1)"
      Tab(1).Control(4)=   "txtAux2(0)"
      Tab(1).Control(5)=   "Text1(53)"
      Tab(1).Control(6)=   "Text2(52)"
      Tab(1).Control(7)=   "Text1(52)"
      Tab(1).Control(8)=   "FrameToolAux(1)"
      Tab(1).Control(9)=   "Text2(51)"
      Tab(1).Control(10)=   "Text1(51)"
      Tab(1).Control(11)=   "Text2(1)"
      Tab(1).Control(12)=   "Text1(44)"
      Tab(1).Control(13)=   "Text2(43)"
      Tab(1).Control(14)=   "Text1(43)"
      Tab(1).Control(15)=   "chkDocArchi"
      Tab(1).Control(16)=   "Text1(41)"
      Tab(1).Control(17)=   "Text1(39)"
      Tab(1).Control(18)=   "Text1(29)"
      Tab(1).Control(19)=   "Text2(29)"
      Tab(1).Control(20)=   "Text1(28)"
      Tab(1).Control(21)=   "Text2(28)"
      Tab(1).Control(22)=   "Text1(27)"
      Tab(1).Control(23)=   "Text2(27)"
      Tab(1).Control(24)=   "Text1(2)"
      Tab(1).Control(25)=   "Text1(25)"
      Tab(1).Control(26)=   "Text1(26)"
      Tab(1).Control(27)=   "Text1(24)"
      Tab(1).Control(28)=   "Text1(23)"
      Tab(1).Control(29)=   "Text1(22)"
      Tab(1).Control(30)=   "Text1(21)"
      Tab(1).Control(31)=   "Text1(20)"
      Tab(1).Control(32)=   "Text1(19)"
      Tab(1).Control(33)=   "Text1(18)"
      Tab(1).Control(34)=   "Text1(38)"
      Tab(1).Control(35)=   "chkImpreso"
      Tab(1).Control(36)=   "DataGrid2"
      Tab(1).Control(37)=   "imgBuscar(18)"
      Tab(1).Control(38)=   "Label1(60)"
      Tab(1).Control(39)=   "Label1(67)"
      Tab(1).Control(40)=   "Label1(66)"
      Tab(1).Control(41)=   "imgBuscar(17)"
      Tab(1).Control(42)=   "Label1(36)"
      Tab(1).Control(43)=   "imgBuscar(16)"
      Tab(1).Control(44)=   "Label1(65)"
      Tab(1).Control(45)=   "Label1(63)"
      Tab(1).Control(46)=   "Label1(55)"
      Tab(1).Control(47)=   "Label1(54)"
      Tab(1).Control(48)=   "imgBuscar(13)"
      Tab(1).Control(49)=   "imgBuscar(9)"
      Tab(1).Control(50)=   "imgFecha(40)"
      Tab(1).Control(51)=   "Label1(52)"
      Tab(1).Control(52)=   "Label1(24)"
      Tab(1).Control(53)=   "Label1(23)"
      Tab(1).Control(54)=   "imgBuscar(8)"
      Tab(1).Control(55)=   "Label1(9)"
      Tab(1).Control(56)=   "imgBuscar(7)"
      Tab(1).Control(57)=   "Label1(12)"
      Tab(1).Control(58)=   "Label1(11)"
      Tab(1).Control(59)=   "Label1(10)"
      Tab(1).Control(60)=   "Label1(5)"
      Tab(1).Control(61)=   "Label1(3)"
      Tab(1).Control(62)=   "Label1(45)"
      Tab(1).ControlCount=   63
      TabCaption(2)   =   "Fitosanitarios / Campos"
      TabPicture(2)   =   "frmFacEntAlbaranesGR.frx":0081
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameCampos"
      Tab(2).Control(1)=   "FrameManipulador"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Carta de portes"
      TabPicture(3)   =   "frmFacEntAlbaranesGR.frx":009D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Totales"
      TabPicture(4)   =   "frmFacEntAlbaranesGR.frx":00B9
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameFactura"
      Tab(4).Control(1)=   "FrameHco"
      Tab(4).Control(2)=   "FrameFacRec"
      Tab(4).Control(3)=   "Text1(40)"
      Tab(4).Control(4)=   "Label1(49)"
      Tab(4).ControlCount=   5
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
         Index           =   54
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   223
         Tag             =   "Chofer|N|S|0||scaalb|chofer|000||"
         Text            =   "Text1"
         Top             =   6240
         Width           =   1020
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
         Index           =   54
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   222
         Text            =   "Text2"
         Top             =   6240
         Width           =   5445
      End
      Begin VB.CommandButton cmdAux2 
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
         Left            =   -65880
         TabIndex        =   218
         ToolTipText     =   "Buscar artículo"
         Top             =   6120
         Visible         =   0   'False
         Width           =   195
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
         Height          =   360
         Index           =   1
         Left            =   -65640
         MaxLength       =   15
         TabIndex        =   217
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   6120
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
         Height          =   360
         Index           =   0
         Left            =   -66480
         MaxLength       =   15
         TabIndex        =   216
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   6120
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
         Height          =   1920
         Index           =   53
         Left            =   -62520
         MaxLength       =   30
         TabIndex        =   48
         Tag             =   "T|T|S|||scaalb|notasportes|||"
         Text            =   "Text1"
         Top             =   4560
         Width           =   5100
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
         Index           =   52
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   213
         Text            =   "Text2"
         Top             =   5400
         Width           =   5445
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
         Index           =   52
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   41
         Tag             =   "Naturaleza|N|S|0|999|scaalb|codnatura|000||"
         Text            =   "Text1"
         Top             =   5400
         Width           =   1020
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
         Index           =   1
         Left            =   -68040
         TabIndex        =   210
         Top             =   4440
         Width           =   1005
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   660
            Index           =   1
            Left            =   120
            TabIndex        =   211
            Top             =   150
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1164
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
         Index           =   51
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   208
         Text            =   "51"
         Top             =   4560
         Width           =   5445
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
         Index           =   51
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   40
         Tag             =   "intermediario|N|S|0|999|scaalb|codinter|000||"
         Text            =   "Text1"
         Top             =   4560
         Width           =   1020
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
         Height          =   3780
         Left            =   -73560
         TabIndex        =   170
         Top             =   1200
         Width           =   9495
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
            Index           =   54
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   193
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   1485
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
            Index           =   51
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   192
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   645
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
            Index           =   53
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   191
            Text            =   "Text1 7"
            Top             =   2040
            Width           =   1485
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
            Index           =   50
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   190
            Text            =   "Text1 7"
            Top             =   2040
            Width           =   645
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
            Index           =   52
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   189
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
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
            Index           =   49
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   188
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   645
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
            TabIndex        =   187
            Text            =   "Text1 7"
            Top             =   480
            Width           =   1485
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
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   186
            Text            =   "Text1 7"
            Top             =   480
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
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   185
            Text            =   "Text1 7"
            Top             =   480
            Width           =   1365
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
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   184
            Text            =   "Text1 7"
            Top             =   480
            Width           =   1485
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
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   183
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1260
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   182
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   645
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
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   181
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   645
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
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   180
            Text            =   "Text1 7"
            Top             =   1560
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
            Index           =   44
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   179
            Text            =   "Text1 7"
            Top             =   2040
            Width           =   1260
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   178
            Text            =   "Text1 7"
            Top             =   2040
            Width           =   645
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
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   177
            Text            =   "Text1 7"
            Top             =   2040
            Width           =   645
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
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   176
            Text            =   "Text1 7"
            Top             =   2040
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
            Index           =   45
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   175
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   1260
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   174
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   645
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
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   173
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   645
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
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   172
            Text            =   "Text1 7"
            Top             =   2520
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
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
            Index           =   55
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   171
            Text            =   "Text1 7"
            Top             =   3240
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
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
            Left            =   6480
            TabIndex        =   207
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
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
            Left            =   7560
            TabIndex        =   206
            Top             =   1200
            Width           =   1455
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
            Index           =   28
            Left            =   2520
            TabIndex        =   205
            Top             =   1200
            Width           =   1575
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
            Index           =   27
            Left            =   360
            TabIndex        =   204
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto.PP"
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
            Left            =   2280
            TabIndex        =   203
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto.Gn"
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
            Left            =   3960
            TabIndex        =   202
            Top             =   120
            Width           =   1455
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
            Index           =   2
            Left            =   5760
            TabIndex        =   201
            Top             =   120
            Width           =   1695
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
            TabIndex        =   200
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
            TabIndex        =   199
            Top             =   480
            Width           =   135
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
            TabIndex        =   198
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
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
            Left            =   5160
            TabIndex        =   197
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   8880
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "TOTAL ALBARAN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   39
            Left            =   4320
            TabIndex        =   196
            Top             =   3240
            Width           =   2610
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
            Left            =   4080
            TabIndex        =   195
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
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
            Left            =   1920
            TabIndex        =   194
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame FrameHco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -63840
         TabIndex        =   160
         Top             =   1200
         Width           =   6375
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
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   165
            Top             =   240
            Width           =   1785
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
            Left            =   135
            MaxLength       =   30
            TabIndex        =   164
            Text            =   "Text1"
            Top             =   840
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
            Index           =   32
            Left            =   1035
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   163
            Text            =   "Text2"
            Top             =   840
            Width           =   5085
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
            Index           =   33
            Left            =   135
            MaxLength       =   30
            TabIndex        =   162
            Text            =   "Text1"
            Top             =   1560
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
            Index           =   33
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   161
            Text            =   "Text2"
            Top             =   1560
            Width           =   5085
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   3120
            TabIndex        =   169
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
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
            TabIndex        =   168
            Top             =   240
            Width           =   1815
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
            Height          =   240
            Index           =   38
            Left            =   120
            TabIndex        =   167
            Top             =   615
            Width           =   1065
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1320
            ToolTipText     =   "Buscar trabajador"
            Top             =   600
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
            Height          =   240
            Index           =   40
            Left            =   120
            TabIndex        =   166
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1320
            ToolTipText     =   "Buscar incidencia"
            Top             =   1320
            Width           =   240
         End
      End
      Begin VB.Frame FrameFacRec 
         Caption         =   "Datos Factura a rectificar "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -73560
         TabIndex        =   152
         Top             =   5280
         Width           =   9495
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
            Index           =   35
            Left            =   7560
            MaxLength       =   10
            TabIndex        =   155
            Tag             =   "Fecha Factura|F|S|||scaalb|fecfactu|dd/mm/yyyy|N|"
            Top             =   360
            Width           =   1425
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
            Index           =   36
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   154
            Tag             =   "Nº. Factura|N|S|0||scaalb|numfactu|0000000|N|"
            Top             =   360
            Width           =   1305
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
            Index           =   37
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   153
            Tag             =   "Tipo Mov. Factura|T|S|||scaalb|codtipmf||N|"
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fact."
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
            Index           =   44
            Left            =   6120
            TabIndex        =   158
            Top             =   420
            Width           =   1200
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
            Index           =   46
            Left            =   3240
            TabIndex        =   157
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Mov."
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
            Left            =   960
            TabIndex        =   156
            Top             =   420
            Width           =   855
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
         Index           =   40
         Left            =   -71160
         MaxLength       =   7
         TabIndex        =   149
         Tag             =   "Descuento General|N|S|||scaalb|aportacion|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   600
         Width           =   1140
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
         TabIndex        =   147
         Top             =   3120
         Width           =   2325
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   148
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
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
                  Object.ToolTipText     =   "Lotes"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   -68000
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   136
         Text            =   "Text2"
         Top             =   2400
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame FrameCampos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -67680
         TabIndex        =   132
         Top             =   600
         Width           =   10335
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
            Left            =   120
            TabIndex        =   219
            Top             =   150
            Width           =   1005
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   220
               Top             =   150
               Width           =   735
               _ExtentX        =   1296
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
                     Object.Visible         =   0   'False
                     Object.ToolTipText     =   "Modificar"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Eliminar"
                  EndProperty
               EndProperty
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   120
            TabIndex        =   133
            Top             =   720
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Campo"
               Object.Width           =   2381
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Partida"
               Object.Width           =   5230
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Variedad"
               Object.Width           =   5247
            EndProperty
         End
         Begin VB.Label lblFramePp 
            Caption         =   "Campos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   1
            Left            =   1200
            TabIndex        =   221
            Top             =   240
            Width           =   5265
         End
      End
      Begin VB.Frame FrameManipulador 
         Caption         =   "Manipulador fitosanitarios  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74760
         TabIndex        =   121
         Top             =   600
         Width           =   6855
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
            Index           =   45
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   123
            Tag             =   "ManipuladorNumCarnet|T|S|||scaalb|ManipuladorNumCarnet||N|"
            Text            =   "123456789"
            Top             =   480
            Width           =   1815
         End
         Begin VB.Frame FrameMani2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   240
            TabIndex        =   122
            Top             =   1080
            Width           =   6375
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
               Left            =   480
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   131
               Text            =   "Tiene bajo el text1 vinculado"
               Top             =   2370
               Width           =   4680
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   48
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   126
               Tag             =   "TipoCarnet|N|S|||scaalb|TipoCarnet||N|"
               Text            =   "123456789"
               Top             =   2400
               Width           =   255
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
               Index           =   47
               Left            =   480
               MaxLength       =   15
               TabIndex        =   125
               Tag             =   "ManipuladorFecCaducidad|F|S|||scaalb|ManipuladorFecCaducidad||N|"
               Text            =   "123456789"
               Top             =   1440
               Width           =   1710
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
               Index           =   46
               Left            =   480
               MaxLength       =   60
               TabIndex        =   124
               Tag             =   "ManipuladorNombre|T|S|||scaalb|ManipuladorNombre||N|"
               Text            =   "123456789"
               Top             =   600
               Width           =   5175
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo carnet"
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
               Index           =   59
               Left            =   120
               TabIndex        =   130
               Top             =   2040
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha caducidad"
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
               Index           =   58
               Left            =   120
               TabIndex        =   129
               Top             =   1200
               Width           =   1680
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nombre"
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
               Index           =   57
               Left            =   120
               TabIndex        =   128
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Carnet"
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
            Index           =   56
            Left            =   360
            TabIndex        =   127
            Top             =   480
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1680
            ToolTipText     =   "Buscar cliente varios"
            Top             =   480
            Width           =   240
         End
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
         Height          =   360
         Index           =   12
         Left            =   13200
         MaxLength       =   15
         TabIndex        =   120
         Text            =   "comision"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
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
         Height          =   1365
         Index           =   44
         Left            =   -65400
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   47
         Tag             =   "ObIn|T|S|||scaalb|observacrm||N|"
         Top             =   2880
         Width           =   7965
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
         Index           =   43
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   117
         Text            =   "Text2"
         Top             =   2400
         Width           =   5445
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
         Index           =   43
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   36
         Tag             =   "Cod. zona|N|S|0||scaalb|codzonas|000|N|"
         Text            =   "Text1"
         Top             =   2400
         Width           =   1020
      End
      Begin VB.CheckBox chkDocArchi 
         Caption         =   "Doc. Archivado"
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
         Left            =   -74880
         TabIndex        =   37
         Tag             =   "Docar|N|N|||scaalb|docarchiv||N|"
         Top             =   3000
         Width           =   2055
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
         Index           =   41
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Fecha envio|F|S|||scaalb|fecenvio|dd/mm/yyyy|N|"
         Top             =   3720
         Width           =   1545
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
         Height          =   360
         Index           =   11
         Left            =   12600
         MaxLength       =   15
         TabIndex        =   112
         Text            =   "numlote"
         Top             =   3960
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
         Height          =   360
         Index           =   10
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   63
         Tag             =   "Bultos"
         Text            =   "12345"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
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
         Index           =   9
         Left            =   12360
         TabIndex        =   109
         ToolTipText     =   "Buscar proveedor"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
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
         Height          =   360
         Index           =   9
         Left            =   11640
         MaxLength       =   6
         TabIndex        =   69
         Tag             =   "proveedor"
         Text            =   "codc"
         Top             =   3960
         Visible         =   0   'False
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
         Index           =   39
         Left            =   -66720
         MaxLength       =   7
         TabIndex        =   33
         Tag             =   "Nº Venta|N|S|||scaalb|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   720
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
         Index           =   29
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   38
         Tag             =   "Cod. Envío|N|N|0|999999|scaalb|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   3720
         Width           =   1020
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
         Index           =   29
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   99
         Text            =   "Text2"
         Top             =   3720
         Width           =   5445
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
         Left            =   -70320
         MaxLength       =   30
         TabIndex        =   35
         Tag             =   "Preparador Material|N|N|0|9999|scaalb|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1560
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
         Index           =   28
         Left            =   -69360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   1560
         Width           =   3645
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
         Index           =   27
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "Trabajador pedido|N|S|0|9999|scaalb|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1560
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
         Index           =   27
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   95
         Text            =   "Text2"
         Top             =   1560
         Width           =   3525
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
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Semana Entrega|N|S|||scaalb|sementre||N|"
         Top             =   720
         Width           =   705
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
         Left            =   -74880
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Nº Pedido|N|S|||scaalb|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   720
         Width           =   1125
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
         Left            =   -73680
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Fecha Pedido|F|S|||scaalb|fecpedcl|dd/mm/yyyy|N|"
         Top             =   720
         Width           =   1545
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
         Left            =   -69840
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "Fecha Oferta|F|S|||scaalb|fecofert|dd/mm/yyyy|N|"
         Top             =   720
         Width           =   1305
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
         Left            =   -71040
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Nº Oferta|N|S|||scaalb|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   720
         Width           =   1125
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
         Height          =   360
         Index           =   5
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   65
         Tag             =   "OP"
         Text            =   "OF"
         Top             =   3960
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
         Height          =   2775
         Left            =   200
         TabIndex        =   74
         Top             =   360
         Width           =   17460
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
            Index           =   50
            Left            =   15960
            MaxLength       =   30
            TabIndex        =   26
            Tag             =   "Bultos|N|S|||scaalb|pesoalba||N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
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
            Index           =   49
            Left            =   14040
            MaxLength       =   30
            TabIndex        =   25
            Tag             =   "Bultos|N|S|0|32000|scaalb|numbultos||N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox chkPideCliente 
            Caption         =   "Pedido por cliente"
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
            Left            =   10920
            TabIndex        =   20
            Tag             =   "P|N|N|||scaalb|PideCliente||N|"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox chkConTransporte 
            Caption         =   "Con transporte"
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
            Left            =   13320
            TabIndex        =   24
            Tag             =   "Trans|N|N|||scaalb|tipAlbaran||N|"
            Top             =   1800
            Width           =   2055
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
            Index           =   42
            Left            =   8640
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   1800
            Width           =   4365
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
            Index           =   42
            Left            =   7680
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Dir envio|N|S|0|9999|scaalb|coddiren|000|N|"
            Text            =   "Text1"
            Top             =   1800
            Width           =   900
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
            Index           =   34
            Left            =   14760
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "Cant. Km|N|S|0|99999|scaalb|cantidkm||N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   950
         End
         Begin VB.CheckBox chkFacturarKm 
            Caption         =   "Facturar Km"
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
            Left            =   13320
            TabIndex        =   22
            Tag             =   "Facturar Km|N|N|||scaalb|facturkm||N|"
            Top             =   840
            Visible         =   0   'False
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
            Index           =   13
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Referencia Cliente|T|S|||scaalb|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   2280
            Width           =   3645
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
            Left            =   8640
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   86
            Tag             =   "Direccion/Dpto.|T|S|||scaalb|nomdirec||N|"
            Text            =   "Text2"
            Top             =   360
            Width           =   4365
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
            Left            =   7680
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaalb|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   900
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
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scaalb|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1800
            Width           =   2445
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
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scaalb|codpobla||N|"
            Text            =   "Text15"
            Top             =   1320
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
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scaalb|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1320
            Width           =   3765
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
            Left            =   4320
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scaalb|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   360
            Width           =   1725
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
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scaalb|nifclien||N|"
            Text            =   "123456789"
            Top             =   360
            Width           =   2055
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
            Index           =   17
            Left            =   7680
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Cod. Agente|N|N|0|9999|scaalb|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   840
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
            Index           =   17
            Left            =   8640
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   81
            Text            =   "Text2"
            Top             =   840
            Width           =   4365
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
            Left            =   7680
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Forma de Pago|N|N|0|999|scaalb|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1320
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
            Index           =   14
            Left            =   8640
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   76
            Text            =   "Text2"
            Top             =   1320
            Width           =   4365
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
            Left            =   7680
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalb|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2280
            Width           =   780
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
            Left            =   9840
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scaalb|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2280
            Width           =   780
         End
         Begin VB.ComboBox cboFacturacion 
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
            ItemData        =   "frmFacEntAlbaranesGR.frx":00D5
            Left            =   15000
            List            =   "frmFacEntAlbaranesGR.frx":00D7
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
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
            Index           =   8
            Left            =   1320
            MaxLength       =   60
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scaalb|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   840
            Width           =   4755
         End
         Begin VB.Label Label1 
            Caption         =   "Peso(Kg)"
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
            Index           =   64
            Left            =   15090
            TabIndex        =   151
            Top             =   2295
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Bultos"
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
            Index           =   62
            Left            =   13320
            TabIndex        =   134
            Top             =   2295
            UseMnemonic     =   0   'False
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   7440
            ToolTipText     =   "Dirección envio"
            Top             =   1860
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Dir. envio"
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
            Index           =   53
            Left            =   6240
            TabIndex        =   116
            Top             =   1793
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Km a facturar"
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
            Index           =   43
            Left            =   13320
            TabIndex        =   108
            Top             =   405
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
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
            Left            =   120
            TabIndex        =   91
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   1080
            ToolTipText     =   "Buscar población"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
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
            Left            =   6240
            TabIndex        =   88
            Top             =   420
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   7440
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   420
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
            Height          =   240
            Index           =   17
            Left            =   120
            TabIndex        =   87
            Top             =   1800
            Width           =   885
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
            Height          =   240
            Index           =   16
            Left            =   120
            TabIndex        =   85
            Top             =   1320
            Width           =   930
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
            Height          =   240
            Index           =   19
            Left            =   3450
            TabIndex        =   84
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "N.I.F."
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
            Index           =   20
            Left            =   120
            TabIndex        =   83
            Top             =   420
            Width           =   555
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1080
            ToolTipText     =   "Buscar cliente varios"
            Top             =   420
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   7440
            ToolTipText     =   "Buscar agente"
            Top             =   900
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
            Height          =   240
            Index           =   15
            Left            =   6240
            TabIndex        =   80
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.Pago"
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
            Left            =   6240
            TabIndex        =   79
            Top             =   2295
            Width           =   1335
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
            Left            =   8880
            TabIndex        =   78
            Top             =   2295
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo facturacion"
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
            Left            =   13320
            TabIndex        =   77
            Top             =   1305
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   7440
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1350
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
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   75
            Top             =   870
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
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
            Index           =   34
            Left            =   6240
            TabIndex        =   82
            Top             =   870
            Width           =   945
         End
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
         TabIndex        =   73
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
         TabIndex        =   72
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
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
         Height          =   360
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3165
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
         Height          =   360
         Index           =   8
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   68
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
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
         Height          =   360
         Index           =   7
         Left            =   9960
         MaxLength       =   30
         TabIndex        =   67
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
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
         Height          =   360
         Index           =   6
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   66
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
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
         Height          =   360
         Index           =   4
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   64
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
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
         Height          =   360
         Index           =   3
         Left            =   6120
         MaxLength       =   16
         TabIndex        =   62
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
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
         Height          =   360
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   60
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
         Left            =   360
         MaxLength       =   15
         TabIndex        =   59
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
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
         Index           =   22
         Left            =   -65400
         MaxLength       =   80
         TabIndex        =   46
         Tag             =   "Observación 5|T|S|||scaalb|observa05||N|"
         Top             =   2160
         Width           =   7965
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
         Left            =   -65400
         MaxLength       =   80
         TabIndex        =   45
         Tag             =   "Observación 4|T|S|||scaalb|observa04||N|"
         Top             =   1800
         Width           =   7965
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
         Left            =   -65400
         MaxLength       =   80
         TabIndex        =   44
         Tag             =   "Observación 3|T|S|||scaalb|observa03||N|"
         Top             =   1440
         Width           =   7965
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
         Left            =   -65400
         MaxLength       =   80
         TabIndex        =   43
         Tag             =   "Observación 2|T|S|||scaalb|observa02||N|"
         Top             =   1080
         Width           =   7965
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
         Left            =   -65400
         MaxLength       =   80
         TabIndex        =   42
         Tag             =   "Observación 1|T|S|||scaalb|observa01||N|"
         Top             =   720
         Width           =   7965
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntAlbaranesGR.frx":00D9
         Height          =   2880
         Left            =   240
         TabIndex        =   71
         Top             =   3720
         Width           =   17220
         _ExtentX        =   30374
         _ExtentY        =   5080
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   16
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
            Size            =   9
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
         Index           =   38
         Left            =   -67440
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Nº terminal|N|S|||scaalb|numtermi||N|"
         Top             =   720
         Width           =   705
      End
      Begin VB.CheckBox chkImpreso 
         Caption         =   "Impreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65280
         TabIndex        =   52
         Tag             =   "Impr|N|N|||scaalb|albImpreso||N|"
         Top             =   2880
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1935
         Left            =   -66960
         TabIndex        =   212
         Top             =   4560
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3413
         _Version        =   393216
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
               LCID            =   1034
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
               LCID            =   1034
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
         Index           =   18
         Left            =   -73560
         ToolTipText     =   "Buscar forma de envio"
         Top             =   6000
         Width           =   240
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
         Index           =   60
         Left            =   -74880
         TabIndex        =   224
         Top             =   6000
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Index           =   67
         Left            =   -62520
         TabIndex        =   215
         Top             =   4320
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Naturaleza"
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
         Index           =   66
         Left            =   -74880
         TabIndex        =   214
         Top             =   5160
         Width           =   1290
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   -73440
         ToolTipText     =   "Buscar forma de envio"
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Intermediario"
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
         Index           =   36
         Left            =   -74880
         TabIndex        =   209
         Top             =   4320
         Width           =   1290
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   -73320
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "TPV(ticket)"
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
         Index           =   65
         Left            =   -67440
         TabIndex        =   159
         ToolTipText     =   "Terminal / Nor de ticket"
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "APORTACION TERMINAL"
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
         Left            =   -73560
         TabIndex        =   150
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Puntos"
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
         Index           =   63
         Left            =   -67995
         TabIndex        =   135
         Top             =   2160
         Visible         =   0   'False
         Width           =   675
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
         Height          =   240
         Index           =   55
         Left            =   -65400
         TabIndex        =   119
         Top             =   2640
         Width           =   3480
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Zona"
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
         Index           =   54
         Left            =   -74880
         TabIndex        =   118
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   -73680
         ToolTipText     =   "Buscar forma de envio"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -73440
         ToolTipText     =   "Buscar forma de envio"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   40
         Left            =   -66720
         Picture         =   "frmFacEntAlbaranesGR.frx":00EE
         ToolTipText     =   "Buscar fecha"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha envio"
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
         Index           =   52
         Left            =   -68040
         TabIndex        =   113
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Envío"
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
         Index           =   24
         Left            =   -74880
         TabIndex        =   100
         Top             =   3480
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Preparador Material"
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
         Index           =   23
         Left            =   -70320
         TabIndex        =   98
         Top             =   1320
         Width           =   1920
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -68280
         ToolTipText     =   "Buscar trabajador"
         Top             =   1320
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
         Height          =   240
         Index           =   9
         Left            =   -74880
         TabIndex        =   96
         Top             =   1260
         Width           =   1785
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -73080
         ToolTipText     =   "Buscar trabajador"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sem."
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
         Left            =   -72120
         TabIndex        =   94
         ToolTipText     =   "Semana de entrega"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   " Pedido"
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
         Left            =   -74880
         TabIndex        =   93
         Top             =   480
         Width           =   960
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
         Height          =   240
         Index           =   10
         Left            =   -73680
         TabIndex        =   92
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
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
         Left            =   -69840
         TabIndex        =   90
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Oferta nº"
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
         Left            =   -71040
         TabIndex        =   89
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones "
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
         Index           =   45
         Left            =   -65400
         TabIndex        =   58
         Top             =   480
         Width           =   2400
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
      Left            =   16530
      TabIndex        =   53
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   101
      Top             =   840
      Width           =   17775
      Begin VB.CheckBox chkFacturar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   51
         Tag             =   "Facturar|N|N|||scaalb|factursn||N|"
         Top             =   330
         Width           =   255
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
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   300
         Width           =   1005
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
         Left            =   6120
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "Nombre Cliente|T|N|||scaalb|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   300
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
         Index           =   3
         Left            =   12120
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Realizada Por|N|N|0|9999|scaalb|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   300
         Width           =   760
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
         Left            =   12960
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   106
         Text            =   "Text2"
         Top             =   300
         Width           =   4560
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
         Index           =   30
         Left            =   1440
         TabIndex        =   1
         Tag             =   "Tipo Albaran|T|N|||scaalb|codtipom||S|"
         Text            =   "Text3"
         Top             =   300
         Width           =   615
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Albaran|F|N|||scaalb|fechaalb|dd/mm/yyyy|N|"
         Top             =   300
         Width           =   1425
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
         Left            =   120
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Albaran|N|S|0||scaalb|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Facturar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   3
         Top             =   390
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5880
         ToolTipText     =   "Buscar cliente"
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Left            =   5040
         TabIndex        =   107
         Top             =   60
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Realizada Por"
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
         Index           =   21
         Left            =   12120
         TabIndex        =   105
         Top             =   60
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   13440
         ToolTipText     =   "Buscar trabajador"
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fec. Alb."
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
         Left            =   2160
         TabIndex        =   104
         Top             =   60
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3120
         Picture         =   "frmFacEntAlbaranesGR.frx":0179
         ToolTipText     =   "Buscar fecha"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   103
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   102
         Top             =   60
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   1560
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
      Caption         =   "data3"
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
   Begin VB.Label LblMostr 
      BackStyle       =   0  'Transparent
      Caption         =   "MOSTRADOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   8760
      TabIndex        =   146
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   14
      Left            =   5160
      ToolTipText     =   "Ver ampliación"
      Top             =   8640
      Width           =   240
   End
   Begin VB.Label lblF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   114
      Top             =   9000
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Index           =   51
      Left            =   3960
      TabIndex        =   111
      Top             =   9000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   57
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
Attribute VB_Name = "frmFacEntAlbaranesGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran  de Venta de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALV,ALR,ALS)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
                        
Public AlbAvisoGenerado As Long 'Cuando desde aviso cierro reparacion, creo un albaran y llamo a este form
                                'Entonces lo cargo el albaran y lo meto insertando lineas
                                
'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmFacClientes3 'Form M7to Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1

Private WithEvents frmInt As frmPortesIntermediario
Attribute frmInt.VB_VarHelpID = -1
Private WithEvents frmNat As frmPortesNaturaleza
Attribute frmNat.VB_VarHelpID = -1
Private WithEvents frmMatr As frmPortesMatriculasChofer
Attribute frmMatr.VB_VarHelpID = -1


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

'   6.-

'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Byte   '0 cabecera   1-direc    2 direnv
'Para saber en MandaBusquedaPrevia si busca en la tabla scaalb o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean



Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim cadList As String 'cadena para pasar al historico
Dim motivo As String 'cadena para el motivo si es factura Rectificativa


Dim PulsadoMas2 As Boolean

Dim txtAnterior As String

Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab


'PORTES
' Tipo fontenas
Dim KilosAnteriores As Currency
Dim RutaCliente As Integer
Dim ZonaCliente As Integer

Dim LineaIntercalar As Integer 'NO reutilizar


Dim AlmacenLineas As Integer
'Dim ElArticulo As String

'Para buscar por los chks
Private BuscaChekc As String


'Si lleva control LOG de quien cambia precio
Dim GrabaLogCambioPrecioDto As Boolean
Dim VendeAMenorPrecio As Byte   ' 0.- Normal     1.- Menor precio     2-super eco
Dim GrabaCambioTrabajador As Integer  '-1: No hay cambio  Si no inidcara que trabajador habia


'Herbelca. Nuevo sistema comisiones
Dim ComisionCliente As Currency
Dim vAgent As cAgente



'Para los puntos. El canje solo deberia hacerse en albaran NUEVO
Dim EsNuevoAlbaran As Boolean

Dim MostrarComision As Boolean 'SOLO herbelca, y super usuarios, veran las comisiones

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    'PonerFocoBtn cmdAceptar
End Sub




Private Sub chkConTransporte_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkConTransporte, BuscaChekc
End Sub
Private Sub chkConTransporte_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkDocArchi_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkDocArchi, BuscaChekc
End Sub
Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkFacturar_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkFacturar, BuscaChekc
End Sub
Private Sub chkFacturar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkFacturarKm_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkFacturarKm, BuscaChekc
End Sub

Private Sub chkFacturarKm_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkImpreso_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkImpreso, BuscaChekc
End Sub
Private Sub chkImpreso_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPideCliente_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkPideCliente, BuscaChekc
End Sub

Private Sub chkPideCliente_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificarCabAlbaran Then
                
                
                
                    TerminaBloquear
                    
                    UpdateaNomDirec
                    
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea(numlinea, False) Then
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    If PrimeraLin Then
                        CargaGrid DataGrid1, data2, True
                    Else
                        CargaGrid2 DataGrid1, data2
                    End If
                    
                    If LineaIntercalar > 0 Then
                        'HA intercalado la linea. Ponemos luego en normal
                        Me.DataGrid1.Enabled = True
                        DataGrid1.AllowAddNew = False
                        NumRegElim = LineaIntercalar
                        CargaTxtAux False, False
                        CargaGrid2 DataGrid1, data2
                        PosicionarData2 1
                        ModificaLineas = 0
                        PonerBotonCabecera True
                        BloquearTxt Text2(16), True
                        cmdRegresar_Click
                    Else
                        'Que meta otra
                        PrimeraLin = True
                        If vParamAplic.PtosAsignar > 0 Then
                            'Si la linea es la de canje NO hacemos mas
                            If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then PrimeraLin = False
                        End If
                        
                        If PrimeraLin Then
                            BotonAnyadirLinea False
                        Else
                            cmdCancelar_Click
                        End If
                    End If
                    
                    
                    
                    
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    numlinea = data2.Recordset!numlinea
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    TerminaBloquear
                    NumRegElim = Val(data2.Recordset!numlinea)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, data2
                    PosicionarData2 1
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(9), True
                    lblF.Caption = ""
                    cmdCancelar_Click
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
        Case 6
            'Matriculas
             
             
                If InsertarModificarMatricula() Then
                    
                        TerminaBloquear
                        Me.DataGrid2.Enabled = True
                        DataGrid2.AllowAddNew = False
                        
                        CargaGrid2 DataGrid2, data3
                        PosicionarData2 2
                        
                        CargaTxtAux2 False, False
                        
                        lblF.Caption = ""
                        ModificaLineas = 0
                        PonerBotonCabecera True
                        cmdCancelar_Click
                End If
                Me.DataGrid2.Enabled = True
            
            
            
            
            
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ModificarCabAlbaran() As Boolean
Dim B As Boolean
Dim SQL As String

    On Error GoTo EModificaAlb
    conn.BeginTrans
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
    B = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    
    If B Then
        B = ModificaDesdeFormulario(Me, 1)
        If B Then
            SQL = "UPDATE scaalb SET nomdirec=" & DBSet(Text2(12).Text, "T") & " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!NumAlbar
            conn.Execute SQL
        End If

        If B Then
            'comprobar si se ha cambiado el cliente
            'o si se ha cambiado la fecha del albaran
            'If (CInt(Me.Data1.Recordset!CodClien) <> CInt(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
            'DAVID.   No es un CINT. Tiene que ser un clng o val
            If (Val(Me.Data1.Recordset!codClien) <> Val(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
                'si hay numeros de serie en ese albaran, actualizamos el cliente
                'al nuevo cliente
                SQL = "UPDATE sserie SET codclien=" & DBSet(Text1(4).Text, "N") & ","
                SQL = SQL & " fechavta=" & DBSet(Text1(1).Text, "F")
                SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
                SQL = SQL & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!NumAlbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute SQL
                
                'Modificar el cliente en la smoval
                SQL = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N") & ","
                SQL = SQL & " fechamov=" & DBSet(Text1(1).Text, "F")
                SQL = SQL & ", horamovi= concat(" & DBSet(Text1(1).Text, "F") & ",' ',hour(horamovi),':',minute(horamovi),':',second(horamovi))"
                SQL = SQL & " WHERE detamovi='" & CodTipoMov & "'"
                SQL = SQL & " AND document='" & Text1(0).Text & "'"
                SQL = SQL & " and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute SQL
            End If
            
            
            
            'LOG GrabaCambioTrabajador
            If GrabaCambioTrabajador >= 0 Then
                'Ha cambiado el trabajador
                '------------------------------------------------------------------------------
                '  LOG de acciones.
                Set LOG = New cLOG
                SQL = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", CStr(GrabaCambioTrabajador))
                SQL = "Antes: " & Format(GrabaCambioTrabajador, "0000") & " " & SQL
                SQL = "Ahora: " & Text1(3).Text & " " & Text2(3).Text & vbCrLf & SQL
                SQL = "Albaran: " & Text1(0).Text & " Fecha " & Text1(1).Text & vbCrLf & SQL
                LOG.Insertar 20, vUsu, SQL
                Set LOG = Nothing
            End If
        End If
    
    
    
        If vParamAplic.PtosAsignar > 0 Then
            'Sistema de puntos
   
  
                
                If Val(Data1.Recordset!codClien) <> Val(Text1(4).Text) Then
                    'Si cambia el cliente, hay que ver
                    SQL = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text)
                    If Val(SQL) = "1" Then
                        'El nuevo cliente tiene puntos
                        SQL = "+"
                        If Data1.Recordset!Puntos < 0 Then SQL = "-"
                        SQL = "UPDATE sclien set puntos=coalesce(puntos,0) " & SQL & DBSet(Abs(Data1.Recordset!Puntos), "N")
                        SQL = SQL & " WHERE codclien =" & Text1(4).Text
                        conn.Execute SQL
                    
                        BuscaChekc = "U"
                       
                    Else
                        SQL = "UPDATE scaalb SET puntos=0 WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!NumAlbar
                        conn.Execute SQL
                        BuscaChekc = "D"
                        Text2(1).Text = ""
                        
                    End If
                    
                    
                    'Le quito los puntos al cliente origen. Osea al reves de arriba
                    SQL = "-"
                    If Data1.Recordset!Puntos < 0 Then SQL = "+"
                    SQL = "UPDATE sclien set puntos=puntos " & SQL & DBSet(Abs(Data1.Recordset!Puntos), "N")
                    SQL = SQL & " WHERE codclien =" & Data1.Recordset!codClien
                    conn.Execute SQL
                    
                    
                    
                    'O borro o updateo movimientos puntos
                    SQL = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                    SQL = SQL & " AND codclien = " & Data1.Recordset!codClien
                    If BuscaChekc = "U" Then
                        BuscaChekc = "UPDATE smovalpuntos SET codclien = " & Text1(4).Text
                    Else
                        BuscaChekc = "DELETE FROM smovalpuntos "
                    End If
                    SQL = BuscaChekc & SQL
                    
                    ejecutar SQL, False
                    
                    BuscaChekc = ""
                    
                    
                End If
                If CDate(Me.Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                    SQL = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                    SQL = SQL & " AND codclien = " & Data1.Recordset!codClien
                    SQL = "UPDATE smovalpuntos set fechaalb= " & DBSet(Text1(1).Text, "F") & SQL
                    ejecutar SQL, False
                    
                End If

            End If
                
    
    
    End If
    
    
    
    
    
EModificaAlb:
    If Err.Number <> 0 Then B = False
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarCabAlbaran = B
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar cabecera Albaran.", Err.Description
End Function




Private Sub cmdAux_Click(Index As Integer)
Dim B As Boolean

    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            
        Case 1 'Busqueda de Cod. Artic
            B = True
            If CodTipoMov = "ART" Then
                If MsgBox("¿Desea traer líneas de la factura que va a rectificar?", vbQuestion + vbYesNo) = vbYes Then
                
                    'si es Albaran de Factura rectificativa cargar un listview con todas las
                    'lineas de la factura y marcar las que queremos seleccionar para
                    'cargarlas en las lineas del Albaran rectificativo
                    If Text1(36).Text = "" Then
                        MsgBox "No se ha encotrado la factura a la que rectifica", vbExclamation
                    Else
                        B = False
                        Set frmMen = New frmMensajes
                        frmMen.cadWhere = " codtipom=" & DBSet(Text1(37).Text, "T") & " and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
                        frmMen.OpcionMensaje = 11 'Lineas Factura a Rectificar
                        frmMen.Show vbModal
                        Set frmMen = Nothing
                        CargaGrid Me.DataGrid1, Me.data2, True
                        cmdCancelar_Click
                    End If
                End If
            End If
            
            
            If B Then
                Set FrmArt = New frmAlmArticu2
                'FrmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo busqueda
                FrmArt.DesdeTPV = False
                FrmArt.Show vbModal
                Set FrmArt = Nothing

            End If
            
    Case 9 'CENTRO COSTE/ PROVEEDOR
        If vEmpresa.TieneAnalitica Then
            'centro de coste
            EsCabecera = 3
            AbrirForm_CentroCoste
        Else
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "1"
            frmProv.Show vbModal
            Set frmProv = Nothing
        End If
    End Select
    PonerFoco txtAux(Index)
End Sub



Private Sub AbreFormMatrChofer(Matricula As Boolean)

    BuscaChekc = ""
    Set frmMatr = New frmPortesMatriculasChofer
    NumRegElim = -1
    If Trim(Text1(29).Text) <> "" Then NumRegElim = Val(Text1(29).Text)
    frmMatr.DatosADevolverBusqueda = "S"
    frmMatr.VerMatriculas = Matricula
    frmMatr.Transportista = NumRegElim
    frmMatr.Show vbModal
    Set frmMatr = Nothing

End Sub


Private Sub cmdAux2_Click(Index As Integer)
    If Modo <> 6 Then Exit Sub
    
    AbreFormMatrChofer True
    If BuscaChekc <> "" Then
        txtAux2(0).Text = RecuperaValor(BuscaChekc, 1)
        txtAux2(1).Text = "   " & RecuperaValor(BuscaChekc, 2)
        cmdAceptar_Click
    End If
End Sub

Private Sub cmdCancelar_Click()
Dim EraNuevaLinea As Boolean
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
            BloquearTxt Text2(9), True
            DataGrid1.Columns(4).Caption = "Artículo"
            EraNuevaLinea = False
            If ModificaLineas = 1 Then 'INSERTAR
                EraNuevaLinea = True
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                Text2(16).Text = ""
                If Not data2.Recordset.EOF Then data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LineaIntercalar = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            
            
            
            
            
            
            If Me.hcoCodTipoM = "ALM" And vParamAplic.EntradaRapidaFacturasMostrador Then
                'Estamos en facturas mostrador, cont entrada rapida. Simularemos el boton
                'de facturar. cuando pulse cancelar
                If EraNuevaLinea Then HacerToolbar 12
                
                
                
            Else
                cmdRegresar_Click
            End If
            
            
            
        Case 6
            CargaTxtAux2 False, False
            TerminaBloquear
            LineaIntercalar = 0
            DataGrid2.AllowAddNew = False
            If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
            PonerBotonCabecera True
            Me.DataGrid2.Enabled = True
            PonerModo 2
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim cad As String
Dim Rs As ADODB.Recordset
Dim TxtMotivoFra As String 'AMESA


    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Si es Albaran para factura RECTIFICATIVA pedir la Factura que se va
    'a Rectificar y si existe en el historico, tabla "scafac", entonces dejamos
    'que inserte el Albaran Rectificativo, si no salimos
    If CodTipoMov = "ART" Then
        cadList = ""

        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 225
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        
        If Trim(Mid(cadList, 1, 12)) = "codtipom=''" Then
            Unload Me
            Exit Sub
        End If
        
        'cargar los datos de la factura recuperada en el formulario
        NomTraba = "select codtipom as codtipmf,numfactu,fecfactu,codclien,nomclien,domclien,scafac.codpobla,pobclien,proclien,nifclien,telclien,"
        NomTraba = NomTraba & "coddirec,nomdirec,scafac.codagent,nomagent,scafac.codforpa, nomforpa,dtoppago,dtognral "  'JUNIO 2010 añado el envio
        NomTraba = NomTraba & " from (scafac inner join sforpa on scafac.codforpa=sforpa.codforpa) "
        NomTraba = NomTraba & " inner join sagent on scafac.codagent=sagent.codagent where " & cadList
        
        Set Rs = New ADODB.Recordset
        Rs.Open NomTraba, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        PonerModo 3
        
        If Not Rs.EOF Then
            Text1(4).Text = Rs!codClien
            FormateaCampo Text1(4)
            Text1(5).Text = Rs!NomClien
            Text1(6).Text = Rs!nifClien
            Text1(7).Text = DBLet(Rs!telclien, "T")
            Text1(8).Text = Rs!domclien
            Text1(9).Text = Rs!codpobla
            Text1(10).Text = Rs!pobclien
            Text1(11).Text = DBLet(Rs!proclien, "T")
            Text1(12).Text = DBLet(Rs!CodDirec, "T")
            FormateaCampo Text1(12)
            Text2(12).Text = DBLet(Rs!nomdirec, "T")
            Text1(14).Text = Rs!codforpa
            FormateaCampo Text1(14)
            Text2(14).Text = Rs!nomforpa
            Text1(15).Text = DBLet(Rs!DtoPPago, "N")
            FormateaCampo Text1(15)
            Text1(16).Text = DBLet(Rs!DtoGnral, "N")
            FormateaCampo Text1(16)
            Text1(17).Text = DBLet(Rs!CodAgent, "T")
            FormateaCampo Text1(17)
            Text2(17).Text = Rs!NomAgent
            Text1(37).Text = Rs!codtipmf
            Text1(36).Text = DBLet(Rs!Numfactu, "N")
            FormateaCampo Text1(36)
            Text1(35).Text = Rs!FecFactu
            
            
            Text1(18).Text = Rs!Numfactu & ", " & Rs!FecFactu
          
            
            NomTraba = "tipofact"
            cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            If cad = "0" Then BloquearDatosCliente (False)
            
            
            
            
            
            
            
            'recuperamos el tipo de facturacion del cliente
            Me.cboFacturacion.ListIndex = CInt(NomTraba)
            
            'Si es factura de AGUA, debe traer la referencia, que sera el contador
            If Rs!codtipmf = "FAG" Then
                'cadlist=>>  codtipom='FAG' and numfactu=0000001 and fecfactu='2014-05-16'
                cad = cadList & " AND 1"
                cad = DevuelveDesdeBD(conAri, "referenc", "scafac1", cad, "1 ORDER BY 1 DESC")
                Text1(13).Text = cad
            End If
            
            
            
            'ULTIMO
            'Memorizo cad con codtipom
            cad = Rs!codtipmf
            
            
            
            
            If vParamAplic.ManipuladorFitosanitarios2 Then
                Rs.Close
                NomTraba = "select ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet from scafac1 "
                NomTraba = NomTraba & " WHERE " & cadList & " ORDER BY 1 DESC"
                Rs.Open NomTraba, conn, adOpenForwardOnly, adCmdText
                If Not Rs.EOF Then
                    '
                    If Not IsNull(Rs!ManipuladorNumCarnet) Then
                        Me.Text1(45).Text = Rs!ManipuladorNumCarnet
                        Me.Text1(46).Text = Rs!ManipuladorNombre
                        Me.Text1(47).Text = Rs!ManipuladorFecCaducidad
                        Text2(0).Text = IIf(Rs!TipoCarnet = 2, "Cualificado", "Básico")
                        '
                        Me.Text1(48).Text = RecuperaValor(Rs!TipoCarnet, 4)
                        
                    End If
                End If
                
            End If
            
            Rs.Close
        Else
            cad = "N" 'para que la busqueda de despues no de error
            Text1(18).Text = ""
            Rs.Close
        End If
        
        
        'Observacion 2
        Text1(19).Text = motivo
        
        'DAVID
        'Para que meta la letra de serie, NO el tipo moviemiento
        Rs.Open "SELECT * FROM stipom WHERE codtipom='" & cad & "'"
        If Not Rs.EOF Then cad = DBLet(Rs!LetraSer, "T")
        Rs.Close
        If cad = "" Then cad = CodTipoMov
        If Text1(18).Text <> "" Then
            TxtMotivoFra = DevuelveDesdeBD(conAri, "texto", "sparaidioma", "codigo", 1)  '1.- Rectifica a
            If TxtMotivoFra = "" Then TxtMotivoFra = "RECTIFICA A FACTURA"
            Text1(18).Text = TxtMotivoFra & ": " & cad & ", " & Text1(18).Text
        End If
        
            
        'DAVID
        'JUNIO 2010
        'Envio por defecto del cliente
        If Text1(4).Text <> "" Then
            cad = "select sclien.codenvio,nomenvio from  sclien,senvio where sclien.codenvio=senvio.codenvio AND sclien.codclien= " & Text1(4).Text
            Rs.Open cad, conn, adOpenForwardOnly, adCmdText
            If Not Rs.EOF Then
                Text1(29).Text = Rs!CodEnvio
                Text2(29).Text = Rs!nomEnvio
            Else
                Text1(29).Text = ""
                Text2(29).Text = ""
            End If
            Rs.Close
            
            
           
            
            
            
        End If
            
        
        Me.chkFacturar.Value = 1
        
        Set Rs = Nothing
    Else
        'Añadiremos el boton de aceptar y demas objetos para insertar
        PonerModo 3
        
        
        If vParamAplic.NumeroInstalacion = 2 Then
            If MsgBox("Albarán realizado por el cliente?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                chkPideCliente.Value = 1
                Text1_GotFocus 1
            End If
        End If

         
        
    End If
    
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    'El preparador del material lo hacemos tb al trabajador actual
    Text1(28).Text = Text1(3).Text
    Text2(28).Text = Text2(3).Text


    'Marca de para facturar
    If vParamAplic.MarcarAlbaranFacturar Then Me.chkFacturar.Value = 1

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(30).Text = CodTipoMov
    
    If vParamAplic.CartaPortes Then
        Text1(51).Text = "1"
        'poner datos
        Text2(51).Text = PonerNombreDeCod(Text1(51), conAri, "sintermediario", "nominter", "codinter")
        
        
        
        
        
    End If
    
        
    
    'Mayo2014
    cad = "1"
    If CodTipoMov = "ALM" Then
        If vParamAplic.EntradaRapidaFacturasMostrador Then cad = "4"
    End If
    PonerFoco Text1(Val(cad))
End Sub


Private Sub BotonAnyadirLinea(Intercalando As Boolean)
Dim Aux As String
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
       
       
    If vParamAplic.PtosAsignar > 0 Then
        'Si ya esta el canje, no dejo insertar mas lineas
        Aux = Replace(ObtenerWhereCP(False), "scaalb", "slialb") & " AND codartic "
        Aux = DevuelveDesdeBD(conAri, "codartic", "slialb", Aux, vParamAplic.PtosArticuloCanje, "T")
        If Aux <> "" Then
            MsgBox "Ya esta el articulo de canje en este albarán." & vbCrLf & "No se pueden insertar mas lineas", vbExclamation
            Exit Sub
        End If
        
    End If
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    
    
    If Intercalando Then
        lblIndicador.Caption = "** INTERCALAR **"
        If Not data2.Recordset.EOF Then
            LineaIntercalar = data2.Recordset!numlinea
        End If
    Else
        LineaIntercalar = 0
        lblIndicador.Caption = "INSERTAR"
    End If
    
    
    
    
    AnyadirLinea DataGrid1, data2
    CargaTxtAux True, True
    'Poner el Almacen por defecto del Trabajador
    'txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    txtAux(0).Text = Format(AlmacenLineas, "000")
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    Text2(9).Text = ""
    
    BloquearTxt Text2(16), hcoCodTipoM = "DEV" ' False
    BloquearTxt Text2(9), True
   '' BloquearTxt txtAux(6), True
   ' BloquearTxt txtAux(7), True
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
    End If
    
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
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
    
    If Me.EsHistorico = False Then
        'Hacer busquedar del tipo de movimiento de albaran en el que estamos
        Text1(30).Text = CodTipoMov
        BloquearTxt Text1(30), True
    End If
End Sub


Private Sub BotonVerTodos()
Dim CadB As String
    
    CadB = " 1 = 1"
    If Not EsHistorico Then CadB = " codtipom='" & CodTipoMov & "'"

    If vUsu.CodigoAgente > 0 Then CadB = CadB & " AND codagent = " & vUsu.CodigoAgente
    


'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = 0
        
        MandaBusquedaPrevia CadB
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index - 1
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
   
    'Si es Cliente de Varios no se pueden modificar sus datos
    
    GrabaCambioTrabajador = -1
    
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If data2.Recordset.EOF Then Exit Sub
    
        
    'Herbelca. Tania 21/07/2016
    '--------------------------
    ' De varios no dejo modificar la linea. Segun ella esto ya lo hacia.
    'Version: 4_6_51 de Feb16 No lo hace    Solo era para eliminar linea
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.Nivel > 0 Then

            vWhere = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", CStr(data2.Recordset!codArtic), "T")
            If Val(vWhere) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
    End If
        
    
    
    
    
    'bloqueamos el registro a modificar
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    txtAux(0).BackColor = vbWhite
    
    'si es factura rectificativa y es una linea de la factura que rectificamos
    'solo podremos modificar la cantidad el resto de campos bloqueados
    If CodTipoMov = "ART" Then '(Albaran Rectificativo)
        vWhere = "codtipom='" & Text1(37).Text & "' and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
        vWhere = vWhere & " and codartic=" & DBSet(txtAux(1).Text, "T")
        vWhere = "SELECT COUNT(*) FROM slifac WHERE " & vWhere
        If RegistrosAListar(vWhere) > 0 Then
            'modificamos una linea de factura a rectificar y solo podemos modificar cantidad
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(1), True
            BloquearTxt txtAux(2), True
            BloquearTxt txtAux(4), True
            BloquearTxt txtAux(6), True
            BloquearTxt txtAux(7), True
            Me.cmdAux(0).Enabled = False
            Me.cmdAux(1).Enabled = False
        End If
    End If
    
    
    ' ---- [21/10/2009] [LAURA]: añadir campo centro de coste por trabajador
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
    End If
    
    
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    
    BloquearTxt Text2(16), hcoCodTipoM = "DEV" 'Campo Ampliacion Linea  Para los albarnes esta desbloqueado
    BloquearTxt Text2(9), True 'Campo nomprove
    BloquearTxt txtAux(2), True
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
Dim NumAlbElim As Long

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

        'Nov 2014
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        If vUsu.Nivel > 0 Then
            
            cad = "slialb.codartic=sartic.codartic and artvario=1 AND codtipom = '" & Data1.Recordset!codtipom & "' AND numalbar "
            cad = DevuelveDesdeBD(conAri, "count(*)", "slialb,sartic", cad, CStr(Data1.Recordset!NumAlbar))
            If Val(cad) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
    End If


    If vParamAplic.PtosAsignar > 0 Then
        cad = Replace(ObtenerWhereCP(False), "scaalb", "slialb") & " AND codartic "
        cad = DevuelveDesdeBD(conAri, "codartic", "slialb", cad, vParamAplic.PtosArticuloCanje, "T")
        If cad <> "" Then
            MsgBox "Tiene  articulo canje. ", vbExclamation
            Exit Sub
        End If

    End If
    If hcoCodTipoM = "DEV" Then
        cadList = "devolucion"
    Else
        cadList = "albarán"
    End If
    cad = "Cabecera de " & cadList & "." & vbCrLf
    cad = cad & "------------------------------------       " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar " & cadList & ":            "
    cad = cad & vbCrLf & "Tipo:  " & Text1(30).Text
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Text1(1).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
          
    If hcoCodTipoM <> "DEV" Then
        If Not ComprobarInventario Then Exit Sub
    End If
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
    
        If hcoCodTipoM <> "DEV" Then
            'Abrir frame de informes para pedir datos antes de grabar en el historico
            cadList = ""
            Set frmList = New frmListadoOfer
            frmList.OpcionListado = 80
            frmList.Show vbModal
            Set frmList = Nothing
            If cadList = "" Then Exit Sub
        
        End If
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumAlbElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(30).Text
        
        If Not Eliminar(NumAlbElim) Then
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            PosicionarDataTrasEliminar
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
Dim CodproveHerbelca As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If data2.Recordset.EOF Then Exit Sub
        
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        CodproveHerbelca = "codprove"
        
        SQL = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", CStr(data2.Recordset!codArtic), "T", CodproveHerbelca)
        If vUsu.Nivel > 0 Then
            
            If Val(SQL) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
        
        
        If CodproveHerbelca = 5000 Then
            'Proveedor de varios
             If vUsu.AlmacenPorDefecto2 > 1 And data2.Recordset!codArtic <> vParamAplic.PtosArticuloCanje Then
                MsgBox "No puede eliminar linea", vbExclamation
                Exit Sub
            End If
        End If
        
        
        'SI es de portes tampoco dejo
        If vParamAplic.ArtPortesN = CStr(data2.Recordset!codArtic) Then
            If vUsu.AlmacenPorDefecto2 > 1 Then
                MsgBox "No puede eliminar linea", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de Albaran?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & data2.Recordset!codArtic & " - " & data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            CargaGrid2 DataGrid1, data2
            SituarDataTrasEliminar data2, NumRegElim
            CalcularDatosFactura
            cmdRegresar_Click
        End If
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub




Private Sub BotonesCampos(Nuevo As Boolean)
    If Nuevo Then
        'Añadir mas campos
            CadenaDesdeOtroForm = ""
            frmADVvarios.Opcion = 0
            frmADVvarios.vCampos = Text1(4).Text
            frmADVvarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                
                
                    
                MultiInsercionCampos
                
                'Cargamos GRID
                
            End If
            CargaDatosCampos
                
    Else
        cadList = ""
        If Me.ListView1.ListItems.Count > 0 Then
            If Not Me.ListView1.SelectedItem Is Nothing Then
                cadList = "Va a eliminar el campo: "
                cadList = cadList & vbCrLf & "Codigo : " & Me.ListView1.SelectedItem.Text
                cadList = cadList & vbCrLf & "Partida : " & Me.ListView1.SelectedItem.SubItems(1)
                cadList = cadList & vbCrLf & "Variedad : " & Me.ListView1.SelectedItem.SubItems(2)
                cadList = cadList & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(cadList, vbQuestion + vbYesNo) = vbYes Then
                    'El tag tiene codcampo
                    cadList = "DELETE FROM slialbcampos WHERE  codtipom = " & DBSet(Data1.Recordset!codtipom, "T")
                    cadList = cadList & " AND numalbar = " & Data1.Recordset!NumAlbar
                    cadList = cadList & " AND codcampo  = " & CStr(Val(Me.ListView1.SelectedItem.Text))
                    conn.Execute cadList
                    
                    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
    
                End If
            End If
        End If
    End If
End Sub
Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String
Dim Port As Integer      'Port: para saber si ha metido/Modificado el articulo de portes
Dim Puntos As Currency
Dim PtosCliente As Currency
Dim PtosAnt As Currency
Dim Aux As String
    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
        'If vParamAplic.OperacionesAseguradas And Me.hcoCodTipoM <> "ALR" Then
        If vParamAplic.OperacionesAseguradas Then
            If Not Riesgo Then Exit Sub
        End If
            
            
        'Fontenas
        If vParamAplic.TipoPortes = 1 Then
            'Si lleva portes haremos varias cosas
            Port = HacerAccionesPortes
            CargaGrid DataGrid1, data2, True
            Set miRsAux = Nothing
        End If
    
        If vParamAplic.NumeroInstalacion = 2 Then ComprobarComisionesAlbaranes
        If vParamAplic.PtosAsignar > 0 Then
            cad = "puntos"
            If DevuelveDesdeBD(conAri, "tienepuntos", "sclien", "codclien", CStr(Data1.Recordset!codClien), "N", cad) = "1" Then
                
                If Me.hcoCodTipoM <> "ART" Then
                    Aux = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " AND 1 "
                    Aux = DevuelveDesdeBD(conAri, "concat(sum(if(codartic=" & DBSet(vParamAplic.PtosArticuloCanje, "T") & ",1,0)),'|',sum(importel),'|') ", "slialb", Aux, "1")
                    If Mid(Aux, 1, 1) = "1" Then
                        'llev articulo canje
                        Aux = RecuperaValor(Aux, 2)
                        If CCur(Aux) < 0 Then
                            'El albaran se queda en negativo. NO dejo continuar
                            MsgBox "El albaran se queda con importe en negativo. ", vbExclamation
                            Exit Sub
                        End If
                    End If
                End If
                If cad = "" Then
                    PtosCliente = 0
                Else
                    PtosCliente = CCur(cad)
                End If
                Puntos = CalcularPuntosAlbaran(Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas), Data1.Recordset!FechaAlb)
                
                
                If Me.Text2(1).Text = "" Then
                    PtosAnt = 0
                Else
                    PtosAnt = ImporteFormateado(Text2(1).Text)
                End If
                Text2(1).Text = Puntos
                
                'Updateamos el albaran
                If PtosAnt <> Puntos Then
                    cad = "UPDATE scaalb set puntos = " & DBSet(Puntos, "N")
                    cad = cad & ObtenerWhereCP(True)
                    conn.Execute cad
                    
                    'los moviimeintos PUNTOS
                    cad = Replace(ObtenerWhereCP(False), "scaalb", "smovalpuntos") & " and concepto=0 AND codclien"
                    
                    cad = DevuelveDesdeBD(conAri, "numero", "smovalpuntos", cad, CStr(Data1.Recordset!codClien))
                    If cad = "" Then
                        
                        cad = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", CStr(Data1.Recordset!codClien))
                        'NUEVA LINEA
                        
                        cad = " VALUES (" & Data1.Recordset!codClien & "," & Val(cad) + 1 & "," & DBSet(Data1.Recordset!codtipom, "T") & "," & Data1.Recordset!NumAlbar
                        
                        cad = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov)" & cad
                        cad = cad & " ," & DBSet(Data1.Recordset!FechaAlb, "F") & ",0," & DBSet(Puntos, "N") & ",now())"
                    Else
                        'UPDATE
                        cad = "UPDATE smovalpuntos set puntos=" & DBSet(Puntos, "N") & " WHERE codclien=" & Data1.Recordset!codClien & " AND numero=" & cad
                    End If
                    
                    ejecutar cad, False
                    
                    
                    Puntos = Puntos - PtosAnt
    
                    PtosCliente = PtosCliente + Puntos
                    cad = "UPDATE sclien set puntos = " & DBSet(PtosCliente, "N")
                    cad = cad & " WHERE codclien = " & Data1.Recordset!codClien
                    conn.Execute cad
                    
                End If
                
            End If
        End If
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
        EsNuevoAlbaran = False
        
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            If Port = 0 Then    'Si  ha metido/modifgicado portes no hago nada (port>0)
            
                'Enero 2010
                'para que no se vuelva a la primera linea
                'DeseleccionaGrid DataGrid1
                'DataGrid1.Bookmark = 1
            Else
                data2.Recordset.MoveLast  'El ultimo es el porte
            End If
        End If
        
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


Private Sub DataGrid1_DblClick()
    If Modo = 2 Then
        If Not data2.Recordset.EOF Then AbrirForm_Articulos DBLet(data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Modo = 2 And KeyCode = 113 Then
        If Not data2.Recordset.EOF Then AbrirForm_Articulos DBLet(data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If data2.Recordset Is Nothing Then Exit Sub
    If data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 1660 And X < 7950 Then
            If IsNull(Me.data2.Recordset!origpre) Then Exit Sub
            Select Case DataGrid1.Columns(10).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
            Me.DataGrid1.ToolTipText = Trim(DBLet(DataGrid1.Columns(5).Value, "T") & "    " & Me.DataGrid1.ToolTipText)
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
        
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo Error1

    If Not data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        '- ampliacion lineas
        SQL = "select ampliaci from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & data2.Recordset!numlinea
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text2(16).Text = DBLet(Rs.Fields(0).Value, "T")
        End If
        Rs.Close
        Set Rs = Nothing
        
        If vEmpresa.TieneAnalitica Then
            '- centro de coste
            ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = DBLet(data2.Recordset!CodCCost, "T")
            Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            '- nombre proveedor
            Text2(9).Text = DBLet(Me.data2.Recordset!nomprove, "T")
        End If
    Else
        Text2(16).Text = ""
        Text2(9).Text = ""
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    
    If PrimeraVez Then
        
        If AlbAvisoGenerado > 0 Then
            PonerCadenaBusqueda
            'Simulo que pulsa lineas
            mnLineas_Click
            
            'Simulo que le da a insertar nueva
            mnNuevo_Click
            
            'AlbAvisoGenerado
            AlbAvisoGenerado = 0
        End If
            
        'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
        If hcoCodMovim <> "" And Not Data1.Recordset.EOF And Modo <> 5 Then PonerCadenaBusqueda
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()


    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon


    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(1).Picture
    Next kCampo
    




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
        '.Buttons(7).Image = I + 5
    End With

    
    If vParamAplic.Ariagro <> "" Then
        With Me.ToolbarAux(2)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            
            .Buttons(1).Image = 3
            .Buttons(3).Image = 5
        End With
    
    End If
    
    If vParamAplic.CartaPortes Then
        With Me.ToolbarAux(1)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).visible = False
            .Buttons(3).Image = 5
        End With
        Label1(24).Caption = "Transportista"
        Label1(52).Caption = "Fecha carga"
        
    Else
        Label1(24).Caption = "Cod. envio"
        Label1(52).Caption = "Fecha envio"
    
    End If
    
    'Carta de portes
    FrameToolAux(1).visible = vParamAplic.CartaPortes
    DataGrid2.visible = vParamAplic.CartaPortes
    Text1(51).visible = vParamAplic.CartaPortes
    Text2(51).visible = vParamAplic.CartaPortes
    Label1(36).visible = vParamAplic.CartaPortes
    imgBuscar(16).visible = vParamAplic.CartaPortes
    Label1(66).visible = vParamAplic.CartaPortes
    Text1(52).visible = vParamAplic.CartaPortes
    Text2(52).visible = vParamAplic.CartaPortes
    imgBuscar(17).visible = vParamAplic.CartaPortes
    Label1(66).visible = vParamAplic.CartaPortes
    Label1(67).visible = vParamAplic.CartaPortes
    Text1(53).visible = vParamAplic.CartaPortes
    
    imgBuscar(18).visible = vParamAplic.CartaPortes
    Label1(60).visible = vParamAplic.CartaPortes
    Text1(54).visible = vParamAplic.CartaPortes
    Text2(54).visible = vParamAplic.CartaPortes
    
    If vParamAplic.CartaPortes Then
        SSTab1.TabCaption(1) = "Portes / Observaciones"
    Else
        SSTab1.TabCaption(1) = "Otros datos"
    End If
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        '11(30   21    20   16
        
        
        '                                                                                   Indice  antiguo
        .Buttons(1).Image = 30 'Nº Serie si lineas con articulos de control Nº serie  Ant 11
        .Buttons(2).Image = 21 'GEnerar factura ant 12
        .Buttons(3).Image = 20  'Marcar a facturar 13
        

        'MAYO 2015  Herbelca. ALbran ruta Castellon
        .Buttons(6).Image = 16 'Imprimir Pedido ant 16
        
        
        If vParamAplic.TipoPortes <> 1 Then
            If vParamAplic.PathFirmasAlbaran <> "" Then
                .Buttons(4).ToolTipText = "Imprimir albaran firmado"
                .Buttons(4).Style = tbrDefault
                .Buttons(4).Image = 54  '54
            Else
                .Buttons(4).Style = tbrSeparator
                .Buttons(4).visible = False
                .Buttons(4).ToolTipText = ""
            End If
        Else
            .Buttons(4).Style = tbrDefault
            .Buttons(4).ToolTipText = "Imprimir portes"
        End If

        
        
        
        
    End With


    Me.SSTab1.Tab = 0
    SSTab1.TabVisible(3) = False
      
    'Direcion envio SOLO si esta en parametros
    Label1(53).visible = vParamAplic.DireccionesEnvio
    imgBuscar(12).visible = vParamAplic.DireccionesEnvio
    Text1(42).visible = vParamAplic.DireccionesEnvio
    Text2(42).visible = vParamAplic.DireccionesEnvio
      
      
      
    LimpiarCampos   'Limpia los campos TextBox
    
    CargarComboFacturacion
    VieneDeBuscar = False
    CodTipoMov = hcoCodTipoM
    
    If CodTipoMov = "ALR" Then
        Me.Caption = "Albaranes Reparación"
        Label1(3).visible = False
        Label1(5).visible = False
        Text1(23).visible = False
        Text1(24).visible = False
        Label1(12).visible = False
        Text1(2).visible = False
        'Captions
        Label1(11).Caption = "Nº Repa."
        Label1(10).Caption = "Fecha repara."
        Text1(24).visible = False
        'Terminal
        Text1(38).visible = False
        Text1(39).visible = False
        Label1(65).visible = False
        
    Else
        Label1(11).Caption = "Nº Ped."
        Label1(10).Caption = "F. pedido"
    End If
   
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = DevuelveTextoDepto(True)
    
    If vParamAplic.TieneCRM Then
        Label1(55).Caption = "Observaciones CRM"
    Else
        Label1(55).Caption = "Observaciones internas"
    End If
    Label1(55).visible = Not (CodTipoMov = "ART")
    Text1(44).visible = Not (CodTipoMov = "ART")
    
    ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
    If vEmpresa.TieneAnalitica Then
        cmdAux(9).ToolTipText = "Buscar centro coste"
        txtAux(9).Tag = "centro coste"
        Label1(51).Caption = "Centro coste"
    Else
        Label1(51).Caption = "Proveedor"
    End If
    
    
    If vParamAplic.PtosAsignar > 0 Then
        Label1(63).visible = True
        Text2(1).visible = True
    End If
    
    
    Dim B As Boolean
    B = False
    If vParamAplic.NumeroInstalacion = 3 Or vParamAplic.NumeroInstalacion = 2 Then
        B = True
    Else
        If hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> "" Then B = True
    End If
    
    If B Then
    
    Else
        Text1(13).Width = 4125
        Text1(13).MaxLength = 255
    End If
        
    '## A mano
    If EsHistorico Then
        Me.FrameHco.visible = True
        kCampo = 1440
    Else
        Me.FrameHco.visible = False
        kCampo = 3440
    End If
   ' FrameFacRec.Left = kCampo
   ' FrameFactura.Left = kCampo
    kCampo = 0
        
    Me.FrameFacRec.visible = (CodTipoMov = "ART")
    
    
    
    
    MostrarComision = False
    If vParamAplic.NumeroInstalacion = 2 And vUsu.Nivel = 0 Then MostrarComision = True
        
    
    'Aportacion a terminal
    Label1(49).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    Text1(40).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    'Ajusta segun sea un tipo de albaran u otro
    PuntosMenusQuitadosPorDEV hcoCodTipoM = "DEV"
   
    'campos asociados
    Me.FrameCampos.visible = False
    If vParamAplic.Ariagro <> "" Then FrameCampos.visible = True
    'Frame manipulador
    
    Me.FrameManipulador.visible = False
    If vParamAplic.ManipuladorFitosanitarios2 Then
        FrameManipulador.visible = True
        FrameMani2.BorderStyle = 0
    Else
        If vParamAplic.Ariagro <> "" Then
            FrameCampos.Left = 120
            SSTab1.TabCaption(2) = "Campos"
        End If
    End If
    SSTab1.TabVisible(2) = vParamAplic.Ariagro <> "" Or vParamAplic.ManipuladorFitosanitarios2
    
    
    If AlbAvisoGenerado > 0 Then hcoCodMovim = AlbAvisoGenerado
        
     'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    'Lo que hacia antes, todo normal
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar=-1"
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
            Text1(0).BackColor = vbYellow
            End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
    End If
    
   
    
End Sub


Private Sub PuntosMenusQuitadosPorDEV(esDEV As Boolean)
Dim Color As Long
    
    
    
    Label1(50).Caption = IIf(esDEV, "Devolucion", "Nº Albaran")
    'If esDEV Then
    '    Color = Me.BackColor 'El que tengo en desarrollo
    'Else
        Color = &H8000000F
    'End If
    
   ' Me.chkFacturar.FontBold = esDEV
   ' Me.chkFacturar.ForeColor = IIf(esDEV, vbWhite, vbBlack)
    
   ' lblDevolucion.visible = esDEV
    Frame2.BackColor = Color
    Frame1(0).BackColor = Color
    chkFacturar.BackColor = Color
    
    
    LblMostr.visible = False
    If Not EsHistorico Then
        NombreTabla = "scaalb"
        NomTablaLineas = "slialb" 'Tabla lineas de Albaranes
        Ordenacion = " ORDER BY codtipom, numalbar "
        
        If hcoCodTipoM = "ALV" Then
            Me.Caption = "Albaranes Clientes"
        ElseIf hcoCodTipoM = "ALM" Then
            Me.Caption = "Albaranes de Mostrador"
            LblMostr.visible = True
           
        ElseIf hcoCodTipoM = "ART" Then
            Me.Caption = "Albaranes Rectificativos"
        ElseIf hcoCodTipoM = "ALI" Then
            Me.Caption = "Albaranes internos"
            
        ElseIf hcoCodTipoM = "ALT" Then
            Me.Caption = "Albaranes de telefonía"
            
        ElseIf hcoCodTipoM = "DEV" Then
            LblMostr.Caption = "DEVOLUCION"
            LblMostr.visible = True
            Me.Caption = "Devolución de mercancia CLIENTE"
        End If
    Else
        NombreTabla = "schalb"
        NomTablaLineas = "slhalb"
        CargarTagsHco Me, "scaalb", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(31).Tag = "Fecha Eliminación|F|N|||schalb|fechelim|dd/mm/yyyy|N|"
        Text1(32).Tag = "Trabajador Eliminación|N|N|0|9999|schalb|trabelim|0000|N|"
        Text1(33).Tag = "Incidencia elim.|T|N|||schalb|codincid||N|"
        Me.Caption = "Histórico Albaranes Clientes"
        Ordenacion = " ORDER BY codtipom, numalbar,fechaalb "
    End If
    Me.BackColor = Color
    SSTab1.BackColor = Color
    
    
    
    
    
    
    
    
End Sub

Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    Me.chkFacturar.Value = 0
    Me.chkFacturarKm.Value = 0
    Me.chkDocArchi.Value = 0
    Me.chkConTransporte.Value = 0
    Me.chkImpreso.Value = 0
    chkPideCliente.Value = 0
    
    If Me.FrameCampos.visible Then Me.ListView1.ListItems.Clear
    Text3(0).Text = "BASE IMPONIBLE"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo = 5 Then
        Cancel = 1
        Exit Sub
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    AlbAvisoGenerado = 0   'por si acaso
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
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
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If EsCabecera = 0 Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(30), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
                CadB = CadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 2), "0000000")
            
        Else
            If EsCabecera = 3 Then
                'Llama desde boton busqueda centros de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux(9).Text = RecuperaValor(CadenaDevuelta, 1)
                Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
                
            ElseIf EsCabecera = 1 Then
                'Llama desde Prismatico Direcciones/Departamentos
                Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                'DIRECCIONES escabecera2=2
                Text1(42).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(42).Text = RecuperaValor(CadenaDevuelta, 2)
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
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
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

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Envio
Dim Indice As Byte
    Indice = 29
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Envio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Envio
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmInt_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico
' o para recuperar una factura que vamos a Rectificar

    cadList = ""
    
    If frmList.OpcionListado = 225 Then  'Factura Rectificativa
        If CadenaSeleccion <> "" Then
            'codtipom
            cadList = " codtipom='" & RecuperaValor(CadenaSeleccion, 1) & "' and numfactu="
            'numfactu
            cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " and fecfactu="
            'fecfactu
            cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
            
            'campos observaciones
            motivo = DevuelveDesdeBD(conAri, "texto", "sparaidioma", "codigo", 2)   '2.- Motivo Rectifica
            If motivo = "" Then motivo = "MOTIVO"
            motivo = motivo & ": " & RecuperaValor(CadenaSeleccion, 4)
        End If
        
    Else 'Para recoger los Datos de Eliminacion que se introdujeron
        cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
        cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
        cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
    End If
End Sub


Private Sub frmMatr_DatoSeleccionado(CadenaSeleccion As String)
    BuscaChekc = CadenaSeleccion
End Sub

Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mensaje de Nº de Serie disponibles
'En cadena seleccion estan concatenados los seleccionados

    If frmMen.OpcionMensaje = 11 Then
        'En cadenaseleccion tenemos la WHERE que selecciona las lineas de la factura
        'que nos queremos traer para generar un albaran de rectificacion
        'Insertaremos estas lineas en la tabla slialb, y luego se podran eliminar,modificar,etc. (son de apoyo)
         InsertarLineasFactu (CadenaSeleccion)
    Else
        If Text1(30).Text = "ART" Then
            'Albaran de factura rectificativa
            If Not QuitarNumSeriesAlbVenta(CadenaSeleccion) Then MsgBox "Los nº de serie a rectificar no se han actualizado correctamente.", vbExclamation
        Else
            If Not AsignarNumSeriesAlbVenta(CadenaSeleccion) Then
                MsgBox "Los nº de serie del albaran no se han actualizado correctamente.", vbExclamation
            End If
        End If
    End If
End Sub


Private Sub frmNat_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmNSerie_CargarNumSeries()
Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim TieneMan As String * 1

    'Estamos en VENTAS e insertamos datos venta vacios
    If ModificaLineas = 4 Then
        CargarNumSeries
    Else
        'Viene de insertar Nº de series al insertar una linea

        'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
        'VAMOS A LEERLO dentro de insertarnumerioserie   22/12/2011
        'Ahora que vaya a 0
        TieneMan = "0": devuelve = ""    'bug. Estaba Tieneman="9"
        'devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
        ''El cliente tiene Mantenimientos
        'If devuelve <> "" Then TieneMan = "1"
        
        'cadena para INSERT
        'Estamos en VENTAS e insertamos datos de Cliente
        CadValues = ""
        CadValues = CadValues & Text1(4).Text & ", " & DBSet(Text1(12).Text, "T") & ", " & TieneMan & ", " & DBSet(devuelve, "T") & ", "
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " 'Fecha ult. Repar y Fin Garantia
        'Datos Venta
        CadValues = CadValues & DBSet(Text1(30).Text, "T") & ", " & ValorNulo & ", '" & Format(Text1(1).Text, FormatoFecha) & "', " & Text1(0).Text & ", " & Me.cmdAux(0).Tag & ", "
        'Rellenar los datos COMPRA del Proveedor a NULO
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo
        
        'cadena para UPDATE.  Faltara en la funciona añadir el nummante(si tiene)
        cadValuesU = " codclien=" & Text1(4).Text & ", coddirec=" & DBSet(Text1(12).Text, "T")
        cadValuesU = cadValuesU & ", codtipom=" & DBSet(Text1(30).Text, "T")
        cadValuesU = cadValuesU & ", fechavta='" & Format(Text1(1).Text, FormatoFecha) & "' "
        cadValuesU = cadValuesU & ", numalbar=" & Text1(0).Text & ", numline1=" & Me.cmdAux(0).Tag
        InsertarNSeries txtAux(1).Text, CadValues, cadValuesU, True
    End If
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = Val(Me.imgBuscar(3).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    txtAnterior = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte


    If Modo = 0 Then Exit Sub
    If Modo = 2 And Index <> 14 Then Exit Sub
    
    TerminaBloquear
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            Set frmC = New frmFacClientes3
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            Indice = 5
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 4
                txtAnterior = Text1(4).Text
            End If
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            Indice = 6
            
        Case 2 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = 1
                   'ANTES
                '01/DICIEMBRE/2010   DAVID
                'MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
                LanzaBusquedaDpto True, CInt(Indice)
                
             End If
             
        Case 3, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            If Index = 7 Then
                Indice = 27
            ElseIf Index = 8 Then
                Indice = 28
            Else
                Indice = Index
            End If
            Me.imgBuscar(3).Tag = Indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            Indice = 14
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'Agente
            Indice = 17
            PonerFoco Text1(Indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            
        Case 9 'Cod Envio
            Indice = 29
            PonerFoco Text1(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            
        Case 12
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = 2
                'MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 42
                LanzaBusquedaDpto False, CInt(Indice)
                
             End If
        Case 13
            Indice = 43
            txtAnterior = ""
            Set frmZ = New frmFacZonas
            frmZ.DatosADevolverBusqueda = "1|2|"
            frmZ.Show vbModal
            Set frmZ = Nothing
            If txtAnterior <> "" Then
                Text1(43).Text = RecuperaValor(txtAnterior, 1)
                Text2(43).Text = RecuperaValor(txtAnterior, 2)
                txtAnterior = Text1(43).Text
            End If
            
            
    Case 14
                CadenaDesdeOtroForm = Text2(16).Text
                frmFacClienteObser.Modificar = Modo = 5 And ModificaLineas > 0
                If hcoCodTipoM = "DEV" Then frmFacClienteObser.Modificar = False
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 5 And ModificaLineas > 0 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text2(16).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
                
                
    Case 15
        'Llamamos al manipulador de carnet fitosnaitarios
        CadenaDesdeOtroForm = ""
        frmFitoCarnet.Cliente = Val(Text1(4).Text)
        frmFitoCarnet.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
        
            If CDate(RecuperaValor(CadenaDesdeOtroForm, 2)) < CDate(Text1(1).Text) Then
                If MsgBox("Carnet caducado.  ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
            Me.Text1(45).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.Text1(46).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
            Me.Text1(47).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            Text2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            Me.Text1(48).Text = IIf(UCase(Text2(0).Text) = "CUALIFICADO", 2, 1)
        End If
        
        
    Case 16
        CadenaDesdeOtroForm = ""
        Set frmInt = New frmPortesIntermediario
        frmInt.DatosADevolverBusqueda = "0|1|"
        frmInt.Show vbModal
        Set frmInt = Nothing
        
        If CadenaDesdeOtroForm <> "" Then
            Me.Text1(51).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.Text2(51).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            
        End If
        
    Case 17
        
        CadenaDesdeOtroForm = ""
        Set frmNat = New frmPortesNaturaleza
        frmNat.DatosADevolverBusqueda = "0|1|"
        frmNat.Show vbModal
        Set frmNat = Nothing
        
        If CadenaDesdeOtroForm <> "" Then
            Me.Text1(52).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.Text2(52).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            
        End If
        
    Case 18
        AbreFormMatrChofer False   'CHOFER
        If BuscaChekc <> "" Then
            Text1(54).Text = RecuperaValor(BuscaChekc, 1)
            Text2(54).Text = RecuperaValor(BuscaChekc, 2)
        End If
        
        
    End Select
    
    

    If Index = 0 And hcoCodTipoM = "ALM" Then
        If HaDevueltoDatos Then
            If vParamAplic.EntradaRapidaFacturasMostrador Then Indice = 14
        End If
    End If
    
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


Private Sub Label2_Click()
    If Modo > 2 And Modo < 5 Then chkFacturar.Value = IIf(chkFacturar.Value = 0, 1, 0)
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
        BotonEliminarLinea
    Else
        'Eliminar Albaran
        BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
    'Imprimir Albaran
    BotonImprimir_ 45, False '45: Informe de Albaranes
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar albaran
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea False
    Else 'Añadir Cabecera
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


Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.Label1(35).visible = Me.SSTab1.Tab = 0
    Me.Text2(16).visible = Me.SSTab1.Tab = 0
    Me.Label1(51).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And SSTab1.Tab = 0
    Me.Text2(9).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And Me.SSTab1.Tab = 0
    Me.imgBuscar(14).visible = Me.SSTab1.Tab = 0
End Sub



Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True        'Cod. Postal
    
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    txtAnterior = Text1(Index).Text
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
   
    If Not (Index = 30 And Modo = 1) Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 38 Then KEYdown KeyCode
    
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
            Case 27, 28, 29
                Ind = Index - 20
            Case 42
                Ind = 12
            Case 43
                Ind = 13
            End Select
            If Ind >= 0 Then
                PulsadoMas2 = True
                PulsarTeclaMas True, Ind
            End If
        End If
    End If
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
Dim campo As String
        
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
    If txtAnterior = Text1(Index).Text Then
        
        
        'Entrad rapida
        If Index = 15 And vParamAplic.EntradaRapidaFacturasMostrador Then
            If Modo = 3 And hcoCodTipoM = "ALM" Then
                If ImporteFormateado(Text1(Index).Text) = 0 Then PonerFocoBtn cmdAceptar
            End If
        End If

        Exit Sub
    End If
          
    
          
          
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 41 'Fecha Albaran,fecenvio
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                If Text2(Index).Text = "" And Modo >= 3 Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
                If Index = 3 And Modo = 3 Then
                    Text1(28).Text = Text1(Index).Text
                    Text2(28).Text = Text2(Index).Text
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                   
                    Text1(5).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "N")
                Else 'If Modo = 3 Then 'Modo Insertar
                    'si es ART-Albaran de factura Rectificativa ya he cargado los
                    'datos de la factura
                     
                    If CodTipoMov <> "ART" Then
                        PonerDatosCliente (Text1(Index).Text)
                    Else
                        campo = "nomclien"
                        devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", campo)
                        If campo <> Text1(5).Text Then PonerDatosCliente Text1(Index).Text
                    End If
                    If Text1(Index).Text = "" Then
                        PonerFoco Text1(Index)
                    Else
                        If Text1(5).Locked Then
                            'Nos vamos a la forma de PAGO
                            If vParamAplic.EntradaRapidaFacturasMostrador Then
                                PonerFoco Text1(17)
                            Else
                                PonerFoco Text1(13)
                            End If
                        Else
                            PonerFoco Text1(5)
                        End If
                    End If
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
            If Text1(6).Locked Then Exit Sub
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
            If (Modo = 4) Then
                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
                    
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
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then ComprobarRefObligatoria
            
        Case 14 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 29 'Cod envio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
            Else
                Text2(Index).Text = ""
            End If
        Case 40
            PonerFormatoDecimal Text1(Index), 3
            
        Case 42, 43
            'Codigo envio y ZONA
            devuelve = ""
            If Text1(Index).Text <> "" Then
                
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "Campo numerico: " & Text1(Index), vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                Else
                    'Comprobar codenvio
                    If Index = 42 Then
                        devuelve = DevuelveDesdeBDNew(1, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", "", "coddiren", Text1(42).Text, "N")
                    Else
                        'zona
                        devuelve = PonerNombreDeCod(Text1(Index), conAri, "szonas", "nomzonas")
                    End If
                    If devuelve = "" Then
                        If Modo > 2 Then
                            MsgBox "No existe el codigo:" & Text1(Index).Text, vbInformation
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                    If Modo > 2 Then Me.Text2(Index).Text = devuelve
                End If
                
            Else
                PonerFoco Text1(Index)
            End If
            Text2(Index).Text = devuelve
    Case 49
        If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
        
    Case 50
        If Not PonerFormatoDecimal(Text1(Index), 6) Then Text1(Index).Text = ""
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If vUsu.CodigoAgente > 0 Then
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & " codagent = " & vUsu.CodigoAgente
    End If
    
    
    If chkVistaPrevia = 1 Then
        EsCabecera = 0
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        If Me.EsHistorico = False Then
            CadB = CadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
        End If
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera = 0 Then
        cad = cad & ParaGrid(Text1(30), 10, "Tipo Alb.")
        cad = cad & ParaGrid(Text1(0), 15, "Nº Albaran")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Alb.")
        cad = cad & ParaGrid(Text1(4), 10, "Cliente")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        Tabla = NombreTabla
        Titulo = "Albaranes"
        
        If EsHistorico Then
            Titulo = "Histórico de Albaranes"
            devuelve = "0|1|2|"
        Else
            Titulo = "Albaranes"
            devuelve = "0|1|"
        End If
    Else
        If EsCabecera = 1 Then
                'DIRECION DEPARTAMENTO
                If vParamAplic.HayDeparNuevo = 1 Then
                    Titulo = "Dptos Cliente: "
                    Desc = "Dpto."
                ElseIf vParamAplic.HayDeparNuevo = 0 Then
                    Titulo = "Direc. Cliente: "
                    Desc = "Direc."
                Else
                    Titulo = "Obra Cliente: "
                    Desc = "Obra"
                End If
                Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
                cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|18·"
                cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
                Tabla = "sdirec"
                devuelve = "0|1|"
                
        ElseIf EsCabecera = 2 Then
            'DIRENVIO
            '--------------------
            Titulo = "Dirección de envio cliente: "
            Desc = " envio"
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            cad = cad & "Codigo" & Desc & "|sdirenvio|coddiren|N|000|18·"
            cad = cad & "Descripción" & Desc & "|sdirenvio|nomdiren|T||65·"
            Tabla = "sdirenvio"
            devuelve = "0|1|"
        
        Else
            Stop
        
        End If
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        If EsCabecera > 0 Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing

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


    Screen.MousePointer = vbHourglass
    On Error GoTo EPonerLineas

    'Datos de la tabla slipre
    CargaGrid DataGrid1, data2, True

    If vParamAplic.CartaPortes Then CargaGrid DataGrid2, data3, True

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim B As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
     'Si es un Albaran de Ticket visualizamos unos datos y sino otros
    B = (Data1.Recordset!EsTicket = 1)
    Me.Toolbar1.Buttons(11).Enabled = (Not B) And (Not EsHistorico)
    

    If hcoCodTipoM <> "ALR" Then
        'sem. entrega pedido
        Label1(12).visible = Not B
        Text1(2).visible = Not B
        'num oferta
        Text1(23).visible = Not B And hcoCodTipoM <> "ALR"
        'fecha oferta
        Text1(24).visible = Not B
        'nº terminal
        Text1(38).visible = B
        'nº venta
        Text1(39).visible = B
        Label1(65).visible = B
    
        If B Then
        'El albaran se genero a partir de un ticket
            Me.Label1(11).Caption = "Nº Ticket"
            Me.Label1(10).Caption = "Fecha Ticket"
            Me.Label1(9).Caption = "Trabajador Ticket"
        
            'ocultamos los datos de la oferta
            Me.Label1(3).Caption = "Nº Venta"
            Label1(5).Caption = "Nº Terminal"
        Else
            Me.Label1(11).Caption = "Nº Pedido"
            Me.Label1(10).Caption = "Fecha Pedido"
            Me.Label1(9).Caption = "Trabajador Pedido"
    
            'Mostramos los datos de la oferta
            Me.Label1(3).Caption = "Nº Oferta"
            Label1(5).Caption = "Fecha Oferta"
        End If
        
    End If
    PonerCamposForma Me, Data1
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
    Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "straba", "nomtraba", "codtraba")
    Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    Text2(43).Text = PonerNombreDeCod(Text1(43), conAri, "szonas", "nomzonas")
     
    'Direccion de envio
    If vParamAplic.DireccionesEnvio Then Text2(42).Text = PonerNombreDeCod(Text1(42), conAri, "sdirenvio", "nomdiren", "codclien = " & Text1(4).Text & " AND coddiren")
    
    
    
    
    
    Text2(16).Text = ""
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "straba", "nomtraba", "codtraba")
        Text2(33).Text = PonerNombreDeCod(Text1(33), conAri, "sincid", "nomincid", "codincid")
    End If
    
    
    If vParamAplic.PtosAsignar > 0 Then ObtenerPuntos
    
   
    If Me.FrameCampos.visible Then CargaDatosCampos
    If vParamAplic.ManipuladorFitosanitarios2 Then
        If Val(DBLet(Data1.Recordset!TipoCarnet, "N")) > 0 Then
            Text2(0).Text = IIf(Val(Data1.Recordset!TipoCarnet) = "2", "Cualificado", "Básico")
        Else
            Text2(0).Text = ""
        End If
    End If
    
    
    If vParamAplic.CartaPortes Then
        Text2(51).Text = PonerNombreDeCod(Text1(51), conAri, "sintermediario", "nominter", "codinter")
        Text2(52).Text = PonerNombreDeCod(Text1(52), conAri, "snaturalezas", "descnatura", "codnatura")
        Text2(54).Text = PonerNombreDeCod(Text1(54), conAri, "sconductor", "concat(nombre,' (',dni,')')", "chofer")
    End If

    
    
    
    CalcularDatosFactura
    
    BotonesToolBarAux
    
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

    lblF.Caption = ""
    BuscaChekc = ""
    

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
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
        PonLblIndicador lblIndicador, Data1
    End If
    
    DespalzamientoVisible NumReg > 1
        
    
    If B Then 'modo=2
        If Me.FrameCampos.visible Then
            'Tiene campos visibles
            If Not Data1.Recordset.EOF Then B = True
        Else
            B = False
        End If
    End If
    If vParamAplic.Ariagro <> "" Then
        ToolbarAux(2).Buttons(1).Enabled = B
        ToolbarAux(2).Buttons(3).Enabled = B
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(30), B
    'Bloquear los campos de Oferta
    If Text1(23).visible Then
        BloquearTxt Text1(23), B
        BloquearTxt Text1(24), B
    End If
    'Bloquear los campos de Pedido
    For i = 25 To 27
        BloquearTxt Text1(i), B
    Next i
    BloquearTxt Text1(2), B
    
    'Si lleva lotes
    If vParamAplic.ManipuladorFitosanitarios2 Then
         If Modo = 3 Or Modo = 4 Then
            For i = 45 To 48
                BloquearTxt Text1(i), True
            Next i
        End If
    End If
        'ZAFIR/101
    
    
    
    'bloquea los datos de venta del TPV (si hay)
    If Text1(38).visible Then
        BloquearTxt Text1(38), B
        BloquearTxt Text1(39), B
    End If
    
    'Bloquea los campos de Factura (si visibles, ed, si es Rectificativa)
    For i = 35 To 37
        BloquearTxt Text1(i), B
    Next i
  
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
    Me.chkFacturar.Enabled = B
    Me.chkFacturarKm.Enabled = B
    Me.chkDocArchi.Enabled = B
    Me.chkConTransporte.Enabled = B
    Me.chkImpreso.Enabled = B
    
    chkPideCliente.Enabled = B 'Modo = 1 Or (B And vUsu.Nivel < 1)
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    BloquearTxt Text2(9), (Modo <> 5)
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    Me.imgFecha(0).Enabled = B
    Me.imgFecha(40).Enabled = B
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(1).visible = False
    Me.imgBuscar(7).Enabled = (Modo = 1)
    Me.imgBuscar(14).Enabled = (Modo <> 1)
              
              
    'Modo Linea de Albaranes
    '- poner visible ampliacion linea
    BloquearTxt Text2(16), True
    '- poner visible nombre proveedor linea
    BloquearTxt Text2(9), True
    SSTab1_Click 0
      
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
       
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    
    
    'Para remarcar el cliente
    '&H00C0FFFF&
    If Modo = 2 Then Text1(5).BackColor = &HC0FFFF
    
    
    
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
Dim B As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    
            
    
    'Asignarle el valor del Combo Tipo de Movimiento al texto oculto text1(30)
'    Text1(30).Text = ObtenerCodTipom
    
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    
    
    
    
    
    'En devoluciones NO dejo modificar el cliente SI ya tiene lineas
    If Modo = 4 Then
        If hcoCodTipoM = "DEV" Then
            If Not data2.Recordset.EOF Then
                If Val(Data1.Recordset!codClien) <> Val(Text1(4).Text) Then
                
                    MsgBox "No puede cambiar el cliente en la devolucion", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
     If Trim(Text1(4).Text) <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            B = False
        End If
    End If
    If Not B Then Exit Function
    
    
    '2014 Dicimebre
    'HERBELCA. No dejo cambiar el agente, ya que si no las comisiones no corresponderian
    If vParamAplic.NumeroInstalacion = 2 Then
        If Modo = 4 Then
        
            'Si tiene lineas de articulos, no puede cambiar por tema comision
            If Not data2.Recordset.EOF Then
                If Val(Text1(17).Text) <> Val(Data1.Recordset!CodAgent) Then
                    If vUsu.Nivel < 1 Then
                        'AVisamos, pero dejamos continuar
                        MsgBox "Las comisones podrían ser incorrectas.", vbExclamation
                        
                    Else
                        MsgBox "No puede cambiar el agente", vbExclamation
                        Text1(17).Text = Data1.Recordset!CodAgent
                        Text1_LostFocus 17
                    End If
                End If
            End If
        End If
    End If
     
    
    
    If vParamAplic.CartaPortes Then
        If Text1(52).Text = "" Or Text2(52).Text = "" Then
            MsgBox "Debe seleccionar naturaleza de la mercancia", vbExclamation
            B = False
        End If
    End If
    
    If Modo = 4 Then
    
    
          If vParamAplic.ManipuladorFitosanitarios2 Then
                'No dejo cambiar el cliente SI lleva fitosnaitarios
                devuelve = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                devuelve = " select distinct codartic from slialb " & devuelve
                devuelve = " numserie<>'' and codartic in (" & devuelve & ") AND 1"
                devuelve = DevuelveDesdeBD(conAri, "count(*)", "sartic", devuelve, "1")
                If Val(devuelve) > 0 Then
                
                    If Me.Text1(45).Text <> "" Then
                        'OK
                        
                    Else
                        MsgBox "El albarán tiene productos fitosanitarios. Debe indicar el carnet", vbExclamation
                        Exit Function
                    End If
                End If
            End If

    
    
    
    
        
            'Modificando
            GrabaCambioTrabajador = -1
            
                
             If Val(Text1(3).Text) <> Val(Data1.Recordset!CodTraba) Then
                 'MsgBox "No puede cambiar el trabajador", vbExclamation
                 'Text2(3).Text = ""
                 'Text1(3).Text = CStr(Data1.Recordset!CodTraba)
                 'PonerFoco Text1(3)
                 GrabaCambioTrabajador = Data1.Recordset!CodTraba
             End If
         
         
         
           'En herbelca SOLO los supersuser quitan la marca de facturar
        
            If hcoCodTipoM <> "DEV" And vParamAplic.NumeroInstalacion = 2 And vUsu.Nivel > 0 Then
                
                    If DBLet(Data1.Recordset!factursn, "N") = 1 And Me.chkFacturar.Value = 0 Then
                    
                        MsgBox "No puede quitar la marca de facturar", vbExclamation
                        Exit Function
                    End If
    
            End If
            
            If vParamAplic.PtosAsignar > 0 Then
            
         
            
                devuelve = ""
                If Val(Data1.Recordset!codClien) <> Val(Text1(4).Text) Then
                    If CDate(Me.Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                        If CDate(Me.Text1(1).Text) > vParamAplic.PtosFechaIncio Then
                            If Data1.Recordset!FechaAlb < vParamAplic.PtosFechaIncio Then devuelve = "N"
                        Else
                            If Data1.Recordset!FechaAlb > vParamAplic.PtosFechaIncio Then devuelve = "N"
                        End If
                    End If
                End If
                
                If devuelve <> "" Then
                    MsgBox "No puede cambiar el cliente y ademas cambiar las fecha del albaran a peridod de puntos distinto", vbExclamation
                    Exit Function
                End If
            
            
                If CDate(Me.Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                    devuelve = Replace(ObtenerWhereCP(False), "scaalb", "slialb") & " AND codartic "
                    devuelve = DevuelveDesdeBD(conAri, "codartic", "slialb", devuelve, vParamAplic.PtosArticuloCanje, "T")
                    If devuelve <> "" Then
                        MsgBox "Tiene  articulo canje. ", vbExclamation
                        Exit Function
                    End If
                End If
            
    
            
            
            End If
    End If  'modificando
    
    
    
    
    'Lleva direcciones de envio. Comprobamos que la que ha puesto existe...
    If vParamAplic.DireccionesEnvio Then
        If Text1(42).Text = "" Xor Text2(42).Text = "" Then
            MsgBox "Dirección de envio INCORRECTA", vbExclamation
            B = False
            PonerFoco Text1(42)
        End If
        'Ha puesto un codenvio y parece ser que existe... LO COMPURBEO que no hay referenciales
        If B And Text1(42).Text <> "" Then
            BuscaChekc = DevuelveDesdeBDNew(1, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", "", "coddiren", Text1(42).Text, "N")
            If BuscaChekc = "" Then
                MsgBox "NO existe la dirección de envio: " & Text1(42).Text, vbExclamation
                PonerFoco Text1(42)
                B = False
            End If
            BuscaChekc = ""
        End If
     End If 'de direnvii

    
    'Estamos en facturas mostrador
    'El cliente esta bloqueado (le hemos dejado pasar, pese a dar el mensaje)
    'La forma de pago solo puede ser EFECTIVO o TARJETA
    If Not ClienteBloqueadoYFormaPagoCorrecta Then B = False

    If B Then
        If Me.hcoCodTipoM = "ALM" Then
            If vParamAplic.FrasMostradorSerieDistinta Then
                'Tiene contadores distintos.... FORMA DE PAGO deberia ser efec o tartje
                BuscaChekc = DevuelveDesdeBDNew(1, " sforpa", "tipforpa", "codforpa", Text1(14).Text)
                If BuscaChekc <> "0" And BuscaChekc <> "6" Then
                    If MsgBox("La forma pago deberia ser efectivo o tarjeta.   ¿Continuar? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
                    If Not B Then PonerFoco Text1(14)
                End If
                BuscaChekc = ""
            End If
        End If
    End If
    
    
    
    If B And Modo = 3 Then
        'En devoluciones NO compruebo
        If hcoCodTipoM <> "ALR" Then
            'Si esta bien y estamos insertando
            If vParamAplic.OperacionesAseguradas Then
                'Si tiene Operaciones ASEGURADAS
                If Not Riesgo Then B = False
            End If
        End If
    End If
        
        
    'Abril 2015.
    'El NIF no puede ser el de la empresa para albaranes normales, y solo puede ser el de la empresa para
    ' albranes interno
    If B Then
        devuelve = ""
        If hcoCodTipoM = "ALI" Then
            If Text1(6).Text <> vParam.CifEmpresa Then devuelve = "Facturas internas sólo pueden ser a NIF empresa(" & vParam.CifEmpresa & ")"
        ElseIf Text1(6).Text = vParam.CifEmpresa Then
            devuelve = "No puede facturarse a si mismo. NIF debe ser distinto empresa(" & vParam.CifEmpresa & ")"
        End If
        If devuelve <> "" Then
            MsgBox devuelve, vbExclamation
            B = False
        End If
    End If
    
    'Albaranes de TELEFONIA.  TIENE QUE existir el telefono, y este debe estar en
    'el campo referencia
    If B Then
        If hcoCodTipoM = "ALT" Then
            'Albaranes de telefonia introducidos a mano, la marca del cliente debe de estar,
            'Cuando se genereren autmaticamente (facturacion desde fichero) pondre un 0
            Me.chkPideCliente.Value = 1
        
            'Albaranes de TELEFONIA.  TIENE QUE existir el telefono, y este debe estar en
            devuelve = ""
            If Text1(13).Text = "" Then
                   devuelve = "Debe poner el teléfono asociado"
            Else
                
                devuelve = "concat(codclien,'|',if(coddirec is null,'',coddirec),'|')"
                devuelve = DevuelveDesdeBD(conAri, devuelve, "sclientfno", "idtelefono", Text1(13).Text, "T")
                If devuelve = "||" Then devuelve = ""
                If devuelve = "" Then
                    devuelve = "No existe el telefono en la BD"
                Else
                    BuscaChekc = RecuperaValor(devuelve, 2)
                    devuelve = RecuperaValor(devuelve, 1)
                    If Val(devuelve) <> Val(Text1(4).Text) Then
                        devuelve = "No esta asociado al cliente"
                    Else
                        'OK existe y es de el cliente. Veamos si lleva coddirec
                                                
                        If BuscaChekc = Me.Text1(12).Text Then
                            devuelve = "" 'ok
                        Else
                            If BuscaChekc = "" Then
                                devuelve = "No debe asignarse a esta direccion"
                            Else
                                If Val(BuscaChekc) = Val(Me.Text1(12).Text) Then
                                    devuelve = ""
                                Else
                                    devuelve = "Asociado a otra direccion del cliente"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            If devuelve <> "" Then
                MsgBox devuelve, vbExclamation
                PonerFoco Text1(13)
                B = False
            End If
            
        End If
    End If
    
    
    
    'HERBELCA.
    ' Modificar.  Los trabajadores de GANDIA-CASTELLON no pueden desmarcar FACTURAR
    If vParamAplic.NumeroInstalacion = 2 And Modo = 4 And Me.chkFacturar.Value = 0 Then
        If DBLet(Me.Data1.Recordset!factursn, "N") = 1 Then
            'ERA facturar y ahora NO tienen la marca.
            If vUsu.AlmacenPorDefecto2 > 1 Then Me.chkFacturar.Value = 1   'NO PREGUNTAMOS ni damos error ni nada de nada
        End If
    End If
    
    
    
    
    
    
    
'    If Modo = 3 And b Then
'         If vParamAplic.ManipuladorFitosanitarios2 And hcoCodTipoM = "ALM" Then
'                'Esto sera para el CHOLI , en Navarrres
'                devuelve = DevuelveDesdeBD(conAri, "ManipuladorNumCarnet", "sclien", "codclien", Text1(4).Text)
'                If devuelve = "" Then
'                    'Veo si tiene autirzados
'                    devuelve = DevuelveDesdeBD(conAri, "numcarnet", "sclienmani", "codclien", Text1(4).Text)
'                End If
'
'                If devuelve <> "" Then
'                    'Llamamos al manipulador de carnet fitosnaitarios
'                    CadenaDesdeOtroForm = ""
'                    frmFitoCarnet.Cliente = Val(Text1(4).Text)
'                    frmFitoCarnet.Show vbModal
'                    If CadenaDesdeOtroForm <> "" Then
'                        Me.Text1(45).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
'                        Me.Text1(46).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
'                        Me.Text1(47).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
'                        Text2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
'                        'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
'                        Me.Text1(48).Text = IIf(UCase(Text2(0).Text) = "CUALIFICADO", 2, 1)
'                    End If
'                End If
'        End If
'    End If
    
    
    DatosOk = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ClienteBloqueadoYFormaPagoCorrecta() As Boolean
    ClienteBloqueadoYFormaPagoCorrecta = True
    If Me.hcoCodTipoM = "ALM" Then
        If EsClienteBloqueado(Text1(4).Text, False, True) Then
            'LA forma de pago solo pude ser efectivo o tarjeta   (0 o 6)
            BuscaChekc = DevuelveDesdeBDNew(1, " sforpa", "tipforpa", "codforpa", Text1(14).Text)
            If BuscaChekc <> "0" And BuscaChekc <> "6" Then
                MsgBox "Cliente bloqueado.  Forma pago INVALIDA(solo efectivo o tarjeta) ", vbExclamation
                PonerFoco Text1(14)
                ClienteBloqueadoYFormaPagoCorrecta = False
            End If
            BuscaChekc = ""
        End If
    End If
    
End Function

Private Function DatosOkLinea(ByRef vCStock As CStock, ByRef ARticuloFitosantiario As Boolean) As Boolean
Dim B As Boolean
Dim i As Byte
Dim Aux As String
Dim AUx3 As String
Dim vArtic As CArticulo
Dim PuntosCliente As Currency
Dim C2 As Currency
Dim Comision As String
Dim CanDispo As Currency
Dim vPrecioFact As CPreciosFact
Dim PrMinimo As Currency

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    ARticuloFitosantiario = False
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(6), txtAux(7), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(8).Text Then txtAux(8).Text = Aux
    
    
        
    B = True
    For i = 0 To 10
        If txtAux(i).Text = "" And i <> 5 Then
            'El campo 5= origpre puede ser nulo (en alb.repar)
            
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            B = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
    
    

    If hcoCodTipoM = "DEV" Then
        If ImporteFormateado(txtAux(3).Text) > 0 Then
            MsgBox "Cantidad debe ser negativa", vbExclamation
            PonerFoco txtAux(3)
            Exit Function
        End If
    End If


    
    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        B = vCStock.MoverStock(False, False)
    End If
    
    
    
    
    If B Then


        'Octubre 2011
        ' Comprobar que este articulo, para este cliente, no esta en otro pedido
        Set vArtic = New CArticulo
        vArtic.LeerDatos txtAux(1).Text
        If vArtic.EsDeVarios = 0 Then
            Aux = "scaped.numpedcl=sliped.numpedcl  AND codclien = " & Text1(4).Text & "  AND sliped.codartic"
            Aux = DevuelveDesdeBD(conAri, "concat(sliped.numpedcl,"" de fecha "",DATE_FORMAT(fecpedcl,'%d %M %Y'))", "scaped,sliped", Aux, txtAux(1).Text, "T")
            If Aux <> "" Then
                
                Aux = "Artículo: " & vArtic.codigo & "   " & vArtic.Nombre & vbCrLf & vbCrLf & "Esta en el pedido: " & Aux
                Aux = "Cliente: " & Text1(4).Text & "   " & Text1(5).Text & vbCrLf & vbCrLf & Aux
                Aux = Aux & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then B = False
            End If
        End If
        'Set vArtic = Nothing  lo pongo a nothing bajo

    End If
    
    
    
    
    
    
    'Articulo canje.
    If vParamAplic.PtosAsignar > 0 Then
        If ModificaLineas = 2 Then
            'Esta modificando el articulo de canje
            If data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
                'El articulo ahora debe ser el de canje. NO lo puede modificar
                If txtAux(1).Text <> vParamAplic.PtosArticuloCanje Then
                    MsgBox "No puede reemplzar el articulo de canje en esta linea", vbExclamation
                    B = False
                End If
            Else
                If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                    'No podemos sustituir un articulo por el articulo de canje
                    MsgBox "No puede reemplzar un articulo por el de canje ", vbExclamation
                    B = False
                End If
            End If
        Else
            'Dando de alta
            If Me.txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
            
             
                'El cliente tiene puntos
                AUx3 = "puntos"
                Aux = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text, "N", AUx3)
                If Val(Aux) = 0 Then
                    MsgBox "El cliente no tiene marca de Puntos", vbExclamation
                    B = False
                Else
                    If AUx3 = "" Then AUx3 = "0"
                    PuntosCliente = CCur(AUx3)
                    'Si no es nuevo albaran, solo superusuarios pueden insertar canje
                    If Not EsNuevoAlbaran Then
                        If vUsu.Nivel > 0 Then
                            MsgBox "No es albaran nuevo. No puede utilizar articulo canje", vbExclamation
                            B = False
                        Else
                            If MsgBox("No es un albaran nuevo. ¿Desea continuar con el canje?", vbQuestion + vbYesNo) <> vbYes Then B = False
                        End If
                    End If
                    
                    If Not B Then Exit Function
                    
                     Aux = Replace(ObtenerWhereCP(False), "scaalb", "slialb") & " AND codartic "
                     Aux = DevuelveDesdeBD(conAri, "codartic", "slialb", Aux, txtAux(1).Text, "T")
                     If Aux <> "" Then
                         MsgBox "Ya esta el articulo de canje en este albarán", vbExclamation
                         B = False
                     
                     Else
                         'De momento Veo si hay algun articulo de familias de canje
                         Aux = CalcularPuntosAlbaranCABEL(Replace(ObtenerWhereCP(False), "scaalb", "slialb"), Data1.Recordset!FechaAlb, AUx3, Comision)
                         
                         If Aux = "" Then Aux = "0"
                         If CCur(Aux) = 0 Then
                            'No hacemos nada
                            MsgBox "No tiene articulos de las familias de canje", vbInformation
                            
                            If Not data2.Recordset.EOF Then B = False
                         Else
                           
                            C2 = Round2(CCur(AUx3) / vParamAplic.PtosEquivalencia, 2) '-> necesito como mucho estos puntos
                            If C2 > PuntosCliente Then
                                Aux = PuntosCliente
                            Else
                                Aux = C2
                            End If
                           
                           
                         
                         
                            If -ImporteFormateado(txtAux(3).Text) > CCur(Aux) Then
                                MsgBox "No puede canjear mas de " & Aux, vbExclamation
                                B = False
                            End If
                            
                            
                            txtAux(12).Text = Comision
                            
                         End If
                     End If
                End If
            End If
        End If
    
    
        If B And vParamAplic.NumeroInstalacion = 2 Then
            If B And vArtic.EsDeVarios = 0 And txtAux(5).Text <> "P" Then     'en herbelca. Precio minimo
                '------------------------------------------
                
                'If Not vArtic.EstablecidoPrecioMinimo Then vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
                vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
                
                If vArtic.EstablecidoPrecioMinimo Then
                    C2 = ImporteFormateado(txtAux(3).Text)
                    If C2 <> 0 Then
                        C2 = Round2(ImporteFormateado(txtAux(8).Text) / C2, 4)
                        If C2 < vArtic.PrecioMinimo Then
                            C2 = C2 - vArtic.PrecioMinimo
                            If Abs(C2) > 0.01 Then
                                B = False
                                Aux = "Precio inferior al mínimo permitido" & vbCrLf
                                If vUsu.Nivel = 0 Then
                                    Aux = Aux & vbCrLf & vbCrLf & "¿Continuar?"
                                    If MsgBox(Aux, vbQuestion + vbYesNoCancel) = vbYes Then B = True
                                Else
                                    MsgBox Aux, vbExclamation
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    
    
        'Es correcto, y vemos la cantidad
        If B Then
            If Me.txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                'Los puntos son positivos de normal . El ya los restar (o sumara en los de devolucion
                Aux = ""
                If ImporteFormateado(txtAux(3).Text) < 0 Then
                    'Cantidad negativa. En albaranes rectificativos
                    If Data1.Recordset!codtipom = "ALR" Then Aux = "Canje puntos en positivo"
                Else
                    If Data1.Recordset!codtipom <> "ALR" Then Aux = "Canje puntos en negativo"
                End If
                If Aux <> "" Then
                    MsgBox Aux, vbExclamation
                    B = False
                End If
            End If
        End If
    
    End If
    
    
    If Not B Then Exit Function
    
    If vParamAplic.NumeroInstalacion = 2 And hcoCodTipoM = "ALM" Then
        AUx3 = ""
        If Val(txtAux(0).Text) = 1 Then
            AUx3 = DevuelveDesdeBD(conAri, "ctrstock", "sartic", "codartic", txtAux(1).Text, "T")
        End If
        If AUx3 = "1" Then
            AUx3 = "sum(cantidad)"
            '                                           cualqueir almacen
            Aux = "scaped.numpedcl=sliped.numpedcl  AND codalmac>=" & txtAux(0).Text & " AND codartic "
            Aux = DevuelveDesdeBD(conAri, "max(fecpedcl)", "scaped,sliped", Aux, txtAux(1).Text, "T", AUx3)
            If Aux <> "" Then
                C2 = CCur(AUx3)
                
               
                AUx3 = "codalmac>=" & txtAux(0).Text & " AND codartic "
                AUx3 = DevuelveDesdeBD(conAri, "sum(canstock)", "salmac", AUx3, txtAux(1).Text, "T")
                CanDispo = CCur(AUx3) 'en stock para el almacen hay candispo unidades
                
                If CanDispo < vCStock.cantidad + C2 Then
                  'Las cantiodades que ha pedido mas las hay en pedidos superan
                  AUx3 = "Hay pedidos con este articulo pendientes de servir" & vbCrLf & vbCrLf
                  AUx3 = AUx3 & "Stock articulo:     " & CanDispo & vbCrLf
                  AUx3 = AUx3 & "Unidades en pedidos: " & C2
                  
                  AUx3 = AUx3 & vbCrLf & vbCrLf & "¿Continuar?"
                  If MsgBox(AUx3, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
                End If

            End If
         End If
    End If
    
    
    
    If B Then
        If hcoCodTipoM <> "FMO" Then
            'En facturas de mostrador NO lo compurbeo para acelerar el prloceso
            vCStock.ComprobarFechaInventario True, ""
              
        End If
        
        
           
        GrabaLogCambioPrecioDto = False
        
        If vParamAplic.LogCambioPrecDto Then ComprobarCambioPrecioDtoyVtaBajoPrecio
   
        
        
    End If
        
    
    
    ' Articulos de varios en negativo NO pueden
    If B Then
        If vParamAplic.NumeroInstalacion = 2 Then
            'HERBELCA
            If vUsu.Nivel > 0 Then
                If ImporteFormateado(Me.txtAux(3).Text) < 0 Then
                    Aux = "artvario=1 AND sartic.codartic"
                    Aux = DevuelveDesdeBD(conAri, "count(*)", "sartic", Aux, txtAux(1).Text, "T")
                    If Val(Aux) > 0 Then
                        MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                        B = False
                    End If
                
                
                    If B And (vUsu.AlmacenPorDefecto2 = 3 Or vUsu.AlmacenPorDefecto2 = 2) Then
                        'Los usuarios de CASTELLON NO pueden realizar abonos sobre materia no rotacion
                        Aux = "artvario=0 AND sartic.codartic"
                        Aux = DevuelveDesdeBD(conAri, "rotacion", "sartic", Aux, txtAux(1).Text, "T")
                        If Val(Aux) = 0 Then
                            MsgBox "Material de NO rotación. No se permite el abono", vbExclamation
                            B = False
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    If B Then
        'Fitosanitarios para los que llevan control
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If ModificaLineas = 1 Then
                 B = DatosLotesFitosOk(ARticuloFitosantiario)
            Else
                'select * from sartic,scateg where
                Aux = " sartic.codcateg=scateg.codcateg and ctrlotes =1 and codartic"
                Aux = DevuelveDesdeBD(conAri, "numserie", "sartic,scateg", Aux, txtAux(1).Text, "T")
                If Aux <> "" Then ARticuloFitosantiario = True
            
                
            
            
                'No me puede cambiar ni la cantidad ni el articuo
                If ARticuloFitosantiario Then
                    Aux = ""
                    If data2.Recordset!codArtic <> txtAux(1).Text Then Aux = Aux & "-Codigo de articulo" & vbCrLf
                    If data2.Recordset!cantidad <> ImporteFormateado(txtAux(3).Text) Then Aux = Aux & "-Cantidad" & vbCrLf
                    If Aux <> "" Then
                        Aux = "Error lotes fitosanitarios. No puede cambiar: " & vbCrLf & vbCrLf & Aux
                        MsgBox Aux, vbExclamation
                        B = False
                    End If
                End If
            End If
        End If
    End If
    
    DatosOkLinea = B
    Set vArtic = Nothing
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Set vArtic = Nothing
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolbar Button.Index
End Sub

Private Sub HacerToolbar(Indice As Integer)
    Select Case Indice
        
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click 'Modificar
        Case 3: mnEliminar_Click  'Borrar
            
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos
        
            
        Case 8:
                mnImprimir_Click 'Imprimir Albaran
        
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

  
'DesdeRecuperaParaRectificativa:  Para que no inserte el punto verde
Private Function InsertarLinea(numlinea As String, DesdeRecuperaParaRectificativa2 As Boolean) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim SQL As String
Dim vWhere As String
Dim B As Boolean
Dim vCStock As CStock
Dim ImpReciclado As Single
Dim DentroTRANS As Boolean
Dim ArtFitosnatiarios As Boolean

Dim SqlIntercalar As String
Dim SqlIntercalar2 As String


    InsertarLinea = False
    SQL = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    
    
    If LineaIntercalar = 0 Then
        'INSERCION NORMAL
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
     
        SqlIntercalar = ""
        SqlIntercalar2 = ""
    Else
        
                                                'por si acaso lleva tasa reciclaje
        SQL = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 2 WHERE " & vWhere & " and numlinea >= " & LineaIntercalar
        SQL = SQL & " order by numlinea desc " 'Para que empieza por las ultimas
        SqlIntercalar = SQL
        
        'ENERO 2018. ERROR GRAVE
        ' No actualizaba la smoval
        SQL = "UPDATE smoval SET numlinea=numlinea + 2"
        SQL = SQL & "  WHERE detamovi ='" & Text1(30).Text & "' AND document ='" & Text1(0).Text & "' AND "
        SQL = SQL & " fechamov= " & DBSet(Text1(1).Text, "F") & " and numlinea >= " & LineaIntercalar
        SQL = SQL & " order by numlinea desc"
        SqlIntercalar2 = SQL
        
        numlinea = LineaIntercalar
    End If
    
    
    
    Me.cmdAux(0).Tag = numlinea 'Aqui almaceno el Nº linea que acabo de Insertar
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", numlinea) Then Exit Function
    
    '24 Febrero 2015
    If DesdeRecuperaParaRectificativa2 Then
        
        B = True
    Else
        B = DatosOkLinea(vCStock, ArtFitosnatiarios)
    End If
    
    If B Then 'Lineas de Albaranes
    
        'Inserta en tabla "slialb"
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos,precioar, dtoline1, dtoline2, importel, origpre,codprovex,numlote,codccost,pvpInferior,comisionagente) "
        SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        '- cantidad,numbultos
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", " & DBSet(txtAux(10).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(5).Text, "T", "N") & ","
        
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If ArtFitosnatiarios Then
                'Pongo un "*" al articulo para indicarme que lleva lotes
                txtAux(11).Text = "*"
            End If
        End If
        
        
        If vEmpresa.TieneAnalitica Then
            '- codprove,numlote,codccost
            SQL = SQL & "0," & DBSet(txtAux(11).Text, "T", "S") & "," & DBSet(txtAux(9).Text, "T", "S")
        Else
            '- codprove,numlote,codccost
            SQL = SQL & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(11).Text, "T", "S") & "," & ValorNulo
        End If
        SQL = SQL & "," & VendeAMenorPrecio & ","
        If vParamAplic.NumeroInstalacion = 2 Then
            SQL = SQL & DBSet(txtAux(12).Text, "N", "S")
        Else
            SQL = SQL & "null"
        End If
        SQL = SQL & ")"
        
     Else
        Exit Function
     End If
    
    If SQL <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        
        'Enero 2018
        'Las lineas de intercalar
        If SqlIntercalar <> "" Then
            conn.Execute SqlIntercalar
            conn.Execute SqlIntercalar2
            Espera 0.1
        End If
        
        'insertar la linea
        conn.Execute SQL
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If hcoCodTipoM <> "DEV" Then
            B = vCStock.ActualizarStock(False, True)
        
            'Si ha cambiado, si tiene el parametro... todo esta ahi
            TrataCambioPrecioDto
        End If
        
        'Sera TRUE, si (y solo si)tiene lo de manipuladore de fitosanitarios y el articulo tiene seire ...
        If ArtFitosnatiarios Then
            
            '#Este codigo esta copiado en ModificarLote
            SQL = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
            SQL = SQL & " SELECT '" & Data1.Recordset!codtipom & "'," & Data1.Recordset!NumAlbar & "," & numlinea
            SQL = SQL & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
            SQL = SQL & " FROM tmpnlotes  WHERE codusu = " & vUsu.codigo & " and cantidad <>0 "

            conn.Execute SQL
            
            'Tengo que updatear la cantidad vendida
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open "Select * FROM tmpnlotes  WHERE codusu = " & vUsu.codigo & " and cantidad >0 ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux!cantidad <> 0 Then
                    If miRsAux!cantidad > 0 Then
                        SQL = "+"
                    Else
                        SQL = "-"
                    End If
                    SQL = "UPDATE slotes SET vendida=vendida " & SQL & DBSet(Abs(miRsAux!cantidad), "N")
                    SQL = SQL & " WHERE numlotes =" & DBSet(miRsAux!numlotes, "T") & " AND codartic= " & DBSet(miRsAux!codArtic, "T")
                    SQL = SQL & " AND fecentra= " & DBSet(miRsAux!FechaAlb, "F")
                
                    conn.Execute SQL
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
        'Si ha actualizado el sctock
        If B Then
            If ClienteConTasaReciclado And Not DesdeRecuperaParaRectificativa2 Then
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                    'Insertamos la linea del reciclado
                    
                    vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,  precioar,"
                    SQL = SQL & "dtoline1, dtoline2, importel, origpre,comisionagente) "
                    SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
                    SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
                    'Importe linea
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A',"
                    'Comision
                    ImpReciclado = 0
                    SQL = SQL & DBSet(ImpReciclado, "N") & ")"
                    conn.Execute SQL
                        
                    
                End If 'articulo con sunida reciclado
            End If  'Cliente con tasa reciclado
        End If 'ok actualiza stock
        
        If B Then
            If vParamAplic.PtosAsignar > 0 Then
                If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                    
                    ImpReciclado = ImporteFormateado(txtAux(3).Text)
    
                    SQL = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", CStr(Data1.Recordset!codClien))
                    SQL = " VALUES (" & Data1.Recordset!codClien & "," & Val(SQL) + 1 & "," & DBSet(Data1.Recordset!codtipom, "T") & "," & Data1.Recordset!NumAlbar
                    
                    SQL = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov)" & SQL
                    SQL = SQL & " ," & DBSet(Data1.Recordset!FechaAlb, "F") & ",1," & DBSet(ImpReciclado, "N") & ",now())"
                    conn.Execute SQL
                
                    
                    SQL = "UPDATE sclien set puntos=" & DBSet(ImpReciclado, "N") & " + coalesce(puntos,0) "
                    SQL = SQL & " WHERE codclien =" & Text1(4).Text
                    conn.Execute SQL
                End If
            End If
        End If
    
    End If
    Set vCStock = Nothing
    
    
    
    If B Then
        conn.CommitTrans
        InsertarLinea = True
        AlmacenLineas = CInt(txtAux(0).Text)
        
        
    Else
        conn.RollbackTrans
         InsertarLinea = False
    End If
    Set miRsAux = Nothing
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & Err.Description
    End If
    
End Function




Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vCStock As CStock
Dim B As Boolean
Dim ImpReciclado As Single
Dim ArtFitosnatiarios As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    '#### LAURA 15/11/2006
    conn.BeginTrans
        
    If DatosOkLinea(vCStock, ArtFitosnatiarios) Then
        
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        B = InicializarCStock(vCStock, "E")
        If B Then
            If hcoCodTipoM <> "DEV" Then
                B = vCStock.DevolverStock2 'eliminamos de smoval y devolvemos stock valores anteriores
                If B Then
                    'si se ha modificado el articulo
                    If CStr(data2.Recordset!codArtic) <> txtAux(1).Text Then
                        'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
                        SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
                        SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
                        SQL = SQL & " WHERE codartic=" & DBSet(data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                        SQL = SQL & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & data2.Recordset!numlinea
                        conn.Execute SQL
                    End If
                End If
                'ahora leemos los valores nuevos
                If B Then B = InicializarCStock(vCStock, "S")
                'insertamos en smoval y actualizamos stock a los valores nuevos
                vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
                If B Then B = vCStock.ActualizarStock(False, True)
                
            Else
                B = True
            End If
    
            'actualizar la linea de Albaran
            If B Then
                SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
                SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", numbultos=" & DBSet(txtAux(10).Text, "N") & ","
                SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", " 'precio
                SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
                SQL = SQL & "importel= " & DBSet(txtAux(8).Text, "N") & ", " 'Importe
                SQL = SQL & "origpre=" & DBSet(txtAux(5).Text, "T", "S") & ","
                ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
                If vEmpresa.TieneAnalitica Then
                    SQL = SQL & "codccost=" & DBSet(txtAux(9).Text, "T", "S") & ","
                Else
                    SQL = SQL & "codprovex=" & DBSet(txtAux(9).Text, "N", "N") & ","
                End If
                SQL = SQL & "numlote=" & DBSet(txtAux(11).Text, "T", "S") & ", "
                
                'Junio2013
                SQL = SQL & "pvpInferior=" & DBSet(VendeAMenorPrecio, "N") & ""
                If vParamAplic.NumeroInstalacion = 2 Then SQL = SQL & " , comisionagente =" & DBSet(txtAux(12).Text, "N", "S")
                
                
                SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & data2.Recordset!numlinea
                conn.Execute SQL
                
                                
                'Si ha cambiado, si tiene el parametro... todo esta ahi
                If hcoCodTipoM <> "DEV" Then TrataCambioPrecioDto
                
                'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       NumRegElim = Val(DBLet(data2.Recordset!numlinea, "N")) + 1
                       SQL = DevuelveDesdeBD(conAri, "codartic", "slialb", SQL, CStr(NumRegElim))
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
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
                
                If vParamAplic.PtosAsignar > 0 Then
                    If data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
                        'Ha cambiado el articulo de puntos
                        'Actualizamos los puntos en el cliente, el el movimputs
                        ImpReciclado = data2.Recordset!cantidad
                        ImpReciclado = ImporteFormateado(txtAux(3).Text) - ImpReciclado
                        SQL = "UPDATE sclien set puntos=" & DBSet(ImpReciclado, "N") & " + coalesce(puntos,0) "
                        SQL = SQL & " WHERE codclien =" & Text1(4).Text
                        conn.Execute SQL
                                                
                        SQL = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                        SQL = SQL & " AND codclien = " & Data1.Recordset!codClien & " AND concepto=1"
                        SQL = "UPDATE smovalpuntos  set puntos = " & DBSet(ImporteFormateado(txtAux(3).Text), "N") & SQL
                        conn.Execute SQL
                                                
                                                
                                                
                    End If
                End If
            End If
'        If SQL <> "" Then
'
'            vCStock.Cantidad = CSng(txtAux(3).Text)
'            b = vCStock.ModificarStock(Data2.Recordset!Cantidad)
'        End If
        End If
    End If
    Set vCStock = Nothing
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        ModificarLinea = True
        
    Else
        conn.RollbackTrans
         ModificarLinea = False
    End If
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
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim SQL As String
    
    On Error GoTo ECargaGrid

    B = vDataGrid.Enabled
    
    SQL = MontaSQLCarga(enlaza, IIf(vDataGrid.Name = "DataGrid2", 2, 1))
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez, 330
    
    CargaGrid2 vDataGrid, vData
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B
    vDataGrid.ScrollBars = dbgAutomatic
    PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
    
    On Error GoTo ECargaGrid

    vData.Refresh
   
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    

    Select Case vDataGrid.Name
    Case "DataGrid1" 'Cod. Almacen
            vDataGrid.Columns(2).visible = False
            i = 3
            vDataGrid.Columns(i).Caption = "Alm."
            vDataGrid.Columns(i).Width = 610
            vDataGrid.Columns(i).NumberFormat = "000"
            
            i = i + 1 '4
            vDataGrid.Columns(i).Caption = "Articulo"
            vDataGrid.Columns(i).Width = 2000
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Descripción Artículo"
            vDataGrid.Columns(i).Width = 4800

            i = 6
            vDataGrid.Columns(i).visible = False
            i = 7
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 1050
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = 8
            vDataGrid.Columns(i).Caption = "Bultos"
            vDataGrid.Columns(i).Width = 750
            vDataGrid.Columns(i).Alignment = dbgRight
                
            i = i + 1 '9
            vDataGrid.Columns(i).Caption = "Precio"
            vDataGrid.Columns(i).Width = 1150
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoPrecio
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "OP"
            vDataGrid.Columns(i).Width = 450
            vDataGrid.Columns(i).Alignment = dbgCenter
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.1"
            vDataGrid.Columns(i).Width = 750
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.2"
            vDataGrid.Columns(i).Width = 750
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Importe"
            vDataGrid.Columns(i).Width = 1300
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            If vEmpresa.TieneAnalitica Then
                i = i + 1
                vDataGrid.Columns(i).Caption = "CCost"
                vDataGrid.Columns(i).Width = 800
                vDataGrid.Columns(i).Alignment = dbgRight
            Else
                i = i + 1
                
                'If vParamAplic.NumeroInstalacion = 2 Then
                If MostrarComision Then
                    vDataGrid.Columns(i).Caption = "Comi."
                    vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                Else
                    vDataGrid.Columns(i).Caption = "Prov"
                End If
                vDataGrid.Columns(i).Width = 800
                vDataGrid.Columns(i).Alignment = dbgRight
            
                '- nombre proveedor
                i = i + 1
                vDataGrid.Columns(i).visible = False
    '            vDataGrid.Columns(i).Caption = "Nom. prove"
    '            vDataGrid.Columns(i).Width = 2100
            End If
            
            '- numlote
            i = i + 1
            vDataGrid.Columns(i).Caption = "Nº Lote"
            vDataGrid.Columns(i).Width = 1400

            'Solo HERBELCA. Acaba con el codprove
            If vParamAplic.NumeroInstalacion = 2 Then
                i = i + 1
                vDataGrid.Columns(i).visible = False
                i = i + 1
                vDataGrid.Columns(i).visible = False   'comision
                
                If Not MostrarComision Then
                    i = i + 1
                    vDataGrid.Columns(i).visible = False   'comision
                End If
            End If
            
    Case "DataGrid2"
            
            i = 2
            vDataGrid.Columns(i).Caption = "matricula"
            vDataGrid.Columns(i).Width = 1500
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Desc. "
            vDataGrid.Columns(i).Width = 1800

    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
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
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(9).visible = visible
        FrameCliente.Refresh
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
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                ElseIf i = 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 4).Text
                ElseIf i >= 4 And i < 9 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                    
                ElseIf i = 9 Then
                    'Proveedor/centro de coste
                    ' o en herbelca es el descuento.- El proveedor esta al final
                    If vParamAplic.NumeroInstalacion = 2 Then
                        txtAux(i).Text = DataGrid1.Columns(17).Text
                    Else
                        'Como estaba
                        txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                    End If
                    
                ElseIf i = 10 Then
                    txtAux(i).Text = DataGrid1.Columns(8).Text
                ElseIf i = 11 Then
                    ' ---- [19/10/2009] [LAURA] : centro de coste si hay conta analitica
                    If vEmpresa.TieneAnalitica Then
                        txtAux(i).Text = DataGrid1.Columns(i + 4).Text
                    Else
                        txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                    End If
                ElseIf i = 12 Then
                    'Comision solo herbelca
                    If vParamAplic.NumeroInstalacion = 2 Then
                        If CStr(data2.Recordset!tipoprecio) <> "*" Then
                            txtAux(i).Text = ""
                        Else
                            If Not MostrarComision Then
                                txtAux(i).Text = DataGrid1.Columns(19).Text 'esta en el campo 19
                            Else
                                txtAux(i).Text = DataGrid1.Columns(14).Text
                            End If
                        End If
                        
                    Else
                        'Resto
                        txtAux(i).Text = ""
                    End If
                End If
                txtAux(i).Locked = False
                
  
                
                
            Next i
        End If
        
        cmdAux(0).Enabled = True
        cmdAux(1).Enabled = True
'        cmdAux(9).Enabled = True
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
        
        
        If vParamAplic.NumeroInstalacion = 2 Then
            Me.cmdAux(9).visible = False
            txtAux(12).visible = True
        Else
            'Como estaba
            BloquearTxt txtAux(9), (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica <> 2)
            Me.cmdAux(9).Enabled = Not txtAux(9).Locked
            Me.cmdAux(9).visible = Me.cmdAux(9).Enabled
            txtAux(12).visible = False
        End If
            
        
        
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 30)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(9).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(9).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 50
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 10
        txtAux(1).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width + 20
        txtAux(2).Width = DataGrid1.Columns(5).Width - 20
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 20
        txtAux(3).Width = DataGrid1.Columns(7).Width - 20
        'Bultos
        txtAux(10).Left = txtAux(3).Left + txtAux(3).Width + 20
        txtAux(10).Width = DataGrid1.Columns(8).Width - 20

        txtAux(4).Left = txtAux(10).Left + txtAux(10).Width + 20
        txtAux(4).Width = DataGrid1.Columns(9).Width - 20
        
        'OP, Dto1, Dto2, Precio, (codProve/codccost)
        For i = 5 To 9
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 26
            txtAux(i).Width = DataGrid1.Columns(i + 5).Width - 26
        Next i
        
        'El boton 3 lo superpongo un poquito
'        cmdAux(9).Left = txtAux(10).Left - 15
        If cmdAux(9).visible Then
            txtAux(9).Width = txtAux(9).Width - cmdAux(9).Width
            cmdAux(9).Left = txtAux(9).Left + txtAux(9).Width
            txtAux(11).Left = cmdAux(9).Left + cmdAux(9).Width + 26 'numlote
        Else
            txtAux(11).Left = txtAux(9).Left + txtAux(9).Width + 26 'numlote
        End If
        
        '- numlote
'        txtAux(11).Left = cmdAux(9).Left + cmdAux(9).Width + 20
        If vEmpresa.TieneAnalitica Then
            txtAux(11).Width = DataGrid1.Columns(15).Width - 26
        Else
            txtAux(11).Width = DataGrid1.Columns(16).Width - 26
        End If
        
        
        'Solo herbelca
        If vParamAplic.NumeroInstalacion = 2 Then
            txtAux(12).Left = DataGrid1.Columns(14).Left + 240
            txtAux(12).Width = DataGrid1.Columns(14).Width
            BloquearTxt txtAux(12), True
            If Not MostrarComision Then txtAux(12).Top = 20000
        End If
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 2
            txtAux(i).visible = visible
        Next i
        'El 12 solo es visible si es visible y herbelca
        txtAux(12).visible = visible And vParamAplic.NumeroInstalacion = 2
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
 
        If vParamAplic.NumeroInstalacion = 2 Then txtAux(9).Left = 25000
        

    End If
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1
            
            If Not vParamAplic.ManipuladorFitosanitarios2 Then BotonNSeries 'Control Nº Series
            
    Case 2
    
            If data2.Recordset Is Nothing Then Exit Sub
            If data2.Recordset.RecordCount = 0 Then
                MsgBox "No tiene lineas de albarán", vbExclamation
                Exit Sub
            End If
        
            If Data1.Recordset!codtipom = "ALT" Then
                MsgBox "No puede facturar directamente un albarán de telefonía", vbExclamation
                Exit Sub
            End If
                
            'procedimiento normal
            If Data1.Recordset!codtipom = "ART" Then
                'Comprobar nº serie de las facturas rectificativas
                DevolverNumSeries
            End If
                
                
                
            If hcoCodTipoM = "DEV" Then
            
                'Proceso que pasaremos la devoucion a ALV o ART . Venta o rectificativa
                'En el caso de rectificativa llamaremos a al trozo de abajo
                
                
                
                ''Cargamos el ado otra vez
                If GeneraAlbaranDesdeDevolucion Then
                        'Sea bueno o malo
                    
                    If hcoCodTipoM = "ART" Then
                        If Not Data1.Recordset.EOF Then HacerToolbar 12
                    Else
                        MsgBox "Albaran venta generado", vbInformation
                    End If
                    
                End If
                Exit Sub
            End If
            
            If Me.chkFacturar.Value = 1 Then
            
                If Not ClienteBloqueadoYFormaPagoCorrecta Then Exit Sub
                
                If vParamAplic.ManipuladorFitosanitarios2 Then
                    If Not VerCarnetManipulador Then Exit Sub
                End If
                
                NumRegElim = Data1.Recordset.AbsolutePosition
                
                
                If vParamAplic.NumeroInstalacion = 2 Then
                    If Not PrecioMinimoAlbaran Then Exit Sub
                End If
                
                'Facturacion de Albaran de Mostrador
                frmListadoPed.codClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
                frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
                AbrirListadoPed (222)
                
                PosicionarDataTrasEliminar
                
                
                'Si es rectificativa. salir
                If hcoCodTipoM = "ART" Then
                    Unload Me
                    Exit Sub
                End If
            Else
                MsgBox "El Albaran no esta marcado para facturar", vbInformation
            End If


    
    Case 3
        'Marca los albaranes que esten como NO facturar a facturar
        If Modo = 5 Then
            If ModificaLineas = 0 Then
                If vParamAplic.ManipuladorFitosanitarios2 Then ModificaLote
            End If
        Else
            MarcarAlbaranes
        End If
        
        
    Case 4
        If vParamAplic.TipoPortes <> 1 Then
            If vParamAplic.PathFirmasAlbaran <> "" Then
                Screen.MousePointer = vbHourglass
                ImprimirAlbaranFirmado
                Screen.MousePointer = vbDefault
            End If
        Else
            BotonImprimir_ 45, True
        End If
        
    
    
            
    Case 6:
        'Mayo 2015.  Impresion albaranes ruta CASTELLON   HERBELCA
        ImpresionAlbaranRutaCliente

    End Select
    
End Sub

Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    If Modo <> 2 And Modo < 5 Then Exit Sub

    If Modo >= 5 And ModificaLineas > 0 Then Exit Sub
    
    Select Case Index
    Case 0
    
        BotonMtoLineas 0, "Albaranes"
        
        
        If Button.Index = 3 Then
            BotonEliminarLinea
        Else
           ' PonerModo 5
            If Button.Index = 1 Or Button.Index = 5 Then
                'AÑADIR linea factura
                BotonAnyadirLinea Button.Index <> 1
                
            ElseIf Button.Index = 2 Then
                'MODIFICAR linea factura
                BotonModificarLinea
                
            ElseIf Button.Index = 6 Then
                If ModificaLineas = 0 Then
                    If vParamAplic.ManipuladorFitosanitarios2 Then
                        ModificaLote
                        cmdRegresar_Click
                    End If
                End If
            End If
            
            
            
            
            
            
            
            
        End If
        
        
    Case 1
    
        If Button.Index = 3 Then
            If data3.Recordset.EOF Then Exit Sub
            
            If MsgBox("Eliminar la matricula del transporte: " & data3.Recordset!Matricula, vbQuestion + vbYesNoCancel) = vbYes Then
                BuscaChekc = "Delete from scaalb_portes " & Replace(ObtenerWhereCP(True), NombreTabla, "scaalb_portes")
                BuscaChekc = BuscaChekc & " and matricula=" & DBSet(data3.Recordset!Matricula, "T")
                                
                                
                'Hay que eliminar
                NumRegElim = data3.Recordset.AbsolutePosition
                If ejecutar(BuscaChekc, False) Then
                    ModificaLineas = 0
                    CargaGrid2 DataGrid2, data3
                    SituarDataTrasEliminar data3, NumRegElim
                End If
                
                BuscaChekc = ""
                                
                                
            End If
        Else
            PonerModo 6
            PonerBotonCabecera True
            If Button.Index = 1 Then
                BotonAnyadirLineaMatricula
            Else
            
            End If
        End If
        
    Case 2
            'Campos ariagro
            BotonesCampos Button.Index = 1
        
        
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento Button.Index
End Sub

Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    LabelAyudatxtAux Index, lblF
    
    'Pierde el foco el importe. Si es herbelca, pasamos al txt
    If Index = 9 Then
         If vParamAplic.NumeroInstalacion = 2 Then PonerFoco Text2(16)
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'campo almacen y flecha arriba
    
    If Index < 2 Or Index = 9 Then  'Para los que tienen busqueda
    
    
    
        
            'Insertando linea albaran
            
            If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                
                If Modo = 5 And ModificaLineas = 1 Then
                    If txtAux(Index).Text = "" Then
                        PulsadoMas2 = True
                        KeyCode = 0
                
                        PulsarTeclaMas False, Index
                    End If
                End If
            Else
                'Ha pulsado F2
                If KeyCode = 113 Then Me.DataGrid1.Columns(4).Caption = "EAN"
            End If
    
        
    ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
    ElseIf KeyCode = 113 Then AccionesF2 Index
    ' ----
    End If
    KEYdown KeyCode
End Sub

Private Sub AccionesF2(Index As Integer)
    If Index = 3 Then
        AbrirForm_Articulos txtAux(1).Text
    Else
        If Index = 4 Then
            AbrirConsultaPrecio Text1(4).Text, txtAux(1).Text, Text1(1).Text
        Else
            If Index = 6 Or Index = 7 Then AbrirFormularioDtos txtAux(1).Text
        End If
    End If
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim cantidad As String
Dim vCStock As CStock
Dim B As Boolean
Dim okArticulo As Boolean
Dim DtoPermitido As Boolean
Dim AbrirDevoluciones As Boolean
Dim StatusArticMayorCero As Boolean
Dim TieneDescuentos As String
Dim AUx3 As String
Dim PtosAuxiliar As Currency



    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        If txtAux(Index).Text <> "" Then txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. ARTICULO
            If txtAux(Index).Text = "" Then
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
                If Not data2.Recordset.EOF Then devuelve = data2.Recordset!codArtic
            End If
            cantidad = txtAux(9).Text
            
            If Me.DataGrid1.Columns(4).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEan(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , cantidad, StatusArticMayorCero)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , cantidad, StatusArticMayorCero)
            End If
            If Not okArticulo Then
                If Me.DataGrid1.Columns(4).Caption = "EAN" Then txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            Else
                If devuelve <> txtAux(1).Text Then
                    'ha cambiado el articulo
                    Me.txtAux(3).Text = ""
                    Me.txtAux(4).Text = ""
                    Me.txtAux(5).Text = ""
                    Me.txtAux(6).Text = ""
                    Me.txtAux(7).Text = ""
                    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
                        'NO TOCAMOS txtaux
                    Else
                        Me.txtAux(9).Text = ""
                    End If
                End If
                B = (Me.ActiveControl.Name = "txtAux")
                If B Then B = (Me.ActiveControl.Index = 0)
                If Not B Then
                    If txtAux(2).Locked Then
                        If StatusArticMayorCero Then PonerFoco txtAux(3)
                    End If
                        
                Else
                    PonerFoco txtAux(0)
                End If
                
                
                '---- [20/10/2009] [LAURA] : añadir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    'Si  ha cambiado el articulo, el proveedor
                    If txtAux(9).Text = "" Then
                        txtAux(9).Text = cantidad
                        'Fuerzo el lostfocus para que carge el proveedor
                        txtAux_LostFocus 9
                    End If
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'Por familia
                    txtAux(9).Text = cantidad
                    Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
                End If
                '----
            
                'Se ira a la pantalla de devolucion
                If CodTipoMov = "DEV" Then
                
                    devuelve = "Select slifac.* from scafac,slifac     where "
                    devuelve = devuelve & " scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu "
                    devuelve = devuelve & " AND codclien = " & Text1(4).Text & " and scafac.fecfactu>='2011-01-01'"
                    devuelve = devuelve & " AND codartic<>" & DBSet(vParamAplic.ArtReciclado, "T")
                    devuelve = devuelve & " AND codtipoa like 'A%' "    'para quitar los que no sean albaranes
                    devuelve = devuelve & " AND codartic = " & DBSet(txtAux(1).Text, "T")
                    devuelve = devuelve & " ORDER BY scafac.fecfactu desc ,scafac.codtipom,scafac.numfactu,numlinea "    'para quitar los que no sean albaranes
                    CadenaDesdeOtroForm = ""
                    AbrirDevoluciones = False
                    Set miRsAux = New ADODB.Recordset
                    miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    If miRsAux.EOF Then
                        
                        'Modificacion Herbelca
                        'Comprobamos si esta en ALBARNES
                        If vParamAplic.NumeroInstalacion = 2 Then
                            miRsAux.Close
                            
                            devuelve = " Select slialb.*,FechaAlb from slialb,scaalb     where  scaalb.codtipom=slialb.codtipom and"
                            devuelve = devuelve & " scaalb.NumAlbar = slialb.NumAlbar  AND codclien = " & Text1(4).Text
                            devuelve = devuelve & " AND codartic<>" & DBSet(vParamAplic.ArtReciclado, "T")
                            devuelve = devuelve & " AND scaalb.codtipom <>'ALZ' "    'para quitar los que no sean albaranes
                            devuelve = devuelve & " AND codartic = " & DBSet(txtAux(1).Text, "T")
                            devuelve = devuelve & " ORDER BY fechaalb,numlinea"
                            miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                            If Not miRsAux.EOF Then AbrirDevoluciones = True
                            
                        
                        End If
                    Else
                        AbrirDevoluciones = True
                    End If
                    
                    
                    
                    If AbrirDevoluciones Then
                        frmListado5.OpcionListado = 16
                        frmListado5.OtrosDatos = txtAux(1).Text & "|" & txtAux(2).Text & "|"
                        frmListado5.Show vbModal
                    Else
                        MsgBox "El cliente no compró este articulo", vbExclamation
                    End If
                        
                    If CadenaDesdeOtroForm = "" Then
                        'PONGO TODOS LOS VALORES a ""
                        txtAux(1).Text = ""
                        txtAux(2).Text = ""
                        Me.txtAux(3).Text = ""
                        Me.txtAux(4).Text = ""
                        Me.txtAux(5).Text = ""
                        Me.txtAux(6).Text = ""
                        Me.txtAux(7).Text = ""
                        Me.txtAux(9).Text = ""
                        Text2(16).Text = ""
                        DoEvents
                        PonerFoco txtAux(1)
                    Else
                        'Traemos los valores de la linea devuelta
                        Me.txtAux(2).Text = miRsAux!NomArtic  'por si acaso es de varios
                        Me.txtAux(3).Text = Format(-miRsAux!cantidad, FormatoCantidad)
                        Me.txtAux(4).Text = Format(miRsAux!precioar, FormatoPrecio)
                        
                        Me.txtAux(6).Text = Format(miRsAux!dtoline1, FormatoDescuento)
                        Me.txtAux(7).Text = Format(miRsAux!dtoline2, FormatoDescuento)
                        Me.txtAux(8).Text = Format(-miRsAux!ImporteL, FormatoDescuento)
                        
                        
                        If vEmpresa.TieneAnalitica Then
                            txtAux(9).Text = DBLet(miRsAux!CodCCost, "T")
                        Else
                            txtAux(9).Text = DBLet(miRsAux!codProvex, "N")
                        End If
                        Me.txtAux(5).Text = miRsAux!origpre
                        txtAux(11).Text = DBLet(miRsAux!numLote, "T")
                        txtAux(10).Text = Abs(miRsAux!NumBultos)
                        
                        TieneDescuentos = "concat(dtognral,'|',dtoppago,'|')"
                        If Mid(miRsAux!codtipom, 1, 1) = "F" Then
                            'Es una factura
                            'Vere los descuentos
                            devuelve = "fecfactu=" & DBSet(miRsAux!FecFactu, "F") & " AND codtipom =" & DBSet(miRsAux!codtipom, "T") & " AND numfactu"
                            TieneDescuentos = DevuelveDesdeBD(conAri, TieneDescuentos, "scafac", devuelve, CStr(miRsAux!Numfactu))
                     
                            'Par el resto
                            devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", miRsAux!codtipom, "T")
                            devuelve = "Rectifica a factura: " & devuelve & Format(miRsAux!Numfactu, "000000") & " de " & Format(miRsAux!FecFactu, "dd/mm/yyyy")
                        Else
                            'Es un albaran. Solo HERBELCA demomento
                            devuelve = "codtipom = " & DBSet(miRsAux!codtipom, "T") & " AND numalbar"
                            devuelve = DevuelveDesdeBD(conAri, "fechaalb", "scaalb", devuelve, miRsAux!NumAlbar, "N", TieneDescuentos)
                            devuelve = "Rectifica al albarán: " & miRsAux!codtipom & Format(miRsAux!NumAlbar, "000000") & " de " & Format(devuelve, "dd/mm/yyyy")
                        End If
                        Text2(16).Text = devuelve
                        If TieneDescuentos = "" Then TieneDescuentos = "0|0|"
                            
                        If CCur(RecuperaValor(TieneDescuentos, 1)) <> Data1.Recordset!DtoGnral Then
                           TieneDescuentos = "En albaran origen: " & RecuperaValor(TieneDescuentos, 1) & vbCrLf
                           TieneDescuentos = TieneDescuentos & "En albaran ACTUAL: " & Data1.Recordset!DtoGnral
                           TieneDescuentos = "Descuento general" & vbCrLf & TieneDescuentos
                        Else
                            If CCur(RecuperaValor(TieneDescuentos, 2)) <> Data1.Recordset!DtoPPago Then
                                TieneDescuentos = "En albaran/factura origen: " & RecuperaValor(TieneDescuentos, 2) & vbCrLf
                                TieneDescuentos = TieneDescuentos & "En albaran devolucion: " & Data1.Recordset!DtoPPago & vbCrLf
                                TieneDescuentos = "Descuento pronto pago" & vbCrLf & vbCrLf & TieneDescuentos
                            Else
                                TieneDescuentos = ""
                            End If
                        End If
                        If TieneDescuentos <> "" Then MsgBox TieneDescuentos, vbExclamation
                        
                    End If

                End If
                
                If vParamAplic.PtosAsignar > 0 And hcoCodTipoM <> "ALR" Then
                    If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                    
                        devuelve = "puntos"
                        If DevuelveDesdeBD(conAri, "tienepuntos", "sclien", "codclien", CStr(Data1.Recordset!codClien), "T", devuelve) = "0" Then
                              MsgBox "NO tiene puntos", vbExclamation
                              txtAux(1).Text = ""
                        Else
                            If devuelve = "" Then devuelve = "0"
                              
                            'Diciembre. Lo quitamos. Puede canjear todos los puntos que tenga el cliente en la ficha
                            If CCur(devuelve) > 0 Then
                                cantidad = CalcularPuntosAlbaranCABEL(Replace(Replace(ObtenerWhereCP(False), "scaalb", "slialb"), NombreTabla, NomTablaLineas), Data1.Recordset!FechaAlb, AUx3, "")
                                If cantidad = "" Then cantidad = "0"
                                
                                'Siginifica que tiene articvulos CABEL
                                If CCur(cantidad) > 0 Then
                                    'Tiene cable. Vamos a ver cuantos puntos necesita como maximo para el importe de este albaran
                                    PtosAuxiliar = Round2(CCur(AUx3) / vParamAplic.PtosEquivalencia, 2) '-> necesito como mucho estos puntos
                                    If PtosAuxiliar > devuelve Then
                                        txtAux(3).Text = Format(-1 * CCur(devuelve), FormatoCantidad)
                                    Else
                                        txtAux(3).Text = Format(-1 * PtosAuxiliar, FormatoCantidad)
                                    End If
                                    
                                    
                    
                                Else
                                      MsgBox "No existen articulos CABEL", vbExclamation
                                      txtAux(1).Text = ""
                                End If
                                
                              
                            Else
                                txtAux(1).Text = ""
                            End If
                            
                            txtAux(6).Text = "0"
                            txtAux(7).Text = "0"
                            txtAux(10).Text = "1"
                            txtAux(6).Enabled = False
                            txtAux(7).Enabled = False
                            txtAux(7).Enabled = False
                            If txtAux(1).Text = "" Then txtAux(2).Text = ""

                        End If
                    End If
                End If
            End If
        
        Case 2 'Nombre Articulo
           If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Si es factura rectifica la cantidad solo puede ser negativa
                If CodTipoMov = "ART" Or CodTipoMov = "DEV" Then
                    If CCur(txtAux(Index)) >= 0 Then
                        MsgBox "La cantidad debe ser negativa.", vbExclamation
                        'De momento lo quito
                        'PonerFoco txtAux(Index)
                        'Exit Sub
                    End If
                End If
            
                'Comprobar si hay suficiente stock
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub
                If vCStock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
                  If Not vCStock.MoverStock(False, False, False) Then
                    PonerFoco txtAux(Index)
                    Set vCStock = Nothing
                    Exit Sub
                  End If
                End If
                
                B = False
                If Modo = 5 Then 'Modo lineas
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    
                    If devuelve <> "" Then
                        '- obtener el nº bultos: cantidad/unids.caja
                        txtAux(10).Text = CalcularNumBultos2(CCur(txtAux(3).Text), CInt(devuelve))
                    End If
                    
                    If ModificaLineas = 1 Then 'insertar linea
                        B = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If data2.Recordset!codArtic <> txtAux(1).Text Then B = True
                    End If
                End If
                
                If B Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    If devuelve <> "" Then
'                        '- obtener el nº bultos: cantidad/unids.caja
'                        txtAux(10).Text = CalcularNumBultos(CCur(txtAux(3).Text), CInt(devuelve))
                        
                    
                        Set CPrecioFact = New CPreciosFact
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        cantidad = txtAux(Index).Text
                        
                        
                        'Mayo 2009
                        'Si este parametro esta a FALSE, siempre cojera precio ud
                        If vParamAplic.CajasCompletas Then
                            NumCajas = CPrecioFact.ObtenerNumCajas(cantidad, devuelve)
                            RestoUnid = CInt(ComprobarCero(cantidad)) - NumCajas * CInt(devuelve)
                        Else
                            NumCajas = 0
                            If CCur(devuelve) > 1 Then
                                If CCur(txtAux(3).Text) >= CCur(devuelve) Then NumCajas = 1
                            End If
                            RestoUnid = 0
                        End If
                        
                        CPrecioFact.CodigoClien = Text1(4).Text
                        
                        'Obtenemos la Tarifa del Cliente
                        'AHORA ESTA DENTRO DE LA CLASE
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        'CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.FijarTarifaActividad
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        
                        Dim Comision As String
                        
                        PorCaja = (NumCajas > 0)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, Comision)
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Artículo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And B) Then
                                txtAux(4).Text = Precio
                                'txtAux(5).Text = OrigP 'De donde viene el precio
                            Else
                                OrigP = txtAux(5).Text
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                            txtAux(5).Text = OrigP 'De donde viene el precio
                        End If
                        
                        
                        If Comision <> "" Then
                            If ComisionCliente > 0 Then
                                If CCur(Comision) > ComisionCliente Then Comision = ComisionCliente
                            End If
                        End If
                        txtAux(12).Text = Comision
                        
                        'Si tiene dto permitido
                        If Not CPrecioFact.DtoPermitido Then
                            txtAux(6).Text = "0"
                            txtAux(7).Text = "0"
                            txtAux(6).Enabled = False
                            txtAux(7).Enabled = False
                        End If
                        
'                            PonerFoco txtAux(Index + 1)

                        'Si es articulo de canje
                        If vParamAplic.PtosAsignar > 0 Then
                            If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                                txtAux(4).Text = Format(vParamAplic.PtosEquivalencia, FormatoPrecio)
                                txtAux(6).Text = "0"
                                txtAux(7).Text = "0"
                                txtAux(10).Text = "1"
                                txtAux(6).Enabled = False
                                txtAux(7).Enabled = False
                                txtAux(7).Enabled = False
                            
                            
                                devuelve = ""
                                If CStr(Data1.Recordset!codtipom) <> "ALR" Then
                                    devuelve = DevuelveDesdeBD(conAri, "puntos", "sclien", "codclien", CStr(Data1.Recordset!codClien))
                                    If devuelve = "" Then devuelve = 0
                                    If Abs(cantidad) > CCur(devuelve) Then
                                        MsgBox "Utiliza mas puntos de los que tiene", vbExclamation
                                        txtAux(3).Text = "-" & Format(devuelve, FormatoCantidad)
                                    Else
                                        devuelve = ""
                                    End If
                                End If
                                If devuelve = "" Then PonerFocoBtn Me.cmdAceptar
                            End If
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
            End If
            
            
        Case 4 'Precio
             If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then txtAux(5).Text = "M"
                End If
            End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
            
        Case 9
            ' ---- [19/10/2009] [LAURA]: añadir centro de coste a la linea
            If txtAux(9).Text = "" Then
                 Text2(9).Text = ""
            Else
                If vEmpresa.TieneAnalitica Then
                    'centro de coste
                    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                    Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            
                Else
                    'Cod proveeee
'                    If txtAux(9).Text = "" Then
'                        devuelve = ""
'                    Else
                        If Not IsNumeric(txtAux(9).Text) Then
                            MsgBox "Campo proveedor debe ser numérico", vbExclamation
                            devuelve = ""
                        Else
                                
                            devuelve = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtAux(9).Text, "N")
                            If devuelve = "" Then MsgBox "No existe el proveedor: " & txtAux(9).Text, vbExclamation
                        End If
                        If devuelve = "" Then
                            txtAux(9).Text = ""
                            PonerFoco txtAux(9)
                        End If
'                    End If
                    Text2(9).Text = devuelve
                End If
            End If
        
           

           
    End Select
    
     If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    TituloLinea = cad
    ModificaLineas = 0
    
    cad = "comision"
    ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text, "N", cad)) = 1
    If vParamAplic.ArtReciclado = "" Then ClienteConTasaReciclado = False
    
    If cad = "" Then cad = "0"
    ComisionCliente = CCur(cad)
    Set vAgent = Nothing
    Set vAgent = New cAgente
    If Not vAgent.LeerDatos(CStr(Data1.Recordset!CodAgent)) Then Exit Sub

    If vParamAplic.TipoPortes = 1 Then
        KilosAnteriores = SumaKilosLineas
    End If
    
    PonerModo 5
    PonerBotonCabecera True
    
    AlmacenLineas = -1
    PonerUltAlmacen
    EsNuevoAlbaran = False
End Sub


Private Function Eliminar(NumAlbElim As Long) As Boolean
Dim SQL As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim MenError As String
Dim ParaElLog As String

    On Error GoTo FinEliminar
    conn.BeginTrans
    
    SQL = ObtenerWhereCP(False)
    
    MenError = DevuelveDesdeBD(conAri, "concat(count(*),'|',sum(importel),'|')", "slialb", Replace(SQL, "scaalb", "slialb") & " AND 1", "1")
    If MenError = "" Then MenError = "0|0|"
    ParaElLog = "Albaran: " & Text1(30).Text & Text1(0).Text & " de " & Text1(1).Text & vbCrLf
    ParaElLog = ParaElLog & Text1(4).Text & " " & Text1(5).Text & vbCrLf
    ParaElLog = ParaElLog & "Base " & Text3(56) & " TOTAL " & Text3(55).Text & vbCrLf
    ParaElLog = ParaElLog & "Lineas " & RecuperaValor(MenError, 1) & ".  Importe: " & RecuperaValor(MenError, 2)
    
    
    'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
    MenError = "Restableciendo stocks de almacen."
    
    If CodTipoMov = "DEV" Then
        B = True 'No reestblecemos stock
    Else
        B = ReestablecerStock
    End If
    
    
    
    If B Then
        'eliminamos de albaranes y pasamos al historico
        'Para los DEV NO
        If CodTipoMov <> "DEV" Then
            Screen.MousePointer = vbHourglass
            B = ActualizarElTraspaso(MenError, SQL, CodTipoMov, cadList)
            Screen.MousePointer = vbDefault
        Else
            'Borramos de scaalb
            SQL = ObtenerWhereCP(True)
            conn.Execute "DELETE from slialb " & Replace(SQL, "scaalb", "slialb")
            conn.Execute "DELETE from scaalb " & SQL
            
        End If
        
        If B Then
            MenError = "Actualizando numeros de serie."
            'Actualizar los posibles num. serie de ese albaran. vaciar los campos
            SQL = "UPDATE  sserie SET codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
            SQL = SQL & " fechavta=" & ValorNulo & ", numalbar=" & ValorNulo & ", numline1=" & ValorNulo
            SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
            SQL = SQL & " WHERE codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!NumAlbar & " AND fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
            conn.Execute SQL
            
            
           
            
            
            
            If B Then
                'Actualiamos el riesgo
                If CodTipoMov <> "DEV" Then
                    If vParamAplic.OperacionesAseguradas Then
                        lblIndicador.Caption = "Riesgo"
                         lblIndicador.Refresh
                        SQL = DevuelveDesdeBD(conAri, "credipriv", "sclien", "codclien", Text1(4).Text, "N")
                        If SQL = "" Then SQL = "9"
                        If Val(SQL) < 9 Then
                            'Febrero 2018 . YA NO
                            'OK tiene credito. Que actualice
                            'ActualizaRiesgoCliente CLng(Text1(4).Text)
                        End If
                        lblIndicador.Caption = ""
                    End If
                End If
            End If
        End If
    End If
        
        
    If B Then
         If vParamAplic.PtosAsignar > 0 Then
            'Sistema de puntos
            If DBLet(Data1.Recordset!Puntos, "N") <> 0 Then
                
                
                'Si cambia el cliente, hay que ver
                SQL = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text)
                If Val(SQL) = "1" Then
                    'El nuevo cliente tiene puntos
                    SQL = "-"
                    If Data1.Recordset!Puntos < 0 Then SQL = "+"
                    SQL = "UPDATE sclien set puntos=coalesce(puntos,0) " & SQL & DBSet(Abs(Data1.Recordset!Puntos), "N")
                    SQL = SQL & " WHERE codclien =" & Text1(4).Text
                    conn.Execute SQL
                End If
            
            End If
            
            'Borro de smovalpuntos
            SQL = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
            SQL = SQL & " AND codclien = " & Data1.Recordset!codClien
            SQL = SQL & " AND concepto = 0"
            
                
            SQL = "DELETE FROM smovalpuntos " & SQL
            conn.Execute SQL
            
            
        End If
    End If

FinEliminar:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, MenError, Err.Description
    End If
    If Not B Then
        conn.RollbackTrans
        
        
    Else
        conn.CommitTrans
        
        '////////////////
        Set LOG = New cLOG
        LOG.Insertar 34, vUsu, ParaElLog
        Set LOG = Nothing
        
        
        
        'Lo ponermos FUERA de la transaccion YA que lleva commits y demas
        Set vTipoMov = New CTiposMov
        B = CBool(vTipoMov.DevolverContador(CodTipoMov, NumAlbElim))
        Set vTipoMov = Nothing
        
        
    End If
    Eliminar = B
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid1, data2, False
    If vParamAplic.CartaPortes Then CargaGrid DataGrid2, data3, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
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


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " " & NombreTabla & ".codtipom= '" & Text1(30).Text & "' and " & NombreTabla & ".numalbar= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NombreTabla & ".fechaalb=" & DBSet(Text1(1).Text, "F")
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean, QueGRid As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    
    If QueGRid = 1 Then
       SQL = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
       If vEmpresa.TieneAnalitica Then
           SQL = SQL & ",codccost"
       Else
           'COTUBRE
           If vParamAplic.NumeroInstalacion = 2 Then
               
               If MostrarComision Then
                   SQL = SQL & ",comisionagente,'' nomprove"   'LO que habia antes de Sep 2018
               Else
                   SQL = SQL & ",codprovex,'' nomprove"   'NUEVO
               End If
           Else
               SQL = SQL & ",codprovex,nomprove"
           End If
       End If
       SQL = SQL & ",numlote "
       
       
       'Para herbelca, ponemos el codprove al final
       If vParamAplic.NumeroInstalacion = 2 Then
           SQL = SQL & ",codprovex,if(pvpinferior=0,'',if(pvpinferior=1,'*','m')) tipoprecio"
           If Not MostrarComision Then SQL = SQL & ",comisionagente"   'para updaear los campos
       End If
    
       
       
       SQL = SQL & " FROM " & NomTablaLineas
       'traza
       If vEmpresa.TieneAnalitica = False Then
           SQL = SQL & " LEFT JOIN sprove on codprovex=codprove "
       End If
       
       If enlaza Then
           SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
           If EsHistorico Then SQL = SQL & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
       Else
           SQL = SQL & " WHERE numalbar = -1"
       End If
       SQL = SQL & " Order by codtipom, numalbar, numlinea"
    
    Else
        'Matriculas en portes
        SQL = "SELECT codtipom,numalbar,matricula,descr FROM "
        SQL = SQL & Trim(NombreTabla) & "_portes as " & NombreTabla & " WHERE "
         
        If enlaza Then
           SQL = SQL & ObtenerWhereCP(False)
           If EsHistorico Then SQL = SQL & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
       Else
           SQL = SQL & " false "
       End If
       
        
    End If
    MontaSQLCarga = SQL
End Function





Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = False
        If Not EsHistorico Then B = (Modo = 2)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (B Or Modo = 0)
        Me.mnNuevo.Enabled = (B Or Modo = 0)
        'Modificar
        If B Then
            If Me.Data1.Recordset.EOF Then B = False
        End If
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
            
            
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
            
        Toolbar1.Buttons(8).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
            
            
            
            
        B = (Modo = 2) And Not EsHistorico
        
        'Nº Series
        Toolbar2.Buttons(1).Enabled = B
        
        'Generar Factura
        Toolbar2.Buttons(2).Enabled = B
        Toolbar2.Buttons(3).Enabled = B
        If Toolbar2.Buttons(4).Style = tbrDefault Then Toolbar2.Buttons(4).Enabled = B
        
        
        'Imprimir
        If vParamAplic.TipoPortes = 1 Then
            Toolbar2.Buttons(4).Enabled = B  'Toolbar1.Buttons(15).Enabled And vParamAplic.TipoPortes = 1
        Else
            If Toolbar2.Buttons(4).Style = tbrDefault Then Toolbar1.Buttons(4).Enabled = B
        End If
        
        
        
        BotonesToolBarAux
        
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
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numalbar", "codtipom", Text1(30).Text, "T", , "numalbar", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
           
    If bol Then
        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
        MenError = "Actualizando Fecha Movimiento del Cliente."
        bol = ActualizarFecMovCliente
        
        MenError = "Error al actualizar el contador del Pedido."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
            
            'Variable globarl que utiliza cavevimun
            InsertadoAlbaran = CLng(Text1(0).Text)
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim i As Byte

    For i = 4 To 17
        Text1(i).Text = ""
    Next i
    Text2(12).Text = ""
    Text2(14).Text = ""
    Text2(17).Text = ""
    Me.cboFacturacion.ListIndex = -1
    For i = 42 To 44
        Text1(i).Text = ""
        If i <> 44 Then Text2(i).Text = ""
    Next
    Me.Text1(45).Text = ""
    Me.Text1(46).Text = ""
    Me.Text1(47).Text = ""
    Me.Text1(48).Text = ""
    Text2(0).Text = ""
    
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim SQL As String
Dim B As Boolean
Dim ImpReciclado As Single



    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    SQL = SQL & " and numlinea=" & data2.Recordset!numlinea
    
    
    
    'Inicilizar la clase para Actualizar los stocks
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""
    
    
    On Error GoTo EEliminarLinea
    
    conn.BeginTrans
    conn.Execute SQL 'Eliminar linea
    If hcoCodTipoM <> "DEV" Then
        B = vCStock.DevolverStock2
    Else
        B = True
    End If
    Set vCStock = Nothing
    
    If B Then
        'Ha borrado la linea y ha devuelvto correctamente el sctock
                   'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    SQL = CStr(data2.Recordset!codArtic)
                    If ArticuloConTasaReciclado(SQL, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       NumRegElim = Val(DBLet(data2.Recordset!numlinea, "N")) + 1
                       SQL = DevuelveDesdeBD(conAri, "codartic", "slialb", SQL, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                       If SQL = vParamAplic.ArtReciclado Then
                       
                            SQL = "DELETE FROM " & NomTablaLineas
                            'WHERE
                            SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                            conn.Execute SQL
                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
    End If


    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
    If hcoCodTipoM <> "DEV" Then
        SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
        SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
        SQL = SQL & " WHERE codartic=" & DBSet(data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & data2.Recordset!numlinea
        conn.Execute SQL
    End If
    
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        ReestablecerLotesArticulo data2.Recordset!numlinea
        
        'Borramos de slialblotes
        SQL = "Delete from slialblotes " & Replace(ObtenerWhereCP(True), NombreTabla, "slialblotes")
        SQL = SQL & " and numlinea=" & data2.Recordset!numlinea
        conn.Execute SQL 'Eliminar linea
    End If
        

    
    
    If vParamAplic.PtosAsignar > 0 Then
        
        'Si es la linea de canje, hay que quitarla de movimientos
        If data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
            SQL = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
            SQL = SQL & " AND codclien = " & Data1.Recordset!codClien & " AND concepto=1"
            SQL = "DELETE FROM smovalpuntos  " & SQL
            conn.Execute SQL
      
            SQL = "UPDATE sclien set puntos=" & DBSet(-1 * data2.Recordset!cantidad, "N") & " + coalesce(puntos,0) "
            SQL = SQL & " WHERE codclien =" & Text1(4).Text
            conn.Execute SQL
        End If
    End If
    
    SQL = "Albarán: " & Text1(30).Text & "-" & Text1(0).Text & " de " & Text1(1).Text & vbCrLf
    SQL = SQL & "Linea: " & data2.Recordset!codArtic & " " & DBSet(data2.Recordset!NomArtic, "T")
    SQL = SQL & "Uds: " & data2.Recordset!cantidad & "    Importe:" & DBSet(data2.Recordset!ImporteL, "N")

    Set LOG = New cLOG
    ' 17 Venta a sabiendas riesgo
    LOG.Insertar 37, vUsu, SQL
    Set LOG = Nothing
            
        
        
        
    'Si tiene
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        EliminarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        DescuentosCantidad ElArticulo
        ' ----
        
        
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


Private Sub ReestablecerLotesArticulo(linea As Integer)
        If linea >= 0 Then
            BuscaChekc = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", data2.Recordset!codArtic, "T")
        Else
            BuscaChekc = "OK"
        End If
        If Trim(BuscaChekc) <> "" Then
            Set miRsAux = New ADODB.Recordset
            BuscaChekc = "Select * from slialblotes WHERE codtipom= '" & Data1.Recordset!codtipom & "' AND numalbar = " & Data1.Recordset!NumAlbar
            If linea >= 0 Then BuscaChekc = BuscaChekc & " AND numlinea =" & data2.Recordset!numlinea
            miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux!cantidad <> 0 Then
                    'Vamos a actualizar VENDIDA
                    'Luego si estoy reestableciendo
                    If miRsAux!cantidad < 0 Then
                        BuscaChekc = "+"
                    Else
                        BuscaChekc = "-"
                    End If
                    BuscaChekc = "UPDATE slotes SET vendida = vendida " & BuscaChekc & " " & DBSet(Abs(miRsAux!cantidad), "N")
                    BuscaChekc = BuscaChekc & " WHERE numlotes= " & DBSet(miRsAux!numLote, "T")
                    BuscaChekc = BuscaChekc & " AND codArtic= " & DBSet(miRsAux!codArtic, "T")
                    BuscaChekc = BuscaChekc & " AND fecentra = " & DBSet(miRsAux!fecentra, "F")
                    conn.Execute BuscaChekc
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
End Sub


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente del albaran
    vCStock.Documento = Text1(0).Text 'Nº Albaran
    vCStock.FechaMov = Text1(1).Text 'Fecha del Albaran
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If data2.Recordset!codArtic = txtAux(1).Text Then
                vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text)) - data2.Recordset!cantidad
            Else
                vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codArtic = data2.Recordset!codArtic
        vCStock.codAlmac = CInt(data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(data2.Recordset!cantidad)
        vCStock.Importe = CCur(data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(data2.Recordset!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Function ComprobarInventario() As Boolean
Dim CadFecInv As String
Dim vCStock As CStock

    ComprobarInventario = True
    CadFecInv = ""
    If data2.Recordset.RecordCount > 0 Then
       data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not data2.Recordset.EOF)
           Set vCStock = New CStock
           If InicializarCStock(vCStock, "E", data2.Recordset!numlinea) Then vCStock.ComprobarFechaInventario False, CadFecInv
           Set vCStock = Nothing
           
           data2.Recordset.MoveNext
        Wend
    End If

    
        If CadFecInv <> "" Then
            CadFecInv = "Fechas inventario posterior: " & CadFecInv & vbCrLf
            CadFecInv = CadFecInv & "¿Continuar?" & vbCrLf
            CadFecInv = String(40, "*") & vbCrLf & CadFecInv & String(40, "*")
            If MsgBox(CadFecInv, vbQuestion + vbYesNo) = vbNo Then ComprobarInventario = False
        End If

End Function

Private Function ReestablecerStock() As Boolean
Dim vCStock As CStock
Dim B As Boolean

    On Error GoTo ERestablecer
    
    ReestablecerStock = False
    B = True
    
    If data2.Recordset.RecordCount > 0 Then
       data2.Refresh
       data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not data2.Recordset.EOF) And B
           Set vCStock = New CStock
           If InicializarCStock(vCStock, "E", data2.Recordset!numlinea) Then
                
               'Actualiza el stock en salmac y borra de smoval
               If Not vCStock.DevolverStock2() Then B = False
           Else
               B = False
           End If
           data2.Recordset.MoveNext
           Set vCStock = Nothing
       Wend
    End If
    
    'Para tabla slotes
    If vParamAplic.ManipuladorFitosanitarios2 Then ReestablecerLotesArticulo -1
        
        
        
ERestablecer:
    If Err.Number <> 0 Then B = False
    ReestablecerStock = B
End Function


Private Sub BotonImprimir_(OpcionListado As Byte, EsInformePortes As Boolean)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 45) Then
        If EsInformePortes Then
            'Es el de portes
             indRPT = 34
        Else
            'ALBARANES
            If hcoCodTipoM = "ALZ" Then
                indRPT = 29   'Albaranes B
            ElseIf hcoCodTipoM = "ALR" Then
                indRPT = 36
            ElseIf hcoCodTipoM = "ALS" Then
                indRPT = 39
            ElseIf hcoCodTipoM = "ALI" Then
                indRPT = 56
            Else
                If EsHistorico Then
                    indRPT = 11 'Hist. Albaranes clientes
                Else
                    indRPT = 10 'Albaran Clientes
                End If
            End If
        End If
    End If
    
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, ImpresionDirecta, pPdfRpt, pRptvMultiInforme) Then Exit Sub
   
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
    
    'PORTES
    cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
    numParam = numParam + 1
    
    'PUNTO VERDE
    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
    numParam = numParam + 1
    
    'Si se imprimen importes y/o
    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
    If devuelve = "" Then devuelve = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    cadParam = cadParam & "Albarcon=" & devuelve & "|"
    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    If Not ImpresionDirecta Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        
        If EsHistorico Then
            devuelve = "{" & NombreTabla & ".codtipom}=" & DBSet(Data1.Recordset!codtipom, "T")
        Else
            devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"  'lo que habia
        End If
        
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        If EsHistorico Then
            'El campo fecha tambien es clave primaria
            devuelve = Text1(1).Text
            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            
            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(Text1(1).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
        
    End If
   
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    If CodTipoMov = "ALI" Then
        'facturas internas VAN sin IVA         Si los ALZ no
        cadParam = cadParam & "pTipoIVA=2|"
        numParam = numParam + 1
    Else
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
    End If
        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda.  -----------
        If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
    
        'En visreport hay un sub para imprmir
        davidNumalbar = 0
        If Not EsInformePortes Then
            davidCodtipom = CodTipoMov
            davidNumalbar = Val(Text1(0).Text)
        End If
    
        With frmImprimir
            'Febrero 2010
            If indRPT = 34 Then
                .outTipoDocumento = 0
            Else
                .outTipoDocumento = 4
                .outClaveNombreArchiv = Text1(30).Text & Text1(0).Text
                .outCodigoCliProv = CLng(Text1(4).Text)
                .NumeroCopias = vParamAplic.NumCop_AlbaranNormal
            End If
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            If indRPT = 34 Then
                .Titulo = "Portes albaran "
            Else
                .Titulo = "Albaran de Cliente"
            End If
            .ConSubInforme = True
            .Show vbModal
            
            
            
            
            If Not EsHistorico Then
                If Not EsInformePortes Then
                    If HaPulsadoElBotonDeImprimir Then
                        'UPDATEAMOS scaalb para que no reimpimrpima los albaranes
                        'Cod Tipo Movimiento
                        devuelve = "scaalb.codtipom = '" & CodTipoMov & "' AND scaalb.numalbar = " & Val(Text1(0).Text)
                        devuelve = "UPDATE scaalb SET albImpreso = 1 WHERE " & devuelve
                        Me.chkImpreso.Value = 1
                        ejecutar devuelve, False
                    End If
                End If
            End If
        End With
    End If
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset, Optional Dif As String, Optional cadSel As String)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'Dif: si se ha modificado la cantidad pasamos la difencia con lo que habia
Dim SQL As String
Dim Campos As String

    On Error GoTo EMostrarNSeries

    If Text1(30).Text = "ART" Then
        SQL = MostrarNSeriesGnral(RSLineas, Campos, True)
    Else
        SQL = MostrarNSeriesGnral(RSLineas, Campos)
    End If
    
   If SQL <> "" Then
        Set frmMen = New frmMensajes
        frmMen.cadWhere = SQL
        
        If Dif <> "" Then
            SQL = " WHERE (codtipom=" & DBSet(CodTipoMov, "T") & " and "
            SQL = SQL & " numalbar=" & Text1(0).Text & " and numline1=" & data2.Recordset!numlinea & ")"
            frmMen.cadWHERE2 = Dif & "|" & SQL & "|"
        Else
            If cadSel <> "" Then
                'seleccionar lineas de nº serie de la factura a rectificar
                frmMen.cadWHERE2 = cadSel
            Else
                frmMen.cadWHERE2 = ""
            End If
        End If
        frmMen.OpcionMensaje = 4 'Nº Series Articulo
        frmMen.vCampos = Campos
        frmMen.Show vbModal
        Set frmMen = Nothing
    End If
    
EMostrarNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PedirNSeries(ByRef Rs As ADODB.Recordset)
Dim SQL As String

    On Error GoTo EPedirNSeries

        SQL = "El artículo tienen control de Nº de Serie." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los Nº De Serie." & vbCrLf
        MsgBox SQL, vbInformation
        PedirNSeriesGnral Rs, False
        
        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
    
        'Herbelca , control puntos
        'If vParamAplic.PtosAsignar > 0 Then Text1(43).Text = CalcularPuntosAlbaran
    
    
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
            
                If vParamAplic.CartaPortes Then
                    'Vemos si el transportista tiene matricula por defecto
                    'Si es asi, la traemos
                    SQL = "codenvio =" & Text1(29).Text & " AND defecto "
                    SQL = DevuelveDesdeBD(conAri, "matricula", "smatriculas", SQL, "1", "T")
                    If SQL <> "" Then
                        'Tienen un vehiculo marcado com por defecto
                        
                        SQL = "," & DBSet(SQL, "T") & ", NULL)"
                        SQL = "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & SQL
                        SQL = "INSERT INTO scaalb_portes(codtipom,numalbar,matricula,descr)" & SQL
                        If Not ejecutar(SQL, False) Then MsgBox "No se ha podido insertar el vehiculo por defecto. " & vbCrLf & SQL, vbExclamation
                            
                    End If
                    
                    
                    
                    
                End If
            
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 0, "Albaranes"
                BotonAnyadirLinea False
                EsNuevoAlbaran = True
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    Me.SSTab1.Tab = 0
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub ComprobarNSeriesLineas(numlinea As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'Si NO se realiza control Nº series en compras pedirlos ahora
'Si se realiza control Nº Series en compras verificar que efectivamente estan introducidos
'y mostrarlos para seleccionarlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
Dim Dif As Single

    'Comprobar si el Articulo tiene control de Nº de Serie
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If SQL = "1" Then 'Hay NºSerie para el Articulo
        'si estamos insertando
        If Modo = 5 Then
            If ModificaLineas = 1 Then 'Insertar
                'Comprobar que la cantidad comprada es >0
                If ComprobarCero(txtAux(3).Text) <= 0 Then Exit Sub
            ElseIf ModificaLineas = 2 Then 'Modificar
                'si se ha modificado la cantidad, habrá que quitar algun nº serie
                'de los seleccionado o anyadir alguno mas
                Dif = CSng(txtAux(3).Text) - CSng(data2.Recordset!cantidad)
                If Dif = 0 Then Exit Sub
                If Text1(30).Text = "ART" Then Exit Sub
'                    Dif = CSng(Data2.Recordset!Cantidad) - CSng(txtAux(3).Text)
            End If
        End If
        
        cadWhere = " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and "
        cadWhere = cadWhere & " numalbar=" & Text1(0).Text & " and numlinea=" & numlinea
    
        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
        SQL = "SELECT slialb.codartic, sum(cantidad) as cantidad, numlinea "
        SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " And nseriesn = 1 "
        SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Me.cmdAux(1).Tag = Text1(0).Text 'Num Albaran
        Me.cmdAux(0).Tag = numlinea 'Num Linea
        
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        If Not vParamAplic.NumSeries And ModificaLineas = 1 Then
            PedirNSeries RSLineas
        Else 'Se realizo contro en COMPRAS, Mostramos los Nº y seleccionamos
            If ModificaLineas = 1 Then 'Insertando la linea
                'Comprobar que efectivamente estan en tabla sserie los NºSerie del Articulo
                ' y que no esten asignados ya a otro albaran de venta
                SQL = " select distinct count(numserie) from sserie where codartic=" & DBSet(txtAux(1).Text, "T") & " and (numalbar='' or isnull(numalbar))"
                '=== Laura 17/01/2007
                'y que no este asignados a una factura
                SQL = SQL & " and (numfactu='' or isnull(numfactu))"
                '===
                If RegistrosAListar(SQL) = 0 Then 'No hay Nº de Serie para elegir
                    PedirNSeries RSLineas
                Else
                    MostrarNSeries RSLineas
                End If
            ElseIf ModificaLineas = 2 Then
                SQL = " select distinct count(numserie) from sserie " & Replace(cadWhere, "numlinea", "numline1")
                If RegistrosAListar(SQL) > 0 Then
                    MostrarNSeries RSLineas, CStr(Dif)
                End If
            End If
        End If

        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub BotonNSeries()
Dim cadWhere As String, SQL As String
Dim RSLineas As ADODB.Recordset

    If Me.Data1.Recordset!EsTicket Then
        MsgBox "Albaranes provenientes de Ticket no tienen control de Nº Serie.", vbInformation
        Exit Sub
    End If

    'Si es Albaran para Factura rectificativa (ART)
    If CodTipoMov = "ART" Then
'      'Si es una Factura Venta(FAV) generada desde un ticket del TPV entonces
'      'no hay numseries
'      SQL = DevuelveDesdeBDNew(conAri, "scafac1", "codtipoa", "codtipom", Data1.Recordset!codtipmf, "T", , "numfactu", Data1.Recordset!NumFactu, "N", "fecfactu", Data1.Recordset!FecFactu, "F")
'      If SQL = "FTI" Then
'        MsgBox "Facturas provenientes de Ticket no tienen control de Nº Serie.", vbInformation
'        Exit Sub
'      Else
        Exit Sub
'      End If
    End If
    
    
    
    ModificaLineas = 4

    cadWhere = " WHERE codtipom='" & Text1(30).Text & "'"
    cadWhere = cadWhere & " and numalbar=" & Text1(0).Text
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    SQL = "SELECT numlinea, slialb.codartic, sum(cantidad) as cantidad "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    
    'Pudioera ser que tuvieran un mismo articulo wen dos lineas, y por lo tanto
    'el articulo tendria numeros de sr asociados a distintas lineas
    'por lo tanto hay que agrupar por numlinea TB
    SQL = SQL & " GROUP BY codartic,numlinea ORDER BY Codartic "
    

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        PedirNSeriesT RSLineas
    Else
        MsgBox "No hay ninguna linea de Articulo con Control de Nº Serie", vbInformation
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    ModificaLineas = 0
End Sub


Private Sub PedirNSeriesT(ByRef Rs As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim SQL As String
Dim linea As Integer

    On Error GoTo EPedirNSeries


        PedirNSeriesGnral Rs, False
        Rs.MoveFirst
        While Not Rs.EOF
            linea = 0
            'Cargar los Nº de serie asignados
            SQL = "SELECT numserie, codartic,nummante FROM sserie "
            SQL = SQL & " WHERE codtipom='" & Text1(30).Text & "' and "
            SQL = SQL & "numalbar=" & Text1(0).Text
            SQL = SQL & " and numline1=" & Rs!numlinea
            SQL = SQL & " ORDER BY codartic "
            Set RSseries = New ADODB.Recordset
            RSseries.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                SQL = SQL & ", nummante = " & DBSet(RSseries!nummante, "T")
                SQL = SQL & " WHERE codartic=" & DBSet(Rs!codArtic, "T")
                SQL = SQL & " and numlinealb=" & Rs!numlinea
                SQL = SQL & " and numlinea=" & linea
                conn.Execute SQL
                RSseries.MoveNext
            Wend
            Rs.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
        
        
        'Igual aqui deberiamos poner si Linea=0 NO seguimos
        
        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
        frmNSerie.NumAlb = Text1(0).Text
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal o actualizarlo
Dim RStmp As ADODB.Recordset
Dim SQL As String
Dim B As Boolean

    On Error GoTo ECargar
    
    conn.BeginTrans
    
    'Limpiar primero los Nº de serie asignados al ALV y luego volver a cargarlos
    SQL = "UPDATE sserie SET codtipom=" & ValorNulo & ", numalbar=" & ValorNulo & ", fechavta="
    SQL = SQL & ValorNulo & ", numline1=" & ValorNulo
    'Enero 2010
    'Tambien reestablezco los valores de tieneman y numeromantenimiento
     SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ", coddirec= " & ValorNulo
    
    SQL = SQL & " WHERE codtipom=" & DBSet(Text1(30).Text, "T") & " and numalbar=" & Text1(0).Text & " AND year(fechavta)=" & Year(Text1(1).Text)
    conn.Execute SQL
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.codigo
    SQL = SQL & " ORDER BY codartic, numlinealb, numlinea "
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
    B = True
    While Not RStmp.EOF And B
        B = InsertarNSerie(RStmp!numSerie, RStmp!codArtic, RStmp!numlinealb, DBLet(RStmp!nummante, "T"))
        RStmp.MoveNext
    Wend
    RStmp.Close
    Set RStmp = Nothing
    
ECargar:
    If Err.Number <> 0 Then B = False
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
End Sub


Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String, nummante As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
Dim devuelve As String
Dim TieneMan As Boolean
Dim NumAlbar As String
Dim nSerie As CNumSerie
Dim B As Boolean

    On Error GoTo EInsertarNSerie


    'Enero 2010
    'AHora si tiene mantenimiento lo habra indicado en la introduccion de numero de serie
    '
    ''Comprobar que el cliente tiene mantenimientos en esa direc/dpto
    'TieneMan = "0"
    'devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    ''El cliente tiene Mantenimientos
    'If devuelve <> "" Then TieneMan = "1"
    nummante = Trim(nummante)
    TieneMan = nummante <> ""



    Set nSerie = New CNumSerie
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    
    nSerie.Cliente = CLng(Text1(4).Text)
    nSerie.DirDpto = Text1(12).Text
    nSerie.conMante = TieneMan
    nSerie.tipoMov = CodTipoMov
    nSerie.FechaVta = Text1(1).Text
    nSerie.NumAlbaran = Text1(0).Text
    nSerie.NumLinAlb = numlinea
    nSerie.ObtenFechaFinGarantia codArtic, Text1(1).Text
    nSerie.nummante = nummante   'Si ha indicado el numero de mantenimiento
    
    
    

    
    'Comprobar si existe en la tabla sserie
     NumAlbar = "numalbar" 'Nº albaran de Venta
     devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codArtic, "T")
     If devuelve <> "" Then 'EXISTE en tabla sserie
        If NumAlbar = "" Then B = nSerie.ActualizarNumSerie(True)
     Else
        B = nSerie.InsertarNumSerie
    End If
    InsertarNSerie = True
    Set nSerie = Nothing
    
EInsertarNSerie:
    If Err.Number <> 0 Then B = False
    If B Then
        InsertarNSerie = True
    Else
        InsertarNSerie = False
    End If
End Function




Private Sub PosicionarDataTrasEliminar()
Dim HayDatos As Boolean
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    HayDatos = SituarDataTrasEliminar(Data1, NumRegElim)
    If HayDatos Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            Data1.Recordset.MoveLast
            If Data1.Recordset.EOF Then HayDatos = False
        End If
    End If
    If HayDatos Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim B As Boolean
    
    
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
            If vCliente.ClienteBloqueado Then
                If Me.hcoCodTipoM <> "ALM" Then
                    LimpiarDatosCliente
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)v
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 3 And EsDeVarios Then 'NUEVO
                If Me.hcoCodTipoM = "ALV" Then
                    If vParamAplic.FrasMostradorSerieDistinta Then
                        'Es de varios y tienen serie de facturacion distinta....
                        Observaciones = "Esta realizando un albaran de venta(FAV) a un cliente de varios." & vbCrLf
                        Observaciones = Observaciones & "Debería ser una factura de mostrador.      ¿Continuar?"
                        If MsgBox(Observaciones, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            LimpiarDatosCliente
                            PonerFoco Text1(4)
                            Set vCliente = Nothing
                            Exit Sub
                        End If
                        Observaciones = ""
                    End If
                End If
            End If
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
                
            End If
            
            If Modo = 3 Or Modo = 4 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Text1(34).Text = vCliente.Kilometros
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
                Text1(29).Text = vCliente.FEnvio
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
                
                Text1(43).Text = vCliente.Zona
                Text2(43).Text = PonerNombreDeCod(Text1(43), conAri, "szonas", "nomzonas")
                
                
                'Si tiene portes
                If vParamAplic.CartaPortes Then POnerChoferDefecto
                    
                
                
                vCliente.PonDatosDireccionEnvio Text1(42), Text2(42)
                
                
                'Febrero 2013
                'Si tiene observaciones del departamento de comercial, van a observaCRM
                Text1(44).Text = DevuelveDesdeBD(conAri, "observa", "scrmobsclien", "dpto=2 AND codclien", codClien)
                
            End If
            Me.Text1(45).Text = "": Me.Text1(46).Text = "": Me.Text1(47).Text = "": Me.Text1(48).Text = "": Text2(0).Text = ""
            
           

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                           
            
            'Comprobar si el cliente tiene cobros pendientes
            'MARZO 2014
            'Para HERBELCA, en mostrador, no comprobaremos los cobros pendientes
            B = True
            If hcoCodTipoM = "ALM" Then
            
                If vParamAplic.NumeroInstalacion = 2 Then
                'If vParamAplic.AlmacenB > 90 Then
                    B = False
                Else
                    If vParamAplic.EntradaRapidaFacturasMostrador Then B = False
                End If
            End If
            If B Then ComprobarCobrosCliente codClien, Text1(1).Text
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
Dim RN As ADODB.Recordset
Dim Aux As String

    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    B = vCliente.LeerDatosCliVario(nifClien)
    If B Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
            
    'Si tiene manipulador de fitosnaitarios
    If B Then
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Set RN = New ADODB.Recordset
            Aux = "Select ManipuladorNumCarnet , fcaducidad "
            Aux = Aux & ",ManipuladortipoCarnet from sclvar WHERE nifclien = " & DBSet(nifClien, "T")
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Aux = "|||||"
            If Not RN.EOF Then
                Aux = DBLet(RN!ManipuladorNumCarnet, "T") & "|"
                Aux = Aux & vCliente.Nombre & "|"
                If Not IsNull(RN!fcaducidad) Then Aux = Aux & Format(RN!fcaducidad, "dd/mm/yyyy")
                Aux = Aux & "|"
                'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
                If Val(DBLet(RN!ManipuladortipoCarnet, "N")) > 0 Then
                    Aux = Aux & IIf(RN!ManipuladortipoCarnet = 2, "Cualificado", "Básico") & "|"
                    Aux = Aux & RN!ManipuladortipoCarnet & "|"
                Else
                    Aux = Aux & "||"
                End If
            End If
            RN.Close
            Set RN = Nothing
            Me.Text1(45).Text = RecuperaValor(Aux, 1)
            Me.Text1(46).Text = RecuperaValor(Aux, 2)
            Me.Text1(47).Text = RecuperaValor(Aux, 3)
            Text2(0).Text = RecuperaValor(Aux, 4)
            'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            Me.Text1(48).Text = RecuperaValor(Aux, 5)
        End If
    End If
            
'    If Not b Then PonerFoco Text1(6)
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
         If Not Comprobar_NIF(NIF) Then
            If MsgBox("El NIF es incorrecto. ¿Continuar de igual modo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        
        End If
       
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


Private Function ActualizarFecMovCliente() As Boolean
Dim vCliente As CCliente
Dim B As Boolean

    On Error GoTo EActFecha

    ActualizarFecMovCliente = False
    Set vCliente = New CCliente
    vCliente.codigo = Text1(4).Text
    B = vCliente.ActualizaUltFecMovim(Text1(1).Text)
    Set vCliente = Nothing
    
EActFecha:
    If Err.Number <> 0 Then B = False
    ActualizarFecMovCliente = B
End Function


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim cadWhere As String, SQL As String
Dim vFactu As CFactura
Dim CambiarValoresIVA As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 56
         Text3(i).Text = ""
    Next i

    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP(False)
    SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(SQL) = 0 Then Exit Sub
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    'Si es Presupuesto*  o es ART tb el codtipom
    CambiarValoresIVA = False
    If hcoCodTipoM = "ALZ" Or hcoCodTipoM = "ALI" Then vFactu.codtipom = hcoCodTipoM
    
    If hcoCodTipoM = "ART" Then
        If Text1(35).Text <> "" Then CambiarValoresIVA = CDate(Text1(35).Text) < vParamAplic.FechaCambioIva
    End If
        

    
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, CambiarValoresIVA) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = vFactu.TipoIVA1
        Text3(38).Text = vFactu.TipoIVA2
        Text3(39).Text = vFactu.TipoIVA3
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
                '---- Laura: Modificado 27/09/2006
'                Text3(i + 3).Text = QuitarCero(Text3(i).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
                '----
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
End Function



Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.codigo = Text1(4).Text
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
    vClien.codigo = Text1(4).Text
    If vClien.TieneRefObligatoria(Text1(13).Text) Then
        If Text1(13).Text = "" Then PonerFoco Text1(13)
    End If
    Set vClien = Nothing
End Sub



 Private Sub InsertarLineasFactu(cadWhere)
'cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
'cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
 Dim Rs As ADODB.Recordset
 Dim SQL As String
 Dim i As Integer
 Dim cadI As String
 Dim numlin As String
 Dim CCos As String   'por si acaso lleva analitica y la linea NO lo llevaba
 
 Dim RL As ADODB.Recordset
 
    On Error GoTo EInsFactu
    Screen.MousePointer = vbHourglass
    
    If cadWhere <> "" Then
        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialb", "numlinea", SQL)
'        i = Int(SQL)
            
        
        cadI = ""
    
        SQL = "SELECT * FROM slifac WHERE " & cadWhere
        Set RL = New ADODB.Recordset
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            txtAux(0).Text = Rs!codAlmac
            txtAux(1).Text = Rs!codArtic
            txtAux(2).Text = Rs!NomArtic
            Text2(16).Text = DBLet(Rs!Ampliaci, "T")
'            Text2(9).Text = DBLet(RS!nomprove, "T")
            txtAux(3).Text = CStr(Rs!cantidad * -1)
            txtAux(4).Text = Rs!precioar
            txtAux(5).Text = DBLet(Rs!origpre, "T")
            txtAux(6).Text = Rs!dtoline1
            txtAux(7).Text = Rs!dtoline2
            txtAux(8).Text = CStr(Rs!ImporteL * -1)
            
            ' ---- [21/10/2009] [LAURA] : se añade el centro de coste
            If Not vEmpresa.TieneAnalitica Then
                txtAux(9).Text = DBLet(Rs!codProvex, "N")
                Text2(9).Text = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(Rs!codProvex), "N")
            Else
                CCos = DBLet(Rs!CodCCost)
                If CCos = "" Then
                    'MAL. DEBERIA tener Analitica.
                    If vParamAplic.ModoAnalitica = 1 Then CCos = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND codartic", CStr(Rs!codArtic), "T")
                    If CCos = "" Then CCos = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Text1(3).Text)
                End If
                txtAux(9).Text = CCos
                Me.Text2(9).Text = PonerNombreCCoste(txtAux(9))
            End If
            
            'para no tener que traer ahora el proveedor pongo en txt(10) un texto
'            txtAux(10).Text = "*"
'            Text2(9).Text = "*"
            
            'numbultos
            txtAux(10).Text = CStr(Rs!NumBultos * -1)
            'numlote
            txtAux(11).Text = DBLet(Rs!numLote, "T")
            
            
            txtAux(12).Text = ""
            VendeAMenorPrecio = 0
            If vParamAplic.NumeroInstalacion = 2 Then
                'HERBELCA
                VendeAMenorPrecio = DBLet(Rs!PVPInferior, "N")
                txtAux(12).Text = DBLet(Rs!comisionagente, "N")
            End If
            
            
            
            If InsertarLinea(numlin, True) Then
                If vParamAplic.ManipuladorFitosanitarios2 Then
                    'Vere si esa linea tenia fitosanitarios y los meto
                    SQL = "SELECT * FROM slifaclotes WHERE " & cadWhere
                    SQL = SQL & " AND numlinea= " & Rs!numlinea
                    
                    RL.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RL.EOF
                        SQL = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic) VALUES "
                        SQL = SQL & "('" & Text1(30).Text & "', " & Val(Text1(0).Text) & "," & Rs!numlinea & ","
                        SQL = SQL & RL!sublinea & "," & DBSet(-RL!cantidad, "N") & ","
                        SQL = SQL & DBSet(RL!numLote, "T") & "," & DBSet(RL!fecentra, "F") & "," & DBSet(Rs!codArtic, "T") & ")"
                        If ejecutar(SQL, False) Then
                            SQL = "UPDATE slotes SET vendida=vendida "
                            If RL!cantidad > 0 Then
                                SQL = SQL & " - "
                            Else
                                SQL = SQL & " + "
                            End If
                            SQL = SQL & DBSet(Abs(RL!cantidad), "N") & " WHERE codartic = " & DBSet(Rs!codArtic, "T")
                            SQL = SQL & " AND numlotes = " & DBSet(RL!numLote, "T")
                            SQL = SQL & " AND fecentra = " & DBSet(RL!fecentra, "F")
                            ejecutar SQL, False
                        End If
                        RL.MoveNext
                    Wend
                    RL.Close
                End If
            End If
            
        
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        
      
        
        
        
        
        
        Set RL = Nothing
        Set Rs = Nothing
        
        CalcularDatosFactura
        
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
'            SQL = SQL & cadI
'            Conn.Execute SQL
'        End If
    End If
    Screen.MousePointer = vbDefault
    
EInsFactu:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Lineas Factura", Err.Description
    End If
End Sub



Private Function AsignarNumSeriesAlbVenta(cadSel As String) As Boolean
Dim i As Integer
Dim Cant As Integer
Dim cadSerie As String
Dim nSerie As CNumSerie
Dim devuelve As String
Dim B As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
    
    
    Cant = CInt(ComprobarCero(txtAux(3).Text))
    
    On Error GoTo ErrorNSerie
    conn.BeginTrans
        
    If ModificaLineas = 2 Then 'Venimos de modificar la cantidad de una linea
        'Borramos los numeros de serie que tenia asignada la linea del albaran
        Set nSerie = New CNumSerie
        nSerie.tipoMov = CodTipoMov
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.FechaVta = Text1(1).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
        B = nSerie.BorrarNumSeriesAlbVta
        Set nSerie = Nothing
    Else
        B = True
    End If
        
    If B Then
        Set nSerie = New CNumSerie
        nSerie.Articulo = txtAux(1).Text
        nSerie.Cliente = CLng(Text1(4).Text)
        nSerie.DirDpto = Text1(12).Text
        nSerie.tipoMov = CodTipoMov
        nSerie.FechaVta = Text1(1).Text
        If nSerie.FechaVta <> "" Then
            devuelve = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", txtAux(1).Text, "T")
            nSerie.FinGarantia = CStr(CDate(nSerie.FechaVta) + CInt(ComprobarCero(devuelve)))
        End If
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
                
        For i = 1 To Cant
            cadSerie = RecuperaValor(cadSel, i + 1)
            If cadSerie <> "" Then
                nSerie.numSerie = cadSerie
                If nSerie.ActualizarNumSerie(True) = False And B Then B = False
            End If
        Next i
        Set nSerie = Nothing
    End If
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    AsignarNumSeriesAlbVenta = B
End Function




Private Sub DevolverNumSeries()
Dim SQL As String
Dim cadWhere As String
Dim Rs As ADODB.Recordset

    On Error GoTo EDevNumSerie
        
        
    If IsNull(Me.Data1.Recordset!Numfactu) Then Exit Sub
        
    cadWhere = ObtenerWhereCP(True)
    SQL = "select slialb.codartic,abs(cantidad) as cantidad,numlinea"
    SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
    SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu and sserie.codclien=scaalb.codclien "
    '-- LAURA: 02/07/2007
'    SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'    SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
    SQL = SQL & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!Numfactu)
'    If Me.Data1.Recordset!codtipmf = "FAV" Then SQL = SQL & " AND codtipom='ALV'"
    '--

    
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Hay articulos con nº de serie en las lineas del albaran rectificativo
    'que hay que quitar de los nº de serie que tenia asignados
    'estamos devolviendo nº serie y pedimos los que vamos a devolver y a estos
    'le limpiamos los campos de venta de la tabla sserie
    If Not Rs.EOF Then
        SQL = "select sserie.numserie, sserie.codartic, sartic.nomartic"
        SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
        '-- LAURA: 02/07/2007
'        SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'        SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
        SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu  and sserie.codclien=scaalb.codclien "
        '--
        SQL = SQL & " inner join sartic on sserie.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!Numfactu)
    
        MostrarNSeries Rs, , SQL
    End If
    Rs.Close
    Set Rs = Nothing
    
EDevNumSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizando Nº Serie.", Err.Description
    End If
End Sub




Private Function QuitarNumSeriesAlbVenta(cadSel As String) As Boolean
Dim i As Integer
Dim numSerie As String
Dim codArtic As String
Dim nSerie As CNumSerie
Dim Grupo As String
Dim B As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
   
    On Error GoTo ErrorNSerie
    
    B = True
    While cadSel <> ""
        i = InStr(1, cadSel, "·")
        If i > 0 Then
            Grupo = Mid(cadSel, 1, i - 1)
            cadSel = Mid(cadSel, i + 1, Len(cadSel))
            If Grupo <> "" Then
                codArtic = RecuperaValor(Grupo, 1)
                numSerie = RecuperaValor(Grupo, 2)
                
                Set nSerie = New CNumSerie
                nSerie.numSerie = numSerie
                nSerie.Articulo = codArtic
                B = B And nSerie.ActualizarNumSerie(True)
                Set nSerie = Nothing
            End If
        End If
    Wend
   
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        Set nSerie = Nothing
        B = False
    End If
    QuitarNumSeriesAlbVenta = B
End Function





Private Sub MarcarAlbaranes()
    
        If EsHistorico Then Exit Sub
    
        If hcoCodTipoM = "ALT" Then
            MsgBox "No se puede realizar sobre albaranes de telefonía", vbExclamation
            Exit Sub
        End If
    
        'Lanzara un desde hasta y los marcara
        frmListado.NumCod = hcoCodTipoM
        CadenaDesdeOtroForm = ""
        AbrirListado 82
        If CadenaDesdeOtroForm = "OK" Then
            'OK. Cambiadas las marcas. Refrescamos y situamos
            Screen.MousePointer = vbHourglass
            DoEvents
            PonerCadenaBusqueda
            PosicionarData
            Screen.MousePointer = vbDefault
        End If
        
End Sub

'FALTA###
'No se el porque del importe
Private Function SumaKilosLineas(Optional ImporteL As Currency) As Currency
Dim C As String
    On Error GoTo ESumaKilosLineas
    SumaKilosLineas = 0
    Set miRsAux = New ADODB.Recordset
    C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    C = C & " AND slialb.codartic=sartic.codartic"
    C = C & " AND slialb.codartic <> " & DBSet(vParamAplic.ArtPortesN, "T")
    C = "select sum(cantidad*pesoarti),sum(importel) from slialb,sartic " & C
    
    
    'El enlzace
    
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        SumaKilosLineas = DBLet(miRsAux.Fields(0), "N")
        ImporteL = DBLet(miRsAux.Fields(1), "N")
    End If
    miRsAux.Close
    
    
    'Fijo la zona y la ruta del cliente
    
    RutaCliente = -1
    ZonaCliente = -1
    C = "Select codzonas,codrutas from sclien where codclien = " & Val(Text1(4).Text)
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ZonaCliente = DBLet(miRsAux!codzonas, "N")
        RutaCliente = DBLet(miRsAux!codrutas, "N")
    End If
    miRsAux.Close
    
    
ESumaKilosLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function

'Si devuelve cero nada
'si devuelve >0 marcara la linea de portes
Private Function HacerAccionesPortes() As Integer
Dim ImporteLineas As Currency
Dim KilosAhora As Currency
Dim C As String
Dim CodEnvio As Integer
Dim PrecioKilo As Currency
Dim DtoPorte As Currency
Dim DesdeKilo As Currency
Dim ImporteL_Portes As Currency
Dim codCCoste As String 'centro de coste

    HacerAccionesPortes = 0
    KilosAhora = SumaKilosLineas(ImporteLineas)
    
    
    ' Si no cambia los kilos me salgo
    '-----------------------------------------------
    'If KilosAhora = KilosAnteriores Then Exit Function
    If data2.Recordset.EOF Then Exit Function
    
    If MsgBox("Desea recalcular los portes?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    
    
    Set miRsAux = New ADODB.Recordset
    
    
    If ZonaCliente > 0 Then
        'Ha encontrado la zona /ruta. Miro en sportes
        C = "select sporte.codenvio,nomenvio,PrecioKg,desdekgs from sporte,senvio where sporte.codenvio=senvio.codenvio "
        C = C & " AND codcentr = " & ZonaCliente
        'Los kilos  hastakgs
        C = C & " AND desdekgs <= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " AND hastakgs >= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " group by sporte.codenvio"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumRegElim = 0
        CodEnvio = 0
        If Not miRsAux.EOF Then
            'Por si acaso hay mas de uno
            CadenaDesdeOtroForm = ""
            While Not miRsAux.EOF
                CodEnvio = miRsAux!CodEnvio
                PrecioKilo = miRsAux!preciokg
                DesdeKilo = DBLet(miRsAux!DesdeKgs, "N")
                NumRegElim = NumRegElim + 1
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux!CodEnvio & "<" & miRsAux!nomEnvio & "<" & miRsAux!preciokg & "<" & DBLet(miRsAux!DesdeKgs, "N") & "|"
                miRsAux.MoveNext
            Wend
            
            
            If NumRegElim > 1 Then
                'Mostraremos un form para que seleccione la opcion correspondiente
                frmVarios.Opcion = 3
                frmVarios.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    C = RecuperaValor(CadenaDesdeOtroForm, 1)
                    CodEnvio = Val(C)
                    
                    C = RecuperaValor(CadenaDesdeOtroForm, 3)
                    PrecioKilo = CCur(C)
                    
                    DesdeKilo = CCur(RecuperaValor(CadenaDesdeOtroForm, 4))
                End If
            Else
                    CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, "<", "|")
                    CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
            
            End If
            
        End If
        miRsAux.Close
        
        
        'Dto en portes
        DtoPorte = 0
        ImporteL_Portes = 0
        If RutaCliente = 1 Or RutaCliente = 3 Or RutaCliente = 4 Then DtoPorte = vParamAplic.AbonoKilos
        If RutaCliente = 1 Or RutaCliente = 2 Then PrecioKilo = 0
        If RutaCliente = 4 And ImporteLineas < vParamAplic.ImporteMinimo Then 'importe pedido menor que importe minimo todo a cero(preciokilo, dtokilo)
               PrecioKilo = 0
               DtoPorte = 0
               ImporteL_Portes = 0
        Else
            If RutaCliente = 4 Then ImporteL_Portes = PrecioKilo
        End If
        
        If DesdeKilo = 1 Then
            If RutaCliente <> 4 Then
                ImporteL_Portes = PrecioKilo
                KilosAhora = 1
            End If
        Else
            ImporteL_Portes = (PrecioKilo - DtoPorte) * KilosAhora
        End If
        If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 'masl 090709
        
        'Ahora compruebo si tiene la linea de portes para aplicarle el importe
        C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        C = "Select numlinea,codccost from slialb " & C & " and codartic ='" & vParamAplic.ArtPortesN & "'"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux!numlinea, "N")
            codCCoste = DBLet(miRsAux!CodCCost, "T")
            
        Else
            codCCoste = ""
            If vEmpresa.TieneAnalitica Then
                If vParamAplic.ModoAnalitica = 0 Then
                    'Del trabajador
                    codCCoste = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Text1(3).Text)
                    If codCCoste = "" Then MsgBox "Trabajador sin centro de coste", vbExclamation
                        
                End If
                If vParamAplic.ModoAnalitica <> 0 Or codCCoste = "" Then
                    codCCoste = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND codartic", vParamAplic.ArtPortesN, "T")
                    If codCCoste = "" Then MsgBox "Familia sin centro de coste", vbExclamation
                End If
            End If
                
        End If
        miRsAux.Close
        
        
        'SI ya existe la borro, para que siempre aparezca al final
        If NumRegElim > 0 Then
            C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            C = C & " AND numlinea = " & NumRegElim
            C = "DELETE FROM  slialb " & C
            conn.Execute C
            Espera 0.1
            
        
        End If
        
     'If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 masl 090709
        
        
            'Si el precio es mayor k cero entonces SI pongo la linea
            C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            C = "select codtipom,codalmac,max(numlinea) from slialb " & C
            C = C & " GROUP BY codalmac"
            miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "NO deberia haberse producido", vbExclamation
                Exit Function
            End If
            NumRegElim = miRsAux.Fields(2) + 1
            HacerAccionesPortes = NumRegElim
    '            SQL = "INSERT INTO " & NomTablaLineas
    '            SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex) "
    '            SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", "
            
            C = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtPortesN, "T")
            C = DevNombreSQL(C)
            C = miRsAux!codAlmac & ",'" & vParamAplic.ArtPortesN & "','" & C & "','"
            
            C = "INSERT INTO " & NomTablaLineas & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex,codccost)" & _
                "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", " & C
            
            
            'Amplicacion
            C = C & CadenaDesdeOtroForm & "',"
            
            
            If RutaCliente <> 1 And RutaCliente <> 3 And RutaCliente <> 4 Then  'masl 090709
                        'Cantidad, precioar dto1 dto2
                        C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(PrecioKilo))
                        C = C & "," & TransformaComasPuntos(CStr(DtoPorte)) & ",0,"
                                           
            Else
                'masl 090709
                'REctificado Marzo 2011
                        'C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(DtoPorte * (-1)))
                        'pintaba mal el preciolinea
                If PrecioKilo - DtoPorte < 0 Then
                    C = C & TransformaComasPuntos(CStr(KilosAhora)) & ",0,0,0,"
                Else
                    C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(PrecioKilo - DtoPorte))
                    C = C & ",0" & ",0,"
                End If
            End If  'masl 090709
            
            'importel
            C = C & TransformaComasPuntos(CStr(ImporteL_Portes))
            
            'origpre,codprovex,codccost
            C = C & ",'M',0," & DBSet(codCCoste, "T") & ")"
        
        
        
            'Noviembre 2009.    Enero 2010.  SIEMPRE hay que meter la linea de portes
            'If ImporteL_Portes <> 0 Then conn.Execute C
            conn.Execute C
        
    End If
            
End Function


'Para obtener los dtos por cantidad lo que hace es a partir de un
'subtring del articulo(poscion 3 a 6) va a sdesca con la suma de la cantidad
'si en sdesca y dentro de los desde /hasta cantidad encuentra un dto lo aplica


Private Sub DescuentosCantidad(Articulo As String)
Dim cad As String
Dim R As ADODB.Recordset
Dim NuevoDto As Currency
Dim Importe As Currency
Dim bAct As Boolean

    On Error GoTo EDescuentosCantidad
    
    If Not vParamAplic.DtoxCantidad Then Exit Sub ' ---- [14/09/2009] (LAURA)
     
    If MsgBox("¿Desea recalcular los descuentos por cantidad?", vbQuestion + vbYesNo) = vbYes Then    'masl 140909
    
        
        'Si no  tenemos portes, ni nos pasamos
    '    If vParamAplic.ArtPortes = "" Then Exit Sub
        
        
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from slialb " & cad
        If Articulo <> "" Then cad = cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        'Y origen PRECIO no es precio especial
        cad = cad & " AND origpre <> 'E'"
        cad = cad & " group by 1"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                cad = "select * from sdesca where desdecan <=" & cad & " and " & cad & " <= hastacan and envagran = '"
                cad = cad & miRsAux!raiz & "'"
                R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                cad = ""
                If Not R.EOF Then cad = R!dtolinea
                R.Close
                
                
                If cad <> "" Then
                    'OK tiene nuevo descuento
                    NuevoDto = CCur(cad)
                    
                    'Cojo los articulos del albaran y le meto el dto
                    cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                    cad = "select * from slialb " & cad
                    '                                 a partir de la 3era posicion
                    cad = cad & " AND codartic like '__" & miRsAux!raiz & "%'"
                    R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not R.EOF
                        '-- comprobar si admite descuento
                        If R!origpre = "T" Then
                            cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            cad = DevuelveDesdeBDNew(conAri, "slista", "dtopermi", "codartic", R!codArtic, "T", , "codlista", cad, "N")
                            bAct = (cad = "1")
                        ElseIf R!origpre = "A" Or R!origpre = "M" Then
                            bAct = True
                        Else
                            bAct = False
                        End If
                        
                        If bAct Then
                            cad = CalcularImporte(CStr(R!cantidad), CStr(R!precioar), CStr(NuevoDto), CStr(R!dtoline2), vParamAplic.TipoDtos)
                            Importe = CCur(cad)
                            cad = "UPDATE slialb set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            cad = cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            cad = cad & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                            cad = cad & " and numlinea = " & R!numlinea
                            conn.Execute cad
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
    
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub





Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    

    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
    frmB.vTabla = "cabccost"
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Centros de coste"
    frmB.vselElem = 0
    frmB.vConexionGrid = conConta
    
    frmB.Show vbModal
    Set frmB = Nothing

End Sub



Private Sub PosicionarData2(QueGRid As Byte)
    On Error GoTo EPosicionarData2
    
    If QueGRid = 1 Then
        data2.Recordset.Find "numlinea = " & NumRegElim
        If data2.Recordset.EOF Then data2.Recordset.MoveFirst
    Else
        data3.Recordset.Find "matricula = " & DBSet(txtAux2(0).Text, "T")
        If data3.Recordset.EOF Then data3.Recordset.MoveFirst
    End If
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub



Private Sub PonerUltAlmacen()
Dim C As String
    If AlmacenLineas < 0 Then
       If Not data2.Recordset.EOF Then
            C = ObtenerWhereCP(True)
            C = Replace(C, NombreTabla, NomTablaLineas)
            AlmacenLineas = DevuelveUltimoAlmacen(NomTablaLineas, C)
       End If
            
       If AlmacenLineas < 0 Then
            'No hay datos todavia
            '                                                                trabajador
            C = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
            If C <> "" Then AlmacenLineas = Val(C)
        End If
    End If
End Sub



Private Sub UpdateaNomDirec()
Dim N As Integer
Dim Ol As Integer
Dim C As String

    N = -1
    If Not IsNull(Data1.Recordset!CodDirec) Then N = Data1.Recordset!CodDirec
    
    Ol = -1
    If Text1(12).Text <> "" Then Ol = CInt(Text1(12).Text)
    
    If N <> Ol Then
        If Ol < 0 Then
            C = "NULL"
        Else
            C = DBSet(Text2(12).Text, "T")
        End If
        C = "UPDATE scaalb set nomdirec=" & C
        C = C & " WHERE codtipom = '" & Text1(30).Text & "' AND numalbar=" & Text1(0).Text
        ejecutar C, False
    End If
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera = 0
        imgBuscar_Click Index
        
    Else
        'Lineas
        
        cmdAux_Click Index
        
        
    End If
        
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
        If EsCabecera = 1 Then 'Llama desde VerTodos del Form
            Text1(12).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
        Else
            'DESDE ENVIO
            Text1(42).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
            Text2(42).Text = RecuperaValor(CadenaSeleccion, 2)
        End If
End Sub

Private Sub LanzaBusquedaDpto(Departamento As Boolean, Indice As Integer)

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


Private Sub ComprobarCambioPrecioDtoyVtaBajoPrecio()
Dim vPreFact As CPreciosFact
Dim SQ As String
Dim Impo As Currency
Dim Particular As Boolean
Dim Cajas As String
Dim Dto1 As String
Dim Dto2 As String
Dim DeVarios As Boolean
Dim Otro As Currency
Dim Comi_ As String
Dim Insertando_o_cambiandoArticulo As Boolean
Dim TipoComisionVarios As String

    On Error GoTo EComprobarCambioPrecioDto


    
  

    'Si es articulo de varios
    'Eso lo sabemos PQ el txtaux(2) NO esta locked

    'Al modificar puede ser que no haya pasado por codartic
    Cajas = "unicajas"
    SQ = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", txtAux(1).Text, "T", Cajas)

    DeVarios = False
    If SQ = "1" Then DeVarios = True
    
 
 
    
    
    If Not DeVarios Then
    
                SQ = DevuelveDesdeBD(conAri, "particular", "sclien", "codclien", Text1(4).Text)
                Particular = SQ = "1"
                
                    
                Insertando_o_cambiandoArticulo = True
                If Insertando_o_cambiandoArticulo Then
                
                    'ESTAMOS INSERTANDO
                    If Me.txtAux(5).Text = "M" Then
                        'seguro que ha cambiado el precio
                        GrabaLogCambioPrecioDto = True
                    Else
                    
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
                            Set vPreFact = New CPreciosFact
                            vPreFact.CodigoClien = Text1(4).Text
                            vPreFact.FijarTarifaActividad
                            vPreFact.CodigoArtic = txtAux(1).Text
                            If Val(Cajas) > 1 Then
                                Impo = Val(CCur(txtAux(3).Text)) - Val(Cajas)
                                If Impo >= 0 Then Cajas = ""
                            End If
                            
                            'Septiembre 2014
                            Comi_ = ""
                            SQ = vPreFact.ObtenerPrecio(Cajas = "", Text1(1).Text, "", Comi_)
                            
                
                            
                            Impo = ImporteFormateado(txtAux(6).Text)
                            If Impo <> CCur(vPreFact.Descuento1) Then
                                GrabaLogCambioPrecioDto = True
                            Else
                                Impo = ImporteFormateado(txtAux(7).Text)
                                If Impo <> CCur(vPreFact.Descuento2) Then GrabaLogCambioPrecioDto = True
                            End If
                            
                        End If
                        
                    End If
                    
                    
                    
                    
                    
                    
                Else
                    'MODIFICANDO
                    'Si ha cambiado el precio,dto1 o dto
                    
                    
                    
                    Impo = ImporteFormateado(txtAux(4).Text)
                    If Impo <> CCur(data2.Recordset!precioar) Then
                        GrabaLogCambioPrecioDto = True
                    Else
                        Impo = ImporteFormateado(txtAux(6).Text)
                        If Impo <> CCur(data2.Recordset!dtoline1) Then
                            GrabaLogCambioPrecioDto = True
                        Else
                            Impo = ImporteFormateado(txtAux(7).Text)
                            If Impo <> CCur(data2.Recordset!dtoline2) Then GrabaLogCambioPrecioDto = True
                        End If
                    End If
                End If
                
    End If
    
    
    
    If vParamAplic.GrabaModificarPrecioAlaBaja Then
            VendeAMenorPrecio = 0
            'Vera el importe calculado y si es inferior dará error
            If DeVarios Then
                
                SQ = DevuelveDesdeBD(conAri, "TipoComiArtVario", "sartic", "codartic", txtAux(1).Text, "T")
                If SQ = "" Then SQ = "0"
                
                VendeAMenorPrecio = CByte(SQ)
                
            Else
                If Me.txtAux(5).Text = "E" Then
                    'precio especial SIEMPRE
                    VendeAMenorPrecio = 1
                Else
                    
                    
                    'OCTUBRE 2014
                    'Para los abonos, rectificativas.
                    'Comparamos con el inmportel absoluto
                    Impo = Abs(ImporteFormateado(txtAux(8).Text))

                    
                    '------------------------------------------
                    If True Then
                        If vPreFact Is Nothing Then
                            Set vPreFact = New CPreciosFact
                                        
                            vPreFact.CodigoClien = Text1(4).Text
                            
                            vPreFact.FijarTarifaActividad
                            vPreFact.CodigoArtic = txtAux(1).Text
                            If Val(Cajas) > 1 Then
                                Otro = Val(CCur(txtAux(3).Text)) - Val(Cajas)
                                If Otro >= 0 Then Cajas = ""
                            End If
                           
                
                    
                        End If
                        SQ = vPreFact.ObtenerPrecioDtoFamilia(Cajas = "", Text1(1).Text, "")
                        SQ = CalcularImporte(txtAux(3).Text, SQ, vPreFact.Descuento1, vPreFact.Descuento2, vParamAplic.TipoDtos)
                        SQ = Abs(SQ)
                        
                        'Vende por debajo precio
                        
                        
                        If ImporteFormateado(txtAux(3).Text) < 0 Then
                            'ABONO
                            If CCur(SQ) < Impo Then VendeAMenorPrecio = 1
                        Else
                            If CCur(SQ) > Impo Then VendeAMenorPrecio = 1
                        End If
                    End If
                End If
            End If
    End If
    
    If vParamAplic.NumeroInstalacion = 2 Then
        If DeVarios Then
            If VendeAMenorPrecio = 1 Then
                Comi_ = vAgent.ComsionEco
            Else
                Comi_ = vAgent.ComsionNormal
            End If
        Else

            
            If Comi_ = "" Then
                'No hay comision especial, con lo cual
                If ComisionCliente <> 0 Then Comi_ = ComisionCliente
            Else
                'la mas baja de todas
                If ComisionCliente > 0 Then
                    If CCur(Comi_) > ComisionCliente Then Comi_ = ComisionCliente
                End If

                
            End If
            
            If Comi_ <> "" Then
                VendeAMenorPrecio = 2 'comision supereco
               
            Else
                'Vemos el precio minimo del articulo
                SQ = DevuelveDesdeBD(conAri, "preciominvta", "sartic", "codartic", vPreFact.CodigoArtic, "T")
                If SQ <> "" Then
                    'Si tiene lo comparamos con lo que ha puesto
                    SQ = CalcularImporte(txtAux(3).Text, SQ, vPreFact.Descuento1, vPreFact.Descuento2, vParamAplic.TipoDtos)
                    SQ = Abs(SQ)
                    If ImporteFormateado(txtAux(3).Text) < 0 Then
                        'ABONO
                        If CCur(SQ) < Impo Then VendeAMenorPrecio = 2
                    Else
                        If CCur(SQ) > Impo Then VendeAMenorPrecio = 2
                    End If
                End If
                    
                If VendeAMenorPrecio = 2 Then
                    
                    Comi_ = vAgent.ComsionPVPMin
                ElseIf VendeAMenorPrecio = 1 Then
                    Comi_ = vAgent.ComsionEco
                Else
                    Comi_ = vAgent.ComsionNormal
                End If
                
            End If
        End If
        
        If txtAux(1).Text = vParamAplic.ArtReciclado Then Comi_ = ""
        If txtAux(1).Text = vParamAplic.ArtPortesN Then Comi_ = ""
        
        
        If vParamAplic.PtosArticuloCanje <> "" Then
            If vParamAplic.PtosArticuloCanje = txtAux(1).Text Then Comi_ = txtAux(12).Text 'DEJO la que ya hemos puesto
        End If
        txtAux(12).Text = Comi_
    End If
    Set vPreFact = Nothing
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
        frmListado3.OtrosDatos = "Nueva"
    Else
        frmListado3.OtrosDatos = "Modificar"
    End If
    frmListado3.OtrosDatos = frmListado3.OtrosDatos & " Alb " & Text1(0).Text & Text1(1).Text & " Articulo " & txtAux(1).Text
    
    'para que herbelca NO vea la comision
    txtAux(12).visible = False
    frmListado3.Show vbModal
    txtAux(12).visible = True
    
    Screen.MousePointer = Rc
    
    
End Sub

'Si MODO=3 , es decir INSERTANDO, sacara un mensaje diciendo los datos del riesgo
Private Function Riesgo() As Boolean
Dim ImpAlb As Currency, ImpTesor As Currency
Dim miSQL As String
Dim ImportePedido As Currency
Dim Aux As String

    Riesgo = True
    Set miRsAux = New ADODB.Recordset
    '              *^**********                ponia credisol  para todo
    miSQL = "Select codclien,tipoiva,credipriv,if(limcredi is null,0,limcredi) limcredi from sclien where codclien =" & Text1(4).Text
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOG
    miSQL = "9"
    If Not IsNull(miRsAux!credipriv) Then miSQL = miRsAux!credipriv
    
    If Val(miSQL) < 9 Then
        
        
        RiesgoCliente miRsAux!codClien, CByte(miRsAux!TipoIVA), Now, ImpTesor, ImpAlb, Nothing, 60
        Aux = Format(miRsAux!limcredi, FormatoImporte)
        If Len(Aux) < 9 Then Aux = "   " & Aux
        miSQL = "Crédito concedido:        " & Aux & vbCrLf
        
        Aux = Format(ImpTesor, FormatoImporte)
        If Len(Aux) < 9 Then Aux = Space(9 - Len(Aux)) & Aux
        miSQL = miSQL & "Tesorería: " & Right(Space(30) & Aux, 30) & vbCrLf
        
        Aux = Format(ImpAlb, FormatoImporte)
        If Len(Aux) < 9 Then Aux = Space(9 - Len(Aux)) & Aux
        miSQL = miSQL & "Albaranes: " & Right(Space(30) & Aux, 30) & vbCrLf

        
        ImpTesor = ImpTesor + ImpAlb
        
        
        
        If Modo = 3 Then
             'Disponible
                
             If ImpTesor > miRsAux!limcredi Then
                'NO deberia haber entrado aqui
                miSQL = miSQL & vbCrLf & "** EXCEDE CREDITO CONCEDIDO **"
             Else
                ImpTesor = miRsAux!limcredi - ImpTesor
                Aux = Format(ImpTesor, FormatoImporte)
                If Len(Aux) < 9 Then Aux = Space(9 - Len(Aux)) & Aux
                miSQL = miSQL & vbCrLf & vbCrLf & "DISPONIBLE: " & Right(Space(30) & Aux, 30) & vbCrLf
             End If
                
             MsgBox miSQL, vbInformation
        
        Else
            'Pasando a cabecera. Comprobara que no se ha sobrepasado el limite de credito
            
            
            'Tesoreria + albaranes + este albaran.....
            'ImpTesor = ImpTesor + ImportePedido
            'miSQL = miSQL & "Pedido:        " & Format(ImportePedido, FormatoImporte) & vbCrLf
            
            If ImpTesor > miRsAux!limcredi Then
                miSQL = miSQL & vbCrLf & "** EXCEDE CREDITO CONCEDIDO **" & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(miSQL, vbQuestion + vbYesNo + vbMsgBoxRight) = vbNo Then
                    Riesgo = False
                Else
                    'Metemos el LOG
                    miSQL = "Cliente: " & Text1(4).Text & " - " & Text1(5).Text & vbCrLf
                    miSQL = miSQL & "Albarán: " & Text1(30).Text & "-" & Text1(0).Text & " de " & Text1(1).Text & vbCrLf
                    miSQL = miSQL & "Importe TOTAL albaran: " & Text3(55).Text

                    Set LOG = New cLOG
                    ' 17 Venta a sabiendas riesgo
                    LOG.Insertar 17, vUsu, miSQL
                    Set LOG = Nothing
            
                    
                End If
                
            End If
            'Actualziamos riesgo
            'Febrero2018
            'ActualizaRiesgoCliente CLng(Text1(4).Text)
        End If

        

    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
End Function





'Campos ALZIRA MOIXENT
Private Sub MultiInsercionCampos()
Dim i As Integer
Dim VariedadPartida As String

        'Quito el indicador # de multi campo
        If InStr(1, CadenaDesdeOtroForm, 1) > 0 Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)



        cadList = cadList & "(select codcampo from slialbcampos where numalbar=" & Data1.Recordset!NumAlbar
        cadList = cadList & " AND "


        cadList = "codtipom = " & DBSet(Data1.Recordset!codtipom, "T")
        cadList = cadList & " AND numalbar"
        cadList = DevuelveDesdeBD(conAri, "max(numlinea)", "slialbcampos", cadList, CStr(Data1.Recordset!NumAlbar), "N")
        NumRegElim = 0
        If cadList <> "" Then NumRegElim = Val(cadList)
        NumRegElim = NumRegElim + 1
        motivo = ""
        While CadenaDesdeOtroForm <> ""
            i = InStr(1, CadenaDesdeOtroForm, "·#")

            If i = 0 Then
                CadenaDesdeOtroForm = ""
            Else
                cadList = Mid(CadenaDesdeOtroForm, 1, i - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 2)
                
                VariedadPartida = "," & DBSet(RecuperaValor(cadList, 2), "T", "S") & "," & DBSet(RecuperaValor(cadList, 3), "T", "S")
                cadList = RecuperaValor(cadList, 1) 'cdocampo

                For i = 1 To Me.ListView1.ListItems.Count
                    'Si no lo ha insertado YA
                    If Val(Me.ListView1.ListItems(i).Text) = Val(cadList) Then
                        cadList = ""
                        Exit For
                    End If

                Next i

                If cadList <> "" Then

                        '  '
                        motivo = motivo & ", (" & Data1.Recordset!NumAlbar & ","
                        motivo = motivo & DBSet(Data1.Recordset!codtipom, "T") & "," & NumRegElim & "," & cadList
                        motivo = motivo & VariedadPartida & ")"
                        NumRegElim = NumRegElim + 1
                End If
            End If
        Wend
        If motivo <> "" Then
            motivo = Mid(motivo, 2) 'quito la primera coma
            '
            motivo = "INSERT INTO slialbcampos(numalbar,codtipom,numlinea,codcampo,nomvarie,nompartida) VALUES " & motivo
            If ejecutar(motivo, False) Then
                'Hay que refrescar y boton anyadir

            End If
        End If

        cadList = ""
        motivo = ""

        '
        
End Sub


Private Sub CargaDatosCampos()
Dim IT
    'Para no meter MUCHOS ariagro.ss
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
'    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie"
'    SQL = SQL & " from (@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
'    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie"
'    'where socio
'    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
'
    
    
    
    cadList = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.codsitua"
    cadList = cadList & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    cadList = cadList & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    cadList = cadList & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    
    cadList = Replace(cadList, "@#", vParamAplic.Ariagro & ".")
    
    cadList = cadList & " WHERE codcampo IN "
    cadList = cadList & "(select codcampo from slialbcampos where numalbar=" & Data1.Recordset!NumAlbar
    cadList = cadList & " AND codtipom = " & DBSet(Data1.Recordset!codtipom, "T")
    cadList = cadList & ")"
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadList, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = Format(miRsAux!codCampo, "000000")
        IT.SubItems(1) = DBLet(miRsAux!nomparti, "T")
        IT.SubItems(2) = DBLet(miRsAux!nomvarie, "T")
        IT.Tag = miRsAux!codCampo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cadList = ""
  
End Sub



Private Sub ImpresionAlbaranRutaCliente()
    
    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    frmListado4.vCadena = Text1(4).Text
    frmListado4.Opcion = 14
    frmListado4.Show vbModal
    
End Sub


Private Function GeneraAlbaranDesdeDevolucion() As Boolean


Dim vT As CTiposMov

    On Error GoTo eGeneraAlbaranDesdeDevolucion
    Set miRsAux = New ADODB.Recordset
    
    GeneraAlbaranDesdeDevolucion = False
    

    'Primera comprobacion. Que todos los articulos estan en status NORMAL
    ' y todas las cantidades en negativo
    txtAnterior = ""
    
    BuscaChekc = "select slialb.*,codstatu from slialb,sartic where slialb.codartic = sartic.codartic  and slialb.codtipom='DEV' AND numalbar = " & Text1(0).Text & " ORDER BY numlinea desc"
    miRsAux.Open BuscaChekc, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
        If miRsAux!codstatu > 0 Then
            BuscaChekc = IIf(miRsAux!codstatu = 1, "OBSOLETO", "CADUCADO")
            BuscaChekc = BuscaChekc & " -> " & miRsAux!codArtic & " " & miRsAux!NomArtic
            txtAnterior = BuscaChekc & vbCrLf & txtAnterior
        End If
    
        If miRsAux!cantidad > 0 Then
            MsgBox "Cantidad debe ser positiva:" & miRsAux!NomArtic, vbExclamation
            GoTo eGeneraAlbaranDesdeDevolucion
            
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If txtAnterior <> "" Then
        txtAnterior = txtAnterior & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
        If MsgBox(txtAnterior, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then GoTo eGeneraAlbaranDesdeDevolucion
    End If
    
    'OK. Lanzamos proceso
    'Preguntaremos si va a venta o factura rectificativa
    CadenaDesdeOtroForm = ""
    frmListado5.OpcionListado = 17
    frmListado5.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
    
        BuscaChekc = CadenaDesdeOtroForm
        Set vT = New CTiposMov
        If Not vT.Leer(BuscaChekc) Then Err.Raise 513, , "Obtener contador " & BuscaChekc
        vT.ConseguirContador vT.TipoMovimiento
        vT.IncrementarContador vT.TipoMovimiento
        
        'Ahora generaremos todo el proceso
        Screen.MousePointer = vbHourglass
        conn.BeginTrans
        If ProcesoDevolucion(vT) Then
            conn.CommitTrans
            GeneraAlbaranDesdeDevolucion = True
            
            'Ahora cargaremos el alvaran de venta o la rectificativa
            hcoCodMovim = vT.Contador
            hcoCodTipoM = vT.TipoMovimiento
            CodTipoMov = hcoCodTipoM
            
            'Ponemos el menu como estaba
            PuntosMenusQuitadosPorDEV False
            DoEvents
            
            'Cargamos los valores de la fra creada
            txtAnterior = "codtipom=" & DBSet(hcoCodTipoM, "T") & " AND numalbar =" & hcoCodMovim
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & txtAnterior & " " & Ordenacion
            PonerCadenaBusqueda
        Else
            conn.RollbackTrans
            'Devolvemos el contador
            vT.DevolverContador vT.TipoMovimiento, vT.Contador
            
        End If
    End If
    
eGeneraAlbaranDesdeDevolucion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
    Set miRsAux = Nothing
    Set vT = Nothing
    BuscaChekc = ""
    txtAnterior = ""
    Screen.MousePointer = vbDefault
End Function


Private Function ProcesoDevolucion(Tip As CTiposMov) As Boolean
Dim vCStock As CStock
    On Error GoTo eProcesoDevolucion
    
    ProcesoDevolucion = False
    
    
    
    
    
    'Metemos la cabecera del albaran. Por si da error, que se salga
    'A falta de `codtipom`,`numalbar`
    BuscaChekc = ",`fechaalb`,`factursn`,`codclien`,`nomclien`,`domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`,"
    BuscaChekc = BuscaChekc & "`telclien`,`coddirec`,`nomdirec`,`referenc`,`facturkm`,`cantidkm`,`codtraba`,`codtrab1`,`codtrab2`,"
    BuscaChekc = BuscaChekc & "`codagent`,`codforpa`,`codenvio`,`dtoppago`,`dtognral`,`tipofact`,`observa01`,`observa02`,`observa03`,"
    BuscaChekc = BuscaChekc & "`observa04`,`observa05`,`numofert`,`fecofert`,`numpedcl`,`fecpedcl`,`fecentre`,`sementre`,"
    BuscaChekc = BuscaChekc & "`codtipmf`,`numfactu`,`fecfactu`,`esticket`,`numtermi`,`numventa`,`aportacion`,`pesoalba`,"
    BuscaChekc = BuscaChekc & "`portes`,`fecenvio`,`docarchiv`,`tipliquid`,`actuacion`,`tipoimp`,`origdat`,`coddiren`,"
    BuscaChekc = BuscaChekc & "`tipAlbaran`,`albImpreso`,`codzonas`,`observacrm`"

    
    txtAnterior = "INSERT INTO scaalb(`codtipom`,`numalbar`" & BuscaChekc & ") SELECT "
    txtAnterior = txtAnterior & DBSet(Tip.TipoMovimiento, "T") & "," & Tip.Contador & BuscaChekc
    txtAnterior = txtAnterior & " FROM scaalb WHERE codtipom='DEV' AND numalbar = " & Text1(0).Text
    conn.Execute txtAnterior
    
    
    BuscaChekc = "select slialb.*,codstatu from slialb,sartic where slialb.codartic = sartic.codartic  and slialb.codtipom='DEV' AND numalbar = " & Text1(0).Text & " ORDER BY numlinea "
    miRsAux.Open BuscaChekc, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Set vCStock = New CStock
    txtAnterior = ""
    
    vCStock.tipoMov = "S"  'SALIDA en negativo
    vCStock.DetaMov = Tip.TipoMovimiento
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente del albaran
    vCStock.Documento = Tip.Contador
    vCStock.FechaMov = Text1(1).Text 'Fecha del Albaran
    vCStock.FechaMov = Text1(1).Text & " " & Format(Now, "hh:mm:ss")
    
    While Not miRsAux.EOF
        
        vCStock.codArtic = miRsAux!codArtic
        vCStock.codAlmac = miRsAux!codAlmac
        vCStock.cantidad = miRsAux!cantidad
        vCStock.Importe = miRsAux!ImporteL
        vCStock.LineaDocu = miRsAux!numlinea
    
        If Not vCStock.ActualizarStock(False, True) Then Err.Raise 513, , "Actualizadno stock almacen " & vCStock.codAlmac & "  art: " & vCStock.codArtic
        
        '(codtipom, numalbar,numlinea, codalmac,
        BuscaChekc = ", ('" & Tip.TipoMovimiento & "', " & Tip.Contador & ", " & miRsAux!numlinea & ", " & vCStock.codAlmac & ","
        'codartic, nomartic, ampliaci,
        BuscaChekc = BuscaChekc & DBSet(vCStock.codArtic, "T") & ", " & DBSet(miRsAux!NomArtic, "T") & ", " & DBSet(miRsAux!Ampliaci, "T") & ", "
        'cantidad,numbultos,precioar,
        BuscaChekc = BuscaChekc & DBSet(vCStock.cantidad, "N") & ", " & DBSet(miRsAux!NumBultos, "N") & ", " & DBSet(miRsAux!precioar, "N") & ", "
        'dtoline1, dtoline2, importel, origpre,
        BuscaChekc = BuscaChekc & DBSet(miRsAux!dtoline1, "N") & "," & DBSet(miRsAux!dtoline2, "N") & ", " & DBSet(miRsAux!ImporteL, "N") & "," & DBSet(miRsAux!origpre, "T") & ","
        
        'codprovex,numlote,codccost,pvpInferior,comisionagente) "
        If vEmpresa.TieneAnalitica Then
            '- codprove,numlote,codccost
            BuscaChekc = BuscaChekc & "0," & DBSet(miRsAux!numLote, "T", "S") & "," & DBSet(miRsAux!CodCCost, "T", "S")
        Else
            '- codprove,numlote,codccost
            BuscaChekc = BuscaChekc & DBSet(miRsAux!codProvex, "N", "N") & "," & DBSet(miRsAux!numLote, "T", "S") & "," & ValorNulo
        End If
        BuscaChekc = BuscaChekc & "," & miRsAux!PVPInferior & ","
        
        If vParamAplic.NumeroInstalacion = 2 Then
            BuscaChekc = BuscaChekc & DBSet(miRsAux!comisionagente, "N")
        Else
            BuscaChekc = BuscaChekc & "null"
        End If
        BuscaChekc = BuscaChekc & ")"
        txtAnterior = txtAnterior & BuscaChekc
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'Ahora crearemos la cabecera antes de todo
    BuscaChekc = "INSERT INTO slialb (codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, "
    BuscaChekc = BuscaChekc & "cantidad,numbultos,precioar, dtoline1, dtoline2, importel, origpre,codprovex,numlote,"
    BuscaChekc = BuscaChekc & "codccost,pvpInferior,comisionagente) VALUES "
    
    BuscaChekc = BuscaChekc & Mid(txtAnterior, 2)  'quitando la primera coma
    conn.Execute BuscaChekc
    
    
    
    BuscaChekc = "DELETE FROM @@ WHERE codtipom='DEV' AND numalbar = " & Text1(0).Text
    
    conn.Execute Replace(BuscaChekc, "@@", "slialb")
    If vParamAplic.Ariagro <> "" Then conn.Execute Replace(BuscaChekc, "@@", "slialbcampos")
    If vParamAplic.ManipuladorFitosanitarios2 Then conn.Execute Replace(BuscaChekc, "@@", "slialblotes")
    conn.Execute Replace(BuscaChekc, "@@", "scaalb")
    
    
    
    ProcesoDevolucion = True
    
    
    
eProcesoDevolucion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set vCStock = Nothing
End Function








'----------------------------------------------------------------------------------------------
'Lotes fitosantiarios
' Los pide, O no segun hay muchos o pocos
Private Function DatosLotesFitosOk(ByRef ArticuloConFitosantiarios As Boolean) As Boolean
Dim SQL As String
Dim CadenaInsertTmpLotes  As String
Dim NumerloLote As String
Dim CuantosLotesDistintos As Integer
Dim CantidadEnTotal As Currency
Dim B As Boolean

    
    ArticuloConFitosantiarios = False
    SQL = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", txtAux(1).Text, "T")
    If SQL = "" Then
        'OK. Salimos YA
        DatosLotesFitosOk = True
        Exit Function
    End If
    B = False
    DatosLotesFitosOk = B
    
    
    
    'Si llega aqui, y no tiene manipulador de Fitosantarios
    If Trim(Text1(45).Text) = "" Then
        
        'Esto sera para el CHOLI , en Navarrres
        SQL = DevuelveDesdeBD(conAri, "ManipuladorNumCarnet", "sclien", "codclien", Text1(4).Text)
        If SQL = "" Then
            'Veo si tiene autirzados
            SQL = DevuelveDesdeBD(conAri, "numcarnet", "sclienmani", "codclien", Text1(4).Text)
        End If
        
        If SQL <> "" Then
            'Llamamos al manipulador de carnet fitosnaitarios
            CadenaDesdeOtroForm = ""
            frmFitoCarnet.Cliente = Val(Text1(4).Text)
            frmFitoCarnet.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                
                SQL = "UPDATE scaalb SET ManipuladorNumCarnet = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "T")
                SQL = SQL & ",ManipuladorFecCaducidad =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T")
                SQL = SQL & ",ManipuladorNombre = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
                SQL = SQL & ", TipoCarnet =" & IIf(UCase(RecuperaValor(CadenaDesdeOtroForm, 4)) = "CUALIFICADO", 2, 1)
                SQL = SQL & ObtenerWhereCP(True)
                If ejecutar(SQL, False) Then
                    ' ManipuladorFecCaducidad  ManipuladorNombre  TipoCarnet
                    Me.Text1(45).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                    Me.Text1(46).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
                    Me.Text1(47).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
                    Text2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
                    'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
                    Me.Text1(48).Text = IIf(UCase(Text2(0).Text) = "CUALIFICADO", 2, 1)
                End If
                
            End If
        End If

    

        If Me.Text1(45).Text = "" Then
            MsgBox "No tiene carnet de manipulador", vbExclamation
            If hcoCodTipoM <> "ALM" Then Exit Function
        End If
    End If
    
    
    ArticuloConFitosantiarios = True
        
        'Los que no lleven el nuevo controlo sigue como antes
        
        CadenaInsertTmpLotes = ""
        
        
        
        
        
        SQL = "select numlotes,fecentra,Codartic,canentra - vendida"
        SQL = SQL & "  disponible from slotes where "
        SQL = SQL & " codartic=" & DBSet(txtAux(1).Text, "T") & " and canentra - vendida  >0 order by fecentra "
      
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        

        NumerloLote = ""
        CuantosLotesDistintos = 0
        If Not miRsAux.EOF Then
            CuantosLotesDistintos = 1
            NumerloLote = miRsAux!numlotes
            CantidadEnTotal = 0
            
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                If miRsAux!numlotes <> NumerloLote Then
                    'Otro lote. No controlaremos nada
                    CuantosLotesDistintos = CuantosLotesDistintos + 1
                Else
                    CantidadEnTotal = CantidadEnTotal + miRsAux!disponible
                End If
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ",0," & DBSet(miRsAux!numlotes, "T") & ")"
                SQL = miRsAux!disponible 'por si solo hay uno
               
                miRsAux.MoveNext
               
            Wend
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        
        
    
    'Si hay mas de uno mostraremos cual y cuanto puede coger
    If NumRegElim = 0 Then
        MsgBox "Ningun lote disponible para el artículo", vbExclamation
        
    Else
        'Este sera el numero de lote asignadao
        'Con lo cual, que haremos, pondremos
        
        If NumRegElim = 1 Then
            If CCur(SQL) < ImporteFormateado(txtAux(3).Text) Then
                MsgBox "Cantidad en el lote insuficiente:" & SQL, vbExclamation
                
            Else
                'Donde va la cantidad asignada en el SQL es en : ,0,'
                'Luego reeplazo esto por la cantidad del albaran
                SQL = TransformaComasPuntos(CStr(ImporteFormateado(txtAux(3).Text)))
                CadenaInsertTmpLotes = Replace(CadenaInsertTmpLotes, ",0,'", "," & SQL & ",'")
                B = True
            End If
        Else
            'Hay mas de un LOTE - Fecha entrada
            'Veremos si por lo menos es el mismo lote
            'Si es el mismo lote reasignaremos las cantidades
            If CuantosLotesDistintos = 1 Then
                'Hay mas de un lote pero de dsitintas fechas de entrada
                'Veremos i hay suficiente  o no
                If CantidadEnTotal < ImporteFormateado(txtAux(3).Text) Then
                    MsgBox "Cantidad en el lote insuficiente:" & CantidadEnTotal & "(+)", vbExclamation
                    
                Else
                    'Hay suficiente en este LOTE. Volvemos a abri , PARA este lote y volvemos a cargar el SQL
                    SQL = "select numlotes,fecentra,Codartic,canentra - vendida"
                    SQL = SQL & "  disponible from slotes where "
                    SQL = SQL & " codartic=" & DBSet(txtAux(1).Text, "T") & " and canentra - vendida  >0"
                    SQL = SQL & " AND numlotes= " & DBSet(NumerloLote, "T") & " order by fecentra "
                    Set miRsAux = New ADODB.Recordset
                    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    CadenaInsertTmpLotes = "" 'Vamos a volver a cargar el SQL de insert in lotes
                    CantidadEnTotal = ImporteFormateado(txtAux(3).Text)
                    NumRegElim = 0
                    While Not miRsAux.EOF
                        NumRegElim = NumRegElim + 1
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                         'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N")
                         
                         If CantidadEnTotal > miRsAux!disponible Then
                             CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible, "N")
                             CantidadEnTotal = CantidadEnTotal - miRsAux!disponible
                             CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                             miRsAux.MoveNext
                         Else
                            CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(CantidadEnTotal, "N")
                            CantidadEnTotal = 0
                            CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                            
                            While Not miRsAux.EOF
                                miRsAux.MoveNext
                            Wend
                         End If
                         'Vuevlo a dejar numregelim por lo menos a DOS, para que no vea que es lote unico
                         NumRegElim = 2
                    Wend
                    miRsAux.Close
                    Set miRsAux = Nothing
                    B = True
                End If 'Cantidad suficoente
            Else
                B = True 'Que lanze frmasignarlotes
            End If  'mas de un lote
        End If
    
        'Mas de un  lote disponible
        Screen.MousePointer = vbHourglass
        
        
        If B Then
            conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.codigo
            Espera 0.3
            CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
            CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
            conn.Execute CadenaInsertTmpLotes
            
            
            
            If NumRegElim = 1 Then
                CadenaDesdeOtroForm = "OK"
                Espera 0.3
            Else
                If CuantosLotesDistintos = 1 Then
                    CadenaDesdeOtroForm = "OK"
                    Espera 0.3
                Else
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = ImporteFormateado(txtAux(3).Text)
                    frmFacTPVLotes.NombreArticulo = txtAux(2).Text
                    frmFacTPVLotes.Show vbModal
                End If
            End If
            If CadenaDesdeOtroForm <> "OK" Then
                'Ha cancelado el proceso
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.codigo
                Espera 0.3
            Else
                DatosLotesFitosOk = True
            End If
        End If
        Screen.MousePointer = vbDefault
    End If   'Numregeleim0
    
End Function




Private Sub ModificaLote()
Dim CadenaInsertTmpLotes As String
Dim J As Integer
Dim LotesArticulos
Dim IncioBusqueda As Integer
Dim fin As Boolean
Dim SQL As String
Dim CadenaOR As String

        Set miRsAux = New ADODB.Recordset
          
        If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Sub   'Por si acaso se ha metido aqui
        If DBLet(data2.Recordset!numLote, "T") = "" Then Exit Sub
          
        CadenaInsertTmpLotes = "codtipom ='" & Data1.Recordset!codtipom & "' AND numalbar =" & Data1.Recordset!NumAlbar
        CadenaInsertTmpLotes = CadenaInsertTmpLotes & " AND numlinea =" & data2.Recordset!numlinea
        CadenaInsertTmpLotes = "Select numlote,cantidad,fecentra from slialblotes  WHERE " & CadenaInsertTmpLotes & "  order by sublinea"
 
        miRsAux.Open CadenaInsertTmpLotes, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        LotesArticulos = "|"
        While Not miRsAux.EOF
            LotesArticulos = LotesArticulos & miRsAux!numLote & "#@#" & Format(miRsAux!fecentra, "dd/mm/yyyy") & Mid(miRsAux!cantidad & Space(10), 1, 10) & "|"
            CadenaOR = CadenaOR & ", " & DBSet(miRsAux!numLote, "T")
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
            CadenaInsertTmpLotes = ""
            SQL = "select codartic,numlotes,fecentra,canentra-vendida disponible from slotes where "
            SQL = SQL & " codartic=" & DBSet(data2.Recordset!codArtic, "T") & " and canentra-vendida >0  "
            If CadenaOR <> "" Then
                CadenaOR = Mid(CadenaOR, 2)
                SQL = SQL & "  OR numlotes in(" & CadenaOR & ")"
            End If
            SQL = SQL & "  order by fecentra "
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N") & ","
                                
                SQL = "|" & miRsAux!numlotes & "#@#"
                fin = False
                IncioBusqueda = 1
                
                While Not fin
                    
                     
                    J = InStr(IncioBusqueda, LotesArticulos, SQL)
                    If J > 0 Then
                        J = J + Len(SQL)
                        SQL = Mid(LotesArticulos, J, 10)
                        
                        If SQL = Format(miRsAux!fecentra, "dd/mm/yyyy") Then
                            'OK, esta es la linea
                            SQL = Trim(Mid(LotesArticulos, J + 10, 10))
                            fin = True
                        Else
                            SQL = "|" & miRsAux!numlotes & "#@#"   'Vuelve a la busqueda
                            IncioBusqueda = InStr(J, LotesArticulos, "|")
                        End If
                    Else
                        SQL = "0"
                        fin = True
                    End If
                Wend
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & SQL
                
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
                
                
            'Si hay mas de uno mostraremos cual y cuanto puede coger
            If NumRegElim = 0 Then
                MsgBox "Ningun lote disponible para el artículo", vbExclamation
              
            Else
              
                'Mas de un  lote disponible
                Screen.MousePointer = vbHourglass
                
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.codigo
                Espera 0.3
                CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
                CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
                conn.Execute CadenaInsertTmpLotes
                
                
                
              
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = data2.Recordset!cantidad
                    frmFacTPVLotes.NombreArticulo = data2.Recordset!NomArtic
                    frmFacTPVLotes.Show vbModal
              
                    If CadenaDesdeOtroForm = "OK" Then
                    
                        'Primero devolveremos la cantidad que tenia la linea
                        ReestablecerLotesArticulo data2.Recordset!numlinea
                        
                        'Borramos la linea de lotes
                        SQL = "codtipom ='" & Data1.Recordset!codtipom & "' AND numalbar =" & Data1.Recordset!NumAlbar
                        SQL = SQL & " AND numlinea =" & data2.Recordset!numlinea
                        SQL = "DELETE FROM slialblotes WHERE " & SQL
                        conn.Execute SQL
                        Espera 0.4
                        
                        SQL = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                        SQL = SQL & " SELECT '" & Data1.Recordset!codtipom & "'," & Data1.Recordset!NumAlbar & "," & data2.Recordset!numlinea
                        SQL = SQL & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
                        SQL = SQL & " FROM tmpnlotes  WHERE codusu = " & vUsu.codigo & " and cantidad <>0 "
            
                        conn.Execute SQL
                        
                        'Tengo que updatear la cantidad vendida
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open "Select * FROM tmpnlotes  WHERE codusu = " & vUsu.codigo & " and cantidad <>0 ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not miRsAux.EOF
                            If miRsAux!cantidad <> 0 Then
                                If miRsAux!cantidad > 0 Then
                                    SQL = "+"
                                Else
                                    SQL = "-"
                                End If
                                SQL = "UPDATE slotes SET vendida=vendida " & SQL & DBSet(Abs(miRsAux!cantidad), "N")
                                SQL = SQL & " WHERE numlotes =" & DBSet(miRsAux!numlotes, "T") & " AND codartic= " & DBSet(miRsAux!codArtic, "T")
                                SQL = SQL & " AND fecentra= " & DBSet(miRsAux!FechaAlb, "F")
                            
                                conn.Execute SQL
                            End If
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                    End If
            
            

                    Espera 0.3
                        
                        
                    
              
            End If


    

End Sub




'Aqui solo entra si manipuladorfitos es true
Private Function VerCarnetManipulador() As Boolean
Dim LlevaLotes As Boolean
'Veremos cuanto suman las cantidades de los articulos que llevan loes

    Set miRsAux = New ADODB.Recordset
    
    
    If vParamAplic.NumeroInstalacion = 1 Then
        ' En alzira, los ALZ y los ALI     no , repito, no es obligado
        If UCase(Text1(30).Text) = "ALS" Or UCase(Text1(30).Text) = "ALI" Then
            VerCarnetManipulador = True
            Exit Function
        End If
    End If
    
    
    VerCarnetManipulador = False
    'Veremos si hay articulos con fitosanitarios
    
    If Text1(47).Text <> "" Then
        If CDate(Text1(47).Text) < CDate(Text1(1).Text) Then
            If MsgBox("Carnet caducado. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
    End If
    
    
    BuscaChekc = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    BuscaChekc = NomTablaLineas & ",sartic WHERE " & NomTablaLineas & ".codartic =sartic.codartic AND " & BuscaChekc
    BuscaChekc = BuscaChekc & " AND numserie <> ''"
    BuscaChekc = "Select sartic.codartic,sum(cantidad) Cuantos FROM " & BuscaChekc & " GROUP BY 1"
    
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    txtAnterior = ""
    LlevaLotes = False
    If miRsAux.EOF Then
        'OK
        
        VerCarnetManipulador = True
    Else
        LlevaLotes = True
        While Not miRsAux.EOF
            'Si hay lotes, si suman los mismos que los introducidos
            
            
            BuscaChekc = Replace(ObtenerWhereCP(False), NombreTabla, "slialblotes")
            BuscaChekc = BuscaChekc & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND 1"
            BuscaChekc = DevuelveDesdeBD(conAri, "sum(cantidad)", "slialblotes", BuscaChekc, "1")
            If BuscaChekc = "" Then BuscaChekc = "0.0"
            If CCur(BuscaChekc) <> miRsAux!Cuantos Then
                txtAnterior = txtAnterior & "- " & Mid(miRsAux!codArtic & Space(20), 1, 20) & miRsAux!Cuantos & " // " & BuscaChekc & vbCrLf
            End If
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If txtAnterior <> "" Then
        txtAnterior = "Error en lotes: " & vbCrLf & vbCrLf & txtAnterior
        MsgBox txtAnterior, vbExclamation
    
    Else
        If LlevaLotes Then
            'Veamos si ha puesto el carnet
            If Me.Text1(45).Text = "" Or Me.Text1(46).Text = "" Then
                txtAnterior = ""
                If Me.hcoCodTipoM = "ART" Then
                    'Factura rectificativa.
                    'Ess de servicio?
                    NumRegElim = InStr(1, Text1(18).Text, "RECTIFICA A FACTURA:")
                    If NumRegElim > 0 Then
                        BuscaChekc = Mid(Text1(18).Text, 21)
                        NumRegElim = InStr(1, BuscaChekc, ",")
                        If NumRegElim > 0 Then
                            BuscaChekc = Trim(Mid(BuscaChekc, 1, NumRegElim - 1))
                            BuscaChekc = DevuelveDesdeBD(conAri, "codtipom", "stipom", "letraser", BuscaChekc, "T")
                            If BuscaChekc = "FAS" Then txtAnterior = "OK"
                        End If
                    End If
                    If txtAnterior = "" Then
                        MsgBox "Falta carnet de manipulador", vbExclamation
                    Else
                        VerCarnetManipulador = True
                    End If
                Else
                    MsgBox "Falta carnet de manipulador", vbExclamation
                End If
            Else
                VerCarnetManipulador = True
            End If
        End If
    End If
    txtAnterior = "": BuscaChekc = ""
End Function






Private Sub ComprobarComisionesAlbaranes()
Dim Aux As String

    On Error GoTo eComprobarComisionesAlbaranes
    

    Me.lblIndicador.Caption = "Comisones"
    Me.lblIndicador.Refresh
    Espera 0.2
    
    
    Set miRsAux = New ADODB.Recordset
    Aux = " AND NOT codartic IN (" & DBSet(vParamAplic.ArtReciclado, "T")
    If vParamAplic.ArtPortesN Then Aux = Aux & "," & DBSet(vParamAplic.ArtPortesN, "T")
    
    
    Aux = " AND codtipom = " & DBSet(Data1.Recordset!codtipom, "T") & Aux & ")"
    Aux = "Select * from slialB WHERE coalesce(comisionagente,0)=0 AND numalbar = " & Data1.Recordset!NumAlbar & Aux
        
    
    
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Aux = ""
    While Not miRsAux.EOF
        Aux = Aux & "A:" & miRsAux!codArtic & "    L:" & miRsAux!numlinea & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close


    If Aux <> "" Then
        Set LOG = New cLOG
        Aux = "Albaran con lineas sin comision" & vbCrLf
        LOG.Insertar 29, vUsu, Aux
        Set LOG = Nothing
        Espera 0.5
    End If
    
eComprobarComisionesAlbaranes:
    If Err.Number <> 0 Then Err.Clear
    Set miRsAux = Nothing
End Sub






Private Sub ObtenerPuntos()
Dim C As String
    C = ObtenerWhereCP(False) & " AND 1"
    C = DevuelveDesdeBD(conAri, "puntos", "scaalb", C, "1")
    Text2(1).Text = C
End Sub




Private Sub ImprimirAlbaranFirmado()
Dim Cade As String
    On Error GoTo eImprimirAlbaranFirmado
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    Cade = "_" & Data1.Recordset!codtipom & "-" & Format(Data1.Recordset!NumAlbar, "0000000")
    Cade = vParamAplic.PathFirmasAlbaran & "\" & Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\*" & Cade & "*.pdf"
    Cade = Dir(Cade, vbArchive)
    If Cade <> "" Then
        Cade = vParamAplic.PathFirmasAlbaran & "\" & Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\" & Cade
        LanzaVisorMimeDocumento Me.hwnd, Cade
    Else
        Cade = Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\*" & Data1.Recordset!codtipom & "-" & Format(Data1.Recordset!NumAlbar, "0000000") & " *"
        Cade = "No existe ninguna documento de albaran firmado. " & vbCrLf & Cade
        MsgBox Cade, vbExclamation
    End If
    Exit Sub
eImprimirAlbaranFirmado:
        MuestraError Err.Number, , Err.Description
End Sub




Private Function PrecioMinimoAlbaran() As Boolean
Dim vPrecioFact As CPreciosFact
Dim vArtic As CArticulo
Dim RN As ADODB.Recordset
Dim B As Boolean
Dim Aux As String
Dim PrMinimo As Currency
Dim ErroresPrMinimo As String
    On Error GoTo ePrecioMinimoAlbaran
    
    PrecioMinimoAlbaran = False
    
    Set vPrecioFact = New CPreciosFact
    vPrecioFact.CodigoClien = Text1(4).Text
    vPrecioFact.FijarTarifaActividad
    
    Aux = "slialb.codartic=sartic.codartic and artvario=0 and origpre<>'P'  and origpre<>'E' AND codtipom = '" & Data1.Recordset!codtipom & "' AND numalbar =" & CStr(Data1.Recordset!NumAlbar)
    Aux = "Select slialb.* from slialb,sartic WHERE " & Aux
    Set RN = New ADODB.Recordset
    ErroresPrMinimo = ""
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF

        If RN!cantidad <> 0 Then
            PrMinimo = Round2(RN!ImporteL / RN!cantidad, 4)
            
            Set vArtic = New CArticulo
            If Not vArtic.LeerDatos(RN!codArtic) Then Err.Raise 513, , "Imposible leer datos articulo: " & RN!codArtic
            'If Not vArtic.EstablecidoPrecioMinimo Then vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
            vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
            If vArtic.EstablecidoPrecioMinimo Then
                
                If PrMinimo < vArtic.PrecioMinimo Then
                    Aux = vbCrLf & " .- " & RN!codArtic & "   " & RN!NomArtic & " [" & vArtic.PrecioMinimo & "]"
                    
                    PrMinimo = PrMinimo - vArtic.PrecioMinimo
                    If Abs(PrMinimo) > 0.01 Then ErroresPrMinimo = ErroresPrMinimo & Aux
                End If
            End If

            Set vArtic = Nothing
        End If
        RN.MoveNext
    Wend
    RN.Close

    If ErroresPrMinimo <> "" Then
        If Len(ErroresPrMinimo) > 400 Then ErroresPrMinimo = Mid(ErroresPrMinimo, 1, 400) & "......."
        Aux = "Precio inferior al mínimo permitido" & vbCrLf & ErroresPrMinimo
        If vUsu.Nivel = 0 Then
            Aux = Aux & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Aux, vbQuestion + vbYesNoCancel) = vbYes Then PrecioMinimoAlbaran = True
        Else
            MsgBox Aux, vbExclamation
        End If
    Else
        PrecioMinimoAlbaran = True
    End If
ePrecioMinimoAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set vArtic = Nothing
    Set vPrecioFact = Nothing
    Set RN = Nothing
End Function



Private Sub BotonesToolBarAux()
Dim B As Boolean


    '   5.-  Mantenimiento Lineas
    B = False
    If Not EsHistorico Then B = Modo = 2 Or Modo = 5
        
    
    ToolbarAux(0).Buttons(1).Enabled = B
    If B Then B = Me.data2.Recordset.RecordCount > 0

    
    ToolbarAux(0).Buttons(2).Enabled = B
    ToolbarAux(0).Buttons(3).Enabled = B
    
    ToolbarAux(0).Buttons(5).Enabled = B
    ToolbarAux(0).Buttons(6).Enabled = B
    ToolbarAux(0).Buttons(7).Enabled = B
    
    
    
    '   6.-  Mantenimiento Lineas
    B = False
    If Not EsHistorico Then B = Modo = 2 Or Modo = 6
        
    
    
    If vParamAplic.CartaPortes Then
        ToolbarAux(1).Buttons(1).Enabled = B
        If B Then B = Me.data3.Recordset.RecordCount > 0
        ToolbarAux(1).Buttons(3).Enabled = B
        
    End If
    
    

End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub




Private Sub BotonAnyadirLineaMatricula()

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
       
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    
    lblIndicador.Caption = "INSERTAR matricula"
    
    AnyadirLinea DataGrid2, data3
    CargaTxtAux2 True, True
    
    'VEr si para el transportista tiene matricula por defecto
    
    
    
    PonerFoco txtAux2(0)
    Me.DataGrid2.Enabled = False
End Sub


'Lineas
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)

Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux2.Count - 1 'TextBox
            txtAux2(i).Top = 330
            txtAux2(i).visible = visible
        Next i
        cmdAux2(0).visible = visible
        
    Else
        
        
        
        alto = ObtenerAlto(DataGrid2, 30)
        For i = 0 To txtAux2.Count - 1
            
            txtAux2(i).Left = DataGrid2.Columns(i + 2).Left + DataGrid2.Left
            txtAux2(i).Width = DataGrid2.Columns(i + 2).Width - 30
        
            If limpiar Then
                txtAux2(i).Text = ""
            Else
                txtAux2(i).Text = DataGrid1.Columns(i + 2).Text
                txtAux2(i).Locked = False
            End If
            txtAux2(i).Top = alto
            txtAux2(i).Height = DataGrid1.RowHeight
            txtAux2(i).visible = True
            BloquearTxt txtAux2(i), IIf(limpiar, False, i = 0)
        Next i
        cmdAux2(0).Left = txtAux2(1).Left - 60
        cmdAux2(0).Enabled = True
        cmdAux2(0).visible = True
        cmdAux2(0).Top = alto
        
        
                

    End If
End Sub




Private Function InsertarModificarMatricula() As Boolean
Dim SQL As String
Dim Rc As Byte
    'Pequeña comprobacion
    InsertarModificarMatricula = False
    txtAux2(0).Text = Trim(txtAux2(0).Text)
    
    If txtAux2(0).Text = "" Then
        MsgBox "Matricula obligatoria", vbExclamation
        Exit Function
    End If
        'matriculas
        SQL = "codenvio=" & Text1(29).Text & " AND matricula"
        SQL = DevuelveDesdeBD(conAri, "codenvio", "smatriculas", SQL, txtAux2(0).Text, "T")
        If SQL = "" Then
            SQL = "Transportista " & Text1(29).Text & " - " & Text2(29).Text & vbCrLf & vbCrLf
            SQL = SQL & "Matricula: " & txtAux2(0).Text & "    ¿Crear en base de datos?"
            Rc = MsgBox(SQL, vbQuestion + vbYesNoCancel)
            If Rc = vbCancel Then Exit Function
            If Rc = vbYes Then
                SQL = "INSERT INTO smatriculas (matricula  ,codenvio  ,titulo ,defecto) VALUES ("
                SQL = SQL & DBSet(txtAux2(0).Text, "T") & "," & Text1(29).Text & "," & DBSet(txtAux2(1).Text, "T", "N") & ",0)"
                If Not ejecutar(SQL, False) Then Exit Function
            End If
        End If
        SQL = "REPLACE INTO scaalb_portes(codtipom,numalbar,matricula,descr)"
        SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text)
        SQL = SQL & "," & DBSet(txtAux2(0).Text, "T") & ", " & DBSet(txtAux2(1).Text, "T") & ")"
        
    If Not ejecutar(SQL, True) Then
        MsgBox SQL, vbExclamation
    Else
        InsertarModificarMatricula = True
    End If


End Function

Private Sub TxtAux2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo = 6 Then
        If Index = 0 And KeyCode = 43 Or KeyCode = 107 Then
            cmdAux2_Click 1
            If InStr(1, txtAux2(Index).Text, "+") > 0 Then txtAux2(Index).Text = ""
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    
End Sub

Private Sub POnerChoferDefecto()
Dim C As String
    C = ""
    If Text1(29).Text <> "" Then
        C = "defecto = 1 AND codenvio"
        C = DevuelveDesdeBD(conAri, "concat(chofer,'|',nombre,' (',dni,')','|')", "sconductor", C, Text1(29).Text, "T")
        If C <> "" Then
            Text1(54).Text = RecuperaValor(C, 1)
            Text2(54).Text = RecuperaValor(C, 2)
        End If
    End If
    
End Sub
