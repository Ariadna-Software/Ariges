VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacClienPot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes potenciales"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ForeColor       =   &H00800000&
   Icon            =   "frmFacClienPot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   580
      Left            =   120
      TabIndex        =   32
      Top             =   450
      Width           =   11655
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Nombre Comercial|T|N|||sclipot|nomcomer||N|"
         Text            =   "Text1"
         Top             =   135
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre Cliente|T|N|||sclipot|nomclien||N|"
         Text            =   "Text1"
         Top             =   135
         Width           =   4125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   670
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Cliente|N|N|0|999999|sclipot|codclien|000000|S|"
         Text            =   "Text1"
         Top             =   135
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Nom.Comercial"
         Height          =   255
         Index           =   122
         Left            =   6600
         TabIndex        =   35
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   34
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   1
      Left            =   2880
      TabIndex        =   30
      Top             =   6960
      Width           =   4575
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   6960
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10440
      TabIndex        =   25
      Top             =   7080
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   7080
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6960
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      TabIndex        =   28
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pasar a cliente"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Direccion de envio"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Datos contacto"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Etiquetas"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cartas"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impresion CRM"
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
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   9360
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   5655
      Left            =   120
      TabIndex        =   36
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFacClienPot.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgBuscar(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgFecha(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(34)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgWeb"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgBuscar(9)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgBuscar(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(111)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(37)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(13)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text2(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text2(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(10)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text2(12)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text2(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(11)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "frameComercial"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "frameAdmon"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(22)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(8)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(7)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(6)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(5)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(4)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(3)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Contactos"
      TabPicture(1)   =   "frmFacClienPot.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(60)"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "imgBuscar(14)"
      Tab(1).Control(3)=   "Label1(61)"
      Tab(1).Control(4)=   "Label1(78)"
      Tab(1).Control(5)=   "Label1(63)"
      Tab(1).Control(6)=   "Label1(67)"
      Tab(1).Control(7)=   "ImgMail(3)"
      Tab(1).Control(8)=   "Label1(77)"
      Tab(1).Control(9)=   "DataGrid1"
      Tab(1).Control(10)=   "txtauxDC(0)"
      Tab(1).Control(11)=   "txtauxDC(1)"
      Tab(1).Control(12)=   "txtauxDC(2)"
      Tab(1).Control(13)=   "cboCargo"
      Tab(1).Control(14)=   "txtauxDC(4)"
      Tab(1).Control(15)=   "txtauxDC(3)"
      Tab(1).Control(16)=   "txtauxDC(5)"
      Tab(1).Control(17)=   "txtauxDC(8)"
      Tab(1).Control(18)=   "txtauxDC(6)"
      Tab(1).Control(19)=   "txtauxDC(7)"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "CRM"
      TabPicture(2)   =   "frmFacClienPot.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LabelCRM"
      Tab(2).Control(1)=   "Label1(2)"
      Tab(2).Control(2)=   "Frame3(1)"
      Tab(2).Control(3)=   "cmdAccCRM(0)"
      Tab(2).Control(4)=   "cmdAccCRM(1)"
      Tab(2).Control(5)=   "cmdAccCRM(2)"
      Tab(2).Control(6)=   "lwCRM"
      Tab(2).ControlCount=   7
      Begin MSComctlLib.ListView lwCRM 
         Height          =   4335
         Left            =   -74040
         TabIndex        =   83
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   2
         Left            =   -64320
         Picture         =   "frmFacClienPot.frx":0060
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Eliminar"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   1
         Left            =   -64920
         Picture         =   "frmFacClienPot.frx":0A62
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Impresion CRM"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   0
         Left            =   -65520
         Picture         =   "frmFacClienPot.frx":0FEC
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Acciones CRM"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4215
         Index           =   1
         Left            =   -74880
         TabIndex        =   81
         Top             =   960
         Width           =   735
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   1050
            Left            =   120
            TabIndex        =   82
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1852
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   13
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Acciones comerciales"
                  Object.Tag             =   "0"
                  Style           =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Llamadas"
                  Object.Tag             =   "1"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Correo electronico"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Cobros"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Observaciones departamento"
                  Object.Tag             =   "4"
                  Style           =   2
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Reclamaciones"
                  Object.Tag             =   "5"
                  Style           =   2
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Historial"
                  Object.Tag             =   "6"
                  Style           =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtauxDC 
         Height          =   1635
         Index           =   7
         Left            =   -67680
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   79
         Tag             =   "N|T|S|||sclipotdp|observa|||"
         Text            =   "frmFacClienPot.frx":19EE
         Top             =   3600
         Width           =   3765
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   6
         Left            =   -67680
         MaxLength       =   60
         TabIndex        =   77
         Tag             =   "N|T|S|||sclipotdp|maidirec|||"
         Text            =   "email"
         Top             =   3000
         Width           =   3765
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   8
         Left            =   -65400
         MaxLength       =   30
         TabIndex        =   75
         Tag             =   "N|T|S|||sclipotdp|id|||"
         Text            =   "id Este esta fuera de vista "
         Top             =   2280
         Width           =   1125
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   5
         Left            =   -67680
         MaxLength       =   12
         TabIndex        =   74
         Tag             =   "N|T|S|||sclipotdp|movil|||"
         Text            =   "movil"
         Top             =   2280
         Width           =   2085
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   3
         Left            =   -67680
         MaxLength       =   12
         TabIndex        =   71
         Tag             =   "N|T|S|||sclipotdp|Telefono|||"
         Text            =   "Tfno"
         Top             =   1560
         Width           =   2085
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   4
         Left            =   -65400
         MaxLength       =   5
         TabIndex        =   70
         Tag             =   "N|T|S|||sclipotdp|ext|||"
         Text            =   "extension"
         Top             =   1560
         Width           =   765
      End
      Begin VB.ComboBox cboCargo 
         Height          =   315
         Left            =   -67680
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtauxDC 
         Height          =   315
         Index           =   2
         Left            =   -67680
         MaxLength       =   40
         TabIndex        =   66
         Tag             =   "N|T|S|||sclipotdp|cargo|||"
         Text            =   "cargo"
         Top             =   840
         Width           =   3765
      End
      Begin VB.TextBox txtauxDC 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -70680
         MaxLength       =   30
         TabIndex        =   65
         Tag             =   "N|T|S|||sclipotdp|dpto|||"
         Text            =   "dpto"
         Top             =   4680
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.TextBox txtauxDC 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   64
         Tag             =   "Nombre|T|N|||sclipotdp|nombre|||"
         Text            =   "nombre"
         Top             =   4680
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Domicilio|T|S|||sclipot|domclien||N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "C.Postal|T|S|||sclipot|codpobla||N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   700
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Población|T|S|||sclipot|pobclien||N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   2340
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Provincia|T|S|||sclipot|proclien||N|"
         Text            =   "Text1"
         Top             =   1920
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "N.I.F.|T|S|||sclipot|nifclien||N|"
         Text            =   "Text1"
         Top             =   2400
         Width           =   1725
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Web|T|S|||sclipot|wwwclien||N|"
         Text            =   "Text1"
         Top             =   2880
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   1695
         Index           =   22
         Left            =   7320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Tag             =   "Observaciones|T|S|||sclipot|observac|||"
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Frame frameAdmon 
         Caption         =   "Administración"
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
         Height          =   1335
         Left            =   5880
         TabIndex        =   46
         Top             =   480
         Width           =   5415
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   14
            Left            =   960
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Contacto Admon.|T|S|||sclipot|perclie1||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   960
            MaxLength       =   20
            TabIndex        =   12
            Tag             =   "Teléfono Admon.|T|S|||sclipot|telclie1||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "Fax Admon.|T|S|||sclipot|faxclie1||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1710
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   960
            MaxLength       =   40
            TabIndex        =   14
            Tag             =   "e-mail Admon.|T|S|||sclipot|maiclie1||N|"
            Text            =   "maiclie1"
            Top             =   960
            Width           =   3990
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   39
            Left            =   2880
            TabIndex        =   48
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   495
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   960
            Width           =   240
         End
      End
      Begin VB.Frame frameComercial 
         Caption         =   "Comercial"
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
         Height          =   1335
         Left            =   5880
         TabIndex        =   41
         Top             =   2040
         Width           =   5415
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   21
            Left            =   960
            MaxLength       =   40
            TabIndex        =   18
            Tag             =   "e-mail Comercial|T|S|||sclipot|maiclie2||N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Fax Comercial|T|S|||sclipot|faxclie2||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1710
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   960
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "Teléfono Comercial|T|S|||sclipot|telclie2||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   960
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Contacto Comercial|T|S|||sclipot|perclie2||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   3990
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   42
            Left            =   2880
            TabIndex        =   44
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   645
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   960
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "Cod. Zona|N|S|0|999|sclipot|codzonas|000|N|"
         Text            =   "Tex"
         Top             =   4710
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   22
         Tag             =   "Cod. Ruta|N|S|0|999|sclipot|codrutas|000|N|"
         Text            =   "Tex"
         Top             =   5160
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   19
         Tag             =   "Cod.Actividad|N|N|0|999|sclipot|codactiv|000|N|"
         Text            =   "Tex"
         Top             =   3810
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   3810
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   12
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   5160
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "Cod. Envío|N|S|0|999|sclipot|codenvio|000|N|"
         Text            =   "Tex"
         Top             =   4260
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   4710
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   4260
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha de Alta|F|N|||sclipot|fechaalt|dd/mm/yyyy|N|"
         Top             =   540
         Width           =   1230
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   63
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8493
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Label Label1 
         Caption         =   "Esta escondido el tool3"
         Height          =   255
         Index           =   2
         Left            =   -67920
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label LabelCRM 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -74040
         TabIndex        =   84
         Top             =   600
         Width           =   5745
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   77
         Left            =   -67680
         TabIndex        =   80
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   3
         Left            =   -67200
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Email"
         Height          =   255
         Index           =   67
         Left            =   -67680
         TabIndex        =   78
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Movil"
         Height          =   255
         Index           =   63
         Left            =   -67680
         TabIndex        =   76
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Extension"
         Height          =   255
         Index           =   78
         Left            =   -65400
         TabIndex        =   73
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   61
         Left            =   -67680
         TabIndex        =   72
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   -67200
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "el cbo oculta el text dc(2)"
         Height          =   255
         Left            =   -65880
         TabIndex        =   68
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo"
         Height          =   255
         Index           =   60
         Left            =   -67680
         TabIndex        =   67
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   62
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "C. Postal"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   61
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   60
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   59
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   58
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Web"
         Height          =   255
         Index           =   37
         Left            =   240
         TabIndex        =   57
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   111
         Left            =   5880
         TabIndex        =   56
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Zona"
         Height          =   255
         Index           =   11
         Left            =   225
         TabIndex        =   55
         Top             =   4740
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Activ."
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   54
         Top             =   3840
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1245
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1245
         ToolTipText     =   "Buscar ruta"
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Ruta"
         Height          =   255
         Index           =   12
         Left            =   225
         TabIndex        =   53
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Envio"
         Height          =   255
         Index           =   10
         Left            =   225
         TabIndex        =   52
         Top             =   4320
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1245
         ToolTipText     =   "Buscar zona"
         Top             =   4800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1245
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1200
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   1200
         Picture         =   "frmFacClienPot.frx":19F6
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alta"
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   51
         Top             =   540
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1245
         Picture         =   "frmFacClienPot.frx":1F80
         ToolTipText     =   "Buscar fecha"
         Top             =   540
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   7080
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   3720
         Width           =   240
      End
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacClienPot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmA As frmFacActividades
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmE As frmFacFormasEnvio
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmR As frmFacRutas
Attribute frmR.VB_VarHelpID = -1
'Private WithEvents frmFP As frmFacFormasPago
'Private WithEvents frmAc As frmFacAgentesCom
'Private WithEvents frmT As frmFacTarifas
'Private WithEvents frmS As frmFacSituaciones
'Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
'Private WithEvents frmMtoTipCo As frmManTiposContrato



Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas de direcciones/dpto
'   6.-  "              "     de direcciones de envio
'   7.-  Per. contacto
'   8.-  Renting
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Private ModoFrame2 As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
    
Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal



Private BuscaChekc As String
Private PriVezForm As Boolean
Private ModoFrame  As Byte








Private Sub cmdAccCRM_Click(Index As Integer)
    
    'Acciones parar el CRM
    Select Case Index
    Case 1
        If Modo <> 2 Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
        If Text1(0).Text = "" Then Exit Sub
        
        frmListadoOfer.NumCod = Format(Val(Text1(0).Text), "0000") & "|" & Text1(1).Text & "|"
        frmListadoOfer.OpcionListado = 402
        frmListadoOfer.Show vbModal
     
        
    Case 0
    
        Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
        Case 0
'            'NUEVA, modificar o insertar acciones comerciales
'            frmCRMMto.DesdeElCliente = Data1.Recordset!codclien
'            frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
'            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
'            frmCRMMto.Show vbModal
        Case 1
            'NUEVA llamda EFECTUADA
            frmcrmMantePot.DesdeElCliente = Data1.Recordset!codClien
            frmcrmMantePot.TipoPredefinido = 1  'llamada
            frmcrmMantePot.DatosADevolverBusqueda = ""   'NUEVA
            frmcrmMantePot.Show vbModal
            
        Case 2
            'Emails
'            LanzarProgramaEmails
'            If MsgBox("Refrescar datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Case 3
            'NO puede insertar nada.
            Exit Sub
        Case 4
'            frmCrmObsDpto.Nuevo = True
'            frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
'            frmCrmObsDpto.Tag = Data1.Recordset!codclien
'            frmCrmObsDpto.Show vbModal
            
        Case 5
'            BuscaChekc = ""
'            If Text1(35).Text = "" Then
'                BuscaChekc = "No tiene cta contable"
'            Else
'                If Text2(35).Text = "" Then BuscaChekc = "Cta contable incorrecta"
'            End If
'            If BuscaChekc < "" Then
'                MsgBox BuscaChekc, vbExclamation
'                Exit Sub
'            End If
'            BuscaChekc = "-1|" & Text1(1).Text & "|" & Text1(35).Text & "|" & Text2(35).Text & "|"
'            frmCRMReclamas.Intercambio = BuscaChekc  'nueva
'            frmCRMReclamas.Show vbModal
'            BuscaChekc = ""
        Case 6
            'NUEVA entrada en Historial
            frmcrmMantePot.DesdeElCliente = Data1.Recordset!codClien
            frmcrmMantePot.TipoPredefinido = 2  'Historial
            frmcrmMantePot.DatosADevolverBusqueda = ""   'NUEVA
            frmcrmMantePot.Show vbModal
        End Select
        Me.Refresh
        DoEvents
        CargaDatosLWCRM
        Screen.MousePointer = vbDefault
    Case 2
    
'        If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
'            If lwCRM.SelectedItem Is Nothing Then Exit Sub
'            If MsgBox("¿Desea eliminar las observaciones del departamento " & Me.lwCRM.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'
'            BuscaChekc = "DELETE from scrmobsclipot  WHERE codclien = " & Me.Data1.Recordset!codclien & " AND dpto=" & lwCRM.SelectedItem.SubItems(3)
'            If Ejecutar(BuscaChekc, False) Then CargaDatosLWCRM
'            BuscaChekc = ""
'        ElseIf CByte(RecuperaValor(lwCRM.Tag, 1)) = 6 Then
'
'        End If
    End Select
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim b As Boolean
Dim EraNuevaLinea As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me, 1) Then
                
                 PosicionarData
                 'CargaFrameDirec2 0   'los dos
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear 'Adelante transacciones....
                    
        
                    PosicionarData
                End If
            End If
                
         Case 5, 6, 7, 8 'InsertarModificar linea
         
            'falta dessarrollar el 7
          
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            If InsertarModificarLinea Then
                If Modo = 5 Then
          '          Cad = "coddirec = " & Text3(0).Text
                ElseIf Modo = 6 Then
          '          Cad = "coddiren = " & Text4(0).Text
                Else
                    If Modo = 7 Then
                        cad = "id = " & txtauxDC(8).Text
                    Else
          '              Cad = "id = " & Me.txtauxRent(0).Text
                    End If
                End If
           '     If Modo < 7 Then CargaFrameDirec2 Modo - 4              'modo 5-> 1      modo 6-> 2
           '     If Modo = 5 Then
           '         b = SituarData(Data2, Cad, Indicador)
           '     ElseIf Modo = 6 Then
           '         b = SituarData(data3, Cad, Indicador)
                    
           '     ElseIf Modo = 7 Then
                
                        
                    LLamaLineas 0, 0
                    DataGrid1.AllowAddNew = False
                    CargaLineas True, 0
                
                    If ModificaLineas = 1 Then
                        data4.Recordset.MoveLast
                    Else
                        data4.Recordset.Find cad
                    End If
                    b = True
            '    Else
            '        '8.- Rentings
            '
            '        EraNuevaLinea = ModificaLineas = 1
            '        LLamaLineasRenting 0, 0
            '        DataGrid2.AllowAddNew = False
            '        CargaLineas True, 1
            '
            '       If ModificaLineas = 1 Then
            '            data5.Recordset.MoveLast
            '        Else
            '            data5.Recordset.Find Cad
            '        End If
            '        b = True
            '    End If
                If b Then
              '      If Modo = 5 Then
             '           PonerCamposDirecciones
              '      ElseIf Modo = 6 Then
              '          PonerCamposDireccionesEnvio
              '      ElseIf Modo = 7 Then
                        PonerDatosForaGrid False
              '      Else
              '          PonerDatosForaGridRent False
                        
                        'Pregunta para generar la factura
              '          If EraNuevaLinea Then
                        
              '              'Deberiamos comprobar si la proxima fecha de facturacion para este cliente es
              '              'anterior a la fecha de alta
              '              BuscaChekc = DevuelveDesdeBD(conAri, "max(ultfec)", "sclipotrenting", "codclien", CStr(Data1.Recordset!codclien))
              '              If BuscaChekc <> "" Then
              '                  If data5.Recordset!fecalta > CDate(BuscaChekc) Then
              '                      'No muesto el msg. Ya lo he hecho en datosoklinea
              '                      'MsgBox "Pendiente facturacion proximo periodo", vbInformation
              '                  Else
              '                      BuscaChekc = ""
              '                  End If
              '              End If
              '              If BuscaChekc = "" Then
              '                  frmListado3.Opcion = 22
              '                  frmListado3.OtrosDatos = "sclipotrenting.codclien = " & Text1(0).Text & " AND " & Cad
              '                  frmListado3.Show vbModal
              '              End If
              '              BuscaChekc = ""
              '          End If
              '      End If
                    ModificaLineas = 0
                    
                    
                    
                    
'                    lblIndicador.Caption = Indicador
                    PonerModoFrame 0, Modo
                    
                    
                    
                    
                End If
                
                PonerBotonCabecera True
                PonerFocoBtn Me.cmdRegresar
                
            End If
      
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdCancelar_Click()
Dim cad As String
Dim Indicador As String

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5  'Lineas Detalle
      '      PonerModoFrame 0, Modo
      '      If ModificaLineas = 1 Then '1 = Insertar
      '          If Not Data2.Recordset.EOF Then
      '              Data2.Recordset.MoveFirst
      '              PonerCamposDirecciones
      '          Else
      '              LimpiarCamposDirecciones2 False
      '          End If
      '      ElseIf ModificaLineas = 2 Then 'Modificar
      '           Cad = "(coddirec=" & Text3(0).Text & ")"
      '           If SituarData(Data2, Cad, Indicador) Then
      '              PonerCamposDirecciones
'     '               lblIndicador.Caption = Indicador
      '           End If
      '      End If
      '      ModificaLineas = 0
      '      PonerModoOpcionesMenu
      '      PonerFoco Text3(1)
        Case 6
            'Modificar direcciones de envio
      '      PonerModoFrame 0, Modo
      '      If ModificaLineas = 1 Then '1 = Insertar
      '          If Not data3.Recordset.EOF Then
      '              data3.Recordset.MoveFirst
      '              PonerCamposDireccionesEnvio
      '          Else
      '              LimpiarCamposDirecciones2 True
      '          End If
      '      ElseIf ModificaLineas = 2 Then 'Modificar
      '           Cad = "(coddiren=" & Text4(0).Text & ")"
      '           If SituarData(data3, Cad, Indicador) Then PonerCamposDireccionesEnvio
      '      End If
      '      ModificaLineas = 0
      '      PonerModoOpcionesMenu
      '      PonerFoco Text4(1)
        Case 7
           'Modificar direcciones de envio
            PonerModoFrame 0, Modo
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                
                If Not data4.Recordset.EOF Then data4.Recordset.MoveFirst
                
                    
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(id=" & Me.txtauxDC(8).Text & ")"
                 CargaLineas True, 0
                 data4.Recordset.Find cad
                 
                 
            End If
            PonerDatosForaGrid False
            LLamaLineas 0, 0
            ModificaLineas = 0
            PonerModoOpcionesMenu
            'PonerFoco Text4(1)
       Case 8
           'Modificar direcciones de envio
      '      PonerModoFrame 0, Modo
      '      DataGrid2.AllowAddNew = False
      '      If ModificaLineas = 1 Then '1 = Insertar
      '
      '          If Not data5.Recordset.EOF Then data5.Recordset.MoveFirst
      '
      '
      '      ElseIf ModificaLineas = 2 Then 'Modificar
      '           Cad = "(id=" & CStr(data5.Recordset!Id) & ")"
      '           CargaLineas True, 1
      '           data5.Recordset.Find Cad
      '
      '
      '      End If
      '      PonerDatosForaGridRent False
      '      LLamaLineasRenting 0, 0
      '      ModificaLineas = 0
      '      PonerModoOpcionesMenu
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    'MostrarSituacion False
    
    Text1(0).Text = SugerirCodigoSiguienteStr("sclipot", "codclien")
    FormateaCampo Text1(0)
    Text1(13).Text = Format(Now, "dd/mm/yyyy")
    'Sugerir el tipo de IVA como NORMAL
   
   
   
 'Valores por defecto desde parametros
    If vParamAplic.PorDefecto_Activ > 0 Then
        Text1(9).Text = vParamAplic.PorDefecto_Activ
        Text1_LostFocus 9
    End If
    If vParamAplic.PorDefecto_Envio > 0 Then
        Text1(10).Text = vParamAplic.PorDefecto_Envio
        Text1_LostFocus 10
    End If
    If vParamAplic.PorDefecto_Zona > 0 Then
        Text1(11).Text = vParamAplic.PorDefecto_Zona
        Text1_LostFocus 11
    End If
    If vParamAplic.PorDefecto_Ruta > 0 Then
        Text1(12).Text = vParamAplic.PorDefecto_Ruta
        Text1_LostFocus 12
    End If
    
    Me.SSTab1.Tab = 0
    PonerFoco Text1(0)
    ConseguirFoco Text1(0), Modo
   
   
   
   
   
   
   
   
   
   
   
   
   
   

    Me.SSTab1.Tab = 0
    PonerFoco Text1(0)
    ConseguirFoco Text1(0), Modo
End Sub


Private Sub BotonAnyadirLinea()
Dim aModo As Byte
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    aModo = Modo
 '   If aModo = 5 Then
 '       Me.SSTab1.Tab = 2
 '   ElseIf aModo = 6 Then
 '       Me.SSTab1.Tab = 3
 '  ElseIf aModo = 7 Then
        Me.SSTab1.Tab = 1

    PonerModoFrame 3, aModo  '3: Insertar
    ModificaLineas = 1 'Insertar
    lblIndicador.Caption = "Insertar Linea"
    PonerModoOpcionesMenu

    'Obtenemos la siguiente numero de Direc./Dpto
    vWhere = "codclien=" & Text1(0).Text
    If aModo = 5 Then
  '      Text3(0).Text = SugerirCodigoSiguienteStr("sdirec", "coddirec", vWhere)
  '      PonerFoco Text3(0)
    ElseIf aModo = 6 Then
  '      Text4(0).Text = SugerirCodigoSiguienteStr("sdirenvio", "coddiren", vWhere)
  '      PonerFoco Text4(0)
    ElseIf Modo = 7 Then
        'Situamos el grid al final
        AnyadirLinea DataGrid1, data4
        LLamaLineas ObtenerAlto(DataGrid1, 20), 1
        txtauxDC(8).Text = SugerirCodigoSiguienteStr("sclipotdp", "id", vWhere)
        PonerFoco Me.txtauxDC(0)
        cboCargo.ListIndex = 0 'el vacio
    Else
  '      AnyadirLinea DataGrid2, data5
  '      LLamaLineasRenting ObtenerAlto(DataGrid2, 20), 1
  '      txtauxRent(0).Text = SugerirCodigoSiguienteStr("sclipotrenting", "id", vWhere)
  '      PonerFoco Me.txtauxRent(1)
  '
    End If
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
        Text1(1).BackColor = vbYellow
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
Dim cad As String
    
    cad = "1=1"
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then cad = "codagent = " & vUsu.CodigoAgente
    End If
'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia2 cad
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & cad & Ordenacion
        PonerCadenaBusqueda
    End If
    
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data

    Select Case Modo
'        Case 5 'Modo Mantenimiento de Direcc./Dptos (Lineas)
'            If Data2.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data2, Index
'            PonerCamposDirecciones
'
'        Case 6
'            If Data3.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data3, Index
'            PonerCamposDirecciones
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
'            MostrarSituacion True
            'CargaFrameDirec2 0   'los dos
            
'            PonerModoOpcionesMenu
    End Select
End Sub


'0- Departamentos.    1- Direccioens de envio
Private Sub DesplazamientoLineas(Index As Integer, Cual As Byte)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
'    If Cual = 0 Then
'            If Data2.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data2, Index
'            PonerCamposDirecciones
'            If Modo = 5 And ModoFrame2 = 0 Then
'                Me.lblIndicador.Caption = "Lineas Detalle"
'                If Not Data2.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & ": " & Me.Data2.Recordset.AbsolutePosition & " de " & Me.Data2.Recordset.RecordCount
'            End If
'
'    Else
'            If Data3.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data3, Index
'            PonerCamposDireccionesEnvio
'            If Modo = 6 And ModoFrame2 = 0 Then
'                Me.lblIndicador.Caption = "Lineas envio"
'                If Not Data3.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & ": " & Me.Data3.Recordset.AbsolutePosition & " de " & Me.Data3.Recordset.RecordCount
'            End If
'    End If
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.SSTab1.Tab = 0
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
Dim aModo As Byte
'Modificar una linea
    aModo = Modo
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
'    If aModo = 5 Then
'        If Data2.Recordset.EOF Then Exit Sub
'        If Data2.Recordset.RecordCount < 1 Then Exit Sub
'        Me.SSTab1.Tab = 2
'    ElseIf aModo = 6 Then
'        If Data3.Recordset.EOF Then Exit Sub
'        If Data3.Recordset.RecordCount < 1 Then Exit Sub
'        Me.SSTab1.Tab = 3
'    ElseIf aModo = 7 Then
        If data4.Recordset.EOF Then Exit Sub
        If data4.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 1
'
'    Else
'        'Renting
'        If data5.Recordset.EOF Then Exit Sub
'        If data5.Recordset.RecordCount < 1 Then Exit Sub
'        Me.SSTab1.Tab = 8
'    End If
    
    
    
    
    
       
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 4, aModo 'ModoFrame=4 -> Modificar
    Me.lblIndicador.Caption = "Modificar Linea"
    ModificaLineas = 2 'Modificar
    PonerModoOpcionesMenu
    
'    'Como el campo 1 es clave primaria, NO se puede modificar
'    If aModo = 5 Then
'        BloquearTxt Text3(0), True
'        PonerFoco Text3(1)
'    ElseIf aModo = 6 Then
'        BloquearTxt Text4(0), True
'        PonerFoco Text4(1)
'    ElseIf aModo = 7 Then
    
                
        LLamaLineas ObtenerAlto(DataGrid1, 20), 2
        txtauxDC(0).Text = data4.Recordset!Nombre
        txtauxDC(1).Text = DBLet(data4.Recordset!Dpto, "T")
        
        PonerFoco Me.txtauxDC(0)
        
'    ElseIf aModo = 8 Then
'
'        LLamaLineasRenting ObtenerAlto(DataGrid2, 20), 2
'
'        For NumRegElim = 0 To txtauxRent.Count - 1
'
'
'                If IsNull(data5.Recordset.Fields(NumRegElim)) Then
'                    txtauxRent(NumRegElim).Text = ""
'                Else
'                    txtauxRent(NumRegElim).Text = data5.Recordset.Fields(NumRegElim)
'                End If
'
'        Next
'
'
'
'        PonerFoco Me.txtauxRent(1)
'    End If
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Modo <> 2 Then Exit Sub
    
    If Not EliminarCliente() Then Exit Sub


    '### a mano
    cad = "¿Seguro que desea ELIMINAR el cliente  potencial?"
    cad = cad & vbCrLf & "Cod. : " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Contactos
        cad = "DELETE FROM sclipotdp WHERE codclien = " & Data1.Recordset!codClien
        conn.Execute cad
        'CRM de potenciales
        cad = "DELETE FROM scrmaccionclipot WHERE codclipot = " & Data1.Recordset!codClien
        conn.Execute cad
        
        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Cliente", Err.Description
    End If
End Sub


'Private Sub BotonEliminarLinea()
'Dim Cad As String, cad2 As String
'Dim I As Integer
'
'    If Data2.Recordset.EOF Then Exit Sub
'    If Data2.Recordset.RecordCount < 1 Then Exit Sub
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'
'    If vParamAplic.Renting Then
'        Cad = "codclien = " & Data1.Recordset!codclien & " AND coddirec"
'        Cad = DevuelveDesdeBD(conAri, "count(*)", "sclipotrenting", Cad, CStr(Data2.Recordset.Fields(1)), "N")
'        If Cad = "" Then Cad = "0"
'        If Val(Cad) > 0 Then
'            MsgBox "Existen rentings de clientes asociados a este departamento/direccion", vbExclamation
'            Exit Sub
'        End If
'    End If
'
'
'
'
'
'    ModificaLineas = 3 'Eliminar
'
'    'Dependiendo del parametro de la aplicacion trabajamos con Dpto o Direc.
'    If vParamAplic.HayDeparNuevo = 1 Then
'        cad2 = " Dpto. "
'        Cad = " el Departamento?"
'    ElseIf vParamAplic.HayDeparNuevo = 0 Then
'        cad2 = " Direc. "
'        Cad = " la Dirección?"
'    Else
'        cad2 = " Obra "
'        Cad = " la obra?"
'    End If
'
'    Cad = "¿Seguro que desea eliminar " & Cad & vbCrLf
'    Cad = Cad & vbCrLf & "Cod." & cad2 & ": " & Format(Data2.Recordset.Fields(1), FormatoCampo(Text3(0)))
'    Cad = Cad & vbCrLf & "Nombre" & cad2 & ": " & Data2.Recordset.Fields(2)
'
'    'Borramos
'    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        On Error GoTo Error2
'        Screen.MousePointer = vbHourglass
'        NumRegElim = Data2.Recordset.AbsolutePosition
'        Data2.Recordset.Delete
'
'
'        'Para borrar en arimoeny
'        If Text1(35).Text <> "" Then
'            'SI NO tiene cta contable NO tiene dpto
'            cad2 = " WHERE codmacta= '" & Text1(35).Text & "' AND Dpto = " & Text3(0).Text
'            ConnConta.Execute "DELETE FROM departamentos " & cad2
'        End If
'
'
'        If SituarDataTrasEliminar(Data2, NumRegElim) Then
'            PonerCamposDirecciones
'        Else
'             'Solo habia un registro
'            LimpiarCamposDirecciones2 False
''            PonerModoFrame 0
'        End If
'
'        ModificaLineas = 0
'        PonerModoFrame 0, 5
'    End If
'
'    Screen.MousePointer = vbDefault
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then
'        Data2.Recordset.CancelUpdate
'        MsgBox Err.Number & ": " & Err.Description, vbExclamation
'    End If
'End Sub



'Private Sub BotonEliminarLineaDirEnvio()
''Eliminar una linea De ArticulosxAlmacen
'Dim Cad As String
'Dim I As Integer
'
'    If Data3.Recordset.EOF Then Exit Sub
'    If Data3.Recordset.RecordCount < 1 Then Exit Sub
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'    If Not PuedeEliminarDirecEnvio(True, Text1(0).Text, CInt(Data3.Recordset!coddiren)) Then Exit Sub
'
'    ModificaLineas = 3 'Eliminar
'
'
'    Cad = "¿Seguro que desea eliminar la direccion de envio" & Cad & vbCrLf
'    Cad = Cad & vbCrLf & "Codigo:  " & Format(Data3.Recordset.Fields(1), FormatoCampo(Text4(0)))
'    Cad = Cad & vbCrLf & "Nombre:  " & Data3.Recordset.Fields(2)
'
'    'Borramos
'    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        On Error GoTo Error2
'        Screen.MousePointer = vbHourglass
'        NumRegElim = Data3.Recordset.AbsolutePosition
'        Data3.Recordset.Delete
'
'        If SituarDataTrasEliminar(Data3, NumRegElim) Then
'            PonerCamposDireccionesEnvio
'        Else
'             'Solo habia un registro
'            LimpiarCamposDirecciones2 True
'
'        End If
'
'        ModificaLineas = 0
'        PonerModoFrame 0, 6
'    End If
'
'
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then
'        Data3.Recordset.CancelUpdate
'        MsgBox Err.Number & ": " & Err.Description, vbExclamation
'    End If
'End Sub


Private Sub BotonDirecciones(ElModo As Byte)

    
    On Error GoTo ErrorDirec
'    If ElModo = 7 Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    If ElModo = 5 Then
     '   Me.SSTab1.Tab = 2
    ElseIf ElModo = 6 Then
    '    Me.SSTab1.Tab = 3
    ElseIf ElModo = 7 Then
        Me.SSTab1.Tab = 1
        
        'Si primera vez qu pulsa boton..
        If Me.cboCargo.ListCount <= 0 Then CargaComboCargos
    Else
'        'Renting, si no tiene establecido el periodo de facturacion de renting, tendremos que avisarlo y NO dejarle pasar
'        If Me.cboFraRenting.ListIndex < 0 Then
'            MsgBox "No tiene establecido el periodo de facturación de renting", vbExclamation
'            Me.SSTab1.Tab = 1
'            Exit Sub
'        End If
'        Me.SSTab1.Tab = 8
        

    End If
    
    Screen.MousePointer = vbHourglass
    ModificaLineas = 0
    'Poner el modo en el formulario
    PonerModo (ElModo) 'Modo 5: Modificar lineas
    PonerModoFrame 0, ElModo 'TextBox Bloqueados inicialmente
    
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault

    Exit Sub
ErrorDirec:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim Indicador As String

    'Quitar lineas y volver a la cabecera
    If Modo >= 5 Then  'modo 5: Lineas Direcciones/Departamentos
        cad = "(codclien=" & Text1(0).Text & ")"
        If SituarData(Data1, cad, Indicador) Then
'            PonerLineaVisible False
            PonerModo 2
            lblIndicador.Caption = Indicador
        End If
    Else 'Regresar
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = cad & Data1.Recordset!perclie1 & "|"
        cad = cad & Data1.Recordset!maiclie1 & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub



'Private Sub cmdRenting_Click(Index As Integer)
'
'   If Index = 0 Then
'        'Departamento
'        imgBuscar(0).Tag = 1000
'        MandaBusquedaPrevia2 "codclien=" & Text1(0).Text
'
'
'
'
'   ElseIf Index = 3 Then
'        'tipco
'        BuscaChekc = ""
'        Set frmMtoTipCo = New frmManTiposContrato
'        frmMtoTipCo.DatosADevolverBusqueda = "0"
'        frmMtoTipCo.Show vbModal
'        Set frmMtoTipCo = Nothing
'        If BuscaChekc <> "" Then
'            Me.txtauxRent(8).Text = RecuperaValor(BuscaChekc, 1)
'            Me.txtauxRent(9).Text = RecuperaValor(BuscaChekc, 2)
'            PonerFoco txtauxRent(10)
'            BuscaChekc = ""
'        End If
'
'
'
'   Else
'        'FECHAS
'        If Index = 1 Then
'            imgFecha(0).Tag = 1004
'        Else
'            imgFecha(0).Tag = 1006
'        End If
'        Set frmF = New frmCal
'        frmF.Fecha = Now
'
'
'
'       'PonerFormatoFecha Text1(Indice)
'       'If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)
'
'       Screen.MousePointer = vbDefault
'       frmF.Show vbModal
'       Set frmF = Nothing
'
'    End If
'End Sub

Private Sub Data4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 7 And ModificaLineas > 0 Then Exit Sub
    If Not data4.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGrid False
    Else
       ' Caption = "EOF"
         PonerDatosForaGrid True
    End If
End Sub

'Private Sub data5_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If Modo = 8 And ModificaLineas > 0 Then Exit Sub
'    If Not data5.Recordset.EOF Then
'        'Caption = data4.Recordset!Id
'        PonerDatosForaGridRent False
'    Else
'       ' Caption = "EOF"
'         PonerDatosForaGridRent True
'    End If
'End Sub

Private Sub DataGrid1_Click()
    If Not data4.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGrid False
End Sub

'Private Sub DataGrid2_Click()
'    If Not data4.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridRent False
'End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PriVezForm Then
        PriVezForm = False
        If Not Data1.Recordset.EOF Then
            PonerModo 2
            PonerCampos
        Else
        
            'BotonAnyadir
            PonerModo 0
        End If
    End If
        
    If Modo = 1 Then PonerFoco Text1(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim Imag
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PriVezForm = True
    
    'Icono de busqueda
    For Each Imag In Me.imgBuscar
    'For kCampo = 0 To Me.imgBuscar.Count - 1
        'Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Imag.Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next 'kCampo
    
    'Icono de e-mail
'    For kCampo = 0 To Me.ImgMail.Count - 1
'        Me.ImgMail(kCampo).Picture = frmPpal.imgListComun.ListImages(20).Picture
'    Next kCampo
    For Each Imag In Me.ImgMail
    'For kCampo = 0 To Me.imgBuscar.Count - 1
        'Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Imag.Picture = frmPpal.imgListComun.ListImages(20).Picture
    Next 'kCampo


    ' ICONITOS DE LA BARRA
    btnAnyadir = 6
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(6).Image = 3   'Insertar Nuevo
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        .Buttons(10).Image = 17 'Pasar a cliente
        
        'octubre 2010
'        .Buttons(11).Image = 29 'Direcciones de envio
        .Buttons(12).Image = 37 'Datos contacto
        .Buttons(13).Image = 29 'cartas
        .Buttons(14).Image = 47 'etiq
        .Buttons(15).Image = 40 'Impresion CRM
        
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
    
    
'    'BARRA DE LAS LINEAS de DIRECCION/DPTOS

   'Ene18 YA dejo pasar Toolbar1.Buttons(10).visible = vParamAplic.NumeroInstalacion = 0   'Se lo prohibo a Herbelca....
    Toolbar1.Buttons(11).visible = False 'vParamAplic.DireccionesEnvio


    ImagenesNavegacion
    
  
    
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    ModificaLineas = 0
       
    'Si hay algun combo los cargamos
 '   CargarComboAlbaran
 '   CargarComboFacturacion
 '   CargarComboTipoIVA
 '   CargaComboTipoCliente
 '   CargaComboFrarRenting
    
    Me.lblSituacion.visible = False
    Me.Frame1(1).visible = False
    

    
    
    
    
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sclipot, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: cuentas, BD: Conta.
    imgBuscar(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "sclipot"
    Ordenacion = " ORDER BY codclien"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    
    If DatosADevolverBusqueda <> "" Then
        If InStr(1, DatosADevolverBusqueda, "|") = 0 Then
            'QUIERO VER EL CLIENTE
            Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=" & DatosADevolverBusqueda
        End If
    End If
    Data1.Refresh
    
    'Asignamos un SQL al DATA2
    'CargaFrameDirec2 0   'los dos
    txtauxDC(8).Left = 23000 'para que no se vea
    

    
    CargaColumnasCRM 1
    

    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        If Data1.Recordset.EOF Then
            PonerModo 1
        Else
           'LO pondra en el activatre
            
        End If
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    'Me.chkClienteV.Value = 0
    'chkCredPriv.Value = 0
    'Me.chkAbonos.Value = 0
    'Me.chkPromociones.Value = 0
    'Me.chkRentingDpto.Value = 0
    'Me.chkReferencia.Value = 0
    'Me.chkTasaReciclado.Value = 0
    'Me.chkCorreo.Value = 0
    'Me.chkPortesFac.Value = 0
    'Me.chkRecargFinan.Value = 0
    'Me.chkParticular.Value = 0
    'Me.cboAlbaran.ListIndex = -1
    'Me.cboFacturacion.ListIndex = -1
    'Me.cboTipoIVA.ListIndex = -1
    'Me.cboFraRenting.ListIndex = -1
    'cboTipocliente.ListIndex = -1
    CargaLineas False, 2
    If Err.Number <> 0 Then Err.Clear
End Sub


'Private Sub LimpiarCamposDirecciones2(DeEnvio As Boolean)
'Dim I As Byte
'    'Limpia los controles TextBox3
'    If Not DeEnvio Then
'        For I = 0 To Text3.Count - 1
'            Text3(I).Text = ""
'        Next I
'        txtZona(14).Text = ""
'    Else
'        For I = 0 To Text4.Count - 1
'            Text4(I).Text = ""
'        Next I
'        txtZona(10).Text = ""
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Actividades
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

'Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
''Agentes Comerciales
'    Text1(36).Text = RecuperaValor(CadenaSeleccion, 1)
'    FormateaCampo Text1(36)
'    Text2(36).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
  
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If Val(imgBuscar(0).Tag) >= 0 Then
            If Val(imgBuscar(0).Tag) = 1000 Then
                'Departamentos en RENTING
                
      '          txtauxRent(1).Text = RecuperaValor(CadenaDevuelta, 1)
      '          txtauxRent(2).Text = RecuperaValor(CadenaDevuelta, 2)
                
            Else
                'Se llama desde el botón de busqueda del campo Tipos de IVA
                'Recuperar solo el campo código y Descripción
    '            Indice = Val(Me.imgBuscar(0).Tag)
                Text1(35).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(35).Text = RecuperaValor(CadenaDevuelta, 2)
        
            End If
        Else
            'Recupera todo el registro de Artículos
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
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

    If CByte(Me.imgBuscar(0).Tag) = 9 Then Indice = 4
    If Indice = 4 Then 'Form Principal de Clientes
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        'Poblacion
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
        'provincia
        Text1(Indice + 2).Text = devuelve

    Else 'Lineas de Direcciones/Dptos
      '  If Me.imgBuscar(0).Tag = 10 Then
      '      Text3(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
      '      Text3(4).Text = ObtenerPoblacion(Text3(3).Text, devuelve)  'Poblacion
      '      'provincia
      '      Text3(5).Text = devuelve
      '  Else
      '      'DIRECCIONES DE ENVIO
      '      Text4(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
      '      Text4(4).Text = ObtenerPoblacion(Text3(4).Text, devuelve)  'Poblacion
      '      'provincia
      '      Text4(5).Text = devuelve
      '  End If
    End If
End Sub

'Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
'    Text1(52).Text = RecuperaValor(CadenaSeleccion, 1)
'    Text2(52).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'Formas de Envío
    Text1(10).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(10)
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Select Case Val(imgFecha(0).Tag)
        Case 0
            Indice = 13
        Case 1
            Indice = 40
        Case 2
            Indice = 41
        Case 3
            Indice = 46
        Case 4
            Indice = 48
            
        Case 5
            Indice = 53
            
        Case 1004, 1006
            'Son las fechas del RENTING
            'Me.txtauxRent(Val(imgFecha(0).Tag) - 1000).Text = Format(vFecha, "dd/mm/yyyy")
            'Exit Sub
    End Select
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmR_DatoSeleccionado(CadenaSeleccion As String)
 'Formas de Envío
    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(12)
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

'Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
''Formas de Pago
'    Text1(23).Text = RecuperaValor(CadenaSeleccion, 1)
'    FormateaCampo Text1(23)
'    Text2(23).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub
'
'Private Sub frmMtoTipCo_DatoSeleccionado(CadenaSeleccion As String)
'    BuscaChekc = CadenaSeleccion 'luego, alli(.show) lo ponemos en los txt
'End Sub

'Private Sub frmR_DatoSeleccionado(CadenaSeleccion As String)
''Rutas
'    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1)
'    FormateaCampo Text1(12)
'    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub
'
'Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
''Situaciones
'    Text1(42).Text = RecuperaValor(CadenaSeleccion, 1)
'    FormateaCampo Text1(42)
'    Text2(42).Text = RecuperaValor(CadenaSeleccion, 2)
'
'End Sub

'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Tarifas
'    Text1(37).Text = RecuperaValor(CadenaSeleccion, 1)
'    FormateaCampo Text1(37)
'    Text2(37).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
'Zonas
    If BuscaChekc = "" Then
        Text1(11).Text = RecuperaValor(CadenaSeleccion, 1)
        FormateaCampo Text1(11)
        Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
        
    Else
        'If BuscaChekc = "15" Then
        '    Text3(14).Text = RecuperaValor(CadenaSeleccion, 1)
        '    Me.txtZona(14).Text = RecuperaValor(CadenaSeleccion, 2)
        'Else
        '    Text4(10).Text = RecuperaValor(CadenaSeleccion, 1)
        '    Me.txtZona(10).Text = RecuperaValor(CadenaSeleccion, 2)
        'End If
    End If
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    'Disitnto de Observaciones
    If Index = 11 Or Index = 17 Then
    
    
    Else
        If Modo = 2 Or Modo = 0 Then Exit Sub
        
        
        If Index = 13 Then
            'En insertar NO VA direccion envio habitual
            If Modo = 3 Then
                MsgBox "Hasta que no cree el cliente no podra tener direcciones envio", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
        
    
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Actividad
            Indice = 9
            Set frmA = New frmFacActividades
            frmA.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 1  'Cod. Envio
            Indice = 10
            Set frmE = New frmFacFormasEnvio
            frmE.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmE.Show vbModal
            Set frmE = Nothing
            
            
            'Cod. Zona
        Case 2, 15, 16
            ' 2.- Zona del cliente
            ' 15.- zona del departamento
            ' 16.- De la direccion de envio
            Indice = 11
            BuscaChekc = ""
            Set frmZ = New frmFacZonas
            frmZ.DatosADevolverBusqueda = "0"
            If Index = 2 Then
                If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            Else
                BuscaChekc = Index
                Indice = 101 'para que bajo no haga ponerofo
            End If
            
            frmZ.Show vbModal
            Set frmZ = Nothing
            
        Case 3  'Cod. Ruta
            Indice = 12
            Set frmR = New frmFacRutas
            frmR.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmR.Show vbModal
            Set frmR = Nothing
            
        Case 4  'Cod. Forma de Pago
'            indice = 23
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
'            frmFP.Show vbModal
'            Set frmFP = Nothing
            
        Case 5  'Cuenta Contable
            imgBuscar(0).Tag = Index
            MandaBusquedaPrevia2 "apudirec= 'S'"
            imgBuscar(0).Tag = -1
            Indice = 35
            
        Case 6 'Código de Agente
'            indice = 36
'            Set frmAc = New frmFacAgentesCom
'            frmAc.DatosADevolverBusqueda = "0"
'            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
'            frmAc.Show vbModal
'            Set frmAc = Nothing
            
        Case 7 'Código de Tarifa
'            indice = 37
'            Set frmT = New frmFacTarifas
'            frmT.DatosADevolverBusqueda = "0"
'            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
'            frmT.Show vbModal
'            Set frmT = Nothing
            
        Case 8 'Código de Situación
'            indice = 42
'            Set frmS = New frmFacSituaciones
'            frmS.DatosADevolverBusqueda = "0"
'             If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
'            frmS.Show vbModal
'            Set frmS = Nothing
            
        Case 9, 10, 12 'CPostal
            Me.imgBuscar(0).Tag = Index
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                Indice = 4
            Else
            '    PonerFoco Text3(3)
            End If
            Me.imgBuscar(0).Tag = -1
            VieneDeBuscar = True
        Case 11, 17
            'Campos MEMO
        
            If Modo = 5 Or Modo = 0 Then
            
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    If Index = 11 Then
                        CadenaDesdeOtroForm = Text1(22).Text
                    Else
                        CadenaDesdeOtroForm = Text1(54).Text
                    End If
                        
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then
                        If Index = 11 Then
                            CadenaDesdeOtroForm = DBLet(Data1.Recordset!observac, "T")
                        Else
                            CadenaDesdeOtroForm = DBLet(Data1.Recordset!obsfacturacion, "T")
                        End If
                    End If
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then
                        If Index = 11 Then
                            Text1(22).Text = Mid(CadenaDesdeOtroForm, 3)
                        Else
                            Text1(54).Text = Mid(CadenaDesdeOtroForm, 3)
                        End If
                    End If
                End If
                CadenaDesdeOtroForm = ""
            End If
            
            
         Case 13
            
'                Set frmDptoEnvio = New frmFacCliEnvDpto
'                frmDptoEnvio.DireccionesEnvio = True
'                frmDptoEnvio.VerDatoDpto = -1
'                frmDptoEnvio.codclien = CLng(Text1(0).Text)
'                frmDptoEnvio.NomClien = Text1(1).Text
'                frmDptoEnvio.Show vbModal
'                Set frmDptoEnvio = Nothing
                
                
        Case 14
                
                frmFacCargos.Show vbModal
                CargaComboCargos
                SituarCboCargo
    End Select
    If Index <> 10 Or Index <> 12 Or Index < 100 Then PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

  
   
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
     Case 0
        Indice = 13
  
   End Select
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   

End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(17).Text
        Case 1: dirMail = Text1(21).Text
        'Case 2: dirMail = Text3(9).Text
        Case 3: dirMail = Me.txtauxDC(6).Text
    End Select

    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(8).Text) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub



'Private Sub lw1_DblClick()
'Dim Seleccionado As Long
'    If Modo <> 2 Then Exit Sub
'    If lw1.ListItems.Count = 0 Then Exit Sub
'    If lw1.SelectedItem Is Nothing Then Exit Sub
'
'
'    If Me.DatosADevolverBusqueda <> "" Then
'        'De momento NO dejo continuar
'        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    'Llegados aqui
'    Select Case CByte(RecuperaValor(lw1.Tag, 1))
'    Case 2
'        'ALBARANES
'        If vParamAplic.TipoFormularioClientes = 0 Then
'            Set frmAlb = New frmFacEntAlbaranes2
'            frmAlb.hcoCodMovim = lw1.SelectedItem.SubItems(1)
'            frmAlb.hcoCodTipoM = lw1.SelectedItem.Text
'            frmAlb.RecuperarFactu = False
'            frmAlb.Show vbModal
'            Set frmAlb = Nothing
'
'        Else
'            Set frmAlbS = New frmFacEntAlbSAIL
'            frmAlbS.hcoCodMovim = lw1.SelectedItem.SubItems(1)
'            frmAlbS.hcoCodTipoM = lw1.SelectedItem.Text
'            frmAlbS.RecuperarFactu = False
'            frmAlbS.Show vbModal
'            Set frmAlbS = Nothing
'
'
'        End If
'
'    Case 0
'        'OFERTAS
'        If vParamAplic.TipoFormularioClientes = 0 Then
'            Set frmOfe = New frmFacEntOfertas2
'            frmOfe.DatosOferta = lw1.SelectedItem.Text
'            frmOfe.Show vbModal
'            Set frmOfe = Nothing
'        Else
'            'SAIL
'            Set frmOfeS = New frmFacEntOferSAIL
'            frmOfeS.DatosOferta = lw1.SelectedItem.Text
'            frmOfeS.Show vbModal
'            Set frmOfeS = Nothing
'
'        End If
'
'    Case 1
'        'PEDIDOS
'        If vParamAplic.TipoFormularioClientes = 0 Then
'            Set frmPed = New frmFacEntPedidos
'            frmPed.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
'            frmPed.EsHistorico = False
'            frmPed.Show vbModal
'            Set frmPed = Nothing
'
'        Else
'            'SAIL
'            Set frmPedS = New frmFacEntPedSail
'            frmPedS.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
'            frmPedS.EsHistorico = False
'            frmPedS.Show vbModal
'            Set frmPedS = Nothing
'
'
'        End If
'    Case 3
'        'FACTURAS
'        'Este no necesitamos crear instancias
'
'        'Lo que ocurre que esta preparado para abrir la factura a partir de un albaran, con lo cual
'        'En la funcion abrir factura, buscare un albaran de la factura para abrirlo
'        AbrirFacturaLW
'
'
'    Case 4
'        'Precios especiales
'        'No creamos instancias
'
'        frmFacPreciosEspecial.CadenaSituarData = "'" & DevNombreSQL(lw1.SelectedItem.Text) & "'|" & Data1.Recordset!codclien & "|"
'        frmFacPreciosEspecial.Show vbModal
'    End Select
'
'    'Pase lo que pase, por si acaso, cargamos el lw
'    lw1.SetFocus
'    Seleccionado = lw1.SelectedItem.Index
'    CargaDatosLWDoc
'    lw1.SelectedItem.Selected = False
'    Set lw1.SelectedItem = Nothing
'    If lw1.ListItems.Count >= Seleccionado Then
'            lw1.ListItems(Seleccionado).Selected = True
'            lw1.ListItems(Seleccionado).EnsureVisible
'    End If
'    Screen.MousePointer = vbDefault
'End Sub




Private Sub lwCRM_DblClick()
Dim Clave As String
Dim i As Integer
    If Modo <> 2 Then Exit Sub
    If lwCRM.ListItems.Count = 0 Then Exit Sub
    If lwCRM.SelectedItem Is Nothing Then Exit Sub




     'Llegados aqui
    Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
'    Case 0
'        'Aciones comerciales
'        ' modificar o insertar acciones comerciales
'        frmCRMMto.DesdeElCliente = Data1.Recordset!codclien
'        frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
'        frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & _
'            " AND scrmacciones.Tipo = " & lwCRM.SelectedItem.SubItems(4) & " And codClien = " & Data1.Recordset!codclien
'        frmCRMMto.Show vbModal
    Case 1
        'Llamadas
'        If lwCRM.SelectedItem.SmallIcon = 27 Then
            'Lee de sllama

'            CadenaDesdeOtroForm = "`feholla`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " and `usuario`=" & DBSet(lwCRM.SelectedItem.SubItems(1), "T")
'            frmLLamadasDatos2.SoloVer = True
'            frmLLamadasDatos2.vModo = 4
'            frmLLamadasDatos2.Show vbModal
'        Else
            'Lee de acciones realizadas con tipo=1 .....

            frmcrmMantePot.DesdeElCliente = Data1.Recordset!codClien
            frmcrmMantePot.TipoPredefinido = 1 'Llamadas realizadas
            frmcrmMantePot.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmaccionclipot.Tipo = 1 And codclipot = " & Data1.Recordset!codClien
            frmcrmMantePot.Show vbModal
'
'        End If
'    Case 2
'        'MAIL
'        frmMensajes.OpcionMensaje = 21
'        If lwCRM.SelectedItem.SmallIcon = 28 Then
'            frmMensajes.cadWHERE2 = "0"
'        Else
'            frmMensajes.cadWHERE2 = "1"
'        End If
'        frmMensajes.cadWhere = "codclien = " & Text1(0).Text & " AND  entryID = '" & lwCRM.SelectedItem.SubItems(5) & "'"
'        frmMensajes.Show vbModal
'    Case 3
'        'Cobros. NO HACEMOS NADA
'        'Nos piramos
'        Exit Sub
'
'    Case 4
'        frmCrmObsDpto.Nuevo = False
'        BuscaChekc = "dpto = " & Me.lwCRM.SelectedItem.SubItems(3) & " AND codclien "
'        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "observa", "scrmobsclipot", BuscaChekc, CStr(Data1.Recordset!codclien))
'
'        frmCrmObsDpto.Dpto = CByte(Me.lwCRM.SelectedItem.SubItems(3))
'        frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
'        frmCrmObsDpto.Tag = Data1.Recordset!codclien
'        frmCrmObsDpto.Show vbModal
'
'    Case 5
'        'Reclamas n
'            BuscaChekc = lwCRM.SelectedItem.SubItems(4) & "|" & Text1(1).Text & "|"
'            frmCRMReclamas.Intercambio = BuscaChekc
'            frmCRMReclamas.Show vbModal
    
    Case 6
            frmcrmMantePot.DesdeElCliente = Data1.Recordset!codClien
            frmcrmMantePot.TipoPredefinido = 2 'Historial
            frmcrmMantePot.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmaccionclipot.Tipo = 2 And codclipot = " & Data1.Recordset!codClien
            frmcrmMantePot.Show vbModal
    End Select
    Me.Refresh
    DoEvents
    
    
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 5 Then
        Clave = lwCRM.SelectedItem.SubItems(4)
    Else
        Clave = lwCRM.SelectedItem.Text
    End If
    CargaDatosLWCRM
    
    Set lwCRM.SelectedItem = Nothing
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 5 Then
        'para encontrar en las reclamas debe buscar por el campo codigo 4
        For i = 1 To lwCRM.ListItems.Count
            If Clave = lwCRM.ListItems(i).SubItems(4) Then
                Set lwCRM.SelectedItem = lwCRM.ListItems(i)
            Else
                lwCRM.ListItems(i).Selected = False
            End If
        Next
    Else
        For i = 1 To lwCRM.ListItems.Count
            If Clave = lwCRM.ListItems(i).Text Then
                Set lwCRM.SelectedItem = lwCRM.ListItems(i)
            Else
                lwCRM.ListItems(i).Selected = False
            End If
        Next
    End If
    BuscaChekc = ""
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
     If Modo >= 5 Then 'Eliminar lineas Artículos x Almacen
      '  If Modo = 5 Then BotonEliminarLinea
      '  If Modo = 6 Then BotonEliminarLineaDirEnvio
        If Modo = 7 Then BotonEliminarLineaContacto
      '  If Modo = 8 Then BotonEliminarRenting
     Else   'Eliminar Artículo
        BotonEliminar
     End If
End Sub

Private Sub mnModificar_Click()
     If Modo >= 5 Then 'Modificar lineas Artículos x Almacen
        'FALTA: bloquear la linea !!!!
        BotonModificarLinea
     Else   'Modificar Artículos
        If BLOQUEADesdeFormulario(Me, 1) Then BotonModificar
     End If
End Sub

Private Sub mnNuevo_Click()
     If Modo >= 5 Then          'Añadir lineas Artículos x Almacen
        BotonAnyadirLinea
    Else 'Añadir Artículos
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



Private Sub Text1_Change(Index As Integer)
    If Index = 4 Then HaCambiadoCP = True 'CPostal ha cambiado
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 4 Then HaCambiadoCP = False
    'If Index <> 22 Then ConseguirFoco Text1(Index), Modo
    If Not EsCampoMemo(Index) Then ConseguirFoco Text1(Index), Modo
End Sub

Private Function EsCampoMemo(Indice As Integer) As Boolean
    EsCampoMemo = False
    If Indice = 22 Or Indice = 54 Then EsCampoMemo = True
End Function


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If EsCampoMemo(Index) And KeyCode = 40 Then 'Flecha abajo
        Me.SSTab1.Tab = 1
        PonerFoco Text1(Index)
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not EsCampoMemo(Index) Then KEYpress KeyAscii
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
Dim campo As String
Dim Codigo As String
Dim tabla As String
Dim Titulo As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
        Case 1
            If Text1(2).Text = "" Then Text1(2).Text = Text1(1).Text
            
        Case 4 'CPostal
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, campo)
                Text1(Index + 2).Text = campo
             End If
             VieneDeBuscar = False
        
        Case 7 'NIF
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                ValidarNIF Text1(Index).Text
                
            End If
        
        Case 9 'Codigo de Actividad
            campo = "nomactiv"
            Codigo = "codactiv"
            tabla = "sactiv"
            Titulo = "Actividades"
            
        Case 10 'Código de Envío
            campo = "nomenvio"
            Codigo = "codenvio"
            tabla = "senvio"
            Titulo = "Formas de Envío"
            
         Case 11 'Código de zona
            campo = "nomzonas"
            Codigo = "codzonas"
            tabla = "szonas"
            Titulo = "Zonas de Clientes"
                       
         Case 12 'Código de Rutas
             campo = "nomrutas"
             Codigo = "codrutas"
             tabla = "srutas"
             Titulo = "Rutas de Asistencia"

        Case 22 'Observaciones
            If Modo = 3 Or Modo = 4 Then 'Insertando o modificando
                'si se pierde el foco con un TAB y pasaria al siguiente campo que
                'esta en la otra pestaña. si movemos foco a otro campo de la
                'misma pestaña no cambiamos
                If Screen.ActiveControl.Name = "Text1" Then
                    If Screen.ActiveControl.Index = 23 Then
                        Me.SSTab1.Tab = 1
                        PonerFoco Text1(23)
                    End If
                End If
            End If

         Case 23 'Codigo Formas de pago
            campo = "nomforpa"
            tabla = "sforpa"
            Codigo = "codforpa"
            Titulo = "Forma de Pago"
            
        Case 24, 25 'Descuento Pronto Pago, Descuento General
                'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(Index), 4
            
        Case 31, 32 'codbanco, sucursal
            PonerFormatoEntero Text1(Index)
            
        Case 35 'Cuenta contable
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            
        Case 36 'Codigo Agente Comercial
            campo = "nomagent"
            tabla = "sagent"
            Codigo = "codagent"
            Titulo = "Agente Comercial"
            
        Case 37 'Codigo Tarifa
            campo = "nomlista"
            Codigo = "codlista"
            tabla = "starif"
            Titulo = "Tarifa"
                                    
        Case 13, 40, 41, 48, 53 'Fecha alta, Fecha último mov.,fecha reclamación solicredito
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 42 'Código Situación
            campo = "nomsitua"
            Codigo = "codsitua"
            tabla = "ssitua"
            Titulo = "Situación"
            
        Case 43, 47, 49 'Límite Crédito , solicitado y riesgo actual
            'Formato tipo 1: Decimal(12,2)
            If Text1(Index).Text <> "" Then
                If Not PonerFormatoDecimal(Text1(Index), 1) Then Text1(Index).Text = ""
            End If
        Case 44 'Distancia Km
'            PonerFormatoDecimal Text1(Index), 5
            PonerFormatoEntero Text1(Index)
            
        
        
        Case 52
            If Modo = 1 Then Exit Sub
            'Buscara direcciones envio
            'sdirenvio nomdiren  coddiren
            campo = "nomdiren"
            tabla = "sdirenvio"
            Codigo = "codclien = " & Val(Text1(0).Text) & " AND coddiren "
            Titulo = "Direccion envio"
        
    End Select
    
    If (Index >= 9 And Index <= 12) Or Index = 23 Or Index = 36 Or Index = 37 Or Index = 42 Or Index = 52 Then
        If PonerFormatoEntero(Text1(Index)) Then
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, tabla, campo, Codigo, Titulo)
            If Text2(Index).Text = "" Then
                PonerFoco Text1(Index)
                If Index = 52 Then Text1(Index).Text = ""
            End If
            
        Else
            Text2(Index).Text = ""
        End If
        
        
        
    End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then
            If CadB <> "" Then CadB = CadB & " AND "
            CadB = CadB & " codagent = " & vUsu.CodigoAgente
        End If
    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia2 CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia2(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case Val(Me.imgBuscar(0).Tag)
        Case 5  'Cuenta Contable
            'Se llama a Busqueda desde el campo Cuenta contable
            '#A MANO: Porque busca en la tabla cuentas
            'de la base de datos de Contabilidad
            cad = cad & "Código|cuentas|codmacta|T||30·Denominacion|cuentas|nommacta|T||70·"
            tabla = "cuentas"
            Titulo = "Cuentas Contables"
            Conexion = conConta    'Conexión a BD: Conta
            
            
        Case 1000
            'Departamento en RENTING  Marzo 2012
            cad = cad & "Código|sdirec|coddirec|N||30·Denominacion|sdirec|nomdirec|T||70·"
            tabla = "sdirec"
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Departamentos"
            Else
                Titulo = "Direccion"
            End If
            Conexion = conAri    'Conexión a BD: Ariges
            
        Case Else   'Registro de la tabla de cabeceras: sartic
            cad = cad & ParaGrid(Text1(0), 10, "Código")
            cad = cad & ParaGrid(Text1(1), 50, "Nombre")
            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
            tabla = "sclipot"
            Titulo = "Clientes potenciales"
            Conexion = conAri    'Conexión a BD: Ariges
    End Select
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
        frmB.vCargaFrame = (Conexion = 2)
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
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
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        
        PonerCampos
        'CargaFrameDirec2 0   'los dos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(9).Text = PonerNombreDeCod(Text1(9), conAri, "sactiv", "nomactiv")
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "senvio", "nomenvio")
    Text2(11).Text = PonerNombreDeCod(Text1(11), conAri, "szonas", "nomzonas")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "srutas", "nomrutas")
    
    
    
    BloquearChecks Me, Modo
    
    CargaLineas True, 2

    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
   
    CargaDatosLWCRM
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub


'Private Sub PonerCamposDirecciones()
'Dim X As Boolean
'
'    If Data2.Recordset.EOF Then Exit Sub
'
'    X = PonerCamposFormaFrame(Me, "Text3", Data2)
'
'
'    Me.txtZona(14).Text = ""
'    If Text3(14).Text <> "" Then
'        txtZona(14).Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text3(14).Text, "N")
'    End If
'    '-- Esto permanece para saber donde estamos
''    lblIndicador.Caption = Data2.Recordset.AbsolutePosition & " de " & Data2.Recordset.RecordCount
'End Sub
'
'
'Private Sub PonerCamposDireccionesEnvio()
'Dim X As Boolean
'
'    If Data3.Recordset.EOF Then Exit Sub
'
'    X = PonerCamposFormaFrame(Me, "Text4", Data3)
'
'    Me.txtZona(10).Text = ""
'    If Text4(10).Text <> "" Then
'        txtZona(10).Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text4(10).Text, "N")
'    End If
'
'    '-- Esto permanece para saber donde estamos
''    lblIndicador.Caption = Data2.Recordset.AbsolutePosition & " de " & Data2.Recordset.RecordCount
'End Sub




'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diversos campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Long
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    BuscaChekc = ""
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        Me.cmdRegresar.Caption = "&Regresar"
    Else
        cmdRegresar.visible = False
    End If
    
     'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, CLng(NumReg)
    
         
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
   
    
'    'Bloquear los Text3
'    For I = 0 To Me.Text3.Count - 1
'        BloquearTxt Me.Text3(I), Not (Modo = 5)
'    Next I
        
    'Bloquear los Text3
'    If vParamAplic.DireccionesEnvio Then
'        For I = 0 To Me.Text4.Count - 1
'            BloquearTxt Me.Text4(I), Not (Modo = 6)
'        Next I
'
'
'        'Si tiene direcciones de envio y el modo=4 entonces esta habilitado
'        BloquearTxt Me.Text1(52), Not (Modo = 1 Or Modo = 4)
'
'    End If
            
    'Bloquear los Text3
    If Modo < 7 Then
        For i = 0 To Me.txtauxDC.Count - 1
            BloquearTxt Me.txtauxDC(i), True
        Next i
    End If
    
    

        
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    'cboAlbaran.Enabled = b
    'cboFacturacion.Enabled = b
    'cboTipoIVA.Enabled = b
    'cboTipocliente.Enabled = b
    'If vParamAplic.Renting Then cboFraRenting.Enabled = b
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    b = Modo <> 0 And Modo <> 2 And Modo < 5
    Me.imgFecha(0).Enabled = b
   
    
    
    For i = 0 To 3
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(9).Enabled = Me.imgBuscar(3).Enabled
    imgBuscar(11).Enabled = Modo >= 2 And Modo < 5
    
    'CRM
    cmdAccCRM(0).visible = Modo = 2
    cmdAccCRM(1).visible = False  'Modo = 2
    
    
    '-----------------------------
    'cmdActRiesgo.visible = Modo = 2 And vUsu.Nivel = 0

    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opcines de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
                        
                        
    'El listview
    If Modo <> 2 Then
       ' lw1.ListItems.Clear
        lwCRM.ListItems.Clear
    End If

                        
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2 Or Modo = 0 Or (Modo >= 5 And ModificaLineas = 0))
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 Or (Modo >= 5 And ModificaLineas = 0))
    'Los que sean AGENTES no pueden entrar
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Lineas Direcciones/Departamentos
    b = Modo = 2
    Toolbar1.Buttons(10).Enabled = b
    
    If vParamAplic.DireccionesEnvio Then Toolbar1.Buttons(11).Enabled = b
    Toolbar1.Buttons(12).Enabled = b 'Datos contacto
    If vParamAplic.Renting Then Toolbar1.Buttons(13).Enabled = b  'Datos contacto
    
    
    
    '-----------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    
    
    b = Modo = 2 Or Modo = 0
    Toolbar1.Buttons(13).Enabled = b
    Toolbar1.Buttons(14).Enabled = b
    Toolbar1.Buttons(15).Enabled = b
'    'BARRA DE DIRECCIONES
'    Me.ToolAux.visible = (Modo <> 0)
'    If Me.ToolAux.visible Then Me.ToolAux.visible = (Me.Data2.Recordset.RecordCount > 0)
'    If Me.ToolAux.visible Then
'        b = Not (Modo = 5 And (ModoFrame2 = 3 Or ModoFrame2 = 4))
'        Me.ToolAux.Buttons(1).Enabled = b
'        Me.ToolAux.Buttons(2).Enabled = b
'        Me.ToolAux.Buttons(3).Enabled = b
'        Me.ToolAux.Buttons(4).Enabled = b
'    End If
'
'    If vParamAplic.DireccionesEnvio Then
'            Me.Toolaux2.visible = (Modo <> 0)
'            If Me.Toolaux2.visible Then Me.Toolaux2.visible = (Me.Data3.Recordset.RecordCount > 0)
'            If Me.Toolaux2.visible Then
'                b = Not (Modo = 6 And (ModoFrame2 = 3 Or ModoFrame2 = 4))
'                Me.Toolaux2.Buttons(1).Enabled = b
'                Me.Toolaux2.Buttons(2).Enabled = b
'                Me.Toolaux2.Buttons(3).Enabled = b
'                Me.Toolaux2.Buttons(4).Enabled = b
'            End If
'    End If
    
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoFrame(Kmodo As Byte, ModoGral As Byte)
Dim i As Byte
On Error GoTo EPonerModoFr

    ModoFrame2 = Kmodo
    PonerModo ModoGral
    
    Select Case ModoGral
    Case 5
'
'        If ModoFrame2 = 0 Then
'
'            If Data2.Recordset.RecordCount > 5 Then
'                I = 5
'            Else
'                I = Data2.Recordset.RecordCount
'            End If
'            DesplazamientoVisible Me.ToolAux, 1, True, I
'        Else
'            DesplazamientoVisible Me.Toolbar1, btnPrimero, False, 1
'
'        End If
    Case 6
'        If ModoFrame2 = 0 Then
'            If Data3.Recordset.RecordCount > 5 Then
'                I = 5
'            Else
'                I = Data3.Recordset.RecordCount
'            End If
'            DesplazamientoVisible Me.Toolaux2, 1, True, I
'        Else
'            DesplazamientoVisible Me.Toolbar1, btnPrimero, False, 1
'
'        End If
        
    End Select
    
    'Bloquear TextBox sino modo 3 o 4
    Select Case ModoGral
'    Case 5
'        For I = 0 To Me.Text3.Count - 1
'            If ModoFrame2 = 3 Then Text3(I).Text = ""
'            BloquearTxt Text3(I), (ModoFrame2 = 0)
'        Next I
'        If ModoFrame2 = 4 Then BloquearTxt Text3(0), True
'
'        imgBuscar(15).Enabled = ModoFrame2 > 0
'    Case 6
'        'direnvio
'        For I = 0 To Me.Text4.Count - 1
'            If ModoFrame2 = 3 Then Text4(I).Text = ""
'            BloquearTxt Text4(I), (ModoFrame2 = 0)
'        Next I
'        If ModoFrame2 = 4 Then BloquearTxt Text4(0), True
'        imgBuscar(16).Enabled = ModoFrame2 > 0
    Case 7
        'Perosna de contacto
        For i = 0 To Me.txtauxDC.Count - 1
            If ModoFrame2 = 3 Then txtauxDC(i).Text = ""
            BloquearTxt txtauxDC(i), (ModoFrame2 = 0)
        Next i
       
       
       imgBuscar(14).visible = ModoFrame2 > 0
       Me.cboCargo.visible = ModoFrame2 > 0
       
     Case 8
'        'Perosna de contacto
'        For I = 0 To Me.txtauxRent.Count - 1
'            If ModoFrame2 = 3 Then txtauxRent(I).Text = ""
'            'Campos SIEMPRE BLOQUEADOS
'            If I = 0 Or I = 2 Then
'                BloquearTxt txtauxRent(I), True
'            Else
'                BloquearTxt txtauxRent(I), (ModoFrame2 = 0)
'            End If
'        Next I
'
'
'       cmdRenting(0).visible = ModoFrame2 > 0
'       cmdRenting(1).visible = ModoFrame2 > 0
'       cmdRenting(2).visible = ModoFrame2 > 0
'       Me.DataGrid2.Enabled = ModoFrame2 = 0
    End Select
    
    'Indice del prismatico del codpostal
    i = 10
    If ModoGral = 6 Then i = 12
    Select Case ModoFrame2
        Case 0  'MODO INICIAL
            'Me.imgBuscar(i).Enabled = False
            PonerBotonCabecera True
        Case 3, 4 'Modo INSERTAR o MODIFICAR
            '3=Insertar,  4=Modificar
            'Me.imgBuscar(i).Enabled = True
            If Modo = 3 Then
                If ModoGral = 5 Then
                '    PonerFoco Text3(0)
                Else
                '    PonerFoco Text4(0)
                End If
            End If
            PonerBotonCabecera False
    End Select

EPonerModoFr:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLineaVisible(bol As Boolean)
'bol=true : Se pone visible el frame ArticulosxAlmacen
'bol=false : se pone visible Datos Articulos
'On Error Resume Next
'
'    Me.frameComercial.visible = Not bol
'
'    Me.Label1(37).visible = Not bol 'Web
'    Me.Text1(8).visible = Not bol
'
'    Me.Label1(5).visible = Not bol 'Cod Actividad
'    Me.imgBuscar(0).visible = Not bol
'    Me.Text1(9).visible = Not bol
'    Me.Text2(0).visible = Not bol
'
'    Me.Label1(6).visible = Not bol 'Cod. Envío
'    Me.imgBuscar(1).visible = Not bol
'    Me.Text1(10).visible = Not bol
'    Me.Text2(1).visible = Not bol
'
'    Me.Label1(7).visible = Not bol 'Cod. Zona
'    Me.imgBuscar(2).visible = Not bol
'    Me.Text1(11).visible = Not bol
'    Me.Text2(2).visible = Not bol
'
'    Me.Label1(17).visible = Not bol 'Cod Ruta
'    Me.imgBuscar(3).visible = Not bol
'    Me.Text1(12).visible = Not bol
'    Me.Text2(3).visible = Not bol
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim fec As Date

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
       
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If
    If Not b Then Exit Function
    



    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function DatosOkLinea() As Boolean
    DatosOkLinea = False
    Select Case Modo
'    Case 5
'        DatosOkLinea = DatosOkLineaDpto
'    Case 6
'        DatosOkLinea = DatosOkLineaEnvio
    Case 7
    
       
        
        'En el text2 opongo el combo
        txtauxDC(2).Text = cboCargo.Text
        'Para datos personales SOLO necesito el nombre
        If Trim(txtauxDC(0).Text) = "" Then
            MsgBox "Nombre obligatorio", vbExclamation
        Else
            DatosOkLinea = True
        End If
        
'    Case 8
'        'renting
'         'desde el 2
'        For NumRegElim = 3 To Me.txtauxRent.Count - 1
'            If NumRegElim <> 10 And NumRegElim <> 11 Then '7= ult fecha factura
'                If Me.txtauxRent(NumRegElim).Text = "" Then
'                        MsgBox "Campos obligatorios", vbExclamation
'                        PonerFoco txtauxRent(NumRegElim)
'                        Exit Function
'                End If
'            End If
'        Next
'        'Si pone coddirec, tiene que existir nomdirec
'        If Me.txtauxRent(1).Text = "" Xor txtauxRent(2).Text = "" Then
'            MsgBox "Error departamento/direccion", vbExclamation
'            Exit Function
'        End If
'
'        'Comprobaremos que la linea que ha puesto no es mayor que uno ya facturado
'        BuscaChekc = DevuelveDesdeBD(conAri, "max(ultfec)", "sclipotrenting", "codclien", CStr(Data1.Recordset!codclien))
'        If BuscaChekc <> "" Then
'            If CDate(txtauxRent(4).Text) >= CDate(BuscaChekc) Then
'                If MsgBox("Peridodo no facturado.No se generara factura. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
'            End If
'            BuscaChekc = ""
'
'        End If
'
'
'
'        DatosOkLinea = True
        
        
    End Select
End Function

'Private Function DatosOkLineaDpto() As Boolean
'Dim b As Boolean
'Dim devuelve As String
'Dim I As Integer
'
'On Error GoTo EDatosOkLinea
'
'    DatosOkLineaDpto = False
'    b = True
'    devuelve = ""
'    'Campo Nombre Direc./Dpto
'    If Text3(1).Text = "" Then devuelve = devuelve & vbCrLf & "-Nombre"
'
'    'Campo Domicilio Direc./Dpto
'    If Text3(2).Text = "" Then devuelve = devuelve & vbCrLf & "-Domicilio"
'
'    'Campo CPostal Direc./Dpto
'    If Text3(3).Text = "" Then devuelve = devuelve & vbCrLf & "-C.Postal"
'
'    'Campo Población Direc./Dpto
'    If Text3(4).Text = "" Then devuelve = devuelve & vbCrLf & "-Población"
'
'    'Campo Provincia Direc./Dpto
'    If Text3(5).Text = "" Then devuelve = devuelve & vbCrLf & "-Provincia"
'
'    'Campo ZONA
'    If Text3(14).Text = "" Then devuelve = devuelve & vbCrLf & "-ZONA "
'
'    If devuelve <> "" Then
'        devuelve = "Campos vacios: " & vbCrLf & devuelve
'        MsgBox devuelve, vbExclamation
'        devuelve = ""
'        Exit Function
'    End If
'
'
'
'    'Comprobamos  si ya existe Si estamos insertando
'    'conAri: conexion a BD Ariges
'    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "coddirec", "codclien", Text1(0).Text, "N", , "coddirec", Text3(0).Text, "N")
'    'If ModificaLineas = 1 And DevuelveExisteEnBD(conAri, "sdirec", "codclien", Text1(0).Text, "N", "coddirec", Text3(0).Text, "N") Then
'    If ModificaLineas = 1 And devuelve <> "" Then
'        b = False
'        devuelve = DevuelveTextoDepto(False)
'        devuelve = "Ya existe" & devuelve & " del Cliente: " & vbCrLf
'        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
'        MsgBox devuelve, vbExclamation
'    End If
'
'
'    'comprobar los datos de la cuenta bancaria si param. de departamentos
'    If Me.FrameCtaBanDpto.visible And b Then
'        'Validar que la cuenta bancaria es correcta
'        For I = 10 To 13
'            If Text3(I).Text <> "" Then
'                If IsNumeric(Text3(I).Text) Then
'                    If Val(Text3(I).Text) = "0" Then Text3(I).Text = ""
'                End If
'            End If
'        Next
'
'
'        If Text3(13).Text <> "" Then
'            'Ha puesto codbanco
'
'                For I = 11 To 13
'                    If Text3(I).Text = "" Then Exit For
'                Next
'                If I <= 13 Then
'                    'Se ha salido
'                    MsgBox "Faltan datos para la cuenta bancaria", vbExclamation
'                    b = False
'                Else
'                    b = Comprueba_CuentaBan2(Text3(10).Text & Text3(11).Text & Text3(12).Text & Text3(13).Text, True)
'                    If Not b Then
'                        If MsgBox("Cuenta bancaria incorrecta.   ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then b = True
'                    End If
'                End If
'        End If
'
'
'
'
'    End If
'
'
'
'
'
'
'    DatosOkLineaDpto = b
'
'EDatosOkLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Function
'
'
'Private Function DatosOkLineaEnvio() As Boolean
'Dim devuelve As String
'On Error GoTo EDatosOkLinea
'
'    DatosOkLineaEnvio = False
'
'    devuelve = ""
'    'Campo Nombre Direc./Dpto
'    If Text4(1).Text = "" Then devuelve = devuelve & "    -Nombre "
'
'    'Campo Domicilio Direc./Dpto
'    If Text4(2).Text = "" Then devuelve = devuelve & "    -Domicilio"
'
'    'Campo CPostal Direc./Dpto
'    If Text4(3).Text = "" Then devuelve = devuelve & "    -C.Postal "
'
'    'Campo Población Direc./Dpto
'    If Text4(4).Text = "" Then devuelve = devuelve & "    -Población"
'
'    'Campo Provincia Direc./Dpto
'    If Text4(5).Text = "" Then devuelve = devuelve & "    -Provincia"
'
'    If Text4(10).Text = "" Then devuelve = devuelve & "    -Zona"
'
'    If devuelve <> "" Then
'        MsgBox "Campos no pueden ser nulos: " & vbCrLf & devuelve, vbExclamation
'        Exit Function
'    End If
'
'    'Comprobamos  si ya existe Si estamos insertando
'    'conAri: conexion a BD Ariges
'    devuelve = DevuelveDesdeBDNew(conAri, "sdirenvio", "coddiren", "codclien", Text1(0).Text, "N", , "coddiren", Text4(0).Text, "N")
'    If ModificaLineas = 1 And devuelve <> "" Then
'        devuelve = "Ya existe la direccion de envio del Cliente: " & vbCrLf
'        devuelve = devuelve & "Codigo: " & Text4(0).Text & vbCrLf
'        MsgBox devuelve, vbExclamation
'        Exit Function
'    End If
'
'
'    DatosOkLineaEnvio = True
'
'EDatosOkLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Function


'
'Private Sub Text3_Change(Index As Integer)
'    If Index = 3 Then HaCambiadoCP = True
'End Sub
'
'Private Sub Text3_GotFocus(Index As Integer)
'    If Index = 3 Then HaCambiadoCP = False
'    ConseguirFoco Text3(Index), 3
'End Sub
'
'Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub
'
'Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = 13 Then 'ENTER
'        KeyAscii = 0
'        If (Index = 9 And Me.FrameCtaBanDpto.visible = False) Or Index = 13 Then
'            PonerFocoBtn Me.cmdAceptar
'        Else
'            SendKeys "{tab}"
'        End If
'    End If
'    If Err.Number <> 0 Then Err.Clear
'End Sub
'
'
'Private Sub Text3_LostFocus(Index As Integer)
'Dim devuelve As String
'
'    On Error Resume Next
'
'    If Not PerderFocoGnralLineas(Text3(Index), ModificaLineas) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    Select Case Index
'        Case 0 'Codigo Direc/Dpto
'            If Trim(Text3(Index).Text) = "" Then Exit Sub
'            FormateaCampo Text3(Index)
'
'        Case 3 'Cod. Postal
'            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
'                Text3(Index + 1).Text = ObtenerPoblacion(Text3(Index).Text, devuelve)
'                Text3(Index + 2).Text = devuelve
'            End If
'            VieneDeBuscar = False
'
'        Case 10, 11 'codbanco, sucursal
'            PonerFormatoEntero Text3(Index)
'
'        Case 12, 13 'DC, cta banco
'            FormateaCampo Text3(Index)
'            If Index = 13 Then PonerFocoBtn Me.cmdAceptar
'
'        Case 14
'            devuelve = ""
'            If PonerFormatoEntero(Text3(Index)) Then
'                devuelve = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text3(Index).Text, "N")
'                If devuelve = "" Then
'                    MsgBox "No existe la zona", vbExclamation
'                    Text3(Index).Text = ""
'                    PonerFoco Text3(Index)
'                End If
'            Else
'                Text3(Index).Text = ""
'            End If
'            Me.txtZona(Index).Text = devuelve
'    End Select
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
'
'
'
''Text4    Direnvio
'Private Sub Text4_Change(Index As Integer)
'    If Index = 3 Then HaCambiadoCP = True
'End Sub
'
'Private Sub Text4_GotFocus(Index As Integer)
'    If Index = 3 Then HaCambiadoCP = False
'    ConseguirFoco Text4(Index), 3
'End Sub
'
'Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub
'
'Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
'
'    If KeyAscii = 13 Then 'ENTER
'
'        If Index <> 9 Then
'            KeyAscii = 0
'            SendKeys "{tab}"
'        Else
'            PonerFocoBtn cmdAceptar
'        End If
'    End If
'
'End Sub
'
'
'Private Sub Text4_LostFocus(Index As Integer)
'Dim devuelve As String
'
'    On Error Resume Next
'
'    If Not PerderFocoGnralLineas(Text4(Index), ModificaLineas) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    Select Case Index
'        Case 0 'Codigo Direc/Dpto
'            If Trim(Text4(Index).Text) = "" Then Exit Sub
'            FormateaCampo Text4(Index)
'
'        Case 3 'Cod. Postal
'            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
'                Text4(Index + 1).Text = ObtenerPoblacion(Text4(Index).Text, devuelve)
'                Text4(Index + 2).Text = devuelve
'            End If
'            VieneDeBuscar = False
'        Case 8
'            'PonerFocoBtn cmdAceptar
'
'        Case 10
'            If PonerFormatoEntero(Text4(Index)) Then
'                devuelve = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text4(Index).Text, "N")
'                If devuelve = "" Then
'                    MsgBox "No existe la zona", vbExclamation
'                    Text4(Index).Text = ""
'                    PonerFoco Text4(Index)
'                End If
'            Else
'                Text4(Index).Text = ""
'            End If
'            Me.txtZona(Index).Text = devuelve
'    End Select
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
'
'
'
'
'
'Private Sub ToolAux_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If Modo < 2 Or Modo = 3 Then Exit Sub
'    Select Case Button.Index
'        Case 1 To 4 'Flechas Desplazamiento
'            DesplazamientoLineas (Button.Index - 1), 0
'        Case 6
'            frmObraListado.Opcion = 2
'            frmObraListado.Show vbModal
'    End Select
'End Sub
'
'Private Sub Toolaux2_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If Modo < 2 Or Modo = 3 Then Exit Sub
'    Select Case Button.Index
'        Case 1 To 4 'Flechas Desplazamiento
'            DesplazamientoLineas (Button.Index - 1), 1
'        Case 6
'            frmObraListado.Opcion = 2
'            frmObraListado.Show vbModal
'    End Select
'End Sub
'


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 6  'Nuevo
           mnNuevo_Click
        Case 7  'Modificar
           mnModificar_Click
        Case 8  'Borrar
           mnEliminar_Click
           
        Case 10
            If Modo <> 2 Then Exit Sub
            If Data1.Recordset.EOF Then Exit Sub
            
            PasarACliente
        Case 11, 12
            BotonDirecciones Button.Index - 5   'sera 5 o 6
        
        Case 13, 14, 15
            If Modo = 2 Or Modo = 0 Then
                frmListadoOfer.NumCod = ""
                If Button.Index = 15 Then
                    If Text1(0).Text <> "" Then frmListadoOfer.NumCod = Format(Val(Text1(0).Text), "0000") & "|" & Text1(1).Text & "|"
                End If
                frmListadoOfer.OpcionListado = 400 + (Button.Index - 13)
                frmListadoOfer.Show vbModal
            End If
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

'
Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


'Private Sub CargarComboAlbaran()
''### Combo Valorar Albaran con
''Cargaremos el combo, o bien desde una tabla o con valores fijos o como
''se quiera, la cuestion es cargarlo
'' El estilo del combo debe de ser 2 - Dropdown List
'' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
'' o marcamos la opcion sorted del combo
''0-Todo, 1-Cantidad y Precio, 2-Cantidad
'
'    cboAlbaran.Clear
'    cboAlbaran.AddItem "Todo"
'    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 0
'
'    cboAlbaran.AddItem "Cantidad y Precio"
'    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 1
'
'    cboAlbaran.AddItem "Cantidad"
'    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 2
'
'End Sub


'Private Sub CargarComboFacturacion()
''### Combo Tipo Facturación
''Cargaremos el combo, o bien desde una tabla o con valores fijos o como
''se quiera, la cuestion es cargarlo
'' El estilo del combo debe de ser 2 - Dropdown List
'' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
'' o marcamos la opcion sorted del combo
''0-Factura Colectiva, 1-Factura x Albaran
'
'    cboFacturacion.Clear
'    cboFacturacion.AddItem "Factura Colectiva"
'    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0
'
'    cboFacturacion.AddItem "Factura x Albaran"
'    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1
'
'End Sub
'
'
'Private Sub CargarComboTipoIVA()
''### Combo Tipo de IVA a Aplicar
''Cargaremos el combo, o bien desde una tabla o con valores fijos o como
''se quiera, la cuestion es cargarlo
'' El estilo del combo debe de ser 2 - Dropdown List
'' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
'' o marcamos la opcion sorted del combo
''0-Normal, 1-Con Recargo de Equivalencia, 2-Exento de IVA
'
'    Me.cboTipoIVA.Clear
'    cboTipoIVA.AddItem "Normal"
'    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 0
'
'    cboTipoIVA.AddItem "Recargo Equivalencia"
'    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 1
'
'    cboTipoIVA.AddItem "Exento de IVA"
'    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 2
'
'    cboTipoIVA.AddItem "Intracomunitario"
'    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 3
'
'    'Junio 2012 Reducido
'    cboTipoIVA.AddItem "Reducido"
'    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 4
'
'End Sub

Private Function InsertarModificarLinea() As Boolean
    Select Case Modo
'    Case 5
'        InsertarModificarLinea = InsertarModificarLineaDpto
'    Case 6
'        InsertarModificarLinea = InsertarModificarLineaEnvio
    Case 7
        InsertarModificarLinea = InsertarModificarLineaDatosConctacto
'    Case 8
'        InsertarModificarLinea = InsertarModificarLineaRenting
    End Select
    If InsertarModificarLinea Then
        Me.Refresh
        Espera 0.25
    End If
End Function
    
'Private Function InsertarModificarLineaDpto() As Boolean
'Dim I As Byte
'Dim SQL As String
'
'    On Error GoTo EInsertarModificarLinea
'
'    InsertarModificarLineaDpto = False
'    SQL = ""
'    Select Case ModificaLineas
'    Case 1  'INSERTAR
'        If DatosOkLinea Then
'            SQL = "INSERT INTO sdirec (codclien,coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba) VALUES ("
'            SQL = SQL & Text1(0).Text & ", "
'            SQL = SQL & Text3(0).Text
'            For I = 1 To 5
'                SQL = SQL & ", "
'                SQL = SQL & DBSet(Text3(I).Text, "T")
'            Next I
'
'            For I = 6 To 13 'campos opcionales
'                SQL = SQL & ", "
'                SQL = SQL & DBSet(Text3(I).Text, "T", "S")
''                If i <> 13 Then SQL = SQL & ", "
'            Next I
'
'            SQL = SQL & ")"
'        End If
'
'    Case 2  'MODIFICAR
'        If DatosOkLinea Then
'            SQL = "UPDATE sdirec Set nomdirec = " & DBSet(Text3(1).Text, "T")
'            SQL = SQL & ", domdirec = " & DBSet(Text3(2).Text, "T")
'            SQL = SQL & ", codpobla = " & DBSet(Text3(3).Text, "T")
'            SQL = SQL & ", pobdirec = " & DBSet(Text3(4).Text, "T")
'            SQL = SQL & ", prodirec = " & DBSet(Text3(5).Text, "T")
'            SQL = SQL & ", perdirec = " & DBSet(Text3(6).Text, "T")
'            'If Text3(7).Text <> "" Then SQL = SQL & ", fechainv = '" & Format(Text3(7).Text, "yyyy-mm-dd") & "'"
'            'If Text3(8).Text <> "" Then SQL = SQL & ", horainve = '" & Format(Text3(8).Text, "hh:mm:ss") & "'"
'            SQL = SQL & ", teldirec = " & DBSet(Text3(7).Text, "T")
'            SQL = SQL & ", faxdirec = " & DBSet(Text3(8).Text, "T")
'            SQL = SQL & ", maidirec = " & DBSet(Text3(9).Text, "T")
'            'datos cuenta bancaria
'            If Me.FrameCtaBanDpto.visible Then
'                SQL = SQL & ", codbanco = " & DBSet(Text3(10).Text, "N", "S")
'                SQL = SQL & ", codsucur = " & DBSet(Text3(11).Text, "N", "S")
'                SQL = SQL & ", digcontr = " & DBSet(Text3(12).Text, "T")
'                SQL = SQL & ", cuentaba = " & DBSet(Text3(13).Text, "T")
'            End If
'            SQL = SQL & ", codzona = " & DBSet(Text3(14).Text, "N", "S")
'            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
'            SQL = SQL & " coddirec =" & (Text3(0).Text)
'        End If
'    End Select
'
'    If SQL <> "" Then
'        conn.Execute SQL
'        InsertarModificarLineaDpto = True
'        TratarDptoEnTesoreria   'TESOERIA
'    Else
'        PonerFoco Text3(1)
'    End If
'    Exit Function
'EInsertarModificarLinea:
'    MuestraError Err.Number, "Insertar/Modificar Direcciones/Departamentos" & vbCrLf & Err.Description
'End Function
'
'
'
'Private Function InsertarModificarLineaEnvio() As Boolean
'Dim I As Byte
'Dim SQL As String
'
'    On Error GoTo EInsertarModificarLinea
'
'    InsertarModificarLineaEnvio = False
'    SQL = ""
'    Select Case ModificaLineas
'    Case 1  'INSERTAR
'        If DatosOkLinea Then
'            SQL = "INSERT INTO sdirenvio (codclien,coddiren,nomdiren,domdiren,codpobla,pobdiren,prodiren,perdiren,teldiren,faxdiren,observa,codzona) VALUES ("
'            SQL = SQL & Text1(0).Text & ", "
'            SQL = SQL & Text4(0).Text
'            For I = 1 To 5
'                SQL = SQL & ", "
'                SQL = SQL & DBSet(Text4(I).Text, "T")
'            Next I
'
'            For I = 6 To 9 'campos opcionales
'                SQL = SQL & ", "
'                SQL = SQL & DBSet(Text4(I).Text, "T", "S")
''                If i <> 13 Then SQL = SQL & ", "
'            Next I
'            SQL = SQL & "," & DBSet(Text4(10).Text, "N", "S")
'            SQL = SQL & ")"
'        End If
'
'    Case 2  'MODIFICAR
'        If DatosOkLinea Then
'            SQL = "UPDATE sdirenvio Set nomdiren = " & DBSet(Text4(1).Text, "T")
'            SQL = SQL & ", domdiren = " & DBSet(Text4(2).Text, "T")
'            SQL = SQL & ", codpobla = " & DBSet(Text4(3).Text, "T")
'            SQL = SQL & ", pobdiren = " & DBSet(Text4(4).Text, "T")
'            SQL = SQL & ", prodiren = " & DBSet(Text4(5).Text, "T")
'            SQL = SQL & ", perdiren = " & DBSet(Text4(6).Text, "T")
'            SQL = SQL & ", teldiren = " & DBSet(Text4(7).Text, "T")
'            SQL = SQL & ", faxdiren = " & DBSet(Text4(8).Text, "T")
'            SQL = SQL & ", observa = " & DBSet(Text4(9).Text, "T")
'            SQL = SQL & ", codzona = " & DBSet(Text4(10).Text, "N", "S")
'            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
'            SQL = SQL & " coddiren =" & (Text4(0).Text)
'        End If
'    End Select
'
'    If SQL <> "" Then
'        conn.Execute SQL
'        InsertarModificarLineaEnvio = True
'    Else
'        PonerFoco Text4(1)
'    End If
'    Exit Function
'EInsertarModificarLinea:
'    MuestraError Err.Number, "Insertar/Modificar Direcciones de envio" & vbCrLf & Err.Description
'End Function

Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "&Cabecera"
    If b Then
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
        
    If b Then
        
'        If Modo = 5 Then
'            Me.lblIndicador.Caption = "Lineas Detalle"
'            If Not Data2.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & ": " & Me.Data2.Recordset.AbsolutePosition & " de " & Me.Data2.Recordset.RecordCount
'        ElseIf Modo = 6 Then
'            Me.lblIndicador.Caption = "Lineas direnvio:"
'            If Not Data3.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & Me.Data3.Recordset.AbsolutePosition & " de " & Me.Data3.Recordset.RecordCount
'        ElseIf Modo = 7 Then
            Me.lblIndicador.Caption = "Datos contacto"
'        Else
'            Me.lblIndicador.Caption = "Renting"
'        End If
    End If
End Sub


'Private Sub MostrarSituacion(vMostrar As Boolean)
'Dim codigo As Integer
'Dim Bloquea As String
'Dim DescBloqueo As String
'
'    On Error GoTo EMostrarSitu
'
'    If Data1.Recordset.EOF Then Exit Sub
'    If vMostrar Then
'        codigo = Data1.Recordset!codsitua
'        If Not IsNull(codigo) Then
'            Me.lblSituacion.visible = (codigo <> 0)
'            Me.Frame1(1).visible = (codigo <> 0)
'            If Not (codigo = 0) Then
'            'Si situacion=0 (activo) no mostrar situacion
'                Bloquea = DevuelveDesdeBDNew(conAri, "ssitua", "tipositu", "codsitua", CStr(codigo), "N")
'                DescBloqueo = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", CStr(codigo), "N")
'                If Val(Bloquea) = 0 Then
'                    'Cliente NO Bloqueado
'                    Me.lblSituacion.Caption = UCase(DescBloqueo)
'                    Me.lblSituacion.ForeColor = vbBlue
'                Else
'                    'Cliente Bloqueado
'                    Me.lblSituacion.Caption = "BLOQUEADO POR: " & UCase(DescBloqueo)
'                    Me.lblSituacion.ForeColor = vbRed
'                End If
'            End If
'        End If
'    Else
'        Me.lblSituacion.visible = False
'        Me.Frame1(1).visible = False
'    End If
'EMostrarSitu:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub


Private Sub PosicionarData()
Dim Indicador As String, cad As String

    cad = "(codclien=" & Val(Text1(0).Text) & ")"
    If SituarData(Data1, cad, Indicador) Then
'       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
    PonerModo 2
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE  codclien= " & Val(Text1(0).Text)
End Function

'Cual 0.- Los dos (si parametros son dos)   1. Solo dpto    2. Solo envio
'Private Sub CargaFrameDirec2(Cual As Byte)
'    If Cual < 2 Then CargaFrame_Direc
'    If vParamAplic.DireccionesEnvio And Cual <> 1 Then CargaFrame_DirecEnv
'End Sub



'Private Sub CargaFrame_Direc()
'Dim cadCli As String
'
'    'Crear las lineas de Direcciones/Departamentos para el cliente
'    'ASignamos un SQL al DATA2
'    Me.Data2.ConnectionString = conn
'    If Text1(0).Text = "" Then
'        cadCli = -1
'    Else
'        cadCli = Val(Text1(0).Text)
'    End If
'    Data2.RecordSource = "Select * from sdirec where codclien = " & cadCli & ";"
'    Data2.Refresh
'
'    cadCli = "0"
'    If Data2.Recordset.RecordCount > 0 Then
'        If Data2.Recordset.RecordCount > 1 Then cadCli = "2"
'        Data2.Recordset.MoveFirst
'        PonerCamposDirecciones
'    Else
'        LimpiarCamposDirecciones2 False
'    End If
'    PonerModoOpcionesMenu
'
'
'
'    DesplazamientoVisible Me.ToolAux, 1, True, CByte(cadCli)
'End Sub
'
'
'Private Sub CargaFrame_DirecEnv()
'Dim cadCli As String
'
'    'Crear las lineas de Direcciones/Departamentos para el cliente
'    'ASignamos un SQL al DATA2
'    Me.Data3.ConnectionString = conn
'    If Text1(0).Text = "" Then
'        cadCli = -1
'    Else
'        cadCli = Val(Text1(0).Text)
'    End If
'    Data3.RecordSource = "Select * from sdirenvio where codclien = " & cadCli & " ORDER BY coddiren;"
'    Data3.Refresh
'
'
'    If Data3.Recordset.RecordCount > 0 Then
'        Data3.Recordset.MoveFirst
'        PonerCamposDireccionesEnvio
'    Else
'        LimpiarCamposDirecciones2 True
'    End If
'    PonerModoOpcionesMenu
'    DesplazamientoVisible Me.Toolaux2, 1, True, Data3.Recordset.RecordCount
'End Sub

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
'    With Me.Toolbar2
'        .ImageList = frmPpal.ImgListPpal
'        .Buttons(1).Image = 5
'        .Buttons(3).Image = 6
'        .Buttons(5).Image = 7
'        .Buttons(7).Image = 8
'        .Buttons(9).Image = 1
'        .Buttons(11).Image = 12
'    End With
    
'    Set lw1.SmallIcons = frmPpal.ImgListPpal
    
  
    'If vParamAplic.TieneCRM Then
    
        With Me.Toolbar3
            .ImageList = frmPpal.ImgListPpal
            .Buttons(1).Image = 3
            .Buttons(3).Image = 30
            .Buttons(5).Image = 25
            .Buttons(7).Image = 13
            .Buttons(9).Image = 31
            .Buttons(11).Image = 32
            .Buttons(13).Image = 33
            '.Buttons(1).visible = False
        End With
        
        Set lwCRM.SmallIcons = frmPpal.ImgListPpal
        
   ' End If
    
    
    'Direcciones envio (NO es la solapa de departamento / direccion
'    SSTab1.TabVisible(3) = vParamAplic.DireccionesEnvio
'    With Me.Toolaux2
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 6 'primero
'        .Buttons(2).Image = 7 'Anterior
'        .Buttons(3).Image = 8 'Siguiente
'        .Buttons(4).Image = 9 'Último
'        .Buttons(6).Image = 16 'Último
'    End With
    
End Sub



'Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

'
'    If Button.Tag = "" Then Exit Sub
'    LabelDoc.Caption = ""
'    'Levantamos todos los botones y dejamos pulsado el de ahora
'    For NumRegElim = 1 To Toolbar2.Buttons.Count
'        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
'            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
'        End If
'    Next NumRegElim
'    CargaColumnas CByte(Button.Tag)
'
'    'Hacemos las acciones
'    If Modo = 2 Then CargaDatosLWDoc
'End Sub













'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'  CRM
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "" Then Exit Sub
    LabelCRM.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar3.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar3.Buttons(NumRegElim).Index <> Button.Index Then Toolbar3.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnasCRM CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWCRM
End Sub





Private Sub CargaColumnasCRM(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader
Dim Ordena As Integer

    'EN POTENCIALES solo utilizo el 1(llamadas) y el 6


    'Las llamadas cogera las llamadas recibidas desde sllama y las efectuadas desde acciones comerciales con tipoaccion=1
    'para poder ordenarlas tendremos una columna viiblefalse con yyymmddhhmmss
    Ordena = -1
    Select Case OpcionList
    Case 0
        'Acciones comerciales
        LabelCRM.Caption = "Acciones comerciales"
        
        Columnas = "Fecha|Usuario|Estado|Medio|Tipo|Descripcion|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|1200|1200|800|2300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||0000||"
        Ncol = 6
               
    Case 1
        'Llamadas
        LabelCRM.Caption = "Llamadas "
        
        Columnas = "Fecha|Usuario|Tipo/Trab|Observaciones|Orden|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|0|4000|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 5
    
        Ordena = 5
        
    Case 2
        LabelCRM.Caption = "E-mail"
        Columnas = "Fecha|Enviado|Email|Asunto|Adj|entryID|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1800|825|2565|3899|495|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm||||||"
        Ncol = 6
    
    Case 3
        'COBROS
        LabelCRM.Caption = "Cobros pendientes"
        Columnas = "Fecha Vto.|Factura|Fecha factura|Forma pago|Pendiente|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|1300|2400|1495|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|1|0|0|1|"
        'Formatos
        Formato = "dd/mm/yyyy||dd/mm/yyyy||" & FormatoImporte & "|"
        Ncol = 5
        
    Case 4
        'COBROS
        LabelCRM.Caption = "Observaciones departamento"
        Columnas = "Departamento|Fecha|Observaciones||"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|6500|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|||"
        Ncol = 4
        
        
    Case 5
        'Reclamaciones
        LabelCRM.Caption = "Reclamaciones"
        Columnas = "Fecha|Factura|Observaciones|Importe|codigo|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|4500|1500|0|"  'La ultima esta oculta
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy|||" & FormatoImporte & "||"
        Ncol = 5
        
    
    Case 6
        'H I S T O R I A L
        LabelCRM.Caption = "Historial"
        Columnas = "Fecha|Usuario|Trabajador|Observaciones|"
        Ancho = "2300|1000|0|5200|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 4
        
    
    End Select
    
    
    cmdAccCRM(2).visible = OpcionList = 4 'Or OpcionList = 6
    lwCRM.ColumnHeaders.Clear
    
    'Guardo la opcion en el tag
    lwCRM.Tag = OpcionList & "|" & Ncol & "|"
    
    
    
    For NumRegElim = 1 To Ncol
         Set C = lwCRM.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
    
    If Ordena < 0 Then
        lwCRM.Sorted = False
    Else
        lwCRM.Sorted = True
        lwCRM.SortKey = 4
        lwCRM.SortOrder = lvwDescending
    End If
    
End Sub







Private Sub CargaDatosLWCRM()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & LabelCRM.Caption
    lblIndicador.Refresh
    CargaDatosLWcrm2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLWcrm2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Kopc As Byte
Dim MeteIT As Boolean
Dim ConexionConta As Boolean  'Si no es conta es ARIGES( conn)
Dim Reemplazar As Boolean
    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar3.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
   

    'EL where del codclien     lo lleva cada sql
    Kopc = CByte(RecuperaValor(lwCRM.Tag, 1))
    ConexionConta = False
    Select Case Kopc
'    Case 0
'        'Acciones comerciales
'        Cad = "select fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion from scrmaccionclipot,scrmtipo WHERE scrmacciones.tipo= scrmtipo.codigo "
'        Cad = Cad & " and codclien=" & Data1.Recordset!codclien & " and tipo > 20"  'las 20 primerasprobablemebne no sepongan aqui
'        GroupBy = ""
'        BuscaChekc = "fechora"
    Case 1
        'Llamadas
        cad = "select fechora ,usuario,'' nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmaccionclipot "   'left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmaccionclipot.tipo=1  and codclipot= " & Data1.Recordset!codClien   '2 DE historial
        GroupBy = ""
        BuscaChekc = "fechora"

'    Case 2
'
'        'eMAIL
'        Cad = "select fechahora, if(enviado=1,""Enviado"",""Recibido""),email,asunto,"
'        Cad = Cad & "if(adjuntos<>"""",""*"","""") ,entryID from scrmmail"
'        Cad = Cad & " WHERE codclien=" & Data1.Recordset!codclien
'        GroupBy = ""
'        BuscaChekc = "fechahora"
'
'    Case 3
'        'Cobros pendientes
'        Cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",codfaccl),7)),fecfaccl,nomforpa,"
'        Cad = Cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
'        Cad = Cad & " FROM  scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
'        Cad = Cad & " WHERE scobro.codmacta = '" & Text1(35).Text & "' "
'
'        'PARA TEINSA
'        If vParamAplic.Frecuencias Then Cad = Cad & " AND (sforpa.tipforpa between 0 and 3) "
'        BuscaChekc = "fecvenci"
'        ConexionConta = True
'
'    Case 4
'        'Observaciones departamento
'        Cad = "select if(dpto=1,""Administracion"",if(dpto=2,""Comercial"",if(dpto=3,""SAT"",""Dirección""))),fecha,observa,dpto from scrmobsclipot"
'        Cad = Cad & " WHERE codclien=" & Data1.Recordset!codclien
'        BuscaChekc = "dpto"
'
'    Case 5
'        'Reclamaciones
'        'Cobros pendientes
'        Cad = "select fecreclama,concat(numserie,right(concat(""00000000"",codfaccl),7)),observaciones,impvenci,codigo"
'        Cad = Cad & " from shcocob where codmacta='" & Text1(35).Text & "' "
'        BuscaChekc = "fecreclama desc ,codigo "
'        ConexionConta = True
'
'
    Case 6
        'Historial
        cad = "select fechora ,usuario,'' nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmaccionclipot "   'left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmaccionclipot.tipo=2  and codclipot= " & Data1.Recordset!codClien   '2 DE historial
        GroupBy = ""
        BuscaChekc = "fechora"
    End Select
    
    
    
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc
    If Kopc <> 4 Then cad = cad & " DESC"

    
    BuscaChekc = ""
    
    lwCRM.ListItems.Clear
   
    Set RS = New ADODB.Recordset
    If Not ConexionConta Then
        'Conn  ariges
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        'Va contra la contabilidad  connconta
        RS.Open cad, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    End If
    While Not RS.EOF
        If Kopc <> 3 Then
            MeteIT = True
        Else
            If RS!Tot <> 0 Then
                MeteIT = True
            Else
                MeteIT = False
            End If
        End If
        
        If MeteIT Then
                Set It = lwCRM.ListItems.Add()
                 
                If lwCRM.ColumnHeaders(1).Tag <> "" Then
                    It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
                Else
                    It.Text = RS.Fields(0)
                End If
                'El resto de cmpos
                For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
                    If IsNull(RS.Fields(NumRegElim - 1)) Then
                        It.SubItems(NumRegElim - 1) = " "
                    Else
                    
                        If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                            It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                        Else
                                                
                            Reemplazar = False
                            'Cad = RS.Fields(NumRegElim - 1)
                            cad = DBLetMemo(RS.Fields(NumRegElim - 1))
                            'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                            If NumRegElim = 4 And (Kopc = 1 Or Kopc = 6) Then Reemplazar = True
                            'para las observaciones de la reclamacion tb quito los vbcrlf
                            If NumRegElim = 3 And Kopc = 5 Then Reemplazar = True
                            
                            'Medio
                            If NumRegElim = 3 And Kopc = 0 Then DevuelveMedio cad
                            If NumRegElim = 3 And Kopc = 4 Then Reemplazar = True
                            If NumRegElim = 3 And Kopc = 4 Then Reemplazar = True
                            
                            If Reemplazar Then
                                cad = Replace(cad, vbCrLf, " ")
                                cad = Replace(cad, vbTab, "   ")
                            End If
                            It.SubItems(NumRegElim - 1) = cad
                        
                            
                            
                        End If
                    End If
                Next
                'El icono
                If Kopc = 1 Then
                    It.SmallIcon = 27
                ElseIf Kopc = 2 Then

                    If RS.Fields(1) = "Enviado" Then
                        It.SmallIcon = 28
                    Else
                        It.SmallIcon = 29
                    End If
                Else
                    'el resto ponemos el del toolbar
                    It.SmallIcon = ElIcono
                End If
        End If
        
        
    
        RS.MoveNext
    Wend
    RS.Close
    
    
    If Kopc = 1 Then
'        'Llamadas. Las efectuadas las hago desde este punto
'    '    cad = "select fechora ,usuario,'' nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
'    '    cad = cad & " scrmaccionclipot " 'left join straba on scrmacciones.codtraba=straba.codtraba "
'    '    cad = cad & " WHERE scrmaccionclipot.tipo=1  and codclien= " & Data1.Recordset!codclien
'        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF
'            '
'            'Coje datos desde dos tablas
'            Set It = lwCRM.ListItems.Add()
'            It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
'
'            For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
'                If IsNull(RS.Fields(NumRegElim - 1)) Then
'                    It.SubItems(NumRegElim - 1) = " "
'                Else
'
'                    If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
'                        It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
'                    Else
'
'
'                        cad = RS.Fields(NumRegElim - 1)
'                        'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
'                        If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
'
'                        It.SubItems(NumRegElim - 1) = cad
'
'
'
'                    End If
'                End If
'            Next
'            It.SmallIcon = 26
'            RS.MoveNext
'        Wend
'        RS.Close
    End If
    Set RS = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub

Private Sub DevuelveMedio(ByRef cad As String)
    'pendiente,en curso finalizada
    If cad = "0" Then
        cad = "Pendiente"
    ElseIf cad = "1" Then
        cad = "En curso"
    Else
        cad = "Finalizada"
    End If
End Sub


Private Sub LanzarProgramaEmails()
'Dim TieneDatosDpto As Boolean

    On Error GoTo ELanzarProgramaEmails

    If Dir(App.Path & "\AriOutlook.exe", vbArchive) = "" Then
        MsgBox "No tienen el programa de asignacion de mails al CRM de Ariadna", vbExclamation
        Exit Sub
    End If
    
    'TieneDatosDpto = False
   ' If Not Data2.Recordset Is Nothing Then
   '     If Not Data2.Recordset.EOF Then TieneDatosDpto = True
   ' End If
        
        
    'Como lanzamos el programa
    '*************************  dbariges|codclien|nombre||||mails que se utlizan|
   ' If TieneDatosDpto Then
   '     BuscaChekc = "Select distinct(maidirec) from sdirec where codclien=" & Data1.Recordset!codclien & " AND maidirec <>"""""
   '     Set miRsAux = New ADODB.Recordset
   '     miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   ' End If
    
    BuscaChekc = ""
    If Text1(17).Text <> "" Then BuscaChekc = BuscaChekc & Text1(17).Text & "|"  'mail1
    If Text1(18).Text <> "" Then BuscaChekc = BuscaChekc & Text1(18).Text & "|"  'mail1
        
        
    'If TieneDatosDpto Then
    '    While Not miRsAux.EOF
    '        If Not IsNull(miRsAux!maidirec) Then
    '            If miRsAux!maidirec <> "" Then BuscaChekc = BuscaChekc & miRsAux!maidirec & "|"
    '        End If
    '        miRsAux.MoveNext
    '    Wend
    '    miRsAux.Close
    'End If
    
    BuscaChekc = vUsu.CadenaConexion & "|" & Data1.Recordset!codClien & "|" & CStr(Data1.Recordset!NomClien) & "||||" & BuscaChekc
    
    Shell App.Path & "\AriOutlook.exe " & BuscaChekc, vbNormalFocus
    
    Espera 2
    
    
ELanzarProgramaEmails:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lanzar Programa Email"
    Set miRsAux = Nothing
    BuscaChekc = ""
End Sub






Private Sub CargaLineas(enlaza As Boolean, Cual As Byte)   'cual=0  percontac, 1:  renting   , 2 los dos
Dim SQL As String


        If Cual <> 1 Then
            SQL = "SELECT nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id,codclien FROM sclipotdp where codclien = "
            
                       SQL = "SELECT nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id,codclien FROM sclipotdp where codclien = "
            
            If enlaza Then
                SQL = SQL & Text1(0).Text
                
            Else
                SQL = SQL & " -1"
            End If
            SQL = SQL & " ORDER BY  id"
            CargaGridGnral DataGrid1, Me.data4, SQL, True
            SQL = "S|txtauxDC(0)|T|Nombre|3600|;S|txtauxDC(1)|T|Departamento|2600|;"
            'Los campos que no se ven que van FUERA DEL GRID
            SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla SQL, DataGrid1, Me
            DataGrid1.ScrollBars = dbgAutomatic
        End If
        
       ' If vParamAplic.Renting Then
            If Cual <> 0 Then
       '         SQL = "SELECT id,sclipotrenting.coddirec,nomdirec,referencia,fecalta,numcuotas,fecbaja,importe"
       '         SQL = SQL & ",sclipotrenting.codtipco,nomtipco,obser,ultfec"
       '         SQL = SQL & " from (sclipotrenting left join sdirec on sclipotrenting.codclien=sdirec.codclien"
       '         SQL = SQL & " and sdirec.coddirec=sclipotrenting.coddirec ) "
       '         SQL = SQL & " inner join stipco on stipco.codtipco=sclipotrenting.codtipco"
       '         SQL = SQL & " WHERE sclipotrenting.codclien = "
       '         If enlaza Then
       '             SQL = SQL & Text1(0).Text
       '
       '         Else
       '             SQL = SQL & " -1"
       '         End If
       '         SQL = SQL & " ORDER BY  id"
       '         CargaGridGnral DataGrid2, Me.data5, SQL, True
       '
       '         SQL = "S|txtauxRent(0)|T|ID|600|;"
       '         If vParamAplic.HayDeparNuevo = 1 Then
       '             SQL = SQL & "S|txtauxRent(1)|T|Dpto|600|"
       '         Else
       '             SQL = SQL & "S|txtauxRent(1)|T|Dir.|600|"
       '         End If
       '         SQL = SQL & ";S|cmdRenting(0)|B||0|;S|txtauxRent(2)|T|Departamento|2950|;"
       '         SQL = SQL & "S|txtauxRent(3)|T|Referencia|1600|;S|txtauxRent(4)|T|Fecha alta|1300|;S|cmdRenting(1)|B||0|;"
       '         SQL = SQL & "S|txtauxRent(5)|T|Cuotas|650|;S|txtauxRent(6)|T|Fecha baja|1300|;S|cmdRenting(2)|B||0|;"
       '         SQL = SQL & "S|txtauxRent(7)|T|Importe|1050|;"
       '         'no se ven
       '         SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;"
       '         arregla SQL, DataGrid2, Me
       '         DataGrid1.ScrollBars = dbgAutomatic
       '         'Como el lo pone a la derecha
       '         txtauxRent(1).Alignment = 0 'a la izda
            End If
        
       ' End If
        
        
End Sub


Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim i As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (data4.Recordset Is Nothing) Then
            If Not data4.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For i = 0 To txtauxDC.Count - 1
            txtauxDC(i).Text = ""
        Next i
        
    Else
        'EL
        
        PonerCamposFormaFrame Me, "txtauxDC", data4
        
        
    End If
End Sub


'Private Sub PonerDatosForaGridRent(ForzarLimpiar As Boolean)
'Dim I As Integer
'Dim Limp As Boolean
'
'    Limp = True
'    If Not ForzarLimpiar Then
'        If Not (data5.Recordset Is Nothing) Then
'            If Not data5.Recordset.EOF Then Limp = False
'        End If
'    End If
'
'
'    If Limp Then
'
'        'Limpiamos
'        For I = 8 To txtauxRent.Count - 1
'            txtauxRent(I).Text = ""
'        Next I
'
'    Else
'        'EL
'
'        PonerCamposFormaFrame Me, "txtauxRent", data5
'
'
'    End If
'End Sub







Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim b As Boolean

    ModificaLineas = xModo
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    b = Modo = 7 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas


    DeseleccionaGrid Me.DataGrid1
    
    txtauxDC(0).Height = DataGrid1.RowHeight
    txtauxDC(0).visible = b
    txtauxDC(0).Top = alto
    txtauxDC(1).Height = DataGrid1.RowHeight
    txtauxDC(1).visible = b
    txtauxDC(1).Top = alto
    SituarCboCargo
End Sub







Private Function InsertarModificarLineaDatosConctacto() As Boolean
Dim i As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    'codclien,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id FROM sclipotdp
    InsertarModificarLineaDatosConctacto = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO sclipotdp (codclien,nombre,dpto,cargo,telefono,ext,movil,maidirec,observa,id) VALUES ("
            SQL = SQL & Text1(0).Text

                    
            For i = 0 To 7 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(txtauxDC(i).Text, "T", "S")
            Next i
            SQL = SQL & ", " & txtauxDC(8).Text & ")"
  
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            'codclien,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id
            SQL = "UPDATE sclipotdp Set nombre = " & DBSet(txtauxDC(0).Text, "T")
            SQL = SQL & ", dpto = " & DBSet(txtauxDC(1).Text, "T", "S")
            SQL = SQL & ", cargo = " & DBSet(txtauxDC(2).Text, "T", "S")
            SQL = SQL & ", telefono = " & DBSet(txtauxDC(3).Text, "T", "S")
            SQL = SQL & ", ext = " & DBSet(txtauxDC(4).Text, "T", "S")
            SQL = SQL & ", movil  = " & DBSet(txtauxDC(5).Text, "T", "S")
            SQL = SQL & ", maidirec= " & DBSet(txtauxDC(6).Text, "T", "S")
            SQL = SQL & ", observa = " & DBSet(txtauxDC(7).Text, "T", "S")
            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
            SQL = SQL & " id =" & (txtauxDC(8).Text)
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaDatosConctacto = True
    Else
        PonerFoco txtauxDC(0)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar datos contacto" & vbCrLf & Err.Description
End Function



Private Sub txtauxDC_GotFocus(Index As Integer)
    If Index <> 7 Then ConseguirFoco txtauxDC(Index), 3
End Sub

Private Sub txtauxDC_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index <> 7 Then
            KeyAscii = 0
            SendKeys "{tab}"
        Else
            PonerFocoBtn cmdAceptar
        End If
    End If
End Sub

Private Sub txtauxDC_LostFocus(Index As Integer)
    'Si quisieramos comprobar algo
    txtauxDC(Index).Text = Trim(txtauxDC(Index).Text)
End Sub


Private Sub BotonEliminarLineaContacto()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String
Dim i As Integer

    If data4.Recordset.EOF Then Exit Sub
    If data4.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    
    cad = "¿Seguro que desea eliminar el contacto?"
    cad = cad & vbCrLf & "Nombre:  " & data4.Recordset!Nombre
    cad = cad & vbCrLf & "Departamento:  " & DBLet(data4.Recordset!Dpto, "T")
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = data4.Recordset.AbsolutePosition
        data4.Recordset.Delete
        CargaLineas True, 0
        
        PonerDatosForaGrid False
            
        ModificaLineas = 0
        PonerModoFrame 0, 7
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        data4.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


'Private Sub BotonEliminarRenting()
''Eliminar una linea De ArticulosxAlmacen
'Dim Cad As String
'
'
'    If data5.Recordset.EOF Then Exit Sub
'    If data5.Recordset.RecordCount < 1 Then Exit Sub
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'    ModificaLineas = 3 'Eliminar
'
'
'    Cad = "¿Seguro que desea eliminar el elemento ?"
'    Cad = Cad & vbCrLf & "ID:  " & data5.Recordset!Id
'    If Not IsNull(data5.Recordset!CodDirec) Then Cad = Cad & vbCrLf & "Departamento:  " & DBLet(data5.Recordset!CodDirec, "T") & " " & DBLet(data5.Recordset!nomdirec, "T")
'    Cad = Cad & vbCrLf & "Referencia:  " & data5.Recordset!Referencia
'    Cad = Cad & vbCrLf & "Importe:  " & data5.Recordset!Importe
'    'Borramos
'    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        On Error GoTo Error2
'        Screen.MousePointer = vbHourglass
'        NumRegElim = data5.Recordset.AbsolutePosition
'        Cad = "DELETE FROM sclipotrenting where codclien = " & Text1(0).Text & " AND ID= " & data5.Recordset!Id
'        conn.Execute Cad
'        CargaLineas True, 1
'        PonerDatosForaGridRent False
'
'        ModificaLineas = 0
'        PonerModoFrame 0, 8
'    End If
'
'
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then
'        data5.Recordset.CancelUpdate
'        MsgBox Err.Number & ": " & Err.Description, vbExclamation
'    End If
'End Sub


'Private Sub CargaComboTipoCliente()
'    CargarCombo_Tabla Me.cboTipocliente, "stipclien", "tipclien", "descclien"
'End Sub

'Private Sub CargaComboFrarRenting()
'    cboFraRenting.Clear
'    cboFraRenting.AddItem "Mensual"
'    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 1
'
'    cboFraRenting.AddItem "Trimestral"
'    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 3
'
'    cboFraRenting.AddItem "Semestral"
'    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 6
'
'    cboFraRenting.AddItem "Anual"
'    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 12
'
'End Sub





Private Sub CargaComboCargos()

    cboCargo.Clear
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select * from scargoscli order by cargo", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'El prinmero vacio
    cboCargo.AddItem ""
    While Not miRsAux.EOF
        cboCargo.AddItem miRsAux!cargo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

Private Sub SituarCboCargo()
Dim i As Integer

    If data4.Recordset Is Nothing Then Exit Sub
    If data4.Recordset.EOF Then Exit Sub

    cboCargo.ListIndex = -1
    For i = 1 To cboCargo.ListCount - 1
        If cboCargo.List(i) = UCase(DBLet(data4.Recordset!cargo, "T")) Then
            cboCargo.ListIndex = i
            Exit For
        End If
    Next
End Sub




Private Function EliminarCliente() As Boolean
    On Error Resume Next
    EliminarCliente = True
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from whoexpedientepot WHERE codclien =" & Me.Text1(0).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Err.Number <> 0 Then
        'NO existe la tabla. NO pasa nada
        Err.Clear
        conn.Errors.Clear
    Else
        'Si que existe
        If Not miRsAux.EOF Then
            MsgBox "Expedientes relacionados", vbExclamation
            EliminarCliente = False
        End If
        miRsAux.Close
    End If
    Set miRsAux = Nothing
End Function


Private Sub PasarACliente()

    BuscaChekc = ""
    For kCampo = 3 To 12
        If kCampo <> 8 Then
            If Text1(kCampo).Text = "" Then BuscaChekc = BuscaChekc & " -" & Label1(kCampo).Caption & vbCrLf
        End If
    Next kCampo

    If BuscaChekc <> "" Then
        BuscaChekc = "Campos obligatorios" & vbCrLf & String(20, "=") & vbCrLf & vbCrLf & BuscaChekc
        MsgBox BuscaChekc, vbExclamation
        Exit Sub
    End If
        
    If Not Comprobar_NIF(Text1(7).Text) Then
        BuscaChekc = "El NIF: " & Text1(7).Text & "   no parece correcto. ¿Continuar?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    
    BuscaChekc = DevuelveDesdeBD(conAri, "concat(codclien,' ' ,nomclien)", "sclien", "nifclien", Text1(7).Text, "T")
    If BuscaChekc <> "" Then
        BuscaChekc = "El NIF ya existe para el cliente: " & BuscaChekc
        MsgBox BuscaChekc, vbExclamation
        Exit Sub
    End If

    'CadenaDesdeOtroForm = Format(Text1(0).Text, "0000") & " - " & Text1(1).Text
    'frmListado2.Opcion = 47
    'frmListado2.Show vbModal
    
    CadenaDesdeOtroForm = Text1(0).Text
    frmFacClientes.VerCliente = -1
    frmFacClientes.Show vbModal
    
    BuscaChekc = DevuelveDesdeBD(conAri, "codclien", "sclien", "nifclien", Text1(7).Text, "T")
    If BuscaChekc <> "" Then
        NumRegElim = Val(BuscaChekc)
        'Le paso los contactos
            'Si huberia o hubiesen metido mas contactos
        
        BuscaChekc = "select " & NumRegElim & ",id,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa  from sclipotdp WHERE codclien = " & Text1(0).Text
        BuscaChekc = "INSERT INTO scliendp(codclien,id,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa ) " & BuscaChekc
        ejecutar BuscaChekc, True

        MsgBox "Se ha creado con exito el cliente: " & NumRegElim, vbInformation
    End If
End Sub




