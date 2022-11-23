VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntAlbaranes2 
   BackColor       =   &H00AEAEAE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Clientes"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   13995
   Icon            =   "frmFacEntAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   126
      Text            =   "nom ccoste"
      Top             =   6435
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   330
      Index           =   16
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   57
      Text            =   "frmFacEntAlbaranes.frx":000C
      Top             =   6000
      Visible         =   0   'False
      Width           =   7005
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   5940
      Width           =   2175
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   15
         Top             =   30
         Width           =   2095
      End
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   120
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12600
      TabIndex        =   59
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11400
      TabIndex        =   58
      Top             =   6000
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   6480
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
      TabIndex        =   41
      Top             =   0
      Width           =   13995
      _ExtentX        =   24686
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Series"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Marcar facturar"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Albaran"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir albaran transporte"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
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
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   8400
         MaxLength       =   15
         TabIndex        =   103
         Text            =   "BASE IMP."
         Top             =   100
         Width           =   1490
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   102
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   960
      Top             =   6480
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
      Height          =   4500
      Left            =   120
      TabIndex        =   43
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1395
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmFacEntAlbaranes.frx":0049
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
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntAlbaranes.frx":0065
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2(1)"
      Tab(1).Control(1)=   "Text1(44)"
      Tab(1).Control(2)=   "Text2(43)"
      Tab(1).Control(3)=   "Text1(43)"
      Tab(1).Control(4)=   "FrameFacRec"
      Tab(1).Control(5)=   "chkDocArchi"
      Tab(1).Control(6)=   "Text1(41)"
      Tab(1).Control(7)=   "Text1(39)"
      Tab(1).Control(8)=   "FrameHco"
      Tab(1).Control(9)=   "Text1(29)"
      Tab(1).Control(10)=   "Text2(29)"
      Tab(1).Control(11)=   "Text1(28)"
      Tab(1).Control(12)=   "Text2(28)"
      Tab(1).Control(13)=   "Text1(27)"
      Tab(1).Control(14)=   "Text2(27)"
      Tab(1).Control(15)=   "Text1(2)"
      Tab(1).Control(16)=   "Text1(25)"
      Tab(1).Control(17)=   "Text1(26)"
      Tab(1).Control(18)=   "Text1(24)"
      Tab(1).Control(19)=   "Text1(23)"
      Tab(1).Control(20)=   "Text1(22)"
      Tab(1).Control(21)=   "Text1(21)"
      Tab(1).Control(22)=   "Text1(20)"
      Tab(1).Control(23)=   "Text1(19)"
      Tab(1).Control(24)=   "Text1(18)"
      Tab(1).Control(25)=   "Text1(38)"
      Tab(1).Control(26)=   "chkImpreso"
      Tab(1).Control(27)=   "Label1(63)"
      Tab(1).Control(28)=   "Label1(55)"
      Tab(1).Control(29)=   "Label1(54)"
      Tab(1).Control(30)=   "imgBuscar(13)"
      Tab(1).Control(31)=   "imgBuscar(9)"
      Tab(1).Control(32)=   "Shape1"
      Tab(1).Control(33)=   "imgFecha(40)"
      Tab(1).Control(34)=   "Label1(52)"
      Tab(1).Control(35)=   "Label1(24)"
      Tab(1).Control(36)=   "Label1(23)"
      Tab(1).Control(37)=   "imgBuscar(8)"
      Tab(1).Control(38)=   "Label1(9)"
      Tab(1).Control(39)=   "imgBuscar(7)"
      Tab(1).Control(40)=   "Label1(12)"
      Tab(1).Control(41)=   "Label1(11)"
      Tab(1).Control(42)=   "Label1(10)"
      Tab(1).Control(43)=   "Label1(5)"
      Tab(1).Control(44)=   "Label1(3)"
      Tab(1).Control(45)=   "Label1(45)"
      Tab(1).ControlCount=   46
      TabCaption(2)   =   "Fitosanitarios / Campos"
      TabPicture(2)   =   "frmFacEntAlbaranes.frx":0081
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameCampos"
      Tab(2).Control(1)=   "FrameManipulador"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Carta de portes"
      TabPicture(3)   =   "frmFacEntAlbaranes.frx":009D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Totales"
      TabPicture(4)   =   "frmFacEntAlbaranes.frx":00B9
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameFactura"
      Tab(4).ControlCount=   1
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -73560
         TabIndex        =   161
         Top             =   600
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   184
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   183
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   182
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   181
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   180
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   179
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   178
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   177
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   176
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   175
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   174
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   173
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   172
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   171
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   170
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   169
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   168
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   167
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   166
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   165
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   164
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   163
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
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
            Index           =   55
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   162
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   6120
            TabIndex        =   199
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   6
            Left            =   6960
            TabIndex        =   198
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   28
            Left            =   2760
            TabIndex        =   197
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   196
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   22
            Left            =   2160
            TabIndex        =   195
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   18
            Left            =   3960
            TabIndex        =   194
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   193
            Top             =   360
            Width           =   1215
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
            TabIndex        =   192
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
            TabIndex        =   191
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
            TabIndex        =   190
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   4800
            TabIndex        =   189
            Top             =   1200
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   8040
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
            Index           =   36
            Left            =   11880
            TabIndex        =   188
            Top             =   2160
            Width           =   135
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
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   39
            Left            =   4200
            TabIndex        =   187
            Top             =   2655
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4080
            TabIndex        =   186
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
            Height          =   255
            Index           =   42
            Left            =   2040
            TabIndex        =   185
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -68880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   160
         Text            =   "Text2"
         Top             =   675
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame FrameCampos 
         Caption         =   "Campos"
         Height          =   3495
         Left            =   -66720
         TabIndex        =   151
         Top             =   600
         Width           =   4695
         Begin VB.CommandButton cmdMtoCampos 
            Height          =   375
            Index           =   0
            Left            =   2520
            Picture         =   "frmFacEntAlbaranes.frx":00D5
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Añadir campo"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdMtoCampos 
            Height          =   375
            Index           =   1
            Left            =   4200
            Picture         =   "frmFacEntAlbaranes.frx":6927
            Style           =   1  'Graphical
            TabIndex        =   152
            ToolTipText     =   "Eliminar campo"
            Top             =   240
            Width           =   375
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   154
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4683
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Campo"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Partida"
               Object.Width           =   2937
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Variedad"
               Object.Width           =   2778
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Nuevo"
            Height          =   195
            Index           =   60
            Left            =   1920
            TabIndex        =   156
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label1 
            Caption         =   "Eliminar"
            Height          =   195
            Index           =   61
            Left            =   3600
            TabIndex        =   155
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Frame FrameManipulador 
         Caption         =   "Manipulador fitosanitarios  "
         Height          =   3495
         Left            =   -74760
         TabIndex        =   140
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   45
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   142
            Tag             =   "ManipuladorNumCarnet|T|S|||scaalb|ManipuladorNumCarnet||N|"
            Text            =   "123456789"
            Top             =   480
            Width           =   1815
         End
         Begin VB.Frame FrameMani2 
            Height          =   2295
            Left            =   240
            TabIndex        =   141
            Top             =   840
            Width           =   6495
            Begin VB.TextBox Text2 
               BackColor       =   &H80000018&
               Height          =   315
               Index           =   0
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   150
               Text            =   "Tiene bajo el text1 vinculado"
               Top             =   1650
               Width           =   4080
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   48
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   145
               Tag             =   "TipoCarnet|N|S|||scaalb|TipoCarnet||N|"
               Text            =   "123456789"
               Top             =   1680
               Width           =   255
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   47
               Left            =   1440
               MaxLength       =   15
               TabIndex        =   144
               Tag             =   "ManipuladorFecCaducidad|F|S|||scaalb|ManipuladorFecCaducidad||N|"
               Text            =   "123456789"
               Top             =   960
               Width           =   1110
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   46
               Left            =   1440
               MaxLength       =   60
               TabIndex        =   143
               Tag             =   "ManipuladorNombre|T|S|||scaalb|ManipuladorNombre||N|"
               Text            =   "123456789"
               Top             =   360
               Width           =   4455
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo carnet"
               Height          =   195
               Index           =   59
               Left            =   120
               TabIndex        =   149
               Top             =   1680
               Width           =   1245
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha caducidad"
               Height          =   195
               Index           =   58
               Left            =   120
               TabIndex        =   148
               Top             =   960
               Width           =   1245
            End
            Begin VB.Label Label1 
               Caption         =   "Nombre"
               Height          =   195
               Index           =   57
               Left            =   120
               TabIndex        =   147
               Top             =   360
               Width           =   690
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Carnet"
            Height          =   195
            Index           =   56
            Left            =   360
            TabIndex        =   146
            Top             =   480
            Width           =   690
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1440
            Picture         =   "frmFacEntAlbaranes.frx":7329
            ToolTipText     =   "Buscar cliente varios"
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   12
         Left            =   13200
         MaxLength       =   15
         TabIndex        =   138
         Text            =   "comision"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   44
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   37
         Tag             =   "ObIn|T|S|||scaalb|observacrm||N|"
         Top             =   3840
         Width           =   7485
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   43
         Left            =   -66360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   134
         Text            =   "Text2"
         Top             =   1440
         Width           =   3045
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   43
         Left            =   -67200
         MaxLength       =   30
         TabIndex        =   133
         Tag             =   "Cod. zona|N|S|0||scaalb|codzonas|000|N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   660
      End
      Begin VB.Frame FrameFacRec 
         Caption         =   "Datos Factura a rectificar "
         Height          =   615
         Left            =   -74760
         TabIndex        =   115
         Top             =   3720
         Width           =   7935
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   37
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   120
            Tag             =   "Tipo Mov. Factura|T|S|||scaalb|codtipmf||N|"
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   36
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   118
            Tag             =   "Nº. Factura|N|S|0||scaalb|numfactu|0000000|N|"
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   35
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   116
            Tag             =   "Fecha Factura|F|S|||scaalb|fecfactu|dd/mm/yyyy|N|"
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Mov."
            Height          =   255
            Index           =   47
            Left            =   1800
            TabIndex        =   121
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
            Height          =   255
            Index           =   46
            Left            =   3480
            TabIndex        =   119
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fact."
            Height          =   255
            Index           =   44
            Left            =   5640
            TabIndex        =   117
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox chkDocArchi 
         Caption         =   "Docu. archivado"
         Height          =   375
         Left            =   -63240
         TabIndex        =   31
         Tag             =   "Docar|N|N|||scaalb|docarchiv||N|"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   41
         Left            =   -63240
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "Fecha envio|F|S|||scaalb|fecenvio|dd/mm/yyyy|N|"
         Top             =   840
         Width           =   1185
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   12600
         MaxLength       =   15
         TabIndex        =   128
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
         Height          =   315
         Index           =   10
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   50
         Tag             =   "Bultos"
         Text            =   "12345"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   9
         Left            =   12360
         TabIndex        =   124
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
         Height          =   315
         Index           =   9
         Left            =   11640
         MaxLength       =   6
         TabIndex        =   56
         Tag             =   "proveedor"
         Text            =   "codc"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   39
         Left            =   -71160
         MaxLength       =   7
         TabIndex        =   123
         Tag             =   "Nº Venta|N|S|||scaalb|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.Frame FrameHco 
         Height          =   2055
         Left            =   -65880
         TabIndex        =   104
         Top             =   2160
         Width           =   4455
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   33
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   112
            Text            =   "Text2"
            Top             =   1560
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   33
            Left            =   135
            MaxLength       =   30
            TabIndex        =   111
            Text            =   "Text1"
            Top             =   1560
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   109
            Text            =   "Text2"
            Top             =   840
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   135
            MaxLength       =   30
            TabIndex        =   108
            Text            =   "Text1"
            Top             =   840
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   106
            Top             =   240
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1080
            Picture         =   "frmFacEntAlbaranes.frx":742B
            ToolTipText     =   "Buscar incidencia"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   113
            Top             =   1320
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmFacEntAlbaranes.frx":752D
            ToolTipText     =   "Buscar trabajador"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   110
            Top             =   615
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   105
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   29
         Left            =   -67200
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "Cod. Envío|N|N|0|999|scaalb|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   840
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   -66360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   840
         Width           =   3045
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   -73185
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Preparador Material|N|N|0|9999|scaalb|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   -72480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   1560
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -73185
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "Trabajador pedido|N|S|0|9999|scaalb|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1200
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -72480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   89
         Text            =   "Text2"
         Top             =   1200
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -72645
         MaxLength       =   10
         TabIndex        =   87
         Tag             =   "Semana Entrega|N|S|||scaalb|sementre||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -74880
         MaxLength       =   7
         TabIndex        =   84
         Tag             =   "Nº Pedido|N|S|||scaalb|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   -73905
         MaxLength       =   10
         TabIndex        =   83
         Tag             =   "Fecha Pedido|F|S|||scaalb|fecpedcl|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   79
         Tag             =   "Fecha Oferta|F|S|||scaalb|fecofert|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   -71160
         MaxLength       =   7
         TabIndex        =   78
         Tag             =   "Nº Oferta|N|S|||scaalb|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   52
         Tag             =   "OP"
         Text            =   "OF"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   1950
         Left            =   200
         TabIndex        =   63
         Top             =   315
         Width           =   13400
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   49
            Left            =   12360
            MaxLength       =   30
            TabIndex        =   23
            Tag             =   "Bultos|N|S|0|32000|scaalb|numbultos||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   950
         End
         Begin VB.CheckBox chkPideCliente 
            Caption         =   "Pedido por cliente"
            Height          =   330
            Left            =   11640
            TabIndex        =   157
            Tag             =   "P|N|N|||scaalb|PideCliente||N|"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox chkConTransporte 
            Caption         =   "Con transporte"
            Height          =   330
            Left            =   8880
            TabIndex        =   20
            Tag             =   "Trans|N|N|||scaalb|tipAlbaran||N|"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   42
            Left            =   7590
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   131
            Text            =   "Text2"
            Top             =   1200
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   42
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Dir envio|N|S|0|9999|scaalb|coddiren|000|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   40
            Left            =   4080
            MaxLength       =   7
            TabIndex        =   25
            Tag             =   "Descuento General|N|S|||scaalb|aportacion|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1020
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   34
            Left            =   12360
            MaxLength       =   30
            TabIndex        =   22
            Tag             =   "Cant. Km|N|S|0|99999|scaalb|cantidkm||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   950
         End
         Begin VB.CheckBox chkFacturarKm 
            Caption         =   "Facturar Km"
            Height          =   330
            Left            =   10320
            TabIndex        =   21
            Tag             =   "Facturar Km|N|N|||scaalb|facturkm||N|"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Referencia Cliente|T|S|||scaalb|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   1725
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7590
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   75
            Tag             =   "Direccion/Dpto.|T|S|||scaalb|nomdirec||N|"
            Text            =   "Text2"
            Top             =   165
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaalb|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scaalb|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1209
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scaalb|codpobla||N|"
            Text            =   "Text15"
            Top             =   861
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scaalb|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   861
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3195
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scaalb|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   165
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scaalb|nifclien||N|"
            Text            =   "123456789"
            Top             =   165
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Cod. Agente|N|N|0|9999|scaalb|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   513
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7590
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   70
            Text            =   "Text2"
            Top             =   513
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Forma de Pago|N|N|0|999|scaalb|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   861
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7590
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   65
            Text            =   "Text2"
            Top             =   861
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalb|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   8280
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scaalb|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   540
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            ItemData        =   "frmFacEntAlbaranes.frx":762F
            Left            =   11640
            List            =   "frmFacEntAlbaranes.frx":7631
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scaalb|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   513
            Width           =   4030
         End
         Begin VB.Label Label1 
            Caption         =   "Bultos"
            Height          =   255
            Index           =   62
            Left            =   11160
            TabIndex        =   158
            Top             =   600
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":7633
            ToolTipText     =   "Dirección envio"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Dir. envio"
            Height          =   255
            Index           =   53
            Left            =   5700
            TabIndex        =   132
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "APORTACION TERMINAL"
            Height          =   255
            Index           =   49
            Left            =   3480
            TabIndex        =   125
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Km a facturar"
            Height          =   255
            Index           =   43
            Left            =   11160
            TabIndex        =   114
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   82
            Top             =   1560
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   855
            Picture         =   "frmFacEntAlbaranes.frx":7735
            ToolTipText     =   "Buscar población"
            Top             =   867
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   5700
            TabIndex        =   77
            Top             =   165
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":7837
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   165
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   76
            Top             =   1209
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   74
            Top             =   861
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   73
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   72
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            Picture         =   "frmFacEntAlbaranes.frx":7939
            ToolTipText     =   "Buscar cliente varios"
            Top             =   165
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5700
            TabIndex        =   71
            Top             =   513
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":7A3B
            ToolTipText     =   "Buscar agente"
            Top             =   516
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   69
            Top             =   861
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.Pago"
            Height          =   255
            Index           =   25
            Left            =   5700
            TabIndex        =   68
            Top             =   1610
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   7560
            TabIndex        =   67
            Top             =   1635
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   11160
            TabIndex        =   66
            Top             =   960
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":7B3D
            ToolTipText     =   "Buscar forma de pago"
            Top             =   867
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   64
            Top             =   513
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   62
         ToolTipText     =   "Buscar artículo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   61
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
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
         TabIndex        =   48
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
         Height          =   315
         Index           =   8
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   55
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
         Height          =   315
         Index           =   7
         Left            =   9960
         MaxLength       =   30
         TabIndex        =   54
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
         Height          =   315
         Index           =   6
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   53
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
         Height          =   315
         Index           =   4
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   51
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
         Height          =   315
         Index           =   3
         Left            =   6120
         MaxLength       =   16
         TabIndex        =   49
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
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   47
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
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   46
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   36
         Tag             =   "Observación 5|T|S|||scaalb|observa05||N|"
         Top             =   3480
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   35
         Tag             =   "Observación 4|T|S|||scaalb|observa04||N|"
         Top             =   3120
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   34
         Tag             =   "Observación 3|T|S|||scaalb|observa03||N|"
         Top             =   2760
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   33
         Tag             =   "Observación 2|T|S|||scaalb|observa02||N|"
         Top             =   2400
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación 1|T|S|||scaalb|observa01||N|"
         Top             =   2040
         Width           =   7485
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntAlbaranes.frx":7C3F
         Height          =   2040
         Left            =   195
         TabIndex        =   60
         Top             =   2325
         Width           =   13400
         _ExtentX        =   23627
         _ExtentY        =   3598
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   -70080
         MaxLength       =   10
         TabIndex        =   122
         Tag             =   "Nº terminal|N|S|||scaalb|numtermi||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.CheckBox chkImpreso 
         Caption         =   "Impreso"
         Height          =   375
         Left            =   -70080
         TabIndex        =   28
         Tag             =   "Impr|N|N|||scaalb|albImpreso||N|"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Puntos"
         Height          =   255
         Index           =   63
         Left            =   -68880
         TabIndex        =   159
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   495
         Index           =   55
         Left            =   -74880
         TabIndex        =   136
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. zona"
         Height          =   195
         Index           =   54
         Left            =   -67200
         TabIndex        =   135
         Top             =   1200
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   -66240
         Picture         =   "frmFacEntAlbaranes.frx":7C54
         ToolTipText     =   "Buscar forma de envio"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -66240
         Picture         =   "frmFacEntAlbaranes.frx":7D56
         ToolTipText     =   "Buscar forma de envio"
         Top             =   600
         Width           =   240
      End
      Begin VB.Shape Shape1 
         Height          =   1395
         Left            =   -67320
         Top             =   480
         Width           =   5895
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   40
         Left            =   -62280
         Picture         =   "frmFacEntAlbaranes.frx":7E58
         ToolTipText     =   "Buscar fecha"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha envio"
         Height          =   195
         Index           =   52
         Left            =   -63240
         TabIndex        =   129
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Envío"
         Height          =   195
         Index           =   24
         Left            =   -67200
         TabIndex        =   94
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Preparador Material"
         Height          =   255
         Index           =   23
         Left            =   -74880
         TabIndex        =   92
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -73455
         Picture         =   "frmFacEntAlbaranes.frx":7EE3
         ToolTipText     =   "Buscar trabajador"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   90
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -73440
         Picture         =   "frmFacEntAlbaranes.frx":7FE5
         ToolTipText     =   "Buscar trabajador"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   12
         Left            =   -72645
         TabIndex        =   88
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   86
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   10
         Left            =   -73905
         TabIndex        =   85
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   -70200
         TabIndex        =   81
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   -71160
         TabIndex        =   80
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74880
         TabIndex        =   45
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12570
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   95
      Top             =   480
      Width           =   13695
      Begin VB.CheckBox chkFacturar 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Tag             =   "Facturar|N|N|||scaalb|factursn||N|"
         Top             =   330
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   760
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7305
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scaalb|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   495
         Width           =   4080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scaalb|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   120
         Width           =   760
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   7305
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   100
         Text            =   "Text2"
         Top             =   120
         Width           =   4080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   30
         Left            =   1275
         TabIndex        =   1
         Tag             =   "Tipo Albaran|T|N|||scaalb|codtipom||S|"
         Text            =   "Text3"
         Top             =   345
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Albaran|F|N|||scaalb|fechaalb|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Albaran|N|S|0||scaalb|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   345
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Facturar"
         Height          =   195
         Left            =   3720
         TabIndex        =   139
         Top             =   390
         Width           =   585
      End
      Begin VB.Label LblMostr 
         BackStyle       =   0  'Transparent
         Caption         =   "MOSTR."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   11640
         TabIndex        =   137
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6180
         Picture         =   "frmFacEntAlbaranes.frx":80E7
         ToolTipText     =   "Buscar cliente"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   5115
         TabIndex        =   101
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Realizada Por"
         Height          =   255
         Index           =   21
         Left            =   5115
         TabIndex        =   99
         Top             =   165
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   6180
         Picture         =   "frmFacEntAlbaranes.frx":81E9
         ToolTipText     =   "Buscar trabajador"
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fec. Alb."
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   98
         Top             =   150
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2835
         Picture         =   "frmFacEntAlbaranes.frx":82EB
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   97
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   96
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   14
      Left            =   3720
      Picture         =   "frmFacEntAlbaranes.frx":8376
      ToolTipText     =   "Ver ampliación"
      Top             =   6000
      Width           =   240
   End
   Begin VB.Label lblF 
      BackStyle       =   0  'Transparent
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
      Left            =   9720
      TabIndex        =   130
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Centro coste"
      Height          =   255
      Index           =   51
      Left            =   2400
      TabIndex        =   127
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
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
Attribute VB_Name = "frmFacEntAlbaranes2"
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

Private WithEvents frmC As frmBasico2 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1

'Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios   MOni, no deja nuevo en clivarios
Private WithEvents frmCV As frmFacClientesV  'Form Mto Clientes Varios
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

Private WithEvents frmProv As frmBasico2 '%=%=frmComProveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1


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
    'KEYpress KeyAscii
    PonerFocoBtn cmdAceptar
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
                    
                    PosicionarData True
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea(numlinea, False) Then
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    
                    If PrimeraLin Then
                        
                        
                        If hcoCodTipoM = "DEV" Then
                            txtAnterior = ""
                            If ImporteFormateado(Text1(16).Text) <> Data1.Recordset!DtoGnral Then txtAnterior = "s"
                            If ImporteFormateado(Text1(15).Text) <> Data1.Recordset!DtoPPago Then txtAnterior = "s"
                            If txtAnterior <> "" Then
                                Screen.MousePointer = vbHourglass
                                'UPDATE
                                txtAnterior = "UPDATE scaalb SET DtoPPago = " & DBSet(Text1(15).Text, "N")
                                txtAnterior = txtAnterior & " , DtoGnral = " & DBSet(Text1(16).Text, "N")
                                txtAnterior = txtAnterior & ObtenerWhereCP(True)
                                If ejecutar(txtAnterior, False) Then
                                    Espera 0.2
                                    PosicionarData False
                                End If
                                Screen.MousePointer = vbDefault
                            End If
                        End If
                        
                    End If
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
                    numlinea = Data2.Recordset!numlinea
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    
                    'If vParamAplic.NumeroInstalacion = vbHerbelca Then InsertarCosteHerbelca numlinea
                    
                    TerminaBloquear
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(9), True
                    lblF.Caption = ""
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



Private Function ModificarCabAlbaran() As Boolean
Dim B As Boolean
Dim Sql As String

    On Error GoTo EModificaAlb
    conn.BeginTrans
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
    B = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    
    If B Then
        B = ModificaDesdeFormulario(Me, 1)
        If B Then
            Sql = "UPDATE scaalb SET nomdirec=" & DBSet(Text2(12).Text, "T") & " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!Numalbar
            conn.Execute Sql
        End If

        If B Then
            'comprobar si se ha cambiado el cliente
            'o si se ha cambiado la fecha del albaran
            'If (CInt(Me.Data1.Recordset!CodClien) <> CInt(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
            'DAVID.   No es un CINT. Tiene que ser un clng o val
            If (Val(Me.Data1.Recordset!codClien) <> Val(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
                'si hay numeros de serie en ese albaran, actualizamos el cliente
                'al nuevo cliente
                Sql = "UPDATE sserie SET codclien=" & DBSet(Text1(4).Text, "N") & ","
                Sql = Sql & " fechavta=" & DBSet(Text1(1).Text, "F")
                Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
                Sql = Sql & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!Numalbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute Sql
                
                'Modificar el cliente en la smoval
                Sql = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N") & ","
                Sql = Sql & " fechamov=" & DBSet(Text1(1).Text, "F")
                Sql = Sql & ", horamovi= concat(" & DBSet(Text1(1).Text, "F") & ",' ',hour(horamovi),':',minute(horamovi),':',second(horamovi))"
                Sql = Sql & " WHERE detamovi='" & CodTipoMov & "'"
                Sql = Sql & " AND document='" & Text1(0).Text & "'"
                Sql = Sql & " and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute Sql
            End If
            
            
            
            'LOG GrabaCambioTrabajador
            If GrabaCambioTrabajador >= 0 Then
                'Ha cambiado el trabajador
                '------------------------------------------------------------------------------
                '  LOG de acciones.
                Set LOG = New cLOG
                Sql = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", CStr(GrabaCambioTrabajador))
                Sql = "Antes: " & Format(GrabaCambioTrabajador, "0000") & " " & Sql
                Sql = "Ahora: " & Text1(3).Text & " " & Text2(3).Text & vbCrLf & Sql
                Sql = "Albaran: " & Text1(0).Text & " Fecha " & Text1(1).Text & vbCrLf & Sql
                LOG.Insertar 20, vUsu, Sql
                Set LOG = Nothing
            End If
        End If
    
    
    
        If vParamAplic.PtosAsignar > 0 Then
            'Sistema de puntos
   
  
                
                If Val(Data1.Recordset!codClien) <> Val(Text1(4).Text) Then
                    'Si cambia el cliente, hay que ver
                    Sql = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text)
                    If Val(Sql) = "1" Then
                        'El nuevo cliente tiene puntos
                        Sql = "+"
                        If Data1.Recordset!Puntos < 0 Then Sql = "-"
                        Sql = "UPDATE sclien set puntos=coalesce(puntos,0) " & Sql & DBSet(Abs(Data1.Recordset!Puntos), "N")
                        Sql = Sql & " WHERE codclien =" & Text1(4).Text
                        conn.Execute Sql
                    
                        BuscaChekc = "U"
                       
                    Else
                        Sql = "UPDATE scaalb SET puntos=0 WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!Numalbar
                        conn.Execute Sql
                        BuscaChekc = "D"
                        Text2(1).Text = ""
                        
                    End If
                    
                    
                    'Le quito los puntos al cliente origen. Osea al reves de arriba
                    Sql = "-"
                    If Data1.Recordset!Puntos < 0 Then Sql = "+"
                    Sql = "UPDATE sclien set puntos=puntos " & Sql & DBSet(Abs(Data1.Recordset!Puntos), "N")
                    Sql = Sql & " WHERE codclien =" & Data1.Recordset!codClien
                    conn.Execute Sql
                    
                    
                    
                    'O borro o updateo movimientos puntos
                    Sql = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                    Sql = Sql & " AND codclien = " & Data1.Recordset!codClien
                    If BuscaChekc = "U" Then
                        BuscaChekc = "UPDATE smovalpuntos SET codclien = " & Text1(4).Text
                    Else
                        BuscaChekc = "DELETE FROM smovalpuntos "
                    End If
                    Sql = BuscaChekc & Sql
                    
                    ejecutar Sql, False
                    
                    BuscaChekc = ""
                    
                    
                End If
                If CDate(Me.Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                    Sql = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                    Sql = Sql & " AND codclien = " & Data1.Recordset!codClien
                    Sql = "UPDATE smovalpuntos set fechaalb= " & DBSet(Text1(1).Text, "F") & Sql
                    ejecutar Sql, False
                    
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
                        CargaGrid Me.DataGrid1, Me.Data2, True
                        cmdCancelar_Click
                    End If
                End If
            End If
            
            
            If B Then
                Set FrmArt = New frmBasico2
                'FrmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo busqueda
'                FrmArt.DesdeTPV = False
'                FrmArt.Show vbModal
                AyudaArticulos FrmArt, txtAux(1).Text
                Set FrmArt = Nothing
'                txtAux_LostFocus (1)
            End If
            
    Case 9 'CENTRO COSTE/ PROVEEDOR
        If vEmpresa.TieneAnalitica Then
            'centro de coste
            EsCabecera = 3
            AbrirForm_CentroCoste
        Else
'            Set frmProv = New frmComProveedores
'            frmProv.DatosADevolverBusqueda = "1"
'            frmProv.Show vbModal
            Set frmProv = New frmBasico2
            AyudaProveedores frmProv, txtAux(Index)
            Set frmProv = Nothing
        End If
    End Select
    PonerFoco txtAux(Index)
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
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LineaIntercalar = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            
            
            If hcoCodTipoM = "DEV" Then
                'Por si acaso ha tocado dtos
                Text1(15).Text = DBLet(Data1.Recordset!DtoPPago, "N")
                FormateaCampo Text1(15)
                Text1(16).Text = DBLet(Data1.Recordset!DtoGnral, "N")
                FormateaCampo Text1(16)
            End If
            
            
            If Me.hcoCodTipoM = "ALM" And vParamAplic.EntradaRapidaFacturasMostrador Then
                'Estamos en facturas mostrador, cont entrada rapida. Simularemos el boton
                'de facturar. cuando pulse cancelar
                If EraNuevaLinea Then HacerToolbar 12
            End If
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim Cad As String
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
        
            Text1(1).Text = Format(Now, "dd/mm/yyyy")  'Julio 19
        
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
            Cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            If Cad = "0" Then BloquearDatosCliente (False)
            
            
            
            
            
            
            
            'recuperamos el tipo de facturacion del cliente
            Me.cboFacturacion.ListIndex = CInt(NomTraba)
            
            'Si es factura de AGUA, debe traer la referencia, que sera el contador
            If Rs!codtipmf = "FAG" Then
                'cadlist=>>  codtipom='FAG' and numfactu=0000001 and fecfactu='2014-05-16'
                Cad = cadList & " AND 1"
                Cad = DevuelveDesdeBD(conAri, "referenc", "scafac1", Cad, "1 ORDER BY 1 DESC")
                Text1(13).Text = Cad
            End If
            
            
            
            'ULTIMO
            'Memorizo cad con codtipom
            Cad = Rs!codtipmf
            
            
            
            
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
            Cad = "N" 'para que la busqueda de despues no de error
            Text1(18).Text = ""
            Rs.Close
        End If
        
        
        'Observacion 2
        Text1(19).Text = motivo
        
        'DAVID
        'Para que meta la letra de serie, NO el tipo moviemiento
        Rs.Open "SELECT * FROM stipom WHERE codtipom='" & Cad & "'"
        If Not Rs.EOF Then Cad = DBLet(Rs!LetraSer, "T")
        Rs.Close
        If Cad = "" Then Cad = CodTipoMov
        If Text1(18).Text <> "" Then
            TxtMotivoFra = DevuelveDesdeBD(conAri, "texto", "sparaidioma", "codigo", 1)  '1.- Rectifica a
            If TxtMotivoFra = "" Then TxtMotivoFra = "RECTIFICA A FACTURA"
            Text1(18).Text = TxtMotivoFra & ": " & Cad & ", " & Text1(18).Text
        End If
        
            
        'DAVID
        'JUNIO 2010
        'Envio por defecto del cliente
        If Text1(4).Text <> "" Then
            Cad = "select sclien.codenvio,nomenvio from  sclien,senvio where sclien.codenvio=senvio.codenvio AND sclien.codclien= " & Text1(4).Text
            Rs.Open Cad, conn, adOpenForwardOnly, adCmdText
            If Not Rs.EOF Then
                Text1(29).Text = Rs!CodEnvio
                Text2(29).Text = Rs!nomEnvio
            Else
                Text1(29).Text = ""
                Text2(29).Text = ""
            End If
            Rs.Close
        End If
            
        
        
        
        
        chkFacturar.Value = IIf(vParamAplic.MarcarAlbaranFacturar, 1, 0)
        
        
        Set Rs = Nothing
    Else
        'Añadiremos el boton de aceptar y demas objetos para insertar
        PonerModo 3
        
        
        If vParamAplic.NumeroInstalacion = 2 Then
            Me.chkFacturar.Enabled = False
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
    
    
    
        
    
    'Mayo2014
    Cad = "1"
    If CodTipoMov = "ALM" Then
        If vParamAplic.EntradaRapidaFacturasMostrador Then Cad = "4"
    End If
    PonerFoco Text1(Val(Cad))
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
        If Not Data2.Recordset.EOF Then
            LineaIntercalar = Data2.Recordset!numlinea
        End If
    Else
        LineaIntercalar = 0
        lblIndicador.Caption = "INSERTAR"
    End If
    
    
    
    
    AnyadirLinea DataGrid1, Data2
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
Dim cadB As String
    
    cadB = " 1 = 1"
    If Not EsHistorico Then cadB = " codtipom='" & CodTipoMov & "'"

    If vParamAplic.NumeroInstalacion = vbHerbelca And vUsu.CodigoAgente > 0 Then
        
        cadB = cadB & " AND (codagent = " & vUsu.CodigoAgente
        If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then cadB = cadB & " OR codclien in (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
        cadB = cadB & ")"
        
    End If

'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = 0
        
        MandaBusquedaPrevia cadB
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB
        
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
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
    
    If Data2.Recordset.EOF Then Exit Sub
    
        
    'Herbelca. Tania 21/07/2016
    '--------------------------
    ' De varios no dejo modificar la linea. Segun ella esto ya lo hacia.
    'Version: 4_6_51 de Feb16 No lo hace    Solo era para eliminar linea
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.Nivel > 0 Then

            vWhere = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", CStr(Data2.Recordset!codArtic), "T")
            If Val(vWhere) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
    End If
        
    
    
    
    
    'bloqueamos el registro a modificar
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
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
Dim Cad As String
Dim NumAlbElim As Long

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

        'Nov 2014
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        If vUsu.Nivel > 0 Then
            
            Cad = "slialb.codartic=sartic.codartic and artvario=1 AND codtipom = '" & Data1.Recordset!codtipom & "' AND numalbar "
            Cad = DevuelveDesdeBD(conAri, "count(*)", "slialb,sartic", Cad, CStr(Data1.Recordset!Numalbar))
            If Val(Cad) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
    End If


    If vParamAplic.PtosAsignar > 0 Then
        Cad = Replace(ObtenerWhereCP(False), "scaalb", "slialb") & " AND codartic "
        Cad = DevuelveDesdeBD(conAri, "codartic", "slialb", Cad, vParamAplic.PtosArticuloCanje, "T")
        If Cad <> "" Then
            MsgBox "Tiene  articulo canje. ", vbExclamation
            Exit Sub
        End If

    End If
    If hcoCodTipoM = "DEV" Then
        cadList = "devolucion"
    Else
        cadList = "albarán"
    End If
    Cad = "Cabecera de " & cadList & "." & vbCrLf
    Cad = Cad & "------------------------------------       " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar " & cadList & ":            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(30).Text
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Text1(1).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
          
    If hcoCodTipoM <> "DEV" Then
        If Not ComprobarInventario Then Exit Sub
    End If
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
    
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
Dim Sql As String
Dim CodproveHerbelca As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        CodproveHerbelca = "codprove"
        
        Sql = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", CStr(Data2.Recordset!codArtic), "T", CodproveHerbelca)
        If vUsu.Nivel > 0 Then
            
            If Val(Sql) > 0 Then
                MsgBox MensajeHerbelcaEliminarVarios, vbExclamation
                Exit Sub
            End If
        End If
        
        
        If CodproveHerbelca = 5000 Then
            'Proveedor de varios
             If vUsu.AlmacenPorDefecto2 > 1 And Data2.Recordset!codArtic <> vParamAplic.PtosArticuloCanje Then
                MsgBox "No puede eliminar linea", vbExclamation
                Exit Sub
            End If
        End If
        
        
        'SI es de portes tampoco dejo
        If vParamAplic.ArtPortesN = CStr(Data2.Recordset!codArtic) Then
            If vUsu.AlmacenPorDefecto2 > 1 Then
                MsgBox "No puede eliminar linea", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    ModificaLineas = 3 'Eliminar
    Sql = "¿Seguro que desea eliminar la línea de Albaran?     "
    Sql = Sql & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    Sql = Sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    Sql = Sql & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
            SituarDataTrasEliminar Data2, NumRegElim
            CalcularDatosFactura
        End If
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub




Private Sub cmdMtoCampos_Click(Index As Integer)
    If Index = 0 Then
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
                    cadList = cadList & " AND numalbar = " & Data1.Recordset!Numalbar
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
Dim Cad As String
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
            CargaGrid DataGrid1, Data2, True
            Set miRsAux = Nothing
        End If
        
        Cad = ""
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Cad = "SI"
        Else
            If vParamAplic.NumeroInstalacion = vbHerbelca Then Cad = "S"
        End If
        If Cad <> "" Then
                Cad = ObtenerWhereCP(False)
                'select count(*) from slialb where numalbar =albaran and codartic in (select codartic from sarticexplosivos) and 1=1
                Cad = Replace(Cad, "scaalb", "slialb")
                Cad = Cad & " AND codartic IN (select codartic from sarticexplosivos) and 1"
                Cad = DevuelveDesdeBD(conAri, "count(*)", "slialb", Cad, "1")
                If Val(Cad) > 0 Then
                    Cad = "1"
                Else
                    Cad = "0"
                End If
                
                If DBLet(Data1.Recordset!codnatura, "N") <> Val(Cad) Then
                    'Hay que actualizar
                
                    Cad = "UPDATE scaalb set codnatura = " & Cad
                    Cad = Cad & ObtenerWhereCP(True)
                    If Not ejecutar(Cad, False) Then MsgBox "Valor del precursor explosivos del albaran es incorrecto", vbExclamation
                        
                 End If
                        
            
        End If
    
        If vParamAplic.NumeroInstalacion = 2 Then ComprobarComisionesAlbaranes
        If vParamAplic.PtosAsignar > 0 Then
            Cad = "puntos"
            If DevuelveDesdeBD(conAri, "tienepuntos", "sclien", "codclien", CStr(Data1.Recordset!codClien), "N", Cad) = "1" Then
                
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
                If Cad = "" Then
                    PtosCliente = 0
                Else
                    PtosCliente = CCur(Cad)
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
                    Cad = "UPDATE scaalb set puntos = " & DBSet(Puntos, "N")
                    Cad = Cad & ObtenerWhereCP(True)
                    conn.Execute Cad
                    
                    'los movimimeintos PUNTOS
                    Cad = Replace(ObtenerWhereCP(False), "scaalb", "smovalpuntos") & " and concepto=0 AND codclien"
                    
                    Cad = DevuelveDesdeBD(conAri, "numero", "smovalpuntos", Cad, CStr(Data1.Recordset!codClien))
                    If Cad = "" Then
                        
                        Cad = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", CStr(Data1.Recordset!codClien))
                        'NUEVA LINEA
                        
                        Cad = " VALUES (" & Data1.Recordset!codClien & "," & Val(Cad) + 1 & "," & DBSet(Data1.Recordset!codtipom, "T") & "," & Data1.Recordset!Numalbar
                        
                        Cad = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov)" & Cad
                        Cad = Cad & " ," & DBSet(Data1.Recordset!FechaAlb, "F") & ",0," & DBSet(Puntos, "N") & ",now())"
                    Else
                        'UPDATE
                        Cad = "UPDATE smovalpuntos set puntos=" & DBSet(Puntos, "N") & " WHERE codclien=" & Data1.Recordset!codClien & " AND numero=" & Cad
                    End If
                    
                    ejecutar Cad, False
                    
                    
                    Puntos = Puntos - PtosAnt
    
                    PtosCliente = PtosCliente + Puntos
                    Cad = "UPDATE sclien set puntos = " & DBSet(PtosCliente, "N")
                    Cad = Cad & " WHERE codclien = " & Data1.Recordset!codClien
                    conn.Execute Cad
                    
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
                Data2.Recordset.MoveLast  'El ultimo es el porte
            End If
        End If
        Me.cmdCancelar.Cancel = True
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


Private Sub DataGrid1_DblClick()
    If Modo = 2 Then
        If Not Data2.Recordset.EOF Then AbrirForm_Articulos DBLet(Data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Modo = 2 And KeyCode = 113 Then
        If Not Data2.Recordset.EOF Then AbrirForm_Articulos DBLet(Data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 1660 And X < 7950 Then
            If IsNull(Me.Data2.Recordset!origpre) Then Exit Sub
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
Dim Sql As String

    On Error GoTo Error1

    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        '- ampliacion lineas
        Sql = "select ampliaci from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text2(16).Text = DBLet(Rs.Fields(0).Value, "T")
        End If
        Rs.Close
        Set Rs = Nothing
        
        If vEmpresa.TieneAnalitica Then
            '- centro de coste
            ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            '- nombre proveedor
            Text2(9).Text = DBLet(Me.Data2.Recordset!nomprove, "T")
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

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 33 'Nº Serie si lineas con articulos de control Nº serie
        .Buttons(12).Image = 26 'GEnerar factura
        .Buttons(13).Image = 30 'Marcar a facturar
        
        .Buttons(14).Image = 27 'Imprimir portes
        .Buttons(15).Image = 16 'Imprimir Pedido
        
        'MAYO 2015  Herbelca. ALbran ruta Castellon
        .Buttons(16).Image = 40 'Imprimir Pedido
        
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
        
        If vParamAplic.TipoPortes <> 1 Then
            If vParamAplic.PathFirmasAlbaran <> "" Then
                .Buttons(14).ToolTipText = "Imprimir albaran firmado"
                .Buttons(14).Style = tbrDefault
                .Buttons(14).Image = 54
            Else
                .Buttons(14).Style = tbrSeparator
                .Buttons(14).ToolTipText = ""
            End If
        Else
            .Buttons(14).Style = tbrDefault
            .Buttons(14).ToolTipText = "Imprimir portes"
        End If
    End With
    Me.SSTab1.Tab = 0
    SSTab1.TabVisible(3) = vParamAplic.CartaPortes
      
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
        
        
    Else
        Label1(11).Caption = "Nº Pedido"
        Label1(10).Caption = "Fecha pedido"
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
    Me.FrameHco.visible = EsHistorico
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
    If esDEV Then
        Color = Me.BackColor 'El que tengo en desarrollo
    Else
        Color = &H8000000F
    End If
    
    Me.chkFacturar.FontBold = esDEV
    Me.chkFacturar.ForeColor = IIf(esDEV, vbWhite, vbBlack)
    
    'Botones que SEGURO qee no estaran visibles
    Toolbar1.Buttons(11).visible = Not esDEV 'Nº Serie si lineas con articulos de control Nº serie
    Toolbar1.Buttons(13).visible = Not esDEV 'Marcar a facturar
    Toolbar1.Buttons(14).visible = Not esDEV 'Imprimir portes
    Toolbar1.Buttons(15).visible = Not esDEV 'Imprimir Pedido
    
    'MAYO 2015  Herbelca. ALbran ruta Castellon
    Toolbar1.Buttons(16).visible = Not esDEV 'Imprimir Pedido
    
    
    
    
    
    
    
    
    'Cambios de lbl
    Toolbar1.Buttons(12).ToolTipText = IIf(esDEV, "Pasar a albaran", "Generar Factura")
    
    
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
            LblMostr.Caption = "DEVOL."
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
    Text3(0).Text = "BASE IMP."
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
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If EsCabecera = 0 Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(30), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
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
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing

            'JUNIO22
            'Indice = 6
            'Set frmCV = New frmBasico2
            'AyudaClientesV frmCV, Text1(Indice)
            'Set frmCV = Nothing
            
            
                        
            
            
            
            
            
            
            
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
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
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
                If CadenaDesdeOtroForm = vbCrLf Then CadenaDesdeOtroForm = ""
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
            If Text1(Index).Locked Then Exit Sub
            Text1(Index).Text = UCase(Trim(Text1(Index).Text))
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
            If (Modo = 4) Then
                If (Text1(Index).Text = Data1.Recordset!nifClien) Then Exit Sub
            End If
            
            'Diciembre 2021-
            'Hay que validar el NIF
            If Text1(Index).Text <> "" Then
                If Not ValidarNIF_(Text1(Index).Text, True) Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
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
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If vParamAplic.NumeroInstalacion = vbHerbelca And vUsu.CodigoAgente > 0 Then
        If cadB = "" Then cadB = " TRUE "
        cadB = cadB & " AND (codagent = " & vUsu.CodigoAgente
        If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then cadB = cadB & " OR codclien in (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
        cadB = cadB & ")"
    End If
    
    
    
    
    
    If chkVistaPrevia = 1 Then
        EsCabecera = 0
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        If Me.EsHistorico = False Then
            cadB = cadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
        End If
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
    If EsCabecera = 0 Then
        Cad = Cad & ParaGrid(Text1(30), 10, "Tipo Alb.")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Albaran")
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha Alb.")
        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        tabla = NombreTabla
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
                Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|18·"
                Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
                tabla = "sdirec"
                devuelve = "0|1|"
                
        ElseIf EsCabecera = 2 Then
            'DIRENVIO
            '--------------------
            Titulo = "Dirección de envio cliente: "
            Desc = " envio"
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Codigo" & Desc & "|sdirenvio|coddiren|N|000|18·"
            Cad = Cad & "Descripción" & Desc & "|sdirenvio|nomdiren|T||65·"
            tabla = "sdirenvio"
            devuelve = "0|1|"
        
        Else
            MsgBox "ERROR manda busqueda previa. En frmalbaranes.   Avise soporte técnico", vbExclamation
            Cad = ""
        End If
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
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
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & ".     " & Me.Caption, vbInformation
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
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar

    Screen.MousePointer = vbHourglass
    On Error GoTo EPonerLineas

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
    
    
    If vParamAplic.PtosAsignar > 0 And Not EsHistorico Then ObtenerPuntos
    
   
    If Me.FrameCampos.visible Then CargaDatosCampos
    If vParamAplic.ManipuladorFitosanitarios2 Then
        If Val(DBLet(Data1.Recordset!TipoCarnet, "N")) > 0 Then
            Text2(0).Text = IIf(Val(Data1.Recordset!TipoCarnet) = "2", "Cualificado", "Básico")
        Else
            Text2(0).Text = ""
        End If
    End If
    CalcularDatosFactura
    
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
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
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
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
        
        
        
    
    If B Then 'modo=2
        If Me.FrameCampos.visible Then
            'Tiene campos visibles
            If Not Data1.Recordset.EOF Then B = True
        Else
            B = False
        End If
    End If
    Me.cmdMtoCampos(0).Enabled = B
    Me.cmdMtoCampos(1).Enabled = B
    
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
            If Not Data2.Recordset.EOF Then
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
            If Not Data2.Recordset.EOF Then
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
Dim SoloEnEfec As Boolean

    SoloEnEfec = False
    ClienteBloqueadoYFormaPagoCorrecta = True
    If Me.hcoCodTipoM = "ALM" Then
        'Para facturas mostrador siempre comprueban
        If EsClienteBloqueado2(Text1(4).Text, 0, True, False) Then SoloEnEfec = True
    Else
        If SoloEnEfectivoAlbaranes Then SoloEnEfec = True
    End If
    
    If SoloEnEfec Then
            'LA forma de pago solo pude ser efectivo o tarjeta   (0 o 6)
            BuscaChekc = DevuelveDesdeBDNew(1, " sforpa", "tipforpa", "codforpa", Text1(14).Text)
            If BuscaChekc <> "0" And BuscaChekc <> "6" Then
                MsgBox "Cliente bloqueado.  Forma pago INVALIDA(solo efectivo o tarjeta) ", vbExclamation
                PonerFoco Text1(14)
                ClienteBloqueadoYFormaPagoCorrecta = False
            End If
            BuscaChekc = "   "

    End If
    
End Function

Private Function DatosOkLinea(ByRef vCStock As CStock, ByRef ARticuloFitosantiario As Boolean) As Boolean
Dim B As Boolean
Dim i As Byte
Dim Aux As String
Dim Aux3 As String
Dim vArtic As CArticulo
Dim PuntosCliente As Currency
Dim C2 As Currency
Dim Comision As String
Dim CanDispo As Currency
Dim vPrecioFact As CPreciosFact
Dim PrMinimo As Currency
Dim ComprobarPrMinimo2 As Boolean
Dim HacerComprobarPrecio As Boolean

Dim B_Aux As Boolean

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

        If hcoCodTipoM <> "ART" Then
            'Octubre 2011
            ' Comprobar que este articulo, para este cliente, no esta en otro pedido
            Set vArtic = New CArticulo
            vArtic.LeerDatos txtAux(1).Text
            If vArtic.EsDeVarios = 0 Then
                Aux = "scaped.numpedcl=sliped.numpedcl  AND codclien = " & Text1(4).Text & "  AND sliped.codartic"
                Aux = DevuelveDesdeBD(conAri, "concat(sliped.numpedcl,"" de fecha "",DATE_FORMAT(fecpedcl,'%d %M %Y'))", "scaped,sliped", Aux, txtAux(1).Text, "T")
                If Aux <> "" Then
                    
                    Aux = "Artículo: " & vArtic.Codigo & "   " & vArtic.Nombre & vbCrLf & vbCrLf & "Esta en el pedido: " & Aux
                    Aux = "Cliente: " & Text1(4).Text & "   " & Text1(5).Text & vbCrLf & vbCrLf & Aux
                    Aux = Aux & vbCrLf & vbCrLf & "¿Continuar?"
                    If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then B = False
                End If
            End If
            'Set vArtic = Nothing  lo pongo a nothing bajo
        End If
    End If
    
    
    
    
    
    
    'Articulo canje.
    If vParamAplic.PtosAsignar > 0 Then
        If ModificaLineas = 2 Then
            'Esta modificando el articulo de canje
            If Data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
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
                Aux3 = "puntos"
                Aux = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text, "N", Aux3)
                If Val(Aux) = 0 Then
                    MsgBox "El cliente no tiene marca de Puntos", vbExclamation
                    B = False
                Else
                    If Aux3 = "" Then Aux3 = "0"
                    PuntosCliente = CCur(Aux3)
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
                         Aux = CalcularPuntosAlbaranCABEL(Replace(ObtenerWhereCP(False), "scaalb", "slialb"), Data1.Recordset!FechaAlb, Aux3, Comision)
                         
                         If Aux = "" Then Aux = "0"
                         If CCur(Aux) = 0 Then
                            'No hacemos nada
                            MsgBox "No tiene articulos de las familias de canje", vbInformation
                            
                            If Not Data2.Recordset.EOF Then B = False
                         Else
                           
                            C2 = Round2(CCur(Aux3) / vParamAplic.PtosEquivalencia, 2) '-> necesito como mucho estos puntos
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
        
        '#####
     
        If B Then
            GrabaLogCambioPrecioDto = False
            
            'En el procedimiento vemos prminimo, cambio en descuentos, comision agente
            If vParamAplic.LogCambioPrecDto Then
                HacerComprobarPrecio = True
                If hcoCodTipoM = "DEV" Then
                    'En pedido de devolucion, no tengo que comprabr nada
                    If Me.txtAux(6).Locked Then HacerComprobarPrecio = False
                End If
                
                If HacerComprobarPrecio Then
                    ComprobarCambioPrecioDtoyVtaBajoPrecio
                End If
            End If
        End If
    
        If B And vParamAplic.NumeroInstalacion = vbHerbelca Then
            ComprobarPrMinimo2 = False
            
            'Añadido el 5 septembre
            If vArtic Is Nothing Then
                Set vArtic = New CArticulo
                vArtic.LeerDatos txtAux(1).Text
            
            End If
            
            If B And vArtic.EsDeVarios = 0 And txtAux(5).Text <> "P" Then     'en herbelca. Precio minimo
                'Si es Precio especial
                If txtAux(5).Text <> "E" Then
                    ComprobarPrMinimo2 = True
                    
                Else
                    'Es precio especial.
                    'Comporbar que no ha tocado el descuento
                    If GrabaLogCambioPrecioDto Then ComprobarPrMinimo2 = True
                
                End If
            End If
            
            
            If Not HacerComprobarPrecio Then ComprobarPrMinimo2 = False  'en albara devolucion NO lo hace
            
            
            If ComprobarPrMinimo2 Then
                'If Not vArtic.EstablecidoPrecioMinimo Then vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
                vArtic.FijarprecioMinimo_ CDate(Text1(1).Text), Val(Text1(4).Text)
                
                Set LOG = New cLOG
                If vArtic.EstablecidoPrecioMinimo Then
                    C2 = ImporteFormateado(txtAux(3).Text)
                    If C2 <> 0 Then
                        C2 = Round2(ImporteFormateado(txtAux(8).Text) / C2, 4)
                        If C2 < vArtic.PrecioMinimo Then
                            Aux = "Albaran: " & Text1(30).Text & Text1(0).Text & " " & Text1(1).Text & vbCrLf
                            Aux = Aux & "Cliente: " & Text1(4).Text & " " & Text1(5).Text & vbCrLf
                            Aux = Aux & "Articulo: " & vArtic.Codigo & "  - " & vArtic.Nombre & vbCrLf
                            Aux = Aux & "PVM: " & vArtic.PrecioMinimo & vbCrLf
                            Aux = Aux & "Venta: " & C2
                            LOG.VariableTemporal = Aux
                            
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
                                If B Then
                                    Aux = LOG.VariableTemporal
                                    LOG.Insertar 43, vUsu, Aux
                                End If
                                Aux = ""
                            End If
                        End If
                    End If
                End If
                Set LOG = Nothing
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
        Aux3 = ""
        If Val(txtAux(0).Text) = 1 Then
            Aux3 = DevuelveDesdeBD(conAri, "ctrstock", "sartic", "codartic", txtAux(1).Text, "T")
        End If
        If Aux3 = "1" Then
            Aux3 = "sum(cantidad)"
            '                                           cualqueir almacen
            Aux = "scaped.numpedcl=sliped.numpedcl  AND codalmac>=" & txtAux(0).Text & " AND codartic "
            Aux = DevuelveDesdeBD(conAri, "max(fecpedcl)", "scaped,sliped", Aux, txtAux(1).Text, "T", Aux3)
            If Aux <> "" Then
                C2 = CCur(Aux3)
                
               
                Aux3 = "codalmac>=" & txtAux(0).Text & " AND codartic "
                Aux3 = DevuelveDesdeBD(conAri, "sum(canstock)", "salmac", Aux3, txtAux(1).Text, "T")
                CanDispo = CCur(Aux3) 'en stock para el almacen hay candispo unidades
                
                If CanDispo < vCStock.cantidad + C2 Then
                  'Las cantiodades que ha pedido mas las hay en pedidos superan
                  Aux3 = "Hay pedidos con este articulo pendientes de servir" & vbCrLf & vbCrLf
                  Aux3 = Aux3 & "Stock articulo:     " & CanDispo & vbCrLf
                  Aux3 = Aux3 & "Unidades en pedidos: " & C2
                  
                  Aux3 = Aux3 & vbCrLf & vbCrLf & "¿Continuar?"
                  If MsgBox(Aux3, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
                End If

            End If
         End If
    End If
    
    
    
    If B Then
        If hcoCodTipoM <> "FMO" Then
            'En facturas de mostrador NO lo compurbeo para acelerar el prloceso
            vCStock.ComprobarFechaInventario True, ""
              
        End If
        
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
                    If Data2.Recordset!codArtic <> txtAux(1).Text Then Aux = Aux & "-Codigo de articulo" & vbCrLf
                    If Data2.Recordset!cantidad <> ImporteFormateado(txtAux(3).Text) Then Aux = Aux & "-Cantidad" & vbCrLf
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
Dim Aux As String
Dim EsEfectivo As Boolean
    Select Case Indice
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos
            
        Case 5: mnNuevo_Click 'Nuevo
        
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 10: mnLineas_Click  'Lineas
        Case 11:
        
                'De momento NINGUna de estas dos acciones se puede hacer para fitosanitarios
                If Not vParamAplic.ManipuladorFitosanitarios2 Then
                    If Modo = 5 Then
                        'Insertar intercalando
                        BotonAnyadirLinea True
                    Else
                        BotonNSeries 'Control Nº Series
                    End If
                End If
        Case 12 'Generar Factura Mostrador
                'o Factura Rectificativa (FRT)
           
            
                'Septiebmre2009
                If Data2.Recordset Is Nothing Then Exit Sub
                If Data2.Recordset.RecordCount = 0 Then
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
                Else
                    
                    If Not CuantosTiposDeIva Then Exit Sub
                
                    'Factura NORMAl
                    EsEfectivo = False
                    If vParamAplic.MaximoImporteEfectivo > 0 Then
                    
                        Aux = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(14).Text)
                        If Aux <> "" Then
                            If Val(Aux) = 0 Then
                                EsEfectivo = True
                                Aux = Text3(55).Text
                                If Aux = "" Then Aux = "0"
                                If ImporteFormateado(Aux) > vParamAplic.MaximoImporteEfectivo Then
                                    MsgBox "Excede importe máximo de pago en efectivo", vbExclamation
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                
                If Me.chkFacturar.Value = 1 Then
                
                    If Not ClienteBloqueadoYFormaPagoCorrecta Then Exit Sub
                    
                    If vParamAplic.ManipuladorFitosanitarios2 Then
                        If Not VerCarnetManipulador Then Exit Sub
                        
                    End If
                    
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    
                    
                    If vParamAplic.NumeroInstalacion = vbHerbelca Then
                        If Not PrecioMinimoAlbaran Then Exit Sub
                    End If
                    
                    
                    
                    'Febrero 2022
                    If hcoCodTipoM = "ALM" And Not EsHistorico And vParamAplic.NumeroInstalacion = vbHerbelca Then
                    
                        BuscaChekc = ObtenerWhereCP(False) & " AND 1"
                        BuscaChekc = DevuelveDesdeBD(conAri, "codnatura", "scaalb", BuscaChekc, "1")
                
                        If Val(BuscaChekc) > 0 Then
                            BuscaChekc = "{" & NombreTabla & ".codtipom}='" & hcoCodTipoM & "'"  'lo que habia
                            BuscaChekc = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
                            With frmImprimir
                                .outTipoDocumento = 0
                                .NombreRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", 101)  ' "HerAlbaranPrecursor.rpt""HerAlbaranPrecursor.rpt"
                                .FormulaSeleccion = BuscaChekc
                                .OtrosParametros = ""
                                .NumeroParametros = 0
                                .SoloImprimir = False
                                .EnvioEMail = False
                                .Opcion = 250
                                .Titulo = "Precursor explosivos"
                                .ConSubInforme = True
                                .Show vbModal
                        End With
                    End If
                 End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    'Facturacion de Albaran de Mostrador
                    '-----------------------------------
                    frmListadoPed.FacturaSuperaImporteTicket = False
                    
                    BuscaChekc = "NO"
                    If CodTipoMov = "ALM" Then
                        BuscaChekc = ""
                    Else
                        'Cuando herbelca factura un unico albaran de venta, tambien lo pasa a ALM (factura simplicada)
                        'Por lo tanto
                        If vParamAplic.NumeroInstalacion = vbHerbelca And CodTipoMov = "ALV" Then BuscaChekc = ""
                    End If
                    If BuscaChekc = "" Then
                        If vParamAplic.SerieFacturaDirecta <> "" Then
                            'No puede hacer facturas simplicadas a partuculares / empresas que superen cierta cantidad(maximporticketparticular / maximporticketempresa)
                            txtAnterior = "clivario"
                            BuscaChekc = DevuelveDesdeBD(conAri, "particular", "sclien", "codclien", Text1(4).Text, "N", txtAnterior)
                            
                            
                            '17 ENERO   ARIGES-502
                            If txtAnterior = "0" Then
                                'Cualquier cliente IDENDTIICADO hace factura FAD
                                frmListadoPed.FacturaSuperaImporteTicket = True
                            Else
                                If Comprobar_NIF(Text1(6).Text) Then
                                    'De varios per NIF correcto
                                    frmListadoPed.FacturaSuperaImporteTicket = True  'hace FAD
                                Else
                                    Aux = Text3(55).Text
                                    If Aux = "" Then Aux = "0"
                                    If BuscaChekc = "1" Then
                                        'particular
                                        If ImporteFormateado(Aux) > vParamAplic.MaxImporteTicketParticular Then Aux = "S" 'Supera
                                    Else
                                        'Empresa
                                        If ImporteFormateado(Aux) > vParamAplic.MaxImporteTicketEmpresa Then Aux = "S"  'Supera
                                    End If
                                    If Aux = "S" Then
                                        'Supera el importe maximo para NO hacer factura simplificada
                                        MsgBox "Supera el maximo para realizar factura simplificada y no esta identificado el cliente", vbExclamation
                                        Exit Sub
                                    End If
                                End If
                                    
                                
                            End If ' txtAnterior = "0"
                        End If 'factura direta
                        
                        txtAnterior = ""
                    End If  'buscachek
                    BuscaChekc = ""
                    
                    frmListadoPed.TotalFacDesdeAlbaran = 0
                    If Text3(55).Text <> "" Then
                        If ImporteFormateado(Text3(55).Text) > 0 Then
                            'Si la forma de pago es Efectivo
                           If EsEfectivo Then frmListadoPed.TotalFacDesdeAlbaran = ImporteFormateado(Text3(55).Text)
                        End If
                    End If
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
    
            
        Case 13
            'DAVID
            'Marca los albaranes que esten como NO facturar a facturar
            If Modo = 5 Then
                If ModificaLineas = 0 Then
                    If vParamAplic.ManipuladorFitosanitarios2 Then ModificaLote
                End If
            Else
                MarcarAlbaranes
            End If
            
            
        Case 14
            If vParamAplic.TipoPortes <> 1 Then
                If vParamAplic.PathFirmasAlbaran <> "" Then
                    Screen.MousePointer = vbHourglass
                    ImprimirAlbaranFirmado
                    Screen.MousePointer = vbDefault
                End If
            Else
                BotonImprimir_ 45, True
            End If
            
        Case 15: mnImprimir_Click 'Imprimir Albaran
        
                
        Case 16:
            'Mayo 2015.  Impresion albaranes ruta CASTELLON   HERBELCA
            ImpresionAlbaranRutaCliente
        Case 17: mnSalir_Click   'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Indice - btnPrimero)
            'Desplazamiento (Button.Index - btnPrimero)
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
Dim Sql As String
Dim vWhere As String
Dim B As Boolean
Dim vCStock As CStock
Dim ImpReciclado As Single
Dim DentroTRANS As Boolean
Dim ArtFitosnatiarios As Boolean

Dim SqlIntercalar As String
Dim SqlIntercalar2 As String


    InsertarLinea = False
    Sql = ""
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
        Sql = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 2 WHERE " & vWhere & " and numlinea >= " & LineaIntercalar
        Sql = Sql & " order by numlinea desc " 'Para que empieza por las ultimas
        SqlIntercalar = Sql
        
        'ENERO 2018. ERROR GRAVE
        ' No actualizaba la smoval
        Sql = "UPDATE smoval SET numlinea=numlinea + 2"
        Sql = Sql & "  WHERE detamovi ='" & Text1(30).Text & "' AND document ='" & Text1(0).Text & "' AND "
        Sql = Sql & " fechamov= " & DBSet(Text1(1).Text, "F") & " and numlinea >= " & LineaIntercalar
        Sql = Sql & " order by numlinea desc"
        SqlIntercalar2 = Sql
        
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
        Sql = "INSERT INTO " & NomTablaLineas
            Sql = Sql & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos,precioar, dtoline1, dtoline2, importel, origpre,codprovex,numlote,codccost,pvpInferior,comisionagente) "
        Sql = Sql & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        '- cantidad,numbultos
        Sql = Sql & DBSet(txtAux(3).Text, "N") & ", " & DBSet(txtAux(10).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(5).Text, "T", "N") & ","
        
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If ArtFitosnatiarios Then
                'Pongo un "*" al articulo para indicarme que lleva lotes
                txtAux(11).Text = "*"
            End If
        End If
        
        
        If vEmpresa.TieneAnalitica Then
            '- codprove,numlote,codccost
            Sql = Sql & "0," & DBSet(txtAux(11).Text, "T", "S") & "," & DBSet(txtAux(9).Text, "T", "S")
        Else
            '- codprove,numlote,codccost
            Sql = Sql & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(11).Text, "T", "S") & "," & ValorNulo
        End If
        Sql = Sql & "," & VendeAMenorPrecio & ","
        If vParamAplic.NumeroInstalacion = 2 Then
            Sql = Sql & DBSet(txtAux(12).Text, "N", "S")
        Else
            Sql = Sql & "null"
        End If
        Sql = Sql & ")"
        
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
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
        conn.Execute Sql
        
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
            Sql = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
            Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "'," & Data1.Recordset!Numalbar & "," & numlinea
            Sql = Sql & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
            Sql = Sql & " FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 "

            conn.Execute Sql
            
            'Tengo que updatear la cantidad vendida
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open "Select * FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad >0 ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux!cantidad <> 0 Then
                    If miRsAux!cantidad > 0 Then
                        Sql = "+"
                    Else
                        Sql = "-"
                    End If
                    Sql = "UPDATE slotes SET vendida=vendida " & Sql & DBSet(Abs(miRsAux!cantidad), "N")
                    Sql = Sql & " WHERE numlotes =" & DBSet(miRsAux!numlotes, "T") & " AND codartic= " & DBSet(miRsAux!codArtic, "T")
                    Sql = Sql & " AND fecentra= " & DBSet(miRsAux!FechaAlb, "F")
                
                    conn.Execute Sql
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
                    Sql = "INSERT INTO " & NomTablaLineas
                    Sql = Sql & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,  precioar,"
                    Sql = Sql & "dtoline1, dtoline2, importel, origpre,comisionagente) "
                    Sql = Sql & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    Sql = Sql & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                    Sql = Sql & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
                    Sql = Sql & DBSet(ImpReciclado, "N") & ",0,0,"
                    'Importe linea
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    Sql = Sql & DBSet(ImpReciclado, "N") & ", 'A',"
                    'Comision
                    ImpReciclado = 0
                    Sql = Sql & DBSet(ImpReciclado, "N") & ")"
                    conn.Execute Sql
                        
                    
                End If 'articulo con sunida reciclado
            End If  'Cliente con tasa reciclado
        End If 'ok actualiza stock
        
        If B Then
            If vParamAplic.PtosAsignar > 0 Then
                If txtAux(1).Text = vParamAplic.PtosArticuloCanje Then
                    
                    ImpReciclado = ImporteFormateado(txtAux(3).Text)
    
                    Sql = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", CStr(Data1.Recordset!codClien))
                    Sql = " VALUES (" & Data1.Recordset!codClien & "," & Val(Sql) + 1 & "," & DBSet(Data1.Recordset!codtipom, "T") & "," & Data1.Recordset!Numalbar
                    
                    Sql = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov)" & Sql
                    Sql = Sql & " ," & DBSet(Data1.Recordset!FechaAlb, "F") & ",1," & DBSet(ImpReciclado, "N") & ",now())"
                    conn.Execute Sql
                
                    
                    Sql = "UPDATE sclien set puntos=" & DBSet(ImpReciclado, "N") & " + coalesce(puntos,0) "
                    Sql = Sql & " WHERE codclien =" & Text1(4).Text
                    conn.Execute Sql
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
Dim Sql As String
Dim vCStock As CStock
Dim B As Boolean
Dim ImpReciclado As Single
Dim ArtFitosnatiarios As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    
    
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
                    If CStr(Data2.Recordset!codArtic) <> txtAux(1).Text Then
                        'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
                        Sql = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
                        Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
                        Sql = Sql & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                        Sql = Sql & " AND numalbar=" & Data1.Recordset!Numalbar & " AND numline1=" & Data2.Recordset!numlinea
                        conn.Execute Sql
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
                Sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
                Sql = Sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                Sql = Sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", numbultos=" & DBSet(txtAux(10).Text, "N") & ","
                Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", " 'precio
                Sql = Sql & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
                Sql = Sql & "importel= " & DBSet(txtAux(8).Text, "N") & ", " 'Importe
                Sql = Sql & "origpre=" & DBSet(txtAux(5).Text, "T", "S") & ","
                ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & "codccost=" & DBSet(txtAux(9).Text, "T", "S") & ","
                Else
                    Sql = Sql & "codprovex=" & DBSet(txtAux(9).Text, "N", "N") & ","
                End If
                Sql = Sql & "numlote=" & DBSet(txtAux(11).Text, "T", "S") & ", "
                
                'Junio2013
                Sql = Sql & "pvpInferior=" & DBSet(VendeAMenorPrecio, "N") & ""
                If vParamAplic.NumeroInstalacion = 2 Then Sql = Sql & " , comisionagente =" & DBSet(txtAux(12).Text, "N", "S")
                
                
                Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
                conn.Execute Sql
                
                                
                'Si ha cambiado, si tiene el parametro... todo esta ahi
                If hcoCodTipoM <> "DEV" Then TrataCambioPrecioDto
                
                'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       Sql = Mid(Sql, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       Sql = DevuelveDesdeBD(conAri, "codartic", "slialb", Sql, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                       If Sql = vParamAplic.ArtReciclado Then
                       
                            Sql = "UPDATE " & NomTablaLineas & " SET "
                            Sql = Sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
                            Sql = Sql & "precioar= " & DBSet(ImpReciclado, "N") & ", " 'precio
                            ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                            Sql = Sql & "importel= " & DBSet(ImpReciclado, "N")  'Importe
                            'WHERE
                            Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                            conn.Execute Sql
                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
                
                If vParamAplic.PtosAsignar > 0 Then
                    If Data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
                        'Ha cambiado el articulo de puntos
                        'Actualizamos los puntos en el cliente, el el movimputs
                        ImpReciclado = Data2.Recordset!cantidad
                        ImpReciclado = ImporteFormateado(txtAux(3).Text) - ImpReciclado
                        Sql = "UPDATE sclien set puntos=" & DBSet(ImpReciclado, "N") & " + coalesce(puntos,0) "
                        Sql = Sql & " WHERE codclien =" & Text1(4).Text
                        conn.Execute Sql
                                                
                        Sql = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
                        Sql = Sql & " AND codclien = " & Data1.Recordset!codClien & " AND concepto=1"
                        Sql = "UPDATE smovalpuntos  set puntos = " & DBSet(ImporteFormateado(txtAux(3).Text), "N") & Sql
                        conn.Execute Sql
                                                
                                                
                                                
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
Dim Sql As String
    
    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
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
    vDataGrid.Columns(2).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            i = 3
            vDataGrid.Columns(i).Caption = "Alm."
            vDataGrid.Columns(i).Width = 470
            vDataGrid.Columns(i).NumberFormat = "000"
            
            i = i + 1 '4
            vDataGrid.Columns(i).Caption = "Articulo"
            vDataGrid.Columns(i).Width = 1700
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Desc. Artículo"
            vDataGrid.Columns(i).Width = 3500

            i = 6
            vDataGrid.Columns(i).visible = False
            i = 7
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 850
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = 8
            vDataGrid.Columns(i).Caption = "Bultos"
            vDataGrid.Columns(i).Width = 650
            vDataGrid.Columns(i).Alignment = dbgRight
                
            i = i + 1 '9
            vDataGrid.Columns(i).Caption = "Precio"
            vDataGrid.Columns(i).Width = 950
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoPrecio
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "OP"
            vDataGrid.Columns(i).Width = 350
            vDataGrid.Columns(i).Alignment = dbgCenter
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto. 1"
            vDataGrid.Columns(i).Width = 600
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto. 2"
            vDataGrid.Columns(i).Width = 600
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Importe lin"
            vDataGrid.Columns(i).Width = 1100
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            If vEmpresa.TieneAnalitica Then
                i = i + 1
                vDataGrid.Columns(i).Caption = "CCost"
                vDataGrid.Columns(i).Width = 680
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
                vDataGrid.Columns(i).Width = 680
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
                        If CStr(Data2.Recordset!tipoprecio) <> "*" Then
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
        alto = ObtenerAlto(DataGrid1, 10)
        
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
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 20
            txtAux(i).Width = DataGrid1.Columns(i + 5).Width - 20
        Next i
        
        'El boton 3 lo superpongo un poquito
'        cmdAux(9).Left = txtAux(10).Left - 15
        If cmdAux(9).visible Then
            txtAux(9).Width = txtAux(9).Width - cmdAux(9).Width
            cmdAux(9).Left = txtAux(9).Left + txtAux(9).Width
            txtAux(11).Left = cmdAux(9).Left + cmdAux(9).Width + 20 'numlote
        Else
            txtAux(11).Left = txtAux(9).Left + txtAux(9).Width + 20 'numlote
        End If
        
        '- numlote
'        txtAux(11).Left = cmdAux(9).Left + cmdAux(9).Width + 20
        If vEmpresa.TieneAnalitica Then
            txtAux(11).Width = DataGrid1.Columns(15).Width - 20
        Else
            txtAux(11).Width = DataGrid1.Columns(16).Width - 20
        End If
        
        
        'Solo herbelca
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            txtAux(12).Left = DataGrid1.Columns(14).Left + 210
            txtAux(12).Width = DataGrid1.Columns(14).Width - 9
            BloquearTxt txtAux(12), True
            If Not MostrarComision Then txtAux(12).Top = 6000
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


Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
        
        BloquearTxt txtAux(6), False
        BloquearTxt txtAux(7), False
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
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

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
            AbrirConsultaPrecio2 Text1(4).Text, txtAux(1).Text, Text1(1).Text, Text1(13).Text
        Else
            If Index = 6 Or Index = 7 Then AbrirFormularioDtos txtAux(1).Text
        End If
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
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
Dim Aux3 As String
Dim PtosAuxiliar As Currency



    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
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
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
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
                        If vParamAplic.NumeroInstalacion = vbHerbelca Then
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
                        BloquearTxt txtAux(4), True
                        Me.txtAux(6).Text = Format(miRsAux!dtoline1, FormatoDescuento)
                        BloquearTxt txtAux(6), True
                        Me.txtAux(7).Text = Format(miRsAux!dtoline2, FormatoDescuento)
                        BloquearTxt txtAux(7), True
                        Me.txtAux(8).Text = Format(-miRsAux!ImporteL, FormatoDescuento)
                        
                        
                        If vEmpresa.TieneAnalitica Then
                            txtAux(9).Text = DBLet(miRsAux!CodCCost, "T")
                        Else
                            txtAux(9).Text = DBLet(miRsAux!codProvex, "N")
                        End If
                        Me.txtAux(5).Text = miRsAux!origpre
                        txtAux(11).Text = DBLet(miRsAux!numLote, "T")
                        txtAux(10).Text = Abs(miRsAux!NumBultos)
                        
                        TieneDescuentos = "concat(dtoppago,'|',dtognral,'|')"
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
                            devuelve = DevuelveDesdeBD(conAri, "fechaalb", "scaalb", devuelve, miRsAux!Numalbar, "N", TieneDescuentos)
                            devuelve = "Rectifica al albarán: " & miRsAux!codtipom & Format(miRsAux!Numalbar, "000000") & " de " & Format(devuelve, "dd/mm/yyyy")
                        End If
                        Text2(16).Text = CStr(devuelve)
                        If TieneDescuentos = "" Then TieneDescuentos = "0|0|"
                        devuelve = ""
                        
                        
                        If CCur(Replace(RecuperaValor(TieneDescuentos, 1), ".", ",")) <> Data1.Recordset!DtoPPago Then
                            devuelve = "Descuento pronto pago" & vbCrLf
                            devuelve = devuelve & "Origen: " & RecuperaValor(TieneDescuentos, 1)
                            devuelve = devuelve & "       Devolucion: " & Data1.Recordset!DtoPPago & vbCrLf
                        End If
                        If CCur(Replace(RecuperaValor(TieneDescuentos, 2), ".", ",")) <> Data1.Recordset!DtoGnral Then
                           devuelve = devuelve & IIf(devuelve = "", "", vbCrLf & vbCrLf) & "Descuento general" & vbCrLf
                           devuelve = devuelve & "Origen: " & RecuperaValor(TieneDescuentos, 2)
                           devuelve = devuelve & "     Devolucion: " & Data1.Recordset!DtoGnral
                        End If
                        If devuelve <> "" Then
                        
                            'Si es la primera LINEA, los inserto directamente
                            If Me.Data2.Recordset.EOF Then
                                Text1(15).Text = Replace(RecuperaValor(TieneDescuentos, 1), ".", ",")
                                Text1(16).Text = Replace(RecuperaValor(TieneDescuentos, 2), ".", ",")
                                Aux3 = ""
                            Else
                                MsgBox devuelve, vbExclamation
                            End If
                        End If
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
                                cantidad = CalcularPuntosAlbaranCABEL(Replace(Replace(ObtenerWhereCP(False), "scaalb", "slialb"), NombreTabla, NomTablaLineas), Data1.Recordset!FechaAlb, Aux3, "")
                                If cantidad = "" Then cantidad = "0"
                                
                                'Siginifica que tiene articvulos CABEL
                                If CCur(cantidad) > 0 Then
                                    'Tiene cabel. Vamos a ver cuantos puntos necesita como maximo para el importe de este albaran
                                    PtosAuxiliar = Round2(CCur(Aux3) / vParamAplic.PtosEquivalencia, 2) '-> necesito como mucho estos puntos
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
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then B = True
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
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then
                         txtAux(5).Text = "M"
                         BloquearTxt txtAux(6), False
                         BloquearTxt txtAux(7), False
                    End If
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


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    Me.SSTab1.Tab = numTab
    TituloLinea = Cad
    ModificaLineas = 0
    
    Cad = "comision"
    ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text, "N", Cad)) = 1
    If vParamAplic.ArtReciclado = "" Then ClienteConTasaReciclado = False
    
    If Cad = "" Then Cad = "0"
    ComisionCliente = CCur(Cad)
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
Dim Sql As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim MenError As String
Dim ParaElLog As String

    On Error GoTo FinEliminar
    conn.BeginTrans
    
    Sql = ObtenerWhereCP(False)
    
    MenError = DevuelveDesdeBD(conAri, "concat(count(*),'|',sum(importel),'|')", "slialb", Replace(Sql, "scaalb", "slialb") & " AND 1", "1")
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
    
    
    Screen.MousePointer = vbHourglass
    If B Then
        'eliminamos de albaranes y pasamos al historico
        'Para los DEV NO
        If CodTipoMov <> "DEV" Then
            B = ActualizarElTraspaso(MenError, Sql, CodTipoMov, cadList)
        Else
            'Borramos de scaalb
            Sql = ObtenerWhereCP(True)
            conn.Execute "DELETE from slialb " & Replace(Sql, "scaalb", "slialb")
            conn.Execute "DELETE from scaalb " & Sql
            
        End If
        
        If B Then
            MenError = "Actualizando numeros de serie."
            'Actualizar los posibles num. serie de ese albaran. vaciar los campos
            Sql = "UPDATE  sserie SET codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
            Sql = Sql & " fechavta=" & ValorNulo & ", numalbar=" & ValorNulo & ", numline1=" & ValorNulo
            Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
            Sql = Sql & " WHERE codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!Numalbar & " AND fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
            conn.Execute Sql
            
            
           Screen.MousePointer = vbHourglass
            
            
            
            If B Then
                'Actualiamos el riesgo
                If CodTipoMov <> "DEV" Then
                    If vParamAplic.OperacionesAseguradas Then
                        lblIndicador.Caption = "Riesgo"
                         lblIndicador.Refresh
                        Sql = DevuelveDesdeBD(conAri, "credipriv", "sclien", "codclien", Text1(4).Text, "N")
                        If Sql = "" Then Sql = "9"
                        If Val(Sql) < 9 Then
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
                Sql = DevuelveDesdeBD(conAri, "tienePuntos", "sclien", "codclien", Text1(4).Text)
                If Val(Sql) = "1" Then
                    'El nuevo cliente tiene puntos
                    Sql = "-"
                    If Data1.Recordset!Puntos < 0 Then Sql = "+"
                    Sql = "UPDATE sclien set puntos=coalesce(puntos,0) " & Sql & DBSet(Abs(Data1.Recordset!Puntos), "N")
                    Sql = Sql & " WHERE codclien =" & Text1(4).Text
                    conn.Execute Sql
                End If
            
            End If
            
            'Borro de smovalpuntos
            Sql = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
            Sql = Sql & " AND codclien = " & Data1.Recordset!codClien
            Sql = Sql & " AND concepto = 0"
            
                
            Sql = "DELETE FROM smovalpuntos " & Sql
            conn.Execute Sql
            
            
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

    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData(CambiarModo As Boolean)
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
            If CambiarModo Then
                PonerModo 2
                lblIndicador.Caption = Indicador
            End If
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
Dim Sql As String

    On Error Resume Next
    
    Sql = " " & NombreTabla & ".codtipom= '" & Text1(30).Text & "' and " & NombreTabla & ".numalbar= " & Val(Text1(0).Text)
    If EsHistorico Then Sql = Sql & " AND " & NombreTabla & ".fechaalb=" & DBSet(Text1(1).Text, "F")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then Err.Clear
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
    
    
    'Enero 2008. David
    'Para la trazabilidad con repescto al codproveedor en las lineas
    Sql = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & ",codccost"
    Else
        'COTUBRE
        If vParamAplic.NumeroInstalacion = 2 Then
            
            If MostrarComision Then
                Sql = Sql & ",comisionagente,'' nomprove"   'LO que habia antes de Sep 2018
            Else
                Sql = Sql & ",codprovex,'' nomprove"   'NUEVO
            End If
        Else
            Sql = Sql & ",codprovex,nomprove"
        End If
    End If
    Sql = Sql & ",numlote "
    
    
    'Para herbelca, ponemos el codprove al final
    If vParamAplic.NumeroInstalacion = 2 Then
        Sql = Sql & ",codprovex,if(pvpinferior=0,'',if(pvpinferior=1,'*','m')) tipoprecio"
        If Not MostrarComision Then Sql = Sql & ",comisionagente"   'para updaear los campos
    End If
 
    
    
    Sql = Sql & " FROM " & NomTablaLineas
    'traza
    If vEmpresa.TieneAnalitica = False Then
        Sql = Sql & " LEFT JOIN sprove on codprovex=codprove "
    End If
    
    If enlaza Then
        Sql = Sql & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then Sql = Sql & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        Sql = Sql & " WHERE false " 'numalbar = -1"
    End If
    Sql = Sql & " Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0))
        'Insertar
        Toolbar1.Buttons(5).Enabled = (B Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (B Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = B And Not EsHistorico
        Me.mnModificar.Enabled = B And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = B And Not EsHistorico
        Me.mnEliminar.Enabled = B And Not EsHistorico
            
        B = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = B
        Me.mnLineas.Enabled = B
        
        'Nº Series
        Toolbar1.Buttons(11).Enabled = B And Not EsHistorico
        
        'Generar Factura
        'DAVID###
        'Antes:
        'Toolbar1.Buttons(12).Enabled = b And (CodTipoMov = "ALM" Or CodTipoMov = "ART")
        'Ahora.  Cualquier tipo se puede generar la factura
        Toolbar1.Buttons(12).Enabled = B
        Toolbar1.Buttons(13).Enabled = B
        If Toolbar1.Buttons(14).Style = tbrDefault Then Toolbar1.Buttons(14).Enabled = B
        
        
        'Imprimir
        Toolbar1.Buttons(15).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        If vParamAplic.TipoPortes = 1 Then
            Toolbar1.Buttons(14).Enabled = Toolbar1.Buttons(15).Enabled And vParamAplic.TipoPortes = 1
        Else
            If Toolbar1.Buttons(14).Style = tbrDefault Then Toolbar1.Buttons(14).Enabled = B
        End If
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
        
        'Insertar intercalando
        If Modo = 5 Then
            B = (ModificaLineas = 0)
            Toolbar1.Buttons(11).Image = 34 '.Buttons(11).Image = 26
            Toolbar1.Buttons(11).ToolTipText = "Insertar intercalando"
            '
            Toolbar1.Buttons(13).ToolTipText = "Lotes articulo"
            Toolbar1.Buttons(13).Image = 52
            Toolbar1.Buttons(13).Enabled = B And vParamAplic.ManipuladorFitosanitarios2
            
            
            
            
            
            
        Else
            'b=modo=2
            B = B And Not EsHistorico
            Toolbar1.Buttons(11).Image = 33
            Toolbar1.Buttons(11).ToolTipText = "Nº de serie"
            
            Toolbar1.Buttons(13).ToolTipText = "Marcar para facturar"
            Toolbar1.Buttons(13).Image = 30
            Toolbar1.Buttons(13).Enabled = Modo <= 2 And Not EsHistorico
        End If
        Toolbar1.Buttons(11).Enabled = B
        
        
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
Dim Sql As String
Dim B As Boolean
Dim ImpReciclado As Single



    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
    
    
    
    'Inicilizar la clase para Actualizar los stocks
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""
    
    
    On Error GoTo EEliminarLinea
    
    conn.BeginTrans
    conn.Execute Sql 'Eliminar linea
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
                    Sql = CStr(Data2.Recordset!codArtic)
                    If ArticuloConTasaReciclado(Sql, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       Sql = Mid(Sql, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       Sql = DevuelveDesdeBD(conAri, "codartic", "slialb", Sql, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                       If Sql = vParamAplic.ArtReciclado Then
                       
                            Sql = "DELETE FROM " & NomTablaLineas
                            'WHERE
                            Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                            conn.Execute Sql
                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
    End If


    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
    If hcoCodTipoM <> "DEV" Then
        Sql = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
        Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
        Sql = Sql & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
        Sql = Sql & " AND numalbar=" & Data1.Recordset!Numalbar & " AND numline1=" & Data2.Recordset!numlinea
        conn.Execute Sql
    End If
    
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        ReestablecerLotesArticulo Data2.Recordset!numlinea
        
        'Borramos de slialblotes
        Sql = "Delete from slialblotes " & Replace(ObtenerWhereCP(True), NombreTabla, "slialblotes")
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute Sql 'Eliminar linea
    End If
        

    
    
    If vParamAplic.PtosAsignar > 0 Then
        
        'Si es la linea de canje, hay que quitarla de movimientos
        If Data2.Recordset!codArtic = vParamAplic.PtosArticuloCanje Then
            Sql = Replace(ObtenerWhereCP(True), "scaalb", "smovalpuntos")
            Sql = Sql & " AND codclien = " & Data1.Recordset!codClien & " AND concepto=1"
            Sql = "DELETE FROM smovalpuntos  " & Sql
            conn.Execute Sql
      
            Sql = "UPDATE sclien set puntos=" & DBSet(-1 * Data2.Recordset!cantidad, "N") & " + coalesce(puntos,0) "
            Sql = Sql & " WHERE codclien =" & Text1(4).Text
            conn.Execute Sql
        End If
    End If
    
    Sql = "Albarán: " & Text1(30).Text & "-" & Text1(0).Text & " de " & Text1(1).Text & vbCrLf
    Sql = Sql & "Linea: " & Data2.Recordset!codArtic & " " & DBSet(Data2.Recordset!NomArtic, "T")
    Sql = Sql & "Uds: " & Data2.Recordset!cantidad & "    Importe:" & DBSet(Data2.Recordset!ImporteL, "N")

    Set LOG = New cLOG
    ' 17 Venta a sabiendas riesgo
    LOG.Insertar 37, vUsu, Sql
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
            BuscaChekc = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", Data2.Recordset!codArtic, "T")
        Else
            BuscaChekc = "OK"
        End If
        If Trim(BuscaChekc) <> "" Then
            Set miRsAux = New ADODB.Recordset
            BuscaChekc = "Select * from slialblotes WHERE codtipom= '" & Data1.Recordset!codtipom & "' AND numalbar = " & Data1.Recordset!Numalbar
            If linea >= 0 Then BuscaChekc = BuscaChekc & " AND numlinea =" & Data2.Recordset!numlinea
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
            If Data2.Recordset!codArtic = txtAux(1).Text Then
                vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text)) - Data2.Recordset!cantidad
            Else
                vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
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

Private Function ComprobarInventario() As Boolean
Dim CadFecInv As String
Dim vCStock As CStock

    ComprobarInventario = True
    CadFecInv = ""
    If Data2.Recordset.RecordCount > 0 Then
       Data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not Data2.Recordset.EOF)
           Set vCStock = New CStock
           If InicializarCStock(vCStock, "E", Data2.Recordset!numlinea) Then vCStock.ComprobarFechaInventario False, CadFecInv
           Set vCStock = Nothing
           
           Data2.Recordset.MoveNext
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
    
    If Data2.Recordset.RecordCount > 0 Then
       Data2.Refresh
       Data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not Data2.Recordset.EOF) And B
           Set vCStock = New CStock
           If InicializarCStock(vCStock, "E", Data2.Recordset!numlinea) Then
                
               'Actualiza el stock en salmac y borra de smoval
               If Not vCStock.DevolverStock2() Then B = False
           Else
               B = False
           End If
           Data2.Recordset.MoveNext
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
Dim HabiaImpresoYaElAlbaran As Boolean
Dim ImprimeFluorados As Boolean


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
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
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
    cadParam = cadParam & "Albarcon=" & devuelve & "|"  'NO TOCAR. En frmimprimir hay una vinculacion
    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    
    
    
    If Not ImpresionDirecta Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If
        
    ImprimeFluorados = False
    
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
                    HabiaImpresoYaElAlbaran = chkImpreso.Value = 1
                    If HaPulsadoElBotonDeImprimir Then
                    
                    
                        
                        
                        'UPDATEAMOS scaalb para que no reimpimrpima los albaranes
                        'Cod Tipo Movimiento
                        If Not HabiaImpresoYaElAlbaran Then
                            devuelve = "scaalb.codtipom = '" & CodTipoMov & "' AND scaalb.numalbar = " & Val(Text1(0).Text)
                            devuelve = "UPDATE scaalb SET albImpreso = 1 WHERE " & devuelve
                            Me.chkImpreso.Value = 1
                            ejecutar devuelve, False
                        End If
                        
                        
                        'Cunado imprimen un albaran. Lo marca, erpo no actualiza el Data1.
                        ' Se van al siguiente Y VUELVEN, y modifican algo de la cabecera, vuelve a dejar como NO impreso
                        If vParamAplic.NumeroInstalacion = vbHerbelca Then
                            If HaPulsadoElBotonDeImprimir And Not HabiaImpresoYaElAlbaran Then PosicionarData True
                        End If
                                                
                    End If
                    
                    'Solo herbelca
                    If vParamAplic.NumeroInstalacion = vbHerbelca Then
                        devuelve = "codtipom = '" & CodTipoMov & "' AND numalbar = " & Val(Text1(0).Text) & " AND numfactu is null AND 1"
                        devuelve = DevuelveDesdeBD(conAri, "count(*)", "sserie", devuelve, "1")
                        
                        If Val(devuelve) > 0 Then
                            'Lleva numeros de serie
                            'Veremos si alguno es de gases fluorados. tipar:
                            devuelve = "slialb inner join sartic on slialb.codartic=sartic.codartic inner join stipar on sartic.codtipar=stipar.codtipar"
                            
                            devuelve = DevuelveDesdeBD(conAri, "count(*)", devuelve, "codtipom = '" & CodTipoMov & "' AND numalbar = " & Val(Text1(0).Text) & " and impresionauxiliar", "1")
                            If devuelve = "" Then devuelve = 0
                        
                        End If
                        
                        If Val(devuelve) > 0 Then
                            If HabiaImpresoYaElAlbaran Then
                                'pregunta
                                If MsgBox("Desea volver a imprimir hoja ""gases fluorados""", vbQuestion + vbYesNoCancel) <> vbYes Then devuelve = ""
                            End If
                            
                            If devuelve <> "" Then ImprimeFluorados = True
                            
                        End If
                    End If
                    
                    
                End If
            End If
                                
        End With
        
        If ImprimeFluorados Then
            With frmImprimir
                
                .outTipoDocumento = 0
                .NombreRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", 102) ' "HerFluorados.rpt"
                .NombrePDF = .NombreRPT
                .NumeroCopias = 1
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = OpcionListado
                
                .Titulo = "Gases fluorados "
                
                .ConSubInforme = True
                .Show vbModal
            End With
        End If
        
        If indRPT <> 34 And Not EsHistorico Then
            If vParamAplic.NumeroInstalacion = vbHerbelca Then
                    
                cadSelect = ObtenerWhereCP(False) & " AND 1"
                cadSelect = DevuelveDesdeBD(conAri, "codnatura", "scaalb", cadSelect, "1")
                
                If Val(cadSelect) > 0 Then
                    
                    With frmImprimir
            
                        .outTipoDocumento = 0
                        .NombreRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", 101) ' "HerAlbaranPrecursor.rpt"
            
                        .FormulaSeleccion = cadFormula
                        .OtrosParametros = cadParam
                        .NumeroParametros = numParam
                        .SoloImprimir = False
                        .EnvioEMail = False
                        .Opcion = 250
                        
                        .Titulo = "Precursor explosivos"
                        
                        .ConSubInforme = True
                        .Show vbModal
            
                    End With
                End If
            End If
        End If
        
        
        
    End If
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset, Optional Dif As String, Optional cadSel As String)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'Dif: si se ha modificado la cantidad pasamos la difencia con lo que habia
Dim Sql As String
Dim Campos As String

    On Error GoTo EMostrarNSeries

    If Text1(30).Text = "ART" Then
        Sql = MostrarNSeriesGnral(RSLineas, Campos, True)
    Else
        Sql = MostrarNSeriesGnral(RSLineas, Campos)
    End If
    
   If Sql <> "" Then
        Set frmMen = New frmMensajes
        frmMen.cadWhere = Sql
        
        If Dif <> "" Then
            Sql = " WHERE (codtipom=" & DBSet(CodTipoMov, "T") & " and "
            Sql = Sql & " numalbar=" & Text1(0).Text & " and numline1=" & Data2.Recordset!numlinea & ")"
            frmMen.cadWHERE2 = Dif & "|" & Sql & "|"
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
Dim Sql As String

    On Error GoTo EPedirNSeries

        Sql = "El artículo tienen control de Nº de Serie." & vbCrLf & vbCrLf
        Sql = Sql & "Introduzca los Nº De Serie." & vbCrLf
        MsgBox Sql, vbInformation
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
Dim Sql As String
    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
    
        'Herbelca , control puntos
        'If vParamAplic.PtosAsignar > 0 Then Text1(43).Text = CalcularPuntosAlbaran
    
    
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        If Val(Text1(0).Text) <= 0 Then
            MsgBox "Error obteniendo contador", vbExclamation
            Sql = ""
        Else
            Sql = CadenaInsertarDesdeForm(Me)
        End If
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
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
Dim Sql As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
Dim Dif As Single

    'Comprobar si el Articulo tiene control de Nº de Serie
    Sql = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If Sql = "1" Then 'Hay NºSerie para el Articulo
        'si estamos insertando
        If Modo = 5 Then
            If ModificaLineas = 1 Then 'Insertar
                'Comprobar que la cantidad comprada es >0
                If ComprobarCero(txtAux(3).Text) <= 0 Then Exit Sub
            ElseIf ModificaLineas = 2 Then 'Modificar
                'si se ha modificado la cantidad, habrá que quitar algun nº serie
                'de los seleccionado o anyadir alguno mas
                Dif = CSng(txtAux(3).Text) - CSng(Data2.Recordset!cantidad)
                If Dif = 0 Then Exit Sub
                If Text1(30).Text = "ART" Then Exit Sub
'                    Dif = CSng(Data2.Recordset!Cantidad) - CSng(txtAux(3).Text)
            End If
        End If
        
        cadWhere = " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and "
        cadWhere = cadWhere & " numalbar=" & Text1(0).Text & " and numlinea=" & numlinea
    
        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
        Sql = "SELECT slialb.codartic, sum(cantidad) as cantidad, numlinea "
        Sql = Sql & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
        Sql = Sql & cadWhere & " And nseriesn = 1 "
        Sql = Sql & " GROUP BY codartic ORDER BY Codartic "

        Set RSLineas = New ADODB.Recordset
        RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
                Sql = " select distinct count(numserie) from sserie where codartic=" & DBSet(txtAux(1).Text, "T") & " and (numalbar='' or isnull(numalbar))"
                '=== Laura 17/01/2007
                'y que no este asignados a una factura
                Sql = Sql & " and (numfactu='' or isnull(numfactu))"
                '===
                If RegistrosAListar(Sql) = 0 Then 'No hay Nº de Serie para elegir
                    PedirNSeries RSLineas
                Else
                    MostrarNSeries RSLineas
                End If
            ElseIf ModificaLineas = 2 Then
                Sql = " select distinct count(numserie) from sserie " & Replace(cadWhere, "numlinea", "numline1")
                If RegistrosAListar(Sql) > 0 Then
                    MostrarNSeries RSLineas, CStr(Dif)
                End If
            End If
        End If

        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub BotonNSeries()
Dim cadWhere As String, Sql As String
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
    Sql = "SELECT numlinea, slialb.codartic, sum(cantidad) as cantidad "
    Sql = Sql & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    Sql = Sql & cadWhere & " And nseriesn = 1 "
    
    'Pudioera ser que tuvieran un mismo articulo wen dos lineas, y por lo tanto
    'el articulo tendria numeros de sr asociados a distintas lineas
    'por lo tanto hay que agrupar por numlinea TB
    Sql = Sql & " GROUP BY codartic,numlinea ORDER BY Codartic "
    

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim Sql As String
Dim linea As Integer

    On Error GoTo EPedirNSeries


        PedirNSeriesGnral Rs, False
        Rs.MoveFirst
        While Not Rs.EOF
            linea = 0
            'Cargar los Nº de serie asignados
            Sql = "SELECT numserie, codartic,nummante FROM sserie "
            Sql = Sql & " WHERE codtipom='" & Text1(30).Text & "' and "
            Sql = Sql & "numalbar=" & Text1(0).Text
            Sql = Sql & " and numline1=" & Rs!numlinea
            Sql = Sql & " ORDER BY codartic "
            Set RSseries = New ADODB.Recordset
            RSseries.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                Sql = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                Sql = Sql & ", nummante = " & DBSet(RSseries!nummante, "T")
                Sql = Sql & " WHERE codartic=" & DBSet(Rs!codArtic, "T")
                Sql = Sql & " and numlinealb=" & Rs!numlinea
                Sql = Sql & " and numlinea=" & linea
                conn.Execute Sql
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
Dim Sql As String
Dim B As Boolean

    On Error GoTo ECargar
    
    conn.BeginTrans
    
    'Limpiar primero los Nº de serie asignados al ALV y luego volver a cargarlos
    Sql = "UPDATE sserie SET codtipom=" & ValorNulo & ", numalbar=" & ValorNulo & ", fechavta="
    Sql = Sql & ValorNulo & ", numline1=" & ValorNulo
    'Enero 2010
    'Tambien reestablezco los valores de tieneman y numeromantenimiento
     Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo & ", coddirec= " & ValorNulo
    
    Sql = Sql & " WHERE codtipom=" & DBSet(Text1(30).Text, "T") & " and numalbar=" & Text1(0).Text & " AND year(fechavta)=" & Year(Text1(1).Text)
    conn.Execute Sql
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    Sql = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    Sql = Sql & " ORDER BY codartic, numlinealb, numlinea "
    Set RStmp = New ADODB.Recordset
    RStmp.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
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
Dim Numalbar As String
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
     Numalbar = "numalbar" 'Nº albaran de Venta
     devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", Numalbar, "codartic", codArtic, "T")
     If devuelve <> "" Then 'EXISTE en tabla sserie
        If Numalbar = "" Then B = nSerie.ActualizarNumSerie(True)
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
            SoloEnEfectivoAlbaranes = False
            If vCliente.ClienteBloqueado(0, SoloEnEfectivoAlbaranes) Then
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
            Text1(4).Text = vCliente.Codigo
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
                
                vCliente.PonDatosDireccionEnvio Text1(42), Text2(42)
                
                
                'Febrero 2013
                'Si tiene observaciones del departamento de comercial, van a observaCRM
                Text1(44).Text = DevuelveDesdeBD(conAri, "observa", "scrmobsclien", "dpto=2 AND codclien", codClien)
                
            End If
            Me.Text1(45).Text = "": Me.Text1(46).Text = "": Me.Text1(47).Text = "": Me.Text1(48).Text = "": Text2(0).Text = ""
            
           
             Me.chkFacturar.Value = vCliente.MarcaFacturar()
             
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
    vCliente.Codigo = Text1(4).Text
    B = vCliente.ActualizaUltFecMovim(Text1(1).Text)
    Set vCliente = Nothing
    
EActFecha:
    If Err.Number <> 0 Then B = False
    ActualizarFecMovCliente = B
End Function


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim cadWhere As String, Sql As String
Dim vFactu As CFactura
Dim CambiarValoresIVA As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 56
         Text3(i).Text = ""
    Next i

    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP(False)
    Sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(Sql) = 0 Then Exit Sub
    
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



 Private Sub InsertarLineasFactu(cadWhere)
'cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
'cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
 Dim Rs As ADODB.Recordset
 Dim Sql As String
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
    
        Sql = "SELECT * FROM slifac WHERE " & cadWhere
        Set RL = New ADODB.Recordset
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                    Sql = "SELECT * FROM slifaclotes WHERE " & cadWhere
                    Sql = Sql & " AND numlinea= " & Rs!numlinea
                    
                    RL.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RL.EOF
                        Sql = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic) VALUES "
                        Sql = Sql & "('" & Text1(30).Text & "', " & Val(Text1(0).Text) & "," & Rs!numlinea & ","
                        Sql = Sql & RL!sublinea & "," & DBSet(-RL!cantidad, "N") & ","
                        Sql = Sql & DBSet(RL!numLote, "T") & "," & DBSet(RL!fecentra, "F") & "," & DBSet(Rs!codArtic, "T") & ")"
                        If ejecutar(Sql, False) Then
                            Sql = "UPDATE slotes SET vendida=vendida "
                            If RL!cantidad > 0 Then
                                Sql = Sql & " - "
                            Else
                                Sql = Sql & " + "
                            End If
                            Sql = Sql & DBSet(Abs(RL!cantidad), "N") & " WHERE codartic = " & DBSet(Rs!codArtic, "T")
                            Sql = Sql & " AND numlotes = " & DBSet(RL!numLote, "T")
                            Sql = Sql & " AND fecentra = " & DBSet(RL!fecentra, "F")
                            ejecutar Sql, False
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
Dim Sql As String
Dim cadWhere As String
Dim Rs As ADODB.Recordset

    On Error GoTo EDevNumSerie
        
        
    If IsNull(Me.Data1.Recordset!Numfactu) Then Exit Sub
        
    cadWhere = ObtenerWhereCP(True)
    Sql = "select slialb.codartic,abs(cantidad) as cantidad,numlinea"
    Sql = Sql & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
    Sql = Sql & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu and sserie.codclien=scaalb.codclien "
    '-- LAURA: 02/07/2007
'    SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'    SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
    Sql = Sql & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!Numfactu)
'    If Me.Data1.Recordset!codtipmf = "FAV" Then SQL = SQL & " AND codtipom='ALV'"
    '--

    
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Hay articulos con nº de serie en las lineas del albaran rectificativo
    'que hay que quitar de los nº de serie que tenia asignados
    'estamos devolviendo nº serie y pedimos los que vamos a devolver y a estos
    'le limpiamos los campos de venta de la tabla sserie
    If Not Rs.EOF Then
        Sql = "select sserie.numserie, sserie.codartic, sartic.nomartic"
        Sql = Sql & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
        '-- LAURA: 02/07/2007
'        SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'        SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
        Sql = Sql & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu  and sserie.codclien=scaalb.codclien "
        '--
        Sql = Sql & " inner join sartic on sserie.codartic=sartic.codartic "
        Sql = Sql & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!Numfactu)
    
        MostrarNSeries Rs, , Sql
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
        If CadenaDesdeOtroForm = "OK" And Modo <> 0 Then
            'OK. Cambiadas las marcas. Refrescamos y situamos
            Screen.MousePointer = vbHourglass
            DoEvents
            PonerCadenaBusqueda
            PosicionarData True
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
    If Data2.Recordset.EOF Then Exit Function
    
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
Dim Cad As String
Dim R As ADODB.Recordset
Dim NuevoDto As Currency
Dim Importe As Currency
Dim bAct As Boolean

    On Error GoTo EDescuentosCantidad
    
    If Not vParamAplic.DtoxCantidad Then Exit Sub ' ---- [14/09/2009] (LAURA)
     
    If MsgBox("¿Desea recalcular los descuentos por cantidad?", vbQuestion + vbYesNo) = vbYes Then    'masl 140909
    
        
        'Si no  tenemos portes, ni nos pasamos
    '    If vParamAplic.ArtPortes = "" Then Exit Sub
        
        Importe = 0
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        Cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        'cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from slialb " & cad
        'If Articulo <> "" Then cad = cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        Cad = "select codartic,sum(cantidad) suma from slialb " & Cad
        If Articulo <> "" Then Cad = Cad & " AND codartic= " & DBSet(Articulo, "T")
        
        
        'Y origen PRECIO no es precio especial
        Cad = Cad & " AND origpre <> 'E'"
        Cad = Cad & " group by 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                Cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                Cad = "select * from sdesca where desdecan <=" & Cad & " and " & Cad & " <= hastacan and envagran = "
                Cad = Cad & DBSet(miRsAux!codArtic, "T")
                R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Cad = ""
                If Not R.EOF Then Cad = R!dtolinea
                R.Close
                
                
                If Cad <> "" Then
                    'OK tiene nuevo descuento
                    NuevoDto = CCur(Cad)
                    
                    'Cojo los articulos del albaran y le meto el dto
                    Cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                    Cad = "select * from slialb " & Cad
                    '
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
                            Cad = "UPDATE slialb set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            Cad = Cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            Cad = Cad & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                            Cad = Cad & " and numlinea = " & R!numlinea
                            conn.Execute Cad
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
    
    
    
    If Importe <> 0 Then
        CargaGrid DataGrid1, Data2, True
        CalcularDatosFactura
    End If
    
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub





Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    

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
                        If Impo > CCur(vPreFact.Descuento1) Then
                            GrabaLogCambioPrecioDto = True
                        Else
                            Impo = ImporteFormateado(txtAux(7).Text)
                            If Impo > CCur(vPreFact.Descuento2) Then GrabaLogCambioPrecioDto = True
                        End If
                        
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
    
    If vParamAplic.NumeroInstalacion = vbHerbelca Then
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



        cadList = cadList & "(select codcampo from slialbcampos where numalbar=" & Data1.Recordset!Numalbar
        cadList = cadList & " AND "


        cadList = "codtipom = " & DBSet(Data1.Recordset!codtipom, "T")
        cadList = cadList & " AND numalbar"
        cadList = DevuelveDesdeBD(conAri, "max(numlinea)", "slialbcampos", cadList, CStr(Data1.Recordset!Numalbar), "N")
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
                        motivo = motivo & ", (" & Data1.Recordset!Numalbar & ","
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
    cadList = cadList & "(select codcampo from slialbcampos where numalbar=" & Data1.Recordset!Numalbar
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
    vCStock.HoraMov = Text1(1).Text & " " & Format(Now, "hh:mm:ss")
    
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
        BuscaChekc = BuscaChekc & "," & DBLet(miRsAux!PVPInferior, "N") & ","
        
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
Dim Sql As String
Dim CadenaInsertTmpLotes  As String
Dim NumerloLote As String
Dim CuantosLotesDistintos As Integer
Dim CantidadEnTotal As Currency
Dim B As Boolean

    
    ArticuloConFitosantiarios = False
    Sql = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", txtAux(1).Text, "T")
    If Sql = "" Then
        'OK. Salimos YA
        DatosLotesFitosOk = True
        Exit Function
    End If
    B = False
    DatosLotesFitosOk = B
    
    
    
    'Si llega aqui, y no tiene manipulador de Fitosantarios
    If Trim(Text1(45).Text) = "" Then
        
        'Esto sera para el CHOLI , en Navarrres
        Sql = DevuelveDesdeBD(conAri, "ManipuladorNumCarnet", "sclien", "codclien", Text1(4).Text)
        If Sql = "" Then
            'Veo si tiene autirzados
            Sql = DevuelveDesdeBD(conAri, "numcarnet", "sclienmani", "codclien", Text1(4).Text)
        End If
        
        If Sql <> "" Then
            'Llamamos al manipulador de carnet fitosnaitarios
            CadenaDesdeOtroForm = ""
            frmFitoCarnet.Cliente = Val(Text1(4).Text)
            frmFitoCarnet.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                
                Sql = "UPDATE scaalb SET ManipuladorNumCarnet = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "T")
                Sql = Sql & ",ManipuladorFecCaducidad =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F")
                Sql = Sql & ",ManipuladorNombre = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
                Sql = Sql & ", TipoCarnet =" & IIf(UCase(RecuperaValor(CadenaDesdeOtroForm, 4)) = "CUALIFICADO", 2, 1)
                Sql = Sql & ObtenerWhereCP(True)
                If ejecutar(Sql, False) Then
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
        
        
        
        
        
        Sql = "select numlotes,fecentra,Codartic,canentra - vendida"
        Sql = Sql & "  disponible from slotes where "
        Sql = Sql & " codartic=" & DBSet(txtAux(1).Text, "T") & " and canentra - vendida  >0 order by fecentra "
      
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ",0," & DBSet(miRsAux!numlotes, "T") & ")"
                Sql = miRsAux!disponible 'por si solo hay uno
               
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
            If CCur(Sql) < ImporteFormateado(txtAux(3).Text) Then
                MsgBox "Cantidad en el lote insuficiente:" & Sql, vbExclamation
                
            Else
                'Donde va la cantidad asignada en el SQL es en : ,0,'
                'Luego reeplazo esto por la cantidad del albaran
                Sql = TransformaComasPuntos(CStr(ImporteFormateado(txtAux(3).Text)))
                CadenaInsertTmpLotes = Replace(CadenaInsertTmpLotes, ",0,'", "," & Sql & ",'")
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
                    Sql = "select numlotes,fecentra,Codartic,canentra - vendida"
                    Sql = Sql & "  disponible from slotes where "
                    Sql = Sql & " codartic=" & DBSet(txtAux(1).Text, "T") & " and canentra - vendida  >0"
                    Sql = Sql & " AND numlotes= " & DBSet(NumerloLote, "T") & " order by fecentra "
                    Set miRsAux = New ADODB.Recordset
                    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    CadenaInsertTmpLotes = "" 'Vamos a volver a cargar el SQL de insert in lotes
                    CantidadEnTotal = ImporteFormateado(txtAux(3).Text)
                    NumRegElim = 0
                    While Not miRsAux.EOF
                        NumRegElim = NumRegElim + 1
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
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
            conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
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
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
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
Dim Sql As String
Dim CadenaOR As String

        Set miRsAux = New ADODB.Recordset
          
        If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Sub   'Por si acaso se ha metido aqui
        If DBLet(Data2.Recordset!numLote, "T") = "" Then Exit Sub
          
        CadenaInsertTmpLotes = "codtipom ='" & Data1.Recordset!codtipom & "' AND numalbar =" & Data1.Recordset!Numalbar
        CadenaInsertTmpLotes = CadenaInsertTmpLotes & " AND numlinea =" & Data2.Recordset!numlinea
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
            Sql = "select codartic,numlotes,fecentra,canentra-vendida disponible from slotes where "
            Sql = Sql & " codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and canentra-vendida >0  "
            If CadenaOR <> "" Then
                CadenaOR = Mid(CadenaOR, 2)
                Sql = Sql & "  OR numlotes in(" & CadenaOR & ")"
            End If
            Sql = Sql & "  order by fecentra "
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N") & ","
                                
                Sql = "|" & miRsAux!numlotes & "#@#"
                fin = False
                IncioBusqueda = 1
                
                While Not fin
                    
                     
                    J = InStr(IncioBusqueda, LotesArticulos, Sql)
                    If J > 0 Then
                        J = J + Len(Sql)
                        Sql = Mid(LotesArticulos, J, 10)
                        
                        If Sql = Format(miRsAux!fecentra, "dd/mm/yyyy") Then
                            'OK, esta es la linea
                            Sql = Trim(Mid(LotesArticulos, J + 10, 10))
                            Sql = TransformaComasPuntos(Sql)
                            fin = True
                        Else
                            Sql = "|" & miRsAux!numlotes & "#@#"   'Vuelve a la busqueda
                            IncioBusqueda = InStr(J, LotesArticulos, "|")
                        End If
                    Else
                        Sql = "0"
                        fin = True
                    End If
                Wend
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & Sql
                
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
                
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
                Espera 0.3
                CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
                CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
                conn.Execute CadenaInsertTmpLotes
                
                
                
              
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = Data2.Recordset!cantidad
                    frmFacTPVLotes.NombreArticulo = Data2.Recordset!NomArtic
                    frmFacTPVLotes.Show vbModal
              
                    If CadenaDesdeOtroForm = "OK" Then
                    
                        'Primero devolveremos la cantidad que tenia la linea
                        ReestablecerLotesArticulo Data2.Recordset!numlinea
                        
                        'Borramos la linea de lotes
                        Sql = "codtipom ='" & Data1.Recordset!codtipom & "' AND numalbar =" & Data1.Recordset!Numalbar
                        Sql = Sql & " AND numlinea =" & Data2.Recordset!numlinea
                        Sql = "DELETE FROM slialblotes WHERE " & Sql
                        conn.Execute Sql
                        Espera 0.4
                        
                        Sql = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                        Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "'," & Data1.Recordset!Numalbar & "," & Data2.Recordset!numlinea
                        Sql = Sql & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
                        Sql = Sql & " FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 "
            
                        conn.Execute Sql
                        
                        'Tengo que updatear la cantidad vendida
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open "Select * FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not miRsAux.EOF
                            If miRsAux!cantidad <> 0 Then
                                If miRsAux!cantidad > 0 Then
                                    Sql = "+"
                                Else
                                    Sql = "-"
                                End If
                                Sql = "UPDATE slotes SET vendida=vendida " & Sql & DBSet(Abs(miRsAux!cantidad), "N")
                                Sql = Sql & " WHERE numlotes =" & DBSet(miRsAux!numlotes, "T") & " AND codartic= " & DBSet(miRsAux!codArtic, "T")
                                Sql = Sql & " AND fecentra= " & DBSet(miRsAux!FechaAlb, "F")
                            
                                conn.Execute Sql
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
Dim B As Boolean
'Veremos cuanto suman las cantidades de los articulos que llevan loes

    Set miRsAux = New ADODB.Recordset
    
    B = False
    If vParamAplic.NumeroInstalacion = 1 Then B = True
    If vParamAplic.ManipuladorFitosanitarios2 Then B = True
    
    If B Then
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
    Aux = "Select * from slialb WHERE coalesce(comisionagente,0)=0 AND numalbar = " & Data1.Recordset!Numalbar & Aux
        
    
    
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Aux = ""
    While Not miRsAux.EOF
        Aux = Aux & "A:" & miRsAux!codArtic & "    L:" & miRsAux!numlinea & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close


    If Aux <> "" Then
        Set LOG = New cLOG
        Aux = "Albaran con lineas sin comision" & vbCrLf & Aux
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
    Cade = "_" & Data1.Recordset!codtipom & "-" & Format(Data1.Recordset!Numalbar, "0000000")
    Cade = vParamAplic.PathFirmasAlbaran & "\" & Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\*" & Cade & "*.pdf"
    Cade = Dir(Cade, vbArchive)
    If Cade <> "" Then
        Cade = vParamAplic.PathFirmasAlbaran & "\" & Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\" & Cade
        LanzaVisorMimeDocumento Me.hwnd, Cade
    Else
        Cade = Format(Data1.Recordset!FechaAlb, FormatoFecha) & "\*" & Data1.Recordset!codtipom & "-" & Format(Data1.Recordset!Numalbar, "0000000") & " *"
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
Dim Permi As String


    On Error GoTo ePrecioMinimoAlbaran
    
    PrecioMinimoAlbaran = False
    
    Set vPrecioFact = New CPreciosFact
    vPrecioFact.CodigoClien = Text1(4).Text
    vPrecioFact.FijarTarifaActividad
    
    Aux = "slialb.codartic=sartic.codartic and artvario=0 and origpre<>'P' and origpre<>'E'  AND codtipom = '" & Data1.Recordset!codtipom & "' AND numalbar =" & CStr(Data1.Recordset!Numalbar)
    Aux = "Select slialb.* from slialb,sartic WHERE " & Aux
    
    'Agosto2022
    'Las rectificativas NO entran aqui
    If hcoCodTipoM = "ART" Then Aux = Aux & " AND false "
    
    Set RN = New ADODB.Recordset
    ErroresPrMinimo = ""
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF

        If RN!cantidad <> 0 Then
        
            
        
        
        
        
            PrMinimo = Round2(RN!ImporteL / RN!cantidad, 4)
            
            Set vArtic = New CArticulo
            If Not vArtic.LeerDatos(RN!codArtic) Then Err.Raise 513, , "Imposible leer datos articulo: " & RN!codArtic
            'If Not vArtic.EstablecidoPrecioMinimo Then vArtic.FijarprecioMinimo CDate(Text1(1).Text), Val(Text1(4).Text)
            vArtic.FijarprecioMinimo_ CDate(Text1(1).Text), Val(Text1(4).Text)
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
            Permi = ""
        Else
                'Si el pedido esta generado por un usuario NIVELUSU=0
                'dejaremos contiuar
                If IsNull(Data1.Recordset!CodTrab1) Then
                    Permi = -1
                Else
                    Permi = Data1.Recordset!CodTrab1   'EL TRABAJADOR DEL PEDIDO si es super usuario, puedo continuar
                End If
                
                Permi = DevuelveDesdeBD(conAri, "login", "straba", "codtraba", Permi)
                If Permi <> "" Then
                    Permi = UCase(Permi)
                    Permi = DevuelveDesdeBD(conAri, "nivelariges", "usuarios.usuarios", "ucase(login)", Permi, "T")
                    If Permi <> "" Then
                        If Val(Permi) > 0 Then
                            Permi = "N"
                        Else
                            Permi = ""
                        End If
                    End If
                Else
                    Permi = "N" 'No ha encontrado el usuario
                End If
                        
        End If
        
        
        
        If Permi = "" Then
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





Private Function CuantosTiposDeIva() As Boolean
    
    CuantosTiposDeIva = True
    BuscaChekc = Replace(ObtenerWhereCP(False), "scaalb", "slialb")
    BuscaChekc = "Select codigiva from slialb ,sartic where slialb.codartic =sartic.codartic and  " & BuscaChekc & " GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BuscaChekc = ""
    While Not miRsAux.EOF
        BuscaChekc = BuscaChekc & "X"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Len(BuscaChekc) > 3 Then
        MsgBox "Mas de 3 IVAS en factura", vbExclamation
        CuantosTiposDeIva = False
    End If
    BuscaChekc = ""
End Function


'##  Octubre Lo ponemos y quitamos el mismo mes
'Private Sub InsertarCosteHerbelca(linea As String)
'
'
'    If Val(DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", txtAux(1).Text, "T")) <> 1 Then Exit Sub
'
'    'Vamos por posiciones
'    If ModificaLineas = 1 Then
'        CadenaDesdeOtroForm = ""
'    Else
'
'        CadenaDesdeOtroForm = "codtipom =" & DBSet(Text1(30).Text, "T") & " AND numalbar = " & Val(Text1(0).Text) & " AND numlinea "
'        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "precoste", "slialb", CadenaDesdeOtroForm, linea)
'        If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Format(CadenaDesdeOtroForm, FormatoCantidad)
'    End If
'    frmListado5.OtrosDatos = Mid(txtAux(1).Text & Space(16), 1, 16) & Mid(txtAux(2).Text & Space(50), 1, 50) & Mid(txtAux(4).Text & Space(14), 1, 14) & Text2(16).Text
'    frmListado5.OpcionListado = 38
'    frmListado5.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        If ModificaLineas = 1 Then Espera 0.25
'        CadenaDesdeOtroForm = "UPDATE slialb SET precoste = " & DBSet(CadenaDesdeOtroForm, "N")
'        CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE codtipom =" & DBSet(Text1(30).Text, "T") & " AND numalbar= " & Val(Text1(0).Text) & " AND numlinea =" & linea
'        ejecutar CadenaDesdeOtroForm, False
'        CadenaDesdeOtroForm = ""
'    End If
'End Sub



Private Sub BloquearCampoTrabajador()

    imgBuscar(3).Enabled = Modo = 1
    BloquearTxt Text1(3), Modo <> 1
End Sub

