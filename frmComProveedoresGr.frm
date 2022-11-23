VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComProveedoresGr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores "
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   14505
   Icon            =   "frmComProveedoresGr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   79
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   80
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3840
      TabIndex        =   77
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   78
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
      Left            =   12000
      TabIndex        =   76
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   3360
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "lineas"
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
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   120
      TabIndex        =   64
      Top             =   720
      Width           =   14175
      Begin VB.CheckBox chkProveV 
         Caption         =   "Proveedor de Varios"
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
         Left            =   11280
         TabIndex        =   2
         Tag             =   "Proveedor Varios|N|N|||sprove|provario||N|"
         Top             =   248
         Width           =   2655
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
         Index           =   1
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre Proveedor|T|N|||sprove|nomprove||N|"
         Text            =   "Text1"
         Top             =   165
         Width           =   6045
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
         Index           =   0
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Código Proveedor|N|N|0|999999|sprove|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   3120
         TabIndex        =   66
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Proveedor"
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
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   225
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5640
      Top             =   10200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame Frame3 
      Height          =   540
      Left            =   120
      TabIndex        =   59
      Top             =   9960
      Width           =   3000
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
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   2715
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   42
      Top             =   1560
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   14631
      _Version        =   393216
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
      TabPicture(0)   =   "frmComProveedoresGr.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgCuentas(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(20)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgCuentas(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(19)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgCuentas(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "imgFecha(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "imgFecha(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "imgCuentas(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(21)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(62)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "imgCuentas(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(15)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ImgMail(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "imgWeb"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label2(10)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label2(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label2(13)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(16)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(39)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(6)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(3)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(7)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(8)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(9)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(10)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cboTipoDto"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(14)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(15)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(16)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(17)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(18)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(13)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(12)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text2(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text2(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text1(11)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(5)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cboTipoProv"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text2(0)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text1(29)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text2(29)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "checkAlbFac"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "ckhOcultarEnListado"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text1(37)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cboPais"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Frame1"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Frame2(13)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Text1(36)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Text1(40)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Text1(27)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Text1(41)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).ControlCount=   66
      TabCaption(1)   =   "Direcciones/Observaciones"
      TabPicture(1)   =   "frmComProveedoresGr.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(31)"
      Tab(1).Control(1)=   "Text1(32)"
      Tab(1).Control(2)=   "Text1(33)"
      Tab(1).Control(3)=   "Text1(34)"
      Tab(1).Control(4)=   "Text1(35)"
      Tab(1).Control(5)=   "Text1(38)"
      Tab(1).Control(6)=   "Text1(28)"
      Tab(1).Control(7)=   "FrameToolAux(1)"
      Tab(1).Control(8)=   "DataGrid1"
      Tab(1).Control(9)=   "Label2(12)"
      Tab(1).Control(10)=   "imgCuentas(6)"
      Tab(1).Control(11)=   "Label2(14)"
      Tab(1).Control(12)=   "Label2(11)"
      Tab(1).Control(13)=   "imgCuentas(5)"
      Tab(1).Control(14)=   "Label2(1)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Documentos"
      TabPicture(2)   =   "frmComProveedoresGr.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameNavegaDoc"
      Tab(2).Control(1)=   "Text1(30)"
      Tab(2).Control(2)=   "lw1"
      Tab(2).Control(3)=   "LabelDOC"
      Tab(2).Control(4)=   "imgDocumentos"
      Tab(2).Control(5)=   "imgFecha(2)"
      Tab(2).Control(6)=   "Label3"
      Tab(2).ControlCount=   7
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
         Index           =   41
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "R|T|S|||sprove|referencia|||"
         Text            =   "Text1"
         Top             =   3960
         Width           =   3270
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
         Index           =   31
         Left            =   -68160
         MaxLength       =   80
         TabIndex        =   108
         Tag             =   "O|T|S|||sprove|observa1|||"
         Text            =   "Text1"
         Top             =   4680
         Width           =   7080
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
         Index           =   32
         Left            =   -68160
         MaxLength       =   80
         TabIndex        =   107
         Tag             =   "O|T|S|||sprove|observa2|||"
         Text            =   "Text1"
         Top             =   5160
         Width           =   7080
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
         Index           =   33
         Left            =   -68160
         MaxLength       =   80
         TabIndex        =   106
         Tag             =   "O|T|S|||sprove|observa3|||"
         Text            =   "Text1"
         Top             =   5640
         Width           =   7080
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
         Index           =   34
         Left            =   -68160
         MaxLength       =   80
         TabIndex        =   105
         Tag             =   "O|T|S|||sprove|observa4|||"
         Text            =   "Text1"
         Top             =   6120
         Width           =   7080
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
         Index           =   35
         Left            =   -68160
         MaxLength       =   80
         TabIndex        =   104
         Tag             =   "O|T|S|||sprove|observa5|||"
         Text            =   "Text1"
         Top             =   6600
         Width           =   7080
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
         Height          =   1545
         Index           =   38
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   102
         Tag             =   "Observaciones|T|S|||sprove|observaComer|||"
         Text            =   "frmComProveedoresGr.frx":0060
         Top             =   6600
         Width           =   6000
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
         Height          =   1545
         Index           =   28
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Tag             =   "Observaciones|T|S|||sprove|observac|||"
         Text            =   "frmComProveedoresGr.frx":0067
         Top             =   4680
         Width           =   6000
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
         Index           =   27
         Left            =   8760
         MaxLength       =   40
         TabIndex        =   38
         Tag             =   "Web|T|S|||sprove|wwwprove|||"
         Text            =   "Text1"
         Top             =   7680
         Width           =   5160
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
         Index           =   40
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   37
         Tag             =   "eMail Administración|T|S|||sprove|emailPed|||"
         Text            =   "Text1"
         Top             =   7680
         Width           =   5400
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
         Index           =   36
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   36
         Tag             =   "H|T|S|||sprove|horario|||"
         Text            =   "Text1"
         Top             =   7080
         Width           =   12480
      End
      Begin VB.Frame Frame2 
         Caption         =   "Compras"
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
         Height          =   2205
         Index           =   13
         Left            =   7200
         TabIndex        =   92
         Top             =   4500
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
            Index           =   26
            Left            =   4560
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Fax Compras|T|S|||sprove|faxprov2|||"
            Text            =   "Text1"
            Top             =   1680
            Width           =   2040
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
            Index           =   25
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "Teléfono Compras|T|S|||sprove|telprov2|||"
            Text            =   "Text1"
            Top             =   1680
            Width           =   2040
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
            Index           =   24
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   33
            Tag             =   "eMail Compras|T|S|||sprove|maiprov2|||"
            Text            =   "Text1"
            Top             =   1080
            Width           =   5400
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
            Index           =   23
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   32
            Tag             =   "Persona de Contacto Compras|T|S|||sprove|perprov2|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   5400
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Index           =   9
            Left            =   4080
            TabIndex        =   96
            Top             =   1680
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   8
            Left            =   120
            TabIndex        =   95
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
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
            TabIndex        =   94
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   2085
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   840
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   1080
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Administración"
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
         Height          =   2205
         Left            =   240
         TabIndex        =   87
         Top             =   4500
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
            Index           =   19
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   28
            Tag             =   "Persona de Contacto Administración|T|S|||sprove|perprov1|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   5400
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
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   29
            Tag             =   "eMail Administración|T|S|||sprove|maiprov1|||"
            Text            =   "Text1"
            Top             =   1080
            Width           =   5400
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
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Telefono Administración|T|S|||sprove|telprov1|||"
            Text            =   "Text1"
            Top             =   1680
            Width           =   1800
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
            Left            =   4200
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Fax Administración|T|S|||sprove|faxprov1|||"
            Text            =   "Text1"
            Top             =   1680
            Width           =   2400
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   90
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label2 
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
            Index           =   4
            Left            =   120
            TabIndex        =   89
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label2 
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
            Height          =   240
            Index           =   5
            Left            =   3480
            TabIndex        =   88
            Top             =   1680
            Width           =   615
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   840
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
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
            Left            =   120
            TabIndex        =   91
            Top             =   1080
            Width           =   3495
         End
      End
      Begin VB.Frame FrameNavegaDoc 
         Enabled         =   0   'False
         Height          =   735
         Left            =   -74700
         TabIndex        =   83
         Top             =   855
         Width           =   7695
         Begin VB.OptionButton optDoc 
            Caption         =   "Facturas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   5040
            TabIndex        =   86
            Tag             =   "7"
            Top             =   315
            Width           =   2175
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Albaranes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2520
            TabIndex        =   85
            Tag             =   "6"
            Top             =   315
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Pedidos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   84
            Tag             =   "5"
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   1
         Left            =   -74760
         TabIndex        =   81
         Top             =   360
         Width           =   645
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   990
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   150
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   1746
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
      Begin VB.ComboBox cboPais 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2460
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
         Index           =   37
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   10
         Tag             =   "IBAN|T|S|||sprove|iban|||"
         Text            =   "Text1"
         Top             =   2955
         Width           =   735
      End
      Begin VB.CheckBox ckhOcultarEnListado 
         Caption         =   "No listar en Dtos."
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
         Left            =   11040
         TabIndex        =   19
         Tag             =   "s|N|N|||sprove|OcultarEnListDto||N|"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text1 
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
         Index           =   30
         Left            =   -63225
         TabIndex        =   69
         Text            =   "Text4"
         Top             =   1005
         Width           =   1575
      End
      Begin VB.CheckBox checkAlbFac 
         Caption         =   "Albaran x Factura"
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
         Left            =   11040
         TabIndex        =   21
         Tag             =   "s|N|N|||sprove|albaranxfactura||N|"
         Top             =   1440
         Width           =   2775
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
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   67
         Text            =   "Text2"
         Top             =   3960
         Width           =   4005
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
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   16
         Tag             =   "Cod. Situación|N|N|0|99|sprove|codsitua|0|N|"
         Text            =   "Te"
         Top             =   3960
         Width           =   765
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   3450
         Width           =   3375
      End
      Begin VB.ComboBox cboTipoProv 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         TabIndex        =   17
         Tag             =   "Tipo de Proveedor|N|N|||sprove|tipprove||N|"
         Text            =   "Combo1"
         Top             =   480
         Width           =   3735
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Población|T|N|||sprove|pobprove||N|"
         Text            =   "Text1"
         Top             =   1470
         Width           =   2790
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
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
         Left            =   11760
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "Dto. General|N|S|0|99.90|sprove|dtognral|#0.00||"
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
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
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   3450
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
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   2955
         Width           =   4095
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
         Index           =   12
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Cuenta Contable|T|N|||sprove|codmacta|||"
         Text            =   "Text1"
         Top             =   3450
         Width           =   1335
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
         Left            =   9240
         MaxLength       =   3
         TabIndex        =   25
         Tag             =   "Forma Pago|N|N|0|999|sprove|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   2955
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
         Height          =   360
         Index           =   18
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Cuenta Bancaria|T|S|||sprove|cuentaba|0000000000||"
         Text            =   "Text1"
         Top             =   2955
         Width           =   2055
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
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "Digito Control|T|S|||sprove|digcontr|00||"
         Text            =   "Text1"
         Top             =   2955
         Width           =   495
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
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Sucursal|N|S|0|9999|sprove|codsucur|0000||"
         Text            =   "Text1"
         Top             =   2955
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
         Height          =   360
         Index           =   15
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "Banco|N|S|0|9999|sprove|codbanco|0000||"
         Text            =   "Text1"
         Top             =   2955
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
         Index           =   14
         Left            =   9240
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "Banco Propio|N|N|0|9999|sprove|codbanpr|0000||"
         Text            =   "Text1"
         Top             =   3450
         Width           =   735
      End
      Begin VB.ComboBox cboTipoDto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "Tipo Descuento|N|N|||sprove|tipodtos||N|"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   5
         TabIndex        =   23
         Tag             =   "Dto. Pronto Pago|N|S|0|99.90|sprove|dtoppago|#0.00||"
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fecha última compra|F|S|||sprove|fechamov|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Fecha de Alta|F|N|||sprove|fecprove|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "N.I.F.|T|N|||sprove|nifprove|||"
         Text            =   "Text1"
         Top             =   2460
         Width           =   1815
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
         Index           =   2
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Nombre Comercial|T|N|||sprove|nomcomer||N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   4845
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
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   4
         Tag             =   "Domicilio|T|S|||sprove|domprove||N|"
         Text            =   "Text1"
         Top             =   975
         Width           =   4815
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "CPostal|T|N|||sprove|codpobla||N|"
         Text            =   "Text1"
         Top             =   1470
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
         Index           =   6
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Provincia|T|N|||sprove|proprove|||"
         Text            =   "Text1"
         Top             =   1965
         Width           =   3270
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   6375
         Left            =   -74730
         TabIndex        =   71
         Top             =   1695
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   11245
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   72
         Top             =   960
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   39
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   75
         Tag             =   "pais|T|S|||sprove|codpais|||"
         Text            =   "Text1"
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia"
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
         Left            =   7320
         TabIndex        =   111
         Top             =   3960
         Width           =   1125
      End
      Begin VB.Label LabelDOC 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   420
         Left            =   -74100
         TabIndex        =   110
         Top             =   450
         Width           =   5745
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones del pedido"
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
         Left            =   -68160
         TabIndex        =   109
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   6
         Left            =   -72000
         ToolTipText     =   "Buscar forma de pago"
         Top             =   6360
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones comerciales"
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   103
         Top             =   6360
         Width           =   2670
      End
      Begin VB.Label Label2 
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
         Index           =   11
         Left            =   -74760
         TabIndex        =   101
         Top             =   4440
         Width           =   1440
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   -72960
         ToolTipText     =   "Buscar forma de pago"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Horario"
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
         Left            =   360
         TabIndex        =   99
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "eMail"
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
         Index           =   0
         Left            =   360
         TabIndex        =   98
         Top             =   7680
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   10
         Left            =   7800
         TabIndex        =   97
         Top             =   7800
         Width           =   495
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   8400
         Picture         =   "frmComProveedoresGr.frx":006E
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   7800
         Width           =   255
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   2
         Left            =   960
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   7680
         Width           =   240
      End
      Begin VB.Image imgDocumentos 
         Height          =   375
         Left            =   -74670
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Pais"
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
         Left            =   3960
         TabIndex        =   74
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Direcciones del proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   1
         Left            =   -74040
         TabIndex        =   73
         Top             =   480
         Width           =   5145
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   -63585
         Picture         =   "frmComProveedoresGr.frx":05F8
         ToolTipText     =   "Buscar fecha"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Left            =   -64185
         TabIndex        =   70
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   1800
         ToolTipText     =   "Buscar situación"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Situación"
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
         Index           =   62
         Left            =   240
         TabIndex        =   68
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN Proveedor"
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
         Index           =   21
         Left            =   255
         TabIndex        =   63
         Top             =   2990
         Width           =   1320
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   1680
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   8880
         Picture         =   "frmComProveedoresGr.frx":0683
         ToolTipText     =   "Buscar fecha"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   8880
         Picture         =   "frmComProveedoresGr.frx":070E
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1800
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Proveedor"
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
         Left            =   7320
         TabIndex        =   62
         Top             =   510
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. General"
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
         Left            =   10440
         TabIndex        =   61
         Top             =   2474
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Propio"
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
         Index           =   14
         Left            =   7320
         TabIndex        =   58
         Top             =   3480
         Width           =   1275
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   8880
         ToolTipText     =   "Buscar banco propio"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. Pronto Pago"
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
         Left            =   7320
         TabIndex        =   57
         Top             =   2474
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Descuento"
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
         Left            =   7320
         TabIndex        =   56
         Top             =   1983
         Width           =   1545
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   8880
         ToolTipText     =   "Buscar forma de pago"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
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
         Left            =   7320
         TabIndex        =   55
         Top             =   3000
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Cta Contable"
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
         Left            =   255
         TabIndex        =   54
         Top             =   3480
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Ult. Compra"
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
         Left            =   7320
         TabIndex        =   53
         Top             =   1492
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
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
         Index           =   8
         Left            =   7320
         TabIndex        =   52
         Top             =   1001
         Width           =   1380
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
         Index           =   7
         Left            =   255
         TabIndex        =   51
         Top             =   2494
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Comercial"
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
         Index           =   2
         Left            =   255
         TabIndex        =   50
         Top             =   510
         Width           =   1755
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
         Index           =   5
         Left            =   3000
         TabIndex        =   49
         Top             =   1560
         Width           =   1200
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
         Index           =   3
         Left            =   255
         TabIndex        =   47
         Top             =   1006
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. postal"
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
         Index           =   4
         Left            =   255
         TabIndex        =   45
         Top             =   1502
         Width           =   1125
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
         Index           =   6
         Left            =   255
         TabIndex        =   43
         Top             =   1998
         Width           =   885
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
      Height          =   390
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   10080
      Width           =   1125
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
      Height          =   390
      Left            =   11760
      TabIndex        =   39
      Top             =   10080
      Width           =   1125
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
      Height          =   390
      Left            =   13080
      TabIndex        =   41
      Top             =   10080
      Visible         =   0   'False
      Width           =   1125
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmComProveedoresGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private HaDevueltoDatos As Boolean
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
Private Ordenacion As String
Private CadenaConsulta As String

Private WithEvents frmB2 As frmBasico2 'Form para busquedas
Attribute frmB2.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Formas de Pago en menu Facturacion
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmBP As frmBasico2 'frmFacBancosPropios
Attribute frmBP.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1



Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim Modo As Byte


Private Sub cboPais_KeyPress(KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub

Private Sub cboTipoDto_KeyPress(KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub cboTipoProv_KeyPress(KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub checkAlbFac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkProveV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ckhOcultarEnListado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim CambioNombreProveedor As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 1 'BUSCAR
        HacerBusqueda
   
    Case 2, 4 'MODIFICAR
        If DatosOk Then
            If Data1.Recordset.EOF Then
                i = InsertarDesdeForm(Me)
            Else
                CambioNombreProveedor = False
                                                                            'EL NOMBRE DEL PROVEEDOR HA CAMBIADO
                If Trim(DevNombreSQL(Data1.Recordset!nomprove)) <> Trim(Text1(1).Text) Then CambioNombreProveedor = True
                
                
                i = ModificaDesdeFormulario(Me, 1)
                TerminaBloquear
                
                'Actualizadmos en contabilidad
                
    '                                                                'Hay datos contables. Actualizamos?
                If HayQueActualizarenContabilidad Then
                    ModificarCtaContabilidad False, Text1(12).Text, Val(Text1(0).Text)
                    Text2(0).Text = Text1(1).Text
                End If
                    
                
                If CambioNombreProveedor Then UpdatearNomProve
                
                
                PosicionarData
            End If
        End If

    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                ComprobarCrearCuenta
                PosicionarData
            End If
        End If
    End Select
    


Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ComprobarCrearCuenta()
        'Si pone en la cuenta contable, crear nueva cta contable
         If Text2(0).Text = vbCrearNuevaCta Then
            If Not InsertarCuentaCble(Text1(12).Text, "", Text1(0).Text) Then
                MsgBox "Se ha producido un error insertando la cuenta: " & Text1(0).Text & ". Compruebelo", vbExclamation
                Exit Sub
            Else
                Text1_LostFocus 12
            End If
        End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
        Case 2
            PonerCampos
            lblIndicador.Caption = ""
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
    
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    VariePublic = Text1(0).Text
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    VerLineasDirecciones True
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_Load()
    'Icono del formularios
    Me.Icon = frmPpal.Icon
    Modo = 0
   
    'Icono de busqueda
    For kCampo = 0 To Me.imgCuentas.Count - 1
        'Me.imgCuentas(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Me.imgCuentas(kCampo).Picture = frmPpal.imgListComun.ListImages(1).Picture
    Next kCampo
   
   'Icono de e-mail
    For kCampo = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(kCampo).Picture = frmPpal.imgListComun.ListImages(20).Picture
    Next kCampo
   
   
    
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
       
   With Me.ToolbarAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        
        '.Buttons(1).Image = 3
        .Buttons(1).visible = False
        .Buttons(2).Image = 4
        '.Buttons(3).Image = 5
        .Buttons(3).visible = False
    End With
     
    
    
    'Documentos
    ImagenesNavegacion
    
    'Solo si puede tener REA, entonces mostraremos el check este
    checkAlbFac.visible = vParamAplic.IVA_REA > 0
    
    
    limpiar Me
    Me.SSTab1.Tab = 0
    VieneDeBuscar = False

    NombreTabla = "sprove"
    Ordenacion = " ORDER BY codprove"
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codprove=-1"
    Data1.Refresh
    CargaGrid False
    Toolbar1.Buttons(6).Enabled = Not Data1.Recordset.EOF
    Toolbar1.Buttons(7).Enabled = Not Data1.Recordset.EOF
     
     
     
    Text1(39).visible = vParamAplic.ContabilidadNueva
    Label1(15).visible = vParamAplic.ContabilidadNueva
    cboPais.visible = vParamAplic.ContabilidadNueva
     
     
    CargarComboTipoProveedor
    CargarComboTipoDto
    CargaComboPais
    
    optDoc_Click 0
      
    'Ponemos los datos del listview
    imgFecha(2).Tag = vEmpresa.FechaIni
    CargaColumnas 1   'Por defecto albaranes
    ImagenDocumento CByte(optDoc(1).Tag)
      
    If vParamAplic.ManipuladorFitosanitarios2 Then Label1(16).Caption = "Nº ROPO"
      
      
     '=======Modif.
     If DatosADevolverBusqueda = "" Then
        PonerModo 0
     Else
        PonerModo 1
     End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Byte
      
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un botón de busqueda de los campos
            'Cuenta Contable, Forma Pago, Banco Propio
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 12).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(Indice).Text = RecuperaValor(CadenaDevuelta, 2)

        Else
            'Recupera todo el registro de Proveedor
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
    
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub




Private Sub frmB2_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Byte

    If CadenaSeleccion <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un botón de busqueda de los campos
            'Cuenta Contable, Forma Pago, Banco Propio
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 12).Text = RecuperaValor(CadenaSeleccion, 1)
            Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)

        Else
            'Recupera todo el registro de Proveedor
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
            cadB = Aux
    
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If

End Sub

Private Sub frmBP_DatoSeleccionado(CadenaSeleccion As String)
'Banco Propio
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(14)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    
    Indice = CByte(Val(imgFecha(0).Tag))
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
    
    If Indice = 30 Then imgFecha(2).Tag = vFecha
    
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Forma de Pago
    Text1(13).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(29).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(29).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim Indice As Byte
    
    If Index <> 5 And Index <> 6 Then
        If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cuenta Contable
            imgCuentas(0).Tag = Index
            'Conexión a BD: Conta, Tabla: Cuentas
            MandaBusquedaPrevia "apudirec='S'"
            imgCuentas(0).Tag = -1 'Abre el frmBuscaGrid para la conexión
                                   'de la BD: Ariges
            Indice = 12
        Case 1 'Forma de Pago
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Indice = 13
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
        Case 2 'Banco Propio
'            Set frmBP = New frmFacBancosPropios
'            frmBP.DatosADevolverBusqueda = "0"
'            frmBP.Show vbModal
'            Set frmBP = Nothing
            Indice = 14
            Set frmBP = New frmBasico2
            AyudaBancosPropios frmBP, Text1(Indice)
            Set frmBP = Nothing
            
        Case 3 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            VieneDeBuscar = True
            Indice = 4
        Case 4
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
            frmS.Show vbModal
            Set frmS = Nothing
            
            
        Case 5, 6
            If Modo = 5 Or Modo = 0 Then
               Else
                
                    If Modo = 3 Or Modo = 4 Then
                        If Index = 5 Then
                            CadenaDesdeOtroForm = Text1(28).Text
                        Else
                            CadenaDesdeOtroForm = Text1(38).Text
                        End If
                    Else
                        CadenaDesdeOtroForm = ""
                        If Not Data1.Recordset.EOF Then
                            If Index = 5 Then
                                CadenaDesdeOtroForm = DBLet(Data1.Recordset!observac, "T")
                            Else
                                CadenaDesdeOtroForm = DBLet(Data1.Recordset!observacomer, "T")
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
                            If Index = 5 Then
                                Text1(28).Text = Mid(CadenaDesdeOtroForm, 3)
                            Else
                                Text1(38).Text = Mid(CadenaDesdeOtroForm, 3)
                            End If
                        End If
                    End If
                    CadenaDesdeOtroForm = ""
            End If
        
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Integer

   If Modo = 2 Or Modo = 5 Or Modo = 0 Then
        If Index <> 2 Then Exit Sub
    End If
   Screen.MousePointer = vbHourglass
   
   
   If Index < 2 Then
        Indice = 8 + Index
   Else
        'text1)30)
        Indice = 30
   End If
   imgFecha(0).Tag = Indice
   'FECHA
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   frmF.Show vbModal
   Set frmF = Nothing
   If Index <> 2 Then
        PonerFoco Text1(Indice)
    Else
        CargaDatosLW
    End If
End Sub


Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        dirMail = Text1(20).Text
    ElseIf Index = 1 Then
        dirMail = Text1(24).Text
        
    ElseIf Index = 2 Then
        dirMail = Text1(40).Text
    End If
    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(27).Text) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub lw1_DblClick()
Dim Seleccionado As Long
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un proveedor. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 1
        'ALBARANES
        'Set frmAlb = New frmFacEntAlbaranes
        If vParamAplic.TipoFormularioClientes = 0 Then
            
            frmComEntAlbaranesGR.cadSelAlbaranes = " numalbar='" & DevNombreSQL(lw1.SelectedItem.Text) & _
            "' AND fechaalb= '" & Format(lw1.SelectedItem.SubItems(1), "yyyy-mm-dd") & _
            "' AND codprove = " & Data1.Recordset!Codprove
        
            frmComEntAlbaranesGR.Show vbModal
            frmComEntAlbaranesGR.cadSelAlbaranes = ""
        Else
            frmComEntAlbaranSA.cadSelAlbaranes = " numalbar='" & DevNombreSQL(lw1.SelectedItem.Text) & _
            "' AND fechaalb= '" & Format(lw1.SelectedItem.SubItems(1), "yyyy-mm-dd") & _
            "' AND codprove = " & Data1.Recordset!Codprove
        
            frmComEntAlbaranSA.Show vbModal
            frmComEntAlbaranSA.cadSelAlbaranes = ""
        End If
    
    Case 0
        'PEDIDOS
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComEntPedidos2.MostrarDatos = lw1.SelectedItem.Text
            frmComEntPedidos2.EsHistorico = False
            frmComEntPedidos2.Show vbModal
        Else
            'SAIL
            
        End If
    Case 2
        'FACTURAS
        'Este no necesitamos crear instancias
        AbrirFacturaLW
        
        
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If

End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
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

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub optDoc_Click(Index As Integer)
 Dim ElTag As Byte
    If Modo = 0 Then Exit Sub
    ElTag = CByte(optDoc(Index).Tag)
    ImagenDocumento ElTag
    lw1.ListItems.Clear
    CargaColumnas CByte(Index)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
     kCampo = Index
     ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
On Error Resume Next

    'Si no estamos Insertando o Modificando no hacemos nada
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'proveedor
            PonerFormatoEntero Text1(Index)
            
        Case 1
            If Modo = 3 Then
                If Me.Text1(Index).Text <> "" Then Text1(2).Text = Text1(Index).Text
            End If
                
            
        Case 4 'Cod. Postal
            If Text1(Index).Locked Then Exit Sub
            If Text1(Index).Text <> "" And Not VieneDeBuscar Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            Else
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
            End If
            VieneDeBuscar = False
            
        Case 7 'NIF
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                If ValidarNIF_(Text1(Index).Text, False) Then
                     'select codprove, nomprove ,nifprove  from sprove
                     devuelve = DevuelveDesdeBD(conAri, "concat(codprove,' - ',nomprove)", "sprove", "nifprove", Text1(Index).Text, "T")
                     If devuelve <> "" Then MsgBox "Ya existe un proveedor con este NIF" & vbCrLf & devuelve, vbExclamation
                     devuelve = ""
                End If
            End If
       
        Case 8, 9 'Fechas
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 10, 11 'Descuentos
            'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
        Case 12 'Cta Contable
            Text2(0).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            
        Case 13 ' Forma Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(1).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa")
            Else
                Text2(1).Text = ""
            End If
            
        Case 14 'Banco Propio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(2).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr")
            Else
                Text2(2).Text = ""
            End If
            
        Case 15, 16, 17, 18 'cuenta bancaria
            PonerFormatoEntero Text1(Index)
            If Index = 18 Then
                devuelve = Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text
                If Len(devuelve) = 20 Then
                    DevuelveIBAN2 "ES", devuelve, devuelve
                    If Len(devuelve) = 2 Then
                        devuelve = "ES" & devuelve
                        If Me.Text1(37).Text = "" Then
                            Text1(37).Text = devuelve
                        Else
                            If Me.Text1(37).Text <> devuelve Then MsgBox "Codigo IBAN distinto del calculado [" & devuelve & "]", vbExclamation
                        End If
                    End If
                    
                End If
            End If
        Case 27
            PonerFocoBtn Me.cmdAceptar
            
        Case 29
            If PonerFormatoEntero(Text1(29)) Then
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "ssitua", "nomsitua", "codsitua")
            Else
                Text2(29).Text = ""
            End If
        Case 37
            If Me.Text1(Index).Text <> "" Then IBAN_Correcto Me.Text1(Index)
    End Select
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Poner botones de desplazamiento visible si Modo 2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If




    
    '----------------------------------------------------------------
    B = (Kmodo >= 3) Or Modo = 1 'Modo: Insertar/Modificar o Busqueda
    Me.cboTipoProv.Enabled = B
    Me.cboTipoDto.Enabled = B
    Me.chkProveV.Enabled = B 'proveedor varios
    checkAlbFac.Enabled = B           'Solo si al aplicacion lleva REA veremos este check
    Me.ckhOcultarEnListado.Enabled = B
    cmdAceptar.visible = B
    cmdCancelar.visible = B
    If vParamAplic.ContabilidadNueva Then cboPais.Enabled = B
    
    FrameNavegaDoc.Enabled = Modo = 2 Or Modo = 0
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Fecha ult. compra bloqueada pq se modifica por programa
    BloquearTxt Text1(9), (Modo <> 1)
    'La fecha esta NUNCA se puede escribir
    Text1(30).Enabled = False
    Text1(30).Text = Me.imgFecha(2).Tag
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    
    B = False
    Me.ToolbarAux(0).Buttons(2).Enabled = False
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            Me.ToolbarAux(0).Buttons(2).Enabled = True
            If Data1.Recordset.RecordCount > 1 Then B = True
        End If
    End If
    DespalzamientoVisible B
    
    
    'Para remarcar el cliente
    '&H00C0FFFF&
    If Modo = 2 Then Text1(1).BackColor = &HC0FFFF
        
    
    'Poner longitud de los campos
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    
    B = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
      
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    mnEliminar.Enabled = B

    

    B = (Modo >= 3) 'Modo: Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerCampos()
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Text2(0).Text = PonerNombreCuenta(Text1(12), Modo)

    'Rellenar Text2 con nombre asociado al codigo
    Text2(1).Text = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(13).Text, "N")
    Text2(2).Text = DevuelveDesdeBDNew(conAri, "sbanpr", "nombanpr", "codbanpr", Text1(14).Text, "N")
        
        
    If vParamAplic.ContabilidadNueva Then PonerPais
        
    'Poner la situacion
    Modo = 3   'pequeña trampa para que haga el losfocus
    Text1_LostFocus 29
    Modo = 2
    
    CargaGrid True
    
    Me.Refresh
    DoEvents
    CargaDatosLW
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim X As String
        
    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
        
        
        
    'Campos nombre direccion... NO pueden tener *
    If ComprobarTieneAsteriscosEnTextbox("1|2|") Then
        If Modo = 3 Then
            B = False
            Exit Function
        End If
    End If
    
        
        
    'Validar que la cuenta bancaria es correcta
    
    If Comprueba_CuentaBan2(Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text, False) Then
        CadenaConsulta = Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text
        If Len(CadenaConsulta) = 20 Then
            
            X = ""  'BuscaChekc
            If Me.Text1(37).Text <> "" Then X = Mid(Text1(37).Text, 1, 2)
            
                
            If DevuelveIBAN2(X, CadenaConsulta, CadenaConsulta) Then
                If Me.Text1(37).Text = "" Then
                    If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(37).Text = X & CadenaConsulta
                Else
                    If Mid(Text1(37).Text, 3) <> CadenaConsulta Then
                        CadenaConsulta = "Calculado : " & X & CadenaConsulta
                        CadenaConsulta = "Introducido: " & Me.Text1(37).Text & vbCrLf & CadenaConsulta & vbCrLf
                        CadenaConsulta = "Error en codigo IBAN" & vbCrLf & CadenaConsulta & "Continuar?"
                        If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then Exit Function
                    End If
                End If
            End If
                    
        End If
        CadenaConsulta = ""

    End If
    
    
    
    
    If Modo = 4 Then
        'Modificar
        If Text1(12).Text <> DBLet(Data1.Recordset!Codmacta, "N") Then
            If MsgBox("Va a cambiar la cuenta en contabilidad. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then B = False
        End If
        
    ElseIf Modo = 3 Then
        X = DevuelveDesdeBD(conAri, "concat(codprove,' - ',nomprove)", "sprove", "nifprove", Text1(7).Text, "T")
        If X <> "" Then
            X = "Ya existe un proveedor con este NIF:" & vbCrLf & Text1(7).Text & vbCrLf & X & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(X, vbQuestion + vbYesNo) = vbNo Then B = False
        End If
        
    End If
    
    If B And vParamAplic.ContabilidadNueva Then Me.Text1(39).Text = PaisSeleccionado
    
    DatosOk = B
End Function


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        
        Case 1  'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
           If BLOQUEADesdeFormulario(Me) Then BotonModificar
        Case 3  'Borrar
           mnEliminar_Click
           
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
        
        Case 8
            'IMPRMIR
            'AbrirListado (58)   ': Informe Proveedores
            frmInformesNew.OpcionListado = 58
            frmInformesNew.Show vbModal

    End Select


End Sub

Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3 'Modo 3: Insertar
    SSTab1.Tab = 0
    
    'Obtenemos la siguiente numero de codigo de Proveedor
    If vParamAplic.NumeroInstalacion = vbFenollar Then
    
        If MsgBox("Acreedor?", vbQuestion + vbYesNo) = vbYes Then
            Text1(0).Text = SugerirCodigoSiguienteStr("sprove", "codprove", "codprove>9999")
        Else
            Text1(0).Text = SugerirCodigoSiguienteStr("sprove", "codprove", "codprove<9999")
        End If
    Else
        Text1(0).Text = SugerirCodigoSiguienteStr("sprove", "codprove")
    End If
    Text1(8).Text = Format(Now, "dd/mm/yyyy")
    Me.cboTipoProv.ListIndex = 0
    
    Text1(29).Text = vParamAplic.PorDefecto_Situ
    Text1_LostFocus 29
    
    If vParamAplic.ContabilidadNueva Then cboPais.ListIndex = 0
    
    PonerFoco Text1(0)   'Ponemos el foco
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    If PuedeEliminarProveedor Then
    
    
    
        '### a mano
        Cad = "¿Seguro que desea eliminar el Proveedor?"
        Cad = Cad & vbCrLf & "Cod. : " & Data1.Recordset.Fields(0)
        Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
        'Borramos
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            On Error GoTo Error2
            NumRegElim = Data1.Recordset.AbsolutePosition
            conn.BeginTrans
            Cad = "DELETE FROM sdirrecog where codprove=" & Data1.Recordset!Codprove
            If ejecutar(Cad, False) Then
                Cad = "DELETE FROM sprove where codprove=" & Data1.Recordset!Codprove
                If ejecutar(Cad, False) Then Cad = ""
            End If
            If Cad = "" Then
                'OK. Ha ido bien
                conn.CommitTrans
            
                If SituarDataTrasEliminar(Data1, NumRegElim) Then
                    PonerCampos
                Else
                    LimpiarCampos
                    PonerModo 0
                End If
            
            Else
                'Error es
                conn.RollbackTrans
            End If
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            Data1.Recordset.CancelUpdate
            MuestraError Err.Number, "Eliminar Proveedor", Err.Description
        End If
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Me.Text1(2)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            PonerFoco Text1(kCampo)
            Text1(kCampo).BackColor = vbYellow
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub CargarComboTipoProveedor()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Nacional, 1-Intracomunitario, 2-Extranjero  3- Regimen especial agrario
'4- Estimación directa

    cboTipoProv.Clear
    cboTipoProv.AddItem "Nacional"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 0
    
    cboTipoProv.AddItem "Intracomunitario"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 1
    
    cboTipoProv.AddItem "Extranjero"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 2
    
    cboTipoProv.AddItem "R.E.A."
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 3
    
    cboTipoProv.AddItem "Estimación directa"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 4
End Sub

Private Sub CargarComboTipoDto()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-Resto

    cboTipoDto.Clear
    cboTipoDto.AddItem "Aditivo"
    cboTipoDto.ItemData(cboTipoDto.NewIndex) = 0
    
    cboTipoDto.AddItem "Resto"
    cboTipoDto.ItemData(cboTipoDto.NewIndex) = 1
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    cboTipoDto.ListIndex = -1
    cboTipoProv.ListIndex = -1
    Me.chkProveV.Value = 0
    Me.checkAlbFac.Value = 0
    Me.ckhOcultarEnListado.Value = 0
    If vParamAplic.ContabilidadNueva Then cboPais.ListIndex = -1
    CargaGrid False
End Sub


Private Sub HacerBusqueda()
Dim cadB As String


    If vParamAplic.ContabilidadNueva Then Text1(39).Text = PaisSeleccionado


    cadB = ObtenerBusqueda(Me, False)
    
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Conexion As String

    'Llamamos a al form
    Cad = ""
    Select Case Val(Me.imgCuentas(0).Tag)
        Case 0 'Se llama a Busqueda desde el campo: Cuenta Contable
            '#A MANO: Porque busca en la Tabla: Cuentas
            'de la BDatos de Contabilidad
            Cad = Cad & "Código|cuentas|codmacta|T||30·Denominacion|cuentas|nommacta|T||70·"
            tabla = "cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexión a BD: Conta
            
            Set frmB2 = New frmBasico2
            
            AyudaCtasContables frmB2, Text1(12)
            
            Set frmB2 = Nothing
            
            
        Case Else 'Se llama a Busqueda desde el registro Proveedor
            Cad = Cad & ParaGrid(Text1(0), 20, "Código")
            Cad = Cad & ParaGrid(Text1(1), 40, "Nombre")
            Cad = Cad & ParaGrid(Text1(2), 41, "Nombre Comercial")
            tabla = "sprove"
            Titulo = "Proveedores"
            Conexion = conAri    'Conexión a BD: Ariges
    
            Set frmB2 = New frmBasico2
            
            AyudaProveedores frmB2, , cadB, True
            
            Set frmB2 = Nothing
    
    
    End Select
        
    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = Conexion
'        frmB.vCargaFrame = (Conexion = 2)
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            PonerFoco Text1(kCampo + 1)
''                If (Not adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                    cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If


    End If
End Sub


Private Sub PonerCadenaBusqueda()
    On Error GoTo EEPonerBusq

    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        If cmdRegresar.visible Then
            PonerFocoBtn Me.cmdRegresar
        Else
            PonerFoco Text1(0)
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub



Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codprove=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
'       LimpiarCampos
        PonerModo 0
    End If
End Sub





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
'        .Buttons(3).Image = 9
'        .Buttons(5).Image = 10
'        .Buttons(7).Image = 11
'    End With
'
    Set lw1.SmallIcons = frmPpal.ImgListPpal
End Sub







Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 1
        LabelDoc.Caption = "Albaranes"

        'ALBARANES
        'Label2(0).Caption = "Albaranes"
        Columnas = "Albaran|Fecha|Cod.|Forma de pago|Dto.gral|Dto.pp|Importe|"
        Ancho = "2200|1600|1300|3300|1100|1100|2000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|1|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|000||" & FormatoImporte & "|" & FormatoImporte & "|" & FormatoImporte & "|"
        Ncol = 7
               
               
    Case 2
        LabelDoc.Caption = "Facturas"
        
        'Label2(0).Caption = "Facturas"
        Columnas = "Numero|Fecha|F. recepcion|Codigo|Forma de pago|Dto.gral|Dto.pp|Importe|"
        Ancho = "2000|1500|1500|800|3000|1000|900|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|1|1|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|dd/mm/yyyy|000||" & FormatoImporte & "|" & FormatoImporte & "|" & FormatoImporte & "|"
        Ncol = 8
               
    Case 0
        'PEDIDOS
        LabelDoc.Caption = "Pedidos"

        'Label2(0).Caption = "Pedidos"
        Columnas = "Visado"
        
        Columnas = "Numero|Fecha|Codigo|Cliente|Trab.|Solicitado por|Importe|"
        Ancho = "1500|1500|1200|3200|900|2800|1800|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|0|1|"
        'Formatos
        Formato = "00000000|dd/mm/yyyy|000000||000||" & FormatoImporte & "|"
        Ncol = 7
    'Case 2
        '
    End Select
    
    
    'Fecha incio busquedas
    Text1(30).Text = Format(imgFecha(2).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte

        If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
            If vUsu.CodigoAgente > 0 Then Exit Sub
        End If


    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo  doc."
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub



Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim BuscaChekc
    On Error GoTo ECargaDatosLW
    
    
    If Modo <> 2 Then Exit Sub
    
    ElIcono = 0
    For NumRegElim = 0 To Me.optDoc.Count - 1
        If Me.optDoc(NumRegElim).Value Then
            ElIcono = Me.optDoc(NumRegElim).Tag
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text1(30).Text = Format(imgFecha(2).Tag, "dd/mm/yyyy")
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 1
        'ALBARANES
        Cad = " select c.numalbar,c.fechaalb,c.codforpa,nomforpa,dtognral,dtoppago,Sum(ImporteL)"
        Cad = Cad & " from scaalp c left join sforpa on c.codforpa=sforpa.codforpa"
        Cad = Cad & " inner join slialp l on c.codprove=l.codprove and c.numalbar=l.numalbar"
        
        
        GroupBy = "1,2,3"
        BuscaChekc = "c.fechaalb"
        
    
    Case 0
        'PEDIDOS,

        Cad = "select c.numpedpr,c.fecpedpr,c.codclien,nomclien,codtrab1,nomtraba,sum(importel) from scappr c "
        Cad = Cad & " left join straba on codtrab1=straba.codtraba"
        Cad = Cad & " left join sclien on c.codclien=sclien.codclien"
        Cad = Cad & " inner join slippr on c.numpedpr=slippr.numpedpr  "

         Cad = Cad & " WHERE true "
        BuscaChekc = "fecpedpr"
        GroupBy = "1"
    Case 2
        Cad = "select numfactu,fecfactu,fecrecep,c.codforpa,nomforpa,dtognral,dtoppago,totalfac from scafpc c "
        Cad = Cad & " left join sforpa on c.codforpa=sforpa.codforpa WHERE true "
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    Cad = Cad & " and c.codprove=" & Data1.Recordset!Codprove
    
    'La fecha
    Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFecha(2).Tag, FormatoFecha) & "'"
    
    
    'El group by
    Cad = Cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            IT.Text = Format(Rs.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            IT.Text = Rs.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(Rs.Fields(NumRegElim - 1)) Then
                IT.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    IT.SubItems(NumRegElim - 1) = Format(Rs.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    IT.SubItems(NumRegElim - 1) = Rs.Fields(NumRegElim - 1)
                End If
            End If
        Next
        IT.SmallIcon = ElIcono
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub



Private Sub AbrirFacturaLW()
Dim s As String
    
    Set miRsAux = New ADODB.Recordset

    s = "select numalbar,fechaalb from scafpa where numfactu='" & DevNombreSQL(lw1.SelectedItem.Text)
    s = s & "' and fecfactu='" & Format(lw1.SelectedItem.SubItems(1), FormatoFecha) & "' ORDER BY numalbar desc"
    miRsAux.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    s = ""
    If Not miRsAux.EOF Then
        s = DevNombreSQL(miRsAux.Fields(0)) & "|" & miRsAux.Fields(1) & "|"
    End If
    miRsAux.Close
    Set miRsAux = Nothing

    
    If s <> "" Then
        If vParamAplic.TipoFormularioClientes = 0 Then
            With frmComHcoFacturas2GR
                .hcoCodMovim = RecuperaValor(s, 1)
                .hcoFechaMovim = RecuperaValor(s, 2)
                .hcoCodProve = Data1.Recordset!Codprove
                .Show vbModal
            End With
        Else
            'SAIL
            
        End If
    
    Else
        MsgBox "No se han encontrado los albaranes de la factura", vbExclamation
    End If
End Sub




'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'
'   Grid de direcciones recogida
'
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim Sql As String
    
    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    Sql = "select  `coddirre`,`nomdirre`,`codpobla`,`pobdirre`,`teldirre` from sdirrecog  WHERE codprove = "
    If enlaza Then
        Sql = Sql & Data1.Recordset!Codprove
    Else
        Sql = Sql & " -1"
    End If
    
    
    CargaGridGnral DataGrid1, Data2, Sql, True
    
    CargaGrid2 DataGrid1, Data2
    
'    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
'    vDataGrid.Enabled = Not B
    DataGrid1.ScrollBars = dbgAutomatic

    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
    
    On Error GoTo ECargaGrid

   

            i = 0
            vDataGrid.Columns(i).Caption = "Cod."
            vDataGrid.Columns(i).Width = 970
            vDataGrid.Columns(i).NumberFormat = "000"
            
            i = i + 1 '4
            vDataGrid.Columns(i).Caption = "Descripcion"
            vDataGrid.Columns(i).Width = 4200
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "C.P."
            vDataGrid.Columns(i).Width = 1100

            i = i + 1
            vDataGrid.Columns(i).Caption = "Poblacion"
            vDataGrid.Columns(i).Width = 3200
            

            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Telefono"
            vDataGrid.Columns(i).Width = 2900
            
            
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i

    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub



Private Sub VerLineasDirecciones(DesdeDobleClick As Boolean)

    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If DesdeDobleClick Then
            If Data2.Recordset.EOF Then Exit Sub
    End If
    
    If BloqueoManual("lindirprov", Data1.Recordset!Codprove, False) Then

        frmComDirRecogida.Codprove = Data1.Recordset!Codprove
        frmComDirRecogida.nomprove = CStr(Data1.Recordset!nomprove)
        
        If DesdeDobleClick Then frmComDirRecogida.VerDatoDpto = Data2.Recordset!coddirre
   
        frmComDirRecogida.Show vbModal
        CargaGrid True
        
    End If
    DesBloqueoManual "lindirprov"
End Sub


Private Function PuedeEliminarProveedor() As Boolean
Dim C As String

    PuedeEliminarProveedor = False
    CadenaConsulta = "Fra. proveedor"
    C = DevuelveDesdeBD(conAri, "numfactu", "scafpc", "codprove", CStr(Data1.Recordset!Codprove))
    If C = "" Then
        CadenaConsulta = "Alb. proveedor"
        C = DevuelveDesdeBD(conAri, "numalbar", "scaalp", "codprove", CStr(Data1.Recordset!Codprove))
        If C = "" Then
            CadenaConsulta = "Pedido proveedor"
            C = DevuelveDesdeBD(conAri, "numpedpr", "scappr", "codprove", CStr(Data1.Recordset!Codprove))
            If C = "" Then
                CadenaConsulta = "Articulos asignados"
                C = DevuelveDesdeBD(conAri, "numpedpr", "scappr", "codprove", CStr(Data1.Recordset!Codprove))
                If C = "" Then
                    'FLOTAS
                    CadenaConsulta = "Flotas."
                    C = DevuelveDesdeBD(conAri, "concat(codflota ,' ',nomflota)", "sflotas", "codprove", CStr(Data1.Recordset!Codprove))
                    If C = "" Then PuedeEliminarProveedor = True
                End If
            End If
        End If
    End If
    
    If C <> "" Then
        C = "Existen datos relacionados con el proveedor: " & CadenaConsulta
        MsgBox C, vbExclamation
    End If
    CadenaConsulta = ""
        
End Function





'Comprobaremos que ha cambiado los campos que enlazan con conta. nombre nif.....
Private Function HayQueActualizarenContabilidad() As Boolean
Dim QueCampos As String
Dim mTag As cTag
Dim i As Integer
Dim fin As Boolean
Dim txt As String
Dim Valor
    HayQueActualizarenContabilidad = False
    If Text1(12).Text = "" Or Text2(0).Text = "" Then Exit Function

    'Vere si el campo que habia al que hay ha cambiado
    QueCampos = "0|1|3|4|5|6|7|12|13|15|16|17|18|37|"
    fin = False
    Set mTag = New cTag
    
    While Not fin
      i = InStr(1, QueCampos, "|")
      'NO puede ser ccero
      txt = Mid(QueCampos, 1, i - 1)
      QueCampos = Mid(QueCampos, i + 1)
      i = CInt(txt)
      mTag.Cargar Text1(i)
    'TIENE QUE ESTAR CARGADO  If mTag.Cargado Then

                Debug.Print mTag.columna
                        
                        
                If mTag.Vacio = "S" Then
                    Valor = DBLet(Data1.Recordset.Fields(mTag.columna))
                Else
                    Valor = Data1.Recordset.Fields(mTag.columna)
                End If
                If mTag.Formato <> "" And CStr(Valor) <> "" Then
                    If mTag.TipoDato = "N" Then
                        'Es numerico, entonces formatearemos y sustituiremos
                        ' La coma por el punto
                        txt = Format(Valor, mTag.Formato)
                        
                    Else
                        txt = Format(Valor, mTag.Formato)
                    End If
                Else
                    If mTag.TipoDato = "N" Then
                        If Val(Valor) = 0 Then
                            txt = ""
                        Else
                           txt = Valor
                        End If
                    Else
                        txt = Valor
                    End If
                End If

                If Text1(i).Text <> txt Then
                    fin = True
                    'Por si acaso el campo que cambia ES EL ULTIMO
                    If QueCampos = "" Then QueCampos = "NO"
                Else
                    fin = QueCampos = ""
                End If
    Wend
    

    'PREGUNTA
    If QueCampos <> "" Then
        'Significa que ha cambiado algo
        If MsgBox("Actualizar datos cuenta en contabilidad", vbQuestion + vbYesNo) = vbYes Then HayQueActualizarenContabilidad = True
        
    End If
End Function



Private Sub UpdatearNomProve()
Dim i As Byte
    
    For i = 1 To 5
        CadenaConsulta = RecuperaValor("scappr|scaalp|scafpc|schalp|schppr|", CInt(i))
        lblIndicador.Caption = "Actualiza " & CadenaConsulta
        lblIndicador.Refresh
        CadenaConsulta = "UPDATE " & CadenaConsulta & " SET nomprove=" & DBSet(Text1(1).Text, "T")
        CadenaConsulta = CadenaConsulta & " WHERE codprove = " & Text1(0).Text
        conn.Execute CadenaConsulta
        DoEvents
    Next
    
    CadenaConsulta = "PROV.   " & Format(Text1(0).Text, "000000") & "-> " & Text1(1).Text
    Set LOG = New cLOG
    LOG.Insertar 21, vUsu, CadenaConsulta
    Set LOG = Nothing
End Sub

                    

Private Sub CargaComboPais()
    cboPais.Clear
    If Not vParamAplic.ContabilidadNueva Then Exit Sub
    
    cboPais.AddItem "ESPAÑA  (ES)"
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from paises where codpais <>'ES' and nompais<>'' order by nompais", ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboPais.AddItem miRsAux!nompais & "   (" & miRsAux!codpais & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Function PaisSeleccionado() As String

    If cboPais.ListIndex < 0 Then
        PaisSeleccionado = ""
    Else
        PaisSeleccionado = Mid(cboPais.Text, InStr(1, cboPais.Text, "(") + 1, 2)
    End If
End Function

Private Sub PonerPais()
Dim i As Integer

    
    
    If DBLet(Data1.Recordset!codpais, "T") = "" Then
        i = -1
    Else
        For i = 0 To cboPais.ListCount - 1
            If InStr(1, cboPais.List(i), "(" & Data1.Recordset!codpais & ")") > 0 Then
                'Este es el pais
                Exit For
            End If
        Next
        If i >= cboPais.ListCount Then i = -1
    End If
    
    cboPais.ListIndex = i
End Sub

Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    VerLineasDirecciones False
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento Button.Index - 1
End Sub


Private Sub ImagenDocumento(DatoEnElTag As Byte)

    On Error Resume Next
    
    imgDocumentos.Picture = frmPpal.ImgListPpal.ListImages(DatoEnElTag).Picture
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub





'  Vamos a ver que en los campos Text1(1) , Text1(2)  TExt1(4) NO hayan *
'  ComprobarAsteriscosEnTextbox Text1(1) , "1|2|4|"
Private Function ComprobarTieneAsteriscosEnTextbox(ByVal secuencia As String) As Boolean
Dim i As Integer
Dim N As Integer
Dim C As String

    ComprobarTieneAsteriscosEnTextbox = False
    Do
        i = InStr(1, secuencia, "|")
        If i = 0 Then
            secuencia = ""
        Else
            C = Mid(secuencia, 1, i - 1)
            secuencia = Mid(secuencia, i + 1)
            N = CInt(C)
            If TieneCampoTextoAsterisco(Text1(N)) Then
                ComprobarTieneAsteriscosEnTextbox = True
                MsgBox "Carcater asterisco NO permitido: " & vbCrLf & Text1(N).Text, vbExclamation
                secuencia = ""
                PonerFoco Text1(N)
            End If
        End If
    Loop Until secuencia = ""
End Function




