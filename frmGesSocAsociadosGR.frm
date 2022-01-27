VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGesSocAsociadosGR 
   Caption         =   "Datos básicos"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   105
   ClientWidth     =   12615
   Icon            =   "frmGesSocAsociadosGR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   12615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   86
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   87
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
      Left            =   3870
      TabIndex        =   84
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   85
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
      Height          =   330
      Left            =   10890
      TabIndex        =   83
      Top             =   270
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   240
      TabIndex        =   47
      Top             =   900
      Width           =   12210
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
         Left            =   1065
         TabIndex        =   0
         Tag             =   "Código|N|N|0|99999|asociados|idasoc|000000|S|"
         Text            =   "Text"
         Top             =   210
         Width           =   1155
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
         Left            =   3375
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre Trabajador|T|N|||asociados|NomLargo||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   8625
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Left            =   285
         TabIndex        =   49
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   780
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
      Left            =   11430
      TabIndex        =   33
      Top             =   9780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   210
      TabIndex        =   35
      Top             =   9735
      Width           =   2895
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
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   2715
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
      Left            =   11430
      TabIndex        =   34
      Top             =   9780
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
      Left            =   10110
      TabIndex        =   32
      Top             =   9780
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3000
      Top             =   9360
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
      Left            =   4440
      Top             =   9360
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
      Height          =   7935
      Left            =   240
      TabIndex        =   37
      Top             =   1755
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   13996
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
      TabCaption(0)   =   "Asociados gestion"
      TabPicture(0)   =   "frmGesSocAsociadosGR.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(34)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(36)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(37)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(25)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ImgMail(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgFecha(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(16)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgFecha(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgFecha(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(40)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "imgBuscar(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "imgBuscar(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DataGrid1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(8)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(2)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(19)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(15)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(16)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(17)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(18)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(12)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(11)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(13)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(14)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(10)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(20)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Combo1(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Check1(4)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame3"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "FrameSocio"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtAux1(0)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtAux1(1)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cboEntidades"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cboSeccionGesoc"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "FrameToolAux0"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Email  /  Histórico"
      TabPicture(1)   =   "frmGesSocAsociadosGR.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1(0)"
      Tab(1).Control(1)=   "Check1(1)"
      Tab(1).Control(2)=   "Check1(2)"
      Tab(1).Control(3)=   "Check1(3)"
      Tab(1).Control(4)=   "txtHco(1)"
      Tab(1).Control(5)=   "txtHco(0)"
      Tab(1).Control(6)=   "DataGrid2"
      Tab(1).Control(7)=   "lblDpto(0)"
      Tab(1).Control(8)=   "Line1"
      Tab(1).Control(9)=   "lblDpto(5)"
      Tab(1).Control(10)=   "lblDpto(6)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Datos III"
      TabPicture(2)   =   "frmGesSocAsociadosGR.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAux2"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   135
         TabIndex        =   88
         Top             =   5040
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   89
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
      Begin VB.ComboBox cboSeccionGesoc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   5760
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cboEntidades 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   5760
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Se le enviarán habitualmente los comunicados de la Cooperativa"
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   81
         Tag             =   "Envio de correo normal|N|N|||asociados|Correo|||"
         Top             =   465
         Width           =   7155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autorización para comunicaciones comerciales a otras empresas"
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
         Left            =   -74640
         TabIndex        =   80
         Tag             =   "Aurorizo comunicaciones|N|N|||asociados|Auto3|||"""
         Top             =   1125
         Width           =   6780
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autorización para envios de correo electrónico"
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
         Index           =   2
         Left            =   -74640
         TabIndex        =   79
         Tag             =   "Autorizo correo electrónico|N|N|||asociados|Auto2|||"
         Top             =   780
         Width           =   5595
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autorización para procesar y almacenar los datos informáticamente"
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
         Left            =   -74640
         TabIndex        =   78
         Tag             =   "Autorizo tratamiento informático|N|N|||asociados|Auto1|||"
         Top             =   1440
         Width           =   7560
      End
      Begin VB.TextBox txtAux1 
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
         Height          =   315
         Index           =   1
         Left            =   2280
         MaxLength       =   70
         TabIndex        =   76
         Tag             =   "Formación|T|N|||strab1|formacio||N|"
         Text            =   "Formacion Formacion Formacion Formacion Formacion Formacion Formacion "
         Top             =   5640
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.TextBox txtAux1 
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
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   15
         TabIndex        =   75
         Tag             =   "Periodo|T|N|||strab1|periodos||N|"
         Text            =   "Periodo"
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtHco 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   1
         Left            =   -67080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Text            =   "frmGesSocAsociadosGR.frx":0060
         Top             =   5160
         Width           =   3855
      End
      Begin VB.TextBox txtHco 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   -67200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   70
         Text            =   "frmGesSocAsociadosGR.frx":0066
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Frame FrameSocio 
         Caption         =   "Ariagro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         TabIndex        =   63
         Top             =   4215
         Width           =   11850
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
            Height          =   315
            Index           =   27
            Left            =   840
            MaxLength       =   10
            TabIndex        =   27
            Tag             =   "Codigo EUROAGRO|N|S|||asociados|CodSocEuroagro|||"
            Text            =   "Text1"
            Top             =   315
            Width           =   885
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
            Left            =   8115
            MaxLength       =   10
            TabIndex        =   30
            Tag             =   "Aportación obligatoria|N|S|||asociados|AportObligatoria|##,###,##0.00||"
            Text            =   "123456.36"
            Top             =   315
            Width           =   1020
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
            Left            =   10665
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Aportación voluntaria|N|S|||asociados|AportVoluntaria|##,###,##0.00||"
            Top             =   315
            Width           =   1020
         End
         Begin VB.ComboBox Combo1 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "Tipo de IVA|N|S|||asociados|CodIva|||"
            Top             =   315
            Width           =   1605
         End
         Begin VB.ComboBox Combo1 
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
            Left            =   4410
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "Tipo para IRPF|T|N|||asociados|TipoIRPF|||"
            Top             =   315
            Width           =   1770
         End
         Begin VB.Label Label2 
            Caption         =   "Cuota entrada"
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
            Index           =   4
            Left            =   9195
            TabIndex        =   68
            Top             =   375
            Width           =   1440
         End
         Begin VB.Label Label2 
            Caption         =   "Aportación capital"
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
            Left            =   6270
            TabIndex        =   67
            Top             =   375
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "IRPF"
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
            Index           =   2
            Left            =   3930
            TabIndex        =   66
            Top             =   375
            Width           =   600
         End
         Begin VB.Label Label2 
            Caption         =   "Código"
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
            Left            =   105
            TabIndex        =   65
            Top             =   375
            Width           =   840
         End
         Begin VB.Label Label2 
            Caption         =   "IVA"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   64
            Top             =   375
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Suministros / Gasolinera "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   57
         Top             =   3465
         Width           =   11895
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
            Left            =   6975
            MaxLength       =   4
            TabIndex        =   25
            Tag             =   "Tarifa|N|S|||asociados|tarifaprecio|0|N|"
            Text            =   "Text1"
            Top             =   225
            Width           =   570
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   375
            Left            =   8265
            TabIndex        =   58
            Top             =   225
            Width           =   3495
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
               Left            =   1935
               MaxLength       =   10
               TabIndex        =   26
               Tag             =   "Cta|T|S|||asociados|codmacta||N|"
               Text            =   "Text1"
               Top             =   0
               Width           =   1530
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta gasolinera"
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
               Left            =   45
               TabIndex        =   62
               Top             =   60
               Width           =   1770
            End
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
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "Dto Gnral|N|N|||asociados|DtoPpago|0.00|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   570
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
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "Dto Gnral|N|N|||asociados|DtoGnral|0.00|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Actividad en gestión"
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
            Left            =   4860
            TabIndex        =   61
            Top             =   270
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Dto P.P."
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
            Left            =   3000
            TabIndex        =   60
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Dto gnral"
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
            Left            =   480
            TabIndex        =   59
            Top             =   270
            Width           =   960
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Es socio"
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
         Left            =   2850
         TabIndex        =   3
         Tag             =   "Socio|N|N|||asociados|essocio|||"
         Top             =   600
         Width           =   1695
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
         Index           =   0
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Situacion|N|N|||asociados|Estado|||"
         Top             =   600
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
         Index           =   20
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "T|T|S|||asociados|movil|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   1530
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
         Left            =   4050
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "T|T|S|||asociados|telefono3|||"
         Text            =   "Text1"
         Top             =   2745
         Width           =   1530
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
         Height          =   1470
         Index           =   14
         Left            =   6960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "o|T|S|||asociados|Observaciones||N|"
         Text            =   "frmGesSocAsociadosGR.frx":006C
         Top             =   2040
         Width           =   5055
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
         Left            =   9855
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Fecha de Baja|F|S|||asociados|fechabaja|dd/mm/yyyy|N|"
         Top             =   1080
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
         Index           =   11
         Left            =   4230
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Fecha de Nacimiento|F|S|||asociados|fechanac|dd/mm/yyyy|N|"
         Top             =   3150
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
         Index           =   12
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha de Alta|F|S|||asociados|FechaAlta|dd/mm/yyyy|N|"
         Top             =   1080
         Width           =   1350
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
         Left            =   9435
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Cuenta|T|S|||asociados|NUmcc|0000000000||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1755
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
         Left            =   9045
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "DC|T|S|||asociados|DC|00||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   360
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
         Left            =   8340
         MaxLength       =   4
         TabIndex        =   19
         Tag             =   "Sucur.|N|S|0|9999|asociados|sucursal|0000|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   660
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
         Left            =   7650
         MaxLength       =   4
         TabIndex        =   18
         Tag             =   "Ban|N|S|0|9999|asociados|entidad|0000|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   660
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
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   17
         Tag             =   "IBAN|T|S|||asociados|iban|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   660
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
         Left            =   1170
         MaxLength       =   35
         TabIndex        =   4
         Tag             =   "D|T|S|||asociados|Direccion|||"
         Text            =   "Text1"
         Top             =   1065
         Width           =   4365
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
         Left            =   1170
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Teléfono|T|S|||asociados|telefono1|||"
         Text            =   "Text1"
         Top             =   2745
         Width           =   1260
      End
      Begin VB.TextBox txtAux2 
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
         Height          =   315
         Left            =   -74160
         MaxLength       =   70
         TabIndex        =   45
         Tag             =   "Habilidad|T|N|||strab2|habilida||N|"
         Text            =   "Habilidad"
         Top             =   3120
         Visible         =   0   'False
         Width           =   6735
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
         Left            =   2625
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "T|T|S|||asociados|Telefono2||N|"
         Text            =   "Text1"
         Top             =   2745
         Width           =   1290
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
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "e-mail|T|S|||asociados|Mail||N|"
         Text            =   "Text1"
         Top             =   2325
         Width           =   4410
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
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "N.I.F.|T|N|||asociados|nif|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1410
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
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Pro.|T|S|||asociados|provincia||N|"
         Text            =   "Text1"
         Top             =   1905
         Width           =   4410
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
         Left            =   3090
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Pob|T|S|||asociados|Poblacion||N|"
         Text            =   "Text1"
         Top             =   1530
         Width           =   2460
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
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "C.Postal|T|N|||asociados|CodPostal||N|"
         Text            =   "Text1"
         Top             =   1530
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmGesSocAsociadosGR.frx":0072
         Height          =   1815
         Left            =   165
         TabIndex        =   56
         Top             =   5880
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   3201
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmGesSocAsociadosGR.frx":0087
         Height          =   5535
         Left            =   -74880
         TabIndex        =   69
         Top             =   2160
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
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
               LCID            =   1033
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
               LCID            =   1033
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
         Index           =   0
         Left            =   900
         Picture         =   "frmGesSocAsociadosGR.frx":009C
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   -74760
         TabIndex        =   82
         Top             =   1920
         Width           =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -63210
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   -67200
         TabIndex        =   72
         Top             =   1995
         Width           =   900
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   -67080
         TabIndex        =   71
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   6675
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
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
         Left            =   5760
         TabIndex        =   55
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
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
         TabIndex        =   54
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Observ."
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
         Left            =   5760
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   9495
         Picture         =   "frmGesSocAsociadosGR.frx":0A9E
         ToolTipText     =   "Buscar fecha"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Baja"
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
         Left            =   8535
         TabIndex        =   52
         Top             =   1110
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3990
         Picture         =   "frmGesSocAsociadosGR.frx":0B29
         ToolTipText     =   "Buscar fecha"
         Top             =   3150
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "F.Nacim."
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
         Left            =   2955
         TabIndex        =   51
         Top             =   3150
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   6675
         Picture         =   "frmGesSocAsociadosGR.frx":0BB4
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Alta"
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
         Left            =   5760
         TabIndex        =   50
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   885
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2355
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Móvil"
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
         Left            =   120
         TabIndex        =   46
         Top             =   3150
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfonos"
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
         TabIndex        =   44
         Top             =   2775
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   43
         Top             =   2355
         Width           =   630
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
         Height          =   255
         Index           =   36
         Left            =   120
         TabIndex        =   42
         Top             =   630
         Width           =   600
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
         Index           =   15
         Left            =   120
         TabIndex        =   41
         Top             =   1950
         Width           =   960
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
         Index           =   34
         Left            =   2130
         TabIndex        =   40
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "CPostal"
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
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   780
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
         Index           =   13
         Left            =   120
         TabIndex        =   38
         Top             =   1095
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   5760
      Top             =   9360
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
   Begin VB.Menu mnMtoLineas 
      Caption         =   "A&vanzadas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnEstudios 
         Caption         =   "&Crear asociado en sección"
         HelpContextID   =   2
      End
      Begin VB.Menu mnHabilidades 
         Caption         =   "&Habilidades"
         HelpContextID   =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnExperiencia 
         Caption         =   "Experiencia &Laboral"
         HelpContextID   =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnFormRealizada 
         Caption         =   "&Formación Realizada"
         HelpContextID   =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnFormEmpresa 
         Caption         =   "F&ormacion Empresa"
         HelpContextID   =   2
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmGesSocAsociadosGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB1 As frmBasico2 'frmBuscaGrid 'Form para busquedas
Attribute frmB1.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1


Private Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Byte 'Indica que numero de Tab que esta en modo Mantenimiento

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas del Mantenimiento en que estemos

Private kCampo As Integer

Dim BuscaChekc  As String

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim miSQL As String
Private InsertarEnCambiosSituacion As String

'BD ArigasolXX
' Para saber a que empresa ataca
Private BD_Arigaso_l As String
Private udNegocioGasol As Byte



Private Sub cboEntidades_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboEntidades_LostFocus()
    PonerFocoBtn cmdAceptar
End Sub

Private Sub cboSeccionGesoc_Click()
    If Modo <> 5 Then Exit Sub
    If cboSeccionGesoc.visible Then
        Me.cboEntidades.visible = cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex) = 1
    Else
        cboEntidades.visible = False
    End If
    
End Sub

Private Sub cboSeccionGesoc_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub Check1_Click(Index As Integer)

    
    If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
    
    
    
    If Index = 4 And Modo > 2 Then
        'INSERTANDO MODIFICANDO
        If Me.Check1(4).Value = 1 Then
            CadenaConsulta = ""
            If Modo = 4 Then
                If DBLet(Me.Data1.Recordset!CodSocEuroagro, "N") > 0 Then CadenaConsulta = "NO"
            End If
            
            If CadenaConsulta = "" Then
                'OK. Vamos a crear NUEVO socio. Codigo menor que select * from asociados where IdAsoc = 5582 and (fechabaja is null) and CodSocEuroagro < 10000 and EsSocio = 1
                CadenaConsulta = DevuelveDesdeBD(conAri, "max(codsocio)", "ariagro.rsocios", "codsocio<10000 AND 1", "1")
                CadenaConsulta = Val(CadenaConsulta) + 1
                Text1(27).Text = Format(CadenaConsulta, "0000")
                
                
                
            End If
        Else
            If Modo = 3 Then Text1(27).Text = ""
        End If
    End If
    
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'===========================================================================
'       PROCEDIMIENTOS
'============================================================================

Private Sub cmdAceptar_Click()
Dim Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then
                    ActAriFacElec CLng(Text1(0).Text)
                    Data1.RecordSource = "Select * from asociados where idasoc=" & Text1(0).Text
                    PosicionarData
              End If
            End If
            
        Case 4  'MODIFICAR
            InsertarEnCambiosSituacion = ""
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                
                    'Noviembre 2014
                    If InsertarEnCambiosSituacion <> "" Then
                        'Ha cambiado algo en la situacion. Le ofertamos un forma para "explicarlo"
                        BuscaChekc = Me.Text1(1).Text & "|" & InsertarEnCambiosSituacion & "|"
                        CadenaDesdeOtroForm = ""
                        frmListado5.OpcionListado = 5
                        frmListado5.OtrosDatos = BuscaChekc
                        frmListado5.Show vbModal
                        
                        'Guardamos en el LOG de hco de cambios de situacion
                        'NUMREGELIM lleva vlor para cambios
                        BuscaChekc = "INSERT INTO asociados_hcocambios(IdAsoc,FechaCambio,usuario,cambios,situacion,observaciones,tipoCambio,FechaCampo) VALUES (" & Data1.Recordset!IdAsoc
                        BuscaChekc = BuscaChekc & ",now()," & DBSet(vUsu.Login, "T") & "," & NumRegElim & ","
                        BuscaChekc = BuscaChekc & DBSet(InsertarEnCambiosSituacion, "T") & "," & DBSet(CadenaDesdeOtroForm, "T", "N") & ","
                        BuscaChekc = BuscaChekc & DBSet(TituloLinea, "T") & "," & DBSet(miSQL, "F", "S") & ")"
                        ejecutar BuscaChekc, False
                        Espera 0.5
                        TituloLinea = ""
                    End If
                
                    Screen.MousePointer = vbHourglass
                    'EN Arifacelec, siempre
                    ActAriFacElec CLng(Text1(0).Text)
                
                
                    'ACtualizar
                    ActualizarEnUnidadesDeNegocio
                
                    TerminaBloquear
                    PosicionarData
                    'Refrescamos
                    If InsertarEnCambiosSituacion <> "" Then
                        PonerCamposLineas
                        InsertarEnCambiosSituacion = ""
                    End If
                    
                    Screen.MousePointer = vbDefault
                End If
            End If
                
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)

            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then
                    Select Case Me.SSTab1.Tab
                        Case 0 'Estudios/Formacion - Datos de la tabla strab1
                            CargaGrid1 DataGrid1, Data2, DevSQLGrid(1, True)
                        Case 1 'Habilidades
                            ' CargaGrid DataGrid2, Data3, Cad
                       
                    End Select
                    'BotonAnyadirLinea
                    
                    cmdCancelar_Click
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
'--
'                    PonerBotonCabecera True
                    ModificaLineas = 0
                    Select Case Me.SSTab1.Tab
                        Case 0 'Estudios/Formacion - Datos de la tabla strab1
                            NumRegElim = Data2.Recordset.AbsolutePosition
                            CargaTxtAux1 False, False
                            'CargaGrid DataGrid1, Data2, cad
                            CargaGrid2 DataGrid1, Data2
                            SituarDataPosicion Data2, NumRegElim, Indicador
                            '++
                            PonerModo 2
                            Me.lblIndicador.Caption = ""
'                        Case 2 'Habilidades
'                            NumRegElim = Data3.Recordset.AbsolutePosition
'                            CargaTxtAux2 False, False
'                            'CargaGrid DataGrid2, Data3, cad
'                            CargaGrid2 DataGrid2, Data3
'                            SituarDataPosicion Data3, NumRegElim, Indicador
                       
                    End Select
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
            Select Case Me.SSTab1.Tab
                Case 0 'Estudios/Formacion
                    CargaTxtAux1 False, False
                    DataGrid1.Enabled = True
                    If ModificaLineas = 1 Then 'Insertar
                        DataGrid1.AllowAddNew = False
                        If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                    End If
                Case 22 'Habilidades
                    CargaTxtAux2 False, False
                    DataGrid2.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid2.AllowAddNew = False
                        If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                    End If
            End Select
'--
'            PonerBotonCabecera True
            PonerModo 2
            ModificaLineas = 0
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de asociados: asociados (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Me.Check1(4).Value = 0
    
    Text1(0).Text = SugerirCodigoSiguienteStr("asociados", "idasoc", "")    '"idasoc<8500")  --> Sep 2021  Chelo dice que lo quitemos. ademas ya no hay hueco
    Text1(12).Text = Format(Now, "dd/mm/yyyy")
    Text1(21).Text = "0": Text1(22).Text = "0": Text1(23).Text = "1"
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
'    If NumTabMto <> Me.SSTab1.Tab Then
'        MsgBox "No puede Añadir. Esta en Modo Mantenimiento de otra linea.", vbExclamation
'        Exit Sub
'    End If
    
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
'--
'    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    Select Case Me.SSTab1.Tab
        Case 0 'Estudios / Formacion
                'Situamos el grid al final
                CargaUnidadesdeNegocio True
                If Me.cboSeccionGesoc.ListCount > 0 Then
                    AnyadirLinea DataGrid1, Data2
                    CargaTxtAux1 True, True
                    PonerFocoCbo Me.cboSeccionGesoc
                    
                Else
                    CargaTxtAux1 False, False
                    Me.cmdAceptar.visible = False
                    MsgBox "Se ha dado de alta en todas las unidades de negocio", vbExclamation
'--
'                    PonerBotonCabecera True
                    PonerModo 2
                    Me.lblIndicador.Caption = ""
                End If
        Case 22 'Habilidades
        
                AnyadirLinea DataGrid2, Data3
                CargaTxtAux2 True, True
                PonerFoco txtAux2
     
    End Select
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        
        LimpiarCampos
        BuscaChekc = ""
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
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
'Ver todos
    LimpiarCampos
    Me.SSTab1.Tab = 0
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea


    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Modificar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    
    Select Case Me.SSTab1.Tab
        Case 0 'Estudios/Formacion
                If Data2.Recordset.EOF Then Exit Sub
                
                
                
                CargaTxtAux1 True, False
                DataGrid1.Enabled = False
                PonerFoco txtAux1(1)
        Case 2 'Habilidades
        
        
        'vWhere = "idasoc=" & Val(Text1(0).Text) & " and numlinea="
'                If Data3.Recordset.EOF Then Exit Sub
'                vWhere = vWhere & Data3.Recordset!numlinea
'                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
'                CargaTxtAux2 True, False
'                DataGrid2.Enabled = False
'                PonerFoco txtAux2

    End Select
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'--
'    PonerBotonCabecera False
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de asociados (asociados)
Dim Cad As String
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Not PuedeEliminarTrabajador Then Exit Sub
    
    
    Cad = "Cabecera de asociados." & vbCrLf
    Cad = Cad & "------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Trabajador:"
    Cad = Cad & vbCrLf & "Código:   " & Format(Data1.Recordset.Fields(0), "000000")
    Cad = Cad & vbCrLf & "Descripción:   " & Data1.Recordset.Fields(1)
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    
    
    
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        If Not Eliminar Then
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Trabajador", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Trabajador. Tablas: strab1, strab2, strab3, strab4, strab5
Dim SQL As String
Dim numlinea As Integer
On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

     If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede eliminar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If

    Select Case Me.SSTab1.Tab
        Case 1 'EStudios/Formacion
        
            MsgBox "No se puede eliminar socios en secciones", vbExclamation
            Exit Sub
        
            If Data2.Recordset.EOF Then Exit Sub
            numlinea = Data2.Recordset!numlinea
        Case 2 'Habilidades
            If Data3.Recordset.EOF Then Exit Sub
            numlinea = Data3.Recordset!numlinea
'        Case 3 'Experiencia Laboral
'            If Data4.Recordset.EOF Then Exit Sub
'            numlinea = Data4.Recordset!numlinea
'        Case 4 'Formacion Realizada
'            If Data5.Recordset.EOF Then Exit Sub
'            numlinea = Data5.Recordset!numlinea
'        Case 5 'Formacion Empresa
'            If Data6.Recordset.EOF Then Exit Sub
'            numlinea = Data6.Recordset!numlinea
    End Select
    
'    ModificaLineas = 3 'Eliminar
'    SQL = "¿Seguro que desea eliminar la línea de " & TituloLinea & "?"
'    SQL = SQL & vbCrLf & "Cod. Traba.: " & Format(Data1.Recordset!IdAsoc, "000000")
'    SQL = SQL & vbCrLf & "Nombre: " & Data1.Recordset!NomTraba
'    SQL = SQL & vbCrLf & "Numlinea: " & numlinea
'    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        SQL = "Delete from " & NomTablaLineas & " where idasoc=" & Data1.Recordset!IdAsoc
'        SQL = SQL & " and numlinea=" & numlinea
'        conn.Execute SQL
'
'        ModificaLineas = 0
'        Select Case Me.SSTab1.Tab
'            Case 1: 'Estudios/Formacion
''                CancelaADODC (Data2)
'                CargaGrid2 DataGrid1, Data2
''                CancelaADODC (Data2)
'            Case 2: 'Habilidades
'                CargaGrid2 DataGrid2, Data3
''            Case 3: 'Experiencia Laboral
''                CargaGrid2 DataGrid3, Data4
''            Case 4 'Formacion Realizada
''                CargaGrid2 DataGrid4, Data5
''            Case 5 'Formacion Empresa
''                CargaGrid2 DataGrid5, Data6
'        End Select
''        CancelaADODC
'    End If
'    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Trabajador", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera tambien
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de asociados
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


Private Sub Data6_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub data3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim Lim As Boolean

    On Error GoTo EM
    Me.txtHco(0).Text = "": Me.txtHco(1).Text = ""
    Lim = True
    If Not Data3.Recordset.EOF Then
            Me.txtHco(0).Text = DBLet(Data3.Recordset!Situacion, "T")
            Me.txtHco(1).Text = Data3.Recordset!Observaciones
            Lim = False
    End If
    
    
    
    
    
   Exit Sub
EM:
    If Err.Number <> 0 Then Err.Clear

    If Lim Then Me.txtHco(0).Text = "": Me.txtHco(1).Text = ""
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    
    'Icono de imagen de e-mail
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(20).Picture
    '++
    Me.imgBuscar(1).Picture = Me.imgBuscar(0).Picture
    
    BuscaChekc = "UdNegocioGasol"
    BD_Arigaso_l = Trim(DevuelveDesdeBD(conAri, "BDGasol", "parametros", "1", "1", , BuscaChekc))
    'If Val(BD_Arigaso_l) = 0 Then Err.Raise 513, , "Mal configurado el programa. Falta campo en parametros: BD_Arigasol"
    BD_Arigaso_l = "arigasol" & BD_Arigaso_l
    udNegocioGasol = CByte(BuscaChekc)
    BuscaChekc = ""
    
    'ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 19
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 25 'Estudios/Formacion
'        .Buttons(11).Image = 27 'Habilidades
'        .Buttons(12).Image = 37 'Experiencia Laboral
'        .Buttons(13).Image = 28 'Formacion Realizada
'        .Buttons(14).Image = 29 'Formacion Empresa
'
'        .Buttons(16).Image = 16  'Salir
'        .Buttons(17).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
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
        
    With Me.ToolAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
    End With
    
    Me.SSTab1.Tab = 0
    'SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    
    CargaCombos
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
         
         
    
         
    '## A mano
    NombreTabla = "asociados"
    Ordenacion = " ORDER BY idasoc"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where IdAsoc=-1"
    Data1.Refresh
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    Combo1(2).ListIndex = -1
    Me.Check1(4).Value = 0: Check1(1).Value = 0: Check1(2).Value = 0: Check1(3).Value = 0: Check1(0).Value = 0
    
    Me.txtHco(0).Text = "": Me.txtHco(1).Text = ""
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB1_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub



Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = Val(imgFecha(0).Tag) + 11
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte
    
    If Index = 1 Then
        VerObservaciones
        Exit Sub
    End If
    
    
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'CPostal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 4
            'VieneDeBuscar = True
            If CadenaConsulta <> "" Then
                Text1(3).Text = RecuperaValor(CadenaConsulta, 1)
                Text1_LostFocus 3
            End If
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub

Private Sub VerObservaciones()
    CadenaDesdeOtroForm = Me.Text1(14).Text
    frmFacClienteObser.Modificar = Modo >= 3
    frmFacClienteObser.Text1 = CadenaDesdeOtroForm
    frmFacClienteObser.Show vbModal
    'Llevara DOS VALORES.
    'Si modifica y el texto
    If Modo = 3 Or Modo = 4 Then
        If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(14).Text = Mid(CadenaDesdeOtroForm, 3)
    End If
    CadenaDesdeOtroForm = ""

End Sub

Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   Me.imgFecha(0).Tag = Index
   Indice = Index + 11
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        dirMail = Text1(9).Text
    End If
    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de asociados
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnEstudios_Click()


    If Modo <> 2 Then Exit Sub
    
    
    'Si la situacion es de baja no dejo pasar
    If Combo1(0).ListIndex > 0 Then
        MsgBox "Solo se puede modificar en asociados activos", vbExclamation
        Exit Sub
    End If
    
    'frmListado5.OtrosDatos = Me.Text1(0).Text
    'frmListado5.OpcionListado = 1
    'frmListado5.Show vbModal
   '
    'CargaGrid1 DataGrid1, Data2, DevSQLGrid(1, True)
    
    

    
'Abre Mantenimiento de lineas  Estudios/Formacion
    BotonMtoLineas 1, "Entidades"
    NomTablaLineas = "Entidades"
End Sub

Private Sub mnExperiencia_Click()
'Abre Mantenimiento de lineas Experiencia Laboral
    BotonMtoLineas 3, "Experiencia Laboral"
    NomTablaLineas = "strab3"
End Sub

Private Sub mnFormEmpresa_Click()
'Abre Mantenimiento de lineas Formacion Empresa
    BotonMtoLineas 5, "Formación Empresa"
    NomTablaLineas = "strab5"
End Sub

Private Sub mnFormRealizada_Click()
'Abre Mantenimiento de lineas Formacion Realizada
    BotonMtoLineas 4, "Formación Realizada"
    NomTablaLineas = "strab4"
End Sub

Private Sub mnHabilidades_Click()
'Abre Mantenimiento de lineas Habilidades
    BotonMtoLineas 2, "Habilidades"
    NomTablaLineas = "strab2"
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         Me.SSTab1.Tab = 0
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Trabajador
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
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
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index <> 14 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
    If Index = 14 Then Exit Sub
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 0 'poblacion
            Case 11: KEYFecha KeyAscii, 0 'fecha nacimiento
            Case 12: KEYFecha KeyAscii, 1 'fecha alta
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
    imgFecha_Click (Indice)
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
    
    If Modo = 3 Or Modo = 4 Then
        If Index <> 14 Then Text1(Index).Text = UCase(Text1(Index).Text)
    End If
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod. Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de trabajador en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
        Case 3 'CPostal
             If Not VieneDeBuscar Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 6 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF_ Text1(Index).Text, False
            devuelve = ""
            If Text1(Index).Text <> "" Then
                If Modo = 3 Then
                    devuelve = " TRUE "
                ElseIf Modo = 4 Then
                    devuelve = " idasoc <> " & Data1.Recordset!IdAsoc
                End If
                
                If devuelve <> "" Then
                    devuelve = devuelve & " AND nif "
                    devuelve = DevuelveDesdeBD(conAri, "concat(idasoc,' ',coalesce(nomlargo,'N/D'))", "asociados", devuelve, Text1(Index).Text, "T")
                    If devuelve <> "" Then MsgBox "Ya tiene un socio ese NIF: " & devuelve, vbExclamation
                    
                End If
            End If
        Case 10
            
            
        Case 11, 12, 13 'Fecha Nacimiento, Fecha alta, Fecha baja
            'Si no es modo de Busqueda poner el formato
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 24 'Cod almacen
            
        Case 25
            'Agente
            
        Case 15, 16, 17, 18 'cod. banco, cod. sucursal
            PonerFormatoEntero Text1(Index)
            
                        
            
            If Index = 18 Then
                If Me.Text1(Index).Text <> "" Then
                    Me.Text1(Index).Text = Right(String(10, "0") & Text1(Index).Text, 10)
                    devuelve = Text1(15).Text & Me.Text1(16).Text & Me.Text1(17).Text & Me.Text1(18).Text
                
                    If Len(devuelve) = 20 Then
                        DevuelveIBAN2 "ES", devuelve, devuelve
                        If Len(devuelve) = 2 Then
                            devuelve = "ES" & devuelve
                            If Me.Text1(19).Text = "" Then
                                Text1(19).Text = devuelve
                            Else
                                If Me.Text1(19).Text <> devuelve Then MsgBox "Codigo IBAN distinto del calculado [" & devuelve & "]", vbExclamation
                            End If
                        End If
                    End If
                    devuelve = ""
                End If
            End If
                        
            
            
            
        Case 18, 22
            'Cuenta banco.
            'Si esta bien puesta la cuenta calculamos el iban
            If Me.Text1(Index).Text <> "" Then
                Me.Text1(Index).Text = Right(String(10, "0") & Text1(Index).Text, 10)
                devuelve = Text1(Index - 3).Text & Me.Text1(Index - 2).Text & Me.Text1(Index - 1).Text & Me.Text1(Index).Text
                kCampo = 26
                If Index = 22 Then kCampo = 27
                If Len(devuelve) = 20 Then
                    DevuelveIBAN2 "ES", devuelve, devuelve
                    If Len(devuelve) = 2 Then
                        devuelve = "ES" & devuelve
                        If Me.Text1(kCampo).Text = "" Then
                            Text1(kCampo).Text = devuelve
                        Else
                            If Me.Text1(kCampo).Text <> devuelve Then MsgBox "Codigo IBAN distinto del calculado [" & devuelve & "]", vbExclamation
                        End If
                    End If
                End If
                devuelve = ""
                kCampo = Index
            End If

            
            
    End Select
End Sub

' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
'Private Sub PonceCentroCoste()
'Dim C As String
'    text1(10).Text = Trim(text1(10).Text)
'    C = ""
'    If text1(10).Text <> "" Then
'        C = PonerNombreDeCod(text1(10), conConta, "cabccost", "nomccost", "codccost")
'        If C = "" Then
'            MsgBox "No existe centro de coste", vbExclamation
'            text1(10).Text = ""
'        End If
'    End If
'    Text2(10).Text = C
'
'End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    'cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(0)
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)

'    'Llamamos a al form
'    '##A mano
'    'Cad = ""
'    'Cad = Cad & ParaGrid(Text1(0), 14, "Código")
'    'Cad = Cad & ParaGrid(Text1(1), 65, "Nombre")
'    'Cad = Cad & ParaGrid(Text1(6), 18, "NIF")
'    CadenaConsulta = ParaGrid(Text1(0), 14, "Código")
'    CadenaConsulta = CadenaConsulta & ParaGrid(Text1(1), 65, "Nombre")
'    CadenaConsulta = CadenaConsulta & ParaGrid(Text1(6), 18, "NIF")
'
''            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
'    'Tabla = "asociados"
'    'Titulo = "asociados"
'    'Me.imgFecha(0).Tag = 0
'
'        Screen.MousePointer = vbHourglass
'        Set frmB1 = New frmBuscaGrid
'        frmB1.vCampos = CadenaConsulta
'        frmB1.vTabla = "asociados"
'        frmB1.vSQL = cadB
'
'        '###A mano
'        frmB1.vDevuelve = "0|1|"
'        frmB1.vTitulo = "Asociados"
'        frmB1.vselElem = 1
'        frmB1.vConexionGrid = conAri
'
'        CadenaConsulta = ""
'        frmB1.Show vbModal
'        Set frmB1 = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If CadenaConsulta <> "" Then
'            CadenaConsulta = RecuperaValor(CadenaConsulta, 1)
'            CadenaConsulta = "Select * from asociados where idasoc = " & CadenaConsulta
'            PonerCadenaBusqueda
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If

    CadenaConsulta = ""
    
    Set frmB1 = New frmBasico2
    AyudaAsociadosGesSoc frmB1, Text1(0), cadB
    Set frmB1 = Nothing
    
    If CadenaConsulta <> "" Then
        CadenaConsulta = RecuperaValor(CadenaConsulta, 1)
        CadenaConsulta = "Select * from asociados where idasoc = " & CadenaConsulta
        PonerCadenaBusqueda
    Else   'de ha devuelto datos, es decir NO ha devuelto datos
        PonerFoco Text1(kCampo)
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
        PonerCampos
        PonerModo 2
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
    'Estudios/Formacion - Datos de la tabla strab1
    CargaGrid1 DataGrid1, Data2, DevSQLGrid(1, True)
    
    
    'Hco acciones realizadas
    
    BuscaChekc = "select IdAsoc,FechaCambio,usuario,tipoCambio,FechaCampo ,situacion,observaciones "
    BuscaChekc = BuscaChekc & " from asociados_hcocambios where IdAsoc= " & Data1.Recordset!IdAsoc
    BuscaChekc = BuscaChekc & " order by FechaCambio desc"
    CargaGrid1 DataGrid2, Data3, BuscaChekc

    
    
    'Habilidades
    'SQL = "Select * from strab2 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    'SQL = SQL & " order by numlinea"
    'CargaGrid DataGrid2, Data3, SQL

    'Experiencia Laboral
'    SQL = "Select * from strab3 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
'    SQL = SQL & " order by numlinea"
'    CargaGrid DataGrid3, Data4, SQL
'
'    'Formacion Realizada
'    SQL = "Select * from strab4 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
'    SQL = SQL & " order by numlinea"
'    CargaGrid DataGrid4, Data5, SQL
'
'    'Formacion Empresa
'    SQL = "Select * from strab5 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
'    SQL = SQL & " order by numlinea"
'    CargaGrid DataGrid5, Data6, SQL

    PrimeraVez = False
    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
    
    
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

    'Visualizar el login solo si es administrador o root
    B = (vUsu.Nivel < 2)
   
'--
'    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    
        
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
        cmdCancelar.Cancel = True
    End If
    
    '=======================================
    B = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1


    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    B = Not (Modo = 0 Or Modo = 2 Or Modo = 5)
    For i = 0 To Me.Check1.Count - 1
        Me.Check1(i).Enabled = B
    Next
    
    
    
    B = Modo = 1 Or Modo = 3 Or Modo = 4 'busqueda o inser/mod
    BloquearCmb Combo1(0), Not B
    BloquearCmb Combo1(1), Not B
    BloquearCmb Combo1(2), Not B
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2 '--And Modo <> 5
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B
    Next i
    
    'For I = 0 To Me.imgBuscar.Count - 1
    '    Me.imgBuscar(I).Enabled = B
    'Next I
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    'Solo en modo1
    Frame4.Enabled = Modo = 1
    'Solo buscar
    BloquearTxt Text1(24), Modo <> 1   'codmacta
    BloquearTxt Text1(27), Modo <> 1   'socio  ariagro
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean
Dim i As Byte
Dim EsBusqueda As Boolean
Dim bAux As Boolean


    EsBusqueda = Me.DatosADevolverBusqueda <> ""
    B = (Modo = 2 Or Modo = 0 Or Modo = 1) And Not EsBusqueda
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2) And Not EsBusqueda
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Mantenimiento lineas
    B = (Modo = 2) And Not EsBusqueda
'--
'    For i = 10 To 14
'        Toolbar1.Buttons(i).Enabled = B
'    Next i
'
'    Toolbar1.Buttons(16).Enabled = B Or Modo = 0
'    Me.mnEstudios.Enabled = B
'    Me.mnExperiencia.Enabled = B
'    Me.mnFormEmpresa.Enabled = B
'    Me.mnFormRealizada.Enabled = B
'    Me.mnHabilidades.Enabled = B
    
    '------------------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    
    '++
    Toolbar1.Buttons(8).Enabled = False
    
    
    If Not Data1.Recordset Is Nothing And Not Data1.Recordset.EOF Then B = (Data1.Recordset!Estado = 0)
    B = B And (Modo = 2) And DatosADevolverBusqueda = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        If B Then
            bAux = False
            If Not Me.Data2.Recordset Is Nothing Then
                bAux = (B And Me.Data2.Recordset.RecordCount > 0)
            End If
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
  PonerLongCamposGnral Me, Modo, 1
End Sub





Private Function DatosOk() As Boolean
Dim B As Boolean
Dim CambioSitu As String

On Error GoTo EDatosOK

    DatosOk = False
    B = True
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
          
          
    'Comprobamos CCC
    BuscaChekc = Format(Me.Text1(15).Text, "0000") & Format(Me.Text1(16).Text, "0000") & Format(Me.Text1(17).Text, "00") & Format(Me.Text1(18).Text, "0000000000")
    
    
    If BuscaChekc <> "" Then
        If Not Comprueba_CuentaBan2(BuscaChekc, False) Then
            If MsgBox("¿Continuar?", vbQuestion + vbYesNo) = vbNo Then B = False
        End If
    End If
        
    If B Then
        If BuscaChekc <> "" Then
            If DevuelveIBAN2("ES", BuscaChekc, CadenaConsulta) Then
                If Text1(19).Text = "" Then
                    Text1(19).Text = "ES" & CadenaConsulta
                Else
                    If Mid(Text1(19).Text, 3) <> CadenaConsulta Then
                        CadenaConsulta = "Calculado : " & "ES" & CadenaConsulta
                        CadenaConsulta = "Introducido: " & Me.Text1(19).Text & vbCrLf & CadenaConsulta & vbCrLf
                        CadenaConsulta = "Error en codigo IBAN" & vbCrLf & CadenaConsulta & "Continuar?"
                        If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then B = False
                    End If
                End If
            End If
        End If
    End If
          
          
    'Situacion de baja, y no tiene fecha de baja
    If B Then
        If Trim(Text1(13).Text) <> "" Xor Combo1(0).ListIndex > 0 Then
            If MsgBox("Fecha de baja deberia llevar situacion de baja." & vbCrLf & " ¿Continuar de igual modo?", vbQuestion + vbYesNoCancel) <> vbYes Then B = False
        End If
    End If
          
          
          
    If Text1(6).Text <> "" Then
        If Modo = 3 Then
            BuscaChekc = " TRUE "
        ElseIf Modo = 4 Then
            BuscaChekc = " idasoc <> " & Data1.Recordset!IdAsoc
        End If
        
        If BuscaChekc <> "" Then
            BuscaChekc = BuscaChekc & " AND nif "
            BuscaChekc = DevuelveDesdeBD(conAri, "concat(idasoc,' ',coalesce(nomlargo,'N/D'))", "asociados", BuscaChekc, Text1(6).Text, "T")
            If BuscaChekc <> "" Then
                BuscaChekc = "Ya tiene un socio ese NIF: " & BuscaChekc & vbCrLf & "¿Continuar?"
                If MsgBox(BuscaChekc, vbQuestion + vbYesNoCancel) <> vbYes Then B = False
            End If
            
        End If
    End If
          
          
          
          
          
          
          
          
    'Noviembre 2014
    '-------------------------------------
    ' Si va bien guardaremos en el hco de cambios de situacion
    ' si ha cambiado algo sobre: socio - alta - situacion - fechabaja
    
    If Modo = 4 And B Then
        'Meto la cadena para el insert
        NumRegElim = 0
        
        'Es socio
        If Check1(4).Value = 1 Then
            BuscaChekc = "SOCIO "
        Else
            BuscaChekc = "ASOCIADO "
        End If
        
        If Val(DBLet(Data1.Recordset!essocio, "N")) <> Abs(Val(Me.Check1(4).Value)) Then
            NumRegElim = 1
            BuscaChekc = BuscaChekc & "    [ Era "
            If Val(Data1.Recordset!essocio) = 1 Then
                BuscaChekc = BuscaChekc & "    Socio"
            Else
                BuscaChekc = BuscaChekc & "    Asociado"
            End If
            BuscaChekc = BuscaChekc & "]"

        End If
        InsertarEnCambiosSituacion = BuscaChekc & vbCrLf
        
        
       'Codactiv
        BuscaChekc = "Actividad(Ges): " & Text1(23).Text
        If Val(Data1.Recordset!tarifaprecio) <> Val(Text1(23).Text) Then
            NumRegElim = 1
            BuscaChekc = BuscaChekc & "    [ Tenia la " & CStr(Data1.Recordset!tarifaprecio) & "]"
        End If
        InsertarEnCambiosSituacion = InsertarEnCambiosSituacion & BuscaChekc & vbCrLf
        
        
        BuscaChekc = "F. alta: " & Text1(12).Text
        If DBLet(Data1.Recordset!fechaalta, "F") <> Text1(12).Text Then
            NumRegElim = NumRegElim + 2
            BuscaChekc = BuscaChekc & "    [ Antes era " & DBLet(Data1.Recordset!fechaalta, "F") & "]"
        End If
        InsertarEnCambiosSituacion = InsertarEnCambiosSituacion & BuscaChekc & vbCrLf
        
        
        BuscaChekc = "Situación: " & Me.Combo1(0).List(Combo1(0).ListIndex)
        If DBLet(Data1.Recordset!Estado, "N") <> Combo1(0).ListIndex Then
            NumRegElim = NumRegElim + 4
            BuscaChekc = BuscaChekc & "    [Antes " & SituacionAnterior & "]"
        End If
        InsertarEnCambiosSituacion = InsertarEnCambiosSituacion & BuscaChekc & vbCrLf
        
        
        
        
        
        BuscaChekc = "F. baja: " & Text1(13).Text
        If DBLet(Data1.Recordset!fechabaja, "F") <> Text1(13).Text Then
            NumRegElim = NumRegElim + 8
            BuscaChekc = BuscaChekc & "    [Antes era " & DBLet(Data1.Recordset!fechabaja, "T") & "]"
        End If
        InsertarEnCambiosSituacion = InsertarEnCambiosSituacion & BuscaChekc & vbCrLf
        
        'No hay cambios
        TituloLinea = ""
        miSQL = ""
        If NumRegElim = 0 Then
            InsertarEnCambiosSituacion = ""
        Else
            'Ha habido cambio de situacion.
            
            If DBLet(Data1.Recordset!Estado, "N") <> Combo1(0).ListIndex Then
                'Cambio de situacion
                
                If DBLet(Data1.Recordset!fechaalta, "F") <> Text1(12).Text Then
                    'hay fecha de alta
                    TituloLinea = "ALTA "
                    miSQL = Text1(12).Text
                
                
                ElseIf DBLet(Data1.Recordset!fechabaja, "F") <> Text1(13).Text Then
                    'Hay fecha de baja
                    TituloLinea = "BAJA "
                    miSQL = Text1(13).Text
                
                End If
                If TituloLinea <> "" Then
                    If Check1(4).Value = 0 Then
                        TituloLinea = TituloLinea & " ASOCIADO"
                    Else
                        TituloLinea = TituloLinea & " SOCIO"
                    End If
                End If
            
            Else
                'NO ha cambiado el ESTADO, pero ha cambiado fechas de alta o baja
                If DBLet(Data1.Recordset!fechaalta, "F") <> Text1(12).Text Then
                    'hay fecha de alta
                    TituloLinea = "ALTA "
                    miSQL = Text1(12).Text
                
                
                ElseIf DBLet(Data1.Recordset!fechabaja, "F") <> Text1(13).Text Then
                    'Hay fecha de baja
                    TituloLinea = "BAJA "
                    miSQL = Text1(13).Text
                
                End If
                If TituloLinea <> "" Then
                    If Check1(4).Value = 0 Then
                        TituloLinea = TituloLinea & " asociado*"
                    Else
                        TituloLinea = TituloLinea & " socio*"
                    End If
                    
                End If

            
            
            End If
            If TituloLinea = "" Then TituloLinea = "GENERICO"
        End If
    End If
          
    BuscaChekc = ""
    DatosOk = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SituacionAnterior() As String
    On Error Resume Next
    If IsNull(Data1.Recordset!Estado) Then
        SituacionAnterior = "No tenia"
    Else
        SituacionAnterior = Me.Combo1(0).List(Data1.Recordset!Estado)
        If Err.Number <> 0 Then
            SituacionAnterior = "Error leyendo anterior: " & Data1.Recordset!Estado
            Err.Clear
        End If
    End If
End Function


Private Function DatosOkLinea() As Boolean

On Error GoTo EDatosOkLinea

    DatosOkLinea = False


    Select Case Me.SSTab1.Tab
        Case 0 '
            miSQL = ""
            
            If Trim(txtAux1(0).Text) = "" Then
                miSQL = "Fecha alta no puede ser nula" & vbCrLf
                PonerFoco txtAux1(0)
            End If
            
            If ModificaLineas = 1 Then
                'ALTA
                If Me.cboSeccionGesoc.ListIndex < 0 Then
                    If miSQL = "" Then PonerFocoCbo cboSeccionGesoc
                    miSQL = miSQL & "Seleccione la seccion" & vbCrLf
                    
                Else
                    If cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex) = 1 Then
                        If cboEntidades.ListIndex < 0 Then
                            If miSQL = "" Then PonerFocoCbo cboEntidades
                            miSQL = miSQL & "Seleccione la entidad de gasolinera" & vbCrLf
                        End If
                    End If
                End If
            
            Else
                'Solo dejamos dar de baja
                'If Trim(txtAux1(0).Text) = "" Then
                '    miSQL = "Fecha alta no puede ser nula" & vbCrLf
                '    PonerFoco txtAux1(1)
                'End If
                
            End If
            
        Case 2 'Habilidades
            
            miSQL = "No desarrollado"
            
            
    End Select
    
    If miSQL <> "" Then
        miSQL = "Error en campos: " & vbCrLf & vbCrLf & miSQL
        MsgBox miSQL, vbExclamation
        Exit Function
    End If
    
    
    DatosOkLinea = True
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    PonerModo 5
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
            '++no ha hecho nada y lo hemos dejado en modo 5
            PonerModo 2
        Case Else
    End Select
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
            
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
            
'        Case 10  'Estudios/Formacion
'            mnEstudios_Click
'        Case 11  'Habilidades
'            mnHabilidades_Click
'        Case 12  'Experiencia Laboral
'            mnExperiencia_Click
'        Case 13 'Formacion Realizada
'            mnFormRealizada_Click
'        Case 14  'Formacion Empresa
'            mnFormEmpresa_Click
'
'        Case 16
'            'frmListado2.opcion = 17
'            'frmListado2.Show vbModal
'
'        Case 17    'Salir
'            mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento (Button.Index - btnPrimero)
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
   
    
Private Function InsertarLinea() As Boolean

On Error GoTo EInsertarLinea

    InsertarLinea = False
    
    If DatosOkLinea Then
          
          Select Case Me.SSTab1.Tab
             Case 0
                 
                 
                  If DarAltaUnidadNegocio_ Then InsertarLinea = True
            Case 2 'Habilidades
                    
          End Select
     End If
    
   
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
On Error GoTo EModificarLinea

    ModificarLinea = False

    If DatosOkLinea Then

         Select Case Me.SSTab1.Tab
            Case 0 'Secciones
                'Para dar de baja, lo otro no hacemos nada
                
            
                If Me.txtAux1(1).Text = "" Then
                    'NO ha puesto fecha de baja. NO hago NADA de nada
            
                Else
                
                End If
                ActualizarEnSecciones
                
            
            
            
        End Select
        
        ModificarLinea = True
    End If

  
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas" & vbCrLf & Err.Description
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
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid1(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String)
On Error GoTo ECargaGrid

    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    vDataGrid.RowHeight = 320
     
    CargaGrid2 vDataGrid, vData
    vDataGrid.Enabled = (Modo = 0 Or Modo = 2)
    vDataGrid.ScrollBars = dbgAutomatic
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False 'idasoc


    Select Case vDataGrid.Name
        Case "DataGrid1" 'Estudios / Formacion
                vDataGrid.Columns(1).Caption = "Seccion"
                vDataGrid.Columns(1).Width = 2500
                vDataGrid.Columns(2).Caption = "F. Alta"
                vDataGrid.Columns(2).Width = 1350
                vDataGrid.Columns(3).Caption = "F. Baja"
                vDataGrid.Columns(3).Width = 1350
                vDataGrid.Columns(4).Caption = "Colectivo"
                vDataGrid.Columns(4).Width = 5650
        Case "DataGrid2" 'Habilidades
                vDataGrid.Columns(1).Caption = "Fecha hco"
                vDataGrid.Columns(1).Width = 2200
                vDataGrid.Columns(1).visible = True
                vDataGrid.Columns(2).Caption = "Usuario"
                vDataGrid.Columns(2).Width = 1250
                vDataGrid.Columns(2).visible = True
                vDataGrid.Columns(3).Caption = "Tipo modificación"
                vDataGrid.Columns(3).Width = 2100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(4).Caption = "Fecha"
                vDataGrid.Columns(4).Width = 1300
                vDataGrid.Columns(4).visible = True
                vDataGrid.Columns(5).visible = False
                vDataGrid.Columns(6).visible = False
'        Case "DataGrid3" 'Experiencia Laboral
'                vDataGrid.Columns(2).Caption = "Periodo"
'                vDataGrid.Columns(2).Width = 2100
'                vDataGrid.Columns(3).visible = True
'                vDataGrid.Columns(3).Caption = "Experiencia"
'                vDataGrid.Columns(3).Width = 6450
'        Case "DataGrid4" 'Formacion Realizada
'                vDataGrid.Columns(2).Caption = "Fecha Formac."
'                vDataGrid.Columns(2).Width = 1450
'                vDataGrid.Columns(3).Caption = "Fecha Eval."
'                vDataGrid.Columns(3).Width = 1450
'                vDataGrid.Columns(4).Caption = "Formación"
'                vDataGrid.Columns(4).Width = 4000
'                vDataGrid.Columns(5).Caption = "Centro"
'                vDataGrid.Columns(5).Width = 1670
'                vDataGrid.Columns(6).Caption = "Evaluación"
'                vDataGrid.Columns(6).Width = 1160
'        Case "DataGrid5" 'Formacion Empresa
'                vDataGrid.Columns(2).Caption = "Fecha Formac."
'                vDataGrid.Columns(2).Width = 1500
'                vDataGrid.Columns(3).Caption = "Formación"
'                vDataGrid.Columns(3).Width = 4670
'                vDataGrid.Columns(4).Caption = "Resultado"
'                vDataGrid.Columns(4).Width = 1900
    End Select

    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.RowHeight = 350
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux1(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(i).Top = 290
            txtAux1(i).visible = visible
        Next i
        cboEntidades.visible = visible
        cboSeccionGesoc.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
                BloquearTxt txtAux1(i), False
            Next i
            cboEntidades.ListIndex = -1
            cboSeccionGesoc.ListIndex = -1
        Else
        
            'MODificar
            'Solo dejamos modificar la fecha de baja
            For i = 0 To 1
                txtAux1(i).Text = Trim(DataGrid1.Columns(i + 2).Text)
                BloquearTxt txtAux1(i), i = 0
            Next i
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 8)
        cboEntidades.Top = alto
        cboSeccionGesoc.Top = alto
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        cboSeccionGesoc.Left = DataGrid1.Left + 320
        cboSeccionGesoc.Width = DataGrid1.Columns(1).Width
        txtAux1(0).Left = DataGrid1.Columns(2).Left + DataGrid1.Left
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 20
        'Formacion
        txtAux1(1).Left = DataGrid1.Columns(3).Left + DataGrid1.Left
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 20
        
        cboEntidades.Left = DataGrid1.Columns(4).Left + DataGrid1.Left
        cboEntidades.Width = DataGrid1.Columns(4).Width
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
        cboEntidades.visible = False
        cboSeccionGesoc.visible = False
        If visible And limpiar Then cboSeccionGesoc.visible = True
    End If
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
            txtAux2.Top = 290
            txtAux2.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            txtAux2.Text = ""
            BloquearTxt txtAux2, False
        Else
            txtAux2.Text = DataGrid2.Columns(2).Text
            BloquearTxt txtAux2, False
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 8)
        
        txtAux2.Top = alto
        txtAux2.Height = DataGrid2.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Habilidad
        txtAux2.Left = DataGrid2.Left + 320
        txtAux2.Width = DataGrid2.Columns(2).Width - 20
            
        'Los ponemos Visibles o No
        '--------------------------
        txtAux2.visible = visible
    End If
End Sub






Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), Modo
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
      If Not (Index = 0 And KeyCode = 38) Then
            KEYdown KeyCode
      End If
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        'Me.SSTab1.Tab = numTab
        NumTabMto = numTab - 1
        TituloLinea = Cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Sub TxtAux1_LostFocus(Index As Integer)
    
    If txtAux1(Index).Text <> "" Then
        PonerFormatoFecha txtAux1(Index)
           
            'PonerFoco txtAux1(Index)
        
    End If

    If Index = 1 Then
        'If txtAux1(Index).Text <> "" Then
            If Not Me.cboEntidades.visible Then PonerFocoBtn Me.cmdAceptar
      '  End If
    End If
End Sub

Private Sub txtAux2_GotFocus()
    ConseguirFoco txtAux2, Modo
End Sub

Private Sub txtAux2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub



Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar

        conn.BeginTrans
        SQL = " WHERE  idasoc=" & Data1.Recordset!IdAsoc

        'Lineas Estudios/Formacion
        conn.Execute "Delete from asociados_hcocambios " & SQL
        conn.Execute "Delete from asociados " & SQL
        
        
        'Agosto 2020.  NO esta facelec en el server
        'conn.Execute "DELETE from facelec_ariadna.cliente where cod_gessoc = " & CStr(Data1.Recordset!IdAsoc)
        
        
        'Lineas Experiencia Laboral
'        conn.Execute "Delete from strab3 " & SQL
'        'Lineas Formacion Realizada
'        conn.Execute "Delete from strab4 " & SQL
'        'Lineas Experiencia Empresa
'        conn.Execute "Delete from strab5 " & SQL






        

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar" & Err.Description
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function



Private Sub LimpiarDataGrids()

On Error Resume Next
  
    
    CargaGrid1 DataGrid1, Data2, DevSQLGrid(1, False)
    
      
    BuscaChekc = "select IdAsoc,FechaCambio,usuario,tipoCambio,FechaCampo,situacion,observaciones from asociados_hcocambios where IdAsoc= -1"
    CargaGrid1 DataGrid2, Data3, BuscaChekc
    BuscaChekc = ""
    
    PrimeraVez = False
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function DevSQLGrid(Cual As Byte, enlaza As Boolean) As String
    If Cual = 1 Then
        DevSQLGrid = "select asociados_unidadesnegocio.idunidad,substring(nombre,20),"
        DevSQLGrid = DevSQLGrid & "asociados_unidadesnegocio.fechaalta,fechabaja "
        'Si tiene gasolinera, de momento es aLZIRA
        DevSQLGrid = DevSQLGrid & " , if(asociados_unidadesnegocio.idunidad=1,nomcoope,'')"
        DevSQLGrid = DevSQLGrid & " FROM (asociados_unidadesnegocio inner join unidadesnegocio on "
        DevSQLGrid = DevSQLGrid & "  asociados_unidadesnegocio.idunidad=unidadesnegocio.idunidad )"
        DevSQLGrid = DevSQLGrid & "  left join  asociados_entidades on asociados_entidades.IdAsoc= "
        DevSQLGrid = DevSQLGrid & "  asociados_unidadesnegocio.IdAsoc "
        'Si tiene gasolinera, de momento es aLZIRA
        DevSQLGrid = DevSQLGrid & "  left join " & BD_Arigaso_l & ".scoope  on asociados_entidades.identidad =codcoope"
        
        
        DevSQLGrid = DevSQLGrid & " where unidadesnegocio.idunidad and asociados_unidadesnegocio.idasoc="
        
        If enlaza Then
            DevSQLGrid = DevSQLGrid & Text1(0).Text
        Else
            DevSQLGrid = DevSQLGrid & "-1"
        End If
        DevSQLGrid = DevSQLGrid & " ORDER BY asociados_unidadesnegocio.idunidad"
        
    End If
    
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(idasoc=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       LimpiarDataGrids
       PonerModo 0
    End If
End Sub



Private Function PuedeEliminarTrabajador() As Boolean
    PuedeEliminarTrabajador = False
    
    If Me.Data2.Recordset.EOF Then PuedeEliminarTrabajador = True
    
    
    
End Function


Private Sub CargaCombos()


    Combo1(0).AddItem "Activo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Baja voluntaria"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Baja justificada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    
'-- Tipos de IVA (NN: Hay que cargar ahora de la conta que corresonda)
    Combo1(1).Clear
    
    Set miRsAux = New ADODB.Recordset
    
    CadenaConsulta = "Select ContIva from parametros"   'de gesscoial
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO puede ser EOF
    CadenaConsulta = CStr(miRsAux!contiva)
    miRsAux.Close
    
    CadenaConsulta = "SELECT * FROM " & IIf(vParamAplic.ContabilidadNueva, "ariconta", "conta") & CadenaConsulta & ".tiposiva"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1(1).AddItem miRsAux!nombriva
        Combo1(1).ItemData(Combo1(1).NewIndex) = miRsAux!Codigiva
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
'-- Tipos de IRPF
    Combo1(2).Clear
    Combo1(2).AddItem "MODULOS"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "EST. DIRECTA"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "ENTIDAD"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
    
    
    
    '#####
    'EntidadesArigasol
    cboEntidades.Clear
    
    miRsAux.Open "Select codcoope,nomcoope FROM " & BD_Arigaso_l & ".scoope ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboEntidades.AddItem miRsAux!Nomcoope
        cboEntidades.ItemData(cboEntidades.NewIndex) = miRsAux!codcoope
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    Set miRsAux = Nothing

    
    
End Sub



Private Sub ActualizarEnUnidadesDeNegocio()
Dim RN As ADODB.Recordset
Dim Donde As String

        On Error GoTo eActualizarEnUnidadesDeNegocio


        '-- actualiza cuentas contables
        Donde = "Cuentas contables"
        ActualizaCuentasAsociado Text1(0), 0, udNegocioGasol
        
         
        Set RN = New ADODB.Recordset
        CadenaConsulta = "select unidadesnegocio.*,asociados_unidadesnegocio.fechaalta from asociados_unidadesnegocio,unidadesnegocio where "
        CadenaConsulta = CadenaConsulta & " asociados_unidadesnegocio.IdUnidad= unidadesnegocio.idunidad and idasoc=" & Text1(0).Text
        
        RN.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RN.EOF
        
            Donde = "Unidad de negocio: " & RN!IdUnidad
            Select Case RN!IdUnidad
            Case 1
                '-- Actualizamos el asociado en la gasolinera
                'select * from asociados_entidades where IdAsoc IdEntidad FechaAlta
                
                
                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "identidad", "asociados_entidades", "idasoc", Text1(0).Text)
                If CadenaDesdeOtroForm = "" Then
                    MsgBox "Error obteniendo entidad de gasolinera", vbExclamation
                Else
                    ActGasolineraAsociadoSocio CLng(Text1(0)), CInt(CadenaDesdeOtroForm), RN!fechaalta, BD_Arigaso_l
                End If
            
            Case 2, 4
            
                '-- actualiza el cliente en Ariges
                TraspasaAsociadoAriges Text1(0), RN, CDate(Text1(12).Text)
            
            
                
            Case 3
                '-- Actualiza en el nuevo AriAgro
                ActualizaSocioAriagro CLng(Text1(0))
            
            End Select
            RN.MoveNext
        Wend
        RN.Close
        
        
eActualizarEnUnidadesDeNegocio:
    If Err.Number <> 0 Then
        Donde = Donde & vbCrLf & String(40, "=") & vbCrLf & vbCrLf & Err.Description
        conn.Errors.Clear
        Err.Clear
        MsgBox Donde, vbExclamation
    End If
        
    Set RN = Nothing
        
End Sub



Private Sub CargaUnidadesdeNegocio(Nuevas As Boolean)
    
    
    cboSeccionGesoc.Clear
    BuscaChekc = "select IdUnidad,substring(nombre,20) nombre from unidadesnegocio where "
    'Crear
    If Nuevas Then BuscaChekc = BuscaChekc & " NOT "
    BuscaChekc = BuscaChekc & " idunidad in (select idunidad from asociados_unidadesnegocio"
    BuscaChekc = BuscaChekc & " WHERE idasoc=" & Data1.Recordset!IdAsoc & ") ORDER By idUnidad"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboSeccionGesoc.AddItem miRsAux!Nombre
        cboSeccionGesoc.ItemData(cboSeccionGesoc.NewIndex) = miRsAux!IdUnidad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    BuscaChekc = ""
End Sub




Private Sub ActualizarEnSecciones()
Dim Ud As Byte
Dim SituBajaAgro As String
    
    
    If ModificaLineas = 2 Then
        'Modificando
        Ud = CByte(Data2.Recordset!IdUnidad)
    Else
        'INSERTANDO
        Ud = Me.cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex)
    End If
    
    
    If Me.txtAux1(1).Text = "" Then
        BuscaChekc = "null"  'en minuscula. Abajo se compara con el valor null
    Else
        BuscaChekc = DBSet(txtAux1(1).Text, "F")
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    Select Case Ud
    Case 1
        'GASOLINERA
        miSQL = "UPDATE " & BD_Arigaso_l & ".ssocio SET fechabaj=" & BuscaChekc
        miSQL = miSQL & " where codsocio = " & CStr(Data1.Recordset!IdAsoc)
        conn.Execute miSQL
        
        
         
    
    Case 2, 4
        'Suministros 2,  4 ARIGES gesocical
         
         miSQL = "select empresa_conta,codsituabaja  from unidadesnegocio where IdUnidad =" & CStr(Ud)
         miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         'NO puede ser eof
         miSQL = "UPDATE ariges" & miRsAux!empresa_conta & ".sclien SET codsitua="
         If BuscaChekc = "null" Then
            miSQL = miSQL & "0"
         Else
            miSQL = miSQL & miRsAux!codsituabaja
         End If
         miSQL = miSQL & " where codclien = " & CStr(Data1.Recordset!IdAsoc)
         miRsAux.Close
         
         conn.Execute miSQL
         
    Case 3
        'HORTOFRU
        miSQL = DBLet(Data1.Recordset!CodSocEuroagro, "T")
        If miSQL = "" Then
            MsgBox "Error en Codigo socio euroagro", vbExclamation
            Exit Sub
        Else
            SituBajaAgro = SituacionBajaAriagro(BuscaChekc <> "null")
            If SituBajaAgro <> "" Then SituBajaAgro = ", codsitua = " & SituBajaAgro & "   "
            miSQL = " WHERE codsocio = " & miSQL
            miSQL = "SET fechabaja = " & BuscaChekc & SituBajaAgro & miSQL
            miSQL = "UPDATE Ariagro.rsocios  " & miSQL
            conn.Execute miSQL
            
        End If
    
    End Select
    Set miRsAux = Nothing
    miSQL = "UPDATE asociados_unidadesnegocio SET fechabaja=" & BuscaChekc
    miSQL = miSQL & " where idasoc = " & CStr(Data1.Recordset!IdAsoc)
    miSQL = miSQL & " AND IdUnidad = " & Ud
    conn.Execute miSQL
    
End Sub

'Dara la situacion mas comun para la baja o el alta segun estemos haciendo
Private Function SituacionBajaAriagro(DarBaja As Boolean) As String
Dim C As String
Dim Aux As String
    
    'select codsitua,count(*) from rsocios where not fechabaja is null group by 1 order by 2 desc
    C = ""
    If DarBaja Then C = " NOT "
    C = C & " fechabaja is null AND 1"
    
    Aux = "count(*)"
    C = DevuelveDesdeBD(conAri, "codsitua", "ariagro.rsocios", C, "1  group by 1 order by 2 desc", "", Aux)

        
    SituacionBajaAriagro = C
    
End Function





Private Sub InsertaAsociadoEnAriges2(Insertar As Boolean)

    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If ModificaLineas = 2 Then
        'Modificando
        miSQL = CByte(Data2.Recordset!IdUnidad)
    Else
        'INSERTANDO
        miSQL = Me.cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex)
    End If
    
    miRsAux.Open "Select * from unidadesnegocio where IdUnidad = " & miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If TraspasaAsociadoAriges(Data1.Recordset!IdAsoc, miRsAux, Data1.Recordset!fechaalt) Then
        
        
        
        'Metemos el registro en la de uidades de negocio... si es crear
        If Insertar Then
            miRsAux.Close
            miSQL = "Select fechaalta,fechabaja FROM asociados where IdAsoc=" & Data1.Recordset!IdAsoc
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            miSQL = "REPLACE asociados_unidadesnegocio(IdAsoc,IdUnidad,FechaAlta) VALUES (" & Data1.Recordset!IdAsoc
            miSQL = miSQL & "," & Me.cboSeccionGesoc.ItemData(Me.cboSeccionGesoc.ListIndex) & ","
            miSQL = miSQL & DBSet(miRsAux!fechaalta, "F", "N") & ")"
        
        
            conn.Execute miSQL
            
            
            'Es nuevo en la seccion. mandamos a que meta en la conta
            'Actualizamos datos en contabilidad
            Espera 0.7
            ActualizaCuentasAsociado CLng(Data1.Recordset!IdAsoc), CByte(Me.cboSeccionGesoc.ItemData(Me.cboSeccionGesoc.ListIndex)), udNegocioGasol
            
            'QUitamos de aqui el dato
            Me.cboSeccionGesoc.RemoveItem Me.cboSeccionGesoc.ListIndex
        End If
        
       
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub



Private Sub CargaEntidadesArigasol()
    Set miRsAux = New ADODB.Recordset
    cboEntidades.Clear
    

    miRsAux.Open "Select codcoope,nomcoope FROM " & BD_Arigaso_l & ".scoope ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboEntidades.AddItem miRsAux!Nomcoope
        cboEntidades.ItemData(cboEntidades.NewIndex) = miRsAux!codcoope
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    miRsAux.Open "select IdEntidad from asociados_entidades where idasoc=" & Data1.Recordset!IdAsoc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        For kCampo = 0 To cboEntidades.ListCount - 1
            If cboEntidades.ItemData(kCampo) = Val(miRsAux!identidad) Then cboEntidades.ListIndex = kCampo
        Next
    End If
    miRsAux.Close
    Set miRsAux = Nothing
        
End Sub


Private Sub ActualizaSocioEnAriagro(Insertar)

    
    Screen.MousePointer = vbHourglass
    
    Set miRsAux = New ADODB.Recordset
    
    If ActualizaSocioAriagro(CLng(Data1.Recordset!IdAsoc)) Then
        
        
        
        'Metemos el registro en la de uidades de negocio... si es crear
        If Insertar Then
           
            miSQL = "Select fechaalta,fechabaja FROM asociados where IdAsoc=" & Data1.Recordset!IdAsoc
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            miSQL = "REPLACE asociados_unidadesnegocio(IdAsoc,IdUnidad,FechaAlta) VALUES (" & Data1.Recordset!IdAsoc
            miSQL = miSQL & "," & Me.cboSeccionGesoc.ItemData(Me.cboSeccionGesoc.ListIndex) & ","
            miSQL = miSQL & DBSet(miRsAux!fechaalta, "F", "N") & ")"
        
          
            conn.Execute miSQL
            
               
            'QUitamos de aqui el dato
            Me.cboSeccionGesoc.RemoveItem Me.cboSeccionGesoc.ListIndex
            
            
            miRsAux.Close
            
        End If
        
       
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub ActualizaSocioEnAriagasol(Insertar As Boolean)
Dim EsNuevo As Boolean
Dim Codmacta As String
Dim rUd As ADODB.Recordset
Dim UltimoNivel As Integer
Dim i As Integer

    If Me.cboEntidades.ListIndex < 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Set miRsAux = New ADODB.Recordset
    

    miSQL = DevuelveDesdeBD(conAri, "codsocio", BD_Arigaso_l & ".ssocio", "codsocio", Data1.Recordset!IdAsoc)
    EsNuevo = miSQL = ""
    
        
    'Pueden ser varias cuentas a actualizar
    'Si es nuevo meto su cuenta contable en ssocio
    If EsNuevo Then
        miSQL = " select unidadesnegocio.*,asociados.essocio from asociados,unidadesnegocio where "
        '                       socio                                       gasolinera
        miSQL = miSQL & " IdAsoc = " & Data1.Recordset!IdAsoc & " And unidadesnegocio.IdUnidad = 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Codmacta = ""
        If miRsAux.EOF Then
            MsgBox "Error obteniendo cuenta contable", vbExclamation
        Else
        
            miSQL = "Select * from " & IIf(vParamAplic.ContabilidadNueva, "ariconta", "conta") & miRsAux!empresa_conta & ".empresa"
            Set rUd = New ADODB.Recordset
            rUd.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOF
            i = rUd!numnivel
            UltimoNivel = rUd.Fields("numdigi" & CStr(i))
            rUd.Close
            Set rUd = Nothing
            If UltimoNivel = 0 Then
                MsgBox "Error obteniendo ultimo nivel", vbExclamation
            Else
                If miRsAux!raiz_cliente_socio <> "" And miRsAux!essocio = 1 Then
                    '
                     i = UltimoNivel - Len(miRsAux!raiz_cliente_socio)
                     Codmacta = String(CLng(i), "0")
                     Codmacta = miRsAux!raiz_cliente_socio & Format(miRsAux!CodSocEuroagro, Codmacta)
                     
                End If
                
                If miRsAux!raiz_cliente_asociado <> "" And miRsAux!essocio = 0 Then
                    i = UltimoNivel - Len(miRsAux!raiz_cliente_asociado)
                    Codmacta = String(CLng(i), "0")
                     
                    Codmacta = miRsAux!raiz_cliente_asociado & Format(Data1.Recordset!IdAsoc, Codmacta)
                    
                End If
            End If 'ultimo nivel
            
        End If
        miRsAux.Close
        
        If Codmacta <> "" Then
            miSQL = "UPDATE asociados set Codmacta='" & Codmacta & "' WHERE idasoc =" & Data1.Recordset!IdAsoc
            conn.Execute miSQL
            ConnConta.Execute "commit"
            Espera 1
        End If
    End If
    
    If ActGasolineraAsociadoSocio(CLng(Data1.Recordset!IdAsoc), Me.cboEntidades.ItemData(Me.cboEntidades.ListIndex), Now, BD_Arigaso_l) Then
        
        
        
        'Metemos el registro en la de uidades de negocio... si es crear
        If Insertar Then
           
            miSQL = "Select fechaalta,fechabaja FROM asociados where IdAsoc=" & Data1.Recordset!IdAsoc
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            miSQL = "REPLACE asociados_unidadesnegocio(IdAsoc,IdUnidad,FechaAlta) VALUES (" & Data1.Recordset!IdAsoc
            miSQL = miSQL & "," & Me.cboSeccionGesoc.ItemData(Me.cboSeccionGesoc.ListIndex) & ","
            miSQL = miSQL & DBSet(miRsAux!fechaalta, "F", "N") & ")"
        
          
            conn.Execute miSQL
            
            
            
            miSQL = "REPLACE asociados_entidades(IdAsoc,IdEntidad,FechaAlta) VALUES (" & Me.Data1.Recordset!IdAsoc
            miSQL = miSQL & "," & Me.cboEntidades.ItemData(Me.cboEntidades.ListIndex) & ","
            miSQL = miSQL & DBSet(miRsAux!fechaalta, "F", "N") & ")"
            conn.Execute miSQL
            
            
                            
               
            
            miRsAux.Close
            
            
                        
        End If
        'QUitamos de aqui el dato
        Me.cboSeccionGesoc.RemoveItem Me.cboSeccionGesoc.ListIndex
   
    
       
       'Es nuevo en la seccion. mandamos a que meta en la conta
        'Actualizamos datos en contabilidad
        If EsNuevo Then
            Espera 0.7
            ActualizaCuentasAsociado CLng(Data1.Recordset!IdAsoc), 1, udNegocioGasol
        End If
       
       
    End If
    
    Screen.MousePointer = vbDefault
End Sub






Private Function DarAltaUnidadNegocio_() As Boolean
Dim RN As ADODB.Recordset

    
    Select Case cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex)
    Case 1
        '-- Actualizamos el asociado en la gasolinera
        'select * from asociados_entidades where IdAsoc IdEntidad FechaAlta
        
        
        
        ActGasolineraAsociadoSocio CLng(Text1(0)), CInt(cboEntidades.ItemData(cboEntidades.ListIndex)), CDate(Me.txtAux1(0).Text), BD_Arigaso_l
        miSQL = "REPLACE asociados_entidades(IdAsoc,IdEntidad,FechaAlta) VALUES (" & CLng(Text1(0))
        miSQL = miSQL & "," & Me.cboEntidades.ItemData(Me.cboEntidades.ListIndex) & ","
        miSQL = miSQL & DBSet(CDate(Me.txtAux1(0).Text), "F", "N") & ")"
        conn.Execute miSQL
            
    
    Case 2, 4
        Set RN = New ADODB.Recordset
        
        
        RN.Open "Select * from unidadesnegocio where idunidad = " & cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        '-- actualiza el cliente en Ariges
        TraspasaAsociadoAriges Text1(0), RN, CDate(Me.txtAux1(0).Text)
        RN.Close
        Set RN = Nothing
    
        
    Case 3
        '-- Actualiza en el nuevo AriAgro
        ActualizaSocioAriagro CLng(Text1(0))
    
    End Select
    
    miSQL = "insert into asociados_unidadesnegocio(IdAsoc,IdUnidad,FechaAlta) VALUES (" & Text1(0).Text & ","
    miSQL = miSQL & cboSeccionGesoc.ItemData(cboSeccionGesoc.ListIndex) & "," & DBSet(txtAux1(0).Text, "F") & ")"
    ejecutar miSQL, False
     DarAltaUnidadNegocio_ = True
End Function
