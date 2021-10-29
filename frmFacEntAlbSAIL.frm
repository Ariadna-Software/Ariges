VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntAlbSAIL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C"
   ClientHeight    =   9090
   ClientLeft      =   -150
   ClientTop       =   345
   ClientWidth     =   15300
   Icon            =   "frmFacEntAlbSAIL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   8535
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   38
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14160
      TabIndex        =   99
      Top             =   8640
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12960
      TabIndex        =   98
      Top             =   8640
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   960
      Top             =   8520
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
      TabIndex        =   39
      Top             =   0
      Width           =   15300
      _ExtentX        =   26988
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
            Object.ToolTipText     =   "Cambiar TIPO albaran"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Albaran"
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
         Left            =   11880
         MaxLength       =   15
         TabIndex        =   144
         Text            =   "BASE IMP."
         Top             =   100
         Width           =   1490
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   13440
         MaxLength       =   15
         TabIndex        =   143
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   720
      Top             =   8520
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
      Height          =   7260
      Left            =   120
      TabIndex        =   41
      Top             =   1275
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   12806
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   7
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
      TabPicture(0)   =   "frmFacEntAlbSAIL.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(51)"
      Tab(0).Control(1)=   "Label1(35)"
      Tab(0).Control(2)=   "Label1(53)"
      Tab(0).Control(3)=   "imgBuscar2(12)"
      Tab(0).Control(4)=   "imgBuscar2(9)"
      Tab(0).Control(5)=   "Label1(54)"
      Tab(0).Control(6)=   "imgBuscar2(13)"
      Tab(0).Control(7)=   "Label1(55)"
      Tab(0).Control(8)=   "imgBuscar2(0)"
      Tab(0).Control(9)=   "Label1(56)"
      Tab(0).Control(10)=   "Label1(58)"
      Tab(0).Control(11)=   "imgObserva(1)"
      Tab(0).Control(12)=   "imgObserva(0)"
      Tab(0).Control(13)=   "DataGrid1"
      Tab(0).Control(14)=   "txtaux(0)"
      Tab(0).Control(15)=   "txtaux(1)"
      Tab(0).Control(16)=   "txtaux(3)"
      Tab(0).Control(17)=   "txtaux(4)"
      Tab(0).Control(18)=   "txtaux(6)"
      Tab(0).Control(19)=   "txtaux(7)"
      Tab(0).Control(20)=   "txtaux(8)"
      Tab(0).Control(21)=   "txtaux(2)"
      Tab(0).Control(22)=   "cmdAux(0)"
      Tab(0).Control(23)=   "cmdAux(1)"
      Tab(0).Control(24)=   "FrameCliente"
      Tab(0).Control(25)=   "txtaux(5)"
      Tab(0).Control(26)=   "txtaux(9)"
      Tab(0).Control(27)=   "txtaux(10)"
      Tab(0).Control(28)=   "txtaux(11)"
      Tab(0).Control(29)=   "Text2(9)"
      Tab(0).Control(30)=   "txtaux(16)"
      Tab(0).Control(31)=   "txtaux(12)"
      Tab(0).Control(32)=   "Text2(0)"
      Tab(0).Control(33)=   "Text2(13)"
      Tab(0).Control(34)=   "txtaux(14)"
      Tab(0).Control(35)=   "Text2(2)"
      Tab(0).Control(36)=   "txtaux(15)"
      Tab(0).Control(37)=   "txtaux(17)"
      Tab(0).Control(38)=   "txtaux(13)"
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntAlbSAIL.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(77)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(78)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(79)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(73)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(45)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(10)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(11)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(12)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "imgBuscar(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(9)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "imgBuscar(8)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(23)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(24)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "imgBuscar(9)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(43)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(61)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(72)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "imgFecha(44)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(80)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(81)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(82)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(83)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "imgGeolocalizacion"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "imgFirmaRecep"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Text1(39)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "chkDocArchi"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text1(38)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text1(18)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text1(19)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Text1(20)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Text1(21)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text1(22)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text1(23)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Text1(24)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Text1(26)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Text1(25)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Text1(2)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Text2(27)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Text1(27)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Text2(28)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Text1(28)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Text2(29)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Text1(29)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "FrameFacRec"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "FrameFactura"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Text1(34)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "chkFacturarKm"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Text1(43)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "cboTipopEnvio"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "FrameHco"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Text1(45)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Text1(46)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Text1(47)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Text1(48)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Text1(49)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).ControlCount=   57
      TabCaption(2)   =   "O.trab /ext"
      TabPicture(2)   =   "frmFacEntAlbSAIL.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTituloEst"
      Tab(2).Control(1)=   "Label3(0)"
      Tab(2).Control(2)=   "Label3(1)"
      Tab(2).Control(3)=   "Label3(7)"
      Tab(2).Control(4)=   "Label3(6)"
      Tab(2).Control(5)=   "Label3(5)"
      Tab(2).Control(6)=   "Label3(8)"
      Tab(2).Control(7)=   "txtEuler(6)"
      Tab(2).Control(8)=   "txtEuler(7)"
      Tab(2).Control(9)=   "FrameOT_Euler"
      Tab(2).Control(10)=   "txtEuler(9)"
      Tab(2).Control(11)=   "txtEuler(8)"
      Tab(2).Control(12)=   "txtEuler(10)"
      Tab(2).Control(13)=   "FrameOT_Taxco"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Reparaciones"
      TabPicture(3)   =   "frmFacEntAlbSAIL.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtEule_R(22)"
      Tab(3).Control(1)=   "cboRecepAgenClien"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).Control(3)=   "txtEule_R(21)"
      Tab(3).Control(4)=   "txtEule_R(20)"
      Tab(3).Control(5)=   "txtEule_R(2)"
      Tab(3).Control(6)=   "txtEule_R(1)"
      Tab(3).Control(7)=   "txtEule_R(0)"
      Tab(3).Control(8)=   "Frame4R"
      Tab(3).Control(9)=   "txtEule_R(4)"
      Tab(3).Control(10)=   "txtEule_R(3)"
      Tab(3).Control(11)=   "chkEuler(9)"
      Tab(3).Control(12)=   "chkEuler(8)"
      Tab(3).Control(13)=   "chkEuler(7)"
      Tab(3).Control(14)=   "chkEuler(6)"
      Tab(3).Control(15)=   "chkEuler(5)"
      Tab(3).Control(16)=   "chkEuler(4)"
      Tab(3).Control(17)=   "chkEuler(3)"
      Tab(3).Control(18)=   "chkEuler(2)"
      Tab(3).Control(19)=   "chkEuler(1)"
      Tab(3).Control(20)=   "chkEuler(0)"
      Tab(3).Control(21)=   "txtEule_R(15)"
      Tab(3).Control(22)=   "txtEule_R(16)"
      Tab(3).Control(23)=   "txtEule_R(14)"
      Tab(3).Control(24)=   "txtEule_R(13)"
      Tab(3).Control(25)=   "txtEule_R(12)"
      Tab(3).Control(26)=   "txtEule_R(9)"
      Tab(3).Control(27)=   "txtEule_R(10)"
      Tab(3).Control(28)=   "txtEule_R(8)"
      Tab(3).Control(29)=   "txtEule_R(6)"
      Tab(3).Control(30)=   "txtEule_R(5)"
      Tab(3).Control(31)=   "txtEule_R(7)"
      Tab(3).Control(32)=   "optEule_R(7)"
      Tab(3).Control(33)=   "optEule_R(6)"
      Tab(3).Control(34)=   "optEule_R(5)"
      Tab(3).Control(35)=   "optEule_R(4)"
      Tab(3).Control(36)=   "txtEule_R(19)"
      Tab(3).Control(37)=   "txtEule_R(18)"
      Tab(3).Control(38)=   "txtEule_R(17)"
      Tab(3).Control(39)=   "cboEulerUdR"
      Tab(3).Control(40)=   "txtEule_R(11)"
      Tab(3).Control(41)=   "Label3E(38)"
      Tab(3).Control(42)=   "Label3E(37)"
      Tab(3).Control(43)=   "Label3E(36)"
      Tab(3).Control(44)=   "Label3E(24)"
      Tab(3).Control(45)=   "Label3E(23)"
      Tab(3).Control(46)=   "Label3E(20)"
      Tab(3).Control(47)=   "Label3E(15)"
      Tab(3).Control(48)=   "Label3E(10)"
      Tab(3).Control(49)=   "Label3E(9)"
      Tab(3).Control(50)=   "Label3E(8)"
      Tab(3).Control(51)=   "Label3E(7)"
      Tab(3).Control(52)=   "Label3E(6)"
      Tab(3).Control(53)=   "Label3E(5)"
      Tab(3).Control(54)=   "Label3E(4)"
      Tab(3).Control(55)=   "Label3E(3)"
      Tab(3).Control(56)=   "Label3E(2)"
      Tab(3).Control(57)=   "Label3E(1)"
      Tab(3).Control(58)=   "Label3E(30)"
      Tab(3).Control(59)=   "Label3E(29)"
      Tab(3).Control(60)=   "Label3E(28)"
      Tab(3).Control(61)=   "Label3E(27)"
      Tab(3).Control(62)=   "Label3E(26)"
      Tab(3).Control(63)=   "Label3E(25)"
      Tab(3).Control(64)=   "Label3E(11)"
      Tab(3).Control(65)=   "Label3E(19)"
      Tab(3).Control(66)=   "Label3E(18)"
      Tab(3).Control(67)=   "Label3E(17)"
      Tab(3).Control(68)=   "Label3E(16)"
      Tab(3).Control(69)=   "Label3E(14)"
      Tab(3).Control(70)=   "Label3E(13)"
      Tab(3).Control(71)=   "Label3E(12)"
      Tab(3).Control(72)=   "Label3E(32)"
      Tab(3).Control(73)=   "Label3E(31)"
      Tab(3).Control(74)=   "Label3E(35)"
      Tab(3).Control(75)=   "Label3E(34)"
      Tab(3).Control(76)=   "Label3E(33)"
      Tab(3).Control(77)=   "Line4"
      Tab(3).Control(78)=   "Line3"
      Tab(3).Control(79)=   "Line5"
      Tab(3).ControlCount=   80
      TabCaption(4)   =   "Costes"
      TabPicture(4)   =   "frmFacEntAlbSAIL.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3E(22)"
      Tab(4).Control(1)=   "Label1(66)"
      Tab(4).Control(2)=   "Label1(67)"
      Tab(4).Control(3)=   "Label1(68)"
      Tab(4).Control(4)=   "Label1(69)"
      Tab(4).Control(5)=   "Label1(70)"
      Tab(4).Control(6)=   "Label1(71)"
      Tab(4).Control(7)=   "Line2"
      Tab(4).Control(8)=   "Line6"
      Tab(4).Control(9)=   "Label1(84)"
      Tab(4).Control(10)=   "Label1(85)"
      Tab(4).Control(11)=   "ListView2"
      Tab(4).Control(12)=   "cmdLineasCostes(1)"
      Tab(4).Control(13)=   "cmdLineasCostes(0)"
      Tab(4).Control(14)=   "cmdLineasCostes(2)"
      Tab(4).Control(15)=   "cmdLineasCostes(3)"
      Tab(4).Control(16)=   "cmdLineasCostes(4)"
      Tab(4).Control(17)=   "cmdLineasCostes(5)"
      Tab(4).ControlCount=   18
      TabCaption(5)   =   "Fichadas"
      TabPicture(5)   =   "frmFacEntAlbSAIL.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(62)"
      Tab(5).Control(1)=   "Label1(63)"
      Tab(5).Control(2)=   "Label1(64)"
      Tab(5).Control(3)=   "Label3E(0)"
      Tab(5).Control(4)=   "ListView1"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Lineas albaran"
      TabPicture(6)   =   "frmFacEntAlbSAIL.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdLineasImpresion(0)"
      Tab(6).Control(1)=   "cmdLineasImpresion(1)"
      Tab(6).Control(2)=   "cmdLineasImpresion(2)"
      Tab(6).Control(3)=   "cmdLineasImpresion(3)"
      Tab(6).Control(4)=   "lwEulerLineas"
      Tab(6).ControlCount=   5
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   5
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   370
         ToolTipText     =   "Importar lineas proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.Frame FrameOT_Taxco 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   355
         Top             =   1320
         Visible         =   0   'False
         Width           =   7095
         Begin VB.CommandButton cmdOrdenTallerTaxco 
            Height          =   375
            Left            =   6480
            Picture         =   "frmFacEntAlbSAIL.frx":00D0
            Style           =   1  'Graphical
            TabIndex        =   367
            ToolTipText     =   "Imprimir puntos"
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   7
            Left            =   2100
            TabIndex        =   240
            Text            =   "Text5"
            Top             =   5040
            Width           =   1335
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   6
            Left            =   2100
            TabIndex        =   239
            Text            =   "Text5"
            Top             =   3840
            Width           =   1815
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   5
            Left            =   2100
            TabIndex        =   238
            Text            =   "Text5"
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   4
            Left            =   2100
            TabIndex        =   237
            Text            =   "Text5"
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   3
            Left            =   2100
            TabIndex        =   236
            Text            =   "Text5"
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   2
            Left            =   2100
            TabIndex        =   235
            Text            =   "Text5"
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   234
            Text            =   "Text5"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtTaxco 
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   233
            Text            =   "Text5"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "LICENCIA"
            Height          =   195
            Index           =   30
            Left            =   720
            TabIndex        =   368
            Top             =   3420
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Orden taller taxco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   5040
            TabIndex        =   366
            Top             =   240
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label3 
            Caption         =   "Kms"
            Height          =   195
            Index           =   28
            Left            =   720
            TabIndex        =   365
            Top             =   5040
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Kilómetros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   240
            TabIndex        =   364
            Top             =   4680
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Modelo taximetro"
            Height          =   195
            Index           =   22
            Left            =   720
            TabIndex        =   363
            Top             =   3840
            Width           =   1305
         End
         Begin VB.Label Label3 
            Caption         =   "Neumáticos"
            Height          =   195
            Index           =   20
            Left            =   720
            TabIndex        =   362
            Top             =   2640
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Taxi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   361
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Motor"
            Height          =   195
            Index           =   18
            Left            =   720
            TabIndex        =   360
            Top             =   2160
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Marca/Modelo"
            Height          =   195
            Index           =   17
            Left            =   720
            TabIndex        =   359
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label Label3 
            Caption         =   "Bastidor"
            Height          =   195
            Index           =   15
            Left            =   720
            TabIndex        =   358
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Matrícula"
            Height          =   195
            Index           =   13
            Left            =   720
            TabIndex        =   357
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Datos vehiculo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   356
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.TextBox txtEuler 
         Height          =   315
         Index           =   10
         Left            =   -70080
         TabIndex        =   244
         Text            =   "Text4"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   22
         Left            =   -70920
         MaxLength       =   16
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   0
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   346
         ToolTipText     =   "Insertar "
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   1
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   345
         ToolTipText     =   "Modificar"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   2
         Left            =   -73800
         Style           =   1  'Graphical
         TabIndex        =   344
         ToolTipText     =   "Eliminar"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   3
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   343
         ToolTipText     =   "Imprimir factura lineas"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   4
         Left            =   -72120
         Style           =   1  'Graphical
         TabIndex        =   342
         ToolTipText     =   "Pedir proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   49
         Left            =   9120
         MaxLength       =   35
         TabIndex        =   340
         Tag             =   "latitud|N|S|||scaalb|longitud|#0.00000|N|"
         Text            =   "01/01/2019 18:50:89"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   48
         Left            =   7800
         MaxLength       =   35
         TabIndex        =   338
         Tag             =   "Fecha auxiliar  Albaran|N|S|||scaalb|latitud|#0.00000|N|"
         Text            =   "01/01/2019 18:50:89"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   47
         Left            =   12720
         MaxLength       =   35
         TabIndex        =   336
         Tag             =   "DNO entraga|T|S|||scaalb|dnirecep||N|"
         Text            =   "01/01/2019 18:50:89"
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   46
         Left            =   9720
         MaxLength       =   35
         TabIndex        =   334
         Tag             =   "R|T|S|||scaalb|perrecep|||"
         Text            =   "01/01/2019 18:50:89"
         Top             =   1440
         Width           =   2865
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   45
         Left            =   7800
         MaxLength       =   20
         TabIndex        =   331
         Tag             =   "Fecha entrega mercancia|FH|S|||scaalb|fechaent|dd/mm/yyyy hh:nn:ss|N|"
         Text            =   "01/01/2019 18:50:89"
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Frame FrameHco 
         Height          =   615
         Left            =   120
         TabIndex        =   145
         Top             =   6480
         Width           =   14895
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   33
            Left            =   11160
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   153
            Text            =   "Text2"
            Top             =   180
            Width           =   3645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   33
            Left            =   10440
            MaxLength       =   30
            TabIndex        =   152
            Text            =   "Text1"
            Top             =   180
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   150
            Text            =   "Text2"
            Top             =   180
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   4680
            MaxLength       =   30
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   180
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   147
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Eliminacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   75
            Left            =   120
            TabIndex        =   328
            Top             =   217
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   10200
            Picture         =   "frmFacEntAlbSAIL.frx":065A
            ToolTipText     =   "Buscar incidencia"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   9240
            TabIndex        =   154
            Top             =   210
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   4320
            Picture         =   "frmFacEntAlbSAIL.frx":075C
            ToolTipText     =   "Buscar trabajador"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   3480
            TabIndex        =   151
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha "
            Height          =   255
            Index           =   37
            Left            =   1560
            TabIndex        =   148
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   146
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cboTipopEnvio 
         Height          =   315
         ItemData        =   "frmFacEntAlbSAIL.frx":085E
         Left            =   12360
         List            =   "frmFacEntAlbSAIL.frx":086E
         Style           =   2  'Dropdown List
         TabIndex        =   325
         Tag             =   "Origen Datos|N|S|||scaalb|codinter|||"
         Top             =   675
         Width           =   2415
      End
      Begin VB.ComboBox cboRecepAgenClien 
         Height          =   315
         ItemData        =   "frmFacEntAlbSAIL.frx":089B
         Left            =   -69240
         List            =   "frmFacEntAlbSAIL.frx":08AB
         Style           =   2  'Dropdown List
         TabIndex        =   323
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   375
         Left            =   -63480
         TabIndex        =   320
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
         Begin VB.OptionButton optEule_R 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   322
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optEule_R 
            Caption         =   "Agencia"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   321
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtEuler 
         Height          =   315
         Index           =   8
         Left            =   -73560
         TabIndex        =   242
         Text            =   "Text4"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtEuler 
         Height          =   315
         Index           =   9
         Left            =   -71880
         TabIndex        =   243
         Text            =   "Text4"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   3
         Left            =   -70440
         Style           =   1  'Graphical
         TabIndex        =   312
         ToolTipText     =   "Imprimir costes"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   2
         Left            =   -72600
         Style           =   1  'Graphical
         TabIndex        =   311
         ToolTipText     =   "eliminar articulo"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   0
         Left            =   -73560
         Style           =   1  'Graphical
         TabIndex        =   310
         ToolTipText     =   "Insertar articulo"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   1
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   309
         ToolTipText     =   "Modificar articulo"
         Top             =   600
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4575
         Left            =   -74640
         TabIndex        =   299
         Top             =   960
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   8070
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Trab."
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tarea"
            Object.Width           =   1429
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Vehiculo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2195
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Tiempo"
            Object.Width           =   1850
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Horas"
            Object.Width           =   1850
         EndProperty
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   21
         Left            =   -72240
         MaxLength       =   16
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   20
         Left            =   -73680
         MaxLength       =   16
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   2
         Left            =   -62280
         MaxLength       =   16
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   1
         Left            =   -64680
         MaxLength       =   50
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   0
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Frame Frame4R 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   375
         Left            =   -73680
         TabIndex        =   262
         Top             =   1680
         Width           =   3015
         Begin VB.OptionButton optEule_R 
            Caption         =   "Pagados"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   62
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optEule_R 
            Caption         =   "Debidos"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   61
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label3E 
            Caption         =   "Portes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   0
            TabIndex        =   263
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   4
         Left            =   -65280
         MaxLength       =   50
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   3360
         Width           =   3975
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   3
         Left            =   -65280
         MaxLength       =   50
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   3000
         Width           =   3975
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   9
         Left            =   -66120
         TabIndex        =   76
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   8
         Left            =   -67680
         TabIndex        =   75
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   7
         Left            =   -68760
         TabIndex        =   74
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   6
         Left            =   -70080
         TabIndex        =   73
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   5
         Left            =   -71160
         TabIndex        =   72
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   4
         Left            =   -66120
         TabIndex        =   70
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   3
         Left            =   -67680
         TabIndex        =   69
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   2
         Left            =   -68760
         TabIndex        =   68
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   1
         Left            =   -70080
         TabIndex        =   67
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
         Height          =   255
         Index           =   0
         Left            =   -71160
         TabIndex        =   66
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   15
         Left            =   -66000
         MaxLength       =   50
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   5880
         Width           =   1575
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   16
         Left            =   -63000
         MaxLength       =   16
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   14
         Left            =   -66000
         MaxLength       =   50
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   5400
         Width           =   4695
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   13
         Left            =   -66000
         MaxLength       =   50
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   4920
         Width           =   4695
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   12
         Left            =   -66000
         MaxLength       =   50
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   4440
         Width           =   2175
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   9
         Left            =   -72960
         MaxLength       =   50
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   5880
         Width           =   1575
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   10
         Left            =   -69960
         MaxLength       =   50
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   8
         Left            =   -72960
         MaxLength       =   50
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   5400
         Width           =   4815
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   6
         Left            =   -69840
         MaxLength       =   50
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   5
         Left            =   -72960
         MaxLength       =   50
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   4440
         Width           =   2175
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   7
         Left            =   -72960
         MaxLength       =   50
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   4920
         Width           =   4815
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "V"
         Height          =   195
         Index           =   7
         Left            =   -71760
         TabIndex        =   86
         Top             =   6600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "Otro"
         Height          =   195
         Index           =   6
         Left            =   -71160
         TabIndex        =   87
         Top             =   6600
         Width           =   615
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "N"
         Height          =   195
         Index           =   5
         Left            =   -72960
         TabIndex        =   84
         Top             =   6600
         Width           =   615
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "C"
         Height          =   195
         Index           =   4
         Left            =   -72360
         TabIndex        =   85
         Top             =   6600
         Width           =   615
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   19
         Left            =   -62160
         MaxLength       =   16
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   18
         Left            =   -64080
         MaxLength       =   16
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   17
         Left            =   -66000
         MaxLength       =   16
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   6360
         Width           =   855
      End
      Begin VB.ComboBox cboEulerUdR 
         Height          =   315
         ItemData        =   "frmFacEntAlbSAIL.frx":08D8
         Left            =   -68880
         List            =   "frmFacEntAlbSAIL.frx":08E5
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   6480
         Width           =   735
      End
      Begin VB.TextBox txtEule_R 
         Height          =   315
         Index           =   11
         Left            =   -69840
         MaxLength       =   16
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   6480
         Width           =   855
      End
      Begin VB.Frame FrameOT_Euler 
         Height          =   5535
         Left            =   -74640
         TabIndex        =   247
         Top             =   1320
         Width           =   6975
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   227
            Text            =   "Text4"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   5
            Left            =   1320
            TabIndex        =   232
            Text            =   "Text1"
            Top             =   4560
            Width           =   4815
         End
         Begin VB.TextBox txtEuler 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   5040
            TabIndex        =   228
            Text            =   "Text1"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   375
            Left            =   240
            TabIndex        =   255
            Top             =   360
            Width           =   4935
            Begin VB.OptionButton optEuler 
               Caption         =   "Debidos"
               Height          =   195
               Index           =   0
               Left            =   1080
               TabIndex        =   246
               Top             =   0
               Width           =   975
            End
            Begin VB.OptionButton optEuler 
               Caption         =   "Pagados"
               Height          =   195
               Index           =   1
               Left            =   2400
               TabIndex        =   248
               Top             =   0
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Portes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   0
               TabIndex        =   256
               Top             =   0
               Width           =   1935
            End
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   230
            Text            =   "Text1"
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   229
            Text            =   "Text1"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   4
            Left            =   1320
            TabIndex        =   231
            Text            =   "Text1"
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha"
            Height          =   195
            Index           =   4
            Left            =   4080
            TabIndex        =   261
            Top             =   1380
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   260
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Referencia"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   259
            Top             =   1380
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Marca"
            Height          =   195
            Index           =   12
            Left            =   480
            TabIndex        =   254
            Top             =   2640
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Modelo"
            Height          =   195
            Index           =   14
            Left            =   480
            TabIndex        =   253
            Top             =   3120
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Bombas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   252
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Motor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   251
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Modelo"
            Height          =   195
            Index           =   26
            Left            =   480
            TabIndex        =   250
            Top             =   4560
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Marca"
            Height          =   195
            Index           =   27
            Left            =   480
            TabIndex        =   249
            Top             =   4080
            Width           =   705
         End
      End
      Begin VB.TextBox txtEuler 
         Height          =   4995
         Index           =   7
         Left            =   -67200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   241
         Text            =   "frmFacEntAlbSAIL.frx":08F9
         Top             =   1560
         Width           =   6975
      End
      Begin VB.TextBox txtEuler 
         Height          =   4995
         Index           =   6
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   245
         Text            =   "frmFacEntAlbSAIL.frx":08FF
         Top             =   1560
         Width           =   6975
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   43
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "ObsCRM|T|S|||scaalb|observacrm|||"
         Top             =   5760
         Width           =   7125
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   13
         Left            =   -63960
         MaxLength       =   6
         TabIndex        =   55
         Text            =   "codc"
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox txtaux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   17
         Left            =   -62520
         TabIndex        =   220
         Text            =   "codc"
         Top             =   6600
         Width           =   2175
      End
      Begin VB.CheckBox chkFacturarKm 
         Caption         =   "Facturar Km"
         Height          =   375
         Left            =   3960
         TabIndex        =   216
         Tag             =   "Facturar Km|N|N|||scaalb|facturkm||N|"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   215
         Tag             =   "Cant. Km|N|S|0|99999|scaalb|cantidkm||N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   15
         Left            =   -63960
         MaxLength       =   12
         TabIndex        =   57
         Text            =   "codc"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -63360
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   213
         Text            =   "nom capit"
         Top             =   5880
         Width           =   3165
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   14
         Left            =   -63960
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "codc"
         Top             =   5880
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   13
         Left            =   -63360
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   211
         Text            =   "nom capit"
         Top             =   5160
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -63360
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   209
         Text            =   "nom capit"
         Top             =   4440
         Width           =   3165
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   12
         Left            =   -63960
         MaxLength       =   6
         TabIndex        =   54
         Text            =   "codc"
         Top             =   4440
         Width           =   615
      End
      Begin VB.Frame FrameFactura 
         Height          =   3060
         Left            =   120
         TabIndex        =   169
         Top             =   3240
         Width           =   7455
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
            Left            =   5400
            MaxLength       =   15
            TabIndex        =   192
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   191
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   190
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   189
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   188
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   187
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   186
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   185
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   184
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   183
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   182
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   181
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   180
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   179
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
            TabIndex        =   178
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
            TabIndex        =   177
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
            TabIndex        =   176
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   175
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   174
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   173
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   172
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   171
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   170
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   74
            Left            =   240
            TabIndex        =   327
            Top             =   0
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
            Height          =   255
            Index           =   42
            Left            =   1320
            TabIndex        =   207
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   3360
            TabIndex        =   206
            Top             =   1200
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
            ForeColor       =   &H00008000&
            Height          =   285
            Index           =   39
            Left            =   3480
            TabIndex        =   205
            Top             =   2655
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
            TabIndex        =   204
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   1320
            X2              =   7320
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   4080
            TabIndex        =   203
            Top             =   1200
            Width           =   735
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
            TabIndex        =   202
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
            TabIndex        =   201
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
            TabIndex        =   200
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   199
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   18
            Left            =   3960
            TabIndex        =   198
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   22
            Left            =   2160
            TabIndex        =   197
            Top             =   360
            Width           =   855
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
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   28
            Left            =   2040
            TabIndex        =   195
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   6
            Left            =   6240
            TabIndex        =   194
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   5400
            TabIndex        =   193
            Top             =   1200
            Width           =   495
         End
      End
      Begin VB.TextBox txtaux 
         BackColor       =   &H80000018&
         Height          =   675
         Index           =   16
         Left            =   -63960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Text            =   "frmFacEntAlbSAIL.frx":0905
         Top             =   3360
         Width           =   3765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   -63240
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   166
         Text            =   "nom ccoste"
         Top             =   2640
         Width           =   2925
      End
      Begin VB.TextBox txtaux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   -61800
         MaxLength       =   15
         TabIndex        =   164
         Text            =   "numlote"
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   -61200
         MaxLength       =   5
         TabIndex        =   47
         Text            =   "bulto"
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   -63960
         MaxLength       =   6
         TabIndex        =   52
         Text            =   "codc"
         Top             =   2640
         Width           =   735
      End
      Begin VB.Frame FrameFacRec 
         Height          =   615
         Left            =   120
         TabIndex        =   155
         Top             =   2520
         Width           =   7455
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   37
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   160
            Tag             =   "Tipo Mov. Factura|T|S|||scaalb|codtipmf||N|"
            Text            =   "MMM"
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   36
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   158
            Tag             =   "Nº. Factura|N|S|0||scaalb|numfactu|0000000|N|"
            Text            =   "4000000000"
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   35
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   156
            Tag             =   "Fecha Factura|F|S|||scaalb|fecfactu|dd/mm/yyyy|N|"
            Text            =   "99/99/9999"
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Factura a rectificar "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   76
            Left            =   120
            TabIndex        =   329
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   47
            Left            =   1920
            TabIndex        =   161
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
            Height          =   195
            Index           =   46
            Left            =   3000
            TabIndex        =   159
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fact."
            Height          =   195
            Index           =   44
            Left            =   5280
            TabIndex        =   157
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   29
         Left            =   7800
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Cod. Envío|N|N|0|999|scaalb|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   134
         Text            =   "Text2"
         Top             =   675
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   240
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Preparador Material|N|N|0|9999|scaalb|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   2160
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   132
         Text            =   "Text2"
         Top             =   2160
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   240
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "Trabajador pedido|N|S|0|9999|scaalb|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   130
         Text            =   "Text2"
         Top             =   1440
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   128
         Tag             =   "Semana Entrega|N|S|||scaalb|sementre||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   240
         MaxLength       =   7
         TabIndex        =   125
         Tag             =   "Nº Pedido|N|S|||scaalb|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   124
         Tag             =   "Fecha Pedido|F|S|||scaalb|fecpedcl|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   120
         Tag             =   "Fecha Oferta|F|S|||scaalb|fecofert|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   119
         Tag             =   "Nº Oferta|N|S|||scaalb|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   49
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   1950
         Left            =   -74800
         TabIndex        =   104
         Top             =   315
         Width           =   14580
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   40
            Left            =   9840
            MaxLength       =   7
            TabIndex        =   21
            Tag             =   "Aportacion|N|S|||scaalb|aportacion|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1020
         End
         Begin VB.ComboBox cboTipoImpr 
            Height          =   315
            ItemData        =   "frmFacEntAlbSAIL.frx":0942
            Left            =   12600
            List            =   "frmFacEntAlbSAIL.frx":0955
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "TipoImp|N|S|||scaalb|tipoimp|||"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.ComboBox cboTipoDat 
            Height          =   315
            ItemData        =   "frmFacEntAlbSAIL.frx":098D
            Left            =   12600
            List            =   "frmFacEntAlbSAIL.frx":099A
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Tag             =   "Origen Datos|N|S|||scaalb|origdat|||"
            Top             =   1140
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   41
            Left            =   12600
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Fecha envio|F|S|||scaalb|fecenvio|dd/mm/yyyy|N|"
            Top             =   720
            Width           =   1185
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   1
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   218
            Text            =   "Text2"
            Top             =   513
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   42
            Left            =   6885
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "Actuacion|T|S|||scaalb|actuacion|||"
            Text            =   "Text1"
            Top             =   513
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1125
            MaxLength       =   255
            TabIndex        =   14
            Tag             =   "Referencia Cliente|T|S|||scaalb|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   4125
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   116
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
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaalb|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   13
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   17
            Tag             =   "Cod. Agente|N|N|0|9999|scaalb|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   861
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   111
            Text            =   "Text2"
            Top             =   861
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "Forma de Pago|N|N|0|999|scaalb|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1215
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   106
            Text            =   "Text2"
            Top             =   1215
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   19
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
            TabIndex        =   20
            Tag             =   "Descuento General|N|N|0|99.90|scaalb|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   540
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   12600
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   10
            Tag             =   "Domicilio|T|N|||scaalb|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   513
            Width           =   4030
         End
         Begin VB.Label Label1 
            Caption         =   "Tasas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   49
            Left            =   9240
            TabIndex        =   369
            Top             =   1620
            Width           =   555
         End
         Begin VB.Image imgNull 
            Height          =   240
            Index           =   0
            Left            =   14160
            Picture         =   "frmFacEntAlbSAIL.frx":09B6
            ToolTipText     =   "Limpiar campo"
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNull 
            Height          =   240
            Index           =   1
            Left            =   14160
            Picture         =   "frmFacEntAlbSAIL.frx":7208
            ToolTipText     =   "Limpiar campo"
            Top             =   1560
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Origen datos"
            Height          =   195
            Index           =   60
            Left            =   11160
            TabIndex        =   224
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo impresion"
            Height          =   195
            Index           =   59
            Left            =   11160
            TabIndex        =   223
            Top             =   1620
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha liq."
            Height          =   195
            Index           =   52
            Left            =   11160
            TabIndex        =   222
            Top             =   720
            Width           =   975
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   40
            Left            =   12240
            Picture         =   "frmFacEntAlbSAIL.frx":DA5A
            ToolTipText     =   "Buscar fecha"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   6600
            Picture         =   "frmFacEntAlbSAIL.frx":DAE5
            ToolTipText     =   "Buscar forma de pago"
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Actuacion"
            Height          =   255
            Index           =   57
            Left            =   5580
            TabIndex        =   219
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   123
            Top             =   1590
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   855
            Picture         =   "frmFacEntAlbSAIL.frx":DBE7
            ToolTipText     =   "Buscar población"
            Top             =   867
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   5580
            TabIndex        =   118
            Top             =   165
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            Picture         =   "frmFacEntAlbSAIL.frx":DCE9
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   165
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   117
            Top             =   1209
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   115
            Top             =   861
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   114
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   113
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            Picture         =   "frmFacEntAlbSAIL.frx":DDEB
            ToolTipText     =   "Buscar cliente varios"
            Top             =   165
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5580
            TabIndex        =   112
            Top             =   870
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            Picture         =   "frmFacEntAlbSAIL.frx":DEED
            ToolTipText     =   "Buscar agente"
            Top             =   870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5580
            TabIndex        =   110
            Top             =   1215
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.Pago"
            Height          =   255
            Index           =   25
            Left            =   5580
            TabIndex        =   109
            Top             =   1590
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   7560
            TabIndex        =   108
            Top             =   1590
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   11160
            TabIndex        =   107
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            Picture         =   "frmFacEntAlbSAIL.frx":DFEF
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1230
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   105
            Top             =   513
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   -72360
         TabIndex        =   103
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
         Left            =   -74040
         TabIndex        =   102
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   45
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   -65760
         MaxLength       =   12
         TabIndex        =   100
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -67320
         MaxLength       =   30
         TabIndex        =   51
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -67800
         MaxLength       =   5
         TabIndex        =   50
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -66840
         MaxLength       =   12
         TabIndex        =   48
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -68880
         MaxLength       =   16
         TabIndex        =   46
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73800
         MaxLength       =   18
         TabIndex        =   44
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   15
         TabIndex        =   43
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   34
         Tag             =   "Observación 5|T|S|||scaalb|observa05||N|"
         Top             =   5040
         Width           =   7125
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   33
         Tag             =   "Observación 4|T|S|||scaalb|observa04||N|"
         Top             =   4680
         Width           =   7125
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación 3|T|S|||scaalb|observa03||N|"
         Top             =   4320
         Width           =   7125
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observación 2|T|S|||scaalb|observa02||N|"
         Top             =   3960
         Width           =   7125
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   7800
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 1|T|S|||scaalb|observa01||N|"
         Text            =   "ABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJ"
         Top             =   3600
         Width           =   7125
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntAlbSAIL.frx":E0F1
         Height          =   4680
         Left            =   -74805
         TabIndex        =   101
         Top             =   2325
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   8255
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
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   162
         Tag             =   "Nº terminal|N|S|||scaalb|numtermi||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.CheckBox chkDocArchi 
         Caption         =   "Documento archivado"
         Height          =   375
         Left            =   9480
         TabIndex        =   35
         Tag             =   "Docar|N|N|||scaalb|docarchiv||N|"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   39
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   163
         Tag             =   "Nº Venta|N|S|||scaalb|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   304
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9128
         SortKey         =   8
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5503
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Precio"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "codartic"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lwEulerLineas 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   347
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   8916
         SortKey         =   5
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Articulo"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   11642
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Descuento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "descripcionReal"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   85
         Left            =   -69240
         TabIndex        =   372
         Top             =   6720
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   84
         Left            =   -71160
         TabIndex        =   371
         Top             =   6720
         Width           =   1860
      End
      Begin VB.Line Line6 
         X1              =   -70920
         X2              =   -70920
         Y1              =   600
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   -71520
         X2              =   -71520
         Y1              =   600
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Venta"
         Height          =   195
         Index           =   8
         Left            =   -70080
         TabIndex        =   349
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label3E 
         Caption         =   "Alb. venta"
         Height          =   195
         Index           =   38
         Left            =   -70920
         TabIndex        =   348
         Top             =   960
         Width           =   1005
      End
      Begin VB.Image imgFirmaRecep 
         Height          =   1095
         Left            =   10680
         Picture         =   "frmFacEntAlbSAIL.frx":E106
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image imgGeolocalizacion 
         Height          =   255
         Left            =   8760
         Picture         =   "frmFacEntAlbSAIL.frx":11A29
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Longitud"
         Height          =   195
         Index           =   83
         Left            =   9120
         TabIndex        =   341
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Latitud"
         Height          =   195
         Index           =   82
         Left            =   7800
         TabIndex        =   339
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "DNI"
         Height          =   195
         Index           =   81
         Left            =   12720
         TabIndex        =   337
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Persona recepcion"
         Height          =   195
         Index           =   80
         Left            =   9720
         TabIndex        =   335
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   44
         Left            =   8280
         Picture         =   "frmFacEntAlbSAIL.frx":11FB3
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo envio"
         Height          =   255
         Index           =   72
         Left            =   12360
         TabIndex        =   324
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Albaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   -74640
         TabIndex        =   319
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Reparación"
         Height          =   195
         Index           =   6
         Left            =   -73560
         TabIndex        =   318
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Tr. exterior"
         Height          =   195
         Index           =   7
         Left            =   -71880
         TabIndex        =   317
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   71
         Left            =   -65400
         TabIndex        =   316
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   70
         Left            =   -67200
         TabIndex        =   315
         Top             =   6720
         Width           =   2100
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   69
         Left            =   -73320
         TabIndex        =   314
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   68
         Left            =   -74640
         TabIndex        =   313
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   67
         Left            =   -63360
         TabIndex        =   308
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   66
         Left            =   -61920
         TabIndex        =   307
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label3E 
         Caption         =   "Costes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   22
         Left            =   -74760
         TabIndex        =   306
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label3E 
         Caption         =   "Fichadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   305
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   -63960
         TabIndex        =   302
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   -62160
         TabIndex        =   301
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Horas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   -65640
         TabIndex        =   300
         Top             =   5640
         Width           =   975
      End
      Begin VB.Image imgObserva 
         Height          =   255
         Index           =   0
         Left            =   -62640
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label3E 
         Caption         =   "T. Externo"
         Height          =   195
         Index           =   37
         Left            =   -72240
         TabIndex        =   298
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label3E 
         AutoSize        =   -1  'True
         Caption         =   "Orden de trabajo"
         Height          =   195
         Index           =   36
         Left            =   -73680
         TabIndex        =   297
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3E 
         Caption         =   "F. Recepcion"
         Height          =   195
         Index           =   24
         Left            =   -62280
         TabIndex        =   296
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label3E 
         Caption         =   "Nº Expedicion"
         Height          =   195
         Index           =   23
         Left            =   -64680
         TabIndex        =   295
         Top             =   1320
         Width           =   2025
      End
      Begin VB.Label Label3E 
         Caption         =   "Recepcion del equipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   20
         Left            =   -73680
         TabIndex        =   294
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3E 
         Caption         =   "Otros equipos / tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   -64680
         TabIndex        =   293
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label3E 
         AutoSize        =   -1  'True
         Caption         =   "Vertical"
         Height          =   195
         Index           =   10
         Left            =   -67800
         TabIndex        =   292
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label3E 
         AutoSize        =   -1  'True
         Caption         =   "Pozo"
         Height          =   195
         Index           =   9
         Left            =   -68880
         TabIndex        =   291
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label3E 
         Caption         =   "Vertical"
         Height          =   195
         Index           =   8
         Left            =   -70200
         TabIndex        =   290
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label3E 
         Caption         =   "Horizontal"
         Height          =   195
         Index           =   7
         Left            =   -71400
         TabIndex        =   289
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Agitador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   -66360
         TabIndex        =   288
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3E 
         Caption         =   "Bombas sumegibles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   -68880
         TabIndex        =   287
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3E 
         Caption         =   "Bombas superficie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -71280
         TabIndex        =   286
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label3E 
         Caption         =   "Aguas limpias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   -73800
         TabIndex        =   285
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label3E 
         Caption         =   "Aguas residuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   284
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label3E 
         Caption         =   "Tipo de bombas recepcionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Index           =   1
         Left            =   -73800
         TabIndex        =   283
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label Label3E 
         Caption         =   "V"
         Height          =   195
         Index           =   30
         Left            =   -66840
         TabIndex        =   282
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "I (A)"
         Height          =   195
         Index           =   29
         Left            =   -63600
         TabIndex        =   281
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   28
         Left            =   -66840
         TabIndex        =   280
         Top             =   5400
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Marca"
         Height          =   195
         Index           =   27
         Left            =   -66840
         TabIndex        =   279
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   26
         Left            =   -66840
         TabIndex        =   278
         Top             =   4920
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Motor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   -65040
         TabIndex        =   277
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label3E 
         Caption         =   "Datos equipo / bomba recepcionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   11
         Left            =   -73800
         TabIndex        =   276
         Top             =   3720
         Width           =   4095
      End
      Begin VB.Label Label3E 
         Caption         =   "Año"
         Height          =   195
         Index           =   19
         Left            =   -73800
         TabIndex        =   275
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "H (m.c.a)"
         Height          =   195
         Index           =   18
         Left            =   -70920
         TabIndex        =   274
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   17
         Left            =   -73800
         TabIndex        =   273
         Top             =   5400
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Bombas(Parte hidraulica)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   -72000
         TabIndex        =   272
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label3E 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   14
         Left            =   -73800
         TabIndex        =   271
         Top             =   4920
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "NºCurva"
         Height          =   195
         Index           =   13
         Left            =   -70680
         TabIndex        =   270
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Marca"
         Height          =   195
         Index           =   12
         Left            =   -73800
         TabIndex        =   269
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Caudal"
         Height          =   195
         Index           =   32
         Left            =   -70440
         TabIndex        =   268
         Top             =   6360
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Tipo de rodete"
         Height          =   195
         Index           =   31
         Left            =   -73800
         TabIndex        =   267
         Top             =   6360
         Width           =   1035
      End
      Begin VB.Label Label3E 
         Caption         =   "RPM"
         Height          =   195
         Index           =   35
         Left            =   -63000
         TabIndex        =   266
         Top             =   6360
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Pot (Kw)"
         Height          =   195
         Index           =   34
         Left            =   -64800
         TabIndex        =   265
         Top             =   6360
         Width           =   705
      End
      Begin VB.Label Label3E 
         Caption         =   "Pot(CV)"
         Height          =   195
         Index           =   33
         Left            =   -66840
         TabIndex        =   264
         Top             =   6360
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -67200
         TabIndex        =   258
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Notas operario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   257
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblTituloEst 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   -66120
         TabIndex        =   226
         Top             =   360
         Width           =   5850
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones CRM"
         Height          =   255
         Index           =   61
         Left            =   7800
         TabIndex        =   225
         Top             =   5520
         Width           =   3975
      End
      Begin VB.Image imgObserva 
         Height          =   255
         Index           =   1
         Left            =   -61320
         Top             =   6360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   58
         Left            =   -62520
         TabIndex        =   221
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Km a facturar"
         Height          =   255
         Index           =   43
         Left            =   4920
         TabIndex        =   217
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P. Coste"
         Height          =   255
         Index           =   56
         Left            =   -63960
         TabIndex        =   214
         Top             =   6360
         Width           =   975
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   0
         Left            =   -63120
         Picture         =   "frmFacEntAlbSAIL.frx":1203E
         ToolTipText     =   "Buscar forma de pago"
         Top             =   5640
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   55
         Left            =   -63960
         TabIndex        =   212
         Top             =   5640
         Width           =   975
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   13
         Left            =   -63600
         Picture         =   "frmFacEntAlbSAIL.frx":12140
         ToolTipText     =   "Buscar forma de pago"
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "T.O."
         Height          =   255
         Index           =   54
         Left            =   -63960
         TabIndex        =   210
         Top             =   4920
         Width           =   975
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   9
         Left            =   -63000
         Picture         =   "frmFacEntAlbSAIL.frx":12242
         ToolTipText     =   "Buscar forma de pago"
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   12
         Left            =   -63240
         Picture         =   "frmFacEntAlbSAIL.frx":12344
         ToolTipText     =   "Buscar forma de pago"
         Top             =   4185
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Capitulo"
         Height          =   255
         Index           =   53
         Left            =   -63960
         TabIndex        =   208
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   -63960
         TabIndex        =   168
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Centro coste"
         Height          =   255
         Index           =   51
         Left            =   -63960
         TabIndex        =   167
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   8400
         Picture         =   "frmFacEntAlbSAIL.frx":12446
         ToolTipText     =   "Buscar forma de envio"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Envío merc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   6360
         TabIndex        =   135
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Preparador Material"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   133
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1800
         Picture         =   "frmFacEntAlbSAIL.frx":12548
         ToolTipText     =   "Buscar trabajador"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   131
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1800
         Picture         =   "frmFacEntAlbSAIL.frx":1264A
         ToolTipText     =   "Buscar trabajador"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   129
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   127
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   126
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   122
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   121
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   7800
         TabIndex        =   42
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   -73560
         X2              =   -61200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         X1              =   -73560
         X2              =   -61320
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   -73680
         X2              =   -61200
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo envio: codinter  Puede ser NULL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   73
         Left            =   8880
         TabIndex        =   326
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   79
         Left            =   7800
         TabIndex        =   333
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   78
         Left            =   7800
         TabIndex        =   332
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Entrega merca."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   77
         Left            =   6360
         TabIndex        =   330
         ToolTipText     =   "Entrega mercancia"
         Top             =   1200
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   14160
      TabIndex        =   36
      Top             =   8640
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   136
      Top             =   360
      Width           =   15015
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   44
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha auxiliar  Albaran|F|S|||scaalb|fechaaux|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   760
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   9105
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Nombre Cliente|T|N|||scaalb|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   495
         Width           =   4080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Realizada Por|N|N|0|9999|scaalb|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   120
         Width           =   760
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   9105
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   141
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
      Begin VB.CheckBox chkFacturar 
         Caption         =   "Facturar"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Tag             =   "Facturar|N|N|||scaalb|factursn||N|"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   43
         Left            =   4035
         Picture         =   "frmFacEntAlbSAIL.frx":1274C
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. entrega"
         Height          =   255
         Index           =   65
         Left            =   3240
         TabIndex        =   303
         Top             =   150
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   7980
         Picture         =   "frmFacEntAlbSAIL.frx":127D7
         ToolTipText     =   "Buscar cliente"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   6915
         TabIndex        =   142
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada Por"
         Height          =   255
         Index           =   21
         Left            =   6915
         TabIndex        =   140
         Top             =   165
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   7980
         Picture         =   "frmFacEntAlbSAIL.frx":128D9
         ToolTipText     =   "Buscar trabajador"
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Alb."
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   139
         Top             =   150
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2835
         Picture         =   "frmFacEntAlbSAIL.frx":129DB
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   138
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   137
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Frame FrameTAXCO 
      Caption         =   "TAXCO"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   350
      Top             =   8520
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   1
         Left            =   3600
         TabIndex        =   354
         Text            =   "Text4"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   0
         Left            =   1320
         TabIndex        =   352
         Text            =   "Text4"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "KM"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   353
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula"
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
         Index           =   9
         Left            =   240
         TabIndex        =   351
         Top             =   120
         Width           =   1005
      End
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
      Left            =   7800
      TabIndex        =   165
      Top             =   8640
      Width           =   4935
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
      Begin VB.Menu mnLineaAccion 
         Caption         =   "Nueva linea"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnLineaAccion 
         Caption         =   "Modificar linea"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnLineaAccion 
         Caption         =   "Eliminar linea"
         Index           =   2
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacEntAlbSAIL"
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
Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmTr As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmTr.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 '%=%=frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticulos   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1

Private WithEvents frmProv2 As frmBasico2 '%=%=frmComProveedores
Attribute frmProv2.VB_VarHelpID = -1


Private WithEvents frmOT As frmObraOT
Attribute frmOT.VB_VarHelpID = -1
Private WithEvents frmOC As frmObraCapitulo
Attribute frmOC.VB_VarHelpID = -1
Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents FrmArtEul As frmAlmArticuEUL
Attribute FrmArtEul.VB_VarHelpID = -1


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

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

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

Dim EulerParam As String


'PORTES
' Tipo fontenas
Dim KilosAnteriores As Currency
Dim RutaCliente As Integer
Dim ZonaCliente As Integer


Dim AlmacenLineas As Integer

Dim LineaIntercalar As Integer 'NO reutilizar

Dim CarpetaImagenesEULER As String
Dim BuscaChekc As String

'Costes matriculas
Dim CostesMatriculas As Collection


Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cboTipoDat_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub cboTipoImpr_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


  
Private Sub cboTipopEnvio_KeyPress(KeyAscii As Integer)
        KEYpress KeyAscii
End Sub

Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEuler_KeyPress(Index As Integer, KeyAscii As Integer)
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


Private Sub chkFacturarKm_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    
End Sub

Private Sub cmdAccDocs_Click(Index As Integer)

End Sub



Private Function EsTipoAlbaranFacturarInmediatoTaxco() As Boolean
    EsTipoAlbaranFacturarInmediatoTaxco = hcoCodTipoM = "ALO" Or hcoCodTipoM = "ALV" Or hcoCodTipoM = "ALE"
End Function


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
                    If Not EsHistorico And vParamAplic.NumeroInstalacion = vbTaxco And EsTipoAlbaranFacturarInmediatoTaxco Then
                        If Me.chkFacturar.Value Then
                            If Text1(14).Text <> 2 Then
                                davidNumalbar = 1   'NO se podra salir SIN factura. SEA COMO SEA
                                HacerButton 12
                                davidNumalbar = 0
                            End If
                        End If
                    End If
                    
                    
                    
                    
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
                        BotonAnyadirLinea False
                    End If
                    
                    
                    
                    
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    numlinea = Data2.Recordset!numlinea
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    TerminaBloquear
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt txtAux(16), True
                    BloquearTxt Text2(9), True
                    Dim J As Integer
                    For J = 12 To 17
                        BloquearTxt txtAux(J), True
                    Next
                    'BloquearTxt Text2(9), True
                    BloquearTxt txtAux(9), True
                    Me.DataGrid1.Enabled = True
                End If
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
Dim SQL As String
Dim CambiaCerradaFacturarTaxco As String

    On Error GoTo EModificaAlb
    conn.BeginTrans
    
    CambiaCerradaFacturarTaxco = ""
    If vParamAplic.NumeroInstalacion = vbTaxco And EsTipoAlbaranFacturarInmediatoTaxco Then
        If DBLet(Data1.Recordset!factursn, "N") <> Me.chkFacturar.Value Then CambiaCerradaFacturarTaxco = "Antes: " & DBLet(Data1.Recordset!factursn, "N") & "   Despues: " & Me.chkFacturar.Value
    End If
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
    B = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    
    If B Then
        B = ModificaDesdeFormulario(Me, 1)
        
        
        
        If B Then
            'Ficha tecnica
            If SSTab1.TabVisible(2) = True Then ActualizaBDFicha
            If SSTab1.TabVisible(3) = True Then ActualizaBDFicha
        
            SQL = "UPDATE scaalb SET nomdirec=" & DBSet(Text2(12).Text, "T") & " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!Numalbar
            conn.Execute SQL
            
            
            If CambiaCerradaFacturarTaxco <> "" Then
                Set LOG = New cLOG
                CambiaCerradaFacturarTaxco = Replace(CambiaCerradaFacturarTaxco, "1", "Abierto")
                CambiaCerradaFacturarTaxco = Replace(CambiaCerradaFacturarTaxco, "0", "Cerrado")
                CambiaCerradaFacturarTaxco = "[Cerrar/abrir taller] " & CambiaCerradaFacturarTaxco
                LOG.Insertar 29, vUsu, CambiaCerradaFacturarTaxco
                Set LOG = Nothing
            End If
            
            
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
                SQL = SQL & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!Numalbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
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
            
            CrearCarpetaComun2 True, 0
            
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
            
            If B Then
            
                If InstalacionEsEulerTaxco Then
                'EULER  As
                    Set FrmArtEul = New frmAlmArticuEUL
                    'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
                    FrmArtEul.FechaDoc = CDate(Text1(1).Text)
                    FrmArtEul.Codprove = -1
                    FrmArtEul.DesdeVentas = True
                    FrmArtEul.Show vbModal
                    Set FrmArtEul = Nothing
            
                Else
            
                    Set FrmArt = New frmAlmArticulos
                    FrmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo busqueda
                    FrmArt.DeConsulta = True
                    FrmArt.Show vbModal
                    Set FrmArt = Nothing
                End If
                PonerFoco txtAux(1)
            End If
            
'    Case 9 'CENTRO COSTE/ PROVEEDOR
'        If vEmpresa.TieneAnalitica Then
'            'centro de coste
'            AbrirForm_CentroCoste
'        Else
'            Set frmProv2 = New frmComProveedores
'            frmProv2.DatosADevolverBusqueda = "1"
'            frmProv2.Show vbModal
'            Set frmProv2 = Nothing
'            If CadenaDesdeOtroForm <> "" Then
'                txtaux(9).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
'                Text2(9).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
'            End If
'        End If
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
Dim J As Integer

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
            For J = 12 To 17
                BloquearTxt txtAux(J), True
            Next
            'BloquearTxt Text2(9), True
            BloquearTxt txtAux(9), True
            DataGrid1.Columns(4).Caption = "Artículo"
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            PonerForaGrid
            ModificaLineas = 0
            LineaIntercalar = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                cmdRegresar_Click
                PonerFocoOBj DataGrid1
            End If
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim Cad As String
Dim RS As ADODB.Recordset
Dim DatosVehiculo As String


    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Si es Albaran para factura RECTIFICATIVA pedir la Factura que se va
    'a Rectificar y si existe en el historico, tabla "scafac", entonces dejamos
    'que inserte el Albaran Rectificativo, si no salimos
    DatosVehiculo = ""
    If CodTipoMov = "ART" Then
        cadList = ""

        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 225
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        If Trim(Mid(cadList, 1, 12)) = "codtipom=''" Then
            PonerModo 0
            Exit Sub
        End If
        
        
        'cargar los datos de la factura recuperada en el formulario
        NomTraba = "select codtipom as codtipmf,numfactu,fecfactu,codclien,nomclien,domclien,scafac.codpobla,pobclien,proclien,nifclien,telclien,"
        NomTraba = NomTraba & "coddirec,nomdirec,scafac.codagent,nomagent,scafac.codforpa, nomforpa,dtoppago,dtognral "  'JUNIO 2010 añado el envio
        NomTraba = NomTraba & " from (scafac inner join sforpa on scafac.codforpa=sforpa.codforpa) "
        NomTraba = NomTraba & " inner join sagent on scafac.codagent=sagent.codagent where " & cadList
        
        Set RS = New ADODB.Recordset
        RS.Open NomTraba, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        PonerModo 3
        
        If Not RS.EOF Then
            Text1(4).Text = RS!codClien
            FormateaCampo Text1(4)
            Text1(5).Text = RS!NomClien
            Text1(6).Text = RS!nifClien
            Text1(7).Text = DBLet(RS!telclien, "T")
            Text1(8).Text = RS!domclien
            Text1(9).Text = RS!codpobla
            Text1(10).Text = RS!pobclien
            Text1(11).Text = DBLet(RS!proclien, "T")
            Text1(12).Text = DBLet(RS!CodDirec, "T")
            FormateaCampo Text1(12)
            Text2(12).Text = DBLet(RS!nomdirec, "T")
            Text1(14).Text = RS!codforpa
            FormateaCampo Text1(14)
            Text2(14).Text = RS!nomforpa
            Text1(15).Text = DBLet(RS!DtoPPago, "N")
            FormateaCampo Text1(15)
            Text1(16).Text = DBLet(RS!DtoGnral, "N")
            FormateaCampo Text1(16)
            Text1(17).Text = DBLet(RS!CodAgent, "T")
            FormateaCampo Text1(17)
            Text2(17).Text = RS!NomAgent
            Text1(37).Text = RS!codtipmf
            Text1(36).Text = DBLet(RS!Numfactu, "N")
            FormateaCampo Text1(36)
            Text1(35).Text = RS!FecFactu
            
            'Observacion 1   'DAVID
            'Text1(18).Text = "RECTIFICA A FACTURA: " & RS!codtipmf & ", " & RS!NumFactu & ", " & RS!FecFactu
            Text1(18).Text = RS!Numfactu & ", " & RS!FecFactu
            'Observacion 2
            Text1(19).Text = motivo
            
            NomTraba = "tipofact"
            Cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            If Cad = "0" Then BloquearDatosCliente (False)
            
            
            'Memorizo cad con codtipom
            Cad = RS!codtipmf
            
            
            
            
            
            'recuperamos el tipo de facturacion del cliente
            Me.cboFacturacion.ListIndex = CInt(NomTraba)
            
             
            
        Else
            Cad = Replace(cadList, " and ", "|") & "|"
            Cad = Replace(Cad, "codtipom=", "")
            Cad = Replace(Cad, "numfactu=", "")
            Cad = Replace(Cad, "fecfactu=", "")
            cadList = Replace(Cad, "'", "")
            Text1(37).Text = Trim(RecuperaValor(cadList, 1))
            Cad = RecuperaValor(cadList, 2)
            Text1(36).Text = Val(Cad)
            FormateaCampo Text1(36)
            Cad = RecuperaValor(cadList, 3)
            Text1(35).Text = Format(Cad, "dd/mm/yyyy")
            Text1(19).Text = motivo
        
            Text1(18).Text = Text1(36).Text & ", " & Text1(35).Text
        
            Cad = "N" 'para que la busqueda de despues no de error
        End If
        RS.Close
        
        'DAVID
        'Para que meta la letra de serie, NO el tipo moviemiento
        RS.Open "SELECT * FROM stipom WHERE codtipom='" & Cad & "'"
        If Not RS.EOF Then Cad = DBLet(RS!LetraSer, "T")
        RS.Close
        If Cad = "" Then Cad = CodTipoMov
        Text1(18).Text = "RECTIFICA A FACTURA: " & Cad & ", " & Text1(18).Text
        
            
        'DAVID
        'JUNIO 2010
        'Envio por defecto del cliente
        Text1(29).Text = ""
        Text2(29).Text = ""
        If Text1(4).Text <> "" Then
            Cad = "select sclien.codenvio,nomenvio from  sclien,senvio where sclien.codenvio=senvio.codenvio AND sclien.codclien= " & Text1(4).Text
            RS.Open Cad, conn, adOpenForwardOnly, adCmdText
            If Not RS.EOF Then
                Text1(29).Text = RS!CodEnvio
                Text2(29).Text = RS!nomEnvio
            End If
            RS.Close
       End If
            
        
        
        
        
        
        
        Set RS = Nothing
    Else
        DatosVehiculo = ""
        If CodTipoMov = "ALO" And vParamAplic.NumeroInstalacion = vbTaxco Then
           CadenaDesdeOtroForm = ""
           frmListado5.OpcionListado = 29
           frmListado5.Show vbModal
           
           
           DatosVehiculo = CadenaDesdeOtroForm
           
        End If
    
        'Añadiremos el boton de aceptar y demas objetos para insertar
        PonerModo 3
        
        
        If DatosVehiculo <> "" Then
            If DatosVehiculo <> "OK" Then AsignarDatosVehiculoReparacion DatosVehiculo
        End If
        
        
    End If
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = CStr(NomTraba)

    'El preparador del material lo hacemos tb al trabajador actual
    Text1(28).Text = Text1(3).Text
    Text2(28).Text = Text2(3).Text

    'Marca de para facturar
    If vParamAplic.MarcarAlbaranFacturar Then
        NomTraba = "S"
        If vParamAplic.NumeroInstalacion = vbTaxco And EsTipoAlbaranFacturarInmediatoTaxco Then NomTraba = ""
        If NomTraba <> "" Then chkFacturar.Value = 1
    End If



    If InstalacionEsEulerTaxco Then cboTipoImpr.ListIndex = 4 'por defecto ALBARAN

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(44).Text = Text1(1).Text
    Text1(30).Text = CodTipoMov
    
    
    If DatosVehiculo <> "" Then
        'Solo puede pasar en TAXCO, en ALOS
        SSTab1.Tab = 2
        PonerFoco Me.txtTaxco(6)
    Else
        PonerFoco Text1(1)
    End If
End Sub


Private Sub BotonAnyadirLinea(Intercalando As Boolean)
Dim J As Integer
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
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

    For J = 12 To 17
        txtAux(J).Text = ""
        If J < 15 Then PonerDatosNuevosLineaAlbaran False, J
    Next J
    Text2(9).Text = ""
    BloquearTxt txtAux(16), False
    BloquearTxt Text2(9), True
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
Dim B As Boolean

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
     
        
       FocoBuscar
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then FocoBuscar
        
    End If
    
    If Me.EsHistorico = False Then
        'Hacer busquedar del tipo de movimiento de albaran en el que estamos
        Text1(30).Text = CodTipoMov
        BloquearTxt Text1(30), True
    End If
End Sub


Private Sub FocoBuscar()
Dim B As Boolean

    'Si pasamos el control aqui lo ponemos en amarillo
     B = False
     If vParamAplic.NumeroInstalacion = vbTaxco And CodTipoMov = "ALO" Then B = True
    If B Then
        Me.SSTab1.Tab = 2
        PonerFoco txtTaxco(0)
        DoEvents
        Text1(0).BackColor = vbWhite
        txtTaxco(0).BackColor = vbYellow
    Else
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    End If
End Sub


Private Sub BotonVerTodos()
Dim Cad As String
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        Cad = ""
        If Not EsHistorico Then Cad = " codtipom='" & CodTipoMov & "'"
        MandaBusquedaPrevia Cad, False
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        If EsHistorico = False Then
            CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & CodTipoMov & "'"
        End If
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda False
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
    
    
    
    'bloqueamos el registro a modificar
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    
    BloquearTxt txtAux(12), False
    BloquearTxt txtAux(13), False
    BloquearTxt txtAux(14), False
    BloquearTxt txtAux(15), False
    BloquearTxt txtAux(16), False
    
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
    
    BloquearTxt txtAux(16), False 'Campo Ampliacion Linea
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

    If vParamAplic.NumeroInstalacion = vbTaxco Then
        If hcoCodTipoM = "ALD" Or hcoCodTipoM = "ALB" Then
            If vUsu.Nivel > 0 Then
                MsgBox "Sin permiso", vbExclamation
                Exit Sub
            End If
        End If
                
    End If
    If vParamAplic.NumeroInstalacion = vbEuler Then
        Cad = "  codtipoa='" & Text1(30).Text & "' AND numalbar"
        Cad = DevuelveDesdeBD(conAri, "numproyec", "sproyectolin", Cad, Text1(0).Text)
        If Cad <> "" Then
            MsgBox "Albaran vinculado al proyecto: " & Cad, vbExclamation
            Exit Sub
        End If
    End If
    
    Cad = "Cabecera de Albaran+es." & vbCrLf
    Cad = Cad & "------------------------------------       " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Albaran:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(30).Text
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Text1(1).Text
'    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
      
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        
        If Not Me.Data2.Recordset.EOF Then
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

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        

    
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la LÍNEA de Albaran?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
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


Private Sub cmdLineasCostes_Click(Index As Integer)
Dim C As String
Dim R As Boolean
Dim Cant As Currency
Dim TotalImporte As Currency
Dim ImporteADescontar As Currency
Dim Aux2 As String
Dim CPrecioFact As CPreciosFact
Dim Precio  As String

    If Modo <> 2 Then Exit Sub
    If EsHistorico Then Exit Sub
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Index = 3 Then
        'IMPRESION
        If ListView2.ListItems.Count = 0 Then Exit Sub
        ImprimirCostesEuler
        Exit Sub
    End If
    
    If Index = 4 Then
        'NUEVO PEDIDO PROVEEDOR DESDE ALbaran
        frmComEntPedidosSa.AlbaranVenta = Data1.Recordset!codtipom & Format(Data1.Recordset!FechaAlb, "ddmmyyyy") & Data1.Recordset!Numalbar
        frmComEntPedidosSa.Show vbModal
        
        If True Then CargaCostesEuler False
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    C = ""
    R = False
    If Index > 0 Then
        If ListView2.ListItems.Count = 0 Then Exit Sub
        If ListView2.SelectedItem Is Nothing Then Exit Sub
        
        If Index = 5 Then
                    
                'Todo lo que sean lineas de proveedor , que no esten importadas  las meto en slialb
                'ALC  FAC
                'Veré que no estan puestas todavia
                Aux2 = ""
                For NumRegElim = 1 To ListView2.ListItems.Count
                    
                    If ListView2.ListItems(NumRegElim).Text = "ALC" Or ListView2.ListItems(NumRegElim).Text = "FAC" Then
                    
                        Aux2 = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and codartic "
                        Aux2 = DevuelveDesdeBD(conAri, "codartic", "slialb", Aux2, ListView2.ListItems(NumRegElim).SubItems(9), "T")
                        If Aux2 = "" Then
                            'Perfecto. La añadimos
                            C = C & ", " & DBSet(ListView2.ListItems(NumRegElim).SubItems(9), "T")
                        Else
                            'Mal, ya esta añadida
                            'No hacemos nada
                            
                        End If
                    End If
                    
                Next
                
                If C = "" Then
                    MsgBox "Ningun coste Proveedor para inserar en venta", vbExclamation
                    Exit Sub
                End If
                
                
                
                'Veremos si hay una primera linea de costes articulo vario
                Set miRsAux = New ADODB.Recordset
                Aux2 = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and  sartic.codartic=slialb.codartic order by numlinea"
                Aux2 = "Select slialb.codartic,slialb.nomartic,artvario,numlinea,ImporteL,cantidad,dtoline1,dtoline2 from slialb,sartic " & Aux2
                miRsAux.Open Aux2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Aux2 = ""
                Cant = 0  'linea 1
                ImporteADescontar = 0
                kCampo = 0
                C = "|"
                While Not miRsAux.EOF
                    Cant = Cant + 1
                    If miRsAux!artvario = 1 Then
                        'Si es la primera linea
                        If Cant = 1 Then
                            'primera linea. Puede descontar de aqui.
                            'la cantidad debe ser 1
                            ' y los descuentos 0
                            ' precioar * 1 = importel
                            If miRsAux!cantidad = 1 And miRsAux!dtoline1 = 0 And miRsAux!dtoline2 = 0 Then
                                Aux2 = miRsAux!codArtic & "   " & miRsAux!NomArtic
                                ImporteADescontar = miRsAux!ImporteL
                                kCampo = miRsAux!numlinea
                                
                                
                                
                            Else
                                MsgBox "Primer linea varios no es compensable(cantidad /descuentos)", vbExclamation
                                Cant = 0
                            End If
                        Else
                            If Cant = 2 Then
                                'Segunda linea tambien es de varios. Tenemos que ofertar que eliga la linea
                                MsgBox "Mas de una linea para poder descontar. **", vbExclamation
                            End If
                        End If
                    
                    End If
                    C = C & miRsAux!codArtic & "|"   'para saber si hay que insertar o no la linea de coste
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                
                If kCampo > 0 Then
                    If kCampo = 1000 Then
                        'Habrá que mostrar frmmensaje para seleccionar linea
                        
                    Else
                        'Sobre la linea
                        Aux2 = "Linea articulo descontar: " & vbCrLf & Aux2 & vbCrLf
                        Aux2 = Aux2 & "Importe linea:     " & Format(ImporteADescontar, FormatoImporte) & "€" & vbCrLf
                        Aux2 = Aux2 & vbCrLf & "¿Descontar importes de esta linea?"
                        Aux2 = MsgBox(Aux2, vbQuestion + vbYesNoCancel)
                        If CByte(Aux2) = vbCancel Then Exit Sub
                        
                        If CByte(Aux2) = vbNo Then kCampo = 0
                        
                    End If
                    
                End If
                
                
                
                
                
                'Perfecto , ahora vamos a añadir estos articulos pero en el txtaux(preciocoste)=0 para que no los vuelva a poner como coste
                AlmacenLineas = -1
                PonerUltAlmacen
                BuscaChekc = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
                If BuscaChekc = "" Or AlmacenLineas <= 0 Then
                    MsgBox "Error almacen // Centro de coste: " & AlmacenLineas & " // " & BuscaChekc, vbExclamation
                    Exit Sub
                End If
                
                Set CPrecioFact = New CPreciosFact
                CPrecioFact.CodigoClien = Text1(4).Text
                CPrecioFact.FijarTarifaActividad
                Aux2 = C 'artiuclos que NO hay que insertar
                For NumRegElim = 1 To ListView2.ListItems.Count
                    
                    If ListView2.ListItems(NumRegElim).Text = "ALC" Or ListView2.ListItems(NumRegElim).Text = "FAC" Then
                        
                        
                        
                        If InStr(1, Aux2, "|" & ListView2.ListItems(NumRegElim).SubItems(9) & "|") = 0 Then
                            'OK, esta linea de coste aun no se ha añadido
                            
                            
                            'codAlmac, codArtic NomArtic Ampliaci
                            C = AlmacenLineas & " codalmac , " & DBSet(ListView2.ListItems(NumRegElim).SubItems(9), "T") & " codartic,"
                            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & " nomartic, "
                            C = C & DBSet("Alb: " & ListView2.ListItems(NumRegElim).SubItems(2), "T") & " Ampliaci, "
                            
                            
                            'cantidad  butlos
                            CPrecioFact.CodigoArtic = ListView2.ListItems(NumRegElim).SubItems(9)                                                                                    'ORIGPRE
                            Precio = CPrecioFact.ObtenerPrecio(False, Text1(1).Text, txtAnterior, "")   'txtAnterior: ORIGPRE
                            
                            If Precio = "0" Then
                                'Como minimo, el precio de compra
                                txtAnterior = "M"
                                CPrecioFact.DtoPermitido = False
                                Precio = ListView2.ListItems(NumRegElim).SubItems(6)
                            End If
                            Cant = ImporteFormateado(ListView2.ListItems(NumRegElim).SubItems(5))
                            Cant = Cant * -1 'En la funcion Insertarlineafactu cambia el signo otra vez
                            C = C & DBSet(Cant, "N") & " cantidad, "
                            C = C & DBSet(Cant, "N") & " numbultos, "

                            'precioar  origpre   dtoline1     dtoline2 Importe CodCCost
                            C = C & DBSet(Precio, "N") & " precioar, "
                            C = C & "'" & txtAnterior & "' origpre, "
                                    
                            If Not CPrecioFact.DtoPermitido Then
                                CPrecioFact.Descuento1 = 0
                                CPrecioFact.Descuento2 = 0
                            End If
                            C = C & DBSet(CPrecioFact.Descuento1, "N") & " dtoline1 ," & DBSet(CPrecioFact.Descuento2, "N") & " dtoline2, "
                            
                            txtAnterior = CalcularImporte(ImporteSinFormato(CStr(Cant)), Precio, ImporteSinFormato(CStr(CPrecioFact.Descuento1)), ImporteSinFormato(CStr(CPrecioFact.Descuento2)), vParamAplic.TipoDtos)
                        
                            TotalImporte = ImporteSinFormato(txtAnterior)
                     
                            
                            C = C & DBSet(TotalImporte, "N") & " importel, " & DBSet(BuscaChekc, "T") & " CodCCost, '' numlote,"
                            
                            'Codprovex
                            C = C & RecuperaValor(ListView2.ListItems(NumRegElim).Tag, 3) & " Codprovex"
                            
                            
                            C = "  SELECT " & C
                            If InsertarLineasFactu("SI", C) Then
                        
                                'Descontamos
                                If kCampo > 0 Then
                                    'Voy a desconar del total de la linea que hemos dicho
                                    If Abs(TotalImporte) < ImporteADescontar Then
                                        C = "UPDATE slialb SET importel=importel " & DBSet(TotalImporte, "N")
                                        C = C & " , precioar=precioar " & DBSet(TotalImporte, "N")
                                        C = C & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                                        C = C & " AND numlinea =" & kCampo
                                        If Not ejecutar(C, False) Then Exit For
                                         ImporteADescontar = ImporteADescontar + TotalImporte
                                    Else
                                        'No puede descontar mas importe
                                        
                                        C = ListView2.ListItems(NumRegElim).SubItems(9) & " - " & ListView2.ListItems(NumRegElim).SubItems(4)
                                        C = "No se puede descontar mas importe" & vbCrLf & C & vbCrLf & "Importe coste: " & TotalImporte & " -- " & ImporteADescontar
                                        MsgBox C, vbExclamation
                                        kCampo = 0 'para que no intente compensar mas
                                    End If
                                End If
                            End If
                        End If  'INSERTAMOS
                    End If
                Next
                
                
                
                CargaGrid DataGrid1, Data2, True
                MsgBox "Proceso finalizado", vbInformation
                Exit Sub  'Ya esta todo
                

        End If
        
        
        If ListView2.SelectedItem.Text <> "MAT" Then
            MsgBox "No se puede modificar este dato", vbExclamation
            Exit Sub
        End If
        
        
        
    End If
        
    
        
    
    'OK. Abrimos para cargar los datos
    If Index <> 2 Then
        CadenaDesdeOtroForm = ""
        C = ""
        If Index = 1 Then
            C = Trim(Mid(ListView2.SelectedItem.SubItems(2), 3))
            C = "numlinea=" & C & " and codtipom ='" & CodTipoMov & "' and numalbar"
            C = DevuelveDesdeBD(conAri, "codartic", "slialb_eu", C, Text1(0).Text)
            C = ListView2.SelectedItem.SubItems(3) & "|" & C & "|" & ListView2.SelectedItem.SubItems(4) & "|"
            C = C & ListView2.SelectedItem.SubItems(5) & "|" & ListView2.SelectedItem.SubItems(6) & "|" & ListView2.SelectedItem.SubItems(7) & "|"
        End If
        frmListado3.OtrosDatos = CStr(C)
        frmListado3.Opcion = 67
        frmListado3.Show vbModal
        If CadenaDesdeOtroForm <> "" Then InsertarModicarArticuloCostesEuler CByte(Index)
        
    Else
        
            'Eliminar
            C = "Va a eliminar la linea seleccionada:" & vbCrLf & ListView2.SelectedItem.SubItems(4) & "   " & ListView2.SelectedItem.SubItems(7)
            If MsgBox(C, vbQuestion + vbYesNoCancel) = vbYes Then
                InsertarModicarArticuloCostesEuler CByte(Index)
               
            End If
            
        
        
    End If
    
    If True Then CargaCostesEuler False
End Sub

'0 insertar
'1 modifiar
'2 eliminar

Private Sub InsertarModicarArticuloCostesEuler(Accion As Byte)
Dim cS As CStock
Dim C As String
Dim cantidad As Currency

    Set cS = New CStock
    On Error GoTo eInsertarModicarArticuloCostesEuler
        
        
    conn.BeginTrans
    cS.codAlmac = 1
    cS.DetaMov = "MAT"
    cS.Documento = CodTipoMov & Format(Text1(0).Text, "0000000")
    
            
    'C = PonerTrabajadorConectado("")
    'If C = "" Then C = Text1(3).Text
    cS.Trabajador = Val(Text1(4).Text)
        
    If Accion < 2 Then 'insertar modificar
    
            
           If Accion = 0 Then
                'INSERTAR
                C = " codtipom ='" & CodTipoMov & "' and numalbar"
                C = DevuelveDesdeBD(conAri, "max(numlinea)", "slialb_eu", C, Text1(0).Text)
                C = Val(C) + 1
                
            Else
                'modificar
                C = Trim(Mid(ListView2.SelectedItem.SubItems(2), 3))
            End If
            cS.LineaDocu = C
            If Accion = 1 Then
                'Primero que nada borramos el movimiento anterior
                cS.codArtic = ListView2.SelectedItem.SubItems(9)
                cS.FechaMov = ListView2.SelectedItem.SubItems(3)
                cS.cantidad = CCur(ListView2.SelectedItem.SubItems(5))
                cS.tipoMov = "E"
                If Not cS.DevolverStock2 Then Err.Raise 513, , "No puede eliminar movimiento anterior"
                
            End If
        
            C = " VALUES ('" & CodTipoMov & "'," & Text1(0).Text & "," & C & ",1,"
            C = "REPLACE INTO slialb_eu(codtipom,numalbar,numlinea,codalmac,fechamov,codartic,nomartic,cantidad,precioar) " & C
            C = C & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F") & ","
            C = C & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T") & ","
            C = C & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ","
            C = C & DBSet(RecuperaValor(CadenaDesdeOtroForm, 4), "N") & ","
            C = C & DBSet(RecuperaValor(CadenaDesdeOtroForm, 5), "N") & ")"
            conn.Execute C
          
            cS.codArtic = RecuperaValor(CadenaDesdeOtroForm, 2)
            
            cS.FechaMov = RecuperaValor(CadenaDesdeOtroForm, 1)
            cS.HoraMov = cS.FechaMov & Format(Now, " hh:mm:ss")
            
            cantidad = RecuperaValor(CadenaDesdeOtroForm, 5)
            cantidad = cantidad * RecuperaValor(CadenaDesdeOtroForm, 4)
            cS.Importe = cantidad
            cS.tipoMov = "S"
            cantidad = RecuperaValor(CadenaDesdeOtroForm, 4)
    
    
            
            C = ""
            cS.ComprobarFechaInventario False, C
            If C <> "" Then Err.Raise 513, , C
            
            cS.cantidad = cantidad
            If cS.MueveStock Then
                If Not cS.MoverStock(False, True, True) Then Err.Raise 513, , "Actualizando stock"
                If Not cS.ActualizarStock(False, True) Then Err.Raise 513, , "Actualizando stock(2)"
            End If
            
    Else
        'Eliminar


        C = Trim(Mid(ListView2.SelectedItem.SubItems(2), 3))
        cS.LineaDocu = C
        
        'Primero que nada borramos el movimiento anterior
        cS.codArtic = ListView2.SelectedItem.SubItems(9)
        cS.FechaMov = ListView2.SelectedItem.SubItems(3)
        cS.cantidad = CCur(ListView2.SelectedItem.SubItems(5))
        cS.tipoMov = "E"
        If Not cS.DevolverStock2 Then Err.Raise 513, , "No puede eliminar movimiento anterior"
        
        C = Trim(Mid(ListView2.SelectedItem.SubItems(2), 3))
        C = "numlinea=" & C & " and codtipom ='" & CodTipoMov & "' and numalbar =" & Text1(0).Text
        C = "DELETE FROM slialb_eu WHERE " & C
        conn.Execute C
    End If
            
eInsertarModicarArticuloCostesEuler:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Acciones costes EULER"
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    Set cS = Nothing
End Sub





Private Sub cmdLineasImpresion_Click(Index As Integer)
Dim Cad As String
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Index > 0 Then
        If lwEulerLineas.ListItems.Count = 0 Then
            MsgBox "Ningun dato", vbExclamation
            Exit Sub
        End If
        If Index < 3 Then
            'Modificar eliminar.
            'el seleccionado
            If Me.lwEulerLineas.SelectedItem Is Nothing Then
                MsgBox "Seleccione una linea", vbExclamation
                Exit Sub
            End If
        End If
    Else
        If Data2.Recordset.EOF Then Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    
    If Index < 2 Then
        'nuevo modificar
        If Index = 1 Then
            CadenaDesdeOtroForm = Mid(lwEulerLineas.SelectedItem.Key, 2, 3)
        Else
            CadenaDesdeOtroForm = ""  '"" = nuevo   id= linea
        End If
        frmListado5.OtrosDatos = Data1.Recordset!codtipom & "|" & Data1.Recordset!Numalbar & "|"
        frmListado5.OpcionListado = 28
        frmListado5.Show vbModal
        
    
    Else
        If Index = 2 Then
            'Eliminar
            Cad = "Va a eliminar linea impresion" & vbCrLf & "Articulo : " & Me.lwEulerLineas.SelectedItem.Text & vbCrLf
            Cad = Cad & "Descripcion : " & Me.lwEulerLineas.SelectedItem.SubItems(1) & vbCrLf
            Cad = Cad & "Importe : " & Me.lwEulerLineas.SelectedItem.SubItems(4) & vbCrLf
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                Cad = " WHERE codtipom='" & Data1.Recordset!codtipom & "' AND numalbar = " & Data1.Recordset!Numalbar
                 Cad = "DELETE FROM slialb_eu2 " & Cad & " AND numlinea= " & Mid(Me.lwEulerLineas.SelectedItem.Key, 2, 3)
                If ejecutar(Cad, False) Then CadenaDesdeOtroForm = "OK"
            End If

        Else
            'imprimir
            If lwEulerLineas.Tag <> "" Then
                MsgBox lwEulerLineas.Tag, vbExclamation
            Else
                BotonImprimir 90, False '90: Informe de Albaranes lineas especiales
            End If
        End If
    End If
    
    If CadenaDesdeOtroForm <> "" Then PonerLineasAlbaranEULER
    
    
End Sub

Private Sub cmdOrdenTallerTaxco_Click()
    If Modo <> 2 Then Exit Sub
    
     BotonImprimir 91, False
       
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String
Dim Port As Integer      'Port: para saber si ha metido/Modificado el articulo de portes
Dim LanzaFacurar As Boolean

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
        'Fontenas
        
        If vParamAplic.TipoPortes = 1 Then
            'Si lleva portes haremos varias cosas
            Port = HacerAccionesPortes
            CargaGrid DataGrid1, Data2, True
            Set miRsAux = Nothing
        End If
        
        
        If InstalacionEsEulerTaxco Then
            If ListView2.ListItems.Count > 0 Then ListView2.ListItems.Clear
            
            
            If lwEulerLineas.ListItems.Count > 0 Then PonerLineasAlbaranEULER   'para que vuelva a ver si la suma de importes coincide
        End If
        
        LanzaFacurar = False
        If AvisosAlbaranesEfectifivoPendienteFacturar(True) Then LanzaFacurar = True
        
        
        
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
    
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
        cmdCancelar.Cancel = True
        
        
        If LanzaFacurar Then HacerButton 12
        
        
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


Private Sub Command1_Click()
'    frmListado5.OpcionListado = 27
'    frmListado5.Show vbModal
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
        If X > 7660 And X < 7950 Then
            If IsNull(Me.Data2.Recordset!origpre) Then Exit Sub
            Select Case DataGrid1.Columns(9).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
'                Case Else
'                    Me.DataGrid1.ToolTipText = ""
            End Select
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
        
    End If
End Sub

Private Sub PonerForaGrid()
    'Dim RS As ADODB.Recordset
    'Dim SQL As String
    Dim Borrar As Boolean
    Dim J As Integer
    Dim Desde As Integer
    Dim Base As Integer
    Dim C As String
    
On Error GoTo Error1
  
        Borrar = True
  
        If vEmpresa.TieneAnalitica Then
            Desde = 14
            
        Else
            Desde = 15
        End If
'        For J = 0 To Data2.Recordset.Fields.Count - 1
'            Debug.Print J & "  : " & Data2.Recordset.Fields(J).Name
'        Next J
        'Nuevo SAIL. codtipom numalbar numlinea
        'SQL = "select codcapit,codtipor, codtipor as codtraba,precoste,ampliaci from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
        If Not Data2.Recordset.EOF Then
            Borrar = False
            For J = Desde To Data2.Recordset.Fields.Count - 1
                Base = J - Desde + 10
                'Para que no vaya a buscar a las tablas(capitulos,OPT, trabaj) cada vez.....
                If Base < 12 Or Base > 14 Then
                    If Not IsNull(Data2.Recordset.Fields(J).Value) Then
                        txtAux(Base).Text = Data2.Recordset.Fields(J).Value
                        'Numero
                        If Base = 15 Then PonerFormatoDecimal txtAux(15), 2
                    Else
                        txtAux(Base).Text = ""
                    End If
                End If
            Next J
    
        
            J = ModificaLineas
            
            
            C = DBLet(Data2.Recordset!codcapit, "T")
            If txtAux(12).Text <> C Then
                txtAux(12).Text = C
                PonerDatosNuevosLineaAlbaran False, 12
            End If
            
            
                 
            C = DBLet(Data2.Recordset!codtipor, "T")
            If txtAux(13).Text <> C Then
                txtAux(13).Text = C
                PonerDatosNuevosLineaAlbaran False, 13
            End If
            
            C = DBLet(Data2.Recordset!CodTraba, "T")
            If txtAux(14).Text <> C Then
                txtAux(14).Text = C
                PonerDatosNuevosLineaAlbaran False, 14
            End If
            
            
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            Else
                '- nombre proveedor
                Me.txtAux(9).Text = DBLet(Data2.Recordset!codProvex, "T")
                Text2(9).Text = DBLet(Me.Data2.Recordset!nomprove, "T")
            End If
        
            ModificaLineas = J
            
            If Not EsHistorico Then
                C = DevuelveDesdeBD(conAri, "observa", "slialt", "codtipom= '" & CodTipoMov & "' AND numalbar = " & Text1(0).Text & " AND numlinea", Data2.Recordset!numlinea, "N")
                txtAux(17).Text = C
            End If
            
      Else
        'EOF
        For J = 9 To 17
            txtAux(J).Text = ""
        Next
        Text2(0).Text = "": Text2(9).Text = ""
        Text2(2).Text = "": Text2(13).Text = ""
        
      End If   'De EOF
        
    

    
    
    

    
Error1:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Borrar = True
    End If
    
    If Borrar Then
        Text2(9).Text = ""
        Text2(0).Text = ""
        Text2(2).Text = ""
        Text2(13).Text = ""
        For J = 9 To 17
            txtAux(J).Text = ""
        Next
    End If

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    
        If Not Data2.Recordset.EOF Then
            If Not DGrid_CambiarFila(DataGrid1) Then Exit Sub
        End If
        
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then PonerForaGrid
        
        
        
      
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If AlbAvisoGenerado > 0 Then
        PonerCadenaBusqueda False
        'Simulo que pulsa lineas
        mnLineas_Click
        
        'Simulo que le da a insertar nueva
        mnNuevo_Click
        
        'AlbAvisoGenerado
        AlbAvisoGenerado = 0
    End If
        
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF And Modo <> 5 Then PonerCadenaBusqueda False
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon


    '
    imgObserva(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
    imgObserva(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    
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
       
        
        .Buttons(16).Image = 16 'Imprimir
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
        
        
        
        'ANTES ERAN PORTES. EN sail y EULER NO hay portes
        'ES el boton de cambiar de tipo
        'If vParamAplic.TipoPortes <> 1 Then
        '    .Buttons(14).Style = tbrSeparator
        '    .Buttons(14).ToolTipText = ""
        'Else
        '    .Buttons(14).Style = tbrDefault
        '    .Buttons(14).ToolTipText = "Imprimir portes"
        'End If
        
    End With
    Me.SSTab1.Tab = 0
    Set CostesMatriculas = New Collection
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    CargarCombos  'CargarComboFacturacion  + cbotipestado
    
    CodTipoMov = hcoCodTipoM
            
                
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    
    'TABS visibles para euler
    Modo = 0
    'If InstalacionEsEulerTaxco Then
    If vParamAplic.NumeroInstalacion = vbEuler Then
        If CodTipoMov = "ALO" Or CodTipoMov = "ALE" Or CodTipoMov = "ALR" Or CodTipoMov = "ALV" Then Modo = 1
    End If
    SSTab1.TabVisible(5) = Modo = 1 'Fichadas. PARA ALE,ALO,
    SSTab1.TabVisible(4) = Modo = 1 'Costes. PARA ALE,ALO,
    SSTab1.TabVisible(6) = Modo = 1 'Lineas impresion de albaranes
    
    PrimeraVez = True
    Modo = 0
    cadList = "Albaranes Clientes"
    EulerParam = ""
    CarpetaImagenesEULER = ""
    Toolbar1.Buttons(14).visible = False
    
    If InstalacionEsEulerTaxco Then
        
        Label1(60).Caption = "Estado"
        FrameOT_Euler.visible = False
        CarpetaImagenesEULER = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
        lblTituloEst.Caption = ""
        If vParamAplic.NumeroInstalacion = vbEuler Then Toolbar1.Buttons(14).visible = True
        FrameTAXCO.visible = False
        
        If CodTipoMov = "ALO" Then
            SSTab1.TabVisible(2) = True
            SSTab1.TabCaption(2) = "Orden de trabajo"
            lblTituloEst.ForeColor = &H800000
            Toolbar1.Buttons(14).visible = True
            Toolbar1.Buttons(14).ToolTipText = "Ver histórico reparaciones"
            Label3(7).Caption = "Tr. exterior"
            
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then
               SSTab1.TabCaption(2) = "Orden de taller"
                FrameTAXCO.BorderStyle = 0
                FrameTAXCO.visible = True
                chkFacturar.Caption = "CERRADO"
                Me.FrameOT_Taxco.visible = True
                
                
                Label1(57).visible = False
                imgBuscar(12).visible = False
                Text2(1).visible = False
                Text1(42).visible = False
                
                
            Else
                FrameOT_Euler.visible = True
            End If

            
            cadList = UCase(SSTab1.TabCaption(2))

            If vParamAplic.NumeroInstalacion = vbEuler Then EulerParam = CarpetaImagenesEULER
            
        ElseIf CodTipoMov = "ALE" Then
            SSTab1.TabVisible(2) = True
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                SSTab1.TabCaption(2) = "Gestoria"
            Else
                SSTab1.TabCaption(2) = "Trabajo exterior"
            End If
            cadList = UCase(SSTab1.TabCaption(2))
            lblTituloEst.ForeColor = &H80&
            Label3(7).Caption = "Orden trabajo"
            
            If vParamAplic.NumeroInstalacion = vbEuler Then EulerParam = CarpetaImagenesEULER
            
        ElseIf CodTipoMov = "ALR" Then
            SSTab1.TabVisible(3) = True
            If vParamAplic.NumeroInstalacion = vbEuler Then EulerParam = CarpetaImagenesEULER
           
        ElseIf CodTipoMov = "ALD" Then
            cadList = "Albaranes gasolinera"
            Text1(0).MaxLength = 8
        ElseIf CodTipoMov = "ALB" Then
            cadList = "Albaranes tienda ALVIC"
            Text1(0).MaxLength = 8
        ElseIf CodTipoMov = "ALI" Then
            cadList = "Albaranes de garantias"
        End If
        lblTituloEst.Caption = cadList
        
        
        PonerImagenFirma
        
        'Referencia cliente
        Text1(13).Width = 4005
        Text1(13).MaxLength = 255
        
        
        'Iconitos de costes
        Me.cmdLineasCostes(0).Picture = frmPpal.imgListComun.ListImages(3).Picture
        Me.cmdLineasCostes(1).Picture = frmPpal.imgListComun.ListImages(4).Picture
        Me.cmdLineasCostes(2).Picture = frmPpal.imgListComun.ListImages(14).Picture
        cmdLineasCostes(3).Picture = frmPpal.imgListComun.ListImages(40).Picture
        cmdLineasCostes(4).Picture = frmPpal.ImgListPpal.ListImages(9).Picture
        cmdLineasCostes(5).Picture = frmPpal.ImgListPpal.ListImages(33).Picture
        Me.cmdLineasImpresion(0).Picture = frmPpal.imgListComun.ListImages(3).Picture
        Me.cmdLineasImpresion(1).Picture = frmPpal.imgListComun.ListImages(4).Picture
        Me.cmdLineasImpresion(2).Picture = frmPpal.imgListComun.ListImages(14).Picture
        Me.cmdLineasImpresion(3).Picture = frmPpal.imgListComun.ListImages(40).Picture
            
        
        If EsHistorico Then
            cmdLineasCostes(0).Enabled = False
            cmdLineasCostes(1).Enabled = False
            cmdLineasCostes(2).Enabled = False
            cmdLineasCostes(3).Enabled = False
            cmdLineasCostes(4).Enabled = False
            cmdLineasCostes(5).Enabled = False
        
            cmdLineasImpresion(0).Enabled = False
            cmdLineasImpresion(1).Enabled = False
            cmdLineasImpresion(2).Enabled = False
            cmdLineasImpresion(3).Enabled = False
        End If
        
    End If
    Me.Caption = cadList
    
    
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
    If vParamAplic.NumeroInstalacion = vbTaxco Then Label1(1).Caption = "Dpto."
    
    If vParamAplic.TieneCRM Then
        Label1(61).Caption = "Observaciones CRM"
    Else
        Label1(61).Caption = "Observaciones internas"
    End If
    ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
    If vEmpresa.TieneAnalitica Then
        'cmdAux(9).ToolTipText = "Buscar centro coste"
        imgBuscar2(9).ToolTipText = "Buscar centro coste"
        txtAux(9).Tag = "centro coste"
        Label1(51).Caption = "Centro coste"
    Else
        Label1(51).Caption = "Proveedor"
    End If
    imgBuscar2(9).Tag = -1
        
        
    '## A mano
    Me.FrameHco.visible = EsHistorico
    Me.FrameFacRec.visible = (CodTipoMov = "ART")
    
    
    
    'Aportacion a terminal
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        Label1(49).visible = hcoCodTipoM = "ALE" And vParamAplic.ctaAportacion <> ""
        Text1(40).visible = hcoCodTipoM = "ALE" And vParamAplic.ctaAportacion <> ""
    Else
        'Label1(49).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
        'Text1(40).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
        Label1(49).visible = False
        Text1(40).visible = False
    End If
    
    If Not EsHistorico Then
        NombreTabla = "scaalb"
        NomTablaLineas = "slialb" 'Tabla lineas de Albaranes
        Ordenacion = " ORDER BY codtipom, numalbar "
        
        If CodTipoMov = "ALV" Then
            Me.Caption = "Albaranes Clientes"
        ElseIf CodTipoMov = "ALM" Then
            Me.Caption = "Albaranes de Mostrador"
        ElseIf CodTipoMov = "ART" Then
            Me.Caption = "Albaranes Rectificativos"
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
 
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If AlbAvisoGenerado > 0 Then hcoCodMovim = AlbAvisoGenerado
        
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
    
        
        
        
    
    'Lo que era TEXTO para SAIL pasa a ser CERRADA para Euler
    'If vParamAplic.NumeroInstalacion = vbEuler Then cboTipoDat.List(2) = "CERRADA"
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    cboTipoDat.ListIndex = -1
    cboTipoImpr.ListIndex = -1
    cboTipopEnvio.ListIndex = -1
    Me.chkFacturar.Value = 0
    Me.chkFacturarKm.Value = 0
    Me.chkDocArchi.Value = 0
    Text3(0).Text = "BASE IMP."
    
    
    If SSTab1.TabVisible(2) Or SSTab1.TabVisible(3) Then LimpiarFichaTecnica False
    If SSTab1.TabVisible(5) Then ListView1.ListItems.Clear: Label1(63).Caption = "": Label1(64).Caption = "": LimpiaCostes
    If SSTab1.TabVisible(6) Then
        lwEulerLineas.ListItems.Clear
        lwEulerLineas.Tag = ""
    End If
    
    CargaCostesEuler True
    
    If InstalacionEsEulerTaxco Then
        'Me.imgFirmaRecep.Picture = LoadPicture("")
        imgFirmaRecep.visible = False
        imgFirmaRecep.Tag = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiaCostes()
Dim N As Integer

    If CostesMatriculas.Count = 0 Then Exit Sub

    For N = CostesMatriculas.Count To 1 Step -1
        CostesMatriculas.Remove N
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    AvisosAlbaranesEfectifivoPendienteFacturar False
    
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    AlbAvisoGenerado = 0   'por si acaso
    
    Set CostesMatriculas = Nothing
    
    If vParamAplic.TotalAlbaranSobreTablaTmp Then ejecutar "DELETE FROM tmptotalalbaran WHERE codusu = " & vUsu.Codigo, True
    
    
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(42).Text = RecuperaValor(CadenaSeleccion, 3)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 4) & "  " & RecuperaValor(CadenaSeleccion, 5)
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


Private Sub FrmArtEul_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = ""
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If EsCabecera Then 'Llama desde VerTodos del Form
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
            PonerCadenaBusqueda False
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 2), "0000000")
            
        ElseIf Val(imgBuscar2(9).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            'Llama desde Prismatico Direcciones/Departamentos   o de actuaciones
            Precio = CadenaDevuelta
            
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
    Indice = CByte(Me.imgFecha(0).Tag)
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
            motivo = "MOTIVO: " & RecuperaValor(CadenaSeleccion, 4)
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
         InsertarLineasFactu CadenaSeleccion, ""
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
        TieneMan = "0": devuelve = ""
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
        
        'cadena para UPDATE
        cadValuesU = " codclien=" & Text1(4).Text & ", coddirec=" & DBSet(Text1(12).Text, "T")
        cadValuesU = cadValuesU & ", codtipom=" & DBSet(Text1(30).Text, "T")
        cadValuesU = cadValuesU & ", fechavta='" & Format(Text1(1).Text, FormatoFecha) & "' "
        cadValuesU = cadValuesU & ", numalbar=" & Text1(0).Text & ", numline1=" & Me.cmdAux(0).Tag
        InsertarNSeries txtAux(1).Text, CadValues, cadValuesU, True
    End If
End Sub


Private Sub frmOC_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmOT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmProv2_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmTr_DatoSeleccionado(CadenaSeleccion As String)
    txtAnterior = CadenaSeleccion
End Sub

'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'Dim Indice As Byte
'
'End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    TerminaBloquear
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            Set frmC = New frmBasico2
'            frmC.DatosADevolverBusqueda = "0"
'            frmC.Show vbModal
            AyudaClientes frmC, Text1(4)
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
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text), True
                Indice = 12
             End If
             
        Case 3, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            If Index = 7 Then
                Indice = 27
            ElseIf Index = 8 Then
                Indice = 28
            Else
                Indice = Index
            End If
            
            txtAnterior = ""
            
'            Set frmTr = New frmAdmTrabajadores
'            frmTr.DatosADevolverBusqueda = "0"
'            frmTr.Show vbModal
            Set frmTr = New frmBasico2
            AyudaTrabajadores frmTr, Text1(Indice)
            Set frmTr = Nothing

            If txtAnterior <> "" Then
                Text1(Indice).Text = Format(RecuperaValor(txtAnterior, 1), "0000") 'Cod Trabajador
                Text2(Indice).Text = RecuperaValor(txtAnterior, 2) 'Nom Trabajador
                txtAnterior = Text1(Indice).Text
            End If
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
            End If
            If Trim(Text1(12).Text) = "" Then
                MsgBox "Debe seleccionar un obra para el cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            EsCabecera = False
            Set frmAc = New frmObraActua
            frmAc.DatosADevolverBusqueda = Text1(4).Text & "|" & Text1(12).Text & "|"
            frmAc.Show vbModal
            Set frmAc = Nothing
            Indice = 12
            'MandaBusquedaPrevia "codclien = " & Text1(4).Text & " AND coddirec = " & Text1(12).Text, False
    End Select
    
    If Indice > 0 Then
        If Indice = 12 Then
            PonerFoco Text1(15)
        Else
            PonerFoco Text1(Indice)
        End If
    End If
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
        If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
End Sub


Private Sub imgBuscar2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    Select Case Index
    Case 0
    
            
'            Set frmTr = New frmAdmTrabajadores
'            frmTr.DatosADevolverBusqueda = "0"
'            frmTr.Show vbModal
            Set frmTr = New frmBasico2
            AyudaTrabajadores frmTr, txtAux(14)
            Set frmTr = Nothing
            If txtAnterior <> "" Then
            
                txtAux(14).Text = Format(RecuperaValor(txtAnterior, 1), "0000") 'Cod Trabajador
                Text2(2).Text = RecuperaValor(txtAnterior, 2) 'Nom Trabajador
                txtAnterior = txtAux(14)
                PonerFoco txtAux(14)
            End If
    Case 12
        Set frmOC = New frmObraCapitulo
        frmOC.DatosADevolverBusqueda = "0|1|"
        frmOC.Show vbModal
        Set frmOC = Nothing
        If CadenaDesdeOtroForm <> "" Then
            txtAux(12).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Text2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtAnterior = txtAux(12)
            PonerFoco txtAux(12)
        End If
    
    Case 13
        Set frmOT = New frmObraOT
        frmOT.DatosADevolverBusqueda = "0|1|"
        frmOT.Show vbModal
        Set frmOT = Nothing
        If CadenaDesdeOtroForm <> "" Then
            txtAux(13).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Text2(13).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtAnterior = txtAux(13)
            PonerFoco txtAux(13)
        End If
    Case 9
         If vEmpresa.TieneAnalitica Then
            'centro de coste
            AbrirForm_CentroCoste
        Else
'            Set frmProv2 = New frmComProveedores
'            frmProv2.DatosADevolverBusqueda = "1"
'            frmProv2.Show vbModal
            Set frmProv2 = New frmBasico2
            AyudaProveedores frmProv2, txtAux(9).Text
            Set frmProv2 = Nothing
            PonerFoco txtAux(9)
            If CadenaDesdeOtroForm <> "" Then
                txtAux(9).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(9).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        End If
    
    
        
        txtAnterior = txtAux(9)
    End Select
    CadenaDesdeOtroForm = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte
Dim hora As String


    



   If Modo = 2 Or Modo = 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   hora = ""
   Indice = Index + 1
   If Index = 44 Then
        If Text1(Indice).Text <> "" Then hora = Trim(Mid(Text1(Indice).Text, 11))
        If Not EsHoraOK(hora) Then hora = ""
    End If
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   
  
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   If Index = 44 Then
     If hora = "" Then hora = Format(Now, "hh:mm:ss")
     If hora = "00:00:00" Then hora = Format(Now, "hh:mm:ss")
        If Len(Text1(Indice).Text) = 10 Then Text1(Indice).Text = Text1(Indice).Text & " " & hora
   End If
   PonerFoco Text1(Indice)
   
End Sub


Private Sub imgFirmaRecep_DblClick()
      If Modo <> 2 Then Exit Sub
    If Me.imgFirmaRecep.Tag = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    LanzaVisorMimeDocumento Me.hwnd, imgFirmaRecep.Tag
    Espera 0.35
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgGeolocalizacion_Click()
    If Text1(48).Text <> "" And Text1(49).Text <> "" Then
        txtAnterior = TransformaComasPuntos(Text1(48).Text) & "," & TransformaComasPuntos(Text1(49).Text)
        AbrirGeolocalizacion txtAnterior
        
        txtAnterior = ""
    End If
End Sub

Private Sub imgNull_Click(Index As Integer)
    If Index = 0 Then
        Me.cboTipoDat.ListIndex = -1
    Else
        Me.cboTipoImpr.ListIndex = -1
    End If
End Sub

Private Sub imgObserva_Click(Index As Integer)
Dim txtAsociado As Integer
Dim OpcionObserva As Byte
    '0. Nada
    '1. Abrir
    '2. Abrir y modificar
    'aBRE LAS OBSERVACIONES DE LA LINEA
    OpcionObserva = 0
    txtAsociado = 16 + Index 'Son el 16 y el 17
    
    
    If Modo <> 5 Then
        If txtAux(txtAsociado).Text <> "" Then OpcionObserva = 1
    Else
        OpcionObserva = 1
        If ModificaLineas > 0 Then OpcionObserva = 2
    End If
    
    If OpcionObserva > 0 Then
        'Abrir
                If EsHistorico Then OpcionObserva = 1
                CadenaDesdeOtroForm = txtAux(txtAsociado).Text
               
                frmFacAlbObser.Modificar = OpcionObserva = 2
                frmFacAlbObser.Text1 = CadenaDesdeOtroForm
                frmFacAlbObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If OpcionObserva = 2 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then txtAux(txtAsociado).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
     
        
    End If
End Sub

Private Sub Label2_Click()

End Sub



Private Sub ListView2_DblClick()

    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    
    
  
    
    
    If ListView2.SelectedItem.Text <> "ALC" Then
        If ListView2.SelectedItem.Text <> "FAC" Then
            If ListView2.SelectedItem.Text <> "PED" Then Exit Sub
        End If
    End If
    
    
    If ListView2.SelectedItem.Text = "ALC" Then
    
      'IT.Tag = "numalbar =" & DBSet(Rs!NUmAlbar, "T") & " AND  fechaalb =" & DBSet(Rs!FechaAlb, "F") & " AND codprove =" & Rs!Codprove
       With frmComEntAlbaranSA
            .hcoCodMovim = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            .hcoFechaMovim = RecuperaValor(ListView2.SelectedItem.Tag, 2)
            .hcoCodProve = RecuperaValor(ListView2.SelectedItem.Tag, 3)
            .EsHistorico = False
            .Show vbModal
        End With
    
    ElseIf ListView2.SelectedItem.Text = "PED" Then
        'PEDIDOS
            frmComEntPedidosSa.MostrarDatos = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            frmComEntPedidosSa.EsHistorico = False
            frmComEntPedidosSa.Show vbModal
    
    Else
    
        
        'IT.Tag = "numfactu =" & DBSet(Rs!Numfactu, "T") & " AND  fecfactu=" & DBSet(Rs!FecFactu, "F") & " AND codprove =" & Rs!Codprove
         With frmComHcoFacturSA
            .hcoCodMovim = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            .hcoFechaMovim = RecuperaValor(ListView2.SelectedItem.Tag, 2)
            .hcoCodProve = RecuperaValor(ListView2.SelectedItem.Tag, 3)
            .Show vbModal
        End With
    End If
    
    
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Albaran
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Albaran

    If vParamAplic.TotalAlbaranSobreTablaTmp Then CargaTablaTmpTotales


    If vParamAplic.NumeroInstalacion = vbTaxco And hcoCodTipoM = "ALO" And Me.chkFacturar.Value = 0 Then
        cmdOrdenTallerTaxco_Click
    Else
        BotonImprimir 45, False '45: Informe de Albaranes
    End If
End Sub


Private Sub mnLineaAccion_Click(Index As Integer)
    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Index <> 0 Then
        If Me.Data2.Recordset.EOF Then Exit Sub
        If SSTab1.Tab <> 0 Then Exit Sub
        
    End If
    
    If True Then
    
        If Not PuedeMeterLineas Then Exit Sub
    
        'Desde F1, f2 y f3
        If Index <> 2 Then
                    
        
        
            BotonMtoLineas 0, "Albaranes"
            
            If Index = 0 Then
                BotonAnyadirLinea False
            Else
                BotonModificarLinea
            End If
        Else
            BotonEliminarLinea
        End If
    End If
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




Private Sub optEule_R_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub optEuler_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optEuler_LostFocus(Index As Integer)
    If Index = 1 Then PonerFocoOBj txtEuler(0)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'    Me.Label1(35).visible = Me.SSTab1.Tab = 0
'    Me.Text2(16).visible = Me.SSTab1.Tab = 0
'    Me.Label1(51).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And SSTab1.Tab = 0
'    Me.Text2(9).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And Me.SSTab1.Tab = 0
'
    If SSTab1.Tab = 4 Then
        If PreviousTab <> 4 And Modo = 2 Then
            'Si ya han sido cargados
            If ListView2.ListItems.Count = 0 Then CargaCostesEuler False
        End If
    End If
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
Dim ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
     If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
    
        If Text1(Index).Text = "" Then
            ind = -1
            Select Case Index
            Case 3
                ind = 3
            Case 4
                ind = 0
            Case 6
                ind = 1
            Case 9
                ind = 6
            Case 12
                ind = 2
            Case 17
                ind = 5
            Case 14
                ind = 4
            Case 27, 28, 29
                ind = Index - 20
            End Select
            If ind >= 0 Then
                PulsadoMas2 = True
                PulsarTeclaMas True, ind
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
    If txtAnterior = Text1(Index).Text Then Exit Sub
          
    
          
          
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 41, 44 'Fecha Albaran,fecenvio
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                    
                
        Case 45
                If Text1(Index).Text <> "" Then
                    If IsDate(Text1(Index).Text) Then
                        Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy hh:nn:ss")
                    Else
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
        Case 3, 27, 28 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
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
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
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
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
'            If (Modo = 4) Then
'                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
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
        Case 42
            If Text1(42).Text = "" Then
                Text2(1).Text = ""
            Else
                PonerCampoActuacion
                If Text1(42).Text = "" Then PonerFoco Text1(42)
            End If
            
        
        Case 48, 49
            
            If Text1(Index).Text <> "" Then
                If InStr(1, Text1(Index).Text, ".") > 0 Then Text1(Index).Text = Replace(Text1(Index).Text, ".", ",")
                If IsNumeric(Text1(Index).Text) Then
                    
                    Text1(Index).Text = Format(Text1(Index).Text, "0.000000")
                Else
                    Text1(Index).Text = ""
                End If
            End If
    End Select
End Sub




Private Sub HacerBusqueda()
Dim cadB As String
Dim C As String


    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    C = ""
    If InstalacionEsEulerTaxco Then
        If Not EsHistorico Then
            C = BuscaEnBDFicha
        Else
            C = ""
        End If
        'Cadb siempre llevara codtipom=hcodtipom
        If C <> "" Then
            If cadB <> "" Then cadB = cadB & " AND "
            cadB = cadB & C
        End If
    End If
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB, False
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        If Me.EsHistorico = False Then
            cadB = cadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
        End If
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String, Dpto As Boolean)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera Then
        Cad = Cad & ParaGrid(Text1(30), 10, "Tipo Alb.")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Albaran")
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha Ped.")
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
        Precio = ""
        If Dpto Then
            'DEMPARTAMENTO
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            Else
                Titulo = "Obras Cliente: "
                Desc = "Obra"
                
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    Titulo = "Dptos Cliente: "
                    Desc = "Dpto."
                End If
            End If
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
            Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||55·"
            tabla = "sdirec"
            devuelve = "0|1|"
    
        Else
            'Actuacion
            Titulo = "Actuaciones en obra: " & Text1(4).Text & " - " & Text1(5).Text & " // " & Text1(12).Text
            Cad = Cad & "Actuacion " & "|sactuaobra|actuacion|T||25·"
            Cad = Cad & "Fec. Ini. |sactuaobra|fechaini|F||15·"
            Cad = Cad & "Obs|sactuaobra|observa|T||55·"
            tabla = "sactuaobra"
            devuelve = "0|1|"
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
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        
        If Not EsCabecera Then
            'Dpto o actuacion
            If Precio <> "" Then
                If Dpto Then
                    Text1(12).Text = Format(RecuperaValor(Precio, 1), "000")
                    Text2(12).Text = RecuperaValor(Precio, 2)
                Else
                    Text1(42).Text = RecuperaValor(Precio, 1)
                    Text2(1).Text = RecuperaValor(Precio, 2)
                End If
                Precio = ""
            End If
            If Dpto Then
                PonerFoco Text1(12)
            Else
                PonerFoco Text1(42)
            End If
        End If

    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda(DesdeFaturacion As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Not DesdeFaturacion Then MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            LimpiarCampos
            
            FocoBuscar
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

    PonerForaGrid

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim B As Boolean
Dim Movi As String
    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    
    
    
    
    
     'Si es un Albaran de Ticket visualizamos unos datos y sino otros
    B = (Data1.Recordset!EsTicket = 1)
    Me.Toolbar1.Buttons(11).Enabled = (Not B) And (Not EsHistorico)
        
        
    Movi = hcoCodTipoM
        
    If EsHistorico Then
        If Modo = 2 Then
            If Not Data1.Recordset.EOF Then Movi = Data1.Recordset!codtipom
            
            SSTab1.TabVisible(3) = Movi = "ALR"
            SSTab1.TabVisible(2) = Movi = "ALO" Or Movi = "ALE"
            FrameOT_Euler.visible = Movi = "ALO"
        End If
    End If
        

    If Movi <> "ALR" Then
        'sem. entrega pedido
        Label1(12).visible = Not B
        Text1(2).visible = Not B
        'num oferta
        Text1(23).visible = Not B And Movi <> "ALR"
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
    
    
    If SSTab1.TabVisible(2) Then PonerCamposFicha
    If SSTab1.TabVisible(3) Then PonerCamposFichaReparacion
    If SSTab1.TabVisible(5) Then PonerTareasAsociadas
    
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
    Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "straba", "nomtraba", "codtraba")
    Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    

    'Septiembre. Si tipomimp es NULL NO debe poner valor en el combo
    If IsNull(Data1.Recordset!tipoimp) Then cboTipoImpr.ListIndex = -1
    If IsNull(Data1.Recordset!origdat) Then Me.cboTipoDat.ListIndex = -1
    If IsNull(Data1.Recordset!codinter) Then Me.cboTipopEnvio.ListIndex = -1
    



    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "straba", "nomtraba", "codtraba")
        Text2(33).Text = PonerNombreDeCod(Text1(33), conAri, "sincid", "nomincid", "codincid")
    End If
    
    PonerCampoActuacion
    
    CargaCostesEuler False
    
    
    PonerImagenFirma
    
    'Me.lblTipo(0).Caption = PonerCampoTipoSail(43)
    'Me.lblTipo(1).Caption = PonerCampoTipoSail(44)
    CalcularDatosFactura
    
    
    'Lo mpongo al final pq utiliza el total albaran
    If SSTab1.TabVisible(6) Then PonerLineasAlbaranEULER
    
    
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
        
    If EsHistorico Then
        If Modo < 2 And InstalacionEsEulerTaxco Then
            SSTab1.TabVisible(3) = False
            SSTab1.TabVisible(2) = False
        End If
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    If InstalacionEsEulerTaxco Then BloquearFicha Modo = 0 Or Modo = 2 Or Modo = 5
    
    
    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
    
    i = 0
    If InstalacionEsEulerTaxco Then
        'para EULER
        i = 1
        B = True
        If Modo = 1 Then
            B = False
        Else
            If Modo = 3 And hcoCodTipoM = "ALR" Then B = False
        End If
    End If
    If i = 0 Then
        BloquearTxt Text1(0), Modo <> 1, True
    Else
        'EULER EN MODO
        BloquearTxt Text1(0), B, True
    End If
    B = (Modo <> 1)
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
    imgNull(0).visible = B
    imgNull(1).visible = B
    cboTipoDat.Enabled = B
    cboTipoImpr.Enabled = B
    cboTipopEnvio.Enabled = B
    imgObserva(0).visible = Modo > 1
    imgObserva(1).visible = Modo > 1
    
    imgGeolocalizacion.visible = Modo = 2
    
    If CodTipoMov = "ALO" And vParamAplic.NumeroInstalacion = vbTaxco Then
        If Modo = 0 Then
            Me.FrameTAXCO.visible = False
        Else
            FrameTAXCO.visible = True
            FrameTAXCO.Enabled = Modo = 1
            cmdOrdenTallerTaxco.visible = Not EsHistorico And Modo = 2
            
            If Modo <> 1 Then
                Text5(0).BackColor = vbWhite
                Text5(1).BackColor = vbWhite
            End If
        End If
    End If
    
    
    'If vParamAplic.NumeroInstalacion = vbTaxco And b Then BloquearTxt Text1(14), Modo <> 1, False

    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    B = True
    
    If Modo = 5 Then B = ModificaLineas = 0
    
    'BloquearTxt Text2(9), b
    BloquearTxt txtAux(9), B
    For i = 12 To 17
        BloquearTxt txtAux(i), B
    Next i
    
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    Me.imgFecha(0).Enabled = B
    Me.imgFecha(43).Enabled = B
    Me.imgFecha(40).Enabled = B
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(1).visible = False
    Me.imgBuscar(7).Enabled = (Modo = 1)
              
    'If vParamAplic.NumeroInstalacion = vbTaxco And b Then imgBuscar(4).Enabled = False
    B = (Modo = 0 Or Modo = 2)
    For i = 0 To cmdLineasCostes.Count - 1
        cmdLineasCostes(i).Enabled = B
    Next i
    
    
              
              
''    'Modo Linea de Albaranes
''    '- poner visible ampliacion linea
''    BloquearTxt txtAux(16), True
''    '- poner visible nombre proveedor linea
''    BloquearTxt Text2(9), True
    SSTab1_Click 0
      
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
Dim B As Boolean
Dim devuelve As String
Dim i As Byte
    On Error GoTo EDatosOK

    DatosOk = False
    
    'Asignarle el valor del Combo Tipo de Movimiento al texto oculto text1(30)
'    Text1(30).Text = ObtenerCodTipom
    
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
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
    
        
    
    If InstalacionEsEulerTaxco Then
        
        If Modo = 3 Then
            'En euler, los albaranes de reparacion pueden introducir MANUALMENTE el numero
            If Me.hcoCodTipoM = "ALR" Then
                If Me.Text1(0).Text <> "" Then
                    devuelve = "codtipom = " & DBSet(hcoCodTipoM, "T") & " AND numalbar  "
                    devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", devuelve, Text1(0).Text, "N")
                    If devuelve <> "" Then
                        MsgBox "Ya existe el albaran " & Me.hcoCodTipoM & Text1(0).Text, vbExclamation
                        B = False
                    End If
                End If
                
            End If
    
        End If
        
        If Text1(44).Text = "" Then
            MsgBox "Fecha auxiliar obligatoria", vbExclamation
            B = False
        End If
    
        
        'Si el cliente es de vARIOS
        ' NO debe dejar si no esta identificada la venta con un  NIF validoo
        
        If Me.Text1(6).Enabled Then
            If Not Me.Text1(6).Locked Then
            
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    BuscaChekc = ""
                    If Len(Text1(5).Text) < 15 Then
                        BuscaChekc = "Nombre incorrrecto(1)"
                    Else
                        If InStr(1, Text1(5).Text, "VARIOS ") Then BuscaChekc = "Nombre incorrrecto(2)"
                    End If
                    If BuscaChekc <> "" Then
                        MsgBox BuscaChekc, vbExclamation
                        BuscaChekc = ""
                        B = False
                    End If
                    
                End If
            
            
                If Text1(6).Text = "22222222J" Or Text1(6).Text = "222222J" Then
                    MsgBox "HAZ UN TICKET !!!!!!!!!!!", vbCritical
                    Exit Function
                End If
                If Not Comprobar_NIF(Text1(6).Text) Then
                    BuscaChekc = "El NIF: " & Text1(6).Text & "   no parece correcto."
                    If vParamAplic.NumeroInstalacion = vbTaxco Then
                        MsgBox BuscaChekc, vbExclamation
                        Exit Function
                    Else
                        BuscaChekc = BuscaChekc & vbCrLf & vbCrLf & "¿Continuar?"
                        If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Function
                    End If
                End If
                
                
                
                
                
                
                
                
            End If
        End If
            
    End If
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
    
    
        'TAXCO.   Forma de pago =2 SOLO si el cliente tiene CREDITO
        If Val(Text1(14).Text) = 2 Then
            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codforpa", "codclien", Text1(4).Text, "N")
            If Val(devuelve) <> 2 Then
                MsgBox "El cliente NO tiene crédito. ", vbExclamation
                Text1(14).Text = "1"
                Text2(14).Text = "EFECTIVO"
                Exit Function
            End If
        End If
    
    
    
        If hcoCodTipoM = "ALO" And Not EsHistorico Then
    
            If Val(txtTaxco(7).Text) <= 0 Then
                MsgBox "Introduzca kilometros", vbExclamation
                B = False
            End If
            
            If Trim(txtTaxco(0).Text) = "" Then
                MsgBox "Introduzca matrícula", vbExclamation
                B = False
            End If
        End If
        
        If hcoCodTipoM = "ALE" And Not EsHistorico Then txtTaxco(7).Text = "0"
        
        If vUsu.Nivel2 > 1 Then
            'Si cambia el codigo de trabjador.. y no tiene permiso....
            If Modo = 4 Then
                If Val(Text1(3).Text) <> Val(Data1.Recordset!CodTraba) Then
                    MsgBox "NO puede cambiar el trabajador asignado", vbExclamation
                    B = False
                End If
            End If
        End If
    End If
    
    
    If B Then
        B = False
        EsClienteBloqueado2 Text1(4).Text, 0, True, B
        If B Then
            'Solo efectivo
            BuscaChekc = DevuelveDesdeBDNew(1, " sforpa", "tipforpa", "codforpa", Text1(14).Text)
            If BuscaChekc <> "0" And BuscaChekc <> "6" Then
                MsgBox "Cliente bloqueado.  Forma pago INVALIDA(solo efectivo o tarjeta) ", vbExclamation
                PonerFoco Text1(14)
                B = False 'Para que de false el datosok
            End If
        Else
            B = True
        End If

    End If
    
    
    
    
    
    
    DatosOk = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

 
Private Function DatosOkLinea(ByRef vCStock As CStock) As Boolean
Dim B As Boolean
Dim i As Byte
Dim Aux As String

    On Error GoTo EDatosOkLinea
    txtAux(10).Text = 1 'en sail los bultos a pelo
    DatosOkLinea = False
    
    
        'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(6), txtAux(7), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(8).Text Then txtAux(8).Text = Aux
    
    
    
    B = True
    'De los datos basicos NINGUNO puede ser nulo
    For i = 0 To 8
        'Debug.Print i & " " & txtaux(i).Tag
        If txtAux(i).Text = "" And i <> 5 Then
            'El campo 5= origpre puede ser nulo (en alb.repar)
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            B = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
    
    
    

    
    
    If txtAux(9).Text = "" Or Text2(9).Text = "" Then
    
        If vEmpresa.TieneAnalitica Then
            MsgBox "Centro de coste incorrecto.", vbExclamation
        Else
            MsgBox "Proveedor incorrecto", vbExclamation
        End If
        PonerFoco txtAux(9)
        Exit Function
    End If
    
    
    
    
    
    
    
    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        B = vCStock.MoverStock(False, False)
    End If
    DatosOkLinea = B
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
  '  If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub Text5_GotFocus(Index As Integer)
    ConseguirFoco Text5(Index), Modo
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    If Modo <> 1 Then Exit Sub
    
    If Index = 1 Then
        
        Me.txtTaxco(7).Text = Text5(Index).Text
    Else
        Me.txtTaxco(Index).Text = Text5(Index).Text
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    HacerButton Button.Index

End Sub

Private Sub HacerButton(IndiceBoton As Integer)

    Select Case IndiceBoton
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos
            
        Case 5: mnNuevo_Click 'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 10: mnLineas_Click  'Lineas
        Case 11:
            
                    If Modo = 5 Then
                        'Insertar intercalando
                        BotonAnyadirLinea True
                    Else
                        BotonNSeries 'Control Nº Series
                    End If
            
            
            
            
        Case 12 'Generar Factura Mostrador
                'o Factura Rectificativa (FRT)
          
            
                'Septiebmre2009
                If Data2.Recordset Is Nothing Then Exit Sub
                If Data2.Recordset.RecordCount = 0 Then
                    MsgBox "No tiene lineas de albarán", vbExclamation
                    Exit Sub
                End If
            
            
                'EN EULER no dejamos facturar los albarenes internos
                If hcoCodTipoM = "ALI" Then
                    If InstalacionEsEulerTaxco Then Exit Sub
                End If
                'procedimiento normal
                If Data1.Recordset!codtipom = "ART" Then
                    'Comprobar nº serie de las facturas rectificativas
                    DevolverNumSeries
                End If
                    
                If InstalacionEsEulerTaxco Then
                    If lwEulerLineas.Tag <> "" Then
                        MsgBox lwEulerLineas.Tag, vbExclamation
                        Exit Sub
                    End If
                    
                    
                    'No puedo fcturar ALD y ALB
                    'Vienen del traspaso de alvic. Se facturan final de mes
                    
                    If vParamAplic.NumeroInstalacion = vbTaxco Then
                        
                        If (hcoCodTipoM = "ALD" Or hcoCodTipoM = "ALB") Then
                            TaxcoFacurarUnAlbaranALVIC_Facturado = ""
                            If vUsu.Nivel = 0 And DBLet(Data1.Recordset!observa02, "T") <> "" Then
                                TaxcoFacurarUnAlbaranALVIC_Facturado = Text1(1).Text & "|" & Data1.Recordset!observa02 & "|" & Data1.Recordset!codforpa & "|"
                            Else
                                MsgBox "Facturar fin de mes", vbExclamation
                                Exit Sub
                            End If
                        End If
                    End If
                        
                        
                        
                    If vParamAplic.NumeroInstalacion = vbEuler Then
                        BuscaChekc = "  codtipoa='" & Text1(30).Text & "' AND numalbar"
                        BuscaChekc = DevuelveDesdeBD(conAri, "numproyec", "sproyectolin", BuscaChekc, Text1(0).Text)
                        If BuscaChekc <> "" Then
                            MsgBox "Albaran vinculado al proyecto: " & BuscaChekc, vbExclamation
                            BuscaChekc = ""
                            Exit Sub
                        End If
                    End If
                End If
                 
                 
                'Comprobar que esta marcada para facturar
'                If Data1.Recordset!codTipoM <> "ALM" Then Exit Sub
                If Me.chkFacturar.Value = 1 Then
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    
                    'Facturacion de Albaran de Mostrador
                    frmListadoPed.codClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
                    frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
                    AbrirListadoPed (222)
                    davidNumalbar = 0
                    PosicionarDataTrasEliminar
                Else
                    BuscaChekc = "El Albaran no esta marcado para facturar"
                    If vParamAplic.NumeroInstalacion = vbTaxco And EsTipoAlbaranFacturarInmediatoTaxco Then BuscaChekc = "El albaran/OT NO está cerrado"
                    MsgBox BuscaChekc, vbInformation
                    BuscaChekc = ""
                End If
            
        Case 13
            'DAVID
            'Marca los albaranes que esten como NO facturar a facturar
            If vParamAplic.NumeroInstalacion = vbTaxco And Not EsHistorico And EsTipoAlbaranFacturarInmediatoTaxco Then
                MsgBox "No se pueden cerrar desde aqui", vbExclamation
                Exit Sub
            End If
            
            
            Screen.MousePointer = vbHourglass
            MarcarAlbaranes
            Screen.MousePointer = vbDefault
            
        Case 14
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                VerHistoricoTaxco
            Else
                If Modo <> 2 Then Exit Sub
                If vUsu.Nivel > 1 Then
                    MsgBox "No tiene pemiso para realizar la operacion", vbExclamation
                Else
                    HacerCambioTipoAlbaran
                End If
            End If
            'Cuando eran portes
            'If vParamAplic.TipoPortes <> 1 Then Exit Sub
            'BotonImprimir 45, True
            
            
        Case 16: mnImprimir_Click 'Imprimir Albaran
        Case 17: mnSalir_Click   'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (IndiceBoton - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim B As Boolean
    PonerOpcionesMenuGeneral Me
    
     B = (Modo >= 3) Or Modo = 1
    If Modo = 5 Then
            B = (ModificaLineas = 0)
            Toolbar1.Buttons(11).Image = 34 '.Buttons(11).Image = 26
            Toolbar1.Buttons(11).ToolTipText = "Insertar intercalando"
            
            
    Else
            'b=modo=2
            B = B And Not EsHistorico
            Toolbar1.Buttons(11).Image = 33
            Toolbar1.Buttons(11).ToolTipText = "Nº de serie"
            
    End If
    Toolbar1.Buttons(11).Enabled = B
    
    
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

  
'DesdeRecuperaParaRectificativa:  Para que no inserte el punto verde
Private Function InsertarLinea(numlinea As String, DesdeRecuperaParaRectificativa As Boolean) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim SQL As String
Dim vWhere As String
Dim B As Boolean
Dim vCStock As CStock
Dim ImpReciclado As Single
Dim DentroTRANS As Boolean
Dim ImporteRecicladoSigaus As Currency
Dim ImporteRecicladoSigausT As String

    InsertarLinea = False
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    
    
    If LineaIntercalar = 0 Then
        'INSERCION NORMAL
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
     
    Else
                                              'por si acaso lleva tasa reciclaje
        SQL = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 2 WHERE " & vWhere & " and numlinea >= " & LineaIntercalar
        SQL = SQL & " order by numlinea desc " 'Para que empieza por las ultimas
        conn.Execute SQL
        numlinea = LineaIntercalar
    End If
    
    
    
    SQL = ""
    Me.cmdAux(0).Tag = numlinea 'Aqui almaceno el Nº linea que acabo de Insertar
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", numlinea) Then Exit Function
    
    If DatosOkLinea(vCStock) Then 'Lineas de Albaranes
        'Inserta en tabla "slialb"
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos,precioar, dtoline1, dtoline2, importel, origpre,codprovex,numlote,codccost "
        'SAIL
        'codcapit codtipor CodTraba precoste
        SQL = SQL & ",codcapit, codtipor, CodTraba, precoste)"
        SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(txtAux(16).Text, "T") & ", "
        '- cantidad,numbultos
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", " & DBSet(txtAux(10).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(5).Text, "T", "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            '- codprove,numlote,codccost
            SQL = SQL & "0," & DBSet(txtAux(11).Text, "T", "S") & "," & DBSet(txtAux(9).Text, "T", "S")
        Else
            '- codprove,numlote,codccost
            SQL = SQL & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(11).Text, "T", "S") & "," & ValorNulo
        End If
        'codcapit codtipor CodTraba precoste
        SQL = SQL & "," & DBSet(txtAux(12).Text, "N", "S") & ","
        SQL = SQL & DBSet(txtAux(13).Text, "T", "S") & ","
        SQL = SQL & DBSet(txtAux(14).Text, "N", "S") & ","
        SQL = SQL & DBSet(txtAux(15).Text, "N", "S") & ")"
        '-
'        sql = sql & DBSet(txtAux(11).Text, "T", "S") & ")"
     Else
        Exit Function
     End If
    
    If SQL <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute SQL
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        B = vCStock.ActualizarStock(False, True)
        
        
        
        
        'Si ha actualizado el sctock
        If B Then
            If ClienteConTasaReciclado And Not DesdeRecuperaParaRectificativa Then
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                    'Insertamos la linea del reciclado
                    ImporteRecicladoSigausT = "preciove"
                    vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T", ImporteRecicladoSigausT)
                    If vParamAplic.NumeroInstalacion = vbTaxco Then
                        If ImporteRecicladoSigausT = "" Then ImporteRecicladoSigausT = "0"
                        ImpReciclado = CCur(ImporteRecicladoSigausT)
                    End If
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,  precioar,"
                    SQL = SQL & "dtoline1, dtoline2, importel, origpre,codccost) "
                    SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
                    SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
                    'Importe linea
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A',"
                    If vEmpresa.TieneAnalitica Then
                        SQL = SQL & DBSet(txtAux(9).Text, "N")
                    Else
                        SQL = SQL & "null"
                    End If
                    SQL = SQL & ")"
                    conn.Execute SQL
                        
                    
                End If 'articulo con sunida reciclado
            End If  'Cliente con tasa reciclado
        End If 'ok actualiza stock
        
        
    
    End If
    Set vCStock = Nothing
    
    
    
    If B Then
        conn.CommitTrans
        InsertarLinea = True
        
        DatosObservaciones SQL, 0, CInt(numlinea)
        
        AlmacenLineas = CInt(txtAux(0).Text)
        
        ' ---- [15/09/2009] (LAURA)
'        'Miramos en los descuentos
'        'Hacer sdesca
'        ElArticulo = txtAux(1).Text
'        DescuentosCantidad ElArticulo
        ' ----
        
    Else
        conn.RollbackTrans
         InsertarLinea = False
    End If
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

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    '## LAURA 15/11/2006
    'si se ha modificado el articulo eliminar de la smoval y reestablecer stock
    'Inicilizar la clase para Actualizar los stocks
    
    
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    '#### LAURA 15/11/2006
    conn.BeginTrans
        
    If DatosOkLinea(vCStock) Then
        
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        B = InicializarCStock(vCStock, "E")
        If B Then
            B = vCStock.DevolverStock2 'eliminamos de smoval y devolvemos stock valores anteriores
            If B Then
                'si se ha modificado el articulo
                If CStr(Data2.Recordset!codArtic) <> txtAux(1).Text Then
                    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
                    SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
                    SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
                    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                    SQL = SQL & " AND numalbar=" & Data1.Recordset!Numalbar & " AND numline1=" & Data2.Recordset!numlinea
                    conn.Execute SQL
                End If
            End If
            'ahora leemos los valores nuevos
            If B Then B = InicializarCStock(vCStock, "S")
            'insertamos en smoval y actualizamos stock a los valores nuevos
            vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
            If B Then B = vCStock.ActualizarStock(False, True)
    
            'actualizar la linea de Albaran
            If B Then
                SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
                SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(txtAux(16).Text, "T") & ", "
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
                SQL = SQL & "numlote=" & DBSet(txtAux(11).Text, "T", "S") & ","
                
                
                'SAIL
                'codcapit codtipor CodTraba precoste
                SQL = SQL & "codcapit= " & DBSet(txtAux(12).Text, "N", "S") & ", " 'precio
                SQL = SQL & "codtipor= " & DBSet(txtAux(13).Text, "T", "S") & ", "
                SQL = SQL & "CodTraba= " & DBSet(txtAux(14).Text, "N", "S") & ", "
                SQL = SQL & "precoste= " & DBSet(txtAux(15).Text, "N", "S") & " " 'Importe
                
                
                
                SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
                conn.Execute SQL
                
                
                'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    
                    If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       
                       If vParamAplic.NumeroInstalacion = vbTaxco Then NumRegElim = -1   'Para que no devuelve nada
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
        
        
        DatosObservaciones SQL, 1, CInt(Data2.Recordset!numlinea)
        
        ' ---- [15/09/2009] (LAURA)
'        If txtAux(1).Text = Data2.Recordset!codArtic Then
'            ElArticulo = Data2.Recordset!codArtic
'        Else
'            'Son distintos. Que recalcule todo
'            ElArticulo = ""
'        End If
'        DescuentosCantidad ElArticulo
        ' ----
        
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

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
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
            vDataGrid.Columns(i).Width = 1600
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Desc. Artículo"
            vDataGrid.Columns(i).Width = 3500

            i = 6
            vDataGrid.Columns(i).visible = False
            i = 7
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 810
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            'SAIL. ESTA NO ESTA
            'i = 8
            'vDataGrid.Columns(i).Caption = "Bultos"
            'vDataGrid.Columns(i).Width = 650
            'vDataGrid.Columns(i).Alignment = dbgRight
                
            i = i + 1 '8
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
            vDataGrid.Columns(i).Caption = "Importe"
            vDataGrid.Columns(i).Width = 1100
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            
            'SAIL. REsot a visible a false
            i = i + 1
            Do
                
                vDataGrid.Columns(i).visible = False
'                                    If vEmpresa.TieneAnalitica Then
'                                        i = i + 1
'                                        vDataGrid.Columns(i).Caption = "CCost"
'                                        vDataGrid.Columns(i).Width = 680
'                                        vDataGrid.Columns(i).Alignment = dbgRight
'                                    Else
'                                        i = i + 1
'                                        vDataGrid.Columns(i).Caption = "Prov"
'                                        vDataGrid.Columns(i).Width = 680
'                                        vDataGrid.Columns(i).Alignment = dbgRight
'
'                                        '- nombre proveedor
'                                        i = i + 1
'                                        vDataGrid.Columns(i).visible = False
'                            '            vDataGrid.Columns(i).Caption = "Nom. prove"
'                            '            vDataGrid.Columns(i).Width = 2100
'                                    End If
            
                i = i + 1
            Loop Until i > vDataGrid.Columns.Count - 1
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

    imgBuscar2(0).visible = visible
    imgBuscar2(12).visible = visible
    imgBuscar2(13).visible = visible
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To 8
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        imgBuscar2(9).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
            
        Else 'Vamos a modificar
            For i = 0 To 8
                If i < 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                ElseIf i >= 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 4).Text
               
                End If
                txtAux(i).Locked = False
            Next
            
            If False Then
'                ElseIf i > 10 Then
'                    ' ---- [19/10/2009] [LAURA] : centro de coste si hay conta analitica
'                    If vEmpresa.TieneAnalitica Then
'                        txtAux(i).Text = DataGrid1.Columns(i + 4).Text
'                    Else
'                        txtAux(i).Text = DataGrid1.Columns(i + 5).Text
'                    End If
'                End If
            End If
            
        End If
        
        cmdAux(0).Enabled = True
        cmdAux(1).Enabled = True
'        cmdAux(9).Enabled = True
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
        '#####
'        'Bloquear campo numbultos q es calculado
'        BloquearTxt txtAux(10), True
        
        ' ---- [20/10/2009] [LAURA] : añadir centro de coste
'        BloquearTxt txtAux(9), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        i = 0
        If InstalacionEsEulerTaxco Then
                
        Else
            If (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica <> 2) Then i = 1
        End If
        BloquearTxt txtAux(9), i = 1
        
        'Me.cmdAux(9).Enabled = Not txtAux(9).Locked
        'Me.cmdAux(9).visible = Me.cmdAux(9).Enabled
        imgBuscar2(9).visible = Not txtAux(9).Locked


        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To 8
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
      
        
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
        'Bultos    ESTE NO ESTA EN el grid para SAIL
        'Precio
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 20
        'txtAux(4).Left = txtAux(10).Left + txtAux(10).Width + 20
        txtAux(4).Width = DataGrid1.Columns(8).Width - 20
        
        'OP, Dto1, Dto2, Precio, (codProve/codccost)
        For i = 5 To 8
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 20
            txtAux(i).Width = DataGrid1.Columns(i + 4).Width - 20
        Next i
        
        
        '- numlote
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To 8
            txtAux(i).visible = visible
        Next i
        
        If vUsu.Nivel2 = 2 And visible Then
            For i = 4 To 8
                BloquearTxt txtAux(i), True
            Next
        End If
        
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
'        cmdAux(9).visible = visible
    End If
End Sub


Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer
    If txtAux(Index).Locked Then Exit Sub
    
    
    
    txtAnterior = txtAux(Index).Text
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 16 Then
        'Campo observaciones. NO, repito NO, se selecciona todo
        If txtAux(Index).Text <> "" Then
            txtAux(Index).Text = txtAux(Index).Text & " "
            txtAux(Index).SelStart = Len(txtAux(Index).Text)
        End If
    Else
        ConseguirFocoLin txtAux(Index), cadkey
    End If
    LabelAyudatxtAux Index, lblF
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim HacerPulsadoMas As Boolean
Dim i As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'campo almacen y flecha arriba
    
    If Index < 2 Or Index = 9 Or Index = 12 Or Index = 13 Or Index = 14 Then 'Para los que tienen busqueda
    
    
    
        
            'Insertando linea albaran
            
            If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                
                If Modo = 5 Then
                    If txtAux(Index).Text = "" Then
                        HacerPulsadoMas = False
                        If ModificaLineas = 1 Then
                            HacerPulsadoMas = True
                        Else
                            If Index > 2 Then HacerPulsadoMas = True
                        End If
                        If HacerPulsadoMas Then
                            PulsadoMas2 = True
                            KeyCode = 0
                            PulsarTeclaMas False, Index
                        End If
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
        If Index = 4 Then AbrirConsultaPrecio2 Text1(4).Text, txtAux(1).Text, Text1(1).Text, ""
            
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
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            cantidad = txtAux(9).Text
            
            If Me.DataGrid1.Columns(4).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEan(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , cantidad)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , cantidad)
            End If
            If Not okArticulo Then
                If Me.DataGrid1.Columns(4).Caption = "EAN" Then txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            Else
                'Si ha cambiado el articulo, quito todo menos cantitdad
                If ModificaLineas = 2 Then
                    If txtAux(1).Text <> Data2.Recordset!codArtic Then
                        For NumCajas = 4 To 8
                            txtAux(NumCajas).Text = ""
                        Next
                        NumCajas = 0
                        If InstalacionEsEulerTaxco Then txtAux(15).Text = ""
                    End If
                    
                End If
                B = (Me.ActiveControl.Name = "txtAux")
                If B Then B = (Me.ActiveControl.Index = 0)
                If Not B Then
                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
                
                
                'EULER. Precio uc para los costes
                If InstalacionEsEulerTaxco Then
                    'Si no es de varios
                    If txtAux(2).Locked Then
                        codTarif = DevuelveDesdeBD(conAri, "preciouc", "sartic", "codartic", txtAux(1).Text, "T")
                        If codTarif <> "" Then
                            If CCur(codTarif) = 0 Then
                                codTarif = ""
                            Else
                                codTarif = Format(codTarif, FormatoPrecio)
                            End If
                        End If
                        txtAux(15).Text = codTarif
                        codTarif = ""
                    End If
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
                
                
'                'Si  ha cambiado el articulo, el proveedor
'                If txtAux(9).Text = "" Then
'                    txtAux(9).Text = Cantidad
'                    'Fuerzo el lostfocus para que carge el proveedor
'                    txtAux_LostFocus 9
'                End If
            End If
            
            If txtAux(16).Text = "" Then _
                txtAux(16).Text = DevuelveDesdeBD(conAri, "txtauxdocumento", "sartic", "codartic", txtAux(1).Text, "T")
            
        Case 2 'Nombre Articulo
           If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Si es factura rectifica la cantidad solo puede ser negativa
                If CodTipoMov = "ART" Then
                    If CCur(txtAux(Index)) >= 0 Then
                        MsgBox "En facturas rectificativas la cantidad debe ser negativa.", vbExclamation
                        PonerFoco txtAux(Index)
                        Exit Sub
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
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then
                             B = True
                        Else
                            If CStr(DBLet(Data2.Recordset!origpre, "T")) <> "M" Then B = True
                        End If
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
                            RestoUnid = 0
                        End If
                        
                        CPrecioFact.CodigoClien = Text1(4).Text
                        
                        'Obtenemos la Tarifa del Cliente
                        'AHORA ESTA DENTRO DE LA CLASE
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        'CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.FijarTarifaActividad
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        
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
                        End If
                        
                        'Si tiene dto permitido
                        If Not CPrecioFact.DtoPermitido Then
                            txtAux(6).Text = "0"
                            txtAux(7).Text = "0"
                            txtAux(6).Enabled = False
                            txtAux(7).Enabled = False
                        End If
                        
'                            PonerFoco txtAux(Index + 1)
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
                If vUsu.Nivel2 = 2 Then PonerFocoBtn cmdAceptar
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
                    If Text2(9).Text = "" Then PonerFoco txtAux(9)
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
        
           
    Case 12, 13, 14
        If txtAnterior <> txtAux(Index).Text Then PonerDatosNuevosLineaAlbaran True, Index
        
    Case 15
        PonerFormatoDecimal txtAux(15), 3
    End Select
    
     If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub



Private Function PuedeMeterLineas() As Boolean
    PuedeMeterLineas = True
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
         If EsTipoAlbaranFacturarInmediatoTaxco Then
            If chkFacturar.Value = 1 Then
                MsgBox Me.Caption & "    Situación: cerrado", vbExclamation
                PuedeMeterLineas = False
            End If
        End If
        
        
        
        If hcoCodTipoM = "ALD" Or hcoCodTipoM = "ALB" Then
            MsgBox "Proceso restringido. Avise a Ariadna", vbExclamation
            PuedeMeterLineas = False
        End If
                        
        
    End If


End Function

Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        
    If Not PuedeMeterLineas Then Exit Sub
    
    
    Me.SSTab1.Tab = numTab
    TituloLinea = Cad
    ModificaLineas = 0
    
        If vParamAplic.ArtReciclado <> "" Then
        
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                'Para acelerar , TODOS los clientes llevan tasa reciclado que vendra a ser el SIGAUS
                ClienteConTasaReciclado = True
                
            Else
                ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
            End If
        Else
            ClienteConTasaReciclado = False
        End If
    
        
    If vParamAplic.TipoPortes = 1 Then
        KilosAnteriores = SumaKilosLineas
    End If
    
    PonerModo 5
    PonerBotonCabecera True
    
    AlmacenLineas = -1
    PonerUltAlmacen
    
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
    B = ReestablecerStock
    
    
    If B Then
        'eliminamos de albaranes y pasamos al historico
        If Data2.Recordset.RecordCount > 0 Then
            B = ActualizarElTraspaso(MenError, SQL, CodTipoMov, cadList)
        Else
                SQL = " codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!Numalbar
                conn.Execute "DELETE from slialb_eu where " & SQL
                                
                conn.Execute "DELETE from slialb_eu2 where " & SQL
                
                conn.Execute "DELETE from scaalb_eu where " & SQL
                
                conn.Execute "DELETE from scaalb where " & SQL
                
                
            B = True
        End If
        If B Then
            MenError = "Observaciones linea."
            SQL = "DELETE from slialt where "
            SQL = SQL & " codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!Numalbar
            conn.Execute SQL
            
            

            
            MenError = "Actualizando numeros de serie."
            'Actualizar los posibles num. serie de ese albaran. vaciar los campos
            SQL = "UPDATE  sserie SET codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
            SQL = SQL & " fechavta=" & ValorNulo & ", numalbar=" & ValorNulo & ", numline1=" & ValorNulo
            SQL = SQL & " WHERE codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!Numalbar & " AND fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
            conn.Execute SQL
            
            
            'Devolvemos contador, si no estamos actualizando
            ' Y si no es ALR ya que con los ALR tenemos el problema de EULER
            If InstalacionEsEulerTaxco Then
                If CodTipoMov = "ALR" Then SQL = ""
            End If
            
            If SQL <> "" Then
                Set vTipoMov = New CTiposMov
                B = CBool(vTipoMov.DevolverContador(CodTipoMov, NumAlbElim))
                Set vTipoMov = Nothing
            End If
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
        
        
    End If
    Eliminar = B
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData(DesdeHacerFactura As Boolean)
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
        PonerCadenaBusqueda DesdeHacerFactura
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
    
    
    'Enero 2008. David
    'Para la trazabilidad con repescto al codproveedor en las lineas
    SQL = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,"

    SQL = SQL & "precioar, origpre, dtoline1, dtoline2, importel "
    
    If vEmpresa.TieneAnalitica Then
        SQL = SQL & ",codccost"
    Else
        SQL = SQL & ",codprovex,nomprove"
    End If
    
    
    'SAIL
    SQL = SQL & ",numlote,numbultos"
    SQL = SQL & ",codcapit,codtipor, codtraba,precoste,ampliaci"
    
    SQL = SQL & " FROM " & NomTablaLineas
    'traza
    If vEmpresa.TieneAnalitica = False Then
        SQL = SQL & " LEFT JOIN sprove on codprovex=codprove "
    End If
    
    If enlaza Then
        SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE false "   'numalbar = -1"
    End If
    SQL = SQL & " Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = SQL
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
        
        If Toolbar1.Buttons(14).visible Then Toolbar1.Buttons(14).Enabled = Not EsHistorico
        
        
        'Imprimir
        Toolbar1.Buttons(16).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        'Toolbar1.Buttons(14).Enabled = Toolbar1.Buttons(15).Enabled And vParamAplic.TipoPortes = 1
        
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
End Sub

Private Sub CargarCombos()
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



    CargarCombo_Tabla cboTipoDat, "stipEstadoAlbaran", "tipestado", "desestado", , True, "tipestado"


End Sub






Private Function InsertarAlbaran(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim codtipomAUX As String
Dim ObtenerContador As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    
    
    ObtenerContador = True   'Obtener contador
    codtipomAUX = CodTipoMov
    If InstalacionEsEulerTaxco Then
        If CodTipoMov = "ALR" Then
        
            'Si ha metido a mano el numero de albaran, lo dejo
            If Trim(Text1(0).Text) <> "" Then
                ObtenerContador = False
            Else
                'Si el trabajador es de Valencia sera los ALR, si es de EUSAKADI seran CAR
                devuelve = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", Text1(3).Text)
                If devuelve = "10" Then codtipomAUX = "CAR"
            End If
        End If
    End If
    
    If ObtenerContador Then Text1(0).Text = vTipoMov.ConseguirContador(codtipomAUX)
    
    
    
    cambiaSQL = False
    Do
        'Pero en scaalb, en el caso de los albaranes de reparacion de EULER, siempre graba el ALR
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numalbar", "codtipom", Text1(30).Text, "T", , "numalbar", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            
            If Not ObtenerContador Then Err.Raise 513, , "Entrada duplicada en BD"  'EULER. Pueden poner contador a mano
            
            vTipoMov.IncrementarContador (codtipomAUX)
            Text1(0).Text = vTipoMov.ConseguirContador(codtipomAUX)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    vSQL = CadenaInsertarDesdeForm(Me)
    
    
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
        
        MenError = "Error al actualizar el contador del albaran."

        If ObtenerContador Then vTipoMov.IncrementarContador (codtipomAUX)    'del leedio en la variable
    End If
    
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


Private Sub LimpiarDatosCliente()
Dim i As Byte

    For i = 4 To 17
        Text1(i).Text = ""
    Next i
    Text2(12).Text = ""
    Text2(14).Text = ""
    Text2(17).Text = ""
    Me.cboFacturacion.ListIndex = -1

End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim SQL As String
Dim B As Boolean
Dim ImpReciclado As Single



    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    On Error GoTo EEliminarLinea
    
    conn.BeginTrans
    conn.Execute SQL 'Eliminar linea
    B = vCStock.DevolverStock2
    Set vCStock = Nothing

    ' ---- [15/09/2009] (LAURA)
    'El articulo
'    ElArticulo = Data2.Recordset!codArtic
    ' ----
    
    If B Then
        'Ha borrado la linea y ha devuelvto correctamente el sctock
                   'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    SQL = CStr(Data2.Recordset!codArtic)
                    If vParamAplic.NumeroInstalacion = vbTaxco Then SQL = "NNNN"  'para que nunca entre
                    If ArticuloConTasaReciclado(SQL, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
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
    SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
    SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
    SQL = SQL & " AND numalbar=" & Data1.Recordset!Numalbar & " AND numline1=" & Data2.Recordset!numlinea
    conn.Execute SQL
    

    
     
    SQL = "Albarán: " & Text1(30).Text & "-" & Text1(0).Text & " de " & Text1(1).Text & vbCrLf
    SQL = SQL & "Linea: " & Data2.Recordset!codArtic & " " & DBSet(Data2.Recordset!NomArtic, "T")
    SQL = SQL & "   Uds: " & Data2.Recordset!cantidad & "    Importe:" & DBSet(Data2.Recordset!ImporteL, "T")

    Set LOG = New cLOG
    ' 17 Venta a sabiendas riesgo
    LOG.Insertar 37, vUsu, SQL
    Set LOG = Nothing
    
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        EliminarLinea = True
        
        
        DatosObservaciones SQL, 2, CInt(Data2.Recordset!numlinea)
        
        ' ---- [15/09/2009] (LAURA)
'        DescuentosCantidad ElArticulo
        ' ----
        
        
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


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
    
ERestablecer:
    If Err.Number <> 0 Then B = False
    ReestablecerStock = B
End Function


Private Sub BotonImprimir(OpcionListado As Byte, EsInformePortes As Boolean)
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
    pRptvMultiInforme = 0
    
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
            ElseIf hcoCodTipoM = "ALO" Then
                indRPT = 76
            ElseIf hcoCodTipoM = "ALE" Then
                indRPT = 77
            Else
                If EsHistorico Then
                    indRPT = 11 'Hist. Albaranes clientes
                Else
                    indRPT = 10 'Albaran Clientes
                End If
            End If
        End If
    Else
        
        indRPT = OpcionListado
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
    cadParam = cadParam & "Albarcon=" & devuelve & "|"
    numParam = numParam + 1
    
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    'Nombre fichero .rpt a Imprimir
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
        devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"
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
    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
    If devuelve <> "" Then
        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
        numParam = numParam + 1
    End If

        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        devuelve = NombreTabla & " LEFT JOIN " & NomTablaLineas & " ON "
    Else
        devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    End If
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda.  -----------
        If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
        With frmImprimir
            'Febrero 2010
            If indRPT = 34 Then
                .outTipoDocumento = 0
            Else
                .outTipoDocumento = 4
                .outClaveNombreArchiv = Text1(30).Text & Text1(0).Text
                .outCodigoCliProv = CLng(Text1(4).Text)
            End If
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            If indRPT = 34 Then
                .Titulo = "Portes albaran "
            ElseIf indRPT = 85 Then
                .Titulo = "Costes albaran"
            ElseIf indRPT = 91 Then
                .Titulo = "ORDEN TALLER"
                .NumeroCopias = 2
            Else
                .Titulo = "Albaran de Cliente"
            End If
            .ConSubInforme = True
            .Show vbModal
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
            SQL = SQL & " numalbar=" & Text1(0).Text & " and numline1=" & Data2.Recordset!numlinea & ")"
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


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
Dim SQL As String

    On Error GoTo EPedirNSeries

        SQL = "El artículo tienen control de Nº de Serie." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los Nº De Serie." & vbCrLf
        MsgBox SQL, vbInformation
        PedirNSeriesGnral RS, False
        
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
Dim TbVi As Boolean
    On Error GoTo EInsertarCab
    

    
    
    
    
    Set vTipoMov = New CTiposMov
        
    TbVi = True
        If InsertarAlbaran(SQL, vTipoMov) Then
            Text1(0).Text = Format(Text1(0).Text, "0000000")
            
             'Ficha tecnica
            If SSTab1.TabVisible(2) = True Then ActualizaBDFicha
            If SSTab1.TabVisible(3) = True Then ActualizaBDFicha
        
        
            CrearCarpetaComun2 False, 0
        
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            
            PonerCadenaBusqueda False
            PonerModo 2
            'Ponerse en Modo Insertar Lineas
            
            If vParamAplic.NumeroInstalacion = vbTaxco And CodTipoMov = "ALO" Then
                'NADA
                TbVi = False
            Else
                BotonMtoLineas 0, "Albaranes"
                BotonAnyadirLinea False
            End If
        End If

        
    
    Set vTipoMov = Nothing
    If TbVi Then Me.SSTab1.Tab = 0
    
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
                Dif = CSng(txtAux(3).Text) - CSng(Data2.Recordset!cantidad)
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


Private Sub PedirNSeriesT(ByRef RS As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim SQL As String
Dim linea As Integer

    On Error GoTo EPedirNSeries


        PedirNSeriesGnral RS, False
        RS.MoveFirst
        While Not RS.EOF
            linea = 0
            'Cargar los Nº de serie asignados
            SQL = "SELECT numserie, codartic,nummante FROM sserie "
            SQL = SQL & " WHERE codtipom='" & Text1(30).Text & "' and "
            SQL = SQL & "numalbar=" & Text1(0).Text
            SQL = SQL & " and numline1=" & RS!numlinea
            SQL = SQL & " ORDER BY codartic "
            Set RSseries = New ADODB.Recordset
            RSseries.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                SQL = SQL & ", nummante = " & DBSet(RSseries!nummante, "T")
                SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
                SQL = SQL & " and numlinealb=" & RS!numlinea
                SQL = SQL & " and numlinea=" & linea
                conn.Execute SQL
                RSseries.MoveNext
            Wend
            RS.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
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
     SQL = SQL & ", TieneMan=0 , NumMante= " & ValorNulo & ",coddirec=" & ValorNulo
    
    SQL = SQL & " WHERE codtipom=" & DBSet(Text1(30).Text, "T") & " and numalbar=" & Text1(0).Text & " AND year(fechavta)=" & Year(Text1(1).Text)
    conn.Execute SQL
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
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
            B = False
            SoloEnEfectivoAlbaranes = False
            If vCliente.ClienteBloqueado(0, SoloEnEfectivoAlbaranes) Then
                B = True
                
                ' Abril 2020  Hemos añadido SoloEnEfectivoAlbaranes
                'If vParamAplic.NumeroInstalacion = vbTaxco Then b = False
                
                If B Then
                    LimpiarDatosCliente
                    PonerFoco Text1(4)
                    Set vCliente = Nothing
                    Exit Sub
                End If
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
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    Text1(14).Text = 1  'SIEMPRE EFECTIVO
                Else
                    Text1(14).Text = vCliente.ForPago
                End If
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Text1(34).Text = IIf(vParamAplic.NumeroInstalacion = vbTaxco, 0, vCliente.Kilometros)
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
                Text1(29).Text = vCliente.FEnvio
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
            End If
                
'                SituacionCliente = RS.Fields!codsitua

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                           
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(1).Text
            
            
            
            If vParamAplic.NumeroInstalacion = vbTaxco And Modo = 3 And CodTipoMov = "ALO" Then
                
                Me.SSTab1.Tab = 2
                PonerFoco txtTaxco(0)
            
            End If
            
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
    If B Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
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
    If hcoCodTipoM = "ALZ" Then vFactu.codtipom = "ALZ"
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



 Private Function InsertarLineasFactu(cadWhere, DesdeCostes As String) As Boolean
'cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
'cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
 Dim RS As ADODB.Recordset
 Dim SQL As String
 Dim i As Integer
 Dim cadI As String
 Dim numlin As String
 Dim CCos As String   'por si acaso lleva analitica y la linea NO lo llevaba
 
 
    On Error GoTo EInsFactu
    Screen.MousePointer = vbHourglass
    InsertarLineasFactu = False
    
    If cadWhere <> "" Then
        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialb", "numlinea", SQL)
'        i = Int(SQL)
            
        
        cadI = ""
        If DesdeCostes <> "" Then
            'Mayo 2021
            'Estamos insertando desde costes
            SQL = DesdeCostes
        Else
            'desde factura. Lo que habia
            SQL = "SELECT * FROM slifac WHERE " & cadWhere
        End If
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            txtAux(0).Text = RS!codAlmac
            txtAux(1).Text = RS!codArtic
            txtAux(2).Text = RS!NomArtic
            txtAux(16).Text = DBLet(RS!Ampliaci, "T")
'            Text2(9).Text = DBLet(RS!nomprove, "T")
            txtAux(3).Text = CStr(RS!cantidad * -1)
            txtAux(4).Text = RS!precioar
            txtAux(5).Text = DBLet(RS!origpre, "T")
            txtAux(6).Text = RS!dtoline1
            txtAux(7).Text = RS!dtoline2
            txtAux(8).Text = CStr(RS!ImporteL * -1)
            
            ' ---- [21/10/2009] [LAURA] : se añade el centro de coste
            If Not vEmpresa.TieneAnalitica Then
                txtAux(9).Text = DBLet(RS!codProvex, "N")
                Text2(9).Text = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(RS!codProvex), "N")
            Else
                CCos = DBLet(RS!CodCCost)
                If CCos = "" Then
                    'MAL. DEBERIA tener Analitica.
                    If vParamAplic.ModoAnalitica = 1 Then CCos = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND codartic", CStr(RS!codArtic), "T")
                    If CCos = "" Then CCos = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Text1(3).Text)
                End If
                txtAux(9).Text = CCos
                Me.Text2(9).Text = PonerNombreCCoste(txtAux(9))
            End If
            
            'para no tener que traer ahora el proveedor pongo en txt(10) un texto
'            txtAux(10).Text = "*"
'            Text2(9).Text = "*"
            
            'numbultos
            txtAux(10).Text = CStr(RS!NumBultos * -1)
            'numlote
            txtAux(11).Text = DBLet(RS!numLote, "T")
            
            
            If DesdeCostes <> "" Then
                For i = 12 To 15
                    txtAux(i).Text = ""
                Next
            End If
            
            
            If InsertarLinea(numlin, True) Then
            
            End If
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        CalcularDatosFactura
        
        InsertarLineasFactu = True
    End If
    Screen.MousePointer = vbDefault
    
EInsFactu:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Lineas Factura", Err.Description
    End If
End Function




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
Dim RS As ADODB.Recordset

    On Error GoTo EDevNumSerie
        
    cadWhere = ObtenerWhereCP(True)
    SQL = "select slialb.codartic,abs(cantidad) as cantidad,numlinea"
    SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
    SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu and sserie.codclien=scaalb.codclien "
    '-- LAURA: 02/07/2007
'    SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'    SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
    SQL = SQL & cadWhere & " and scaalb.numfactu="
    If IsNull(Data1.Recordset!Numfactu) Then
         SQL = SQL & " -1"
    Else
     SQL = SQL & CStr(Me.Data1.Recordset!Numfactu)
    End If
'    If Me.Data1.Recordset!codtipmf = "FAV" Then SQL = SQL & " AND codtipom='ALV'"
    '--

    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Hay articulos con nº de serie en las lineas del albaran rectificativo
    'que hay que quitar de los nº de serie que tenia asignados
    'estamos devolviendo nº serie y pedimos los que vamos a devolver y a estos
    'le limpiamos los campos de venta de la tabla sserie
    If Not RS.EOF Then
        SQL = "select sserie.numserie, sserie.codartic, sartic.nomartic"
        SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
        '-- LAURA: 02/07/2007
'        SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'        SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
        SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu  and sserie.codclien=scaalb.codclien "
        '--
        SQL = SQL & " inner join sartic on sserie.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!Numfactu)
    
        MostrarNSeries RS, , SQL
    End If
    RS.Close
    Set RS = Nothing
    
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


Private Sub BotonRecuperarFactura()
'Genera una factura a partir del Albaran de Mostrador
'pero sin coger contador de factura lo pide en un form

    'Comprobar que esta marcada para facturar
    If Me.chkFacturar.Value = 1 Then
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Facturacion de Albaran de Mostrador
        frmListadoPed.codClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
        frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
        AbrirListadoPed (222)
        
        PosicionarDataTrasEliminar
    Else
        MsgBox "El Albaran no esta marcado para facturar", vbInformation
    End If
End Sub


Private Sub MarcarAlbaranes()

        'Lanzara un desde hasta y los marcara
        frmListado.NumCod = hcoCodTipoM
        CadenaDesdeOtroForm = ""
        AbrirListado 82
        If CadenaDesdeOtroForm = "OK" Then
            'OK. Cambiadas las marcas. Refrescamos y situamos
            Screen.MousePointer = vbHourglass
            DoEvents
            PonerCadenaBusqueda False
            PosicionarData False
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
                                           
                                                                       Else 'masl 090709
                        C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(DtoPorte * (-1)))
                        C = C & ",0" & ",0,"
                                                                       
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
        
        
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        Cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from slialb " & Cad
        If Articulo <> "" Then Cad = Cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        'Y origen PRECIO no es precio especial
        Cad = Cad & " AND origpre <> 'E'"
        Cad = Cad & " group by 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                Cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                Cad = "select * from sdesca where desdecan <=" & Cad & " and " & Cad & " <= hastacan and envagran = '"
                Cad = Cad & miRsAux!raiz & "'"
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
                    '                                 a partir de la 3era posicion
                    Cad = Cad & " AND codartic like '__" & miRsAux!raiz & "%'"
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
    
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub





Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    imgBuscar2(9).Tag = "9"
    
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
    imgBuscar2(9).Tag = "-1"
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
        EsCabecera = True
        imgBuscar_Click Index
        
    Else
        'Lineas
        EsCabecera = False
        If Index >= 9 Then
            If Index = 14 Then
                imgBuscar2_Click 0
            Else
                imgBuscar2_Click Index
            End If
        Else
            cmdAux_Click Index
        End If
        
        
    End If
        
End Sub



Private Sub PonerDatosNuevosLineaAlbaran(Edicion As Boolean, Index As Integer)
Dim devuelve As String
Dim J As Integer
       devuelve = ""
            
                'Si es numerico
                'ORDEN TRABAJO=13
                
                If Index <> 13 Then
                    J = Index - 12  'Sera el index del text2
                    If txtAux(Index).Text <> "" Then
                        If Not EsNumerico(txtAux(Index).Text) Then
                            txtAux(Index).Text = ""
                            If Edicion Then PonerFoco txtAux(Index)
                        End If
                    End If
                Else
                    J = Index
                End If
                
                If txtAux(Index).Text <> "" Then
                    If Index = 12 Then
                        'codcapit nomcapit scapitulos
                        devuelve = DevuelveDesdeBD(conAri, "nomcapit", "scapitulos", "codcapit", txtAux(Index).Text, "N")
                    ElseIf Index = 13 Then
                        'stipor codtipor nomtipor
                        devuelve = DevuelveDesdeBD(conAri, "nomtipor", "stipor", "codtipor", txtAux(Index).Text, "T")
                    Else
                        devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtAux(Index).Text, "N")
                    End If
                    If devuelve = "" Then
                        MsgBox "No existe el registro para el campo: " & txtAux(Index).Text & " en la tabla de " & txtAux(Index).Tag, vbExclamation
                        txtAux(Index).Text = ""
                        If Edicion Then PonerFoco txtAux(Index)
                    End If
                End If
                
                Text2(J).Text = devuelve
                


End Sub

'Hacer:  0. inser   1. modfi   2. borrar
Private Sub DatosObservaciones(ByRef SQL As String, Hacer As Byte, linea As Integer)


    If Hacer > 0 Then
        'Borrar y update
        If Hacer = 1 Then
            'update
            If txtAux(17).Text = "" Then
                'La borro
                Hacer = 2
            Else
                SQL = DevuelveDesdeBD(conAri, "numlinea", "slialt", "codtipom= '" & CodTipoMov & "' AND numalbar = " & Text1(0).Text & " AND numlinea", Data2.Recordset!numlinea, "N")
            
                'UPDATE
                If SQL = "" Then
                    'Insertamos NUEVO
                    Hacer = 0
                Else
                    SQL = "UPDATE slialt set observa=" & DBSet(txtAux(17).Text, "T")
                End If
            End If
        End If
        
        'dele
        If Hacer = 2 Then SQL = "DELETE FROM slialt "
        
        SQL = SQL & " WHERE codtipom='" & CodTipoMov & "' AND numalbar = " & Text1(0).Text & " AND numlinea = " & Data2.Recordset!numlinea
    
    End If
    
    If Hacer = 0 Then
        If txtAux(17).Text <> "" Then
            SQL = "INSERT INTO slialt (codtipom, numalbar,numlinea,observa) VALUES ('"
            SQL = SQL & CodTipoMov & "'," & Text1(0).Text & "," & linea & "," & DBSet(txtAux(17).Text, "T") & ")"
        Else
            Exit Sub
        End If
    End If
    
    ejecutar SQL, False
End Sub




Private Sub PonerCampoActuacion()
Dim CADENA As String
            If Modo = 1 Then Exit Sub
            CADENA = ""
            
            If vParamAplic.NumeroInstalacion = vbEuler Then
                If Text1(42).Text <> "" Then
                    CADENA = "codclien =" & Text1(4).Text & " AND coddirec is null  " & " AND actuacion "
                    CADENA = DevuelveDesdeBDNew(conAri, "sactuaobra", "coalesce(observa,fechaini,actuacion)", CADENA, Text1(42).Text, "T")
                     If CADENA = "" Then
                         MsgBox "Ningun proyecto con ese valor:" & Text1(42).Text, vbInformation
                         Text1(42).Text = ""
                     End If
                End If
            
            Else
                If Text1(42).Text <> "" Then
                    Text1(42).Text = UCase(Text1(42).Text)
                    If Text1(4).Text = "" Or Text1(12).Text = "" Then
                        MsgBox "Falta cliente/obra", vbExclamation
                        Text1(42).Text = ""
                    Else
                        CADENA = "codclien =" & Text1(4).Text & " AND coddirec= " & Text1(12).Text & " AND actuacion "
                    
                        CADENA = DevuelveDesdeBDNew(conAri, "sactuaobra", "concat(fechaini,' ',if(observa is null,'',observa))", CADENA, Text1(42).Text, "T")
                        If CADENA = "" Then
                            MsgBox "Ninguna actuacion con ese valor:" & Text1(42).Text, vbInformation
                            Text1(42).Text = ""
                        End If
                    End If
                    
                End If
            
            End If
            
            Text2(1).Text = CADENA

End Sub



Private Sub txtEule_R_GotFocus(Index As Integer)
    ConseguirFoco txtEule_R(Index), Modo
End Sub

Private Sub txtEule_R_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtEule_R_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtEule_R(Index), Modo) Then Exit Sub
    
    If Index = 20 Or Index = 21 Or Index = 22 Then
        txtEule_R(Index).Text = Trim(txtEule_R(Index).Text)
        If txtEule_R(Index).Text <> "" Then
            If Not PonerFormatoEntero(txtEule_R(Index)) Then
                txtEule_R(Index).Text = ""
            Else
                txtAnterior = IIf(Index = 20, "ALE", IIf(Index = 21, "ALO", "ALV"))
                txtAnterior = "codtipom = '" & txtAnterior & "' AND numalbar"
                If DevuelveDesdeBD(conAri, "numalbar", "scaalb", txtAnterior, txtEule_R(Index).Text) = "" Then
                    'Label3(36 o 37
                    MsgBox "El albaran de " & Label3E(Index + 16).Caption & " NO existe", vbExclamation
                End If
                txtAnterior = ""
            End If
        End If
    Else
        If Index = 2 Then
            txtEule_R(Index).Text = Trim(txtEule_R(Index).Text)
            If txtEule_R(Index).Text <> "" Then
            
                txtAnterior = txtEule_R(Index).Text
                If Not EsFechaOK(txtAnterior) Then
                    MsgBox "Fecha incorrecta", vbExclamation
                    txtEule_R(Index).Text = ""
                    PonerFoco txtEule_R(Index)
                Else
                    txtEule_R(Index).Text = txtAnterior
                End If
                txtAnterior = ""
            End If
        End If
    End If
        
End Sub


Private Sub txtEuler_GotFocus(Index As Integer)
    If Index <> 7 Then ConseguirFoco txtEuler(Index), Modo
End Sub

Private Sub txtEuler_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 7 And Index <> 6 Then KEYpress KeyAscii
End Sub

Private Sub txtEuler_LostFocus(Index As Integer)
        If Not PerderFocoGnral(txtEuler(Index), Modo) Then Exit Sub
        
        
        If Index = 8 Or Index = 9 Or Index = 10 Then
            If Not PonerFormatoEntero(txtEuler(Index)) Then txtEuler(Index).Text = ""
            
            txtEule_R(Index).Text = Trim(txtEuler(Index).Text)
            If txtEuler(Index).Text <> "" Then
                If Not PonerFormatoEntero(txtEuler(Index)) Then
                    txtEuler(Index).Text = ""
                Else
                
                    'El campo es ALO si es un trabajao exterior
                    ' o  ALE si es una orden de trabajo
                    txtAnterior = IIf(hcoCodTipoM = "ALO", "ALE", "ALO")
                
                    txtAnterior = IIf(Index = 8, "ALR", IIf(Index = 9, txtAnterior, "ALV"))
                    txtAnterior = "codtipom = '" & txtAnterior & "' AND numalbar"
                    If DevuelveDesdeBD(conAri, "numalbar", "scaalb", txtAnterior, txtEuler(Index).Text) = "" Then
                        MsgBox "El albaran de " & Label3(Index - 2).Caption & " NO existe", vbExclamation
                    End If
                    txtAnterior = ""
                End If
            End If
            
            
        Else
            If Index = 7 Then PonerFocoBtn cmdAceptar
        End If
        
        
End Sub


Private Sub LimpiarFichaTecnica(SinTxts As Boolean)
    If vParamAplic.NumeroInstalacion = vbEuler Then
        LimpiarFichaTecnicaEuler SinTxts
    Else
        LimpiarFichaTecnicaTaxco SinTxts
    End If
End Sub

Private Sub LimpiarFichaTecnicaEuler(SinTxts As Boolean)
Dim N As Byte
Dim AuxTipom As String
    
    AuxTipom = CodTipoMov
    If EsHistorico Then
        If Modo > 0 Then
            If Data1.Recordset.EOF Then AuxTipom = Data1.Recordset!codtipom
        End If
    End If
    If AuxTipom = "ALR" Then
         If Not SinTxts Then
            For N = 0 To Me.txtEule_R.Count - 1
                txtEule_R(N).Text = ""
            Next
         End If
         
         For N = 0 To chkEuler.Count - 1
             chkEuler(N).Value = 0
         Next
         
         For N = 0 To Me.optEule_R.Count - 1
            Me.optEule_R(N).Value = False
         Next N
         
         cboEulerUdR.ListIndex = -1
        
         cboRecepAgenClien.ListIndex = -1
    Else
        If Not SinTxts Then
            For N = 0 To Me.txtEuler.Count - 1
                txtEuler(N).Text = ""
            Next
        End If
        
        Me.optEuler(0).Value = True
        Me.optEuler(0).Value = False  'Ninguno seleccionado
    
    End If
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        Text5(0).Text = ""
        Text5(1).Text = ""
    End If
End Sub
Private Sub LimpiarFichaTecnicaTaxco(SinTxts As Boolean)
Dim N As Byte
Dim AuxTipom As String
    
    AuxTipom = CodTipoMov
    If EsHistorico Then
        If Modo > 0 Then
            If Data1.Recordset.EOF Then AuxTipom = Data1.Recordset!codtipom
        End If
    End If
    If Not SinTxts Then
        For N = 0 To Me.txtTaxco.Count - 1
            txtTaxco(N).Text = ""
            txtTaxco(N).BackColor = vbWhite
        Next
        Text5(0).Text = ""
        Text5(1).Text = ""
    
        For N = 7 To 10
            txtEuler(N).Text = ""
            txtEuler(N).BackColor = vbWhite
        Next
    End If
        

    
End Sub





Private Sub BloquearFicha(Bloquea As Boolean)
Dim N As Byte
Dim AuxTipom As String
        
        AuxTipom = CodTipoMov
        If EsHistorico Then
            If Not Data1.Recordset.EOF Then AuxTipom = Data1.Recordset!codtipom
        End If
        
        
        
        If AuxTipom = "ALR" Then
            
                cboEulerUdR.Enabled = Not Bloquea
                 
                For N = 0 To Me.txtEule_R.Count - 1
                    BloquearTxt txtEule_R(N), Bloquea
                Next
                
                For N = 0 To Me.optEule_R.Count - 1
                    Me.optEule_R(N).Enabled = Not Bloquea
                Next N
                
                For N = 0 To chkEuler.Count - 1
                    chkEuler(N).Enabled = Not Bloquea
                Next
        
                
                cboRecepAgenClien.Locked = Bloquea
        Else
            
            If vParamAplic.NumeroInstalacion = vbEuler Then
        
                For N = 0 To Me.txtEuler.Count - 1
                    BloquearTxt txtEuler(N), Bloquea
                Next
            
                For N = 0 To Me.optEuler.Count - 1
                    Me.optEuler(N).Enabled = Not Bloquea
                Next N
            
            
            Else
                
                For N = 0 To Me.txtTaxco.Count - 1
                    BloquearTxt txtTaxco(N), Bloquea
                Next
                
                For N = 7 To 10
                    BloquearTxt txtEuler(N), Bloquea
                Next
                
                
                If Not Bloquea Then txtTaxco(7).BackColor = &HC0E0FF
                
            End If
            
        End If
End Sub



Private Function CamposSQlFicha() As String

    If vParamAplic.NumeroInstalacion = vbEuler Then

            'Primero iran todos los txts juntos y por orden de index
            CamposSQlFicha = "ReferPedido,FechaPed,bombamarca,bombaModelo,motormarca,motorModelo"
            CamposSQlFicha = CamposSQlFicha & ",TrabajoExterior,observaciones,"
            
            'Resto
            CamposSQlFicha = CamposSQlFicha & "NumReparacion, "
            
            'NumReparacion NumParteTrabajo NumTrabajExterno  numero de albara de para vincular OT a TE REp, TE a OT y REp ...
            CamposSQlFicha = CamposSQlFicha & IIf(hcoCodTipoM = "ALE", "NumParteTrabajo", "NumTrabajExterno")
            CamposSQlFicha = CamposSQlFicha & ", NumAlbaranVenta "
            
            CamposSQlFicha = CamposSQlFicha & ",TipoPortes,codtipom,numalbar"
    
    Else
        
'''        bombamarca -Matricula
'''        bombaModelo -bastidor
'''        motormarca -motor
'''        motorModelo -marca modelo
'''        ReferPedido -neumaticos
'''        RecepAgenCliMat -licencia
'''        RecpNumExp -taximetro
'''        numrepar -kms
'''        Observaciones
        CamposSQlFicha = "bombamarca,bombaModelo,motorModelo,motormarca,ReferPedido,RecepAgenCliMat,RecpNumExp,numrepar"
        CamposSQlFicha = CamposSQlFicha & ", observaciones, NumReparacion ,NumTrabajExterno"
        CamposSQlFicha = CamposSQlFicha & ", NumAlbaranVenta "
        CamposSQlFicha = CamposSQlFicha & ", TipoPortes, codtipom, numalbar"
        
    End If
    
    
End Function

Private Sub PonerCamposFicha()
Dim N As Byte
Dim SQL As String


    SQL = CamposSQlFicha()
    'schalb_eu
    SQL = "Select " & SQL & " FROM " & IIf(EsHistorico, "schalb_eu", "scaalb_eu") & " WHERE numalbar = " & Text1(0).Text & " AND codtipom = " & DBSet(Text1(30).Text, "T")
    
        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        LimpiarFichaTecnica False
        
    Else
        
        
        If vParamAplic.NumeroInstalacion = vbEuler Then
            'EL SQL estara montaddo para que coincida el orden del columna con el index
            For N = 0 To txtEuler.Count - 1
                txtEuler(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            Next
            
            
            'Agencia cliente
            N = 1
            If DBLet(miRsAux!TipoPortes, "N") = 0 Then N = 0
            optEuler(N).Value = True
        
        
        Else
'            bombamarca -Matricula
'            bombaModelo -bastidor
'            motormarca -marca
'            motorModelo -modelo
'            ReferPedido -neumaticos
'            RecepAgenCliMat -licencia
'            RecpNumExp -taximetro
'            numrepar -kms
'            Observaciones
 
            For N = 0 To txtTaxco.Count - 1
                txtTaxco(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            Next
           
        
            For N = 7 To 10
                txtEuler(N).Text = DBLet(miRsAux.Fields(CInt(N + 1)), "T")
            Next
        
            If Modo = 2 And hcoCodTipoM = "ALO" Then
                Text5(0).Text = txtTaxco(0).Text
                Text5(1).Text = txtTaxco(7).Text
            End If
       End If
        
        
        
    End If
    miRsAux.Close
    
    
    
   
    
    
    
    
    Set miRsAux = Nothing
End Sub

Private Sub ActualizaBDFicha()
    If CodTipoMov = "ALR" Then
        ActualizaBDFichaALR
    Else
        ActualizaBDFichaNor
    End If
End Sub

Private Sub ActualizaBDFichaALR()
Dim s As String
Dim N As Byte

    s = CamposSQlFichaReparacion()
    s = "REPLACE INTO scaalb_eu(" & s & ") VALUES ("
    For N = 0 To txtEule_R.Count - 1
        s = s & DBSet(txtEule_R(N).Text, "T", "S") & ","
    Next

    For N = 0 To Me.chkEuler.Count - 1
        s = s & DBSet(chkEuler(N), "T", "S") & ","
    Next
    
    
    'numrepar ,  RecepAgenClien,RecepPortes, DatosBomUdCaudal,DatosBomTipoRodete"
    If cboRecepAgenClien.ListIndex < 0 Then
        s = s & "NULL"
    Else
        s = s & cboRecepAgenClien.ListIndex
    End If
    's = s & Abs(Me.optEule_R(1).Value) & "," & Abs(Me.optEule_R(2).Value) & ","
    s = s & "," & Abs(Me.optEule_R(2).Value) & ","
    
    s = s & IIf(Me.cboEulerUdR.ListIndex < 1, "null", cboEulerUdR.ListIndex) & ","
    'Rodete
    kCampo = 6
    For N = 4 To 7
        If Me.optEule_R(N).Value Then kCampo = N
    Next
    s = s & kCampo & "," & DBSet(Text1(30).Text, "T") & "," & Text1(0).Text & ")"
    
    
   
   conn.Execute s
    
End Sub



Private Sub ActualizaBDFichaNor()
Dim s As String
Dim N As Byte

    s = CamposSQlFicha()
    
    If vParamAplic.NumeroInstalacion = vbEuler Then
    
        s = "REPLACE INTO scaalb_eu(" & s & ") VALUES ("
        For N = 0 To txtEuler.Count - 1
            s = s & DBSet(txtEuler(N).Text, "T", "S") & ","
        Next
    
        'For N = 1 To Me.chkEuler.Count
        '    s = s & DBSet(chkEuler(N - 1), "T", "S") & ","
        'Next
        
        
        'numlbar
        s = s & Abs(Me.optEuler(1).Value) & ",'" & Text1(30).Text & "'," & Text1(0).Text & ")"
    
   
    Else
        'TAXCO
         
        s = "REPLACE INTO scaalb_eu(" & s & ") VALUES ("
        'bombamarca,bombaModelo,motormarca,motorModelo,ReferPedido,RecepAgenCliMat,RecpNumExp,numrepar
        For N = 0 To txtTaxco.Count - 1
            s = s & DBSet(txtTaxco(N).Text, "T", "S") & ","
        Next
    
        'observaciones,NumReparacion ,NumTrabajExterno, NumAlbaranVenta
        For N = 7 To 10
            s = s & DBSet(txtEuler(N).Text, "T", "S") & ","
        Next
    
        ',TipoPortes,codtipom,numalbar
        s = s & "0,'" & Text1(30).Text & "'," & Text1(0).Text & ")"
   
           
    End If
    ejecutar s, False
    
End Sub

Private Function BuscaEnBDFicha() As String
    If vParamAplic.NumeroInstalacion = vbEuler Then
        BuscaEnBDFicha = BuscaEnBDFichaEuler
    Else
        BuscaEnBDFicha = BuscaEnBDFichaTaxco
    End If
    
End Function


Private Function BuscaEnBDFichaEuler() As String
Dim Columnas As String
Dim SQ As String
Dim N As Byte

    
    BuscaEnBDFichaEuler = ""
    
    
    If hcoCodTipoM = "ALO" Or hcoCodTipoM = "ALE" Then
    
        Columnas = CamposSQlFicha()
        Columnas = Replace(Columnas, ",", "|")
    
    
        For N = 0 To txtEuler.Count - 1
            If Trim(txtEuler(N).Text) <> "" Then
                
                'Numericos. "", "")
                SQ = ""
                If N = 8 Or N = 9 Then
                    
                    If SeparaCampoBusqueda("N", RecuperaValor("NumReparacion|" & IIf(hcoCodTipoM = "ALE", "NumParteTrabajo", "NumTrabajExterno") & "|", N - 7), txtEuler(N).Text, SQ) > 0 Then
                        
                    End If
                Else
                
                    SQ = RecuperaValor(Columnas, CInt(N + 1))
                    If InStr(1, txtEuler(N).Text, "*") > 0 Then
                        SQ = SQ & " like " & DBSet(Replace(Me.txtEuler(N).Text, "*", "%"), "T")
                    Else
                        SQ = SQ & " = " & DBSet(txtEuler(N), "T", "S")
                    End If
                
                End If
                If SQ <> "" Then BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND " & SQ
            End If
        Next
        
        'Portes debidos, pagados
        kCampo = -1
        For N = 0 To 1
            If Me.optEuler(N).Value Then kCampo = N
        Next
        If kCampo >= 0 Then BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND TipoPortes = " & kCampo
            
        If BuscaEnBDFichaEuler <> "" Then BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND codtipom = '" & hcoCodTipoM & "'"
            
   Else
   
        'Albaranes de reparacion
        If hcoCodTipoM = "ALR" Then
            Columnas = CamposSQlFichaReparacion()
            Columnas = Replace(Columnas, ",", "|")
        
            For N = 0 To txtEule_R.Count - 1
                If Trim(txtEule_R(N).Text) <> "" Then
                    SQ = RecuperaValor(Columnas, CInt(N + 1))
                    If InStr(1, txtEule_R(N).Text, "*") > 0 Then
                        SQ = SQ & " like " & DBSet(Replace(Me.txtEule_R(N).Text, "*", "%"), "T")
                    Else
                        SQ = SQ & " = " & DBSet(txtEule_R(N), "T", "S")
                    End If
                    BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND " & SQ
                End If
            Next
        
            For N = 0 To Me.chkEuler.Count - 1
                If chkEuler(N).Value = 1 Then
                    'El primero es Bombas horizontal superficie
                    SQ = RecuperaValor(Columnas, CInt(N + 23)) & " = 1"
                    BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND " & SQ
                End If
            Next
            
            'DatosBomTipoRodete ,DatosBomCaudal DatosBomUdCaudal
            SQ = ""
            For N = 4 To 7
                If Me.optEule_R(N).Value Then SQ = "DatosBomTipoRodete =" & N
            Next
            If SQ <> "" Then BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND " & SQ
            
            'Bomba caudal
            N = 11
            If Trim(txtEule_R(N).Text) <> "" Then
                    If InStr(1, txtEule_R(N).Text, "*") > 0 Then
                        SQ = SQ & " like " & DBSet(Replace(Me.txtEule_R(N).Text, "*", "%"), "T")
                    Else
                        SQ = SQ & " = " & DBSet(txtEule_R(N), "T", "S")
                    End If
                    BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND " & SQ
            End If
                    
            If Me.cboEulerUdR.ListIndex >= 0 Then BuscaEnBDFichaEuler = BuscaEnBDFichaEuler & " AND  codtipom='ALR' AND DatosBomUdCaudal = " & cboEulerUdR.ListIndex
            
        End If 'ALR
        
        
   
   End If
            
    
    If BuscaEnBDFichaEuler <> "" Then
        BuscaEnBDFichaEuler = Mid(BuscaEnBDFichaEuler, 5)
        BuscaEnBDFichaEuler = "Select numalbar from scaalb_eu WHERE " & BuscaEnBDFichaEuler
        BuscaEnBDFichaEuler = " numalbar IN (" & BuscaEnBDFichaEuler & ")"
    End If
      
End Function
Private Function BuscaEnBDFichaTaxco() As String
Dim Columnas As String
Dim SQ As String
Dim N As Byte

    
    BuscaEnBDFichaTaxco = ""
    
    If hcoCodTipoM <> "ALO" Then Exit Function
    
    
    
    
        Columnas = CamposSQlFicha()
        Columnas = Replace(Columnas, ",", "|")
    
    
        For N = 0 To txtTaxco.Count - 1
            If Trim(txtTaxco(N).Text) <> "" Then
                
                'Numericos. "", "")
                SQ = ""
                If N = 7 Then
                    
                    If SeparaCampoBusqueda("N", RecuperaValor(Columnas, N + 1), Me.txtTaxco(N).Text, SQ) > 0 Then
                        
                    End If
                Else
                
                    SQ = RecuperaValor(Columnas, CInt(N + 1))
                    If InStr(1, txtTaxco(N).Text, "*") > 0 Then
                        SQ = SQ & " like " & DBSet(Replace(Me.txtTaxco(N).Text, "*", "%"), "T")
                    Else
                        SQ = SQ & " = " & DBSet(txtTaxco(N), "T", "S")
                    End If
                
                End If
                If SQ <> "" Then BuscaEnBDFichaTaxco = BuscaEnBDFichaTaxco & " AND " & SQ
            End If
            
            
        Next
       
        
        For N = 7 To 10
        
            If Me.txtEuler(N).Text <> "" Then
                
                SQ = RecuperaValor(Columnas, CInt(N + 9))
                If InStr(1, txtEuler(N).Text, "*") > 0 Then
                    SQ = SQ & " like " & DBSet(Replace(Me.txtEuler(N).Text, "*", "%"), "T")
                Else
                    SQ = SQ & " = " & DBSet(txtEuler(N), "T", "S")
                End If
                If SQ <> "" Then BuscaEnBDFichaTaxco = BuscaEnBDFichaTaxco & " AND " & SQ
            End If
        Next
        
        
        If BuscaEnBDFichaTaxco <> "" Then BuscaEnBDFichaTaxco = BuscaEnBDFichaTaxco & " AND codtipom = '" & hcoCodTipoM & "'"
            
        If BuscaEnBDFichaTaxco <> "" Then
            BuscaEnBDFichaTaxco = Mid(BuscaEnBDFichaTaxco, 5)
            BuscaEnBDFichaTaxco = "Select numalbar from scaalb_eu WHERE " & BuscaEnBDFichaTaxco
            BuscaEnBDFichaTaxco = " numalbar IN (" & BuscaEnBDFichaTaxco & ")"
        End If


End Function



'*************************************************************************************
' Ficha reparacion

Private Function CamposSQlFichaReparacion() As String
    'Primero iran todos los txts juntos y por orden de index
    CamposSQlFichaReparacion = "RecepAgenCliMat,RecpNumExp,FechaAlb,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,DatosBommarca"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosBomNumCurva,DatosBomModelo,DatosBomNumSerie,DatosBomAno,DatosBomH,DatosBomCaudal"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosMotorMarca , DatosMotorModelo, DatosMotorNumSerie, DatosMotorV, DatosMotorI"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosMotorCV, DatosMotorKw, DatosMotorrpm,NumTrabajExterno,NumParteTrabajo,NumAlbaranVenta"

    'Tipo bomba recepcionada
    'Son los check. Tambien vmos con el ordern
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", TipoBombResSuperHor,TipoBombResSuperVer,TipoBombResSumPoz, TipoBombResSumVer, TipoBomAgitadorRes"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombLimSumPoz, TipoBombLimSumVer, TipoBomAgitadorLim "
    

    'Luego resto campos
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", RecepAgenClien,RecepPortes, DatosBomUdCaudal,DatosBomTipoRodete"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",codtipom,numalbar "
    
End Function





Private Sub PonerCamposFichaReparacion()
Dim N As Byte
Dim SQL As String
    
    SQL = CamposSQlFichaReparacion()
'    If EsHistorico Then
'       SQL = "Select " & SQL & " FROM scaalb_eu WHERE numrepar = " & Text1(2).Text & " AND fecrepar =" & DBSet(Text1(4).Text, "F")
'    Else
     SQL = "Select " & SQL & " FROM scaalb_eu WHERE numalbar = " & Text1(0).Text & " AND codtipom = " & DBSet(Text1(30).Text, "T")
'    End If
        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        LimpiarFichaTecnica False
        
    Else
        
        
        
        'EL SQL estara montaddo para que coincida el orden del columna con el index
        For N = 0 To txtEule_R.Count - 1
            txtEule_R(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            'If N = 20 Or N = 21 Then
            If N >= 20 And N <= 22 Then
                'NUmerico
                If txtEule_R(N).Text <> "" Then txtEule_R(N).Text = Format(txtEule_R(N).Text, "000000")
            End If
        Next
    
        'Agencia cliente
        'N = 1
        'If DBLet(miRsAux!RecepAgenClien, "N") = 0 Then N = 0
        'optEule_R(N).Value = True
        
        If IsNull(miRsAux!RecepAgenClien) Then
            Me.cboRecepAgenClien.ListIndex = -1
        Else
            Me.cboRecepAgenClien.ListIndex = miRsAux!RecepAgenClien
        End If
        
        N = 3
        If DBLet(miRsAux!RecepPortes, "N") = 1 Then N = 2
        optEule_R(N).Value = True
        
        'Empieza en la 20
        For N = 1 To Me.chkEuler.Count
            chkEuler(N - 1).Value = DBLet(miRsAux.Fields(CInt(N) + 22), "N")
        Next
        
        ' DatosBomUdCaudal,DatosBomTipoRodete"
        kCampo = DBLet(miRsAux!DatosBomTipoRodete, "N")
        If kCampo = 0 Then kCampo = 6 'OTROS
        For N = 4 To 7
            If N = kCampo Then Me.optEule_R(N).Value = True
        Next
        

        cboEulerUdR.ListIndex = -1
        If Not IsNull(miRsAux!DatosBomUdCaudal) Then cboEulerUdR.ListIndex = miRsAux!DatosBomUdCaudal
        
       
        'Combo1.ListIndex = kCampo
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub







Private Function ActualizaBDFichaRep() As String
Dim s As String
Dim N As Byte

    s = CamposSQlFicha()
    s = "REPLACE INTO scarepeu(" & s & ") VALUES ("
    For N = 0 To txtEuler.Count - 1
        s = s & DBSet(txtEuler(N).Text, "T", "S") & ","
    Next

    For N = 1 To Me.chkEuler.Count
        s = s & DBSet(chkEuler(N - 1), "T", "S") & ","
    Next
    
    
    'numrepar ,  RecepAgenClien,RecepPortes, DatosBomUdCaudal,DatosBomTipoRodete"
    s = s & Text1(2).Text & "," & Abs(Me.optEuler(1).Value) & "," & Abs(Me.optEuler(2).Value) & ","
    s = s & Me.cboEulerUdR.ListIndex & ","
    'Rodete
    kCampo = 6
    For N = 4 To 7
        If Me.optEuler(N).Value Then kCampo = N
    Next
    s = s & kCampo & ")"
    
    
   
   conn.Execute s
    
End Function







Private Sub PonerTareasAsociadas()
Dim N As Integer
Dim SQL As String
Dim Horas As Currency
Dim HorasDec As Currency
Dim Vehiculo As String

    SQL = "select sreloj.codtraba,nomtraba,fecha,sreloj.codtipor,nomtipor,horainicio,horafin,calculadas,codflota from sreloj left join stipor on sreloj.codtipor=stipor.codtipor"
    SQL = SQL & " left join straba on straba.codtraba=sreloj.codtraba"
    SQL = SQL & " WHERE codtipom = '" & CodTipoMov & "' and numalbar =" & Text1(0).Text
    SQL = SQL & " ORDER BY fecha,horainicio"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ListView1.ListItems.Clear
    Horas = 0
    N = 0
    Vehiculo = ""
    While Not miRsAux.EOF
        N = N + 1
        ListView1.ListItems.Add , , Format(miRsAux!CodTraba, "0000")
        ListView1.ListItems(N).SubItems(1) = DBLet(miRsAux!NomTraba, "T")
        ListView1.ListItems(N).SubItems(2) = DBLet(miRsAux!codtipor, "T")
        ListView1.ListItems(N).SubItems(3) = DBLet(miRsAux!NomTipor, "T")
        
        ListView1.ListItems(N).SubItems(4) = DBLet(miRsAux!codflota, "T") & " "
        
        If Trim(ListView1.ListItems(N).SubItems(4)) <> "" Then
            If InStr(1, Vehiculo, miRsAux!codflota) = 0 Then Vehiculo = Vehiculo & ", " & DBSet(miRsAux!codflota, "T")
        End If
        
        ListView1.ListItems(N).SubItems(5) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        
        If Not IsNull(miRsAux!calculadas) Then
            Horas = Horas + miRsAux!calculadas
            ListView1.ListItems(N).SubItems(7) = Format(miRsAux!calculadas, FormatoCantidad)
            SQL = Format(Int(miRsAux!calculadas), "00") & ":"
            
            
            HorasDec = Int((miRsAux!calculadas - Int(miRsAux!calculadas)) * 100)
            HorasDec = Round(HorasDec * 0.6, 2)
            SQL = SQL & Format(HorasDec, "00")
            ListView1.ListItems(N).SubItems(6) = SQL
            
            
            
        Else
            ListView1.ListItems(N).SubItems(7) = " "
            ListView1.ListItems(N).SubItems(6) = " "
        End If
        miRsAux.MoveNext
    Wend
    Label1(63).Caption = Format(Horas, FormatoCantidad)
    
    If Horas = 0 Then
        SQL = ""
    Else
        SQL = Format(Int(Horas), "00") & ":"
        HorasDec = Int((Horas - Int(Horas)) * 100)
        HorasDec = Round(HorasDec * 0.6, 2)
        SQL = SQL & Format(HorasDec, "00")
    End If
    Label1(64).Caption = SQL
    
    miRsAux.Close
    
    LimpiaCostes
    If Vehiculo <> "" Then
        Vehiculo = Mid(Vehiculo, 2)
        Vehiculo = "Select codflota,precostehora from sflotas where codflota IN (" & Vehiculo & ") AND precostehora >0 ORDER BY 1"
        miRsAux.Open Vehiculo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
            
            HorasDec = 0
            Vehiculo = miRsAux!codflota
            'Recorremos las fichadas en busca del
            For N = 1 To ListView1.ListItems.Count
                If Trim(ListView1.ListItems(N).SubItems(4)) = Vehiculo Then
                    Horas = ImporteFormateado(ListView1.ListItems(N).SubItems(7))
                    HorasDec = HorasDec + Horas
                End If
            Next
            
            
            Horas = miRsAux!precostehora
            Horas = Horas * HorasDec
            Vehiculo = Vehiculo & "|" & Format(HorasDec, FormatoImporte) & "|" & Format(miRsAux!precostehora, FormatoImporte) & "|" & Format(Horas, FormatoImporte) & "|"
            CostesMatriculas.Add Vehiculo
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    
    
    
    
    
    
    
    
    Set miRsAux = Nothing
    
    
'    For N = 1 To ListView1.ColumnHeaders.Count
'        Debug.Print N & ": " & ListView1.ColumnHeaders(N).Width
'    Next
End Sub





'Costes EULER
Private Sub CargaCostesEuler(limpiar As Boolean)
Dim oldC As Byte
Dim C1 As String
Dim RS As ADODB.Recordset
Dim N As Integer
Dim H As Currency
Dim TotalCostes As Currency
Dim CostesHoras As Currency
Dim CostesVehiculo As Currency
Dim IT As ListItem
Dim Aux1 As Currency
Dim J As Integer

    On Error GoTo eCargaCostesEuler
    
    Me.ListView2.ListItems.Clear
    For N = 66 To 71
        Label1(N).Caption = ""
    Next
    Label1(84).Caption = ""
    Label1(85).Caption = ""
    If limpiar Then Exit Sub
    If Text1(0).Text = "" Then Exit Sub
    
    If Me.SSTab1.Tab <> 4 Then Exit Sub
    
    oldC = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    lblIndicador.Tag = lblIndicador.Caption
    
    lblIndicador.Caption = "Costes alb."
    lblIndicador.Refresh
    N = 0
    TotalCostes = 0
    CostesHoras = 0
    CostesVehiculo = 0
    'Si tiene horas, las aplicamos aqui
    H = 0
    If Label1(63).Caption <> "" And Label1(63).Caption <> "0,00" Then
      
        C1 = ImporteFormateado(Label1(63).Caption)
        H = CCur(C1)
        ListView2.ListItems.Add , , "HOR"
        ListView2.ListItems(1).SubItems(1) = "Horas trabajadas"
        
        For N = 2 To 4
            ListView2.ListItems(1).SubItems(N) = " "
        Next
        ListView2.ListItems(1).SubItems(5) = Format(H, FormatoImporte)
        ListView2.ListItems(1).SubItems(6) = Format(vParamAplic.PrecioHoraCosteEUL, FormatoPrecio)
        H = H * vParamAplic.PrecioHoraCosteEUL
        TotalCostes = TotalCostes + H
        CostesHoras = H
        ListView2.ListItems(1).SubItems(7) = Format(H, FormatoImporte)
        ListView2.ListItems(1).SubItems(8) = "0000"  'ordenacion
        N = 1
    End If
    
    For J = 1 To CostesMatriculas.Count
        N = N + 1
        
        C1 = CostesMatriculas.Item(J)
        Set IT = ListView2.ListItems.Add(, "V" & N, "VEH")
        IT.SubItems(1) = "Coste vehiculo "
        IT.SubItems(2) = RecuperaValor(C1, 1)
        
        
        
        IT.SubItems(5) = RecuperaValor(C1, 2)
        IT.SubItems(6) = RecuperaValor(C1, 3)
        H = ImporteFormateado(RecuperaValor(C1, 4))
        IT.SubItems(7) = Format(H, FormatoImporte)
        TotalCostes = TotalCostes + H
        CostesVehiculo = CostesVehiculo + H
        
        IT.SubItems(8) = Format(N, "0000") 'ordenacion
        
    Next
    
    
    
    'En albaranes
    C1 = "select scaalp.numalbar,scaalp.fechaalb,nomprove,codartic,nomartic,cantidad,precioar,importel,scaalp.Codprove from scaalp,slialp  where"
    C1 = C1 & " scaalp.NumAlbar = slialp.NumAlbar And scaalp.FechaAlb = slialp.FechaAlb And scaalp.Codprove = slialp.Codprove"
    'C1 = C1 & " and codtipomv='" & CodTipoMov
    C1 = C1 & " and codtipomv='" & IIf(EsHistorico, Text1(30).Text, CodTipoMov)
    C1 = C1 & "' and numalbarV=" & Text1(0).Text
    C1 = C1 & " ORDER BY Fechaalb"
    
    Set RS = New ADODB.Recordset
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "ALC"
        IT.SubItems(1) = DBLet(RS!nomprove, "T")
        IT.SubItems(2) = DBLet(RS!Numalbar, "T")
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
               
        If RS!cantidad = 0 Then
            Aux1 = 0
        Else
            Aux1 = RS!ImporteL / RS!cantidad
        End If
        IT.SubItems(6) = Format(Aux1, FormatoPrecio)
        Aux1 = Aux1 - RS!precioar
        If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
        IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numalbar  'ordenacion
        IT.SubItems(9) = RS!codArtic
        
        IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
        TotalCostes = TotalCostes + RS!ImporteL
         
        RS.MoveNext
    Wend
    RS.Close


    'ALbaranes vinculados al pedido
    If Me.Text1(25).Text <> "" Then
        C1 = "select scaalp.numalbar,scaalp.fechaalb,nomprove,codartic,nomartic,cantidad,precioar,importel,scaalp.Codprove from scaalp,slialp  where"
        C1 = C1 & " scaalp.NumAlbar = slialp.NumAlbar And scaalp.FechaAlb = slialp.FechaAlb And scaalp.Codprove = slialp.Codprove"
        C1 = C1 & " and slialp.codclien =" & Text1(4).Text
        C1 = C1 & " and numpedV=" & Text1(25).Text
        C1 = C1 & " ORDER BY Fechaalb"
        
        Set RS = New ADODB.Recordset
        RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            N = N + 1
            
            Set IT = ListView2.ListItems.Add
            IT.Text = "ALC"
            IT.SubItems(1) = DBLet(RS!nomprove, "T")
            IT.SubItems(2) = DBLet(RS!Numalbar, "T")
            IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
            IT.SubItems(4) = DBLet(RS!NomArtic, "T")
            IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
                   
            If RS!cantidad = 0 Then
                Aux1 = 0
            Else
                Aux1 = RS!ImporteL / RS!cantidad
            End If
            IT.SubItems(6) = Format(Aux1, FormatoPrecio)
            Aux1 = Aux1 - RS!precioar
            If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
            IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
            IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numalbar  'ordenacion
            IT.SubItems(9) = RS!codArtic
            
            IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
            TotalCostes = TotalCostes + RS!ImporteL
             
            RS.MoveNext
        Wend
        RS.Close
    End If


    'FACTURAS PROVEEDOR
    lblIndicador.Caption = "Costes fact."
    lblIndicador.Refresh
    C1 = "select scafpc.numfactu,scafpc.fecfactu,nomprove,codartic,nomartic,cantidad,precioar,importel,scafpc.Codprove,slifpc.numalbar,scafpa.fechaalb from"
    C1 = C1 & " scafpc,scafpa,slifpc  where "
    C1 = C1 & " scafpc.codprove = scafpa.codprove And scafpc.numfactu = scafpa.numfactu And scafpc.fecfactu = scafpa.fecfactu "
    C1 = C1 & " AND scafpc.codprove = slifpc.codprove And scafpc.numfactu = slifpc.numfactu And scafpc.fecfactu = slifpc.fecfactu "
    C1 = C1 & " and scafpa.numalbar = slifpc.numalbar"
    C1 = C1 & " and codtipomv='" & IIf(EsHistorico, Text1(30).Text, CodTipoMov)
    C1 = C1 & "' and numalbarV=" & Text1(0).Text
    C1 = C1 & " ORDER BY fecfactu"
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "FAC"
        IT.SubItems(1) = DBLet(RS!nomprove, "T")
        IT.SubItems(2) = DBLet(RS!Numfactu, "T")
        IT.SubItems(3) = Format(RS!FecFactu, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        
        
        If RS!cantidad = 0 Then
            Aux1 = 0
        Else
            Aux1 = RS!ImporteL / RS!cantidad
        End If
        IT.SubItems(6) = Format(Aux1, FormatoPrecio)
        Aux1 = Aux1 - RS!precioar
        If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
        
        
        
        IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
        IT.SubItems(8) = Format(RS!FecFactu, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numfactu  'ordenacion
        TotalCostes = TotalCostes + RS!ImporteL
        IT.SubItems(9) = RS!codArtic
        
        
        IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
        RS.MoveNext
    Wend
    RS.Close

    If Text1(25).Text <> "" Then
        'FACTURAS PROVEEDOR
        lblIndicador.Caption = "Costes fact."
        lblIndicador.Refresh
        C1 = "select scafpc.numfactu,scafpc.fecfactu,nomprove,codartic,nomartic,cantidad,precioar,importel,scafpc.Codprove,slifpc.numalbar,scafpa.fechaalb from"
        C1 = C1 & " scafpc,scafpa,slifpc  where "
        C1 = C1 & " scafpc.codprove = scafpa.codprove And scafpc.numfactu = scafpa.numfactu And scafpc.fecfactu = scafpa.fecfactu "
        C1 = C1 & " AND scafpc.codprove = slifpc.codprove And scafpc.numfactu = slifpc.numfactu And scafpc.fecfactu = slifpc.fecfactu "
        C1 = C1 & " and scafpa.numalbar = slifpc.numalbar"
        C1 = C1 & " and slifpc.codclien=" & Text1(4).Text
        C1 = C1 & " and numpedV=" & Text1(25).Text
        C1 = C1 & " ORDER BY fecfactu"
        RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            N = N + 1
            
            Set IT = ListView2.ListItems.Add
            IT.Text = "FAC"
            IT.SubItems(1) = DBLet(RS!nomprove, "T")
            IT.SubItems(2) = DBLet(RS!Numfactu, "T")
            IT.SubItems(3) = Format(RS!FecFactu, "dd/mm/yyyy")
            IT.SubItems(4) = DBLet(RS!NomArtic, "T")
            IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
            
            
            If RS!cantidad = 0 Then
                Aux1 = 0
            Else
                Aux1 = RS!ImporteL / RS!cantidad
            End If
            IT.SubItems(6) = Format(Aux1, FormatoPrecio)
            Aux1 = Aux1 - RS!precioar
            If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
            
            
            
            IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
            IT.SubItems(8) = Format(RS!FecFactu, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numfactu  'ordenacion
            TotalCostes = TotalCostes + RS!ImporteL
            IT.SubItems(9) = RS!codArtic
            
            
            IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
            RS.MoveNext
        Wend
        RS.Close
    End If


    lblIndicador.Caption = "Adicionales"
    lblIndicador.Refresh
    C1 = "select fechamov ,codartic,numlinea ,nomartic ,cantidad ,precioar,round(cantidad *precioar,2) implin FROM slialb_eu "
    C1 = C1 & " WHERE codtipom='" & CodTipoMov
    C1 = C1 & "' and numalbar=" & Text1(0).Text
    C1 = C1 & " ORDER BY fechamov"
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "MAT"
        IT.SubItems(1) = " "
        IT.SubItems(2) = "L " & RS!numlinea
        IT.SubItems(3) = Format(RS!FechaMov, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        IT.SubItems(6) = Format(RS!precioar, FormatoPrecio)
        IT.SubItems(7) = Format(RS!implin, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaMov, "yymmdd") & "   " & Format(RS!numlinea, "00") 'ordenacion
        IT.SubItems(9) = RS!codArtic
        TotalCostes = TotalCostes + RS!implin
                 
        RS.MoveNext
    Wend
    RS.Close






    'En este albarane.   NO haria falta linkar con sartic
    C1 = "select scaalb.numalbar,scaalb.fechaalb,nomclien,slialb.codartic,slialb.nomartic,cantidad,preciouc,precoste"
    C1 = C1 & " From scaalb, slialb, sartic"
    C1 = C1 & " Where scaalb.NumAlbar = slialb.NumAlbar And scaalb.codtipom = slialb.codtipom And slialb.codArtic = sartic.codArtic"
    C1 = C1 & " and scaalb.codtipom='" & CodTipoMov
    C1 = C1 & "' and scaalb.numalbar=" & Text1(0).Text
    C1 = C1 & " and slialb.precoste<>0"
    C1 = C1 & " ORDER BY Fechaalb"
    

    
    
    Set RS = New ADODB.Recordset
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "ALV"
        IT.SubItems(1) = " "
        IT.SubItems(2) = DBLet(RS!Numalbar, "T")
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        
        Aux1 = DBLet(RS!precoste, "N")   ' DBLet(Rs!precioUC, "N")
        Aux1 = Aux1 * DBLet(RS!cantidad, "N")
        Aux1 = Round(Aux1, 2)
        'IT.SubItems(6) = " " & Format(DBLet(Rs!precioUC, "N"), FormatoPrecio)
        IT.SubItems(6) = " " & Format(DBLet(RS!precoste, "N"), FormatoPrecio)
    
        IT.SubItems(7) = Format(Aux1, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & CodTipoMov & RS!Numalbar  'ordenacion
        TotalCostes = TotalCostes + Aux1
         
        RS.MoveNext
    Wend
    RS.Close




    'SEPT 2018.  Abril 2021.  Miramos por numped, si tiene
    
    For H = 1 To 2
        'Los qu
        lblIndicador.Caption = "Pedidos proveedor. " & H
        lblIndicador.Refresh
        C1 = " select scappr.numpedpr,fecpedpr,nomprove,codartic,nomartic,cantidad,precioar,importel,scappr.Codprove"
        C1 = C1 & " From scappr, slippr where  scappr.numpedpr = slippr.numpedpr "
        If H = 1 Then
            'ALbaran
            C1 = C1 & " and codtipomv='" & CodTipoMov
            C1 = C1 & "' and numalbarV=" & Text1(0).Text
        Else
            'Pedido
            If Me.Text1(25).Text <> "" Then
                C1 = C1 & " and codtipomv IS NULL AND numpedV=" & Text1(25).Text
            Else
                C1 = C1 & " and FALSE"
            End If
        End If
        C1 = C1 & " ORDER BY fecpedpr"
        RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            N = N + 1
            
            Set IT = ListView2.ListItems.Add
            IT.Text = "PED"
            IT.SubItems(1) = DBLet(RS!nomprove, "T")
            IT.SubItems(2) = DBLet(RS!numpedpr, "T")
            IT.SubItems(3) = Format(RS!fecpedpr, "dd/mm/yyyy")
            IT.SubItems(4) = DBLet(RS!NomArtic, "T")
            IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
            
            
            If RS!cantidad = 0 Then
                Aux1 = 0
            Else
                Aux1 = RS!ImporteL / RS!cantidad
            End If
            IT.SubItems(6) = Format(Aux1, FormatoPrecio)
            Aux1 = Aux1 - RS!precioar
            If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
            
            
            
            IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
            IT.SubItems(8) = Format(RS!fecpedpr, "yymmdd") & Format(RS!Codprove, "00000") & Format(RS!numpedpr, "000000") 'ordenacion
            TotalCostes = TotalCostes + RS!ImporteL
            IT.SubItems(9) = RS!codArtic
            
            
            IT.Tag = RS!numpedpr & "|" & RS!fecpedpr & "|" & RS!Codprove & "|"
            RS.MoveNext
        Wend
        RS.Close

    Next


        
    If ListView2.ListItems.Count > 0 Then
    
        Label1(67).Caption = "Total costes"
        Label1(66).Caption = Format(TotalCostes, FormatoImporte)
        Label1(68).Caption = "Costes horas"
        Label1(69).Caption = Format(CostesHoras, FormatoImporte)
        CostesHoras = TotalCostes - CostesHoras - CostesVehiculo
        Label1(70).Caption = "Costes materiales"
        Label1(71).Caption = Format(CostesHoras, FormatoImporte)
        
        Label1(84).Caption = "Costes vehiculos"
        Label1(85).Caption = Format(CostesVehiculo, FormatoImporte)
        
    End If
    
eCargaCostesEuler:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    lblIndicador.Caption = lblIndicador.Tag
    Screen.MousePointer = oldC
End Sub













Private Sub ImprimirCostesEuler()



    Screen.MousePointer = vbHourglass
    If ImprimirLosCostesAlbaranEuler(ListView2, hcoCodTipoM, Text1(0).Text) Then
        
        BotonImprimir 85, False '45: Informe de Albaranes
    End If
    Screen.MousePointer = vbDefault
        
    
End Sub


'Esto mismo esta copiado en el RELOJ, para que creen las carpetas necesarias.
'Tambien la  funcion de cambiar tipo de ALBARAN, si tiene datos las carpetas entonces he copiado en la funcion parte del fuente
Private Sub CrearCarpetaComun2(Modificando As Boolean, NUmeroNuevoDeAlbaranRepacacion As Long)
Dim Referencia As String
Dim OtraCadena As String
Dim CarpetaTipoAlbaran As String
Dim J As Byte
Dim Numalbar As Long

    On Error GoTo eCrearCarpetaComun

    If Not InstalacionEsEulerTaxco Then Exit Sub
    If EsHistorico Then Exit Sub
    
 '   If NUmeroNuevoDeAlbaranRepacacion = 0 Then
        J = 0
        CarpetaTipoAlbaran = ""
        If hcoCodTipoM = "ALR" Then J = 1: CarpetaTipoAlbaran = "REPARACIONES"
        If hcoCodTipoM = "ALO" Then J = 2: CarpetaTipoAlbaran = "orden de trabajo"
        If hcoCodTipoM = "ALE" Then J = 3: CarpetaTipoAlbaran = "trabajo exterior"
        If J = 0 Then Exit Sub
 '    End If
    
    ' Abril 2021
    'Se añaden tambien las opciones de carpeta para
    'trabajo exterior / orden de trabajo
    If EulerParam = "" Then Exit Sub

     'Z:(VALENCIA O ZALDIVIA)REPARACIONES\YYYY\NNNNNNN\
     
    If hcoCodTipoM = "ALR" Then
        'ALR. Van por Valencia-Zaldibia
        CadenaConsulta = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        If CadenaConsulta = "" Then Err.Raise 513, , "Obteniendo centro trabajo (coste) trabajador " & Text1(3).Text
        
        If CadenaConsulta = "1" Then
            CadenaConsulta = "ZALDIBIA"
        Else
            CadenaConsulta = "VALENCIA"
        End If
    
    
        'txtAnterior = EulerParam & "\REPARACIONES\" & CadenaConsulta & "\" & Year(CDate(Text1(1).Text))  Abril2021
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & CadenaConsulta & "\" & Year(CDate(Text1(1).Text))
    Else
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & Year(CDate(Text1(1).Text))
    End If
        
    If Dir(txtAnterior, vbDirectory) = "" Then MkDir txtAnterior
    
    
    'Reemplazamos los carteres especiales de carpeta \/:*?<>| por espacios en blanco
    Referencia = Text1(13).Text
    For J = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", J, 1), " ")
    Next
    
    If NUmeroNuevoDeAlbaranRepacacion = 0 Then
        Numalbar = Val(Text1(0).Text)
    Else
        Numalbar = NUmeroNuevoDeAlbaranRepacacion
    End If
    
    Referencia = Trim(Format(Numalbar, "0000000") & " " & Referencia)
     
    
    
    CadenaConsulta = txtAnterior & "\" & Referencia
    If Dir(CadenaConsulta, vbDirectory) <> "" Then
        If Not Modificando Then Err.Raise 513, , "Ya existe la carpeta: " & CadenaConsulta
    Else
        'NO existe vemos, si existe una carpeta que empieza por
        OtraCadena = txtAnterior & "\" & Format(Numalbar, "0000000") & "*"
        OtraCadena = Dir(OtraCadena, vbDirectory)
        If OtraCadena <> "" Then
            OtraCadena = txtAnterior & "\" & OtraCadena
            Name OtraCadena As CadenaConsulta
            OtraCadena = ""
        Else
            MkDir CadenaConsulta
        End If
    End If
    
    'Documentacion interna
    CadenaConsulta = CadenaConsulta & "\Documentacion interna" 'Si cambias aqui, cambiar tambien en mas sitios !!!!
    If Dir(CadenaConsulta, vbDirectory) <> "" Then
        If Not Modificando Then Err.Raise 513, , "Ya existe la carpeta: " & CadenaConsulta
    Else
        MkDir CadenaConsulta
    End If
eCrearCarpetaComun:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description & vbCrLf & CadenaConsulta & vbCrLf & OtraCadena
End Sub

Private Function ComprobarSiHayDocumentosEnCarpeta() As Boolean
Dim Referencia As String
Dim OtraCadena As String
Dim CarpetaTipoAlbaran As String
Dim J As Byte
Dim Numalbar As Long

    On Error GoTo eCrearCarpetaComun
    
    ComprobarSiHayDocumentosEnCarpeta = False

    If Not InstalacionEsEulerTaxco Then Exit Function
    If EsHistorico Then Exit Function
    
    If EulerParam = "" Then Exit Function
    
    J = 0
    CarpetaTipoAlbaran = ""
    If hcoCodTipoM = "ALR" Then J = 1: CarpetaTipoAlbaran = "REPARACIONES"
    If hcoCodTipoM = "ALO" Then J = 2: CarpetaTipoAlbaran = "orden de trabajo"
    If hcoCodTipoM = "ALE" Then J = 3: CarpetaTipoAlbaran = "trabajo exterior"
    If J = 0 Then Exit Function

    
    

     'Z:(VALENCIA O ZALDIVIA)REPARACIONES\YYYY\NNNNNNN\
     
    If hcoCodTipoM = "ALR" Then
        'ALR. Van por Valencia-Zaldibia
        CadenaConsulta = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        If CadenaConsulta = "" Then Err.Raise 513, , "Obteniendo centro trabajo (coste) trabajador del albaran " & Text1(3).Text
        
        If CadenaConsulta = "1" Then
            CadenaConsulta = "ZALDIBIA"
        Else
            CadenaConsulta = "VALENCIA"
        End If
    
    
        'txtAnterior = EulerParam & "\REPARACIONES\" & CadenaConsulta & "\" & Year(CDate(Text1(1).Text))  Abril2021
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & CadenaConsulta & "\" & Year(CDate(Text1(1).Text))
    Else
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & Year(CDate(Text1(1).Text))
    End If
        
    If Dir(txtAnterior, vbDirectory) = "" Then Exit Function
    
    
    'Reemplazamos los carteres especiales de carpeta \/:*?<>| por espacios en blanco
    Referencia = Text1(13).Text
    For J = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", J, 1), " ")
    Next
    
    
    Numalbar = Val(Text1(0).Text)
    Referencia = Trim(Format(Numalbar, "0000000") & " " & Referencia)
     
    
    
    CadenaConsulta = txtAnterior & "\" & CStr(Referencia) & "\"
    Referencia = Dir(CadenaConsulta, vbDirectory)      ' Recupera la primera entrada.
    
    
    If Referencia = "" Then
        Referencia = Trim(Format(Numalbar, "0000000") & "*")
        CadenaConsulta = txtAnterior & "\" & CStr(Referencia)
        Referencia = Dir(CadenaConsulta, vbDirectory)      ' Recupera la primera entrada.
        If Referencia <> "" Then
            CadenaConsulta = txtAnterior & "\" & CStr(Referencia) & "\"
            
        Else
            Exit Function 'no hay ninguna carpeta para ese año y numero de reparacion
        End If
    End If
    
    
    For J = 1 To 2
        
        If J = 2 Then
            CadenaConsulta = CadenaConsulta & "Documentacion interna\"
            Referencia = "OK"
        End If
        
        
        If Referencia <> "" Then Referencia = Dir(CadenaConsulta & "*.*", vbArchive)
        
        
        Do While Referencia <> ""   ' Inicia el bucle.
            ' Ignora el directorio actual y el que lo abarca.
            
              
                    
            ComprobarSiHayDocumentosEnCarpeta = True
            Exit For
    
            'Referencia = Dir   ' Obtiene siguiente entrada.
        Loop
    
    Next J
    
    
eCrearCarpetaComun:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description & vbCrLf & CadenaConsulta & vbCrLf & OtraCadena
End Function





Private Sub PonerImagenFirma()
Dim C As String
    On Error GoTo ePonerImagenFirma
    
    If CarpetaImagenesEULER = "" Then Exit Sub
    
    If PrimeraVez Then
        If Dir(CarpetaImagenesEULER, vbDirectory) = "" Then
            MsgBox "No existe carpeta: " & CarpetaImagenesEULER, vbExclamation
            CarpetaImagenesEULER = ""
        
        Else
            C = DevuelveDesdeBD(conAri, "carpetafirmas", "eulerparam", "1", "1")
            CarpetaImagenesEULER = CarpetaImagenesEULER & "\" & C
            
                
            If Dir(CarpetaImagenesEULER, vbDirectory) = "" Then
                MsgBox "No existe carpeta: " & CarpetaImagenesEULER, vbExclamation
                CarpetaImagenesEULER = ""
            End If
        End If
        Exit Sub
    End If
    
    If Modo <> 2 Then Exit Sub
    
    C = CarpetaImagenesEULER & "\" & Mid(Data1.Recordset!codtipom & "   ", 1, 3) & Format(Data1.Recordset!Numalbar, "0000000") & ".png"
    If Dir(C, vbArchive) = "" Then C = ""
    If C <> "" Then
        imgFirmaRecep.visible = True
        imgFirmaRecep.Tag = C
    Else
        imgFirmaRecep.Tag = ""
         imgFirmaRecep.visible = False
    End If
    
    Exit Sub
ePonerImagenFirma:
    Err.Clear
    CarpetaImagenesEULER = ""
End Sub







Private Sub HacerCambioTipoAlbaran()
        
    'Primera comprpobacion
    If ComprobarSiHayDocumentosEnCarpeta Then
        MsgBox "Tiene documentos asociados al albaran", vbExclamation
        Exit Sub
    End If
        
        
        
    cadList = DevuelveDesdeBD(conAri, "IncidenciaCambiarTipo", "eulerparam", "1", "1")
    If cadList = "" Then Exit Sub
        
        
        
        
        
        
    CadenaDesdeOtroForm = Data1.Recordset!codtipom
    frmVarios.Opcion = 15
    frmVarios.Show vbModal
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    
    If Not hazCambioTipodeAlbaran Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
        PosicionarDataTrasEliminar
        Espera 0.25
        ejecutar txtAnterior, False
        
        
        Screen.MousePointer = vbHourglass
     
        
        cadList = "Cambio realizado.  " & cadList
        
        
        MsgBox cadList, vbInformation
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Function hazCambioTipodeAlbaran() As Boolean
Dim vTipoMov As CTiposMov
Dim Aux As String
Dim SQL As String
Dim A_Reparacion As Boolean
On Error GoTo ehazCambioTipodeAlbaran

    AlmacenLineas = 0
    Set vTipoMov = New CTiposMov
    Aux = CadenaDesdeOtroForm
    A_Reparacion = False
    If CadenaDesdeOtroForm = "ALR" Then
        
        Aux = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", Text1(3).Text)
        If Aux = "10" Then
            Aux = "CAR"
        Else
            Aux = "ALR"
        End If
        A_Reparacion = True
    End If
    If Not vTipoMov.Leer(Aux) Then Err.Raise 513, , "Error leyendo tipo movimiento: " & Aux
    Aux = CadenaDesdeOtroForm
    'Fecha elim y codinicidencia eliminar
    SQL = PonerTrabajadorConectado("")
    If SQL = "" Then Err.Raise 513, , "Imposible obtener trabajador conectado"
    cadList = DBSet(Now, "F") & " as fechaelim, " & SQL & " as trabaelim ,'" & cadList & "' as codinicid"
    
    'Lo pasamos a HISTORICO
    SQL = ObtenerWhereCP(False)
    If Not ActualizarElTraspasoSinBorrar(txtAnterior, SQL, CodTipoMov, cadList) Then Err.Raise 513, , txtAnterior
    
    
  
    
    txtAnterior = "Actualizando TIPO y numero albaran"
    conn.Execute "SET FOREIGN_KEY_CHECKS=0"
    AlmacenLineas = 1
    
    cadList = Trim("TipoAnt: " & Data1.Recordset!codtipom & Format(Data1.Recordset!Numalbar, "0000000") & "     " & Text1(43).Text)
    
    cadList = ", observacrm = " & DBSet(cadList, "T")
    cadList = "UPDATE scaalb set codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & cadList
    cadList = cadList & " WHERE " & SQL
    conn.Execute cadList
    
    
    
    'UPDATES y ELMINAR scaalb eu y slialb_eu
    txtAnterior = "cambiando datos slialb_eu"
    cadList = " WHERE " & Replace(SQL, "scaalb", "slialb_eu")
    cadList = "UPDATE slialb_eu set codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & cadList
    conn.Execute cadList
    
    txtAnterior = "cambiando datos slialb_eu2"
    cadList = " WHERE " & Replace(SQL, "scaalb", "slialb_eu2")
    cadList = "UPDATE slialb_eu2 set codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & cadList
    conn.Execute cadList
    
    
    
    txtAnterior = "actualizando datos scaalb_eu"
    cadList = " WHERE " & Replace(SQL, "scaalb", "scaalb_eu")
    cadList = "UPDATE scaalb_eu set codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & cadList
    conn.Execute cadList
    
    
    
    
    
    
    
    
    
    
    
    cadList = Replace(SQL, "scaalb", "slialb")
    cadList = " WHERE " & cadList
    cadList = "UPDATE slialb set codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & " " & cadList
    conn.Execute cadList
    
    
    
    txtAnterior = "Actualizando movimientos"
    SQL = " fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND detamovi=" & DBSet(Data1.Recordset!codtipom, "T")
    SQL = SQL & " AND document='" & Format(Data1.Recordset!Numalbar, "0000000") & "'"
    cadList = "detamovi =" & DBSet(Aux, "T") & " , document ='" & Format(vTipoMov.Contador + 1, "0000000") & "'"
    SQL = "UPDATE smoval SET " & cadList & " WHERE " & SQL
    conn.Execute SQL
    
    
    
    'Las horas
    txtAnterior = "Tareas reloj"
    SQL = ObtenerWhereCP(True)
    cadList = Replace(SQL, "scaalb", "sreloj")
    SQL = "update sreloj  SET codtipom=" & DBSet(Aux, "T") & ", numalbar = " & vTipoMov.Contador + 1 & cadList
    conn.Execute SQL
    
    txtAnterior = "Pedidos proveedor"
    SQL = " FechaAlbV=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND codtipomV=" & DBSet(Data1.Recordset!codtipom, "T")
    SQL = SQL & " AND NumalbarV=" & Format(Data1.Recordset!Numalbar, "0000000")
    cadList = "codtipomV =" & DBSet(Aux, "T") & " , NumalbarV ='" & Format(vTipoMov.Contador + 1, "0000000") & "'"
    SQL = "UPDATE slippr SET " & cadList & " WHERE " & SQL
    conn.Execute SQL
    
    
        
    txtAnterior = "Albaranes proveedor"
    SQL = " FechaAlbV=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND codtipomV=" & DBSet(Data1.Recordset!codtipom, "T")
    SQL = SQL & " AND NumalbarV=" & Format(Data1.Recordset!Numalbar, "0000000")
    cadList = "codtipomV =" & DBSet(Aux, "T") & " , NumalbarV ='" & Format(vTipoMov.Contador + 1, "0000000") & "'"
    SQL = "UPDATE slialp SET " & cadList & " WHERE " & SQL
    conn.Execute SQL
    
    
    txtAnterior = "Facturas proveedor"
    SQL = " FechaAlbV=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND codtipomV=" & DBSet(Data1.Recordset!codtipom, "T")
    SQL = SQL & " AND NumalbarV=" & Format(Data1.Recordset!Numalbar, "0000000")
    cadList = "codtipomV =" & DBSet(Aux, "T") & " , NumalbarV ='" & Format(vTipoMov.Contador + 1, "0000000") & "'"
    SQL = "UPDATE slialp SET " & cadList & " WHERE " & SQL
    conn.Execute SQL
    
    
    
    txtAnterior = "LOG"
    SQL = "Origen: " & Data1.Recordset!codtipom & " " & Format(Data1.Recordset!Numalbar, "0000000") & "    " & Format(Data1.Recordset!FechaAlb, "dd/mm/yyyy") & vbCrLf
    SQL = SQL & "Destino: " & Aux & " " & Format(vTipoMov.Contador + 1, "0000000")
    cadList = "Albaran creado: " & Aux & Format(vTipoMov.Contador + 1, "0000000")
    Set LOG = New cLOG
    LOG.Insertar 38, vUsu, SQL
    Set LOG = Nothing
        
    'VA todo bien    NumTrabajExterno,NumParteTrabajo,NumAlbaranVenta
    'Ahora grabaremos en scaal_eu, el dato de DONDE viene
    
    'EN reparaciones , los campos que muestra SON:
    '        NumTrabajExterno,NumParteTrabajo

    'ALO.ALE.AL.
    '       NumParteTrabajo NumTrabajExterno
    
    
    SQL = ""
    
    If hcoCodTipoM = "ALV" Then
        SQL = "NumAlbaranVenta"
        
    ElseIf hcoCodTipoM = "ALR" Then
        SQL = "NumReparacion"
    ElseIf hcoCodTipoM = "ALO" Then
        SQL = "NumParteTrabajo"
    ElseIf hcoCodTipoM = "ALE" Then
        SQL = "NumTrabajExterno"
    End If
    If SQL <> "" Then
     
        
    
        'Cuando acabe de actualziar el todo, pdatearemos scaalb_eu poniendo DE DONDE viene eñ nuevo albaran
        SQL = SQL & " = " & Data1.Recordset!Numalbar
        SQL = "UPDATE scaalb_eu SET " & SQL & " WHERE codtipom=" & DBSet(Aux, "T") & " AND numalbar = " & vTipoMov.Contador + 1
        'execute despues de todo, despues de hacer el begin trans
    End If
    txtAnterior = SQL
    
    vTipoMov.IncrementarContador vTipoMov.TipoMovimiento
    
    
        
    
    
    Espera 0.5
    Aux = ""
    If A_Reparacion Then
        Aux = "OK"
    Else
        If vTipoMov.TipoMovimiento = "ALE" Or vTipoMov.TipoMovimiento = "ALO" Then Aux = "OK"
    End If
    
    If Aux <> "" Then
        Aux = hcoCodTipoM
         
        hcoCodTipoM = vTipoMov.TipoMovimiento
        If EulerParam = "" Then EulerParam = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
        CrearCarpetaComun2 True, vTipoMov.Contador
        
        
        hcoCodTipoM = Aux
        txtAnterior = SQL
        
    End If
        
    
    
    
    
    
    hazCambioTipodeAlbaran = True
    
ehazCambioTipodeAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, txtAnterior, Err.Description
    Set vTipoMov = Nothing
    If AlmacenLineas = 1 Then ejecutar "SET FOREIGN_KEY_CHECKS=1", False
    AlmacenLineas = 0
    
    If hcoCodTipoM <> "ALR" Then EulerParam = ""
    
    
    'txtAnterior: llevará un utlimo UPDATE
    If Not hazCambioTipodeAlbaran Then txtAnterior = ""
    
    
    
End Function


Private Sub PonerLineasAlbaranEULER()
Dim SQL As String
Dim N As Integer
Dim vImpo As Currency

    Me.lwEulerLineas.ListItems.Clear
    lwEulerLineas.Tag = ""
    
    If EsHistorico Then
         SQL = " WHERE codtipom='" & Data1.Recordset!codtipom & "' AND numalbar =" & Data1.Recordset!Numalbar
        SQL = "slhalb_eu2" & SQL
    Else
        SQL = ObtenerWhereCP(True)
        SQL = Replace(SQL, "scaalb.", "")
        SQL = "slialb_eu2" & SQL
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    SQL = "Select codtipom,numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel FROM  " & SQL
    SQL = SQL & " order by numlinea"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vImpo = 0
    N = 0
    While Not miRsAux.EOF
        N = N + 1
        lwEulerLineas.ListItems.Add , "k" & Format(miRsAux!numlinea, "000") & miRsAux!Numalbar, miRsAux!Articulo
        lwEulerLineas.ListItems(N).SubItems(1) = Replace(miRsAux!descrarticulo, vbCrLf, " ")
        lwEulerLineas.ListItems(N).SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
        lwEulerLineas.ListItems(N).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        lwEulerLineas.ListItems(N).SubItems(4) = Format(miRsAux!dtoline1, FormatoCantidad)
        lwEulerLineas.ListItems(N).SubItems(5) = Format(miRsAux!ImporteL, FormatoCantidad)
        lwEulerLineas.ListItems(N).ToolTipText = miRsAux!descrarticulo
        vImpo = vImpo + miRsAux!ImporteL
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If N > 0 Then
        'Tiene lineas y NO suma el burto
        If ImporteFormateado(Text3(56).Text) <> vImpo Then
            SQL = "B.imponible albaran: " & Text3(56).Text & vbCrLf
            SQL = SQL & "Suma lineas:           " & Format(vImpo, FormatoImporte)
            vImpo = ImporteFormateado(Text3(56).Text) - vImpo
            SQL = SQL & vbCrLf & "Diferencia:              " & Format(vImpo, FormatoImporte)
            lwEulerLineas.Tag = "Importes lineas impresion: " & vbCrLf & vbCrLf & SQL
            MsgBox SQL, vbExclamation
        End If
    End If
        

 
End Sub





Private Sub txtTaxco_GotFocus(Index As Integer)
    ConseguirFoco txtTaxco(Index), Modo
End Sub

Private Sub txtTaxco_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtTaxco_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtTaxco(Index), Modo) Then Exit Sub
    
    If Index = 7 Then
        If Not PonerFormatoEntero(txtTaxco(Index)) Then txtTaxco(Index).Text = ""
        Me.Text5(1).Text = txtTaxco(Index).Text
        
    Else
        txtTaxco(Index).Text = UCase(txtTaxco(Index).Text)
        If Index = 0 Then Me.Text5(0).Text = txtTaxco(Index).Text
    End If
    
End Sub



Private Sub AsignarDatosVehiculoReparacion(ElSQL As String)
Dim SQL As String
Dim N As Integer
Dim AlbaranFactura As Boolean

    
    
    N = 0
    SQL = ""
    AlbaranFactura = InStr(1, ElSQL, "smatriculataller") = 0
    
    Set miRsAux = New ADODB.Recordset
    If AlbaranFactura Then
    
    
            SQL = "'' vacio0, '' vacio1, '' vacio2,'' codtraba,codclien,nomclien,nifclien,telclien,domclien,codpobla,pobclien,proclien,"
            SQL = SQL & " coddirec,'' referenc,codforpa,dtoppago,dtognral,codagent,'' observa01,'' observa02,'' observa03, ''observa04, '' observa05,"
            SQL = SQL & " '' numofert,'' fecofert,'' numpedcl,'' fecpedcl,'' codtrab1,'' codtrab2,'' codenvio, '' vacio30, '' vacio31, '' vacio32, '' vacio33,"
            SQL = SQL & " '' cantidkm,'' fecfactu,'' numfactu,'' codtipmf,'' numtermi,'' numventa,'' aportacion,'' fecenvio,'' actuacion,'' observacrm,'' fechaaux,'' fechaent"
            SQL = SQL & "  ,'' perrecep,'' dnirecep,'' latitud,'' longitud"
            
            
            
            txtAnterior = "scafac WHERE "
            If InStr(1, ElSQL, "scaalb") > 0 Then
                txtAnterior = "scaalb WHERE "
            End If
                
            SQL = "SELECT " & SQL & " FROM " & txtAnterior & ElSQL
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "MAL . Es EOF"
            Else
                 For N = 0 To Text1.Count - 1
                    Text1(N).Text = DBLet(miRsAux.Fields(N), "T")
                Next
            End If
            miRsAux.Close
        
            'Y ahora del albaran
            
            txtAnterior = "scafac_eu scafac"
            If InStr(1, ElSQL, "scaalb") > 0 Then txtAnterior = "scaalb_eu scaalb "
            SQL = CamposSQlFicha
            
            
            
            SQL = "SELECT " & SQL & " FROM " & txtAnterior & " WHERE " & Replace(ElSQL, "#", "")
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "MAL . Es EOF"
            Else
                For N = 0 To Me.txtTaxco.Count - 1
                    If N <> 7 Then
                        txtTaxco(N).Text = DBLet(miRsAux.Fields(N), "T")
                    End If
                Next
                
                For N = 7 To 10
                    txtEuler(N).Text = "" ' DBLet(miRsAux.Fields(CInt(N + 1)), "T")
                Next
                
                Text5(0).Text = txtTaxco(0).Text
                
            End If
            miRsAux.Close
    Else
    
            Text1(4).Text = ""
            'bombamarca,bombaModelo,motorModelo,motormarca,ReferPedido,RecepAgenCliMat,RecpNumExp,numrepar, observaciones,
            'NumReparacion ,NumTrabajExterno, NumAlbaranVenta , TipoPortes, codtipom, numalbar
            SQL = "matricula bombamarca,bastidor bombaModelo,marca_modelo motorModelo,motor motormarca,'' ReferPedido,'' RecepAgenCliMat,'' RecpNumExp,0 numrepar,'' observaciones,"
            SQL = SQL & " '' NumReparacion , '' NumTrabajExterno, '' NumAlbaranVenta, '' TipoPortes,codclien"
    
            SQL = "SELECT " & SQL & " FROM smatriculataller WHERE " & ElSQL
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "MAL . Es EOF"
            Else
                For N = 0 To Me.txtTaxco.Count - 1
                    If N <> 7 Then
                        txtTaxco(N).Text = DBLet(miRsAux.Fields(N), "T")
                    End If
                Next
                
                For N = 7 To 10
                    txtEuler(N).Text = "" ' DBLet(miRsAux.Fields(CInt(N + 1)), "T")
                Next
                
                Text5(0).Text = txtTaxco(0).Text
                
                
                Text1(4).Text = miRsAux!codClien
                
            
                miRsAux.Close
    
    
                SQL = "'' vacio0, '' vacio1, '' vacio2,'' codtraba,codclien,nomclien,nifclien,telclie1,domclien,codpobla,pobclien,proclien,"
                SQL = SQL & " '' coddirec,'' referenc,codforpa,dtoppago,dtognral,codagent,'' observa01,'' observa02,'' observa03, ''observa04, '' observa05,"
                SQL = SQL & " '' numofert,'' fecofert,'' numpedcl,'' fecpedcl,'' codtrab1,'' codtrab2,'' codenvio, '' vacio30, '' vacio31, '' vacio32, '' vacio33,"
                SQL = SQL & " '' cantidkm,'' fecfactu,'' numfactu,'' codtipmf,'' numtermi,'' numventa,'' aportacion,'' fecenvio,'' actuacion,'' observacrm,'' fechaaux,'' fechaent"
                SQL = SQL & "  ,'' perrecep,'' dnirecep,'' latitud,'' longitud"
                
                
                
             
                    
                SQL = "SELECT " & SQL & " FROM sclien where codclien = " & Text1(4).Text
                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "MAL . Es EOF el Cliente :" & Text1(4).Text, vbExclamation
                Else
                     For N = 0 To Text1.Count - 1
                        Text1(N).Text = DBLet(miRsAux.Fields(N), "T")
                    Next
                End If
                
            End If
            miRsAux.Close
        
            
            
            
            
            
            
    
    
    
    
    End If
     
    BloquearDatosCliente False
    
    
     'Tipo facturacion y codigo envio
     If Text1(4).Text <> "" Then
        SQL = "Select sclien.codenvio,nomenvio,TipoFact from sclien,senvio where senvio.codenvio=sclien.codenvio and codclien= " & Text1(4).Text
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then
            Text1(29).Text = miRsAux!CodEnvio
            Text2(29).Text = miRsAux!nomEnvio
            cboFacturacion.ListIndex = miRsAux!TipoFact
        End If
        miRsAux.Close
    End If
        
    txtAnterior = ""
    LineaIntercalar = 0
    Set miRsAux = Nothing
End Sub


Private Function AvisosAlbaranesEfectifivoPendienteFacturar(EsteAlbaran As Boolean) As Boolean
Dim RN As ADODB.Recordset
Dim C As String
    
    
    AvisosAlbaranesEfectifivoPendienteFacturar = False
    If EsHistorico Then Exit Function
    If vParamAplic.NumeroInstalacion <> vbTaxco Then Exit Function
    
    If EsteAlbaran Then
        If Data2.Recordset.EOF Then Exit Function
    End If
    
    C = "Select codtipom,numalbar,codclien,nomclien from scaalb where codtipom=" & DBSet(hcoCodTipoM, "T")
    If EsteAlbaran Then
        'Estoy pasando de lineas a cabecera
        C = C & " and numalbar=" & Text1(0).Text
    Else
        'He mandado cerrar el form
        C = C & " and fechaalb=" & DBSet(Now, "F")
        'If hcoCodTipoM = "ALO" Then C = C & "  and factursn=0"
        If EsTipoAlbaranFacturarInmediatoTaxco Then C = C & "  and factursn=0"
        
        C = C & " and numalbar IN (select numalbar from slialb WHERE codtipom=" & DBSet(hcoCodTipoM, "T") & ")"
    End If
    
    C = C & " and codforpa in (select codforpa from sforpa where tipforpa =0)"
    Set RN = New ADODB.Recordset
    RN.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    While Not RN.EOF
        If Len(C) < 700 Then C = C & RN!codtipom & Format(RN!Numalbar, "000000") & "  Cli:" & Format(RN!codClien, "00000") & " " & RN!NomClien & vbCrLf
        
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
    If C <> "" Then
        C = "PENDIENTE DE FACTURAR" & vbCrLf & String(40, "=") & vbCrLf & C
        If Not EsteAlbaran Then MsgBox C, vbExclamation
        AvisosAlbaranesEfectifivoPendienteFacturar = True
    End If
End Function

Private Sub VerHistoricoTaxco()
    frmMensajes.OpcionMensaje = 26
    frmMensajes.Show vbModal
End Sub



Private Sub CargaTablaTmpTotales()
Dim C As String
Dim K As Byte
    On Error GoTo eCargaTablaTmpTotales
    
    C = "DELETE FROM tmptotalalbaran WHERE codtipom = " & DBSet(Text1(30).Text, "T") & " AND numalbar =" & Text1(0).Text
    conn.Execute C
    
    C = ""
    For K = 1 To 3
        kCampo = 36 + K '37-38  -39
        
        If Trim(Text3(kCampo).Text) <> "" Then
            '(`codusu`,`codtipom`,`numalbar`,`numlinea`,
            C = C & ", (" & vUsu.Codigo & "," & DBSet(Text1(30).Text, "T") & "," & Text1(0).Text & "," & K
            '`base`, ,`porcentaje`,`imporiva`,
            C = C & ", " & DBSet(Text3(kCampo + 6).Text, "N", "N") & "," & DBSet(Text3(kCampo + 3).Text, "N") & "," & DBSet(Text3(kCampo + 9).Text, "N")
            
            '`porc_recargo`,`ivarecargo`,`totallinea`)
            C = C & "," & DBSet(Text3(kCampo + 12).Text, "N") & "," & DBSet(Text3(kCampo + 15).Text, "N", "N") & ","
            
            'En el primero, K=1  , ponemos totalfac
            C = C & IIf(K = 1, DBSet(Text3(55).Text, "N", "N"), 0) & ")"
            
        End If
    Next
    
    If C <> "" Then
        C = Mid(C, 2) 'quitamos la primera coma
        C = "INSERT INTO tmptotalalbaran (`codusu`,`codtipom`,`numalbar`,`numlinea`,`base`,`porcentaje`,`imporiva`,`porc_recargo`,`ivarecargo`,`totallinea`) VALUES " & C
        
        conn.Execute C
    End If
    Exit Sub
eCargaTablaTmpTotales:
    MuestraError Err.Number, , Err.Description
End Sub
