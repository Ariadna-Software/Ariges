VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacClientesGr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   18060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmFacClientesGr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   18060
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
      Left            =   7200
      TabIndex        =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   361
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
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cobros"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Errores N�Factura"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas sin Asiento"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   195
      Left            =   11400
      TabIndex        =   352
      Top             =   360
      Width           =   2415
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
      Left            =   3840
      TabIndex        =   292
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   293
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
      TabIndex        =   291
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   362
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   120
      TabIndex        =   105
      Top             =   1560
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   16113
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      TabsPerRow      =   12
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
      TabPicture(0)   =   "frmFacClientesGr.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(114)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(14)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(34)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(15)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(37)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgBuscar(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgBuscar(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgBuscar(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgWeb"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(16)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "imgFecha(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "imgBuscar(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(93)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "imgBuscar(17)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(36)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(19)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(4)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(5)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(6)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(8)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(22)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(11)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(12)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(9)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text2(12)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(10)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text2(11)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text2(10)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(13)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "chkClienteV"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(54)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(60)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(7)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboPais"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(45)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "frameDptoAdmon"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "frameDptoVentas"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "frameDptoDirec"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "Asegur"
      TabPicture(1)   =   "frmFacClientesGr.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAsegurados"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Direcciones"
      TabPicture(2)   =   "frmFacClientesGr.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameDirecciones"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Dir. Envio"
      TabPicture(3)   =   "frmFacClientesGr.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameDireccionEnvio"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Documentos"
      TabPicture(4)   =   "frmFacClientesGr.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameVisorDocumentos"
      Tab(4).Control(1)=   "FrameNavegaDoc"
      Tab(4).Control(2)=   "FramePuntos"
      Tab(4).Control(3)=   "Text1(46)"
      Tab(4).Control(4)=   "lw1"
      Tab(4).Control(5)=   "cmdCatalogo"
      Tab(4).Control(6)=   "imgDocumentos"
      Tab(4).Control(7)=   "LabelDoc"
      Tab(4).Control(8)=   "imgFecha(3)"
      Tab(4).Control(9)=   "Label3"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "CRM"
      TabPicture(5)   =   "frmFacClientesGr.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LabelCRM"
      Tab(5).Control(1)=   "imgCrm"
      Tab(5).Control(2)=   "lwCRM"
      Tab(5).Control(3)=   "cmdAccCRM(2)"
      Tab(5).Control(4)=   "cmdAccCRM(1)"
      Tab(5).Control(5)=   "cmdAccCRM(0)"
      Tab(5).Control(6)=   "FrameNavegaCRM"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "Datos contacto"
      TabPicture(6)   =   "frmFacClientesGr.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame4"
      Tab(6).Control(1)=   "frameAdmon"
      Tab(6).Control(2)=   "frameComercial"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Renting"
      TabPicture(7)   =   "frmFacClientesGr.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(88)"
      Tab(7).Control(1)=   "Label1(89)"
      Tab(7).Control(2)=   "Label1(90)"
      Tab(7).Control(3)=   "Label4(0)"
      Tab(7).Control(4)=   "imgBuscar(24)"
      Tab(7).Control(5)=   "DataGrid2"
      Tab(7).Control(6)=   "txtauxRent(1)"
      Tab(7).Control(7)=   "txtauxRent(0)"
      Tab(7).Control(8)=   "txtauxRent(2)"
      Tab(7).Control(9)=   "cmdRenting(0)"
      Tab(7).Control(10)=   "cmdRenting(1)"
      Tab(7).Control(11)=   "txtauxRent(3)"
      Tab(7).Control(12)=   "txtauxRent(4)"
      Tab(7).Control(13)=   "txtauxRent(5)"
      Tab(7).Control(14)=   "txtauxRent(6)"
      Tab(7).Control(15)=   "txtauxRent(11)"
      Tab(7).Control(16)=   "txtauxRent(7)"
      Tab(7).Control(17)=   "cmdRenting(2)"
      Tab(7).Control(18)=   "txtauxRent(8)"
      Tab(7).Control(19)=   "txtauxRent(9)"
      Tab(7).Control(20)=   "txtauxRent(10)"
      Tab(7).Control(21)=   "FrameToolAux(4)"
      Tab(7).ControlCount=   22
      TabCaption(8)   =   "tfno"
      TabPicture(8)   =   "frmFacClientesGr.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label1(95)"
      Tab(8).Control(1)=   "Label1(99)"
      Tab(8).Control(2)=   "Label1(98)"
      Tab(8).Control(3)=   "Label2(1)"
      Tab(8).Control(4)=   "Label1(96)"
      Tab(8).Control(5)=   "Label1(97)"
      Tab(8).Control(6)=   "imgBuscar(18)"
      Tab(8).Control(7)=   "imgBuscar(19)"
      Tab(8).Control(8)=   "imgBuscar(20)"
      Tab(8).Control(9)=   "Label1(100)"
      Tab(8).Control(10)=   "Label1(101)"
      Tab(8).Control(11)=   "Label1(102)"
      Tab(8).Control(12)=   "imgBuscar(21)"
      Tab(8).Control(13)=   "imgFechaTf(9)"
      Tab(8).Control(14)=   "imgFechaTf(10)"
      Tab(8).Control(15)=   "Label1(103)"
      Tab(8).Control(16)=   "Label1(20)"
      Tab(8).Control(17)=   "Line1"
      Tab(8).Control(18)=   "Label5"
      Tab(8).Control(19)=   "lwTfnoCuotas"
      Tab(8).Control(20)=   "DataGrid3"
      Tab(8).Control(21)=   "txtauxTfno(0)"
      Tab(8).Control(22)=   "txtauxTfno(1)"
      Tab(8).Control(23)=   "txtauxTfno(2)"
      Tab(8).Control(24)=   "txtauxTfno(3)"
      Tab(8).Control(25)=   "FrameTelefonia(0)"
      Tab(8).Control(26)=   "txtauxTfno(4)"
      Tab(8).Control(27)=   "txtauxTfno(5)"
      Tab(8).Control(28)=   "Text5(4)"
      Tab(8).Control(29)=   "Text5(5)"
      Tab(8).Control(30)=   "txtauxTfno(6)"
      Tab(8).Control(31)=   "Text5(6)"
      Tab(8).Control(32)=   "txtauxTfno(7)"
      Tab(8).Control(33)=   "txtauxTfno(8)"
      Tab(8).Control(34)=   "txtauxTfno(9)"
      Tab(8).Control(35)=   "txtauxTfno(10)"
      Tab(8).Control(36)=   "FrameTelefonia(1)"
      Tab(8).Control(37)=   "cboOperadorTfnnia2(1)"
      Tab(8).Control(38)=   "cboOperadorTfnnia2(0)"
      Tab(8).Control(39)=   "FrameModuloVtaPlazos"
      Tab(8).Control(40)=   "FrameToolAux(5)"
      Tab(8).ControlCount=   41
      TabCaption(9)   =   "Fito"
      TabPicture(9)   =   "frmFacClientesGr.frx":0108
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label1(33)"
      Tab(9).Control(1)=   "Label1(35)"
      Tab(9).Control(2)=   "Label1(104)"
      Tab(9).Control(3)=   "imgFecha(6)"
      Tab(9).Control(4)=   "Label1(105)"
      Tab(9).Control(5)=   "Label1(107)"
      Tab(9).Control(6)=   "Label1(108)"
      Tab(9).Control(7)=   "Label1(109)"
      Tab(9).Control(8)=   "ImageFito(0)"
      Tab(9).Control(9)=   "ImageFito(1)"
      Tab(9).Control(10)=   "ImageFito(2)"
      Tab(9).Control(11)=   "ImageFito(3)"
      Tab(9).Control(12)=   "Label1(115)"
      Tab(9).Control(13)=   "ImageFito(4)"
      Tab(9).Control(14)=   "Label1(124)"
      Tab(9).Control(15)=   "DataGrid4"
      Tab(9).Control(16)=   "cboManipulador"
      Tab(9).Control(17)=   "Text1(57)"
      Tab(9).Control(18)=   "txtauxFito(3)"
      Tab(9).Control(19)=   "txtauxFito(2)"
      Tab(9).Control(20)=   "txtauxFito(1)"
      Tab(9).Control(21)=   "cboFitos(0)"
      Tab(9).Control(22)=   "txtauxFito(0)"
      Tab(9).Control(23)=   "txtauxFito(4)"
      Tab(9).Control(24)=   "Text1(58)"
      Tab(9).Control(25)=   "cmdFitos(0)"
      Tab(9).Control(26)=   "txtauxFito(5)"
      Tab(9).Control(27)=   "chkManiProv"
      Tab(9).Control(28)=   "cboFitos(1)"
      Tab(9).Control(29)=   "FrameToolAux(3)"
      Tab(9).ControlCount=   30
      TabCaption(10)  =   "Marja"
      TabPicture(10)  =   "frmFacClientesGr.frx":0124
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "imgFechaCampos(7)"
      Tab(10).Control(1)=   "Label1(110)"
      Tab(10).Control(2)=   "Label1(111)"
      Tab(10).Control(3)=   "imgFechaCampos(8)"
      Tab(10).Control(4)=   "Label1(112)"
      Tab(10).Control(5)=   "imgFechaCampos(9)"
      Tab(10).Control(6)=   "Label1(113)"
      Tab(10).Control(7)=   "Label4(1)"
      Tab(10).Control(8)=   "DataGrid5"
      Tab(10).Control(9)=   "txtauxMarja(0)"
      Tab(10).Control(10)=   "txtauxMarja(1)"
      Tab(10).Control(11)=   "txtauxMarja(2)"
      Tab(10).Control(12)=   "txtauxMarja(3)"
      Tab(10).Control(13)=   "txtauxMarja(4)"
      Tab(10).Control(14)=   "txtauxMarja(7)"
      Tab(10).Control(15)=   "txtauxMarja(5)"
      Tab(10).Control(16)=   "txtauxMarja(9)"
      Tab(10).Control(17)=   "txtauxMarja(8)"
      Tab(10).Control(18)=   "txtauxMarja(6)"
      Tab(10).Control(19)=   "cbomarjal"
      Tab(10).Control(20)=   "FrameToolAux(6)"
      Tab(10).ControlCount=   21
      TabCaption(11)  =   "Tab 11"
      TabPicture(11)  =   "frmFacClientesGr.frx":0140
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      Begin VB.Frame Frame3 
         Caption         =   "Codigos CDIR / FACE  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1695
         Left            =   -74640
         TabIndex        =   425
         Top             =   600
         Width           =   17175
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   68
            Left            =   10800
            MaxLength       =   16
            TabIndex        =   398
            Tag             =   "Codigo aseg.|T|S|||sclien|oficinacontable||N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   67
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   397
            Tag             =   "Prop.|T|S|||sclien|orgproponente||N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   66
            Left            =   10800
            MaxLength       =   16
            TabIndex        =   396
            Tag             =   "UT|T|S|||sclien|unidadtramitadora||N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   65
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   395
            Tag             =   "O|T|S|||sclien|organogestor||N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Oficina contable"
            Height          =   240
            Index           =   64
            Left            =   8400
            TabIndex        =   429
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "�rgano proponente"
            Height          =   240
            Index           =   50
            Left            =   360
            TabIndex        =   428
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Unidad tramitadora"
            Height          =   240
            Index           =   49
            Left            =   8400
            TabIndex        =   427
            Top             =   480
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "�rgano gestor"
            Height          =   240
            Index           =   32
            Left            =   360
            TabIndex        =   426
            Top             =   480
            Width           =   1950
         End
      End
      Begin VB.Frame FrameAsegurados 
         Caption         =   "Datos asegurados / riesgo   "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   5415
         Left            =   -74640
         TabIndex        =   394
         Top             =   2880
         Width           =   17175
         Begin VB.CommandButton cmdActRiesgo 
            Caption         =   "Actualizar riesgo"
            Height          =   495
            Left            =   12000
            TabIndex        =   424
            Top             =   4560
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtSit 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   423
            Text            =   "Text2"
            Top             =   4680
            Width           =   6045
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   51
            Left            =   14160
            MaxLength       =   10
            TabIndex        =   421
            Tag             =   "Fecha Reclamaci�n|F|S|||sclien|UtFecrecal|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   3360
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   49
            Left            =   8760
            MaxLength       =   16
            TabIndex        =   419
            Tag             =   "Riesgo|N|S|||sclien|riesgoact|#,###,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   3360
            Width           =   1570
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   53
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   417
            Tag             =   "Fecha concesion|F|S|||sclien|fecbajcre|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   3360
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   43
            Left            =   14160
            MaxLength       =   16
            TabIndex        =   415
            Tag             =   "L�mite cr�dito|N|S|0||sclien|limcredi|#,###,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2520
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   63
            Left            =   8760
            MaxLength       =   16
            TabIndex        =   413
            Tag             =   "Cr�dito concedidp|N|S|0||sclien|CreditoConcedido|#,###,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2520
            Width           =   1570
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   41
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   411
            Tag             =   "Fecha concesion|F|S|||sclien|fechaulr|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   64
            Left            =   14160
            MaxLength       =   3
            TabIndex        =   409
            Tag             =   "N�Grupo|T|S|||sclien|TipoCredito|||"
            Text            =   "Text1"
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   47
            Left            =   8760
            MaxLength       =   16
            TabIndex        =   407
            Tag             =   "L�mite cr�dito|N|S|0||sclien|credisol|#,###,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1680
            Width           =   1570
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   48
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   405
            Tag             =   "Fecha Reclamaci�n|F|S|||sclien|FechaSol|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   55
            Left            =   14160
            MaxLength       =   16
            TabIndex        =   403
            Tag             =   "N�Grupo|T|S|||sclien|NumGrupo|||"
            Text            =   "Text1"
            Top             =   840
            Width           =   1470
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   50
            Left            =   8760
            MaxLength       =   16
            TabIndex        =   401
            Tag             =   "Codigo aseg.|T|S|||sclien|codaseg||N|"
            Text            =   "Text1"
            Top             =   847
            Width           =   1570
         End
         Begin VB.ComboBox cboTipoASeg 
            Height          =   360
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   399
            Tag             =   "Tipo credito|N|S|||sclien|credipriv||N|"
            Top             =   847
            Width           =   2415
         End
         Begin VB.Label lblPerfil 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   14160
            TabIndex        =   434
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. fec. recalculo"
            Height          =   240
            Index           =   83
            Left            =   12120
            TabIndex        =   422
            Top             =   3390
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Riesgo actual"
            Height          =   240
            Index           =   81
            Left            =   6360
            TabIndex        =   420
            Top             =   3360
            Width           =   1320
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   2400
            Picture         =   "frmFacClientesGr.frx":015C
            ToolTipText     =   "Buscar fecha"
            Top             =   3390
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha baja"
            Height          =   240
            Index           =   92
            Left            =   600
            TabIndex        =   418
            Top             =   3390
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "L�mite Cr�dito"
            Height          =   240
            Index           =   45
            Left            =   12120
            TabIndex        =   416
            Top             =   2550
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Credito concedido"
            Height          =   240
            Index           =   118
            Left            =   6360
            TabIndex        =   414
            Top             =   2580
            Width           =   1785
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   2400
            Picture         =   "frmFacClientesGr.frx":01E7
            ToolTipText     =   "Buscar fecha"
            Top             =   2550
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha concesi�n"
            Height          =   240
            Index           =   66
            Left            =   600
            TabIndex        =   412
            Top             =   2550
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo cr�dito"
            Height          =   240
            Index           =   22
            Left            =   12120
            TabIndex        =   410
            Top             =   1710
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cr�dito solicitado"
            Height          =   240
            Index           =   79
            Left            =   6360
            TabIndex        =   408
            Top             =   1710
            Width           =   1710
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   2400
            Picture         =   "frmFacClientesGr.frx":0272
            ToolTipText     =   "Buscar fecha"
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha solicitud"
            Height          =   240
            Index           =   80
            Left            =   600
            TabIndex        =   406
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "N� Grupo:"
            Height          =   240
            Index           =   94
            Left            =   12120
            TabIndex        =   404
            Top             =   900
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "C�digo aseguradora"
            Height          =   240
            Index           =   82
            Left            =   6360
            TabIndex        =   402
            Top             =   900
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo cr�dito"
            Height          =   255
            Index           =   117
            Left            =   600
            TabIndex        =   400
            Top             =   900
            Width           =   1335
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   6
         Left            =   -74760
         TabIndex        =   389
         Top             =   480
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   6
            Left            =   120
            TabIndex        =   390
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   5
         Left            =   -74640
         TabIndex        =   386
         Top             =   360
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   5
            Left            =   120
            TabIndex        =   387
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   4
         Left            =   -74760
         TabIndex        =   384
         Top             =   480
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   385
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
      Begin VB.Frame FrameNavegaCRM 
         Height          =   735
         Left            =   -74640
         TabIndex        =   377
         Top             =   840
         Width           =   13455
         Begin VB.OptionButton optCRM 
            Caption         =   "Historial"
            Height          =   240
            Index           =   5
            Left            =   11880
            TabIndex        =   383
            Tag             =   "33"
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optCRM 
            Caption         =   "Reclamaciones"
            Height          =   240
            Index           =   4
            Left            =   9552
            TabIndex        =   382
            Tag             =   "32"
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optCRM 
            Caption         =   "Obs. departamento"
            Height          =   240
            Index           =   3
            Left            =   6720
            TabIndex        =   381
            Tag             =   "31"
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton optCRM 
            Caption         =   "Cobros"
            Height          =   240
            Index           =   2
            Left            =   5160
            TabIndex        =   380
            Tag             =   "13"
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optCRM 
            Caption         =   "Llamadas"
            Height          =   240
            Index           =   1
            Left            =   3360
            TabIndex        =   379
            Tag             =   "30"
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optCRM 
            Caption         =   "Acciones comerciales"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   378
            Tag             =   "3"
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.Frame FrameVisorDocumentos 
         BorderStyle     =   0  'None
         Caption         =   "Visor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   -63720
         TabIndex        =   224
         Top             =   960
         Width           =   5055
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   720
            Picture         =   "frmFacClientesGr.frx":02FD
            Style           =   1  'Graphical
            TabIndex        =   227
            ToolTipText     =   "Eliminar"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   1320
            Picture         =   "frmFacClientesGr.frx":0CFF
            Style           =   1  'Graphical
            TabIndex        =   226
            ToolTipText     =   "Ver Documento"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "frmFacClientesGr.frx":1289
            Style           =   1  'Graphical
            TabIndex        =   225
            ToolTipText     =   "Insertar Im�gen"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Image1 
            Height          =   6495
            Left            =   120
            Stretch         =   -1  'True
            Top             =   720
            Width           =   5535
         End
      End
      Begin VB.Frame FrameNavegaDoc 
         Enabled         =   0   'False
         Height          =   735
         Left            =   -74880
         TabIndex        =   368
         Top             =   840
         Width           =   11055
         Begin VB.OptionButton optDoc 
            Caption         =   "Puntos"
            Height          =   240
            Index           =   7
            Left            =   9840
            TabIndex        =   376
            Tag             =   "39"
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Documentos"
            Height          =   240
            Index           =   6
            Left            =   8085
            TabIndex        =   375
            Tag             =   "36"
            Top             =   360
            Width           =   1680
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Dtos."
            Height          =   240
            Index           =   5
            Left            =   7035
            TabIndex        =   374
            Tag             =   "12"
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Precios esp."
            Height          =   240
            Index           =   4
            Left            =   5385
            TabIndex        =   373
            Tag             =   "1"
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Facturas"
            Height          =   240
            Index           =   3
            Left            =   4095
            TabIndex        =   372
            Tag             =   "8"
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Albaranes"
            Height          =   240
            Index           =   2
            Left            =   2685
            TabIndex        =   371
            Tag             =   "7"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Pedidos"
            Height          =   240
            Index           =   1
            Left            =   1410
            TabIndex        =   370
            Tag             =   "6"
            Top             =   360
            Width           =   1200
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Ofertas"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   369
            Tag             =   "5"
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   3
         Left            =   -74760
         TabIndex        =   366
         Top             =   2880
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   367
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
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
      Begin VB.Frame frameDptoDirec 
         Caption         =   "Datos Relacionados con Dpto. Direcci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1860
         Left            =   120
         TabIndex        =   343
         Top             =   7200
         Width           =   8115
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   42
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   344
            Text            =   "Text2"
            Top             =   840
            Width           =   4485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   42
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "Cod. Situaci�n|N|N|0|99|sclien|codsitua|00|N|"
            Text            =   "Te"
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   40
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Fecha ult. movim.|F|S|||sclien|fechamov|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   44
            Left            =   4680
            MaxLength       =   5
            TabIndex        =   20
            Tag             =   "Distancia Km.|N|S|0|99999|sclien|kilometr||N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboTipocliente 
            Height          =   360
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "tipclien|N|N|||sclien|tipclien||N|"
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Distancia Km."
            Height          =   195
            Index           =   56
            Left            =   3600
            TabIndex        =   348
            Top             =   443
            Width           =   1080
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1635
            ToolTipText     =   "Buscar situaci�n"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Situaci�n"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   347
            Top             =   840
            Width           =   1455
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1635
            Picture         =   "frmFacClientesGr.frx":1C8B
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. movimiento"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   346
            Top             =   420
            Width           =   2130
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo cliente"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   345
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame frameDptoVentas 
         Caption         =   "Datos Relacionados con Dpto. Ventas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   4215
         Left            =   8520
         TabIndex        =   329
         Top             =   4800
         Width           =   9135
         Begin VB.ComboBox cboPrioridad 
            Height          =   360
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Tag             =   "Prioridad|N|N|||sclien|prioridad||N|"
            Top             =   3315
            Width           =   3615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   36
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   333
            Text            =   "Text2"
            Top             =   360
            Width           =   5205
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   36
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   41
            Tag             =   "Cod. Agente|N|N|0|9999|sclien|codagent|0000|N|"
            Text            =   "Text"
            Top             =   360
            Width           =   1000
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   360
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Tag             =   "Tipo Facturaci�n|N|N|||sclien|tipofact||N|"
            Top             =   2820
            Width           =   2295
         End
         Begin VB.ComboBox cboAlbaran 
            Height          =   360
            ItemData        =   "frmFacClientesGr.frx":1D16
            Left            =   2160
            List            =   "frmFacClientesGr.frx":1D18
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "Valorar albaran con|N|N|||sclien|albarcon||N|"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   39
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   50
            Tag             =   "Repeticiones Fact.|N|S|1|9|sclien|numrepet|#|N|"
            Text            =   "T"
            Top             =   2820
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   37
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   43
            Tag             =   "Cod. Tarifa|N|N|0|999|sclien|codtarif|000|N|"
            Text            =   "Tex"
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   37
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   332
            Text            =   "Text2"
            Top             =   1320
            Width           =   5205
         End
         Begin VB.CheckBox chkPromociones 
            Caption         =   "Aplicar Promociones"
            Height          =   315
            Left            =   2640
            TabIndex        =   54
            Tag             =   "Aplicar Promociones|N|N|||sclien|promocio||N|"
            Top             =   3780
            Width           =   2535
         End
         Begin VB.CheckBox chkReferencia 
            Caption         =   "Referencia Obligada"
            Height          =   315
            Left            =   240
            TabIndex        =   53
            Tag             =   "Referencia obligada|N|N|||sclien|referobl||N|"
            Top             =   3780
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   38
            Left            =   8520
            MaxLength       =   1
            TabIndex        =   51
            Tag             =   "Per�odo Facturaci�n|N|N|0|9|sclien|periodof|0|N|"
            Text            =   "T"
            Top             =   2820
            Width           =   390
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   52
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   331
            Text            =   "Text2"
            Top             =   1830
            Width           =   5205
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   52
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   44
            Tag             =   "Dir. envio habitual|N|S|0||sclien|coddirenhab|||"
            Text            =   "Tex"
            Top             =   1830
            Width           =   1000
         End
         Begin VB.CheckBox chkParticular 
            Caption         =   "Particular"
            Height          =   315
            Left            =   7560
            TabIndex        =   56
            Tag             =   "Particular|N|N|||sclien|particular||N|"
            Top             =   3780
            Width           =   1335
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
            Height          =   360
            Index           =   59
            Left            =   8160
            MaxLength       =   5
            TabIndex        =   48
            Tag             =   "Comision|N|S|0|99.90|sclien|Comision|#0.00||"
            Text            =   "Text1"
            Top             =   2280
            Width           =   765
         End
         Begin VB.CheckBox chkRecargFinan 
            Caption         =   "Recargo financiero"
            Height          =   315
            Left            =   5160
            TabIndex        =   55
            Tag             =   "Recargo financiero|N|N|||sclien|Recargofinanciero||N|"
            Top             =   3780
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   61
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   42
            Tag             =   "Visitador|N|N|0|9999|sclien|visitador|0000|N|"
            Text            =   "Text"
            Top             =   840
            Width           =   1000
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   61
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   330
            Text            =   "Text2"
            Top             =   840
            Width           =   5205
         End
         Begin VB.CheckBox chkPuntos 
            Caption         =   "Puntos "
            Height          =   315
            Left            =   4920
            TabIndex        =   46
            Tag             =   "Puntos|N|N|||sclien|TienePuntos||N|"
            Top             =   2280
            Width           =   1095
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
            Height          =   360
            Index           =   62
            Left            =   6000
            TabIndex        =   47
            Tag             =   "Puntos|N|S|||sclien|puntos|||"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prioridad"
            Height          =   240
            Index           =   126
            Left            =   240
            TabIndex        =   435
            Top             =   3360
            Width           =   840
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1800
            ToolTipText     =   "Buscar agente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
            Height          =   240
            Index           =   9
            Left            =   240
            TabIndex        =   342
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Facturaci�n"
            Height          =   240
            Index           =   4
            Left            =   240
            TabIndex        =   341
            Top             =   2880
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valorar Albaran "
            Height          =   240
            Index           =   18
            Left            =   240
            TabIndex        =   340
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rep.factura"
            Height          =   240
            Index           =   55
            Left            =   4560
            TabIndex        =   339
            Top             =   2880
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tarifa"
            Height          =   240
            Index           =   59
            Left            =   240
            TabIndex        =   338
            Top             =   1320
            Width           =   570
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1800
            ToolTipText     =   "Buscar tarifa"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Per�odo Facturaci�n"
            Height          =   240
            Index           =   46
            Left            =   6480
            TabIndex        =   337
            Top             =   2880
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dir. envio hab."
            Height          =   240
            Index           =   84
            Left            =   240
            TabIndex        =   336
            ToolTipText     =   "Direccion envio habitual"
            Top             =   1800
            Width           =   1440
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1800
            ToolTipText     =   "Buscar tarifa"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comision"
            Height          =   240
            Index           =   106
            Left            =   7200
            TabIndex        =   335
            Top             =   2340
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Visitador"
            Height          =   240
            Index           =   116
            Left            =   240
            TabIndex        =   334
            Top             =   840
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   22
            Left            =   1800
            ToolTipText     =   "Buscar agente"
            Top             =   840
            Width           =   240
         End
      End
      Begin VB.Frame frameDptoAdmon 
         Caption         =   "Datos Relacionados con Dpto. Administraci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   4215
         Left            =   8520
         TabIndex        =   314
         Top             =   480
         Width           =   9135
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   23
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   318
            Text            =   "Text2"
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   23
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   23
            Tag             =   "Cod. F. Pago|N|N|0|999|sclien|codforpa|000|N|"
            Text            =   "Tex"
            Top             =   360
            Width           =   765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   30
            Left            =   5520
            MaxLength       =   2
            TabIndex        =   317
            Tag             =   "D�a Pago 3|N|S|0|99|sclien|diapago3||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   29
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   316
            Tag             =   "D�a de Pago 2|N|S|0|99|sclien|diapago2||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   26
            Left            =   8160
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Mes a no girar|N|S|0|12|sclien|mesnogir||N|"
            Text            =   "Te"
            Top             =   360
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   31
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   31
            Tag             =   "C�digo Banco|N|S|0|9999|sclien|codbanco|0000|N|"
            Text            =   "Text"
            Top             =   1920
            Width           =   765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   32
            Left            =   3720
            MaxLength       =   4
            TabIndex        =   32
            Tag             =   "Sucursal|N|S|0|9999|sclien|codsucur|0000|N|"
            Text            =   "Text"
            Top             =   1920
            Width           =   765
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   33
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   57
            Tag             =   "D�gito Control|T|S|||sclien|digcontr|00||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   435
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   34
            Left            =   5040
            MaxLength       =   10
            TabIndex        =   33
            Tag             =   "Cuenta Bancaria|T|S|||sclien|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   35
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   34
            Tag             =   "Cta. Contable|T|N|||sclien|codmacta||N|"
            Text            =   "Text1"
            Top             =   2400
            Width           =   1725
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   35
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   315
            Text            =   "Text2"
            Top             =   2400
            Width           =   5205
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   28
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "D�a Pago 1|N|S|0|99|sclien|diapago1||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   27
            Left            =   8160
            MaxLength       =   2
            TabIndex        =   27
            Tag             =   "D�a Vto. Atrasado|N|S|0|31|sclien|diavtoat||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   615
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
            Height          =   360
            Index           =   24
            Left            =   2040
            MaxLength       =   5
            TabIndex        =   25
            Tag             =   "Dto. Pronto Pago|N|N|0|99.90|sclien|dtoppago|#0.00||"
            Text            =   "Text1"
            Top             =   840
            Width           =   765
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
            Height          =   360
            Index           =   25
            Left            =   2040
            MaxLength       =   5
            TabIndex        =   28
            Tag             =   "Dto. General|N|N|0|99.90|sclien|dtognral|#0.00||"
            Text            =   "Text1"
            Top             =   1380
            Width           =   765
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Cuenta. Ventas alternativa"
            Height          =   315
            Left            =   5880
            TabIndex        =   38
            Tag             =   "Cancela abonos|N|N|||sclien|cliabono||N|"
            Top             =   3360
            Width           =   3135
         End
         Begin VB.ComboBox cboTipoIVA 
            Height          =   360
            ItemData        =   "frmFacClientesGr.frx":1D1A
            Left            =   3960
            List            =   "frmFacClientesGr.frx":1D1C
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "Tipo de IVA|N|N|||sclien|tipoiva||N|"
            Top             =   1380
            Width           =   2415
         End
         Begin VB.CheckBox chkTasaReciclado 
            Caption         =   "Tas......"
            Height          =   315
            Left            =   3000
            TabIndex        =   37
            Tag             =   "tasareciclado|N|N|0|1|sclien|tasareciclado||N|"
            Top             =   3360
            Width           =   2535
         End
         Begin VB.CheckBox chkCorreo 
            Caption         =   "Se le envia correo"
            Height          =   315
            Left            =   240
            TabIndex        =   39
            Tag             =   "Referencia obligada|N|N|||sclien|enviocorreo||N|"
            Top             =   3720
            Width           =   2295
         End
         Begin VB.CheckBox chkPortesFac 
            Caption         =   "Portes al facturar"
            Height          =   315
            Left            =   240
            TabIndex        =   36
            Tag             =   "Portes al facturar|N|N|||sclien|AplicaPortesFactura||N|"
            Top             =   3360
            Width           =   2295
         End
         Begin VB.ComboBox cboFraRenting 
            Height          =   360
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Tag             =   "tipclien|N|S|||sclien|TipoFraRenting||N|"
            Top             =   2880
            Width           =   2655
         End
         Begin VB.CheckBox chkRentingDpto 
            Caption         =   "Por dpto."
            Height          =   315
            Left            =   4800
            TabIndex        =   35
            Tag             =   "Renting x departamento|N|N|||sclien|Rentin_x_dpto||N|"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   56
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   30
            Tag             =   "IBAN|T|S|||sclien|iban|||"
            Text            =   "Text"
            Top             =   1920
            Width           =   765
         End
         Begin VB.CheckBox chkEnvioFraEmail 
            Caption         =   "Envio factura por email"
            Height          =   315
            Left            =   3000
            TabIndex        =   40
            Tag             =   "Recargo financiero|N|N|||sclien|EnvFraEmail||N|"
            Top             =   3720
            Width           =   2775
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1800
            ToolTipText     =   "Buscar forma de pago"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de pago"
            Height          =   240
            Index           =   68
            Left            =   240
            TabIndex        =   328
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "D�as de Pago"
            Height          =   240
            Index           =   31
            Left            =   3000
            TabIndex        =   327
            Top             =   900
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   240
            Index           =   48
            Left            =   240
            TabIndex        =   326
            Top             =   1920
            Width           =   465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1800
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
            Height          =   240
            Index           =   8
            Left            =   6720
            TabIndex        =   325
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "D�a Vt. atrasado"
            Height          =   240
            Index           =   52
            Left            =   6480
            TabIndex        =   324
            Top             =   900
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Pronto Pago"
            Height          =   240
            Index           =   53
            Left            =   240
            TabIndex        =   323
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. General"
            Height          =   240
            Index           =   54
            Left            =   240
            TabIndex        =   322
            Top             =   1410
            Width           =   1230
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable"
            Height          =   240
            Index           =   51
            Left            =   240
            TabIndex        =   321
            Top             =   2400
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo IVA"
            Height          =   240
            Index           =   29
            Left            =   3000
            TabIndex        =   320
            Top             =   1410
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Fact. "
            Height          =   240
            Index           =   91
            Left            =   240
            TabIndex        =   319
            Top             =   2880
            Width           =   1485
         End
      End
      Begin VB.Frame frameComercial 
         Caption         =   "Comercial"
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
         Height          =   1935
         Left            =   -65880
         TabIndex        =   305
         Top             =   480
         Width           =   8295
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   21
            Left            =   2520
            MaxLength       =   60
            TabIndex        =   309
            Tag             =   "e-mail Comercial|T|S|||sclien|maiclie2||N|"
            Text            =   "Text1"
            Top             =   1440
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   20
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   308
            Tag             =   "Fax Comercial|T|S|||sclien|faxclie2||N|"
            Text            =   "Text1"
            Top             =   900
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   19
            Left            =   2520
            MaxLength       =   20
            TabIndex        =   307
            Tag             =   "Tel�fono Comercial|T|S|||sclien|telclie2||N|"
            Text            =   "Text1"
            Top             =   900
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   18
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   306
            Tag             =   "Contacto Comercial|T|S|||sclien|perclie2||N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   41
            Left            =   1080
            TabIndex        =   313
            Top             =   1500
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   240
            Index           =   42
            Left            =   5040
            TabIndex        =   312
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   240
            Index           =   43
            Left            =   1080
            TabIndex        =   311
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   240
            Index           =   44
            Left            =   1080
            TabIndex        =   310
            Top             =   360
            Width           =   915
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   2280
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   1500
            Width           =   240
         End
      End
      Begin VB.Frame frameAdmon 
         Caption         =   "Administraci�n"
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
         Height          =   1935
         Left            =   -74760
         TabIndex        =   296
         Top             =   480
         Width           =   8295
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   14
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   300
            Tag             =   "Contacto Admon.|T|S|||sclien|perclie1||N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   15
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   299
            Tag             =   "Tel�fono Admon.|T|S|||sclien|telclie1||N|"
            Text            =   "Text1"
            Top             =   900
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   16
            Left            =   5280
            MaxLength       =   15
            TabIndex        =   298
            Tag             =   "Fax Admon.|T|S|||sclien|faxclie1||N|"
            Text            =   "Text1"
            Top             =   900
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Index           =   17
            Left            =   2160
            MaxLength       =   60
            TabIndex        =   297
            Tag             =   "e-mail Admon.|T|S|||sclien|maiclie1||N|"
            Text            =   "maiclie1"
            Top             =   1440
            Width           =   5055
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   240
            Index           =   3
            Left            =   720
            TabIndex        =   304
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   240
            Index           =   38
            Left            =   720
            TabIndex        =   303
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   240
            Index           =   39
            Left            =   4680
            TabIndex        =   302
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   40
            Left            =   720
            TabIndex        =   301
            Top             =   1440
            Width           =   600
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   1800
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   1452
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   45
         Left            =   6600
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   11
         Tag             =   "Password cliente|T|N|||sclien|pasclien|||"
         Text            =   "3"
         Top             =   2520
         Width           =   1380
      End
      Begin VB.ComboBox cboPais 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   7
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "N.I.F.|T|N|||sclien|nifclien||N|"
         Text            =   "Text1"
         Top             =   540
         Width           =   1965
      End
      Begin VB.Frame FrameModuloVtaPlazos 
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
         Left            =   -64440
         TabIndex        =   284
         Top             =   6600
         Width           =   6900
         Begin VB.TextBox txtauxTfno 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   14
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   207
            Text            =   "1.2562"
            Top             =   1560
            Width           =   555
         End
         Begin VB.TextBox txtauxTfno 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   13
            Left            =   5160
            MaxLength       =   40
            TabIndex        =   209
            Text            =   "1.2562"
            Top             =   1560
            Width           =   915
         End
         Begin VB.TextBox txtauxTfno 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   12
            Left            =   3120
            MaxLength       =   40
            TabIndex        =   208
            Text            =   "1.2562"
            Top             =   1560
            Width           =   555
         End
         Begin VB.TextBox txtauxTfno 
            Height          =   360
            Index           =   11
            Left            =   240
            MaxLength       =   40
            TabIndex        =   206
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   11
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   285
            Text            =   "Text5"
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Venta a plazos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   125
            Left            =   120
            TabIndex        =   295
            Top             =   0
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Financiacion"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   123
            Left            =   240
            TabIndex        =   290
            Top             =   1200
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Meses"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   119
            Left            =   1800
            TabIndex        =   289
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Importe mes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   122
            Left            =   4920
            TabIndex        =   288
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Restantes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   120
            Left            =   2880
            TabIndex        =   286
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   23
            Left            =   1080
            Tag             =   "-1"
            ToolTipText     =   "Buscar actividad"
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Art�culo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   121
            Left            =   240
            TabIndex        =   287
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Frame FramePuntos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -63600
         TabIndex        =   279
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdAccDocs 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2160
            Picture         =   "frmFacClientesGr.frx":1D1E
            Style           =   1  'Graphical
            TabIndex        =   283
            ToolTipText     =   "Insertar puntos masivo"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            Picture         =   "frmFacClientesGr.frx":2720
            Style           =   1  'Graphical
            TabIndex        =   282
            ToolTipText     =   "Insertar puntos"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            Picture         =   "frmFacClientesGr.frx":3122
            Style           =   1  'Graphical
            TabIndex        =   281
            ToolTipText     =   "Imprimir puntos"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdAccDocs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   720
            Picture         =   "frmFacClientesGr.frx":36AC
            Style           =   1  'Graphical
            TabIndex        =   280
            ToolTipText     =   "Eliminar puntos"
            Top             =   240
            Width           =   495
         End
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
         Index           =   60
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   277
         Tag             =   "Pais|T|S|||sclien|codpais|||"
         Text            =   "Text1"
         Top             =   2520
         Width           =   165
      End
      Begin VB.ComboBox cbomarjal 
         Height          =   360
         Left            =   -63360
         TabIndex        =   267
         Tag             =   "-1"
         Text            =   "cbomarjal"
         Top             =   1560
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox txtauxMarja 
         Height          =   360
         Index           =   6
         Left            =   -63360
         MaxLength       =   30
         TabIndex        =   271
         Tag             =   "Partida|T|S||||partida|||"
         Text            =   "partida"
         Top             =   1560
         Width           =   3765
      End
      Begin VB.TextBox txtauxMarja 
         Height          =   360
         Index           =   8
         Left            =   -61200
         TabIndex        =   269
         Text            =   "nombre"
         Top             =   2400
         Width           =   1605
      End
      Begin VB.TextBox txtauxMarja 
         Height          =   4995
         Index           =   9
         Left            =   -63360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   270
         Text            =   "frmFacClientesGr.frx":40AE
         Top             =   3240
         Width           =   5565
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -69360
         MaxLength       =   40
         TabIndex        =   266
         Tag             =   "Sup.derechos|N|N||||dchos|#,##0.00||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtauxMarja 
         Height          =   360
         Index           =   7
         Left            =   -63360
         TabIndex        =   268
         Text            =   "nombre"
         Top             =   2400
         Width           =   1845
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -69840
         MaxLength       =   40
         TabIndex        =   265
         Tag             =   "Sup.SIGPAC|N|N||||poligno|#,##0.00||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -70800
         MaxLength       =   40
         TabIndex        =   264
         Tag             =   "Poligono|N|N|||||00000||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   263
         Tag             =   "Partida|N|N|||||00000||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   262
         Tag             =   "Poligono|N|N||||poligno|00000||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtauxMarja 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   261
         Tag             =   "id|N|N||||id|000||"
         Text            =   "nombre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cboFitos 
         Height          =   360
         Index           =   1
         ItemData        =   "frmFacClientesGr.frx":40B5
         Left            =   -67440
         List            =   "frmFacClientesGr.frx":40BF
         Style           =   2  'Dropdown List
         TabIndex        =   259
         Top             =   3480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkManiProv 
         Caption         =   "Provisional"
         Height          =   240
         Left            =   -67440
         TabIndex        =   240
         Tag             =   "Mani. provisional|N|N|||sclien|Manipuladorprovisional||N|"
         Top             =   1883
         Width           =   1815
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -68400
         MaxLength       =   40
         TabIndex        =   251
         Text            =   "Fecha"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdFitos 
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
         Height          =   375
         Index           =   0
         Left            =   -68880
         TabIndex        =   250
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   58
         Left            =   -69480
         MaxLength       =   10
         TabIndex        =   239
         Tag             =   "Fecha de caducidad|F|S|||sclien|ManipuladorFecCaducidad|dd/mm/yyyy||"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -65520
         MaxLength       =   40
         TabIndex        =   253
         Text            =   "nombre"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -72720
         MaxLength       =   40
         TabIndex        =   246
         Text            =   "nombre"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cboFitos 
         Height          =   360
         Index           =   0
         ItemData        =   "frmFacClientesGr.frx":40CB
         Left            =   -71520
         List            =   "frmFacClientesGr.frx":40D5
         Style           =   2  'Dropdown List
         TabIndex        =   248
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   40
         TabIndex        =   247
         Text            =   "nombre"
         Top             =   3480
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -70560
         MaxLength       =   40
         TabIndex        =   249
         Text            =   "nombre"
         Top             =   3360
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtauxFito 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -66960
         MaxLength       =   40
         TabIndex        =   252
         Text            =   "nombre"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   57
         Left            =   -71640
         TabIndex        =   238
         Tag             =   "Referencia|T|S|||sclien|ManipuladorNumCarnet|||"
         Text            =   "Te"
         Top             =   1800
         Width           =   2085
      End
      Begin VB.ComboBox cboManipulador 
         Height          =   360
         ItemData        =   "frmFacClientesGr.frx":40EE
         Left            =   -74760
         List            =   "frmFacClientesGr.frx":40F0
         Style           =   2  'Dropdown List
         TabIndex        =   237
         Tag             =   "Manipulador|N|N|||sclien|ManipuladortipoCarnet||N|"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox cboOperadorTfnnia2 
         Height          =   360
         Index           =   0
         ItemData        =   "frmFacClientesGr.frx":40F2
         Left            =   -73680
         List            =   "frmFacClientesGr.frx":40F4
         Style           =   2  'Dropdown List
         TabIndex        =   191
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboOperadorTfnnia2 
         Height          =   360
         Index           =   1
         ItemData        =   "frmFacClientesGr.frx":40F6
         Left            =   -59760
         List            =   "frmFacClientesGr.frx":40F8
         Style           =   2  'Dropdown List
         TabIndex        =   210
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Frame FrameTelefonia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -74760
         TabIndex        =   229
         Top             =   6600
         Width           =   3495
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   5
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   241
            ToolTipText     =   "Cambiar de titular"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   4
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   236
            ToolTipText     =   "CUOTA. Eliminar"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   235
            ToolTipText     =   "CUOTA. Modificar"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   2
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   234
            ToolTipText     =   "CUOTA. Nueva"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   231
            ToolTipText     =   "Imprimir Contrato"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdAccionesTfno 
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
            Index           =   0
            Left            =   120
            Picture         =   "frmFacClientesGr.frx":40FA
            Style           =   1  'Graphical
            TabIndex        =   230
            ToolTipText     =   "Renovar tel�fono"
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.TextBox txtauxTfno 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   -61200
         MaxLength       =   40
         TabIndex        =   200
         Text            =   "31/12/2018"
         Top             =   3240
         Width           =   1275
      End
      Begin VB.TextBox txtauxTfno 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   9
         Left            =   -64440
         MaxLength       =   40
         ScrollBars      =   1  'Horizontal
         TabIndex        =   197
         Text            =   "31/12/2018"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtauxTfno 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   8
         Left            =   -62160
         MaxLength       =   40
         TabIndex        =   199
         Text            =   "1.2562"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtauxTfno 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   -63000
         MaxLength       =   40
         TabIndex        =   198
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   6
         Left            =   -63120
         Locked          =   -1  'True
         TabIndex        =   220
         Text            =   "Text5"
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txtauxTfno 
         Height          =   360
         Index           =   6
         Left            =   -64440
         MaxLength       =   40
         TabIndex        =   196
         Top             =   2400
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   5
         Left            =   -63120
         Locked          =   -1  'True
         TabIndex        =   216
         Text            =   "Text5"
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   4
         Left            =   -63120
         Locked          =   -1  'True
         TabIndex        =   215
         Text            =   "Text5"
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox txtauxTfno 
         Height          =   360
         Index           =   5
         Left            =   -64440
         MaxLength       =   40
         TabIndex        =   195
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtauxTfno 
         Height          =   360
         Index           =   4
         Left            =   -64440
         MaxLength       =   40
         TabIndex        =   194
         Top             =   840
         Width           =   1245
      End
      Begin VB.Frame FrameTelefonia 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   -64560
         TabIndex        =   212
         Top             =   3720
         Width           =   6855
         Begin VB.CheckBox chkTelefonia 
            Caption         =   "Internet"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   203
            Top             =   120
            Width           =   1500
         End
         Begin VB.CheckBox chkTelefonia 
            Caption         =   "Inactivo"
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   204
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chkTelefonia 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   202
            Top             =   120
            Width           =   1275
         End
         Begin VB.CheckBox chkTelefonia 
            Caption         =   "Imprime factura"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   201
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.TextBox txtauxTfno 
         Height          =   1395
         Index           =   3
         Left            =   -64440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   205
         Text            =   "frmFacClientesGr.frx":4AFC
         Top             =   4680
         Width           =   6885
      End
      Begin VB.TextBox txtauxTfno 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   -70080
         MaxLength       =   40
         TabIndex        =   193
         Text            =   "nombre"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txtauxTfno 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   -72360
         MaxLength       =   40
         TabIndex        =   192
         Text            =   "nombre"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtauxTfno 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   -74400
         MaxLength       =   40
         TabIndex        =   190
         Text            =   "nombre"
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Height          =   1320
         Index           =   54
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "Observaciones facturacion|T|S|||sclien|obsfacturacion|||"
         Top             =   5760
         Width           =   3975
      End
      Begin VB.TextBox txtauxRent 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   10
         Left            =   -61800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   178
         Tag             =   "ID|T|S|||sclienrenting|obser|||"
         Text            =   "Ffin"
         Top             =   3120
         Width           =   4365
      End
      Begin VB.TextBox txtauxRent 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   9
         Left            =   -61200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   186
         Tag             =   "ID|T|N|||sclienrenting|nomtipco|||"
         Text            =   "Ffin"
         Top             =   1680
         Width           =   3765
      End
      Begin VB.TextBox txtauxRent 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   -61800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   177
         Tag             =   "ID|N|N|||sclienrenting|codtipco|0||"
         Text            =   "Ffin"
         Top             =   1680
         Width           =   525
      End
      Begin VB.CommandButton cmdRenting 
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
         Height          =   375
         Index           =   2
         Left            =   -71280
         TabIndex        =   184
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -65280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   176
         Tag             =   "Importe|N|N|||sclienrenting|importe|#,##0.00||"
         Text            =   "imp"
         Top             =   4440
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtauxRent 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   11
         Left            =   -61800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   179
         Tag             =   "Nombre|F|S||||ultfec|dd/mm/yyyy||"
         Text            =   "Ultima"
         Top             =   4320
         Width           =   1365
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -66600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   175
         Tag             =   "ID|F|N|||sclienrenting|fecbaja|dd/mm/yyyy||"
         Text            =   "Ffin"
         Top             =   4320
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -67680
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   174
         Tag             =   "Cutoas|N|N|||sclienrenting|numcuotas|0||"
         Text            =   "Cuotas"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -68760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   173
         Tag             =   "ID|F|N|||sclienrenting|fecalta|dd/mm/yyyy||"
         Text            =   "Alta"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   172
         Tag             =   "Ref|T|N|||sclienrenting|referencia|||"
         Text            =   "Referencia"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdRenting 
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
         Height          =   375
         Index           =   1
         Left            =   -67320
         TabIndex        =   183
         Top             =   4560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdRenting 
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
         Height          =   375
         Index           =   0
         Left            =   -69360
         TabIndex        =   182
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -71280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   181
         Text            =   "nomdpto"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   170
         Tag             =   "ID|N|N|||sclienrenting|ID|0||"
         Text            =   "id"
         Top             =   4320
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtauxRent 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73800
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   171
         Tag             =   "Dpto|N|S|||sclienrenting|coddirec|0||"
         Text            =   "dpto"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contactos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   6375
         Left            =   -74760
         TabIndex        =   157
         Top             =   2520
         Width           =   17175
         Begin VB.CommandButton cmdCargos 
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
            Height          =   375
            Left            =   5160
            TabIndex        =   392
            ToolTipText     =   "Editiar/Modificar cargos"
            Top             =   4200
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Index           =   2
            Left            =   120
            TabIndex        =   364
            Top             =   240
            Width           =   1605
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   365
               Top             =   180
               Width           =   1095
               _ExtentX        =   1931
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
         Begin VB.ComboBox cboCargo 
            Height          =   360
            Left            =   5520
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtauxDC 
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
            Index           =   8
            Left            =   15720
            MaxLength       =   30
            TabIndex        =   59
            Tag             =   "N|T|S|||scliendp|id|||"
            Text            =   "id Este esta fuera de vista "
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtauxDC 
            Height          =   360
            Index           =   3
            Left            =   12360
            MaxLength       =   12
            TabIndex        =   94
            Tag             =   "N|T|S|||scliendp|Telefono|||"
            Text            =   "Tfno"
            Top             =   1680
            Width           =   2085
         End
         Begin VB.TextBox txtauxDC 
            Height          =   360
            Index           =   4
            Left            =   15840
            MaxLength       =   5
            TabIndex        =   95
            Tag             =   "N|T|S|||scliendp|ext|||"
            Text            =   "extension"
            Top             =   1680
            Width           =   885
         End
         Begin VB.TextBox txtauxDC 
            Height          =   360
            Index           =   5
            Left            =   12360
            MaxLength       =   12
            TabIndex        =   96
            Tag             =   "N|T|S|||scliendp|movil|||"
            Text            =   "movil"
            Top             =   2355
            Width           =   2325
         End
         Begin VB.TextBox txtauxDC 
            Height          =   360
            Index           =   6
            Left            =   12360
            MaxLength       =   60
            TabIndex        =   97
            Tag             =   "N|T|S|||scliendp|maidirec|||"
            Text            =   "email"
            Top             =   3000
            Width           =   4485
         End
         Begin VB.TextBox txtauxDC 
            Height          =   2115
            Index           =   7
            Left            =   12360
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   98
            Tag             =   "N|T|S|||scliendp|observa|||"
            Text            =   "frmFacClientesGr.frx":4B03
            Top             =   3720
            Width           =   4485
         End
         Begin VB.TextBox txtauxDC 
            Height          =   315
            Index           =   1
            Left            =   12360
            MaxLength       =   30
            TabIndex        =   93
            Tag             =   "N|T|S|||scliendp|dpto|||"
            Text            =   "dpto"
            Top             =   960
            Width           =   4365
         End
         Begin VB.TextBox txtauxDC 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   91
            Tag             =   "Nombre|T|N|||scliendp|nombre|||"
            Text            =   "nombre"
            Top             =   4200
            Visible         =   0   'False
            Width           =   4005
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5175
            Left            =   120
            TabIndex        =   161
            Top             =   840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   9128
            _Version        =   393216
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
         Begin VB.TextBox txtauxDC 
            Height          =   315
            Index           =   2
            Left            =   6600
            MaxLength       =   40
            TabIndex        =   99
            Tag             =   "N|T|S|||scliendp|cargo|||"
            Text            =   "cargo"
            Top             =   5640
            Width           =   4005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   10320
            Tag             =   "-1"
            ToolTipText     =   "Buscar actividad"
            Top             =   840
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "el cbo oculta el text dc(2)"
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
            Index           =   0
            Left            =   9120
            TabIndex        =   165
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Extension"
            Height          =   255
            Index           =   78
            Left            =   14760
            TabIndex        =   164
            Top             =   1680
            Width           =   975
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   3
            Left            =   12000
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   3000
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones"
            Height          =   255
            Index           =   77
            Left            =   10800
            TabIndex        =   163
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Email"
            Height          =   255
            Index           =   67
            Left            =   10800
            TabIndex        =   162
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Departamento"
            Height          =   255
            Index           =   60
            Left            =   10800
            TabIndex        =   160
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   61
            Left            =   10800
            TabIndex        =   159
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Movil"
            Height          =   255
            Index           =   63
            Left            =   10800
            TabIndex        =   158
            Top             =   2400
            Width           =   855
         End
      End
      Begin VB.Frame FrameDireccionEnvio 
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
         Height          =   8115
         Left            =   -74880
         TabIndex        =   150
         Top             =   600
         Width           =   17415
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   10
            Left            =   11280
            MaxLength       =   30
            TabIndex        =   83
            Tag             =   "Cl|T|N|||sdirenvio|domdiren||N|"
            Text            =   "Text3"
            Top             =   960
            Width           =   5445
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Index           =   1
            Left            =   360
            TabIndex        =   356
            Top             =   120
            Width           =   1605
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   357
               Top             =   180
               Width           =   1095
               _ExtentX        =   1931
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
         Begin VB.TextBox txtZona 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   9
            Left            =   12240
            Locked          =   -1  'True
            TabIndex        =   169
            Text            =   "Text5"
            Top             =   3405
            Width           =   3735
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   9
            Left            =   11280
            MaxLength       =   6
            TabIndex        =   87
            Tag             =   "Zona|N|S|0||sdirenvio|codzona||N|"
            Text            =   "Text3"
            Top             =   3405
            Width           =   765
         End
         Begin VB.TextBox Text4 
            Height          =   2400
            Index           =   8
            Left            =   11280
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   90
            Tag             =   "Obs|T|S|||sdirenvio|observa||N|"
            Text            =   "frmFacClientesGr.frx":4B0B
            Top             =   5160
            Width           =   5805
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   0
            Left            =   360
            MaxLength       =   4
            TabIndex        =   80
            Tag             =   "C�digo|N|N|0|9999|sdirenvio|coddiren|0000|S|"
            Text            =   "Text3"
            Top             =   2520
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   2
            Left            =   4920
            MaxLength       =   30
            TabIndex        =   82
            Tag             =   "Domicilio|T|N|||sdirenvio|domdiren||N|"
            Text            =   "Text3"
            Top             =   2520
            Visible         =   0   'False
            Width           =   3270
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   3
            Left            =   11280
            MaxLength       =   30
            TabIndex        =   85
            Tag             =   "Poblaci�n|T|N|||sdirenvio|pobdiren||N|"
            Text            =   "Text3"
            Top             =   2160
            Width           =   5445
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   5
            Left            =   11280
            MaxLength       =   30
            TabIndex        =   86
            Tag             =   "Provincia|T|N|||sdirenvio|prodiren||N|"
            Text            =   "Text3"
            Top             =   2790
            Width           =   3405
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   6
            Left            =   11280
            MaxLength       =   10
            TabIndex        =   88
            Tag             =   "Tel�fono|T|S|||sdirenvio|teldiren||N|"
            Text            =   "Text3"
            Top             =   3960
            Width           =   2085
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   1
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   81
            Tag             =   "Nombre Direc|T|N|||sdirenvio|nomdiren||N|"
            Text            =   "Text3"
            Top             =   2520
            Visible         =   0   'False
            Width           =   3270
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   7
            Left            =   11280
            MaxLength       =   10
            TabIndex        =   89
            Tag             =   "Fax|T|S|||sdirenvio|faxdiren||N|"
            Text            =   "Text3"
            Top             =   4560
            Width           =   2085
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Index           =   4
            Left            =   11280
            MaxLength       =   6
            TabIndex        =   84
            Tag             =   "C.Postal|T|N|||sdirenvio|codpobla||N|"
            Text            =   "Text3"
            Top             =   1560
            Width           =   1005
         End
         Begin MSDataGridLib.DataGrid DataGrid7 
            Height          =   6615
            Left            =   360
            TabIndex        =   351
            Top             =   840
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   11668
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
         Begin VB.Label Label1 
            Caption         =   "Direcci�n"
            Height          =   240
            Index           =   21
            Left            =   9720
            TabIndex        =   363
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Label lblFramePp 
            Caption         =   "Direcciones de envio"
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
            Left            =   2160
            TabIndex        =   358
            Top             =   210
            Width           =   5265
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   11040
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   3465
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Zona"
            Height          =   240
            Index           =   87
            Left            =   9720
            TabIndex        =   167
            Top             =   3495
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones"
            Height          =   240
            Index           =   58
            Left            =   9720
            TabIndex        =   156
            Top             =   5160
            Width           =   1440
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   10920
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "C.Postal"
            Height          =   240
            Index           =   73
            Left            =   9720
            TabIndex        =   155
            Top             =   1650
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   240
            Index           =   72
            Left            =   9720
            TabIndex        =   154
            Top             =   2160
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   240
            Index           =   71
            Left            =   9720
            TabIndex        =   153
            Top             =   2790
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   240
            Index           =   70
            Left            =   9720
            TabIndex        =   152
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   240
            Index           =   65
            Left            =   9720
            TabIndex        =   151
            Top             =   4680
            Width           =   345
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   46
         Left            =   -62160
         TabIndex        =   146
         Text            =   "Text4"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdAccCRM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -60360
         Picture         =   "frmFacClientesGr.frx":4B11
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Acciones CRM"
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAccCRM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -58440
         Picture         =   "frmFacClientesGr.frx":5513
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Impresion CRM"
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAccCRM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -59760
         Picture         =   "frmFacClientesGr.frx":5A9D
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Eliminar"
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame FrameDirecciones 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   8235
         Left            =   -74760
         TabIndex        =   129
         Top             =   480
         Width           =   17175
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Index           =   0
            Left            =   240
            TabIndex        =   353
            Top             =   240
            Width           =   1605
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   354
               Top             =   150
               Width           =   1095
               _ExtentX        =   1931
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
         Begin VB.TextBox txtZona 
            BackColor       =   &H80000018&
            Height          =   360
            Index           =   14
            Left            =   11880
            Locked          =   -1  'True
            TabIndex        =   168
            Text            =   "Text5"
            Top             =   5880
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   14
            Left            =   11160
            MaxLength       =   6
            TabIndex        =   70
            Tag             =   "Zona|N|S|0||sdirec|codzona||N|"
            Text            =   "Text3"
            Top             =   5880
            Width           =   645
         End
         Begin VB.Frame FrameCtaBanDpto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1800
            Left            =   9360
            TabIndex        =   138
            Top             =   6240
            Width           =   7695
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   19
               Left            =   5520
               MaxLength       =   10
               TabIndex        =   79
               Tag             =   "Cuenta Bancaria|T|S|||sdirec|oficinacontable|||"
               Top             =   1320
               Width           =   1965
            End
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   18
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   78
               Tag             =   "Prop.|T|S|||sdirec|orgproponente|||"
               Top             =   1320
               Width           =   1965
            End
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   17
               Left            =   5520
               MaxLength       =   10
               TabIndex        =   77
               Tag             =   "Ud|T|S|||sdirec|unidadtramitadora|||"
               Top             =   780
               Width           =   1965
            End
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   16
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   76
               Tag             =   "o|T|S|||sdirec|organogestor|||"
               Top             =   780
               Width           =   1965
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   360
               Index           =   15
               Left            =   1800
               MaxLength       =   4
               TabIndex        =   71
               Tag             =   "IBAN|T|S|||sdirec|iban|||"
               Text            =   "Text"
               Top             =   240
               Width           =   765
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   360
               Index           =   10
               Left            =   2640
               MaxLength       =   4
               TabIndex        =   72
               Tag             =   "C�digo Banco|N|S|0|9999|sdirec|codbanco|0000|N|"
               Text            =   "Text"
               Top             =   240
               Width           =   765
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   360
               Index           =   11
               Left            =   3480
               MaxLength       =   4
               TabIndex        =   73
               Tag             =   "Sucursal|N|S|0|9999|sdirec|codsucur|0000|N|"
               Text            =   "Text"
               Top             =   240
               Width           =   765
            End
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   12
               Left            =   4320
               MaxLength       =   2
               TabIndex        =   74
               Tag             =   "D�gito Control|T|S|||sdirec|digcontr|00||"
               Text            =   "Text1"
               Top             =   240
               Width           =   405
            End
            Begin VB.TextBox Text3 
               Height          =   360
               Index           =   13
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   75
               Tag             =   "Cuenta Bancaria|T|S|||sdirec|cuentaba|0000000000||"
               Top             =   240
               Width           =   2685
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ofi. contable"
               Height          =   240
               Index           =   76
               Left            =   3840
               TabIndex        =   433
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label1 
               Caption         =   "�rg. proponente"
               Height          =   240
               Index           =   75
               Left            =   120
               TabIndex        =   432
               Top             =   1350
               Width           =   1650
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ud. tramitadora"
               Height          =   240
               Index           =   74
               Left            =   3840
               TabIndex        =   431
               Top             =   780
               Width           =   1545
            End
            Begin VB.Label Label1 
               Caption         =   "�rgano gestor"
               Height          =   240
               Index           =   69
               Left            =   120
               TabIndex        =   430
               Top             =   780
               Width           =   1410
            End
            Begin VB.Label Label1 
               Caption         =   "IBAN"
               Height          =   255
               Index           =   47
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   3
            Left            =   11160
            MaxLength       =   6
            TabIndex        =   64
            Tag             =   "C.Postal|T|N|||sdirec|codpobla||N|"
            Text            =   "Text3"
            Top             =   2190
            Width           =   1005
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   8
            Left            =   11160
            MaxLength       =   10
            TabIndex        =   68
            Tag             =   "Fax|T|S|||sdirec|faxdirec||N|"
            Text            =   "Text3"
            Top             =   4650
            Width           =   2565
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   1
            Left            =   3000
            MaxLength       =   30
            TabIndex        =   61
            Tag             =   "Nombre Direc./Dpto|T|N|||sdirec|nomdirec||N|"
            Text            =   "Text3"
            Top             =   3840
            Visible         =   0   'False
            Width           =   3270
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   9
            Left            =   11160
            MaxLength       =   40
            TabIndex        =   69
            Tag             =   "e-mail|T|S|||sdirec|maidirec||N|"
            Text            =   "Text3"
            Top             =   5265
            Width           =   5655
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   6
            Left            =   11160
            MaxLength       =   30
            TabIndex        =   62
            Tag             =   "Persona Contacto|T|S|||sdirec|perdirec||N|"
            Text            =   "Text3"
            Top             =   960
            Width           =   5055
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   7
            Left            =   11160
            MaxLength       =   10
            TabIndex        =   67
            Tag             =   "Tel�fono|T|S|||sdirec|teldirec||N|"
            Text            =   "Text3"
            Top             =   4035
            Width           =   2565
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   5
            Left            =   11160
            MaxLength       =   30
            TabIndex        =   66
            Tag             =   "Provincia|T|N|||sdirec|prodirec||N|"
            Text            =   "Text3"
            Top             =   3420
            Width           =   3885
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   4
            Left            =   11160
            MaxLength       =   30
            TabIndex        =   65
            Tag             =   "Poblaci�n|T|N|||sdirec|pobdirec||N|"
            Text            =   "Text3"
            Top             =   2805
            Width           =   3885
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   2
            Left            =   11160
            MaxLength       =   100
            TabIndex        =   63
            Tag             =   "Domicilio|T|N|||sdirec|domdirec||N|"
            Text            =   "Text3"
            Top             =   1575
            Width           =   5775
         End
         Begin VB.TextBox Text3 
            Height          =   360
            Index           =   0
            Left            =   960
            MaxLength       =   3
            TabIndex        =   60
            Tag             =   "C�digo Direc./Dpto|N|N|0|999|sdirec|coddirec|000|S|"
            Text            =   "Text3"
            Top             =   3960
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSDataGridLib.DataGrid DataGrid6 
            Height          =   6975
            Left            =   240
            TabIndex        =   350
            Top             =   960
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   12303
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
         Begin VB.Label lblFramePp 
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
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   0
            Left            =   2040
            TabIndex        =   355
            Top             =   360
            Width           =   6825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   10920
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   5940
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Zona"
            Height          =   255
            Index           =   86
            Left            =   9480
            TabIndex        =   166
            Top             =   6000
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "0 es la direcci�n de envio de facturaci�n"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   57
            Left            =   360
            TabIndex        =   140
            Top             =   7920
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   30
            Left            =   9480
            TabIndex        =   137
            Top             =   4731
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   10
            Left            =   9480
            TabIndex        =   136
            Top             =   5362
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Pers. Contacto"
            Height          =   255
            Index           =   27
            Left            =   9480
            TabIndex        =   135
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   28
            Left            =   9480
            TabIndex        =   134
            Top             =   4100
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   26
            Left            =   9480
            TabIndex        =   133
            Top             =   3469
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   25
            Left            =   9480
            TabIndex        =   132
            Top             =   2838
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo postal"
            Height          =   240
            Index           =   24
            Left            =   9480
            TabIndex        =   131
            Top             =   2222
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   23
            Left            =   9480
            TabIndex        =   130
            Top             =   1591
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   10920
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   2220
            Width           =   240
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   2
            Left            =   10920
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   5340
            Width           =   240
         End
      End
      Begin VB.CheckBox chkClienteV 
         Caption         =   "Cliente Varios"
         Height          =   240
         Left            =   6240
         TabIndex        =   5
         Tag             =   "Cliente Varios|N|N|||sclien|clivario||N|"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   13
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha de Alta|F|N|||sclien|fechaalt|dd/mm/yyyy|N|"
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   10
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   118
         Text            =   "Text2"
         Top             =   3960
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   11
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   117
         Text            =   "Text2"
         Top             =   4440
         Width           =   5325
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   14
         Tag             =   "Cod. Env�o|N|S|0|99999|sclien|codenvio|000|N|"
         Text            =   "Tex"
         Top             =   3960
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   12
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   116
         Text            =   "Text2"
         Top             =   4920
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   9
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   114
         Text            =   "Text2"
         Top             =   3480
         Width           =   5325
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   9
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "Cod.Actividad|N|N|0|999|sclien|codactiv|000|N|"
         Text            =   "Tex"
         Top             =   3480
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   12
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   16
         Tag             =   "Cod. Ruta|N|S|0|999|sclien|codrutas|000|N|"
         Text            =   "Tex"
         Top             =   4920
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   11
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "Cod. Zona|N|S|0|999|sclien|codzonas|000|N|"
         Text            =   "Tex"
         Top             =   4440
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Height          =   1320
         Index           =   22
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Tag             =   "Observaciones|T|S|||sclien|observac|||"
         Top             =   5760
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   8
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Tag             =   "Web|T|S|||sclien|wwwclien||N|"
         Text            =   "Text1"
         Top             =   3000
         Width           =   6285
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   6
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Provincia|T|N|||sclien|proclien||N|"
         Text            =   "Text1"
         Top             =   2040
         Width           =   6285
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   5
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Poblaci�n|T|N|||sclien|pobclien||N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   3780
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   4
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "C.Postal|T|N|||sclien|codpobla||N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   3
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   6
         Tag             =   "Domicilio|T|N|||sclien|domclien||N|"
         Text            =   "Text1"
         Top             =   1050
         Width           =   6285
      End
      Begin MSComctlLib.ListView lwCRM 
         Height          =   6615
         Left            =   -74640
         TabIndex        =   145
         Top             =   1800
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   11668
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
      Begin MSComctlLib.ListView lw1 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   149
         Top             =   1680
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12515
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   6975
         Left            =   -74760
         TabIndex        =   180
         Top             =   1320
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12303
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   5055
         Left            =   -74640
         TabIndex        =   211
         Top             =   1080
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8916
         _Version        =   393216
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
      Begin MSComctlLib.ListView lwTfnoCuotas 
         Height          =   1815
         Left            =   -70920
         TabIndex        =   232
         Top             =   6720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3201
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   7135
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2373
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   244
         Top             =   3600
         Width           =   17295
         _ExtentX        =   30506
         _ExtentY        =   8916
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   6975
         Left            =   -74760
         TabIndex        =   260
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   12303
         _Version        =   393216
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
      Begin VB.CommandButton cmdCatalogo 
         Height          =   495
         Left            =   -60000
         Picture         =   "frmFacClientesGr.frx":649F
         Style           =   1  'Graphical
         TabIndex        =   393
         ToolTipText     =   "Catalogos"
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   -60360
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Huertos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   1
         Left            =   -72960
         TabIndex        =   391
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Telefon�a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   -72840
         TabIndex        =   388
         Top             =   480
         Width           =   1845
      End
      Begin VB.Image imgCrm 
         Height          =   375
         Left            =   -74640
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgDocumentos 
         Height          =   375
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Renting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   -72960
         TabIndex        =   349
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   124
         Left            =   -74760
         TabIndex        =   294
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Password web"
         Height          =   240
         Index           =   19
         Left            =   4920
         TabIndex        =   128
         Top             =   2580
         Width           =   1410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   -64800
         X2              =   -64800
         Y1              =   720
         Y2              =   8640
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   240
         Index           =   36
         Left            =   3360
         TabIndex        =   110
         Top             =   600
         Width           =   555
      End
      Begin VB.Image ImageFito 
         Height          =   255
         Index           =   4
         Left            =   -58440
         Top             =   1853
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Imprimir listado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   115
         Left            =   -59520
         TabIndex        =   278
         Top             =   1823
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Partida"
         Height          =   195
         Index           =   113
         Left            =   -63360
         TabIndex        =   275
         Top             =   1320
         Width           =   975
      End
      Begin VB.Image imgFechaCampos 
         Height          =   240
         Index           =   9
         Left            =   -61800
         ToolTipText     =   "Buscar fecha"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   112
         Left            =   -63360
         TabIndex        =   274
         Top             =   3000
         Width           =   1425
      End
      Begin VB.Image imgFechaCampos 
         Height          =   240
         Index           =   8
         Left            =   -59880
         Picture         =   "frmFacClientesGr.frx":CCF1
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha baja"
         Height          =   195
         Index           =   111
         Left            =   -61200
         TabIndex        =   273
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha alta"
         Height          =   195
         Index           =   110
         Left            =   -63360
         TabIndex        =   272
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Image imgFechaCampos 
         Height          =   240
         Index           =   7
         Left            =   -61920
         Picture         =   "frmFacClientesGr.frx":CD7C
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image ImageFito 
         Height          =   255
         Index           =   3
         Left            =   -72840
         ToolTipText     =   "Carnet.  Insertar / Ver"
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image ImageFito 
         Height          =   255
         Index           =   2
         Left            =   -73200
         ToolTipText     =   "DNI.  Insertar / Ver"
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image ImageFito 
         Height          =   255
         Index           =   1
         Left            =   -60840
         Top             =   1860
         Width           =   255
      End
      Begin VB.Image ImageFito 
         Height          =   255
         Index           =   0
         Left            =   -62640
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Carnet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   109
         Left            =   -61800
         TabIndex        =   258
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "D.N.I."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   108
         Left            =   -63240
         TabIndex        =   257
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   107
         Left            =   -65040
         TabIndex        =   256
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec. caducidad"
         Height          =   240
         Index           =   105
         Left            =   -69480
         TabIndex        =   255
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   -67920
         Picture         =   "frmFacClientesGr.frx":CE07
         ToolTipText     =   "Buscar fecha"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Autorizados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   104
         Left            =   -74760
         TabIndex        =   254
         Top             =   2520
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia"
         Height          =   240
         Index           =   35
         Left            =   -71640
         TabIndex        =   245
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Carnet manipulador"
         Height          =   240
         Index           =   33
         Left            =   -74760
         TabIndex        =   243
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Procedencia"
         Height          =   195
         Index           =   20
         Left            =   -59640
         TabIndex        =   242
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cuotas propias"
         Height          =   240
         Index           =   103
         Left            =   -70920
         TabIndex        =   233
         Top             =   6360
         Width           =   1470
      End
      Begin VB.Image imgFechaTf 
         Height          =   240
         Index           =   10
         Left            =   -60240
         Picture         =   "frmFacClientesGr.frx":CE92
         ToolTipText     =   "Buscar fecha"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFechaTf 
         Height          =   240
         Index           =   9
         Left            =   -63480
         Picture         =   "frmFacClientesGr.frx":CF1D
         ToolTipText     =   "Buscar fecha"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   -62880
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. alta"
         Height          =   240
         Index           =   102
         Left            =   -64440
         TabIndex        =   223
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Puntos"
         Height          =   240
         Index           =   101
         Left            =   -62040
         TabIndex        =   222
         Top             =   3000
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Min�"
         Height          =   240
         Index           =   100
         Left            =   -62880
         TabIndex        =   221
         Top             =   3000
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   -63600
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   -62640
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   -62160
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Asociado ppal"
         Height          =   240
         Index           =   97
         Left            =   -64440
         TabIndex        =   218
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Direccion facturaci�n"
         Height          =   240
         Index           =   96
         Left            =   -64440
         TabIndex        =   217
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   "Los chk tienen que estar ocultos al ins/modif cliente"
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
         Index           =   1
         Left            =   -62400
         TabIndex        =   213
         Top             =   4320
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   6000
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Obs. facturacion"
         Height          =   240
         Index           =   93
         Left            =   4320
         TabIndex        =   189
         Top             =   5520
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Ult. factura"
         Height          =   255
         Index           =   90
         Left            =   -61800
         TabIndex        =   188
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   89
         Left            =   -61800
         TabIndex        =   187
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label LabelDoc 
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
         ForeColor       =   &H00000080&
         Height          =   540
         Left            =   -74280
         TabIndex        =   148
         Top             =   480
         Width           =   7065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -62640
         Picture         =   "frmFacClientesGr.frx":CFA8
         ToolTipText     =   "Buscar fecha"
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -63480
         TabIndex        =   147
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label LabelCRM 
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
         Left            =   -74040
         TabIndex        =   144
         Top             =   480
         Width           =   5745
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1680
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   5520
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmFacClientesGr.frx":D033
         ToolTipText     =   "Buscar fecha"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alta"
         Height          =   240
         Index           =   16
         Left            =   120
         TabIndex        =   127
         Top             =   600
         Width           =   1065
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   1440
         Picture         =   "frmFacClientesGr.frx":D0BE
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1440
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1440
         ToolTipText     =   "Buscar zona"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Envio"
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   120
         Top             =   4020
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Ruta"
         Height          =   240
         Index           =   17
         Left            =   120
         TabIndex        =   119
         Top             =   5010
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1440
         ToolTipText     =   "Buscar ruta"
         Top             =   5010
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Actividad"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   115
         Top             =   3510
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   113
         Top             =   4530
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   11
         Left            =   120
         TabIndex        =   112
         Top             =   5520
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Web"
         Height          =   240
         Index           =   37
         Left            =   120
         TabIndex        =   111
         Top             =   3030
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   240
         Index           =   15
         Left            =   120
         TabIndex        =   109
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n"
         Height          =   255
         Index           =   34
         Left            =   3120
         TabIndex        =   108
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "C. Postal"
         Height          =   240
         Index           =   14
         Left            =   120
         TabIndex        =   107
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   240
         Index           =   13
         Left            =   120
         TabIndex        =   106
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Modelo"
         Height          =   240
         Index           =   98
         Left            =   -64440
         TabIndex        =   219
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "F. Renov."
         Height          =   240
         Index           =   99
         Left            =   -61200
         TabIndex        =   228
         Top             =   3000
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   95
         Left            =   -64440
         TabIndex        =   214
         Top             =   4320
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Pais"
         Height          =   255
         Index           =   114
         Left            =   120
         TabIndex        =   276
         Top             =   2573
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "TIPO contrato"
         Height          =   240
         Index           =   88
         Left            =   -61800
         TabIndex        =   185
         Top             =   1320
         Width           =   1410
      End
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
      Height          =   705
      Left            =   120
      TabIndex        =   123
      Top             =   720
      Width           =   17895
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   2
         Left            =   11760
         MaxLength       =   60
         TabIndex        =   2
         Tag             =   "Nombre Comercial|T|N|||sclien|nomcomer||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   6045
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   1
         Left            =   3120
         MaxLength       =   60
         TabIndex        =   1
         Tag             =   "Nombre Cliente|T|N|||sclien|nomclien||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   6645
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   0
         Left            =   840
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "C�digo Cliente|N|N|0|999999|sclien|codclien|000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   60
         Width           =   17895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom.Comercial"
         Height          =   240
         Index           =   12
         Left            =   10200
         TabIndex        =   126
         Top             =   300
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   240
         Index           =   1
         Left            =   2280
         TabIndex        =   125
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C�digo"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   124
         Top             =   300
         Width           =   660
      End
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
      Height          =   475
      Index           =   1
      Left            =   5400
      TabIndex        =   121
      Top             =   10800
      Width           =   7815
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   122
         Top             =   180
         Width           =   6915
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   16440
      TabIndex        =   101
      Top             =   10800
      Visible         =   0   'False
      Width           =   1155
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
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   103
      Top             =   10800
      Width           =   4935
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   104
         Top             =   180
         Width           =   4275
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   16440
      TabIndex        =   102
      Top             =   10800
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   15000
      TabIndex        =   100
      Top             =   10800
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5880
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   3960
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   11400
      Top             =   5040
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   10560
      Top             =   5760
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   810
      Left            =   9000
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1429
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   1890
      Left            =   11760
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3334
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
   Begin MSAdodcLib.Adodc Adodc1IMG 
      Height          =   495
      Left            =   9600
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin MSAdodcLib.Adodc data7 
      Height          =   1890
      Left            =   9240
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3334
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   1890
      Left            =   10920
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3334
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   17520
      TabIndex        =   359
      Top             =   120
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFacClientesGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Nuevo para WHOSE
'Quiero ver el cliente en cuestion
Public VerCliente As Long
 

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
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAc As frmFacAgentesCom
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio2 As frmFacCliEnvDpto
Attribute frmDptoEnvio2.VB_VarHelpID = -1
Private WithEvents frmMtoTipCo As frmManTiposContrato
Attribute frmMtoTipCo.VB_VarHelpID = -1
Private WithEvents frmModeloTel As frmTelefoniaModelos
Attribute frmModeloTel.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2
Attribute FrmArt.VB_VarHelpID = -1


'Para los documentos
Private frmAlb As frmFacEntAlbaranes2
Private frmAlbG  As frmFacEntAlbaranesGR
Private frmAlbS As frmFacEntAlbSAIL
Private frmOfe As frmFacEntOfertas2
Private frmOfeS As frmFacEntOferSAIL
Private frmPed As frmFacEntPedidos
Private frmPedS As frmFacEntPedSail



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
'   9.-  Telefonia
'   10.- Fitosan
'   11.- Campos
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Private ModoFrame2 As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar     5: BUSCAR(Enero2014) para direnvio

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
    
Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal



Private CambiaCCC_Ariadna As Boolean 'Por si tiene que actualizar en resto aplicaciones ariadna

'NUEVO: JULIO 2007. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String
Private PriVezForm As Boolean
Private ModoFrame  As Byte



Private Sub cboAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cboCargo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cboOperadorTfnnia_KeyPress(index As Integer, KeyAscii As Integer)
    
End Sub



Private Sub cboFitos_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboManipulador_Click()
    If Modo = 4 Then
        'Modificando socio
        If Text1(57).Text = "" Then
            Text1(57).Text = Text1(7).Text
            PonerFoco Text1(57)
        End If
    End If
End Sub

Private Sub cbomarjal_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboOperadorTfnnia2_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 1 Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            PonerFocoBtn cmdAceptar
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub cboPais_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboPrioridad_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboTipoASeg_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipocliente_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub cboTipoIVA_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkAbonos_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkAbonos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkClienteV_Click()
If Modo = 1 Then CheckCadenaBusqueda chkClienteV, BuscaChekc
End Sub

Private Sub chkClienteV_GotFocus()
   ConseguirfocoChk Modo
End Sub

Private Sub chkClienteV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCorreo_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkCorreo, BuscaChekc
End Sub

Private Sub chkCorreo_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkCorreo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkEnvioFraEmail_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkEnvioFraEmail, BuscaChekc
End Sub

Private Sub chkEnvioFraEmail_GotFocus()
    
    ConseguirfocoChk Modo
End Sub

Private Sub chkEnvioFraEmail_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkManiProv_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkManiProv, BuscaChekc
End Sub

Private Sub chkManiProv_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkManiProv_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkParticular_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkParticular, BuscaChekc
End Sub

Private Sub chkParticular_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkParticular_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPortesFac_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkPortesFac, BuscaChekc
End Sub

Private Sub chkPortesFac_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPortesFac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPromociones_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkPromociones, BuscaChekc
End Sub

Private Sub chkPromociones_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPromociones_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPuntos_Click()
   If Modo = 1 Then CheckCadenaBusqueda chkPuntos, BuscaChekc
   
End Sub

Private Sub chkPuntos_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPuntos_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkRecargFinan_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkRecargFinan, BuscaChekc
End Sub

Private Sub chkRecargFinan_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkRecargFinan_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkReferencia_Click()
    
    'Buscqueda
    If Modo = 1 Then CheckCadenaBusqueda chkReferencia, BuscaChekc
    
End Sub

Private Sub chkReferencia_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkReferencia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkRentingDpto_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkRentingDpto, BuscaChekc
End Sub

Private Sub chkRentingDpto_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkRentingDpto_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkTasaReciclado_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkTasaReciclado, BuscaChekc
End Sub

Private Sub chkTasaReciclado_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkTasaReciclado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkTelefonia_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccCRM_Click(index As Integer)
    
    'Acciones parar el CRM
    Select Case index
    Case 1
        If Modo <> 2 Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
        If Text1(0).Text = "" Then Exit Sub
        
        
        frmCRMImprimir.Text1 = Text1(0).Text
        frmCRMImprimir.Text2 = Text1(1).Text
        frmCRMImprimir.Show vbModal
        
    Case 0
    
        Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
        Case 0
            'NUEVA, modificar o insertar acciones comerciales
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
        Case 1
            'NUEVA llamda EFECTUADA
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 1  'Llamada efectuada
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
            
        'Case 2
        '    'Emails
        '    LanzarProgramaEmails
        '    If MsgBox("Refrescar datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Case 2
            'NO puede insertar nada.
            Exit Sub
        Case 3
            frmCrmObsDpto.Nuevo = True
            frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
            frmCrmObsDpto.Label2.Font.SIZE = IIf(Len(Data1.Recordset!NomClien) > 30, 11, 13)
            frmCrmObsDpto.Tag = Data1.Recordset!codClien
            frmCrmObsDpto.Show vbModal
            
        Case 4
           
        
            BuscaChekc = ""
            If Text1(35).Text = "" Then
                BuscaChekc = "No tiene cta contable"
            Else
                If Text2(35).Text = "" Then BuscaChekc = "Cta contable incorrecta"
            End If
            If BuscaChekc < "" Then
                MsgBox BuscaChekc, vbExclamation
                Exit Sub
            End If
            
            
            
            BuscaChekc = "-1|" & Text1(1).Text & "|" & Text1(35).Text & "|" & Text2(35).Text & "|"
            frmCRMReclamas.Intercambio = BuscaChekc  'nueva
            frmCRMReclamas.Show vbModal
            BuscaChekc = ""
        Case 5
            'NUEVA entrada en Historial
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 2  'Historial
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
        End Select
        Me.Refresh
        DoEvents
        CargaDatosLWCRM
        Screen.MousePointer = vbDefault
    Case 2
    
        If CByte(RecuperaValor(lwCRM.Tag, 1)) = 3 Then
            If lwCRM.SelectedItem Is Nothing Then Exit Sub
            If MsgBox("�Desea eliminar las observaciones del departamento " & Me.lwCRM.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            BuscaChekc = "DELETE from scrmobsclien  WHERE codclien = " & Me.Data1.Recordset!codClien & " AND dpto=" & lwCRM.SelectedItem.SubItems(3)
            If ejecutar(BuscaChekc, False) Then CargaDatosLWCRM
            BuscaChekc = ""
        ElseIf CByte(RecuperaValor(lwCRM.Tag, 1)) = 6 Then
        
        End If
    End Select
End Sub

Private Sub cmdAccDocs_Click(index As Integer)
Dim SQL As String
    If index <> 2 And index <> 4 Then
        If Modo <> 2 Then Exit Sub
    End If
    Select Case index
        Case 0
            
            LanzaAnyadirImagenDocumento 0
            
            
        Case 1, 2
            
            
            
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
            
            
            If index = 2 Then
                ImprimirImagen
            Else
                EliminarImagen
            End If
            
            
        'PUNTOS
        Case 3, 4, 5, 6
        
            If index = 4 Then
                ImprimirHcoPuntos
                Exit Sub
            End If
            
            If Text1(0).Text = "" Then Exit Sub
            
            If index = 3 Then
                'Sin definir
            
            Else
                If vUsu.Nivel > 0 Then
                    MsgBox "No tiene permiso", vbExclamation
                    Exit Sub
                End If
            
                CadenaDesdeOtroForm = ""
                If index = 5 Then
                    'Nuevo
                    
                    frmListado5.OtrosDatos = Text1(0).Text & "|" & Text1(1).Text & "|"
                    frmListado5.OpcionListado = 19
                    frmListado5.Show vbModal
                    
            
                Else
                    'QUitar
                    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
                    
                    If Me.lw1.SelectedItem.Tag = 0 Then
                        MsgBox "No son incrementos manuales de puntos", vbExclamation
                        Exit Sub
                    End If
                    
                    If MsgBox("Seguro que desea eliminar los puntos?", vbQuestion + vbYesNoCancel) = vbYes Then
                        If DesHacerIncrementoPuntosCliente Then CadenaDesdeOtroForm = "OK"
                    End If
                    
                End If
                If CadenaDesdeOtroForm <> "" Then
                    
                        PosicionarData
                        PonerCampos
                    
                End If
            End If

            
    End Select
End Sub

Private Sub cmdAccionesTfno_Click(index As Integer)
Dim Seguir As Boolean

    If Me.data6.Recordset.EOF Then Exit Sub

    Seguir = False
   ' If index < 2 Or index > 4 Then
        If Modo = 2 Or Modo = 9 Then Seguir = True
   ' Else
   '     If Modo = 9 And ModificaLineas = 0 Then Seguir = True
   ' End If
    
    If Not Seguir Then Exit Sub
    Select Case index
    Case 0, 5
        Renovar_Cambiar_Telefono index = 0
        
    Case 1
    
        BuscaChekc = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "73")
        
        If BuscaChekc <> "" Then

            CadenaDesdeOtroForm = ""
            frmListado5.OpcionListado = 18  'pedir importe y `precio terminal
            frmListado5.Show vbModal
            If CadenaDesdeOtroForm = "" Then Exit Sub
                
            'Lanzar rpt de documento
            With frmImprimir
                .FormulaSeleccion = "({sclientfno.IdTelefono}=""" & data6.Recordset!idtelefono & """) "
                .OtrosParametros = "|Duracion=""" & RecuperaValor(CadenaDesdeOtroForm, 2) & """|"
                
                CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 1)
                If CadenaDesdeOtroForm = "" Then
                    CadenaDesdeOtroForm = "           "
                Else
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " �"
                End If
                CadenaDesdeOtroForm = "PrecioTerminal=""" & CadenaDesdeOtroForm & " ""|"
                .OtrosParametros = .OtrosParametros & CadenaDesdeOtroForm
                
                .NumeroParametros = 2
        
                .SoloImprimir = False
                .EnvioEMail = False
                .Titulo = "Contrato telefono"
                .Opcion = 3000   'VAN TODOS EN ESTE SACO
                .NombrePDF = ""
                .NombrePDF = BuscaChekc
                .NombreRPT = BuscaChekc
                .ConSubInforme = True
                .MostrarTreeDesdeFuera = False
                .Show vbModal
            End With
            BuscaChekc = ""
            CadenaDesdeOtroForm = ""
       Else
            MsgBox "Falta personalizar. Llame a Ariadna", vbExclamation
            
       End If
    Case 2, 3
        'Insertar modificar cuota propia de telefonia
        
        If index = 2 Then
            'NUEVO
            kCampo = Me.lwTfnoCuotas.ListItems.Count
            If kCampo > 0 Then
                kCampo = CInt(Val(Mid(Me.lwTfnoCuotas.ListItems(kCampo).Key, 2)))
            End If
            BuscaChekc = "||"
            kCampo = kCampo + 1
        Else
            If lwTfnoCuotas.SelectedItem Is Nothing Then Exit Sub
            kCampo = Mid(lwTfnoCuotas.SelectedItem.Key, 2)
            BuscaChekc = lwTfnoCuotas.SelectedItem.Text & "|" & lwTfnoCuotas.SelectedItem.SubItems(1) & "|"
        End If
        CadenaDesdeOtroForm = data6.Recordset!idtelefono & "|" & kCampo & "|" & BuscaChekc
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & data6.Recordset!Operador & "|"
        frmVarios3.Opcion = 2
        frmVarios3.Show vbModal
        
        
        CargaCuotasTelefonia kCampo
        
        CadenaDesdeOtroForm = ""
    Case 4
        'Eliminar la cuota de telefonia
         If lwTfnoCuotas.SelectedItem Is Nothing Then Exit Sub
         
         BuscaChekc = "Va a eliminar la cuota: " & lwTfnoCuotas.SelectedItem.Text & " (" & lwTfnoCuotas.SelectedItem.SubItems(1) & ")"
         BuscaChekc = BuscaChekc & vbCrLf & "Tel�fono: " & data6.Recordset!idtelefono & vbCrLf
         BuscaChekc = BuscaChekc & vbCrLf & "�Continuar?"
         If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then
            BuscaChekc = "DELETE from  sclientfnoCuotas WHERE IdTelefono = '" & data6.Recordset!idtelefono & "'"
            BuscaChekc = BuscaChekc & " AND numlinea = " & Mid(Me.lwTfnoCuotas.SelectedItem.Key, 2)
            conn.Execute BuscaChekc
            CargaCuotasTelefonia 0
         End If
    End Select

    BuscaChekc = ""
    
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim b As Boolean
Dim EraNuevaLinea As Boolean
Dim NombreModificado As Boolean

     Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me, 1) Then
                 'Si pone en la cuenta contable, crear nueva cta contable
                 If Text2(35).Text = vbCrearNuevaCta Then
                    If Not InsertarCuentaCble(Text1(35).Text, Text1(0).Text) Then
                        MsgBox "Se ha producido un error insertando la cuenta: " & Text1(1).Text & ". Compruebelo", vbExclamation
                        Exit Sub
                    Else
                        Text2(35).Text = Text1(1).Text
                    End If
                End If
                 ActualizarAsegurados_
                 PosicionarData
                 
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
            
                NombreModificado = False
                If DBLet(Data1.Recordset!Clivario, "N") = 0 Then
                    'EL NOMBRE DEL cliente HA CAMBIADO. Los de varios NO los contemplamos
                    If Trim(DevNombreSQL(Data1.Recordset!NomClien)) <> Trim(Text1(1).Text) Then NombreModificado = True
                End If
                
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear 'Adelante transacciones....
                    
                    'Si ha cambiado la situacion de bloqueo
                    If Val(Data1.Recordset!codsitua) <> Val(Text1(42).Text) Then
                        'SI. Grabamos en LOG
                        Set LOG = New cLOG
                        cad = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", CStr(Val(Data1.Recordset!codsitua)))
                        cad = "Anterior: " & Val(Data1.Recordset!codsitua) & " - " & cad
                        cad = "Actual: " & Text1(42).Text & " - " & Text2(12).Text & vbCrLf & cad
                        LOG.Insertar 31, vUsu, cad
                        Set LOG = Nothing
                    End If
                    
                    'Actualizadmos en contabilidad    'Hay datos contables. Actualizamos?
                    If HayQueActualizarenContabilidad Then
                        ModificarCtaContabilidad True, Text1(35).Text, Val(Text1(0).Text)
                        ActualizarAsegurados_
                        
                        If CambiaCCC_Ariadna Then
                            cad = "codclien = " & Me.Text1(0).Text
                            If ComprobarDatosProcesoCCC(cad, lblIndicador, True) Then
                                frmVarios3.Opcion = 1
                                frmVarios3.Show vbModal
                            End If
                        End If
                    End If

                    If NombreModificado Then UpdatearNomClien

                    PosicionarData
                End If
            End If
                
         Case 5, 6, 7, 8, 9, 10, 11 'InsertarModificar linea
            'Enero 2014
            'Puede buscar dentro de un cliente por direccion de envio
            If Modo = 6 And ModificaLineas = 5 Then
                'OK. Esta buscando por direccion de envio
                'Buscaremos y si retorna haremos un truco.
                
            End If
          
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            If InsertarModificarLinea Then
                Select Case Modo
                Case 5
                    cad = "coddirec = " & Text3(0).Text
                Case 6
                    cad = "coddiren = " & Text4(0).Text
                Case 7
                    cad = "id = " & txtauxDC(8).Text
                Case 8
                    cad = "id = " & Me.txtauxRent(0).Text
                Case 9
                    cad = "IdTelefono = '" & Me.txtauxTfno(0).Text & "'"
                Case 10
                    cad = "id = " & Me.txtauxFito(4).Text
                Case 11
                    cad = "id = " & Me.txtauxMarja(0).Text
                End Select
                
                
                If Modo = 5 Then
                
                    LLamaLineasDirec 0, 0
                    DataGrid6.AllowAddNew = False
                    CargaLineas True, 5
                
                    If ModificaLineas = 1 Then
                        Data2.Recordset.MoveLast
                    Else
                        Data2.Recordset.Find cad
                    End If
                    b = True

                
                ElseIf Modo = 6 Then
                    
                    LLamaLineasDirenEvio 0, 0
                    DataGrid7.AllowAddNew = False
                    CargaLineas True, 6
                
                    If ModificaLineas = 1 Then
                        Data3.Recordset.MoveLast
                    Else
                        Data3.Recordset.Find cad
                    End If
                    b = True
                    
                    
                    
                ElseIf Modo = 7 Then
                
                        
                    LLamaLineasDatosContacto 0, 0
                    DataGrid1.AllowAddNew = False
                    CargaLineas True, 0
                
                    If ModificaLineas = 1 Then
                        Data4.Recordset.MoveLast
                    Else
                        Data4.Recordset.Find cad
                    End If
                    b = True
                ElseIf Modo = 8 Then
                    '8.- Rentings
                    
                    EraNuevaLinea = ModificaLineas = 1
                    LLamaLineasRenting 0, 0
                    DataGrid2.AllowAddNew = False
                    CargaLineas True, 1
                
                    If ModificaLineas = 1 Then
                        Data5.Recordset.MoveLast
                    Else
                        Data5.Recordset.Find cad
                    End If
                    b = True
                ElseIf Modo = 9 Then
                    '9.- Telefonia
                    
                    LLamaLineasTfnia 0, 0
                    DataGrid3.AllowAddNew = False
                    CargaLineas True, 2
                
                    If ModificaLineas = 1 Then
                        data6.Recordset.MoveLast
                    Else
                        data6.Recordset.Find cad
                    End If
                    b = True
                ElseIf Modo = 10 Then
                    '10.- Fitos
                    LLamaLineasFito 0, 0
                    DataGrid4.AllowAddNew = False
                    CargaLineas True, 3
                
                    If ModificaLineas = 1 Then
                        Data7.Recordset.MoveLast
                    Else
                        Data7.Recordset.Find cad
                    End If
                    b = True
                ElseIf Modo = 11 Then
                    '11.- Campos huertos
                    LLamaLineasCamposHuertos 0, 0
                    DataGrid5.AllowAddNew = False
                    CargaLineas True, 4
                
                    If ModificaLineas = 1 Then
                        data8.Recordset.MoveLast
                    Else
                        data8.Recordset.Find cad
                    End If
                    b = True
                    
                    
                    
                End If
                If b Then
                    If Modo = 5 Then
                        PonerDatosForaGridDepartamentos False
                    ElseIf Modo = 6 Then
                        PonerDatosForaGridDirEnvio False
                    ElseIf Modo = 7 Then
                        PonerDatosForaGridContacto False
                        
                    ElseIf Modo = 9 Then
                        'Telefonia
                        PonerDatosForaGridTfno False
                    ElseIf Modo = 10 Then
                    
                    
                    ElseIf Modo = 11 Then
                        'datos
                        PonerDatosForaGridCamposHuertos False
                        
                    Else
                        PonerDatosForaGridRent False
                        
                        'Pregunta para generar la factura
                        If EraNuevaLinea Then
                        
                            'Deberiamos comprobar si la proxima fecha de facturacion para este cliente es
                            'anterior a la fecha de alta
                            BuscaChekc = DevuelveDesdeBD(conAri, "max(ultfec)", "sclienrenting", "codclien", CStr(Data1.Recordset!codClien))
                            If BuscaChekc <> "" Then
                                If Data5.Recordset!fecalta > CDate(BuscaChekc) Then
                                    'No muesto el msg. Ya lo he hecho en datosoklinea
                                    'MsgBox "Pendiente facturacion proximo periodo", vbInformation
                                Else
                                    BuscaChekc = ""
                                End If
                            End If
                            If BuscaChekc = "" Then
                                frmListado3.Opcion = 22
                                frmListado3.OtrosDatos = "sclienrenting.codclien = " & Text1(0).Text & " AND " & cad
                                frmListado3.Show vbModal
                            End If
                            BuscaChekc = ""
                        End If
                    End If
                    ModificaLineas = 0
                    
                    
                    
                    
'                    lblIndicador.Caption = Indicador
                    'PonerModoFrame 0, Modo
                    
                    
                    
                    
                End If
                
                'PonerBotonCabecera True
                'PonerFocoBtn Me.cmdRegresar
                 PonerModo 2
                 Indicador_
            End If
      
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Indicador_()
    On Error Resume Next
    Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then
        Err.Clear
        lblIndicador.Caption = ""
    End If
End Sub

Private Sub cmdActRiesgo_Click()
    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    
    If DBLet(Data1.Recordset!Clivario, "N") = 1 Then
        'No recalculamos a clivarios
        MsgBox "Cliente de varios", vbExclamation
        Exit Sub
    End If
    
    
    If Me.cboTipoASeg.ListIndex < 0 Then
        MsgBox "Tipo credito incorrecto", vbExclamation
        Exit Sub
    End If
    
    If cboTipoASeg.ItemData(cboTipoASeg.ListIndex) = 9 Then
        MsgBox "Sin asegurar", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("�Calcular el riesgo del cliente " & Text1(1).Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set miRsAux = Nothing
    
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "Calculando riesgo"
    Me.lblIndicador.Refresh
    Riesgo
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Riesgo()
Dim ImpAlb As Currency, ImpTesor As Currency
Dim miSQL As String


    RiesgoCliente CLng(Text1(0).Text), Me.cboTipoIVA.ItemData(cboTipoIVA.ListIndex), Now, ImpTesor, ImpAlb, miRsAux, 90
    ImpTesor = ImpTesor + ImpAlb
    miSQL = "UPDATE sclien SET UtFecrecal = " & DBSet(Now, "F")
    miSQL = miSQL & ", riesgoact = " & DBSet(ImpTesor, "N")
        
    ImpAlb = ImporteFormateado(Text1(43).Text)
    
    If ImpTesor <= ImpAlb Then
    
        'NO supera el limite
        If CInt(Text1(42).Text) > 0 Then
            'Estaba bloqueado por riesgo. Le quito la marca
            If CInt(Text1(42).Text) = vParamAplic.SituacionBloqueoOpAseg Then miSQL = miSQL & " ,codsitua = 0"
        End If
    Else
        'SUPERA EL RIESGO
        If CInt(Text1(42).Text) = 0 Then miSQL = miSQL & " ,codsitua = " & vParamAplic.SituacionBloqueoOpAseg
        
    End If
    miSQL = miSQL & " WHERE codclien = " & Text1(0).Text
    conn.Execute miSQL
    Espera 0.5
    PosicionarData
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then PonerCampos
    End If
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
            PonerModoFrame 0, Modo
            DataGrid6.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                If Not Data2.Recordset.EOF Then
                    Data2.Recordset.MoveFirst
                    PonerDatosForaGridDepartamentos False
                Else
                    PonerDatosForaGridDepartamentos True
                End If
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(coddirec=" & Text3(0).Text & ")"
                 CargaLineas True, 5
                 Data2.Recordset.Find cad
                 
                 
            End If
            PonerDatosForaGridDepartamentos False
            ModificaLineas = 0
            LLamaLineasDirec 0, 0
            PonerModo 2 'reestablecemos el modo
        Case 6
            'Modificar direcciones de envio
            PonerModoFrame 0, Modo
            DataGrid7.AllowAddNew = False
            If ModificaLineas = 1 Or ModificaLineas = 5 Then '1 = Insertar
                
                If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                    
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(coddiren=" & Text4(0).Text & ")"
                 CargaLineas True, 6
                 Data3.Recordset.Find cad
                 
            End If
            
            PonerDatosForaGridDirEnvio False
            ModificaLineas = 0
            LLamaLineasDirenEvio 0, 0
            PonerModo 2 'reestablecemos el modo
        Case 7
           'Modificar persona contacto
            PonerModoFrame 0, Modo
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                
                If Not Data4.Recordset.EOF Then Data4.Recordset.MoveFirst
                
                    
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(id=" & Me.txtauxDC(8).Text & ")"
                 CargaLineas True, 0
                 Data4.Recordset.Find cad
                 
                 
            End If
            PonerDatosForaGridContacto False
            ModificaLineas = 0
            LLamaLineasDatosContacto 0, 0
            
            PonerModo 2
       Case 8
           'Modificar renting
            PonerModoFrame 0, Modo
            DataGrid2.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                If Not Data5.Recordset.EOF Then Data5.Recordset.MoveFirst
                
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(id=" & CStr(Data5.Recordset!Id) & ")"
                 CargaLineas True, 1
                 Data5.Recordset.Find cad
                 
            End If
            PonerDatosForaGridRent False
            LLamaLineasRenting 0, 0
            ModificaLineas = 0
            PonerModo 2
    Case 9
            PonerModoFrame 0, Modo
            DataGrid3.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                
                If Not data6.Recordset.EOF Then data6.Recordset.MoveFirst
                
                    
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(IdTelefono='" & Me.txtauxTfno(0).Text & "')"
                 CargaLineas True, 2
                 data6.Recordset.Find cad
                 
                 
            End If
            PonerDatosForaGridTfno False
            LLamaLineasTfnia 0, 0
            ModificaLineas = 0
            PonerModo 2
    Case 10
            PonerModoFrame 0, Modo
            DataGrid4.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                If Not Data7.Recordset.EOF Then Data7.Recordset.MoveFirst
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(id='" & Me.txtauxFito(4).Text & "')"
                 CargaLineas True, 3
                 Data7.Recordset.Find cad
            End If
            LLamaLineasFito 0, 0
            ModificaLineas = 0
            PonerDatosForaGridTfno False
            PonerModo 2
            
    Case 11
            PonerModoFrame 0, Modo
            DataGrid5.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                If Not data8.Recordset.EOF Then data8.Recordset.MoveFirst
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(id='" & Me.txtauxMarja(0).Text & "')"
                 CargaLineas True, 4
                 data8.Recordset.Find cad
            End If
            PonerDatosForaGridCamposHuertos False
            LLamaLineasCamposHuertos 0, 0
            ModificaLineas = 0

            PonerModo 2
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vac�a los TextBox
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    MostrarSituacion False
    
    If VerCliente >= 0 Then
         NumRegElim = -1
        If vParamAplic.NumeroInstalacion = vbFenollar Then SugiereHuecoClienteNormal
            
        If NumRegElim = -1 Then NumRegElim = SugerirCodigoSiguienteStr("sclien", "codclien")
        Text1(0).Text = NumRegElim
    Else
        'Sugerir hueco
        NuevoClienteDesdePotencial
    End If
    
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        Text1(35).Text = "43." & NumRegElim
        Text1(35).Text = RellenaCodigoCuenta(Text1(35).Text)
         Text2(35).Text = vbCrearNuevaCta
    End If
    
    FormateaCampo Text1(0)
    Text1(13).Text = Format(Now, "dd/mm/yyyy")
    'Sugerir el tipo de IVA como NORMAL
    Me.cboTipoIVA.ListIndex = 0
    'Sugerir valorar albaran con: TODO
    Me.cboAlbaran.ListIndex = 0
    'Sugerir tipo facturacion a: Factura colectiva
    Me.cboFacturacion.ListIndex = 0
    'Sugerir tipo cliente
    Me.cboTipocliente.ListIndex = 0
    
    
    If vParamAplic.NumeroInstalacion = vbFontenas Then
        cboPrioridad.ListIndex = 3
    Else
        cboPrioridad.ListIndex = 0
    End If
    
    
    'Fitos
    If vParamAplic.ManipuladorFitosanitarios2 Then cboManipulador.ListIndex = 0
    If vParamAplic.ContabilidadNueva Then cboPais.ListIndex = 0 'Espa�a
    If vParamAplic.OperacionesAseguradas Then cboTipoASeg.ListIndex = Me.cboTipoASeg.ListCount - 1
    Me.chkCorreo.Value = 1
    'Sugerimos periodo y repeticion , a 1
    Text1(38).Text = 1
    Text1(39).Text = 1
    
    'A cero los descuentos
    Text1(24).Text = "0,00"
    Text1(25).Text = "0,00"
    
    'Valores por defecto desde parametros
    If vParamAplic.PorDefecto_Activ > 0 Then
        If Text1(9).Text = "" Then Text1(9).Text = vParamAplic.PorDefecto_Activ
        Text1_LostFocus 9
    End If
    If vParamAplic.PorDefecto_Envio > 0 Then
        If Text1(10).Text = "" Then Text1(10).Text = vParamAplic.PorDefecto_Envio
        Text1_LostFocus 10
    End If
    If vParamAplic.PorDefecto_Zona > 0 Then
        If Text1(11).Text = "" Then Text1(11).Text = vParamAplic.PorDefecto_Zona
        Text1_LostFocus 11
    End If
    If vParamAplic.PorDefecto_Ruta > 0 Then
        If Text1(12).Text = "" Then Text1(12).Text = vParamAplic.PorDefecto_Ruta
        Text1_LostFocus 12
    End If
    If vParamAplic.PorDefecto_Situ >= 0 Then
        Text1(42).Text = vParamAplic.PorDefecto_Situ
        Text1_LostFocus 42
    End If
    If vParamAplic.PorDefecto_Tarifa > 0 Then
        Text1(37).Text = vParamAplic.PorDefecto_Tarifa
        Text1_LostFocus 37
    End If
    If vParamAplic.PorDefecto_Agente > 0 Then
        Text1(36).Text = vParamAplic.PorDefecto_Agente
        Text1_LostFocus 36
        Text1(61).Text = Text1(36).Text
        Text2(61).Text = Text2(36).Text
    End If
    

    
    
    Me.SSTab1.Tab = 0
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        PonerFoco Text1(1)
    Else
        PonerFoco Text1(0)
        ConseguirFoco Text1(0), Modo
    End If
    
End Sub


Private Sub SugiereHuecoClienteNormal()
Dim Id As Long
Dim campo As String

    Set miRsAux = New ADODB.Recordset

    campo = "select codclien from sclien   WHERE codClien > 1900"
    miRsAux.Open campo, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    Id = miRsAux!codClien
     While Not miRsAux.EOF
            
            If miRsAux!codClien = Id Then
              Id = Id + 1
            
         Else
                'No hacemos nada
                NumRegElim = Id
                miRsAux.MoveLast
         End If
         miRsAux.MoveNext
        Wend
  
    miRsAux.Close
    Set miRsAux = Nothing

End Sub


Private Sub NuevoClienteDesdePotencial()
Dim campo As String
On Error GoTo eBuscarHuecoCliente
    Set miRsAux = New ADODB.Recordset

    campo = "select codclien,@rownum:=@rownum+1 AS rownum from sclien, (SELECT @rownum:=0) r  WHERE codClien > 0"
    miRsAux.Open campo, conn, adOpenKeyset, adLockReadOnly, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
        
        If (miRsAux!codClien - miRsAux!rownum) > 0 Then
            NumRegElim = miRsAux!codClien - 1
            'Este es el codigo
            miRsAux.MoveLast
        Else
            'No hacemos nada
            NumRegElim = miRsAux!codClien + 1
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Text1(0).Text = NumRegElim
    
    
    campo = "select * from sclipot where codclien= " & CadenaDesdeOtroForm
    miRsAux.Open campo, conn, adOpenKeyset, adLockReadOnly, adCmdText
        
    
    
    'codactiv, CodEnvio, codzonas, codrutas, perclie1, perclie1, telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac
    Text1(1).Text = DBLet(miRsAux!nomcomer, "T")
    Text1(2).Text = DBLet(miRsAux!nomcomer, "T")
    Text1(3).Text = DBLet(miRsAux!domclien, "T")
    Text1(4).Text = DBLet(miRsAux!codpobla, "T")
    Text1(5).Text = DBLet(miRsAux!pobclien, "T")
    Text1(6).Text = DBLet(miRsAux!proclien, "T")
    Text1(7).Text = DBLet(miRsAux!nifClien, "T") 'pasw
    Text1(45).Text = DBLet(miRsAux!nifClien, "T") 'pasw
    Text1(8).Text = DBLet(miRsAux!wwwclien, "T")
    ''codactiv, CodEnvio, codzonas, codrutas,
    Text1(9).Text = DBLet(miRsAux!codactiv, "T")
    Text1(10).Text = DBLet(miRsAux!CodEnvio, "T")
    Text1(11).Text = DBLet(miRsAux!codzonas, "T")
    Text1(12).Text = DBLet(miRsAux!codrutas, "T")
    'perclie1,  telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac
    Text1(14).Text = DBLet(miRsAux!perclie1, "T")
    Text1(15).Text = DBLet(miRsAux!telclie1, "T")
    Text1(16).Text = DBLet(miRsAux!faxclie1, "T")
    Text1(17).Text = DBLet(miRsAux!maiclie1, "T")
    Text1(18).Text = DBLet(miRsAux!perclie2, "T")
    Text1(19).Text = DBLet(miRsAux!telclie2, "T")
    Text1(20).Text = DBLet(miRsAux!faxclie2, "T")
    Text1(21).Text = DBLet(miRsAux!maiclie2, "T")
    campo = "Cliente potencial: " & CadenaDesdeOtroForm & DBLet(miRsAux!observac, "T")
    Text1(22).Text = campo
    
    miRsAux.Close
    
    
    
    CadenaDesdeOtroForm = ""
    VerCliente = 0
    
    Set miRsAux = Nothing

    Exit Sub

eBuscarHuecoCliente:
    MuestraError Err.Number, Err.Description
End Sub
Private Sub BotonAnyadirLinea()
Dim aModo As Byte
Dim vWhere As String
    
   
    If ModificaLineas = 2 Then Exit Sub
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    '   5.-  Mantenimiento Lineas de direcciones/dpto
'   6.-  "              "     de direcciones de envio
'   7.-  Per. contacto
'   8.-  Renting
'   9.-  Telefonia
'   10.- Fitosan
'   11.- Campos
    aModo = Modo
    If aModo = 5 Then
        Me.SSTab1.Tab = 2
    ElseIf aModo = 6 Then
        Me.SSTab1.Tab = 3
    ElseIf aModo = 7 Then
        Me.SSTab1.Tab = 6
    ElseIf aModo = 9 Then
        Me.SSTab1.Tab = 8
    ElseIf aModo = 10 Then
        Me.SSTab1.Tab = 9
    ElseIf aModo = 11 Then
        Me.SSTab1.Tab = 10
    Else
        Me.SSTab1.Tab = 7
    End If
    PonerModoFrame 3, aModo  '3: Insertar
    ModificaLineas = 1 'Insertar
    lblIndicador.Caption = "Insertar l�nea " & DevuelveTextoModAnyadir(aModo)
    PonerModoOpcionesMenu

    'Obtenemos la siguiente numero de Direc./Dpto
    vWhere = "codclien=" & Text1(0).Text
    If aModo = 5 Then
        AnyadirLinea DataGrid6, Data2
        LLamaLineasDirec ObtenerAlto(DataGrid6, 20), 1
        
        Text3(0).Text = SugerirCodigoSiguienteStr("sdirec", "coddirec", vWhere)
        PonerFoco Text3(0)
        
        'Si no es herbelca, ofertamos la misma zona que el cliente ppal
        txtZona(14).Text = ""
        If Not (vParamAplic.AlmacenB > 1) Then
            Text3(14).Text = Text1(11).Text
            Me.txtZona(14).Text = Text2(11).Text
        End If
        
        PonerFoco Text3(1)
    ElseIf aModo = 6 Then
        Text4(0).Text = SugerirCodigoSiguienteStr("sdirenvio", "coddiren", vWhere)
        
        
        AnyadirLinea DataGrid7, Data3
        LLamaLineasDirenEvio ObtenerAlto(DataGrid7, 20), 1
        

        
        'Si no es herbelca, ofertamos la misma zona que el cliente ppal
        txtZona(9).Text = ""
        If Not (vParamAplic.AlmacenB > 1) Then
            Text4(9).Text = Text1(11).Text
            Me.txtZona(9).Text = Text2(11).Text
        End If
        
        PonerFoco Text4(1)
        
        
        
    ElseIf Modo = 7 Then
        'Situamos el grid al final
        AnyadirLinea DataGrid1, Data4
        LLamaLineasDatosContacto ObtenerAlto(DataGrid1, 20), 1
        txtauxDC(8).Text = SugerirCodigoSiguienteStr("scliendp", "id", vWhere)
        PonerFoco Me.txtauxDC(0)
        'cboCargo.ListIndex = 0 'el vacio
        
    ElseIf Modo = 9 Then
        'Situamos el grid al final
        AnyadirLinea DataGrid3, data6
        LLamaLineasTfnia ObtenerAlto(DataGrid3, 20), 1
        
        
        'Algunos valores por defecto
        Me.cboOperadorTfnnia2(1).ListIndex = 0
        txtauxTfno(9).Text = Format(Now, "dd/mm/yyyy")
        txtauxTfno(7).Text = 0 'cuota minima
        txtauxTfno(8).Text = 0 'puntos
        PonerFoco Me.txtauxTfno(0)
    ElseIf Modo = 10 Then
        'Situamos el grid al final
        AnyadirLinea DataGrid4, Data7
        LLamaLineasFito ObtenerAlto(DataGrid4, 30), 1
        txtauxFito(4).Text = SugerirCodigoSiguienteStr("sclienmani", "id", vWhere)
        PonerFoco txtauxFito(0)
    ElseIf Modo = 11 Then
        'Situamos el grid al final
        AnyadirLinea DataGrid5, data8
        LLamaLineasCamposHuertos ObtenerAlto(DataGrid5, 30), 1
        Me.txtauxMarja(0).Text = SugerirCodigoSiguienteStr("sclienhuertos", "id", vWhere)
        PonerFoco txtauxMarja(1)
    Else
        AnyadirLinea DataGrid2, Data5
        LLamaLineasRenting ObtenerAlto(DataGrid2, 30), 1
        txtauxRent(0).Text = SugerirCodigoSiguienteStr("sclienrenting", "id", vWhere)
        PonerFoco Me.txtauxRent(1)
         
    End If
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        If vParamAplic.TieneTelefonia2 > 0 Then LLamaLineasTfnia ObtenerAlto(DataGrid3, 20), 0
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


Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data

    
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, index
            PonerCampos
            'MostrarSituacion True

            

End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    If Me.SSTab1.Tab > 1 Then Me.SSTab1.Tab = 0
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Function DevuelveTextoModAnyadir(QueModoEs As Byte)
    
    Select Case QueModoEs
    Case 5
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            DevuelveTextoModAnyadir = "Obras"
        Else
            DevuelveTextoModAnyadir = IIf(vParamAplic.TipoDtos, "Dpto.", "direc.")
        End If
        
    Case 6
        DevuelveTextoModAnyadir = "dir. envio"
    Case 7
        DevuelveTextoModAnyadir = "contacto"
    Case 9
        DevuelveTextoModAnyadir = "telefonia"
    Case 10
        DevuelveTextoModAnyadir = "carnet fito"
    Case 11
        DevuelveTextoModAnyadir = "campos"
    Case 8
        DevuelveTextoModAnyadir = "renting"
    Case Else
        DevuelveTextoModAnyadir = ""
    End Select
End Function


Private Sub BotonModificarLinea()
Dim aModo As Byte
'Modificar una linea
    aModo = Modo
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    '   5.-  Mantenimiento Lineas de direcciones/dpto
'   6.-  "              "     de direcciones de envio
'   7.-  Per. contacto
'   8.-  Renting
'   9.-  Telefonia
'   10.- Fitosan
'   11.- Campos
    If aModo = 5 Then
        If Data2.Recordset.EOF Then Exit Sub
        If Data2.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 2
        
    ElseIf aModo = 6 Then
        If Data3.Recordset.EOF Then Exit Sub
        If Data3.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 3
    ElseIf aModo = 7 Then
        If Data4.Recordset.EOF Then Exit Sub
        If Data4.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 6
        
    ElseIf aModo = 9 Then
        If data6.Recordset.EOF Then Exit Sub
        If data6.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 8
        
    ElseIf aModo = 10 Then
        If Data7.Recordset.EOF Then Exit Sub
        If Data7.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 9
        
    ElseIf aModo = 11 Then
        If data8.Recordset.EOF Then Exit Sub
        If data8.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 10
        
    Else
        'Renting
        If Data5.Recordset.EOF Then Exit Sub
        If Data5.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 7
    End If
    
    
    
    
    
       
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 4, aModo 'ModoFrame=4 -> Modificar
    Me.lblIndicador.Caption = "Modificar linea " & DevuelveTextoModAnyadir(aModo)
    ModificaLineas = 2 'Modificar
    PonerModoOpcionesMenu
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    If aModo = 5 Then
        LLamaLineasDirec ObtenerAlto(DataGrid6, 20), 2
        BloquearTxt Text3(0), True
        PonerFoco Text3(1)
    ElseIf aModo = 6 Then
        LLamaLineasDirenEvio ObtenerAlto(DataGrid7, 20), 2
        BloquearTxt Text4(0), True
        PonerFoco Text4(1)
    ElseIf aModo = 7 Then
    
                
        LLamaLineasDatosContacto ObtenerAlto(DataGrid1, 20), 2
        txtauxDC(0).Text = Data4.Recordset!Nombre
        
        
        PonerFoco Me.txtauxDC(0)
        
    ElseIf aModo = 8 Then
    
        LLamaLineasRenting ObtenerAlto(DataGrid2, 20), 2
        
        For NumRegElim = 0 To txtauxRent.Count - 1

           
                If IsNull(Data5.Recordset.Fields(NumRegElim)) Then
                    txtauxRent(NumRegElim).Text = ""
                Else
                    txtauxRent(NumRegElim).Text = Data5.Recordset.Fields(NumRegElim)
                End If

        Next
            
        
        
        PonerFoco Me.txtauxRent(1)
        
    ElseIf aModo = 9 Then
    
                
        LLamaLineasTfnia ObtenerAlto(DataGrid3, 20), 2
        BloquearTxt txtauxTfno(0), True
        txtauxTfno(0).Text = data6.Recordset!idtelefono
        txtauxTfno(1).Text = DBLet(data6.Recordset!IMEI, "T")
        txtauxTfno(2).Text = DBLet(data6.Recordset!SIM, "T")
        NumRegElim = DBLet(data6.Recordset!CodDirec, "N")
        If NumRegElim > 0 Then txtauxTfno(4).Text = NumRegElim
        txtauxTfno_LostFocus 4
        SituarCombo Me.cboOperadorTfnnia2(0), DBLet(data6.Recordset!Operador, "N")
        SituarCombo Me.cboOperadorTfnnia2(1), DBLet(data6.Recordset!procedencia, "N")
        NumRegElim = DBLet(data6.Recordset!clienppal, "N")
        If NumRegElim > 0 Then txtauxTfno(5).Text = NumRegElim
        txtauxTfno_LostFocus 5
        
        If Not IsNull(data6.Recordset!modelo) Then txtauxTfno(6).Text = DBLet(data6.Recordset!modelo, "N")
        txtauxTfno_LostFocus 6
        txtauxTfno(7).Text = DBLet(data6.Recordset!cuotaminima, "T")
        txtauxTfno(8).Text = DBLet(data6.Recordset!Puntos, "T")
        txtauxTfno(9).Text = DBLet(data6.Recordset!fechaalta, "T")
        txtauxTfno(10).Text = DBLet(data6.Recordset!fecharenove, "T")
        
        If vParamAplic.TelefoniaVtaPlazos Then
            If Not IsNull(data6.Recordset!modelo) Then txtauxTfno(11).Text = DBLet(data6.Recordset!artplazos, "T")
            
            If Not IsNull(data6.Recordset!PlazosMeses) Then txtauxTfno(12).Text = data6.Recordset!PlazosMeses
            If Not IsNull(data6.Recordset!ImportePlazo) Then txtauxTfno(13).Text = Format(data6.Recordset!ImportePlazo, FormatoPrecio)
            If Not IsNull(data6.Recordset!PlazosOrigen) Then txtauxTfno(14).Text = data6.Recordset!PlazosOrigen
            
        End If
        'PonerFoco Me.txtauxTfno(1)
        PonerFocoCbo Me.cboOperadorTfnnia2(0)
        
    ElseIf aModo = 10 Then
        LLamaLineasFito ObtenerAlto(DataGrid4, 20), 2
        txtauxFito(0).Text = DBLet(Data7.Recordset!CIF, "T")
        txtauxFito(1).Text = DBLet(Data7.Recordset!Nombre, "T")
        txtauxFito(2).Text = DBLet(Data7.Recordset!numcarnet, "T")
        txtauxFito(3).Text = DBLet(Data7.Recordset!Telefono, "T")
        txtauxFito(4).Text = DBLet(Data7.Recordset!Id, "T")
        txtauxFito(5).Text = DBLet(Data7.Recordset!fcaducidad, "F")
        If DBLet(Data7.Recordset!Tipo, "N") = "Cualificado" Then
            cboFitos(0).ListIndex = 1
        Else
            cboFitos(0).ListIndex = 0
            'SituarCombo Me.cboFitos, DBLet(data7.Recordset!Tipo, "N")
        End If
            
        cboFitos(1).ListIndex = Abs(UCase(DBLet(Data7.Recordset!Prov, "T")) = "SI")
        
        PonerFoco Me.txtauxFito(1)
        
    ElseIf aModo = 11 Then
        'Campos huertos
        LLamaLineasCamposHuertos ObtenerAlto(DataGrid5, 20), 2
        txtauxMarja(0).Text = DBLet(data8.Recordset!Id, "T")
        txtauxMarja(1).Text = Format(DBLet(data8.Recordset!poligono, "N"), "0000")
        txtauxMarja(2).Text = Format(DBLet(data8.Recordset!parcela, "N"), "0000")
        txtauxMarja(3).Text = Format(DBLet(data8.Recordset!recintos, "N"), "0000")
        txtauxMarja(4).Text = DBLet(data8.Recordset!supsigpa, "N")
        txtauxMarja(5).Text = DBLet(data8.Recordset!supderec, "N")
        
        cbomarjal.Text = DBLet(data8.Recordset!partida, "T")
        
        
        BloquearTxt txtauxMarja(0), True
        PonerFoco txtauxMarja(1)
    End If
End Sub


Private Sub BotonEliminar()
Dim cad As String
Dim b As Boolean

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If Not PuedeEliminarCliente(CLng(Data1.Recordset.Fields(0))) Then Exit Sub


    '### a mano
    cad = "�Seguro que desea eliminar el Cliente?"
    cad = cad & vbCrLf & "Cod. : " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        conn.BeginTrans
        b = EliminardeBD
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If

        
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


Private Sub BotonEliminarLinea()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String, cad2 As String
Dim I As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
       
    If vParamAplic.Renting Then
        cad = "codclien = " & Data1.Recordset!codClien & " AND coddirec"
        cad = DevuelveDesdeBD(conAri, "count(*)", "sclienrenting", cad, CStr(Data2.Recordset.Fields(0)), "N")
        If cad = "" Then cad = "0"
        If Val(cad) > 0 Then
            MsgBox "Existen " & RentingLB & " de clientes asociados a este departamento/direccion", vbExclamation
            Exit Sub
        End If
    End If
       
    If vParamAplic.TieneTelefonia2 > 0 Then
        cad = "codclien = " & Data1.Recordset!codClien & " AND coddirec"
        cad = DevuelveDesdeBD(conAri, "count(*)", "sclientfno", cad, CStr(Data2.Recordset.Fields(0)), "N")
        If cad = "" Then cad = "0"
        If Val(cad) > 0 Then
            MsgBox "Existen tel�fonos de clientes asociados a este departamento/direccion", vbExclamation
            Exit Sub
        End If
    End If
       
       
    ModificaLineas = 3 'Eliminar
    
    'Dependiendo del parametro de la aplicacion trabajamos con Dpto o Direc.
    If vParamAplic.HayDeparNuevo = 1 Then
        cad2 = " Dpto. "
        cad = " el Departamento?"
    ElseIf vParamAplic.HayDeparNuevo = 0 Then
        cad2 = " Direc. "
        cad = " la Direcci�n?"
    Else
        cad2 = " Obra "
        cad = " la obra?"
    End If
    
    cad = "�Seguro que desea eliminar " & cad & vbCrLf
    cad = cad & vbCrLf & "Cod." & cad2 & ": " & Format(Data2.Recordset.Fields(0), "000")
    cad = cad & vbCrLf & "Nombre" & cad2 & ": " & Data2.Recordset.Fields(1)
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data2.Recordset.AbsolutePosition
        cad = "DELETE FROM sdirec WHERE codclien =" & Data1.Recordset!codClien
        cad = cad & " AND coddirec=" & Data2.Recordset!CodDirec
        conn.Execute cad
        
        'Para borrar en arimoeny
        If Text1(35).Text <> "" Then
            'SI NO tiene cta contable NO tiene dpto
            cad2 = " WHERE codmacta= '" & Text1(35).Text & "' AND Dpto = " & Text3(0).Text
            ConnConta.Execute "DELETE FROM departamentos " & cad2
        End If
        I = Data2.Recordset.AbsolutePosition
        I = I - 1
        
        CargaLineas True, 5
        
        If I > 0 Then Data2.Recordset.Move I
        'PonerDatosForaGridContacto False
            

        ModificaLineas = 0

    End If
    
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data2.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub



Private Sub BotonEliminarLineaDirEnvio()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String
Dim I As Integer

    If Data3.Recordset.EOF Then Exit Sub
    If Data3.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    If Not PuedeEliminarDirecEnvio(True, Text1(0).Text, CInt(Data3.Recordset!coddiren)) Then Exit Sub
    
    ModificaLineas = 3 'Eliminar
    
    
    cad = "�Seguro que desea eliminar la direccion de envio" & cad & vbCrLf
    cad = cad & vbCrLf & "Codigo:  " & Format(Data3.Recordset.Fields(0), FormatoCampo(Text4(0)))
    cad = cad & vbCrLf & "Nombre:  " & Data3.Recordset.Fields(1)
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data3.Recordset.AbsolutePosition
        
        
        cad = "DELETE FROM sdirenvio WHERE codclien =" & Data1.Recordset!codClien
        cad = cad & " AND coddiren=" & Data3.Recordset!coddiren
        conn.Execute cad
        
        
        
        CargaLineas True, 6
        
        If NumRegElim > 0 Then Data3.Recordset.Move NumRegElim

        
        
        ModificaLineas = 0
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data3.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Private Sub cmdCargos_Click()
    If Modo = 7 Then
        If ModificaLineas > 0 Then imgBuscar_Click 14
    End If
End Sub

Private Sub cmdCatalogo_Click()
    If Modo <> 2 Then Exit Sub
    If Text1(0).Text = "" Then Exit Sub
    
    frmAlmCatalogos.desdeArticulos = False
    frmAlmCatalogos.Codigo = Text1(0).Text
    frmAlmCatalogos.Show vbModal
    CargaDatosLWDoc

End Sub

Private Sub cmdFitos_Click(index As Integer)
     If index = 0 Then
         
        imgFecha(0).Tag = 3000
        Set frmF = New frmCal
        frmF.Fecha = Now
   

       frmF.Show vbModal
       Set frmF = Nothing
       If Me.txtauxFito(5).Text <> "" Then PonerFoco txtauxFito(5)
    End If
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
        Else
            PonerModo 0
        End If
        Me.cmdCancelar.Cancel = True
    Else 'Regresar
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
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




Private Sub Renovar_Cambiar_Telefono(Renovar As Boolean)
    
    
    
    BuscaChekc = PonerTrabajadorConectado(CadenaConsulta)
    
    If BuscaChekc = "" Then
        MsgBox "Imposible asignar trabajador conectado", vbExclamation
    Else
        'Cliente|telefno|compa�ia|modelo|puntos|ultrenovacion|codclien|
        BuscaChekc = Text1(1).Text & "[" & Text1(0).Text & "]|" & CStr(data6.Recordset!idtelefono) & "|"
        BuscaChekc = BuscaChekc & CStr(data6.Recordset!Nombre) & "|"
        If txtauxTfno(6).Text <> "" Then BuscaChekc = BuscaChekc & txtauxTfno(6) & " - " & Text5(6).Text
        BuscaChekc = BuscaChekc & "|" & txtauxTfno(8).Text & "|" & txtauxTfno(10).Text & "|" & Text1(0).Text & "|"
        frmListado3.OtrosDatos = BuscaChekc
        
        If Renovar Then
            frmListado3.Opcion = 42
            frmListado3.Show vbModal
            
            'Para que se situe despues
            CadenaConsulta = "IdTelefono = '" & Me.txtauxTfno(0).Text & "'"
            
            
        Else
            'Cambiar de socio
            frmListado3.Opcion = 44
            frmListado3.Show vbModal
            CadenaConsulta = ""
    
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    CargaLineas True, 2
    If CadenaConsulta <> "" Then data6.Recordset.Find CadenaConsulta



    If RecuperaValor(lwCRM.Tag, 1) = "0" Then
        ModoFrame2 = Modo
        Modo = 2
        CargaDatosLWCRM
        Modo = CByte(ModoFrame2)
        ModoFrame2 = 0
    End If
    
    BuscaChekc = ""
    CadenaConsulta = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRenting_Click(index As Integer)

   If index = 0 Then
        'Departamento
        imgBuscar(0).Tag = 1000
        MandaBusquedaPrevia2 "codclien=" & Text1(0).Text
        
        
        
        
   ElseIf index = 3 Then
        'tipco
        BuscaChekc = ""
        Set frmMtoTipCo = New frmManTiposContrato
        frmMtoTipCo.DatosADevolverBusqueda = "0"
        frmMtoTipCo.Show vbModal
        Set frmMtoTipCo = Nothing
        If BuscaChekc <> "" Then
            Me.txtauxRent(8).Text = RecuperaValor(BuscaChekc, 1)
            Me.txtauxRent(9).Text = RecuperaValor(BuscaChekc, 2)
            PonerFoco txtauxRent(10)
            BuscaChekc = ""
        End If
   
   
   
   Else
        'FECHAS
        If index = 1 Then
            imgFecha(0).Tag = 1004
        Else
            imgFecha(0).Tag = 1006
        End If
        Set frmF = New frmCal
        frmF.Fecha = Now
   
       
       
       'PonerFormatoFecha Text1(Indice)
       'If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)
    
       Screen.MousePointer = vbDefault
       frmF.Show vbModal
       Set frmF = Nothing

    End If
End Sub

Private Sub Data2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 5 And ModificaLineas > 0 Then Exit Sub
    Debug.Print adReason & ": " & adStatus
    If Not Data2.Recordset.EOF Then
        PonerDatosForaGridDepartamentos False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridDepartamentos True
    End If
End Sub

Private Sub data3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Modo = 6 And ModificaLineas > 0 Then Exit Sub
    If Not Data3.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGridDirEnvio False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridDirEnvio True
    End If
End Sub

Private Sub Data4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 7 And ModificaLineas > 0 Then Exit Sub
    If Not Data4.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGridContacto False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridContacto True
    End If
End Sub

Private Sub data5_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 8 And ModificaLineas > 0 Then Exit Sub
    If Not Data5.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGridRent False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridRent True
    End If
End Sub

Private Sub data6_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 9 And ModificaLineas > 0 Then Exit Sub
    If Not data6.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGridTfno False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridTfno True
    End If
End Sub


Private Sub data8_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 11 And ModificaLineas > 0 Then Exit Sub
    
    If Not data8.Recordset.EOF Then
        'Caption = data4.Recordset!Id
        PonerDatosForaGridCamposHuertos False
    Else
       ' Caption = "EOF"
         PonerDatosForaGridCamposHuertos True
    End If
End Sub



Private Sub DataGrid1_Click()
    If Not Data4.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridContacto False
End Sub

Private Sub DataGrid2_Click()
    If Not Data5.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridRent False
End Sub

Private Sub DataGrid3_Click()
     If Not data6.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridTfno False
End Sub

Private Sub DataGrid5_Click()
    If Not data8.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridCamposHuertos False
End Sub

Private Sub DataGrid6_Click()
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridDepartamentos False
End Sub

Private Sub DataGrid7_Click()
    If Not Data3.Recordset.EOF And ModificaLineas <> 1 Then PonerDatosForaGridDirEnvio False

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PriVezForm Then
        PriVezForm = False
        ProcesarCarpetaImagenes
        
        If DatosADevolverBusqueda = "" Then
            If VerCliente <> 0 Then
                If VerCliente > 0 Then
                    'QUiere ver el cliente:VerCliente
                    'Para whose, pero puede ponerse en cualquier sitio
                    CadenaConsulta = "select * from " & NombreTabla & " WHERE codclien = " & VerCliente
                    PonerCadenaBusqueda
                    PonerModo 2
                Else
                    BotonAnyadir
                End If
            End If
        End If
    End If
        
    If Modo = 1 Then PonerFoco Text1(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PriVezForm = True
        
        
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(1).Picture
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
    
'    ' Botonera Principal 2
'    With Me.Toolbar2
'        .HotImageList = frmPpal.imgListComun_OM2
'        .DisabledImageList = frmPpal.imgListComun_BN2
'        .ImageList = frmPpal.ImgListComun2
'        .Buttons(1).Image = 47
'        .Buttons(2).Image = 44
'        .Buttons(3).Image = 42
'        .Buttons(4).Image = 36
'    End With
'
    
    With Me.ToolbarAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    If vParamAplic.DireccionesEnvio Then
        SSTab1.TabVisible(3) = True
        With Me.ToolbarAux(1)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    Else
        SSTab1.TabVisible(3) = False
    End If

    With Me.ToolbarAux(2)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        With Me.ToolbarAux(3)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    End If
    
    
    If vParamAplic.Renting Then
        With Me.ToolbarAux(4)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    End If
    
    If vParamAplic.TieneTelefonia2 Then
        With Me.ToolbarAux(5)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    End If
    
    If vParamAplic.Huertos Then
        With Me.ToolbarAux(6)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
        End With
    End If
    
    
    
    Me.SSTab1.Tab = 0
    
    SSTab1.TabVisible(7) = vParamAplic.Renting
    SSTab1.TabCaption(7) = RentingLB
    SSTab1.TabVisible(5) = vParamAplic.TieneCRM
        
    
            
    'Coultamos password web
    Text1(45).visible = False
    Label1(19).visible = False
            
    'Marjal Chipos
    SSTab1.TabVisible(10) = vParamAplic.Huertos
   
    If vParamAplic.Huertos Then
        SSTab1.TabCaption(10) = "Campos"
       ' Toolbar1.Buttons(16).visible = True
        
        
       Me.imgFechaCampos(9).Picture = Me.imgBuscar(8).Picture
        
        
    End If
    
    'Telefonia
    SSTab1.TabVisible(8) = False
    If vParamAplic.TieneTelefonia2 > 0 Then
        'Toolbar1.Buttons(14).visible = vParamAplic.TieneTelefonia2 > 0
        SSTab1.TabVisible(8) = vParamAplic.TieneTelefonia2 > 0
        SSTab1.TabCaption(8) = "Telefon�a"
        Me.cmdAccionesTfno(1).Picture = frmPpal.imgListComun.ListImages(44).Picture
        Me.cmdAccionesTfno(5).Picture = frmPpal.imgListComun.ListImages(45).Picture
        
        'iconos para las cuotas
        Me.cmdAccionesTfno(2).Picture = frmPpal.imgListComun.ListImages(3).Picture
        Me.cmdAccionesTfno(3).Picture = frmPpal.imgListComun.ListImages(4).Picture
        Me.cmdAccionesTfno(4).Picture = frmPpal.imgListComun.ListImages(43).Picture
        
        
        
        'Venta a plazos telefono
        FrameModuloVtaPlazos.visible = vParamAplic.TelefoniaVtaPlazos
        txtauxTfno(3).Height = IIf(vParamAplic.TelefoniaVtaPlazos, 2925, 4065)
        If vParamAplic.TelefoniaVtaPlazos Then FrameModuloVtaPlazos.BorderStyle = 0
    End If
    'Si tienen renting
    cboFraRenting.visible = vParamAplic.Renting
    Label1(91).visible = vParamAplic.Renting
    Label1(91).Caption = Label1(91).Caption & RentingLB
    'Si NO tiene renting ocultamos el chk
    If vParamAplic.Renting Then
        Me.chkRentingDpto.Top = 4560
    Else
        Me.chkRentingDpto.Top = 14560
    End If
    'Si no tiene reting, el check de cuenta alternativa, lo bajamos
    If Not vParamAplic.Renting Then Me.chkAbonos.Top = 3720
    
    
    'Fitosantiarios
    'Toolbar1.Buttons(15).visible = vParamAplic.ManipuladorFitosanitarios2
    Me.SSTab1.TabVisible(9) = vParamAplic.ManipuladorFitosanitarios2
    If vParamAplic.ManipuladorFitosanitarios2 Then
        CargaComboManipulador
        SSTab1.TabCaption(9) = "Fitosanitarios"
    End If
    cboManipulador.visible = vParamAplic.ManipuladorFitosanitarios2
    Text1(57).visible = vParamAplic.ManipuladorFitosanitarios2
    If vParamAplic.ManipuladorFitosanitarios2 Then
        
        For kCampo = 0 To Me.ImageFito.Count - 3
            Me.ImageFito(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Next kCampo
        Me.ImageFito(4).Picture = frmPpal.imgListComun.ListImages(16).Picture
    End If
    
    
    
    'La nevegacion para albaranes, facturas....
    ImagenesNavegacion
    ImagenDocumento CInt(optDoc(0).Tag)
    If vParamAplic.TieneCRM Then ImagenCRM CByte(Me.optCRM(0).Tag)
    
    
    Me.chkTasaReciclado.Caption = "Tasa reciclado"
    
    'Comprobar si es Departamento o Direccion (segun paramatro)
    kCampo = 0 'DIRECCIONESS
    If vParamAplic.HayDeparNuevo = 1 Then
        Me.Toolbar1.Buttons(10).ToolTipText = "Departamentos"
        lblFramePp(0).Caption = "Departamentos"
      '  Me.Label1(22).Caption = "Cod. Dpto"
        Me.SSTab1.TabCaption(2) = "Departamentos"
        If vParamAplic.NumeroInstalacion = 6 Then
            lblFramePp(0).Caption = lblFramePp(0).Caption & " / OBRAS"
            Me.SSTab1.TabCaption(2) = "Dpto. / Obras"
            Label1(6).Caption = "Transportista"
        End If
        FrameCtaBanDpto.BorderStyle = 0
        Me.FrameCtaBanDpto.visible = True
        kCampo = 1
    ElseIf vParamAplic.HayDeparNuevo = 0 Then

        Me.FrameCtaBanDpto.visible = False
    Else
        'OBRA
        FrameCtaBanDpto.BorderStyle = 0
        Me.FrameCtaBanDpto.visible = True
        If InstalacionEsEulerTaxco Then
            'Pondra direcciones
        Else
            
            Me.Toolbar1.Buttons(10).ToolTipText = "Obras"
            lblFramePp(0).Caption = "Obras"
            Me.Label1(22).Caption = "Cod. obra"
            Me.SSTab1.TabCaption(2) = "Obras"
            
            kCampo = 1
        End If
    End If
    If kCampo = 0 Then
        Me.Toolbar1.Buttons(10).ToolTipText = "Direcciones"
        lblFramePp(0).Caption = "Direcciones"
        Me.SSTab1.TabCaption(2) = "Direcciones"
    End If
    
    
    
    SSTab1.TabVisible(11) = False
    
    'Si lleva puntos
    Text1(62).visible = vParamAplic.PtosAsignar > 0
    Me.chkPuntos.visible = vParamAplic.PtosAsignar > 0
    
    
    
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    ModificaLineas = 0
       
    'Si hay algun combo los cargamos
    CargarComboAlbaran
    CargarComboFacturacion
    CargarComboTipoIVA
    CargaComboTipoCliente
    CargaComboFrarRenting
    If vParamAplic.TieneTelefonia2 > 0 Then CargaComboTfnos_
    CargaComboPais
    CargaComboPrioridad
    
    
    
    'EnvFraEmail     en taxco es para ver si se le comunica facturas por facturaE
    If vParamAplic.NumeroInstalacion = vbTaxco Then chkEnvioFraEmail.Caption = "Fact. electr�nica"
    
    
    Me.lblSituacion.visible = False
    Me.Frame1(1).visible = False
    
    
    'Si no tiene el parametro de direcciones envio, NO se muestra el txt
    Me.Label1(84).visible = vParamAplic.DireccionesEnvio
    Me.imgBuscar(13).visible = vParamAplic.DireccionesEnvio
    Me.Text1(52).visible = vParamAplic.DireccionesEnvio
    Me.Text2(52).visible = vParamAplic.DireccionesEnvio
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        Label1(94).Caption = "Perfil credito"
        Label1(22).visible = True
        Text1(64).visible = True
        Text1(64).MaxLength = 3
        Text1(64).Width = 765
    End If
    
    'Pone el Tag del primer bot�n de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sclien, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: cuentas, BD: Conta.
    imgBuscar(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "sclien"
    Ordenacion = " ORDER BY codclien"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where false"
    Data1.Refresh
    
    
    txtauxDC(8).Left = 23000 'para que no se vea
    
    'Ponemos los datos del listview
    imgFecha(3).Tag = vEmpresa.FechaIni
    CargaColumnas 0
    If vParamAplic.TieneCRM Then CargaColumnasCRM 0
    
    'SSTab1.TabVisible(1) = vParamAplic.OperacionesAseguradas And vUsu.Nivel = 0
    'Me.SSTab1.TabCaption(1) = "Operaciones aseguradas"
    Me.SSTab1.TabCaption(1) = "Op. aseguradas - FACE"
    SSTab1.TabVisible(1) = vUsu.Nivel = 0
    
    
    If vParamAplic.OperacionesAseguradas And vUsu.Nivel = 0 Then CargaComboAseguradora
    
    'oolbarDoc.Buttons(15).visible = vParamAplic.PtosAsignar > 0
    Me.optDoc(7).visible = vParamAplic.PtosAsignar > 0
    
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        If vUsu.CodigoAgente > 0 Then FrameNavegaDoc.visible = False 'rameNavegaDocToolbarDoc.visible = False
        Label1(17).Caption = "Asociaci�n"
    End If
    
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
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkClienteV.Value = 0
    lblPerfil.Caption = ""
    Me.chkAbonos.Value = 0
    Me.chkPromociones.Value = 0
    Me.chkRentingDpto.Value = 0
    Me.chkReferencia.Value = 0
    Me.chkTasaReciclado.Value = 0
    Me.chkCorreo.Value = 0
    Me.chkPortesFac.Value = 0
    Me.chkRecargFinan.Value = 0
    Me.chkParticular.Value = 0
    chkPuntos.Value = 0
    Me.cboAlbaran.ListIndex = -1
    Me.cboFacturacion.ListIndex = -1
    Me.cboTipoIVA.ListIndex = -1
    Me.cboFraRenting.ListIndex = -1
    cboTipocliente.ListIndex = -1
    cboTipoASeg.ListIndex = -1
    cboPais.ListIndex = -1
    cboPrioridad.ListIndex = -1
    CargaLineas False, 8
    If vParamAplic.TieneTelefonia2 > 0 Then
        Me.chkTelefonia(0).Value = 0: Me.chkTelefonia(1).Value = 0: Me.chkTelefonia(2).Value = 0:: Me.chkTelefonia(3).Value = 0
        lwTfnoCuotas.ListItems.Clear
    End If
    If vParamAplic.ManipuladorFitosanitarios2 Then
        Me.chkManiProv.Value = 0
        cboManipulador.ListIndex = -1
    End If
    chkEnvioFraEmail.Value = 0
            
    If RecuperaValor(lw1.Tag, 1) = "6" Then CargarIMG ""
    
    
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    VerCliente = 0

    
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Actividades
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
'Agentes Comerciales   -  visitaor
    
    Text1(CInt(BuscaChekc)).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(CInt(BuscaChekc))
    Text2(CInt(BuscaChekc)).Text = RecuperaValor(CadenaSeleccion, 2)
    BuscaChekc = ""
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    
    CadenaDesdeOtroForm = CadenaSeleccion
    
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
  
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If Val(imgBuscar(0).Tag) >= 0 Then
            If Val(imgBuscar(0).Tag) >= 1000 Then
                'Departamentos en RENTING
                If Val(imgBuscar(0).Tag) = 1000 Then
                    txtauxRent(1).Text = RecuperaValor(CadenaDevuelta, 1)
                    txtauxRent(2).Text = RecuperaValor(CadenaDevuelta, 2)
                ElseIf Val(imgBuscar(0).Tag) = 1001 Then
                    Me.txtauxTfno(4).Text = RecuperaValor(CadenaDevuelta, 1)
                    Me.Text5(4).Text = RecuperaValor(CadenaDevuelta, 2)
                ElseIf Val(imgBuscar(0).Tag) = 1002 Then
                    'telefonia cliente ppal
                    Me.txtauxTfno(5).Text = RecuperaValor(CadenaDevuelta, 1)
                    Me.Text5(5).Text = RecuperaValor(CadenaDevuelta, 2)
                ElseIf Val(imgBuscar(0).Tag) = 1002 Then
                    'Modelo telefono
                    'imgBuscar(0).Tag) = 1003
                    Me.txtauxTfno(6).Text = RecuperaValor(CadenaDevuelta, 1)
                    Me.Text5(6).Text = RecuperaValor(CadenaDevuelta, 2)
                    
                Else
                    '1004
                End If
            Else
                'Se llama desde el bot�n de busqueda del campo Tipos de IVA
                'Recuperar solo el campo c�digo y Descripci�n
    '            Indice = Val(Me.imgBuscar(0).Tag)
                Text1(35).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(35).Text = RecuperaValor(CadenaDevuelta, 2)
        
            End If
        Else
            'Recupera todo el registro de Art�culos
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
        If Me.imgBuscar(0).Tag = 10 Then
            Text3(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
            Text3(4).Text = ObtenerPoblacion(Text3(3).Text, devuelve)  'Poblacion
            'provincia
            Text3(5).Text = devuelve
        Else
            'DIRECCIONES DE ENVIO
            Text4(4).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
            Text4(3).Text = ObtenerPoblacion(Text4(4).Text, devuelve)  'Poblacion
            'provincia
            Text4(5).Text = devuelve
        End If
    End If
End Sub

Private Sub frmDptoEnvio2_DatoSeleccionado(CadenaSeleccion As String)
    'If Modo = 6 Then
    If Modo < 3 Or Modo > 4 Then
        
        BuscaChekc = RecuperaValor(CadenaSeleccion, 1)
    Else
        Text1(52).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(52).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'Formas de Env�o
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
        Case 6
            Indice = 58
        Case 1004, 1006
            'Son las fechas del RENTING
            Me.txtauxRent(Val(imgFecha(0).Tag) - 1000).Text = Format(vFecha, "dd/mm/yyyy")
            Exit Sub
        Case 2000 To 2100
            Me.txtauxTfno(Val(imgFecha(0).Tag) - 2000).Text = Format(vFecha, "dd/mm/yyyy")
            Exit Sub
        Case 3000
            'Me.txtauxTfno(Val(imgFecha(0).Tag) - 2000).Text = Format(vFecha, "dd/mm/yyyy")
            Me.txtauxFito(5).Text = Format(vFecha, "dd/mm/yyyy")
            Exit Sub
        Case 4000 To 4100
             
            Me.txtauxMarja(Val(imgFecha(0).Tag) - 4000).Text = Format(vFecha, "dd/mm/yyyy")
            Exit Sub
        
    End Select
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Pago
    Text1(23).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(23)
    Text2(23).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmModeloTel_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtauxTfno(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.Text5(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTipCo_DatoSeleccionado(CadenaSeleccion As String)
    BuscaChekc = CadenaSeleccion 'luego, alli(.show) lo ponemos en los txt
End Sub

Private Sub frmR_DatoSeleccionado(CadenaSeleccion As String)
'Rutas
    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(12)
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
'Situaciones
    Text1(42).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(42)
    Text2(42).Text = RecuperaValor(CadenaSeleccion, 2)
    txtSit.Text = Text2(42).Text
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(37).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(37)
    Text2(37).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
'Zonas
    If BuscaChekc = "" Then
        Text1(11).Text = RecuperaValor(CadenaSeleccion, 1)
        FormateaCampo Text1(11)
        Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
        
    Else
        If BuscaChekc = "15" Then
            Text3(14).Text = RecuperaValor(CadenaSeleccion, 1)
            Me.txtZona(14).Text = RecuperaValor(CadenaSeleccion, 2)
        Else
            Text4(9).Text = RecuperaValor(CadenaSeleccion, 1)
            Me.txtZona(9).Text = RecuperaValor(CadenaSeleccion, 2)
        End If
    End If
End Sub

Private Sub Image1_Click()
    If Modo <> 2 Then Exit Sub
    If CByte(RecuperaValor(lw1.Tag, 1)) = 6 Then
          LanzaVisorMimeDocumento Me.hwnd, Me.lw1.SelectedItem.SubItems(2)
'        If Not Me.lw1.SelectedItem Is Nothing Then
'            CadenaDesdeOtroForm = ""
'            frmFichaTecIMG.vDatos = Text1(0).Text & "|" & Text1(1).Text & "|" & lw1.SelectedItem.SubItems(2) & "|"
'            frmFichaTecIMG.Opcion = 1
'            frmFichaTecIMG.Show vbModal
'        End If
    End If
End Sub

Private Sub ImageFito_Click(index As Integer)
Dim Puede As Boolean
Dim J As Integer
    
    
    'Listado fito
    If index = 4 Then
        frmListado3.Opcion = 64
        frmListado3.Show vbModal
        Exit Sub
    End If
    
    Puede = False
    If Modo <> 2 Then
        If Modo = 4 Then
            If index <= 1 Then Puede = True
        Else
            If Modo = 10 And ModificaLineas = 0 Then Puede = True
        End If
    Else
        Puede = True
    End If
    
    
    
    'Asociados
    If Puede Then
        If index >= 2 Then
            'Tiene que tener ADO con datos
            If Data7.Recordset.EOF Then Puede = False
        End If
    End If
    
    
    If Not Puede Then Exit Sub
            
    'Si no existe lo metemos
    If index < 2 Then
        'Carnet y DNI del asociado PPAL
        
        CadenaConsulta = DevuelveDesdeBD(conAri, "codigo", "sfichdocs", "codclien = " & Text1(0).Text & " AND TipoDoc ", CStr(index + 1))
        
        If CadenaConsulta = "" Then
            'NO EXISTE. La creamos
            'EXISTE. la vemos
            LanzaAnyadirImagenDocumento index + 1
        Else
            If RecuperaValor(lw1.Tag, 1) <> "6" Then optDoc_Click 6 '  Hacer_ButtonClick 13, 6                  'Ponemos visible los documentos
                 
            'Si existe. Lo busco en los lw
            For J = 1 To lw1.ListItems.Count
                'eN SUBITEM4 TENEMOS 0 DOC  1 dni  2 cARNET
                If lw1.ListItems(J).SubItems(4) = index + 1 Then
                    Set lw1.SelectedItem = lw1.ListItems(J)
                    lw1.ListItems(J).Selected = True
                    
                    Image1_Click
                End If
            Next
            CadenaConsulta = ""
        End If
    Else
        'Del autorizado
        'Si existe, lo traere y lo visualizare
        J = 7
        If index = 3 Then J = 8
        
        If Data7.Recordset.Fields(J) = "" Then
            'NO existe
            LanzaAnyadirImagenDocumento 199 + index
        Else
            'Lo traemos y los mostramos
            If index = 2 Then
                CadenaConsulta = "ImgDNI,DocDNI"
            Else
                CadenaConsulta = "ImgManipula , DocManipula"
            End If
            CadenaConsulta = "Select " & CadenaConsulta & " from sclienmani WHERE codclien = " & Text1(0).Text & " AND id =" & Data7.Recordset!Id
            
            Adodc1IMG.ConnectionString = conn
            Adodc1IMG.RecordSource = CadenaConsulta
            Adodc1IMG.Refresh
            
            CadenaConsulta = Adodc1IMG.Recordset.Fields(1)
            CadenaConsulta = App.Path & "\ImgFicFT\" & CadenaConsulta
            If Dir(CadenaConsulta, vbArchive) <> "" Then Kill CadenaConsulta
            
            If LeerBinary(Adodc1IMG.Recordset.Fields(0), CadenaConsulta) Then LanzaVisorMimeDocumento Me.hwnd, CadenaConsulta
            CadenaConsulta = ""
        End If
        
        
    End If
    
End Sub

Private Sub imgBuscar_Click(index As Integer)
    Dim Indice As Byte

    'Disitnto de Observaciones
    If index = 11 Or index = 17 Or index = 21 Then
        'Observaciones
    
    Else
        'If Not (Index = 18 Or Index = 19 Or Index = 20 Or Index = 23) Then
        '    If Modo = 2 Or Modo = 0 Or Modo > 4 Then Exit Sub
        'End If
        
        If index = 15 Or index = 16 Then
            If Modo <> 5 Then
                Exit Sub
            End If
        End If
        
        If index = 13 Then
            'En insertar NO VA direccion envio habitual
            If Modo = 3 Then
                MsgBox "Hasta que no cree el cliente no podra tener direcciones envio", vbExclamation
                Exit Sub
            End If
        End If
    End If
    If index = 18 Or index = 19 Or index = 20 Or index = 23 Then
        If Modo <> 9 Then
            If Modo <> 1 Then Exit Sub
        Else
            If ModificaLineas = 0 Then Exit Sub
        End If
    End If
    
    If index = 24 Then
        'RENTING
        If Not cmdRenting(0).visible Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

    Select Case index
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
            If index = 2 Then
                If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            Else
                BuscaChekc = index
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
            Indice = 23
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5  'Cuenta Contable
            imgBuscar(0).Tag = index
            MandaBusquedaPrevia2 "apudirec= 'S'"
            imgBuscar(0).Tag = -1
            Indice = 35
            
        Case 6, 22 'C�digo de Agente
            Indice = 36
            If index = 22 Then Indice = 61
            BuscaChekc = Indice
            Set frmAc = New frmFacAgentesCom
            frmAc.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmAc.Show vbModal
            Set frmAc = Nothing
            
        Case 7 'C�digo de Tarifa
            Indice = 37
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 8 'C�digo de Situaci�n
            Indice = 42
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
        Case 9, 10, 12 'CPostal
            Me.imgBuscar(0).Tag = index
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If index = 9 Then
                Indice = 4
            Else
                PonerFoco Text3(3)
            End If
            Me.imgBuscar(0).Tag = -1
            VieneDeBuscar = True
       
        Case 11, 17
            'Campos MEMO
        
            If Modo = 5 Or Modo = 0 Then
            
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    If index = 11 Then
                        CadenaDesdeOtroForm = Text1(22).Text
                    Else
                        CadenaDesdeOtroForm = Text1(54).Text
                    End If
                        
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then
                        If index = 11 Then
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
                        If index = 11 Then
                            Text1(22).Text = Mid(CadenaDesdeOtroForm, 3)
                        Else
                            Text1(54).Text = Mid(CadenaDesdeOtroForm, 3)
                        End If
                    End If
                End If
                CadenaDesdeOtroForm = ""
            End If
            
            
         Case 13
            
                LanzaFrmDireccionEnvio
                
                
        Case 14
                
                frmFacCargos.Show vbModal
                CargaComboCargos
                SituarCboCargo
        Case 18
                imgBuscar(0).Tag = 1001
                MandaBusquedaPrevia2 "codclien=" & Text1(0).Text
        Case 19
                imgBuscar(0).Tag = 1002 'd
                MandaBusquedaPrevia2 ""
        Case 20
               ' imgBuscar(0).Tag = 1003  'modelo
               ' MandaBusquedaPrevia2 ""
               Set frmModeloTel = New frmTelefoniaModelos
               frmModeloTel.DatosADevolverBusqueda = "0|1|"
               frmModeloTel.Show vbModal
               Set frmModeloTel = Nothing
               
               
         Case 21
            'MEMO de tel�fono
            
                frmFacClienteObser.Modificar = False
                If Modo = 9 And ModificaLineas >= 1 Then frmFacClienteObser.Modificar = True
                CadenaDesdeOtroForm = ""
                frmFacClienteObser.Text1 = txtauxTfno(3).Text
                frmFacClienteObser.Show vbModal

                If Mid(CadenaDesdeOtroForm, 1, 1) = "1" Then
                    'Ha modificado
                    txtauxTfno(3).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
               
        Case 23
                'Articulo para telefonia
                CadenaDesdeOtroForm = ""
                Set FrmArt = New frmAlmArticu2
                FrmArt.DesdeTPV = False
                FrmArt.Show vbModal
                Set FrmArt = Nothing
                If CadenaDesdeOtroForm <> "" Then
                    Me.txtauxTfno(11).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                    Me.Text5(11).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
                    PonerFoco txtauxTfno(14)
                End If
        Case 24
            cmdRenting_Click 3
    End Select
    
    If index < 20 Then
        If index <> 10 Or index <> 12 Then PonerFoco Text1(Indice)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Or Modo > 4 Then
        If index <> 3 Then Exit Sub
   End If
   
   If index = 3 And Modo <> 2 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case index
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
    Case 6
        Indice = 58
   End Select
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   
   'Para la fecha de la navegacion
   If index = 3 And Text1(46).Text <> "" Then
        imgFecha(3).Tag = Text1(46).Text
        CargaDatosLWDoc
    End If
End Sub

Private Sub imgFechaCampos_Click(index As Integer)
Dim b As Boolean
        
        b = False
        If Modo = 11 Then
            If ModificaLineas > 0 Then b = True
        Else
            If Modo <> 2 Then Exit Sub
        End If
        
        
        If index = 9 Then
            'Campo mobservaciones
                frmFacClienteObser.Modificar = b
                CadenaDesdeOtroForm = ""
                frmFacClienteObser.Text1 = Me.txtauxMarja(9).Text
                frmFacClienteObser.Show vbModal

                If b Then
                    If Mid(CadenaDesdeOtroForm, 1, 1) = "1" Then
                        'Ha modificado
                        txtauxMarja(9).Text = Mid(CadenaDesdeOtroForm, 3)
                    End If
                End If
            
        Else
                
            If Not b Then Exit Sub
            
            imgFecha(0).Tag = 4000 + index
            Set frmF = New frmCal
            frmF.Fecha = Now
            If Me.txtauxMarja(index).Text <> "" Then frmF.Fecha = CDate(txtauxMarja(index).Text)
            frmF.Show vbModal
        End If
        PonerFoco txtauxTfno(index)
End Sub

Private Sub imgFechaTf_Click(index As Integer)
        
        If Modo <> 1 Then
            If Modo <> 9 Then
                Exit Sub
            Else
                If ModificaLineas = 0 Then Exit Sub
            End If
        End If
                
        
        
        imgFecha(0).Tag = 2000 + index
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Me.txtauxTfno(index).Text <> "" Then frmF.Fecha = CDate(txtauxTfno(index).Text)
        frmF.Show vbModal
        
        PonerFoco txtauxTfno(index)
        
End Sub

Private Sub ImgMail_Click(index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case index
        Case 0: dirMail = Text1(17).Text
        Case 1: dirMail = Text1(21).Text
        Case 2: dirMail = Text3(9).Text
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



Private Sub lw1_Click()
  If RecuperaValor(lw1.Tag, 1) = "6" Then
    If Not lw1.SelectedItem Is Nothing Then CargarIMG lw1.SelectedItem.SubItems(2)

       
  End If
End Sub

Private Sub lw1_DblClick()
Dim Seleccionado As Long
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'ALBARANES
        If vParamAplic.TipoFormularioClientes = 0 Then
        
            If vParamAplic.HaciendoFrmulariosGrandes Then
                Set frmAlbG = New frmFacEntAlbaranesGR
                frmAlbG.hcoCodMovim = lw1.SelectedItem.SubItems(1)
                frmAlbG.hcoCodTipoM = lw1.SelectedItem.Text
                frmAlbG.Show vbModal
                Set frmAlbG = Nothing
                
        
            Else
                Set frmAlb = New frmFacEntAlbaranes2
                frmAlb.hcoCodMovim = lw1.SelectedItem.SubItems(1)
                frmAlb.hcoCodTipoM = lw1.SelectedItem.Text
                frmAlb.Show vbModal
                Set frmAlb = Nothing
            End If
            
            
            If vParamAplic.PtosAsignar > 0 Then
                BuscaChekc = DevuelveDesdeBD(conAri, "puntos", "sclien", "codclien", Data1.Recordset!codClien)
                If BuscaChekc = "" Then BuscaChekc = "0"
                If DBLet(Data1.Recordset!Puntos, "N") <> CCur(BuscaChekc) Then
                    'Ha cambiado puntos.
                    PosicionarData
                    Text1(62).Text = Format(Data1.Recordset!Puntos, FormatoImporte)
                End If
            End If
            
        Else
            Set frmAlbS = New frmFacEntAlbSAIL
            frmAlbS.hcoCodMovim = lw1.SelectedItem.SubItems(1)
            frmAlbS.hcoCodTipoM = lw1.SelectedItem.Text
            frmAlbS.Show vbModal
            Set frmAlbS = Nothing
                 
            
        End If
        
    Case 0
        'OFERTAS
        
        
       
        
        If vParamAplic.TipoFormularioClientes = 0 Then
            Set frmOfe = New frmFacEntOfertas2
            frmOfe.EsHistorico = IIf(lw1.SelectedItem.ToolTipText <> "", True, False)
            frmOfe.DatosOferta = lw1.SelectedItem.Text
            frmOfe.Show vbModal
            Set frmOfe = Nothing
        Else
            'SAIL
            Set frmOfeS = New frmFacEntOferSAIL
            frmOfeS.DatosOferta = lw1.SelectedItem.Text
            frmOfeS.Show vbModal
            Set frmOfeS = Nothing
            
        End If
        
    Case 1
        'PEDIDOS
        If vParamAplic.TipoFormularioClientes = 0 Then
            Set frmPed = New frmFacEntPedidos
            frmPed.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
            frmPed.EsHistorico = False
            frmPed.Show vbModal
            Set frmPed = Nothing
            
        Else
            'SAIL
            Set frmPedS = New frmFacEntPedSail
            frmPedS.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
            frmPedS.EsHistorico = False
            frmPedS.Show vbModal
            Set frmPedS = Nothing
            
            
        End If
    Case 3
        'FACTURAS
        'Este no necesitamos crear instancias
        
        'Lo que ocurre que esta preparado para abrir la factura a partir de un albaran, con lo cual
        'En la funcion abrir factura, buscare un albaran de la factura para abrirlo
        AbrirFacturaLW
        
        
    Case 4
        'Precios especiales
        'No creamos instancias

        frmFacPreciosEspecial.CadenaSituarData = "'" & DevNombreSQL(lw1.SelectedItem.Text) & "'|" & Data1.Recordset!codClien & "|"
        frmFacPreciosEspecial.Show vbModal
        
    Case 6
        ImprimirImagen
        Screen.MousePointer = vbDefault
        Exit Sub
        
    Case 7
            
  
        'PUNTOS. Abre el frm
        AbrirAlbaranesPuntos
        Screen.MousePointer = vbDefault
        Exit Sub
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.index
    CargaDatosLWDoc
    If Not lw1.SelectedItem Is Nothing Then
        lw1.SelectedItem.Selected = False
        Set lw1.SelectedItem = Nothing
    End If
    
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub lwCRM_DblClick()
Dim Clave As String
Dim I As Integer
    If Modo <> 2 Then Exit Sub
    If lwCRM.ListItems.Count = 0 Then Exit Sub
    If lwCRM.SelectedItem Is Nothing Then Exit Sub




     'Llegados aqui
    Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
    Case 0
        'Aciones comerciales
        ' modificar o insertar acciones comerciales
        frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
        
        frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
        If Val(Me.lwCRM.SelectedItem.SubItems(4)) = 3 Then frmCRMMto.TipoPredefinido = 3  'Renovacion
        
        
        frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & _
            " AND scrmacciones.Tipo = " & lwCRM.SelectedItem.SubItems(4) & " And codClien = " & Data1.Recordset!codClien
        frmCRMMto.Show vbModal
    Case 1
        'Llamadas
        If lwCRM.SelectedItem.SmallIcon = 27 Then
            'Lee de sllama
            
            CadenaDesdeOtroForm = "`feholla`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " and `usuario`=" & DBSet(lwCRM.SelectedItem.SubItems(1), "T")
            frmLLamadasDatos2.SoloVer = True
            frmLLamadasDatos2.vModo = 4
            frmLLamadasDatos2.Show vbModal
        Else
            'Lee de acciones realizadas con tipo=1 .....
            
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 1 'Llamadas realizadas
            frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmacciones.Tipo = 1 And codClien = " & Data1.Recordset!codClien
            frmCRMMto.Show vbModal
            
        End If
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
    Case 2
        'Cobros. NO HACEMOS NADA
        'Nos piramos
        Exit Sub
        
    Case 3
        frmCrmObsDpto.Nuevo = False
        BuscaChekc = "dpto = " & Me.lwCRM.SelectedItem.SubItems(3) & " AND codclien "
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "observa", "scrmobsclien", BuscaChekc, CStr(Data1.Recordset!codClien))
        
        frmCrmObsDpto.Dpto = CByte(Me.lwCRM.SelectedItem.SubItems(3))
        frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
        frmCrmObsDpto.Tag = Data1.Recordset!codClien
        frmCrmObsDpto.Show vbModal
        
    Case 4
        'Reclamas n
            BuscaChekc = lwCRM.SelectedItem.SubItems(4) & "|" & Text1(1).Text & "|"
            If vParamAplic.ContabilidadNueva Then BuscaChekc = BuscaChekc & lwCRM.SelectedItem.Tag & "|"  'llevara el numlinea
            frmCRMReclamas.Intercambio = BuscaChekc
            frmCRMReclamas.Show vbModal
    
    Case 5
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 2 'Historial
            frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmacciones.Tipo = 2 And codClien = " & Data1.Recordset!codClien
            frmCRMMto.Show vbModal
    End Select
    Me.Refresh
    DoEvents
    
    
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
        Clave = lwCRM.SelectedItem.SubItems(4)
    Else
        Clave = lwCRM.SelectedItem.Text
    End If
    CargaDatosLWCRM
    
    Set lwCRM.SelectedItem = Nothing
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
        'para encontrar en las reclamas debe buscar por el campo codigo 4
        For I = 1 To lwCRM.ListItems.Count
            If Clave = lwCRM.ListItems(I).SubItems(4) Then
                
                Set lwCRM.SelectedItem = lwCRM.ListItems(I)
                Exit For
            Else
                lwCRM.ListItems(I).Selected = False
            End If
        Next
    Else
        For I = 1 To lwCRM.ListItems.Count
            If Clave = lwCRM.ListItems(I).Text Then
                Set lwCRM.SelectedItem = lwCRM.ListItems(I)
            Else
                lwCRM.ListItems(I).Selected = False
            End If
        Next
    End If
    BuscaChekc = ""
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
     If Modo >= 5 Then 'Eliminar lineas Art�culos x Almacen
        If Modo = 5 Then BotonEliminarLinea
        If Modo = 6 Then BotonEliminarLineaDirEnvio
        If Modo = 7 Then BotonEliminarLineaContacto
        If Modo = 8 Then BotonEliminarRenting
        If Modo = 9 Then BotonEliminarTelefono
        If Modo = 10 Then BotonEliminarManipulador
        If Modo = 11 Then BotonEliminarHuertos
     Else   'Eliminar Art�culo
        BotonEliminar
     End If
End Sub

Private Sub mnModificar_Click()
     If Modo >= 5 Then 'Modificar lineas Art�culos x Almacen
        'FALTA: bloquear la linea !!!!
        BotonModificarLinea
     Else   'Modificar Art�culos
        If BLOQUEADesdeFormulario(Me, 1) Then BotonModificar
     End If
End Sub

Private Sub mnNuevo_Click()
     If Modo >= 5 Then          'A�adir lineas Art�culos x Almacen
        BotonAnyadirLinea
    Else 'A�adir Art�culos
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




Private Sub optCRM_Click(index As Integer)
Dim ElTag As Byte
    
    ElTag = CByte(optCRM(index).Tag)
    ImagenCRM ElTag
    CargaColumnasCRM CByte(index)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWCRM
    
End Sub

Private Sub optDoc_Click(index As Integer)
Dim ElTag As Byte
    
    ElTag = CByte(optDoc(index).Tag)
    ImagenDocumento ElTag
    CargaColumnas CByte(index)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWDoc
    
    
End Sub



Private Sub Text1_Change(index As Integer)
    If index = 4 Then HaCambiadoCP = True 'CPostal ha cambiado
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    If index = 4 Then HaCambiadoCP = False
    'If Index <> 22 Then ConseguirFoco Text1(Index), Modo
    If Not EsCampoMemo(index) Then ConseguirFoco Text1(index), Modo
End Sub

Private Function EsCampoMemo(Indice As Integer) As Boolean
    EsCampoMemo = False
    If Indice = 22 Or Indice = 54 Then EsCampoMemo = True
End Function


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If EsCampoMemo(index) And KeyCode = 40 Then 'Flecha abajo
        Me.SSTab1.Tab = 1
        PonerFoco Text1(index)
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    If Not EsCampoMemo(index) Then KEYpress KeyAscii
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
Dim campo As String
Dim Codigo As String
Dim tabla As String
Dim Titulo As String

    If Not PerderFocoGnral(Text1(index), Modo) Then
        If index <> 46 Then Exit Sub    'En modo 2 , el 46 seguimos
    End If
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case index
        Case 0 'Cod Cliente
            If PonerFormatoEntero(Text1(index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(index)) Then PonerFoco Text1(index)
                End If
            End If
            
            
        Case 1
            If Modo = 3 Then
                If Text1(index).Text <> "" Then Text1(2).Text = Text1(index).Text
            End If
            
            
        Case 4 'CPostal
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(index + 1).Text = ObtenerPoblacion(Text1(index).Text, campo)
                Text1(index + 2).Text = campo
             End If
             VieneDeBuscar = False
        
        Case 7 'NIF
            If Text1(index).Text <> "" And Me.chkClienteV.Value = False Then
                Text1(index).Text = UCase(Text1(index).Text)
                ValidarNIF Text1(index).Text
                If Modo = 3 Then
                    If Text1(45).Text = "" Then Text1(45).Text = Text1(index).Text
                    'Veremos si ya existe un cliente con este NIF
                    Codigo = DevuelveDesdeBD(conAri, "concat(codclien,' - ',nomclien)", "sclien", "nifclien", Text1(index).Text, "T")
                    If Codigo <> "" Then MsgBox "Ya existe un cliente con este NIF" & vbCrLf & Codigo, vbExclamation
                    Codigo = ""
                End If
            End If
        
        Case 9 'Codigo de Actividad
            campo = "nomactiv"
            Codigo = "codactiv"
            tabla = "sactiv"
            Titulo = "Actividades"
            
        Case 10 'C�digo de Env�o
            campo = "nomenvio"
            Codigo = "codenvio"
            tabla = "senvio"
            Titulo = "Formas de Env�o"
            
         Case 11 'C�digo de zona
            campo = "nomzonas"
            Codigo = "codzonas"
            tabla = "szonas"
            Titulo = "Zonas de Clientes"
                       
         Case 12 'C�digo de Rutas
             campo = "nomrutas"
             Codigo = "codrutas"
             tabla = "srutas"
             Titulo = "Rutas de Asistencia"

        Case 22 'Observaciones
            If Modo = 3 Or Modo = 4 Then 'Insertando o modificando
                'si se pierde el foco con un TAB y pasaria al siguiente campo que
                'esta en la otra pesta�a. si movemos foco a otro campo de la
                'misma pesta�a no cambiamos
                If Screen.ActiveControl.Name = "Text1" Then
                    If Screen.ActiveControl.index = 23 Then
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
            
        Case 24, 25, 59 'Descuento Pronto Pago, Descuento General  y comision
                'Formato tipo 4: Decimal(4,2)
            If Text1(index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(index), 4
            
        Case 31, 32 'codbanco, sucursal
            PonerFormatoEntero Text1(index)
            
            
        Case 34
            'Si hay valor en la cuenta le calculamos el IBAN
            If Me.Text1(index).Text <> "" Then
                Me.Text1(index).Text = Right(String(10, "0") & Text1(index).Text, 10)
                campo = Text1(31).Text & Me.Text1(32).Text & Me.Text1(33).Text & Me.Text1(34).Text
            
                If Len(campo) = 20 Then
                    DevuelveIBAN2 "ES", campo, campo
                    If Len(campo) = 2 Then
                        campo = "ES" & campo
                        If Me.Text1(56).Text = "" Then
                            Text1(56).Text = campo
                        Else
                            If Me.Text1(56).Text <> campo Then MsgBox "Codigo IBAN distinto del calculado [" & campo & "]", vbExclamation
                        End If
                    End If
                End If
                campo = ""
            End If
        Case 35 'Cuenta contable
            Text2(index).Text = PonerNombreCuenta(Text1(index), Modo, Text1(0).Text)
            
        Case 36, 61 'Codigo Agente Comercial
            campo = "nomagent"
            tabla = "sagent"
            Codigo = "codagent"
            Titulo = "Agente Comercial"
            If index = 51 Then Titulo = "Visitador"
                
            
        Case 37 'Codigo Tarifa
            campo = "nomlista"
            Codigo = "codlista"
            tabla = "starif"
            Titulo = "Tarifa"
                                    
        Case 13, 40, 41, 48, 53, 58, 46 'Fecha alta, Fecha �ltimo mov.,fecha reclamaci�n solicredito
             If Text1(index).Text <> "" Then PonerFormatoFecha Text1(index)
             If index = 46 Then
                If Text1(index).Text <> "" Then
                    imgFecha(3).Tag = Text1(46).Text
                    CargaDatosLWDoc
                Else
                    Text1(46).Text = imgFecha(3).Tag
                End If
             End If
        Case 42 'C�digo Situaci�n
            campo = "nomsitua"
            Codigo = "codsitua"
            tabla = "ssitua"
            Titulo = "Situaci�n"
            
        Case 43, 47, 49, 63 'L�mite Cr�dito , solicitado y riesgo actual ,credito concedido
            'Formato tipo 1: Decimal(12,2)
            If Text1(index).Text <> "" Then
                If Not PonerFormatoDecimal(Text1(index), 1) Then Text1(index).Text = ""
            End If
        Case 44
            '44   Distancia Km
            
'            PonerFormatoDecimal Text1(Index), 5
            PonerFormatoEntero Text1(index)
            
            
        
        Case 52
            If Modo = 1 Then Exit Sub
            'Buscara direcciones envio
            'sdirenvio nomdiren  coddiren
            campo = "nomdiren"
            tabla = "sdirenvio"
            Codigo = "codclien = " & Val(Text1(0).Text) & " AND coddiren "
            Titulo = "Direccion envio"
        
        
        Case 55
            'En fenollar
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                lblPerfil.Caption = ""
                If Text1(55).Text <> "" Then
                    lblPerfil.Caption = PonerNombreDeCod(Text1(index), conAri, "stipperfil", "titulo", "perfil", "Perfil riesgo")
                    If lblPerfil.Caption = "" Then
                        campo = " GROUP_CONCAT( concat(perfil,':    ',titulo) separator '\n' )"
                        campo = DevuelveDesdeBD(conAri, campo, "stipperfil", "1", "1")
                        If campo <> "" Then
                            campo = "Perfiles admitidos: " & vbCrLf & vbCrLf & campo
                            MsgBox campo, vbExclamation
                        End If
                        Text1(55).Text = ""
                    End If
                End If
            End If
    End Select
    
    If (index >= 9 And index <= 12) Or index = 23 Or index = 36 Or index = 37 Or index = 42 Or index = 52 Or index = 61 Then
        If PonerFormatoEntero(Text1(index)) Then
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, tabla, campo, Codigo, Titulo)
            If Text2(index).Text = "" Then
                PonerFoco Text1(index)
                If index = 52 Then Text1(index).Text = ""
            End If
            
        Else
            Text2(index).Text = ""
        End If
        
        If index = 42 Then txtSit.Text = Text2(index).Text
        If index = 36 Then
            If Modo = 3 And Text2(index).Text <> "" And Text1(61).Text = "" Then Text1(61).Text = Text1(index).Text
        End If
    End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadB2 As String

    If vParamAplic.TieneTelefonia2 > 0 Then
        'Permito hacer busquedas por telefonia
        cadB2 = DevuelveBusquedaTelefonia
    Else
        cadB2 = ""
    End If
    
    If vParamAplic.ContabilidadNueva Then Text1(60).Text = PaisSeleccionado
    
    
    
    
    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
        
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then
            If CadB <> "" Then CadB = CadB & " AND "
            CadB = CadB & " codagent = " & vUsu.CodigoAgente
        End If
    End If
    
    If cadB2 <> "" Then
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & " codclien IN (Select codclien from sclientfno WHERE " & cadB2 & ")"
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
            cad = cad & "C�digo|cuentas|codmacta|T||30�Denominacion|cuentas|nommacta|T||70�"
            tabla = "cuentas"
            Titulo = "Cuentas Contables"
            Conexion = conConta    'Conexi�n a BD: Conta
            
            
        Case 1000, 1001
            'Departamento en RENTING  Marzo 2012      1001: En telefono: Mar13
            cad = cad & "C�digo|sdirec|coddirec|N||30�Denominacion|sdirec|nomdirec|T||70�"
            tabla = "sdirec"
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Departamentos"
            Else
                Titulo = "Direccion"
            End If
            Conexion = conAri    'Conexi�n a BD: Ariges
        
        Case 1003
            cad = cad & "C�digo|stfnoModel|codmodelo|N||30�Descripcion|stfnoModel|descripcion|T||70�"
            Titulo = "Modelo de telefono"
            tabla = "stfnoModel"
            Conexion = conAri    'Conexi�n a BD: Ariges
        
        Case 1004
            cad = cad & "C�digo|sartic|codartic|T||30�Descripcion|sartic|nomartic|T||70�"
            Titulo = "Art. telefonia VTA PLAZOS"
            tabla = "sartic"
            Conexion = conAri    'Conexi�n a BD: Ariges
        Case Else   'Registro de la tabla de cabeceras: sartic
    
            cad = cad & ParaGrid(Text1(0), 10, "C�digo")
            cad = cad & ParaGrid(Text1(1), 50, "Nombre")
            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
            tabla = "sclien"
            Titulo = "Clientes"
            Conexion = conAri    'Conexi?n a BD: Ariges
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
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
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


Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(9).Text = PonerNombreDeCod(Text1(9), conAri, "sactiv", "nomactiv")
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "senvio", "nomenvio")
    Text2(11).Text = PonerNombreDeCod(Text1(11), conAri, "szonas", "nomzonas")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "srutas", "nomrutas")
    Text2(23).Text = PonerNombreDeCod(Text1(23), conAri, "sforpa", "nomforpa")
    Text2(35).Text = PonerNombreDeCod(Text1(35), conConta, "cuentas", "nommacta")
    Text2(36).Text = PonerNombreDeCod(Text1(36), conAri, "sagent", "nomagent")
    Text2(37).Text = PonerNombreDeCod(Text1(37), conAri, "starif", "nomlista", "codlista")
    Text2(42).Text = PonerNombreDeCod(Text1(42), conAri, "ssitua", "nomsitua")
    txtSit.Text = Text2(42).Text
    Text2(61).Text = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", Text1(61).Text)
    
    
    If vParamAplic.DireccionesEnvio Then Text2(52).Text = PonerNombreDeCod(Text1(52), conAri, "sdirenvio", "nomdiren", "codclien = " & Text1(0).Text & " AND coddiren")
    
    'If vParamAplic.ContabilidadNueva Then PonerPais
    PonerPais
    
    Me.lblPerfil.Caption = ""
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        If Text1(55).Text <> "" Then lblPerfil.Caption = DevuelveDesdeBD(conAri, "titulo", "stipperfil", "perfil", Text1(55).Text, "T")
    End If
    
    MostrarSituacion True
    
    BloquearChecks Me, Modo
    
    lblIndicador.Caption = "Clientes aux"
    lblIndicador.Refresh
    CargaLineas True, 8
    BotonesToolBarAux
    
    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    CargaDatosLWDoc
    If vParamAplic.TieneCRM Then CargaDatosLWCRM
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diversos campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim DespVis As Boolean
Dim b As Boolean
Dim N As Integer

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar. Eto era para una barra de tareas para todo
    'ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    BuscaChekc = ""
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    If Modo = 2 Then Indicador_
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        Me.cmdRegresar.Caption = "Regresar"
    Else
        cmdRegresar.visible = False
    End If
        
    FrameNavegaDoc.Enabled = b Or Modo = 0
    FrameNavegaCRM.Enabled = b Or Modo = 0
    If vParamAplic.TieneCRM Then FrameNavegaCRM.Enabled = b Or Modo = 0
    
    
     'Poner Flechas de desplazamiento visibles
     
    DespVis = False
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then DespVis = True
        End If
    End If
    DespalzamientoVisible DespVis
    
         
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'El campo 46 NUNCA se puede escribir en el
    If Text1(46).Text = "" Then Text1(46).Text = Me.imgFecha(3).Tag
    BloquearTxt Text1(46), Modo <> 2
    'la fecha utlimo recalcuo de riesgo tp se escribe
    Text1(51).Enabled = False
    
    If Modo = 2 Then Text1(43).BackColor = &HFFFF80
    
    
    'Bloquear los Text3
    For I = 0 To Me.Text3.Count - 1
        BloquearTxt Me.Text3(I), Not (Modo = 5)
    Next I
        
    'Bloquear los Text3
    If vParamAplic.DireccionesEnvio Then
        For I = 0 To Me.Text4.Count - 1
            BloquearTxt Me.Text4(I), Not (Modo = 6)
        Next I
        
        
        'Si tiene direcciones de envio y el modo=4 entonces esta habilitado
        BloquearTxt Me.Text1(52), Not (Modo = 1 Or Modo = 4)
        
    End If
    'Bloquear los Text3
    If Modo <> 7 Then
        For I = 0 To Me.txtauxDC.Count - 1
            BloquearTxt Me.txtauxDC(I), True
        Next I
    End If
    
    'Campos telefonia
    If vParamAplic.TieneTelefonia2 > 0 Then
        b = Modo = 1

        
        FrameTelefonia(1).Enabled = Modo = 2 Or Modo = 9
        
        FrameTelefonia(0).visible = Not (Modo = 3 Or Modo = 4)  'Insertando o modifiando NO puede estar visible el frame
        Me.cboOperadorTfnnia2(0).Enabled = b
        Me.cboOperadorTfnnia2(1).Enabled = b
        'FrameTelefonia(1).Enabled = Modo = 2 Or Modo = 4
        N = IIf(vParamAplic.TelefoniaVtaPlazos, 14, 10)
        For I = 0 To N
            BloquearTxt Me.txtauxTfno(I), Not b
            If I < 3 Then
                Me.txtauxTfno(I).visible = Modo = 1
                If I = 0 Then Me.cboOperadorTfnnia2(0).visible = Modo = 1
            End If
        Next
        
        If Modo <> 9 Then
            FrameTelefonia(0).Enabled = False
            For I = 2 To 4
             '   Me.cmdAccionesTfno(I).Enabled = Modo = 2
            Next
        Else
            FrameTelefonia(0).Enabled = True
        End If
        
        If Modo <> 1 And Modo <> 9 Then Me.cboOperadorTfnnia2(0).visible = False
    End If
    
    
        
    '---------------------------------------------
    'b = Modo <> 0 And Modo <> 2 And Modo <> 5
    b = Modo = 1 Or Modo = 3 Or Modo = 4
    cboAlbaran.Enabled = b
    cboFacturacion.Enabled = b
    cboTipoIVA.Enabled = b
    cboTipocliente.Enabled = b
    cboPrioridad.Enabled = b
    If vParamAplic.Renting Then cboFraRenting.Enabled = b
    If vParamAplic.ManipuladorFitosanitarios2 Then cboManipulador.Enabled = b
    cboTipoASeg.Enabled = b
    cboPais.Enabled = b
    
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    'Permisos
    I = 0
    If vParamAplic.OperacionesAseguradas And vUsu.Nivel = 0 Then I = 1
    Me.FrameAsegurados.Enabled = I = 1
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    For I = 0 To Me.imgFecha.Count - 1
        If I <> 3 Then Me.imgFecha(I).Enabled = b
    Next I
    
    
    If vParamAplic.PtosAsignar > 0 Then
        'LLEVa puntos
        If Modo = 3 Or Modo = 4 Then
            BloquearTxt Text1(62), vUsu.Login <> "root"
            Me.chkPuntos.Enabled = vUsu.Nivel = 0
        Else
            Me.chkPuntos.Enabled = True
        End If
        
    End If
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        'el 15 y 16 son de zona en direc y envio
        If I = 15 Or I = 16 Then
            Me.imgBuscar(I).Enabled = False
        Else
            Me.imgBuscar(I).Enabled = b
        End If
    Next I
    imgBuscar(11).Enabled = Modo >= 2 And Modo < 5
    imgBuscar(17).Enabled = imgBuscar(11).Enabled
    
    
    If Modo = 2 Or Modo = 9 Then imgBuscar(21).Enabled = True
    'CRM
    cmdAccCRM(0).visible = vParamAplic.TieneCRM And Modo = 2
    cmdAccCRM(1).visible = vParamAplic.TieneCRM And Modo = 2
    
    
    '-----------------------------
    If vParamAplic.OperacionesAseguradas Then cmdActRiesgo.visible = Modo = 2 And vUsu.Nivel = 0

    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opcines de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
                        
   BotonesToolBarAux
                        
    'El listview
    If Modo <> 2 Then
        lw1.ListItems.Clear
        If vParamAplic.TieneCRM Then lwCRM.ListItems.Clear
    End If

                        
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.Nivel > 0 And Modo = 4 Then
            imgBuscar(8).Enabled = False
            BloquearTxt Text1(42), True
        End If
    End If
                        
    cmdCatalogo.visible = False
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean



    b = (Modo = 2 Or Modo = 0)
    Toolbar1.Buttons(5).Enabled = b
    Toolbar1.Buttons(6).Enabled = b
    
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    Else
        If vUsu.Nivel2 > 1 Then b = False
    End If
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    
    Toolbar1.Buttons(8).Enabled = True   'IMprimir
    
    b = (Modo = 2)
    
    'Los que sean AGENTES no pueden entrar
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    Else
        If vUsu.Nivel2 > 1 Then b = False
    End If
    Toolbar1.Buttons(2).Enabled = b  'modificar
    Toolbar1.Buttons(3).Enabled = b     'eliminar
    
    
           
           
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoFrame(Kmodo As Byte, ModoGral As Byte)
Dim I As Byte
On Error GoTo EPonerModoFr

    ModoFrame2 = Kmodo
    PonerModo ModoGral
    
    'Bloquear TextBox sino modo 3 o 4
    Select Case ModoGral
    Case 5
        For I = 0 To Me.Text3.Count - 1
            If ModoFrame2 = 3 Then Text3(I).Text = ""
            BloquearTxt Text3(I), (ModoFrame2 = 0)
        Next I
        
        
        If ModoFrame2 = 4 Then BloquearTxt Text3(0), True
        
        imgBuscar(15).Enabled = ModoFrame2 > 0
    Case 6
        'direnvio
        For I = 0 To Me.Text4.Count - 1
            If ModoFrame2 = 3 Then Text4(I).Text = ""
            BloquearTxt Text4(I), (ModoFrame2 = 0)
        Next I
        If ModoFrame2 = 4 Then BloquearTxt Text4(0), True
        imgBuscar(16).Enabled = ModoFrame2 > 0
        
    Case 7
        'Perosna de contacto
        For I = 0 To Me.txtauxDC.Count - 1
            If ModoFrame2 = 3 Then txtauxDC(I).Text = ""
            BloquearTxt txtauxDC(I), (ModoFrame2 = 0)
        Next I
       
       
       imgBuscar(14).visible = ModoFrame2 > 0
        Me.cboCargo.visible = ModoFrame2 > 0
        imgBuscar(14).Enabled = ModoFrame2 > 0
   
     Case 8
        'renting
        For I = 0 To Me.txtauxRent.Count - 1
            If ModoFrame2 = 3 Then txtauxRent(I).Text = ""
            'Campos SIEMPRE BLOQUEADOS
            If I = 0 Or I = 2 Then
                BloquearTxt txtauxRent(I), True
            Else
                BloquearTxt txtauxRent(I), (ModoFrame2 = 0)
            End If
        Next I
       
         
       cmdRenting(0).visible = ModoFrame2 > 0
       cmdRenting(1).visible = ModoFrame2 > 0
       cmdRenting(2).visible = ModoFrame2 > 0
       Me.DataGrid2.Enabled = ModoFrame2 = 0
    Case 9
        'Telefonia
        For I = 0 To Me.txtauxTfno.Count - 1
            If ModoFrame2 = 3 Then
                txtauxTfno(I).Text = ""
                If I < 4 Then Me.chkTelefonia(I).Value = 0
                If I > 3 And I < 7 Then Text5(I).Text = ""
            End If
            
            
            BloquearTxt txtauxTfno(I), (ModoFrame2 = 0)
            
        Next I
        If ModoFrame2 = 3 Then
            Me.cboOperadorTfnnia2(0).ListIndex = -1
            Me.cboOperadorTfnnia2(1).ListIndex = -1
        End If
        Me.cboOperadorTfnnia2(0).Enabled = ModoFrame2 <> 0
        Me.cboOperadorTfnnia2(1).Enabled = Me.cboOperadorTfnnia2(0).Enabled
        Me.DataGrid3.Enabled = ModoFrame2 = 0
        Me.FrameTelefonia(0).Enabled = ModoFrame2 <> 0
        
        For I = 2 To 4
            Me.cmdAccionesTfno(I).visible = Modo = 2  'ModoFrame2 = 0
        Next
        
        For I = 18 To 20
            Me.imgBuscar(I).Enabled = ModoFrame2 > 2
        Next
        Me.imgBuscar(23).Enabled = vParamAplic.TelefoniaVtaPlazos And ModoFrame2 > 2
    Case 10

        'Fitosanitarios
        For I = 0 To Me.txtauxFito.Count - 1
            If ModoFrame2 = 3 Then txtauxFito(I).Text = ""
            'Campos SIEMPRE BLOQUEADOS
            If I = 4 Then
                BloquearTxt txtauxFito(I), True
            Else
                BloquearTxt txtauxFito(I), (ModoFrame2 = 0)
            End If
        Next I
        If ModoFrame2 = 3 Then
            Me.cboFitos(0).ListIndex = -1
            Me.cboFitos(1).ListIndex = -1
        End If
         
      
       Me.DataGrid4.Enabled = ModoFrame2 = 0

    Case 11
        
        'Campos / huertos
        '-------------------
         
        For I = 0 To Me.txtauxMarja.Count - 1
            If ModoFrame2 = 3 Then
                txtauxMarja(I).Text = ""
                
            End If
            
            
            BloquearTxt txtauxMarja(I), (ModoFrame2 = 0)
            
        Next I
        Me.DataGrid5.Enabled = ModoFrame2 = 0
        
        For I = 7 To 9
            Me.imgFechaCampos(I).Enabled = ModoFrame2 > 2
        Next
    End Select
    
    
    I = 10
    If ModoGral = 6 Then I = 12
    Select Case ModoFrame2
        Case 0  'MODO INICIAL
            Me.imgBuscar(I).Enabled = False
            
        Case 3, 4 'Modo INSERTAR o MODIFICAR
            '3=Insertar,  4=Modificar
            Me.imgBuscar(I).Enabled = True
            If Modo = 3 Then
                If ModoGral = 5 Then
                    PonerFoco Text3(0)
                Else
                    PonerFoco Text4(0)
                End If
            End If
            
    End Select

    BotonesToolBarAux
    Me.cmdCancelar.visible = Kmodo = 3 Or Kmodo = 4
    Me.cmdAceptar.visible = cmdCancelar.visible
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
'    Me.Label1(6).visible = Not bol 'Cod. Env�o
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
    
    
    If vParamAplic.NumeroInstalacion <> vbFontenas Then
        If cboPrioridad.ListIndex < 0 Then
            If cboPrioridad.ListCount > 2 Then
                cboPrioridad.ListIndex = 3
            Else
                cboPrioridad.ListIndex = 0
            End If
        End If
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
       
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
        
        
        If b And vParamAplic.NumeroInstalacion = vbFenollar Then
            If Val(Mid(Text1(35).Text, 3)) <> Val(Text1(0).Text) Then
                If MsgBox("No coincide Id cliente con valor CUENTA CONTABLE.  �Continuar de igual modo?", vbQuestion + vbYesNoCancel) <> vbYes Then b = False
                
        
            End If
        End If
        
    End If
    If Not b Then Exit Function
    
    
                    
    '- Validar que la cuenta bancaria es correcta
    If Comprueba_CuentaBan2(Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text, False) Then
            CadenaConsulta = Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text
            If Len(CadenaConsulta) = 20 Then
                
                BuscaChekc = ""
                If Me.Text1(56).Text <> "" Then BuscaChekc = Mid(Text1(56).Text, 1, 2)
                
                    
                If DevuelveIBAN2(BuscaChekc, CadenaConsulta, CadenaConsulta) Then
                    If Me.Text1(56).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(56).Text = BuscaChekc & CadenaConsulta
                    Else
                        If Mid(Text1(56).Text, 3) <> CadenaConsulta Then
                            CadenaConsulta = "Calculado : " & BuscaChekc & CadenaConsulta
                            CadenaConsulta = "Introducido: " & Me.Text1(56).Text & vbCrLf & CadenaConsulta & vbCrLf
                            CadenaConsulta = "Error en codigo IBAN" & vbCrLf & CadenaConsulta & "Continuar?"
                            If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then Exit Function
                        End If
                    End If
                End If
                        
            End If
            CadenaConsulta = ""
            BuscaChekc = ""
    End If
    



    '- comprobar q dia de vto atrasado tiene valor solo si mes a no girar tiene valor
    If Trim(Text1(26).Text) = "" And Trim(Text1(27).Text) <> "" Then
        b = False
        MsgBox "El d�a de Vto. atrasado solo debe tener valor si hay mes a no girar.", vbInformation
    ElseIf Trim(Text1(26).Text) <> "" And Trim(Text1(27).Text) <> "" Then
        If Trim(Text1(28).Text) <> "" Or Trim(Text1(29).Text) <> "" Or Trim(Text1(30).Text) <> "" Then
            b = False
            MsgBox "Si hay dias de pago no puede haber d�a de vto. atrasado.", vbInformation
        Else
            'comprobar q el dia de vto atrasado introducido existe para
            'el mes siguiente al mes a no girar
              If CInt(Text1(26).Text) + 1 < 13 Then
                If Not IsDate(Text1(27).Text & "/" & CInt(Text1(26).Text) + 1 & "/" & Year(Now)) Then
                    b = False
                    MsgBox "La fecha del dia de vto atrasado para el mes " & CInt(Text1(26).Text) + 1 & " NO es valida.", vbInformation
                End If
              Else
                If Not IsDate(Text1(27).Text & "/1/" & Year(Now) + 1) Then
                    b = False
                    MsgBox "La fecha del dia de vto atrasado para el mes 1" & " NO es valida.", vbInformation
                End If
              End If
        End If
    End If

    'QUito esto   11 Enero 09
    'Text1(22).Text = QuitarCaracterEnter(Text1(22))
    
    
    'Operaciones aseguradas. Si tiene fecha concesion pondre el riesgo, de momento a cero
    If b Then
        If Me.Text1(41).Text <> "" Then
            BuscaChekc = ""
            'Si el valor del limite de credito es nulo o cero aviso
            If Text1(43).Text = "" Then
                BuscaChekc = "N"
            Else
                If ImporteFormateado(Text1(43).Text) = 0 Then BuscaChekc = "N"
            End If
                
                
                
 
                
                
            If BuscaChekc <> "" Then
                If MsgBox("Ha puesto fecha concesi�n y no indica el l�mite concedido" & vbCrLf & "   �Continuar?", vbQuestion + vbYesNo) = vbNo Then b = False
                BuscaChekc = ""
            End If
                    
            If Text1(49).Text = "" Then Text1(49).Text = "0"
        End If
    
    
    
    
        'Operaciones aseguradas FENOLLAR
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            If Text1(55).Text <> "" Then
                BuscaChekc = "|" & Text1(55).Text & "|"
                If InStr(1, "|NORM|30|90|60|NADA|OP|180|120|150|PART|", BuscaChekc) = 0 Then
                    MsgBox "Tipo de perfil incorrecto" & vbCrLf & " 30 - 60  - 90 - 120 - 150 - 180 - NADA - NORM - OP - PART", vbExclamation
                    b = False
                    PonerFoco Text1(55)
                End If
            
                If b Then
                    If Text1(64).Text = "" Then
                        'MsgBox "Debe indicar tipo de credito asegurado"
                        'B = False
                        'PonerFoco Text1(64)
                    End If
                End If
            End If
            BuscaChekc = ""
            
        End If
    
    
    End If
    
    If b And vParamAplic.ManipuladorFitosanitarios2 Then
        If Me.cboManipulador.ListIndex > 0 Then
            BuscaChekc = ""
            
            If Me.Text1(58).Text = "" Then BuscaChekc = "Introduzca la fecha de caducidad del carnet de fitosanitarios" & vbCrLf
            If Me.Text1(57).Text = "" Then BuscaChekc = "Introduzca el numero de carnet fitosanitarios" & vbCrLf & BuscaChekc
            
            If BuscaChekc <> "" Then
                MsgBox BuscaChekc, vbExclamation
                b = False
         
            End If
            
            
        End If
    End If
    
    'Si lleva aseguradas
    If b And vParamAplic.OperacionesAseguradas And vUsu.Nivel = 0 Then
        BuscaChekc = ""
        If Me.cboTipoASeg.ItemData(cboTipoASeg.ListIndex) = 9 Then
            If Me.Text1(43).Text <> "" Then
                If ImporteFormateado(Text1(43).Text) > 0 Then
                    BuscaChekc = "No debe poner limite de cr�dito"
                Else
                    Text1(43).Text = ""
                End If
            End If
        Else
            If Me.Text1(43).Text = "" Then
                BuscaChekc = "Debe poner limite de cr�dito"
            Else
                If ImporteFormateado(Text1(43).Text) = 0 Then BuscaChekc = "Debe poner limite de cr�dito"
            End If
        End If
        If Me.chkClienteV.Value = 1 And Me.cboTipoASeg.ItemData(cboTipoASeg.ListIndex) <> 9 Then BuscaChekc = BuscaChekc & vbCrLf & "NO puede asegurar clientes varios"
        If BuscaChekc <> "" Then
            MsgBox BuscaChekc, vbExclamation
            b = False
        End If
    End If
    
    If b Then
        BuscaChekc = ""
        If Modo = 3 Then
            BuscaChekc = Text1(0).Text
        Else
            If Modo = 4 Then
                'Si ha cambiado el NIF
                If Data1.Recordset!nifClien <> Text1(7).Text Then BuscaChekc = Text1(0).Text
            End If
        End If
        
        If BuscaChekc <> "" Then
            BuscaChekc = DevuelveDesdeBD(conAri, "concat(codclien,' - ',nomclien)", "sclien", "nifclien", Text1(7).Text, "T")
            
            If BuscaChekc <> "" Then
                BuscaChekc = "Ya existe un cliente con este NIF:" & vbCrLf & vbCrLf & Text1(7).Text & "   " & BuscaChekc & vbCrLf & "�Continuar?"
                If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then b = False
                BuscaChekc = ""
            End If
        End If
    End If
    
    If b And vParamAplic.ContabilidadNueva Then Me.Text1(60).Text = PaisSeleccionado
        
 
        
        
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function DatosOkLinea() As Boolean
    DatosOkLinea = False
    Select Case Modo
    Case 5
        DatosOkLinea = DatosOkLineaDpto
    Case 6
        DatosOkLinea = DatosOkLineaEnvio
    Case 7
    
       
        
        'En el text2 opongo el combo
        txtauxDC(2).Text = cboCargo.Text
        'Para datos personales SOLO necesito el nombre
        If Trim(txtauxDC(0).Text) = "" Then
            MsgBox "Nombre obligatorio", vbExclamation
        Else
            DatosOkLinea = True
        End If
        
    Case 8
        'renting
         'desde el 2
        For NumRegElim = 3 To Me.txtauxRent.Count - 1
            If NumRegElim <> 10 And NumRegElim <> 11 Then '7= ult fecha factura
                If Me.txtauxRent(NumRegElim).Text = "" Then
                        MsgBox "Campos obligatorios", vbExclamation
                        PonerFoco txtauxRent(NumRegElim)
                        Exit Function
                End If
            End If
        Next
        'Si pone coddirec, tiene que existir nomdirec
        If Me.txtauxRent(1).Text = "" Xor txtauxRent(2).Text = "" Then
            MsgBox "Error departamento/direccion", vbExclamation
            Exit Function
        End If

        'Comprobaremos que la linea que ha puesto no es mayor que uno ya facturado
        BuscaChekc = DevuelveDesdeBD(conAri, "max(ultfec)", "sclienrenting", "codclien", CStr(Data1.Recordset!codClien))
        If BuscaChekc <> "" Then
            If CDate(txtauxRent(4).Text) >= CDate(BuscaChekc) Then
                If MsgBox("Peridodo no facturado.No se generara factura. �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
            BuscaChekc = ""
            
        End If
        
        
        
        DatosOkLinea = True
        
    Case 9
        'Solo obligamos al TFNO
        
        If Trim(txtauxTfno(0).Text) = "" Then
            MsgBox "Telefono es obligatorio", vbExclamation
        Else
            BuscaChekc = ""
            If Not IsNumeric(txtauxTfno(0).Text) Then BuscaChekc = BuscaChekc & "-No es num�rico" & vbCrLf
            If Len(txtauxTfno(0).Text) <> 9 Then BuscaChekc = BuscaChekc & "-Longitud distinta de 9" & vbCrLf
            If BuscaChekc <> "" Then
                    BuscaChekc = "Error en campo N�mero de tel�fono. " & vbCrLf & vbCrLf & BuscaChekc & vbCrLf & "�Continuar?"
                    If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then BuscaChekc = ""
            End If
            If BuscaChekc = "" Then
                'Es clave UNICA el telefono
                BuscaChekc = "sclientfno left join sclien on  sclientfno.codclien=sclien.codclien"
                BuscaChekc = DevuelveDesdeBD(conAri, "concat(sclientfno.codclien,' - ',nomclien)", BuscaChekc, "sclientfno.codclien<>" & Text1(0).Text & " AND IdTelefono", txtauxTfno(0).Text, "T")
                If BuscaChekc <> "" Then
                    MsgBox "El tel�fono ya pertenece al cliente: " & BuscaChekc, vbExclamation
                Else
                    If cboOperadorTfnnia2(0).ListIndex < 0 Then
                        MsgBox "Seleccione un operador de telefon�a", vbExclamation
                    Else
                        If txtauxTfno(7).Text = "" Then txtauxTfno(7).Text = 0
                    
                        If txtauxTfno(1).Text = "" Or txtauxTfno(2).Text = "" Or txtauxTfno(7).Text = "" Or txtauxTfno(8).Text = "" Or txtauxTfno(9).Text = "" Then
                            MsgBox "Campos     SIM/IMEI/PUNTOS/Cuota minima/Fecha alta    obligatorios", vbExclamation
                        Else
                            If cboOperadorTfnnia2(1).ListIndex < 0 Then
                                MsgBox "Seleccione procedencia", vbExclamation
                            Else
                                DatosOkLinea = True
                            End If
                        End If
                    End If
                End If
            End If
            
            If DatosOkLinea Then
                'Esta yendo bien
                'Si lleva venta plazos
                If vParamAplic.TelefoniaVtaPlazos Then
                    BuscaChekc = ""
                    If txtauxTfno(11).Text = "" Xor txtauxTfno(12).Text = "" Then BuscaChekc = "N"
                    If txtauxTfno(11).Text = "" Xor txtauxTfno(13).Text = "" Then BuscaChekc = "N"
                    If txtauxTfno(13).Text = "" Xor txtauxTfno(14).Text = "" Then BuscaChekc = "N"
                    If BuscaChekc <> "" Then
                        MsgBox "Si indica venta a plazo debe indicar los Articulo / Meses /importe", vbExclamation
                        DatosOkLinea = False
                    End If
                End If
            End If
        End If
        
    Case 10
        
        BuscaChekc = ""
        kCampo = -1
        If Me.cboFitos(0).ListIndex < 0 Then BuscaChekc = BuscaChekc & " - Tipo carnet" & vbCrLf
        For NumRegElim = 0 To Me.txtauxFito.Count - 1
            If NumRegElim <> 2 And NumRegElim <> 3 Then
                If Me.txtauxFito(NumRegElim).Text = "" Then
                        BuscaChekc = BuscaChekc & " - " & RecuperaValor("DNI|Nombre||||Caducidad|", NumRegElim + 1) & vbCrLf
                        If kCampo < 0 Then kCampo = NumRegElim
                End If
            End If
        Next
        If BuscaChekc <> "" Then
            BuscaChekc = "Campos obligatorios: " & vbCrLf & BuscaChekc
            MsgBox BuscaChekc, vbExclamation
            If kCampo >= 0 Then PonerFoco txtauxFito(kCampo)
        Else
            DatosOkLinea = True
        End If

    Case 11
        BuscaChekc = ""
        kCampo = 0
        For NumRegElim = 0 To 7
            If NumRegElim <> 6 Then
                If Trim(txtauxMarja(NumRegElim).Text) = "" Then
                   BuscaChekc = BuscaChekc & " .-" & DataGrid5.Columns(NumRegElim).Caption & vbCrLf
                   kCampo = NumRegElim
                End If
            End If
        Next

        If BuscaChekc <> "" Then
            MsgBox "Campos obligatorios: " & vbCrLf & BuscaChekc, vbExclamation
            PonerFoco txtauxMarja(kCampo)
        Else
            DatosOkLinea = True
            
            Me.txtauxMarja(6).Text = cbomarjal.Text
            
        End If
    End Select
End Function

Private Function DatosOkLineaDpto() As Boolean
Dim b As Boolean
Dim devuelve As String
Dim I As Integer

On Error GoTo EDatosOkLinea

    DatosOkLineaDpto = False
    b = True
    devuelve = ""
    'Campo Nombre Direc./Dpto
    If Text3(1).Text = "" Then devuelve = devuelve & vbCrLf & "-Nombre"
    
    'Campo Domicilio Direc./Dpto
    If Text3(2).Text = "" Then devuelve = devuelve & vbCrLf & "-Domicilio"

    'Campo CPostal Direc./Dpto
    If Text3(3).Text = "" Then devuelve = devuelve & vbCrLf & "-C.Postal"
    
    'Campo Poblaci�n Direc./Dpto
    If Text3(4).Text = "" Then devuelve = devuelve & vbCrLf & "-Poblaci�n"

    'Campo Provincia Direc./Dpto
    If Text3(5).Text = "" Then devuelve = devuelve & vbCrLf & "-Provincia"
        
    'Campo ZONA
    If Text3(14).Text = "" Then devuelve = devuelve & vbCrLf & "-ZONA "
    
    If devuelve <> "" Then
        devuelve = "Campos vacios: " & vbCrLf & devuelve
        MsgBox devuelve, vbExclamation
        devuelve = ""
        Exit Function
    End If
    
   
    
    'Comprobamos  si ya existe Si estamos insertando
    'conAri: conexion a BD Ariges
    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "coddirec", "codclien", Text1(0).Text, "N", , "coddirec", Text3(0).Text, "N")
    'If ModificaLineas = 1 And DevuelveExisteEnBD(conAri, "sdirec", "codclien", Text1(0).Text, "N", "coddirec", Text3(0).Text, "N") Then
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = DevuelveTextoDepto(False)
        devuelve = "Ya existe" & devuelve & " del Cliente: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        MsgBox devuelve, vbExclamation
    End If
    
    
    'comprobar los datos de la cuenta bancaria si param. de departamentos
    If Me.FrameCtaBanDpto.visible And b Then
        'Validar que la cuenta bancaria es correcta
        For I = 10 To 13
            If Text3(I).Text <> "" Then
                If IsNumeric(Text3(I).Text) Then
                    If Val(Text3(I).Text) = "0" Then Text3(I).Text = ""
                End If
            End If
        Next
        
        
        If Text3(13).Text <> "" Then
            'Ha puesto codbanco
          
                For I = 11 To 13
                    If Text3(I).Text = "" Then Exit For
                Next
                If I <= 13 Then
                    'Se ha salido
                    MsgBox "Faltan datos para la cuenta bancaria", vbExclamation
                    b = False
                Else
                    b = Comprueba_CuentaBan2(Text3(10).Text & Text3(11).Text & Text3(12).Text & Text3(13).Text, False)
                    If Not b Then
                        If MsgBox("Cuenta bancaria incorrecta.   �Continuar?", vbQuestion + vbYesNo) = vbYes Then b = True
                    End If
                End If
        End If
        
        
 
        
    End If
    
    
    
    
    
    
    DatosOkLineaDpto = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLineaEnvio() As Boolean
Dim devuelve As String
Dim I As Integer
On Error GoTo EDatosOkLinea

    DatosOkLineaEnvio = False
    
    devuelve = ""
    
    For I = 1 To 10
        Text4(I).Text = Trim(Text4(I).Text)
        If I < 6 Or I > 8 Then
            If Text4(I).Text = "" Then
                devuelve = devuelve & "     -" & RecuperaValor(Text4(I).Tag, 1)
            End If
        End If
    Next
    If devuelve <> "" Then
        MsgBox "Campos no pueden estar vacios: " & vbCrLf & devuelve, vbExclamation
        Exit Function
    End If
    
    'Comprobamos  si ya existe Si estamos insertando
    'conAri: conexion a BD Ariges
    devuelve = DevuelveDesdeBDNew(conAri, "sdirenvio", "coddiren", "codclien", Text1(0).Text, "N", , "coddiren", Text4(0).Text, "N")
    If ModificaLineas = 1 And devuelve <> "" Then
        devuelve = "Ya existe la direccion de envio del Cliente: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text4(0).Text & vbCrLf
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
     
    
    DatosOkLineaEnvio = True
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text3_Change(index As Integer)
    If index = 3 Then HaCambiadoCP = True
    If index = 2 Then Text3(index).ToolTipText = Text3(index).Text
End Sub

Private Sub Text3_GotFocus(index As Integer)
    If index = 3 Then HaCambiadoCP = False
    ConseguirFoco Text3(index), 3
End Sub

Private Sub Text3_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        If (index = 9 And Me.FrameCtaBanDpto.visible = False) Or index = 19 Then
            PonerFocoBtn Me.cmdAceptar
        Else
            SendKeys "{tab}"
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text3_LostFocus(index As Integer)
Dim devuelve As String

    On Error Resume Next
    
    If Not PerderFocoGnralLineas(Text3(index), ModificaLineas) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case index
        Case 0 'Codigo Direc/Dpto
            If Trim(Text3(index).Text) = "" Then Exit Sub
            FormateaCampo Text3(index)

        Case 3 'Cod. Postal
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text3(index + 1).Text = ObtenerPoblacion(Text3(index).Text, devuelve)
                Text3(index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 10, 11 'codbanco, sucursal
            PonerFormatoEntero Text3(index)
            
        Case 12, 13 'DC, cta banco
            FormateaCampo Text3(index)
            If index = 13 Then
                devuelve = Me.Text3(10).Text & Text3(11).Text & Text3(12).Text & Text3(13).Text
                
                If Len(devuelve) = 20 Then
                    DevuelveIBAN2 "ES", devuelve, devuelve
                    If Len(devuelve) = 2 Then
                        devuelve = "ES" & devuelve
                        If Me.Text3(15).Text = "" Then
                            Text3(15).Text = devuelve
                        Else
                            If Me.Text3(15).Text <> devuelve Then MsgBox "Codigo IBAN distinto del calculado [" & devuelve & "]", vbExclamation
                        End If
                    End If
                End If
                
            End If
            
        Case 14
            devuelve = ""
            If PonerFormatoEntero(Text3(index)) Then
                devuelve = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text3(index).Text, "N")
                If devuelve = "" Then
                    MsgBox "No existe la zona", vbExclamation
                    Text3(index).Text = ""
                    PonerFoco Text3(index)
                End If
            Else
                Text3(index).Text = ""
            End If
            Me.txtZona(index).Text = devuelve
        Case 19
            PonerFocoBtn Me.cmdAceptar
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub



'Text4    Direnvio
Private Sub Text4_Change(index As Integer)
    If index = 3 Then HaCambiadoCP = True
End Sub

Private Sub Text4_GotFocus(index As Integer)
    If index = 3 Then HaCambiadoCP = False
    ConseguirFoco Text4(index), 3
End Sub

Private Sub Text4_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text4_KeyPress(index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then 'ENTER
        
        If index <> 9 Then
            KeyAscii = 0
            SendKeys "{tab}"
        Else
            PonerFocoBtn cmdAceptar
        End If
    End If
   
End Sub


Private Sub Text4_LostFocus(index As Integer)
Dim devuelve As String

    On Error Resume Next
    
    If Not PerderFocoGnralLineas(Text4(index), ModificaLineas) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case index
        Case 0 'Codigo Direc/Dpto
            If Trim(Text4(index).Text) = "" Then Exit Sub
            FormateaCampo Text4(index)

        Case 4 'Cod. Postal
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text4(index - 1).Text = ObtenerPoblacion(Text4(index).Text, devuelve)
                Text4(index + 1).Text = devuelve
            End If
            VieneDeBuscar = False
        Case 8
            'PonerFocoBtn cmdAceptar
            
        Case 9
            If PonerFormatoEntero(Text4(index)) Then
                devuelve = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text4(index).Text, "N")
                If devuelve = "" Then
                    MsgBox "No existe la zona", vbExclamation
                    Text4(index).Text = ""
                    PonerFoco Text4(index)
                End If
            Else
                Text4(index).Text = ""
            End If
            Me.txtZona(9).Text = devuelve
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     
    Select Case Button.index
        
        Case 1  'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
           mnModificar_Click
        Case 3  'Borrar
           mnEliminar_Click
           
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
        
        Case 8
            'IMPRMIR
            If Modo = 2 Or Modo = 0 Then AbrirListadoOfer 47
            
            
'
'        Case 10, 11, 12, 13, 14, 15, 16
'            'Direcciones/Departamentos    -----
'            ' y direccion de envio y Renting y telefonia(ene2013)
'            ' campos(huertos) SEPT 2015
'            BotonDirecciones Button.Index - 5   'sera 5 o 6
'
'        Case 23    'Salir
'            mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub CargarComboAlbaran()
'### Combo Valorar Albaran con
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Todo, 1-Cantidad y Precio, 2-Cantidad

    cboAlbaran.Clear
    cboAlbaran.AddItem "Todo"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 0

    cboAlbaran.AddItem "Cantidad y Precio"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 1

    cboAlbaran.AddItem "Cantidad"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 2

End Sub


Private Sub CargarComboFacturacion()
'### Combo Tipo Facturaci�n
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


Private Sub CargarComboTipoIVA()
'### Combo Tipo de IVA a Aplicar
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Con Recargo de Equivalencia, 2-Exento de IVA

    Me.cboTipoIVA.Clear
    cboTipoIVA.AddItem "Normal"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 0

    cboTipoIVA.AddItem "Recargo Equivalencia"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 1

    cboTipoIVA.AddItem "Exento de IVA"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 2

    cboTipoIVA.AddItem "Intracomunitario"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 3
    
    'Junio 2012 Reducido
    cboTipoIVA.AddItem "Reducido"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 4

End Sub

Private Function InsertarModificarLinea() As Boolean
    Select Case Modo
    Case 5
        InsertarModificarLinea = InsertarModificarLineaDpto
    Case 6
        InsertarModificarLinea = InsertarModificarLineaEnvio
    Case 7
        InsertarModificarLinea = InsertarModificarLineaDatosConctacto
    Case 8
        InsertarModificarLinea = InsertarModificarLineaRenting
    Case 9
        InsertarModificarLinea = InsertarModificarLineaTelefonia
    Case 10
        InsertarModificarLinea = InsertarModificarLineamanipuladorFito
    Case 11
        InsertarModificarLinea = InsertarModificarLineaCamposhuertos

    End Select
    
    If InsertarModificarLinea Then
        Me.Refresh
        Espera 0.25
    End If
End Function
    
Private Function InsertarModificarLineaDpto() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLineaDpto = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO sdirec (codclien,coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba,codzona,iban"
            SQL = SQL & " , organogestor,unidadtramitadora,orgproponente,oficinacontable) VALUES ("
            SQL = SQL & Text1(0).Text & ", "
            SQL = SQL & Text3(0).Text
            For I = 1 To 5
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text3(I).Text, "T")
            Next I
                    
            For I = 6 To 19 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text3(I).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next I
                        
            SQL = SQL & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            SQL = "UPDATE sdirec Set nomdirec = " & DBSet(Text3(1).Text, "T")
            SQL = SQL & ", domdirec = " & DBSet(Text3(2).Text, "T")
            SQL = SQL & ", codpobla = " & DBSet(Text3(3).Text, "T")
            SQL = SQL & ", pobdirec = " & DBSet(Text3(4).Text, "T")
            SQL = SQL & ", prodirec = " & DBSet(Text3(5).Text, "T")
            SQL = SQL & ", perdirec = " & DBSet(Text3(6).Text, "T")
            'If Text3(7).Text <> "" Then SQL = SQL & ", fechainv = '" & Format(Text3(7).Text, "yyyy-mm-dd") & "'"
            'If Text3(8).Text <> "" Then SQL = SQL & ", horainve = '" & Format(Text3(8).Text, "hh:mm:ss") & "'"
            SQL = SQL & ", teldirec = " & DBSet(Text3(7).Text, "T")
            SQL = SQL & ", faxdirec = " & DBSet(Text3(8).Text, "T")
            SQL = SQL & ", maidirec = " & DBSet(Text3(9).Text, "T")
            'datos cuenta bancaria
            If Me.FrameCtaBanDpto.visible Then
                SQL = SQL & ", codbanco = " & DBSet(Text3(10).Text, "N", "S")
                SQL = SQL & ", codsucur = " & DBSet(Text3(11).Text, "N", "S")
                SQL = SQL & ", digcontr = " & DBSet(Text3(12).Text, "T")
                SQL = SQL & ", cuentaba = " & DBSet(Text3(13).Text, "T")
                SQL = SQL & ", iban = " & DBSet(Text3(15).Text, "T")
                
                SQL = SQL & ", organogestor = " & DBSet(Text3(16).Text, "N", "S")
                SQL = SQL & ", unidadtramitadora = " & DBSet(Text3(17).Text, "N", "S")
                SQL = SQL & ", orgproponente = " & DBSet(Text3(18).Text, "N", "S")
                SQL = SQL & ", oficinacontable = " & DBSet(Text3(19).Text, "N", "S")
                
                
            End If
            SQL = SQL & ", codzona = " & DBSet(Text3(14).Text, "N", "S")
            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
            SQL = SQL & " coddirec =" & (Text3(0).Text)
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaDpto = True
        TratarDptoEnTesoreria   'TESOERIA
    Else
        PonerFoco Text3(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Direcciones/Departamentos" & vbCrLf & Err.Description
End Function
    


Private Function InsertarModificarLineaEnvio() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLineaEnvio = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO sdirenvio (codclien,coddiren,nomdiren,perdiren,pobdiren,codpobla,prodiren,teldiren,faxdiren,observa,codzona,domdiren) VALUES ("
            SQL = SQL & Text1(0).Text & ", "
            SQL = SQL & Text4(0).Text
            For I = 1 To 5
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text4(I).Text, "T")
            Next I
                    
            For I = 6 To 8 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text4(I).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next I
            SQL = SQL & "," & DBSet(Text4(9).Text, "N", "S") & "," & DBSet(Text4(10).Text, "T", "S")
            SQL = SQL & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            SQL = "UPDATE sdirenvio Set nomdiren = " & DBSet(Text4(1).Text, "T")
            SQL = SQL & ", domdiren = " & DBSet(Text4(10).Text, "T")
            SQL = SQL & ", codpobla = " & DBSet(Text4(4).Text, "T")
            SQL = SQL & ", pobdiren = " & DBSet(Text4(3).Text, "T")
            SQL = SQL & ", prodiren = " & DBSet(Text4(5).Text, "T")
            SQL = SQL & ", perdiren = " & DBSet(Text4(2).Text, "T")
            SQL = SQL & ", teldiren = " & DBSet(Text4(6).Text, "T")
            SQL = SQL & ", faxdiren = " & DBSet(Text4(7).Text, "T")
            SQL = SQL & ", observa = " & DBSet(Text4(8).Text, "T")
            SQL = SQL & ", codzona = " & DBSet(Text4(9).Text, "N", "S")
            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
            SQL = SQL & " coddiren =" & (Text4(0).Text)
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaEnvio = True
    Else
        PonerFoco Text4(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Direcciones de envio" & vbCrLf & Err.Description
End Function




Private Sub MostrarSituacion(vMostrar As Boolean)
Dim Codigo As Integer
Dim Bloquea As String
Dim DescBloqueo As String

    On Error GoTo EMostrarSitu

    If Data1.Recordset.EOF Then Exit Sub
    If vMostrar Then
        Codigo = Data1.Recordset!codsitua
        If Not IsNull(Codigo) Then
            Me.lblSituacion.visible = (Codigo <> 0)
            Me.Frame1(1).visible = (Codigo <> 0)
            If Not (Codigo = 0) Then
            'Si situacion=0 (activo) no mostrar situacion
                Bloquea = DevuelveDesdeBDNew(conAri, "ssitua", "tipositu", "codsitua", CStr(Codigo), "N")
                DescBloqueo = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", CStr(Codigo), "N")
                If Val(Bloquea) = 0 Then
                    'Cliente NO Bloqueado
                    Me.lblSituacion.Caption = UCase(DescBloqueo)
                    Me.lblSituacion.ForeColor = vbBlue
                Else
                    'Cliente Bloqueado
                    Me.lblSituacion.Caption = "BLOQUEADO POR: " & UCase(DescBloqueo)
                    Me.lblSituacion.ForeColor = vbRed
                End If
            End If
        End If
    Else
        Me.lblSituacion.visible = False
        Me.Frame1(1).visible = False
    End If
EMostrarSitu:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


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





Private Sub CargaFrame_Direc()
Dim enlaza As Boolean

    'Crear las lineas de Direcciones/Departamentos para el cliente
    'ASignamos un SQL al DATA2
    Me.Data2.ConnectionString = conn
    enlaza = True
    If Text1(0).Text = "" Then enlaza = False
    
    
    
    
    CargaLineas enlaza, 5
    
    
    PonerModoOpcionesMenu
    
    
    
    'DesplazamientoVisible Me.ToolAux, 1 , True, CByte(cadCli )
End Sub


Private Sub CargaFrame_DirecEnv()
Dim enlaza As Boolean

    'Crear las lineas de Direcciones/Departamentos para el cliente
    'ASignamos un SQL al DATA2
    Me.Data3.ConnectionString = conn
    enlaza = False
    If Text1(0).Text <> "" Then enlaza = True
    CargaLineas enlaza, 6
    PonerModoOpcionesMenu
   
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

Private Sub ImagenDocumento(DatoEnElTag As Byte)

    On Error Resume Next
    
    imgDocumentos.Picture = frmPpal.ImgListPpal.ListImages(DatoEnElTag).Picture
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ImagenCRM(DatoEnElTag As Byte)

    On Error Resume Next
    
    imgCrm.Picture = frmPpal.ImgListPpal.ListImages(DatoEnElTag).Picture
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub ImagenesNavegacion()
'    With Me.ToolbarDoc
'        .ImageList = frmPpal.ImgListPpal
'        .Buttons(1).Image = 5
'        .Buttons(3).Image = 6
'        .Buttons(5).Image = 7
'        .Buttons(7).Image = 8
'        .Buttons(9).Image = 1
'        .Buttons(11).Image = 12
'        .Buttons(13).Image = 36
'        .Buttons(15).Image = 39
'    End With
    
    
    Set lw1.SmallIcons = frmPpal.ImgListPpal
    
    
    If vParamAplic.TieneCRM Then
'
'        With Me.Toolbar3
'            .ImageList = frmPpal.ImgListPpal
'            .Buttons(1).Image = 3
'            .Buttons(3).Image = 30
'            .Buttons(5).Image = 25
'            .Buttons(7).Image = 13
'            .Buttons(9).Image = 31
'            .Buttons(11).Image = 32
'            .Buttons(13).Image = 33
'        End With
'        Toolbar3.Buttons(5).visible = False
        Set lwCRM.SmallIcons = frmPpal.ImgListPpal
        
    End If
    
    
    

End Sub




'Private Sub ToolbarDoc_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Hacer_ButtonClick Button.Index, Button.Tag
'End Sub

'Private Sub Hacer_ButtonClick(indice As Integer, ElTag As String)
'
'    If ElTag = "" Then Exit Sub
'    LabelDoc.Caption = ""
'    'Levantamos todos los botones y dejamos pulsado el de ahora
'    For NumRegElim = 1 To ToolbarDoc.Buttons.Count
'        If ToolbarDoc.Buttons(NumRegElim).Tag <> "" Then
'            If ToolbarDoc.Buttons(NumRegElim).Index <> indice Then ToolbarDoc.Buttons(NumRegElim).Value = tbrUnpressed
'        End If
'    Next NumRegElim
'    CargaColumnas CByte(ElTag)
'
'    'Hacemos las acciones
'    If Modo = 2 Then CargaDatosLWDoc
'End Sub

Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader
    Me.FrameVisorDocumentos.visible = False
    FramePuntos.visible = False
    
    lw1.ListItems.Clear
    Select Case OpcionList
    Case 2, 3
        'ALBARANES
        If OpcionList = 3 Then
            LabelDoc.Caption = "Facturas"
            Columnas = "Forma de pago" & IIf(vParamAplic.TieneTelefonia2 > 0, "/ Tfno", "")
        Else
            LabelDoc.Caption = "Albaranes"
            Columnas = "Referencia"
        End If
        Columnas = "Tipo|Numero|Fecha|" & Columnas & "|Importe|"
        Ancho = "1200|1500|1800|4200|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy||" & FormatoImporte & "|"
        Ncol = 5
               
    Case 0, 1
        'OFERTAS  y PEDIDOS. Tienen la msimas colimnas (aprox)
        If OpcionList = 0 Then
            LabelDoc.Caption = "Ofertas"
            Columnas = "Acep."
        Else
            LabelDoc.Caption = "Pedidos"
            Columnas = "Visado"
        End If
        Columnas = "Numero|Fecha |Fec. entrega|" & Columnas & "|Referencia|Importe|"
        Ancho = "1900|1600|1600|990|2300|1600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|2|0|1|"
        'Formatos
        Formato = "00000000|dd/mm/yyyy|dd/mm/yyyy|||" & FormatoImporte & "|"
        Ncol = 6
    'Case 2
        '
        
    Case 4
        'PRECIOS ESPECIALES
        LabelDoc.Caption = "Precios especiales"
        Columnas = "Art�culo|Descripcion |Precio|F. cambio|Nuevo|"
        Ancho = "1800|4200|1550|1400|1550|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|0|1|"
        'Formatos
        Formato = "||" & FormatoImporte & "|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 5
    Case 5
        'DTO FAMILIA MARCA
        LabelDoc.Caption = "Dto Familia/Marca"
        Columnas = "Fecha|Dto1|Dto2|Familia|Marca|"
        Ancho = "1800|1000|1000|3800|2200|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|1|1|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy|" & FormatoImporte & "|" & FormatoImporte & "|||"
        Ncol = 5
        
    Case 6
        'DOCUMENTOS ASOCIADOS AL CLIENTE
        LabelDoc.Caption = "Documentos asociados"
        Columnas = "orden|Descripci�n|docum|codigo|leido|TipoDoc|"
        Ancho = "1000|8000|0|0|0|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|||"
        Ncol = 6
    
        Me.FrameVisorDocumentos.visible = True
        
        
    Case 7
        LabelDoc.Caption = "Puntos ventas"
        'numero,codtipom,numalbar,fechaalb,concepto,puntos
        Columnas = "Fecha|Descripci�n|Tipo|Id|puntos|saldo|"
        Ancho = "1800|3000|600|1900|1100|2000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|1|"
            'Formatos
        Formato = "dd/mm/yyyy||||" & FormatoImporte & "|" & FormatoImporte & "|"
        Ncol = 6
        FramePuntos.visible = True
    End Select
    
    
    'Fecha incio busquedas
    If Text1(46).Text = "" Then Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
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




Private Sub CargaDatosLWDoc()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & LabelDoc.Caption
    lblIndicador.Refresh
    CargaDatosLWDoc2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLWDoc2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim EsDTOFam As Boolean
Dim Saldo As Currency
Dim TemaPuntos As Boolean

Dim ConversionFechaHco2 As String
Dim CargaCatalogos As Boolean

    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    'For NumRegElim = 1 To ToolbarDoc.Buttons.Count
    '    If ToolbarDoc.Buttons(NumRegElim).Value = tbrPressed Then
    '        ElIcono = me.imgDocumentos.po  ToolbarDoc.Buttons(NumRegElim).Image
    '        Exit For
    '    End If
    'Next
    ElIcono = 0
    For NumRegElim = 0 To Me.optDoc.Count - 1
        If Me.optDoc(NumRegElim).Value Then
            ElIcono = Me.optDoc(NumRegElim).Tag
            Exit For
        End If
    Next
    
    'Fecha incio busquedas
    If Text1(46).Text = "" Then Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    
    cmdCatalogo.visible = False
    EsDTOFam = False
    CargaCatalogos = False
    TemaPuntos = False
    ConversionFechaHco2 = ""
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'ALBARANES
        cad = "select c.codtipom,c.numalbar,fechaalb,referenc,sum(importel) from scaalb c,slialb l where c.codtipom=l.codtipom and c.numalbar=l.numalbar"
        GroupBy = "1,2,3"
        BuscaChekc = "fechaalb"
        
    Case 0
        'OFERTAS
        cad = "select c.numofert,c.fecofert,fecentre,if(aceptado=1,""SI"","" ""),referenc ,sum(importel), 0 HCO from scapre c,slipre l where"
        cad = cad & " c.numofert=l.numofert "
        cad = cad & " and codclien=" & Data1.Recordset!codClien
        cad = cad & " and fecofert >='" & Format(imgFecha(3).Tag, FormatoFecha) & "'"
        
        'Truco. Si es un agente, solo puede ver las suyas
        If vParamAplic.NumeroInstalacion = 2 Then
            'HERBELCA
            If vUsu.CodigoAgente > 0 Then cad = cad & " AND c.codagent= " & vUsu.CodigoAgente
        End If
        cad = cad & " GROUP BY 1,2"
        
        
        If Text1(46).Text <> "" Then
            ConversionFechaHco2 = Text1(46).Text
            Text1(46).Text = DateAdd("yyyy", -2, Now)
        End If
        cad = cad & " UNION select c.numofert,c.fecofert,fecentre,if(aceptado=1,""SI"","" ""),referenc ,sum(importel)  ,1 HCO from schpre c,slhpre l where"
        cad = cad & " c.numofert=l.numofert "
        
        
        'Truco. Si es un agente, solo puede ver las suyas
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            'HERBELCA
            If vUsu.CodigoAgente > 0 Then cad = cad & " AND c.codagent= " & vUsu.CodigoAgente
        End If
        
        
        
        
        
        GroupBy = "1,2"
        BuscaChekc = "c.fecofert"
    Case 1
        'PEDIDOS
        cad = "select c.numpedcl,c.fecpedcl,fecentre,if(visadore=1,""SI"",""""),referenc,sum(importel) from scaped c,sliped l"
        cad = cad & " where c.numpedcl=l.numpedcl AND cerrado=0 "
        BuscaChekc = "fecpedcl"
        GroupBy = "1,2"
    Case 3
        cad = "select codtipom,numfactu,fecfactu,if(codtipom='FAT',telclien, nomforpa) "
        cad = cad & " ,totalfac from scafac,sforpa WHERE scafac.codforpa=sforpa.codforpa"
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
    Case 4
        'PRECIOS ESPECIALES
        cad = "select s.codartic,nomartic,precioac,fechanue,precionu from sprees s,sartic a where s.codartic=a.codartic"
        BuscaChekc = ""
        GroupBy = ""
        
    Case 5
        If vParamAplic.NumeroInstalacion = vbFenollar Then CargaCatalogos = True
    
    
    
        cad = "SELECT fechadto,dtoline1,dtoline2,nomfamia,nommarca,codclien"
        cad = cad & "  FROM (sdtofm LEFT OUTER JOIN sfamia ON sdtofm.codfamia=sfamia.codfamia) LEFT OUTER JOIN smarca ON sdtofm.codmarca=smarca.codmarca"
        cad = cad & " WHERE "
        EsDTOFam = True
    Case 6
        'IMAGENES-DOCUMENTOS
        cad = "select codigo,orden,descripfich,docum,0 from sfichdocs WHERE 1=1 "
        BuscaChekc = ""
        GroupBy = ""
        
    Case 7
        
        cad = "select fechaalb,nomtipom,smovalpuntos.codtipom,numalbar,puntos,0 saldo,concepto,observaciones,numero from smovalpuntos left join stipom on smovalpuntos.codtipom=stipom.codtipom WHERE true "
        BuscaChekc = ""
        GroupBy = ""
        TemaPuntos = True
    End Select
    
    
    'Para todos menos para Dtofamila marca
    
    If Not EsDTOFam Then
            'EL where del codclien
            cad = cad & " and codclien=" & Data1.Recordset!codClien
            
            'La fecha
            If BuscaChekc <> "" Then
                If Text1(46).Text <> "" Then
                    cad = cad & " and " & BuscaChekc & " >='" & Format(Text1(46).Text, FormatoFecha) & "'"
                Else
                    cad = cad & " and " & BuscaChekc & " >='" & Format(imgFecha(3).Tag, FormatoFecha) & "'"
                End If
            End If
            'El group by
            If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
            
            'El ORDER BY
            'BuscaChekc="" si es la opcion de precios especiales
            If CByte(RecuperaValor(lw1.Tag, 1)) = 6 Then
                cad = cad & " ORDER BY orden"
            ElseIf CByte(RecuperaValor(lw1.Tag, 1)) = 7 Then
                'Es PUNTOS
                cad = cad & " ORDER BY fechaalb, concepto"
            Else
                If BuscaChekc = "" Then BuscaChekc = " codartic "
                If BuscaChekc = "fecfactu" Then
                    'ORDENACION FACTURAS
                    cad = cad & " ORDER BY fecfactu desc, codtipom,numfactu desc"
                Else
                    'Ofertas
                    If CByte(RecuperaValor(lw1.Tag, 1)) = 0 Then BuscaChekc = "2" 'Orden by la segunda columna
                
                    cad = cad & " ORDER BY " & BuscaChekc & " DESC"
                End If
            End If
    Else
        'Para familia marca
        cad = cad & " (codclien=" & Data1.Recordset!codClien & " AND codactiv is null)"
        cad = cad & " OR (codactiv = " & Data1.Recordset!codactiv & " AND codclien is null)"
    End If
    BuscaChekc = ""
    
    
    If CByte(RecuperaValor(lw1.Tag, 1)) = 6 Then
        
        CargarArchivos True, 0, True
    
    Else
        Saldo = 0
        lw1.ListItems.Clear
    
        Set RS = New ADODB.Recordset
        
        If CargaCatalogos Then CargaCatalogosCliente
            
        
        
        
        
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set IT = lw1.ListItems.Add()
          
            
            If lw1.ColumnHeaders(1).Tag <> "" Then
                IT.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
            Else
                IT.Text = RS.Fields(0)
            End If
          
            'El resto de cmpos
            For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
                If IsNull(RS.Fields(NumRegElim - 1)) Then
                    IT.SubItems(NumRegElim - 1) = " "
                Else
                    If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                        IT.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                    Else
                        IT.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                    End If
                End If
               
            Next
             
             
            If CByte(RecuperaValor(lw1.Tag, 1)) = 0 Then
                IT.SmallIcon = IIf(RS!HCO = 0, ElIcono, 1)
                If RS!HCO = 1 Then IT.ToolTipText = "Historico"
            Else
                IT.SmallIcon = ElIcono
            End If
            'Para familia /dto
            If EsDTOFam Then
                'Si codclien es >0 then
                If DBLet(RS!codClien, "N") > 0 Then IT.Bold = True
            End If
            
            If TemaPuntos Then
                Saldo = Saldo + RS!Puntos
                IT.SubItems(5) = Format(Saldo, FormatoImporte)
                IT.Tag = 0
                
                'Si el concpto NO es cero, cambio el icono
                If RS!Concepto > 0 Then
                    If RS!Concepto = 1 Then
                        IT.SmallIcon = 2
                        IT.SubItems(1) = "Canje puntos"
                    Else
                        IT.SmallIcon = 3
                        IT.Tag = RS!numero
                        IT.SubItems(1) = Mid(RS!Observaciones, 1, 40)
                    End If
                End If
            End If
            RS.MoveNext
        Wend
        RS.Close
        
        
        If TemaPuntos Then
            If Saldo <> DBLet(Data1.Recordset!Puntos, "N") Then
                Set IT = lw1.ListItems.Add()
                IT.Text = "ERROR"
                IT.ForeColor = vbRed
                IT.Bold = True
                IT.SubItems(1) = " "
                IT.SubItems(2) = " "
                IT.SubItems(3) = "Cliente"
                IT.SubItems(4) = " "
                IT.SubItems(5) = Format(DBLet(Data1.Recordset!Puntos, "N"), FormatoCantidad)
                IT.ListSubItems(5).Bold = True
                IT.ListSubItems(5).ForeColor = vbRed
                
            Else
                'EL ULTIMO ITEM es correcto.
                If Not IT Is Nothing Then
                    IT.ListSubItems(5).Bold = True
                    IT.ListSubItems(5).ForeColor = vbBlue
                    'IT.Text = "NNN"
                End If
            End If
        End If
        
        If Not IT Is Nothing Then IT.EnsureVisible
        
    End If
    
    Set RS = Nothing
    
    
    If Me.lw1.ListItems.Count > 0 Then lw1.ListItems(1).EnsureVisible
    
    
    If ConversionFechaHco2 <> "" Then Text1(46).Text = ConversionFechaHco2
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    If ConversionFechaHco2 <> "" Then Text1(46).Text = ConversionFechaHco2
End Sub



Private Sub AbrirFacturaLW()
Dim s As String
'    Set miRsAux = New ADODB.Recordset
    
'
'    If lw1.SelectedItem.Text = "FAM" Then
        'Van directas
        s = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(2) & "|"
'    Else
'        s = "select codtipoa,numalbar,fechaalb from scafac1 where codtipom='"
'        s = s & lw1.SelectedItem.Text & "' and numfactu=" & lw1.SelectedItem.SubItems(1)
'        s = s & " and fecfactu='" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "' ORDER BY codtipoa desc"
'        miRsAux.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        s = ""
'        If Not miRsAux.EOF Then
'            s = miRsAux.Fields(0) & "|" & miRsAux.Fields(1) & "|" & miRsAux.Fields(2) & "|"
'        End If
'        miRsAux.Close
'        Set miRsAux = Nothing
'    End If
    
    If s <> "" Then
        With frmFacHcoFacturas2
                .DesdeFichaCliente = True
                .hcoCodMovim = RecuperaValor(s, 2)
                .hcoCodTipoM = RecuperaValor(s, 1)
                .hcoFechaMov = RecuperaValor(s, 3)
                .Show vbModal
        End With
    End If
End Sub


Private Function TratarDptoEnTesoreria() As Boolean
Dim Existe As Boolean
Dim C As String


    
    If Text1(35).Text = "" Or Text2(35).Text = "" Then
        
        MsgBox "Cuenta contable erronea.", vbExclamation
        Exit Function
    End If


    Existe = False
    C = "codmacta = '" & Text1(35).Text & "' and Dpto "
    C = DevuelveDesdeBD(conConta, "descripcion", "departamentos", C, Text3(0).Text)
    If C <> "" Then Existe = True
    
    
    If Existe Then
        If ModificaLineas = 1 Then
            'Estamos insertando y ya existe. UPDATEAMOS
            
        End If
        'UPDATEAMOS
        C = "UPDATE  departamentos set Descripcion = " & DBSet(Text3(1).Text, "T")
        C = C & " WHERE codmacta= '" & Text1(35).Text & "' AND Dpto = " & Text3(0).Text
    Else
        'NO EXISTE... creamos
        C = "insert into `departamentos` (`codmacta`,`Dpto`,`Descripcion`) values ('"
        C = C & Text1(35).Text & "'," & Text3(0).Text & "," & DBSet(Text3(1).Text, "T") & ")"
        
    End If
    ConnConta.Execute C
    
End Function


'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'  CRM
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------


Private Sub CargaColumnasCRM(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader
Dim Ordena As Integer
    'Las llamadas cogera las llamadas recibidas desde sllama y las efectuadas desde acciones comerciales con tipoaccion=1
    'para poder ordenarlas tendremos una columna viiblefalse con yyymmddhhmmss
    Ordena = -1
    Select Case OpcionList
    Case 0
        'Acciones comerciales
        LabelCRM.Caption = "Acciones comerciales"
        
        Columnas = "Fecha|Usuario|Estado|Medio|Tipo|Descripcion|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "3100|1200|1800|1400|1200|7300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||0000||"
        Ncol = 6
               
    Case 1
        'Llamadas
        LabelCRM.Caption = "Llamadas "
        
        Columnas = "Fecha|Usuario|Tipo/Trab|Observaciones|Orden|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "3100|1200|3400|8500|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 5
    
        Ordena = 5
        '************************* Movemos el 3 al 2 , ya que el 2 lo quitamos
'    Case 2
'        LabelCRM.Caption = "E-mail"
'        Columnas = "Fecha|Enviado|Email|Asunto|Adj|entryID|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
'        Ancho = "1800|825|2565|3899|495|0|"
'        'vwColumnRight =1  left=0   center=2
'        Alinea = "0|0|0|0|0|"
'        'Formatos
'        Formato = "dd/mm/yyyy hh:mm||||||"
'        Ncol = 6
    
    Case 2
        'COBROS
        LabelCRM.Caption = "Cobros pendientes"
        Columnas = "Fecha Vto.|Factura|Fecha factura|Forma pago|Importe|Cobrado|Pendiente|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1800|1600|1800|3900|2200|1800|2295|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|1|0|0|1|1|1|"
        'Formatos
        Formato = "dd/mm/yyyy||dd/mm/yyyy||" & FormatoImporte & "|" & FormatoImporte & "|" & FormatoImporte & "|"
        Ncol = 7
        
    Case 3
        'COBROS
        LabelCRM.Caption = "Observaciones departamento"
        Columnas = "Departamento|Fecha|Observaciones||"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "2100|1800|10500|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|||"
        Ncol = 4
        
        
    Case 4
        'Reclamaciones
        LabelCRM.Caption = "Reclamaciones"
        Columnas = "Fecha|Factura|Observaciones|Importe|codigo|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1900|1700|8500|1600|0|"  'La ultima esta oculta
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy|||" & FormatoImporte & "||"
        Ncol = 5
        
    
    Case 5
        'H I S T O R I A L
        LabelCRM.Caption = "Historial"
        Columnas = "Fecha|Usuario|Trabajador|Observaciones|"
        Ancho = "2800|1600|3500|8200|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 4
        
    
    End Select
    
    
   
    cmdAccCRM(0).visible = Modo = 2 And OpcionList <> 2
    cmdAccCRM(1).visible = Modo = 2
    cmdAccCRM(2).visible = Modo = 2 And OpcionList = 3 'Or OpcionList = 6
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
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Kopc As Byte
Dim MeteIT As Boolean
Dim ConexionConta As Boolean  'Si no es conta es ARIGES( conn)
    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    
    ElIcono = 0
    For NumRegElim = 0 To Me.optCRM.Count - 1
        If Me.optCRM(NumRegElim).Value Then
            ElIcono = Me.optCRM(NumRegElim).Tag
            Exit For
        End If
    Next
    
    'Fecha incio busquedas
    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")

    'EL where del codclien     lo lleva cada sql
    Kopc = CByte(RecuperaValor(lwCRM.Tag, 1))
    ConexionConta = False
    Select Case Kopc
    Case 0
        'Acciones comerciales
        cad = "select fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion from scrmacciones,scrmtipo WHERE scrmacciones.tipo= scrmtipo.codigo "
        cad = cad & " and codclien=" & Data1.Recordset!codClien
        
        'Los tipo 1 y dos NO van aqui. El 3, que es RENOVACION TFNO SI
        cad = cad & " AND (tipo>=3 or tipo > 20)"  'las 20 primerasprobablemebne no sepongan aqui
        GroupBy = ""
        BuscaChekc = "fechora"
    Case 1
        'Llamadas
        cad = "select feholla,usuario,nomllama1,observac,date_format(feholla,""%Y%m%d%H%i%s"") from sllama,sllama1  where"
        cad = cad & " sllama.codllama1 = sllama1.codllama1"
        cad = cad & " and codclien=" & Data1.Recordset!codClien
        GroupBy = ""
        BuscaChekc = "feholla"
    
'    Case 2
'
'        'eMAIL
'        cad = "select fechahora, if(enviado=1,""Enviado"",""Recibido""),email,asunto,"
'        cad = cad & "if(adjuntos<>"""",""*"","""") ,entryID from scrmmail"
'        cad = cad & " WHERE codclien=" & Data1.Recordset!codClien
'        GroupBy = ""
'        BuscaChekc = "fechahora"
'
    Case 2
        'Cobros pendientes
        If vParamAplic.ContabilidadNueva Then
            cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",numfactu),7)),fecfactu,nomforpa,impvenci,gastos,"
            cad = cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
            cad = cad & " FROM  cobros scobro INNER JOIN formapago sforpa ON scobro.codforpa=sforpa.codforpa "
            
            
        Else
            cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",codfaccl),7)),fecfaccl,nomforpa,impvenci,gastos"
            cad = cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
            cad = cad & " FROM  scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
            
        End If
        cad = cad & " WHERE scobro.codmacta = '" & Text1(35).Text & "' "
        cad = cad & " and recedocu=0 "

        'PARA TEINSA
        If vParamAplic.NumeroInstalacion = 3 Then cad = cad & " AND (sforpa.tipforpa between 0 and 3) "
        BuscaChekc = "fecvenci"
        ConexionConta = True
        
    Case 3
        'Observaciones departamento
        cad = "select if(dpto=1,""Administracion"",if(dpto=2,""Comercial"",if(dpto=3,""SAT"",""Direcci�n""))),fecha,observa,dpto from scrmobsclien"
        cad = cad & " WHERE codclien=" & Data1.Recordset!codClien
        BuscaChekc = "dpto"
        
    Case 4
        'Reclamaciones
        'Cobros pendientes
        If vParamAplic.ContabilidadNueva Then
            cad = "select fecreclama,concat(numserie,right(concat(""00000000"",numfactu),7)),observaciones,if (impvenci is null,importes,impvenci) impvenci,reclama.codigo,numlinea"
            cad = cad & " FROM  reclama  left join reclama_facturas  on reclama.codigo=reclama_facturas.codigo"
            cad = cad & " WHERE codmacta='" & Text1(35).Text & "' "
            BuscaChekc = "fecreclama desc ,reclama.codigo,numlinea "
        Else
            cad = "select fecreclama,concat(numserie,right(concat(""00000000"",codfaccl),7)),observaciones,impvenci,codigo"
            cad = cad & " from shcocob where codmacta='" & Text1(35).Text & "' "
            BuscaChekc = "fecreclama desc ,codigo "
        End If
        ConexionConta = True
        
        
    Case 5
        'Historial
        cad = "select fechora ,usuario,nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmacciones.tipo=2  and codclien= " & Data1.Recordset!codClien   '2 DE historial
        GroupBy = ""
        BuscaChekc = "fechora"
    End Select
    
    
    
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc
    If Kopc <> 3 Then cad = cad & " DESC"

    
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
        If Kopc <> 2 Then
            MeteIT = True
        Else
            If RS!Tot <> 0 Then
                MeteIT = True
            Else
                MeteIT = False
            End If
        End If
        
        If MeteIT Then
                Set IT = lwCRM.ListItems.Add()
                 
                If lwCRM.ColumnHeaders(1).Tag <> "" Then
                    IT.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
                Else
                    IT.Text = RS.Fields(0)
                End If
                'El resto de cmpos
                For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
                    If IsNull(RS.Fields(NumRegElim - 1)) Then
                        IT.SubItems(NumRegElim - 1) = " "
                    Else
                    
                        If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                            IT.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                        Else
                        
                            
                            'Cad = RS.Fields(NumRegElim - 1)
                            cad = DBLetMemo(RS.Fields(NumRegElim - 1))
                            'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                            If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
                            'para las observaciones de la reclamacion tb quito los vbcrlf
                            If NumRegElim = 3 And Kopc = 4 Then cad = Replace(cad, vbCrLf, " ")
                            
                            'Medio
                            If NumRegElim = 3 And Kopc = 0 Then DevuelveMedio cad
                            If NumRegElim = 3 And Kopc = 3 Then cad = Replace(cad, vbCrLf, " ")
                            
                            
                            
                            IT.SubItems(NumRegElim - 1) = cad
                        
                            
                            
                        End If
                    End If
                Next
                
                
                If Kopc = 4 And vParamAplic.ContabilidadNueva Then
                    'Para las reclamaciones, en la contabiiada nueva, PODRIA  llevar lineas
                    IT.Tag = DBLet(RS!numlinea, "T")
                End If
                
                'El icono
                If Kopc = 1 Then
                    IT.SmallIcon = 27
                ElseIf Kopc = 22 Then

                    If RS.Fields(1) = "Enviado" Then
                        IT.SmallIcon = 28
                    Else
                        IT.SmallIcon = 29
                    End If
                Else
                    'el resto ponemos el del toolbar
                    IT.SmallIcon = ElIcono
                End If
        End If
        
        
    
        RS.MoveNext
    Wend
    RS.Close
    
    
    If Kopc = 1 Then
        'Llamadas. Las efectuadas las hago desde este punto
        cad = "select fechora ,usuario,nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmacciones.tipo=1  and codclien= " & Data1.Recordset!codClien
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '
            'Coje datos desde dos tablas
            Set IT = lwCRM.ListItems.Add()
            IT.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
           
            For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
                If IsNull(RS.Fields(NumRegElim - 1)) Then
                    IT.SubItems(NumRegElim - 1) = " "
                Else
                
                    If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                        IT.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                    Else
                    
                        
                        cad = RS.Fields(NumRegElim - 1)
                        'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                        If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
  
                        IT.SubItems(NumRegElim - 1) = cad
                    
                        
                        
                    End If
                End If
            Next
            IT.SmallIcon = 26
            RS.MoveNext
        Wend
        RS.Close
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
Dim TieneDatosDpto As Boolean

    On Error GoTo ELanzarProgramaEmails

    If Dir(App.Path & "\AriOutlook.exe", vbArchive) = "" Then
        MsgBox "No tienen el programa de asignacion de mails al CRM de Ariadna", vbExclamation
        Exit Sub
    End If
    
    TieneDatosDpto = False
    If Not Data2.Recordset Is Nothing Then
        If Not Data2.Recordset.EOF Then TieneDatosDpto = True
    End If
        
        
    'Como lanzamos el programa
    '*************************  dbariges|codclien|nombre||||mails que se utlizan|
    If TieneDatosDpto Then
        BuscaChekc = "Select distinct(maidirec) from sdirec where codclien=" & Data1.Recordset!codClien & " AND maidirec <>"""""
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    BuscaChekc = ""
    If Text1(17).Text <> "" Then BuscaChekc = BuscaChekc & Text1(17).Text & "|"  'mail1
    If Text1(18).Text <> "" Then BuscaChekc = BuscaChekc & Text1(18).Text & "|"  'mail1
        
        
    If TieneDatosDpto Then
        While Not miRsAux.EOF
            If Not IsNull(miRsAux!maidirec) Then
                If miRsAux!maidirec <> "" Then BuscaChekc = BuscaChekc & miRsAux!maidirec & "|"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    BuscaChekc = vUsu.CadenaConexion & "|" & Data1.Recordset!codClien & "|" & CStr(Data1.Recordset!NomClien) & "||||" & BuscaChekc
    
    Shell App.Path & "\AriOutlook.exe " & BuscaChekc, vbNormalFocus
    
    Espera 2
    
    
ELanzarProgramaEmails:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lanzar Programa Email"
    Set miRsAux = Nothing
    BuscaChekc = ""
End Sub






Private Sub CargaLineas(enlaza As Boolean, Cual_ As Byte)
'cual:     0  percontac, 1  renting   , 2 telefonos    3 fitos  4 Campos(huertos)
'          5 departamentos     6  Direcciones de envio
'          8 Todos
Dim SQL As String


        If Cual_ = 0 Or Cual_ = 8 Then
            SQL = "SELECT nombre,cargo,dpto,telefono,ext,maidirec,movil,observa,id,codclien FROM scliendp where "
            If enlaza Then
                SQL = SQL & "codclien = " & Text1(0).Text
                
            Else
                SQL = SQL & " false"
            End If
             
            SQL = SQL & " ORDER BY  id"
            CargaGridGnral DataGrid1, Me.Data4, SQL, PriVezForm, 330
            SQL = "S|txtauxDC(0)|T|Nombre|5000|;S|txtauxDC(2)|T|Cargo|4300|;S|cmdCargos|B||0|;"
            'Los campos que no se ven que van FUERA DEL GRID
            SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla SQL, DataGrid1, Me, 330
            DataGrid1.ScrollBars = dbgAutomatic
            If cboCargo.Width < 2000 Then
                    
                cboCargo.Left = txtauxDC(2).Left
                cboCargo.Width = txtauxDC(2).Width + 60
            End If
            Me.cmdCargos.Left = txtauxDC(2).Left + cboCargo.Width + 15
        End If
        
        If vParamAplic.Renting Then
            If Cual_ = 1 Or Cual_ = 8 Then
                SQL = "SELECT id,sclienrenting.coddirec,nomdirec,referencia,fecalta,numcuotas,fecbaja,importe"
                SQL = SQL & ",sclienrenting.codtipco,nomtipco,obser,ultfec"
                SQL = SQL & " from (sclienrenting left join sdirec on sclienrenting.codclien=sdirec.codclien"
                SQL = SQL & " and sdirec.coddirec=sclienrenting.coddirec ) "
                SQL = SQL & " inner join stipco on stipco.codtipco=sclienrenting.codtipco"
                SQL = SQL & " WHERE "
                If enlaza Then
                    SQL = SQL & " sclienrenting.codclien = " & Text1(0).Text
                    
                Else
                    SQL = SQL & " false"
                End If
                SQL = SQL & " ORDER BY  id"
                CargaGridGnral DataGrid2, Me.Data5, SQL, PriVezForm, 330
                
                SQL = "S|txtauxRent(0)|T|ID|700|;"
                If vParamAplic.HayDeparNuevo = 1 Then
                    SQL = SQL & "S|txtauxRent(1)|T|Dpto|750|"
                Else
                    SQL = SQL & "S|txtauxRent(1)|T|Dir.|750|"
                End If
                SQL = SQL & ";S|cmdRenting(0)|B||0|;S|txtauxRent(2)|T|Departamento|3150|;"
                SQL = SQL & "S|txtauxRent(3)|T|Referencia|2800|;S|txtauxRent(4)|T|Fecha alta|1300|;S|cmdRenting(1)|B||0|;"
                SQL = SQL & "S|txtauxRent(5)|T|Cuotas|750|;S|txtauxRent(6)|T|Fecha baja|1300|;S|cmdRenting(2)|B||0|;"
                SQL = SQL & "S|txtauxRent(7)|T|Importe|1350|;"
                'no se ven
                SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;"
                arregla SQL, DataGrid2, Me, 330
                DataGrid1.ScrollBars = dbgAutomatic
                'Como el lo pone a la derecha
                txtauxRent(1).Alignment = 0 'a la izda
                                
            End If
        
        End If
        
        
        If vParamAplic.TieneTelefonia2 > 0 Then
            If Cual_ = 2 Or Cual_ = 8 Then
                SQL = "select  IdTelefono,stfnooperador.nombre ,operador,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones,coddirec,clienppal,"
                SQL = SQL & " modelo,coninternet,puntos,fechaalta,cuotaminima,fecharenove,procedencia "
                If vParamAplic.TelefoniaVtaPlazos Then SQL = SQL & " ,ArtPlazos,PlazosMeses,ImportePlazo,PlazosOrigen"
                SQL = SQL & " ,if(Inactivo=1,'*','') as B "
                SQL = SQL & "  FROM sclientfno,stfnooperador WHERE sclientfno.operador=stfnooperador.codoperador  AND "
                
                If enlaza Then
                    SQL = SQL & "codclien = " & Text1(0).Text
                Else
                    SQL = SQL & " false"
                End If
                SQL = SQL & " ORDER BY  IdTelefono"
                CargaGridGnral DataGrid3, Me.data6, SQL, PriVezForm, 330
                SQL = "S|txtauxTfno(0)|T|Tel�fono|1380|;S|cboOperadorTfnnia2(0)|C|Operador|1790|;N|||||;"
                SQL = SQL & "S|txtauxTfno(1)|T|IMEI|2590|;S|txtauxTfno(2)|T|SIM|2600|;"
                
                'Los campos que no se ven que van FUERA DEL GRID
                SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
                If vParamAplic.TelefoniaVtaPlazos Then SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;" 'vta a plazos mese importe mesestot
                arregla SQL, DataGrid3, Me, 330
                DataGrid3.ScrollBars = dbgAutomatic
                
                DataGrid3.Columns(IIf(vParamAplic.TelefoniaVtaPlazos, 21, 18)).Width = 340
            End If
        End If
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If Cual_ = 3 Or Cual_ = 8 Then
                SQL = "select  cif,nombre,if(tipocarnet=2,'Cualificado','B�sico') tipo,numcarnet,fcaducidad,telefono"
                SQL = SQL & ", if (Manipuladorprovisional=0,'','Si') PROV,if(ImgDNI is null, '','*') DNI,if(ImgManipula is null, '','*') as 'Car.'"
                SQL = SQL & ",id  FROM sclienmani WHERE  "
                If enlaza Then
                    SQL = SQL & "codclien = " & Text1(0).Text
                Else
                    SQL = SQL & " false"
                End If
                SQL = SQL & " ORDER BY  id"
                CargaGridGnral DataGrid4, Me.Data7, SQL, PriVezForm, 330
                SQL = "S|txtauxFito(0)|T|CIF|1500|;"
                SQL = SQL & "S|txtauxFito(1)|T|Nombre|4800|;"
                SQL = SQL & "S|cboFitos(0)|C|Tipo|1200|;S|txtauxFito(2)|T|Referencia|2810|;"
                SQL = SQL & "S|cmdFitos(0)|B||0|;S|txtauxFito(5)|T|Caducidad|1750|;"
                
                SQL = SQL & "S|txtauxFito(3)|T|Telefono|2100|;"
                SQL = SQL & "S|cboFitos(1)|C|Provi.|600|;||||100|;||||150|;"
                SQL = SQL & "N|txtauxFito(4)|T|id|0|;"
                arregla SQL, DataGrid4, Me, 330
                DataGrid4.ScrollBars = dbgAutomatic
                
                cmdFitos(0).Height = DataGrid4.RowHeight
            End If
        End If
        
        
        
        'Sept 2015
        If vParamAplic.Huertos Then
            If Cual_ = 4 Or Cual_ = 8 Then
                SQL = "select id, poligono,parcela, recintos,supsigpa,supderec,partida,fecaltas,fecbajas,observac"
                'id,codparti,fecaltas,fecbajas,supsigpa,supderec,poligono,parcela,recintos,observac
                SQL = SQL & "  from sclienhuertos WHERE  "
                If enlaza Then
                    SQL = SQL & "codclien = " & Text1(0).Text
                Else
                    SQL = SQL & " false "
                End If
                SQL = SQL & " ORDER BY  1"
                CargaGridGnral DataGrid5, Me.data8, SQL, PriVezForm, 330
                'poligono,codparti, recintos,supsigpa,supderec,fecaltas,fecbajas,observac,id"
                SQL = "S|txtauxMarja(0)|T|id|690|;"
                SQL = SQL & "S|txtauxMarja(1)|T|Pol�gono|1690|;"
                SQL = SQL & "S|txtauxMarja(2)|T|Parcela|1650|;"
                SQL = SQL & "S|txtauxMarja(3)|T|Recintos|1650|;"
                SQL = SQL & "S|txtauxMarja(4)|T|SIGPAC(ha)|2100|;"
            
                SQL = SQL & "S|txtauxMarja(5)|T|Sup.derechos(ha)|2200|;"
                'SQL = SQL & "S|txtauxMarja(6)|T|Partida|900|;"
                SQL = SQL & "N|||||;"
                SQL = SQL & "N|||||;"
                SQL = SQL & "N|||||;"
                SQL = SQL & "N|||||;"
                'Aunque no se vean, pongo el caption de la columna, para despues en el datosok poner que campo me falta
                DataGrid5.Columns(6).Caption = "Fecha alta"
                arregla SQL, DataGrid5, Me, 330
                DataGrid5.ScrollBars = dbgAutomatic
                
               
            End If
        End If
        
                
        
        'Junio18
        If Cual_ = 5 Or Cual_ = 8 Then
            SQL = "SELECT coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba,codzona,iban"
            SQL = SQL & " , organogestor,unidadtramitadora,orgproponente,oficinacontable FROM sdirec WHERE "
            If enlaza Then
                SQL = SQL & " codclien = " & Text1(0).Text
            Else
                SQL = SQL & " false"
            End If
            SQL = SQL & " ORDER BY  coddirec"
            
            CargaGridGnral DataGrid6, Me.Data2, SQL, PriVezForm
            
            SQL = "S|Text3(0)|T|ID|1300|;S|Text3(1)|T|Nombre|5950|;"
            
            SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;"
             arregla SQL, DataGrid6, Me, 330
            DataGrid6.ScrollBars = dbgAutomatic
            'Como el lo pone a la derecha
            Text3(0).Alignment = 0 'a la izda
        End If
    
        If vParamAplic.DireccionesEnvio Then
            If Cual_ = 6 Or Cual_ = 8 Then
                SQL = "SELECT coddiren,nomdiren,perdiren,pobdiren,codpobla,prodiren,teldiren,faxdiren,observa,codzona,domdiren FROM sdirenvio WHERE "
                If enlaza Then
                    SQL = SQL & " codclien = " & Text1(0).Text
                Else
                    SQL = SQL & " false"
                End If
                SQL = SQL & " ORDER BY  coddiren"
                
                CargaGridGnral DataGrid7, Me.Data3, SQL, PriVezForm
                
                SQL = "S|Text4(0)|T|ID|600|;S|Text4(1)|T|Nombre|3950|;S|Text4(2)|T|Contacto|3950|;"
                
                SQL = SQL & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
                arregla SQL, DataGrid7, Me, 330
                DataGrid7.ScrollBars = dbgAutomatic
                'Como el lo pone a la derecha
                Text4(0).Alignment = 0 'a la izda
            End If
        End If
        
                
        
        
End Sub


Private Sub PonerDatosForaGridContacto(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data4.Recordset Is Nothing) Then
            If Not Data4.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For I = 0 To txtauxDC.Count - 1
            txtauxDC(I).Text = ""
        Next I
        
    Else
        'EL
        
        PonerCamposFormaFrame Me, "txtauxDC", Data4
        
        
    End If
End Sub



Private Sub PonerDatosForaGridRent(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data5.Recordset Is Nothing) Then
            If Not Data5.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For I = 8 To txtauxRent.Count - 1
            txtauxRent(I).Text = ""
        Next I
        
    Else
        'EL
        
        PonerCamposFormaFrame Me, "txtauxRent", Data5
        
        
    End If
End Sub



Private Sub PonerDatosForaGridTfno(ForzarLimpiar As Boolean)
Dim I As Integer
Dim J As Integer
Dim Limp As Boolean


    Limp = True
    If Not ForzarLimpiar Then
        If Not (data6.Recordset Is Nothing) Then
            If Not data6.Recordset.EOF Then Limp = False
        End If
    End If
    
    lwTfnoCuotas.ListItems.Clear
    If Limp Then

        'Limpiamos
        For I = 0 To txtauxTfno.Count - 1
            If I < 3 Then Me.chkTelefonia(I).Value = 0
            txtauxTfno(I).Text = ""
            If I > 3 And I < 7 Then Me.Text5(I).Text = "" '4-5-6
        Next I
        cboOperadorTfnnia2(0).ListIndex = -1
        cboOperadorTfnnia2(1).ListIndex = -1
        
        
                
    Else
        'Pongo los campos en los txt
        J = IIf(vParamAplic.TelefoniaVtaPlazos, 14, 10)
        For I = 0 To J
        
                BuscaChekc = RecuperaValor("IdTelefono|IMEI|SIM|Observaciones|coddirec|clienppal|modelo|cuotaminima|puntos|fechaalta|fecharenove|ArtPlazos|PlazosMeses|ImportePlazo|PlazosOrigen|", I + 1)
                Me.txtauxTfno(I).Text = DBLet(data6.Recordset.Fields(BuscaChekc), "T")
                If I > 3 And I < 7 Then txtauxTfno_LostFocus I
                If I >= 11 Then txtauxTfno_LostFocus I
        Next
        SituarCombo Me.cboOperadorTfnnia2(0), DBLet(data6.Recordset!Operador, "N")
        SituarCombo Me.cboOperadorTfnnia2(1), DBLet(data6.Recordset!procedencia, "N")
        For I = 0 To 3

                BuscaChekc = RecuperaValor("Factura|Detalle|Inactivo|coninternet|", I + 1)
                BuscaChekc = DBLet(data6.Recordset.Fields(BuscaChekc), "T")
                Me.chkTelefonia(I).Value = Abs(BuscaChekc = "1")

        Next
        
        
        'Solo para alzira y Bolbaite y demas   2=catadau
        CargaCuotasTelefonia 0
         

        BuscaChekc = ""
    End If
End Sub


Private Sub PonerDatosForaGridDpto(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data2.Recordset Is Nothing) Then
            If Not Data2.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For I = 0 To txtauxDC.Count - 1
            txtauxDC(I).Text = ""
        Next I
        
    Else
        'EL
        
        PonerCamposFormaFrame Me, "txtauxDC", Data4
        
        
    End If
End Sub






Private Sub CargaCuotasTelefonia(QueItem As Integer)
Dim RP As ADODB.Recordset
Dim I As Byte


    Me.lwTfnoCuotas.ListItems.Clear
    Set RP = New ADODB.Recordset
    BuscaChekc = "select * from sclientfnoCuotas where idtelefono=" & DBSet(data6.Recordset!idtelefono, "T") & " ORDER BY numlinea"
    RP.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RP.EOF
        I = I + 1
        Me.lwTfnoCuotas.ListItems.Add , "N" & Format(RP!numlinea, "000"), RP!Descripcion
        lwTfnoCuotas.ListItems(I).SubItems(1) = Format(RP!Precio, FormatoPrecio)
        lwTfnoCuotas.ListItems(I).ToolTipText = RP!Descripcion
        If I = QueItem Then Set Me.lwTfnoCuotas.SelectedItem = lwTfnoCuotas.ListItems(I)
        
        RP.MoveNext
    Wend
    Set RP = Nothing
            
End Sub

Private Sub LLamaLineasDatosContacto(alto As Single, xModo As Byte)
Dim b As Boolean

    ModificaLineas = xModo
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    b = Modo = 7 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas


    DeseleccionaGrid Me.DataGrid1
    
    txtauxDC(0).Height = DataGrid1.RowHeight
    txtauxDC(0).visible = b
    txtauxDC(0).Top = alto
    'txtauxDC(1).Height = DataGrid1.RowHeight
    'txtauxDC(1).visible = B
    'txtauxDC(1).Top = alto
    cmdCargos.visible = b
    cmdCargos.Top = alto
    Me.cboCargo.Top = alto
    If b Then
        SituarCboCargo
    Else
        cboCargo.visible = False
    End If

End Sub


Private Sub LLamaLineasTfnia(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 9 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas


    
    DeseleccionaGrid Me.DataGrid3
     DataGrid3.Enabled = Not b
    For I = 0 To 2
        txtauxTfno(I).Height = DataGrid3.RowHeight
        txtauxTfno(I).visible = b
        txtauxTfno(I).Top = alto
        
    Next
    Me.cboOperadorTfnnia2(0).visible = b
    Me.cboOperadorTfnnia2(0).Top = alto
End Sub



Private Sub LLamaLineasFito(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 10 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    DataGrid4.Enabled = Not b
    
    DeseleccionaGrid Me.DataGrid4
    txtauxFito(4).visible = False 'ID
    For I = 0 To 5
        If I <> 4 Then
            txtauxFito(I).Height = DataGrid4.RowHeight
            txtauxFito(I).visible = b
            txtauxFito(I).Top = alto
        End If
    Next
    Me.cboFitos(0).visible = b
    Me.cboFitos(1).visible = b
    cboFitos(0).Top = alto
    cboFitos(1).Top = alto
    cmdFitos(0).visible = b
    cmdFitos(0).Top = alto
End Sub


Private Sub LLamaLineasDirec(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 5 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    DataGrid6.Enabled = Not b
    
    DeseleccionaGrid Me.DataGrid6
   
    For I = 0 To 1
        
           ' Text3(i).Height = DataGrid6.RowHeight
            Text3(I).visible = b
            Text3(I).Top = alto
        
    Next

End Sub


Private Sub LLamaLineasDirenEvio(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 6 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    DataGrid7.Enabled = Not b
    
    DeseleccionaGrid Me.DataGrid7
   
    For I = 0 To 2
        
           ' Text3(i).Height = DataGrid6.RowHeight
            Text4(I).visible = b
            Text4(I).Top = alto
        
    Next

End Sub






Private Sub LLamaLineasCamposHuertos(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 11 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    
    
    DeseleccionaGrid Me.DataGrid5
    'txtauxFito(4).visible = False 'ID
    For I = 0 To 5
        
        txtauxMarja(I).Height = DataGrid5.RowHeight
        txtauxMarja(I).visible = b
        txtauxMarja(I).Top = alto
    
    Next
     
    cbomarjal.visible = b
End Sub


Private Sub PonerDatosForaGridCamposHuertos(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean


    Limp = True
    If Not ForzarLimpiar Then
        If Not (data8.Recordset Is Nothing) Then
            If Not data8.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For I = 0 To txtauxMarja.Count - 1
            txtauxMarja(I).Text = ""
    
        Next I
     
        
                
    Else
        
        For I = 1 To 2
            BuscaChekc = RecuperaValor("fecaltas|fecbajas|", I)
            BuscaChekc = DBLet(data8.Recordset.Fields(BuscaChekc), "T")
            If BuscaChekc <> "" Then BuscaChekc = Format(CDate(BuscaChekc), "dd/mm/yyyy")
            txtauxMarja(6 + I).Text = BuscaChekc
        Next
        Me.txtauxMarja(9).Text = DBLetMemo(data8.Recordset!observac)
        txtauxMarja(6).Text = DBLet(data8.Recordset!partida, "T")
        BuscaChekc = ""
    End If
End Sub



Private Function InsertarModificarLineaDatosConctacto() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    'codclien,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id FROM scliendp
    InsertarModificarLineaDatosConctacto = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO scliendp (codclien,nombre,dpto,cargo,telefono,ext,movil,maidirec,observa,id) VALUES ("
            SQL = SQL & Text1(0).Text

                    
            For I = 0 To 7 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(txtauxDC(I).Text, "T", "S")
            Next I
            SQL = SQL & ", " & txtauxDC(8).Text & ")"
  
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            'codclien,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id
            SQL = "UPDATE scliendp Set nombre = " & DBSet(txtauxDC(0).Text, "T")
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



Private Function InsertarModificarLineaTelefonia() As Boolean
Dim I As Byte
Dim SQL As String
Dim HaCambiadoFacturaImpresa As Boolean 'Feb 2014

    On Error GoTo EInsertarModificarLinea
    'sclientfno(codclien,IdTelefono,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones)
    InsertarModificarLineaTelefonia = False
    SQL = ""
    HaCambiadoFacturaImpresa = False
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO sclientfno(codclien,IdTelefono,IMEI,SIM,Observaciones,Factura,Detalle,Inactivo,"
            SQL = SQL & "coninternet,coddirec,clienppal,modelo,cuotaminima,puntos,fechaalta,fecharenove,Operador,procedencia"
            SQL = SQL & ",ArtPlazos,PlazosMeses,ImportePlazo,PlazosOrigen) VALUES (" & Text1(0).Text
            
                     
            For I = 0 To 3 '
                SQL = SQL & ", "
                SQL = SQL & DBSet(txtauxTfno(I).Text, "T", "S")
            Next I
            For I = 0 To 3
                SQL = SQL & ", "
                SQL = SQL & Me.chkTelefonia(I).Value
            Next
            For I = 4 To 8 '
                SQL = SQL & ", "
                SQL = SQL & DBSet(txtauxTfno(I).Text, "N", IIf(I >= 7, "N", "S"))

            Next I
            SQL = SQL & "," & DBSet(txtauxTfno(9).Text, "F", "S")
            'Si la fecha renovacion es "" pongo la fecha de alta
            'If Me.txtauxTfno(10).Text = "" Then txtauxTfno(10).Text = txtauxTfno(9).Text feb2020
            SQL = SQL & "," & DBSet(txtauxTfno(10).Text, "F", "S")
            SQL = SQL & "," & cboOperadorTfnnia2(0).ItemData(cboOperadorTfnnia2(0).ListIndex)
            SQL = SQL & "," & cboOperadorTfnnia2(1).ItemData(cboOperadorTfnnia2(1).ListIndex)
            For I = 11 To 14
                If vParamAplic.TelefoniaVtaPlazos Then
                    SQL = SQL & "," & DBSet(txtauxTfno(I).Text, IIf(I = 11, "T", "N"), "S")
                Else
                    SQL = SQL & ",NULL"
                End If
            Next I
            SQL = SQL & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            
            SQL = DBLet(data6.Recordset!Factura, "N")
            If Val(SQL) <> Abs(Me.chkTelefonia(0).Value) Then HaCambiadoFacturaImpresa = True
                        
            'codclien,IdTelefono,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones
            SQL = ""
            For I = 1 To 3  'EL CERO NO
                BuscaChekc = RecuperaValor("IMEI|SIM|Observaciones|", CInt(I))
                SQL = SQL & ", " & BuscaChekc & " = " & DBSet(txtauxTfno(I).Text, "T", "S")
            Next I
            For I = 0 To 3
                BuscaChekc = RecuperaValor("Factura|Detalle|Inactivo|coninternet|", I + 1)
                SQL = SQL & ", " & BuscaChekc & " = " & Me.chkTelefonia(I).Value
            Next
            For I = 4 To 8  'EL CERO NO
                BuscaChekc = RecuperaValor("|||coddirec|clienppal|modelo|cuotaminima|puntos|", CInt(I))
                SQL = SQL & ", " & BuscaChekc & " = " & DBSet(txtauxTfno(I).Text, "N", "S")
            Next I
            
            SQL = SQL & ", fechaalta = " & DBSet(txtauxTfno(9).Text, "F", "S")
            SQL = SQL & ", fecharenove = " & DBSet(txtauxTfno(10).Text, "F", "S")
            SQL = SQL & ", Operador= " & Me.cboOperadorTfnnia2(0).ItemData(cboOperadorTfnnia2(0).ListIndex)
            SQL = SQL & ", procedencia= " & Me.cboOperadorTfnnia2(1).ItemData(cboOperadorTfnnia2(1).ListIndex)
            
            If vParamAplic.TelefoniaVtaPlazos Then
                For I = 11 To 14  ',ArtPlazos,PlazosMeses,ImportePlazoPlazosOrigen
                    BuscaChekc = RecuperaValor("ArtPlazos|PlazosMeses|ImportePlazo|PlazosOrigen|", CInt(I - 10))
                    If I = 12 Then
                        'Cuantos quedam
                        SQL = SQL & ", " & BuscaChekc & " = "
                        If txtauxTfno(I).Text = "" Then
                            SQL = SQL & "NULL"
                        Else
                            SQL = SQL & txtauxTfno(I).Text
                        End If
                    Else
                        SQL = SQL & ", " & BuscaChekc & " = " & DBSet(txtauxTfno(I).Text, IIf(I = 11, "T", "N"), "S")
                    End If
                Next I
            End If
            SQL = Mid(SQL, 2) 'quito la primera coma
            SQL = "UPDATE sclientfno Set " & SQL
            SQL = SQL & " WHERE  IdTelefono = " & DBSet(txtauxTfno(0).Text, "T")
            
            
            
            
            
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaTelefonia = True
        
        If HaCambiadoFacturaImpresa Then
            'Marcamos las facturas como para enviar(o no enviar) segun check
            '#  NUMPEDCL sera para la reimpresion de facturas numpedcl
            '#   0.- SE imprime
            '#   1.- NO. ya que va por email
            SQL = "0"
            If Me.chkTelefonia(0).Value = 0 Then SQL = "1"
            SQL = "UPDATE scafac1 set numpedcl=" & SQL
            SQL = SQL & " WHERE codtipom='FAT' AND observa4=" & DBSet(txtauxTfno(0).Text, "T")
            SQL = SQL & " AND (numfactu,fecfactu) IN (select numfactu,fecfactu from scafac WHERE "
            SQL = SQL & " codclien = " & Me.Text1(0).Text & " and codtipom='FAT')"
            ejecutar SQL, True
            
            
        End If
        
        
        'Si tiene venta plazos telefonia. Compruebo que no ha cambiado nada
        If vParamAplic.TelefoniaVtaPlazos Then
            'ArtPlazos,PlazosMeses,ImportePlazo
            BuscaChekc = ""
            If DBLet(data6.Recordset!artplazos, "T") <> txtauxTfno(11).Text Then BuscaChekc = BuscaChekc & vbCrLf & "Articulo : " & DBLet(data6.Recordset!artplazos, "T") & " // " & txtauxTfno(11).Text
            SQL = ""
            If Not IsNull(data6.Recordset!PlazosMeses) Then SQL = data6.Recordset!PlazosMeses
            If SQL <> txtauxTfno(12).Text Then BuscaChekc = BuscaChekc & vbCrLf & "Plazos restan : " & DBLet(data6.Recordset!PlazosMeses, "T") & " // " & txtauxTfno(12).Text
            
            SQL = ""
            If Not IsNull(data6.Recordset!ImportePlazo) Then SQL = Format(data6.Recordset!ImportePlazo, FormatoImporte)
            If SQL <> txtauxTfno(13).Text Then BuscaChekc = BuscaChekc & vbCrLf & "Imp/mes : " & DBLet(data6.Recordset!ImportePlazo, "T") & " // " & txtauxTfno(13).Text
            
            SQL = ""
            If Not IsNull(data6.Recordset!PlazosMeses) Then SQL = data6.Recordset!PlazosOrigen
            If SQL <> txtauxTfno(14).Text Then BuscaChekc = BuscaChekc & vbCrLf & "Plazos origen: " & DBLet(data6.Recordset!PlazosOrigen, "T") & " // " & txtauxTfno(14).Text
        
            If BuscaChekc <> "" Then
                'GRABAMOS UN LOG
                BuscaChekc = "Telefono: " & txtauxTfno(0).Text & vbCrLf & BuscaChekc
                SQL = "[TELEFONIA] Venta plazos. Cambio en el cliente: " & Text1(0).Text & " " & Text1(1).Text & vbCrLf & "Anterior//Actual" & vbCrLf & BuscaChekc
                Set LOG = New cLOG
                LOG.Insertar 29, vUsu, SQL
                Set LOG = Nothing
                
                
                
                'Acciones comerciales. La 5
                SQL = "NO"
                If ModoFrame2 = 4 Then
                    'Disitnot BD que ahora
                    If DBLet(data6.Recordset!artplazos, "T") <> txtauxTfno(11).Text Then
                        SQL = ""
                        If DBLet(data6.Recordset!artplazos, "T") <> "" Then
                            'Tenia y ahora NO tiene
                            BuscaChekc = "Fin venta plazos " & vbCrLf & BuscaChekc
                        Else
                            'Lo que pone en buscachek es bueno
                        End If
                    End If
                Else
                    SQL = ""
                End If
                If SQL = "" Then
                    
                    SQL = PonerTrabajadorConectado("")
                    SQL = DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & Text1(0).Text & "," & vUsu.CodigoAgente & "," & Val(SQL) & ",2,5,'Otros',"
                    SQL = "INSERT INTO scrmacciones(usuario,fechora,codclien,agente,codtraba,estado,tipo,medio,observaciones) VALUES (" & SQL
                    SQL = SQL & DBSet(BuscaChekc, "T") & ")"
                    If Not ejecutar(SQL, True) Then MsgBox "Error insertando en hco de acciones comerciales", vbExclamation
                        
                End If
                
                
                SQL = ""
                
            End If
            
            
            'INSERTARMENOS LOG DE acciones comerciales
            
            
        End If
        
    Else
        PonerFoco txtauxTfno(1)
    End If
    BuscaChekc = ""
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar datos contacto" & vbCrLf & Err.Description
    BuscaChekc = ""
End Function




Private Function InsertarModificarLineamanipuladorFito() As Boolean
Dim I As Byte
Dim SQL As String


    On Error GoTo EInsertarModificarLinea
    'sclientfno(codclien,IdTelefono,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones)
    InsertarModificarLineamanipuladorFito = False
    SQL = ""
    
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            If Me.cboFitos(0).ListIndex = 1 Then
                I = 2
            Else
                I = 1
            End If
            SQL = "INSERT INTO sclienmani(codclien,tipocarnet,cif,nombre,numcarnet,telefono,id,fcaducidad,Manipuladorprovisional)  VALUES ("
            SQL = SQL & Text1(0).Text & "," & I
            
                     
            For I = 0 To Me.txtauxFito.Count - 1
                If I = 5 Then
                    SQL = SQL & ", " & DBSet(txtauxFito(I).Text, "F", "N")
                Else
                    SQL = SQL & ", " & DBSet(txtauxFito(I).Text, "T", "S")
                End If
            Next I
            I = 0
            If cboFitos(1).ListIndex = 1 Then I = 1
            SQL = SQL & "," & I & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            
            
                        
            'codclien,IdTelefono,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones
            SQL = ""
            
            For I = 1 To 6  'EL CERO NO
                If I <> 5 Then
                    BuscaChekc = RecuperaValor("cif|nombre|numcarnet|telefono||fcaducidad|", CInt(I))
                    SQL = SQL & ", " & BuscaChekc & " = " & DBSet(txtauxFito(I - 1).Text, IIf(I = 6, "F", "T"), "S")
                End If
            Next I
            I = 1
            If Me.cboFitos(0).ListIndex = 1 Then I = 2
            SQL = " tipocarnet = " & I & SQL
            I = Me.cboFitos(1).ListIndex
            SQL = SQL & ", Manipuladorprovisional = " & I
            SQL = "UPDATE sclienmani Set " & SQL
            SQL = SQL & " WHERE  id = " & Data7.Recordset!Id
            SQL = SQL & " AND  codclien = " & DBSet(Text1(0).Text, "T")
            
            
            
            
            
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineamanipuladorFito = True
    Else
        PonerFoco txtauxTfno(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar datos manipulador fitosanitarios" & vbCrLf & Err.Description
End Function


Private Function InsertarModificarLineaCamposhuertos() As Boolean
Dim I As Byte
Dim SQL As String


    On Error GoTo EInsertarModificarLinea
    'sclientfno(codclien,IdTelefono,IMEI,SIM,Factura,Detalle,Inactivo,Observaciones)
    InsertarModificarLineaCamposhuertos = False
    SQL = ""
    
    
    If Not DatosOkLinea Then
       
        Exit Function
    End If
    
    
            
            
    BuscaChekc = "id|poligono|parcela|recintos|supsigpa|supderec|partida|fecaltas|fecbajas|observac|"
                        
    kCampo = 0
    If ModificaLineas = 2 Then kCampo = 1
            
    For I = kCampo To Me.txtauxMarja.Count - 1
        SQL = SQL & ", "
        If ModificaLineas = 2 Then SQL = SQL & RecuperaValor(BuscaChekc, CInt(I + 1)) & " = "
            
        If I < 6 Then
            SQL = SQL & DBSet(txtauxMarja(I), "N")
        ElseIf I = 7 Or I = 8 Then
            SQL = SQL & DBSet(txtauxMarja(I), "F", "S")
        Else
            SQL = SQL & DBSet(txtauxMarja(I), "T", "S")
        End If
    Next I
            
            
    If ModificaLineas = 1 Then
        SQL = Text1(0).Text & SQL
        BuscaChekc = Replace(BuscaChekc, "|", ",")
        BuscaChekc = Mid(BuscaChekc, 1, Len(BuscaChekc) - 1) 'quitamos la ultmia coma
        SQL = "INSERT INTO sclienhuertos(codclien," & BuscaChekc & ") VALUES (" & SQL & ")"
    
    Else
        SQL = Mid(SQL, 2)
        SQL = "UPDATE sclienhuertos SET " & SQL
        SQL = SQL & " WHERE  id = " & data8.Recordset!Id
        SQL = SQL & " AND  codclien = " & DBSet(Text1(0).Text, "T")
    End If
    If SQL <> "" Then
        
        conn.Execute SQL
        InsertarModificarLineaCamposhuertos = True
        
        
        'Voy a tratar el combo, por si lo que ha puesto NO estaba entodavia
        
        SQL = ""
        For NumRegElim = 1 To cbomarjal.ListCount
            If cbomarjal.List(NumRegElim) = Me.txtauxMarja(6).Text Then
                SQL = "X"
                Exit For
            End If
        Next
        If SQL = "" Then Cargacbomarjal
            
       
                
    Else
        PonerFoco txtauxTfno(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar datos manipulador fitosanitarios" & vbCrLf & Err.Description
End Function






Private Sub ToolbarAux_ButtonClick(index As Integer, ByVal Button As MSComctlLib.Button)

    If Modo <> 2 And Modo < 5 Then Exit Sub

    If Modo >= 5 And ModificaLineas > 0 Then Exit Sub
    
    Select Case index
    Case 0
    
        
        'Departamentos
        If Button.index = 3 Then
            BotonEliminarLinea
        Else
            PonerModo 5
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If


    Case 1
        'Direcciones de envio
        If Button.index = 3 Then
            BotonEliminarLineaDirEnvio
        Else
            PonerModo 6
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If


    Case 2
        'CONTACTOS
        
        If Me.cboCargo.ListCount <= 0 Then CargaComboCargos
        If Button.index = 3 Then
            BotonEliminarLineaContacto
        Else
            PonerModo 7
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If

    Case 3
        'MANIPULADOR
        If Button.index = 3 Then
            BotonEliminarManipulador
        Else
            PonerModo 10
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If

    Case 4
        'RENTING
        
         'Renting, si no tiene establecido el periodo de facturacion de renting, tendremos que avisarlo y NO dejarle pasar
        If Me.cboFraRenting.ListIndex < 0 Then
            MsgBox "No tiene establecido el periodo de facturaci�n de " & RentingLB, vbExclamation
            Me.SSTab1.Tab = 0
            Exit Sub
        End If
        
        
        If Button.index = 3 Then
            BotonEliminarRenting
        Else
            PonerModo 8
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If

    Case 5
        'TELEFONIA
        If Button.index = 3 Then
            BotonEliminarTelefono
        Else
            PonerModo 9
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If

    Case 6
        'HUERTOS
        If cbomarjal.Tag = -1 Then Cargacbomarjal
            
        If Button.index = 3 Then
            BotonEliminarTelefono
        Else
            PonerModo 11
            If Button.index = 1 Then
                'A�ADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If

    End Select
    
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento Button.index - 1
End Sub

Private Sub txtauxDC_GotFocus(index As Integer)
    If index <> 7 Then ConseguirFoco txtauxDC(index), 3
End Sub

Private Sub txtauxDC_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If index <> 7 Then
            KeyAscii = 0
            SendKeys "{tab}"
        Else
            PonerFocoBtn cmdAceptar
        End If
    End If
End Sub

Private Sub txtauxDC_LostFocus(index As Integer)
    'Si quisieramos comprobar algo
    txtauxDC(index).Text = Trim(txtauxDC(index).Text)
    If index = 7 Then PonerFocoBtn cmdAceptar
End Sub


Private Sub BotonEliminarLineaContacto()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String
Dim I As Integer

    If Data4.Recordset.EOF Then Exit Sub
    If Data4.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    
    cad = "�Seguro que desea eliminar el contacto?"
    cad = cad & vbCrLf & "Nombre:  " & Data4.Recordset!Nombre
    cad = cad & vbCrLf & "Departamento:  " & DBLet(Data4.Recordset!Dpto, "T")
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data4.Recordset.AbsolutePosition
        Data4.Recordset.Delete
        CargaLineas True, 0
        
        'PonerDatosForaGridContacto False
            
        ModificaLineas = 0
        'PonerModoFrame 0, 2
        
        
        
        
        
        
        
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data4.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Private Sub BotonEliminarRenting()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String


    If Data5.Recordset.EOF Then Exit Sub
    If Data5.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    
    cad = "�Seguro que desea eliminar el elemento ?"
    cad = cad & vbCrLf & "ID:  " & Data5.Recordset!Id
    If Not IsNull(Data5.Recordset!CodDirec) Then cad = cad & vbCrLf & "Departamento:  " & DBLet(Data5.Recordset!CodDirec, "T") & " " & DBLet(Data5.Recordset!nomdirec, "T")
    cad = cad & vbCrLf & "Referencia:  " & Data5.Recordset!Referencia
    cad = cad & vbCrLf & "Importe:  " & Data5.Recordset!Importe
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data5.Recordset.AbsolutePosition
        cad = "DELETE FROM sclienrenting where codclien = " & Text1(0).Text & " AND ID= " & Data5.Recordset!Id
        conn.Execute cad
        CargaLineas True, 1
        PonerDatosForaGridRent False
            
        ModificaLineas = 0
        
        If NumRegElim > 0 Then Data5.Recordset.Move NumRegElim
        ModificaLineas = 0
        
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data5.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub




Private Sub BotonEliminarTelefono()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String


    If data6.Recordset.EOF Then Exit Sub
    If data6.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
       
    'Deberiamos comprobar SI puede eliminar el telefono
    cad = DevuelveDesdeBD(conAri, "count(*)", "tel_cab_factura", "telefono", CStr(data6.Recordset!idtelefono), "T")
    If cad <> "" Then
        If Val(cad) > 0 Then
            MsgBox "Existen facturas relacionadas con este numero", vbExclamation
            Exit Sub
        End If
    End If
       
       
       
       
    '------------------------------
       
    ModificaLineas = 3 'Eliminar
    
    cad = "�Seguro que desea eliminar el tel�fono " & data6.Recordset!idtelefono & "?"
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = data6.Recordset.AbsolutePosition
        

        
        cad = "DELETE FROM sclientfnoCuotas where  IdTelefono= " & data6.Recordset!idtelefono
        conn.Execute cad
        
        cad = "DELETE FROM sclientfno where  IdTelefono= " & data6.Recordset!idtelefono
        conn.Execute cad
        CargaLineas True, 2
        PonerDatosForaGridTfno False
            
        ModificaLineas = 0
        PonerModoFrame 0, 9
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        data6.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Private Sub BotonEliminarManipulador()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String


    If Data7.Recordset.EOF Then Exit Sub
    If Data7.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    
    cad = "�Seguro que desea eliminar al autorizado?"
    cad = cad & vbCrLf & "ID :  " & Data7.Recordset!Id & "    - " & DBLet(Data7.Recordset!CIF, "T")
    
    cad = cad & vbCrLf & "Nombre:  " & DBLet(Data7.Recordset!Nombre, "T")
    cad = cad & vbCrLf & "Carnet:  " & Data7.Recordset!Tipo
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data7.Recordset.AbsolutePosition
        cad = "DELETE FROM sclienmani where codclien = " & Text1(0).Text & " AND ID= " & Data7.Recordset!Id
        conn.Execute cad
        CargaLineas True, 3
        'PonerDatosForaGridRent False
            
        ModificaLineas = 0
        PonerModoFrame 0, 10
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data7.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub



Private Sub BotonEliminarHuertos()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String


    If data8.Recordset.EOF Then Exit Sub
    If data8.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    
    cad = "�Seguro que desea eliminar al campo?"
    cad = cad & vbCrLf & "ID :  " & data8.Recordset!Id
    
    cad = cad & vbCrLf & "Campo:  " & DataGrid5.Columns(1).Text & " - " & DataGrid5.Columns(2).Text & " - " & DataGrid5.Columns(3).Text
    cad = cad & vbCrLf & "partida:  " & DBLet(data8.Recordset!partida, "T")
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = data8.Recordset.AbsolutePosition
        cad = "DELETE FROM sclienhuertos where codclien = " & Text1(0).Text & " AND ID= " & data8.Recordset!Id
        conn.Execute cad
        CargaLineas True, 4
        
            
        ModificaLineas = 0
        PonerModoFrame 0, 11
    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        data8.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub





Private Sub CargaComboTipoCliente()
    CargarCombo_Tabla Me.cboTipocliente, "stipclien", "tipclien", "descclien"
End Sub

Private Sub CargaComboFrarRenting()
    cboFraRenting.Clear
    cboFraRenting.AddItem "Mensual"
    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 1

    cboFraRenting.AddItem "Trimestral"
    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 3

    cboFraRenting.AddItem "Semestral"
    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 6

    cboFraRenting.AddItem "Anual"
    cboFraRenting.ItemData(cboFraRenting.NewIndex) = 12
    
End Sub


Private Sub CargaComboPais()
    cboPais.Clear
    If Not vParamAplic.ContabilidadNueva Then Exit Sub
    
    cboPais.AddItem "ESPA�A  (ES)"
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from paises where codpais <>'ES' and nompais<>'' order by nompais", ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboPais.AddItem miRsAux!nompais & "   (" & miRsAux!codpais & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Sub CargaComboManipulador()

    cboManipulador.Clear
    cboManipulador.AddItem "Sin carnet"
    cboManipulador.ItemData(cboManipulador.NewIndex) = 0

    cboManipulador.AddItem "B�sico"
    cboManipulador.ItemData(cboManipulador.NewIndex) = 1

    cboManipulador.AddItem "Cualificado"
    cboManipulador.ItemData(cboManipulador.NewIndex) = 2

End Sub

Private Sub CargaComboAseguradora()
    On Error Resume Next
    CargarCombo_Tabla cboTipoASeg, "stipocredito", "codTipoCredito", "nomTipoCredito"
    
End Sub

Private Sub CargaComboPrioridad()
    CargarCombo_Tabla Me.cboPrioridad, "sprioridades", "Prioridad", "descripcion", , , "Prioridad "
    If cboPrioridad.ListCount = 0 Then cboPrioridad.AddItem "NORMAL"
End Sub



Private Sub CargaComboTfnos_()
    On Error Resume Next
    CargarCombo_Tabla cboOperadorTfnnia2(0), "stfnoOperador", "codoperador", "nombre"
    CargarCombo_Tabla cboOperadorTfnnia2(1), "tel_procedencias", "codproce", "Descripcion"
End Sub

'Comprobaremos que ha cambiado los campos que enlazan con conta. nombre nif.....
Private Function HayQueActualizarenContabilidad() As Boolean
Dim QueCampos As String
Dim mTag As CTag
Dim I As Integer
Dim fin As Boolean
Dim txt As String
Dim Valor
    HayQueActualizarenContabilidad = False
    CambiaCCC_Ariadna = False
    If Text1(35).Text = "" Or Text2(35).Text = "" Then Exit Function


    'Para CCC en aopliaciones ARIADNA
    If vParamAplic.ComprobarBancoRestoAplicaciones Then
        txt = Format(DBLet(Data1.Recordset.Fields!codbanco, "N"), "0000") & Format(DBLet(Data1.Recordset.Fields!codsucur, "N"), "0000")
        txt = txt & Right("00" & DBLet(Data1.Recordset.Fields!digcontr), 2)
        txt = txt & Right(String(10, "0") & DBLet(Data1.Recordset.Fields!cuentaba), 10)
        'Nov 2013.
        txt = DBLet(Data1.Recordset!Iban, "T") & txt
        QueCampos = Me.Text1(56).Text & Me.Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text
        If txt <> QueCampos Then CambiaCCC_Ariadna = True
    End If
    



    'Vere si el campo que habia al que hay ha cambiado
    QueCampos = "0|1|3|4|5|6|7|31|32|33|34|"
    'Marzo 2012, operaciones aseguradas
    QueCampos = QueCampos & "50|48|47|41|43|23|"
    'Mayo 2012, la fecha baja credito    y IBAN
    QueCampos = QueCampos & "53|56|"
    If vParamAplic.ContabilidadNueva Then QueCampos = QueCampos & "60|"   'PAIS
    
    fin = False
    Set mTag = New CTag
    
    
    
    
    While Not fin
      I = InStr(1, QueCampos, "|")
      'NO puede ser ccero
      txt = Mid(QueCampos, 1, I - 1)
      QueCampos = Mid(QueCampos, I + 1)
      I = CInt(txt)
      mTag.Cargar Text1(I)
      'TIENE QUE ESTAR CARGADO  If mTag.Cargado Then

                'Debug.Print mTag.columna
                        
                        
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

                If Text1(I).Text <> txt Then
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
Dim I As Integer

    If Data4.Recordset Is Nothing Then Exit Sub
    If Data4.Recordset.EOF Then Exit Sub

    cboCargo.ListIndex = -1
    For I = 1 To cboCargo.ListCount - 1
        If cboCargo.List(I) = UCase(DBLet(Data4.Recordset!cargo, "T")) Then
            cboCargo.ListIndex = I
            Exit For
        End If
    Next
End Sub




Private Sub LLamaLineasRenting(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Integer

    ModificaLineas = xModo
    
    b = Modo = 8 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas


    DeseleccionaGrid Me.DataGrid2
    
    DataGrid2.Enabled = Not b
    
    For I = 0 To 7
        If I < 3 Then
            cmdRenting(I).visible = b
            cmdRenting(I).Top = alto
            cmdRenting(I).Height = DataGrid2.RowHeight
        End If
        txtauxRent(I).Height = DataGrid2.RowHeight
        txtauxRent(I).visible = b
        txtauxRent(I).Top = alto
             
        If I = 0 Or I = 2 Then
            BloquearTxt txtauxRent(I), True, I = 0 And ModificaLineas = 1
        End If
    Next I
    imgBuscar(24).Enabled = b
    
    
    
    For I = 8 To 11
   
        If I = 8 Or I = 10 Then
            BloquearTxt txtauxRent(I), Not b, False
            
        Else
            BloquearTxt txtauxRent(I), True, False
        End If
        
        
    Next I
    
End Sub


Private Sub txtauxFito_GotFocus(index As Integer)
    ConseguirFoco txtauxFito(index), 3
End Sub

Private Sub txtauxFito_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If index <> 3 Then
            KeyAscii = 0
            SendKeys "{tab}"
        Else
            'Despues del importe
            PonerFocoBtn cmdAceptar
        End If
    End If
End Sub

Private Sub txtauxFito_LostFocus(index As Integer)
    If Modo <> 10 Then Exit Sub
    'If Index = 2 Then If Not PonerFormatoEntero(txtauxFito(Index)) Then txtauxFito(Index).Text = ""
    If index = 5 Then PonerFormatoFecha txtauxFito(index)
    If index = 0 Then
        If txtauxFito(index).Text <> "" Then
            txtauxFito(index).Text = UCase(txtauxFito(index).Text)
            If Not Comprobar_NIF(txtauxFito(index)) Then MsgBox "El NIF parace incorrecto. ", vbExclamation
            'ManipuladortipoCarnet ManipuladorNumCarnet ManipuladorFecCaducidad
            If ModificaLineas = 1 Then
                BuscaChekc = "concat(coalesce(ManipuladortipoCarnet ,''),'|',coalesce(ManipuladorNumCarnet,''),'|',coalesce(ManipuladorFecCaducidad,''),'|'"
                BuscaChekc = BuscaChekc & ",coalesce(nomclien,''),'|')"
                BuscaChekc = DevuelveDesdeBD(conAri, BuscaChekc, "sclien", "nifclien", txtauxFito(index).Text, "T")
                If BuscaChekc = "" Then BuscaChekc = "0|"
                'A28226256
                If RecuperaValor(BuscaChekc, 1) > 0 Then
                    txtauxFito(1).Text = RecuperaValor(BuscaChekc, 4)
                    txtauxFito(2).Text = RecuperaValor(BuscaChekc, 2)
                    txtauxFito(5).Text = Format(RecuperaValor(BuscaChekc, 3), "dd/mm/yyyy")
                    
                    Me.cboFitos(0).ListIndex = CInt(RecuperaValor(BuscaChekc, 1)) - 1
                End If
            End If
            If txtauxFito(2).Text = "" Then txtauxFito(2).Text = txtauxFito(index).Text
        End If
    End If
End Sub

Private Sub txtauxMarja_GotFocus(index As Integer)
    If index <> 9 Then ConseguirFoco txtauxMarja(index), 3
End Sub


Private Sub txtauxMarja_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If index <> 9 Then
            KeyAscii = 0
            SendKeys "{tab}"
        
        End If
    End If
End Sub

Private Sub txtauxMarja_LostFocus(index As Integer)
    txtauxMarja(index).Text = Trim(txtauxMarja(index).Text)
    Select Case index
    Case 1, 3
           'txtauxRent
           BuscaChekc = ""
           If txtauxMarja(index).Text <> "" Then
              If Not PonerFormatoEntero(txtauxMarja(index)) Then txtauxMarja(index).Text = ""
          End If
        
       
    Case 7, 8
          If txtauxMarja(index).Text <> "" Then PonerFormatoFecha txtauxMarja(index)
    
    Case 4, 5
          
          If Not PonerFormatoDecimal(txtauxMarja(index), 3) Then txtauxMarja(index).Text = ""
    Case 9
        PonerFocoBtn cmdAceptar
        DoEvents
        PonerFocoBtn cmdAceptar
    End Select
End Sub

Private Sub txtauxRent_GotFocus(index As Integer)
    ConseguirFoco txtauxRent(index), 3
End Sub

Private Sub txtauxRent_KeyPress(index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
        If index <> 10 Then
            KeyAscii = 0
            SendKeys "{tab}"
        Else
            'Despues del importe
            PonerFocoBtn cmdAceptar
            
        End If
    End If
End Sub

Private Sub txtauxRent_LostFocus(index As Integer)
      txtauxRent(index).Text = Trim(txtauxRent(index).Text)
      Select Case index
      Case 1
             'txtauxRent
             BuscaChekc = ""
             If txtauxRent(index).Text <> "" Then
                If PonerFormatoEntero(txtauxRent(index)) Then
                    BuscaChekc = "codclien = " & Text1(0).Text & " AND coddirec "
                    BuscaChekc = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", BuscaChekc, txtauxRent(index).Text, "N")
                End If
            End If
            Me.txtauxRent(2).Text = BuscaChekc
            If BuscaChekc = "" Then
                If Me.txtauxRent(index).Text <> "" Then
                    txtauxRent(index).Text = ""
                    
                End If
            End If
         
      Case 4, 6
            If txtauxRent(index).Text <> "" Then PonerFormatoFecha txtauxRent(index)
      Case 5
            If PonerFormatoEntero(txtauxRent(index)) Then
                'Si la fecha es correcta
                If Me.txtauxRent(4).Text <> "" Then
                    'n cutoas
                    txtauxRent(6).Text = Format(DateAdd("m", CInt(txtauxRent(5).Text), CDate(Me.txtauxRent(4).Text)))
                    'menos un dia
                    txtauxRent(6).Text = Format(DateAdd("d", -1, CDate(Me.txtauxRent(6).Text)))
                End If
            End If
        
      Case 7
            If Not PonerFormatoDecimal(txtauxRent(index), 3) Then txtauxRent(index).Text = ""
            
      Case 8
            'tipo de contrato
            BuscaChekc = ""
            If txtauxRent(index).Text <> "" Then
                BuscaChekc = DevuelveDesdeBD(conAri, "nomtipco", "stipco", "codtipco", txtauxRent(index).Text, "T")
                If BuscaChekc = "" Then
                    MsgBox "No existe el tipo de contrato", vbExclamation
                    txtauxRent(index).Text = ""
                    PonerFoco txtauxRent(index)
                End If
            End If
            txtauxRent(9).Text = BuscaChekc
      End Select
      
      BuscaChekc = ""
End Sub




Private Function InsertarModificarLineaRenting() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    'codclien,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa,id FROM scliendp
    InsertarModificarLineaRenting = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            SQL = "INSERT INTO sclienrenting(codclien,id,coddirec,referencia,fecalta,numcuotas,fecbaja,importe,codtipco, obser,ultfec) VALUES ("
            SQL = SQL & Text1(0).Text

                    
            For I = 0 To 11
                If I <> 2 And I <> 9 Then SQL = SQL & ", " 'el 2 no mete en el sql
                If I = 0 Or I = 1 Or I = 5 Then
                    'ENTERO
                    SQL = SQL & DBSet(txtauxRent(I).Text, "N", "S")
                Else
                    If I = 4 Or I = 6 Or I = 11 Then
                        'FECHA
                        SQL = SQL & DBSet(txtauxRent(I).Text, "F", "S")
                    Else
                        If I = 7 Then
                            'DECIMAL
                            SQL = SQL & DBSet(txtauxRent(I).Text, "N", "N")
                        Else
                            'TEXTO
                            If I <> 2 And I <> 9 Then SQL = SQL & DBSet(txtauxRent(I).Text, "T", "S") 'el nomdepartamento NO VA AQUI
                        End If
                    End If
                End If
            Next
                
                
            
            SQL = SQL & ")"
  
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            '(codclien,id,coddirec,referencia,fecalta,numcuotas,fecbaja,ultfec,importe) VALUES ("
            '
            SQL = "UPDATE sclienrenting Set coddirec = " & DBSet(txtauxRent(1).Text, "N", "S")
            SQL = SQL & ", referencia = " & DBSet(txtauxRent(3).Text, "T", "N")
            SQL = SQL & ", fecalta = " & DBSet(txtauxRent(4).Text, "F", "N")
            SQL = SQL & ", numcuotas = " & DBSet(txtauxRent(5).Text, "N", "N")
            SQL = SQL & ", fecbaja = " & DBSet(txtauxRent(6).Text, "F", "N")
            'SQL = SQL & ", ultfec  = " & DBSet(txtauxRent(11).Text, "F", "S")
            SQL = SQL & ", importe= " & DBSet(txtauxRent(7).Text, "N", "N")
            SQL = SQL & ", codtipco= " & DBSet(txtauxRent(8).Text, "T", "N")
            SQL = SQL & ", obser = " & DBSet(txtauxRent(10).Text, "T", "S")
            SQL = SQL & " WHERE codclien =" & (Text1(0).Text) & " AND "
            SQL = SQL & " id =" & (txtauxRent(0).Text)
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaRenting = True
    Else
        PonerFoco txtauxRent(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar datos " & RentingLB & vbCrLf & Err.Description
End Function



Private Sub ActualizarAsegurados_()
Dim Aux As String



    'numpoliz fecsolic credisol fecconce credicon forpa ctabanco
    Aux = "UPDATE cuentas set "
    
    'NULO
    Aux = Aux & " numpoliz =" & DBSet(Text1(50), "T", "S")
    Aux = Aux & ",fecsolic =" & DBSet(Text1(48), "F", "S")
    Aux = Aux & ",credisol =" & DBSet(Text1(47), "N", "S")
    Aux = Aux & ",fecconce =" & DBSet(Text1(41), "F", "S")
    Aux = Aux & ",credicon =" & DBSet(Text1(43), "N", "S")
    Aux = Aux & ",fecbajcre =" & DBSet(Text1(53), "F", "S")
    

    
    'Aux = Aux & ",ctabanco="
    Aux = Aux & " WHERE codmacta = '" & Text1(35).Text & "'"
    
    On Error Resume Next
    ConnConta.Execute Aux
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


Private Function DevuelveBusquedaTelefonia() As String
Dim I As Byte
Dim EsLike As Boolean
Dim Aux As String
Dim J As Integer

    DevuelveBusquedaTelefonia = ""
    J = IIf(vParamAplic.TelefoniaVtaPlazos, 14, 10)
    For I = 0 To J
        Me.txtauxTfno(I).Text = Trim(Me.txtauxTfno(I).Text)
        If Me.txtauxTfno(I).Text <> "" Then
        
            
            'Los textos
            If I < 4 Then
                Aux = RecuperaValor("IdTelefono|IMEI|SIM|Observaciones|", I + 1)
                DevuelveBusquedaTelefonia = DevuelveBusquedaTelefonia & " AND " & Aux
                Aux = txtauxTfno(I).Text
            
                If InStr(1, Aux, "*") > 0 Then
                    Aux = " like " & DBSet(Replace(Me.txtauxTfno(I).Text, "*", "%"), "T")
                Else
                    Aux = " = " & DBSet(Me.txtauxTfno(I).Text, "T")
                End If
            ElseIf I < 9 Then
                
                If SeparaCampoBusqueda("N", RecuperaValor("sclienTfno.coddirec|sclienTfno.clienppal|modelo|cuotaminima|puntos|", I - 3), txtauxTfno(I).Text, Aux) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
            ElseIf I < 11 Then
                'FECHA
                If SeparaCampoBusqueda("F", "sclienTfno.fechaalta", txtauxTfno(I).Text, Aux) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
            Else
                'Venta plazos ,ArtPlazos,PlazosMeses,ImportePlazo
                If SeparaCampoBusqueda(IIf(I = 11, "T", "N"), RecuperaValor("ArtPlazos|sclienTfno.PlazosMeses|ImportePlazo|PlazosOrigen|", I - 10), txtauxTfno(I).Text, Aux) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
            End If
            If Aux <> "" Then DevuelveBusquedaTelefonia = DevuelveBusquedaTelefonia & Aux
        End If
    Next
    
    For I = 0 To 3
        If Me.chkTelefonia(I).Value = 1 Then
            Aux = RecuperaValor("Factura|Detalle|Inactivo|coninternet|", I + 1)
            DevuelveBusquedaTelefonia = DevuelveBusquedaTelefonia & " AND " & Aux & " = 1"
        End If
    Next
    
    If Me.cboOperadorTfnnia2(0).ListIndex >= 0 Then DevuelveBusquedaTelefonia = DevuelveBusquedaTelefonia & " AND OPERADOR = " & cboOperadorTfnnia2(0).ItemData(cboOperadorTfnnia2(0).ListIndex)
    If Me.cboOperadorTfnnia2(1).ListIndex >= 0 Then DevuelveBusquedaTelefonia = DevuelveBusquedaTelefonia & " AND procedencia = " & cboOperadorTfnnia2(1).ItemData(cboOperadorTfnnia2(1).ListIndex)
        
    
    If DevuelveBusquedaTelefonia <> "" Then
        DevuelveBusquedaTelefonia = Mid(DevuelveBusquedaTelefonia, 5) 'quitamos el primer and
    
    
    End If
End Function


Private Sub txtauxTfno_GotFocus(index As Integer)
    If index <> 3 Then ConseguirFoco txtauxTfno(index), 3
End Sub

Private Sub txtauxTfno_KeyPress(index As Integer, KeyAscii As Integer)
    If index <> 3 Then KEYpress KeyAscii
           
End Sub

Private Sub txtauxTfno_LostFocus(index As Integer)
Dim C As String
    
    Select Case index
    Case 3
        'KEYpress 13  'son textos
        'PonerFocoBtn Me.cmdAceptar
    'ElseIf Index > 3 And Index < 9 Then
     Case 3 To 8
        'Cliente, departamento
        
        If Me.txtauxTfno(index).Text <> "" Then
            
            If Modo <> 1 Then
                BuscaChekc = ""
                If Not IsNumeric(txtauxTfno(index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                Else
                    If index < 7 Then
                        If index = 4 Then
                            BuscaChekc = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", "codclien=" & Text1(0).Text & " AND coddirec", Me.txtauxTfno(index).Text)
                        ElseIf index = 5 Then
                            BuscaChekc = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Me.txtauxTfno(index).Text)
                        Else
                            BuscaChekc = DevuelveDesdeBD(conAri, "descripcion", "stfnoModel", "codmodelo", Me.txtauxTfno(index).Text)
                        End If
                        If BuscaChekc = "" Then
                            MsgBox "No existe ningun dato(telefonia:" & index & ") en la BD con ese valor", vbExclamation
                            txtauxTfno(index).Text = ""
                        End If
                    Else
                        'El 8 nada y el
                        BuscaChekc = ""
                    End If
                End If

                If index < 7 Then
                    If BuscaChekc = "" Then PonerFoco Me.txtauxTfno(index)
                    Me.Text5(index).Text = BuscaChekc
                End If
                BuscaChekc = ""
                
            End If
        Else
            If index < 7 Then Text5(index).Text = ""
        End If
    'ElseIf Index >= 9 And Index <= 10 Then
    Case 9, 10
        If Modo > 1 Then
            BuscaChekc = Trim(Me.txtauxTfno(index).Text)
            If BuscaChekc <> "" Then
                If Not EsFechaOK(BuscaChekc) Then
                    MsgBox "Fecha incorrecta: " & txtauxTfno(index).Text, vbExclamation
                    txtauxTfno(index).Text = ""
                    PonerFoco txtauxTfno(index)
                Else
                    txtauxTfno(index).Text = BuscaChekc
                End If
                BuscaChekc = ""
            End If
        End If
    Case 12, 13, 11, 14
        If Me.txtauxTfno(index).Text <> "" And Modo <> 1 Then
            If index = 12 Or index = 14 Then
                If Not PonerFormatoEntero(txtauxTfno(index)) Then Me.txtauxTfno(index).Text = ""
            ElseIf index = 13 Then
                If Not PonerFormatoDecimal(txtauxTfno(index), 2) Then Me.txtauxTfno(index).Text = ""
            Else
                'codartic
                C = "codartic"
                BuscaChekc = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Me.txtauxTfno(index).Text, "T", C)
                If BuscaChekc = "" Then
                    MsgBox "No existe el articulo", vbExclamation
                    Me.txtauxTfno(index).Text = ""
                Else
                    Me.txtauxTfno(index).Text = C
                End If
                Text5(11).Text = BuscaChekc
            End If
            If Me.txtauxTfno(index).Text = "" Then PonerFoco txtauxTfno(index)
        Else
            If index = 11 Then Text5(11).Text = ""
        End If
    End Select
    
    
    
End Sub



Private Sub UpdatearNomClien()
Dim I As Byte
    

    For I = 1 To 9
        CadenaConsulta = RecuperaValor("scaalb|scaavi|scafac|scaped|scapedrma|scapre|schalb|schped|schpre|", CInt(I))
        lblIndicador.Caption = "Actualiza " & CadenaConsulta
        lblIndicador.Refresh
        CadenaConsulta = "UPDATE " & CadenaConsulta & " SET nomclien=" & DBSet(Text1(1).Text, "T")
        CadenaConsulta = CadenaConsulta & " WHERE codclien = " & Text1(0).Text
        conn.Execute CadenaConsulta
        Screen.MousePointer = vbHourglass
        DoEvents
    Next
    
    CadenaConsulta = "CLI.  " & Format(Text1(0).Text, "000000") & "-> " & Text1(1).Text
    Set LOG = New cLOG
    LOG.Insertar 21, vUsu, CadenaConsulta
    Set LOG = Nothing
End Sub



Private Sub ProcesarCarpetaImagenes()
Dim C As String
Dim MiNombre As String

    On Error GoTo EProcesarCarpetaImagenes
    C = App.Path & "\ImgFicFT"
    If Dir(C, vbDirectory) = "" Then
        MkDir C
    Else
        On Error Resume Next
        If Dir(C & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
            MiNombre = Dir(C & "\*.*")   ' Recupera la primera entrada.
            Do While MiNombre <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If MiNombre <> "." And MiNombre <> ".." Then
                    Kill C & "\" & MiNombre
               End If
               MiNombre = Dir   ' Obtiene siguiente entrada.
            Loop
        End If
        On Error GoTo EProcesarCarpetaImagenes
    
    End If
    
    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
End Sub



Private Function CargarIMG(Archivo As String) As Boolean
    On Error Resume Next
    Screen.MousePointer = vbHourglass
'    lblCarga2.Caption = "Cargando ..."
'    lblCarga2.Refresh
    CargarIMG = False
    
    If InStr(1, Archivo, ".pdf") <> 0 Then
        Me.Image1.Picture = LoadPicture(App.Path & "\pdf.dat")
    ElseIf InStr(1, Archivo, ".png") <> 0 Then
        Me.Image1.Picture = LoadPicture(App.Path & "\png.dat")
    ElseIf InStr(1, Archivo, ".tif") <> 0 Then
        Me.Image1.Picture = LoadPicture(App.Path & "\tif.dat")
    Else
        Me.Image1.Picture = LoadPicture(Archivo)
    
    End If

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    Else
        CargarIMG = True
    End If
'    lblCarga2.Caption = lblCarga2.Tag
    Screen.MousePointer = vbDefault
End Function



Private Sub ImprimirImagen()

                
                  
    LanzaVisorMimeDocumento Me.hwnd, Me.lw1.SelectedItem.SubItems(2)
                

End Sub


'VinculaLW --> Normal
'    -> False DNI fitosanitarios
Private Sub CargarArchivos(BorrarAnteriores As Boolean, IndiceSituar As Long, VinculaLW As Boolean)
Dim C As String
Dim L As Long

        
    If VinculaLW Then lw1.ListItems.Clear
    If BorrarAnteriores Then ProcesarCarpetaImagenes
    


    C = "Select * from sfichdocs where codclien=" & DBSet(Text1(0).Text, "N") & " ORDER BY TipoDoc desc, orden"


   
    BuscaChekc = ""
    Adodc1IMG.ConnectionString = conn
    Adodc1IMG.RecordSource = C
    Adodc1IMG.Refresh

    If Adodc1IMG.Recordset.EOF Then
        'NO HAY NINGUNA
        CargarIMG ""
    Else
        'LEEMOS LAS IMAGENES
'        InsertandoImg = True
        While Not Adodc1IMG.Recordset.EOF
            L = Adodc1IMG.Recordset!Codigo

            C = App.Path & "\ImgFicFT\" & L
            If DBLet(Adodc1IMG.Recordset!Docum) <> "0" Then
                C = App.Path & "\ImgFicFT\" & Adodc1IMG.Recordset!Docum
            End If
            If Dir(C) <> "" Then
                If VinculaLW Then AnyadirAlListview C
                C = ""
            Else
           
                If LeerBinary(Adodc1IMG.Recordset!campo, C) Then
                    If VinculaLW Then AnyadirAlListview C
                    C = ""
                End If
            End If
            
            If C = "" And VinculaLW Then
                'Se ha a�adido a listview
                If IndiceSituar > 0 Then
                                        'ULTIMO A�ADIDO
                    If L = IndiceSituar Then BuscaChekc = lw1.ListItems.Count
                
                End If
            End If
            
            Adodc1IMG.Recordset.MoveNext
        Wend
    
        
        
'        InsertandoImg = False
        If VinculaLW Then
            If lw1.ListItems.Count > 0 Then
                L = 1
                If BuscaChekc <> "" Then L = CLng(BuscaChekc)
                CargarIMG lw1.ListItems(L).SubItems(2)
                Set lw1.SelectedItem = lw1.ListItems(L)
            End If
            
        End If
    End If

    Set Adodc1IMG.Recordset = Nothing
End Sub



Private Sub AnyadirAlListview(vpaz As String)
Dim IT
    If Dir(vpaz, vbArchive) = "" Then
        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
      
        Set IT = lw1.ListItems.Add()
        IT.Text = Me.Adodc1IMG.Recordset!Orden '

        IT.SubItems(1) = Me.Adodc1IMG.Recordset.Fields(3)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = vpaz
        IT.SubItems(3) = Me.Adodc1IMG.Recordset.Fields(0)
        IT.SubItems(4) = Me.Adodc1IMG.Recordset!TipoDoc
        Set IT = Nothing
     End If
End Sub


Private Sub EliminarImagen()
    On Error Resume Next

    BuscaChekc = "Va a proceder a eliminar el documento de la lista. " & vbCrLf & vbCrLf & "� Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(BuscaChekc, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If Dir(lw1.SelectedItem.SubItems(2), vbArchive) <> "" Then Kill lw1.SelectedItem.SubItems(2)
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
        Else
            BuscaChekc = "delete from sfichdocs where codigo = " & Me.lw1.SelectedItem.SubItems(3)
            If ejecutar(BuscaChekc, False) Then CargarArchivos False, 0, True
            
            
        End If
    End If


End Sub



Private Sub LanzaFrmDireccionEnvio()
    Set frmDptoEnvio2 = New frmFacCliEnvDpto
    frmDptoEnvio2.DireccionesEnvio = True
    frmDptoEnvio2.VerDatoDpto = -1
    frmDptoEnvio2.codClien = CLng(Text1(0).Text)
    frmDptoEnvio2.NomClien = Text1(1).Text
    frmDptoEnvio2.Show vbModal
    Set frmDptoEnvio2 = Nothing
End Sub

'0. Insertar NORMAL
'   2.- DNI fitosanitarios
'   3.- Carnet fitosantiaruis

'   201- DNI asoci
'   202- Carnet asoc

Private Sub LanzaAnyadirImagenDocumento(TipoDoc As Integer)
    CadenaDesdeOtroForm = ""
    
    If TipoDoc > 200 Then
        frmFichaTecIMG.vDatos = Text1(0).Text & "|" & Data7.Recordset!Nombre & "|" & Data7.Recordset!Id & "|"
        
    Else
        frmFichaTecIMG.vDatos = Text1(0).Text & "|" & Text1(1).Text & "|"
    End If
    frmFichaTecIMG.Opcion_ = TipoDoc
    frmFichaTecIMG.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        'Si esta la solapa de documents
        If TipoDoc < 200 Then
            If RecuperaValor(lw1.Tag, 1) = "6" Then CargarArchivos False, Val(CadenaDesdeOtroForm), True
        Else
            
            CadenaDesdeOtroForm = "id = " & Data7.Recordset!Id
            CargaLineas True, 3
            Data7.Recordset.Find CadenaDesdeOtroForm
        End If
    End If
End Sub


Private Sub Cargacbomarjal()
    
    Set miRsAux = New ADODB.Recordset
    cbomarjal.Clear
    
    miRsAux.Open "Select distinct(partida) from sclienhuertos where partida<>'' ORDER BY 1", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cbomarjal.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cbomarjal.Tag = 1
End Sub


Private Sub PonerPais()
Dim I As Integer

    
    
    If DBLet(Data1.Recordset!codpais, "T") = "" Then
        I = -1
    Else
        For I = 0 To cboPais.ListCount - 1
            If InStr(1, cboPais.List(I), "(" & Data1.Recordset!codpais & ")") > 0 Then
                'Este es el pais
                Exit For
            End If
        Next
        If I >= cboPais.ListCount Then I = -1
    End If
    
    cboPais.ListIndex = I
End Sub



Private Function PaisSeleccionado() As String

    If cboPais.ListIndex < 0 Then
        PaisSeleccionado = ""
    Else
        PaisSeleccionado = Mid(cboPais.Text, InStr(1, cboPais.Text, "(") + 1, 2)
    End If
End Function


Private Sub ImprimirHcoPuntos()
    
    frmListado3.Opcion = 68
    frmListado3.OtrosDatos = ""
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Not IsNull(Data1.Recordset!codClien) Then frmListado3.OtrosDatos = Data1.Recordset!codClien & "|" & Data1.Recordset!NomClien & "|"
        End If
    End If
    frmListado3.Show vbModal
    
End Sub


Private Sub AbrirAlbaranesPuntos()
Dim Documento As String
Dim SQL As String

    Documento = lw1.SelectedItem.SubItems(3)
    Select Case Me.lw1.SelectedItem.SubItems(2)
        Case "ALV", "ART", "ALM", "ALZ", "ALI", "ALS", "ALO", "ALE", "ALR"
                                'ALV:Albaran de Venta (a clientes)
                                'ART: Albaran rectificativo
                                'ALM: ALbaran Mostrador
                                'ALZ: Albaranes "B"
                                'ALI: Albaranes INTERNOS
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
            'si esta ya facturado abrir el hist�rico de facturas: frmFacHcoFacturas





            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", lw1.SelectedItem.SubItems(2), "T", , "numalbar", Documento, "N")
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                         With frmFacEntAlbaranes2
                            If EsNumerico(Documento) Then
                                .hcoCodMovim = Format(Documento, "0000000")
                            Else
                                .hcoCodMovim = Documento
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .Show vbModal
                        End With
                        
                        'CargarPuntos otra vez
                        'Veremos si ha cambiado los puntos
                        
                        SQL = DevuelveDesdeBDNew(conAri, "sclien", "puntos", "codclien", CStr(Data1.Recordset!codClien))
                        If SQL = "" Then SQL = "0"
                        If CCur(SQL) <> DBLet(Data1.Recordset!Puntos, "N") Then
                            PosicionarData
                            PonerCampos
                        End If
                        
                Else
                    'FORMULARIO SAIL
                         With frmFacEntAlbSAIL
                            If EsNumerico(Documento) Then
                                .hcoCodMovim = Format(Documento, "0000000")
                            Else
                                .hcoCodMovim = Documento
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .Show vbModal
                        End With
                End If

            Else 'No existe en albaran, abrir Historico Factura
                With frmFacHcoFacturas2
                    .DesdeFichaCliente = False
                    If EsNumerico(Documento) Then
                        .hcoCodMovim = Format(Documento, "0000000")
                    Else
                        .hcoCodMovim = Documento
                    End If
                    .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                    .hcoFechaMov = lw1.SelectedItem.Text

                    .Show vbModal
                End With
            End If

        Case "ALR" 'Albaran de Reparacion (a clientes)
                If vParamAplic.TipoFormularioClientes = 0 Then
                     With frmFacEntAlbaranes2
                        If EsNumerico(Documento) Then
                            .hcoCodMovim = Format(Documento, "0000000")
                        Else
                            .hcoCodMovim = Documento
                        End If
                        .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                        .Show vbModal
                    End With
                End If

        End Select
End Sub



Private Function DesHacerIncrementoPuntosCliente() As Boolean
Dim Importe As Currency
    On Error GoTo eHacerIncrementoPuntosCliente
    DesHacerIncrementoPuntosCliente = False
    conn.BeginTrans
    
    
    Importe = ImporteFormateado(lw1.SelectedItem.SubItems(4))
    conn.Execute "UPDATE sclien set puntos=" & DBSet(-Importe, "N") & "+ coalesce(puntos,0) WHERE codclien=" & Text1(0).Text
    
    conn.Execute "DELETE from smovalpuntos where codclien=" & Text1(0).Text & " AND numero = " & lw1.SelectedItem.Tag
    conn.CommitTrans
    DesHacerIncrementoPuntosCliente = True
    Exit Function
eHacerIncrementoPuntosCliente:
    MuestraError Err.Number, Err.Description
    conn.RollbackTrans
End Function


Private Sub PonerDatosForaGridDepartamentos(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean
Dim T As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data2.Recordset Is Nothing) Then
            If Not Data2.Recordset.EOF Then Limp = False
        End If
    End If
    
    txtZona(14).Text = ""
    
    If Limp Then

        'Limpiamos
        For I = 0 To Text3.Count - 1
            Text3(I).Text = ""
    
        Next I
     
        
                
    Else
        '  0        1       2           3       4       5           6       7       8       9       10      11          12      13      14      15
        'coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba,codzona,iban
        ' 16            17                  18          19
        'organogestor,unidadtramitadora,orgproponente,oficinacontable
        For I = 0 To Text3.Count - 1
            
            If I > 0 And I < 10 Then
                T = True
            Else
                If I = 15 Then
                    T = True
                Else
                    T = False
                End If
            End If
            If T Then
                Text3(I).Text = DBLet(Data2.Recordset.Fields(I), "T")
            Else
                If IsNull(Data2.Recordset.Fields(I)) Then
                    Text3(I).Text = ""
                Else
                    If I = 13 Then
                        Text3(I).Text = DBLet(Data2.Recordset.Fields(I), "0000000000")
                    Else
                        Text3(I).Text = DBLet(Data2.Recordset.Fields(I), "0000")
                    End If
                End If
            End If
        Next
        
        If Text3(14).Text <> "" Then txtZona(14).Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text3(14).Text, "N")
    
        
    End If
End Sub


Private Sub PonerDatosForaGridDirEnvio(ForzarLimpiar As Boolean)
Dim I As Integer
Dim Limp As Boolean
Dim T As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data3.Recordset Is Nothing) Then
            If Not Data3.Recordset.EOF Then Limp = False
        End If
    End If
    
    txtZona(9).Text = ""
    
    If Limp Then

        'Limpiamos
        For I = 0 To Text4.Count - 1
            Text4(I).Text = ""
    
        Next I
     
        
                
    Else
        '  0        1       2           3       4       5           6       7       8       9
        'coddiren,nomdiren,perdiren,pobdiren,codpobla,prodiren,teldiren,faxdiren,observa,codzona
        For I = 0 To Text4.Count - 1
            
            If I > 0 And I < 10 Then
                T = True
            Else
                If I = 15 Then
                    T = True
                Else
                    T = False
                End If
            End If
            If T Then
                Text4(I).Text = DBLet(Data3.Recordset.Fields(I), "T")
            Else
                If IsNull(Data3.Recordset.Fields(I)) Then
                    Text4(I).Text = ""
                Else
                    If I = 13 Then
                        Text4(I).Text = DBLet(Data3.Recordset.Fields(I), "0000000000")
                    Else
                        Text4(I).Text = DBLet(Data3.Recordset.Fields(I), "0000")
                    End If
                End If
            End If
        Next
        
        If Text4(9).Text <> "" Then txtZona(9).Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text4(9).Text, "N")
    
        
    End If
End Sub



Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub




Private Sub BotonesToolBarAux()
Dim b As Boolean



    b = Modo = 2 Or Modo = 5
    ToolbarAux(0).Buttons(1).Enabled = b
    If b Then b = Me.Data2.Recordset.RecordCount > 0
    ToolbarAux(0).Buttons(2).Enabled = b   '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
    ToolbarAux(0).Buttons(3).Enabled = b  '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
    
    If vParamAplic.DireccionesEnvio Then
        b = Modo = 2 Or Modo = 6
        ToolbarAux(1).Buttons(1).Enabled = b
        If b Then b = Me.Data3.Recordset.RecordCount > 0
        ToolbarAux(1).Buttons(2).Enabled = b
        ToolbarAux(1).Buttons(3).Enabled = b
    End If
    
    
    'Contacto
    b = Modo = 2 Or Modo = 7
    ToolbarAux(2).Buttons(1).Enabled = b
    If b Then b = Me.Data4.Recordset.RecordCount > 0
    ToolbarAux(2).Buttons(2).Enabled = b   '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
    ToolbarAux(2).Buttons(3).Enabled = b  '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
    If Me.cboCargo.ListCount <= 0 Then CargaComboCargos
    
    'Fito
    If vParamAplic.ManipuladorFitosanitarios2 Then
        b = Modo = 2 Or Modo = 10
        ToolbarAux(3).Buttons(1).Enabled = b
        If b Then b = Me.Data7.Recordset.RecordCount > 0
        
        ToolbarAux(3).Buttons(2).Enabled = b
        ToolbarAux(3).Buttons(3).Enabled = b
    End If
    
    If vParamAplic.Renting Then
        b = Modo = 2 Or Modo = 8
        ToolbarAux(4).Buttons(1).Enabled = b
        If b Then b = Me.Data5.Recordset.RecordCount > 0
        
        ToolbarAux(4).Buttons(2).Enabled = b
        ToolbarAux(4).Buttons(3).Enabled = b
    End If
    
    If vParamAplic.TieneTelefonia2 Then
        b = Modo = 2 Or Modo = 9
        ToolbarAux(5).Buttons(1).Enabled = b
        If b Then b = Me.data6.Recordset.RecordCount > 0
        
        ToolbarAux(5).Buttons(2).Enabled = b
        ToolbarAux(5).Buttons(3).Enabled = b
    End If
    
    If vParamAplic.Huertos Then
        b = Modo = 2 Or Modo = 11
        ToolbarAux(6).Buttons(1).Enabled = b
        If b Then b = Me.data8.Recordset.RecordCount > 0
        
        ToolbarAux(6).Buttons(2).Enabled = b
        ToolbarAux(6).Buttons(3).Enabled = b
    End If
    
    
    
    
End Sub



Private Sub CargaCatalogosCliente()
Dim C As String
Dim IT As ListItem
    Me.cmdCatalogo.visible = True
    Set miRsAux = New ADODB.Recordset
    C = "SELECT sagrupa.codagrupa,descagrupa,dto1 FROM sagrupacli inner join sagrupa on sagrupacli.codagrupa=sagrupa.codagrupa"
    C = C & " WHERE codclien =" & Text1(0).Text
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = vEmpresa.FechaIni
        IT.SubItems(1) = Format(miRsAux!Dto1, FormatoImporte)
        IT.SubItems(2) = " "
        IT.SubItems(3) = Replace(miRsAux!descagrupa, "CATALOGO", "CAT:")
        IT.SubItems(4) = "AGR: " & miRsAux!codagrupa
        IT.SmallIcon = 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub



Private Function EliminardeBD() As Boolean
Dim cad As String
    On Error GoTo eEliminardeBD

    EliminardeBD = False

        cad = "DELETE FROM scliendp WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
            
        
        
        cad = "DELETE FROM slimana WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        cad = "DELETE FROM scamana WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        '
        'schped  slhped    numpedcl  fecpedcl   numpedcl  fecpedcl
        cad = "DELETE from slhped WHERE (numpedcl,fecpedcl) IN (select numpedcl,fecpedcl from schped WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from schped WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        'schpre slhpre numofert  fecofert
        cad = "DELETE from slhpre WHERE (numofert,fecofert) IN (select numofert,fecofert from schpre WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from schpre WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        'OFERTAS scapre slipre numofert  fecofert
        cad = "DELETE from slipre WHERE (numofert) IN (select numofert from scapre WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from scapre WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        
        
        'schalb  schalb_eu  slhalb   slhalb_eu   codtipom,numalbar,fechaalb   codtipom,numalbar,fechaalb1
        cad = "DELETE from slhalb WHERE (codtipom,numalbar,fechaalb) IN (select codtipom,numalbar,fechaalb from schalb WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from slhalb_eu WHERE (codtipom,numalbar) IN (select codtipom,numalbar from schalb WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from schalb_eu WHERE (codtipom,numalbar,fechaalb1) IN (select codtipom,numalbar,fechaalb from schalb WHERE codclien =" & Data1.Recordset!codClien & ")"
        conn.Execute cad
        cad = "DELETE from schalb WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        
        
        
        cad = "DELETE FROM scliendp WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
        
        '###El cliente
        cad = "DELETE FROM sclien WHERE codclien =" & Data1.Recordset!codClien
        conn.Execute cad
                
                
                
        EliminardeBD = True
        
        Exit Function

eEliminardeBD:
    MuestraError Err.Number, Err.Description
End Function
