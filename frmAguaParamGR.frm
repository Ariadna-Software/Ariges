VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAguaParamGR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Par�metros facturacion Agua Potable"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13995
   Icon            =   "frmAguaParamGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   1935
      TabIndex        =   135
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   136
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
      Height          =   705
      Left            =   135
      TabIndex        =   133
      Top             =   135
      Width           =   1740
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   134
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
               Object.Visible         =   0   'False
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
      Left            =   12150
      TabIndex        =   132
      Top             =   270
      Width           =   1530
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   135
      TabIndex        =   57
      Top             =   1005
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   13785
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmAguaParamGR.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(32)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(43)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(111)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(16)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(17)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgFec(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(39)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(23)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(22)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(30)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgBuscar(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line2(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(33)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(11)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(18)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(20)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(21)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(24)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(35)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(37)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(38)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(40)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Combo1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(31)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(32)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(33)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(29)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(26)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(25)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text2(15)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(15)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(28)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(27)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(34)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(24)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(36)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(37)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text1(38)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text1(39)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text1(40)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(41)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text1(42)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text1(43)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text1(44)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text1(45)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text1(46)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Text1(47)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text1(48)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Text1(49)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Text1(50)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Text1(51)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).ControlCount=   62
      TabCaption(1)   =   "Dom�stico"
      TabPicture(1)   =   "frmAguaParamGR.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(18)"
      Tab(1).Control(1)=   "Text2(18)"
      Tab(1).Control(2)=   "Text1(16)"
      Tab(1).Control(3)=   "Text2(16)"
      Tab(1).Control(4)=   "Text1(3)"
      Tab(1).Control(5)=   "Text2(3)"
      Tab(1).Control(6)=   "Text1(5)"
      Tab(1).Control(7)=   "Text2(5)"
      Tab(1).Control(8)=   "Text1(11)"
      Tab(1).Control(9)=   "Text2(7)"
      Tab(1).Control(10)=   "Text1(13)"
      Tab(1).Control(11)=   "Text2(9)"
      Tab(1).Control(12)=   "Text2(11)"
      Tab(1).Control(13)=   "Text2(13)"
      Tab(1).Control(14)=   "Text1(7)"
      Tab(1).Control(15)=   "Text1(9)"
      Tab(1).Control(16)=   "Text2(20)"
      Tab(1).Control(17)=   "Text1(20)"
      Tab(1).Control(18)=   "Text1(1)"
      Tab(1).Control(19)=   "Text1(2)"
      Tab(1).Control(20)=   "Text1(30)"
      Tab(1).Control(21)=   "imgBuscar(0)"
      Tab(1).Control(22)=   "Label1(14)"
      Tab(1).Control(23)=   "Label1(31)"
      Tab(1).Control(24)=   "imgBuscar(15)"
      Tab(1).Control(25)=   "Label1(28)"
      Tab(1).Control(26)=   "imgBuscar(13)"
      Tab(1).Control(27)=   "Line2(2)"
      Tab(1).Control(28)=   "imgBuscar(2)"
      Tab(1).Control(29)=   "imgBuscar(4)"
      Tab(1).Control(30)=   "imgBuscar(6)"
      Tab(1).Control(31)=   "imgBuscar(8)"
      Tab(1).Control(32)=   "imgBuscar(10)"
      Tab(1).Control(33)=   "imgBuscar(17)"
      Tab(1).Control(34)=   "Label1(56)"
      Tab(1).Control(35)=   "Label1(54)"
      Tab(1).Control(36)=   "Label1(53)"
      Tab(1).Control(37)=   "Label1(52)"
      Tab(1).Control(38)=   "Label1(51)"
      Tab(1).Control(39)=   "Label1(50)"
      Tab(1).Control(40)=   "Label1(49)"
      Tab(1).Control(41)=   "Label1(48)"
      Tab(1).Control(42)=   "Label1(47)"
      Tab(1).Control(43)=   "Label1(46)"
      Tab(1).Control(44)=   "Label1(2)"
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "Industrial"
      TabPicture(2)   =   "frmAguaParamGR.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(3)"
      Tab(2).Control(1)=   "Label1(19)"
      Tab(2).Control(2)=   "Label1(55)"
      Tab(2).Control(3)=   "Label1(1)"
      Tab(2).Control(4)=   "Label1(36)"
      Tab(2).Control(5)=   "Label1(34)"
      Tab(2).Control(6)=   "Label1(29)"
      Tab(2).Control(7)=   "Label1(26)"
      Tab(2).Control(8)=   "Label1(27)"
      Tab(2).Control(9)=   "Label1(25)"
      Tab(2).Control(10)=   "Label1(42)"
      Tab(2).Control(11)=   "imgBuscar(18)"
      Tab(2).Control(12)=   "imgBuscar(11)"
      Tab(2).Control(13)=   "imgBuscar(9)"
      Tab(2).Control(14)=   "imgBuscar(7)"
      Tab(2).Control(15)=   "imgBuscar(5)"
      Tab(2).Control(16)=   "imgBuscar(3)"
      Tab(2).Control(17)=   "imgBuscar(1)"
      Tab(2).Control(18)=   "Line2(3)"
      Tab(2).Control(19)=   "Label1(4)"
      Tab(2).Control(20)=   "imgBuscar(14)"
      Tab(2).Control(21)=   "imgBuscar(16)"
      Tab(2).Control(22)=   "Label1(5)"
      Tab(2).Control(23)=   "Label1(15)"
      Tab(2).Control(24)=   "Text1(23)"
      Tab(2).Control(25)=   "Text1(22)"
      Tab(2).Control(26)=   "Text1(35)"
      Tab(2).Control(27)=   "Text1(21)"
      Tab(2).Control(28)=   "Text2(21)"
      Tab(2).Control(29)=   "Text1(10)"
      Tab(2).Control(30)=   "Text1(8)"
      Tab(2).Control(31)=   "Text2(14)"
      Tab(2).Control(32)=   "Text2(12)"
      Tab(2).Control(33)=   "Text2(10)"
      Tab(2).Control(34)=   "Text1(14)"
      Tab(2).Control(35)=   "Text2(8)"
      Tab(2).Control(36)=   "Text1(12)"
      Tab(2).Control(37)=   "Text2(6)"
      Tab(2).Control(38)=   "Text1(6)"
      Tab(2).Control(39)=   "Text2(4)"
      Tab(2).Control(40)=   "Text1(4)"
      Tab(2).Control(41)=   "Text2(17)"
      Tab(2).Control(42)=   "Text1(17)"
      Tab(2).Control(43)=   "Text2(19)"
      Tab(2).Control(44)=   "Text1(19)"
      Tab(2).ControlCount=   45
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
         Index           =   51
         Left            =   7995
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "T|T|S|||sparamAgua|CodEntSuministra|||"
         Text            =   "Text1"
         Top             =   840
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
         Index           =   50
         Left            =   5730
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "T|T|S|||sparamAgua|TfnoAverias|||"
         Text            =   "Text1"
         Top             =   840
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
         Index           =   49
         Left            =   11625
         MaxLength       =   15
         TabIndex        =   27
         Tag             =   "Alc m3 Bloque2|N|N|0||sparamAgua|BloqueAlc2|||"
         Text            =   "Text1"
         Top             =   7140
         Width           =   645
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
         Index           =   48
         Left            =   9225
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Alcan m3 bloque 1|N|N|0||sparamAgua|BloqueAlc1|||"
         Text            =   "Text1"
         Top             =   7140
         Width           =   645
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
         Index           =   47
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "Precio dom B1|N|S|0||sparamAgua|preAlcDomB1|0.000||"
         Text            =   "Text1"
         Top             =   4680
         Width           =   1185
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
         Index           =   46
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "Precio dom B2|N|S|0||sparamAgua|preAlcDomB2|0.000||"
         Text            =   "Text1"
         Top             =   5160
         Width           =   1185
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
         Index           =   45
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Precio dom B3|N|S|0||sparamAgua|preAlcDomB3|0.000||"
         Text            =   "Text1"
         Top             =   5640
         Width           =   1185
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
         Index           =   44
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Precio ind B1|N|S|0||sparamAgua|preAlcIndB1|0.000||"
         Text            =   "Text1"
         Top             =   4680
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "Precio ind B2|N|S|0||sparamAgua|preAlcIndB2|0.000||"
         Text            =   "Text1"
         Top             =   5160
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Precio ind B3|N|S|0||sparamAgua|preAlcIndB3|0.000||"
         Text            =   "Text1"
         Top             =   5640
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "Precio ind B3|N|S|0||sparamAgua|PreConIndB3|0.000||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "Precio ind B2|N|S|0||sparamAgua|PreConIndB2|0.000||"
         Text            =   "Text1"
         Top             =   3360
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Precio ind B1|N|S|0||sparamAgua|PreConIndB1|0.000||"
         Text            =   "Text1"
         Top             =   2880
         Width           =   1185
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Precio dom B3|N|S|0||sparamAgua|PreConDomB3|0.000||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   1185
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Precio dom B2|N|S|0||sparamAgua|PreConDomB2|0.000||"
         Text            =   "Text1"
         Top             =   3360
         Width           =   1185
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Precio dom B1|N|S|0||sparamAgua|PreConDomB1|0.000||"
         Text            =   "Text1"
         Top             =   2880
         Width           =   1185
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
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Precio afor indus|N|S|0||sparamAgua|PrecioAfoMaxI|0.000||"
         Text            =   "Text1"
         Top             =   6600
         Visible         =   0   'False
         Width           =   1185
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Precio aforo dom|N|S|0||sparamAgua|PrecioAfoMaxD|0.000||"
         Text            =   "Text1"
         Top             =   6600
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   4965
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Precio consu. gen. Indus|N|S|0||sparamAgua|PrecioServGenI|0.000||"
         Text            =   "Text1"
         Top             =   6600
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
         Index           =   28
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Precio consumo gen. dom|N|S|0||sparamAgua|PrecioServGenD|0.000||"
         Text            =   "Text1"
         Top             =   7200
         Width           =   1185
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
         Left            =   -68025
         MaxLength       =   16
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   6840
         Width           =   2070
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
         Left            =   -65910
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   110
         Text            =   "Text2"
         Top             =   6840
         Width           =   4245
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
         Left            =   -67920
         MaxLength       =   16
         TabIndex        =   51
         Text            =   "19"
         Top             =   6840
         Width           =   2070
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
         Left            =   -65805
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   108
         Text            =   "Text2"
         Top             =   6840
         Width           =   4035
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   6840
         Width           =   2070
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   107
         Text            =   "Text2"
         Top             =   6840
         Width           =   4305
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
         Left            =   -74760
         MaxLength       =   16
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   6840
         Width           =   2070
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
         Left            =   -72645
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   104
         Text            =   "Text2"
         Top             =   6840
         Width           =   4575
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1920
         Width           =   2070
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   101
         Text            =   "Text2"
         Top             =   1920
         Width           =   4305
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   3120
         Width           =   2070
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
         Index           =   6
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   100
         Text            =   "Text2"
         Top             =   3120
         Width           =   4305
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   4320
         Width           =   2070
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
         Index           =   8
         Left            =   -65790
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   99
         Text            =   "Text2"
         Top             =   1920
         Width           =   4050
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
         Index           =   14
         Left            =   -67905
         MaxLength       =   16
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   4320
         Width           =   2070
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
         Left            =   -65790
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   98
         Text            =   "Text2"
         Top             =   3120
         Width           =   4050
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   4320
         Width           =   4305
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
         Left            =   -65790
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   96
         Text            =   "Text2"
         Top             =   4320
         Width           =   4050
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
         Left            =   -67905
         MaxLength       =   16
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1920
         Width           =   2070
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
         Left            =   -67905
         MaxLength       =   16
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   3120
         Width           =   2070
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   95
         Text            =   "Text2"
         Top             =   5160
         Width           =   4305
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   5160
         Width           =   2070
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2055
         Width           =   2070
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   87
         Text            =   "Text2"
         Top             =   2055
         Width           =   4500
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   3120
         Width           =   2070
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
         Index           =   5
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   86
         Text            =   "Text2"
         Top             =   3120
         Width           =   4500
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4320
         Width           =   2070
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
         Index           =   7
         Left            =   -65925
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   2055
         Width           =   4275
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
         Left            =   -67995
         MaxLength       =   16
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   4320
         Width           =   2070
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
         Left            =   -65925
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   84
         Text            =   "Text2"
         Top             =   3120
         Width           =   4275
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
         Index           =   11
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   83
         Text            =   "Text2"
         Top             =   4320
         Width           =   4470
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
         Left            =   -65925
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   4320
         Width           =   4275
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
         Left            =   -67995
         MaxLength       =   16
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   2055
         Width           =   2070
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
         Left            =   -67995
         MaxLength       =   16
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   3120
         Width           =   2070
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
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   81
         Text            =   "Text2"
         Top             =   5400
         Width           =   4470
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
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   5400
         Width           =   2070
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
         Left            =   240
         MaxLength       =   16
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1560
         Width           =   2070
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
         Left            =   2385
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   1560
         Width           =   5085
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
         Left            =   -73080
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Indus Bloque2|N|N|||sparamAgua|bloque2I|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Left            =   -74640
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Indus bloque 1|N|N|0||sparamAgua|bloque1I|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   42
         Tag             =   "AForo indus|N|N|0||sparamAgua|AforoMaxInd|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Left            =   -74640
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "m3 bloque 1|N|N|0||sparamAgua|bloque1D|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Left            =   -73080
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "m3 Bloque2|N|N|0||sparamAgua|bloque2D|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Index           =   30
         Left            =   -71640
         MaxLength       =   15
         TabIndex        =   31
         Tag             =   "Aforo dom|N|N|0||sparamAgua|AforoMaxDom|||"
         Text            =   "Text1"
         Top             =   960
         Width           =   645
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
         Left            =   7905
         MaxLength       =   40
         TabIndex        =   24
         Tag             =   "T|T|S|||sparamAgua|TextoDOGV1|||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   3525
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
         Index           =   26
         Left            =   7905
         MaxLength       =   40
         TabIndex        =   25
         Tag             =   "T|T|S|||sparamAgua|TextoDOGV2|||"
         Text            =   "Text1"
         Top             =   4680
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
         Index           =   29
         Left            =   7905
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Fecha tarifa|F|S|||sparamAgua|fechaServGen|||"
         Text            =   "Text1"
         Top             =   2880
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
         Index           =   33
         Left            =   2700
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Alerta|N|N|0||sparamAgua|AlertaConsumo|||"
         Text            =   "Text1"
         Top             =   840
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
         Index           =   32
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "Ult a�o|N|N|2013|2030|sparamAgua|UltimoAnyoLiquidado|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   780
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "Ult periodo|N|N|1|12|sparamAgua|UtimoPeridoLiquidado|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   405
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
         ItemData        =   "frmAguaParamGR.frx":0060
         Left            =   660
         List            =   "frmAguaParamGR.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Periodo|N|N|||sparamAgua|periodoFac|||"
         Top             =   840
         Width           =   1815
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
         Left            =   240
         MaxLength       =   4
         TabIndex        =   54
         Tag             =   "C�dig|N|N|1|1|sparamAgua|codigo||S|"
         Text            =   "Text"
         Top             =   840
         Width           =   285
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -72435
         Picture         =   "frmAguaParamGR.frx":00BB
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. entidad suminist."
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
         Index           =   40
         Left            =   7995
         TabIndex        =   131
         Top             =   600
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Tel�fono de averias"
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
         Index           =   38
         Left            =   5730
         TabIndex        =   130
         Top             =   600
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque2 hasta"
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
         Index           =   37
         Left            =   10425
         TabIndex        =   129
         Top             =   7200
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque1 hasta"
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
         Index           =   35
         Left            =   8025
         TabIndex        =   128
         Top             =   7200
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque I"
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
         Left            =   2100
         TabIndex        =   126
         Top             =   4740
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pre. alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   21
         Left            =   240
         TabIndex        =   125
         Top             =   4740
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque II"
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
         Index           =   20
         Left            =   2100
         TabIndex        =   124
         Top             =   5220
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque III"
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
         Index           =   18
         Left            =   2100
         TabIndex        =   123
         Top             =   5700
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Industrial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Index           =   15
         Left            =   -64680
         TabIndex        =   122
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dom�stico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   300
         Index           =   14
         Left            =   -64590
         TabIndex        =   121
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio consumo Generalitat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   240
         TabIndex        =   120
         Top             =   7260
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio aforo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   240
         TabIndex        =   119
         Top             =   6720
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque III"
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
         Index           =   11
         Left            =   2100
         TabIndex        =   118
         Top             =   3900
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque II"
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
         Index           =   10
         Left            =   2100
         TabIndex        =   117
         Top             =   3420
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio consumo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   116
         Top             =   2940
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Industrial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   4995
         TabIndex        =   115
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dom�stico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   3150
         TabIndex        =   114
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque I"
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
         Index           =   33
         Left            =   2100
         TabIndex        =   113
         Top             =   2940
         Width           =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Index           =   0
         X1              =   240
         X2              =   6165
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   31
         Left            =   -68025
         TabIndex        =   111
         Top             =   6555
         Width           =   2115
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   -65820
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   6600
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   -67890
         TabIndex        =   109
         Top             =   6570
         Width           =   2115
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   -65685
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   6600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   -72615
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   6555
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   -74640
         TabIndex        =   106
         Top             =   6555
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   28
         Left            =   -74760
         TabIndex        =   105
         Top             =   6555
         Width           =   1980
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   -72735
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   6600
         Width           =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         Index           =   3
         X1              =   -74880
         X2              =   -61860
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         Index           =   2
         X1              =   -74760
         X2              =   -61725
         Y1              =   6375
         Y2              =   6375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -72450
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -72615
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   2835
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   -64890
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -65025
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   2835
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -71460
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   -63720
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   -73305
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   4875
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota aforos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   42
         Left            =   -74640
         TabIndex        =   94
         Top             =   4875
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   25
         Left            =   -74640
         TabIndex        =   93
         Top             =   1635
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio agua "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   27
         Left            =   -74640
         TabIndex        =   92
         Top             =   2835
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio transitoria-amortizaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   26
         Left            =   -67905
         TabIndex        =   91
         Top             =   4035
         Width           =   4140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota mantenimiento contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   29
         Left            =   -74640
         TabIndex        =   90
         Top             =   4035
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   34
         Left            =   -67905
         TabIndex        =   89
         Top             =   1635
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   36
         Left            =   -67905
         TabIndex        =   88
         Top             =   2835
         Width           =   2820
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -72615
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   -64980
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -65115
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -71415
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -63795
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   -73290
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuota"
         Top             =   5130
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   990
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   30
         Left            =   240
         TabIndex        =   80
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo agua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   56
         Left            =   -74640
         TabIndex        =   78
         Top             =   1770
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio agua "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   54
         Left            =   -74640
         TabIndex        =   77
         Top             =   2835
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio transitoria-amortizaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   53
         Left            =   -67995
         TabIndex        =   76
         Top             =   4035
         Width           =   4140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota mantenimiento contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   52
         Left            =   -74640
         TabIndex        =   75
         Top             =   4035
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota consumo alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   51
         Left            =   -67995
         TabIndex        =   74
         Top             =   1770
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota servicio alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   50
         Left            =   -67995
         TabIndex        =   73
         Top             =   2835
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota aforos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   49
         Left            =   -74640
         TabIndex        =   72
         Top             =   5130
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque1 hasta"
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
         Left            =   -74640
         TabIndex        =   71
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque2 hasta"
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
         Left            =   -73080
         TabIndex        =   70
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Aforo max"
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
         Index           =   19
         Left            =   -71640
         TabIndex        =   69
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque1 hasta"
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
         Index           =   48
         Left            =   -74640
         TabIndex        =   68
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque2 hasta"
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
         Index           =   47
         Left            =   -73080
         TabIndex        =   67
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Aforo max"
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
         Index           =   46
         Left            =   -71640
         TabIndex        =   66
         Top             =   720
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   3
         Index           =   1
         X1              =   7830
         X2              =   12630
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label1 
         Caption         =   "Texto DOGV-1"
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
         Index           =   22
         Left            =   7905
         TabIndex        =   64
         Top             =   3600
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Texto DOGV-2"
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
         Index           =   23
         Left            =   7905
         TabIndex        =   63
         Top             =   4440
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha tarifa"
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
         Index           =   39
         Left            =   7905
         TabIndex        =   62
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   9180
         Picture         =   "frmAguaParamGR.frx":0ABD
         ToolTipText     =   "Buscar fecha"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Alerta"
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
         Index           =   17
         Left            =   2700
         TabIndex        =   61
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Ultimo periodo"
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
         Left            =   3840
         TabIndex        =   60
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo facturaci�n"
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
         Index           =   111
         Left            =   660
         TabIndex        =   59
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
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
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Precios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   112
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Generalitat valenciana"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Index           =   3
         Left            =   -74880
         TabIndex        =   103
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Generalitat valenciana"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   102
         Top             =   6000
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Generalitat valenciana"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Index           =   43
         Left            =   7785
         TabIndex        =   65
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Index           =   4
         X1              =   7905
         X2              =   12705
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label1 
         Caption         =   "Bloque alcantarillado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   32
         Left            =   7905
         TabIndex        =   127
         Top             =   6600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   55
      Top             =   9045
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   165
         Width           =   2115
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
      Left            =   12660
      TabIndex        =   53
      Top             =   9120
      Width           =   1065
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
      Left            =   11340
      TabIndex        =   52
      Top             =   9120
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   720
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
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
Attribute VB_Name = "frmAguaParamGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents FrmArt As frmBasico2
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin ningun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Private CadenaConsulta As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1



Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    
                    PosicionarData
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        Case 1  'BUSCAR
            HacerBusqueda
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3 'Insertar
        LimpiarCampos
        PonerModo 0
        PonerOpcionesMenu
    Case 4  'Modificar
        lblIndicador.Caption = ""
        TerminaBloquear
        PonerModo 2
        PonerCampos
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'    'Sugerir el siguiente codigo
'    Text1(0).Text = Format(SugerirCodigoSiguienteStr("sparamAgua", "codagent"), "0000")
'    PonerFoco Text1(0)
'    Text1_GotFocus 0
'End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
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
    'Ver todos
  '  If chkVistaPrevia.Value = 1 Then
  '      MandaBusquedaPrevia ""
  '  Else
        LimpiarCampos
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
  '  End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(33)
End Sub





'Private Sub cmdRegresar_Click()
'Dim cad As String
'
'    If Data1.Recordset.EOF Then
'        MsgBox "Ning�n registro devuelto.", vbExclamation
'        Exit Sub
'    End If
'
'    cad = Data1.Recordset.Fields(0) & "|"
'    cad = cad & Data1.Recordset.Fields(1) & "|"
'    RaiseEvent DatoSeleccionado(cad)
'    Unload Me
'End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If NombreTabla = "" Then
        'ASignamos un SQL al DATA1
        '## A mano
        NombreTabla = "sparamAgua"
        BotonVerTodos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    For I = 1 To imgBuscar.Count - 1
        imgBuscar(I).Picture = imgBuscar(0).Picture
    Next

    ' ICONITOS DE LA BARRA
'    btnPrimero = 13
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 15  'Salir
'        .Buttons(13).Image = 6  'Primero
'        .Buttons(14).Image = 7  'Anterior
'        .Buttons(15).Image = 8  'Siguiente
'        .Buttons(16).Image = 9  '�ltimo
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
    
    LimpiarCampos
        
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "artCuotaConD|artCuotaConI|artCuotaServD|artCuotaServI|artAlcanConD|artAlcanConI|artAlcanServD|artAlcanServI|artContadorD|"
    NombreTabla = NombreTabla & "artContadorI|artAmortD|artAmortI|artVarios|artCuotaConGenD|artCuotaConGenI|artCuotaServGenD|artCuotaServGenI|"
    NombreTabla = NombreTabla & "artAforoD|artAforoI|"
    
    For kCampo = 3 To 21
        Me.Text1(kCampo).Tag = "Articulo: " & kCampo - 2 & "|T|S|||sparamAgua|" & RecuperaValor(NombreTabla, kCampo - 2) & "|||"
        'Text2(kCampo).Text = RecuperaValor(NombreTabla, kCampo - 2)
        'Text1(kCampo).Text = RecuperaValor(NombreTabla, kCampo - 2)
    Next
    NombreTabla = ""
    
    '|T|S|||sparamAgua|artCanon3|||
    
    
    Data1.ConnectionString = conn
'    Data1.RecordSource = "Select * from " & NombreTabla & " where codigo=1"
'    Data1.Refresh
   
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    CadenaConsulta = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    CadenaConsulta = ""
    Set FrmArt = New frmBasico2
'    FrmArt.DesdeTPV = False
'    FrmArt.Show vbModal
    AyudaArticulos FrmArt, Text1(Index + 3)
    Set FrmArt = Nothing
    If CadenaConsulta <> "" Then
        Text1(Index + 3).Text = RecuperaValor(CadenaConsulta, 1)
        Text2(Index + 3).Text = RecuperaValor(CadenaConsulta, 2)
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
 If Modo = 2 Or Modo = 0 Then Exit Sub
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Me.Text1(29).Text <> "" Then frmC.Fecha = CDate(Text1(29).Text)
    frmC.Show vbModal
    If CadenaConsulta <> "" Then Text1(29).Text = CadenaConsulta
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

'Private Sub mnNuevo_Click()
'    BotonAnyadir
'End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
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
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 0 'consumo agua
            Case 7: KEYBusqueda KeyAscii, 4 'consumo alcantarillado
            Case 5: KEYBusqueda KeyAscii, 2 'cuota servicio agua
            Case 9: KEYBusqueda KeyAscii, 6 'cuota serv alcanatarillado
            Case 11: KEYBusqueda KeyAscii, 8 'cuota mto contenedor
            Case 13: KEYBusqueda KeyAscii, 10 '
            Case 16: KEYBusqueda KeyAscii, 13 'cuota servicio agua
            Case 18: KEYBusqueda KeyAscii, 15 'cuota consumo agua
            Case 15: KEYBusqueda KeyAscii, 12 'varios
            
            Case 4: KEYBusqueda KeyAscii, 1 'consumo agua
            Case 8: KEYBusqueda KeyAscii, 5 'cuota consumo alcantarillado
            Case 6: KEYBusqueda KeyAscii, 3 'cuota serv agua
            Case 10: KEYBusqueda KeyAscii, 7 'cuota serv alcantarillado
            Case 12: KEYBusqueda KeyAscii, 9 'cuota mto contador
            Case 14: KEYBusqueda KeyAscii, 11 'cuota serv trans-amort
            Case 21: KEYBusqueda KeyAscii, 18 'cuota aforos
            Case 17: KEYBusqueda KeyAscii, 14 'cuota serv agua
            Case 19: KEYBusqueda KeyAscii, 16 'cuota cons agua
            
            Case 29: KEYFecha KeyAscii, 0  'fecha
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
    
    Select Case Index
    Case 1, 2
       If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
        
    Case 3 To 21
        'LAS CUENTAS
            devuelve = ""
            If Text1(Index).Text <> "" Then
                devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(Index).Text, "T")
                If devuelve = "" Then
                    MsgBox "No existe el articulo", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
            Text2(Index).Text = devuelve
        
    Case 22, 23, 30, 31, 32, 33, 35, 48, 49
       
        If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
        
    Case 24, 27, 28, 34, 36 To 47
        If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 5
        
    Case 29
        If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)

    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB
    PonerCadenaBusqueda
   
End Sub




Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        End If
         Screen.MousePointer = vbDefault
         'PonerModo 0
         Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
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


    For kCampo = 3 To 21
        If Text1(kCampo).Text <> "" Then
            CadenaConsulta = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(kCampo).Text, "T")
            If CadenaConsulta = "" Then CadenaConsulta = "******* ERROR leyendo articulo"
            Text2(kCampo).Text = CadenaConsulta
        End If
    Next
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo

    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
'    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = b
'    Else
'        cmdRegresar.visible = False
'    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    b = Modo = 1 Or Modo >= 3 'busqueda o inser/mod
    BloquearCmb Combo1, Not b
    
    '---------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
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



Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).visible = False
    Me.mnNuevo.visible = False
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).visible = False
    mnEliminar.visible = False
    
    '----------------------------------------
    b = (Modo >= 3) 'Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim cad As String

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar datos OK
    If Not b Then Exit Function
        
    CadenaConsulta = ""
    
    'Bloque 1 no puede ser mayor o igual que bloque 2
    If Val(Text1(1).Text) >= Val(Text1(2).Text) Then CadenaConsulta = CadenaConsulta & "- Bloque uno no puede ser mayor ni igual que el dos" & vbCrLf
        
    For kCampo = 3 To 15
        Select Case kCampo
        Case 3, 4, 5, 11, 12, 13
            If Text1(kCampo).Text = "" Then CadenaConsulta = CadenaConsulta & "- " & RecuperaValor(Text1(kCampo).Tag, 1) & "- no puede estar vacio" & vbCrLf
        
        End Select
    Next
        
        
    'Periodo liquidacion
    '------------------------
    'El periodo ANUAL ya esta bien, el valor que tenga
    If Combo1.ListIndex > 0 Then
        
        kCampo = Val(Me.Text1(31).Text)
        CadenaDesdeOtroForm = ""
        Select Case Combo1.ItemData(Combo1.ListIndex)
        Case 1
          '  If kCampo > 2 Then CadenaDesdeOtroForm = ""
        Case 2
            If kCampo > 6 Then CadenaDesdeOtroForm = "6"
        Case 3
            If kCampo > 4 Then CadenaDesdeOtroForm = "4"
        Case 4
            If kCampo > 3 Then CadenaDesdeOtroForm = "3"
        Case 6
            If kCampo > 2 Then CadenaDesdeOtroForm = "2"
        
        End Select
        If CadenaDesdeOtroForm <> "" Then
            CadenaConsulta = CadenaConsulta & vbCrLf & "- Periodo no puede ser mayor de " & CadenaDesdeOtroForm
            CadenaDesdeOtroForm = ""
        End If
    End If
        
        
        
        
    If CadenaConsulta <> "" Then
        b = False
        MsgBox CadenaConsulta, vbExclamation
    End If
    
    DatosOk = b
End Function

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
          '  mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            'mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
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


Private Sub PosicionarData()
Dim cad As String
Dim Indicador As String

    cad = "(codigo=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


