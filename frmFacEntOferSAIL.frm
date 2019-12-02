VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntOferSAIL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ofertas Clientes"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14400
   Icon            =   "frmFacEntOferSAIL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   920
      Left            =   120
      TabIndex        =   104
      Top             =   390
      Width           =   14175
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   8925
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   109
         Text            =   "Text2"
         Top             =   160
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   8160
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scapre|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   160
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   8925
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scapre|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   520
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scapre|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   520
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1220
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Oferta|F|N|||scapre|fecofert|dd/mm/yyyy|N|"
         Top             =   430
         Width           =   1185
      End
      Begin VB.CheckBox chkAceptado 
         Caption         =   "Aceptada"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Tag             =   "Aceptada|N|N|||scapre|aceptado||N|"
         Top             =   405
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   200
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Oferta|N|S|0||scapre|numofert|0000000|S|"
         Text            =   "Text1 7"
         Top             =   430
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrega|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
         Top             =   430
         Width           =   1305
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   7860
         Picture         =   "frmFacEntOferSAIL.frx":000C
         ToolTipText     =   "Buscar trabajador"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Realiz. por"
         Height          =   255
         Index           =   21
         Left            =   7065
         TabIndex        =   110
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   7860
         Picture         =   "frmFacEntOferSAIL.frx":010E
         ToolTipText     =   "Buscar cliente"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   7065
         TabIndex        =   108
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "F. Oferta"
         Height          =   255
         Index           =   14
         Left            =   1235
         TabIndex        =   107
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2060
         Picture         =   "frmFacEntOferSAIL.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   50
         Left            =   200
         TabIndex        =   106
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3600
         Picture         =   "frmFacEntOferSAIL.frx":029B
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. reclamación"
         Height          =   195
         Index           =   51
         Left            =   2520
         TabIndex        =   105
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   80
      Text            =   "frmFacEntOferSAIL.frx":0326
      Top             =   7800
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   63
      Top             =   7680
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   64
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12960
      TabIndex        =   61
      Top             =   7680
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11760
      TabIndex        =   60
      Top             =   7680
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   360
      Top             =   7320
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
      Left            =   1800
      Top             =   7440
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
      Height          =   6240
      Left            =   120
      TabIndex        =   67
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11007
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmFacEntOferSAIL.frx":0363
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
      Tab(0).Control(14)=   "txtAux(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAux(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboOpcion"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Textos de la Carta"
      TabPicture(1)   =   "frmFacEntOferSAIL.frx":037F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(34)"
      Tab(1).Control(1)=   "Text1(19)"
      Tab(1).Control(2)=   "Text1(20)"
      Tab(1).Control(3)=   "Text1(18)"
      Tab(1).Control(4)=   "Text1(30)"
      Tab(1).Control(5)=   "Text1(29)"
      Tab(1).Control(6)=   "Text1(28)"
      Tab(1).Control(7)=   "Text1(27)"
      Tab(1).Control(8)=   "Text1(26)"
      Tab(1).Control(9)=   "Text1(25)"
      Tab(1).Control(10)=   "Text1(24)"
      Tab(1).Control(11)=   "Text1(23)"
      Tab(1).Control(12)=   "Text1(22)"
      Tab(1).Control(13)=   "Text1(21)"
      Tab(1).Control(14)=   "Label1(46)"
      Tab(1).Control(15)=   "Label1(2)"
      Tab(1).Control(16)=   "Label1(45)"
      Tab(1).Control(17)=   "Label1(5)"
      Tab(1).Control(18)=   "Label1(3)"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Concepto y Gestión Oferta"
      TabPicture(2)   =   "frmFacEntOferSAIL.frx":039B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(32)"
      Tab(2).Control(1)=   "Text1(31)"
      Tab(2).Control(2)=   "Label1(37)"
      Tab(2).Control(3)=   "Label1(38)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Grupo sumergido"
      TabPicture(3)   =   "frmFacEntOferSAIL.frx":03B7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text4(23)"
      Tab(3).Control(1)=   "Text4(22)"
      Tab(3).Control(2)=   "Text4(21)"
      Tab(3).Control(3)=   "Text4(20)"
      Tab(3).Control(4)=   "Text4(19)"
      Tab(3).Control(5)=   "Text4(18)"
      Tab(3).Control(6)=   "Text4(17)"
      Tab(3).Control(7)=   "Text4(16)"
      Tab(3).Control(8)=   "Text4(15)"
      Tab(3).Control(9)=   "Text4(14)"
      Tab(3).Control(10)=   "Text4(13)"
      Tab(3).Control(11)=   "Text4(12)"
      Tab(3).Control(12)=   "Text4(11)"
      Tab(3).Control(13)=   "Text4(10)"
      Tab(3).Control(14)=   "Text4(9)"
      Tab(3).Control(15)=   "Text4(8)"
      Tab(3).Control(16)=   "Text4(7)"
      Tab(3).Control(17)=   "Text4(6)"
      Tab(3).Control(18)=   "Text4(5)"
      Tab(3).Control(19)=   "Text4(4)"
      Tab(3).Control(20)=   "Text4(3)"
      Tab(3).Control(21)=   "Text4(2)"
      Tab(3).Control(22)=   "Text4(1)"
      Tab(3).Control(23)=   "Text4(0)"
      Tab(3).Control(24)=   "imgBuscar(10)"
      Tab(3).Control(25)=   "imgBuscar(9)"
      Tab(3).Control(26)=   "imgBuscar(8)"
      Tab(3).Control(27)=   "Label1(44)"
      Tab(3).Control(28)=   "Label1(43)"
      Tab(3).Control(29)=   "Label1(40)"
      Tab(3).Control(30)=   "Label1(29)"
      Tab(3).Control(31)=   "Label1(28)"
      Tab(3).Control(32)=   "Label1(27)"
      Tab(3).Control(33)=   "Label1(24)"
      Tab(3).Control(34)=   "Label1(23)"
      Tab(3).Control(35)=   "Label1(18)"
      Tab(3).Control(36)=   "Line2"
      Tab(3).ControlCount=   37
      TabCaption(4)   =   "Documentos"
      TabPicture(4)   =   "frmFacEntOferSAIL.frx":03D3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdPDF"
      Tab(4).Control(1)=   "AcroPDF1"
      Tab(4).Control(2)=   "ListView1"
      Tab(4).Control(3)=   "Label1(47)"
      Tab(4).Control(4)=   "ImgPDF(1)"
      Tab(4).Control(5)=   "ImgPDF(2)"
      Tab(4).Control(6)=   "ImgPDF(0)"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Totales"
      TabPicture(5)   =   "frmFacEntOferSAIL.frx":03EF
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameFactura"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton cmdPDF 
         Height          =   315
         Left            =   -67560
         Picture         =   "frmFacEntOferSAIL.frx":040B
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cboOpcion 
         Height          =   315
         ItemData        =   "frmFacEntOferSAIL.frx":167D
         Left            =   13440
         List            =   "frmFacEntOferSAIL.frx":1687
         Style           =   2  'Dropdown List
         TabIndex        =   172
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   4935
         Left            =   -67560
         TabIndex        =   171
         Top             =   960
         Width           =   6375
         _cx             =   5080
         _cy             =   5080
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   -74520
         TabIndex        =   169
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8493
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fichero"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame FrameFactura 
         Height          =   3900
         Left            =   -73560
         TabIndex        =   130
         Top             =   960
         Width           =   10575
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
            Index           =   1
            Left            =   240
            MaxLength       =   15
            TabIndex        =   173
            Text            =   "Text1 7"
            Top             =   3240
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   153
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   152
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   151
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   150
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   149
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   148
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   147
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   146
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   145
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   144
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   143
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   142
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   141
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   140
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   139
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   138
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   137
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   136
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   135
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
            TabIndex        =   134
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
            TabIndex        =   133
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
            TabIndex        =   132
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
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
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   131
            Text            =   "Text1 7"
            Top             =   3240
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Total opciones"
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
            Height          =   255
            Index           =   49
            Left            =   240
            TabIndex        =   174
            Top             =   3000
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   168
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   5520
            TabIndex        =   167
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   2400
            X2              =   9120
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4920
            TabIndex        =   166
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   2520
            TabIndex        =   165
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   7200
            TabIndex        =   164
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   22
            Left            =   8040
            TabIndex        =   163
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   162
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2160
            TabIndex        =   161
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   160
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   8
            Left            =   5760
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
            Top             =   480
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
            Index           =   36
            Left            =   11880
            TabIndex        =   155
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL OFERTA"
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
            Height          =   255
            Index           =   39
            Left            =   5640
            TabIndex        =   154
            Top             =   3240
            Width           =   1530
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1005
         Index           =   34
         Left            =   -73200
         MaxLength       =   255
         TabIndex        =   128
         Tag             =   "Obs CRM|T|S|||scapre|observacrm|||"
         Top             =   4680
         Width           =   7755
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   -63120
         TabIndex        =   59
         Text            =   "Text4"
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   22
         Left            =   -64920
         TabIndex        =   58
         Text            =   "Text4"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   -65760
         TabIndex        =   57
         Text            =   "Text4"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   -66600
         TabIndex        =   56
         Text            =   "Text4"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   -64920
         TabIndex        =   55
         Text            =   "Text4"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   -65760
         TabIndex        =   54
         Text            =   "Text4"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   -66600
         TabIndex        =   53
         Text            =   "Text4"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   16
         Left            =   -74280
         TabIndex        =   52
         Text            =   "Text4"
         Top             =   4800
         Width           =   6495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   -63120
         TabIndex        =   51
         Text            =   "Text4"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   -64920
         TabIndex        =   50
         Text            =   "Text4"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   -65760
         TabIndex        =   49
         Text            =   "Text4"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   -66600
         TabIndex        =   48
         Text            =   "Text4"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   -64920
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   -65760
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   -66600
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   8
         Left            =   -74280
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   3120
         Width           =   6495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -63120
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -64920
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -65760
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -66600
         TabIndex        =   40
         Text            =   "Text4"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -64920
         TabIndex        =   39
         Text            =   "Text4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -65760
         TabIndex        =   38
         Text            =   "Text4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -66600
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   0
         Left            =   -74280
         TabIndex        =   36
         Text            =   "Text4"
         Top             =   1560
         Width           =   6495
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   11880
         TabIndex        =   118
         ToolTipText     =   "Buscar artículo"
         Top             =   4080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   12000
         TabIndex        =   117
         Tag             =   "DescCa"
         Text            =   "DescCap"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   11040
         MaxLength       =   12
         TabIndex        =   79
         Tag             =   "Captilo"
         Text            =   "Capitulo"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   81
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   2000
         Left            =   240
         TabIndex        =   88
         Top             =   370
         Width           =   13815
         Begin VB.ComboBox cboMotTra 
            Height          =   315
            Left            =   6885
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Tag             =   "Motivo|N|N|||scapre|motivoTraspaso||N|"
            Top             =   1620
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   33
            Left            =   8520
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   600
            Width           =   2310
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   33
            Left            =   6885
            MaxLength       =   20
            TabIndex        =   15
            Tag             =   "Actua|T|S|||scapre|actuacion||N|"
            Text            =   "Tex"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   101
            Tag             =   "Direccion/Dpto.|T|S|||scapre|nomdirec||N|"
            Text            =   "Text2"
            Top             =   210
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|9999|scapre|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   210
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scapre|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1266
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scapre|codpobla||N|"
            Text            =   "Text15"
            Top             =   914
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1875
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scapre|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   914
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3315
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scapre|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   210
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scapre|nifclien||N|"
            Text            =   "123456789"
            Top             =   210
            Width           =   990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1200
            MaxLength       =   255
            TabIndex        =   13
            Tag             =   "Referencia Cliente|T|S|||scapre|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1610
            Width           =   4125
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "Cod. Agente|N|N|0|9999|scapre|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   915
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   95
            Text            =   "Text2"
            Top             =   915
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   17
            Tag             =   "Forma de Pago|N|N|0|999|scapre|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1275
            Width           =   510
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   1275
            Width           =   3390
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   12240
            MaxLength       =   5
            TabIndex        =   18
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scapre|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   240
            Width           =   510
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   12240
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scapre|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   560
            Width           =   510
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   12240
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "Tipo Facturación|N|N|||scapre|tipofact||N|"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1200
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scapre|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   562
            Width           =   4070
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo"
            Height          =   255
            Index           =   52
            Left            =   5640
            TabIndex        =   176
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   6600
            Picture         =   "frmFacEntOferSAIL.frx":1693
            ToolTipText     =   "Buscar forma de pago"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Actuacion"
            Height          =   255
            Index           =   6
            Left            =   5640
            TabIndex        =   116
            Top             =   600
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmFacEntOferSAIL.frx":1795
            ToolTipText     =   "Buscar población"
            Top             =   914
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   1
            Left            =   5640
            TabIndex        =   103
            Top             =   210
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            Picture         =   "frmFacEntOferSAIL.frx":1897
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   102
            Top             =   1266
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   100
            Top             =   914
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   99
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   98
            Top             =   210
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmFacEntOferSAIL.frx":1999
            ToolTipText     =   "Buscar cliente varios"
            Top             =   220
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   97
            Top             =   1610
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5640
            TabIndex        =   96
            Top             =   915
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            Picture         =   "frmFacEntOferSAIL.frx":1A9B
            ToolTipText     =   "Buscar agente"
            Top             =   915
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5640
            TabIndex        =   94
            Top             =   1275
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   11040
            TabIndex        =   93
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   11040
            TabIndex        =   92
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   11040
            TabIndex        =   91
            Top             =   960
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            Picture         =   "frmFacEntOferSAIL.frx":1B9D
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1275
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   89
            Top             =   562
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Plazo Entrega 2|T|S|||scapre|plazos02||N|"
         Top             =   740
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Validez de la oferta|T|S|||scapre|plazos03||N|"
         Top             =   1100
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   21
         Tag             =   "Plazo Entrega 1|T|S|||scapre|plazos01||N|"
         Top             =   450
         Width           =   7845
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   74
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   10080
         MaxLength       =   12
         TabIndex        =   82
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   78
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   4080
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
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   77
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   76
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   2355
         Index           =   32
         Left            =   -72000
         MultiLine       =   -1  'True
         TabIndex        =   35
         Tag             =   "Gestión Oferta|T|S|||scapre|seguiofe||N|"
         Text            =   "frmFacEntOferSAIL.frx":1C9F
         Top             =   3360
         Width           =   9285
      End
      Begin VB.TextBox Text1 
         Height          =   1755
         Index           =   31
         Left            =   -72000
         MultiLine       =   -1  'True
         TabIndex        =   34
         Tag             =   "Concepto Oferta|T|S|||scapre|concepto||N|"
         Text            =   "frmFacEntOferSAIL.frx":1CA7
         Top             =   720
         Width           =   9285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6240
         MaxLength       =   16
         TabIndex        =   75
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   4080
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
         TabIndex        =   73
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
         TabIndex        =   72
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   30
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   33
         Tag             =   "Observación 5|T|S|||scapre|observa05||N|"
         Top             =   4100
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación 4|T|S|||scapre|observa04||N|"
         Top             =   3830
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   28
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observación 3|T|S|||scapre|observa03||N|"
         Top             =   3560
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   27
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 2|T|S|||scapre|observa02||N|"
         Top             =   3290
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   26
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 1|T|S|||scapre|observa01||N|"
         Top             =   3020
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   25
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Asunto Carta 5|T|S|||scapre|asunto05||N|"
         Top             =   2600
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   24
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Asunto Carta 4|T|S|||scapre|asunto04||N|"
         Top             =   2330
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   23
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Asunto Carta 3|T|S|||scapre|asunto03||N|"
         Top             =   2060
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Asunto Carta|T|S|||scapre|asunto02||N|"
         Top             =   1790
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Asunto Carta 1|T|S|||scapre|asunto01||N|"
         Top             =   1520
         Width           =   7845
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntOferSAIL.frx":1CAF
         Height          =   3705
         Left            =   240
         TabIndex        =   83
         Top             =   2415
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   6535
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   16777215
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
      Begin VB.Label Label1 
         Caption         =   "PDFs"
         Height          =   255
         Index           =   47
         Left            =   -74400
         TabIndex        =   170
         Top             =   600
         Width           =   735
      End
      Begin VB.Image ImgPDF 
         Height          =   240
         Index           =   1
         Left            =   -73200
         Picture         =   "frmFacEntOferSAIL.frx":1CC4
         ToolTipText     =   "Modificar"
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgPDF 
         Height          =   240
         Index           =   2
         Left            =   -72840
         Picture         =   "frmFacEntOferSAIL.frx":26C6
         ToolTipText     =   "Eliminar accion"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image ImgPDF 
         Height          =   240
         Index           =   0
         Left            =   -73560
         Picture         =   "frmFacEntOferSAIL.frx":30C8
         ToolTipText     =   "Nuevo documento"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   615
         Index           =   46
         Left            =   -74400
         TabIndex        =   129
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   10
         Left            =   -74640
         Picture         =   "frmFacEntOferSAIL.frx":3ACA
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   4800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   -74520
         Picture         =   "frmFacEntOferSAIL.frx":3BCC
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   8
         Left            =   -74520
         Picture         =   "frmFacEntOferSAIL.frx":3CCE
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Atm"
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
         Index           =   44
         Left            =   -67560
         TabIndex        =   127
         Top             =   5340
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Caudal "
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
         Index           =   43
         Left            =   -67560
         TabIndex        =   126
         Top             =   4860
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Atm"
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
         Index           =   40
         Left            =   -67560
         TabIndex        =   125
         Top             =   3660
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Caudal "
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
         Left            =   -67560
         TabIndex        =   124
         Top             =   3180
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Atm"
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
         Index           =   28
         Left            =   -67560
         TabIndex        =   123
         Top             =   2100
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Caudal "
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
         Index           =   27
         Left            =   -67560
         TabIndex        =   122
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
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
         Index           =   24
         Left            =   -62640
         TabIndex        =   121
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Caudal / Presion"
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
         Left            =   -66240
         TabIndex        =   120
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
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
         Index           =   18
         Left            =   -73920
         TabIndex        =   119
         Top             =   720
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   -74520
         X2              =   -61800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Validez Oferta"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   113
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Oferta"
         Height          =   255
         Index           =   37
         Left            =   -74520
         TabIndex        =   85
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Gestión Oferta"
         Height          =   255
         Index           =   38
         Left            =   -74490
         TabIndex        =   84
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74400
         TabIndex        =   71
         Top             =   3020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto Carta"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   69
         Top             =   1520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo Entrega"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   68
         Top             =   450
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12930
      TabIndex        =   62
      Top             =   7680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
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
            Object.ToolTipText     =   "Lineas Oferta"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Pedido"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cargar Plantilla"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Traer de Oferta"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pasar a historico"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recordatorio"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Valoración"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Oferta"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Fact. Pro forma"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         Left            =   10920
         MaxLength       =   15
         TabIndex        =   112
         Text            =   "TOTAL"
         Top             =   120
         Width           =   1490
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   12480
         MaxLength       =   15
         TabIndex        =   111
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8040
         TabIndex        =   66
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   11
      Left            =   3840
      ToolTipText     =   "Buscar forma de pago"
      Top             =   7560
      Width           =   315
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
      Left            =   9360
      TabIndex        =   114
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   70
      Top             =   7560
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
      Begin VB.Menu mnGenPedido 
         Caption         =   "&Generar Pedido"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnPlantillas 
         Caption         =   "&Plantillas"
         HelpContextID   =   2
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnOferta 
         Caption         =   "Traer &Oferta"
         HelpContextID   =   2
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpOferta 
            Caption         =   "&Oferta"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnImpRecordatorio 
            Caption         =   "&Recordatorio"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnImpValoracion 
            Caption         =   "&Valoración"
            HelpContextID   =   2
            Shortcut        =   ^V
         End
         Begin VB.Menu mnImpFactProF 
            Caption         =   "&Factura Pro Forma"
            Shortcut        =   ^T
         End
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
Attribute VB_Name = "frmFacEntOferSAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schpre, y solo en modo de consulta


Public DatosOferta As String   'Para situarla

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
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
Private WithEvents FrmArt As frmAlmArticulos   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmOC As frmObraCapitulo   'Capitulos
Attribute frmOC.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer 'Listados para Ofertas
Attribute frmList.VB_VarHelpID = -1
Private WithEvents FrmArtEul As frmAlmArticuEUL   'Form Articulos
Attribute FrmArtEul.VB_VarHelpID = -1


'Carga de Plantillas en la linea de la Oferta
Private WithEvents frmPlant As frmFacCargaPlantilla  'Form para cargar plantillas
Attribute frmPlant.VB_VarHelpID = -1
'Carga las lineas de otra Oferta
Private WithEvents frmTOferta As frmFacTraerOferta
Attribute frmTOferta.VB_VarHelpID = -1

Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1
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
Private Modo As Byte

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim primeravez As Boolean

Dim EsCabecera As Boolean

'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No


Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private CadenaSQL As String 'Para crear consulta de Generar Pedido a partir de la Oferta

Private Ordenacion As String   'ORDER BY de la cadena consulta
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

Dim txtAnterior As String

Dim LineaIntercalar As Integer 'NO reutilizar

Dim PulsadoMas2 As Boolean
' ---- [15/09/2009] (LAURA)
'Dim ElArticulo As String   'para la sdesca
' ----

'EULER.
' Cada oferta, llevara:
' vparam.path & \AÑO\format(codclien,000000) & " " & nomclien & "\format(numofert ,"000000" & " " & referencia
Dim CarpetaDestinoOferta As String










'=====================================================================================


Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        Me.SSTab1.Tab = 1
        PonerFoco Text1(18)
    End If
End Sub

Private Sub cboOpcion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkAceptado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkAceptado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean
Dim CambiaDpto As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
                Me.SSTab1.Tab = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                
                    'Insertar modificar los datos del grupo sumergido
                    GrupoSumergidoBD
                
                    CadenaSQL = ""
                    
                    CambiaDpto = False
                    If Text1(12).Text = "" Then
                        If Not IsNull(Data1.Recordset!CodDirec) Then CambiaDpto = True
                    Else
                        If Text1(12).Text <> Val(DBLet(Data1.Recordset!CodDirec, "N")) Then CambiaDpto = True
                    End If
                    
                    If CambiaDpto Then
                        If Text2(12).Text = "" Then
                            CadenaSQL = "NULL"
                        Else
                            CadenaSQL = "'" & DevNombreSQL(Text2(12).Text) & "'"
                        End If
                        CadenaSQL = "UPDATE scapre SET nomdirec=" & CadenaSQL & " WHERE"
                        CadenaSQL = CadenaSQL & " numofert= " & Text1(0).Text
                    
                        
                    End If
                    'Actualizar los datos del cliente si es de varios
                    EsDeVarios = EsClienteVarios(Text1(4).Text)
                    If EsDeVarios Then ActualizarClienteVarios Text1(4).Text, Text1(6).Text
                    TerminaBloquear
                    
                    If CadenaSQL <> "" Then
                        Espera 0.2
                        ejecutar CadenaSQL, False
                        CadenaSQL = ""
                    End If
                    PosicionarData
                    
                    
                End If
            End If
            
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'slima1' (Revisiones)
            If ModificaLineas = 1 Then 'INSERTAR lineas Ofertas
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
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
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
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


Private Sub cmdAux_Click(Index As Integer)

    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
        
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
                FrmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo Busqueda
                FrmArt.DeConsulta = True
                FrmArt.Show vbModal
                Set FrmArt = Nothing
            End If
            PonerFoco txtAux(1)
        Case 2
             Set frmOC = New frmObraCapitulo
            frmOC.DatosADevolverBusqueda = "0|1|"
            frmOC.Show vbModal
            Set frmOC = Nothing
        
    End Select
    If Index < 2 Then
        PonerFoco txtAux(Index)
    Else
        PonerFoco txtAux(9)
    End If
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
                DataGrid1.AllowAddNew = False
                ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                DataGrid1.Enabled = True
            End If
            LineaIntercalar = 0
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    BloquearTxt Text1(0), True, True
    
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea(Intercalando As Boolean)


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    Precio = ""
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    Me.SSTab1.Tab = 0
    If Intercalando Then
        lblIndicador.Caption = "** INTERCALAR **"
        If Not Data2.Recordset.EOF Then
            LineaIntercalar = Data2.Recordset!numlinea
        End If
    Else
        lblIndicador.Caption = "INSERTAR"
    End If
    lblIndicador.Refresh
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    Me.cboOpcion.ListIndex = -1
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(1, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    txtAux(9).Text = ""
    
    
    'Para recordar que estamos intercalando
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
End Sub


Private Sub BotonVerTodos()
'Ver todos
    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia2 "", False
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
    Me.SSTab1.Tab = 0
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
Dim vTipoMov As CTiposMov
Dim NumOferElim As Long

    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    cad = "Cabecera de Ofertas." & vbCrLf
    cad = cad & "-----------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Oferta:            "
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumOferElim = Data1.Recordset.Fields(0).Value
        
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
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumOferElim
        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Oferta", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    Me.SSTab1.Tab = 0
    ModificaLineas = 3 'Eliminar
    
    SQL = "¿Seguro que desea eliminar la línea de Oferta?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    ' ---- [15/09/2009] (LAURA)
'    ElArticulo = Data2.Recordset!codArtic
    ' ----
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & " WHERE " & ObtenerWhereCP
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        ' ---- [15/09/2009] (LAURA)
'        DescuentosCantidad ElArticulo
        ' ----
        Text2(16).Text = ""
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdPDF_Click()
    If Not AcroPDF1.visible Then Exit Sub
    If AcroPDF1.src <> "" Then LanzaVisorMimeDocumento Me.hwnd, AcroPDF1.src
    'IF ME.AcroPDF1
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        
        '25 Enero 2010
        'probando. Falta conmprobar que ha ido bien desde entonces
        'If DataGrid1.Row >= 0 Then
        '    DeseleccionaGrid DataGrid1
        '    DataGrid1.Bookmark = 1
        'End If
        
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
        cmdCancelar.Cancel = True
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
'        If Data1.Recordset.EOF Then
'            MsgBox "Ningún registro devuelto.", vbExclamation
'            Exit Sub
'        End If
'        Cad = Data1.Recordset.Fields(0) & "|"
'        Cad = Cad & Data1.Recordset.Fields(1) & "|"
'        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo EKeyPress
    
    If KeyAscii = 27 Then 'ESC
        If Modo = 5 Then 'Modo Lineas
            cmdRegresar_Click
        ElseIf Modo = 0 Or Modo = 2 Then 'Estamos en Cabecera
            Unload Me
        End If
    End If
    
EKeyPress:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If (Modo <> 2 And Modo <> 5) Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        If X > 7750 And X < 8000 Then
            Select Case DataGrid1.Columns(8).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
        Else
            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        'Poner descripcion de ampliacion lineas
        Text2(16).Text = DevuelveDesdeBDNew(1, NomTablaLineas, "ampliaci", "numofert", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
    Else
        Text2(16).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_activate()
        If primeravez Then
        primeravez = False
        If InstalacionEsEulerTaxco Then AccesoA_Carpeta
        If DatosOferta <> "" Then
            PonerModo 1
            Text1(0).Text = DatosOferta
            HacerBusqueda
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub AccesoA_Carpeta()
    On Error Resume Next
    If Dir(EulerParam & "\Ofertas\", vbDirectory) = "" Then MsgBox "No existe la carpeta: " & EulerParam & "\Ofertas\", vbExclamation
    Err.Clear
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    primeravez = True
    
    
    Me.imgBuscar(11).Picture = frmPpal.imgListComun.ListImages(19).Picture
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 25
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Pedido
        .Buttons(12).Image = 32 'Cargar Plantilla
        .Buttons(13).Image = 24 'Traer Lineas de Otra Oferta
        
        .Buttons(15).Image = 13 'Recordatorio
        .Buttons(16).Image = 30 'Recordatorio
        .Buttons(17).Image = 31 'Valoracion
        .Buttons(18).Image = 16 'Imprimir
        .Buttons(19).Image = 40 'Imprimir factura pro forma
        
        .Buttons(22).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
'    CargarComboTipoPago
    CargarComboFacturacion
    CodTipoMov = "OFE"
    VieneDeBuscar = False 'Para el CPostal
   
    SSTab1.TabVisible(4) = InstalacionEsEulerTaxco    'vParamAplic.NumeroInstalacion = vbEuler
    
    If Not InstalacionEsEulerTaxco Then Label1(51).Caption = "F. entrega"  'vParamAplic.NumeroInstalacion <> vbEuler
    
   
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = DevuelveTextoDepto(True)
    If vParamAplic.TieneCRM Then
        Label1(46).Caption = "Observaciones CRM"
    Else
        Label1(46).Caption = "Obs. internas"
    End If
    '## A mano
    If Not EsHistorico Then
        NombreTabla = "scapre"
        NomTablaLineas = "slipre" 'Tabla lineas de Ofertas
        Me.Caption = "Ofertas Clientes"
    Else
        NombreTabla = "schpre"
        NomTablaLineas = "slhpre"
        CargarTagsHco Me, "scapre", NombreTabla
        Me.Caption = "Histórico Ofertas Clientes"
    End If
    Ordenacion = " ORDER BY numofert "


    Label1(52).visible = EsHistorico
    cboMotTra.visible = EsHistorico
    If EsHistorico Then CargarCombo_Tabla cboMotTra, "smotitrasofer", "codmotiv", "descmotivo", , False, "descmotivo"

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where numofert=-1"
    Data1.Refresh
'    If DatosADevolverBusqueda = "" Then
    If Me.DatosOferta = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    primeravez = True

    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    
    

    primeravez = True
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkAceptado(0).Value = False
    Me.cboFacturacion.ListIndex = -1
    cmdPDF.visible = False
    Text3(0).Text = "BASE IMP."
    Me.ListView1.ListItems.Clear
    ListView1_Click
    If EsHistorico Then Me.cboMotTra.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(33).Text = RecuperaValor(CadenaSeleccion, 3)
    Text2(33).Text = RecuperaValor(CadenaSeleccion, 4) & "  " & RecuperaValor(CadenaSeleccion, 5)
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub


Private Sub FrmArtEul_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                CadB = CadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Precio = CadenaDevuelta
            'Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            'Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
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
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
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


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen el Listado de Oferta para generar el Pedido
Dim vSQL As String

    'Construimos parte de la SQL para insertar en Pedidos
    vSQL = ""
    vSQL = " '" & Format(RecuperaValor(CadenaSeleccion, 2), FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    vSQL = vSQL & Format(RecuperaValor(CadenaSeleccion, 3), FormatoFecha) & "' as fecentre, " 'Fecha entrega
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 4) & " as sementre, " 'Sem entrega
    vSQL = vSQL & "0 as visadore, " & "codclien, nomclien, domclien, codpobla, pobclien, proclien, nifclien, "
    vSQL = vSQL & "telclien, coddirec, nomdirec, referenc, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Operador de Pedido
    vSQL = vSQL & "codagent, codforpa, dtoppago, dtognral, tipofact, observa01, observa02, observa03, "
    vSQL = vSQL & "observa04, observa05, 0 as servcomp,0 as restoped, " & Text1(0).Text & " as numofert, '" 'Nº Oferta
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecofert " 'Fecha Oferta. El resto de campos pasa en PASarOfertaPedido
    CadenaSQL = vSQL
End Sub


Private Sub frmOC_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Forma Pago
    txtAux(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom
End Sub

Private Sub frmPlant_CargarPlantillas()
Dim RS As ADODB.Recordset
Dim RSLineas As ADODB.Recordset
Dim SQL As String, devuelve As String
Dim codAlmac As String
'codTarif As String
Dim cantidad As Integer
Dim NumCajas As Integer, RestoUnid As Integer
Dim Precio As String, Dto1 As String, Dto2 As String
Dim OrigP As String 'De donde viene el precio: promocion, precio especial,...
Dim CPrecioFact As CPreciosFact
Dim COntadorLInea As Integer

    Screen.MousePointer = vbHourglass
    
    'Si se ha seleccionado alguna plantilla para añadir sus lineas a la Oferta
    '(cantidad de alguna linea de tmpscapla > 0), entonces añadimos todas las
    'lineas de esa oferta poniendo en cantidad de slipre de lineas de oferta
    'el resultado de multiplicar la cantidad de tmpscapla * cantidad de slipla
    SQL = "SELECT * FROM tmpscapla WHERE codusu=" & vUsu.Codigo & " AND cantidad>0 ORDER BY codplant"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Obtener el almacen por defecto del trabajador
    'Poner el Almacen por defecto del Trabajador
    codAlmac = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    
    
    'Obtener la tarifa del cliente. LO Fijare en la funcion
    'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")



    'Consigo el contador
    COntadorLInea = Val(SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP))
    
    While Not RS.EOF  'Para cada plantilla
        'Añadimos todas las lineas de esa plantilla en la cantidad correcta en las
        'lineas de la Oferta
        SQL = "SELECT * FROM slipla WHERE codplant=" & RS!codplant & " order by numlinea"
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RSLineas.EOF
            'Comprobar si el articulo se vende por cajas antes de entrar a la función
            devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RSLineas!codArtic, "T")
            If devuelve <> "" Then
            'Si se puede vender por cajas(devuelve>1) poner numero de unidades/caja en una linea con el
            'precio de caja, y otra linea con el resto unidades un precio unidad
                cantidad = (RS!cantidad * RSLineas!cantidad)
                NumCajas = ObtenerNumCajas(CStr(cantidad), devuelve)
                RestoUnid = CInt(cantidad) - NumCajas * CInt(devuelve)
                'Obtener el precio a aplicar
                Set CPrecioFact = New CPreciosFact
                'CPrecioFact.CodigoLista = codTarif
                CPrecioFact.CodigoArtic = RSLineas!codArtic
                CPrecioFact.CodigoClien = Text1(4).Text
                CPrecioFact.FijarTarifaActividad
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, "")
                
                'Si PorCaja vuelve de ObtenerPrecio a False se aplica precio
                'de Unidad aunque se venda por cajas, ya que ha regresado con pvp de articulo
                Dto1 = CPrecioFact.Descuento1
                Dto2 = CPrecioFact.Descuento2
                Set CPrecioFact = Nothing
                    
                If PorCaja And NumCajas > 0 Then 'El Articulo se Vende Por Cajas y Cantidad supera la cant en 1 caja
                    'Obtener el precio y los descuentos adecuados
                    'Insertar 2 lineas: 1 linea con la cantidad que se puede vender en cajas y al precio de caja
                    InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, NumCajas * CInt(devuelve), Precio, Dto1, Dto2, OrigP, COntadorLInea
                    '2 linea con el resto de la cantidad que no llega a una caja a precio de unidad
                    If RestoUnid > 0 Then InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, RestoUnid, Precio, Dto1, Dto2, OrigP, COntadorLInea
'                    Else
'                        InsertarLineaDePlantilla rsLineas!codArtic, codAlmac, codTarif, Cantidad, 0
'                    End If
                Else 'No llega a una caja
                    InsertarLineaDePlantilla RSLineas!codArtic, codAlmac, cantidad, Precio, Dto1, Dto2, OrigP, COntadorLInea
                End If
            End If
            RSLineas.MoveNext
        Wend
        RSLineas.Close
        Set RSLineas = Nothing
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

    'Borrar de la Tabla Temporal (tmpscapla) los registros insertados tras añadir
    'las lineas de las plantillas seleccionadas
    DescargarDatosTMP
    'Actualizar el Grid con las lineas Añadidas
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub frmTOferta_CargarOferta2(NumOfe As String)
Dim RS As ADODB.Recordset
Dim SQL As String
Dim numlinea As String, vWhere As String
Dim I  As Integer
Dim CopiaDesdeHco As Boolean
    On Error GoTo ECargarOferta
    
    'Si se ha seleccionado alguna oferta para añadir sus lineas a la Oferta
    If NumOfe = "" Then Exit Sub
    CopiaDesdeHco = False
    If Mid(NumOfe, 1, 1) = "H" Then
        'Desde hco
        CopiaDesdeHco = True
        NumOfe = Mid(NumOfe, 2)
    End If
    
    Screen.MousePointer = vbHourglass
    If CopiaDesdeHco Then
        SQL = DevuelveDesdeBD(conAri, "distinct(fecofert)", "slhpre", "numofert", RecuperaValor(NumOfe, 1))
        'por si hubiera mas de una fecha. Solo cojo una
        
        SQL = "Select * from slhpre where numofert=" & RecuperaValor(NumOfe, 1) & " AND fecofert='" & Format(SQL, FormatoFecha) & "'"
    Else
        SQL = "Select * from " & NomTablaLineas & " where numofert=" & RecuperaValor(NumOfe, 1)
    End If
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF  'Para cada linea de oferta
        'Obtener el siguiente numero de linea
        vWhere = ObtenerWhereCP
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        SQL = "INSERT INTO " & NomTablaLineas & " (numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,origpre,esopcion) "
        SQL = SQL & " VALUES(" & Text1(0).Text & ", " & numlinea & ", " & RS!codAlmac & ", " & DBSet(RS!codArtic, "T") & ", " & DBSet(RS!NomArtic, "T") & ", "
        SQL = SQL & DBSet(RS!Ampliaci, "T", "S")
        SQL = SQL & ", " & DBSet(RS!cantidad, "N") & ", " & DBSet(RS!precioar, "N") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", " & DBSet(RS!ImporteL, "N") & ", "
        SQL = SQL & DBSet(CStr(RS!origpre), "T", "S") & ","
        SQL = SQL & RS!esopcion & ")"
        
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    
    
    SQL = RecuperaValor(NumOfe, 2)  'Copio observaciones
    vWhere = RecuperaValor(NumOfe, 3)  'Copio datos carta
    I = Val(SQL) + Val(vWhere)
    If I > 0 Then
        'Cargo en RS la oferta
        If CopiaDesdeHco Then
            SQL = DevuelveDesdeBD(conAri, "distinct(fecofert)", "slhpre", "numofert", RecuperaValor(NumOfe, 1))
            'por si hubiera mas de una fecha. Solo cojo una
            SQL = "Select * from schpre where numofert=" & RecuperaValor(NumOfe, 1) & " AND fecofert='" & Format(SQL, FormatoFecha) & "'"
        Else
            SQL = "Select * from " & NombreTabla & " where numofert=" & RecuperaValor(NumOfe, 1)
        End If
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            'UPDATEAMOS los campos de la oferta de observaciones
            SQL = ""
            If RecuperaValor(NumOfe, 2) = "1" Then 'Copio observaciones
                
                
                For I = 1 To 5
                    vWhere = "observa0" & I
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(RS.Fields(vWhere), "T"), "T", "S")
                    SQL = SQL & numlinea
                Next I
                
                '15 Marzo 2010. Cunado pone copiar observacinoes TB tiene que copiar el campo concepto
                SQL = SQL & ", concepto = " & DBSet(DBLet(RS!Concepto, "T"), "T")
     
            End If
            
            If RecuperaValor(NumOfe, 3) = "1" Then 'Copio carta
                For I = 1 To 5
                    vWhere = "asunto0" & I
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(RS.Fields(vWhere), "T"), "T", "S")
                    SQL = SQL & numlinea
                Next I
                
            End If
       
            SQL = Mid(SQL, 2)  'quito la primera coma
            SQL = SQL & " WHERE numofert = " & Text1(0).Text
            SQL = "UPDATE " & NombreTabla & " SET " & SQL
            RS.Close
        conn.Execute SQL
        PosicionarData  'vuelvo a cargar los datos
        PonerCampos
        Else
            MsgBox "Error buscando oferta destino: " & Text1(0).Text & ".  EOF", vbExclamation
        End If
    End If
    
    
    Set RS = Nothing

    'Actualizar el Grid con las lineas Añadidas
    If I = 0 Then CalcularDatosFactura   'Si no mete obser y carta que carge los totales
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
    
ECargarOferta:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Traer de otra Oferta.", Err.Description
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then
        If Index <> 11 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    TerminaBloquear

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            Set frmC = New frmFacClientes
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
                EsCabecera = False
                MandaBusquedaPrevia2 " codclien= " & Val(Text1(4).Text), True
                Indice = 12
             End If
             
        Case 3 'Realizada Por Trabajador
            Indice = 3
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
            
        Case 7
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
            
            Indice = 7
            Set frmAc = New frmObraActua
            frmAc.DatosADevolverBusqueda = Text1(4).Text & "|" & Text1(12).Text & "|"
            frmAc.Show vbModal
            Set frmAc = Nothing
            
            
       Case 8, 9, 10
            EsCabecera = False
            MandaBusquedaPrevia2 CStr(Index), False
            
            
            
        Case 11
      
                
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
    If Indice = 7 Then
        PonerFoco Text1(15)
    Else
        If Indice <> 2 Then
            If Index >= 8 Then
            
            Else
                PonerFoco Text1(Indice)
            End If
        End If
    End If
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




Private Sub ImgPDF_Click(Index As Integer)
    
    If EsHistorico Then Exit Sub
    If Modo <> 2 Then Exit Sub
    
    If Index > 0 Then
        CadenaDesdeOtroForm = ""
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then CadenaDesdeOtroForm = "OK"
        End If
        If CadenaDesdeOtroForm = "" Then
            MsgBox "Seleccione un documento", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If Index = 2 Then
        If MsgBox("Seguro que desea eliminar el documento seleccionado?", vbQuestion + vbYesNo) = vbYes Then
            Precio = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
            If EliminarArhivoPDFOferta(Precio) Then
                'Precio = "DELETE FROM `sliprepdfs` WHERE numofert =" & Text1(0).Text
                'Precio = Precio & " AND numlinea = " & Mid(ListView1.SelectedItem.Key, 2)
                'ejecutar Precio, False
                
                cargaDocumentos
            End If
        
        
        End If
    Else
        If Index = 1 Then Exit Sub
        If CarpetaDestinoOferta = "" Then
          '  If Year(Data1.Recordset!fecofert) < 2017 Then
          '      MsgBox "Carpeta destino no esta establecida", vbExclamation
          '      Exit Sub
          '  End If
            
            'Creamos la estructura (Deberia haber sido creada con anterioridad, al dar nuevo. Pero para las de 2016 dejo pasar
            CrearCarpeta
            
        End If
        CadenaDesdeOtroForm = ""
        If Index = 1 Then CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
        frmEuler.Opcion = Index
        frmEuler.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 0 Then
                'Insertamos documento
                Precio = RecuperaValor(CadenaDesdeOtroForm, 1)
                If Dir(Precio, vbArchive) = "" Then
                    MsgBox "No existe fichero", vbExclamation
                    
                Else
                    
                    If CarpetaDestinoOferta = "" Then
                        'Crear estructura
                        
                        MsgBox "falta Crear estructura", vbExclamation
                        Exit Sub
                    End If
                    Precio = RecuperaValor(CadenaDesdeOtroForm, 2)
                    txtAnterior = NombreArchivoEULER(Precio)
                
                    If CopiaArhivoPDFOfertaEuler(CarpetaDestinoOferta, RecuperaValor(CadenaDesdeOtroForm, 1)) Then
                        Precio = DevuelveDesdeBD(conAri, "max(numlinea)", "sliprePdfs", "numofert", Text1(0).Text)
                        If Precio = "" Then Precio = "0"
                        Precio = CStr(Val(Precio) + 1)
                
                        'INSERT INTO BD
                        Precio = "insert into `sliprepdfs` (`numofert`,`numlinea`,`ficheroDesc`,`ficheroNombre`) values ( " & Text1(0).Text & "," & Precio & ","
                        Precio = Precio & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & "," & DBSet(txtAnterior & ".pdf", "T") & ")"
                        'ejecutar Precio, False
                        
                        cargaDocumentos
                    End If
                
                End If
            Else
                'Modificar
                Precio = "UPDATE sliprepdfs SET ficheroDesc = " & DBSet(CadenaDesdeOtroForm, "T")
                Precio = Precio & " WHERE numofert = " & Text1(0).Text & " AND numlinea = " & Mid(ListView1.SelectedItem.Key, 2)
                ejecutar Precio, False
                ListView1.SelectedItem.Text = CadenaDesdeOtroForm
            End If
        End If
    End If
End Sub

Private Sub ListView1_Click()
Dim cad As String
Dim C As String
    cad = ""
    If CarpetaDestinoOferta <> "" Then
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then
                If LCase(Right(ListView1.SelectedItem.SubItems(1), 3)) = "pdf" Then cad = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
            End If
        End If
    End If
    
    CargaArchivo cad
    cmdPDF.visible = cad <> ""
    
        
End Sub

Private Sub ListView1_DblClick()
Dim C As String
    
    C = ""
    If CarpetaDestinoOferta <> "" Then
        If ListView1.ListItems.Count > 0 Then
            If Not ListView1.SelectedItem Is Nothing Then C = CarpetaDestinoOferta & ListView1.SelectedItem.SubItems(1)
        
        End If
    End If
    
    
    If C <> "" Then LanzaVisorMimeDocumento Me.hwnd, C
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub


Private Sub mnGenPedido_Click()
'Pasar una Oferta a Pedido
Dim devuelve As String
Dim LineasQueSepasan As String

    'Comprobar que hay una Oferta seleccionada
    If Text1(0).Text = "" Then Exit Sub
    
    
    '23 Abril 2010
    'NO puede tener la fecha entraga a NULO
    If Trim(Text1(2).Text) = "" Then
        MsgBox "Fecha entrega no puede ser nula", vbExclamation
        Exit Sub
    End If
    
    '17 Diciembre 2010
    If EsClienteBloqueado(Text1(4).Text, True, False) Then Exit Sub
    
    
    'Comprobar que la Oferta seleccionada esta aceptada
    devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "aceptado", "numofert", Text1(0).Text, "N")
    If devuelve = "0" Then
        MsgBox "La Oferta debe estar Aceptada para pasar a Pedido.", vbExclamation
        Exit Sub
    End If
    If devuelve = "1" Then
    
        'Agosto2011
        '---------------------------------------------
        'Las ofertas pueden llevar opciones.
        'Si llevan opciones mostraremos un lw con las opciones
        'LineasQueSepasan: Llevara las lineas que SI se van a pasar
        '                   Si esta vacio son todas
        LineasQueSepasan = ""
        If TieneOpciones2(False) Then
        
            '------------------------------------------------------------
            '
            CadenaDesdeOtroForm = ""
            frmFacTrasOfertaOpciones.Caption = Text1(0).Text
            frmFacTrasOfertaOpciones.Label1.Caption = Text1(5).Text
            frmFacTrasOfertaOpciones.Show vbModal
            If CadenaDesdeOtroForm = "NO" Then Exit Sub
            If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)                  'QUITO LA PRIMERA COMA
            LineasQueSepasan = CadenaDesdeOtroForm
            
            
        End If
    
    
    
        'Pedir: Operador de Pedido, fecha pedido, y fecha entrega (calcular semana)
'        AbrirListadoOfer (37) '37: Pedir datos para Pedido (NO IMPRIME LISTADO)
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 37
        frmList.codClien = Text1(4).Text
        frmList.FecEntre = Text1(2).Text
        frmList.Show vbModal
        Set frmList = Nothing
        
        'Tenemos en CadenaSQL parte de la SELECT para insertar el Pedido
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        PasarOfertaAPedido CadenaSQL, LineasQueSepasan

        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            LimpiarDataGrids
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnImpFactProF_Click()
'Imprime factura pro forma
    BotonImprimirProForma (59) '59: Informe Factura ProForma
End Sub

Private Sub mnImpOferta_Click()
'Imprime una Oferta
       davidCodtipom = CarpetaDestinoOferta    'reutililizo
       frmListadoOfer.NumCod = Text1(0).Text   'Nº de Oferta
       frmListadoOfer.FecEntre = Text1(1).Text 'Fecha de Oferta
       If EsHistorico Then
            AbrirListadoOfer (35) '35: Informe Historico de Ofertas
       Else
            AbrirListadoOfer (31) '31: Informe de Ofertas
       End If
       davidCodtipom = ""
End Sub

Private Sub mnImpRecordatorio_Click()
    frmListadoOfer.NumCod = Text1(0).Text
    frmListadoOfer.codClien = Text1(4).Text
    AbrirListadoOfer (32) '32: Recordatorio de Ofertas
End Sub

Private Sub mnImpValoracion_Click()
    frmListadoOfer.codClien = Text1(4).Text
    frmListadoOfer.NumCod = Text1(0).Text 'Nº de Oferta
    AbrirListadoOfer (33) '33: Valoracion de Ofertas
End Sub

Private Sub mnLineas_Click()
    
    BotonMtoLineas 0, "Ofertas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea False
    Else 'Añadir Cabecera de Ofertas
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub


Private Sub mnOferta_Click()
'Añadir las lineas de otra oferta a la Oferta
    Set frmTOferta = New frmFacTraerOferta
    frmTOferta.Show vbModal
    Set frmTOferta = Nothing
End Sub

Private Sub mnPlantillas_Click()
'Añadir Plantilla de Oferta
    Set frmPlant = New frmFacCargaPlantilla
    frmPlant.Show vbModal
    Set frmPlant = Nothing
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
    Label1(35).visible = SSTab1.Tab = 0
    imgBuscar(11).visible = SSTab1.Tab = 0
    Text2(16).visible = Modo > 3 Or SSTab1.Tab = 0
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer
    txtAnterior = Text1(Index).Text
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False
    If (Index <> 31 And Index <> 32) Then ConseguirFoco Text1(Index), Modo, cadkey
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 30 And KeyCode = 40 Then
        Me.SSTab1.Tab = 2
        PonerFoco Text1(31)
    Else
        If Not Text1(Index).MultiLine Then KEYdown KeyCode
        
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
                Case 33
                    Ind = 7
                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
        
        
        
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then
        If KeyAscii = 13 And (Index = 30 Or Index = 32) Then 'ENTER
            If Index = 32 Then
    '            PonerFocoBtn Me.cmdAceptar
            ElseIf Index = 30 Then
                Me.SSTab1.Tab = 2
                PonerFoco Text1(31)
            ElseIf Index = 32 Then
                Me.SSTab1.Tab = 3
                PonerFoco Text4(0)
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
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
        Case 1, 2 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)
            If Index = 2 And Text1(Index).Text <> "" Then 'Fecha Entrega
                'Comprobar que es posterior a la del pedido
                If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
            End If
            
        Case 3 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else
                    PonerDatosCliente (Text1(Index).Text)
                End If
            Else
                LimpiarDatosCliente
            End If
            If Modo <> 1 Then
                'Si no estamos en busqueda, y se ha equivado poniendo cliente... vuelve a cliente
                If Text1(Index).Text = "" Then PonerFoco Text1(Index)
            End If
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(6).Text = Data1.Recordset!nifClien Then
'                    Exit Sub
'                End If
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
                'Comprobar que el cliente tiene mantenimientos en esa dired/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then
                ComprobarRefObligatoria
            End If
            
        Case 14 'Forma de Pago
            If Me.SSTab1.Tab = 0 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then  'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 33
            If Text1(33).Text = "" Then
                Text2(33).Text = ""
            Else
                PonerCampoActuacion
                If Text1(33).Text = "" Then PonerFoco Text1(33)
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia2 CadB, False
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia2(CadB As String, Dpto As Boolean)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
Dim I As Integer 'Para articuo
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
            cad = cad & ParaGrid(Text1(0), 15, "Nº Oferta")
            cad = cad & ParaGrid(Text1(1), 20, "Fecha Ofer.")
            cad = cad & ParaGrid(Text1(4), 15, "Cliente")
            cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
            tabla = NombreTabla
            If EsHistorico Then
                Titulo = "Histórico de Ofertas"
                devuelve = "0|1|"
            Else
                Titulo = "Ofertas"
                devuelve = "0|"
            End If
'            devuelve = "0|"
    Else 'Llama desde lineas, para cargar solo los depart/direc. del cliente seleccionado
        If Dpto Then
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            Else
                Titulo = "Obras Cliente: "
                Desc = "Obras"
            End If
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
            cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
            tabla = "sdirec"
            devuelve = "0|1|"
        Else
            I = CInt(CadB)
            I = I - 8 '0,1,2
            I = I * 8
            CadB = ""
           'Articulo para Gtupo sumergido
            Titulo = "ARTICULOS: "
            cad = cad & "Cod-Ref " & "|sartic|Codartic|T||27·"
            cad = cad & "Descripcion|sartic|nomartic|F||70·"
            'Cad = Cad & "Obs|sactuaobra|observa|T||55·"
            tabla = "sartic"
            devuelve = "0|1|"
        
        End If
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
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
                  Text4(I).Text = RecuperaValor(Precio, 2)
                  PonerFoco Text4(I)
                End If
                Precio = ""
            End If
            If Dpto Then PonerFoco Text1(12)
            
            
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
        If Modo = 1 Then
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(kCampo)
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

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

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
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    PonerCampoActuacion
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    CalcularDatosFactura
    
    
    PonerCamposGrupoSumergido True
    cmdPDF.visible = False
    If InstalacionEsEulerTaxco Then cargaDocumentos
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
    lblF.Caption = ""
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    

    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    cmdRegresar.visible = Modo = 5 And ModificaLineas = 0
    
    
        
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Poner Flechas de desplazamiento visibles
    b = (Modo = 2)
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    b = Modo = 3 Or Modo = 4
    BloquearGrupoSumergido Not b
    
    'Poner siempre el campo numOferta (contador) bloqueado, excepto cuando
    'estamos en modo de Busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 56
        BloquearTxt Text3(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    For I = 46 To 48
        Text3(I).BackColor = &HFFFFC0
        Text3(I + 6).BackColor = &HFFFFC0
    Next I
    'Campos total Factura en verde
    Text3(55).BackColor = &HC0FFC0
    Text3(56).BackColor = &HC0FFC0    'Tatal factura
    '---------------------------------------------------
    
    b = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Me.cboFacturacion.Enabled = b
    Me.chkAceptado(0).Enabled = b
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    If EsHistorico Then BloquearCmb Me.cboMotTra, Not b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        If I <> 11 Then Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(1).visible = False
    
       
    'Modo Linea de Ofertas
    Me.Label1(35).visible = SSTab1.Tab = 0
    Me.Text2(16).visible = SSTab1.Tab = 0
    BloquearTxt Text2(16), True
       
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
Dim b As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then Exit Function
   
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
     If Trim(Text1(4).Text) <> "" Then
        devuelve = DevuelveDesdeBDNew(1, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            b = False
        End If
    End If
    
    
    'Febrero 2015
    'NO es obligado la actuacion
    If Text1(33).Text <> "" Then
        'MsgBox "Campo actuacion obligatorio", vbExclamation
        'B = False
        'Else
        If Text2(33).Text = "" Then
            MsgBox "Actuacion incorrecta", vbExclamation
            b = False
        End If
    End If
    
    
    
    If Not b Then Exit Function
          
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte
Dim vArtic As CArticulo
Dim Aux As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
        'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(6), txtAux(7), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(8).Text Then txtAux(8).Text = Aux
    
    
    
    b = True
    For I = 0 To txtAux.Count - 3
        If txtAux(I).Text = "" Then
            MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(I)
            Exit Function
        End If
    Next I
        
    
    'Comprobar que existe de el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        b = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
    
    

    
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_GotFocus(Index As Integer)
     If Index = 16 Then
        'Campo observaciones. NO, repito NO, se selecciona todo
        If Text2(Index).Text <> "" Then
            Text2(Index).Text = Text2(Index).Text & " "
            Text2(Index).SelStart = Len(Text2(Index).Text)
        End If
    End If
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then
      '  PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    'If Index = 16 And KeyAscii = 13 Then 'ENTER
    '    PonerFocoBtn Me.cmdAceptar
    'End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
    If Index = 16 And Modo = 5 And ModificaLineas > 0 Then PonerFocoBtn cmdAceptar
End Sub




Private Sub Text4_GotFocus(Index As Integer)
    If (Index Mod 8) <> 0 Then ConseguirFoco Text4(Index), Modo
End Sub



Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus(Index As Integer)
Dim J As Integer
    Text4(Index).Text = Trim(Text4(Index).Text)
    If Text4(Index).Text = "" Then Exit Sub

    J = Index Mod 8
    Select Case J
    Case 0
        'NADA  nomartic
        
    Case 1 To 6
        If Not PonerFormatoEntero(Text4(Index)) Then Text4(Index).Text = ""
    Case 7
        If Not PonerFormatoDecimal(Text4(Index), 3) Then Text4(Index).Text = ""
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    

    If Button.Index = 17 Or Button.Index = 19 Then
        'Valoracion , fra proforma
        'If TieneOpciones Then Exit Sub
        TieneOpciones2 True   'Siempre dejo seguir
    End If
    
    
    Select Case Button.Index
        Case 1: mnBuscar_Click 'Buscar
        Case 2: BotonVerTodos  'Todos
            
        Case 5: mnNuevo_Click  'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
            
        Case 10: mnLineas_Click  'Lineas
        Case 11:
                If Modo = 5 Then
                    'Insertar intercalando
                    BotonAnyadirLinea True
                Else
                    mnGenPedido_Click 'Generar Pedido
                End If
        Case 12: mnPlantillas_Click ' Plantillas. Solo visible en Mantenimiento Lineas.
        Case 13: mnOferta_Click 'Traer Lineas de Otra Oferta
        
        Case 15:
            PasarAHco
            
        Case 16 'Recordatorio
            mnImpRecordatorio_Click
        Case 17 'Valoracion
            mnImpValoracion_Click
        Case 18 'Imprimir
            mnImpOferta_Click
        Case 19 'Imprimir factura por forma
            mnImpFactProF_Click
        
        Case 22: mnSalir_Click    'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    On Error Resume Next
    
    PonerOpcionesMenuGeneral Me
        
    J = Val(Me.mnGenPedido.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenPedido.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean
    
    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim SQL As String
Dim numlinea As String
    
    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea Then 'Lineas de Ofertas
        'Conseguir el siguiente numero de linea
        
        If LineaIntercalar = 0 Then
            'INSERCION NORMAL
            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP)
        
        Else
            SQL = ObtenerWhereCP
            SQL = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 1 WHERE " & SQL & " and numlinea >= " & LineaIntercalar
            SQL = SQL & " order by numlinea desc " 'Para que empieza por las ultimas
            conn.Execute SQL
            numlinea = LineaIntercalar
        End If
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codcapit,esopcion) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " 'Dto 2
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", " 'Importe
        SQL = SQL & DBSet(txtAux(5).Text, "T") & ","
        'SAIL
        SQL = SQL & DBSet(txtAux(9).Text, "N", "S") & ","  'capitulo
        SQL = SQL & Abs(Me.cboOpcion.ListIndex = 1) & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        'sdesca
'        ElArticulo = txtAux(1).Text
'        DescuentosCantidad ElArticulo
        ' ----
        
    End If
    Exit Function

EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta" & vbCrLf & Err.Description
End Function


Private Function InsertarLineaDePlantilla(codArtic As String, codAlmac As String, cantidad As Integer, Precio As String, Dto1 As String, Dto2 As String, OrigP, ByRef numlinea As Integer) As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim SQL As String
'Dim NumLInea As String
Dim NomArtic As String
Dim Importe As String

    On Error GoTo EInsertarLinea

    InsertarLineaDePlantilla = False
    SQL = ""
    
    'Conseguir el siguiente numero de linea
    
    
    
    SQL = "INSERT INTO " & NomTablaLineas
    SQL = SQL & " (numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre) "
    SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", "
    SQL = SQL & codAlmac & ", " & DBSet(codArtic, "T") & ", "
    NomArtic = DevuelveDesdeBDNew(1, "sartic", "nomartic", "codartic", codArtic, "T")
    SQL = SQL & DBSet(NomArtic, "T") & ", " & ValorNulo & ", " & DBSet(cantidad, "N") & ", "
                   
    Importe = CalcularImporte(CStr(cantidad), Precio, Dto1, Dto2, vParamAplic.TipoDtos)
    SQL = SQL & DBSet(Precio, "N") & ", "
    SQL = SQL & DBSet(Dto1, "N") & ", "
    SQL = SQL & DBSet(Dto2, "N") & ", "
    SQL = SQL & DBSet(Importe, "N") & ", '"
    SQL = SQL & OrigP & "')"
     
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLineaDePlantilla = True
        numlinea = numlinea + 1
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta." & vbCrLf & Err.Description
End Function



Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de Revisiones: slima1
Dim SQL As String
    
    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    If DatosOkLinea Then
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & " cantidad = " & DBSet(txtAux(3).Text, "N", "N") & ", precioar = " & DBSet(txtAux(4).Text, "N", "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N", "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N", "N") & ", "
        SQL = SQL & " importel = " & DBSet(txtAux(8).Text, "N") & ", origpre=" & DBSet(txtAux(5).Text, "T")
        'SAIL
        SQL = SQL & ", codcapit = " & DBSet(txtAux(9).Text, "N", "S")
        SQL = SQL & " , esopcion = " & Abs(Me.cboOpcion.ListIndex = 1)
        SQL = SQL & " WHERE " & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea
    End If

    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        If txtAux(1).Text = Data2.Recordset!codArtic Then
'            ElArticulo = Data2.Recordset!codArtic
'        Else
'            'Son distintos. Que recalcule todo
'            ElArticulo = ""
'        End If
'        DescuentosCantidad ElArticulo
        ' ----
        
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Oferta" & vbCrLf & Err.Description
End Function



Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, primeravez

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b

    primeravez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(2).Caption = "Alm."
                vDataGrid.Columns(2).Width = 500
                vDataGrid.Columns(2).NumberFormat = "000"
                
                vDataGrid.Columns(3).Caption = "Articulo"
                vDataGrid.Columns(3).Width = 1700
                
                vDataGrid.Columns(4).Caption = "Desc. Artículo"
                vDataGrid.Columns(4).Width = 3300
                
'                vDataGrid.Columns(5).Caption = "Ampl. Línea"
'                vDataGrid.Columns(5).Width = 7980
                vDataGrid.Columns(5).visible = False
                
                vDataGrid.Columns(6).Caption = "Cant."
                vDataGrid.Columns(6).Width = 750
                vDataGrid.Columns(6).Alignment = dbgRight
                vDataGrid.Columns(6).NumberFormat = FormatoImporte
                
                vDataGrid.Columns(7).Caption = "Precio"
                vDataGrid.Columns(7).Width = 1000
                vDataGrid.Columns(7).Alignment = dbgRight
                vDataGrid.Columns(7).NumberFormat = FormatoPrecio
                
                vDataGrid.Columns(8).Caption = "OP"
                vDataGrid.Columns(8).Width = 350
                vDataGrid.Columns(8).Alignment = dbgCenter
                
                
                vDataGrid.Columns(9).Caption = "Dto. 1"
                vDataGrid.Columns(9).Width = 600
                vDataGrid.Columns(9).Alignment = dbgRight
                vDataGrid.Columns(9).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(10).Caption = "Dto. 2"
                vDataGrid.Columns(10).Width = 600
                vDataGrid.Columns(10).Alignment = dbgRight
                vDataGrid.Columns(10).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(11).Caption = "Importe Línea"
                vDataGrid.Columns(11).Width = 1400
                vDataGrid.Columns(11).Alignment = dbgRight
                vDataGrid.Columns(11).NumberFormat = FormatoImporte
                
                vDataGrid.Columns(12).Caption = "Opc."
                vDataGrid.Columns(12).Width = 600
               
                
                vDataGrid.Columns(13).Caption = "Cap."
                vDataGrid.Columns(13).Width = 720
                vDataGrid.Columns(13).NumberFormat = "0000"
                
                vDataGrid.Columns(14).Caption = "Capitulo"
                vDataGrid.Columns(14).Width = 1800
                
    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
'        txtAux2.visible = visible
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
            cboOpcion.ListIndex = -1
        Else 'Vamos a modificar
            For I = 0 To 8
                If I < 3 Then
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                Else
                    txtAux(I).Text = DataGrid1.Columns(I + 3).Text
                End If
                txtAux(I).Locked = False
            Next I
            For I = 9 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I + 4).Text
                txtAux(I).Locked = False
            Next I
            
            If Me.DataGrid1.Columns(12).Text = "Si" Then
                cboOpcion.ListIndex = 1
            Else
                cboOpcion.ListIndex = -1
            End If
        End If
        
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
        'Desc capitulo
        BloquearTxt txtAux(10), True

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        For I = 0 To 2
            cmdAux(I).Top = alto
            'cmdAux(1).Top = alto
            cmdAux(I).Height = DataGrid1.RowHeight
            'cmdAux(1).Height = DataGrid1.RowHeight
        Next
        cboOpcion.Top = alto
        
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
        For I = 4 To 8
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 3).Width - 10
        Next I
        
        txtAux(9).Left = txtAux(8).Left + txtAux(8).Width + 10
        txtAux(9).Width = DataGrid1.Columns(12).Width - 160
        cmdAux(2).Left = txtAux(9).Left + txtAux(9).Width - 50
        txtAux(10).Left = cmdAux(2).Left + cmdAux(2).Width
        txtAux(10).Width = DataGrid1.Columns(13).Width - 10
        cboOpcion.Left = txtAux(8).Left + txtAux(8).Width + 30
        cboOpcion.Width = DataGrid1.Columns(10).Width - 60
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        
    End If
    For I = 0 To 2
        cmdAux(I).visible = visible
    Next
    cboOpcion.visible = visible
End Sub


Private Sub TxtAux_Change(Index As Integer)
    'Precio y Modo Borrar Lineas
    If Index = 4 And ModificaLineas = 2 Then txtAux(5).Text = "M"
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index

   
        ConseguirFocoLin txtAux(Index), cadkey

    
    LabelAyudatxtAux Index, lblF
    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    
    
    '   F2   F2   F2    F2  F2
    If KeyCode = 113 Then
        If Index = 3 Then
            AbrirForm_Articulos txtAux(1).Text
            
        ElseIf Index = 4 Then
            AbrirConsultaPrecio2 Text1(4).Text, txtAux(1).Text, Text1(1).Text, ""
        End If
        
        
    Else
          If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
            If Index < 2 Or Index = 9 Then  'Para los que tienen busqueda
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
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
'Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As CStock
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim cantidad As String
Dim b As Boolean

    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    'txtAux(Index).Text = Trim(txtAux(Index))
    If txtAux(Index).Text = "" And (Index <> 1 And Index <> 4) Then Exit Sub
    'If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)
            
        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then
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
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve) Then
                
                 If devuelve <> txtAux(1).Text Then
                    'ha cambiado el articulo
                    Me.txtAux(3).Text = ""
                    Me.txtAux(4).Text = ""
                    Me.txtAux(5).Text = ""
                    Me.txtAux(6).Text = ""
                    Me.txtAux(7).Text = ""
                End If
            
            
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
                    If txtAux(2).Locked Then
'                        PonerFoco txtAux(3)
                    Else
                        PonerFoco txtAux(2)
                    End If
                Else
                    PonerFoco txtAux(0)
                End If
                
                
                If Text2(16).Text = "" Then _
                    Text2(16).Text = DevuelveDesdeBD(conAri, "txtauxdocumento", "sartic", "codartic", txtAux(1).Text, "T")

                
                
            Else
                PonerFoco txtAux(Index)
            End If
'
'             'Si es articulo de varios podemos modificar la descripción del articulo, sino bloqueamos.
'            If Not EsArticuloVarios(txtAux(Index).Text) Then
'                BloquearTxt txtAux(2), True
'                PonerFoco txtAux(3)
'            Else
'                BloquearTxt txtAux(2), False
'                PonerFoco txtAux(2)
'            End If
        
        Case 2 'desc articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                b = True
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then b = False
                If vCStock.MueveStock Then
                    If Not vCStock.MoverStock(False, False) Then b = False
                End If
                If Not b Then
                    Set vCStock = Nothing
                    Exit Sub
                End If
                
                b = False
                If Modo = 5 Then 'Modo lineas
                    If ModificaLineas = 1 Then 'insertar linea
                        b = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then
                         b = True
                        Else
                            If CStr(DBLet(Data2.Recordset!origpre, "T")) <> "M" Then b = True
                        End If
                        
                    End If
                End If
                
                If b Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If devuelve <> "" Then
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
                    
                    
                 ''       'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                 ''       'precio de caja, y otra linea con el resto unidades un precio unidad
                 ''
                 ''       NumCajas = ObtenerNumCajas(Cantidad, devuelve)
                 ''       RestoUnid = CInt(Cantidad) - NumCajas * CInt(devuelve)
            
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        Set CPrecioFact = New CPreciosFact
                        'CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        CPrecioFact.CodigoClien = Text1(4).Text
                        CPrecioFact.FijarTarifaActividad
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
    '                        TxtAux(3).Text = NumCajas * CInt(devuelve)
                            PonerFoco txtAux(Index)
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And b) Then
                                txtAux(4).Text = Precio
                                txtAux(5).Text = OrigP 'De donde viene el precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                            PonerFoco txtAux(4)
                            ConseguirFocoLin txtAux(4)
    '                            PonerFoco Text2(16)
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If 'modo 5
                Set vCStock = Nothing
            End If 'formato decimal
            
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
            devuelve = ""
            If txtAux(Index).Text <> "" Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Campo numérico", vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                Else
                    devuelve = DevuelveDesdeBD(conAri, "nomcapit", "scapitulos", "codcapit", txtAux(Index).Text, "N")
                    If devuelve = "" Then
                        MsgBox "No existe el capitulo: " & txtAux(9).Text, vbExclamation
                        txtAux(9).Text = ""
                        PonerFoco txtAux(9)
                    End If
                End If
            End If
            txtAux(10).Text = devuelve
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
        
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        LineaIntercalar = 0
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    SQL = " WHERE  numofert=" & Data1.Recordset!NumOfert

    'Lineas ofertas sail
    conn.Execute "Delete from slipresail " & SQL
    
    conn.Execute "Delete from sliprePdfs " & SQL
    
    'Lineas de Ofertas
    conn.Execute "Delete from " & NomTablaLineas & SQL
    
    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar", Err.Description
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP & ")"
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
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next

    ObtenerWhereCP = " numofert= " & Text1(0).Text
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
    
    SQL = "SELECT numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2,importel, "
    
    SQL = SQL & " if( EsOpcion=1,""Si"",""""),"
    SQL = SQL & NomTablaLineas & ".codcapit , nomcapit"
    SQL = SQL & " FROM " & NomTablaLineas & " LEFT JOIN scapitulos ON " & NomTablaLineas & ".codcapit = scapitulos.codcapit WHERE "
    If enlaza Then
        SQL = SQL & ObtenerWhereCP
        If EsHistorico Then SQL = SQL & " and fecofert='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " numofert = -1"
    End If
    SQL = SQL & " Order by numofert, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bol As Boolean
Dim I As Byte

    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For I = 5 To 17
        Toolbar1.Buttons(I).visible = Not EsHistorico
    Next I
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnLineas.visible = Not EsHistorico
    Me.mnGenPedido.visible = Not EsHistorico
    Me.mnPlantillas.visible = Not EsHistorico
    Me.mnOferta.visible = Not EsHistorico 'Traer de Oferta
    Me.mnImpRecordatorio.visible = Not EsHistorico
    Me.mnImpValoracion.visible = Not EsHistorico
    Me.mnbarra2.visible = Not EsHistorico
    Me.mnbarra3.visible = Not EsHistorico
    Me.mnBarra4.visible = Not EsHistorico
    
    Me.Toolbar1.Buttons(19).Enabled = Not EsHistorico
    Me.mnImpFactProF.Enabled = Not EsHistorico
    'Pasar a hco
    Toolbar1.Buttons(15).Enabled = False
    If Not EsHistorico Then
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
            
        If b And Modo <> 5 Then Toolbar1.Buttons(15).Enabled = b
        
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        'Generar Pedido
        If Modo = 5 Then
            Toolbar1.Buttons(11).Image = 34 '.Buttons(11).Image = 26
            Toolbar1.Buttons(11).ToolTipText = "Insertar intercalando"
            b = (ModificaLineas = 0)
        Else
            'b=modo=2
            Toolbar1.Buttons(11).Image = 26   '26
            Toolbar1.Buttons(11).ToolTipText = "Generar Pedido"
        End If
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenPedido.Enabled = b And Modo <> 5
        
        
        
        
        
        b = (Modo = 5) And (ModificaLineas = 0)
        'Plantillas
        Toolbar1.Buttons(12).visible = b
        Toolbar1.Buttons(12).Enabled = b
        Me.mnPlantillas.visible = b
        Me.mnPlantillas.Enabled = b
        'Traer Lineas de Otra Oferta
        Toolbar1.Buttons(13).visible = b
        Toolbar1.Buttons(13).Enabled = b
        Me.mnOferta.visible = b
        Me.mnOferta.Enabled = b
        
        'Recordatorio
        b = (Modo = 2)
        bol = (Modo <> 5)
        Toolbar1.Buttons(16).visible = bol
        Toolbar1.Buttons(16).Enabled = b
        Me.mnImpRecordatorio.visible = bol
        Me.mnImpRecordatorio.Enabled = b
        'Valoración
        Toolbar1.Buttons(17).visible = bol
        Toolbar1.Buttons(17).Enabled = b
        Me.mnImpValoracion.visible = bol
        Me.mnImpValoracion.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
    End If
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
Dim devuelve As String
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Ofertas
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numofert", "numofert", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Ofertas (scapre)."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    EsDeVarios = EsClienteVarios(Text1(4).Text)
    If EsDeVarios Then
'        MenError = "Error al actualizar el Cliente de Varios (sclvar)."
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador de la Oferta."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Oferta." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim I As Byte

    For I = 4 To 13
        Text1(I).Text = ""
    Next I
    If Modo = 3 Then
        For I = 14 To 17
            Text1(I).Text = ""
        Next I
        Text2(12).Text = ""
        Text2(14).Text = ""
        Text2(17).Text = ""
        Me.cboFacturacion.ListIndex = -1
    End If
End Sub
    

Private Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Integer
Dim NumCajas As Integer
Dim cantidad As Integer, UniPorCaja As Integer

    On Error Resume Next

    cantidad = CInt(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function


Private Function DescargarDatosTMP()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    SQL = "DELETE from tmpscapla" & " where codusu= " & vUsu.Codigo
    conn.Execute SQL
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal", Err.Description
End Function



Private Function InsertarPedido(cadSQL As String, MenError As String, numPed As String, QueLineasPasanPedido As String) As Boolean
'Devuelve el mensane de error si se produce
'OUT -> numPed: Nº Pedido que inserta
'Dim cadError As String
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String
Dim vSQL As String

    On Error GoTo EInsertarPedido
    
    bol = False
    InsertarPedido = bol
    
    'Obtener el Contador de PEDIDO
    codtipom = "PEV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            numPed = vTipoMov.ConseguirContador(codtipom)
            devuelve = DevuelveDesdeBDNew(1, "scaped", "numpedcl", "numpedcl", numPed, "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codtipom)
                numPed = vTipoMov.ConseguirContador(codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    

    'Acabar la sql con el contador seleccionado
    vSQL = "INSERT INTO scaped (numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec, nomdirec, referenc,codtraba,codagent, codforpa, dtoppago, dtognral, tipofact,"
    vSQL = vSQL & "observa01, observa02, observa03, observa04, observa05,servcomp,restoped,numofert,fecofert,actuacion,observacrm)"
    vSQL = vSQL & " SELECT " & numPed & " as numpedcl, " & cadSQL
    
    'SAIL
    'Añado actuacion
    vSQL = vSQL & ",actuacion,observacrm "
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numofert=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (scaped )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Pedido (sliped)."
    If Not InsertarLineasPedido(numPed, QueLineasPasanPedido) Then Exit Function
    
    MenError = "Error al actualizar el contador del Pedido."
'    bol = vTipoMov.IncrementarContador("REG")
    bol = vTipoMov.IncrementarContador(codtipom)
    Set vTipoMov = Nothing
    'bol = True
    
EInsertarPedido:
        If Err.Number <> 0 Then bol = False
        InsertarPedido = bol
End Function


Private Sub PasarOfertaAPedido(vSQL As String, QueLineaspasamos As String)
Dim bol As Boolean
Dim MenError As String
Dim numPed As String

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    'Insertar en tablas de Pedido la Oferta
    bol = InsertarPedido(vSQL, MenError, numPed, QueLineaspasamos)
    If bol Then 'Si se inserta Pedido
       'Pasar la Oferta al Historico de Oferta y Borrarla de Ofertas
       vSQL = " scapre.numofert= " & Text1(0).Text
       bol = ActualizarElTraspaso(MenError, vSQL, "OFE", "0")
       If bol Then
            ejecutar "DELETE FROM slipresail WHERE numofert = " & Text1(0).Text, False
            ejecutar "DELETE FROM sliprePdfs WHERE numofert = " & Text1(0).Text, False
       End If
    End If
    
EGenPedido:
    If Err.Number <> 0 Then
        MenError = "Pasando Oferta a Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        Screen.MousePointer = vbDefault
        MsgBox "La Oferta de Venta Nº: " & Text1(0).Text & vbCrLf & vbCrLf & "ha generado el Pedido Nº: " & Format(numPed, "0000000"), vbInformation
    Else
        conn.RollbackTrans
    End If
End Sub


Private Function InsertarLineasPedido(NumPedido As String, QueLineasPasanPedido As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Long
Dim CCost As String


    On Error GoTo Elin

    'David SEPT2009
    'Falta bultos i bultos servidos.  Con lo cual
    'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
    SQL = ""
    SQL = "SELECT " & NumPedido & " as numpedcl, numlinea, codalmac, codartic, nomartic, ampliaci, "
    SQL = SQL & "cantidad, " & "0 as servidas,"
    SQL = SQL & "cantidad,0,"  'bultos i bultos servidos
    SQL = SQL & " precioar, dtoline1, dtoline2, importel, origpre , NULL as numlote," 'Null de numlote
    
    If Not vEmpresa.TieneAnalitica Then
        CCost = "NULL"
    Else
        'De momento, como es para
        'if vparamaplic.ModoAnalitica then
        CCost = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Text1(3).Text)
        If CCost = "" Then
            CCost = "NULL"
        Else
            CCost = DBSet(CCost, "T")
        End If
    End If
    
    SQL = SQL & CCost & " as codccost," 'centro de coste
    
    'SAIL
    If InstalacionEsEulerTaxco Then
        SQL = SQL & "NULL"
    Else
        SQL = SQL & "'T'"
    End If
    SQL = SQL & ",codcapit,cantidad"    'ENERO 2019  solicitadas
    
    SQL = SQL & ", 0 idL"    'ENERO 2019  solicitadas
    SQL = SQL & " FROM " & NomTablaLineas & " WHERE numofert=" & Text1(0).Text

    
    If vParamAplic.ArtSeparador <> "" Then SQL = SQL & " AND codartic <> " & DBSet(vParamAplic.ArtSeparador, "T")
    'Si llevaba opciones
    If QueLineasPasanPedido <> "" Then SQL = SQL & " AND numlinea IN (" & QueLineasPasanPedido & ")"
    
    
    
    SQL = "INSERT INTO sliped " & SQL
    conn.Execute SQL
    

    
    'Ahora actualizo los bultos
    SQL = "Select cantidad , unicajas,numlinea from sliped,sartic where sliped.codartic = sartic.codartic and unicajas >1 and  sliped.numpedcl = " & NumPedido
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        
        I = RS!cantidad \ CLng(RS!unicajas)
        If (RS!cantidad Mod CLng(RS!unicajas)) > 0 Then I = I + 1
        SQL = "UPDATE sliped Set numbultos=" & I & " WHERE sliped.numpedcl = " & NumPedido & " AND numlinea = " & RS!numlinea
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    InsertarLineasPedido = True
    Exit Function
Elin:
    MsgBox Err.Description, vbExclamation
    Set RS = Nothing
     'Hay error , almacenamos y salimos
    InsertarLineasPedido = False

        
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente de la oferta
    vCStock.Documento = Text1(0).Text 'Nº de oferta
    vCStock.FechaMov = Text1(1).Text 'Fecha oferta
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then '1=Insertar Linea
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



Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
    
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
            If vCliente.ClienteBloqueado(True) Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
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
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(vCliente.Codigo, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
            
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(1).Text
            
            
                If vCliente.DeVarios Then
                    PonerFoco Text1(5)
                Else
                    PonerFoco Text1(17)
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
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
'    Text1(6).Text = vCliente.NIF
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
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


Private Sub BotonImprimirProForma(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vTipoM As CTiposMov

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Oferta para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 17 'Facturas Proforma Clientes
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If
    
    'Pasar la letra serie de la factura como parámetro
    Set vTipoM = New CTiposMov
    If vTipoM.Leer("FAV") Then
        
    End If
    Set vTipoM = Nothing
    
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Oferta
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº Oferta
        devuelve = "{" & NombreTabla & ".numofert}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
    End If
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
    
        'Tb lo vamos a tratar
        .outTipoDocumento = 5
        .outClaveNombreArchiv = Text1(0).Text
        .outCodigoCliProv = CLng(Val(Text1(4).Text))
        
    
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = "Factura ProForma"
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub CalcularDatosFactura()
Dim T
Dim cadWhere As String
Dim SQL As String
Dim vFactu As CFactura

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For Each T In Text3
        T.Text = ""
    Next
    
    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP
    
     
    'Agosto 2011. Solo sumaran aquellas que en esopcion NO sea 1
    cadWhere = cadWhere & " AND esopcion <> 1"
    If Me.Data2.Recordset.RecordCount < 1 Then Exit Sub
    
    'SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    'If RegistrosAListar(SQL) = 0 Then Exit Sub
    SQL = DevuelveDesdeBD(conAri, "sum(importel)", NomTablaLineas, "esopcion=1 AND numofert", Text1(0).Text)
    If SQL <> "" Then
        If CCur(SQL) = 0 Then
            SQL = ""
        Else
            SQL = Format(CCur(SQL), FormatoImporte)
        End If
    End If
    Text3(1).Text = SQL
    
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, False) Then
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
Dim I As Byte

    For I = 33 To 36
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    For I = 49 To 54
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    'Desglose B.Imponible por IVA
    For I = 43 To 45
        If Text3(I).Text <> "" Then
             If CSng(Text3(I).Text) = 0 Then
                Text3(I).Text = QuitarCero(Text3(I).Text)
                Text3(I - 3).Text = QuitarCero(Text3(I - 3).Text)
                Text3(I - 6).Text = QuitarCero(Text3(I - 6).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I).Text)
            Else
                Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
                Text3(I - 3) = Format(Text3(I - 3).Text, FormatoDescuento)
                Text3(I + 3).Text = Format(Text3(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
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


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                
                'Insertar modificar los datos del grupo sumergido
                GrupoSumergidoBD
            
                'El Data esta vacio, desde el modo de inicio se pulsa Insertar
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                
                'Estrucutra SIEMPRE
                CrearCarpeta

                
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 1, "Oferta"
                BotonAnyadirLinea False
            End If
        End If
        FormateaCampo Text1(0)
    End If
    Set vTipoMov = Nothing
End Sub






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
        If vParamAplic.TipoPortes <> 1 Then Exit Sub
        
        
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        cad = " WHERE numofert = " & Text1(0).Text
        cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from " & NomTablaLineas & cad
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
                    cad = " WHERE numofert = " & Text1(0).Text
                    cad = "select * from " & NomTablaLineas & cad
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
                            cad = "UPDATE " & NomTablaLineas & " set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            cad = cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            cad = cad & " WHERE numofert = " & Text1(0).Text
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


Private Sub PosicionarData2()
    On Error GoTo EPosicionarData2
    
    Data2.Recordset.Find "numlinea = " & NumRegElim
    If Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub


'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera = True
        If imgBuscar(Index).visible Then imgBuscar_Click Index
        
    Else
        'Lineas
        EsCabecera = False
        If Index = 9 Then
            cmdAux_Click 2
        Else
            cmdAux_Click Index
        End If
        
        
    End If
        
End Sub

Private Sub PonerCampoActuacion()
Dim Cadena As String
            If Modo = 1 Then Exit Sub
            Cadena = ""
            If Text1(33).Text <> "" Then
                If Text1(4).Text = "" Or Text1(12).Text = "" Then
                    MsgBox "Falta cliente/obra", vbExclamation
                    Text1(33).Text = ""
                Else
                    Cadena = "codclien =" & Text1(4).Text & " AND coddirec= " & Text1(12).Text & " AND actuacion "
                
                    Cadena = DevuelveDesdeBDNew(conAri, "sactuaobra", "concat(fechaini,' ',if(observa is null,'',observa))", Cadena, Text1(33).Text, "N")
                    If Cadena = "" Then
                        MsgBox "Ninguna actuacion con ese valor:" & Text1(33).Text, vbInformation
                        Text1(33).Text = ""
                    End If
                End If
                
            End If
            Text2(33).Text = Cadena

End Sub


Private Sub BloquearGrupoSumergido(Bloquear As Boolean)
Dim T As TextBox
    imgBuscar(8).visible = Not Bloquear
    imgBuscar(9).visible = Not Bloquear
    imgBuscar(10).visible = Not Bloquear
    For Each T In Text4
        BloquearTxt T, Bloquear
    Next
End Sub


Private Sub PonerCamposGrupoSumergido(enlaza As Boolean)
Dim I As Integer
Dim limpiar As Boolean
Dim RT As ADODB.Recordset
Dim H As Integer
Dim cad As String
    limpiar = True
    If enlaza Then
            Set RT = New ADODB.Recordset
            RT.Open "select `nomarti1`,`caudal11`,`caudal12`,`caudal13`,`attm11`,`attm12`,`attm13`,`importe1`,`nomarti2`,`caudal21`,`caudal22`,`caudal23`,`attm21`,`attm22`,`attm23`,`importe2`,`nomarti3`,`caudal31`,`caudal32`,`caudal33`,`attm31`,`attm32`,`attm33`,`importe3` from `slipresail` WHERE numofert = " & Data1.Recordset!NumOfert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RT.EOF Then
                limpiar = False
                For I = 0 To Text4.Count - 1
                    H = I Mod 8
                    
                    If H = 0 Then
                        Text4(I).Text = DBLet(RT.Fields(I), "T")
                    Else
                        If IsNull(RT.Fields(I)) Then
                            cad = ""
                        Else
                            cad = RT.Fields(I)
                            If H = 7 Then
                                cad = Format(cad, FormatoCantidad)
                            Else
                                cad = Format(cad, "0")
                            End If
                        End If
                        Text4(I).Text = cad
                    End If
                Next
            End If
            RT.Close
            Set RT = Nothing
    End If
    
    If limpiar Then
        For I = 0 To Text4.Count - 1
            Text4(I).Text = ""
        Next I
    End If

    
End Sub


Private Sub GrupoSumergidoBD()
Dim SQL As String
Dim I As Integer
Dim J As Integer
Dim K As Integer
    If Modo = 3 Then
        'INSERTAR
        SQL = ""
        'insert into `slipresail` (`numofert`,`nomarti1`,`caudal11`,`caudal12`,`caudal13`,`attm11`,`attm12`,`attm13`,`importe1`,`nomarti2`,`caudal21`,`caudal22`,`caudal23`,`attm21`,`attm22`,`attm23`,`importe2`,`nomarti3`,`caudal31`,`caudal32`,`caudal33`,`attm31`,`attm32`,`attm33`,`importe3`) values ( '12',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.00',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.00',NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.00')
        For I = 0 To Text4.Count - 1
           If (I Mod 8) = 0 Then
                'nomartic
                SQL = SQL & "," & DBSet(Text4(I).Text, "T", "S")
            Else
                SQL = SQL & "," & DBSet(Text4(I).Text, "N", "S")
            End If
        Next I
        
        'INSERTAMOS
        SQL = "insert into `slipresail` (`numofert`,`nomarti1`,`caudal11`,`caudal12`,`caudal13`,`attm11`,`attm12`,`attm13`,`importe1`,`nomarti2`,`caudal21`,`caudal22`,`caudal23`,`attm21`,`attm22`,`attm23`,`importe2`,`nomarti3`,`caudal31`,`caudal32`,`caudal33`,`attm31`,`attm32`,`attm33`,`importe3`) VALUES (" & Text1(0).Text & SQL & ")"
        
    Else
        'MODIFICAR
        SQL = ""
        '`nomarti1`,`caudal11`,`caudal12`,`caudal13`,`attm11`,`attm12`,`attm13`,`importe1`,
        '`nomarti2`,`caudal21`,`caudal22`,`caudal23`,`attm21`,`attm22`,`attm23`,`importe2`,
        '`nomarti3`,`caudal31`,`caudal32`,`caudal33`,`attm31`,`attm32`,`attm33`,`importe3
        For I = 1 To 3
            J = (I - 1) * 8
            K = J  'para caudal y attm
            SQL = SQL & ", nomarti" & I & " = " & DBSet(Text4(J).Text, "T", "S")
            'importe
            J = J + 7
            SQL = SQL & ", importe" & I & " = " & DBSet(Text4(J).Text, "N", "S")
            'Caudal
            For J = 1 To 3
                K = K + 1
                SQL = SQL & ", caudal" & I & J & " = " & DBSet(Text4(K).Text, "N", "S")
            Next J
            'Atmosferas
            For J = 1 To 3
                K = K + 1
                SQL = SQL & ", attm" & I & J & " = " & DBSet(Text4(K).Text, "N", "S")
            Next J
        Next I
        SQL = Mid(SQL, 2) 'quito la primera coma
        SQL = "UPDATE slipresail SET " & SQL & " WHERE numofert = " & Text1(0).Text
        
    End If
    ejecutar SQL, False
End Sub


Private Sub cargaDocumentos()
Dim I As Integer
Dim Archvi As String
    Me.ListView1.ListItems.Clear
    
    
    Archvi = "SI"
    
    I = Year(CDate(Text1(2).Text))
    'If i < 2017 Then Archvi = ""
    
    CarpetaDestinoOferta = ComprobarExisteCarpetaPDFOferta(I, CLng(Text1(4).Text), CLng(Text1(0).Text))
    If CarpetaDestinoOferta = "" Then Archvi = ""
    
    If Archvi = "" Then
        ListView1_Click
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    'txtAnterior = "sliprePdfs"
    'If EsHistorico Then txtAnterior = "slhprePdfs"
   '
   ' txtAnterior = "Select * from " & txtAnterior & " WHERE numofert =" & Text1(0).Text & " ORDER BY numlinea"
   ' miRsAux.Open txtAnterior, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   ' i = 0
   ' While Not miRsAux.EOF
   '     i = i + 1
   '     Me.ListView1.ListItems.Add , "C" & miRsAux!numlinea, miRsAux!ficheroDesc
   '     Me.ListView1.ListItems(i).SubItems(1) = miRsAux!ficheronombre
   '     miRsAux.MoveNext
   '
   '
   ' Wend
   ' miRsAux.Close
     I = 0
    Archvi = Dir(CarpetaDestinoOferta, vbDirectory)   ' Recupera la primera entrada.
    Do While Archvi <> ""   ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If Archvi <> "." And Archvi <> ".." Then
           ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
           If (GetAttr(CarpetaDestinoOferta & Archvi) And vbDirectory) <> vbDirectory Then
                'Debug.Print MiNombre   ' Muestra la entrada
                
                 I = I + 1
                 Me.ListView1.ListItems.Add , "C" & I, Archvi
                 
                 Me.ListView1.ListItems(I).SubItems(1) = Archvi
                 
           End If   ' solamente si representa un directorio.
        End If
        Archvi = Dir   ' Obtiene siguiente entrada.
    Loop

    
    
    ListView1_Click
    Set miRsAux = Nothing
    
End Sub




Private Function CargaArchivo(Archivo As String) As Boolean
    
    On Error GoTo eCargaArchivo
    CargaArchivo = False
    
    If Archivo = "" Then
        AcroPDF1.visible = False
    Else
        AcroPDF1.LoadFile (Archivo)
        AcroPDF1.Tag = Archivo
        AcroPDF1.visible = True
    End If
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function





Private Function TieneOpciones2(MostrarMensaje As Boolean) As Boolean
Dim SQL As String
    TieneOpciones2 = False
    
    'Si tiene opciones NO depodremos listar algunas cosas: fra provorma
    If Text1(0).Text <> "" Then
        SQL = DevuelveDesdeBD(conAri, "count(*)", NomTablaLineas, "esopcion=1 AND numofert", Text1(0).Text)
        If Val(SQL) > 0 Then
            If MostrarMensaje Then MsgBox "Oferta con opciones", vbExclamation
            TieneOpciones2 = True
        End If
    End If
End Function

Private Function CrearCarpeta() As Boolean
Dim Aux As String
 On Error Resume Next
 
    CrearCarpeta = False
    Aux = EulerParam & "\Ofertas\" & Year(Data1.Recordset!fecofert)
    If Dir(Aux, vbDirectory) = "" Then
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Aux
            Exit Function
        End If
    End If
    
    Aux = Aux & "\" & Format(Data1.Recordset!codClien, "000000") & " " & Data1.Recordset!NomClien
    If Dir(Aux, vbDirectory) = "" Then
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Aux
            Exit Function
        End If
    End If
    
    Aux = Aux & "\" & Trim(Format(Data1.Recordset!NumOfert, "0000000") & " " & DBLet(Data1.Recordset!referenc, "T"))
    
    If Dir(Aux, vbDirectory) <> "" Then
        'YA existe
    
    Else
        MkDir Aux
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
            Aux = ""
        Else
            'Intentamos crear documentacion interna
             MkDir Aux & "\Documentacion interna"
             If Err.Number <> 0 Then Err.Clear
        End If
        
    End If
    CarpetaDestinoOferta = Aux & "\"
End Function





Private Sub PasarAHco()
Dim cad As String

    'Ciertas comprobaciones
    If EsHistorico Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub



    CadenaDesdeOtroForm = Text1(4).Text & "|" & Text1(5).Text & "|" & Text1(1).Text & "|" & Text1(0).Text & "|"
    frmListadoOfer.OpcionListado = 36
    AbrirListadoOfer (36)



        NumRegElim = Data1.Recordset.AbsolutePosition
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
     
         

End Sub
