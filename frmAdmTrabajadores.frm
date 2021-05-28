VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmTrabajadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabajadores"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14595
   Icon            =   "frmAdmTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   72
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   73
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
      Left            =   3915
      TabIndex        =   70
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   71
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
      Left            =   12780
      TabIndex        =   67
      Top             =   360
      Width           =   1530
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   240
      TabIndex        =   52
      Top             =   810
      Width           =   14145
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
         Left            =   11205
         MaxLength       =   20
         TabIndex        =   61
         Tag             =   "Login Trabajador|T|S|||straba|login||N|"
         Text            =   "Text aldu dkdo sñsñs"
         Top             =   240
         Width           =   2565
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
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código Trabajador|N|N|0|9999|straba|codtraba|0000|S|"
         Text            =   "Text"
         Top             =   240
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
         Index           =   1
         Left            =   3345
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Nombre Trabajador|T|N|||straba|nomtraba||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   6915
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
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
         Left            =   10440
         TabIndex        =   62
         Top             =   270
         Width           =   1245
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
         TabIndex        =   54
         Top             =   270
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
         Left            =   2475
         TabIndex        =   53
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   10170
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
         TabIndex        =   32
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
      Left            =   13350
      TabIndex        =   30
      Top             =   10335
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
      Left            =   12060
      TabIndex        =   28
      Top             =   10335
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   5400
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
      Left            =   4560
      Top             =   5400
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
      Height          =   8490
      Left            =   240
      TabIndex        =   33
      Top             =   1620
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   14975
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmAdmTrabajadores.frx":000C
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
      Tab(0).Control(7)=   "Label1(24)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(26)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ImgMail(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "imgBuscar(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgBuscar(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgBuscar(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgBuscar(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(6)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(9)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "frameBancos"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "frameDptoPersonal"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(8)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(24)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text2(24)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text2(10)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(25)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(25)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text2(28)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(28)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "FrameAux0"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      Begin VB.Frame FrameAux0 
         Caption         =   "Estudios / Formación"
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
         Height          =   4245
         Left            =   90
         TabIndex        =   63
         Top             =   4140
         Width           =   13785
         Begin VB.Frame FrameToolAux0 
            Height          =   645
            Left            =   135
            TabIndex        =   68
            Top             =   270
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   0
               Left            =   135
               TabIndex        =   69
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
            Left            =   675
            MaxLength       =   15
            TabIndex        =   64
            Tag             =   "Periodo|T|N|||strab1|periodos||N|"
            Text            =   "Periodo"
            Top             =   3495
            Visible         =   0   'False
            Width           =   1935
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
            Left            =   2670
            MaxLength       =   70
            TabIndex        =   66
            Tag             =   "Formación|T|N|||strab1|formacio||N|"
            Text            =   "Formacion Formacion Formacion Formacion Formacion Formacion Formacion "
            Top             =   3495
            Visible         =   0   'False
            Width           =   6135
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmAdmTrabajadores.frx":0028
            Height          =   3105
            Left            =   135
            TabIndex        =   65
            Top             =   990
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   5477
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
         Index           =   28
         Left            =   8745
         MaxLength       =   4
         TabIndex        =   27
         Tag             =   "Agente|N|S|||straba|codagent1|||"
         Top             =   3720
         Width           =   690
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
         Left            =   9480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         Text            =   "Text2"
         Top             =   3720
         Width           =   4215
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
         Index           =   25
         Left            =   9480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   3240
         Width           =   4215
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
         Left            =   8745
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "Agente|N|S|||straba|codagent|||"
         Top             =   3240
         Width           =   690
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
         Left            =   1410
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "Centro de coste|T|S|||straba|codccost||N|"
         Top             =   3795
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
         Index           =   10
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   3795
         Width           =   4380
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
         Index           =   24
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   3390
         Width           =   4380
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
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "Almacen por Defecto|N|N|0|999|straba|codalmac|000|N|"
         Text            =   "Text aldu dkdo sñsñs"
         Top             =   3390
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
         Index           =   2
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Domicilio|T|N|||straba|domtraba||N|"
         Text            =   "Text1"
         Top             =   900
         Width           =   5310
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
         Left            =   1425
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Teléfono|T|N|||straba|teltraba||N|"
         Text            =   "Text1"
         Top             =   2190
         Width           =   1830
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
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Cargo en la empresa|T|S|||straba|cartraba||N|"
         Text            =   "Text1"
         Top             =   2580
         Width           =   5310
      End
      Begin VB.Frame frameDptoPersonal 
         Caption         =   "Datos relacionados con Dpto Personal"
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
         Height          =   1275
         Left            =   6840
         TabIndex        =   40
         Top             =   480
         Width           =   7020
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
            Left            =   5505
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fecha de Baja|F|S|||straba|fechabaj|dd/mm/yyyy|N|"
            Top             =   840
            Width           =   1345
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
            Left            =   5505
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha de Alta|F|N|||straba|fechaalt|dd/mm/yyyy|N|"
            Top             =   405
            Width           =   1345
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
            Left            =   2070
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Fecha de Nacimiento|F|N|||straba|fechanac|dd/mm/yyyy|N|"
            Top             =   420
            Width           =   1345
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
            Left            =   2070
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "Nº SSocial|T|S|||straba|nrosegur||N|"
            Text            =   "000000000000"
            Top             =   840
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Baja"
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
            Left            =   3975
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   5220
            Picture         =   "frmAdmTrabajadores.frx":003D
            ToolTipText     =   "Buscar fecha"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alta"
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
            Left            =   3975
            TabIndex        =   46
            Top             =   420
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   5220
            Picture         =   "frmAdmTrabajadores.frx":00C8
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fec.Nacimiento"
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
            TabIndex        =   45
            Top             =   420
            Width           =   1560
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1785
            Picture         =   "frmAdmTrabajadores.frx":0153
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "NºSeguridad Social"
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
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1920
         End
      End
      Begin VB.Frame frameBancos 
         Caption         =   "Datos Bancarios"
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
         Height          =   1275
         Left            =   6840
         TabIndex        =   42
         Top             =   1800
         Width           =   7020
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "IBAN gastos|T|S|||straba|iban1|||"
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
            Index           =   26
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "IBAN|T|S|||straba|iban|||"
            Text            =   "Text1"
            Top             =   360
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
            Left            =   2655
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Código Banco Nómina|N|S|0|9999|straba|codbanco|0000|N|"
            Text            =   "Text1"
            Top             =   360
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
            Index           =   16
            Left            =   3390
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Sucursal Nómina|N|S|0|9999|straba|codsucur|0000|N|"
            Text            =   "Text1"
            Top             =   360
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
            Index           =   17
            Left            =   4140
            MaxLength       =   2
            TabIndex        =   19
            Tag             =   "Dígito Control Nómina|T|S|||straba|digcontr|00||"
            Text            =   "Text1"
            Top             =   360
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
            Index           =   18
            Left            =   4710
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Cuenta Bancaria Nómina|T|S|||straba|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   360
            Width           =   2160
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
            Left            =   2655
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Código Banco Gastos|N|S|0|9999|straba|codbanc1|0000|N|"
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
            Index           =   20
            Left            =   3390
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "Sucursal Gastos|N|S|0|9999|straba|codsucu1|0000|N|"
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
            Index           =   21
            Left            =   4140
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Dígito Control Gastos|T|S|||straba|digcont1|00||"
            Text            =   "Text1"
            Top             =   840
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
            Index           =   22
            Left            =   4710
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Cuenta Bancaria Gastos|T|S|||straba|cuentab1|0000000000||"
            Text            =   "Text1"
            Top             =   840
            Width           =   2160
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Nómina"
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
            Left            =   105
            TabIndex        =   48
            Top             =   360
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Gastos"
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
            Left            =   105
            TabIndex        =   43
            Top             =   840
            Width           =   1725
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
         Index           =   9
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "e-mail|T|S|||straba|maitraba||N|"
         Text            =   "Text1"
         Top             =   2970
         Width           =   5310
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
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "N.I.F.|T|N|||straba|niftraba||N|"
         Text            =   "Text1"
         Top             =   465
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
         Index           =   5
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Provincia|T|N|||straba|protraba||N|"
         Text            =   "Text1"
         Top             =   1770
         Width           =   5310
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
         Left            =   3675
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Población|T|N|||straba|pobtraba||N|"
         Text            =   "Text1"
         Top             =   1335
         Width           =   3045
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
         Left            =   1425
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "C.Postal|T|N|||straba|codpobla||N|"
         Text            =   "Text1"
         Top             =   1335
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1170
         Picture         =   "frmAdmTrabajadores.frx":01DE
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   8460
         Tag             =   "-1"
         ToolTipText     =   "Buscar centro coste"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "x"
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
         Left            =   6975
         TabIndex        =   60
         Top             =   3720
         Width           =   1440
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   8460
         Tag             =   "-1"
         ToolTipText     =   "Buscar centro coste"
         Top             =   3240
         Width           =   240
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
         Height          =   255
         Index           =   4
         Left            =   6975
         TabIndex        =   58
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "C.Coste"
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
         Left            =   240
         TabIndex        =   56
         Top             =   3795
         Width           =   795
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1140
         Tag             =   "-1"
         ToolTipText     =   "Buscar centro coste"
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   1155
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2970
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1140
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacen"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
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
         Left            =   240
         TabIndex        =   50
         Top             =   3390
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo"
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
         Left            =   240
         TabIndex        =   49
         Top             =   2580
         Width           =   750
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
         Index           =   12
         Left            =   240
         TabIndex        =   44
         Top             =   2190
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
         Left            =   240
         TabIndex        =   39
         Top             =   2970
         Width           =   765
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
         Left            =   240
         TabIndex        =   38
         Top             =   465
         Width           =   555
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
         Left            =   240
         TabIndex        =   37
         Top             =   1770
         Width           =   1005
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
         Left            =   2565
         TabIndex        =   36
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1335
         Width           =   960
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
         Left            =   240
         TabIndex        =   34
         Top             =   900
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   5880
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   7320
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data5 
      Height          =   330
      Left            =   3120
      Top             =   5640
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
   Begin MSAdodcLib.Adodc Data6 
      Height          =   330
      Left            =   4560
      Top             =   5640
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
      Left            =   13380
      TabIndex        =   29
      Top             =   10335
      Visible         =   0   'False
      Width           =   1065
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
      Caption         =   "&Mantenimiento Lineas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnEstudios 
         Caption         =   "&Estudios/Formación"
         HelpContextID   =   2
      End
      Begin VB.Menu mnHabilidades 
         Caption         =   "&Habilidades"
         HelpContextID   =   2
      End
      Begin VB.Menu mnExperiencia 
         Caption         =   "Experiencia &Laboral"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormRealizada 
         Caption         =   "&Formación Realizada"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormEmpresa 
         Caption         =   "F&ormacion Empresa"
         HelpContextID   =   2
      End
   End
End
Attribute VB_Name = "frmAdmTrabajadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBasico2 'frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'Form para busquedas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico2
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios  'Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1

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
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'===========================================================================
'       PROCEDIMIENTOS
'============================================================================

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
                
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            cad = "Select * from " & NomTablaLineas & " where codtraba= " & data1.Recordset!CodTraba
            cad = cad & " order by numlinea"
            
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then
                    CargaGrid DataGrid1, Data2, True 'cad
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
'                    PonerBotonCabecera True
                    ModificaLineas = 0
                    'Estudios/Formacion - Datos de la tabla strab1
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    CargaTxtAux1 False, False
                    'CargaGrid DataGrid1, Data2, cad
                    CargaGrid2 DataGrid1, Data2
                    SituarDataPosicion Data2, NumRegElim, Indicador
                    
                    PonerModo 2
                    Me.lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
                    
                    
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
            'Estudios/Formacion
            CargaTxtAux1 False, False
            DataGrid1.Enabled = True
            If ModificaLineas = 1 Then 'Insertar
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
'            PonerBotonCabecera True
            ModificaLineas = 0
            PonerModo 2
            Me.lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("straba", "codtraba")
    Text1(12).Text = Format(Now, "dd/mm/yyyy")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    'Estudios / Formacion
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux1 True, True
    PonerFoco txtAux1(0)
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
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If data1.Recordset.EOF Then
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
    If data1.Recordset.EOF Then Exit Sub
    DesplazamientoData data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    vWhere = "codtraba=" & Val(Text1(0).Text) & " and numlinea="
    'Estudios/Formacion
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = vWhere & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    CargaTxtAux1 True, False
    PonerFoco txtAux1(0)
    DataGrid1.Enabled = False
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"



End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de trabajadores (straba)
Dim cad As String
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If data1.Recordset.EOF Then Exit Sub
    
    
    If Not PuedeEliminarTrabajador Then Exit Sub
    
    
    cad = "Cabecera de Trabajadores." & vbCrLf
    cad = cad & "------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Trabajador:"
    cad = cad & vbCrLf & "Código:   " & Format(data1.Recordset.Fields(0), "000000")
    cad = cad & vbCrLf & "Descripción:   " & data1.Recordset.Fields(1)
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    
    
    
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = data1.Recordset.AbsolutePosition
        
        If Not Eliminar Then
            Exit Sub
        ElseIf SituarDataTrasEliminar(data1, NumRegElim) Then
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

     'EStudios/Formacion
    If Data2.Recordset.EOF Then Exit Sub
    numlinea = Data2.Recordset!numlinea
    
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de Estudios/Formación?"
    SQL = SQL & vbCrLf & "Trabajador: " & Format(data1.Recordset!CodTraba, "0000")
    SQL = SQL & vbCrLf & "Nombre: " & data1.Recordset!NomTraba
    SQL = SQL & vbCrLf & "Numlinea: " & numlinea
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from " & NomTablaLineas & " where codtraba=" & data1.Recordset!CodTraba
        SQL = SQL & " and numlinea=" & numlinea
        conn.Execute SQL

        ModificaLineas = 0
         'Estudios/Formacion
        CargaGrid2 DataGrid1, Data2
    End If
    'PonerFocoBtn Me.cmdRegresar
    PonerModo 2
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Trabajador", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera tambien
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = data1.Recordset.Fields(0) & "|"
        cad = cad & data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub




Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    'Icono de imagen de e-mail
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(20).Picture

    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next



    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
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
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
         
    '## A mano
    NombreTabla = "straba"
    Ordenacion = " ORDER BY codtraba"
    NomTablaLineas = "strab1"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    data1.ConnectionString = conn
    data1.RecordSource = "Select * from " & NombreTabla & " where codtraba=-1"
    data1.Refresh
    
    
    If InstalacionEsEulerTaxco Then
        Label1(6).Caption = "Reloj"
        
        Label1(40).Caption = "Permiso"
        
        
    Else
        Label1(6).Caption = "Agente APP"
    End If
    
    'PonerAltoForm
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
    CargaGrid DataGrid1, Data2, False


End Sub
Private Sub PonerAltoForm()

    If DatosADevolverBusqueda = "" Then
        SSTab1.visible = True
         
        Me.cmdAceptar.Top = 5520
        'Me.cmdCancelar.Top = 5520
        'Me.cmdRegresar.Top = 5520
        Me.Frame1(0).Top = 5360
        Me.Height = 6765
    Else
        SSTab1.visible = False
        Me.Height = 2400
        Me.cmdAceptar.Top = 1180
        'Me.cmdAceptar.Top = 1180
        'Me.cmdRegresar.Top = 1180
        Me.Frame1(0).Top = 1020
    End If
    cmdCancelar.Top = cmdAceptar.Top
    cmdRegresar.Top = cmdAceptar.Top

End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almac
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Almac
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        If Me.imgFecha(0).Tag = 0 Then
            Screen.MousePointer = vbHourglass
            cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else
            'Centro de coste
            If Me.imgFecha(0).Tag = 10 Then
                Text1(10).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(10).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                'Agente
                Text1(CInt(imgFecha(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(CInt(imgFecha(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String

    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        If Me.imgFecha(0).Tag = 0 Then
            Screen.MousePointer = vbHourglass
            cadB = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else
            'Centro de coste
            If Me.imgFecha(0).Tag = 10 Then
                Text1(10).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(10).Text = RecuperaValor(CadenaSeleccion, 2)
            Else
                'Agente
                Text1(CInt(imgFecha(0).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(CInt(imgFecha(0).Tag)).Text = RecuperaValor(CadenaSeleccion, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    TituloLinea = CadenaSeleccion
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String
    
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = Val(imgFecha(0).Tag) + 11
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    TituloLinea = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'CPostal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 4
            VieneDeBuscar = True
            
        Case 1 'Almacen por defecto del trabajador
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            Indice = 24
            
        Case 2 'Centros de coste de la conta
        
            Screen.MousePointer = vbHourglass
            Set frmCC = New frmBasico2
            TituloLinea = ""
            AyudaCentroCoste frmCC
            Set frmCC = Nothing
            If TituloLinea <> "" Then
                Text1(10).Text = Format(RecuperaValor(TituloLinea, 1), "000") 'Cod Almac
                Text2(10).Text = RecuperaValor(TituloLinea, 2) 'Nom Almac
                TituloLinea = ""
            End If
        
        Case 3, 4 'Agentes comerciales
            If Index = 3 Then
                Indice = 25
                Me.imgFecha(0).Tag = 25
            Else
                Indice = 28
               Me.imgFecha(0).Tag = 28
            End If
          
        
            Screen.MousePointer = vbHourglass
'            Set frmB = New frmBuscaGrid
'            frmB.vCampos = "Codigo|sagent|codagent|T||20·Nombre|sagent|nomagent|T||70·"
'            frmB.vTabla = "sagent"
'            frmB.vSQL = ""
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1|"
'            frmB.vTitulo = "Agentes"
'            frmB.vselElem = 0
'            frmB.vConexionGrid = conAri
'
'            frmB.Show vbModal

            Set frmB = New frmBasico2
            AyudaAgentesComerciales frmB, Text1(Indice)
            Set frmB = Nothing
               
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
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
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnEstudios_Click()
'Abre Mantenimiento de lineas  Estudios/Formacion
    BotonMtoLineas 1, "Estudios/Formacion"
    NomTablaLineas = "strab1"
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
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 0 'codigo postal
            Case 24: KEYBusqueda KeyAscii, 1 'almacen
            Case 10: KEYBusqueda KeyAscii, 2 'centro de coste
            Case 25: KEYBusqueda KeyAscii, 3 'agente
            Case 28: KEYBusqueda KeyAscii, 4 'x
            
            Case 11: KEYFecha KeyAscii, 0 ' fecha nacimiento
            Case 12: KEYFecha KeyAscii, 1 ' fecha alta
            Case 13: KEYFecha KeyAscii, 2 ' fecha baja
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
            ValidarNIF Text1(Index).Text
            
            
        Case 10
            ' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
            Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
'            PonceCentroCoste
            
        Case 11, 12, 13 'Fecha Nacimiento, Fecha alta, Fecha baja
            'Si no es modo de Busqueda poner el formato
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 24 'Cod almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index).Text = ""
            End If
        Case 25, 28
            'Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent", "codagent")
                If Text2(Index).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16, 19, 20 'cod. banco, cod. sucursal
            PonerFormatoEntero Text1(Index)
            
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

    cadB = ObtenerBusqueda(Me, False)
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
'Carga el formulario frmBuscaGrid con los valores correspondientes
    
    TituloLinea = ""
    Set frmT = New frmBasico2
    AyudaTrabajadores frmT, Text1(0), cadB
    Set frmB = Nothing
    If TituloLinea <> "" Then
        Screen.MousePointer = vbHourglass
        cadB = ValorDevueltoFormGrid(Text1(0), TituloLinea, 1)
        TituloLinea = "2"
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If

End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    data1.RecordSource = CadenaConsulta
    data1.Refresh
    If data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        data1.Recordset.MoveFirst
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
Dim SQL As String
Dim vWhere As String
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
   
    vWhere = " WHERE codtraba= " & data1.Recordset!CodTraba
    'Estudios/Formacion - Datos de la tabla strab1
    SQL = "Select * from strab1 " & vWhere
    SQL = SQL & " order by numlinea"
    CargaGrid DataGrid1, Data2, True  'SQL
    
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
    
    If data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, data1
    
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "salmpr", "nomalmac")
    
    ' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
'    PonceCentroCoste
    Me.Text2(10).Text = PonerNombreCCoste(Me.Text1(10))
    Text2(25).Text = PonerNombreDeCod(Text1(25), conAri, "sagent", "nomagent")
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "sagent", "nomagent", "codagent")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    'Visualizar el login solo si es administrador o root
    b = (vUsu.Nivel < 2)
    Me.Label1(25).visible = b
    Text1(23).visible = b

    'Actualiza Iconos Insertar,Modificar,Eliminar
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
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not data1.Recordset.EOF Then
        If data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    b = Modo = 1 Or Modo = 3 Or Modo = 4 Or Modo = 5 'Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
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
Dim b As Boolean
Dim bAux As Boolean
Dim i As Byte
Dim EsBusqueda As Boolean

On Error Resume Next

    EsBusqueda = Me.DatosADevolverBusqueda <> ""
    b = (Modo = 2 Or Modo = 0) And Not EsBusqueda
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2) And Not EsBusqueda
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '------------------------------------------
    b = Not (Modo = 0 Or Modo = 2) '(Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    'imprimir
    Toolbar1.Buttons(8).Enabled = True
    
    b = (Modo = 2) And DatosADevolverBusqueda = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
  PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
On Error GoTo EDatosOK

    DatosOk = False
    b = True
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
          
    '[Monica]15/02/2021: comprobamos que el login no lo tenga otro trabajador
    If Modo = 3 Or Modo = 4 Then
        SQL = "select count(*) from straba where codtraba <> " & DBSet(Text1(0), "N") & " and login = " & DBSet(Text1(23), "T")
        If TotalRegistros(SQL) <> 0 Then
            MsgBox "El login introducido está asignado a otro trabajador. Revise.", vbExclamation
            PonerFoco Text1(23)
            Exit Function
        End If
    End If
          
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Estudios/Formacion
    If Trim(txtAux1(0).Text) = "" Or Trim(txtAux1(1).Text) = "" Then
        MsgBox "Los campos Periodo y Formación no pueden ser nulos", vbExclamation
        b = False
    End If
    
    DatosOkLinea = b
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
        Case Else
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8
            frmListado2.Opcion = 17
            frmListado2.Show vbModal
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
Dim SQL As String
Dim vWhere As String
Dim NumF As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea Then
         vWhere = "codtraba=" & Val(Text1(0).Text)
         NumF = SugerirCodigoSiguienteStr("strab1", "numlinea", vWhere)
         'Estudios/Formacion
         SQL = "INSERT INTO strab1 "
         SQL = SQL & "(codtraba, numlinea, periodos, formacio) "
         SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & NumF & ","
         SQL = SQL & DBSet(txtAux1(0).Text, "T") & "," & DBSet(txtAux1(1).Text, "T") & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
Dim SQL As String
Dim vWhere As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    If DatosOkLinea Then
         vWhere = "codtraba=" & Val(Text1(0).Text)
         'Estudios/Formacion
        SQL = "UPDATE strab1 Set periodos = " & DBSet(txtAux1(0).Text, "T")
        SQL = SQL & ", formacio = " & DBSet(txtAux1(1).Text, "T")
        SQL = SQL & " WHERE " & vWhere & " AND numlinea=" & Data2.Recordset!numlinea
    End If

    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Trabajador" & vbCrLf & Err.Description
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
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, True

    DataGrid1.RowHeight = 350
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b

    'PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


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
    
    SQL = "Select * from strab1 where codtraba= "
    If enlaza Then
        SQL = SQL & data1.Recordset!CodTraba
    Else
        SQL = SQL & "  -1"
    End If
    SQL = SQL & " Order by 1"
    MontaSQLCarga = SQL
End Function


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False 'codtraba
    vDataGrid.Columns(1).visible = False 'numlinea

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Estudios / Formacion
                vDataGrid.Columns(2).Caption = "Período"
                vDataGrid.Columns(2).Width = 4100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(3).Caption = "Formación"
                vDataGrid.Columns(3).Width = 8780
    End Select

    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i

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
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
                BloquearTxt txtAux1(i), False
            Next i
        Else
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = DataGrid1.Columns(i + 2).Text
                BloquearTxt txtAux1(i), False
            Next i
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 6)
        
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        txtAux1(0).Left = DataGrid1.Left + 320
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 20
        'Formacion
        txtAux1(1).Left = txtAux1(0).Left + txtAux1(0).Width + 20
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 20
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
    End If
End Sub



Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), 3
End Sub

'Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
'      If Not (Index = 0 And KeyCode = 38) Then
'            KEYdown KeyCode
'      End If
'End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        NumTabMto = numTab
        TituloLinea = cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Sub TxtAux1_LostFocus(Index As Integer)
    If Index = 1 Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub



Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar

        conn.BeginTrans
        SQL = " WHERE  codtraba=" & data1.Recordset!CodTraba

        'Lineas Estudios/Formacion
        conn.Execute "Delete from strab1 " & SQL
        'Cabeceras
        conn.Execute "Delete from straba " & SQL

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
'Pone los Grids sin datos, apuntando a ningún registro
Dim cad As String
On Error Resume Next

    'cad = "Select * from strab1 where codtraba= -1"
    CargaGrid DataGrid1, Data2, False ' cad
    
    PrimeraVez = False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codtraba=" & Text1(0).Text & ")"
    If SituarData(data1, cad, Indicador) Then
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
Dim Aux As String
    PuedeEliminarTrabajador = False
    
    
    
    If HayRegParaInforme("sctrcompr", "codtraba = " & Text1(0).Text, True) Then Aux = Aux & vbCrLf & " -Control albaranes proveedor(1)"
    If HayRegParaInforme("sctrcompr", "codtrab1 = " & Text1(0).Text, True) Then Aux = Aux & vbCrLf & " -Control albaranes proveedor(2)"
    If InstalacionEsEulerTaxco Then
        If HayRegParaInforme("sreloj", "codtraba = " & Text1(0).Text, True) Then Aux = Aux & vbCrLf & " -Tareas productividad"
    End If
    
    If Aux <> "" Then
        Aux = "Existen registros relacionados" & vbCrLf & Aux
        MsgBox Aux, vbExclamation
    Else
        PuedeEliminarTrabajador = True
    End If
    
    
End Function
