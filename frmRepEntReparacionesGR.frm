VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepEntReparacionesGR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   16410
   ClipControls    =   0   'False
   Icon            =   "frmRepEntReparacionesGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   16410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   214
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   215
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3825
      TabIndex        =   212
      Top             =   45
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   213
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Confirmar Reparaci�n"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4950
      TabIndex        =   210
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   211
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
      Height          =   315
      Left            =   14355
      TabIndex        =   194
      Top             =   270
      Width           =   1620
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
      Left            =   9000
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   104
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   10890
      Visible         =   0   'False
      Width           =   3735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9000
      Left            =   90
      TabIndex        =   49
      Tag             =   "A|N|S|||scarep|contestado||S|"
      ToolTipText     =   "Descliente"
      Top             =   1575
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   15875
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Datos basicos "
      TabPicture(0)   =   "frmRepEntReparacionesGR.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameOtros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameClientes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameAux"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Presupuesto / S.A.T."
      TabPicture(1)   =   "frmRepEntReparacionesGR.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(3)"
      Tab(1).Control(1)=   "Label2(2)"
      Tab(1).Control(2)=   "imgFecha(4)"
      Tab(1).Control(3)=   "Label9(3)"
      Tab(1).Control(4)=   "Label1(13)"
      Tab(1).Control(5)=   "imgFecha(3)"
      Tab(1).Control(6)=   "Label9(2)"
      Tab(1).Control(7)=   "imgFecha(2)"
      Tab(1).Control(8)=   "Label9(1)"
      Tab(1).Control(9)=   "Label2(1)"
      Tab(1).Control(10)=   "Label11(0)"
      Tab(1).Control(11)=   "Label1(12)"
      Tab(1).Control(12)=   "Line1"
      Tab(1).Control(13)=   "Line2"
      Tab(1).Control(14)=   "Label12(0)"
      Tab(1).Control(15)=   "Label12(1)"
      Tab(1).Control(16)=   "Label11(1)"
      Tab(1).Control(17)=   "imgFecha(5)"
      Tab(1).Control(18)=   "Label9(4)"
      Tab(1).Control(19)=   "Label9(5)"
      Tab(1).Control(20)=   "imgBuscar(8)"
      Tab(1).Control(21)=   "Text1(22)"
      Tab(1).Control(22)=   "Text2(21)"
      Tab(1).Control(23)=   "Text1(21)"
      Tab(1).Control(24)=   "Text1(20)"
      Tab(1).Control(25)=   "Text1(19)"
      Tab(1).Control(26)=   "Text1(18)"
      Tab(1).Control(27)=   "Text1(17)"
      Tab(1).Control(28)=   "Text1(16)"
      Tab(1).Control(29)=   "Combo1"
      Tab(1).Control(30)=   "Check1"
      Tab(1).Control(31)=   "Text1(25)"
      Tab(1).Control(32)=   "Text1(26)"
      Tab(1).Control(33)=   "Text1(27)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "L�neas"
      TabPicture(2)   =   "frmRepEntReparacionesGR.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(16)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ficha reparaci�n"
      TabPicture(3)   =   "frmRepEntReparacionesGR.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line3"
      Tab(3).Control(1)=   "Label3(1)"
      Tab(3).Control(2)=   "Label3(2)"
      Tab(3).Control(3)=   "Label3(3)"
      Tab(3).Control(4)=   "Label3(4)"
      Tab(3).Control(5)=   "Label3(5)"
      Tab(3).Control(6)=   "Label3(6)"
      Tab(3).Control(7)=   "Label3(7)"
      Tab(3).Control(8)=   "Label3(8)"
      Tab(3).Control(9)=   "Label3(9)"
      Tab(3).Control(10)=   "Label3(10)"
      Tab(3).Control(11)=   "Label3(12)"
      Tab(3).Control(12)=   "Label3(13)"
      Tab(3).Control(13)=   "Label3(14)"
      Tab(3).Control(14)=   "Label3(15)"
      Tab(3).Control(15)=   "Label3(16)"
      Tab(3).Control(16)=   "Label3(17)"
      Tab(3).Control(17)=   "Label3(18)"
      Tab(3).Control(18)=   "Label3(19)"
      Tab(3).Control(19)=   "Line4"
      Tab(3).Control(20)=   "Label3(20)"
      Tab(3).Control(21)=   "Label3(22)"
      Tab(3).Control(22)=   "Label3(23)"
      Tab(3).Control(23)=   "Label3(24)"
      Tab(3).Control(24)=   "Line5"
      Tab(3).Control(25)=   "Label3(11)"
      Tab(3).Control(26)=   "Label3(25)"
      Tab(3).Control(27)=   "Label3(26)"
      Tab(3).Control(28)=   "Label3(27)"
      Tab(3).Control(29)=   "Label3(28)"
      Tab(3).Control(30)=   "Label3(29)"
      Tab(3).Control(31)=   "Label3(30)"
      Tab(3).Control(32)=   "Label3(31)"
      Tab(3).Control(33)=   "Label3(32)"
      Tab(3).Control(34)=   "Label3(33)"
      Tab(3).Control(35)=   "Label3(34)"
      Tab(3).Control(36)=   "Label3(35)"
      Tab(3).Control(37)=   "Label3(36)"
      Tab(3).Control(38)=   "Label3(37)"
      Tab(3).Control(39)=   "txtEuler(7)"
      Tab(3).Control(40)=   "chkEuler(0)"
      Tab(3).Control(41)=   "chkEuler(1)"
      Tab(3).Control(42)=   "chkEuler(2)"
      Tab(3).Control(43)=   "chkEuler(3)"
      Tab(3).Control(44)=   "chkEuler(4)"
      Tab(3).Control(45)=   "chkEuler(5)"
      Tab(3).Control(46)=   "chkEuler(6)"
      Tab(3).Control(47)=   "chkEuler(7)"
      Tab(3).Control(48)=   "chkEuler(8)"
      Tab(3).Control(49)=   "chkEuler(9)"
      Tab(3).Control(50)=   "txtEuler(5)"
      Tab(3).Control(51)=   "txtEuler(6)"
      Tab(3).Control(52)=   "txtEuler(3)"
      Tab(3).Control(53)=   "txtEuler(4)"
      Tab(3).Control(54)=   "txtEuler(8)"
      Tab(3).Control(55)=   "txtEuler(10)"
      Tab(3).Control(56)=   "txtEuler(9)"
      Tab(3).Control(57)=   "optEuler(0)"
      Tab(3).Control(58)=   "optEuler(1)"
      Tab(3).Control(59)=   "txtEuler(0)"
      Tab(3).Control(60)=   "txtEuler(1)"
      Tab(3).Control(61)=   "Frame4"
      Tab(3).Control(62)=   "txtEuler(2)"
      Tab(3).Control(63)=   "txtEuler(12)"
      Tab(3).Control(64)=   "txtEuler(13)"
      Tab(3).Control(65)=   "txtEuler(14)"
      Tab(3).Control(66)=   "txtEuler(16)"
      Tab(3).Control(67)=   "txtEuler(15)"
      Tab(3).Control(68)=   "Frame5"
      Tab(3).Control(69)=   "txtEuler(11)"
      Tab(3).Control(70)=   "cboEulerUd"
      Tab(3).Control(71)=   "txtEuler(17)"
      Tab(3).Control(72)=   "txtEuler(18)"
      Tab(3).Control(73)=   "txtEuler(19)"
      Tab(3).Control(74)=   "txtEuler(20)"
      Tab(3).Control(75)=   "txtEuler(21)"
      Tab(3).ControlCount=   76
      Begin VB.Frame FrameAux 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   90
         TabIndex        =   195
         Top             =   6075
         Width           =   15810
         Begin VB.Frame FrameToolAux0 
            Height          =   645
            Left            =   135
            TabIndex        =   216
            Top             =   45
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   0
               Left            =   150
               TabIndex        =   217
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
            Height          =   315
            Index           =   0
            Left            =   180
            MaxLength       =   15
            TabIndex        =   200
            Tag             =   "C�digo Almacen"
            Text            =   "codalmac"
            Top             =   2190
            Visible         =   0   'False
            Width           =   615
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
            Height          =   315
            Index           =   1
            Left            =   1020
            MaxLength       =   18
            TabIndex        =   201
            Tag             =   "C�digo Art�culo"
            Text            =   "Artic Artic Artic5"
            Top             =   2190
            Visible         =   0   'False
            Width           =   1455
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
            Height          =   315
            Index           =   3
            Left            =   6060
            MaxLength       =   16
            TabIndex        =   202
            Tag             =   "Cantidad"
            Text            =   "1,234,567,891.25"
            Top             =   2370
            Visible         =   0   'False
            Width           =   1215
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
            Height          =   315
            Index           =   4
            Left            =   7260
            MaxLength       =   12
            TabIndex        =   203
            Tag             =   "Precio"
            Text            =   "123,456.7879"
            Top             =   2370
            Visible         =   0   'False
            Width           =   855
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
            Height          =   315
            Index           =   5
            Left            =   8700
            MaxLength       =   5
            TabIndex        =   204
            Tag             =   "Descuento 1"
            Text            =   "Dto1"
            Top             =   2370
            Visible         =   0   'False
            Width           =   615
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
            Height          =   315
            Index           =   6
            Left            =   9300
            MaxLength       =   30
            TabIndex        =   205
            Tag             =   "Descuento 2"
            Text            =   "Dto2"
            Top             =   2370
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
            Height          =   315
            Index           =   7
            Left            =   9900
            MaxLength       =   12
            TabIndex        =   206
            Tag             =   "Importe"
            Text            =   "Importe"
            Top             =   2370
            Visible         =   0   'False
            Width           =   1095
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
            Height          =   315
            Index           =   2
            Left            =   2700
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   199
            Tag             =   "Nombre Art�culo"
            Text            =   "nomArtic"
            Top             =   2250
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Left            =   780
            TabIndex        =   198
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Left            =   2460
            TabIndex        =   197
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
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
            Height          =   315
            Index           =   8
            Left            =   11070
            MaxLength       =   12
            TabIndex        =   207
            Tag             =   "CC"
            Text            =   "CC"
            Top             =   2340
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Index           =   2
            Left            =   10860
            TabIndex        =   196
            Top             =   2310
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1845
            Left            =   135
            TabIndex        =   208
            Top             =   810
            Visible         =   0   'False
            Width           =   15585
            _ExtentX        =   27490
            _ExtentY        =   3254
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
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
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -72840
         MaxLength       =   16
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   1080
         Width           =   1260
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -74760
         MaxLength       =   16
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   1080
         Width           =   1380
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -63120
         MaxLength       =   16
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   6000
         Width           =   900
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -65040
         MaxLength       =   16
         TabIndex        =   153
         Text            =   "Text1"
         Top             =   6000
         Width           =   900
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66960
         MaxLength       =   16
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   6000
         Width           =   900
      End
      Begin VB.ComboBox cboEulerUd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRepEntReparacionesGR.frx":007C
         Left            =   -69840
         List            =   "frmRepEntReparacionesGR.frx":0089
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -70800
         MaxLength       =   16
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   6120
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
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
         Left            =   -74040
         TabIndex        =   187
         Top             =   6120
         Width           =   3015
         Begin VB.OptionButton optEuler 
            Caption         =   "V"
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
            Index           =   7
            Left            =   1320
            TabIndex        =   143
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "Otro"
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
            Left            =   1920
            TabIndex        =   144
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "N"
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
            Index           =   5
            Left            =   120
            TabIndex        =   141
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "C"
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
            Left            =   720
            TabIndex        =   142
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66960
         MaxLength       =   16
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   5520
         Width           =   1620
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -63960
         MaxLength       =   16
         TabIndex        =   151
         Text            =   "Text1"
         Top             =   5520
         Width           =   1740
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66960
         MaxLength       =   16
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   5040
         Width           =   4740
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66960
         MaxLength       =   16
         TabIndex        =   148
         Text            =   "Text1"
         Top             =   4560
         Width           =   4740
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66960
         MaxLength       =   16
         TabIndex        =   147
         Text            =   "Text1"
         Top             =   4080
         Width           =   2220
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -63360
         MaxLength       =   16
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   375
         Left            =   -74760
         TabIndex        =   175
         Top             =   1560
         Width           =   3015
         Begin VB.OptionButton optEuler 
            Caption         =   "Pagados"
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
            Left            =   1905
            TabIndex        =   117
            Top             =   0
            Width           =   1200
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "Debidos"
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
            Left            =   645
            TabIndex        =   116
            Top             =   0
            Value           =   -1  'True
            Width           =   1200
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
            TabIndex        =   176
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -65760
         MaxLength       =   16
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   1440
         Width           =   2100
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -70200
         MaxLength       =   16
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   1440
         Width           =   4260
      End
      Begin VB.OptionButton optEuler 
         Caption         =   "Cliente"
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
         Left            =   -69000
         TabIndex        =   119
         Top             =   840
         Width           =   1155
      End
      Begin VB.OptionButton optEuler 
         Caption         =   "Agencia"
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
         Left            =   -70320
         TabIndex        =   118
         Top             =   840
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -73920
         MaxLength       =   16
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   5520
         Width           =   1620
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -70920
         MaxLength       =   16
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   5520
         Width           =   1860
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -73920
         MaxLength       =   16
         TabIndex        =   138
         Text            =   "Text1"
         Top             =   5040
         Width           =   4860
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66240
         MaxLength       =   16
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   3000
         Width           =   4110
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -66240
         MaxLength       =   16
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   2640
         Width           =   4110
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -70800
         MaxLength       =   16
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   4080
         Width           =   1740
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -73920
         MaxLength       =   16
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   4080
         Width           =   2220
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Index           =   9
         Left            =   -67080
         TabIndex        =   133
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Index           =   8
         Left            =   -68640
         TabIndex        =   132
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Index           =   7
         Left            =   -69720
         TabIndex        =   131
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -71040
         TabIndex        =   130
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -72120
         TabIndex        =   129
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -67080
         TabIndex        =   127
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -68640
         TabIndex        =   126
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -69720
         TabIndex        =   125
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -71040
         TabIndex        =   124
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkEuler 
         Caption         =   "chkEuler"
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
         Left            =   -72120
         TabIndex        =   123
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox txtEuler 
         BeginProperty Font 
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
         Left            =   -73920
         MaxLength       =   16
         TabIndex        =   137
         Text            =   "Text1"
         Top             =   4560
         Width           =   4860
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
         Height          =   1155
         Index           =   27
         Left            =   -72060
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Tag             =   "F|T|S|||scarep|observasat|||"
         Text            =   "frmRepEntReparacionesGR.frx":009D
         Top             =   5160
         Width           =   10035
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
         Left            =   -72060
         MaxLength       =   10
         TabIndex        =   36
         Tag             =   "Fecha Entrega SAT|F|S|||scarep|fecentresat|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   4680
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
         Index           =   25
         Left            =   -72060
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "Imp reparacion SAT|N|S|||scarep|importesat|0.00||"
         Text            =   "Text1"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente avisado"
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
         Left            =   -66765
         TabIndex        =   31
         Tag             =   "A|N|S|||scarep|avisocli||S|"
         Top             =   1650
         Width           =   1845
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
         ItemData        =   "frmRepEntReparacionesGR.frx":00A3
         Left            =   -65700
         List            =   "frmRepEntReparacionesGR.frx":00B0
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "Aceptado|N|S|||scarep|contestado||N|"
         Top             =   1080
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
         Index           =   16
         Left            =   -72240
         MaxLength       =   7
         TabIndex        =   26
         Tag             =   "Imp pres1|N|S|||scarep|imppresu1|0.00||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1365
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
         Left            =   -72240
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Imp presupuesto 2|N|S|||scarep|impresu2|0.00||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1365
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
         Index           =   18
         Left            =   -68505
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Fecha presupuesto|F|S|||scarep|fecha|dd/mm/yyyy|N|"
         Text            =   "Text1"
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
         Index           =   19
         Left            =   -68505
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Fecha aprobacion|F|S|||scarep|fechaaprob|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1560
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
         Index           =   20
         Left            =   -72060
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Fecha envio|F|S|||scarep|fecenviosat|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   3120
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
         Index           =   21
         Left            =   -72060
         MaxLength       =   4
         TabIndex        =   32
         Tag             =   "Servicio SAT|N|S|0||scarep|codman|0000||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   1125
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
         Left            =   -70905
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   88
         Text            =   "Text2"
         Top             =   2640
         Width           =   3930
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
         Left            =   -72060
         MaxLength       =   15
         TabIndex        =   34
         Tag             =   "N� Reparaci�n|T|S|||scarep|resguardosat|||"
         Text            =   "Text1"
         Top             =   3600
         Width           =   3765
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3360
         Left            =   90
         TabIndex        =   66
         Top             =   2745
         Width           =   15810
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
            Index           =   37
            Left            =   10605
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   110
            Text            =   "Text2"
            Top             =   2970
            Width           =   5145
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
            Left            =   9900
            MaxLength       =   4
            TabIndex        =   109
            Tag             =   "Tecnico|N|N|0|9999|scarep|codtrab1|0000|N|"
            Text            =   "Te"
            Top             =   2970
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
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
            Left            =   11295
            MaxLength       =   80
            TabIndex        =   107
            Tag             =   "F.Aviso|F|S|||scarep|fecaviso|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1350
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
            Left            =   10530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   85
            Text            =   "Text2"
            Top             =   2265
            Width           =   5220
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
            Left            =   9900
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Trabajo realizado|N|S|||scarep|codtrabajo|00|N|"
            Text            =   "Te"
            Top             =   2265
            Width           =   570
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
            Index           =   23
            Left            =   10530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   83
            Text            =   "Text2"
            Top             =   945
            Width           =   5220
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
            Left            =   9900
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Tipo averia|N|S|||scarep|codavi|00|N|"
            Text            =   "Te"
            Top             =   945
            Width           =   570
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
            Index           =   0
            Left            =   13680
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   71
            Tag             =   "Tipo Albaran|T|S|||schrep|codtipom||N|"
            Text            =   "Text2"
            Top             =   270
            Width           =   660
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
            Index           =   14
            Left            =   14400
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   70
            Tag             =   "Fecha Alb|F|S|||schrep|fechaalb|dd/mm/yyyy|N|"
            Text            =   "Text2"
            Top             =   270
            Width           =   1350
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
            Index           =   15
            Left            =   12690
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   69
            Tag             =   "N� Albaran|T|S|||schrep|numalbar|0000000|N|"
            Text            =   "Text2"
            Top             =   270
            Width           =   930
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   23
            Tag             =   "Texto Reparaci�n 1|T|S|||scarep|textore1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2250
            Width           =   9780
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   24
            Tag             =   "Texto Reparaci�n 2|T|S|||scarep|textore2||N|"
            Text            =   "Text1"
            Top             =   2610
            Width           =   9780
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   25
            Tag             =   "Texto Reparaci�n 3|T|S|||scarep|textore3||N|"
            Text            =   "Text1"
            Top             =   2970
            Width           =   9780
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   17
            Tag             =   "Material con el que entra|T|S|||scarep|material||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   270
            Width           =   9765
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   18
            Tag             =   "Aver�a detectada|T|S|||scarep|tipoaver||N|"
            Text            =   "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
            Top             =   945
            Width           =   9765
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
            Left            =   75
            MaxLength       =   80
            TabIndex        =   20
            Tag             =   "Situaci�n de la Reparaci�n|T|S|||scarep|motivore||N|"
            Text            =   "Text1"
            Top             =   1575
            Width           =   9735
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
            Left            =   9900
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "Motivo Pendiente Rep.|N|S|||scarep|codmotre|00|N|"
            Text            =   "Te"
            Top             =   1575
            Width           =   570
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
            Left            =   10485
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   68
            Text            =   "Text2"
            Top             =   1575
            Width           =   5265
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
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
            Left            =   9900
            MaxLength       =   80
            TabIndex        =   67
            Tag             =   "N� aviso|N|S|||scarep|numaviso||N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1380
         End
         Begin VB.Image imgVerAlbaran 
            Height          =   240
            Left            =   15525
            Picture         =   "frmRepEntReparacionesGR.frx":00CE
            Tag             =   "-1"
            ToolTipText     =   "Ver albar�n"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   11925
            ToolTipText     =   "Buscar trabajador"
            Top             =   2655
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "T�cnico reparaci�n"
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
            Left            =   9945
            TabIndex        =   111
            Top             =   2655
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha aviso"
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
            Left            =   11295
            TabIndex        =   108
            Top             =   0
            Width           =   1245
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   11700
            ToolTipText     =   "Buscar tipo trabajo"
            Top             =   1980
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajo realizado"
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
            Left            =   9900
            TabIndex        =   87
            Top             =   1980
            Width           =   1755
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   11070
            ToolTipText     =   "Buscar tipo aver�a"
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo averia"
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
            Left            =   9900
            TabIndex        =   84
            Top             =   660
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Aviso n�"
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
            Index           =   11
            Left            =   9900
            TabIndex        =   82
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Alb."
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
            Index           =   8
            Left            =   13725
            TabIndex        =   80
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alb."
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
            Left            =   14400
            TabIndex        =   79
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "N�Albaran"
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
            Index           =   10
            Left            =   12690
            TabIndex        =   78
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Aver�a detectada"
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
            Left            =   75
            TabIndex        =   76
            Top             =   675
            Width           =   2565
         End
         Begin VB.Label Label8 
            Caption         =   "Situaci�n Reparaci�n"
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
            Left            =   75
            TabIndex        =   75
            Top             =   1305
            Width           =   2700
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo Pendiente Reparaci�n"
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
            Left            =   9900
            TabIndex        =   74
            Top             =   1305
            Width           =   2865
         End
         Begin VB.Label Label1 
            Caption         =   "Texto Reparaci�n"
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
            Left            =   75
            TabIndex        =   73
            Top             =   1980
            Width           =   2415
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   12825
            ToolTipText     =   "Buscar motivo"
            Top             =   1305
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Material con que entra"
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
            Left            =   75
            TabIndex        =   72
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame FrameClientes 
         Caption         =   "Datos Clientes"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   7755
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
            Left            =   2295
            MaxLength       =   40
            TabIndex        =   4
            Tag             =   "NomCliente|T|N|||scarep|nomclien|||"
            Text            =   "Text1"
            Top             =   360
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
            Index           =   33
            Left            =   5190
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Tfno|T|S|||scarep|proclien|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   2220
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
            Left            =   2145
            MaxLength       =   30
            TabIndex        =   9
            Tag             =   "Tfno|T|S|||scarep|pobclien|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   3015
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
            Index           =   31
            Left            =   1365
            MaxLength       =   6
            TabIndex        =   8
            Tag             =   "CP|T|S|||scarep|codpobla|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   750
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
            Left            =   1365
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "D|T|S|||scarep|domclien|||"
            Text            =   "Text1"
            Top             =   1080
            Width           =   6030
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
            Index           =   29
            Left            =   5220
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "Tfno|T|S|||scarep|telclien|||"
            Text            =   "Text1"
            Top             =   720
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
            Index           =   28
            Left            =   1365
            MaxLength       =   16
            TabIndex        =   5
            Tag             =   "NIF|T|N|||scarep|nifdatos|||"
            Text            =   "Text1"
            Top             =   720
            Width           =   2535
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
            Index           =   6
            Left            =   1365
            MaxLength       =   6
            TabIndex        =   3
            Tag             =   "Cod. Cliente|N|N|0|999999|scarep|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   870
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
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   1845
            Width           =   5415
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
            Index           =   7
            Left            =   1365
            MaxLength       =   3
            TabIndex        =   11
            Tag             =   "Direccion/Dpto.|N|S|0|999|scarep|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   1845
            Width           =   585
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1065
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
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
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   720
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1065
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direcci�n"
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
            Left            =   150
            TabIndex        =   64
            Top             =   1845
            Width           =   900
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   1080
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   1860
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
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
            Index           =   19
            Left            =   4275
            TabIndex        =   63
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
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
            Left            =   150
            TabIndex        =   62
            Top             =   720
            Width           =   375
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
            Index           =   7
            Left            =   150
            TabIndex        =   61
            Top             =   1080
            Width           =   915
         End
      End
      Begin VB.Frame FrameOtros 
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
         Height          =   2250
         Left            =   8025
         TabIndex        =   50
         Top             =   405
         Width           =   7815
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   113
            Text            =   "1234567891"
            Top             =   1830
            Visible         =   0   'False
            Width           =   1785
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
            Left            =   1710
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "Ref|T|S|||scarep|refclien|||"
            Text            =   "12345678901234567890"
            Top             =   1785
            Width           =   2580
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
            Index           =   2
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   77
            Text            =   "123456789"
            Top             =   660
            Width           =   1320
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
            Index           =   3
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   81
            Text            =   "1234567891"
            Top             =   1440
            Width           =   1320
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
            Index           =   4
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   86
            Text            =   "1234567891"
            Top             =   1050
            Width           =   1320
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
            Left            =   1710
            MaxLength       =   7
            TabIndex        =   13
            Tag             =   "N� Reparaci�n|N|S|0|9999999|scarep|numrepar|0000000|S|"
            Text            =   "Text1"
            Top             =   645
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
            Index           =   4
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fecha reparacion|F|N|||scarep|fecrepar|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1410
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
            Index           =   3
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha entrada|F|N|||scarep|fecentre|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1035
            Width           =   1350
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
            Left            =   2430
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   51
            Text            =   "Text2"
            Top             =   270
            Width           =   5160
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
            Index           =   5
            Left            =   1710
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "Operador|N|N|0|9999|scarep|codtraba|0000|N|"
            Text            =   "Te"
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label1 
            Caption         =   "Baja"
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
            Index           =   23
            Left            =   4485
            TabIndex        =   112
            Top             =   1830
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Ref. cliente"
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
            Left            =   120
            TabIndex        =   106
            Top             =   1830
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "N�Mantenimiento"
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
            Left            =   4485
            TabIndex        =   58
            Top             =   1110
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Ult.Reparaci�n"
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
            Left            =   4485
            TabIndex        =   57
            Top             =   750
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Fin Garantia"
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
            Left            =   4485
            TabIndex        =   56
            Top             =   1470
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "N�Reparaci�n"
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
            Left            =   120
            TabIndex        =   55
            Top             =   660
            Width           =   1515
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1455
            Picture         =   "frmRepEntReparacionesGR.frx":0AD0
            ToolTipText     =   "Buscar fecha"
            Top             =   1410
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "F.Entrada"
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
            TabIndex        =   54
            Top             =   1050
            Width           =   1065
         End
         Begin VB.Label Label9 
            Caption         =   "F.Reparaci�n"
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
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1440
            Picture         =   "frmRepEntReparacionesGR.frx":0B5B
            ToolTipText     =   "Buscar fecha"
            Top             =   1050
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Operador"
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
            Index           =   9
            Left            =   120
            TabIndex        =   52
            Top             =   270
            Width           =   915
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1440
            ToolTipText     =   "Buscar trabajador"
            Top             =   285
            Width           =   240
         End
      End
      Begin VB.Label Label3 
         Caption         =   "T. Externo"
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
         Index           =   37
         Left            =   -72840
         TabIndex        =   193
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Orden de trabajo"
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
         Left            =   -74760
         TabIndex        =   192
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "RPM"
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
         Index           =   35
         Left            =   -63960
         TabIndex        =   191
         Top             =   6000
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Pot (Kw)"
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
         Index           =   34
         Left            =   -65760
         TabIndex        =   190
         Top             =   6000
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Pot(CV)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   -67800
         TabIndex        =   189
         Top             =   6000
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Caudal"
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
         Index           =   32
         Left            =   -71400
         TabIndex        =   188
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de rodete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   31
         Left            =   -74760
         TabIndex        =   186
         Top             =   5880
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "V"
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
         Index           =   30
         Left            =   -67800
         TabIndex        =   185
         Top             =   5520
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "I (A)"
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
         Index           =   29
         Left            =   -64560
         TabIndex        =   184
         Top             =   5520
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "N� Serie"
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
         Index           =   28
         Left            =   -67800
         TabIndex        =   183
         Top             =   5040
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Marca"
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
         Index           =   27
         Left            =   -67800
         TabIndex        =   182
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Modelo"
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
         Index           =   26
         Left            =   -67800
         TabIndex        =   181
         Top             =   4560
         Width           =   705
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
         Left            =   -66000
         TabIndex        =   180
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   165
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   -74640
         X2              =   -62160
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label3 
         Caption         =   "F. Alb"
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
         Left            =   -63360
         TabIndex        =   179
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "N� Expedicion"
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
         Left            =   -65760
         TabIndex        =   178
         Top             =   1200
         Width           =   2100
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia / Cliente / Matr�cula"
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
         Left            =   -70200
         TabIndex        =   177
         Top             =   1200
         Width           =   2865
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   174
         Top             =   480
         Width           =   2535
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -62280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "A�o"
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
         Left            =   -74760
         TabIndex        =   173
         Top             =   5520
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "H (m.c.a)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   -71880
         TabIndex        =   172
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "N� Serie"
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
         Left            =   -74760
         TabIndex        =   171
         Top             =   5040
         Width           =   885
      End
      Begin VB.Label Label3 
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
         Left            =   -72960
         TabIndex        =   170
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         Left            =   -65640
         TabIndex        =   169
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Modelo"
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   168
         Top             =   4560
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "N�Curva"
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
         Index           =   13
         Left            =   -71640
         TabIndex        =   167
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Marca"
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
         Index           =   12
         Left            =   -74760
         TabIndex        =   166
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vertical"
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
         Left            =   -68760
         TabIndex        =   164
         Top             =   2400
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pozo"
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
         Left            =   -69840
         TabIndex        =   163
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Vertical"
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
         Index           =   8
         Left            =   -71160
         TabIndex        =   162
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Horizontal"
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
         Index           =   7
         Left            =   -72360
         TabIndex        =   161
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label Label3 
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
         Left            =   -67320
         TabIndex        =   160
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   -69840
         TabIndex        =   159
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         Left            =   -72240
         TabIndex        =   158
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   157
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   156
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   155
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -62280
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -72330
         ToolTipText     =   "Buscar servicio"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   16
         Left            =   -74640
         TabIndex        =   103
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label9 
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
         Height          =   195
         Index           =   5
         Left            =   -73680
         TabIndex        =   102
         Top             =   5160
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.entrega"
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
         Left            =   -73680
         TabIndex        =   101
         Top             =   4680
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   -72375
         Picture         =   "frmRepEntReparacionesGR.frx":0BE6
         ToolTipText     =   "Buscar fecha"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Imp. reparaci�n"
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
         Left            =   -73680
         TabIndex        =   100
         Top             =   4110
         Width           =   1590
      End
      Begin VB.Label Label12 
         Height          =   255
         Index           =   1
         Left            =   -67080
         TabIndex        =   99
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Aceptado"
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
         Left            =   -66765
         TabIndex        =   98
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00972E0B&
         BorderWidth     =   2
         X1              =   -71400
         X2              =   -62160
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00972E0B&
         BorderWidth     =   2
         X1              =   -73200
         X2              =   -62160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Presupuesto"
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
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   97
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Importe 1�"
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
         Left            =   -73680
         TabIndex        =   96
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Importe 2�"
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
         Left            =   -73680
         TabIndex        =   95
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha presupuesto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70710
         TabIndex        =   94
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   -68790
         Picture         =   "frmRepEntReparacionesGR.frx":0C71
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha aprobaci�n"
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
         Left            =   -70710
         TabIndex        =   93
         Top             =   1590
         Width           =   1785
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -68790
         Picture         =   "frmRepEntReparacionesGR.frx":0CFC
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio de asistencia t�cnica"
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
         Height          =   240
         Index           =   13
         Left            =   -74760
         TabIndex        =   92
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.envio"
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
         Left            =   -73680
         TabIndex        =   91
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   -72330
         Picture         =   "frmRepEntReparacionesGR.frx":0D87
         ToolTipText     =   "Buscar fecha"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio SAT"
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
         Left            =   -73680
         TabIndex        =   90
         Top             =   2640
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "N� Resguardo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -73680
         TabIndex        =   89
         Top             =   3600
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   135
      TabIndex        =   45
      Top             =   765
      Width           =   16050
      Begin VB.CheckBox chkPresupuesto 
         Caption         =   "Presupuesto"
         Enabled         =   0   'False
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
         Left            =   14235
         TabIndex        =   2
         Tag             =   "Presupuesto|N|N|||scarep|presupue||N|"
         Top             =   285
         Width           =   1560
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
         Left            =   4545
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Art�culo|T|N|||scarep|codartic||N|"
         Text            =   "Text1"
         Top             =   285
         Width           =   2430
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
         Left            =   7005
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   285
         Width           =   7125
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
         Index           =   0
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "N� Serie|T|S|||scarep|numserie||N|"
         Text            =   "Text1"
         Top             =   285
         Width           =   2115
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   990
         Picture         =   "frmRepEntReparacionesGR.frx":0E12
         Tag             =   "-1"
         ToolTipText     =   "Buscar N� Serie"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Art�culo"
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
         Left            =   3450
         TabIndex        =   48
         Top             =   285
         Width           =   795
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   4275
         ToolTipText     =   "Buscar art�culo"
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "N� Serie"
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
         Left            =   120
         TabIndex        =   46
         Top             =   285
         Width           =   930
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
      Index           =   16
      Left            =   2340
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   209
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   10890
      Visible         =   0   'False
      Width           =   6615
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
      Left            =   13995
      TabIndex        =   42
      Top             =   10830
      Width           =   1065
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
      Left            =   15165
      TabIndex        =   43
      Top             =   10845
      Width           =   1065
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
      Left            =   15165
      TabIndex        =   38
      Top             =   10830
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   10740
      Width           =   2175
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
         TabIndex        =   41
         Top             =   180
         Width           =   1875
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8760
      Top             =   3360
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
      Left            =   9960
      Top             =   3720
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
   Begin VB.Label Label1 
      Caption         =   "Centro de coste"
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
      Index           =   17
      Left            =   9045
      TabIndex        =   105
      Top             =   10575
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n"
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
      Left            =   2340
      TabIndex        =   44
      Top             =   10575
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   39
      Top             =   10860
      Visible         =   0   'False
      Width           =   2175
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
End
Attribute VB_Name = "frmRepEntReparacionesGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public ControlRep As Boolean 'Para saber si se llama en el menu ppal desde
                             'Mantenimiento de Reparaciones o desde Control de Reparaciones
Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schrep, y solo en modo de consulta
Public EntradaEquipo As String 'SI desde avisos le han dado a meter equipo.


Private ControlReparacionAjustado As Boolean


Private WithEvents frmB As frmBasico2 'uscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB1 As frmBasico2 'Form para busquedas
Attribute frmB1.VB_VarHelpID = -1

Private WithEvents frmB2 As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB2.VB_VarHelpID = -1
Private WithEvents frmB3 As frmBuscaGrid 'Numeros de serie repetidos
Attribute frmB3.VB_VarHelpID = -1

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'frmFacClientesGr 'Form Mantenimiento Clientes�
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmNSeries As frmBasico2 'frmRepNumSerie2GR 'Form Mantenimiento N� Series
Attribute frmNSeries.VB_VarHelpID = -1
Private WithEvents frmNSeries2 As frmRepNumSerie2GR 'Form Mantenimiento N� Series
Attribute frmNSeries2.VB_VarHelpID = -1
Private WithEvents frmTraba As frmBasico2 'frmAdmTrabajadores  'Form Mantenimiento Trabajadores
Attribute frmTraba.VB_VarHelpID = -1
Private WithEvents frmMoti As frmRepMotivosPend  'Form Mantenimiento Motivos Ptes. Rep.
Attribute frmMoti.VB_VarHelpID = -1

Private WithEvents frmCliV As frmBasico2 'frmFacClientesV
Attribute frmCliV.VB_VarHelpID = -1

Private WithEvents frmTpAve As frmtipave
Attribute frmTpAve.VB_VarHelpID = -1
Private WithEvents frmSAT   As frmManSat
Attribute frmSAT.VB_VarHelpID = -1
Private WithEvents frmTraRea As frmManTraReali
Attribute frmTraRea.VB_VarHelpID = -1

Private WithEvents frmList As frmListadoPed 'Listados para pasar de Pedido -> Albaran
Attribute frmList.VB_VarHelpID = -1


Private Modo As Byte
Private ModoAnterior As Byte

Private ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Private HaDevueltoDatos As Boolean


Dim NombreTabla As String 'Nombre de la Tabla Cabecera
Dim NomTablaLineas As String 'Nombre de la Tabla de lineas

Dim Ordenacion As String
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim PrimeraVez As Boolean
Dim PrimeraVezForm As Boolean
Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla sserie o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim CadenaConsulta As String
Dim CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido
Dim CadenaSQLHco As String
Dim ImprimeAlb As Boolean 'Para saber cuando vuelve de Generar ALbaran si se ha solicitado Imprimir Albaran o no
Dim FechaAlb As String

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo
Dim Indice As Byte
'Dim PrimeraVez As Boolean

Dim ValorAntesFoco As String

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEuler_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPresupuesto_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSCAR
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarCabecera Then
                    'Ficha tecnica
                    If InstalacionEsEulerTaxco Then ActualizaBDFicha
                
                    If EntradaEquipo <> "" Then
                        'Viene de entrada equipo
                        'BotonImprimir (62)
                        BotonImprimir2 True, 0
                        CadenaDesdeOtroForm = "OK"
                        Unload Me
                        Exit Sub
                    End If
                End If
            End If
        Case 4 'MODIFICAR
            If DatosOk Then
                 'El campo numaviso lo tengo que dejar con el valor que tiene
                 'marzo2010 Text1(15).Text = DBLet(Me.Data1.Recordset!numaviso, "T")
                 If ModificaDesdeFormulario(Me, 1) Then
                    If InstalacionEsEulerTaxco Then ActualizaBDFicha
                     TerminaBloquear
                     PosicionarData
                 End If
                 'Vuelvo a
                 'Mostrar SOLO el numero de aviso, no la fecha de donde venia
                 'If Me.Text1(15).Text <> "" Then Text1(15).Text = RecuperaValor(Text1(15).Text, 1)

             End If
             
        Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slirep'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    
                    NumRegElim = Data2.Recordset!numlinea
                    
                    
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
'--                    PonerBotonCabecera True
                    lblIndicador.Caption = ""
                    PosicionarData2

                    PonerModo 2 '++
                    BloquearTxt Text2(16), True
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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



Private Sub cmdAux_Click(Index As Integer)
 Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            
        Case 1 'Busqueda de Cod. Artic
            Indice = 20
            Set frmA = New frmBasico2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo Busqueda
'            frmA.DesdeTPV = False
'            frmA.Show vbModal
            AyudaArticulos frmA, txtAux(Index)
            Set frmA = Nothing
            
        Case 2
            AbrirForm_CentroCoste
            
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    

    Set frmB2 = New frmBuscaGrid
    If vParamAplic.ContabilidadNueva Then
        frmB2.vCampos = "Codigo|ccoste|codccost|T||20�Descripci�n|ccoste|nomccost|T||70�"
        frmB2.vTabla = "ccoste"
    Else
        frmB2.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
        frmB2.vTabla = "cabccost"
    End If
    
    frmB2.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB2.vDevuelve = "0|1|"
    frmB2.vTitulo = "Centros de coste"
    frmB2.vselElem = 0
    frmB2.vConexionGrid = conConta
    
    frmB2.Show vbModal
    Set frmB2 = Nothing
    
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'BUSCAR
            LimpiarCampos
            PonerModo 0
        Case 3 'INSERTAR
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'MODIFICAR
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
         Case 5 'LINEAS Detalle
            TerminaBloquear
            CargaTxtAux False, False
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            BloquearTxt Text2(16), True
            ModificaLineas = 0
'            PonerBotonCabecera True
            Me.lblIndicador.Caption = ""
            PonerModo 2
            Me.DataGrid1.Enabled = True
    End Select
    PonerFoco Text1(0)
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton de cabecera

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
        'DataGrid1.visible = False
        

    End If
End Sub


Private Sub HabilitarFrames(b As Boolean)
    Me.Frame3.Enabled = Not b
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numrepar", Text1(2).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.Text2(8).Text = PonerNombreCCoste(Me.txtAux(8))
            End If
        Else
            Text2(16).Text = ""
            Text2(8).Text = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVezForm Then
        PrimeraVezForm = False
        DoEvents
        Screen.MousePointer = vbHourglass
        '--------------------------------
        
        
'--
'        If ControlReparacionAjustado Then
            'Cargamos el DATA�
            DataGrid1.visible = True
            CargaGrid DataGrid1, Data2, False
'        End If
        
        CargaDatosAviso

    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim I As Integer

    PrimeraVezForm = True

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
'    For kCampo = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
'    Next kCampo
    
    For I = 1 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = imgBuscar(0).Picture
    Next I

    'ICONOS de La toolbar
'    btnAnyadir = 5
'    btnPrimero = 17 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'A�adir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(10).Image = 10 'Mto Lineas
'        .Buttons(11).Image = 26 'Confirmar Reparaci�n
'        .Buttons(12).Image = 16 'Imprimir
'        .Buttons(14).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
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

    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 36 ' confirmar reparacion
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
    
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.visible = False

    'Ocultar los Textos de Reparacion si no en Control de Rep
    If InstalacionEsEulerTaxco Then
        ControlReparacionAjustado = True
    Else
        ControlReparacionAjustado = ControlRep
    End If
'--
'    Label1(6).visible = ControlReparacionAjustado
'    Text1(12).visible = ControlReparacionAjustado
'    Text1(13).visible = ControlReparacionAjustado
'    Text1(14).visible = ControlReparacionAjustado
'++
    Label1(6).Enabled = ControlReparacionAjustado
    Text1(12).Enabled = ControlReparacionAjustado
    Text1(13).Enabled = ControlReparacionAjustado
    Text1(14).Enabled = ControlReparacionAjustado
    If Not Text1(12).Enabled Then Text1(12).Tag = ""
    If Not Text1(13).Enabled Then Text1(13).Tag = ""
    If Not Text1(14).Enabled Then Text1(14).Tag = ""
    
    
    
    'Trabajo realizado  si es control reparacion o en HCO
    kCampo = 0
    'If ControlRep2 Or EsHistorico Then kCampo = 1
    If ControlReparacionAjustado Or EsHistorico Then kCampo = 1
'--
'    Label1(15).visible = (kCampo = 1)
'    Me.imgBuscar(7).visible = (kCampo = 1)
'    Text1(24).visible = (kCampo = 1)
'    Text2(24).visible = (kCampo = 1)
'++
    Me.imgBuscar(7).Enabled = (kCampo = 1)
    Text1(24).Enabled = (kCampo = 1)
    Text2(24).Enabled = (kCampo = 1)
    If Not Text1(24).Enabled Then Text1(24).Tag = ""
    
    'La solapa de las lineas
'--
'    SSTab1.TabVisible(2) = ControlReparacionAjustado
'    SSTab1.TabVisible(3) = InstalacionEsEulerTaxco   'vParamAplic.NumeroInstalacion = vbEuler
'++
    SSTab1.TabEnabled(2) = ControlReparacionAjustado
    SSTab1.TabEnabled(3) = InstalacionEsEulerTaxco   'vParamAplic.NumeroInstalacion = vbEuler
'--
'    '++
'    Me.FrameAux.visible = ControlReparacionAjustado
'    SSTab1.TabVisible(2) = False
'++
    Me.FrameAux.Enabled = ControlReparacionAjustado
    SSTab1.TabEnabled(2) = False
    SSTab1.TabVisible(2) = False

'--    'Si es Hist�rico no aparece codmotre
'    Label1(5).visible = Not EsHistorico
'    imgBuscar(5).visible = Not EsHistorico
'    Text1(11).visible = Not EsHistorico
'    Text2(11).visible = Not EsHistorico
'++
    Label1(5).Enabled = Not EsHistorico
    imgBuscar(5).Enabled = Not EsHistorico
    Text1(11).Enabled = Not EsHistorico
    Text2(11).Enabled = Not EsHistorico
    If Not Text1(11).Enabled Then Text1(11).Tag = ""
    
'--    'Si es hco no tiene el dato de numaviso ni fecaviso
'    Text1(15).visible = Not EsHistorico
'    Text1(36).visible = Not EsHistorico
'    Label1(11).visible = Not EsHistorico
'    Label1(18).visible = Not EsHistorico
'++
    Text1(15).Enabled = Not EsHistorico
    Text1(36).Enabled = Not EsHistorico
    Label1(11).Enabled = Not EsHistorico
    Label1(18).Enabled = Not EsHistorico
    If Not Text1(15).Enabled Then Text1(15).Tag = ""
    If Not Text1(36).Enabled Then Text1(36).Tag = ""
    
    
'--    'Si es hco aparecen el t�cnico de la reparacion
'    Label1(21).visible = EsHistorico
'    imgBuscar(10).visible = EsHistorico
'    Text1(37).visible = EsHistorico
'    Text2(37).visible = EsHistorico
'++
    Label1(21).Enabled = EsHistorico
    imgBuscar(10).Enabled = EsHistorico
    Text1(37).Enabled = EsHistorico
    Text2(37).Enabled = EsHistorico
    If Not Text1(37).Enabled Then Text1(37).Tag = ""
    
    
    'Si es Hist�rico no aparece fecentre 'Fecha Prev. entrega Repar
    'David: Hemos metido el campo en la BD
    'Label4.visible = Not EsHistorico
    'imgFecha(1).visible = Not EsHistorico
    'Text1(4).visible = Not EsHistorico
    
'--    'Si es Hist�rico no aparece Presupuesto
'    Me.chkPresupuesto.visible = Not EsHistorico
'++
    Me.chkPresupuesto.Enabled = Not EsHistorico
    If Not chkPresupuesto.Enabled Then chkPresupuesto.Tag = ""

'--    'Campos que solo aparecen en el Hist�rico
'    Text2(0).visible = EsHistorico
'    Text2(14).visible = EsHistorico
'    Text2(15).visible = EsHistorico
'    Label1(8).visible = EsHistorico
'    Label1(22).visible = EsHistorico
'    Label1(10).visible = EsHistorico
'    imgVerAlbaran.visible = EsHistorico
'++
    Text2(0).Enabled = EsHistorico
    Text2(14).Enabled = EsHistorico
    Text2(15).Enabled = EsHistorico
    Label1(8).Enabled = EsHistorico
    Label1(22).Enabled = EsHistorico
    Label1(10).Enabled = EsHistorico
    imgVerAlbaran.Enabled = EsHistorico
    
'--    'No se ven en hco
'    Label1(23).visible = Not EsHistorico
'    Text2(6).visible = Not EsHistorico
'++
    Label1(23).Enabled = Not EsHistorico
    Text2(6).Enabled = Not EsHistorico
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "REP"
    PrimeraVez = True
    
    'Comprobar si es Departamento o Direccion
    Me.Label1(2).Caption = DevuelveTextoDepto(True)
    
    
    
    If Not EsHistorico Then
        NombreTabla = "scarep" 'Tabla Cabecera Reparaciones
        NomTablaLineas = "slirep" 'Tabla Lineas Reparaciones
        Me.Caption = "Reparaciones"
    Else
        NombreTabla = "schrep"
        NomTablaLineas = "slirep"
        CargarTagsHco Me, "scarep", NombreTabla
        'Leer estos datos almacenados en la tabla del Historico
        Text2(1).Tag = "Cod. Art�culo|T|N|||schrep|nomartic||N|"
        Text2(2).Tag = "Ult. Reparac|F|S|||schrep|ultrepar|dd/mm/yyyy|N|"
        Text2(3).Tag = "Fin Garantia|F|N|||schrep|fingaran|dd/mm/yyyy|N|"
        Text2(4).Tag = "N� Mantenim.|T|S|||schrep|nummante||N|"
        Me.Caption = "Hist�rico Reparaciones"
        
        
        'Datos Albaran
'-- lo dejo donde estan
'        Label1(10).Left = 240
'        Text2(15).Left = 240
'        Label1(8).Left = 1240
'        Text2(0).Left = 1240
'        Label1(22).Left = 1980
'        Text2(14).Left = 1980
'        imgVerAlbaran.Top = Text2(15).Top + 30
'        imgVerAlbaran.Left = 1980 + Text2(14).Width + 120
    End If
    
    
    
    
    
    Ordenacion = " ORDER BY numrepar "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numrepar = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbLightBlue
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Articulos
    If Indice = 1 Then
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            'Estamos en Cabecera
            'Recupera todo el registro de N� Serie
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else  'Llama desde Prismatico Direcciones/Departamentos
            Text1(7).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(7).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Text1(7).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
        Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB1_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    txtAux(8).Text = RecuperaValor(CadenaDevuelta, 1)
    Text2(8).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub frmB3_Selecionado(CadenaDevuelta As String)
     Text1(1).Text = RecuperaValor(CadenaDevuelta, 1)
     Text2(1).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Clientes
    Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    If Modo <> 1 Then Text1(34).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCliV_DatoSeleccionado(CadenaSeleccion As String)
    Text1(28).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(34).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(28).Text)

End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Indice = Val(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Cuando pasa de Reparacion -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String
Dim Rs As ADODB.Recordset
Dim cad1 As String, Cad2 As String

    'Seleccionar algunos campos del Cliente
    vSQL = "Select proclien, codagent, codforpa, dtoppago, dtognral, tipofact "
    vSQL = vSQL & " FROM sclien WHERE codclien=" & Text1(6).Text
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad1 = RecuperaValor(CadenaSeleccion, 1) 'trab. albaran
    Cad2 = RecuperaValor(CadenaSeleccion, 2) 'trab. prepara material
    FechaAlb = RecuperaValor(CadenaSeleccion, 4)

    'Construimos parte de la SQL para insertar en tabla de Albaranes
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "', " 'Fecha Albaran
    vSQL = vSQL & "0, " 'facturar s/n
    vSQL = vSQL & Text1(6).Text & ", " & DBSet(Text1(34).Text, "T") & ", " 'nomclien
    
    
    'Aqui van los datos del cliente
    'Nuevo Dic 2009
    'vSQL = vSQL & DBSet(Text2(10).Text, "T") & ", " & DBSet(Text2(12).Text, "T") & ", " & DBSet(Text2(13).Text, "T") & ", " 'domclien, codpobla, pobclien
    'domclien, codpobla, pobclien
    vSQL = vSQL & DBSet(Text1(30).Text, "T") & ", " & DBSet(Text1(31).Text, "T") & ", " & DBSet(Text1(32).Text, "T") & ", "
    'proclien, nifclien, telclien "
    vSQL = vSQL & DBSet(Text1(33).Text, "T") & ", '" & Text1(28).Text & "', '" & Text1(29).Text & "', "
    vSQL = vSQL & DBSet(Text1(7).Text, "N", "S") & ", " & DBSet(Text2(7).Text, "T") & ", " ' nomdirec
    'Nuevo 13 Enero 10
    'vSQL = vSQL & ValorNulo & ", " & cad1 & ", "  'referenc, codtraba(ped), "
    vSQL = vSQL & DBSet(Text1(35).Text, "T", "S") & ", " & cad1 & ", "   'referenc, codtraba(ped), "
    vSQL = vSQL & DBSet(Text1(5).Text, "N", "S") & ", " 'Trabajador de pedido
    vSQL = vSQL & Cad2 & ", " 'Material Preparado por
    vSQL = vSQL & DBSet(Rs!CodAgent, "N") & ", " & DBSet(Rs!codforpa, "N") & ", " '"codagent, codforpa, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 3) & ", " 'Cod Envio
    vSQL = vSQL & DBSet(Rs!DtoPPago, "N") & ", " & DBLet(Rs!DtoGnral, "N") & ", " & DBLet(Rs!TipoFact, "N") & ", " '" '"dtoppago, dtognral, tipofact,
    
    'ANTIGUAS OBSERVACIONES. 19 JUN 07
    'vSQL = vSQL & DBSet(Text1(8).Text, "T") & ", " & DBSet(Text1(9).Text, "T") & ", " & DBSet(Text1(10).Text, "T") & ", " 'observa01, observa02, observa03,
    'vSQL = vSQL & DBSet(Text1(14).Text, "T") & ", " & DBSet(Text1(13).Text, "T") & ", " 'observa04, observa05, "
    
    'AHORA
    vSQL = vSQL & DBSet(Text1(14).Text, "T") & ", " & DBSet(Text1(13).Text, "T") & ", " & DBSet(Text1(12).Text, "T") & ", " 'observa01, observa02, observa03,
    vSQL = vSQL & DBSet("N�mero serie: " & Text1(0).Text, "T") & ", " & DBSet("Art�culo: " & Text1(1).Text & " - " & Text2(1).Text, "T") & ", " 'observa04, observa05, "
    
    vSQL = vSQL & ValorNulo & ", " & ValorNulo & ", "  'N� Oferta, fecha de la Oferta
    vSQL = vSQL & Text1(2).Text & ", '"  'N� Pedido
    vSQL = vSQL & Format(Text1(3).Text, FormatoFecha) & "', " & ValorNulo 'Fecha Pedido, Semana entrega
    
    'Faltara el nomclien
    
    'vSQL = vSQL & Text1(18).Text 'Semana entrega Pedido
    CadenaSQL = vSQL
    
    Rs.Close
    Set Rs = Nothing
    
    CadenaSQLHco = cad1 & ", " & Cad2 & ", material, tipoaver, motivore, textore1, textore2, textore3 "
    
    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    ImprimeAlb = CBool(RecuperaValor(CadenaSeleccion, 5))
End Sub


Private Sub frmMoti_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Motivos Pendientes Rep.
    Text1(11).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmNSeries_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento N� Serie
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'num serie
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 2) 'cod artic
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 3) ' desc artic
    'DAVID.
    'Si me va a devolver VACIO no lo borro por si , y solo si, viene de los avisos
    If EntradaEquipo = "" Then
        'mantenimiento normal
        Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
    Else
        If RecuperaValor(CadenaSeleccion, 4) = "" Then
            'NO HACEMOS NADA. NO vaciamos el campo codcliente
        Else
            Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
        End If
    End If
End Sub

Private Sub frmNSeries2_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento N� Serie
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'num serie
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 2) 'cod artic
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 3) ' desc artic
    'DAVID.
    'Si me va a devolver VACIO no lo borro por si , y solo si, viene de los avisos
    If EntradaEquipo = "" Then
        'mantenimiento normal
        Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
    Else
        If RecuperaValor(CadenaSeleccion, 4) = "" Then
            'NO HACEMOS NADA. NO vaciamos el campo codcliente
        Else
            Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
        End If
    End If
End Sub

Private Sub frmSAT_DatoSeleccionado(CadenaSeleccion As String)
    PonValoresDatoSeleccionado 21, CadenaSeleccion
End Sub

Private Sub frmTpAve_DatoSeleccionado(CadenaSeleccion As String)
    PonValoresDatoSeleccionado 23, CadenaSeleccion
End Sub

Private Sub frmTraba_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Trabajadores
    'Text1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    'Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
    PonValoresDatoSeleccionado CInt(Indice), CadenaSeleccion
End Sub

Private Sub PonValoresDatoSeleccionado(Indice As Integer, CadenaSeleccion As String)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTraRea_DatoSeleccionado(CadenaSeleccion As String)
PonValoresDatoSeleccionado 24, CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim cadMen As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'N� Serie
'            Set frmNSeries = New frmRepNumSerie2GR
'            frmNSeries.DatosADevolverBusqueda = "0"
'            frmNSeries.DatoAInsertar = ""
'            frmNSeries.Show vbModal
'            Set frmNSeries = Nothing
'            Indice = 0
            
            Set frmNSeries = New frmBasico2
            AyudaNrosSerie frmNSeries, Text1(0), , False
            Set frmNSeries = Nothing
            
            
            
        Case 1 'Codigo Articulo
            If Modo = 3 Or Modo = 4 Then
                If Text1(0).Text <> "" Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            Indice = 1
            Set frmA = New frmBasico2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abrir en Modo busqueda
'            frmA.DesdeTPV = False
'            frmA.Show vbModal
            AyudaArticulos frmA, Text1(Indice)
            Set frmA = Nothing
        
        Case 2, 10 'Cod. Trabajador (Operador)
            If Index = 2 Then
                Indice = 5
            Else
                Indice = 37
            End If
'            Set frmTraba = New frmAdmTrabajadores
'            frmTraba.DatosADevolverBusqueda = "0"
'            frmTraba.Show vbModal
            Set frmTraba = New frmBasico2
            AyudaTrabajadores frmTraba, Text1(Indice)
            Set frmTraba = Nothing
            
            
            
        
        Case 3 'Cod. Cliente
'            Set frmCli = New frmFacClientesGr
'            frmCli.DatosADevolverBusqueda = "0"
'            frmCli.Show vbModal
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(6).Text
            Set frmCli = Nothing
            Indice = 6
            
        Case 4 'Direc/Dpto del Cliente
            'Mostrar las Direc. o Dptos del cliente seleccionado
            If Trim(Text1(6).Text) = "" Then
               cadMen = DevuelveTextoDepto(False)
              
               MsgBox "Debe seleccionar un cliente para mostrar su(s) " & cadMen & ".", vbInformation
               Screen.MousePointer = vbDefault
               Exit Sub
            Else
               EsCabecera = False
               MandaBusquedaPrevia " codclien= " & Val(Text1(6).Text)
               Indice = 7
            End If
             
        Case 5 'Cod. Motivo Pendiente Rep.
            Set frmMoti = New frmRepMotivosPend
            frmMoti.DatosADevolverBusqueda = "0"
            frmMoti.Show vbModal
            Set frmMoti = Nothing
            Indice = 11
            

        Case 6
            Set frmTpAve = New frmtipave
            frmTpAve.DatosADevolverBusqueda = "0"
            frmTpAve.Show vbModal
            Set frmTpAve = Nothing
            Indice = 10 'Para que ponga el foco en el siguiente
        Case 7
            Set frmTraRea = New frmManTraReali
            frmTraRea.DatosADevolverBusqueda = "0"
            frmTraRea.Show vbModal
            Set frmTraRea = Nothing
            Indice = 24 'Para que ponga el foco en el siguiente
        Case 8
            Set frmSAT = New frmManSat
            frmSAT.DatosADevolverBusqueda = "0"
            frmSAT.Show vbModal
            Set frmSAT = Nothing
            Indice = 20 'Para que ponga el foco en el siguiente
            
        Case 9
            'Clientes varios
'            Set frmCliV = New frmFacClientesV
'            frmCliV.DatosADevolverBusqueda = "0|"
'            frmCliV.Show vbModal
'            Set frmCliV = Nothing
            
            Indice = 29
            Set frmCliV = New frmBasico2
            AyudaClientesV frmCliV, Text1(Indice)
            Set frmCliV = Nothing
            
    End Select
    
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0: Indice = 3 'Fecha Reparacion
        Case 1: Indice = 4 'Fecha Entrega
        Case 2 To 4
            Indice = Index + 16
        Case 5
            Indice = 26
   End Select
   imgFecha(0).Tag = Indice

   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub imgVerAlbaran_Click()
    If Modo = 1 Then Exit Sub

    If Text2(15).Text <> "" Then
    
    
        CadenaSQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", Text2(0).Text, "T", , "numalbar", Text2(15).Text, "N")
        If CadenaSQL <> "" Then 'existe el Albaran
            'If vParamAplic.TipoFormularioClientes = 0 Then
                 With frmFacEntAlbaranes2
                    .hcoCodMovim = Format(Text2(15).Text, , "0000000")
                    .hcoCodTipoM = Text2(0).Text ' Comprobar esto
                    .Show vbModal
                End With
            'End If
        
        Else 'No existe en albaran, abrir Historico Factura
            With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                .hcoCodMovim = Format(Text2(15).Text, , "0000000")
                .hcoCodTipoM = Text2(0).Text
                .hcoFechaMov = CDate(Text2(14).Text)
                .Show vbModal
            End With
        End If
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar Linea
        BotonEliminarLinea
    Else
        BotonEliminar
    End If
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modifica linea
        BotonModificarLinea
    Else
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir linea
        BotonAnyadirLinea
    Else
        BotonAnyadir
    End If
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If (ModificaLineas = 1 Or ModificaLineas = 2) Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    If InstalacionEsEulerTaxco Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            If SSTab1.Tab = 3 Then PonerFoco txtEuler(0)
        End If
    End If
        
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ValorAntesFoco = Text1(Index).Text
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
    If Modo = 3 And Index = 1 Then
        'Si coje foco el articulo, entonces, si tiene numero de serie, NO dejo que
        'cambie el articulo, lo paso el foco a otro
        If Text1(0).Text <> "" Then KEYpress 13
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'Primer campo, fecla arriba
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 27 Then KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim totArtic As Integer
Dim PonerEnUno As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'N� serie
            If Text1(Index).Text = "" Then
            
                BloquearPorNumeroSerie False
                If Modo <> 1 Then PonerFoco Text1(1)
                Exit Sub
            End If
            If Modo = 1 Or Modo = 4 Then Exit Sub
            totArtic = ArticulosDelNSerie(Text1(Index).Text)
            PonerEnUno = False
            If totArtic = 0 Then
                'No se encontro ningun registro en la tabla sserie para ese valor de N� de serie
                If MsgBox("No existe el N� de Serie: " & Text1(Index).Text & ". �Desea crearlo?", vbQuestion + vbYesNo) = vbYes Then
                    AbrirNumSerie
                    PonerEnUno = True
                Else
                    BloquearPorNumeroSerie False
                    Exit Sub
                End If
                
                BloquearPorNumeroSerie True
                If PonerEnUno Then PonerFoco Text1(0)
                
            ElseIf totArtic = 1 Then
                'Solo hay un articulo que tiene ese n� de serie: Recuperar datos de
                'la tabla sserie
                Text1(1).Text = DevuelveDesdeBDNew(conAri, "sserie", "codartic", "numserie", Text1(0).Text, "T")
                Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "sartic", "nomartic")
                CargarDatosNSerie Text1(0).Text, Text1(1).Text
                ComprobarReparaciones Modo, Text1(0).Text, Text1(1).Text
                BloquearPorNumeroSerie False
                MensajeBaja
                If PonerEnUno Then PonerFoco Text1(0)
                
                
            Else
                'hay varios art�culos que tienen este n� de serie, hasta que no se
                'seleccione el codartic no se pueden recuperar los datos de la tabla sserie
                If Text1(1).Text = "" Then
                    'Busca numserie/codartic para numserie repetidos
                    BuscaNumserieRepetido
                
                    'Si despues de la funcion anterior, puede ser que se codartic tenga ya valor
                End If
                
                If Text1(1).Text <> "" Then
                    CargarDatosNSerie Text1(0).Text, Text1(1).Text
                    ComprobarReparaciones Modo, Text1(0).Text, Text1(1).Text
                    PonerFoco Text1(2)
                End If
                BloquearPorNumeroSerie True
                If PonerEnUno Then PonerFoco Text1(0)
            End If

        Case 1 'Codigo Articulo
            'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
            PonerDatosCodigoDescripcion Index
            
        'Fechas Reparacion, Fecha Entrega      :  Fecha presupu,aprobacion   : SAT: envio entrega
        Case 3, 4, 18, 19, 20, 26
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)

            'Comprobar que Fecha Rep. es posterior a la de Entrada
            If Index <= 4 Then
                If Not EsFechaIgualPosterior(Text1(3).Text, Text1(4).Text, True, "La Fecha de Reparaci�n debe ser posterior a la Fecha de Entrada.") Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
            End If
                
        Case 5, 37 'Cod Trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            PonerDatosCodigoDescripcion Index
        Case 6 'Cliente
            If Modo <> 1 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    If Modo = 1 Then 'Modo=1 Busqueda
    '                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                        PonerDatosCodigoDescripcion Index
                    Else 'Insertando
                        PonerDatosCliente2 Text1(Index).Text, False
                    End If
                Else
                    LimpiarDatosCliente
                End If
            End If
        Case 7 'Direc/dpto del cliente
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            If Text1(6).Text = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Text1(Index).Text = ""
                PonerFoco Text1(6)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
            Text2(Index).Text = devuelve 'Nombre direc. o dpto
            If devuelve = "" Then 'No existe el dpto
                
                devuelve = DevuelveTextoDepto(False)
                devuelve = "No existe" & devuelve & Text1(Index).Text & " para el cliente: "
                devuelve = devuelve & Text1(6).Text & " - " & Text1(34).Text
                MsgBox devuelve, vbInformation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 11 'Motivo pendiente reparacion
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smotre", "nommotre", "codmotre")
            
        Case 16, 17, 25
            PonerFormatoDecimal Text1(Index), 1 'Tipo 2: Decimal(10,4)
        'Case 21
        '    'Servicio ASISTENCIA TECNICA
        '    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smansat", "nomsat", "codsat")
        'Case 23
        '    'Tipo averia
        '    'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipave", "nomave", "codave")
        '
        'Case 24
        '    'Trabajao realizado
        '    'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smantr", "nomtrabajo", "codtrabajo")
            
           
        Case 21, 23, 24
            PonerDatosCodigoDescripcion Index
            
            
        Case 28
            'NIF de clientes varios
            If Not Text1(28).Locked Then
                If ValorAntesFoco <> Text1(28).Text Then PonerDatosClienteVario Text1(28).Text
            End If
        Case 31
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
            Text1(Index + 2) = devuelve
    End Select
End Sub



Private Sub PonerDatosCodigoDescripcion(Index As Integer)


    Select Case Index
        Case 1
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
            If Modo = 3 And Text1(Index).Text <> "" Then
                If Text1(0).Text = "" Then PonerFechaRepar False
            End If
        Case 5, 37
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            
        Case 6
            Text1(34).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            
        Case 21
            'Servicio ASISTENCIA TECNICA
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smansat", "nomsat", "codsat")
            
        Case 23
            'Tipo averia
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipave", "nomave", "codave")
            
        Case 24
            'Trabajao realizado
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smantr", "nomtrabajo", "codtrabajo")

    
    End Select
End Sub



Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    'campo Ampliaci�n linea y ENTER
    If Index = 16 And KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptar
End Sub

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
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click 'Modificar
        Case 3: mnEliminar_Click 'Eliminar
        
        Case 5: mnBuscar_Click 'Busqueda
        Case 6: mnVerTodos_Click 'Ver Todos
            
        Case 8 'Imprimir
            'If (Not ControlRep) And (Not EsHistorico) Then BotonImprimir (62)
            If InstalacionEsEulerTaxco Then
                If Not EsHistorico Then BotonImprimir2 True, 0
            Else
                If (Not ControlRep) And (Not EsHistorico) Then BotonImprimir2 True, 0
            End If
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then
        CadenaDesdeOtroForm = ""
        Unload Me
    End If
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim b2 As Boolean
Dim NumReg As Byte

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, (Modo = 2), NumReg
    b = (Kmodo = 2)
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    
        
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
        
        
    BloquearPorNumeroSerie Modo = 3
        
    '-------------------------------------------
    'Bloquear todos los Text Box que se llamen Text1
    BloquearText1 Me, Modo
    
    If InstalacionEsEulerTaxco Then BloquearFicha Modo = 0 Or Modo = 2 Or Modo = 5
    
    
    'N� Reparacion siempre bloqueado, es contador, salvo en Modo=Buscar
    If Modo <> 1 Then BloquearTxt Text1(2), True, True
    
    'N� aviso simepre bloeueado salvo buscar
    Text1(15).Enabled = (Modo = 1)
    Text1(36).Enabled = (Modo = 1)
    
    'Si el modo No es insertar o modi
    'el framecli estara activo seguro
    'Insertando/mod ya sera otra h�
    If Modo <> 3 Or Modo <> 4 Then FrameClientes.Enabled = True
       
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    b = b And Modo < 5
    Me.chkPresupuesto.Enabled = b ' (Modo = 3) Or (Modo = 4) 'Insertar o Modificar
    Me.Check1.Enabled = b
    Me.Combo1.Enabled = b
    
    
    'Para casi todas las ayudas
    'b = ((Modo = 3 Or Modo = 4) And (ControlRep2 = False)) Or Modo = 1
    If InstalacionEsEulerTaxco Then
        b2 = False
    Else
        b2 = ControlRep
    End If
    
    b2 = ((Modo = 3 Or Modo = 4) And (b2 = False)) Or Modo = 1
    
    'Sat,tipo...
    b2 = ((Modo = 3 Or Modo = 4) And True) Or Modo = 1

    
    
    For I = 0 To Me.imgBuscar.Count - 1
        If I < 5 Or I = 10 Then
            BloquearImg Me.imgBuscar(I), Not b
        Else
            'SAT, tipo averia...
            BloquearImg Me.imgBuscar(I), Not b2
        End If
    Next I
    'Me.imgBuscar(1).Enabled = (Modo = 1)
    Me.imgBuscar(1).Enabled = b2
    'La imagen del TRABAJO REALIZADO no se tiene que mostrar a no ser que haya entrado como reparacion
    Me.imgBuscar(7).visible = b2 And (ControlReparacionAjustado Or EsHistorico)
    
    Me.imgBuscar(10).visible = Me.imgBuscar(10).visible And EsHistorico
    Me.imgBuscar(5).visible = Me.imgBuscar(5).visible And Not EsHistorico
      
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b 'Si es insertar o modificar
    Next I
    
      '  Text1(1).TabIndex = 1
    
    
    b = Modo = 1 And EsHistorico
    For I = 1 To 4
        'Text2(i).Enabled = EsHistorico
        BloquearTxt Text2(I), Not b
    Next I
    
    
    If EsHistorico Then
        'Tengo visible los campos de albaran.
        'Entonces, si estoy en busqueda habilito los campos
        BloquearTxt Text2(0), Modo <> 1
        BloquearTxt Text2(15), Modo <> 1
        BloquearTxt Text2(14), Modo <> 1
    End If
    
    'Modo Linea de Ofertas
    b = (Modo = 5)
    Me.Label1(35).visible = b
    Me.Text2(16).visible = b
    Label1(17).visible = b And vEmpresa.TieneAnalitica
    Text2(8).visible = b And vEmpresa.TieneAnalitica
    BloquearTxt Text2(8), True
    BloquearTxt Text2(16), True
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu   'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
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
Dim B1 As Boolean
Dim I As Byte
Dim bAux As Boolean



    'Modo 2. Hay datos y estamos visualizandolos
    If InstalacionEsEulerTaxco Then
        B1 = Not EsHistorico
    Else
        B1 = ((Not ControlRep) Or (ControlRep And Modo = 5)) And (Not EsHistorico)
    End If
    Toolbar1.Buttons(1).Enabled = B1
    Me.mnNuevo.Enabled = B1
    Toolbar1.Buttons(2).Enabled = Not EsHistorico
    Me.mnModificar.Enabled = Not EsHistorico
    Toolbar1.Buttons(3).Enabled = B1
    Me.mnEliminar.Enabled = B1
    
    
    b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b And Not EsHistorico
    Me.mnModificar.Enabled = b And Not EsHistorico
    'Insertar
    Toolbar1.Buttons(1).Enabled = (b Or Modo = 0) And B1
    Me.mnNuevo.Enabled = (b Or Modo = 0) And B1
        
    'eliminar
    Toolbar1.Buttons(3).Enabled = b And B1
    Me.mnEliminar.Enabled = b And B1
    
'--
'    Toolbar1.Buttons(8).visible = Not EsHistorico
'    Toolbar1.Buttons(9).visible = Not EsHistorico
'    Me.mnBarra2.visible = Not EsHistorico
    
'    For I = 10 To 11
'        Toolbar1.Buttons(I).visible = ControlReparacionAjustado
'    Next I
    
'++
    Toolbar1.Buttons(8).Enabled = Not ControlReparacionAjustado And Not EsHistorico
    
    
    Toolbar5.Buttons(1).Enabled = ControlReparacionAjustado
    Me.FrameToolAux0.Enabled = ControlReparacionAjustado
    
    
    
    If ControlReparacionAjustado Then
        b = (Modo = 2)
        'Mto Lineas
'        Toolbar1.Buttons(10).Enabled = b
        'Confirmaci�n Reparaci�n
        Toolbar5.Buttons(1).Enabled = b
    End If
    
    '-------------------------------------
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    b = (Modo = 2) And DatosADevolverBusqueda = "" And ControlReparacionAjustado
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If Data2.Recordset Is Nothing Then
            bAux = False
        Else
            bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        End If
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I


End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPresupuesto.Value = 0
    Me.Check1.Value = 0
    Me.Combo1.ListIndex = -1
    
    If InstalacionEsEulerTaxco Then LimpiarFichaTecnica True
    
    
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonBuscar()
'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        If ControlReparacionAjustado Then CargaGrid DataGrid1, Data2, False
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
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    
    If ControlReparacionAjustado Then CargaGrid DataGrid1, Data2, False
    
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
    
    'Bloquear algunos campos
    'Febrero 2010.   NO bloqueamos el articulo. Si tiene numeroi de serie nos salimos de el en seguida
    'BloquearTxt Text1(1), True
    
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    Text1(5).Text = PonerTrabajadorConectado(NomTraba)
    Text2(5).Text = NomTraba
    
    
    HabilitarDatosCliente False
    
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
Dim I As Byte
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    
    
    'HabilitarDatosCliente
    Precio = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(6).Text)
    HabilitarDatosCliente Precio = 1
    Precio = ""
    
    'Como el campo N� Repar. es clave primaria, NO se puede modificar
    BloquearTxt Text1(2), True, True
    BloquearTxt Text1(1), True
    If ControlReparacionAjustado Then
        Me.chkPresupuesto.Enabled = False
        Text1(0).Locked = True
        For I = 3 To 7
            Text1(I).Locked = True
        Next I
        Me.imgBuscar(5).Enabled = True
        PonerFoco Text1(8)
    Else
        PonerFoco Text1(0)
    End If
    
    
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Reparacion (tabla: slirep)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    
    vWhere = Mid(ObtenerWhereCP, 7) & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
        
    SQL = ""
    SQL = SQL & "Va a Eliminar la Reparaci�n: " & Text1(2).Text & vbCrLf
    SQL = SQL & vbCrLf & "N� Serie: " & Text1(0).Text
    SQL = SQL & vbCrLf & "Artic. : " & Text1(1).Text & " - " & Text2(1).Text
    SQL = SQL & vbCrLf & vbCrLf & "�Desea continuar? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar N� Reparaci�n", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

    SQL = " WHERE numrepar=" & Data1.Recordset!numrepar
    
    'Eliminar las Lineas
    conn.Execute "Delete from " & NomTablaLineas & SQL
    
    'Eliminar Cabecera
    conn.Execute "Delete  from " & NombreTabla & SQL
               
               
    'Si es euler la ficha tecnica
    If InstalacionEsEulerTaxco Then conn.Execute "Delete  from scarepeu " & SQL
               
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Sub BotonEliminarLinea()
'Eliminar una linea De la Reparacion. (Tabla: slirep)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "�Seguro que desea eliminar la l�nea de la Reparaci�n?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Art�culo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        
        PonerModo 2
        
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Reparaci�n", Err.Description
End Sub


Private Sub BotonMtoLineas()

    'Por si acaso esta puesto el modo incorrecto
    If EsHistorico Or Not ControlReparacionAjustado Then Exit Sub
    

    SSTab1.Tab = 2

    ModificaLineas = 0
    PonerModo 5
    
 
   
    'Me.DataGrid1.visible = True
    'Esto antes estaba descomentado.  21 Abril de 2008
    'CargaGrid DataGrid1, Data2, True
    
    PonerBotonCabecera True
    
    If vEmpresa.TieneAnalitica Then
        If Not Data2.Recordset.EOF Then
            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.Text2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            Me.Text2(8).Text = ""
        End If
    End If
End Sub


Private Sub BotonConfirmarRep()
'Confirmar Reparacion
Dim b As Boolean
Dim cadMen As String, vWhere As String

    If MsgBox("�Desea Cerrar la Orden de Reparaci�n y Generar Albaran?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        Screen.MousePointer = vbHourglass
        b = SePuedeServirPedido(cadMen)
        If b Then 'Hay suficiente stock
            'Si hay stock generar albaran completo
            GenerarAlbaran
        ElseIf cadMen <> "" Then
            MsgBox cadMen, vbExclamation
        Else
            Screen.MousePointer = vbDefault
            'Si no se puede servir mostrar mensaje detallando y bloquear
            cadMen = "No hay suficiente Stock para servir la Reparaci�n. "
            cadMen = cadMen & vbCrLf & "�Desea Ver Detalle?"
            If MsgBox(cadMen, vbYesNo, "Contol de Stock") = vbYes Then
                vWhere = " WHERE numrepar = " & Text1(2).Text & " And sfamia.instalac = 0 "
                'DAVID### 09/09/2010
                'Ademas tiene que tener CONTROL DE STCOK el articulo
                vWhere = vWhere & " AND sartic.ctrstock=1 "
                frmMensajes.cadWhere = vWhere
                frmMensajes.vCampos = NomTablaLineas
                frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
                frmMensajes.Show vbModal
            End If
            Exit Sub
        End If
        'Pedir Datos para el Albaran: Operador, Fecha, Reparado por
        
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    If Modo = 3 Then
        If EntradaEquipo <> "" Then
            If Val(Text1(6).Text) <> RecuperaValor(EntradaEquipo, 3) Then
                MensajeNoCoinciden Text1(6).Text, True
                b = MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes
                CadenaDesdeOtroForm = ""
                If Not b Then Exit Function
            End If
        End If
        
        
        'AGOSTO 2012
        '-----------
        CadenaSQL = "codartic = " & DBSet(Text1(1).Text, "T") & " AND numserie"
        CadenaSQL = DevuelveDesdeBD(conAri, "codartic", "scarep", CadenaSQL, Text1(0).Text, "T")
        If CadenaSQL <> "" Then
            MsgBox "El art�culo-n�serie ya esta reparandose: " & Text1(1).Text & " / " & Text1(0).Text, vbExclamation
            CadenaSQL = ""
            Exit Function
        End If
    End If
    
        
        
    If b Then
        If InstalacionEsEulerTaxco Then
            'Si ha puesto albaran
            'CadenaSQLHco = ""
            '    If Me.txtEuler(20).Text <> "" Then
            '        CadenaSQL = "codtipom = 'ALO' AND numalbar"
            '        CadenaSQL = DevuelveDesdeBD(conAri, "numalbar", "scaalb", "numalbar", txtEuler(Index).Text)
            '        'Label3(36 o 37
            '        If CadenaSQL = "" Then
            '        MsgBox "El albaran de " & Label3(Index + 16).Caption & " NO existe", vbExclamation
            '

        End If
    End If
        
    If b Then MensajeBaja



    
    DatosOk = True
End Function



Private Sub MensajeBaja()
        'Comprobaremos si el articuo de numero de serie esta de baja
        CadenaSQL = "concat(fechabaja,""|"",desmotiv,""|"")"
        CadenaDesdeOtroForm = "sserie.codmotba=smotba.codmotiv AND not fechabaja is null AND numserie=" & DBSet(Text1(0).Text, "T") & " AND codartic "
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, CadenaSQL, "sserie,smotba", CadenaDesdeOtroForm, Text1(1).Text, "T")
        If CadenaDesdeOtroForm <> "" Then
                CadenaSQL = RecuperaValor(CadenaDesdeOtroForm, 1)
                CadenaDesdeOtroForm = "Motivo: " & RecuperaValor(CadenaDesdeOtroForm, 2) & vbCrLf
                CadenaDesdeOtroForm = "Fecha: " & Format(CadenaSQL, "dd/mm/yyyy") & vbCrLf & CadenaDesdeOtroForm
                CadenaSQL = String(30, "*") & vbCrLf
                CadenaSQL = CadenaSQL & "Numero serie/articulo de BAJA" & vbCrLf & vbCrLf & CadenaDesdeOtroForm & CadenaSQL
                MsgBox CadenaSQL, vbInformation
               
                CadenaDesdeOtroForm = ""
        End If
        CadenaSQL = ""
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String, Desc As String
Dim selElem As Byte

    'Llamamos a al form
    cad = ""
    If EsCabecera Then
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: sserie
'        cad = cad & ParaGrid(Text1(0), 20, "N� Serie")
'        cad = cad & ParaGrid(Text1(1), 25, "Artic.")
'        cad = cad & "Desc. Artic.|sartic|nomartic|T||40�"
'        cad = cad & ParaGrid(Text1(2), 15, "Num Rep.")
''        cad = cad & "Desc. Tipo|stipar|nomtipar|T||20�"
'
'        tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
''        tabla = tabla & " LEFT JOIN stipar ON " & NombreTabla & ".codtipar=stipar.codtipar"
'        If EsHistorico Then
'            Titulo = "Hist�rico Reparaciones"
'        Else
'            Titulo = "Reparaciones"
'        End If
'        selElem = 2
    Set frmB1 = New frmBasico2
    
    AyudaReparaciones frmB1, Text1(0)

    Set frmB1 = Nothing
    
    Exit Sub
    
   Else
        If vParamAplic.HayDeparNuevo = 1 Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        Else
            Titulo = "Obras Cliente: "
            Desc = "Obra"
        End If
        Titulo = Titulo & Text1(6).Text & " - " & Text1(34).Text 'Cod y Desc. Cliente
'        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||20�"
'        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||40�"
'        tabla = "sdirec"
        selElem = 1
        
        Set frmB = New frmBasico2
        AyudaMantenimientosAux frmB, Titulo, Desc, Text1(7), "sdirec.codclien=" & Text1(6)
        Set frmB = Nothing
        
        Exit Sub
        
    End If
'--
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = selElem
'        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco Text1(kCampo)
''            End If
'        End If
'    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim Aux As String

    


    cadB = ObtenerBusqueda(Me, False)
    
    If InstalacionEsEulerTaxco Then
        Aux = BuscaEnBDFicha
        If Aux <> "" Then
            If cadB <> "" Then cadB = cadB & " AND "
            cadB = cadB & Aux
        End If
    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla & ".", vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
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
Dim devuelve As String
'Dim TieneMan As String

    On Error GoTo EPonerCampos
    
    
    If Data1.Recordset.EOF Then Exit Sub
    
    'Por si acaso, como puede ser NULL
    Combo1.ListIndex = -1
    
    PonerCamposForma Me, Data1 'Los Text1
            
    'Poner el nombre del cod. Articulo
    'Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "sartic", "nomartic")
    'Poner el nombre del Trabajador (Operador)
    'Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "straba", "nomtraba")
    'Poner el nombre del cod. Cliente
    'Text2(6).Text = PonerNombreDeCod(Text1(6), conAri, "sclien", "nomclien")
    
    PonerDatosCodigoDescripcion 1
    PonerDatosCodigoDescripcion 5
   ' PonerDatosCodigoDescripcion 6  Nomclien: Va en la BAse de datos
    PonerDatosCodigoDescripcion 21
    PonerDatosCodigoDescripcion 23
    PonerDatosCodigoDescripcion 24
    
    
    
    If InstalacionEsEulerTaxco Then PonerCamposFicha
    
    'PonerDatosCliente Text1(6).Text
    
    If EsHistorico Then
        'Poner datos Albaran
        Text2(15).Text = DBLet(Me.Data1.Recordset!Numalbar, "T")
        FormateaCampo Text2(15)
        Text2(14).Text = DBLet(Me.Data1.Recordset!FechaAlb, "F")
        Text2(0).Text = DBLet(Me.Data1.Recordset!codtipom, "T")
        
        PonerDatosCodigoDescripcion 37
    End If
        
    
    'Poner el nombre del cod. Direc./Dpto
    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
    Text2(7).Text = devuelve
    If Not EsHistorico Then
'        Dim TieneMan As String
'
'        'Poner la fecha fin Garantia y ult. repar
'        devuelve = "ultrepar"
'        Text2(3).Text = DevuelveDesdeBDNew(conAri, "sserie", "fingaran", "numserie", Text1(0).Text, "T", devuelve, "codartic", Text1(1).Text, "T")
'        If Text2(3).Text = "" Then devuelve = ""
'        Text2(2).Text = devuelve
'        'Poner el num mantenimiento
'        TieneMan = "tieneman"
'        Text2(4).Text = DevuelveDesdeBDNew(conAri, "sserie", "nummante", "numserie", Text1(0).Text, "T", TieneMan, "codartic", Text1(1).Text, "T")
'        If TieneMan = "tieneman" Then TieneMan = "0"
'        If TieneMan = "0" Then
'            Text2(4).Text = ""
'        Else
'            If Text2(4).Text = "" Then Text2(4).Text = "SIN ESPC."
'        End If

        PonerDatosNumSerie

    Else
        If IsNull(Data1.Recordset!ultrepar) Then
            devuelve = ""
        Else
            devuelve = DBLet(Data1.Recordset!ultrepar, "F")
        End If
        Text2(2).Text = devuelve
        If IsNull(Data1.Recordset!fingaran) Then
            devuelve = ""
        Else
            devuelve = DBLet(Data1.Recordset!fingaran, "F")
        End If
        Text2(3).Text = devuelve
       
        If IsNull(Data1.Recordset!nummante) Then
            devuelve = ""
        Else
            devuelve = DBLet(Data1.Recordset!nummante, "T")
        End If
        Text2(4).Text = devuelve
        
       
    End If
    'Poner la descripcion del Motivo Pendiente Reparac.
    Text2(11).Text = PonerNombreDeCod(Text1(11), conAri, "smotre", "nommotre")
        
        
        
    'Mostraremos SOLO el numero de aviso, no la fecha de donde venia
    'Marzo Ahora ya tiene solo el numaviso
    'If Me.Text1(15).Text <> "" Then Text1(15).Text = RecuperaValor(Text1(15).Text, 1)
    
    
    If ControlReparacionAjustado Then
        'Cargamos el DATA
        CargaGrid DataGrid1, Data2, True
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu   'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
    
    
    
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerDatosNumSerie()
Dim R As ADODB.Recordset
Dim cad As String
On Error GoTo EP
        
        
       
        
        Text2(3).Text = ""
        Text2(2).Text = ""
        Text2(4).Text = ""
        Text2(6).Text = ""
        cad = "Select fingaran,ultrepar,nummante,tieneman,fechabaja,codmotba from sserie WHERE numserie = " & DBSet(Text1(0).Text, "T") & " AND codartic= " & DBSet(Text1(1).Text, "T")
        Set R = New ADODB.Recordset
        R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        If Not R.EOF Then
            'Cad = "ultrepar"
            Text2(3).Text = DBLet(R!fingaran, "F")
            Text2(2).Text = DBLet(R!ultrepar, "F")
            Text2(4).Text = DBLet(R!nummante, "N")
            cad = DBLet(R!TieneMan, "N") '"tieneman" Then Cad = "0"
            If cad = "0" Then
                Text2(4).Text = ""
            Else
                If Text2(4).Text = "" Then Text2(4).Text = "SIN ESPC."
            End If
            cad = ""
            If Not IsNull(R!fechabaja) Then
                If R!fechabaja <> "0000-00-00" Then
                    cad = DBLet(R!codmotba, "N")
                    Text2(6).Text = Format(R!fechabaja, "dd/mm/yy")
                End If
                    
            End If
            
            If cad <> "" Then
                    R.Close
                    cad = "Select * from smotba where codmotiv=" & cad
                    R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    cad = "Sin espec."
                    If Not R.EOF Then
                        If Not IsNull(R!desmotiv) Then cad = R!desmotiv
                    End If
                    Text2(6).Text = Text2(6).Text & " " & cad
            End If
        End If
        R.Close
            

EP:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set R = Nothing

End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(numrepar=" & Val(Text1(2).Text) & ")"
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Sub LimpiarDatosCliente()
Dim I As Byte

    For I = 28 To 34
        Text1(I).Text = ""
    Next I
    Text1(6).Text = ""
    Text1(7).Text = ""
    
    
    'If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(6)
End Sub


Private Function ArticulosDelNSerie(numSerie As String) As Integer
'Recupera si para ese numero de Serie hay varios articulos que lo tienen
'RETURN -> N� de articulos diferentes que tienen ese numserie
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error Resume Next

    SQL = "select distinct count(codartic) FROM sserie "
    SQL = SQL & "WHERE numserie='" & numSerie & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ArticulosDelNSerie = Rs.Fields(0).Value
    Else
        ArticulosDelNSerie = 0
    End If
    Rs.Close
    Set Rs = Nothing
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub CargarDatosNSerie(numSerie As String, codArtic As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim EnGarantia As Boolean

    SQL = "Select codclien, coddirec, tieneman, nummante, ultrepar, fingaran "
    SQL = SQL & "FROM sserie WHERE numserie=" & DBSet(numSerie, "T") & " and "
    SQL = SQL & " codartic=" & DBSet(codArtic, "T")

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
    
        'Si viene del formulario de AVISO
        'y estamos insertando
        If Modo = 3 Then
        
            If EntradaEquipo <> "" Then
                'Los datos del cliente me viene de reparacion
                If IsNull(Rs!codClien) Then
                    Rs.Close
                    Set Rs = Nothing
                    Exit Sub
                End If
            
                SQL = RecuperaValor(EntradaEquipo, 3)
                If Val(Rs!codClien) <> Val(SQL) Then
                    MensajeNoCoinciden CStr(Val(Rs!codClien)), False
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    CadenaDesdeOtroForm = ""
                End If
            End If
        End If
        
        Text1(6).Text = Format(Rs!codClien, "000000")
        Text1(7).Text = Format(DBLet(Rs!CodDirec), "000")
        If Text1(7).Text <> "" Then Text2(7).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
        
        
        
        Text2(2).Text = DBLet(Rs!ultrepar, "F")
        Text2(3).Text = DBLet(Rs!fingaran, "F")
        
        'Poner fecha prevista reparacion en funcion del param. de la aplicacion (diassiman,diasnoman)
        'dependiendo de si el numserie,codartic tiene mantenimiento (ver tabla sserie)
        If Rs!TieneMan = "1" Then
            Text2(4).Text = DBLet(Rs!nummante, "T")
            If Text2(4).Text = "" Then Text2(4).Text = "SIN ESPEC."
        End If
        PonerFechaRepar (Val(DBLet(Rs!TieneMan, "N")) = 1)
        'Cargar los datos del Cliente
        
        PonerDatosCliente2 (Text1(6).Text), True
        If Text1(6).Text = "" Then
            'NO esta con codigo de cliente
            FrameClientes.Enabled = True
        End If
        
        
        
        
        
        
        'Junio 2015
        'MSG de equipo en garantia, o no
        If Modo = 3 Then
            EnGarantia = False
            If Text2(3).Text <> "" Then
                If CDate(Text2(3).Text) >= CDate(Now) Then EnGarantia = True
            End If
        
            If EnGarantia Then
                
                SQL = "El equipo esta en garant�a. Finaliza el " & Text2(3).Text
                MsgBox SQL, vbInformation
            Else
                'NO esta en garantia
                SQL = String(70, "*") & vbCrLf
                SQL = SQL & vbCrLf & "El equipo NO esta en garant�a. Finaliz� el " & Text2(3).Text & vbCrLf & vbCrLf & SQL
                MsgBox SQL, vbExclamation
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub PonerFechaRepar(TieneManteinimiento As Boolean)
Dim F As Date
Dim N As Byte

    F = Now
    If TieneManteinimiento Then
        F = F + vParamAplic.DiasSiMante
    Else
        F = F + vParamAplic.DiasNoMante
    End If
    N = Weekday(F, vbMonday)
    If N = 6 Then
        N = 2
    Else
        If N = 7 Then
            N = 1
        Else
            N = 0
        End If
    End If
    If N > 0 Then F = F + N
    
    Text1(4).Text = Format(F, "dd/mm/yyyy")

End Sub
Private Sub PonerDatosCliente2(codClien As String, Optional nifClien As String)
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
            If vCliente.ClienteBloqueado(2, False) Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
            End If
                    
                           
            HabilitarDatosCliente vCliente.DeVarios
            
            If Modo = 4 Then
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(6).Text) = CLng(Data1.Recordset!codClien) Then
           '         If Text2(6).Text = Data1.Recordset!nomclien Then
                        Set vCliente = Nothing
                        Exit Sub
           '         End If
                End If
            End If
            Text1(34).Text = vCliente.Nombre
            If Not vCliente.DeVarios Then
                Text1(28).Text = vCliente.NIF
                Text1(29).Text = vCliente.TfnoClien
                Text1(30).Text = vCliente.Domicilio
                Text1(31).Text = vCliente.CPostal
                Text1(32).Text = vCliente.Poblacion
                Text1(33).Text = vCliente.Provincia
                Text1(34).Text = vCliente.Nombre
            End If
            Text1(6).Text = vCliente.Codigo
            FormateaCampo Text1(6)
            

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" And (Modo = 3 Or Modo = 4) Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                           
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(3).Text
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Function InsertarCabecera() As Boolean
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    On Error GoTo EInsertarCab
    InsertarCabecera = False
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(2).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarRepar(SQL, vTipoMov) Then
            
                If InstalacionEsEulerTaxco Then
                    
                    'Hay que a�adir, aunque sea vacio en scarepeu  que lleva los datos de la ficha tecnica
                    ActualizaBDFicha
                End If
            
            
            
            
                InsertarCabecera = True
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE numrepar=" & Text1(2).Text '& ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
'                PonerModo 2
'                PosicionarData
            End If
        End If
        Text1(2).Text = Format(Text1(2).Text, "0000000")
    End If
    
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function InsertarRepar(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertar
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numrepar", "numrepar", Text1(2).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(2).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla de Reparaciones (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
    MenError = "Error al actualizar el contador del Pedido."
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertar:
    If Err.Number <> 0 Then
        MenError = "Insertando Reparaci�n." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarRepar = True
    Else
        conn.RollbackTrans
        InsertarRepar = False
    End If
End Function


Private Function ObtenerWhereCP() As String
Dim SQL As String

    SQL = " WHERE  numrepar= " & Text1(2).Text
    ObtenerWhereCP = SQL
End Function


Private Sub BotonImprimir2(LaReparacion As Boolean, NumeroAlbaran As Long)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim devuelve As String


    cadFormula = ""
    cadParam = ""
    numParam = 0




    If LaReparacion Then

                If Text1(2).Text = "" Then 'N� Reparacion
                    MsgBox "Debe seleccionar una Reparaci�n para Imprimir.", vbInformation
                    Exit Sub
                End If
                
                
                '===================================================
                '============ PARAMETROS ===========================
                'A�adir el parametro de Empresa
                cadParam = cadParam & "|pEmpresa=""" & UCase(vEmpresa.nomempre) & """|"
                numParam = numParam + 1
            
                'A�adir el parametro con el N� de mantenimiento si hay
                If Trim(Text2(4).Text) <> "" Then
                    cadParam = cadParam & "pMantenimiento=""" & Text2(4).Text & """|"
                    numParam = numParam + 1
                End If
                  
                'A�adir el parametro si esta en garantia o no
                If Trim(Text2(3).Text) <> "" Then
                    If Format(Now, "dd/mm/yyyy") > Format(Text2(3).Text, "dd/mm/yyyy") Then
                        cadParam = cadParam & "pGarantia=""NO""|"
                    Else
                        cadParam = cadParam & "pGarantia=""SI""|"
                    End If
                    numParam = numParam + 1
                End If
                  
                'Nombre fichero .rpt a Imprimir
                If Not PonerParamRPT2(24, cadParam, numParam, devuelve, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
                    Exit Sub
                End If
            
                'Nombre fichero .rpt a Imprimir
                frmImprimir.NombreRPT = devuelve
                frmImprimir.NombrePDF = pPdfRpt
                frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
                'frmImprimir.NombreRPT = "rRepResguardo.rpt"
                devuelve = ""
                    
                '===================================================
                '================= FORMULA =========================
                'Cadena para seleccion N� de Reparacion
                '---------------------------------------------------
                If Text1(2).Text <> "" Then
                    'N� Reparacion
                    devuelve = "{" & NombreTabla & ".numrepar}=" & Val(Text1(2).Text)
                    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                End If
                
                 With frmImprimir
                    .FormulaSeleccion = cadFormula
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    .SoloImprimir = False
                    .EnvioEMail = False
                    .Opcion = 62
                    .Titulo = "Resguardo Reparaci�n"
                    .Show vbModal
                End With



    Else
        'El albaran generado anteriormente
            
            '===================================================
            '============ PARAMETROS ===========================
            
            If Not PonerParamRPT2(36, cadParam, numParam, devuelve, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
           
                   
            'Nombre fichero .rpt a Imprimir
            
                frmImprimir.NombreRPT = devuelve
                frmImprimir.NombrePDF = pPdfRpt
                frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
           
            'A�adir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
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
            devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(6).Text, "N")
            If devuelve = "" Then devuelve = "0"
            ' 0 "Todo"
            ' 1 "Cantidad y Precio"
            ' 2 "Cantidad"
            cadParam = cadParam & "Albarcon=" & devuelve & "|"
            numParam = numParam + 1
            
    
            
                
                
            '===================================================
            '================= FORMULA =========================
            'Cadena para seleccion N� de Albaran
            '---------------------------------------------------
                'Cod Tipo Movimiento
                devuelve = "{scaalb.codtipom}='ALR'"
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                'N� Albaran
                devuelve = "{scaalb.numalbar}=" & NumeroAlbaran
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                
                
           
            '=========================================================================
            'Aqui sabemos que valor tiene CodClien y a�adimos a los parametros el tipo de IVA
            'que se aplica a ese cliente
            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(6).Text, "N")
            If devuelve <> "" Then
                cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
                numParam = numParam + 1
            End If
        
                
            '==============================================================
'            'Comprobar si hay registros a Mostrar antes de abrir el Informe
'            devuelve = " scaalb INNER JOIN slialb ON "
'            devuelve = devuelve & "scaalb.codtipom=slialb.codtipom AND scaalb.numalbar= slialb.numalbar "
'            If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
'
            
                With frmImprimir
                    'Febrero 2010
                    'Albaran. Tiene su numero de albara
                        .outTipoDocumento = 4
                        .outClaveNombreArchiv = "ALB" & Format(NumeroAlbaran, "000000")
                        .outCodigoCliProv = CLng(Text1(6).Text)
            
                    
                    .FormulaSeleccion = cadFormula
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    .SoloImprimir = False
                    .EnvioEMail = False
                    .Opcion = 45  'Impresion albaranes
                    .Titulo = "Albaran de Cliente"
                    .ConSubInforme = True
                    .Show vbModal
                End With
        
    End If
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "L�neas Reparaciones"
        PonerFocoBtn Me.cmdRegresar
    Else
        cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu seg�n Modo
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        
        cmdAux(2).visible = visible And vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                If I < 3 Then
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                Else
                    txtAux(I).Text = DataGrid1.Columns(I + 3).Text
                End If
                BloquearTxt txtAux(I), False
            Next I
            'El campo Nom Artic lo bloqueamos inicialmente
            BloquearTxt txtAux(2), True
        End If
            
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 22)
        alto = alto '+ SSTab1.Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(2).Height = DataGrid1.RowHeight
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330 '+ SSTab1.Left
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 180
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 30
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio
        
        For I = 4 To 7
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 3).Width - 10
        Next I
        
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To 7
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(2).visible = False
        If visible Then
            If vEmpresa.TieneAnalitica Then
                I = 8
                txtAux(I).visible = True
                txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
                txtAux(I).Width = DataGrid1.Columns(I + 3).Width - 10
                txtAux(I).Locked = True
                If vParamAplic.ModoAnalitica = 2 Then
                    cmdAux(2).Top = cmdAux(1).Top
                    cmdAux(2).visible = True
                    txtAux(I).Locked = False
                    cmdAux(2).Left = txtAux(I).Left + txtAux(I).Width - 90
                End If
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadirLinea()
Dim CC As String
Dim Aux As String

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
'--
'    PonerBotonCabecera False

    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    
    
    'Veo primero el trabajador conectado
    Aux = PonerTrabajadorConectado(CC)
    If Aux = "" Then Aux = Text1(5).Text
    'Poner el Almacen por defecto del Trabajador
    CC = "codccost"
    'El trabajador conectado NO tiene en trabajadores
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Aux, "N", CC)
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    
    'Si la analitica es por
    If vEmpresa.TieneAnalitica Then
        If vParamAplic.ModoAnalitica = 0 Then txtAux(8).Text = CC
    End If
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    PrimeraVez = True
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    vDataGrid.ScrollBars = dbgAutomatic

    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte
    On Error GoTo ECargaGrid

    vData.Refresh
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    
        'Cod. Almacen
        vDataGrid.Columns(2).Caption = "Alm."

        vDataGrid.Columns(2).Width = 600

        vDataGrid.Columns(2).NumberFormat = "000"
        
        vDataGrid.Columns(3).Caption = "Art�culo"

        vDataGrid.Columns(3).Width = 3000

        
        vDataGrid.Columns(4).Caption = "Descripci�n"
        vDataGrid.Columns(4).Width = 4200

        vDataGrid.Columns(5).visible = False
        
        vDataGrid.Columns(6).Caption = "Cantidad"
        vDataGrid.Columns(6).Width = 1350
        vDataGrid.Columns(6).Alignment = dbgRight
        vDataGrid.Columns(6).NumberFormat = FormatoImporte
        
        I = 7
        vDataGrid.Columns(I).Caption = "Precio"
        vDataGrid.Columns(I).Width = 1450
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoPrecio
        
            
        I = I + 1
        vDataGrid.Columns(I).Caption = "Dto.1"
        vDataGrid.Columns(I).Width = 800
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
        I = I + 1
        vDataGrid.Columns(I).Caption = "Dto.2"
        vDataGrid.Columns(I).Width = 800
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoDescuento
    
        I = I + 1
        vDataGrid.Columns(I).Caption = "Importe L�nea"
'        If conServidas Then
'            vDataGrid.Columns(i).Width = 1250
'        Else
            vDataGrid.Columns(I).Width = 1800
'        End If
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoImporte
    
        If Not vEmpresa.TieneAnalitica Then
            
            'Hay 1 columnas menos
        Else
            I = I + 1
            vDataGrid.Columns(I).Caption = "C.Cos"
            vDataGrid.Columns(I).Width = 800
        End If
    
        For I = 0 To vDataGrid.Columns.Count - 1
            vDataGrid.Columns(I).Locked = True
            vDataGrid.Columns(I).AllowSizing = False
        Next I
        vDataGrid.RowHeight = 350
        vDataGrid.HoldFields
        Exit Sub
        
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub



Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numrepar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    'If conServidas Then SQL = SQL & "servidas, "
    SQL = SQL & "precioar, dtoline1, dtoline2,importel "
    If vEmpresa.TieneAnalitica Then SQL = SQL & ",codccost"
        
        
    
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP
    Else
        SQL = SQL & " WHERE numrepar = -1"
    End If
    SQL = SQL & " Order by numrepar, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: BotonConfirmarRep 'Confirmar Reparaci�n
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub 'en almacen y flecha h. arriba
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
'Dim codTarif As String
Dim vCStock As CStock
Dim CPrecioFact As CPreciosFact
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean

    'Quitar espacios en blanco
    txtAux(Index).Text = Trim(txtAux(Index))
    
    If txtAux(Index).Text = "" And (Index <> 1 And Index <> 8) Then Exit Sub
    
    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
     Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)
            
        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then 'Cod Artic
                txtAux(2).Text = "" 'Nom Artic
                Exit Sub
            End If
            
            If txtAux(0).Text = "" Then 'Cod Almacen
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            
            If Not PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, False, cadMen) Then

                PonerFoco txtAux(Index)
            Else
                
                If vEmpresa.TieneAnalitica Then
                    If vParamAplic.ModoAnalitica = 1 Then '0=trabajador, 1=Familia, 2=Proyecto
                        txtAux(8).Text = cadMen
                        cadMen = ""
                    End If
                End If
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            End If
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub '"S"=Salida de Stock
                vCStock.MoverStock False, False
                If Not PrimeraVez Then Exit Sub
                PrimeraVez = False
                If (Modo = 5 And ModificaLineas = 1) Then 'Modo Insertar en Mto Lineas
                    'Ver si esta en Garantia el Aparato
                    'Si el Articulo esta en garantia pregunta si se facturara la linea o no
                    'Si facturar -> precioar=Precio
                    'Si no facturar -> precioar=0
                    If Text2(3).Text <> "" Then   'a�adido el 18 de diciembre
                        If EsFechaPosterior(Text1(3).Text, Text2(3).Text, False) Then
                           If MsgBox("El aparato esta en Garant�a.�Facturar la linea de Reparacion?", vbYesNo) = vbNo Then
                                txtAux(4).Text = "0,00"
                                txtAux(5).Text = "0,00"
                                txtAux(6).Text = "0,00"
                                Set vCStock = Nothing
                                Exit Sub
                           End If
                        End If
                    End If
                    'Si el aparato tiene Mantenimiento no se cobra la linea de Reparaci�n? Preguntar
                    If Text2(4).Text <> "" Then
                        If MsgBox("El aparato tiene Mantenimiento.�Facturar la linea de Reparaci�n?", vbYesNo) = vbNo Then
                            txtAux(4).Text = "0,00"
                            txtAux(5).Text = "0,00"
                            txtAux(6).Text = "0,00"
                            Set vCStock = Nothing
                            Exit Sub
                        End If
                    End If
                        
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la funci�n
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If devuelve <> "" Then
                        Set CPrecioFact = New CPreciosFact
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        NumCajas = CPrecioFact.ObtenerNumCajas(vCStock.cantidad, devuelve)
                        RestoUnid = CInt(vCStock.cantidad) - NumCajas * CInt(devuelve)
                        'Obtenemos la Tarifa del Cliente
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(6).Text, "N")
                        'CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.CodigoArtic = vCStock.codArtic
                        CPrecioFact.CodigoClien = Text1(6).Text
                        CPrecioFact.FijarTarifaActividad
                        PorCaja = (NumCajas > 0)
                        ' ---- [10/11/2009] [LAURA] : pasamos el codartic en el lugar de una fecha
'                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(4).Text, OrigP, "")
                        ' ----
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Art�culo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(vCStock.cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
                            PonerFoco txtAux(Index)
                        Else
                            If txtAux(4).Text = "" Then
                                txtAux(4).Text = Precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(5).Text = "" Then txtAux(5).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(5), 4
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(6), 4
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
        Case 8
            'CC
             If vEmpresa.TieneAnalitica Then Me.Text2(8).Text = PonerNombreCCoste(Me.txtAux(8))
            
    End Select
    
    If Modo = 5 Then 'Modo Lineas
        If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then 'Cant., Precio, dto1, dto2
            If txtAux(1).Text = "" Then Exit Sub 'Cod artic
            txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
            PonerFormatoDecimal txtAux(7), 1
            'If Index = 6 Then PonerFocoBtn cmdAceptar
        End If
    End If
End Sub


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(6).Text) 'guardamos el cliente
    vCStock.Documento = Text1(2).Text 'N� Albaran
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text)) - Data2.Recordset!cantidad
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(7).Text))
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


Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slirep
Dim SQL As String
Dim numlinea As String, vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Mid(ObtenerWhereCP, 7)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        'Construir la sentencia SQL
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numrepar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) "
        SQL = SQL & "VALUES (" & DBSet(Text1(2).Text, "N") & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N", "N") & ", " 'cantidad
        SQL = SQL & DBSet(txtAux(4).Text, "N", "N") & ", " 'precio
        SQL = SQL & DBSet(txtAux(5).Text, "N", "N") & ", " 'Dto1
        SQL = SQL & DBSet(txtAux(6).Text, "N", "N") & ", " ' Dto2
        SQL = SQL & DBSet(txtAux(7).Text, "N", "N") & ","
        If vEmpresa.TieneAnalitica Then
            SQL = SQL & DBSet(txtAux(8).Text, "T", "N")
        Else
            SQL = SQL & "NULL"
        End If
        SQL = SQL & ")" 'Importe linea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Reparaci�n" & vbCrLf & Err.Description
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim I As Byte
Dim vArtic As CArticulo
Dim Mal As Boolean

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True


    
    
    'Comprobar que los campos NOT NULL tienen valor
    For I = 0 To txtAux.Count - 1
        
        If txtAux(I).Text = "" Then
            If I = 5 Or I = 6 Then
                'LOS DESCUENTOS
                'Si los descuentos estan a blancos, pinto el cero yo
                txtAux(I).Text = "0"
            Else
                Mal = False
                If I = 8 Then
                    If vEmpresa.TieneAnalitica Then Mal = True
                Else
                    Mal = True
                End If
                
                If Mal Then
                    b = False
                    MsgBox "Campo " & txtAux(I).Tag & " no puede estar vacio", vbExclamation
                    PonerFoco txtAux(I)
                    Exit Function
                End If
            End If
        End If
    Next I
   
        
    'Comprobar que existe el articulo en el almacen seleccionado
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


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Reparaciones: slirep
Dim SQL As String
Dim vCStock As CStock
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function

    If Not DatosOkLinea() Then
        Set vCStock = Nothing
        Exit Function
    End If
    
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "cantidad = " & DBSet(txtAux(3).Text, "N", "N") & ", "
        SQL = SQL & "precioar = " & DBSet(txtAux(4).Text, "N", "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N", "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N", "N") & ", "
        SQL = SQL & "importel=" & DBSet(txtAux(7).Text, "N", "N") & " "
        If vEmpresa.TieneAnalitica Then SQL = SQL & ", codccost=" & DBSet(txtAux(8).Text, "T", "N") & " "
        SQL = SQL & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea

        If SQL <> "" Then
            conn.BeginTrans
            conn.Execute SQL
            vCStock.cantidad = CSng(txtAux(3).Text)
            b = vCStock.ModificarStock(Data2.Recordset!cantidad)
        End If
    
    Set vCStock = Nothing
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Reparaci�n" & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarLinea = b
End Function


Private Function SePuedeServirPedido(Optional cadErr As String) As Boolean
'Comprobar Si se puede servir la Reparacion solicitada y pasar a albaran
Dim vCStock As CStock
Dim SQL As String
Dim b As Boolean
Dim Rs As ADODB.Recordset

    On Error GoTo EServir

    SePuedeServirPedido = False
    'Verificar si hay stock para aquellas familias que no son instalacion
    Set vCStock = New CStock
    
    SQL = "SELECT codalmac, codartic, SUM(cantidad) as cantidad from " & NomTablaLineas
    SQL = SQL & ObtenerWhereCP
    SQL = SQL & " GROUP by codalmac, codartic"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'Si no hay lineas para pasar al albaran no seguimos
    If Rs.EOF Then
        cadErr = "No hay lineas para generar el Albaran."
        b = False
        GoTo EServir
    End If
    
    'para cada linea de la Reparacion comprobar el stock si no es instalacion
    b = True
    While (Not Rs.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", , Rs) Then
            cadErr = "No se pudo inicializar la clase Stock"
            b = False
            GoTo EServir
'            Exit Function
        End If
        'Comprobar si se puede mover stock (hay stock, o si no hay pero no control de stock)
        cadErr = ""
        If vCStock.MueveStock Then
            If Not vCStock.MoverStock(False, False, True) Then b = False
        End If
        Rs.MoveNext
    Wend
    Set vCStock = Nothing
    Rs.Close
    Set Rs = Nothing
    SePuedeServirPedido = b
    
EServir:
    If Err.Number <> 0 Then
        b = False
        Set vCStock = Nothing
        Rs.Close
        Set Rs = Nothing
    End If
    
    SePuedeServirPedido = b
End Function


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef Rs As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Reparacion a Albaran de Reparacion
On Error Resume Next
    
    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALR"
    vCStock.Trabajador = CLng(Text1(6).Text) 'guardamos el cliente
    vCStock.Documento = Text1(2).Text
    vCStock.codArtic = Rs!codArtic
    vCStock.codAlmac = CInt(Rs!codAlmac)
    
    vCStock.cantidad = CSng(Rs!cantidad)
    'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
    If Rs.Fields.Count > 3 Then vCStock.Importe = CCur(Rs!ImporteL)
    
    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Sub GenerarAlbaran()
Dim numRep As Long 'N� Reparacion
Dim NumAlb As Long 'N� Albaran
Dim b As Boolean

    'Pedir: Operador de Albaran, Material Preparado por y forma de envio
    DoEvents
    Screen.MousePointer = vbHourglass
    If Bloquearmanualmente Then
        CadenaSQL = ""
        Set frmList = New frmListadoPed
        frmList.NumCod = CodTipoMov
        frmList.OpcionListado = 43
        'Para k no me muestre los checks
        frmList.chkImpEtiq.visible = False
        frmList.chkImpAlbaran.visible = True
         frmList.chkImpAlbaran.Value = 1
        frmList.chkImpHojaExped.visible = False
    
        
        frmList.Show vbModal
        
        Set frmList = Nothing
        b = False
        If CadenaSQL <> "" Then
            NumRegElim = Data1.Recordset.AbsolutePosition
            numRep = Data1.Recordset!numrepar
            b = PasarPedidoAAlbaran(CadenaSQL, NumAlb)
        End If
        
        DesBloqueoManual "GENALBREP"
        Screen.MousePointer = vbDefault
        If b Then
            If ImprimeAlb Then
                BotonImprimir2 False, NumAlb
            Else
                MsgBox "La Reparaci�n N�: " & Format(numRep, "0000000") & " ha generado " & vbCrLf & vbCrLf & "el Albaran de Reparaci�n N�: " & Format(NumAlb, "0000000"), vbInformation
            End If
            PonerModo 2
            'Se habra eliminado el pedido de (scarep, slirep)
            PosicionarDataTrasEliminar
        End If
        
    
        
    Else
        MsgBox "Proceso bloqueado por otro usuario", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        If ControlReparacionAjustado Then
            'Cargamos el DATA
            CargaGrid DataGrid1, Data2, False
        End If
        PonerModo 0
    End If
End Sub


Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As Long) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: N� de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
    
    On Error GoTo EGenPedido

    bol = False
    If vSQL = "" Then Exit Function
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (scaalb, slialb)
    MenError = "Insertando el tablas de albaranes. (scaalb,slialb)"
    bol = InsertarAlbaran(vSQL, MenError, NumAlb)
    
    'Actualizar Stock en salmac, e introducir movimiento en smoval
    If bol Then
        MenError = "Actualizando movimientos de stock."
        bol = InsertarMovStock(NumAlb)
    End If
    
    If bol Then
        MenError = "Pasando al hist�rico de reparaciones."
        'Pasar al Historico de Reparaciones: schrep
        bol = InsertarCabeceraHcoRep(NumAlb)
        If bol Then
            ActualizarFechasElto
         
        'Borrar la Reparacion de las tablas de Reparaciones (scarep, slirep)
            MenError = "Eliminando en tablas de reparaciones.(scarep,slirep)"
            If bol Then bol = Eliminar()
        End If
    End If
    
    
    
    'Si correcto y tiene numnero de aviso, cierro el aviso
    If bol Then
        If Text1(15).Text <> "" Then
            'LLEVA aviso de REPARCION
            
            MenError = "Actualizando avisos."
            CadenaDesdeOtroForm = "UPDATE scaavi SET situacio = 3"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE numaviso =" & Data1.Recordset!numaviso
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND  fechaavi = '" & Format(Data1.Recordset!fecaviso, FormatoFecha) & "'"
            conn.Execute CadenaDesdeOtroForm
        End If

    End If
    
EGenPedido:
    If Err.Number <> 0 Then bol = False
    
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MenError = "Pasando Reparaci�n a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        CadenaDesdeOtroForm = ""
    End If
    PasarPedidoAAlbaran = bol
End Function


Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As Long) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String, SQL As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de ALBARAN de Reparacion
    codtipom = "ALR"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            NumAlb = vTipoMov.ConseguirContador(codtipom)
            devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", codtipom, "T", , "numalbar", CStr(NumAlb), "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codtipom)
                NumAlb = vTipoMov.ConseguirContador(codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    'Acabar la sql con el contador seleccionado
    SQL = "INSERT INTO scaalb (codtipom, numalbar, fechaalb, factursn, codclien, nomclien, domclien, codpobla, pobclien, proclien, "
    SQL = SQL & "nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codtrab1, codtrab2, codagent, codforpa, codenvio, "
    SQL = SQL & "dtoppago, dtognral, tipofact, observa01, observa02, observa03, observa04, observa05, numofert, fecofert, numpedcl, fecpedcl, sementre) "
    SQL = SQL & " VALUES ('" & codtipom & "', " & NumAlb & "," & vSQL & ")"
    
    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
    conn.Execute SQL, , adCmdText
    
    'Insertar Lineas de Albaran
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialb)."
    If Not InsertarLineasAlbaran(codtipom, NumAlb) Then Exit Function
    
    MenError = "Error al actualizar el contador del Albaran."
    vTipoMov.IncrementarContador (codtipom)
    Set vTipoMov = Nothing
    bol = True
    
EInsertarAlbaran:
    If Err.Number <> 0 Then bol = False
    InsertarAlbaran = bol
End Function


Private Function InsertarMovStock(NumAlb As Long) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo EInsMov

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    SQL = "select * from " & NomTablaLineas & ObtenerWhereCP
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not Rs.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", CStr(Rs!numlinea), Rs) Then Exit Function
        vCStock.Documento = CStr(NumAlb)
         'en actualizar stock comprobamos si el articulo tiene control de stock
        'If vCStock.Cantidad <> 0 Then
            b = vCStock.ActualizarStock(False, False)
        Rs.MoveNext
    Wend
    Set vCStock = Nothing
    Rs.Close
    Set Rs = Nothing
    
'    InsertarMovStock = b
    
EInsMov:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando movimiento de stock.", Err.Description
        b = False
    End If
    InsertarMovStock = b
End Function



Private Function InsertarLineasAlbaran(TipoM As String, NumAlb As Long) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim SQL As String

    On Error GoTo EInsertarLin

    'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
    'Cambio por el rollo de la trazabilidad
    SQL = ""
    SQL = "SELECT '" & TipoM & "' as codtipom, " & NumAlb & " as numalbar, numlinea, codalmac,  s.codartic, s.nomartic , ampliaci, "
    '    numbultos
    SQL = SQL & "cantidad,0, precioar, dtoline1, dtoline2, importel, '' as origpre,"
    '  codprove, numlote NULL
    SQL = SQL & "codprove,NULL,"
    If vEmpresa.TieneAnalitica Then
        SQL = SQL & "codccost"
    Else
        SQL = SQL & "NULL"
    End If
    SQL = SQL & " FROM " & NomTablaLineas & " s,sartic WHERE s.codartic=sartic.codartic AND numrepar=" & Text1(2).Text
    
    'Pongo los campos que voy a insertar
    SQL = "(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codproveX,numlote,codccost)" & SQL
    'NO pongo valores para los nuevos campos de sail: codtipor codcapit precoste codtraba
    SQL = "INSERT INTO slialb " & SQL
    conn.Execute SQL
    InsertarLineasAlbaran = True

EInsertarLin:
'    If Err.Number <> 0 Then
'        InsertarLineasAlbaran = False
'    Else
'        InsertarLineasAlbaran = True
'    End If
    InsertarLineasAlbaran = Not (Err.Number <> 0)
End Function


Private Function InsertarCabeceraHcoRep(NumAlb As Long) As Boolean
'Insertar en la Tabla Cabecera de Historico
Dim SQL As String
Dim Aux As String

    On Error GoTo eInsertarCabeceraHcoRep
    
    
    SQL = "SELECT numrepar, fecrepar,fecentre," & NombreTabla & ".numserie, " & NombreTabla & ".codartic, sartic.nomartic, "
    'fecha fin garantia: fingaran, ultrepar
    SQL = SQL & DBSet(Text2(3).Text, "F") & " as fingaran, " & DBSet(Text2(2).Text, "F", "S") & " as ultrepar, "
    SQL = SQL & "codclien, coddirec, " & DBSet(Text2(4).Text, "T") & " as nunmante, " 'nummante
    SQL = SQL & "codtraba, " & CadenaSQLHco & ", "
    SQL = SQL & "'ALR' as codtipom, " & NumAlb & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb "
    
    'Modifiaciones 1 OCTUBRE 2007
    'A�adimos SAT tipo averia y presupuestos
    Aux = ",codman,codavi,codtrabajo,imppresu1,impresu2,contestado,fecha,fechaaprob,avisocli,fecenviosat,resguardosat,importesat,fecentresat,observasat"
    'Modificacion 6 Enero 2009
    ' domclien,codpobla,pobclien,nifdatos,telclien,nomclien  Datos CLIENTE
    Aux = Aux & ",domclien,codpobla,pobclien,nifdatos,telclien,nomclien,refclien"
    SQL = SQL & Aux
    
    SQL = SQL & " FROM " & NombreTabla & " INNER JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic "
    SQL = SQL & ObtenerWhereCP
    
    SQL = "INSERT INTO schrep (numrepar,fecrepar,fecentre,numserie,codartic,nomartic,fingaran,ultrepar,codclien,coddirec,nummante,codtraba,codtrab1,codtrab2,material,tipoaver,motivore,textore1,textore2,textore3,codtipom,numalbar,fechaalb" & Aux & ") " & SQL
    conn.Execute SQL
    
    
    If InstalacionEsEulerTaxco Then
        'schrepeu scarepeu
        'Las tablas son iguales a excepcion de que el Historico lleva la primera columna, la fecha reparacion
        SQL = "INSERT INTO schrepeu Select " & DBSet(Text1(4).Text, "F") & ",scarepeu.* from scarepeu "
        SQL = SQL & ObtenerWhereCP
        conn.Execute SQL
    End If
    
eInsertarCabeceraHcoRep:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        InsertarCabeceraHcoRep = True
    End If
End Function



Private Sub CargaDatosAviso()
    On Error GoTo ECargaDatosAviso
    
    
    
    If EntradaEquipo = "" Then Exit Sub
    
    BotonAnyadir
            
    'Ahora pongo los campos de la entradequipo
    'Numero aviso
    Text1(15).Text = RecuperaValor(EntradaEquipo, 1)
    Text1(36).Text = RecuperaValor(EntradaEquipo, 2)
    
    'Cliente
    Text1(6).Text = RecuperaValor(EntradaEquipo, 3)
    Text1(34).Text = RecuperaValor(EntradaEquipo, 4)
    'NIF
    Text1(28).Text = RecuperaValor(EntradaEquipo, 7)
    'Tfno
    Text1(29).Text = RecuperaValor(EntradaEquipo, 8)
    'Domicilio
    Text1(30).Text = RecuperaValor(EntradaEquipo, 9)
    'Codpostal
    Text1(31).Text = RecuperaValor(EntradaEquipo, 10)
    'Pobla
    Text1(32).Text = RecuperaValor(EntradaEquipo, 11)
    'prov
    Text1(33).Text = RecuperaValor(EntradaEquipo, 12)
    'Dpto
    Text1(7).Text = RecuperaValor(EntradaEquipo, 5)
    Text2(7).Text = RecuperaValor(EntradaEquipo, 6)
    
    
    Exit Sub
ECargaDatosAviso:
    MuestraError Err.Number, "CargaDatosAviso"

End Sub


Private Sub MensajeNoCoinciden(Equipo As String, Pregunta As Boolean)

    CadenaDesdeOtroForm = "############"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = vbCrLf & vbCrLf & CadenaDesdeOtroForm & CadenaDesdeOtroForm & vbCrLf & vbCrLf
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " No coinciden el cliente del aviso (" & RecuperaValor(EntradaEquipo, 3) & ") con el del numero de serie (" & Equipo & ")" & CadenaDesdeOtroForm
    If Pregunta Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf & "�Continuar?"
End Sub


Private Sub HabilitarDatosCliente(Habilitar As Boolean)
Dim I As Integer

    For I = 28 To 34
        BloquearTxt Text1(I), Not Habilitar
    Next I
    imgBuscar(9).visible = Habilitar
    'Text1(28).Text = vCliente.NIF
    'Text1(29).Text = vCliente.TfnoClien
    'Text1(30).Text = vCliente.Domicilio
    'Text1(31).Text = vCliente.CPostal
    'Text1(32).Text = vCliente.Poblacion
    'Text1(33).Text = vCliente.Provincia
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    If Modo = 4 Then
        If DBLet(Data1.Recordset!nifdatos, "T") = nifClien Then Exit Sub
    End If
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    If b Then Text1(34).Text = vCliente.Nombre         'Nom clien
    
    Text1(30).Text = vCliente.Domicilio
    Text1(31).Text = vCliente.CPostal
    Text1(32).Text = vCliente.Poblacion
    Text1(33).Text = vCliente.Provincia
    Text1(29).Text = DBLet(vCliente.TfnoClien, "T")
            
'    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub AbrirNumSerie()
    Set frmNSeries2 = New frmRepNumSerie2GR
    frmNSeries2.DatosADevolverBusqueda = "O"
    frmNSeries2.DatoAInsertar = Text1(0).Text
    frmNSeries2.Show vbModal
    Set frmNSeries2 = Nothing

End Sub


Private Function Bloquearmanualmente() As Boolean
Dim T1 As Single
Dim OK As Boolean

    T1 = Timer
    Bloquearmanualmente = False
    Do
        OK = BloqueoManual("GENALBREP", "1", True)
        If Not OK Then
            If Timer - T1 > 15 Then OK = True
            Espera 1
        Else
            Bloquearmanualmente = True

        End If
    Loop Until OK
End Function


Private Sub BuscaNumserieRepetido()
Dim cad As String
        Set frmB3 = New frmBuscaGrid
            frmB3.vCampos = "N� Serie|sserie|numserie|T||20�Artic.|sserie|codartic|T||25�Desc. Artic.|sartic|nomartic|T||40�"
            
            cad = "sserie LEFT JOIN sartic ON sserie.codartic=sartic.codartic"
            frmB3.vTabla = cad
            frmB3.vBusqueda = " sserie.numserie = '" & DevNombreSQL(Text1(0).Text) & "'"
            frmB3.vTitulo = "N� Serie"
            frmB3.vselElem = 1
            frmB3.vCargaFrame = False
            frmB3.vConexionGrid = 1
            frmB3.vDevuelve = "1|2|"
            frmB3.Show vbModal
            Set frmB3 = Nothing
End Sub


Private Sub ActualizarFechasElto()
Dim C2 As String

    'FALTA### que lea los dias de garantia desde paremtros
    
    C2 = DateAdd("d", vParamAplic.DiasGarantia, CDate(FechaAlb))
    C2 = "'" & Format(C2, FormatoFecha) & "'"
    C2 = "UPDATE sserie SET fingaran = " & C2
    'Ultima feha reparacion
    C2 = C2 & ", ultrepar = '" & Format(FechaAlb, FormatoFecha) & "'"
    C2 = C2 & " WHERE numserie = " & DBSet(Text1(0).Text, "T")
    C2 = C2 & " AND codartic = " & DBSet(Text1(1).Text, "T")
    ejecutar C2, False
End Sub


Private Sub BloquearPorNumeroSerie(Bloquear As Boolean)
        Text1(1).Enabled = Not Bloquear
        Me.FrameClientes.Enabled = Not Bloquear
End Sub



Private Sub LimpiarFichaTecnica(SinTxts As Boolean)
Dim N As Byte
    
    
    If Not SinTxts Then
        For N = 0 To Me.txtEuler.Count - 1
            txtEuler(N).Text = ""
        Next
    End If
    
    For N = 0 To chkEuler.Count - 1
        chkEuler(N).Value = 0
    Next
    
    Me.optEuler(0).Value = True
    Me.optEuler(0).Value = False  'Ninguno seleccionado
    
    Me.optEuler(2).Value = True
    Me.optEuler(2).Value = False
    
    Me.optEuler(7).Value = True
    Me.optEuler(7).Value = False
    
    cboEulerUd.ListIndex = -1
   
    
End Sub

Private Sub BloquearFicha(Bloquea As Boolean)
Dim N As Byte
    
        
        cboEulerUd.Enabled = Not Bloquea
    
        For N = 0 To Me.txtEuler.Count - 1
            BloquearTxt txtEuler(N), Bloquea
        Next
    
        For N = 0 To Me.optEuler.Count - 1
            Me.optEuler(N).Enabled = Not Bloquea
        Next N
        
        For N = 0 To chkEuler.Count - 1
            chkEuler(N).Enabled = Not Bloquea
        Next

End Sub


Private Function CamposSQlFicha() As String
    'Primero iran todos los txts juntos y por orden de index
    CamposSQlFicha = "RecepAgenCliMat,RecpNumExp,FechaAlb,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,DatosBommarca"
    CamposSQlFicha = CamposSQlFicha & ",DatosBomNumCurva,DatosBomModelo,DatosBomNumSerie,DatosBomAno,DatosBomH,DatosBomCaudal"
    CamposSQlFicha = CamposSQlFicha & ",DatosMotorMarca , DatosMotorModelo, DatosMotorNumSerie, DatosMotorV, DatosMotorI"
    CamposSQlFicha = CamposSQlFicha & ",DatosMotorCV, DatosMotorKw, DatosMotorrpm,NumTrabajExterno,NumParteTrabajo"

    'Tipo bomba recepcionada
    'Son los check. Tambien vmos con el ordern
    CamposSQlFicha = CamposSQlFicha & ", TipoBombResSuperHor,TipoBombResSuperVer,TipoBombResSumPoz, TipoBombResSumVer, TipoBomAgitadorRes"
    CamposSQlFicha = CamposSQlFicha & ", TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombLimSumPoz, TipoBombLimSumVer, TipoBomAgitadorLim "
    

    'Luego resto campos
    CamposSQlFicha = CamposSQlFicha & ",numrepar ,  RecepAgenClien,RecepPortes, DatosBomUdCaudal,DatosBomTipoRodete"
    
End Function

Private Sub PonerCamposFicha()
Dim N As Byte
Dim SQL As String
    
    SQL = CamposSQlFicha()
    If EsHistorico Then
       SQL = "Select " & SQL & " FROM schrepeu WHERE numrepar = " & Text1(2).Text & " AND fecrepar =" & DBSet(Text1(4).Text, "F")
    Else
        SQL = "Select " & SQL & " FROM scarepeu WHERE numrepar = " & Text1(2).Text
    End If
        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        LimpiarFichaTecnica False
        
    Else
        
        
        
        'EL SQL estara montaddo para que coincida el orden del columna con el index
        For N = 0 To txtEuler.Count - 1
            txtEuler(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            If N = 20 Or N = 21 Then
                'NUmerico
                If txtEuler(N).Text <> "" Then txtEuler(N).Text = Format(txtEuler(N).Text, "000000")
            End If
        Next
    
        'Agencia cliente
        N = 1
        If DBLet(miRsAux!RecepAgenClien, "N") = 0 Then N = 0
        optEuler(N).Value = True
        
        N = 3
        If DBLet(miRsAux!RecepPortes, "N") = 1 Then N = 2
        optEuler(N).Value = True
        
        'Empieza en la 20
        For N = 1 To Me.chkEuler.Count
            chkEuler(N - 1).Value = DBLet(miRsAux.Fields(CInt(N) + 21), "N")
        Next
        
        ' DatosBomUdCaudal,DatosBomTipoRodete"
        kCampo = DBLet(miRsAux!DatosBomTipoRodete, "N")
        If kCampo = 0 Then kCampo = 6 'OTROS
        For N = 4 To 7
            If N = kCampo Then Me.optEuler(N).Value = True
        Next
        
        If miRsAux!DatosBomUdCaudal >= 0 Then Me.cboEulerUd.ListIndex = miRsAux!DatosBomUdCaudal
            
        kCampo = DBLet(miRsAux!DatosBomTipoRodete, "N")
        'Combo1.ListIndex = kCampo
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Function ActualizaBDFicha() As String
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
    s = s & Me.cboEulerUd.ListIndex & ","
    'Rodete
    kCampo = 6
    For N = 4 To 7
        If Me.optEuler(N).Value Then kCampo = N
    Next
    s = s & kCampo & ")"
    
    
   
   conn.Execute s
    
End Function


Private Function BuscaEnBDFicha() As String
Dim Columnas As String
Dim SQ As String
Dim N As Byte

    Columnas = CamposSQlFicha()
    Columnas = Replace(Columnas, ",", "|")
    
    BuscaEnBDFicha = ""
    
    For N = 0 To txtEuler.Count - 1
        If Trim(txtEuler(N).Text) <> "" Then
            SQ = RecuperaValor(Columnas, CInt(N + 1))
            If N = 20 Or N = 21 Then
                SQ = SQ & " = " & DBSet(txtEuler(N), "N", "S")
            Else
                
                If InStr(1, txtEuler(N).Text, "*") > 0 Then
                    SQ = SQ & " like " & DBSet(Replace(Me.txtEuler(N).Text, "*", "%"), "T")
                Else
                    SQ = SQ & " = " & DBSet(txtEuler(N), "T", "S")
                End If
            End If
            BuscaEnBDFicha = BuscaEnBDFicha & " AND " & SQ
        End If
    Next

    For N = 1 To Me.chkEuler.Count
        If chkEuler(N - 1) = 1 Then
            SQ = RecuperaValor(Columnas, N + 20) & " = 1"
            BuscaEnBDFicha = BuscaEnBDFicha & " AND " & SQ
        End If
    Next
     
    'If Me.cboEulerT.ListIndex >= 0 Then BuscaEnBDFicha = BuscaEnBDFicha & " AND partetrabajo = " & cboEulerT.ListIndex
    If Me.cboEulerUd.ListIndex >= 0 Then BuscaEnBDFicha = BuscaEnBDFicha & " AND DatosBomUdCaudal = " & cboEulerUd.ListIndex
    
    
    
    'Rodete
    kCampo = 0
    For N = 4 To 7
        If Me.optEuler(N).Value Then kCampo = N
    Next
    If kCampo > 0 Then BuscaEnBDFicha = BuscaEnBDFicha & " AND DatosBomTipoRodete = " & kCampo
        
    '
    ' Me.optEuler(1).Value) & "," & Abs(Me.optEuler(3).Value) & ","
    If BuscaEnBDFicha <> "" Then
        BuscaEnBDFicha = Mid(BuscaEnBDFicha, 5)
        BuscaEnBDFicha = "Select numrepar from scarepeu WHERE " & BuscaEnBDFicha
        BuscaEnBDFicha = " numrepar IN (" & BuscaEnBDFicha & ")"
    End If
      
End Function





Private Sub txtEuler_GotFocus(Index As Integer)
    ConseguirFoco txtEuler(Index), Modo
End Sub

Private Sub txtEuler_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtEuler_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtEuler(Index), Modo) Then Exit Sub
    
    If Index = 20 Or Index = 21 Then
        txtEuler(Index).Text = Trim(txtEuler(Index).Text)
        If txtEuler(Index).Text <> "" Then
            If Not PonerFormatoEntero(txtEuler(Index)) Then
                txtEuler(Index).Text = ""
            Else
                CadenaSQL = DevuelveDesdeBD(conAri, "numalbar", "scaalb", "numalbar", txtEuler(Index).Text)
                'Label3(36 o 37
                If CadenaSQL = "" Then MsgBox "El albaran de " & Label3(Index + 16).Caption & " NO existe", vbExclamation
            End If
        End If
    End If
        
End Sub