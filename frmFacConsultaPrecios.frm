VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacConsultaPrecios2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11325
   ClientLeft      =   345
   ClientTop       =   2430
   ClientWidth     =   19110
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   45
      TabIndex        =   73
      Top             =   -45
      Width           =   1410
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   74
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Limpiar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FramePDF 
      Height          =   10560
      Left            =   14055
      TabIndex        =   46
      Top             =   675
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Ver PDF"
         Height          =   255
         Left            =   3690
         TabIndex        =   47
         Top             =   10260
         Visible         =   0   'False
         Width           =   975
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   9945
         Left            =   0
         TabIndex        =   72
         Top             =   270
         Width           =   4635
         _cx             =   5080
         _cy             =   5080
      End
   End
   Begin VB.Frame FrameMostrarDatos 
      Height          =   10545
      Left            =   45
      TabIndex        =   4
      Top             =   675
      Width           =   13950
      Begin VB.Frame FrameNavegaDoc 
         Height          =   645
         Left            =   135
         TabIndex        =   67
         Top             =   6660
         Width           =   6240
         Begin VB.OptionButton optDoc 
            Caption         =   "Facturas"
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
            Left            =   180
            TabIndex        =   71
            Tag             =   "5"
            Top             =   270
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Albaranes"
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
            Left            =   1575
            TabIndex        =   70
            Tag             =   "6"
            Top             =   270
            Width           =   1290
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Pedidos"
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
            Left            =   3105
            TabIndex        =   69
            Tag             =   "7"
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optDoc 
            Caption         =   "Ofertas"
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
            Left            =   4590
            TabIndex        =   68
            Tag             =   "8"
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   3780
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3780
         Width           =   4905
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   3330
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   3330
         Width           =   4905
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   2115
         Width           =   5550
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   2115
         Width           =   5775
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   12415
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   630
         Width           =   1340
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   10455
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   630
         Width           =   1340
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1620
         Width           =   5550
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1620
         Width           =   4905
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   3780
         Width           =   1215
      End
      Begin VB.CheckBox chkCtrolStock 
         Caption         =   "Check1"
         Height          =   195
         Left            =   150
         TabIndex        =   43
         Top             =   4410
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
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
         Left            =   1290
         TabIndex        =   1
         Top             =   2895
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1335
         TabIndex        =   0
         Text            =   "000000"
         Top             =   660
         Width           =   810
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Limpiar"
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
         Index           =   2
         Left            =   11475
         TabIndex        =   2
         Top             =   10800
         Width           =   1065
      End
      Begin VB.TextBox txtResultado 
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
         Height          =   285
         Index           =   17
         Left            =   10110
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4350
         Width           =   1215
      End
      Begin VB.TextBox txtResultado 
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
         Height          =   285
         Index           =   16
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4335
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Height          =   285
         Index           =   15
         Left            =   5580
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   4335
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Height          =   285
         Index           =   14
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   4335
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListStock 
         Height          =   1485
         Left            =   4995
         TabIndex        =   30
         Top             =   4875
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   2619
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
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Almacen"
            Object.Width           =   3703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Ped.cli"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Ped Prov"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Disponible"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   8550
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3780
         Width           =   1215
      End
      Begin MSComctlLib.ListView listTarifa 
         Height          =   1485
         Left            =   120
         TabIndex        =   27
         Top             =   4875
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2619
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
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tarifa"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3780
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   12420
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3780
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtResultado 
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
         Index           =   10
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3780
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2895
         Width           =   6570
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   645
         Width           =   705
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   645
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1140
         Width           =   5550
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1140
         Width           =   4905
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
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
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   660
         Width           =   4905
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Salir"
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
         Index           =   1
         Left            =   12705
         TabIndex        =   3
         Top             =   10785
         Width           =   1065
      End
      Begin MSComctlLib.ListView listDatos 
         Height          =   2895
         Left            =   135
         TabIndex        =   31
         Tag             =   "0"
         Top             =   7365
         Width           =   13620
         _ExtentX        =   24024
         _ExtentY        =   5106
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
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "T"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   2681
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dto1"
            Object.Width           =   1729
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dto2"
            Object.Width           =   1729
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2647
         EndProperty
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   0
         Left            =   1035
         Picture         =   "frmFacConsultaPrecios.frx":0000
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   21
         Left            =   135
         TabIndex        =   66
         Top             =   3810
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Familia"
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
         TabIndex        =   63
         Top             =   3330
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
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
         Left            =   135
         TabIndex        =   59
         Top             =   2115
         Width           =   720
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
         Index           =   18
         Left            =   7200
         TabIndex        =   57
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Ult.Movimiento"
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
         Left            =   11835
         TabIndex        =   56
         Top             =   345
         Width           =   1935
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
         Index           =   16
         Left            =   10440
         TabIndex        =   54
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   15
         Left            =   7200
         TabIndex        =   52
         Top             =   1650
         Width           =   1215
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
         Index           =   8
         Left            =   135
         TabIndex        =   50
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Stock"
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
         Left            =   8550
         TabIndex        =   45
         Top             =   3495
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   1
         Left            =   975
         ToolTipText     =   "Buscar artículo"
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label lblSituacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11100
         TabIndex        =   42
         Top             =   2865
         Width           =   2625
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   11520
         TabIndex        =   41
         Top             =   4380
         Width           =   2100
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   9075
         TabIndex        =   40
         Top             =   4380
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   7020
         TabIndex        =   39
         Top             =   4380
         Width           =   465
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   5010
         TabIndex        =   38
         Top             =   4380
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2250
         TabIndex        =   37
         Top             =   4380
         Width           =   975
      End
      Begin VB.Label lblIndicador 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Left            =   11940
         TabIndex        =   32
         Top             =   10230
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Disponible"
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
         Left            =   9810
         TabIndex        =   29
         Top             =   3495
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "P.V.P."
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
         Left            =   7230
         TabIndex        =   26
         Top             =   3495
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P.M.P"
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
         Left            =   12420
         TabIndex        =   24
         Top             =   3495
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "P.U.Compra"
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
         Left            =   11070
         TabIndex        =   22
         Top             =   3495
         Visible         =   0   'False
         Width           =   1215
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
         Index           =   10
         Left            =   150
         TabIndex        =   20
         Top             =   2895
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
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
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   2535
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. P.P."
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
         Left            =   9255
         TabIndex        =   14
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Dto.Gral"
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
         Left            =   8205
         TabIndex        =   13
         Top             =   345
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
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
         Left            =   150
         TabIndex        =   12
         Top             =   1140
         Width           =   1170
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   4
         Left            =   7215
         TabIndex        =   11
         Top             =   1140
         Width           =   1215
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
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Documentos"
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
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   6390
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00972E0B&
         BorderWidth     =   2
         Index           =   2
         X1              =   1710
         X2              =   13680
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00972E0B&
         BorderWidth     =   2
         Index           =   1
         X1              =   1320
         X2              =   13680
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00972E0B&
         BorderWidth     =   2
         Index           =   0
         X1              =   1320
         X2              =   13725
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00972E0B&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1245
         Top             =   4230
         Width           =   12465
      End
   End
End
Attribute VB_Name = "frmFacConsultaPrecios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Fecha As String   'Podra ser "" o una fecha valida
Public ConsultaDesdeFrm As String
Private WithEvents frmA As frmBasico2
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmC As frmBasico2
Attribute frmC.VB_VarHelpID = -1
Private frmAlb As frmFacEntAlbaranes2

Dim cad As String
Dim IT As ListItem

Dim Valor As Currency
Dim AntiguoTxt As String
Dim PrimeraVez As Boolean



Private Sub LimpiarResultados()
Dim T As TextBox
Dim Index As Integer
    lblIndicador.Caption = ""
    For Each T In Me.txtResultado
        Index = T.Index
        If Index <> 0 And Index <> 1 And Index <> 8 And Index <> 9 Then T.Text = ""
    Next
    Me.listTarifa.ListItems.Clear
    Me.ListStock.ListItems.Clear
    Me.listDatos.ListItems.Clear
    Label2(4).Caption = ""

    lblSituacion.Caption = ""
    
    
    
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 1 Then
        'PonerVisiblePedir True
        Unload Me
    ElseIf Index = 2 Then
        limpiar Me
        LimpiarResultados
    Else
        Unload Me
    End If
End Sub



Private Sub Combo1_Click()
    If PrimeraVez Then Exit Sub
    '------------------------------------------
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "Leyendo BD"
    Set miRsAux = New ADODB.Recordset
    CargarDatosFacturacion
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
    Me.lblIndicador.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Private Sub Command1_Click()
    If Me.AcroPDF1.visible Then
        frmEulerPDF.Tag = Me.AcroPDF1.Tag
        frmEulerPDF.Show vbModal
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Fecha = "" Then Fecha = Now
        Caption = Caption & "     (" & Format(Fecha, "dd/mm/yyyy") & ")"
        If ConsultaDesdeFrm <> "" Then
            txtCodigo(0).Text = RecuperaValor(ConsultaDesdeFrm, 1)
            
            txtCodigo_LostFocus 0
            txtCodigo(1).Text = RecuperaValor(ConsultaDesdeFrm, 2)
            txtCodigo_LostFocus 1
        End If
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    limpiar Me
    Caption = "Consulta precios"
    'ASigno estos iconos
    lblIndicador.Caption = ""
    lblSituacion.Caption = ""
    Me.listDatos.SmallIcons = frmPpal.ImgListPpal
    If ConsultaDesdeFrm <> "" Then
        Me.optDoc(1).Value = True
    Else
        Me.optDoc(0).Value = True
    End If
         
    For i = 1 To imgBuscarG.Count - 1
        imgBuscarG(i).Picture = imgBuscarG(0).Picture
    Next
    
    With Toolbar1
        .ImageList = frmPpal.ImgListComun2
        .DisabledImageList = frmPpal.imgListComun_BN2
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(3).Image = 16  'Imprimir
    End With
    
    
    If InstalacionEsEulerTaxco Then
        Me.Width = 19165
    Else
        Me.Width = 14075
    End If
    
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Fecha = ""
    ConsultaDesdeFrm = ""
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(1).Text = RecuperaValor(CadenaSeleccion, 1)
    txtResultado(9).Text = RecuperaValor(CadenaSeleccion, 2)
    cad = "O"
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtResultado(1).Text = RecuperaValor(CadenaSeleccion, 2)
    cad = "O"
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
Dim KCargo As Integer
    KCargo = -1
    cad = "" 'Para ver si devuelve datos
    If Index = 0 Then
        'Cliente
        
        Set frmC = New frmBasico2
'        frmC.DatosADevolverBusqueda = "1|"
'        frmC.Show vbModal $$$$$
        AyudaClientes frmC, txtCodigo(0)
        Set frmC = Nothing
        'If Cad <> "" Then PonerFoco txtCodigo(1)
        If cad <> "" Then
            'cmdBuscar_Click
            KCargo = 0
            If txtCodigo(1).Text <> "" Then KCargo = 2
        End If
    Else
        'Articulo
        Set frmA = New frmBasico2
        'frmA.DeConsulta = True
        'frmA.DatosADevolverBusqueda3 = "@1@"
'        frmA.DesdeTPV = False
'        frmA.Show vbModal
        AyudaArticulos frmA, txtCodigo(1)
        Set frmA = Nothing
        If cad <> "" Then
            'cmdBuscar_Click
            KCargo = 1
            If txtCodigo(0).Text <> "" Then KCargo = 2
        End If
    End If
    cad = ""
    Set miRsAux = New ADODB.Recordset
    If KCargo >= 0 Then CargarDatos CByte(KCargo)
    Set miRsAux = Nothing
End Sub


'
Private Sub listDatos_DblClick()
Dim Seleccionado As Long
Dim SQL As String

    If listDatos.ListItems.Count = 0 Then Exit Sub
    If listDatos.SelectedItem Is Nothing Then Exit Sub


    If ConsultaDesdeFrm <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta en proceso de generacion de oferta/pedido/albaran.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Dim i As Integer
    If optDoc(0).Value Then i = 0
    If optDoc(1).Value Then i = 1
    If optDoc(2).Value Then i = 2
    If optDoc(3).Value Then i = 3
    
    
    Select Case i
    
    Case 0
            'Si lo primero no es una F no es una factura
            SQL = ""
            
            If Mid(Me.listDatos.SelectedItem.Tag, 1, 1) <> "F" Then
                MsgBox "No es una  factura (F)", vbExclamation
            Else
                SQL = Mid(listDatos.SelectedItem.Tag, 1, 3) & "|" & Mid(listDatos.SelectedItem.Tag, 4) & "|"
                SQL = SQL & listDatos.SelectedItem.SubItems(1) & "|"
            End If
            If SQL = "" Then Exit Sub
            With frmFacHcoFacturas2
                .DesdeFichaCliente = True
                .hcoCodMovim = RecuperaValor(SQL, 2)
                .hcoCodTipoM = RecuperaValor(SQL, 1)
                .hcoFechaMov = RecuperaValor(SQL, 3)
                .Show vbModal
        End With
    Case 1
              'ALBARANES
            'Si lo primero no es una A no es una factura
            SQL = ""
            
            If Mid(Me.listDatos.SelectedItem.Tag, 1, 1) <> "A" Then
                If Mid(Me.listDatos.SelectedItem.Tag, 1, 3) <> "DEV" Then SQL = "N"
            End If
            If SQL <> "" Then
                MsgBox "No es un albaran(A* - DEV)", vbExclamation
                Exit Sub
            End If
            
            SQL = Mid(listDatos.SelectedItem.Tag, 1, 3) & "|" & Mid(listDatos.SelectedItem.Tag, 4) & "|"
            
            If SQL = "" Then Exit Sub
              
        If vParamAplic.TipoFormularioClientes = 0 Then
            
            frmFacEntAlbaranes2.hcoCodMovim = RecuperaValor(SQL, 2)
            frmFacEntAlbaranes2.hcoCodTipoM = RecuperaValor(SQL, 1)
            frmFacEntAlbaranes2.Show vbModal
         
            
        Else
         
            frmFacEntAlbSAIL.hcoCodMovim = listDatos.SelectedItem.SubItems(1)
            frmFacEntAlbSAIL.hcoCodTipoM = listDatos.SelectedItem.Text
            frmFacEntAlbSAIL.Show vbModal
     
                 
            
        End If
    Case 2
        
            'PEDIDO CLIENTE
            If vParamAplic.TipoFormularioClientes = 0 Then
            
                frmFacEntPedidos.DatosADevolverBusqueda2 = listDatos.SelectedItem.Tag
                frmFacEntPedidos.EsHistorico = False
                frmFacEntPedidos.Show vbModal
            Else
                frmFacEntPedSail.DatosADevolverBusqueda2 = listDatos.SelectedItem.Tag
                frmFacEntPedSail.EsHistorico = False
                frmFacEntPedSail.Show vbModal
            End If
    Case 3
        'ofertas
            If vParamAplic.TipoFormularioClientes = 0 Then
             
                frmFacEntOfertas2.DatosOferta = listDatos.SelectedItem.Tag
                frmFacEntOfertas2.Show vbModal
                
            Else
                frmFacEntOferSAIL.DatosOferta = listDatos.SelectedItem.Tag
                frmFacEntOferSAIL.Show vbModal
            End If
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    listDatos.SetFocus
    Seleccionado = listDatos.SelectedItem.Index
    
    Combo1_Click
    If Not listDatos.SelectedItem Is Nothing Then listDatos.SelectedItem.Selected = False
    Set listDatos.SelectedItem = Nothing
    If listDatos.ListItems.Count >= Seleccionado Then
            listDatos.ListItems(Seleccionado).Selected = True
            listDatos.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault


End Sub

Private Sub optDoc_Click(Index As Integer)

    If PrimeraVez Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "Leyendo BD"
    Set miRsAux = New ADODB.Recordset
    CargarDatosFacturacion
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
    Me.lblIndicador.Caption = ""
    Screen.MousePointer = vbDefault


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            cmdCancelar_Click 2
        Case 3 'Imprimir
            'Estadisitcas de veces consultado  precio/cliente
            frmVarios.Opcion = 2
            frmVarios.Show vbModal
    End Select

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    AntiguoTxt = txtCodigo(Index).Text
    ConseguirFoco txtCodigo(Index), 3
End Sub



Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Opc As Byte
    txtCodigo(Index).Text = Trim(txtCodigo(Index))
    If AntiguoTxt = txtCodigo(Index).Text Then Exit Sub
    cad = ""
    Opc = 100
    If Index = 0 Then
        lblIndicador.Caption = ""
        
        'Cliente
        If txtCodigo(Index).Text <> "" Then
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
                
            Else
                cad = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCodigo(Index).Text, "N")
                If cad = "" Then
                    MsgBox "No existe el cliente : " & txtCodigo(Index).Text, vbExclamation
                End If
            End If
        End If
        If cad <> "" Then
            Opc = 0
            If txtCodigo(1).Text <> "" Then Opc = 2
            
        End If
    Else
        'articulo
        If txtCodigo(Index).Text <> "" Then
            cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtCodigo(Index).Text, "T")
            If cad = "" Then
                MsgBox "No existe el articulo: " & txtCodigo(Index).Text, vbExclamation
                PonerFoco txtCodigo(Index)
            End If
        End If
        If cad <> "" Then
            Opc = 1
            If txtCodigo(0).Text <> "" Then Opc = 2
        End If
    End If
    'Me.txtNombre(Index).Text = Cad
    If cad = "" Then
        txtCodigo(Index).Text = ""
        PonerFoco txtCodigo(Index)
    End If
    If Opc = 100 Then
        'Mal. Borramos los campos
        
        LimpiarlosCampos CByte(Index)
        
            
    Else
        Set miRsAux = New ADODB.Recordset
        CargarDatos Opc
        Set miRsAux = Nothing
    End If
End Sub


Private Sub CargaStock()
Dim i As Currency
Dim J As Integer

    Valor = 0
    ListStock.ListItems.Clear
    txtResultado(13).Text = ""
    cad = "select salmac.codalmac,nomalmac,canstock   from salmac,salmpr where salmac.codalmac="
    cad = cad & "salmpr.codalmac AND  codartic=" & DBSet(txtCodigo(1).Text, "T") & " ORDER BY salmac.codalmac"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListStock.ListItems.Add()
        IT.Text = miRsAux!codAlmac
        IT.SubItems(1) = miRsAux!nomalmac
        i = DBLet(miRsAux!CanStock, "N")
        IT.SubItems(2) = Format(i, FormatoCantidad)
        IT.SubItems(3) = " ": IT.SubItems(4) = " "
        IT.SubItems(5) = IT.SubItems(2)
        Valor = Valor + i
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Stock
    txtResultado(13).Text = Format(Valor, FormatoCantidad)
    
    'Cargamos primero los de cliente
    'FALTA###
    'If chkCtrolStock.Value Then
    If True Then
        cad = "select codalmac,sum(cantidad) as cuantos"
        cad = cad & " from sliped where codartic='"
        cad = cad & DevNombreSQL(txtCodigo(1).Text) & "' GROUP BY 1"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
              For J = 1 To ListStock.ListItems.Count
                    If ListStock.ListItems(J).Text = CStr(miRsAux.Fields(0)) Then
                        'ES este
                        i = DBLet(miRsAux.Fields(1), "N")
                        If i <> 0 Then ListStock.ListItems(J).SubItems(3) = Format(i, FormatoCantidad)
                        Valor = Valor - i
                        
                        i = ImporteFormateado(ListStock.ListItems(J).SubItems(2)) - i
                        ListStock.ListItems(J).SubItems(5) = Format(i, FormatoCantidad)
                        Exit For
                    End If
            Next
            miRsAux.MoveNext
        Wend
        miRsAux.Close


'    'Cargamos los comprados
'    C = "select scappr.numpedpr,fecpedpr,codprove,nomprove,sum(cantidad) as cuantos"
'    C = C & " from scappr,slippr where scappr.numpedpr=slippr.numpedpr  and codartic='"
'    C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' group by 1"

        cad = "select codalmac,sum(cantidad) as cuantos"
        cad = cad & " from slippr where codartic='"
        cad = cad & DevNombreSQL(txtCodigo(1).Text) & "' GROUP BY 1"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
              For J = 1 To ListStock.ListItems.Count
                    If ListStock.ListItems(J).Text = CStr(miRsAux.Fields(0)) Then
                        'ES este
                        i = DBLet(miRsAux.Fields(1), "N")
                        If i <> 0 Then ListStock.ListItems(J).SubItems(4) = Format(i, FormatoCantidad)
                        Valor = Valor + i
                        
                        i = ImporteFormateado(ListStock.ListItems(J).SubItems(2)) + i
                                'los pedidos clientes (reservas)
                        i = i - ImporteFormateado(Trim(ListStock.ListItems(J).SubItems(3)))
                        ListStock.ListItems(J).SubItems(5) = Format(i, FormatoCantidad)
          
                        Exit For
                    End If
            Next
            miRsAux.MoveNext
        Wend
        miRsAux.Close


    End If
    
    'Disponible
    txtResultado(0).Text = Format(Valor, FormatoCantidad)
    
End Sub


Private Sub CargaTarifas()
Dim F As Date

    
    
    
    listTarifa.ListItems.Clear
    cad = "select slista.codlista,nomlista,precioac,fechanue,precionu from slista,starif where slista.codlista="
    cad = cad & "starif.codlista and codartic = " & DBSet(txtCodigo(1).Text, "T") & " ORDER BY slista.codlista"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = listTarifa.ListItems.Add()
        IT.Text = miRsAux!codlista
        IT.SubItems(1) = miRsAux!nomlista
        
        F = CDate("01/01/2111")
        If Not IsNull(miRsAux!fechanue) Then
            If Not IsNull(miRsAux!precionu) Then F = miRsAux!fechanue
        End If
        
        If CDate(Fecha) >= F Then
            'Coje el nuevo
            IT.SubItems(2) = Format(DBLet(miRsAux!precionu, "N"), FormatoPrecio) & " *"
        Else
            IT.SubItems(2) = Format(DBLet(miRsAux!precioac, "N"), FormatoPrecio)
        End If
        If miRsAux!codlista = listTarifa.Tag Then
            'Tarifa del cliente
            IT.Bold = True
            IT.ForeColor = vbBlue
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(1).ForeColor = vbBlue
    
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(2).ForeColor = vbBlue
            
        End If
                    
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If listTarifa.ListItems.Count > 6 Then
        listTarifa.ColumnHeaders.Item(2).Width = 2400 '2600
    Else
        listTarifa.ColumnHeaders.Item(2).Width = 2600 '2800
    End If
    
End Sub



'0. Cliente
'1.- Articulop
'2.los dos

Private Sub CargarDatos(Opcion As Byte)
Dim Familia As Integer
Dim marca As Integer

    On Error GoTo EC
    
    cad = "OK"

    If Opcion <> 1 Then
        lblIndicador.Caption = "Datos cliente"
        lblIndicador.Refresh
        
        cad = "select codclien ,nomclien ,dtoppago ,dtognral  ,codsitua ,codmacta,codforpa,codtarif "
        cad = cad & ",codagent, fechamov, fechaalt, perclie1, telclie1, maiclie1 "
        cad = cad & " from sclien where codclien =" & Me.txtCodigo(0).Text
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        'Ponemos los campos
        '--------------------------------------------------------
        Me.txtCodigo(0).Text = miRsAux!codClien
        Me.txtResultado(1).Text = miRsAux!NomClien
        Me.txtResultado(2).Text = miRsAux!codforpa
        Me.txtResultado(3).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", miRsAux!codforpa, "N")
        Me.txtResultado(4).Text = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", miRsAux!codsitua, "N")
        Me.txtResultado(6).Text = Format(miRsAux!DtoGnral, FormatoDescuento)
        Me.txtResultado(7).Text = Format(miRsAux!DtoPPago, FormatoDescuento)
        '15/02/2021: nuevos
        txtResultado(8).Text = miRsAux!CodAgent
        txtResultado(5).Text = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", miRsAux!CodAgent, "N")
        txtResultado(19).Text = Format(miRsAux!fechaalt, "dd/mm/yyyy")
        txtResultado(20).Text = Format(miRsAux!FechaMov, "dd/mm/yyyy")
        txtResultado(18).Text = DBLet(miRsAux!perclie1, "T")
        txtResultado(22).Text = DBLet(miRsAux!telclie1, "T")
        txtResultado(21).Text = DBLet(miRsAux!maiclie1, "T")
        
        
        'Cargo la cta contable
        cad = DBLet(miRsAux!Codmacta, "T")
        
        'Cargo la tarifa
        Me.listTarifa.Tag = miRsAux!codTarif
        
        'Cerramos el RS
        miRsAux.Close
    
    
    
    
        lblIndicador.Caption = "Cobros pendientes"
        lblIndicador.Refresh
    
        PonerCobrosPendientes cad
    
'--        txtResultado(5).Text = Format(Valor, FormatoImporte)
    
        DoEvents
    End If
    'Datos articulo
    If Opcion <> 0 Then
        lblIndicador.Caption = "Articulo"
        lblIndicador.Refresh
        
        cad = "select codartic,nomartic,preciouc,preciomp,preciove,unicajas,codstatu,ctrstock,codfamia,codmarca  from sartic where codartic =" & DBSet(Me.txtCodigo(1).Text, "T")
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Me.txtCodigo(1).Text = miRsAux!codArtic
        Me.txtResultado(9).Text = miRsAux!NomArtic
        Me.txtResultado(10).Text = Format(DBLet(miRsAux!precioUC, "N"), FormatoPrecio)
        Me.txtResultado(11).Text = Format(DBLet(miRsAux!PrecioMP, "N"), FormatoPrecio)
        Me.txtResultado(12).Text = Format(DBLet(miRsAux!PrecioVe, "N"), FormatoPrecio)
        chkCtrolStock.Value = miRsAux!CtrStock  'guardare si lleva control de stock
        Me.txtResultado(9).Tag = miRsAux!unicajas
        Select Case miRsAux!codstatu
        'Abril 2014
        Case 1
            lblSituacion.Caption = "Obsoleto"
        Case 2
            lblSituacion.Caption = "Bloqueado"
        Case 3
            lblSituacion.Caption = "Caducado"
        Case Else
            lblSituacion.Caption = ""
        End Select
        Familia = miRsAux!Codfamia
        marca = miRsAux!codmarca
        miRsAux.Close
        
        txtResultado(24) = Format(Familia, "0000")
        txtResultado(23) = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", CStr(Familia), "N")
        
        txtResultado(26) = Format(marca, "0000")
        txtResultado(25) = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", CStr(marca), "N")
            
        lblIndicador.Caption = "Stock"
        lblIndicador.Refresh
        CargaStock
        
        
        lblIndicador.Caption = "Tarifas"
        lblIndicador.Refresh
        CargaTarifas
    
    
        If InstalacionEsEulerTaxco Then
            'Si la familia, marca tiene catalogo lo mostrara
            cad = "Select * from eulerprecios  WHERE "
            cad = cad & "( codfamia =" & Familia & " AND codmarca =" & marca & ")"
            cad = cad & " OR ( codfamia =" & Familia & " AND codmarca is null )"
            cad = cad & " OR ( codfamia is NULL AND codmarca =" & marca & ")"
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            If Not miRsAux.EOF Then
                'OOOOOOK
                'Tiene un documento asociado
                cad = miRsAux!Documento
                
            End If
            miRsAux.Close
            CargaArchivo cad
            
        End If
        DoEvents
    
    End If
    If Opcion = 2 Then
        'Datos albaranes......
        CargarDatosFacturacion
    
    
        'Ponemos el precio fianl
        CalcularPrecioFinal
    
    End If
    
    'Insertamos log de consulta
    If txtCodigo(0).Text <> "" And txtCodigo(1).Text <> "" Then
        If ConsultaDesdeFrm = "" Then
            lblIndicador.Caption = "Ins. log"
            lblIndicador.Refresh
            cad = "insert into `sconsulta` (`DiaHora`,`Usuario`,`codclien`,`nomclien`,"
            '----------                                       cogera la fecha del mysql
            cad = cad & "`codartic`,`nomartic`) values (" & "concat(curdate(),' ',curtime())" & ","
            cad = cad & DBSet(vUsu.Nombre, "T") & "," & txtCodigo(0).Text & "," & DBSet(txtResultado(1), "T")
            cad = cad & "," & DBSet(txtCodigo(1), "T") & "," & DBSet(txtResultado(9), "T") & ")"
            conn.Execute cad
            Espera 0.3
            
        End If
    End If
    lblIndicador.Caption = ""
        
        
        
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub





Private Sub PonerCobrosPendientes(ByVal Codmacta As String)
    Valor = 0
    If Codmacta = "" Then Exit Sub
    'Obtener a partir de la cuenta del cliente si hay cobros pendientes en Contabilidad
    cad = " WHERE scobro.codmacta = '" & Codmacta & "'"
    cad = cad & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
    cad = cad & " AND (sforpa.tipforpa between 0 and 3)"
    
    If vParamAplic.ContabilidadNueva Then
        cad = " cobros as scobro INNER JOIN formapago as sforpa ON scobro.codforpa=sforpa.codforpa " & cad
    Else
        cad = " scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa " & cad
    End If
    cad = "SELECT sum(impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) FROM " & cad
    'Lee de la Base de Datos de CONTABILIDAD
    miRsAux.Open cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then Valor = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
End Sub



'Cargara los datos de las lineas
'de OFERTAS,PEDIDOS,ALBARANES,FACTURA
Private Sub CargarDatosFacturacion()

    Me.listDatos.ListItems.Clear
    
    If Me.txtCodigo(1).Text <> "" And txtCodigo(0).Text <> "" Then CargaDatosTablas
    
    
    
End Sub



Private Sub CargaDatosTablas()
Dim Aux As String
Dim Ico As Integer
Dim Ktabla As Integer

    Ktabla = 0
    If Me.optDoc(1).Value Then Ktabla = 1
    If Me.optDoc(2).Value Then Ktabla = 2
    If Me.optDoc(3).Value Then Ktabla = 3
    

    Select Case Ktabla
    Case 3
        Ico = 5
        cad = "slipre,scapre WHERE slipre.numofert=scapre.numofert"
        Aux = " '' as Primero,slipre.numofert as elnumero,fecofert as fecha"
        Me.lblIndicador.Caption = "Ofertas"
    Case 2
        Ico = 6
        cad = "sliped,scaped where sliped.numpedcl=scaped.numpedcl"
        Aux = " '' as Primero,sliped.numpedcl as elnumero,fecpedcl  as fecha"
        Me.lblIndicador.Caption = "Pedidos"
    Case 1
        Ico = 7
        cad = "slialb,scaalb where slialb.numalbar=scaalb.numalbar and slialb.codtipom=scaalb.codtipom"
        Aux = " slialb.codtipom as Primero,slialb.numalbar as elnumero,fechaalb as fecha"
        Me.lblIndicador.Caption = "Albaranes"
    Case Else
        'case 0
        Ico = 8
        cad = " slifac,scafac where slifac.numfactu=scafac.numfactu and slifac.codtipom=scafac.codtipom and slifac.fecfactu=scafac.fecfactu"
        Aux = "slifac.codtipom as primero,slifac.numfactu as elnumero,slifac.fecfactu as fecha"
        Me.lblIndicador.Caption = "Facturas"
    End Select
    Me.lblIndicador.Refresh
    
    Aux = "Select " & Aux & ",Cantidad, precioar, dtoline1, dtoline2, ImporteL FROM " & cad
    cad = Aux & " AND codartic = " & DBSet(Me.txtCodigo(1).Text, "T")
    cad = cad & " AND codclien = " & txtCodigo(0).Text
    cad = cad & " ORDER BY 3 desc,2"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set IT = Me.listDatos.ListItems.Add()
        cad = Trim(miRsAux!primero & " " & Format(miRsAux!elnumero, "000000"))
        IT.Text = ""
        
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        
        'Nuevo. El documento
        IT.SubItems(2) = cad
        IT.SubItems(3) = Format(miRsAux!cantidad, FormatoCantidad)
        IT.SubItems(4) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(5) = Format(miRsAux!dtoline1, FormatoDescuento)
        IT.SubItems(6) = Format(miRsAux!dtoline2, FormatoDescuento)
        IT.SubItems(7) = Format(miRsAux!ImporteL, FormatoImporte)
        IT.SmallIcon = Ico
        IT.Tag = cad
        IT.ToolTipText = Me.lblIndicador.Caption & " " & cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub





'Calculamos el precio que se le va a quedar
Private Sub CalcularPrecioFinal()
Dim CPrecioFact As CPreciosFact
Dim Precio As Currency
Dim cantidad As Currency
Dim PorCaja As Boolean
Dim NumCajas As Integer
                Set CPrecioFact = New CPreciosFact
                'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                'precio de caja, y otra linea con el resto unidades un precio unidad
                'Cantidad = txtAux(Index).Text
                cantidad = 1
                NumCajas = CPrecioFact.ObtenerNumCajas(CStr(cantidad), CStr(txtResultado(9).Tag))
                'RestoUnid = CInt(ComprobarCero(Cantidad)) - NumCajas * CInt(devuelve)
                'Obtenemos la Tarifa del Cliente
                CPrecioFact.CodigoArtic = Me.txtCodigo(1).Text
                CPrecioFact.CodigoClien = Me.txtCodigo(0).Text
                CPrecioFact.FijarTarifaActividad
                CPrecioFact.CodigoLista2 = Me.listTarifa.Tag  'la tarifa del cliente
                    
                
               
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, CStr(Fecha), cad, "")
                    
                'En cad TENGO el origen del precio
                Select Case cad
                    Case "P": Label2(4).Caption = "Promoción"
                    Case "E": Label2(4).Caption = "Precio Especial"
                    Case "T": Label2(4).Caption = "Tarifa Artículo"
                    Case "A": Label2(4).Caption = "Precio Artículo"
                    Case "M": Label2(4).Caption = "Manual"
                End Select
                
                    txtResultado(14).Text = Precio
                    PonerFormatoDecimal txtResultado(14), 2
                    txtResultado(15).Text = CPrecioFact.Descuento1
                    PonerFormatoDecimal txtResultado(15), 4
                    txtResultado(16).Text = CPrecioFact.Descuento2
                    PonerFormatoDecimal txtResultado(16), 4
                    txtResultado(17).Text = CalcularImporte(CStr(cantidad), txtResultado(14).Text, txtResultado(15).Text, txtResultado(16).Text, vParamAplic.TipoDtos)
                    PonerFormatoDecimal txtResultado(17), 1
                Set CPrecioFact = Nothing
End Sub




'0  Artic   1 Cliente      2 los dos
Private Sub LimpiarlosCampos(Opcion As Byte)
Dim i As Integer
 
    If Opcion <> 1 Then
        'ARTICULOS
        For i = 1 To 7
            txtResultado(i).Text = ""
        Next i
        chkCtrolStock.Value = 0  'guardare si lleva control de stock
        
        
        
    End If
    If Opcion <> 0 Then
        'CLIENTE
         For i = 9 To 13
            txtResultado(i).Text = ""
        Next i
        txtResultado(0).Text = ""
        Me.listTarifa.ListItems.Clear
        Me.ListStock.ListItems.Clear
        CargaArchivo ""
    End If
    
    
    listDatos.ListItems.Clear
    For i = 14 To 17
            txtResultado(i).Text = ""
    Next i
    Label2(4).Caption = ""
    lblSituacion.Caption = ""
End Sub



Private Function CargaArchivo(Archivo As String) As Boolean
    On Error GoTo eCargaArchivo
    
    If vParamAplic.NumeroInstalacion <> 4 Then Exit Function
    
    CargaArchivo = False
    
    If Archivo = "" Then
        AcroPDF1.visible = False
    Else
        AcroPDF1.LoadFile (Archivo)
        AcroPDF1.visible = True
    End If
    AcroPDF1.Tag = Archivo
    Me.Command1.visible = AcroPDF1.visible
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function
