VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantenusu2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantenusu2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   90
      TabIndex        =   15
      Top             =   45
      Width           =   9255
      Begin VB.ComboBox Combo6 
         Height          =   360
         ItemData        =   "frmMantenusu2.frx":000C
         Left            =   630
         List            =   "frmMantenusu2.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5730
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
         Height          =   360
         ItemData        =   "frmMantenusu2.frx":003E
         Left            =   630
         List            =   "frmMantenusu2.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2820
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   6720
         MaxLength       =   17
         PasswordChar    =   "*"
         TabIndex        =   25
         Text            =   "123456789012345"
         Top             =   4980
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   630
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   4980
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   630
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4260
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   630
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3540
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2670
         Width           =   1575
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7080
         TabIndex        =   27
         Top             =   5940
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5670
         TabIndex        =   26
         Top             =   5940
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2190
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   630
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1410
         Width           =   7725
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frmMantenusu2.frx":0070
         Left            =   630
         List            =   "frmMantenusu2.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2130
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Index           =   0
         Left            =   630
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   690
         Width           =   1365
      End
      Begin VB.Image imgQuitarSkin 
         Height          =   240
         Left            =   2760
         Picture         =   "frmMantenusu2.frx":0074
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Traer menús del usuario"
         Height          =   255
         Left            =   630
         TabIndex        =   55
         Top             =   5490
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Skin"
         Height          =   255
         Left            =   630
         TabIndex        =   52
         Top             =   2580
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "mail-password"
         Height          =   255
         Index           =   7
         Left            =   6810
         TabIndex        =   43
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mail-user"
         Height          =   255
         Index           =   6
         Left            =   630
         TabIndex        =   42
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Servidor SMTP"
         Height          =   255
         Index           =   5
         Left            =   630
         TabIndex        =   41
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   4
         Left            =   630
         TabIndex        =   40
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5190
         TabIndex        =   33
         Top             =   240
         Width           =   3345
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   4770
         Top             =   2070
         Width           =   3525
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Pass."
         Height          =   360
         Index           =   3
         Left            =   5010
         TabIndex        =   32
         Top             =   2670
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   360
         Index           =   2
         Left            =   5040
         TabIndex        =   31
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   630
         TabIndex        =   30
         Top             =   1890
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   29
         Top             =   1170
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   28
         Top             =   450
         Width           =   2295
      End
   End
   Begin VB.Frame FrameNormal 
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
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   9255
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   150
         TabIndex        =   47
         Top             =   0
         Width           =   2655
         Begin VB.CheckBox chkVistaPrevia 
            Caption         =   "Vista previa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3750
            TabIndex        =   48
            Top             =   270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   240
            TabIndex        =   49
            Top             =   180
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
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
                  Object.ToolTipText     =   "Prohibir acceso"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Copiar Menus"
                  Object.Tag             =   "0"
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Index           =   1
         ItemData        =   "frmMantenusu2.frx":0A76
         Left            =   7470
         List            =   "frmMantenusu2.frx":0A83
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1635
      End
      Begin VB.CommandButton cmdUsu 
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
         Height          =   375
         Index           =   3
         Left            =   5640
         Picture         =   "frmMantenusu2.frx":0AA6
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Prohibir acceso"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdConfigMenu 
         Caption         =   "Configurar menu"
         Height          =   375
         Left            =   7170
         TabIndex        =   38
         Top             =   2040
         Width           =   1785
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
         Height          =   1665
         Left            =   3480
         TabIndex        =   6
         Top             =   900
         Width           =   5655
         Begin VB.ComboBox Combo5 
            Height          =   360
            ItemData        =   "frmMantenusu2.frx":72F8
            Left            =   960
            List            =   "frmMantenusu2.frx":72FA
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1140
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   240
            Width           =   4515
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            ItemData        =   "frmMantenusu2.frx":72FC
            Left            =   960
            List            =   "frmMantenusu2.frx":730C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "Skin"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   1170
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   690
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
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
         Height          =   375
         Index           =   0
         Left            =   3960
         Picture         =   "frmMantenusu2.frx":733F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
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
         Height          =   375
         Index           =   1
         Left            =   4440
         Picture         =   "frmMantenusu2.frx":DB91
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modificar usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
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
         Height          =   375
         Index           =   2
         Left            =   4920
         Picture         =   "frmMantenusu2.frx":143E3
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   3480
         TabIndex        =   5
         Tag             =   $"frmMantenusu2.frx":1AC35
         Top             =   3150
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5080
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
            Text            =   "Código"
            Object.Width           =   1763
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resumido"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3352
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   4680
         TabIndex        =   53
         Top             =   2670
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar empresas bloquedas"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   8730
         TabIndex        =   56
         Top             =   240
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
      Begin VB.Label Label8 
         Caption         =   "Acceso"
         Height          =   255
         Index           =   1
         Left            =   6570
         TabIndex        =   46
         Top             =   6390
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3480
         TabIndex        =   13
         Top             =   690
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu2.frx":1ACD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu2.frx":21538
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu2.frx":21F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu2.frx":287AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu2.frx":2F00E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameEditorMenus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   45
      TabIndex        =   34
      Top             =   45
      Width           =   9255
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6015
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10610
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   36
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   35
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   6360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7380
      TabIndex        =   0
      Top             =   5970
      Width           =   975
   End
End
Attribute VB_Name = "frmMantenusu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 106

Dim PrimeraVez As Boolean
Dim SQL As String
Dim i As Integer
Dim UsuarioOrigen As Long

Dim rsEmpresasAriges As ADODB.Recordset



Private Sub cmdConfigMenu_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    CadenaDesdeOtroForm = Me.ListView1.SelectedItem.SubItems(1)
    frmEditorMenus2.CodigoActual = CInt(ListView1.SelectedItem.Text)
    frmEditorMenus2.Show vbModal
    CadenaDesdeOtroForm = ""
End Sub

Private Sub cmdEditorMenus_Click(Index As Integer)
    If Index = 0 Then
        GuardarMenuUsuario
    End If
    Me.FrameEditorMenus.visible = False
    
    
End Sub


Private Sub cmdFrameUsu_Click(Index As Integer)


    If Index = 0 Then
        If Combo6.ListIndex > 0 Then
            If MsgBox("Va a copiar los menus del usuario " & Trim(Text2(0).Text) & " con los del usuario " & Combo6.Text & vbCrLf & vbCrLf & "¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
    
        For i = 0 To Text2.Count - 1
            Text2(i).Text = Trim(Text2(i).Text)
            If i < 4 Then
                If Text2(i).Text = "" Then
                    MsgBox Label4(i).Caption & " requerido.", vbExclamation
                    Exit Sub
                End If
            End If
        Next i
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
            
'        'tipo de skin
'        If Combo4.ListIndex < 0 Then
'            MsgBox "Seleccione un tipo de skin", vbExclamation
'            Exit Sub
'        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        'Ahora vamos con los campos de e-mail
        CadenaDesdeOtroForm = ""
        For i = 4 To 7
            If Text2(i).Text <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
        Next i
        
        If Len(CadenaDesdeOtroForm) > 0 And Len(CadenaDesdeOtroForm) <> 4 Then
            MsgBox "Falta por rellenar correctamente los datos del e-mail.", vbExclamation
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
        
        'Compruebo que el login es unico
        i = 0
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            SQL = "Select login from usuarios.usuarios where login='" & Text2(0).Text & "'"
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not miRsAux.EOF Then SQL = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If SQL <> "" Then
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
            
        Else
            'MODIFICAR
            If FrameUsuario.Tag = 0 Then
                'Estoy modificando un dato normal
                i = CInt(ListView1.SelectedItem.Text)
            Else
                'Estoy agregando un usuario que ya existia en contabiñlidad
                'es decir, le estoy asignando su NIVELUSU de contabilidad
                i = CInt(FrameUsuario.Tag)
            End If
        End If
        
        If Combo6.ListIndex >= 0 Then
            UsuarioOrigen = Combo6.ItemData(Combo6.ListIndex)
        Else
            UsuarioOrigen = 0
        End If
        InsertarModificar i
        
        
    End If
    
    
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        'CargaUsuarios
        CadenaDesdeOtroForm = ""
    Else
        'Pero cargamos el tag como coresponde
        'ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
        
        If Me.FrameUsuario.Tag <> 0 Then
            CadenaDesdeOtroForm = FrameUsuario.Tag
        Else
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        End If
  
    End If
    
    CargaUsuarios
    If CadenaDesdeOtroForm <> "" Then
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Text = CadenaDesdeOtroForm Then
                Set ListView1.SelectedItem = ListView1.ListItems(i)
                Exit For
            End If
        Next i
    End If
    DatosUsusario
    CadenaDesdeOtroForm = ""
    'Para ambos casos
    Me.FrameUsuario.visible = False
    Me.FrameUsuario.Enabled = False
    Me.FrameNormal.visible = True
    Me.FrameNormal.Enabled = True
    
End Sub


Private Sub InsertarModificar(ByVal CodigoUsuario As Integer)
Dim Ant As Integer
Dim fin As Boolean
Dim Sql2 As String
Dim Excepcion As String
Dim CodUsuarioOrigen As Integer

Dim ArigesBD As String

On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset

    CodUsuarioOrigen = 0
    If UsuarioOrigen > 0 Then
        CodUsuarioOrigen = DevuelveValor("select codusu from usuarios.usuarios where login = " & DBSet(Combo6.Text, "T"))
    End If
    
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        SQL = "Select codusu from usuarios.usuarios where codusu > 0"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Ant = 1
        i = 1
        fin = False
        If miRsAux.EOF Then fin = True
        While Not fin
            If miRsAux!CodUsu - Ant > 0 Then
                'Hay un salto
                i = Ant
                fin = True
            Else
                Ant = Ant + 1
            End If
            If Not fin Then
                miRsAux.MoveNext
                If miRsAux.EOF Then
                    fin = True
                    i = Ant
                End If
            End If
        Wend
        miRsAux.Close

        
        SQL = "INSERT INTO usuarios.usuarios (codusu, nomusu,  nivelariges, login, passwordpropio,dirfich,skin, solotesoreria, skinariges) VALUES ("
        SQL = SQL & i
        SQL = SQL & ",'" & Text2(1).Text & "',"
        'Combo
        SQL = SQL & Combo2.ItemData(Combo2.ListIndex) & ",'"
        SQL = SQL & Text2(0).Text & "','"
        SQL = SQL & Text2(3).Text & "',"
        'DIR FICH tiene
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For i = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(i).Text & "|"
            Next i
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        SQL = SQL & CadenaDesdeOtroForm
        
        SQL = SQL & "," & Combo5.ItemData(Combo5.ListIndex) & ","
        SQL = SQL & "0," ' ChkSoloTesoreria.Value & ","
        SQL = SQL & Combo5.ItemData(Combo5.ListIndex) & ")"
        
        ' insercion en el menu_usuarios
        AbrirRsEmpresas
        While Not rsEmpresasAriges.EOF
            
'[Monica]14/04/2020
            ArigesBD = rsEmpresasAriges!AriGes

            Sql2 = "INSERT INTO " & ArigesBD & ".menus_usuarios (codusu,codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido) "
            Sql2 = Sql2 & " select " & i & ",codigo, aplicacion, "
            
            ' insertamos sin partir de ningún usuario
            If UsuarioOrigen <= 0 Then
                Select Case Combo2.ItemData(Combo2.ListIndex)
                    Case 0 ' superusuario
                        Sql2 = Sql2 & "1,1,1,1,1,0"
                    Case 1 ' administrador
                        Sql2 = Sql2 & "1,1,1,1,1,0"
                    Case 2 ' normal
                        Sql2 = Sql2 & "1,1,1,1,1,0"
                    Case 3 ' consulta
                        Sql2 = Sql2 & "1,0,0,1,0,0"
                End Select
                        
                Sql2 = Sql2 & " from " & ArigesBD & ".menus_usuarios "
                Sql2 = Sql2 & " where aplicacion in ('ariges','introcon') and codusu = 0"
            ' insertamos partiendo de un usuario
            Else
                Sql2 = Sql2 & " ver, creareliminar, modificar, imprimir, especial, expandido "
                Sql2 = Sql2 & " from " & ArigesBD & ".menus_usuarios "
                Sql2 = Sql2 & " where aplicacion in ('ariges','introcon') and codusu = " & DBSet(CodUsuarioOrigen, "N")
            End If
            conn.Execute Sql2
            
            
            Excepcion = ""
            ' dependiendo de si es Superusuario, Administrador, Normal o consulta
            Select Case Combo2.ItemData(Combo2.ListIndex)
                Case 0 'superusuario
                    
                Case 1 'administrador
                    Excepcion = "(1)"
                Case 2 'normal
                    Excepcion = "(1)" '"(1,10,12,13,14)" '[Monica]14/04/2020: solo datos basicos
                Case 3 'consulta
                    Excepcion = "(1)" '"(1,10,12,13,14)" '[Monica]14/04/2020: solo datos basicos
            End Select
            
            If Excepcion <> "" Then
'[Monica]14/04/2020
                Sql2 = "update  " & ArigesBD & ".menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0 "
                Sql2 = Sql2 & " where aplicacion in ('ariges') and codusu = " & DBSet(i, "N")
                Sql2 = Sql2 & " and (codigo in " & Excepcion
'[Monica]14/04/2020
'                Sql2 = Sql2 & " or codigo in (select codigo from ariges" & rsEmpresasAriges!codempre & ".menus where padre in " & Excepcion & " and aplicacion in ('ariges')))"
                Sql2 = Sql2 & " or codigo in (select codigo from " & ArigesBD & ".menus where padre in " & Excepcion & " and aplicacion in ('ariges')))"
            
                conn.Execute Sql2
            End If
            
            rsEmpresasAriges.MoveNext
        Wend
        CerrarRsEmpresas
        
    Else
        SQL = "UPDATE usuarios.usuarios Set nomusu='" & Text2(1).Text
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
        If Combo2.ListIndex = 2 Then
            'Si el combo1 es 3 entonces es super
            If Combo1.ListIndex = 3 Then
                i = 0
            Else
                i = 1
            End If
        Else
            i = Combo2.ItemData(Combo2.ListIndex)
        End If
        SQL = SQL & "' , nivelariges =" & i
        'SQL = SQL & "  , login = '" & Text2(2).Text
        SQL = SQL & "  , passwordpropio = '" & Text2(3).Text & "'"
        
        
        'El e-mail
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For i = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(i).Text & "|"
            Next i
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        SQL = SQL & " ,dirfich = " & CadenaDesdeOtroForm
        i = -1
        If Combo4.ListIndex >= 0 Then i = Combo4.ItemData(Combo4.ListIndex)
        SQL = SQL & " ,skinAriges = " & i
        
        
        'aqui, en lugar del selecteditem tengo k pasarle el codigo de usuario
        'ya que cuando es nuevo usario y cojo los datos desde otra aplicacion entonces
        'no lo tengo selected y enonces peta
        
        SQL = SQL & " WHERE codusu = " & CodigoUsuario
        Set LOG = New cLOG
        LOG.Insertar 41, vUsu, "[MODIFICAR] " & SQL
        Set LOG = Nothing
        
        If UsuarioOrigen <= 0 Then
            Sql2 = "update menus_usuarios set "
            Select Case Combo2.ItemData(Combo2.ListIndex)
                Case 0 'super
                    Sql2 = Sql2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 1 'administrador
                    Sql2 = Sql2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 2 'normal
                    Sql2 = Sql2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 3 'consulta
                    Sql2 = Sql2 & " ver=1, creareliminar=0, modificar=0, imprimir=1, especial=0"
            End Select
            Sql2 = Sql2 & " where codusu = " & DBSet(CodigoUsuario, "N")
            Sql2 = Sql2 & " and aplicacion in ('ariges') "
        Else
            Sql2 = "DELETE FROM menus_usuarios WHERE codusu = " & CodigoUsuario
            conn.Execute Sql2
            
            'Preparo el INSERT
            Sql2 = "INSERT INTO menus_usuarios (codusu,codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido,textovisible,vericono) "
            Sql2 = Sql2 & " SELECT " & CodigoUsuario & ",codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido,textovisible,vericono FROM menus_usuarios WHERE codusu = " & UsuarioOrigen
            
        End If
        
        conn.Execute Sql2
        
        
        Excepcion = ""
        ' dependiendo de si es Superusuario, Administrador, Normal o consulta
        Select Case Combo2.ItemData(Combo2.ListIndex)
            Case 0 'superusuario
                
            Case 1 'administrador
                Excepcion = "(1)"
            Case 2 'normal
                Excepcion = "(1,10,12,13,14)"
            Case 3 'consulta
                Excepcion = "(1,10,12,13,14)"
        End Select
        
        If Excepcion <> "" Then
            Sql2 = "update menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0"
            Sql2 = Sql2 & " where aplicacion in ('ariges') and codusu = " & DBSet(CodigoUsuario, "N")
            Sql2 = Sql2 & " and (codigo in " & Excepcion
            Sql2 = Sql2 & " or codigo in (select codigo from menus where padre in " & Excepcion & " and aplicacion in ('ariges')))"
            
            conn.Execute Sql2
        End If
        
    End If
    conn.Execute SQL
    
    
    CadenaDesdeOtroForm = ""
    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub



Private Sub cmdUsu_Click(Index As Integer)
Dim K As Integer

    Select Case Index
    Case 0, 1
        Limpiar Me
        
        If Index = 0 Then
            'Nuevo usuario
            CargaCombo6 0
            
            Label6.Caption = "NUEVO"
            i = 0 'Para el foco
            
            Combo2.ListIndex = -1
            Combo4.ListIndex = -1
        Else
            
            CargaCombo6 ListView1.SelectedItem
            
            'Modificar0
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                'LimpiarCamposUsuario
                PonerDatosUsuario
            End If
            i = 1 'Para el foco
            FrameUsuario.Tag = 0  'Marcamos que es una modificacion desde un usuario existente
        End If
        Text2(0).Enabled = (Index = 0)
        
        Me.FrameNormal.visible = False
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.visible = True
        Me.FrameUsuario.Enabled = True
        Me.FrameEditorMenus.visible = False
        Me.FrameEditorMenus.Enabled = False
        
'        If Not vEmpresa.TieneTesoreria Then Me.ChkSoloTesoreria.Value = 0
        
        Text2(i).SetFocus
        
    Case 2, 3
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        i = vUsu.Codigo Mod 1000
        If i = CInt(ListView1.SelectedItem.Text) Then
            MsgBox "El usuario es el mismo con el que esta trabajando actualmente", vbInformation
            Exit Sub
        End If
        
        If Index = 2 Then
            
            SQL = "El usuario " & ListView1.SelectedItem.SubItems(1) & " será eliminado y no tendra acceso a los programas de Ariadna (Ariconta, ariges....) ." & vbCrLf
            SQL = SQL & vbCrLf & "                              ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            SQL = "DELETE from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            
        Else
            SQL = "Al usuario " & ListView1.SelectedItem.SubItems(1) & " no le estará permitido el acceso al programa Ariges." & vbCrLf
            SQL = SQL & vbCrLf & "                              ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            SQL = "UPDATE usuarios.usuarios SET nivelariges = -1 WHERE codusu = " & ListView1.SelectedItem.Text
        End If
        Screen.MousePointer = vbHourglass
        
        conn.Execute SQL
    
    
        Set LOG = New cLOG
        SQL = IIf(Index = 2, "ELIMINAR", "QUITAR ACCESO")
        SQL = "[" & SQL & "]    " & ListView1.SelectedItem.SubItems(1)
        LOG.Insertar 41, vUsu, SQL
        Set LOG = Nothing
    
    
    
        '//El codigo siguiente seria mas logico meterlo en el modulo de usuario
        '   pero de momento un saco de cemento
        If Index = 2 Then EliminarAuxiliaresUsuario CInt(ListView1.SelectedItem.Text)
    
        CargaUsuarios
        
        Screen.MousePointer = vbDefault
    
        Me.FrameNormal.visible = True
        Me.FrameNormal.Enabled = True
        Me.FrameUsuario.visible = False
        Me.FrameUsuario.Enabled = False
        Me.FrameEditorMenus.visible = False
        Me.FrameEditorMenus.Enabled = False
    
    End Select

End Sub

Private Sub AbrirRsEmpresas()
      Set rsEmpresasAriges = New ADODB.Recordset
      rsEmpresasAriges.Open "Select * from usuarios.empresasariges ORDER BY codempre", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
      
End Sub

Private Sub CerrarRsEmpresas()
    rsEmpresasAriges.Close
    Set rsEmpresasAriges = Nothing
End Sub


Private Sub EliminarAuxiliaresUsuario(CodUsu As Long)

    On Error GoTo EEliminarAuxiliaresUsuario
    SQL = "DELETE FROM usuarios.usuarioempresasariges where codusu =" & CodUsu
    conn.Execute SQL
    
    SQL = "DELETE FROM usuarios.appmenususuario where  codusu =" & CodUsu
    conn.Execute SQL
    
    AbrirRsEmpresas
    While Not rsEmpresasAriges.EOF
        SQL = "DELETE FROM  " & rsEmpresasAriges!AriGes & ".menus_usuarios where codusu = " & CodUsu
        conn.Execute SQL
        rsEmpresasAriges.MoveNext
    Wend
    CerrarRsEmpresas
    
    
    Exit Sub
EEliminarAuxiliaresUsuario:
    MuestraError Err.Number, "Eliminar Auxiliares Usuario"

End Sub

Private Sub PonerDatosUsuario()
        
     Text2(0).Text = miRsAux!Login
     Text2(1).Text = miRsAux!nomusu
     Text2(2).Text = miRsAux!passwordpropio
     Text2(3).Text = miRsAux!passwordpropio
     i = miRsAux!nivelariges

    Select Case i
        Case 0
            i = 3
        Case 1
            i = 2
        Case 2
            i = 1
        Case 3
            i = 0
    End Select

    Combo2.ListIndex = i
    
    i = -1
    If Not IsNull(miRsAux!skinariges) Then i = miRsAux!skinariges
    PosicionarCombo Combo4, i
     
     'Cargamos los datos del correo e-mail
     SQL = Trim(DBLet(miRsAux!Dirfich, "T"))
     If SQL <> "" Then
         For i = 1 To 4
             Text2(3 + i).Text = RecuperaValor(SQL, i)
         Next i
     End If
     
'     Me.ChkSoloTesoreria.Value = DBLet(miRsAux!SoloTesoreria)

End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo3_Click(Index As Integer)
    If Not PrimeraVez Then DatosUsusario
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaUsuarios
    End If
    FrameEditorMenus.visible = False
    LeerEditorMenus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 14
        .Buttons(6).Image = 14
    End With
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 28
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 26
    End With
    
    CargaCombo
   
    
'    Me.ChkSoloTesoreria.visible = vEmpresa.TieneTesoreria
'    Me.ChkSoloTesoreria.Enabled = vEmpresa.TieneTesoreria
    
    Me.ListView1.SmallIcons = ImageList1
    Me.ListView2.SmallIcons = ImageList1
    Me.FrameUsuario.visible = False
    Me.FrameNormal.Enabled = True
    LeerDatosCombo True
    
    
    imgQuitarSkin.visible = vUsu.Id = 0 'root
    
End Sub


Private Sub LeerDatosCombo(Leer As Boolean)
Dim Cad2 As String

' hay que hacer que funcione con combo como la contabilidad

'    On Error GoTo ELe
'    If Leer Then
'
'        Combo3(1).ListIndex = 0
'        I = vControl.UltAccesoBDs  'RecuperaValor(CadenaControl, 6)
'        Combo3(1).ListIndex = I
'    Else
'        'GUARDAR
'        vControl.UltAccesoBDs = Combo3(1).ListIndex
'        vControl.Grabar
'
'            CadenaControl = InsertaValor(CadenaControl, 6, Combo3(1).ListIndex)
'
'    End If
'    Exit Sub
'ELe:
'    Err.Clear
End Sub

Private Sub CargaUsuarios()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    SQL = "Select * from usuarios.usuarios where nivelariges >=0 "
    
    ' solo vemos root si somos root
    If vUsu.Login = "root" Then
        SQL = SQL & " and codusu >= 0 order by codusu"
    Else
        SQL = SQL & " and codusu > 0 order by codusu"
    End If
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set Itm = ListView1.ListItems.Add
        Itm.Text = miRsAux!CodUsu
        Itm.SubItems(1) = miRsAux!Login
        If miRsAux!nivelariges = 0 Then
            Itm.SmallIcon = 4
        Else
            Itm.SmallIcon = 5
        End If
        'Nombre y nivel de usuario
        SQL = "-1"
        If Not IsNull(miRsAux!skinariges) Then SQL = miRsAux!skinariges
        SQL = miRsAux!nivelariges & "|" & miRsAux!nomusu & "|" & SQL & "|"
        
        Itm.Tag = SQL
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub


Private Sub DatosUsusario()
Dim ItmX As ListItem
On Error GoTo EDatosUsu

    ListView2.ListItems.Clear
    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Combo5.ListIndex = -1
        Exit Sub
    End If
    
    Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
    'NIVEL
    SQL = RecuperaValor(ListView1.SelectedItem.Tag, 1)
    '                           COMBO                      en Bd
    '                       0.- Consulta                     3
    '                       1.- Normal                       2
    '                       2.- Administrador                1
    '                       3.- SuperUsuario (root)          0
    If Not IsNumeric(SQL) Then SQL = 3
    Select Case Val(SQL)
    Case 2
        Combo1.ListIndex = 1
    Case 1
        Combo1.ListIndex = 2
    Case 0
        Combo1.ListIndex = 3
    Case Else
        Combo1.ListIndex = 0
    End Select
    
    
    'SQL = DevuelveValor("select skinariges from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text)
    SQL = RecuperaValor(ListView1.SelectedItem.Tag, 3)
    If SQL = "" Then SQL = "-1"
    If SQL = "-1" Then
        Combo5.ListIndex = -1
    Else
        PosicionarCombo Combo5, CInt(Val(SQL))
    End If
    
    SQL = "select  empresasariges.codempre,nomempre,nomresum, usuarioempresasariges.codempre bloqueada "
    SQL = SQL & " from usuarios.empresasariges left join usuarios.usuarioempresasariges on empresasariges.codempre = usuarioempresasariges.codempre And "
    SQL = SQL & "  (usuarioempresasariges.codusu = " & ListView1.SelectedItem.Text & " Or codusu Is Null)"
    SQL = SQL & "   WHERE (1=1)"
    
    If Combo3(1).ListIndex > 0 Then
        SQL = SQL & " AND "
        If Combo3(1).ListIndex = 1 Then SQL = SQL & " NOT "
        SQL = SQL & " (usuarioempresasariges.codempre is null) "
    End If
    
    '[Monica] sólo empresas de ariges nuevas
    SQL = SQL & " and empresasariges.ariges like 'ariges%' "
    
    SQL = SQL & " ORDER BY 1  "
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView2.ListItems.Add
        ItmX.Text = miRsAux.Fields(0)
        ItmX.SubItems(1) = miRsAux!nomempre
        ItmX.SubItems(2) = miRsAux!nomresum
        If miRsAux.Fields(0) > 100 Then
            ItmX.ForeColor = &H808080
            ItmX.ListSubItems(1).ForeColor = &H808080
            ItmX.ListSubItems(2).ForeColor = &H808080
        End If
        
        If IsNull(miRsAux!bloqueada) Then
            ItmX.SmallIcon = 1
        Else
            ItmX.SmallIcon = 2
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LeerDatosCombo False
End Sub

Private Sub imgQuitarSkin_Click()
    Combo4.ListIndex = -1
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim AsignarDatos As Boolean

    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        If UCase(Label6.Caption) = "NUEVO" Then
        
            'Si es nuevo entonces, primero compruebo que no existe el login
            'Si existe, y el usuario tiene nivel conta >=0 entonces
            ' existe en la conta. Si existe pero el nivel conta es -1 entonces
            'lo que hacemos es ponerle los datos y que cambie la opcion de nivel usu
            SQL = "Select * from usuarios.usuarios where login='" & Text2(0).Text & "'"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'Tiene nivel usu
                If miRsAux!nivelariges > 0 Then
                    MsgBox "El usuario ya existe para Ariges", vbExclamation
                    LimpiarCamposUsuario
                    Text2(0).SetFocus
                    
                Else
                    If miRsAux!CodUsu = 0 Then
                        MsgBox "Esta intentando modificar datos del usuario ADMINISTRADOR", vbCritical
                        AsignarDatos = False
                    Else
                        SQL = "El usuario existe para otras aplicaciones de Ariadna Software." & vbCrLf
                        SQL = SQL & "¿Desea agregarlo como usuario a Ariges?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then AsignarDatos = True
                    End If
                    If AsignarDatos Then
                        PonerDatosUsuario
                        'Ahora pongo el label y el campo a disbled
                        Text2(1).SetFocus
                        Label6.Caption = "MODIFICAR"
                        Text2(0).Enabled = False
                        FrameUsuario.Tag = miRsAux!CodUsu 'Pongo el frame al codigo ndel usuario
                    Else
                        LimpiarCamposUsuario
                        Text2(0).SetFocus
                    End If
                End If
            End If
            miRsAux.Close
        End If
    End If
    
End Sub

Private Sub LimpiarCamposUsuario()
    For i = 0 To 7
        Text2(i).Text = ""
    Next i
End Sub


Private Sub LeerEditorMenus()
    On Error GoTo ELeerEditorMenus
    cmdConfigMenu.visible = vUsu.Nivel < 1
    
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' insertar
'            UsuarioOrigen = 0
            cmdUsu_Click (0)
        Case 2 ' modificar
            cmdUsu_Click (1)
        Case 3 ' eliminar
            cmdUsu_Click (2)
        Case 5 ' prohibir acceso
            cmdUsu_Click (3)
        Case 6 ' copiar menus
            If Not ListView1.SelectedItem Is Nothing Then
'                UsuarioOrigen = ListView1.SelectedItem
                cmdUsu_Click (0)
            End If
        Case Else
        
    End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un usuario", vbExclamation
        Exit Sub
    End If

    frmMensajes.OpcionMensaje = 99
    frmMensajes.Parametros = ListView1.SelectedItem.Text
    frmMensajes.Show vbModal
    
    DatosUsusario

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
'            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If Node.Children > 0 Then Recursivo2 Node.Child, Node.Checked
End Sub


Private Sub CheckarNodo(N As Node, Valor As Boolean)
Dim NO As Node
    Set NO = N.LastSibling
    Do
        N.Checked = Valor
        If N.Children > 0 Then CheckarNodo N, Valor
        If N.Next <> NO.LastSibling Then Set N = N.Next
    Loop Until NO = N
End Sub

Private Sub Recursivo2(ByVal Nod As Node, Valor As Boolean)
Dim nx As Node
Dim Aux

    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
      End If
    Set nx = Nothing
End Sub


Private Sub GuardarMenuUsuario()
    SQL = "DELETE from usuarios.appmenusUsuario where aplicacion='Ariges' AND codusu =" & ListView1.SelectedItem.Text
    conn.Execute SQL
    
    i = 0
    SQL = "INSERT INTO usuarios.appmenususuario (aplicacion, codusu, codigo, tag) VALUES ('Ariges'," & ListView1.SelectedItem.Text & ","
    RecursivoBD TreeView1.Nodes(1)
End Sub

Private Sub InsertaBD(vtag As String)
Dim C As String
    i = i + 1
    'SQL = "INSERT INTO appmenususuario (aplicacion, codusu, codigo, tag)
    C = SQL & i & ",'" & vtag & "')"
    conn.Execute C
End Sub


Private Sub RecursivoBD(ByVal Nod As Node)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
      End If
    Set nx = Nothing
End Sub

Private Sub CargaCombo()
    
    'nivel
    Combo2.Clear
    
    Combo2.AddItem "Consulta"
    Combo2.ItemData(Combo2.NewIndex) = 3
    
    Combo2.AddItem "Normal"
    Combo2.ItemData(Combo2.NewIndex) = 2
    
    Combo2.AddItem "Administrador"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Superusuario"
    Combo2.ItemData(Combo2.NewIndex) = 0


    '3 ID_OPTIONS_STYLEBLACK2010
    '2 S_STYLESILVER2010
    '1ID_OPTIONS_STYLEBLUE2010

    'skin
    Combo5.Clear
    
    Combo5.AddItem "Office 2010 Blue"
    Combo5.ItemData(Combo5.NewIndex) = 1
    
    Combo5.AddItem "Office 2010 Silver"
    Combo5.ItemData(Combo5.NewIndex) = 2
    
    Combo5.AddItem "Office 2010 Black"
    Combo5.ItemData(Combo5.NewIndex) = 3
    
    
    'skin
    Combo4.Clear
    
    Combo4.AddItem "Office 2010 Blue"
    Combo4.ItemData(Combo4.NewIndex) = 1
    
    Combo4.AddItem "Office 2010 Silver"
    Combo4.ItemData(Combo4.NewIndex) = 2
    
    Combo4.AddItem "Office 2010 Black"
    Combo4.ItemData(Combo4.NewIndex) = 3


End Sub


Private Sub CargaCombo6(Usuario As Integer)
Dim SQL As String
Dim RS As ADODB.Recordset

    'skin
    Combo6.Clear
    
    SQL = "select codusu, login from usuarios.usuarios where codusu <> " & DBSet(Usuario, "N") & " and login <> 'root' and nivelariges > -1 "
    '[Monica]15/10/2019: cambiamos el orden por nombre de usuario
    SQL = SQL & " order by 2 " 'order by 1"
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Combo6.AddItem "Ninguno"
    Combo6.ItemData(Combo6.NewIndex) = 0
    
    While Not RS.EOF
        Combo6.AddItem RS.Fields(1).Value
        Combo6.ItemData(Combo6.NewIndex) = RS.Fields(0).Value
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
End Sub



