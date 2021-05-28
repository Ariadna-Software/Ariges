VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacCliEnvDpto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dpto"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11790
   Icon            =   "frmFacCliEnvDpto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3735
      TabIndex        =   47
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   48
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   45
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   46
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
      Left            =   10035
      TabIndex        =   44
      Top             =   315
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   4785
      Width           =   11520
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
         Left            =   9000
         TabIndex        =   17
         Top             =   120
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
         Left            =   10320
         TabIndex        =   18
         Top             =   120
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
         Left            =   10305
         TabIndex        =   39
         Top             =   135
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame Frame1 
         Height          =   540
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   2655
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
            TabIndex        =   38
            Top             =   195
            Width           =   2355
         End
      End
   End
   Begin VB.Frame FrameDirecciones 
      Caption         =   "Direcciones"
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
      Height          =   3915
      Left            =   90
      TabIndex        =   19
      Top             =   855
      Width           =   11595
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
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   2745
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
         Index           =   16
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Zona|N|S|||sdirec|codzona||N|"
         Text            =   "Text3"
         Top             =   2745
         Width           =   810
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
         Left            =   10095
         TabIndex        =   40
         Tag             =   "CODCLIEN|N|N|0||sdirec|codclien|000|S|"
         Text            =   "Text3"
         Top             =   165
         Width           =   1260
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
         Height          =   1515
         Index           =   14
         Left            =   7200
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   16
         Tag             =   "Obs|T|S|||sdirenvio|observa||N|"
         Text            =   "frmFacCliEnvDpto.frx":000C
         Top             =   2280
         Width           =   4170
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
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código Direc./Dpto|N|N|0|9999|sdirec|coddirec|0000|S|"
         Text            =   "Text3"
         Top             =   360
         Width           =   835
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Domicilio|T|N|||sdirec|domdirec||N|"
         Text            =   "Text3"
         Top             =   1170
         Width           =   3915
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Población|T|N|||sdirec|pobdirec||N|"
         Text            =   "Text3"
         Top             =   1965
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
         Index           =   5
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Provincia|T|N|||sdirec|prodirec||N|"
         Text            =   "Text3"
         Top             =   2370
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
         Index           =   7
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Teléfono|T|S|||sdirec|teldirec||N|"
         Text            =   "Text3"
         Top             =   1080
         Width           =   1650
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
         Left            =   7200
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Persona Contacto|T|S|||sdirec|perdirec||N|"
         Text            =   "Text3"
         Top             =   720
         Width           =   4170
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
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "e-mail|T|S|||sdirec|maidirec||N|"
         Text            =   "Text3"
         Top             =   1785
         Width           =   4170
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Nombre Direc./Dpto|T|N|||sdirec|nomdirec||N|"
         Text            =   "Text3"
         Top             =   765
         Width           =   3945
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
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fax|T|S|||sdirec|faxdirec||N|"
         Text            =   "Text3"
         Top             =   1425
         Width           =   1650
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
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "C.Postal|T|N|||sdirec|codpobla||N|"
         Text            =   "Text3"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Frame FrameCtaBanDpto 
         Height          =   930
         Left            =   5490
         TabIndex        =   20
         Top             =   2205
         Width           =   5850
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
            Left            =   270
            MaxLength       =   4
            TabIndex        =   11
            Tag             =   "IBAN|T|S|||sdirec|iban|||"
            Text            =   "Text"
            Top             =   405
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
            Index           =   13
            Left            =   3765
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Cuenta Bancaria|T|S|||sdirec|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   405
            Width           =   1650
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
            Left            =   3105
            MaxLength       =   2
            TabIndex        =   14
            Tag             =   "Dígito Control|T|S|||sdirec|digcontr|00||"
            Text            =   "Text1"
            Top             =   405
            Width           =   450
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
            Left            =   2025
            MaxLength       =   4
            TabIndex        =   13
            Tag             =   "Sucursal|N|S|0|9999|sdirec|codsucur|0000|N|"
            Text            =   "Text"
            Top             =   405
            Width           =   690
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
            Index           =   10
            Left            =   1170
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "Código Banco|N|S|0|9999|sdirec|codbanco|0000|N|"
            Text            =   "Text"
            Top             =   405
            Width           =   690
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
            Index           =   1
            Left            =   270
            TabIndex        =   43
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
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
            Left            =   3765
            TabIndex        =   24
            Top             =   165
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
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
            Index           =   33
            Left            =   3105
            TabIndex        =   23
            Top             =   165
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
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
            Left            =   2025
            TabIndex        =   22
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
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
            Index           =   47
            Left            =   1170
            TabIndex        =   21
            Top             =   165
            Width           =   675
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1125
         Picture         =   "frmFacCliEnvDpto.frx":0012
         Tag             =   "-1"
         ToolTipText     =   "Buscar población "
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
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
         Left            =   270
         TabIndex        =   41
         Top             =   2745
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1125
         ToolTipText     =   "Buscar población"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   58
         Left            =   5520
         TabIndex        =   35
         Top             =   2265
         Width           =   1545
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
         Index           =   21
         Left            =   270
         TabIndex        =   34
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
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
         Left            =   270
         TabIndex        =   33
         Top             =   360
         Width           =   1095
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
         Index           =   23
         Left            =   270
         TabIndex        =   32
         Top             =   1170
         Width           =   915
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
         Index           =   24
         Left            =   270
         TabIndex        =   31
         Top             =   1560
         Width           =   795
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
         Index           =   25
         Left            =   270
         TabIndex        =   30
         Top             =   1965
         Width           =   975
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
         Index           =   26
         Left            =   270
         TabIndex        =   29
         Top             =   2370
         Width           =   975
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
         Index           =   28
         Left            =   5520
         TabIndex        =   28
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Pers. Contacto"
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
         Index           =   27
         Left            =   5520
         TabIndex        =   27
         Top             =   720
         Width           =   1545
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
         Index           =   10
         Left            =   5520
         TabIndex        =   26
         Top             =   1785
         Width           =   765
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   30
         Left            =   5520
         TabIndex        =   25
         Top             =   1425
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   3240
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacCliEnvDpto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DireccionesEnvio As Boolean
Public codClien As Long
Public NomClien As String
Public VerDatoDpto As Integer  'Si trae valor es que situaremos en el registro


Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmB As frmBasico2
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1


'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

Private CadenaConsulta2 As String
Private Ordenacion2 As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim PrimVez As Boolean






Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim b As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If Me.DireccionesEnvio Then
                    b = InsertarModificarLineaEnvio
                Else
                    b = InsertarModificarLineaDpto
                End If
                
                    
                If b Then
                    If Data1.Recordset Is Nothing Then
                        If Data1.Recordset.EOF Then b = False
                    End If
                    
                    If Not b Then
                        PonerModo 2
                        BotonVerTodos
                        Exit Sub
                    Else
                        CadenaConsulta2 = " where codclien = " & Me.codClien & " AND " & Ordenacion2 & " = " & Text1(0).Text
                        CadenaConsulta2 = "Select * from " & NombreTabla & CadenaConsulta2
                        Data1.RecordSource = CadenaConsulta2
                    End If
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
            
      
        
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
Dim Hay As Boolean
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
        
        Case 3 'Insertar
            Hay = False
            If Not Data1.Recordset Is Nothing Then
                If Not Data1.Recordset.EOF Then Hay = True
            End If
            If Not Hay Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            

        
    End Select
  
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, Ordenacion2, "codclien = " & Me.codClien)
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
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
    
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia "codclien = " & Me.codClien
    Else
        CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE codclien = " & Me.codClien & " ORDER BY " & Ordenacion2
        PonerCadenaBusqueda
    End If

End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
   
    
    PonerModo 4
    PonerFoco Text1(1)
End Sub





Private Sub cmdRegresar_Click()
Dim cad As String

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
    
        PonerModo 2
        If Not Data1.Recordset.EOF Then Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    Else
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
    
        cad = Data1.Recordset.Fields(1) & "|"
        cad = cad & Data1.Recordset.Fields(2) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub



Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        BotonVerTodos
        
        If Me.VerDatoDpto >= 0 Then
            If SituarData(Data1, Ordenacion2 & "=" & Me.VerDatoDpto, Me.lblIndicador) Then PonerCampos
        Else
            If Data1.Recordset.EOF Then BotonBuscar
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    PrimVez = True
    Me.Icon = frmPpal.Icon
    Me.imgBuscar(1).Picture = Me.imgBuscar(0).Picture
    ' ICONITOS DE LA BARRA
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
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    LimpiarCampos   'Limpia los campos TextBox
    
    
    Text1(15).Top = 8000
    Text1(15).Locked = True
    
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    
  
    '## A mano
    If Not DireccionesEnvio Then
        NombreTabla = DevuelveTextoDepto(False)
        Caption = NombreTabla
        FrameDirecciones.Caption = NombreTabla
        NombreTabla = "sdirec"
        Ordenacion2 = " coddirec"
        NumRegElim = 3195
        Text1(2).MaxLength = 60
    Else
        Caption = "Dir. envio"
        NombreTabla = "sdirenvio"
        Ordenacion2 = " coddiren"
        NumRegElim = 3915
    End If
    Caption = Caption & "  Cliente: " & UCase(NomClien) & "(" & codClien & ")"
    FrameDirecciones.Height = NumRegElim
    Frame2.Top = Me.FrameDirecciones.Top + FrameDirecciones.Height + 120
    Me.Height = Frame2.Top + Frame2.Height + 720
    NumRegElim = 0
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codclien=-1"
    'Data1.Refresh
    
    PonerTags
    'BotonBuscar
    


End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta2 = CadenaSeleccion
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        'Poblacion
        Text1(4).Text = ObtenerPoblacion(Text1(3).Text, CadenaSeleccion)
        'provincia
        Text1(5).Text = CadenaSeleccion
End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    Text1(16).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If Index = 0 Then
        Set frmCP = New frmCPostal
        frmCP.DatosADevolverBusqueda = "0"
        frmCP.Show vbModal
        Set frmCP = Nothing
    Else
        'zona
        Set frmZ = New frmFacZonas
        frmZ.DatosADevolverBusqueda = "0"
        frmZ.Show vbModal
        Set frmZ = Nothing
        
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub





Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    'If BLOQUEADesdeFormulario(Me) Then BotonModificar
   
    BotonModificar


    
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





Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then cadB = cadB & " AND "
    cadB = cadB & " codclien = " & Me.codClien

    'Reemplazamos tabla
    If Me.DireccionesEnvio Then cadB = Replace(cadB, "sdirec.", "sdirenvio.")
        
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        
        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " ORDER BY  " & Ordenacion2
        PonerCadenaBusqueda
    End If
End Sub




Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda." & vbCrLf & Caption, vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & vbCrLf & Caption, vbInformation
            Me.lblIndicador.Caption = ""
            If Modo = 0 Then PonerModo 0
        End If
        
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
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1

    Text2.Text = ""
    If Text1(16).Text <> "" Then Text2.Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text1(16).Text, "N")
    
                
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

 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 5 Then lblIndicador.Caption = "Lineas dto"
    
    
   
    b = Modo < 5
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)

    cmdRegresar.visible = b
  
      
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset Is Nothing Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera b Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    Text1(15).Locked = True  'Siempre bloqueado
    
    
    imgBuscar(0).visible = Modo = 1 Or Modo > 2
    imgBuscar(1).visible = Modo = 1 Or Modo > 2
    BloquearChecks Me, Modo
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
                        
    If Modo <= 2 Then PonerFocoChk chkVistaPrevia
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
On Error Resume Next

    b = (Modo = 0 Or Modo = 2)

    'Añadir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = Modo = 2
    
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    mnEliminar.Enabled = b
    
     '---------------------------------
    b = (Modo >= 3 Or Modo = 1)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear

    Toolbar1.Buttons(8).Enabled = False


End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim I As Integer

    If Modo = 3 Then Text1(15).Text = codClien
    
    
    For I = 10 To 13
        If Text1(I).Text <> "" Then
            If IsNumeric(Text1(I).Text) Then
                If Val(Text1(I).Text) = "0" Then Text1(I).Text = ""
            End If
        End If
    Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    If Text1(16).Text = "" Then
        MsgBox "Indique la zona", vbExclamation
        Exit Function
    End If
    
    'MAYO 2011 . Dia 20
    'Si pone cta bancaria comprobaremos qu esta bien puesta
    'Si ha puesto entidad DEBE completar la cuenta bancaria
    If Text1(13).Text <> "" Then
        For I = 11 To 13
            If Text1(I).Text = "" Then Exit For
        Next
        If I <= 13 Then
            'Se ha salido
            MsgBox "Faltan datos para la cuenta bancaria", vbExclamation
            b = False
        Else
            b = Comprueba_CuentaBan2(Text1(10).Text & Text1(11).Text & Text1(12).Text & Text1(13).Text, False)
            If Not b Then
                If MsgBox("Cuenta bancaria incorrecta.    ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then b = True
            End If
        End If
    End If
    
    
    
    DatosOk = b
End Function






Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 0 'poblacion
            Case 16: KEYBusqueda KeyAscii, 1 'zona
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
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
Dim cto As Byte
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo
'            If Text1(Index).Text <> "" Then
             If PonerFormatoEntero(Text1(Index)) Then

            End If

        Case 3
            If Text1(Index).Text <> "" Then
                If IsNumeric(Text1(Index)) Then
                    Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, CadenaConsulta2)
                    Text1(Index + 2).Text = CadenaConsulta2
                    CadenaConsulta2 = ""
                End If
            End If
        Case 10 To 13
            If Text1(Index).Text <> "" Then
                If Not PonerFormatoEntero(Text1(Index)) Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                Else
                    
                    If Index = 13 Then
                        cto = 10
                    Else
                        If Index = 12 Then
                            cto = 2
                        Else
                            cto = 4
                        End If
                    End If
                    
                    Text1(Index).Text = Right(String("0", cto) & Text1(Index).Text, cto)
                    
                                        
                    If Index = 13 Then
                        
                           
                           CadenaDesdeOtroForm = Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text & Me.Text1(13).Text
                       
                           If Len(CadenaDesdeOtroForm) = 20 Then
                               DevuelveIBAN2 "ES", CadenaDesdeOtroForm, CadenaDesdeOtroForm
                               If Len(CadenaDesdeOtroForm) = 2 Then
                                   CadenaDesdeOtroForm = "ES" & CadenaDesdeOtroForm
                                   If Me.Text1(17).Text = "" Then
                                       Text1(17).Text = CadenaDesdeOtroForm
                                   Else
                                       If Me.Text1(17).Text <> CadenaDesdeOtroForm Then MsgBox "Codigo IBAN distinto del calculado [" & CadenaDesdeOtroForm & "]", vbExclamation
                                   End If
                               End If
                           End If
                           CadenaDesdeOtroForm = ""
            
                    
                    End If
                    
                    
                End If
            End If
        Case 16
            cto = 0
            Text2.Text = ""
            If PonerFormatoEntero(Text1(Index)) Then
                Text2.Text = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", Text1(Index).Text, "N")
                If Text2.Text = "" Then
                    MsgBox "No existe la zona", vbExclamation
                    Text1(16).Text = ""
                    PonerFoco Text1(16)
                End If
            Else
                Text1(16).Text = ""
            End If
    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
                mnNuevo_Click
        Case 2  'Modificar
                mnModificar_Click
        Case 3
          mnEliminar_Click
        
        Case 5: mnBuscar_Click
        Case 6: mnVerTodos_Click
'        Case 11: mnSalir_Click
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


Private Sub PonerBotonCabecera(b As Boolean)

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then PonerFocoBtn Me.cmdRegresar
    
    cmdCancelar.Cancel = True
    
    
   
    Me.cmdRegresar.visible = b
    
    
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear

    
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(" & Ordenacion2 & "=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub





Private Sub PonerTags()

    If Not DireccionesEnvio Then
        
        Text1(0).Tag = "Código Direc./Dpto|N|N|0|999|sdirec|coddirec|000|S|"
        Text1(1).Tag = "Nombre Direc./Dpto|T|N|||sdirec|nomdirec|||"
        Text1(2).Tag = "Domicilio|T|N|||sdirec|domdirec|||"
        Text1(3).Tag = "C.Postal|T|N|||sdirec|codpobla|||"
        Text1(4).Tag = "Población|T|N|||sdirec|pobdirec|||"
        Text1(5).Tag = "Provincia|T|N|||sdirec|prodirec|||"
        Text1(6).Tag = "Persona Contacto|T|S|||sdirec|perdirec|||"
        Text1(7).Tag = "Teléfono|T|S|||sdirec|teldirec|||"
        Text1(8).Tag = "Fax|T|S|||sdirec|faxdirec|||"
        Text1(9).Tag = "e-mail|T|S|||sdirec|maidirec|||"
        Text1(10).Tag = "Código Banco|N|S|0|9999|sdirec|codbanco|0000||"
        Text1(11).Tag = "Sucursal|N|S|0|9999|sdirec|codsucur|0000||"
        Text1(12).Tag = "Dígito Control|T|S|||sdirec|digcontr|00||"
        Text1(13).Tag = "Cuenta Bancaria|T|S|||sdirec|cuentaba|0000000000||"
        
        
        'codclien
        Text1(15).Tag = "CODCLIEN|N|N|0||sdirec|codclien|000|S|"
    Else
        
        Text1(0).Tag = "Código|N|N|0|9999|sdirenvio|coddiren|0000|S|"
        Text1(1).Tag = "Nombre Direc|T|N|||sdirenvio|nomdiren|||"
        Text1(2).Tag = "Domicilio|T|N|||sdirenvio|domdiren|||"
        Text1(3).Tag = "C.Postal|T|N|||sdirenvio|codpobla|||"
        Text1(4).Tag = "Población|T|N|||sdirenvio|pobdiren|||"
        Text1(5).Tag = "Provincia|T|N|||sdirenvio|prodiren|||"
        Text1(6).Tag = "Persona Contacto|T|S|||sdirenvio|perdiren|||"
        Text1(7).Tag = "Teléfono|T|S|||sdirenvio|teldiren|||"
        Text1(8).Tag = "Fax|T|S|||sdirenvio|faxdiren|||"
        Text1(14).Tag = "Obs|T|S|||sdirenvio|observa|||"
        'codclien
        Text1(15).Tag = "CODCLIEN|N|N|0||sdirenvio|codclien|000|S|"
    End If
    
    
    'Cuales estaran visibles
    Text1(14).visible = Me.DireccionesEnvio
    For NumRegElim = 9 To 13
        Text1(NumRegElim).visible = Not Me.DireccionesEnvio
    Next
    FrameCtaBanDpto.visible = Not DireccionesEnvio
End Sub




Private Sub MandaBusquedaPrevia(cadB As String)
'--
'        CadenaConsulta2 = ParaGrid(Text1(0), 10, "Código")
'        CadenaConsulta2 = CadenaConsulta2 & ParaGrid(Text1(1), 75, "Nombre")
'
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = CadenaConsulta2
'        CadenaConsulta2 = ""
'        frmB.vTabla = NombreTabla
'        frmB.vSQL = cadB
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        If Not DireccionesEnvio Then
'            frmB.vTitulo = "Direcciones envio"
'        Else
'            frmB.vTitulo = "Dpto. cliente " & Me.NomClien
'        End If
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri
'        frmB.vCargaFrame = False
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
        
        
        Dim vTitulo As String
        Dim vDesc As String
        If DireccionesEnvio Then
            vTitulo = "Direcciones envio"
            vDesc = "Dirección"
        Else
            vTitulo = "Departamento cliente " & Me.NomClien
            vDesc = "Departamento"
        End If
        CadenaConsulta2 = ""
        
        Set frmB = New frmBasico2
        AyudaMantenimientosAux frmB, vTitulo, vDesc, Text1(0), cadB
        Set frmB = Nothing
        
        If CadenaConsulta2 <> "" Then
            CadenaConsulta2 = RecuperaValor(CadenaConsulta2, 1)
            CadenaConsulta2 = " codclien = " & Me.codClien & " AND " & Ordenacion2 & " = " & CadenaConsulta2
            CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE " & CadenaConsulta2
            PonerCadenaBusqueda
        
        Else   'es decir NO ha devuelto datos
            If Modo = 0 Then
                CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE codclien =" & Me.codClien
                Data1.RecordSource = CadenaConsulta2
                Data1.Refresh
                
                PonerModo 0
            Else
                PonerFoco Text1(kCampo)
            End If
        End If
        CadenaConsulta2 = ""
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub


    
    If Not PuedeEliminarDirecEnvio(DireccionesEnvio, CStr(Me.codClien), CInt(Me.Text1(0).Text)) Then Exit Sub
    
    
    '### a mano
    cad = "¿Seguro que desea eliminar?"
    cad = cad & vbCrLf & "Codigo : " & Data1.Recordset.Fields(1)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(2)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
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




Private Function InsertarModificarLineaDpto() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLineaDpto = False
    SQL = ""
    If Modo = 3 Then 'INSERTAR
        
            SQL = "INSERT INTO sdirec (codclien,coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba) VALUES ("
            SQL = SQL & codClien & ", "
            SQL = SQL & Text1(0).Text
            For I = 1 To 5
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T")
            Next I
                    
            For I = 6 To 13 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next I
                        
            SQL = SQL & ")"
        
        
    Else
        'MODIFICAR
        
            SQL = "UPDATE sdirec Set nomdirec = " & DBSet(Text1(1).Text, "T")
            SQL = SQL & ", domdirec = " & DBSet(Text1(2).Text, "T")
            SQL = SQL & ", codpobla = " & DBSet(Text1(3).Text, "T")
            SQL = SQL & ", pobdirec = " & DBSet(Text1(4).Text, "T")
            SQL = SQL & ", prodirec = " & DBSet(Text1(5).Text, "T")
            SQL = SQL & ", perdirec = " & DBSet(Text1(6).Text, "T")
            'If text1(7).Text <> "" Then SQL = SQL & ", fechainv = '" & Format(text1(7).Text, "yyyy-mm-dd") & "'"
            'If text1(8).Text <> "" Then SQL = SQL & ", horainve = '" & Format(text1(8).Text, "hh:mm:ss") & "'"
            SQL = SQL & ", teldirec = " & DBSet(Text1(7).Text, "T")
            SQL = SQL & ", faxdirec = " & DBSet(Text1(8).Text, "T")
            SQL = SQL & ", maidirec = " & DBSet(Text1(9).Text, "T")
            'datos cuenta bancaria
            If Me.FrameCtaBanDpto.visible Then
                SQL = SQL & ", codbanco = " & DBSet(Text1(10).Text, "N", "S")
                SQL = SQL & ", codsucur = " & DBSet(Text1(11).Text, "N", "S")
                SQL = SQL & ", digcontr = " & DBSet(Text1(12).Text, "T")
                SQL = SQL & ", cuentaba = " & DBSet(Text1(13).Text, "T")
            End If
            
            SQL = SQL & " WHERE codclien =" & codClien & " AND "
            SQL = SQL & " coddirec =" & (Text1(0).Text)
        
    End If
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaDpto = True
        TratarDptoEnTesoreria   'TESOERIA
    Else
        PonerFoco Text1(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar(Modificar) Direcciones/Departamentos" & vbCrLf & Err.Description
End Function
    


Private Function InsertarModificarLineaEnvio() As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLineaEnvio = False
    SQL = ""
   If Modo = 3 Then
        
            SQL = "INSERT INTO sdirenvio (codclien,coddiren,nomdiren,domdiren,codpobla,pobdiren,prodiren,perdiren,teldiren,faxdiren,observa) VALUES ("
            SQL = SQL & codClien & ", "
            SQL = SQL & Text1(0).Text
            For I = 1 To 5
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T")
            Next I
                    
            For I = 6 To 8 'campos opcionales
                SQL = SQL & ", "
                SQL = SQL & DBSet(Text1(I).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next I
                        
            SQL = SQL & "," & DBSet(Text1(I).Text, "T", "S") & ")"
 
        
    Else
            SQL = "UPDATE sdirenvio Set nomdiren = " & DBSet(Text1(1).Text, "T")
            SQL = SQL & ", domdiren = " & DBSet(Text1(2).Text, "T")
            SQL = SQL & ", codpobla = " & DBSet(Text1(3).Text, "T")
            SQL = SQL & ", pobdiren = " & DBSet(Text1(4).Text, "T")
            SQL = SQL & ", prodiren = " & DBSet(Text1(5).Text, "T")
            SQL = SQL & ", perdiren = " & DBSet(Text1(6).Text, "T")
            SQL = SQL & ", teldiren = " & DBSet(Text1(7).Text, "T")
            SQL = SQL & ", faxdiren = " & DBSet(Text1(8).Text, "T")
            SQL = SQL & ", observa = " & DBSet(Text1(14).Text, "T")
            SQL = SQL & " WHERE codclien =" & codClien & " AND "
            SQL = SQL & " coddiren =" & (Text1(0).Text)
    End If
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLineaEnvio = True
    Else
        PonerFoco Text1(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Direcciones de envio" & vbCrLf & Err.Description
End Function




Private Function TratarDptoEnTesoreria() As Boolean
Dim Existe As Boolean
Dim C As String
    
    C = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", CStr(codClien))
    
    If C = "" Then Exit Function
    
    CadenaConsulta2 = DevuelveDesdeBDNew(conConta, "cuentas", "codmacta", "codmacta", C)
    If CadenaConsulta2 = "" Then
        MsgBox "No existe la cuenta contable del cliente " & NomClien, vbExclamation
        Exit Function
    End If


    Existe = False
    CadenaConsulta2 = "codmacta = '" & C & "' and Dpto "
    CadenaConsulta2 = DevuelveDesdeBD(conConta, "descripcion", "departamentos", CadenaConsulta2, Text1(0).Text)
    If CadenaConsulta2 <> "" Then Existe = True
    
    
    If Existe Then
        
        'UPDATEAMOS
        CadenaConsulta2 = "UPDATE  departamentos set Descripcion = " & DBSet(Text1(1).Text, "T")
        CadenaConsulta2 = CadenaConsulta2 & " WHERE codmacta= '" & C & "' AND Dpto = " & Text1(0).Text
    Else
        'NO EXISTE... creamos
        CadenaConsulta2 = "insert into `departamentos` (`codmacta`,`Dpto`,`Descripcion`) values ('"
        CadenaConsulta2 = CadenaConsulta2 & C & "'," & Text1(0).Text & "," & DBSet(Text1(1).Text, "T") & ")"
        
    End If
    ConnConta.Execute CadenaConsulta2
    
End Function

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
