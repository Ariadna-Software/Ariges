VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacBancosPropios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos Propios"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12360
   Icon            =   "frmFacBancosPropios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame Frame2 
      Caption         =   "Datos Contables"
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
      Height          =   2085
      Left            =   135
      TabIndex        =   46
      Top             =   4680
      Width           =   12075
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Cta. Contable|T|N|||sbanpr|codmacta||N|"
         Text            =   "Text1"
         Top             =   360
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
         Index           =   15
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Cta. Efectos negociados|T|N|||sbanpr|codmact1||N|"
         Text            =   "Text1"
         Top             =   765
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
         Index           =   17
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Cta. Gastos Tajeta|T|N|||sbanpr|codmact3||N|"
         Text            =   "Text1"
         Top             =   1575
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
         Index           =   16
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Cta. Gastos remesas|T|N|||sbanpr|codmact2||N|"
         Text            =   "Text1"
         Top             =   1170
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
         Index           =   16
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   1170
         Width           =   7395
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   1575
         Width           =   7395
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   765
         Width           =   7395
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   360
         Width           =   7395
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Gastos Remesas"
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
         Left            =   135
         TabIndex        =   54
         Top             =   1170
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Gastos Tarjeta"
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
         TabIndex        =   53
         Top             =   1575
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Efectos negociados"
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
         Left            =   135
         TabIndex        =   52
         Top             =   765
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
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
         Left            =   135
         TabIndex        =   51
         Top             =   360
         Width           =   1770
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2850
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   2850
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   2850
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2850
         Picture         =   "frmFacBancosPropios.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   405
         Width           =   240
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   44
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   45
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
      TabIndex        =   42
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   225
         TabIndex        =   43
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
      Left            =   10350
      TabIndex        =   41
      Top             =   360
      Width           =   1575
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
      Left            =   1725
      MaxLength       =   4
      TabIndex        =   11
      Tag             =   "IBAN|T|S|||sbanpr|iban|||"
      Text            =   "Text1"
      Top             =   4020
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
      Index           =   18
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Identif. Norma 34|T|S|||sbanpr|idnorma34||N|"
      Text            =   "Text1"
      Top             =   1845
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
      Index           =   0
      Left            =   1725
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código de Banco Propio|N|N|0|9999|sbanpr|codbanpr|0000|S|"
      Text            =   "Text1"
      Top             =   1035
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
      Index           =   1
      Left            =   1725
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominación del Banco Propio|T|N|||sbanpr|nombanpr||N|"
      Text            =   "Text1"
      Top             =   1440
      Width           =   3660
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
      Left            =   1725
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Domicilio del Banco Propio|T|S|||sbanpr|dombanpr||N|"
      Text            =   "Text1"
      Top             =   1845
      Width           =   3660
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
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Código Postal del Banco Propio|N|S|||sbanpr|codpopr||N|"
      Text            =   "Text1"
      Top             =   2250
      Width           =   765
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
      Left            =   1725
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Población|T|S|||sbanpr|pobbanpr||N|"
      Text            =   "Text1"
      Top             =   2670
      Width           =   3660
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
      Left            =   1725
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Teléfono|T|S|||sbanpr|telbanpr||N|"
      Text            =   "Text1"
      Top             =   3075
      Width           =   1365
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
      Left            =   7080
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Web|T|S|||sbanpr|wwwbanpr||N|"
      Text            =   "Text1"
      Top             =   3075
      Width           =   4965
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
      Left            =   7065
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Persona de Contacto|T|S|||sbanpr|perbanpr||N|"
      Text            =   "Text1"
      Top             =   2250
      Width           =   4965
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
      Left            =   7080
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "Identif. Cedente|T|S|||sbanpr|identrem||N|"
      Text            =   "Text1"
      Top             =   1440
      Width           =   1365
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
      Left            =   7080
      MaxLength       =   40
      TabIndex        =   9
      Tag             =   "eMail|T|S|||sbanpr|maibanpr||N|"
      Text            =   "Text1"
      Top             =   2670
      Width           =   4965
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
      Left            =   4590
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "Cuenta Bancaría|T|S|||sbanpr|cuentaba|0000000000||"
      Text            =   "Text1"
      Top             =   4020
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
      Index           =   12
      Left            =   4095
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "Dígito Control|T|S|||sbanpr|digcontr|00||"
      Text            =   "Text1"
      Top             =   4020
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
      Index           =   11
      Left            =   3270
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "Sucursal|N|S|0|9999|sbanpr|codsucur|0000|N|"
      Text            =   "Text1"
      Top             =   4020
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
      Index           =   10
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "Código Banco|N|N|0|9999|sbanpr|codbanco|0000|N|"
      Text            =   "IBAN"
      Top             =   4020
      Width           =   645
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
      Left            =   11010
      TabIndex        =   21
      Top             =   7020
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   135
      TabIndex        =   23
      Top             =   6810
      Width           =   3000
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
         Left            =   240
         TabIndex        =   24
         Top             =   165
         Width           =   2595
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
      Left            =   11025
      TabIndex        =   22
      Top             =   7020
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
      Left            =   9840
      TabIndex        =   20
      Top             =   7020
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   4275
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
      Index           =   19
      Left            =   270
      TabIndex        =   40
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Id. Norma 34"
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
      Left            =   5505
      TabIndex        =   39
      Top             =   1898
      Width           =   1440
   End
   Begin VB.Image ImgMail 
      Height          =   240
      Index           =   0
      Left            =   6795
      Picture         =   "frmFacBancosPropios.frx":0A0E
      Tag             =   "-1"
      ToolTipText     =   "Enviar e-mail"
      Top             =   2715
      Width           =   240
   End
   Begin VB.Image imgWeb 
      Height          =   255
      Left            =   6795
      Picture         =   "frmFacBancosPropios.frx":0F98
      Stretch         =   -1  'True
      Tag             =   "-1"
      ToolTipText     =   "Abrir web"
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Left            =   1410
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   2295
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Web"
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
      Left            =   5505
      TabIndex        =   38
      Top             =   3128
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Pers.Contacto"
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
      Left            =   5505
      TabIndex        =   37
      Top             =   2318
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Identif.Cedente"
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
      Left            =   5505
      TabIndex        =   36
      Top             =   1500
      Width           =   1665
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
      Index           =   14
      Left            =   5505
      TabIndex        =   35
      Top             =   2723
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cta.Bancaria"
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
      Left            =   5085
      TabIndex        =   34
      Top             =   3780
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   270
      TabIndex        =   33
      Top             =   3128
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
      Index           =   11
      Left            =   270
      TabIndex        =   32
      Top             =   2723
      Width           =   1230
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
      Index           =   10
      Left            =   270
      TabIndex        =   31
      Top             =   2318
      Width           =   1005
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
      Index           =   9
      Left            =   270
      TabIndex        =   30
      Top             =   1898
      Width           =   1500
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
      Index           =   4
      Left            =   4500
      TabIndex        =   29
      Top             =   3780
      Visible         =   0   'False
      Width           =   300
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
      Index           =   3
      Left            =   3390
      TabIndex        =   28
      Top             =   3780
      Visible         =   0   'False
      Width           =   930
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
      Index           =   2
      Left            =   2580
      TabIndex        =   27
      Top             =   3780
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
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
      TabIndex        =   26
      Top             =   1493
      Width           =   1545
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
      Left            =   270
      TabIndex        =   25
      Top             =   1035
      Width           =   840
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
Attribute VB_Name = "frmFacBancosPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCtas As frmBasico2 'Form para busquedas
Attribute frmCtas.VB_VarHelpID = -1

Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

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
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim indCodigo As Integer

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                
                
                
                    'Si la ccuenta es CREAR CTA CONTABLE entonces, despues de insertarla, la creamos
                    ComprobarCrearCuentas
                
                
                    If Data1.Recordset.EOF Then 'No estaba cargado Inicialmente
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP
                        Data1.Refresh
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
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Text1(0).Text = Format(SugerirCodigoSiguienteStr("sbanpr", "codbanpr"), "0000")
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
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
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    '### a mano
    Cad = "¿Seguro que desea eliminar el Banco Propio? " & vbCrLf
    Cad = Cad & vbCrLf & "Cod. Banco : " & Format(Data1.Recordset.Fields(0), "0000")
    Cad = Cad & vbCrLf & "Desc. Banco: " & Data1.Recordset.Fields(1)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then     'Borramos
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Cad = "Delete from sbanpr where codbanpr=" & Data1.Recordset!codbanpr
        conn.Execute Cad
'        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
        Screen.MousePointer = vbDefault
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Banco Propio", Err.Description
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'Icono de busqueda
    For i = 1 To imgCuentas.Count - 1
        imgCuentas(i).Picture = imgCuentas(0).Picture
    Next
    imgBuscar.Picture = imgCuentas(0).Picture
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        .Buttons(8).Image = 16  'Imprimir
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


    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "sbanpr"
    Ordenacion = " ORDER BY codbanpr"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codbanpr=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Byte
    
    HaDevueltoDatos = True
    Screen.MousePointer = vbHourglass
    cadB = ""
    Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
    cadB = Aux
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Poblacion
End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'cta contable
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click()
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    'Codigo Postal
    Set frmCP = New frmCPostal
    frmCP.DatosADevolverBusqueda = "0"
    frmCP.Show vbModal
    Set frmCP = Nothing

    PonerFoco Text1(3)
    VieneDeBuscar = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

'    imgCuentas(0).Tag = Index
'    MandaBusquedaPrevia "apudirec='S'"
'    PonerFoco Text1(Index + 14)
'    imgCuentas(0).Tag = -1

    indCodigo = Index + 14

    Set frmCtas = New frmBasico2
    AyudaCtasContables frmCtas, Text1(indCodigo)
    Set frmCtas = Nothing

    PonerFoco Text1(indCodigo)

    Screen.MousePointer = vbDefault
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If Index = 0 Then dirMail = Text1(7).Text
    If LanzaMailGnral(dirMail) Then Espera 2
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2

    LanzaVisorMimeDocumento Me.hwnd, Text1(8).Text


  '  If LanzaHomeGnral(Text1(8).Text) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


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
            'cuentas
            Case 14: KEYBusqueda2 KeyAscii, 0 'cuenta
            Case 15: KEYBusqueda2 KeyAscii, 1 'cuenta efecto negociable
            Case 16: KEYBusqueda2 KeyAscii, 2 'cuenta gastos remesa
            Case 17: KEYBusqueda2 KeyAscii, 3 'cuenta gastos tarjeta
            
            Case 3: KEYBusqueda KeyAscii, 1 'codigo postal
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (Indice)
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click
End Sub

'Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
'    KeyAscii = 0
'    imgFec_Click (indice)
'End Sub
'


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
      
    'en el campo ID de norma 34 no se hace Trim ni nada. Lo q pongan
    If Index = 18 Then Exit Sub
      
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
         Case 0 'Cod. Banco Propio
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    'Detectamos aki si ya existe y no esperamos hasta boton Aceptar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
         Case 3 'CPostal
            If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
            ElseIf Not VieneDeBuscar Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
            End If
            VieneDeBuscar = False
            
         Case 10, 11 'codbanco, codsucursal
            PonerFormatoEntero Text1(Index)
            
         Case 12, 13 'DC, numero cta
            FormateaCampo Text1(Index)
            
         Case 14, 15, 16, 17 'Cuentas
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
            If Text1(Index).Text <> "" And Text2(Index).Text = "" Then PonerFoco Text1(Index)
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
    
    Set frmB = New frmBasico2

    AyudaBancosPropios frmB, , cadB, True

    Set frmB = Nothing

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
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
Dim i As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    For i = 14 To 17
        Text2(i).Text = PonerNombreDeCod(Text1(i), conConta, "cuentas", "nommacta", "codmacta", , "T")
    Next i
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, CLng(NumReg)
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1

    
    '------------------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa botones de la Toolbar segun el Modo
Dim B As Boolean
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    mnEliminar.Enabled = B
    Toolbar1.Buttons(8).Enabled = B
    
    B = B Or Modo = 0
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    

End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    If Not Comprueba_CuentaBan2(Text1(10).Text & Text1(11).Text & Text1(12).Text & Text1(13).Text, True) Then
        If MsgBox("La cuenta bancaria no es correcta. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then B = False
        
    Else
            CadenaConsulta = Text1(10).Text & Text1(11).Text & Text1(12).Text & Text1(13).Text
            If Len(CadenaConsulta) = 20 Then
               
                
                    
                If DevuelveIBAN2("", CadenaConsulta, CadenaConsulta) Then
                    If Me.Text1(19).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(19).Text = "ES" & CadenaConsulta
                    Else
                        If Mid(Text1(19).Text, 3) <> CadenaConsulta Then
                            CadenaConsulta = "Calculado : ES" & CadenaConsulta
                            CadenaConsulta = "Introducido: " & Me.Text1(19).Text & vbCrLf & CadenaConsulta & vbCrLf
                            CadenaConsulta = "Error en codigo IBAN" & vbCrLf & CadenaConsulta & "Continuar?"
                            If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then Exit Function
                        End If
                    End If
                End If
                        
            End If
            CadenaConsulta = ""
        
    End If
 
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If

    DatosOk = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5: mnBuscar_Click  'Buscar
        Case 6: mnVerTodos_Click  'Todos
            
        Case 1: mnNuevo_Click  'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar
            
        Case 8: PrintNou
    End Select
End Sub


Private Sub PrintNou()

    With frmImprimir2
        .cadTabla2 = "sbanpr"
        .Informe2 = "rFacBancos.rpt"
        indCodigo = InStr(1, Data1.RecordSource, " WHERE ")
        If indCodigo > 0 Then
            CadenaConsulta = Mid(Data1.RecordSource, indCodigo + 7)
        Else
            CadenaConsulta = ""
        End If
        
        indCodigo = InStr(1, CadenaConsulta, " ORDER BY ")
        If indCodigo > 0 Then CadenaConsulta = Mid(CadenaConsulta, 1, indCodigo)
        If CadenaConsulta <> "" Then
            .cadRegSelec = SQL2SF_David(CadenaConsulta, "sbanpr|")
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
        .OtrosParametros2 = "|pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With

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
Dim Cad As String
Dim Indicador As String

    Cad = "(codbanpr=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE codbanpr= " & Text1(0).Text
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub ComprobarCrearCuentas()
Dim C As String
Dim i As Integer

    On Error Resume Next
    
    For i = 14 To 17
        If Text2(i).Text = vbCrearNuevaCta Then
            C = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci) VALUES ('"
            C = C & Text1(i).Text & "','" & DevNombreSQL(Text1(1).Text) & "','S',0,'" & DevNombreSQL(Text1(1).Text) & "')"
            ConnConta.Execute C
        End If
    Next i
           
    For i = 14 To 17
        If Text2(i).Text = vbCrearNuevaCta Then
            C = PonerNombreCuenta(Text1(i), 2)
            If C = "" Then
                MsgBox "Error en la cuenta: " & Text1(i).Text, vbExclamation
            Else
                Text2(i).Text = C
            End If
        End If
    Next i
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
