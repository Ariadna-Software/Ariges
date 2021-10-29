VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComFacturarGR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Compra Proveedores"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   16080
   Icon            =   "frmComFacturarGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   61
      Top             =   90
      Width           =   2460
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   62
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Albaranes"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Facturas"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Grid"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFactura 
      Height          =   5760
      Left            =   10095
      TabIndex        =   17
      Top             =   3240
      Width           =   5865
      Begin VB.CommandButton cmdIVA 
         Caption         =   "+"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   57
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton cmdIVA 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   56
         Top             =   3075
         Width           =   255
      End
      Begin VB.CheckBox chkTipoRet 
         Caption         =   "Base + IVA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   54
         Top             =   4650
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdIVA 
         Caption         =   "+"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   2655
         Width           =   255
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Cancelar"
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
         Index           =   1
         Left            =   1380
         TabIndex        =   52
         Top             =   4995
         Width           =   1065
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
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
         Index           =   0
         Left            =   255
         TabIndex        =   51
         Top             =   4995
         Width           =   1065
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
         Left            =   3885
         MaxLength       =   12
         TabIndex        =   48
         Tag             =   "Impret|N|S|||scafac|impret|#,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4125
         Width           =   1770
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
         Left            =   1245
         MaxLength       =   5
         TabIndex        =   47
         Tag             =   "PorRet|N|S|0||scafac|PorRet|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   4125
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3930
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1665
         Width           =   1755
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
         Index           =   8
         Left            =   3930
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1080
         Width           =   1755
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
         Left            =   3930
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   660
         Width           =   1755
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
         Left            =   3930
         MaxLength       =   15
         TabIndex        =   37
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   240
         Width           =   1755
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
         Left            =   600
         MaxLength       =   5
         TabIndex        =   35
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3||N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   540
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
         Left            =   600
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2||N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   540
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
         Left            =   600
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1||N|"
         Text            =   "Text1 7"
         Top             =   2655
         Width           =   540
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
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   27
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "000,222,555.25"
         Top             =   2655
         Width           =   1755
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
         Index           =   13
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   26
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2655
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3885
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2655
         Width           =   1755
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
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   1755
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
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   23
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3885
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   1755
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
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   1755
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
         Index           =   15
         Left            =   1245
         MaxLength       =   5
         TabIndex        =   20
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3885
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3480
         Width           =   1755
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   22
         Left            =   3315
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4995
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Importe retención"
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
         Left            =   2025
         TabIndex        =   50
         Top             =   4125
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "%Reten"
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
         Left            =   345
         TabIndex        =   49
         Top             =   4125
         Width           =   825
      End
      Begin VB.Line Line3 
         X1              =   345
         X2              =   5625
         Y1              =   2175
         Y2              =   2175
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
         Index           =   11
         Left            =   3600
         TabIndex        =   46
         Top             =   1080
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   2070
         X2              =   5670
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Dto.Gnral"
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
         Left            =   2055
         TabIndex        =   45
         Top             =   1080
         Width           =   1530
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
         Index           =   99
         Left            =   3600
         TabIndex        =   44
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Dto. PP"
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
         Left            =   2055
         TabIndex        =   43
         Top             =   660
         Width           =   1530
      End
      Begin VB.Line Line1 
         X1              =   345
         X2              =   5625
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Left            =   2055
         TabIndex        =   42
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
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
         Left            =   2055
         TabIndex        =   38
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
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
         Left            =   600
         TabIndex        =   36
         Top             =   2325
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Left            =   2025
         TabIndex        =   32
         Top             =   2325
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Left            =   3885
         TabIndex        =   31
         Top             =   2355
         Width           =   1335
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
         TabIndex        =   30
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   3315
         TabIndex        =   29
         Top             =   4635
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
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
         Index           =   41
         Left            =   1230
         TabIndex        =   28
         Top             =   2325
         Width           =   720
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   2265
      Left            =   120
      TabIndex        =   7
      Top             =   855
      Width           =   15840
      Begin VB.CheckBox chkInvSujePasivo 
         Caption         =   "Inv. sujeto pasivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9945
         TabIndex        =   60
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox chkLlevarContab 
         Caption         =   "Insertar en contabilidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13140
         TabIndex        =   59
         Top             =   1800
         Width           =   2475
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
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1230
         Width           =   7740
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
         Left            =   4815
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   540
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
         Left            =   10755
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1230
         Width           =   4875
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
         Left            =   10755
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   540
         Width           =   4875
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
         Left            =   9960
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1230
         Width           =   780
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
         Left            =   9960
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Operador|N|N|0|9999|scafpc|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   540
         Width           =   780
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   1230
         Width           =   1050
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
         Index           =   1
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   540
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
         Index           =   0
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "00000111112222244444"
         Top             =   540
         Width           =   2670
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmComFacturarGR.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   945
         Width           =   240
      End
      Begin VB.Label lbTipoProve 
         Caption         =   "asda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   270
         TabIndex        =   58
         Top             =   1665
         Width           =   5265
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   6060
         Picture         =   "frmComFacturarGR.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   285
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4245
         Picture         =   "frmComFacturarGR.frx":0A99
         ToolTipText     =   "Buscar fecha"
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   11655
         ToolTipText     =   "Buscar banco propio"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Recepción"
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
         Left            =   4815
         TabIndex        =   15
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Prev. Pago"
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
         Left            =   9975
         TabIndex        =   12
         Top             =   975
         Width           =   1710
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   10935
         ToolTipText     =   "Buscar trabajador"
         Top             =   270
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
         Index           =   1
         Left            =   9945
         TabIndex        =   11
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Left            =   225
         TabIndex        =   10
         Top             =   975
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
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
         Index           =   29
         Left            =   3180
         TabIndex        =   9
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Left            =   240
         TabIndex        =   8
         Top             =   285
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
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
   Begin VB.Frame FrameList 
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   120
      TabIndex        =   55
      Top             =   3165
      Width           =   9920
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmComFacturarGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public Codprove As Long
Public CadenaAlbaran As String






'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmProv As frmBasico2
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmBanPr As frmBasico2 'frmFacBancosPropios 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1



Private Modo2 As Byte
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

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWhere As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------

Dim dtoGn As Currency
Dim dtoPP As Currency
Dim ForPa As Integer


Private vProve As CProveedor


                    
Private Sub chkInvSujePasivo_Click()
    If chkInvSujePasivo.Value = 1 Then
        'Si esta marcado la de retencion, no dejo marcar esta
        If Me.chkTipoRet.Value = 1 Then
            MsgBox "No puede tener retencion y ser ISP", vbExclamation
            chkInvSujePasivo.Value = 0
        End If
    End If
    CalcularDatosFactura
End Sub



Private Sub chkTipoRet_Click()
    Text1_LostFocus 23  'Como si cambaira la retencion
End Sub

Private Sub cmdGenerar_Click(Index As Integer)
Dim N  As Long



    If Index = 1 Then
        'QUITO EL PORCENTAJE
        Text1(23).Text = ""
        Text1(24).Text = ""
        CalcularDatosFactura
        'Le ha dado a cancelar
        PonerModo2 4
        
    Else
        'Aceptar
        If Text1(23).Text = "" Xor Text1(24).Text = "" Then
            MsgBox "Si pone porcentaje retencion debe poner importe(y viceversa)", vbExclamation
            Exit Sub
        End If
        
        If Text1(10).Text = "" Then
            MsgBox "Error tipo IVA 1", vbExclamation
            Exit Sub
        End If
        N = 0
        If Text1(10).Text = Text1(11).Text Or Text1(10).Text = Text1(12).Text Then
            N = 1
        Else
            If Text1(11).Text <> "" And Text1(11).Text = Text1(12).Text Then N = 1
        End If
        If N > 0 Then
            N = 0
            MsgBox "Mismo tipo de IVA. Deben ser IVAS disitntos", vbExclamation
            Exit Sub
        End If
        
        'Si ha puesto
        If vEmpresa.TieneAnalitica Then
            N = Val(DevuelveDesdeBD(conAri, "count(*)", "slialp", cadWhere & " AND codccost is null AND 1 ", "1"))
            If N > 0 Then
                MsgBox "Existen lineas(" & N & ") de albaranes sin asignar centros de coste", vbExclamation
                Exit Sub
            End If
        End If
        
        
        If Val(DevuelveDesdeBD(conAri, "count(*)", "sproveanticipo", "descontado=0 AND codprove", Text1(3).Text)) = 0 Then
            If MsgBox("¿Generar la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        If cadWhere = "" Then
            
            If Me.Codprove > 0 Then
                If Not DatosOk Then Exit Sub
            
            End If
            If cadWhere = "" Then
                MsgBox "Seleccione albaranes", vbExclamation
                Exit Sub
            End If
        End If
       
        
        If GenerarFactura_ Then
            If vParamAplic.InvSujetoPasivo Then Me.chkInvSujePasivo.Value = 0
            BotonPedirDatos
            If Me.Codprove >= 0 Then Unload Me
        Else
            If Me.Codprove >= 0 Then PonerCamposFacturarAlbaran
        End If
    End If
End Sub

Private Sub cmdIVA_Click(Index As Integer)
Dim Impor As Currency
    'Poner nuevo tipo de IVA
    Set frmB = New frmBuscaGrid
    CadenaDesdeOtroForm = ""
    frmB.vCampos = "Código|tiposiva|codigiva|N||20·Denominacion|tiposiva|nombriva|T||60·Porcentaje|tiposiva|porceiva|N||10·"
    frmB.vTabla = "tiposiva"
    frmB.vTitulo = "Tipos de IVA"
    frmB.vDevuelve = "0|2|"
    
    frmB.vselElem = 1
    frmB.vConexionGrid = conConta
    frmB.vCargaFrame = False
    frmB.Show vbModal
    Set frmB = Nothing
    If CadenaDesdeOtroForm <> "" Then
        
        Text1(10 + Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Impor = CCur(RecuperaValor(CadenaDesdeOtroForm, 2))
        Text1(13 + Index).Text = CStr(Impor) '% iva
        NumRegElim = Impor * 100 'Para no decalrar mas variabnles
        Impor = ImporteFormateado(Text1(16 + Index).Text)
        Impor = Round2((Impor * NumRegElim / 10000), 2)
        Text1(19 + Index).Text = CStr(Impor)
        
        PonerFormatoEntero Text1(10 + Index)
        PonerFormatoDecimal Text1(13 + Index), 3
        PonerFormatoDecimal Text1(19 + Index), 3
        RecalculoDeImportes
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbHourglass
        
    If PrimeraVez Then
        PrimeraVez = False
        If Me.Codprove >= 0 Then
            'Esta facturando UN albaran
            PonerCamposFacturarAlbaran
            
        End If
    Else
        
        If VerAlbaranes Then RefrescarAlbaranes
        VerAlbaranes = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next
    
'    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 18   'Pedir Datos
'        .Buttons(2).Image = 43   'Ver albaranes
'        .Buttons(3).Image = 26   'Generar FActura
'        .Buttons(6).Image = 15   'Salir
'    End With
    
    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2

        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 35   'Ver albaranes
        .Buttons(3).Image = 37   'Generar FActura
        .Buttons(5).Image = 30   'ver grid
    End With
    
    
    Toolbar1.Buttons(2).Enabled = Me.Codprove < 0
    
    cadWhere = ""
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
    chkLlevarContab.Value = CheckValueLeer(Name)
    
    chkInvSujePasivo.Value = 0
    Me.chkInvSujePasivo.visible = vParamAplic.InvSujetoPasivo
    
    '## A mano
    NombreTabla = "scafpc" 'cabecera facturas compras a proveedor
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    PonerModo2 0
    
    
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    lbTipoProve.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkLlevarContab.Value
    DesBloqueoManual "RECFAC"
    TerminaBloquear

'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDesdeOtroForm = CadenaDevuelta
End Sub

Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Proveedores
Dim Indice As Byte
    
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Proveedor
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom proveedor
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo2 = 2 Or Modo2 = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmProv = New frmBasico2
'            frmProv.DatosADevolverBusqueda = "0"
'            frmProv.Show vbModal
            AyudaProveedores frmProv, Text1(3)
            Set frmProv = Nothing
            Indice = 3
            
        Case 1 'Operador. Trabajador
            Indice = 4
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
       
       Case 2 'Bancos Propios
            Indice = 5
'            Set frmBanPr = New frmFacBancosPropios
'            frmBanPr.DatosADevolverBusqueda = "0|1|"
'            frmBanPr.Show vbModal
'            Set frmBanPr = Nothing
            Set frmBanPr = New frmBasico2
            AyudaBancosPropios frmBanPr, Text1(Indice)
            Set frmBanPr = Nothing

    End Select
    
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo2 = 2 Or Modo2 = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index + 1
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub



Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Cuando se selecciona un albaran de la lista
Dim i As Integer
Dim Cad As String
Dim TipoFP As Integer 'Forma de pago
Dim TipoDtoPP As Currency 'descuento pronto pago
Dim tipoDtoGn As Currency 'descuento general

    
    
    Set ListView1.SelectedItem = Item
    
    If Item.Checked Then
        If vEmpresa.TieneAnalitica Then
            Cad = "codccost is null AND fechaalb = " & DBSet(Item.SubItems(1), "F")
            Cad = Cad & " AND numalbar = " & DBSet(Item.Text, "T") & " AND codprove"
        
            i = Val(DevuelveDesdeBD(conAri, "count(*)", "slialp", Cad, Text1(3).Text))
            If i > 0 Then
                MsgBox "Lineas de albaran(" & i & ") sin centro de coste asignado", vbExclamation
                Item.Checked = False
                Exit Sub
            End If
         End If
    End If
    
    If Me.Text1(1).Text = "" Then
        MsgBox "Debe indicar la fecha de factura", vbExclamation
        ListView1.SelectedItem.Checked = False
        PonerFoco Text1(1)
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Inicializamos a cero
    TipoFP = 0
    TipoDtoPP = 0
    tipoDtoGn = 0
    
    'cuando seleccionamos un check vemos si lo podemos seleccionar
    'ya que si ya habia algun albaran selecionado tendremos que comprobar
    'que son de la misma forpa, dtoppago y dtognral.
    'si esto no se cumple no se pueden agrupar en la misma factura
    For i = 1 To ListView1.ListItems.Count
        If Item.Index <> i Then
            If ListView1.ListItems(i).Checked Then
                'ya habia otro albaran seleccionado
                TipoFP = ListView1.ListItems(i).SubItems(2)
                TipoDtoPP = CCur(ListView1.ListItems(i).SubItems(4))
                tipoDtoGn = CCur(ListView1.ListItems(i).SubItems(5))
                Exit For
            End If
        End If
    Next i
    
    If Not (TipoFP = 0 And TipoDtoPP = 0 And tipoDtoGn = 0) Then
    'si ya habia un albaran seleccionado, comprobar que es del mismo tipo
        If Item.SubItems(2) <> TipoFP Or Item.SubItems(4) <> TipoDtoPP Or Item.SubItems(5) <> tipoDtoGn Then
            MsgBox "Se debe seleccionar albaranes de la misma Forma de Pago y Descuentos", vbExclamation
            ListView1.SelectedItem.Checked = False
            Screen.MousePointer = vbDefault
            ListView1.SetFocus
            Exit Sub
        End If
    Else
    End If
    
    ' Calculamos los datos de factura
    If Not VerAlbaranes Then CalcularDatosFactura
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnGenerarFac_Click()
    BotonFacturar
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnVerAlbaran_Click()
    If Me.CadenaAlbaran <> "" Then Exit Sub 'Si factura UN unico albaran, NO permitimos
    BotonVerAlbaranes
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo2
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
Dim Impor As Currency
Dim C As String

    If Modo2 <> 5 Then _
        If Not PerderFocoGnral(Text1(Index), Modo2) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                If Index = 1 Then
                    'la fecha de factura debe ser , como mucho 6 años inferior a la fecha actual
                    
                    If DateAdd("yyyy", -5, Now) > CDate(Text1(Index).Text) Then
                        C = "Fecha factura incorrecta. Fuera plazo años presentacion" & vbCrLf & vbCrLf & Text1(Index).Text & vbCrLf & "¿Continuar?"
                        If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then Text1(Index).Text = ""
                    Else
                        If CDate(Text1(Index).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then
                            If MsgBox("Fecha factura mayor que fin de ejercicios" & vbCrLf & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                        End If
                    End If
                    
                Else
                    'Index=2
                    If CDate(Text1(Index).Text) < vEmpresa.FechaIni Or CDate(Text1(Index).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then
                        MsgBox "Fecha de recepción debe estar dentro de ejercicios contables. ", vbExclamation
                        Text1(Index).Text = ""
                    End If
                End If
                If Text1(Index).Text <> "" Then
                    ' No debe existir el número de factura para el proveedor en hco
                    If ExisteFacturaEnHco Then
                        InicializarListView
                    End If
                End If
            End If
            
        Case 3 'Cod Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerDatosProveedor
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
                    DesBloqueoManual ("RECFAC")
                    If Not BloqueoManual("RECFAC", Text1(3).Text) Then
                        MsgBox "No se puede recepcionar factura de ese proveedor. Hay otro usuario recepcionando.", vbExclamation
                        BotonPedirDatos
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        CargarAlbaranes
                    End If
                    
                End If
                
            Else
                Text2(Index).Text = ""
                lbTipoProve.Caption = ""
            End If

        Case 4 'Cod Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index).Text = ""
            End If
            

        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
        Case 23, 24
            'SON EL IMPORTE DE RETENCION y el porcentaje
            Text1(Index).Text = Trim(Text1(Index).Text)
            
            If Not PonerFormatoDecimal(Text1(Index), 3) Then
                    Text1(Index).Text = ""
                   
            End If
                            
            If Index = 23 Then
                If Text1(23).Text <> "" Then
                    'Ha puesto porcentaje retencion
                    Impor = 0
                    For NumRegElim = 0 To 2
                        'Base imponible
                        If Text1(16 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(16 + NumRegElim).Text)
                        
                        'Si solo es sobre la BASE, esto no lo sumo
                        If Me.chkTipoRet.Value Then
                            If Text1(19 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(19 + NumRegElim).Text)
                        End If
                    Next NumRegElim
                    NumRegElim = ImporteFormateado(Text1(23).Text) * 100
                    Impor = Round2((Impor * NumRegElim / 10000), 2)
                    Text1(24).Text = Format(Impor, FormatoImporte)
                    RecalculoDeImportes
                End If
                
            End If
            
            
    End Select
End Sub


'RECALCULO DATOS FACTURA
'-----------------------------------------------------
Private Sub RecalculoDeImportes()
Dim Impor As Currency
        Impor = 0
        For NumRegElim = 0 To 2
            'Base imponible + iva
            If Text1(16 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(16 + NumRegElim).Text)
            If Text1(19 + NumRegElim).Text <> "" Then Impor = Impor + ImporteFormateado(Text1(19 + NumRegElim).Text)
        Next NumRegElim
        'Memos la retencion
        If Text1(24).Text <> "" Then Impor = Impor - ImporteFormateado(Text1(24).Text)
        
        'TOTAL FACTURA
        Text1(22).Text = CStr(Impor)
        PonerFormatoDecimal Text1(22), 3
End Sub




'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'

'       MODO: 0 pidiendo datos encabezado


Private Sub PonerModo2(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean
On Error GoTo EPonerModo


    Modo2 = Kmodo
    
    
    'GEneral
    B = (Modo2 = 5)
    FrameFactura.Enabled = B   'Solo habilitado al final
    Toolbar1.Enabled = Not B
    'Antes. Para que no se quede en gris
    'ListView1.Enabled = Not b
    'para que no se quede el listview en gris
    FrameList.Enabled = Not B
    FrameIntro.Enabled = Not B
    
  
    cmdGenerar(0).visible = B
    cmdGenerar(1).visible = B
    'chkTipoRet.visible = b
    If Not B Then
        cmdIVA(0).visible = False
        cmdIVA(1).visible = False
        cmdIVA(2).visible = False
    End If
        
    If Modo2 < 5 Then
        If vParamAplic.InvSujetoPasivo Then Me.chkInvSujePasivo.visible = True
    End If
        
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo2 = 2)
        
                 
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo2
    
    'Importes siempre bloqueados
    For i = 6 To 22
        BloquearTxt Text1(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF    'Total factura
        
    
    
    If Modo2 = 4 Then
        For i = 0 To 4
            If i <> 2 Then
                Text1(i).Locked = False
                Text1(i).BackColor = vbWhite
            End If
        Next
    End If
        
    If Modo2 = 5 Then
        BloquearTxt Text1(23), False
        BloquearTxt Text1(24), False
        'Si el tipo de proveedor NO es REA
        'y solo tiene un tipo de IVA , podemos dejar que cambie el iva
        'ANTES JUNIO 2010
        '
        '
        'If Text1(11).Text = "" And Text1(12).Text = "" Then
        '    If vProve.TipoProv <> 2 Then cmdIVA.visible = True
        'End If
        
        If vProve.TipoProv <> 2 Then
                cmdIVA(0).visible = True
                cmdIVA(1).visible = Text1(11).Text <> ""
                cmdIVA(2).visible = Text1(12).Text <> ""
        End If
    End If
    
    
    
    '---------------------------------------------
    B = (Modo2 <> 0 And Modo2 <> 2 And Modo2 <> 5)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
                    

       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo2, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As cTag
Dim Cad As String
Dim i As Byte

    On Error GoTo EDatosOK
    DatosOk = False
    
    
    'Permitimos numeros de factrura de 20 carcateres. PERO en contabilidad nueva
    Text1(0).Text = Trim(Text1(0).Text)
    If Not vParamAplic.ContabilidadNueva Then
        If Len(Text1(0).Text) > 10 Then
            Cad = String(70, "*") & vbCrLf
            Cad = Cad & Cad & Cad & vbCrLf
            Cad = Cad & "En Ariconta4 solo se permiten numero de factura de 10 caracteres" & vbCrLf & vbCrLf & "                   Actualice a la versión ARICONTA6" & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            Exit Function
        End If
    End If
    
    
    
    ' deben de introducirse todos los datos del frame
    For i = 0 To 5
        If Text1(i).Text = "" Then
            If Text1(i).Tag <> "" Then
                Set vtag = New cTag
                If vtag.Cargar(Text1(i)) Then
                    Cad = vtag.Nombre
                Else
                    Cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                Cad = "Campo"
                If i = 5 Then Cad = "Cta. Prev. Pago"
            End If
            MsgBox Cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerFoco Text1(i)
            Exit Function
        End If
    Next i
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    'ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text), True)
    ResultadoFechaContaOK = EsFechaOKConta_SinSII(CDate(Text1(2).Text), True)
    
    If ResultadoFechaContaOK > 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Function

    End If
    
    'comprobar que se han seleccionado lineas para facturar
    
    
    
    
    If cadWhere = "" Then
        Cad = "MAL"
        If Me.Codprove >= 0 Then
            'Esta facturando UN albaran
            If ListView1.ListItems(1).Checked Then
                'Solo hay UN albaran
                If Me.Codprove = Val(Text1(3).Text) Then
                    'Es el mismo proveedor. Solo ha cambiado fechas
                    CalcularDatosFactura
                    If cadWhere <> "" Then Cad = ""
                End If
            End If
        End If
        
        If Cad <> "" Then
            MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
            Exit Function
        End If
    End If
    
    
    ' No debe existir el número de factura para el proveedor en hco
    If ExisteFacturaEnHco Then Exit Function
    
    
    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
    Cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    If RegistrosAListar(Cad) > 1 Then
        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
        Exit Function
    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
    Cad = "select distinct (codforpa) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = miRsAux.Fields(0)
    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    Cad = "Select tipforpa from sforpa where codforpa=" & Cad
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        i = 1
        Cad = miRsAux.Fields(0)
        If Val(Cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            i = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If i = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vProve.CuentaBan = "" Or vProve.DigControl = "" Or vProve.Sucursal = "" Or vProve.Banco = "" Then
            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then i = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If i > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
             
        Case 2 'Ver Albaranes
            mnVerAlbaran_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

        Case 5 ' ver grid
            ListView1.GridLines = Not ListView1.GridLines

    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo2, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String


    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
    InicializarListView
    
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWhere = ""
    
    PonerModo2 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
    'poner trabajador conectado como operador
    Text1(4).Text = PonerTrabajadorConectado(Nombre)
    Text2(4).Text = Nombre
    
    '++
    Text1(4).Enabled = (vParamAplic.NumeroInstalacion <> vbHerbelca)
    Text2(4).Enabled = (vParamAplic.NumeroInstalacion <> vbHerbelca)
    imgBuscar(1).Enabled = (vParamAplic.NumeroInstalacion <> vbHerbelca)
    '++hasta aqui
    
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub

Private Sub BotonVerAlbaranes()

    If Not SeleccionaRegistros Then Exit Sub
    
    VerAlbaranes = True
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmComEntAlbaranesGR.cadSelAlbaranes = cadWhere
        frmComEntAlbaranesGR.EsHistorico = False
        frmComEntAlbaranesGR.Show vbModal
        frmComEntAlbaranesGR.cadSelAlbaranes = ""
    Else
        frmComEntAlbaranSA.cadSelAlbaranes = cadWhere
        frmComEntAlbaranSA.EsHistorico = False
        frmComEntAlbaranSA.Show vbModal
        frmComEntAlbaranSA.cadSelAlbaranes = ""
    End If
End Sub
    


Private Sub CargarAlbaranes()
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Const vbPassionRed = &HC0&      '&H1C47F4

On Error GoTo ECargar

    ListView1.ListItems.Clear
    If VerAlbaranes = False Then cadWhere = ""
    
    
    
    'si no hay proveedor salir
    If Text1(3).Text = "" Then Exit Sub
    
    SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
    SQL = SQL & " sum(slialp.importel) as bruto "
    SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
    SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
    SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text
    If Me.Codprove >= 0 Then
        'Esta facturando direcamente UN albaran
        SQL = SQL & " AND " & Me.CadenaAlbaran
    End If
    
    
    SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
    SQL = SQL & " ORDER BY scaalp.numalbar"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    InicializarListView
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = RS!Numalbar
        ItmX.SubItems(1) = Format(RS!FechaAlb, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(RS!codforpa, "000")
        ItmX.SubItems(3) = RS!nomforpa
        ItmX.SubItems(4) = Format(RS!DtoPPago, "#0.00")
        ItmX.SubItems(5) = Format(RS!DtoGnral, "#0.00")
        ItmX.SubItems(6) = Format(RS!bruto, "#,###,#0.00") '(RAFA/ALZIRA) 12092006
        If DBLet(RS!bruto, "N") < 0 Then
            ItmX.ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(1).ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(2).ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(3).ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(4).ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(5).ForeColor = vbPassionRed
            ItmX.ListSubItems.Item(6).ForeColor = vbPassionRed
            ItmX.ToolTipText = "Abono"
        End If
                             
        If Me.Codprove >= 0 Then ItmX.Checked = True
            
                             
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view
    
    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "NºAlbaran", 1900
    ListView1.ColumnHeaders.Add , , "Fecha", 1400, 2
    ListView1.ColumnHeaders.Add , , "FPag", 0
    ListView1.ColumnHeaders.Add , , "Forma de Pago", 2900
    ListView1.ColumnHeaders.Add , , "DtoPP", 850, 2
    ListView1.ColumnHeaders.Add , , "DtoGr", 850, 2
    ListView1.ColumnHeaders.Add , , "Imp. Bruto", 1600, 1

    ListView1.SmallIcons = frmPpal.ImgListPpal


End Sub



Private Sub CalcularDatosFactura()
Dim i As Integer
Dim SQL As String
Dim cadAux As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 6 To 22
         Text1(i).Text = ""
    Next i

    cadAux = ""
    cadWhere = ""
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
        'para cada albaran seleccionado para la factura
            ForPa = ListView1.ListItems(i).SubItems(2)
            dtoPP = ListView1.ListItems(i).SubItems(4)
            dtoGn = ListView1.ListItems(i).SubItems(5)
            SQL = "(numalbar=" & DBSet(ListView1.ListItems(i).Text, "T") & " and "
            SQL = SQL & "fechaalb=" & DBSet(ListView1.ListItems(i).SubItems(1), "F") & ")"
            If cadAux = "" Then
                cadAux = SQL
            Else
                cadAux = cadAux & " OR " & SQL
            End If
        End If
    Next i
    
    If cadAux <> "" Then
    'se han seleccionado albaranes para facturar
    'Esta el la cadena WHERE de los albaranes seleccionados para obtener
    'el bruto de las lineas de los albaranes agrupadas por tipo de iva
        cadWhere = "slialp.codprove=" & Val(Text1(3).Text)
        cadWhere = cadWhere & " AND (" & cadAux & ")"
    Else
        Exit Sub
    End If
    
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWhere) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = dtoPP
    vFactu.DtoGnral = dtoGn
    
        
    vFactu.FijarTipoIvaProveedor Val(Text1(3).Text)
    vFactu.ISP = False
    If Me.chkInvSujePasivo.visible And Me.chkInvSujePasivo.Value = 1 Then vFactu.ISP = True
        
    
    If vFactu.CalcularDatosFactura2(cadWhere, "scaalp", "slialp", CDate(Text1(1).Text), Me.chkInvSujePasivo.Value = 1) Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        
        For i = 6 To 22
            FormateaCampo Text1(i)
        Next i
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For i = 11 To 20 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For i = 12 To 21 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
   
End Sub



Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWhere = "" Then Exit Function
    cadWhere = Replace(cadWhere, "slialp", "scaalp")
    
    SQL = "Select count(*) FROM scaalp"
    SQL = SQL & " WHERE " & cadWhere
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim Cad As String

        Screen.MousePointer = vbHourglass
    
    
    
    Cad = ""
    If Text1(3).Text = "" Or Text2(3).Text = "" Then
        Cad = "Falta proveedor"
    Else
        If Not IsNumeric(Text1(3).Text) Then Cad = "Campo proveedor debe ser numérico"
        If Text1(4).Text = "" Or Text2(4).Text = "" Then Cad = Cad & "- Operador"
                
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
        
        
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(Text1(3).Text) Then
        Set vProve = Nothing
        Exit Sub
    End If
        
    If vParamAplic.InvSujetoPasivo Then
        If vProve.TipoProv = 1 Then Me.chkInvSujePasivo.visible = False
    End If
    
    
    If Not DatosOk Then Exit Sub
    
    
     If Me.Codprove > 0 Then
        If cadWhere = "" Then
            If Not DatosOk Then Exit Sub
            If cadWhere = "" Then
                ListView1.ListItems(1).Checked = False
                MsgBox "Selecione un albaran", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    
    PonerModo2 5
End Sub


Private Function GenerarFactura_() As Boolean
Dim vFactu As CFacturaCom
Dim Cad  As String
Dim CadenaContab As String
Dim TieneAnticiposPendientesDescontar As String

        On Error GoTo Error1
        GenerarFactura_ = False
        
        
        TieneAnticiposPendientesDescontar = ""
        Cad = DevuelveDesdeBD(conAri, "count(*)", "sproveanticipo", "descontado=0 AND codprove", Text1(3).Text)
        If Val(Cad) > 0 Then
            'Lanzareamos pantalla para seleccionar anticipo o anticipos
            ' si cancela, cancelamos
            CadenaDesdeOtroForm = ""
            frmMensajes.OpcionMensaje = 32
            frmMensajes.cadWhere = Text1(3).Text
            frmMensajes.cadWHERE2 = Text2(3).Text
            frmMensajes.Parametros = Text1(22).Text
            frmMensajes.Show vbModal
            
            
            If CadenaDesdeOtroForm = "CANCEL" Then
                Err.Raise 513, , "Proceso cancelado en anticipos proveedor"
            Else
                TieneAnticiposPendientesDescontar = Trim(CadenaDesdeOtroForm)
                If TieneAnticiposPendientesDescontar <> "" Then TieneAnticiposPendientesDescontar = Mid(TieneAnticiposPendientesDescontar, 2)
                    
            End If
            CadenaDesdeOtroForm = ""
        End If
        
        
        
        
        
        
        
        
        
        'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = Text1(3).Text
        vFactu.Numfactu = Text1(0).Text
        vFactu.FecFactu = Text1(1).Text
        vFactu.FecRecep = Text1(2).Text
        vFactu.Trabajador = Text1(4).Text
        vFactu.BancoPr = Text1(5).Text
        vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
        vFactu.ForPago = ForPa
        vFactu.DtoPPago = dtoPP
        vFactu.DtoGnral = dtoGn
        vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
        vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
        vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
        vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
        vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
        vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
        vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
        vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
        vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
        vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
        vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
        vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
        vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
        vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
        vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
        
        'Sobre que calcual la retencion, si sobre el tota o sobre las bases(sun iva)
        If chkTipoRet.Value = 1 Then
            vFactu.TipoRet = 0
        Else
            vFactu.TipoRet = 1
        End If
        
        
        vFactu.PorRet = ImporteFormateado(Text1(23).Text)
        vFactu.ImpRet2 = ImporteFormateado(Text1(24).Text)
        
        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan
        vFactu.Iban = vProve.Iban
        
        vFactu.ISP = False
        If Me.chkInvSujePasivo.visible And Me.chkInvSujePasivo.Value = 1 Then vFactu.ISP = True
        
        
        '               Check1(0)=inserta tesoreria                             Check1(1) contab
        'If vFactu.TraspasoAlbaranesAFactura(cadWhere, (Check1(0).Value = 1), (Check1(1).Value = 1), False) Then
        If vFactu.TraspasoAlbaranesAFactura(cadWhere, True, False, False, TieneAnticiposPendientesDescontar) Then
        
            '------------------------------------------------------------------------------
            '------------------------------------------------------------------------------
            'Si tiene la marca de contabilizar. La contabilizo
            'ContabilizarFacturas(NomTabla, cadSelect)
            CadenaContab = ""
            If Me.chkLlevarContab.Value = 1 Then
                Espera 0.2
                'Vamos tb a contbilizarla
                frmListado.OpcionListado = 512
                CadenaContab = Text2(3).Text & " (" & Text1(3).Text & ")  | "
                CadenaContab = CadenaContab & Text1(0).Text & "      " & Text1(1).Text & "|"
                
                cadWhere = " scafpc.codprove = " & Val(Text1(3).Text) & " AND scafpc.fecfactu =" & DBSet(Text1(1).Text, "F") & " AND  scafpc.numfactu = " & DBSet(Text1(0).Text, "T")
                
                frmListado.NumCod = CadenaContab
                frmListado.CadTag = cadWhere
                frmListado.Show vbModal
                CadenaContab = DevuelveDesdeBD(conAri, "codigo1", "tmpinformes", "codusu", vUsu.Codigo)
                If CadenaContab = "" Then
                    
                    
                    'Como no se ha contabilizado, pero si que quiero que ponga el vencimiento, lo tenog quieu poner a manao
                    'dav07 @ 01/03/2015
                    Cad = Text1(0).Text & " @ " & Text1(1).Text
                    Cad = DBSet(Cad, "T")
                    Cad = vUsu.Codigo & ",0," & Cad & "," & Text1(3).Text & ")"
                    Cad = "INSERT INTO tmpinformes(codusu,codigo1,nombre1 ,importe1) VALUES (" & Cad
                    ejecutar Cad, True
                    
                    
                    CadenaContab = " Numregis ERROR. No contabizada"
                Else
                    CadenaContab = " Numregis " & CadenaContab
                End If
            Else
                'NO contabiliza
                'Con lo cual METER el registro en la tmp a mano
                conn.Execute "delete from tmpinformes WHERE codusu =" & vUsu.Codigo
                Espera 0.1
                'codusu,codigo1,nombre1 ,importe1
                
                'dav07 @ 01/03/2015
                Cad = Text1(0).Text & " @ " & Text1(1).Text
                Cad = DBSet(Cad, "T")
                Cad = vUsu.Codigo & ",0," & Cad & "," & Text1(3).Text & ")"
                Cad = "INSERT INTO tmpinformes(codusu,codigo1,nombre1 ,importe1) VALUES (" & Cad
                conn.Execute Cad
                Espera 0.2
            End If
            
            
            
            'Marzo 2015. Mostrara en un FORM el numregis asignado y
            ' fecvecni e impvenci
            ' pudiendo cambiar este ultimo
            MensajeNUmeroRegistroYVencimientos
            
            
            
            
            '------------------------------------------------------------------------------
            '  LOG de acciones
            Set LOG = New cLOG
            
            'ALBARANES
            
            Cad = Text1(3).Text & "-" & Mid(Text2(3).Text, 1, 15) & "..." & vbCrLf
            Cad = Cad & "FRA: " & Text1(0).Text & "   " & Text1(1).Text & "  ALB:"
            For NumRegElim = 1 To Me.ListView1.ListItems.Count
                If ListView1.ListItems(NumRegElim).Checked Then Cad = Cad & ListView1.ListItems(NumRegElim).Text & ":"
            Next
            Cad = Cad & CadenaContab
            Cad = Cad & vbCrLf & "B." & Text1(6).Text
            
            NumRegElim = 0
            If Text1(7).Text <> "" Then
                If ImporteFormateado(Text1(7).Text) <> 0 Then
                    Cad = Cad & "  DtoPP " & Text1(7).Text
                    NumRegElim = 1
                End If
            End If
            If Text1(8).Text <> "" Then
                If ImporteFormateado(Text1(8).Text) <> 0 Then
                    Cad = Cad & "  DtoGE " & Text1(8).Text
                    NumRegElim = 1
                End If
            End If
            If NumRegElim = 1 Then Cad = Cad & "  B.I." & Text1(NumRegElim).Text
            Cad = Cad & vbCrLf
                
            For NumRegElim = 0 To 2
                If Text1(10 + NumRegElim).Text <> "" Then
                    Cad = Cad & NumRegElim + 1 & ":- " & Text1(10 + NumRegElim).Text & "(" & Text1(13 + NumRegElim).Text & ")"
                    Cad = Cad & "     " & Text1(16 + NumRegElim).Text & "    " & Text1(19 + NumRegElim).Text & vbCrLf
                End If
            Next
            If Text1(23).Text <> "" Then Cad = Cad & "Ret:" & Text1(23).Text & "   " & Text1(24).Text & vbCrLf
            Cad = Cad & vbCrLf & "TOTAL: " & Text1(22).Text
            If TieneAnticiposPendientesDescontar <> "" Then Cad = Cad & vbCrLf & "Anticipos: " & TieneAnticiposPendientesDescontar
            LOG.Insertar 9, vUsu, Cad
            Set LOG = Nothing
            
            
            
            
        
        
            'Antes
            'BotonPedirDatos
            'AHora
            LimpiarCampos
            Me.ListView1.ListItems.Clear
            PonerModo2 0
            
            
            If vParamAplic.NumeroInstalacion = 1 Then
    
                Cad = DevuelveDesdeBD(conAri, "count(*)", "sflotas", "codprove", CStr(vProve.Codigo))
    
                If Val(Cad) > 0 Then
                    frmComCasarAlbaranes.Codprove = vProve.Codigo
                    frmComCasarAlbaranes.Show vbModal
                End If
            End If
    
            
            
            GenerarFactura_ = True
            
        Else
            
        End If
        
        
        
        
        
        Set vFactu = Nothing
        Set vProve = Nothing
        
 

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim Cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco
    Cad = "SELECT count(*) FROM scafpc "
    Cad = Cad & " WHERE codprove=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(Cad) > 0 Then
        MsgBox "Factura de proveedor ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function

Private Sub RefrescarAlbaranes()
Dim i As Integer
Dim SQL As String
Dim Itm As ListItem
Dim RS As ADODB.Recordset
    

    For i = 1 To ListView1.ListItems.Count
        SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
        SQL = SQL & " sum(slialp.importel) as bruto "
        SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
        SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text & " AND scaalp.numalbar=" & DBSet(ListView1.ListItems(i).Text, "T") & " AND scaalp.fechaalb=" & DBSet(ListView1.ListItems(i).SubItems(1), "F")
        SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
        SQL = SQL & " ORDER BY scaalp.numalbar"

        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        If Not RS.EOF Then 'Actualizamos los datos de este item en el list
            ListView1.ListItems(i).SubItems(2) = RS!codforpa
            ListView1.ListItems(i).SubItems(3) = RS!nomforpa
            ListView1.ListItems(i).SubItems(4) = RS!DtoPPago
            ListView1.ListItems(i).SubItems(5) = RS!DtoGnral
            ListView1.ListItems(i).SubItems(6) = RS!bruto

        End If
        
        If ListView1.ListItems(i).Checked Then 'comprobamos otra vez el chek y recalculamos factura
            Set Itm = ListView1.ListItems(i)
            ListView1_ItemCheck Itm
        End If

        RS.Close
        Set RS = Nothing
    Next i
    
    'recalcular el total de la factura
     For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            CalcularDatosFactura
            Exit For
        End If
     Next i
     
End Sub





Private Function PonerDatosProveedor() As String
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select nomprove,sprove.codbanpr,nombanpr,tipprove from sprove ,sbanpr where sprove.codbanpr= sbanpr.codbanpr  and sprove.codprove =" & Text1(3).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Devolvemos el nombre del prove y fijamos la cadena del banco
    If miRsAux.EOF Then
        PonerDatosProveedor = ""
        Text1(5).Text = ""
        Text2(5).Text = ""
        lbTipoProve.Caption = ""
        
    Else
        PonerDatosProveedor = miRsAux!nomprove
        Text1(5).Text = miRsAux!codbanpr
        Text2(5).Text = miRsAux!nombanpr
        Select Case miRsAux!tipprove
        Case 1
            lbTipoProve.Caption = "Intracomunitario"
            lbTipoProve.ForeColor = &HC0&      '&H80&
            
        Case 2
            lbTipoProve.Caption = "Extranjero"
            lbTipoProve.ForeColor = &H800000
            
        Case Else
            lbTipoProve.Caption = ""
        End Select
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function




Private Sub MensajeNUmeroRegistroYVencimientos()
    
    Espera 0.25
    frmListado.OpcionListado = 514
    frmListado.Show vbModal
    
End Sub



Private Sub PonerCamposFacturarAlbaran()
    
    'Datos Cabecera
    BotonPedirDatos
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(3).Text = Codprove
    Text2(3).Text = PonerDatosProveedor
       
    'Datos Albaranes
    CargarAlbaranes
    
    If Me.ListView1.ListItems.Count > 0 Then ListView1_ItemCheck ListView1.ListItems(1)
    CalcularDatosFactura
    PonerFoco Text1(0)
End Sub
