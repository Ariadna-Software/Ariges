VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   14940
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEstVentasFam 
      Height          =   7605
      Left            =   480
      TabIndex        =   445
      Top             =   0
      Width           =   7035
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Por nombre proveedor"
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   581
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Comparativo"
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   580
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Frame FrameDetalleFacturacion 
         Height          =   3615
         Left            =   12000
         TabIndex        =   692
         Top             =   3120
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   158
            Left            =   1440
            MaxLength       =   16
            TabIndex        =   456
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   158
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   720
            Text            =   "Text5"
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   157
            Left            =   1440
            MaxLength       =   16
            TabIndex        =   455
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   157
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   717
            Text            =   "Text5"
            Top             =   240
            Width           =   3855
         End
         Begin VB.OptionButton optDetalleFacturacion 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   458
            Top             =   1440
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optDetalleFacturacion 
            Caption         =   "Tipo factura"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   459
            Top             =   1800
            Width           =   1695
         End
         Begin MSComctlLib.ListView lwFact 
            Height          =   2415
            Left            =   2400
            TabIndex        =   457
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4260
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tipo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripci�n"
               Object.Width           =   4233
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   58
            Left            =   720
            TabIndex        =   721
            Top             =   600
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   96
            Left            =   1200
            Picture         =   "frmListadoOfer.frx":000C
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   57
            Left            =   720
            TabIndex        =   719
            Top             =   240
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   95
            Left            =   1200
            Picture         =   "frmListadoOfer.frx":010E
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   56
            Left            =   360
            TabIndex        =   718
            Top             =   0
            Width           =   405
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   4
            Left            =   2160
            Picture         =   "frmListadoOfer.frx":0210
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   5
            Left            =   2160
            Picture         =   "frmListadoOfer.frx":035A
            Top             =   1440
            Width           =   240
         End
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Resumen F.P."
         Height          =   255
         Index           =   8
         Left            =   5280
         TabIndex        =   454
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   128
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   591
         Text            =   "Text5"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   128
         Left            =   1620
         MaxLength       =   16
         TabIndex        =   451
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   127
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   588
         Text            =   "Text5"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   127
         Left            =   1620
         MaxLength       =   16
         TabIndex        =   450
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Agrupa proveedor"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   579
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Frame Frame10 
         Caption         =   "Clasificado por "
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   120
         TabIndex        =   546
         Top             =   6600
         Width           =   2175
         Begin VB.OptionButton OptPorCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   1080
            TabIndex        =   548
            Top             =   280
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptPorFamilia 
            Caption         =   "Familia"
            Height          =   195
            Left            =   120
            TabIndex        =   547
            Top             =   280
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   99
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   453
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   98
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   452
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   96
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   475
         Text            =   "Text5"
         Top             =   900
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   448
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   97
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   474
         Text            =   "Text5"
         Top             =   1245
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   449
         Top             =   1245
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarEstVentas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   468
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   5400
         TabIndex        =   469
         Top             =   7080
         Width           =   975
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   240
         TabIndex        =   446
         Top             =   3000
         Width           =   6495
         Begin VB.CheckBox chkDatosAlbaranes 
            Caption         =   "Agrupado"
            Height          =   255
            Index           =   9
            Left            =   5520
            TabIndex        =   756
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CheckBox chkDatosAlbaranes 
            Caption         =   "Detalla proveedor"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   582
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   125
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   577
            Text            =   "Text5"
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   125
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   463
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   124
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   574
            Text            =   "Text5"
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   124
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   462
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chkDatosAlbaranes 
            Caption         =   "Datos albaranes"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   465
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkDetallaArticulo 
            Caption         =   "Detalla articulo"
            Height          =   195
            Left            =   120
            TabIndex        =   464
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Frame FrameDetallaArticulo 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   975
            Left            =   240
            TabIndex        =   529
            Top             =   2400
            Visible         =   0   'False
            Width           =   6135
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   113
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   533
               Text            =   "Text5"
               Top             =   600
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Index           =   113
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   467
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   112
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   530
               Text            =   "Text5"
               Top             =   240
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Index           =   112
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   466
               Top             =   240
               Width           =   1095
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   59
               Left            =   840
               Picture         =   "frmListadoOfer.frx":04A4
               Top             =   600
               Width           =   240
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   37
               Left            =   360
               TabIndex        =   534
               Top             =   600
               Width           =   420
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Art�culo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   36
               Left            =   0
               TabIndex        =   532
               Top             =   0
               Width           =   660
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   58
               Left            =   840
               Picture         =   "frmListadoOfer.frx":05A6
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   35
               Left            =   360
               TabIndex        =   531
               Top             =   240
               Width           =   450
            End
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   101
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   461
            Top             =   705
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   101
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   470
            Text            =   "Text5"
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   100
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   460
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   100
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   447
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   66
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":06A8
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   41
            Left            =   600
            TabIndex        =   578
            Top             =   1680
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   40
            Left            =   240
            TabIndex        =   576
            Top             =   1080
            Width           =   885
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   65
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":07AA
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   600
            TabIndex        =   575
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   600
            TabIndex        =   473
            Top             =   705
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   55
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":08AC
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   600
            TabIndex        =   472
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   54
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":09AE
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   471
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   68
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0AB0
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   840
         TabIndex        =   592
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   44
         Left            =   480
         TabIndex        =   590
         Top             =   1560
         Width           =   795
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   67
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0BB2
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   840
         TabIndex        =   589
         Top             =   1800
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   3600
         Picture         =   "frmListadoOfer.frx":0CB4
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   91
         Left            =   840
         TabIndex        =   482
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   90
         Left            =   480
         TabIndex        =   481
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0D3F
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   89
         Left            =   3120
         TabIndex        =   480
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Ventas por Familia / Art�culo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   31
         Left            =   1200
         TabIndex        =   479
         Top             =   240
         Width           =   4455
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   52
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0DCA
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   30
         Left            =   480
         TabIndex        =   478
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   840
         TabIndex        =   477
         Top             =   900
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   53
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":0ECC
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   840
         TabIndex        =   476
         Top             =   1245
         Width           =   420
      End
   End
   Begin VB.Frame FrameOfertas 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10155
      Begin VB.CheckBox chkVarios 
         Caption         =   "I.V.A. incluido"
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
         Index           =   12
         Left            =   3600
         TabIndex        =   765
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1660
         MaxLength       =   10
         TabIndex        =   6
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdAceptarOfer 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   9
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   4160
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Frame FrameTipoPapel 
         Caption         =   "Tipo de informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   480
         TabIndex        =   1
         Top             =   1720
         Width           =   3015
         Begin VB.OptionButton OptPapelBlanco 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptPapelMembrete 
            Caption         =   "Interna"
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
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3615
         Left            =   5640
         TabIndex        =   694
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.Label Label3 
         Caption         =   "Impresi�n documentos asociados"
         Height          =   195
         Index           =   8
         Left            =   5640
         TabIndex        =   693
         Top             =   840
         Width           =   2625
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Oferta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   19
         Left            =   600
         TabIndex        =   17
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   3360
         TabIndex        =   16
         Top             =   4320
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1340
         Picture         =   "frmListadoOfer.frx":0FCE
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListadoOfer.frx":1059
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   600
         TabIndex        =   15
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   600
         TabIndex        =   13
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   840
         TabIndex        =   12
         Top             =   4320
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":115B
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Imprimir otras Ofertas del Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Ofertas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame FrameComprobarCtaBancoSecciones 
      Height          =   3135
      Left            =   1560
      TabIndex        =   678
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkVarios 
         Caption         =   "Comprobar contabilidades Ariagro"
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
         Left            =   480
         TabIndex        =   690
         Top             =   2160
         Width           =   3720
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   148
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   688
         Text            =   "Text5"
         Top             =   1680
         Width           =   4020
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
         Index           =   148
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   680
         Top             =   1680
         Width           =   900
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
         Index           =   147
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   679
         Top             =   1305
         Width           =   900
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   147
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   684
         Text            =   "Text5"
         Top             =   1305
         Width           =   4020
      End
      Begin VB.CommandButton cmdComprobarCCC_NIF_Secciones 
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
         Left            =   3960
         TabIndex        =   681
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   408
         Left            =   5160
         TabIndex        =   682
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   120
         TabIndex        =   689
         Top             =   2640
         Width           =   3570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   51
         Left            =   300
         TabIndex        =   687
         Top             =   1680
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   86
         Left            =   960
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   50
         Left            =   300
         TabIndex        =   686
         Top             =   1305
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   49
         Left            =   210
         TabIndex        =   685
         Top             =   945
         Width           =   765
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   85
         Left            =   960
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Comprobar cuenta bancaria aplicaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   48
         Left            =   240
         TabIndex        =   683
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   7575
      Left            =   120
      TabIndex        =   496
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CheckBox chkMail 
         Caption         =   "Solo clientes envio email"
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
         Left            =   240
         TabIndex        =   700
         Top             =   5925
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Solo facturas marca enviar por email (Telefon�a)"
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
         TabIndex        =   691
         Top             =   6240
         Visible         =   0   'False
         Width           =   5280
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Incluir ya traspasadas"
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
         Left            =   240
         TabIndex        =   504
         Top             =   6240
         Width           =   2625
      End
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "Enviar"
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
         Left            =   7830
         TabIndex        =   510
         Top             =   7080
         Width           =   1065
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
         Height          =   2355
         Index           =   1
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   509
         Text            =   "frmListadoOfer.frx":11E6
         Top             =   3720
         Width           =   4335
      End
      Begin VB.ListBox ListTipoMov 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   1000
         Left            =   1230
         Style           =   1  'Checkbox
         TabIndex        =   503
         Top             =   4155
         Width           =   4230
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
         Index           =   111
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   498
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   111
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   524
         Text            =   "Text5"
         Top             =   1560
         Width           =   3420
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
         Index           =   110
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   497
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   110
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   521
         Text            =   "Text5"
         Top             =   1185
         Width           =   3420
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
         Height          =   320
         Index           =   0
         Left            =   5760
         TabIndex        =   508
         Text            =   "Text1"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Copia remitente"
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
         Left            =   5760
         TabIndex        =   507
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "administraci�n"
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
         Left            =   7740
         TabIndex        =   506
         Top             =   1320
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "comercial"
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
         Left            =   5760
         TabIndex        =   505
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   18
         Left            =   9000
         TabIndex        =   511
         Top             =   7080
         Width           =   1065
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
         Index           =   109
         Left            =   4155
         MaxLength       =   10
         TabIndex        =   500
         Top             =   2535
         Width           =   1350
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
         Index           =   108
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   499
         Top             =   2520
         Width           =   1350
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
         Index           =   106
         Left            =   1230
         MaxLength       =   7
         TabIndex        =   501
         Text            =   "wwwwwww"
         Top             =   3360
         Width           =   1410
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
         Index           =   107
         Left            =   4110
         MaxLength       =   7
         TabIndex        =   502
         Top             =   3360
         Width           =   1365
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   761
         Top             =   6720
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   2400
         TabIndex        =   762
         Top             =   6480
         Width           =   2925
      End
      Begin VB.Label lblInd 
         Caption         =   "Label11"
         Height          =   195
         Left            =   240
         TabIndex        =   674
         Top             =   6480
         Width           =   2370
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   5205
         Picture         =   "frmListadoOfer.frx":11EC
         ToolTipText     =   "Puntear al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   4845
         Picture         =   "frmListadoOfer.frx":1336
         ToolTipText     =   "Quitar al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
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
         Left            =   5760
         TabIndex        =   527
         Top             =   3435
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   20
         Left            =   120
         TabIndex        =   526
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   34
         Left            =   315
         TabIndex        =   525
         Top             =   1560
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   990
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   33
         Left            =   315
         TabIndex        =   523
         Top             =   1185
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   32
         Left            =   120
         TabIndex        =   522
         Top             =   840
         Width           =   675
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   990
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Index           =   19
         Left            =   5760
         TabIndex        =   520
         Top             =   2475
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   5760
         TabIndex        =   519
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   18
         Left            =   3210
         TabIndex        =   518
         Top             =   2520
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3870
         Picture         =   "frmListadoOfer.frx":1480
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   945
         Picture         =   "frmListadoOfer.frx":150B
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   517
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   17
         Left            =   315
         TabIndex        =   516
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label14 
         Caption         =   "ss"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   16
         Left            =   150
         TabIndex        =   515
         Top             =   360
         Width           =   5400
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N� Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   15
         Left            =   120
         TabIndex        =   514
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   14
         Left            =   315
         TabIndex        =   513
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   13
         Left            =   3315
         TabIndex        =   512
         Top             =   3360
         Width           =   570
      End
   End
   Begin VB.Frame FrameCompras 
      Height          =   5565
      Left            =   2520
      TabIndex        =   411
      Top             =   0
      Width           =   7035
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Solo con rappel"
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   764
         Top             =   4890
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Orden nombre proveedor"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   424
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Frame FrameMinImporte 
         Height          =   735
         Left            =   240
         TabIndex        =   586
         Top             =   3600
         Width           =   2415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   126
            Left            =   480
            MaxLength       =   16
            TabIndex        =   418
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Importe min. familia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   120
            TabIndex        =   587
            Top             =   120
            Width           =   1410
         End
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Resumen proveedor"
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   420
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Comparativo"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   423
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Salto pagina x prov."
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   422
         Top             =   4260
         Width           =   1815
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Datos rappel"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   421
         Top             =   4260
         Width           =   1575
      End
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Datos albaranes"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   419
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "Agrupar por"
         ForeColor       =   &H00000080&
         Height          =   945
         Left            =   360
         TabIndex        =   444
         Top             =   4440
         Width           =   2175
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   427
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia, Art�culo"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   428
            Top             =   550
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   438
         Top             =   2640
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   94
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   440
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   94
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   416
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   95
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   439
            Text            =   "Text5"
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   95
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   417
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   443
            Top             =   120
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   50
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":1596
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   442
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   51
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":1698
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   600
            TabIndex        =   441
            Top             =   705
            Width           =   420
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5880
         TabIndex        =   426
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarCompras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   425
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   91
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   413
         Top             =   1605
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   430
         Text            =   "Text5"
         Top             =   1605
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   90
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   412
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   429
         Text            =   "Text5"
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   92
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   414
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   93
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   415
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   2640
         TabIndex        =   562
         Top             =   5280
         Width           =   2010
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   960
         TabIndex        =   437
         Top             =   1605
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   49
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":179A
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   436
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   22
         Left            =   600
         TabIndex        =   435
         Top             =   1035
         Width           =   885
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   48
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":189C
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Compras por Familia/Art�culo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   21
         Left            =   600
         TabIndex        =   434
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   88
         Left            =   3360
         TabIndex        =   433
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":199E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   87
         Left            =   600
         TabIndex        =   432
         Top             =   2010
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   83
         Left            =   960
         TabIndex        =   431
         Top             =   2280
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":1A29
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.Frame FrameGenAlbCom 
      Height          =   4455
      Left            =   240
      TabIndex        =   209
      Top             =   240
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   168
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   214
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkImprAlbProv 
         Caption         =   "Imprime albaran"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   215
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   48
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   212
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   49
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   213
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAlbCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   216
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   4800
         TabIndex        =   217
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   211
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   1820
         Locked          =   -1  'True
         TabIndex        =   210
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   42
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":1AB4
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F. entrada "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   147
         Left            =   840
         TabIndex        =   763
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   840
         TabIndex        =   232
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   3
         Left            =   600
         TabIndex        =   231
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alb."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   60
         Left            =   840
         TabIndex        =   230
         Top             =   3000
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":1B3F
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Albaran de compra: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   59
         Left            =   600
         TabIndex        =   219
         Top             =   1200
         Width           =   5115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador del Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   58
         Left            =   840
         TabIndex        =   218
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   27
         Left            =   840
         Picture         =   "frmListadoOfer.frx":1BCA
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FrameEnvioMail2 
      Height          =   1215
      Left            =   6840
      TabIndex        =   528
      Top             =   1920
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdCancelaEnvio 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6240
         TabIndex        =   760
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame FrameFacRectif 
      Height          =   6855
      Left            =   3000
      TabIndex        =   314
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   161
         Left            =   600
         MaxLength       =   6
         TabIndex        =   326
         Top             =   5640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   161
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   730
         Text            =   "Text5"
         Top             =   5640
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   160
         Left            =   600
         MaxLength       =   10
         TabIndex        =   324
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   159
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   726
         Text            =   "Text5"
         Top             =   5040
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   159
         Left            =   600
         MaxLength       =   6
         TabIndex        =   325
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   645
         Index           =   87
         Left            =   600
         MaxLength       =   72
         MultiLine       =   -1  'True
         TabIndex        =   322
         Top             =   2760
         Width           =   4575
      End
      Begin VB.ComboBox cboTipomov 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListadoOfer.frx":1CCC
         Left            =   1865
         List            =   "frmListadoOfer.frx":1CCE
         Style           =   2  'Dropdown List
         TabIndex        =   319
         Top             =   1185
         Width           =   3675
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   4200
         TabIndex        =   328
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarFacRect 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   327
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1865
         MaxLength       =   10
         TabIndex        =   321
         Top             =   2115
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   71
         Left            =   1865
         MaxLength       =   10
         TabIndex        =   320
         Top             =   1725
         Width           =   1215
      End
      Begin VB.CheckBox chkFraRectCompleta 
         Caption         =   "Genera factura negativa"
         Height          =   255
         Left            =   600
         TabIndex        =   323
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   98
         Left            =   1920
         Picture         =   "frmListadoOfer.frx":1CD0
         Top             =   5400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   134
         Left            =   600
         TabIndex        =   731
         Top             =   5400
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N� Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   133
         Left            =   2880
         TabIndex        =   729
         Top             =   4440
         Width           =   2220
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   600
         X2              =   5400
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   132
         Left            =   600
         TabIndex        =   728
         Top             =   4200
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   37
         Left            =   1560
         Picture         =   "frmListadoOfer.frx":1DD2
         Top             =   4200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   1
         Left            =   240
         ToolTipText     =   "Factura rectificativa"
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   131
         Left            =   600
         TabIndex        =   727
         Top             =   4800
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   97
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":1E5D
         Top             =   4800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   82
         Left            =   600
         TabIndex        =   397
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   79
         Left            =   600
         TabIndex        =   318
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1500
         Picture         =   "frmListadoOfer.frx":1F5F
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   77
         Left            =   600
         TabIndex        =   317
         Top             =   2115
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Factura a Rectificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   480
         TabIndex        =   316
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   76
         Left            =   600
         TabIndex        =   315
         Top             =   1725
         Width           =   780
      End
   End
   Begin VB.Frame FrameEfectuadas 
      Height          =   7095
      Left            =   4080
      TabIndex        =   92
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   167
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   166
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   46
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkBusHcoOfer 
         Caption         =   "Incluir hist�rico de ofertas en busqueda "
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   6360
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   156
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   715
         Text            =   "Text5"
         Top             =   4680
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   156
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   53
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   155
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   712
         Top             =   4305
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   155
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   52
         Top             =   4305
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   154
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   710
         Text            =   "Text5"
         Top             =   5640
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   154
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   55
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   153
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   707
         Text            =   "Text5"
         Top             =   5280
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   153
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   54
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   118
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   550
         Text            =   "Text5"
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   51
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   117
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "Text5"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   50
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Solo Ofertas Pendientes"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   6720
         Width           =   2535
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   49
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   94
         Text            =   "Text5"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   48
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "Text5"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   45
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   59
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEfect 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   58
         Top             =   6600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   16
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   44
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   146
         Left            =   3480
         TabIndex        =   759
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   41
         Left            =   3960
         Picture         =   "frmListadoOfer.frx":1FEA
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   40
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2075
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   145
         Left            =   240
         TabIndex        =   758
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   144
         Left            =   960
         TabIndex        =   757
         Top             =   1560
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   94
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2100
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   716
         Top             =   4680
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   93
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2202
         Top             =   4305
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   54
         Left            =   240
         TabIndex        =   714
         Top             =   4080
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   53
         Left            =   960
         TabIndex        =   713
         Top             =   4305
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   92
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2304
         Top             =   5670
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   711
         Top             =   5640
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   91
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2406
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   128
         Left            =   240
         TabIndex        =   709
         Top             =   5040
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   708
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   96
         Left            =   120
         TabIndex        =   554
         Top             =   6000
         Width           =   2925
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   93
         Left            =   240
         TabIndex        =   551
         Top             =   3000
         Width           =   945
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   62
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2508
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   61
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":260A
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   101
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   7
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":270C
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   100
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   99
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   6
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":280E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   3960
         Picture         =   "frmListadoOfer.frx":2910
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   98
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   97
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ofertas Efectuadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1200
         TabIndex        =   96
         Top             =   240
         Width           =   3855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":299B
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3480
         TabIndex        =   95
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   95
         Left            =   960
         TabIndex        =   553
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   94
         Left            =   960
         TabIndex        =   552
         Top             =   3720
         Width           =   420
      End
   End
   Begin VB.Frame FramePteRecibir 
      Height          =   5205
      Left            =   480
      TabIndex        =   295
      Top             =   240
      Width           =   7035
      Begin VB.CheckBox chkVarios 
         Caption         =   "Listado resumen por articulo"
         CausesValidation=   0   'False
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   755
         Top             =   3480
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cboPdteFacturarProv 
         Height          =   315
         ItemData        =   "frmListadoOfer.frx":2A26
         Left            =   1680
         List            =   "frmListadoOfer.frx":2A33
         Style           =   2  'Dropdown List
         TabIndex        =   583
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Detalla albar�n"
         CausesValidation=   0   'False
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   584
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ordenar por"
         ForeColor       =   &H00000080&
         Height          =   825
         Left            =   600
         TabIndex        =   311
         Top             =   3840
         Width           =   2055
         Begin VB.OptionButton OptOrdenPed 
            Caption         =   "N� Pedido"
            Height          =   255
            Left            =   240
            TabIndex        =   313
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton OptOrdenArt 
            Caption         =   "Art�culo"
            Height          =   255
            Left            =   240
            TabIndex        =   312
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   305
         Top             =   2760
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   68
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   292
            Top             =   705
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   68
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   307
            Text            =   "Text5"
            Top             =   705
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   67
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   291
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   67
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   306
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   600
            TabIndex        =   310
            Top             =   705
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   44
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":2A4C
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   309
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   43
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":2B4E
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Art�culo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   308
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   290
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   289
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   297
         Text            =   "Text5"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   287
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   296
         Text            =   "Text5"
         Top             =   1725
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   288
         Top             =   1725
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarPte 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   293
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5280
         TabIndex        =   294
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   129
         Left            =   600
         TabIndex        =   722
         Top             =   3000
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":2C50
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   304
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   74
         Left            =   600
         TabIndex        =   303
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":2CDB
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   72
         Left            =   3360
         TabIndex        =   302
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Material pendiente de recibir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   19
         Left            =   600
         TabIndex        =   301
         Top             =   360
         Width           =   4455
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   41
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2D66
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   600
         TabIndex        =   300
         Top             =   1035
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   960
         TabIndex        =   299
         Top             =   1380
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   42
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":2E68
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   960
         TabIndex        =   298
         Top             =   1725
         Width           =   420
      End
   End
   Begin VB.Frame FrameOfertasHcoInforme 
      Height          =   7575
      Left            =   840
      TabIndex        =   735
      Top             =   -120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   162
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   745
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   162
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   740
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   163
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   744
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   163
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   741
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   164
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   738
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   165
         Left            =   3660
         MaxLength       =   10
         TabIndex        =   739
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmListadoOfer.frx":2F6A
         Left            =   1320
         List            =   "frmListadoOfer.frx":2F77
         Style           =   2  'Dropdown List
         TabIndex        =   743
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdListadoOfertasHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   737
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   4920
         TabIndex        =   736
         Top             =   6960
         Width           =   975
      End
      Begin MSComctlLib.ListView lwMotBaja 
         Height          =   3735
         Left            =   1320
         TabIndex        =   742
         Top             =   2880
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MOtivo"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Informe ofertas traspasdas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   1080
         TabIndex        =   754
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   136
         Left            =   240
         TabIndex        =   753
         Top             =   1680
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   99
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":2F85
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   137
         Left            =   480
         TabIndex        =   752
         Top             =   1920
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   100
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":3087
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   138
         Left            =   480
         TabIndex        =   751
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   38
         Left            =   1095
         Picture         =   "frmListadoOfer.frx":3189
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   139
         Left            =   240
         TabIndex        =   750
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   140
         Left            =   480
         TabIndex        =   749
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   141
         Left            =   2880
         TabIndex        =   748
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   39
         Left            =   3360
         Picture         =   "frmListadoOfer.frx":3214
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Aceptada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   142
         Left            =   240
         TabIndex        =   747
         Top             =   6840
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   143
         Left            =   240
         TabIndex        =   746
         Top             =   2880
         Width           =   585
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   8
         Left            =   960
         Picture         =   "frmListadoOfer.frx":329F
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   9
         Left            =   960
         Picture         =   "frmListadoOfer.frx":33E9
         Top             =   2880
         Width           =   240
      End
   End
   Begin VB.Frame FrameClienInactivos 
      Height          =   7005
      Left            =   0
      TabIndex        =   119
      Top             =   -120
      Width           =   10995
      Begin VB.Frame FrameTipoPago 
         Caption         =   "Frame11"
         Height          =   3375
         Left            =   6480
         TabIndex        =   723
         Top             =   600
         Width           =   4215
         Begin MSComctlLib.ListView lwTippa 
            Height          =   2895
            Left            =   120
            TabIndex        =   724
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   5106
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cargo"
               Object.Width           =   6174
            EndProperty
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   7
            Left            =   3000
            Picture         =   "frmListadoOfer.frx":3533
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   6
            Left            =   3360
            Picture         =   "frmListadoOfer.frx":367D
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   130
            Left            =   0
            TabIndex        =   725
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.CheckBox chkEtiqDpto 
         Caption         =   "Contacto/Cargos"
         Height          =   255
         Left            =   5160
         TabIndex        =   593
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CheckBox chkEnviaCorreo 
         Caption         =   "Marca envia correo"
         Height          =   255
         Left            =   3000
         TabIndex        =   585
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   6360
         TabIndex        =   282
         Top             =   1680
         Width           =   4455
         Begin VB.Frame Frame5 
            Caption         =   "e-Mail"
            Height          =   780
            Left            =   600
            TabIndex        =   134
            Top             =   1680
            Width           =   2000
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Comercial"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   286
               Top             =   460
               Width           =   1335
            End
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administraci�n"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   285
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   64
            Left            =   180
            MaxLength       =   6
            TabIndex        =   132
            Top             =   860
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   64
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   283
            Text            =   "Text5"
            Top             =   860
            Width           =   3375
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   133
            Top             =   1395
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   284
            Top             =   650
            Width           =   465
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   40
            Left            =   840
            Picture         =   "frmListadoOfer.frx":37C7
            Top             =   580
            Width           =   240
         End
      End
      Begin VB.Frame FrameImpClien 
         Caption         =   "Imprimir clientes"
         ForeColor       =   &H00000080&
         Height          =   1050
         Left            =   600
         TabIndex        =   131
         Top             =   5760
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton OptCliTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   240
            TabIndex        =   260
            Top             =   735
            Width           =   1215
         End
         Begin VB.OptionButton OptCliSinMante 
            Caption         =   "Sin mantenimiento"
            Height          =   255
            Left            =   240
            TabIndex        =   259
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton OptCliConMante 
            Caption         =   "Con mantenimiento"
            Height          =   255
            Left            =   240
            TabIndex        =   258
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   480
         TabIndex        =   246
         Top             =   2900
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   57
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   261
            Text            =   "Text5"
            Top             =   2025
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   57
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   129
            Top             =   2025
            Width           =   855
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   130
            Top             =   2385
            Width           =   4095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   56
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   128
            Top             =   1470
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   56
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   254
            Text            =   "Text5"
            Top             =   1470
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   127
            Top             =   1130
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   55
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   253
            Text            =   "Text5"
            Top             =   1130
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   126
            Top             =   580
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   54
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   248
            Text            =   "Text5"
            Top             =   580
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   125
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   53
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   247
            Text            =   "Text5"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   34
            Left            =   960
            Picture         =   "frmListadoOfer.frx":38C9
            Top             =   2025
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Situaci�n"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   73
            Left            =   120
            TabIndex        =   262
            Top             =   2025
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "A la atenci�n de:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   71
            Left            =   120
            TabIndex        =   257
            Top             =   2385
            Width           =   1395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   70
            Left            =   480
            TabIndex        =   256
            Top             =   1470
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   33
            Left            =   960
            Picture         =   "frmListadoOfer.frx":39CB
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   69
            Left            =   480
            TabIndex        =   255
            Top             =   1130
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3ACD
            Top             =   1130
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CPostal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   68
            Left            =   120
            TabIndex        =   252
            Top             =   890
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   67
            Left            =   480
            TabIndex        =   251
            Top             =   580
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   31
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3BCF
            Top             =   580
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   66
            Left            =   480
            TabIndex        =   250
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Actividad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   57
            Left            =   120
            TabIndex        =   249
            Top             =   0
            Width           =   795
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   30
            Left            =   960
            Picture         =   "frmListadoOfer.frx":3CD1
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   32
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   137
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   31
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   124
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarClienInac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   135
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5280
         TabIndex        =   136
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   141
         Text            =   "Text5"
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   120
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "Text5"
         Top             =   1600
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   121
         Top             =   1600
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Text5"
         Top             =   2200
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   122
         Top             =   2200
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   138
         Text            =   "Text5"
         Top             =   2550
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   123
         Top             =   2550
         Width           =   855
      End
      Begin VB.Frame frameCliexFacturas 
         Caption         =   "Desde / hasta facturas"
         Height          =   4575
         Left            =   6360
         TabIndex        =   483
         Top             =   1080
         Width           =   4575
         Begin MSComctlLib.ListView lwCargos 
            Height          =   1815
            Left            =   240
            TabIndex        =   595
            Top             =   2520
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cargo"
               Object.Width           =   6174
            EndProperty
         End
         Begin VB.ComboBox cboTipomov 
            Height          =   315
            Index           =   2
            ItemData        =   "frmListadoOfer.frx":3DD3
            Left            =   1680
            List            =   "frmListadoOfer.frx":3DD5
            Style           =   2  'Dropdown List
            TabIndex        =   484
            Top             =   360
            Width           =   1875
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   104
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   487
            Top             =   1755
            Width           =   1080
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   102
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   485
            Text            =   "wwwwwww"
            Top             =   1035
            Width           =   1125
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   103
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   486
            Top             =   1035
            Width           =   1125
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   105
            Left            =   3240
            MaxLength       =   10
            TabIndex        =   488
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   4080
            Picture         =   "frmListadoOfer.frx":3DD7
            Top             =   2160
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   3720
            Picture         =   "frmListadoOfer.frx":3F21
            Top             =   2160
            Width           =   240
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00008000&
            X1              =   3600
            X2              =   840
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cargos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   594
            Top             =   2160
            Width           =   825
         End
         Begin VB.Image imgClearCmbTipomov 
            Height          =   240
            Left            =   3720
            Picture         =   "frmListadoOfer.frx":406B
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   495
            Top             =   360
            Width           =   1410
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   31
            Left            =   720
            Picture         =   "frmListadoOfer.frx":45F5
            Top             =   1770
            Width           =   240
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fact."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   494
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "N� Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   493
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   492
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   2400
            TabIndex        =   491
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   490
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   489
            Top             =   1800
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   32
            Left            =   2880
            Picture         =   "frmListadoOfer.frx":4680
            Top             =   1770
            Width           =   240
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   3250
         TabIndex        =   151
         Top             =   3360
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   3720
         Picture         =   "frmListadoOfer.frx":470B
         Top             =   3375
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   960
         TabIndex        =   150
         Top             =   3360
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4796
         Top             =   3380
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inactividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   36
         Left            =   600
         TabIndex        =   149
         Top             =   3120
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Clientes Inactivos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   148
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4821
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   42
         Left            =   600
         TabIndex        =   147
         Top             =   1040
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   960
         TabIndex        =   146
         Top             =   1260
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4923
         Top             =   1600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   960
         TabIndex        =   145
         Top             =   1600
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4A25
         Top             =   2200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   39
         Left            =   600
         TabIndex        =   144
         Top             =   1940
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   960
         TabIndex        =   143
         Top             =   2200
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4B27
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   960
         TabIndex        =   142
         Top             =   2550
         Width           =   420
      End
   End
   Begin VB.Frame FrameTraspasoHco 
      Height          =   5895
      Left            =   600
      TabIndex        =   102
      Top             =   360
      Width           =   6915
      Begin VB.ComboBox cboMotTras 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4560
         Width           =   4815
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   43
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   205
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   44
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   204
         Text            =   "Text5"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   200
         Text            =   "Text5"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   4140
         MaxLength       =   7
         TabIndex        =   28
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   22
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarTrasHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   69
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   70
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   23
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   25
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   135
         Left            =   600
         TabIndex        =   734
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   56
         Left            =   600
         TabIndex        =   208
         Top             =   1200
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   23
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4C29
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   207
         Top             =   1440
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   24
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4D2B
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   960
         TabIndex        =   206
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   53
         Left            =   600
         TabIndex        =   203
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   25
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4E2D
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   960
         TabIndex        =   202
         Top             =   2400
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   26
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":4F2F
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   960
         TabIndex        =   201
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Oferta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   109
         Top             =   3720
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   108
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   107
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3360
         TabIndex        =   106
         Top             =   3360
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":5031
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Traspaso de Ofertas a Hist�rico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   105
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   104
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   103
         Top             =   3360
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":50BC
         Top             =   3360
         Width           =   240
      End
   End
   Begin VB.Frame FramePasarHco 
      Height          =   4575
      Left            =   120
      TabIndex        =   233
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   52
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   244
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   236
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   51
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   239
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   235
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5400
         TabIndex        =   238
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   237
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   50
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   234
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   29
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":5147
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   720
         TabIndex        =   245
         Top             =   2760
         Width           =   720
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   28
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":5249
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   64
         Left            =   720
         TabIndex        =   243
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el hist�rico: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   63
         Left            =   600
         TabIndex        =   242
         Top             =   1200
         Width           =   4245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   2040
         Picture         =   "frmListadoOfer.frx":534B
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Eliminaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   62
         Left            =   720
         TabIndex        =   241
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Albaran al hist�rico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   4
         Left            =   600
         TabIndex        =   240
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame FramePedidos 
      Height          =   4455
      Left            =   600
      TabIndex        =   329
      Top             =   240
      Width           =   6075
      Begin VB.CheckBox chkVarios 
         Caption         =   "Pedido interno"
         Height          =   195
         Index           =   10
         Left            =   3000
         TabIndex        =   701
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Valorado"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   335
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   331
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   75
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   333
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4080
         TabIndex        =   339
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   337
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   74
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   332
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   330
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Peso carga fenollar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   733
         Top             =   3480
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   732
         Top             =   3480
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   343
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label12 
         Caption         =   "Imprimir otros Pedidos del Proveedor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   342
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   3480
         Picture         =   "frmListadoOfer.frx":53D6
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   341
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   340
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Informe de Pedido Compras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   338
         Top             =   360
         Width           =   4335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":5461
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   336
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "N� Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   334
         Top             =   1320
         Width           =   810
      End
   End
   Begin VB.Frame FrameClientes2 
      Height          =   7695
      Left            =   600
      TabIndex        =   152
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   152
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   705
         Text            =   "Text5"
         Top             =   5040
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   152
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   166
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   151
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   702
         Text            =   "Text5"
         Top             =   4695
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   151
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   165
         Top             =   4695
         Width           =   615
      End
      Begin VB.OptionButton optClienteLis 
         Caption         =   "F.Pago"
         Height          =   195
         Index           =   2
         Left            =   8040
         TabIndex        =   677
         Top             =   5160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optClienteLis 
         Caption         =   "Email"
         Height          =   195
         Index           =   1
         Left            =   7140
         TabIndex        =   676
         Top             =   5160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optClienteLis 
         Caption         =   "Telefonos"
         Height          =   195
         Index           =   0
         Left            =   6000
         TabIndex        =   675
         Top             =   5160
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Poblacion / actividad"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   565
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   130
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   601
         Text            =   "Text5"
         Top             =   6960
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   130
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   170
         Top             =   6960
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   129
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   599
         Text            =   "Text5"
         Top             =   6600
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   129
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   169
         Top             =   6600
         Width           =   855
      End
      Begin VB.Frame FrVolVetasCredito 
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   495
         Left            =   5880
         TabIndex        =   596
         Top             =   4080
         Visible         =   0   'False
         Width           =   2775
         Begin VB.ComboBox cboClienteCredito 
            Height          =   315
            ItemData        =   "frmListadoOfer.frx":54EC
            Left            =   960
            List            =   "frmListadoOfer.frx":54FF
            Style           =   2  'Dropdown List
            TabIndex        =   567
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cr�dito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   597
            Top             =   150
            Width           =   525
         End
      End
      Begin VB.CheckBox chkExportacion 
         Caption         =   "Formato exportaci�n"
         Height          =   255
         Left            =   6000
         TabIndex        =   568
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboOrdVolVta 
         Height          =   315
         ItemData        =   "frmListadoOfer.frx":5541
         Left            =   6000
         List            =   "frmListadoOfer.frx":554B
         Style           =   2  'Dropdown List
         TabIndex        =   566
         Top             =   3720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkVolumen 
         Caption         =   "Inf. con volumen ventas"
         Height          =   255
         Left            =   6000
         TabIndex        =   564
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Frame FrameVolumen 
         Caption         =   "Volumen de ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   6000
         TabIndex        =   563
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   123
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   571
            Top             =   1560
            Width           =   1080
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   122
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   569
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fechas c�lculo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   573
            Top             =   600
            Width           =   1035
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   35
            Left            =   840
            Picture         =   "frmListadoOfer.frx":556C
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   572
            Top             =   1590
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   36
            Left            =   840
            Picture         =   "frmListadoOfer.frx":55F7
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   570
            Top             =   1110
            Width           =   450
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   167
         Top             =   5655
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   168
         Top             =   5970
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   41
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   195
         Text            =   "Text5"
         Top             =   5655
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   42
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   194
         Text            =   "Text5"
         Top             =   5970
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   178
         Text            =   "Text5"
         Top             =   3270
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   177
         Text            =   "Text5"
         Top             =   2955
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   162
         Top             =   3270
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   161
         Top             =   2955
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   7800
         TabIndex        =   172
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6600
         TabIndex        =   171
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   157
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   158
         Top             =   1635
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   33
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   176
         Text            =   "Text5"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   34
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   175
         Text            =   "Text5"
         Top             =   1635
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   159
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   160
         Top             =   2475
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   35
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   174
         Text            =   "Text5"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   173
         Text            =   "Text5"
         Top             =   2475
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   163
         Top             =   3795
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   164
         Top             =   4110
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   3795
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   40
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "Text5"
         Top             =   4110
         Width           =   3135
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":5682
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   2505
         Width           =   435
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":598C
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   1720
         Width           =   435
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   6480
         TabIndex        =   179
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   90
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":5C96
         Top             =   5055
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   1080
         TabIndex        =   706
         Top             =   5040
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   89
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":5D98
         Top             =   4695
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Visitador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   127
         Left            =   600
         TabIndex        =   704
         Top             =   4440
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   1080
         TabIndex        =   703
         Top             =   4695
         Width           =   465
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   0
         Left            =   6120
         ToolTipText     =   "Listados de clientes"
         Top             =   6960
         Width           =   255
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   70
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":5E9A
         Top             =   6960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   1080
         TabIndex        =   602
         Top             =   6960
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   69
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":5F9C
         Top             =   6600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   600
         Top             =   6600
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "C.Postal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   101
         Left            =   600
         TabIndex        =   598
         Top             =   6360
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   198
         Top             =   5655
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   197
         Top             =   5970
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   45
         Left            =   600
         TabIndex        =   196
         Top             =   5400
         Width           =   780
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   21
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":609E
         Top             =   5655
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   22
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":61A0
         Top             =   5985
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   18
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":62A2
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":63A4
         Top             =   2955
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   51
         Left            =   600
         TabIndex        =   193
         Top             =   2715
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   1080
         TabIndex        =   192
         Top             =   3270
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   1080
         TabIndex        =   191
         Top             =   2955
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   1080
         TabIndex        =   190
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   1080
         TabIndex        =   189
         Top             =   1635
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   49
         Left            =   600
         TabIndex        =   188
         Top             =   1080
         Width           =   795
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   13
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":64A6
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":65A8
         Top             =   1650
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   187
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   1080
         TabIndex        =   186
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   1080
         TabIndex        =   185
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   48
         Left            =   600
         TabIndex        =   184
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":66AA
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":67AC
         Top             =   2505
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   1080
         TabIndex        =   183
         Top             =   3795
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   1080
         TabIndex        =   182
         Top             =   4110
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   47
         Left            =   600
         TabIndex        =   181
         Top             =   3540
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   19
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":68AE
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   20
         Left            =   1635
         Picture         =   "frmListadoOfer.frx":69B0
         Top             =   4125
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   46
         Left            =   6480
         TabIndex        =   180
         Top             =   1200
         Width           =   1545
      End
   End
   Begin VB.Frame FrameFacReimprimir 
      Height          =   6375
      Left            =   2520
      TabIndex        =   371
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chk_duplicado2 
         Caption         =   "Excluir clientes envio email"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   383
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   150
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   698
         Text            =   "Text5"
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   150
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   376
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   149
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   695
         Text            =   "Text5"
         Top             =   2655
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   149
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   375
         Top             =   2655
         Width           =   615
      End
      Begin VB.CheckBox chk_duplicado2 
         Caption         =   "Solo factura en papel"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   386
         Top             =   5880
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chk_duplicado2 
         Caption         =   "Ordenado x cliente"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   384
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   121
         Left            =   1200
         TabIndex        =   374
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   121
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   560
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   1200
         TabIndex        =   373
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   120
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   557
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   381
         Top             =   4920
         Width           =   885
      End
      Begin VB.CheckBox chkFormatoTPV 
         Caption         =   "Formato factura TPV"
         Height          =   255
         Left            =   4560
         TabIndex        =   385
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chk_duplicado2 
         Caption         =   "Duplicado"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   382
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   380
         Top             =   4395
         Width           =   1200
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   378
         Top             =   3780
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   377
         Text            =   "wwwwwww"
         Top             =   3780
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   379
         Top             =   4395
         Width           =   1080
      End
      Begin VB.CommandButton cmdAceptarReimpFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   387
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5400
         TabIndex        =   388
         Top             =   5880
         Width           =   975
      End
      Begin VB.ComboBox cboTipomov 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListadoOfer.frx":6AB2
         Left            =   2040
         List            =   "frmListadoOfer.frx":6AB4
         Style           =   2  'Dropdown List
         TabIndex        =   372
         Top             =   840
         Width           =   3555
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   88
         Left            =   915
         Picture         =   "frmListadoOfer.frx":6AB6
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   699
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   87
         Left            =   915
         Picture         =   "frmListadoOfer.frx":6BB8
         Top             =   2655
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   126
         Left            =   240
         TabIndex        =   697
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   696
         Top             =   2655
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   100
         Left            =   360
         TabIndex        =   561
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   64
         Left            =   840
         Picture         =   "frmListadoOfer.frx":6CBA
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   99
         Left            =   360
         TabIndex        =   559
         Top             =   1560
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   63
         Left            =   840
         Picture         =   "frmListadoOfer.frx":6DBC
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   98
         Left            =   240
         TabIndex        =   558
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N� copias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   556
         Top             =   4920
         Width           =   780
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":6EBE
         Top             =   4410
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3360
         TabIndex        =   396
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   395
         Top             =   4440
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3600
         TabIndex        =   394
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   393
         Top             =   3825
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N� Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   392
         Top             =   3540
         Width           =   885
      End
      Begin VB.Label Label14 
         Caption         =   "Reimprimir Facturas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   391
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   390
         Top             =   4170
         Width           =   945
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   960
         Picture         =   "frmListadoOfer.frx":6F49
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   389
         Top             =   840
         Width           =   1410
      End
   End
   Begin VB.Frame FrameEtiqProv 
      Height          =   5925
      Left            =   840
      TabIndex        =   263
      Top             =   360
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   62
         Left            =   1750
         MaxLength       =   50
         TabIndex        =   224
         Top             =   3360
         Width           =   4575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   229
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEtiqProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   228
         Top             =   5400
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   360
         TabIndex        =   276
         Top             =   3720
         Width           =   6255
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   146
            Left            =   1370
            MaxLength       =   50
            TabIndex        =   226
            Top             =   720
            Width           =   4575
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "e-Mail"
            Height          =   420
            Left            =   1800
            TabIndex        =   279
            Top             =   1080
            Width           =   4215
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administraci�n"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   281
               Top             =   120
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   280
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   227
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   63
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   277
            Text            =   "Text5"
            Top             =   105
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   225
            Top             =   105
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Firmado "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   47
            Left            =   240
            TabIndex        =   673
            Top             =   720
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":6FD4
            Top             =   105
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   278
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   60
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   272
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   222
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   61
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   271
         Text            =   "Text5"
         Top             =   2625
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   223
         Top             =   2625
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   221
         Top             =   1605
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   59
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   265
         Text            =   "Text5"
         Top             =   1605
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   220
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   264
         Text            =   "Text5"
         Top             =   1260
         Width           =   3735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A la atenci�n de:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   270
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPostal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   275
         Top             =   2040
         Width           =   630
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   37
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":70D6
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   274
         Top             =   2280
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":71D8
         Top             =   2625
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   273
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   269
         Top             =   1605
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   36
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":72DA
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   268
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   267
         Top             =   915
         Width           =   885
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   35
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":73DC
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Proveedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   266
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameClientesPotenciales 
      Height          =   5655
      Left            =   240
      TabIndex        =   603
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   135
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   612
         Top             =   3840
         Width           =   4095
      End
      Begin VB.CommandButton cmdCliPot 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   614
         Top             =   5160
         Width           =   975
      End
      Begin VB.Frame FrameCartaPot 
         Caption         =   "Frame11"
         Height          =   615
         Left            =   120
         TabIndex        =   630
         Top             =   4080
         Width           =   6255
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   138
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   613
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   138
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   631
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   46
            Left            =   0
            TabIndex        =   632
            Top             =   240
            Width           =   585
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   77
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":74DE
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   5280
         TabIndex        =   615
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   137
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   628
         Text            =   "Text5"
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   137
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   609
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   136
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   625
         Text            =   "Text5"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   136
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   608
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   134
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   623
         Text            =   "Text5"
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   134
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   611
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   133
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   620
         Text            =   "Text5"
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   133
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   610
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   132
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   607
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   132
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   618
         Text            =   "Text5"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   606
         Top             =   1065
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   131
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   605
         Text            =   "Text5"
         Top             =   1065
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A la atenci�n de:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   108
         Left            =   120
         TabIndex        =   672
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   76
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":75E0
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   111
         Left            =   720
         TabIndex        =   629
         Top             =   2400
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   75
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":76E2
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   110
         Left            =   120
         TabIndex        =   627
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   109
         Left            =   720
         TabIndex        =   626
         Top             =   2040
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   74
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":77E4
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   107
         Left            =   720
         TabIndex        =   624
         Top             =   3360
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CPostal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   106
         Left            =   120
         TabIndex        =   622
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   73
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":78E6
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   105
         Left            =   720
         TabIndex        =   621
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   104
         Left            =   720
         TabIndex        =   619
         Top             =   1440
         Width           =   405
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   72
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":79E8
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   103
         Left            =   720
         TabIndex        =   617
         Top             =   1065
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   102
         Left            =   120
         TabIndex        =   616
         Top             =   840
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   71
         Left            =   1200
         Picture         =   "frmListadoOfer.frx":7AEA
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   604
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame FrameCRMProgess 
      Height          =   5055
      Left            =   2640
      TabIndex        =   665
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdPararCRM 
         Caption         =   "Parar"
         Height          =   375
         Left            =   2160
         TabIndex        =   667
         Top             =   4560
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar pbCRM 
         Height          =   375
         Left            =   720
         TabIndex        =   666
         Top             =   4080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proceso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   122
         Left            =   360
         TabIndex        =   671
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   123
         Left            =   1200
         TabIndex        =   670
         Top             =   480
         Width           =   3330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   124
         Left            =   1200
         TabIndex        =   669
         Top             =   960
         Width           =   3330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   125
         Left            =   360
         TabIndex        =   668
         Top             =   960
         Width           =   585
      End
   End
   Begin VB.Frame FrameCRM 
      Height          =   6135
      Left            =   2400
      TabIndex        =   633
      Top             =   -120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkVarios 
         Caption         =   "Pendiente cobros"
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   644
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton cmdImpresionCRM 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   3720
         TabIndex        =   645
         Top             =   5520
         Width           =   975
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Asegurado"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   643
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Privado"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   642
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Credito"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   641
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   145
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   640
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   145
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   663
         Text            =   "Text5"
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   144
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   661
         Text            =   "Text5"
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   144
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   639
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   143
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   658
         Text            =   "Text5"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   143
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   638
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   142
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   637
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   142
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   656
         Text            =   "Text5"
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   141
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   636
         Top             =   2415
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   141
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   653
         Text            =   "Text5"
         Top             =   2415
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   140
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   635
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   140
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   651
         Text            =   "Text5"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   139
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   634
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   139
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   648
         Text            =   "Text5"
         Top             =   1185
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   406
         Left            =   4800
         TabIndex        =   646
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   121
         Left            =   120
         TabIndex        =   664
         Top             =   4320
         Width           =   780
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   84
         Left            =   960
         Picture         =   "frmListadoOfer.frx":7BEC
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   83
         Left            =   960
         Picture         =   "frmListadoOfer.frx":7CEE
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   120
         Left            =   480
         TabIndex        =   662
         Top             =   3840
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   82
         Left            =   960
         Picture         =   "frmListadoOfer.frx":7DF0
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   119
         Left            =   120
         TabIndex        =   660
         Top             =   3240
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   118
         Left            =   480
         TabIndex        =   659
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   117
         Left            =   480
         TabIndex        =   657
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   81
         Left            =   960
         Picture         =   "frmListadoOfer.frx":7EF2
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   116
         Left            =   480
         TabIndex        =   655
         Top             =   2415
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   115
         Left            =   120
         TabIndex        =   654
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   80
         Left            =   960
         Picture         =   "frmListadoOfer.frx":7FF4
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   114
         Left            =   480
         TabIndex        =   652
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   79
         Left            =   960
         Picture         =   "frmListadoOfer.frx":80F6
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   113
         Left            =   480
         TabIndex        =   650
         Top             =   1185
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   112
         Left            =   120
         TabIndex        =   649
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   78
         Left            =   960
         Picture         =   "frmListadoOfer.frx":81F8
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Impresion CRM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   2
         Left            =   1680
         TabIndex        =   647
         Top             =   360
         Width           =   2235
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   3735
      Left            =   0
      TabIndex        =   398
      Top             =   0
      Width           =   6315
      Begin VB.CheckBox chkVarios 
         Caption         =   "Incluir todo tipo de facturas"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   403
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Frame FrameAgrupar 
         Caption         =   "Agrupar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1000
         Left            =   600
         TabIndex        =   410
         Top             =   2160
         Width           =   2055
         Begin VB.OptionButton optForpago 
            Caption         =   "Tipo de pago"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   402
            Top             =   620
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optForpago 
            Caption         =   "Forma de pago"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   401
            Top             =   320
            Width           =   1695
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   88
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   399
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarCierre 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   404
         Top             =   2785
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4440
         TabIndex        =   405
         Top             =   2785
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   89
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   400
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   409
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1480
         Picture         =   "frmListadoOfer.frx":82FA
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Cierre de Caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   408
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   407
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   406
         Top             =   1560
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   3960
         Picture         =   "frmListadoOfer.frx":8385
         Top             =   1560
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameGenPedido 
      Height          =   4455
      Left            =   360
      TabIndex        =   110
      Top             =   120
      Width           =   6315
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   117
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   1820
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   75
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   79
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarGenPed 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   78
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   26
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   77
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   25
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   76
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Semana"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   3600
         TabIndex        =   118
         Top             =   3000
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListadoOfer.frx":8410
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador de Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   840
         TabIndex        =   116
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Pedido: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   114
         Top             =   1200
         Width           =   4080
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1920
         Picture         =   "frmListadoOfer.frx":8512
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrega"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   840
         TabIndex        =   113
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Oferta a Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   112
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   1920
         Picture         =   "frmListadoOfer.frx":859D
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   840
         TabIndex        =   111
         Top             =   2520
         Width           =   960
      End
   End
   Begin VB.Frame FrameRecordatorio 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   6915
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar  coste con:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1695
         Left            =   4080
         TabIndex        =   87
         Top             =   4605
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   680
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   1000
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   2055
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   15
         Left            =   720
         MaxLength       =   80
         TabIndex        =   38
         Top             =   5100
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   14
         Left            =   720
         MaxLength       =   80
         TabIndex        =   37
         Top             =   4800
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   35
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   34
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   3360
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   33
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   32
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3900
         MaxLength       =   7
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame FrameTipoPapel2 
         Caption         =   "Tipo de Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   600
         TabIndex        =   42
         Top             =   5565
         Width           =   2775
         Begin VB.OptionButton OptPapelMembreteR 
            Caption         =   "Con Membrete"
            Height          =   255
            Left            =   1320
            TabIndex        =   60
            Top             =   320
            Width           =   1335
         End
         Begin VB.OptionButton OptPapelBlancoR 
            Caption         =   "Blanco"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   320
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   8
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   40
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdAcetarRecorda 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   39
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   36
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   7
         Left            =   1720
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1720
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lineas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   86
         Top             =   4560
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   85
         Top             =   3720
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8628
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   960
         TabIndex        =   83
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   32
         Left            =   600
         TabIndex        =   82
         Top             =   3120
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":872A
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   80
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":882C
         Top             =   2770
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   960
         TabIndex        =   73
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   29
         Left            =   600
         TabIndex        =   72
         Top             =   2160
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":892E
         Top             =   2410
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   3130
         TabIndex        =   68
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   67
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   3610
         Picture         =   "frmListadoOfer.frx":8A30
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   66
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Oferta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   25
         Left            =   600
         TabIndex        =   65
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Recordatorio de Ofertas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   600
         TabIndex        =   64
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   63
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8ABB
         Top             =   4220
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8BBD
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   3130
         TabIndex        =   62
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Oferta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   600
         TabIndex        =   61
         Top             =   960
         Width           =   780
      End
   End
   Begin VB.Frame FrameConfirmPed 
      Height          =   6255
      Left            =   480
      TabIndex        =   344
      Top             =   120
      Width           =   7035
      Begin VB.CheckBox chkConfirmPed 
         Caption         =   "Adjuntar pedidos"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   355
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CheckBox chkConfirmPed 
         Caption         =   "Enviar por email"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   354
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   80
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   348
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   80
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   361
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   79
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   347
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   79
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   360
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Frame FrameTipoPapel3 
         Caption         =   "Tipo de Papel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   600
         TabIndex        =   359
         Top             =   4485
         Width           =   3375
         Begin VB.OptionButton OptPapelMembrete3 
            Caption         =   "Con Membrete"
            Height          =   255
            Left            =   1800
            TabIndex        =   353
            Top             =   320
            Width           =   1335
         End
         Begin VB.OptionButton OptPapelBlanco3 
            Caption         =   "Blanco"
            Height          =   195
            Left            =   240
            TabIndex        =   352
            Top             =   320
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   78
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   346
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5280
         TabIndex        =   357
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton cmdAcetarConfirm 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   356
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   81
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   349
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   81
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   358
         Text            =   "Text5"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   77
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   345
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   82
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   350
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkImpSaldo 
         Caption         =   "Imprimir saldo"
         Height          =   375
         Left            =   3840
         TabIndex        =   351
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   97
         Left            =   360
         TabIndex        =   555
         Top             =   5760
         Width           =   3525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   86
         Left            =   960
         TabIndex        =   370
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   47
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8C48
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   85
         Left            =   960
         TabIndex        =   369
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   84
         Left            =   600
         TabIndex        =   368
         Top             =   2040
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   46
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8D4A
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   3600
         Picture         =   "frmListadoOfer.frx":8E4C
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   367
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   366
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label13 
         Caption         =   "Cartas Confirmaci�n de Pedidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   365
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   81
         Left            =   600
         TabIndex        =   364
         Top             =   3360
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   45
         Left            =   1155
         Picture         =   "frmListadoOfer.frx":8ED7
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1440
         Picture         =   "frmListadoOfer.frx":8FD9
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   80
         Left            =   3120
         TabIndex        =   363
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   78
         Left            =   600
         TabIndex        =   362
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1680
         Picture         =   "frmListadoOfer.frx":9064
         Top             =   3840
         Width           =   240
      End
   End
   Begin VB.Frame FramePedConfirma 
      Height          =   4095
      Left            =   0
      TabIndex        =   535
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   116
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   544
         Text            =   "Text5"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   116
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   538
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   114
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   536
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarPedConfirma 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   539
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4320
         TabIndex        =   540
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   115
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   537
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   60
         Left            =   1035
         Picture         =   "frmListadoOfer.frx":90EF
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   92
         Left            =   480
         TabIndex        =   545
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "N� Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   600
         TabIndex        =   543
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label12 
         Caption         =   "Confirmaci�n entrega Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   11
         Left            =   600
         TabIndex        =   542
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   541
         Top             =   1560
         Width           =   900
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   11160
      Picture         =   "frmListadoOfer.frx":91F1
      Tag             =   "-1"
      ToolTipText     =   "Buscar cliente"
      Top             =   45
      Width           =   240
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
    '316:  Factura electronica. Realmente copiamos al PATH de parametros
    
    '400:  Clientes potenciales.  Cartas
    '401:                "       Etiquetas
    '402        "               CRM
    
    '406    Impresion masiva CRM
    
    '407 HERBELCA Impresion de un pedido de proveedor. NO tiene pregunta ni na, directamente al vispreort
    
    
    '408 Comprobacion cuentas bancarias(NIF) entre secciones
   
    '409  Listado ofertas historico
    
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta/pedido a imprimir
                    
Public codClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmBasico2
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoProve As frmBasico2
Attribute frmMtoProve.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmBasico2 'frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoZona As frmFacZonas
Attribute frmMtoZona.VB_VarHelpID = -1
Private WithEvents frmMtoRuta As frmFacRutas
Attribute frmMtoRuta.VB_VarHelpID = -1
Private WithEvents frmMtoSitua As frmFacSituaciones
Attribute frmMtoSitua.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoArtic As frmBasico2
Attribute frmMtoArtic.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'codigo postal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1



'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
Private NumeroDeCopias As Integer
'nuevo Febrero 2010
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vImprimedirecto As Boolean '
Private CadenaParaEnvioMail As String
'-------------------------------------



Dim indCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean

Dim SecuenciaEnProgressBar1 As Long
Dim MaxNumerRegistroSEnProgressBar1 As Long


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub cboMotTras_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboOrdVolVta_Click()
    Me.chkExportacion.visible = False
    If Me.chkVolumen.Value = 1 Then
        'Se ha marcado
        Me.chkExportacion.visible = Me.cboOrdVolVta.ListIndex = 1
    End If
End Sub


Private Sub cboPdteFacturarProv_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cboTipomov_Click(Index As Integer)
    If Index = 1 Then
        'Reimpresion de facturas
        'Alzira NO esta en esto
        If vParamAplic.TieneTelefonia2 > 1 Then chk_duplicado2(2).visible = Mid(cboTipomov(Index).Text, 1, 3) = "FAT"
    End If
End Sub

Private Sub cboTipomov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub chk_duplicado2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkBusHcoOfer_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkConfirmPed_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDatosAlbaranes_Click(Index As Integer)
    If Index = 0 Then
        Label4(90).Caption = "Fecha"
        If Me.chkDatosAlbaranes(0).Value = 1 Then
            Label4(90).Caption = Label4(90).Caption & " albaran"
     
        Else
            'Me.chkDatosAlbaranes(10).visible = False
        End If
        
    Else
        Label4(87).Caption = "Fecha"
        If Me.chkDatosAlbaranes(1).Value = 1 Then Label4(87).Caption = Label4(87).Caption & " albaran"
        
        If Me.chkDatosAlbaranes(10).Value = 1 Then Me.chkDatosAlbaranes(4).Value = 0
        
    End If
    
    
    If Index = 3 Then
        chkDatosAlbaranes(12).visible = chkDatosAlbaranes(3).Value = 1
    End If
    
End Sub

Private Sub chkDatosAlbaranes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDetallaArticulo_Click()
    Me.FrameDetallaArticulo.visible = chkDetallaArticulo.Value = 1
End Sub

Private Sub chkDetallaArticulo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub




Private Sub chkEtiqDpto_Click()
    If Me.lwCargos.ListItems.Count = 0 Then CargaCargos
End Sub

Private Sub chkFormatoTPV_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub



Private Sub chkFraRectCompleta_Click()

    Label4(131).visible = Me.chkFraRectCompleta.Value = 1
    imgBuscarOfer(97).visible = Me.chkFraRectCompleta.Value = 1
    txtCodigo(159).visible = Me.chkFraRectCompleta.Value = 1
    txtNombre(159).visible = Me.chkFraRectCompleta.Value = 1
    
    cmdCancel(11).Top = IIf(Me.chkFraRectCompleta.Value = 1, 6111, 4000)
    cmdAceptarFacRect.Top = IIf(Me.chkFraRectCompleta.Value = 1, 6111, 4000)
    FrameFacRectif.Height = IIf(Me.chkFraRectCompleta.Value = 1, 6855, 4695)
    Me.Height = FrameFacRectif.Height + 240
    Line2.BorderStyle = IIf(Me.chkFraRectCompleta.Value = 1, 1, 0)
    Label4(132).visible = Me.chkFraRectCompleta.Value = 1
    imgFecha(37).visible = Me.chkFraRectCompleta.Value = 1
    txtCodigo(160).visible = Me.chkFraRectCompleta.Value = 1
    txtCodigo(161).visible = Me.chkFraRectCompleta.Value = 1
    txtNombre(161).visible = Me.chkFraRectCompleta.Value = 1
    imgBuscarOfer(97).visible = Me.chkFraRectCompleta.Value = 1
    Label4(134).visible = Me.chkFraRectCompleta.Value = 1
    
    If Me.chkFraRectCompleta.Value = 1 Then txtCodigo(160).Text = Format(Now, "dd/mm/yyyy")
    
    DatosFraRectificativa
    
End Sub

Private Sub chkFraRectCompleta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkImprAlbProv_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkImpSaldo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMail_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPendientes_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkVarios_Click(Index As Integer)
    If Index = 3 Then
        'Listado de clientes
        If Me.chkVarios(3).Value = 1 And chkVolumen.Value = 1 Then
            
            MsgBox "No puede marcar a la vez 'con volumen ventas' y 'poblacion'", vbExclamation
            chkVolumen.Value = 0
            chkVolumen_Click
        End If
    End If
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii

End Sub

Private Sub chkVolumen_Click()
   
        'Listado de clientes
        If Me.chkVarios(3).Value = 1 And chkVolumen.Value = 1 Then
            MsgBox "No puede marcar a la vez 'con volumen ventas' y 'poblacion'", vbExclamation
            Me.chkVarios(3).Value = 0
        End If
 


    FrameVolumen.visible = chkVolumen.Value = 1
    cboOrdVolVta.visible = chkVolumen.Value = 1
    optClienteLis(0).visible = FrameVolumen.visible
    optClienteLis(1).visible = FrameVolumen.visible
    optClienteLis(2).visible = FrameVolumen.visible
    FrVolVetasCredito.visible = (chkVolumen.Value = 1) And vParamAplic.OperacionesAseguradas
    
    If Me.chkVolumen.Value = 0 Then Me.chkExportacion.visible = False
End Sub

Private Sub cmdAceptarAlbCom_Click()
'Solicitar datos para Generar Albaran  a partir de Pedido de Compras
Dim Cad As String
Dim CodClienVacio As Boolean

    'Feb. 2011
    Cad = ""
    If txtCodigo(47).Text = "" Or txtNombre(47).Text = "" Then Cad = Cad & "     -Provedor" & vbCrLf
    If txtCodigo(48).Text = "" Then Cad = Cad & "     -N� albar�n" & vbCrLf
    If txtCodigo(49).Text = "" Then Cad = Cad & "     -Fecha albar�n" & vbCrLf
    If txtCodigo(168).Text = "" Then Cad = Cad & "     -Fecha albar�n" & vbCrLf
    If Cad <> "" Then
        Cad = "Campos obligatorios: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    If Abs(DateDiff("m", CDate(txtCodigo(49).Text), CDate(txtCodigo(168).Text))) > 3 Then
        Cad = "La diferencia entre fecha albaran y entrada mercancia mayor 3 meses"
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
            Exit Sub
        Else
            Cad = Cad & vbCrLf & "�Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If

    

     Cad = " year(fechaalb) = " & Year(CDate(txtCodigo(49).Text)) & " AND numalbar=" & DBSet(txtCodigo(48).Text, "T") & " AND codprove"
     Cad = DevuelveDesdeBD(conAri, "fechaalb", "scaalp", Cad, CStr(codClien), "N")
     
     If Cad = "" Then
        Cad = " year(fechaalb)=" & Year(CDate(txtCodigo(49).Text)) & " AND numalbar=" & DBSet(txtCodigo(48).Text, "T") & " AND codprove"
        Cad = DevuelveDesdeBD(conAri, "fechaalb", "scafpa", Cad, CStr(codClien), "N")
        If Cad <> "" Then Cad = Cad & " (Facturado)"
     End If
     
     If Cad <> "" Then
        MsgBox "Ya existe para el proveedor el albaran indicado." & vbCrLf & "Fecha:" & Cad, vbExclamation
        Exit Sub
    End If
    
    Cad = txtCodigo(47).Text & "|"
    Cad = Cad & txtCodigo(48).Text & "|"
    Cad = Cad & txtCodigo(49).Text & "|"
    Cad = Cad & chkImprAlbProv(0).Value & "|"
    Cad = Cad & txtCodigo(168).Text & "|"
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub cmdAceptarCierre_Click()
'Cierre de caja del TPV
Dim campo As String
Dim devuelve As String


    InicializarVbles
    
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1


    'comprobar que se ha introducido FECHA
    '---------------------------------------------------------
    If Trim(txtCodigo(88).Text) <> "" Or Trim(txtCodigo(89).Text) <> "" Then
        'Para Crystal Report
        campo = "{scafac.fecfactu}"
        devuelve = "pDHFecha=""FECHA: " 'Parametro Desde/Hasta Fecha
        If Not PonerDesdeHasta(campo, "F", 88, 89, devuelve) Then Exit Sub
    Else
        MsgBox "Debe introducir la fecha de cierre.    ", vbExclamation
        Exit Sub
    End If
    
    
    '---- Seleccionar solo las facturas que vienen de TICKET del TPV
    If chkVarios(2).Value = 0 Then
        campo = "{scafac1.numventa}"
        campo = "(NOT ISNULL(" & campo & ")) and (" & campo & "<>0)"
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        
        campo = "{scafac1.numtermi} >0"
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        
     
        
    End If
    
    '---- seleccionar solo el tipo pago: 0=efectivo,2=talon, 3=pagare, 6=tarjeta credito
     campo = "{sforpa.tipforpa} in [0,2,3,6]"
     If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
     campo = "{sforpa.tipforpa} in (0,2,3,6)"
     If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    

    
    'ver si hay registros seleccionados para mostrar en el informe
    campo = "(scafac INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu)  INNER JOIN sforpa ON scafac.codforpa = sforpa.codforpa "
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    Titulo = "Cierre de Caja"
    If Me.optForpago(0).Value = True Then
        'informe por Forma de Pago
        nomRPT = "rTPVcierreFP.rpt"
    Else
        'informe por Tipo de Forma de Pago
        nomRPT = "rTPVcierre.rpt"
    End If
    conSubRPT = True
    LlamarImprimir False, False
     
End Sub

Private Sub cmdAceptarClien_Click()
'Listado de Clientes
Dim campo As String, devuelve As String
Dim numOp As Byte
Dim B As Boolean
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H ACTIVIDAD
    '--------------------------------------------
    
    'Aqui metere tb el desde hasta volumen ventas
    Codigo = ""
    If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{sclien.codactiv}"
        'Parametro Desde/Hasta Actividad
        devuelve = " Actividad: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Sub
        Codigo = devuelve
    End If
    If Me.chkVolumen.Value = 1 Then
        devuelve = ""
        If Me.txtCodigo(122).Text <> "" Then devuelve = devuelve & " desde " & Format(txtCodigo(122).Text, "dd/mm/yyyy")
        If Me.txtCodigo(123).Text <> "" Then devuelve = devuelve & " hasta " & Format(txtCodigo(123).Text, "dd/mm/yyyy")
        
        
        
        'Si tiene marcado credito
        If vParamAplic.OperacionesAseguradas Then
            If Me.cboClienteCredito.ListIndex > 0 Then
                devuelve = devuelve & " Credito:"
                If Me.cboClienteCredito.ListIndex = 1 Then
                    devuelve = devuelve & " Privado"
                    campo = "({sclien.credipriv} = 1)"
                    
                ElseIf Me.cboClienteCredito.ListIndex = 2 Then
                    devuelve = devuelve & " Aseguradora"
                    campo = "({sclien.credipriv} = 0)"
                ElseIf Me.cboClienteCredito.ListIndex = 3 Then
                    devuelve = devuelve & " Estudiado"
                    campo = "({sclien.credipriv} = 2)"
                Else
                    devuelve = devuelve & " NO asignado"
                    campo = "({sclien.credipriv} = 9)"
                End If
                If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                If cadSelect <> "" Then cadSelect = cadSelect & " AND "
                cadFormula = cadFormula & campo
                cadSelect = cadSelect & campo
                
            End If
        End If
        devuelve = "              Fecha ventas: " & devuelve
        Codigo = Trim(Codigo & devuelve)
    End If
    If Codigo <> "" Then
        Codigo = "pDHActividad=""" & Codigo & """|"
        cadParam = cadParam & Codigo
        numParam = numParam + 1
    End If
    'Cadena para seleccion D/H ZONA
    '--------------------------------------------
     If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{sclien.codzonas}"
        'Parametro Desde/Hasta Zona
        devuelve = "pDHZona=""Zona: "
        If Not PonerDesdeHasta(campo, "N", 35, 36, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H RUTA
    '--------------------------------------------
     If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = "{sclien.codrutas}"
        'Parametro Desde/Hasta Ruta
        devuelve = "pDHRuta=""Ruta: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H AGENTE    y VISITADOR
    '--------------------------------------------
    Titulo = ""
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "Agente: "
        If Not PonerDesdeHasta(campo, "N", 39, 40, devuelve) Then Exit Sub
        Titulo = Replace(devuelve, """", "")
    End If
    If txtCodigo(151).Text <> "" Or txtCodigo(152).Text <> "" Then
        campo = "{sclien.visitador}"
        'Parametro Desde/Hasta Agente
        devuelve = "Visitador: "
        If Not PonerDesdeHasta(campo, "N", 151, 152, devuelve) Then Exit Sub
        Titulo = Trim(Titulo & "    " & Replace(devuelve, """", ""))
    End If
    
    

    
    If Titulo <> "" Then
        devuelve = "pDHAgente="" " & Titulo & """|"
        cadParam = cadParam & devuelve
        numParam = numParam + 1
        Titulo = ""
    End If
    
    
    
    
    
    'Cadena para seleccion D/H SITUACION
    '--------------------------------------------
    Titulo = ""
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = "{sclien.codsitua}"
        'Parametro Desde/Hasta Situacion
        'devuelve = "pDHSituacion=""Situaci�n: "  '
        devuelve = "Situaci�n: "
        If Not PonerDesdeHasta(campo, "N", 41, 42, devuelve) Then Exit Sub
        Titulo = Replace(devuelve, """", "")
    End If
    
    
    If txtCodigo(129).Text <> "" Or txtCodigo(130).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta Agente
            devuelve = "C.Postal: "
            If Not PonerDesdeHasta(campo, "T", 129, 130, devuelve) Then Exit Sub
            Titulo = Trim(Titulo & "    " & Replace(devuelve, """", ""))
    End If
 
    If Titulo <> "" Then
        devuelve = "pDHSituacion="" " & Titulo & """|"
        cadParam = cadParam & devuelve
        numParam = numParam + 1
        Titulo = ""
    End If
    
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
    If Me.chkVolumen.Value = 0 Then
        If Me.chkVolumen.Value = 0 Then
            numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
            numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
            numOp = PonerGrupo(3, ListView1.ListItems(3).Text)
            numOp = PonerGrupo(4, ListView1.ListItems(4).Text)
        End If
    End If


    cadSelect = cadFormula
    If Not HayRegParaInforme("sclien", cadSelect) Then Exit Sub
     
    Screen.MousePointer = vbHourglass
    B = True
    If Me.chkVolumen.Value = 1 Then B = CalculaVolumenVtas_
    Screen.MousePointer = vbDefault
    
    If Not B Then Exit Sub
    
    If Me.chkVarios(3).Value = 1 Then
        'rFacClienCP.rpt
        nomRPT = "rFacClienCP.rpt"
    Else
        If Me.chkVolumen.Value = 0 Then
            'ESTE ES EL NORMAL
            nomRPT = "rFacClientes.rpt"
            
        Else
            
                
        
        
            
            'A�adimos codusu
            If cadFormula <> "" Then cadFormula = cadFormula & " AND "
            cadFormula = cadFormula & " ({tmpstockfec.codusu} = " & vUsu.Codigo & " )"
            
            'A�adimos el de emial
            devuelve = 0
            If Me.optClienteLis(1).Value Then devuelve = 1
            If Me.optClienteLis(2).Value Then devuelve = 2
            devuelve = "MuestrEmail=" & devuelve & "|"
            cadParam = cadParam & devuelve
            numParam = numParam + 1
        
            'Le calculo el volumen de ventas
            If cboOrdVolVta.ListIndex <= 0 Then
                nomRPT = "rFacClienAgeVol.rpt"
                  
            Else
                If Me.chkExportacion.Value = 1 Then
                    nomRPT = "rFacClienAgeExp.rpt"
                Else
                    nomRPT = "rFacClienAgeVol2.rpt"
                End If
                
            End If
        End If
    End If
    LlamarImprimir False, False
    nomRPT = ""
End Sub


Private Sub cmdAceptarClienInac_Click()
Dim EsInactividad As Boolean
'46: Informe de Clientes Inactivos
'47: Informe de Altas Nuevos Clientes
'90: Informe Etiquetas de clientes
Dim campo As String, devuelve As String
Dim J As Integer


    InicializarVbles
    
    EsInactividad = False     'reultizaremos la opcion 48 para k imprima tb el listado de inactividad
    If OpcionListado = 46 Then
        'Comprobar que se ha introdicido una FECHA de Inactividad
        If txtCodigo(31).Text = "" Then
            MsgBox "Debe introducir la Fecha de Inactividad para el informe.", vbInformation
            Exit Sub
        End If

        EsInactividad = True
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacClienInactivos.rpt"
        Titulo = "Clientes Inactivos"
        
    ElseIf OpcionListado = 48 Then
        'Comprobar si se ha introducido D/H FECHA Alta
        If txtCodigo(31).Text = "" And txtCodigo(32).Text = "" Then
            MsgBox "Debe introducir alg�n intervalo de Fechas de Alta.", vbInformation
            Exit Sub
        End If
        'Nombre fichero .rpt a Imprimir
        Titulo = "Altas Nuevos Clientes"
        nomRPT = "rFacClienAltas.rpt"
    End If
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(27).Text <> "" Or txtCodigo(28).Text <> "" Then
        campo = "{sclien.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 27, 28, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 29, 30, devuelve) Then Exit Sub
    End If
    
    
    
    If OpcionListado = 90 Or OpcionListado = 91 Then '90: Etiquetas de clientes
                                                     '91: Cartas a clientes
        'Cadena para seleccion D/H ACTIVIDAD
        '--------------------------------------------
         If txtCodigo(53).Text <> "" Or txtCodigo(54).Text <> "" Then
            campo = "{sclien.codactiv}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 53, 54, devuelve) Then Exit Sub
        End If
                        
        'Cadena para seleccion D/H COD. POSTAL
        '--------------------------------------------
         If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta cod. Postal
            devuelve = "pDHcpostal=""CPostal: "
            If Not PonerDesdeHasta(campo, "T", 55, 56, devuelve) Then Exit Sub
        End If
        
        'Cadena para seleccion SITUACION
        '--------------------------------------------
        If txtCodigo(57).Text <> "" Then
            campo = "{sclien.codsitua}=" & txtCodigo(57).Text
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        End If
        
        
        'ENERO 2010
        'Si no tiene  la marca de correo NO puede seleccionar cliente
        If OpcionListado = 91 Then
            campo = "{sclien.enviocorreo}=1"
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        End If
        
        
        If OpcionListado = 90 Then
            If Me.chkEnviaCorreo.Value = 1 Then
                campo = "{sclien.enviocorreo}=1"
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
            
            If Me.chkEtiqDpto.Value = 1 Then
                'Va a salir con cargos
                If Me.lwCargos.ListItems.Count > 0 Then
                    campo = ""
                    devuelve = ""
                    For J = 1 To Me.lwCargos.ListItems.Count
                        If lwCargos.ListItems(J).Checked Then
                            campo = campo & ", " & DBSet(lwCargos.ListItems(J).Text, "T")
                        Else
                            'NO ha seleccionado todo
                            devuelve = "NO"
                        End If
                    Next
                    If campo = "" Then
                        MsgBox "Ha seleccionado [" & Me.chkEtiqDpto.Caption & "]" & vbCrLf & "pero no ha seleccionado ninguno", vbExclamation
                        Exit Sub
                    End If
                    
                    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                    If devuelve <> "" Then
                        'Significa que no ha seleccionado todos
                        campo = Mid(campo, 2)
                        cadFormula = cadFormula & "  sclien.codclien IN (select codclien from scliendp where cargo IN (" & campo & "))"
                    Else
                        'Son todos..... pero los que tienen cargos
                        cadFormula = cadFormula & "  sclien.codclien IN (select distinct(codclien) from scliendp WHERE not cargo is null)"
                    End If
                End If
            End If
        End If
        'Parametro a la Atencion de
        cadParam = cadParam & "pAtencion=""Att. " & txtCodigo(0).Text & """|"
        numParam = numParam + 1
        
        'seleccionamos todos los clientes por defecto,
        'pero si seleccionamos clientes con mantenimientos o sin mantenimientos
         'Comprobar si hay registros a Mostrar antes de abrir el Informe
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
        devuelve = ""
        If Me.OptCliConMante Then
            devuelve = ListaClientesMante(cadSelect)
            If devuelve <> "" Then
                cadFormula = "{sclien.codclien} IN [" & devuelve & "]"
                cadSelect = "sclien.codclien IN (" & devuelve & ")"
            Else
                MsgBox "No existen clientes con esos valores", vbExclamation
                Exit Sub
            End If
        ElseIf Me.OptCliSinMante Then
            devuelve = ListaClientesMante(cadSelect)
            If devuelve <> "" Then
                campo = " NOT( {sclien.codclien}  IN [" & devuelve & "])"
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                campo = " sclien.codclien NOT IN (" & devuelve & ")"
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
        End If
        
        If OpcionListado = 90 Then
            
            devuelve = ListaClientesDesdeHastaFactura2()
            'Puede haber puesto desde hasta datos factura
            If devuelve <> "" Then
                campo = " ( {sclien.codclien}  IN [" & devuelve & "])"
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
        End If
        
        
        
        
        
        
        
        
        
        If OpcionListado = 90 Then 'Etiquetas
            'Nombre fichero .rpt a Imprimir
            
            'NUEVO. Igual deberiamos utilizar la clase: CParamTPV
            nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "20")
            If nomRPT = "" Then nomRPT = "rFacClienEtiq.rpt"
            
            'Si tiene la marca de DEPARTAMENTO
            If Me.chkEtiqDpto.Value = 1 Then
                nomRPT = Mid(nomRPT, 1, Len(nomRPT) - 4)
                nomRPT = nomRPT & "dpto.rpt"
            End If
            
            
            Titulo = "Etiquetas de Clientes"
            conSubRPT = False
        Else '91: CARTA/e-MAil
            'Parametro cod. carta
            cadParam = cadParam & "pCodCarta= " & txtCodigo(64).Text & "|"
            numParam = numParam + 1
            
            'Nombre fichero .rpt a Imprimir
            nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "57")
            If nomRPT = "" Then nomRPT = "rFacClienCarta.rpt"
            
           
            Titulo = "Cartas a Clientes"
            conSubRPT = True
        End If
    Else
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
    End If
    
    If OpcionListado = 46 Then
        'Seleccionar aquellos cliente que campo sclien.fechamov < fecha Inactividad
        If txtCodigo(31).Text <> "" Then
            campo = "sclien.fechamov"
            devuelve = txtCodigo(31).Text
            devuelve = "(isnull({sclien.fechamov}) or {" & campo & "} < Date(" & Year(devuelve) & ", " & Month(devuelve) & ", " & Day(devuelve) & "))"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            devuelve = "(" & campo & " < '" & Format(txtCodigo(31).Text, FormatoFecha) & "' OR isnull(sclien.fechamov))"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
            devuelve = "pFechaMov=""" & txtCodigo(31).Text & """|"
            cadParam = cadParam & devuelve
            numParam = numParam + 1
        End If
        
        devuelve = ""
        campo = ""
        Codigo = ""
        For J = 1 To Me.lwTippa.ListItems.Count
            If lwTippa.ListItems(J).Checked Then
                devuelve = devuelve & "1"
                campo = campo & ", " & Mid(lwTippa.ListItems(J).Key, 2)
                Codigo = Codigo & "- " & lwTippa.ListItems(J).Text & " "
                
            End If
        Next
        
        J = Len(devuelve)
        If J <> lwTippa.ListItems.Count Then
            If J = 0 Then
                MsgBox "Seleccione algun tipo de pago", vbExclamation
                Exit Sub
            Else
                'Ha marcado alguno(no todos=
                campo = Mid(campo, 2)
                Codigo = "Tipo de pago: " & Mid(Codigo, 2)
                devuelve = "{sforpa.tipforpa} IN [" & campo & "]"
                
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                devuelve = " sclien.codforpa IN (select codforpa from sforpa WHERE tipforpa IN [" & campo & "])"
                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
                devuelve = "pTipoPago=""" & Codigo & """|"
                cadParam = cadParam & devuelve
                numParam = numParam + 1
            End If
        End If
    ElseIf OpcionListado = 48 Then
        'Cadena para seleccion D/H FECHA
        '--------------------------------------------
        If txtCodigo(31).Text <> "" Or txtCodigo(32).Text <> "" Then
          'Para Crystal Report
            campo = "{sclien.fechaalt}"
            'Parametro Desde/Hasta Fecha
            devuelve = "pDHFecha=""Fecha Alta: "
            If Not PonerDesdeHasta(campo, "F", 31, 32, devuelve) Then Exit Sub
        End If
    End If
        
    If Not HayRegParaInforme("sclien", cadSelect) Then Exit Sub
    
    If OpcionListado = 90 Or OpcionListado = 91 Then 'OpcionListado = 90 'Etiquetas clientes
        Set frmMen = New frmMensajes
        frmMen.cadWhere = cadSelect
        frmMen.OpcionMensaje = 8 'Etiquetas clientes
        frmMen.Show vbModal
        Set frmMen = Nothing
        If cadSelect = "" Then Exit Sub
        
        If OpcionListado = 91 And Me.chkEMAIL(1).Value = 1 Then
            'Enviarlo por e-mail
            campo = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "57")
            If campo = "" Then campo = "rFacClienCarta.rpt"
            EnviarEMailMulti cadSelect, Titulo, campo, "sclien" 'email para clientes
        Else
            If OpcionListado = 90 Then
                'Si ha seleccionado cargos hay que pasar el select
                If Me.chkEtiqDpto.Value = 1 Then
                    'Va a salir con cargos
                    If Me.lwCargos.ListItems.Count > 0 Then
                        campo = ""
                        devuelve = ""
                        For J = 1 To Me.lwCargos.ListItems.Count
                            If lwCargos.ListItems(J).Checked Then
                                campo = campo & ", " & DBSet(lwCargos.ListItems(J).Text, "T")
                            Else
                                'NO ha seleccionado todo
                                devuelve = "NO"
                            End If
                        Next
                        
                        
                        
                        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                        
                        If devuelve <> "" Then
                            'Significa que no ha seleccionado todos
                            campo = Mid(campo, 2)
                            cadFormula = cadFormula & "  cargo IN [" & campo & "]"
                        Else
                            'ha seleccionado todos los cargos
                            cadFormula = cadFormula & "  not cargo is null"
                        End If
                        
                        'Borro tmp
                        conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
                        
                        'Inserto el select
                        devuelve = "insert into tmpinformes(codusu,codigo1,campo1) "
                        devuelve = devuelve & "Select " & vUsu.Codigo & ",sclien.codclien,id from sclien,scliendp WHERE "
                        devuelve = devuelve & " scliendp.codclien=sclien.codclien AND not cargo is null AND "
                        cadSelect = QuitarCaracterACadena(cadFormula, "{")
                        cadSelect = QuitarCaracterACadena(cadSelect, "}")
                        cadSelect = Replace(cadSelect, "[", "(")
                        cadSelect = Replace(cadSelect, "]", ")")
                        devuelve = devuelve & cadSelect
                        conn.Execute devuelve
                        
                        'En el rpt solo tmpinformes.codusu
                        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                     End If
                End If
                
            End If  'de opcion=90... etiquetas cliente
            'Octubre 2011
            conSubRPT = True
            
            LlamarImprimir False, False
        End If
    Else
    
        If OpcionListado = 46 Then OpcionListado = 48             'para que lo imprima
    
        LlamarImprimir False, False
        
        If EsInactividad Then
            OpcionListado = 46 'reestblezco
            
            'Abro un from para que cambia la siutacion de Todos los clientes de ese listado
            cadFormula = cadSelect & " AND 1"
            cadFormula = DevuelveDesdeBD(conAri, "count(*)", "sclien", cadFormula, "1")
            CadenaDesdeOtroForm = ""
            frmListado3.Opcion = 75
            frmListado3.OtrosDatos = cadFormula
            frmListado3.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                CadenaDesdeOtroForm = "UPDATE sclien SET codsitua = " & CadenaDesdeOtroForm & " WHERE "
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & cadSelect
                ejecutar CadenaDesdeOtroForm, False
                
               
            End If
             Unload Me
        End If
    End If
    
End Sub


Private Sub cmdAceptarCompras_Click()
'Listados de Compras
Dim campo As String
Dim Cad As String
Dim tabla As String
Dim HayDatos As Boolean
Dim CadSelConRappel As String


    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    
    CadSelConRappel = ""
    tabla = ""
    If chkDatosAlbaranes(11).Value Then tabla = "(Rappel)"       'SOLO con rappel
        
        
    
    If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{scafpc.codprove}"
        'Parametro Desde/Hasta Proveedor
        Cad = "pDHProve=""Proveedor" & tabla & ": "
        If Not PonerDesdeHasta(campo, "N", 90, 91, Cad) Then Exit Sub
    End If
   
    If chkDatosAlbaranes(11).Value = 1 Then
        If cadSelect <> "" Then
            Cad = Replace(cadSelect, "(", "")
            Cad = Replace(Cad, ")", "")
            Cad = Replace(Cad, "{", "")
            Cad = Replace(Cad, "}", "")
        Else
            Cad = " TRUE "
            cadParam = cadParam & "pDHProve=""Proveedor: RAPPEL |"
        End If
        Cad = "select distinct codprove from scafpc where " & Cad
        Cad = Cad & " and codprove in (select codprove from sdtomp where rap1>0 )"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cad = ""
        While Not miRsAux.EOF
            Cad = Cad & ", " & miRsAux!Codprove
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        If Cad = "" Then
            MsgBox "Ningun proveedor con rappel para mostrar en el listado", vbExclamation
            Exit Sub
        End If
        Cad = Mid(Cad, 2)
        
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & " {scafpc.codprove} IN [" & Cad & "]"
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & " scafpc.codprove IN (" & Cad & ")"
        
        
        Cad = ""
    End If
    
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
        'Para fam/articulo con albaranaes
        campo = "{scafpc.fecfactu}"
        Cad = " "
        If Not PonerDesdeHasta(campo, "F", 92, 93, Cad) Then Exit Sub
        Cad = "pDHFecha=""Fecha Fact.: " & Cad
        If chkDatosAlbaranes(1).Value Then Cad = Cad & "  Albaranes"
        cadParam = cadParam & Cad & """|"
    End If
    
    tabla = "scafpc"
    If OpcionListado = 311 Then
    
        'Si marca la casilla de rappel tiene que tener marcado la de articulo(ya que es familia MARCA)
        If OptCompras(0).Value And Me.chkDatosAlbaranes(2).Value = 1 Then
            'MsgBox "Para mostrar datos rappel debe marcar familia/art�culo", vbExclamation
            'Exit Sub
        End If
        
        'Para los listado que salen articulo no se puede NO detallar
        If chkDatosAlbaranes(7).Value = 1 Then
            If chkDatosAlbaranes(2).Value = 1 Then
                'RAPPEL. SIEMPRE SALDRA EL PROVEEDOR, aunque no este marcado
                
            
            Else
                If OptCompras(1).Value Then
                    MsgBox "Si muestra articulo debe detallar proveedor", vbExclamation
                    Exit Sub
                End If
            End If
        Else
            If chkDatosAlbaranes(2).Value = 1 Then
                'En los de rappel siempre sale el proveedor
                MsgBox "En rappel siempre agrupa por proveedor", vbExclamation
                chkDatosAlbaranes(7).Value = 1
            End If
        End If
        
        
        'Si marca comparativo, tiene que ser por familia
        'y , de momento,
        'admeas debe indcar una fecha dentro de un a�o
        If chkDatosAlbaranes(6).Value = 1 Then
            Cad = ""
            indCodigo = 92
            'Obligado desde/hasta fecha
            If txtCodigo(92).Text = "" Or txtCodigo(93).Text = "" Then
                Cad = "-Debe indicar las fechas del listado"
            Else
                If Year(txtCodigo(92).Text) <> Year(txtCodigo(93).Text) Then Cad = "-Un a�o como maximo"
            End If
            If Cad = "" Then indCodigo = 0
            If OptCompras(1).Value Then Cad = Cad & vbCrLf & "-No puede detallar articulo"
            If chkDatosAlbaranes(2).Value = 1 Then Cad = Cad & vbCrLf & "-No debe marcar el rappel"
            If Cad <> "" Then
                MsgBox "Comparativo: " & vbCrLf & Cad, vbExclamation
                If indCodigo = 92 Then PonerFoco txtCodigo(92)
                Exit Sub
            End If
        End If
        
    
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            Cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 94, 95, Cad) Then Exit Sub
            
            
            If Me.chkDatosAlbaranes(1).Value = 0 Then
                tabla = "( scafpc INNER JOIN slifpc ON scafpc.codprove=slifpc.codprove AND scafpc.numfactu=slifpc.numfactu "
                tabla = tabla & " AND scafpc.fecfactu=slifpc.fecfactu )"
                tabla = tabla & " INNER JOIN sartic ON slifpc.codartic=sartic.codartic "
        
        
            Else
                
            
            
            
            End If
        
        End If
        
        'Si no va con albaranes
        If Me.txtCodigo(126).Text <> "" Then
            'Solo valido para comparativo
            If chkDatosAlbaranes(6).Value = 1 Then
                'Se hace en el tminsertardatos
''''''                campo = "{scafpc.brutofac} >= " & TransformaComasPuntos(CStr(ImporteFormateado(txtCodigo(126).Text)))
''''''                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
''''''                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            Else
                If Me.chkDatosAlbaranes(2).Value = 1 Then
                    'para rappel tab va el minimo
                Else
                    MsgBox "Importe minimo v�lido para listado comparativo/rappel", vbExclamation
                End If
            End If
        End If
    End If
        
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If OpcionListado = 312 Then
        'en esta opcion ver si hay albaranes
        cadSelect = Replace(cadSelect, tabla, "scafpa")
        cadSelect = Replace(cadSelect, "fecfactu", "fechaalb")
        tabla = "scafpa"
    End If
    
    'Para fam/articulo con albaranaes
    If OpcionListado = 311 And Me.chkDatosAlbaranes(1).Value = 1 Then
        'Es un contador de un UNION.
        'Hay que comprobar si hay reg en factuaras Y albaranes
        If Not ContadorDelUnion(False) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    
    Else
        If Not HayRegParaInforme(tabla, cadSelect) Then
            If OpcionListado <> 312 Then Exit Sub
        
            If Not HayRegParaInforme("scaalp", cadSelect) Then Exit Sub
        End If
    End If
    
    If OpcionListado = 312 Then
    
    
        'insertar en la tmpInformes
        BorrarTempInformes
        
        'en esta opcion ver si hay albaranes
        Cad = Replace(cadSelect, tabla, "scaalp")
        Cad = Replace(Cad, "fecfactu", "fechaalb")
        
        'insertar los albaranes q cumple la seleccion
        If Not CargarTmpInformes_Compras_312("scaalp", Cad) Then Exit Sub
        
        
        'insertar los albaranes de facturas q cumple la seleccion
        If Not CargarTmpInformes_Compras_312(tabla, cadSelect) Then Exit Sub
        
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
    End If
    
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    conSubRPT = False
    If OpcionListado = 311 Then
        Cad = "Salta= " & Me.chkDatosAlbaranes(5).Value & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        If Me.chkDatosAlbaranes(2).Value = 0 Then
            'Cuando NO VA con rappel
            Cad = "0"
            If chkDatosAlbaranes(7).Value = 0 Then Cad = "1"
            Cad = "detalla= " & Cad & "|"
            cadParam = cadParam & Cad
            numParam = numParam + 1
       Else
            'RAPPEL, puede mostrar o no los detalles de articulos
            'DetaArtic
            Cad = "0"
            If OptCompras(1).Value Then Cad = "1"
            Cad = "DetaArtic= " & Cad & "|"
            cadParam = cadParam & Cad
            numParam = numParam + 1
            
       End If
        
        
        'El rpt este bien
        If Me.OptCompras(0).Value = True Then
            nomRPT = "rComEstProFam"
            Titulo = "Listado Compras por Familia"
            conSubRPT = True
        Else
            nomRPT = "rComEstProArt"
            Titulo = "Listado Compras por Art�culo"
        End If
        
        'rappel
        HayDatos = True
        If Me.chkDatosAlbaranes(2).Value = 1 Then
            PreparaDatosLineasCompras
            If NumRegElim = 0 Then HayDatos = False
        End If
        'Comparativo
        If Me.chkDatosAlbaranes(6).Value = 1 Then
            ponerLineasComprasComparatativo
            If Me.txtCodigo(126).Text <> "" Then
                'Veo si quedan registros
                Codigo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
                If Codigo = "" Then Codigo = "0"
                If Val(Codigo) = 0 Then
                    HayDatos = False
                     MsgBox "No existen datos", vbExclamation
                End If
            End If
        End If
        If Not HayDatos Then
            Label9(38).Caption = ""
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If Me.chkDatosAlbaranes(1).Value = 1 Then
         
            'Si no va con rrapel
            If Me.chkDatosAlbaranes(2).Value = 0 Then
                'ANTES
                'Utilizamos tmps
                If Not CargaDatosEstadComprasCOMMAND Then Exit Sub
                cadFormula = "{tmpcommand.codusu} = " & vUsu.Codigo
            End If
            nomRPT = nomRPT & "alb"
            Titulo = Titulo & " (con albaranes)"
        End If
        
        '   con rappel                       o compartaivo
        If Me.chkDatosAlbaranes(2).Value = 1 Or Me.chkDatosAlbaranes(6).Value = 1 Then
            If Me.chkDatosAlbaranes(6).Value = 1 Then
                'Comparativo
                Cad = "Detalla= " & Abs(OptCompras(1).Value) & "|"
            Else
                'Rappell
                Cad = 1
  '              If chkDatosAlbaranes(3).Value = 0 Then Cad = 0
                Cad = "Detalla= " & Cad & "|"
            End If
            cadParam = cadParam & Cad
            numParam = numParam + 1
            'Solo hay un rpt para los rappels
            If Me.chkDatosAlbaranes(2).Value = 1 Then
                nomRPT = "rComEstProArtrap"
                Titulo = Titulo & " (rappel)"
                cadFormula = "{tmpcommand.codusu} = " & vUsu.Codigo
            Else
                nomRPT = "rComEstProCompara"
                Titulo = Titulo & " (comparativo)"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            End If
        End If
        
        
        'Ordenado por nombre proveedor
        If Me.chkVarios(9).Value = 1 Then nomRPT = nomRPT & "Nom"
        nomRPT = nomRPT & ".rpt"
        
        
    ElseIf OpcionListado = 310 Then
        nomRPT = "rComEstProImp.rpt"
        Titulo = "Listado Compras por Proveedor"
    Else '312: Albaranes x porveedor
        Cad = "0"
        If chkDatosAlbaranes(7).Value = 0 Then Cad = "1"
        Cad = "detalla= " & Cad & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
    
    
        nomRPT = "rComEstProAlb.rpt"
        Titulo = "Listado albaranes por Proveedor"
    End If
    
    
    LlamarImprimir False, False
    
    If OpcionListado = 312 Then BorrarTempInformes
End Sub

Private Sub cmdAceptarEfect_Click()
    HacerEfectuadas
    Label4(96).Caption = ""
End Sub

Private Sub HacerEfectuadas()

'Ofertas Efectuadas
Dim Cad As String
Dim TotOfeA As String 'N� total de Ofertas Aceptadas del Periodo( en cabecera y en historico)
Dim TotImpBA As String 'Importe Bruto total de Ofertas Aceptadas del Periodo (en cabecera e historico)
Dim TotOfeNA As String 'N� total de Ofertas NO Aceptadas del Periodo( en cabecera y en historico)
Dim TotImpBNA As String 'Importe Bruto total de Ofertas NO Aceptadas del Periodo (en cabecera e historico)
Dim C2 As String

    'Inicializar vbles
    InicializarVbles
    
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "NOHco=" & Abs(Me.chkBusHcoOfer.Value) & "|"
    numParam = numParam + 1
    
    '===================================================
    '============ PARAMETROS ===========================
    If OpcionListado = 34 Then
        'Imprimir solo las Ofertas Pendientes
        If Me.chkPendientes.Value = 1 Then
            Cad = "True"
        Else
            Cad = "False"
        End If
        cadParam = cadParam & "|pPtes=" & Cad & "|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacOfeEfectuadas.rpt"
        Titulo = "Ofertas Efectuadas"
    Else
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rAdmGastosTec.rpt"
        Titulo = "Gastos T�cnicos"
        
        
        
    End If
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    C2 = ""
     If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{schpre.fecofert}"
        Else
            Codigo = "{sgaste.fecgasto}"
        End If
        'Parametro Desde/Hasta FEcha
        'cad = "pDHFecha=""Fecha: "
        Cad = ""
        If Not PonerDesdeHasta(Codigo, "F", 16, 17, Cad) Then Exit Sub
        Cad = AnyadirParametroDH("Fecha:  ", 16, 17)
        C2 = C2 & Cad
    End If
    
    If txtCodigo(166).Text <> "" Or txtCodigo(167).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{schpre.fecentre}"
        Else
            Codigo = "{sgaste.fecgasto}"
        End If
        'Parametro Desde/Hasta FEcha
        'cad = "pDHFecha=""Fecha: "
        Cad = ""
        If Not PonerDesdeHasta(Codigo, "F", 166, 167, Cad) Then Exit Sub
        
        Cad = AnyadirParametroDH(IIf(OpcionListado = 34, IIf(InstalacionEsEulerTaxco, "F. reclamacion: ", "F. entrega: "), "FECHA NO SE QUE "), 166, 167)
        C2 = Trim(C2 & "     " & Cad)
    End If
    
    
    
    
    
    
    'Trabajador
    '--------------------------------------------
     If txtCodigo(117).Text <> "" Or txtCodigo(118).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{schpre.codtraba}"
        Else
            Codigo = "{sgaste.codtecni}"
        End If
        
        'Parametro Desde/Hasta FEcha
        Cad = ""
        If Not PonerDesdeHasta(Codigo, "N", 117, 118, Cad) Then Exit Sub
        Cad = AnyadirParametroDH("Trab: ", 117, 118)
        
        C2 = Trim(C2 & "    " & Cad)
    End If
    
    If C2 <> "" Then
            Cad = "pDHFecha= """ & C2
            cadParam = cadParam & Cad & """|"
            numParam = numParam + 1
    End If
    
    If OpcionListado = 34 Then
            Cad = IIf(OpcionListado = 34, IIf(InstalacionEsEulerTaxco, "F. reclama.", "F. entrega"), "F. No se")
            Cad = "pTextoFentre= """ & Cad
            cadParam = cadParam & Cad & """|"
            numParam = numParam + 1
        
   
   
   
   
        If Me.chkPendientes.Value = 0 Then 'Se muestra resumen si SoloPEndientes=false
        
            Label4(96).Caption = "Obtener datos 1"
            Label4(96).Refresh
            
            
            'Me guardo el cadselect. Primero en Cad y luego en codigo
            Cad = cadSelect
            
            
            Codigo = "scapre.fecofert"
            cadSelect = CadenaDesdeHastaBD(txtCodigo(16).Text, txtCodigo(17).Text, Codigo, "F")
            
            Codigo = Cad
            
            'Obtener total N� Ofertas del Periodo seleccionado y
            'el total Importe Bruto de las Ofertas de Periodo seleccionado
            'y pasarlo al Informe como parametro
            If ObtenerTotalOferPeriodo(cadSelect, TotImpBA, TotImpBNA, TotOfeA, TotOfeNA) Then
                Cad = "pTotPeriodoOfeA="""
                cadParam = cadParam & Cad & TotOfeA & """|"
                Cad = "pTotPeriodoOfeNA="""
                cadParam = cadParam & Cad & TotOfeNA & """|"
                Cad = "pTotPeriodoImpA="""
                cadParam = cadParam & Cad & TotImpBA & """|"
                Cad = "pTotPeriodoImpNA="""
                cadParam = cadParam & Cad & TotImpBNA & """|"
                numParam = numParam + 4
            End If
            
            'Retomamos el cadselcet
            cadSelect = Codigo
            
        End If
    End If
    
    'Cadena para seleccion Desde y Hasta AGENTE
    '--------------------------------------------
    C2 = ""
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        If OpcionListado = 34 Then
            Codigo = "{schpre.codagent}"
        Else
            Codigo = "{sgaste.codtecni}"
        End If
        Cad = "Agente: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, Cad) Then Exit Sub
        Cad = AnyadirParametroDH("Agen: ", 18, 118)
        C2 = Cad
        
    End If
    If OpcionListado = 34 Then
         'ZONA
        If txtCodigo(153).Text <> "" Or txtCodigo(154).Text <> "" Then
            Codigo = "{schpre.codclien}"
            Cad = "Zona : "
            If Not PonerDesdeHasta(Codigo, "N", 153, 154, Cad) Then Exit Sub
            Cad = AnyadirParametroDH("Zona: ", 153, 154)
            C2 = Trim(C2 & "     " & Cad)
        End If
        
        If txtCodigo(155).Text <> "" Or txtCodigo(156).Text <> "" Then
            Codigo = "{schpre.codclien}"
            Cad = "Cliente: "
            If Not PonerDesdeHasta(Codigo, "N", 155, 156, Cad) Then Exit Sub
            
            Cad = AnyadirParametroDH("Cliente : ", 155, 156)
            If C2 <> "" Then C2 = C2 & """ + chr(13) + """
            C2 = Trim(C2 & Cad)
        End If
        
       
    End If
    
    If C2 <> "" Then
            Cad = "pDHAgente= """ & C2
            cadParam = cadParam & Cad & """|"
            numParam = numParam + 1
    End If
     
        
    If Me.chkPendientes.visible And Me.chkPendientes.Value Then
        'Cadena para seleccion de Ofertas no Aceptadas
        Codigo = "{schpre.aceptado}=0"
        If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
        
        
        Codigo = "(schpre.aceptado)=0"
        If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    End If


    '==============================================
    'Modificacion ENERO 2010
    ' creamos dos tmps
    Screen.MousePointer = vbHourglass
    If OpcionListado = 34 Then
        Label4(96).Caption = "Borrando temporales"
        Label4(96).Refresh
        
        conn.Execute "DELETE from tmpscapreu WHERE codusu = " & vUsu.Codigo
        conn.Execute "DELETE from tmpslipreu WHERE codusu = " & vUsu.Codigo

        If cadSelect <> "" Then
            cadSelect = Replace(cadSelect, "scapre", "schpre")
            cadSelect = Replace(cadSelect, "{", "")
            cadSelect = Replace(cadSelect, "}", "")
        End If


        'Cabecera
        For indCodigo = 1 To 2
            NumeroDeCopias = 1
            If indCodigo = 2 Then
                If Me.chkBusHcoOfer.Value = 0 Then NumeroDeCopias = 0
            End If
            
            If NumeroDeCopias = 1 Then
                Label4(96).Caption = "Insertar cab " & indCodigo
                Label4(96).Refresh
                Codigo = "INSERT INTO tmpscapreu(codusu,numofert, fecofert, aceptado, codclien, nomclien, codtraba, codagent, dtoppago, dtognral,fecentre)"
                Codigo = Codigo & " select " & vUsu.Codigo & ","
                Codigo = Codigo & "numofert, fecofert, aceptado, codclien, nomclien, codtraba, codagent, dtoppago, dtognral,fecentre"
                Codigo = Codigo & " FROM "
                If indCodigo = 1 Then Codigo = Codigo & " scapre"
                Codigo = Codigo & " schpre"
                If cadSelect <> "" Then Codigo = Codigo & " WHERE " & cadSelect
                conn.Execute Codigo
                
                Label4(96).Caption = "Insertar lin " & indCodigo
                Label4(96).Refresh
                
                'Las lineas
                Codigo = "INSERT INTO tmpslipreu(codusu,numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel)"
                Codigo = Codigo & " SELECT " & vUsu.Codigo & ",numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel "
                Codigo = Codigo & " FROM "
                If indCodigo = 1 Then Codigo = Codigo & " slipre"
                Codigo = Codigo & " slhpre "
                Codigo = Codigo & " WHERE numofert in (Select numofert FROM "
                If indCodigo = 1 Then Codigo = Codigo & " scapre"
                Codigo = Codigo & " schpre "
                If cadSelect <> "" Then Codigo = Codigo & " WHERE " & cadSelect
                Codigo = Codigo & ")"
                conn.Execute Codigo
            End If
       
       Next indCodigo
       
       
       NumeroDeCopias = 1
       
       'Updateamos ahora los nombres de los agentes y de los trabajadores
       Set miRsAux = New ADODB.Recordset
       Label4(96).Caption = "Trabajadores"
       Label4(96).Refresh
            
       Codigo = "select codtraba from tmpscapreu where codusu = " & vUsu.Codigo & " GROUP BY 1"
       miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       While Not miRsAux.EOF
            Cad = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", CStr(miRsAux.Fields(0)))
            If Cad = "" Then Cad = "  ***"
            Codigo = "UPDATE tmpscapreu SET nomtraba=" & DBSet(Cad, "T") & " WHERE codtraba = " & miRsAux.Fields(0) & "  AND codusu = " & vUsu.Codigo
            conn.Execute Codigo
            miRsAux.MoveNext
       Wend
       miRsAux.Close
       Label4(96).Caption = "Agentes"
       Label4(96).Refresh
       Codigo = "select codagent from tmpscapreu where codusu = " & vUsu.Codigo & " GROUP BY 1"
       miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       While Not miRsAux.EOF
            Cad = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", CStr(miRsAux.Fields(0)))
            If Cad = "" Then Cad = "  ***"
            Codigo = "UPDATE tmpscapreu SET nomagent=" & DBSet(Cad, "T") & " WHERE codagent = " & miRsAux.Fields(0) & "  AND codusu = " & vUsu.Codigo
            conn.Execute Codigo
            miRsAux.MoveNext
       Wend
       miRsAux.Close
       
       
       indCodigo = 0
        Label4(96).Caption = ""
        Label4(96).Refresh
       Screen.MousePointer = vbDefault

       nomRPT = "rFacOfeEfectuadas3.rpt"
       cadFormula = "{tmpscapreu.codusu} = " & vUsu.Codigo
    End If



    '==============================================
    conSubRPT = False
    LlamarImprimir False, False
End Sub


Private Sub cmdAceptarEstVentas_Click()
'Listados estadistica ventas por familia
'Listados de Compras
Dim campo As String
Dim Cad As String
Dim tabla As String
Dim HePuestoElJoinConSclien As Boolean
Dim Cadselect2 As String
Dim F2 As Date
Dim DesdeTmp As Boolean

Dim B As Boolean

    InicializarVbles
    
    Cadselect2 = ""
    DesdeTmp = False
    If OpcionListado = 230 Then
        'Vaciamos la tabla
        conn.Execute "DELETE FROM tmpcommandest WHERE codusu = " & vUsu.Codigo
        
        'Si agrupa por proveedor
        If chkDatosAlbaranes(3).Value = 1 Then
            'y detalla el articulo
            If Me.chkDetallaArticulo.Value = 1 Then
                'TIENE QUE DETALLAR el proveedor
                If chkDatosAlbaranes(4).Value = 0 Then
                    MsgBox "No puede detallar articulo y no detallar proveedor", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        'Si no agrupa por proveedor NO tiene senditod
        If chkDatosAlbaranes(4).Value = 1 Then
            If chkDatosAlbaranes(3).Value = 0 Then
                MsgBox "Detallar proveedor solo disponible para 'agrupa proveedor'", vbExclamation
                Exit Sub
            End If
        End If
        
        'Comparativo. Debe inidcar fechas
        
        If chkDatosAlbaranes(10).Value = 1 Then
            If txtCodigo(98).Text = "" Or txtCodigo(99).Text = "" Then
                MsgBox "Debe indicar fechas para el listado comparativo", vbExclamation
                Exit Sub
            End If
            
            'NO lleva CLIENTE. NO clasifica por cliente
            If Me.OptPorCliente.Value Then OptPorFamilia.Value = True
                
            
        End If
        
        
    End If
    
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------

    tabla = ""
    If txtCodigo(96).Text <> "" Or txtCodigo(97).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "Cli.: "
        If Not PonerDesdeHasta(campo, "N", 96, 97, Cad) Then Exit Sub
        tabla = Cad
    End If
   
    If txtCodigo(127).Text <> "" Or txtCodigo(128).Text <> "" Then
        campo = "{sclien.codactiv} "
        'Parametro Desde/Hasta Cliente
        Cad = "Act.: "
        If Not PonerDesdeHasta(campo, "N", 127, 128, Cad) Then Exit Sub
        If tabla <> "" Then
            tabla = tabla & "   Activ: "
            If txtCodigo(127).Text = txtCodigo(128).Text Then
                tabla = tabla & txtCodigo(127).Text & " - " & Me.txtNombre(127).Text
            Else
                tabla = tabla & "[" & txtCodigo(127).Text & ".." & txtCodigo(128).Text & "]"
            End If
        Else
            tabla = Cad
        End If
    End If
    
    
    If OpcionListado = 231 Then
        Cad = ""
        indCodigo = 0
        campo = ""
        For NumRegElim = 1 To Me.lwFact.ListItems.Count
            If Me.lwFact.ListItems(NumRegElim).Checked Then
                indCodigo = indCodigo + 1
                Cad = Cad & "- " & lwFact.ListItems(NumRegElim).Text
                campo = campo & ", '" & lwFact.ListItems(NumRegElim).Text & "'"
            End If
        Next
        
        If Cad = "" Then
            MsgBox "Seleccione algun tipo de factura", vbExclamation
            Exit Sub
        End If
        
        'No ha seleccionado todos
        If indCodigo <> Me.lwFact.ListItems.Count Then
            campo = Mid(campo, 2)
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            If cadFormula <> "" Then cadFormula = cadFormula & " AND "
            cadSelect = cadSelect & " scafac.codtipom IN (" & campo & ")"
            cadFormula = cadFormula & " {scafac.codtipom} IN [" & campo & "]"
            
            'Para el encabezado de pagina
            Cad = Mid(Cad, 2)
            tabla = Trim(tabla & "   Tipo: " & Cad)
            
        End If
        
        
        
        
        'ZONA - ASociacion
        If txtCodigo(157).Text <> "" Or txtCodigo(158).Text <> "" Then
            campo = "{sclien.codrutas} "
            'Parametro Desde/Hasta Cliente
            Cad = IIf(vParamAplic.NumeroInstalacion = 2, "ASOC: ", "Zona:")
            If Not PonerDesdeHasta(campo, "N", 157, 158, Cad) Then Exit Sub
            If tabla <> "" Then
                tabla = tabla & "   " & IIf(vParamAplic.NumeroInstalacion = 2, "ASOC: ", "Zona:")
                If txtCodigo(157).Text = txtCodigo(158).Text Then
                    tabla = tabla & txtCodigo(157).Text & " - " & Me.txtNombre(157).Text
                Else
                    tabla = tabla & "[" & txtCodigo(157).Text & " hasta " & txtCodigo(158).Text & "]"
                End If
            Else
                tabla = Cad
            End If
        End If

        
        
        
            
        
    End If
    
    Cad = "pDHClien= """ & tabla & """|"
    cadParam = cadParam & Cad
    numParam = numParam + 1

   
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    'MOdificacion  18 Novi 2008
    'Las estadisticas son sobre facturas.... Y ALBARANES!!!!
    'La fecha no se la puedo pasar porque en el union hacer referencia a dos campos
    'fecfactu(factura) y fechaalb (albaranes)
    'para ello hay un parametro en el informe
  
    If txtCodigo(98).Text <> "" Or txtCodigo(99).Text <> "" Then
        
        
        Cad = IIf(Me.chkDatosAlbaranes(0).Value = 1, "alb.", "fact.")
        campo = ""
        If Me.chkDatosAlbaranes(10).Value = 1 Then campo = "[Ver periodo anterior]"
        Cad = "pDHFecha=""" & campo & " Fecha " & Cad & " :"
        
        
        If Me.chkDatosAlbaranes(0).Value = 1 And Me.OptPorFamilia.Value = False Then
            campo = "{scafac1.fechaalb}"
        Else
            campo = "{scafac.fecfactu}"
        End If
        
        
        If Not PonerDesdeHasta(campo, "F", 98, 99, Cad) Then Exit Sub
    End If
    
    tabla = "scafac"

    If OpcionListado = 230 Then
        campo = ""  'Para comprobar que alguno de los campos es distinto de ""
        
        
        '---------------   VENTAS x FAMILIA / ARTICULO
        If Me.chkDetallaArticulo.Value = 1 Then
            If txtCodigo(112).Text <> "" Or txtCodigo(112).Text <> "" Then
                campo = "{slifac.codArtic}"
                Cad = "pDHFamilia=""Art�culo: "
                If Not PonerDesdeHasta(campo, "T", 112, 113, Cad) Then Exit Sub
            End If
        End If
    
    
        'Pondremos en el head del report familia y proveedor juntos
        nomRPT = ""
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtCodigo(100).Text <> "" Or txtCodigo(101).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            Cad = "Fam.: "
            If Not PonerDesdeHasta(campo, "N", 100, 101, Cad) Then Exit Sub
            nomRPT = Cad
        End If
        
        'Cadena para seleccion D/H proveedor
        '--------------------------------------------
         If txtCodigo(124).Text <> "" Or txtCodigo(125).Text <> "" Then
            campo = "{sartic.codprove}"
            'Parametro Desde/Hasta Familia
            Cad = "Prov.: "
            If Not PonerDesdeHasta(campo, "N", 124, 125, Cad) Then Exit Sub
            
            If nomRPT <> "" Then nomRPT = nomRPT & """ + chr(13) + """
            nomRPT = nomRPT & Cad
        End If
        
        
        Cad = "pDHFamilia= """ & nomRPT & """|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        nomRPT = ""
        
        'Si por algun campo de los de arriba es <>"" entonces tenemos que meter esto
        If campo <> "" Then
            If Me.chkDatosAlbaranes(0).Value = 0 Then
                'Sin albaranes
                HePuestoElJoinConSclien = True
                tabla = "(( scafac INNER JOIN slifac ON scafac.codtipom=slifac.codtipom AND scafac.numfactu=slifac.numfactu "
                tabla = tabla & " AND scafac.fecfactu=slifac.fecfactu )"
                tabla = tabla & " INNER JOIN sartic ON slifac.codartic=sartic.codartic) INNER JOIN sclien ON sclien.codclien=scafac.codclien "
            End If
        End If
    End If
    
    
    
    If Me.chkDatosAlbaranes(10).Value = 1 Then
        Cadselect2 = cadSelect
        F2 = CDate(txtCodigo(98).Text)
        Cad = Format(F2, "yyyy-mm-dd")
        F2 = DateAdd("yyyy", -1, F2)
        Cadselect2 = Replace(Cadselect2, Cad, Format(F2, "yyyy-mm-dd"))
        F2 = CDate(txtCodigo(99).Text)
        Cad = Format(F2, "yyyy-mm-dd")
        F2 = DateAdd("yyyy", -1, F2)
        Cadselect2 = Replace(Cadselect2, Cad, Format(F2, "yyyy-mm-dd"))
    End If
        
    
    
    If Me.chkDatosAlbaranes(0).Value = 0 Then
        If Not HePuestoElJoinConSclien Then
            tabla = "scafac,sclien"
            If cadSelect <> "" Then cadSelect = cadSelect & " AND scafac.codclien = sclien.codclien"
        End If
        
        
        B = False
         Screen.MousePointer = vbHourglass
        If Not HayRegParaInforme(tabla, cadSelect) Then B = True
        Screen.MousePointer = vbDefault
        If B Then Exit Sub
    Else
        
    
        'Es un contador de un UNION
        Screen.MousePointer = vbHourglass
        If Not ContadorDelUnion(True) Then
             Screen.MousePointer = vbDefault
            If Cadselect2 = "" Then
                MsgBox "No existen valores con esos parametros", vbExclamation
                Exit Sub
            Else
                B = False
                Screen.MousePointer = vbHourglass
                If Not HayRegParaInforme(tabla, Cadselect2) Then B = True
                Screen.MousePointer = vbDefault
                If B Then Exit Sub
            End If
        End If
    End If
    
    HePuestoElJoinConSclien = False
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    If OpcionListado = 230 Then
        
        
        
        
        If Me.OptPorCliente.Value = True Then 'Agrupar por cliente y familia
                
            If Not Me.chkDatosAlbaranes(9).visible Then Me.chkDatosAlbaranes(9).Value = 0
                
            'If Me.chkDatosAlbaranes(0).Value = 0 Or Me.chkDetallaArticulo.Value = 0 Then
            
            If Me.chkDatosAlbaranes(0).Value = 0 Then Me.chkDatosAlbaranes(9).Value = 0
            
            If Me.chkDetallaArticulo.Value = 0 Then
            
                nomRPT = "rFacEstCliFam"
                Titulo = "Listado Ventas por cliente/familia"
                conSubRPT = True
            Else
                nomRPT = ""
                If Me.chkDatosAlbaranes(9).visible Then
                    If Me.chkDatosAlbaranes(9).Value = 1 Then nomRPT = "gr"
                End If
            
                If nomRPT <> "" Then
                    
                    nomRPT = "rFacEstCliFamArtAgru"
                    Titulo = "Listado ventas por art�culo con albaranes"
                Else
                    nomRPT = "rFacEstCliFamArt"
                    Titulo = "Listado ventas por cliente/familia/art�culo"
                End If
                conSubRPT = False
            End If
            
            
            If Me.chkDatosAlbaranes(0).Value = 1 Then
                 nomRPT = nomRPT & "Alb"
                Titulo = Titulo & "(Con albaranes)"
                DesdeTmp = True
                
            End If
        ' ---- [09/11/2009] [LAURA] : agrupar por cliente/familia o solo por familia
        '                             en ambos casos puede detallar articulo
        ElseIf Me.OptPorFamilia.Value = True Then 'agrupar solo por familia
            If Me.chkDetallaArticulo.Value = 0 Then
                nomRPT = "rFacEstFam"
                Titulo = "Listado Ventas por familia"
                conSubRPT = True
                Me.chkDatosAlbaranes(9).Value = 0
            Else
                nomRPT = "rFacEstFamArt"
                Titulo = "Listado ventas por familia/art�culo"
                conSubRPT = False
            End If
            
        End If
        
        
        
        
        If Me.chkDatosAlbaranes(3).Value = 1 And chkDatosAlbaranes(10).Value = 0 Then
            nomRPT = nomRPT & "Pro"
            If Me.chkDatosAlbaranes(12).visible And chkDatosAlbaranes(12).Value = 1 Then nomRPT = nomRPT & "N"
            cadParam = cadParam & "Detalle= " & Abs(chkDatosAlbaranes(4).Value) & "|"
            numParam = numParam + 1
            
        End If
        
            
        nomRPT = nomRPT & ".rpt"
        
        If Me.chkDatosAlbaranes(10).Value = 1 Then
            DesdeTmp = True
            
            
            nomRPT = "rFacEstCompara.rpt"
            
            
            
            cadParam = cadParam & "Detalle= " & Abs(chkDatosAlbaranes(3).Value = 0) & "|"
            numParam = numParam + 1
        
                
            Titulo = "Listado Ventas por cliente/familia comparativo"
            conSubRPT = True
        End If
        
        If DesdeTmp Then
               
                
                'Abril 2011
                '---------------------------
                
                Screen.MousePointer = vbHourglass
                B = InsertarTmpEstdisticasVtas(Cadselect2)
                DoEvents
                Screen.MousePointer = vbDefault
                If Not B Then Exit Sub
                
                
                'En la cadena seleccion cambiamos las tabla por
                cadFormula = Replace(cadFormula, "scafac1", "tmpcommandest")
                cadFormula = Replace(cadFormula, "scafac", "tmpcommandest")
                cadFormula = Replace(cadFormula, "sartic", "tmpcommandest")
                cadFormula = Replace(cadFormula, "slifac", "tmpcommandest")
                'ahora
                cadFormula = "{tmpcommandest.codusu} = " & vUsu.Codigo
        
        End If
        
        
            
        
        
    Else
    
        'EL QUE HABIA
        If Me.optDetalleFacturacion(0).Value Then
            nomRPT = "rFacEstCliImp.rpt"
            Titulo = "Detalle Facturaci�n Clientes"
            conSubRPT = False
            
            
            
            
            
            If Me.chkDatosAlbaranes(8).Value Then
            
                'Cargamos por tipo datos por tipo IVA
                conn.Execute "DELETE FROM tmpcommand WHERE codusu = " & vUsu.Codigo
                If cadSelect = "" Then cadSelect = " sclien.codclien=scafac.codclien "
                'Cargamos IVAS
                'IVA 1 sin re
                Cad = "INSERT INTO tmpcommand (codusu,rap1,rap2,cantidad,importel)"
                Cad = Cad & " select " & vUsu.Codigo & ", porciva1,porciva1re, sum(baseimp1),sum(imporiv1)+sum(if(imporiv1re is null,0,imporiv1re))"
                Cad = Cad & " from scafac,sclien WHERE "
                Cad = Cad & cadSelect & " and not porciva1 is null group by porciva1,porciva1re"
                conn.Execute Cad
                
                
                Cad = "INSERT INTO tmpcommand (codusu,rap1,rap2,cantidad,importel)"
                Cad = Cad & " select " & vUsu.Codigo & ", porciva2,porciva2re, sum(baseimp2),sum(imporiv2)+sum(if(imporiv2re is null,0,imporiv2re))"
                Cad = Cad & " from scafac,sclien WHERE "
                Cad = Cad & cadSelect & " and not porciva2 is null group by porciva2,porciva2re"
                conn.Execute Cad
                
                Cad = "INSERT INTO tmpcommand (codusu,rap1,rap2,cantidad,importel)"
                Cad = Cad & " select " & vUsu.Codigo & ", porciva3,porciva3re, sum(baseimp3),sum(imporiv3)+sum(if(imporiv3re is null,0,imporiv3re))"
                Cad = Cad & " from scafac,sclien WHERE "
                Cad = Cad & cadSelect & " and not porciva3 is null group by porciva3,porciva3re"
                conn.Execute Cad
            
            
            
            
                nomRPT = "rFacEstCliImpFP.rpt"
                Titulo = "Detalle Facturaci�n por forma pago"
                conSubRPT = True
            
                Cad = "codusu= " & vUsu.Codigo & "|"
                cadParam = cadParam & Cad
                numParam = numParam + 1
            End If
        Else
            nomRPT = "rFacDetalleFacTipom.rpt"
            Titulo = "Detalle Facturaci�n x tipo factura"
            conSubRPT = False
            
            Cad = "VerFP= " & Abs(Me.chkDatosAlbaranes(8).Value) & "|"
            cadParam = cadParam & Cad
            numParam = numParam + 1
            
        End If
    End If
    
    
    LlamarImprimir False, False, , Me.chkDatosAlbaranes(10).Value = 1
    
End Sub

Private Function ContadorDelUnion(Compras As Boolean) As Boolean
Dim C As String

    'Con albaranes
    Codigo = cadSelect
    Codigo = QuitarCaracterACadena(Codigo, "{")
    Codigo = QuitarCaracterACadena(Codigo, "}")
    
    
    ContadorDelUnion = False
    If Compras Then
            C = "(SELECT count(*) FROM   ((((`scafac1` `scafac1` INNER JOIN `scafac` `scafac` ON"
            C = C & " ((`scafac1`.`codtipom`=`scafac`.`codtipom`) AND (`scafac1`.`numfactu`=`scafac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`scafac`.`fecfactu`)) INNER JOIN `slifac` `slifac` ON"
            C = C & " ((((`scafac1`.`codtipom`=`slifac`.`codtipom`) AND (`scafac1`.`numfactu`=`slifac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`slifac`.`fecfactu`)) AND (`scafac1`.`numalbar`=`slifac`.`numalbar`))"
            C = C & " AND (`scafac1`.`codtipoa`=`slifac`.`codtipoa`)) INNER JOIN `sartic` `sartic`"
            C = C & " ON `slifac`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            
            C = C & ")  INNER JOIN `sclien` ON `sclien`.`codclien`=`scafac`.`codclien`"
            
            If Codigo <> "" Then C = C & " WHERE " & Codigo
            C = C & ") + ("
            C = C & " SELECT count(*) from (((`slialb` INNER JOIN scaalb ON ((`slialb`.`codtipom`=`scaalb`.`codtipom`) AND"
            C = C & " (`slialb`.`numalbar`=`scaalb`.`numalbar`)))"
            C = C & " INNER JOIN `sartic` `sartic` ON `slialb`.`codartic`=`sartic`.`codartic`)"
            C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
            C = C & " INNER JOIN `sclien` ON `scaalb`.`codclien`=`sclien`.`codclien`"
            If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafac1", "scaalb")
                Codigo = Replace(Codigo, "scafac", "scaalb")
                Codigo = Replace(Codigo, "slifac", "slialb")
                
                C = C & " WHERE " & Codigo
            End If
            C = C & ")"
    
    Else
    
        'Ventas
        C = "(SELECT Count(*) from (`scafpc` `scafpc` INNER JOIN `scafpa` `scafpa`"
        C = C & " ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`=`scafpa`.`fecfactu`))"
        C = C & " AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
        C = C & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
        C = C & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
        C = C & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"
        If Codigo <> "" Then C = C & " WHERE " & Codigo
        C = C & ") + ("

        C = C & " SELECT count(*)"
        C = C & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
        C = C & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
        If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafpa", "scaalp")
                Codigo = Replace(Codigo, "scafpc", "scaalp")
                Codigo = Replace(Codigo, "slifac", "slialp")
                'Para albaranes, cambio fechafactura por fechaalb
                Codigo = Replace(Codigo, "fecfactu", "fechaalb")
                C = C & " WHERE " & Codigo
        End If
        C = C & ")"
    End If
    
    
    C = "Select " & C & " AS total"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then ContadorDelUnion = True
    End If
    miRsAux.Close
    Codigo = ""
End Function


Private Sub cmdAceptarEtiqProv_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String

    InicializarVbles
    
    'si es listado de CARTAS/eMAIL a proveedores comprobar que se ha seleccionado
    'una carta para imprimir
    If OpcionListado = 306 Then
        If txtCodigo(63).Text = "" Then
            MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
            Exit Sub
        End If
        
        
        If Not PonerParamRPT2(62, cadParam, numParam, nomRPT, vImprimedirecto, campo, pRptvMultiInforme) Then
            nomRPT = "rComProveCarta.rpt"
        End If
        
        
        'Parametro cod. carta
        cadParam = "|pCodCarta= " & txtCodigo(63).Text & "|"
        numParam = numParam + 1
        
        
        'Firmado
        cadParam = cadParam & "pFirmado=""" & Trim(txtCodigo(146).Text) & """|"
        numParam = numParam + 1
        
        
        'Nombre fichero .rpt a Imprimir
        'nomRPT = "rComProveCarta.rpt"
        Titulo = "Cartas a Proveedores"
        conSubRPT = True
        
    Else 'ETIQUETAS
        cadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveEtiq.rpt"
        Titulo = "Etiquetas de Proveedores"
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtCodigo(58).Text <> "" Or txtCodigo(59).Text <> "" Then
        campo = "{sprove.codprove}"
        'Parametro Desde/Hasta Proveedor
        If Not PonerDesdeHasta(campo, "N", 58, 59, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H COD. POSTAL
    '--------------------------------------------
     If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
        campo = "{sprove.codpobla}"
        'Parametro Desde/Hasta cod. Postal
        If Not PonerDesdeHasta(campo, "T", 60, 61, "") Then Exit Sub
    End If
    
    '====================================================
        
        
    'Parametro a la Atencion de
    cadParam = cadParam & "pAtencion=""Att. " & txtCodigo(62).Text & """|"
    numParam = numParam + 1
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme("sprove", cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = cadSelect
    frmMen.OpcionMensaje = 9 'Etiquetas proveedores
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    If OpcionListado = 306 And Me.chkEMAIL(0).Value = 1 Then
        'Enviarlo por e-mail
        EnviarEMailMulti cadSelect, Titulo, "rComProveCarta.rpt", "sprove" 'email para proveedores
    Else
        LlamarImprimir False, False
    End If
    
End Sub


Private Sub cmdAceptarFacRect_Click()
Dim Cad As String
Dim TipoM As String * 3
Dim Cli As Long
Dim CF As CFactura
Dim CCl As CCliente
    'Comprobar que se introdujo el motivo por el que se rectifica la factura
    If Trim(txtCodigo(87).Text) = "" Then
        MsgBox "Debe introducir el motivo de rectificaci�n.", vbExclamation
        PonerFoco txtCodigo(87)
        Exit Sub
    End If


    TipoM = Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3)
    
    'comprobar que existe la factura en tabla "scafac"
    Cad = "select count(*) from scafac where codtipom='" & TipoM & "' AND numfactu="
    If txtCodigo(71).Text <> "" And txtCodigo(72).Text <> "" Then
        Cad = Cad & txtCodigo(71).Text & " AND fecfactu=" & DBSet(txtCodigo(72).Text, "F")
    Else
         Cad = Cad & "-1"  'para que no de error
    End If
    
    
    If RegistrosAListar(Cad) = 0 Then
        
        Cad = vbCrLf & String(40, "*") & vbCrLf
        Cad = Cad & vbCrLf & "No existe la factura que quiere rectificar" & vbCrLf
        If Me.chkFraRectCompleta.Value = 1 Then
            MsgBox Cad, vbExclamation
            Exit Sub
        Else
            Cad = Cad & "�Continuar?" & Cad
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    
    Else
        If Me.chkFraRectCompleta.Value = 1 Then
            Cad = ""
            NumeroDeCopias = 160
            If Me.txtCodigo(160).Text = "" Then
                Cad = Cad & "   -Fecha factura"
            Else
                If Not EsFechaOK(CStr(txtCodigo(160).Text)) Then
                    Cad = Cad & "   -Fecha factura incorrecta"
                Else
                    cadParam = EsFechaOKConta(CDate(txtCodigo(160).Text), True)
                    If cadParam <> "0" Then
                        Cad = Cad & "   - " & MensajeFechaOkConta
                    Else
                        cadParam = FechaLiquidacionIVA
                        If cadParam <> "" Then Cad = Cad & "   - " & cadParam
                    End If
                End If
            End If
            If Cad = "" Then
                If vParamAplic.NumeroInstalacion = vbAlzira Then
                    If Me.txtCodigo(160).Text <> Format(Now, "dd/mm/yyyy") Then
                        If vUsu.Nivel > 0 Then Cad = "Fecha debe ser fecha actual"
                    End If
                End If
            End If
            If Me.txtCodigo(159).Text = "" Or Me.txtNombre(159).Text = "" Then Cad = Cad & vbCrLf & "   -Banco": NumeroDeCopias = 159
            If Me.txtCodigo(161).Text = "" Or Me.txtNombre(161).Text = "" Then Cad = Cad & vbCrLf & "   -Forma de pago": NumeroDeCopias = 161
            
            cadSelect = PonerTrabajadorConectado(Codigo)
            If cadSelect = "" Then Cad = Cad & vbCrLf & "   -Error trabajador conectado"
                        
                        
                        
                        
            If Cad <> "" Then
                Cad = "Campos incorrectos: " & vbCrLf & vbCrLf & Cad
                MsgBox Cad, vbExclamation
                PonerFoco txtCodigo(NumeroDeCopias)
                NumeroDeCopias = 1
                Exit Sub
            End If
            NumeroDeCopias = 1
            
            
            'Comprobaremos que no hay ningun articulo con fecha de inventario posterior a fechaalb
            Set miRsAux = New ADODB.Recordset
            Cad = "select  slifac.codalmac,slifac.codartic,nomartic,fechainv"
            Cad = Cad & " from slifac left join salmac on slifac.codartic=salmac.codartic  and  slifac.codalmac=salmac.codalmac where fechainv>= " & DBSet(txtCodigo(160).Text, "F")
            Cad = Cad & " and codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
            Cad = Cad & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            While Not miRsAux.EOF
                Cad = Cad & Format(miRsAux!FechaINV, "dd/mm/yyyy") & "  " & miRsAux!NomArtic & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            If Cad <> "" Then
                Cad = "Articulos con inventario posterior a la fecha de albaran/factura rectificativo:" & vbCrLf & vbCrLf & Cad
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    'Julio 2015
    'Comprobamos que esa factura NO ha sido rectifcada anteriormente
    Cad = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoM, "T")
    If Cad = "" Then Cad = TipoM
    Cad = "'RECTIFICA A FACTURA: " & Cad & ", " & Val(txtCodigo(71).Text) & ", " & txtCodigo(72).Text & "'"
    Codigo = "observa2 like 'Moti%' AND observa1 = " & Cad & " AND codtipom"
    Cad = DevuelveDesdeBD(conAri, "concat(numfactu,'|',fecfactu , '|')", "scafac1", Codigo, "FRT", "T")
    
    If Cad <> "" Then
        Codigo = vbCrLf & "Factura: " & Format(RecuperaValor(Cad, 1), "00000") & " del " & Format(RecuperaValor(Cad, 2), "dd/mm/yyyy") & vbCrLf
        Codigo = "La factura ya ha sido rectificada. " & Codigo
        Cad = vbCrLf & String(40, "*") & vbCrLf
        Cad = Cad & vbCrLf & Codigo & vbCrLf & "�Continuar?" & Cad
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    'Llegado aqui pongo los datos
    'si existe devolver estos datos para recuperla en el formulario de Albaranes
    If chkFraRectCompleta.Value = 1 Then
    
        Cad = "concat(concat(codtipom,right(concat('0000000',numfactu),7)),'|',concat(codclien,' ',nomclien),'|',brutofac,'|',"
        Cad = Cad & "imporiv1 + coalesce(imporiv2,0) + coalesce(imporiv3,0) ,'|',totalfac,'|')"
        Codigo = "numfactu =" & txtCodigo(71).Text & " AND fecfactu=" & DBSet(txtCodigo(72).Text, "F") & " AND codtipom "
        Cad = DevuelveDesdeBD(conAri, Cad, "scafac", Codigo, TipoM, "T")
        If Cad = "" Then
            MsgBox "NO existe factura a rectificar", vbExclamation
    
        Else
            
            Codigo = "Factura rectificada: " & RecuperaValor(Cad, 1)
            Codigo = Codigo & vbCrLf & "Cliente: " & RecuperaValor(Cad, 2)
            Codigo = Codigo & vbCrLf & "bruto: " & RecuperaValor(Cad, 3)
            Codigo = Codigo & vbCrLf & "IVA: " & RecuperaValor(Cad, 4)
            Codigo = Codigo & vbCrLf & "Total: " & RecuperaValor(Cad, 5) & vbCrLf & vbCrLf
            Codigo = Codigo & "FECHA factura rectificativa: " & txtCodigo(160).Text
            Codigo = Codigo & vbCrLf & "Generar la factura rectificativa?"
            If MsgBox(Codigo, vbQuestion + vbYesNoCancel) = vbYes Then
                Screen.MousePointer = vbHourglass
                
                Cad = RecuperaValor(Cad, 2)
                Cad = Mid(Cad, 1, InStr(1, Cad, " "))
                Cli = Val(Trim(Cad))
                conn.BeginTrans
                If GeneraFacturaRectificativaCompleta Then
                     conn.CommitTrans
                     Espera 0.5
                     Label4(133).Caption = "Calculando vtos."
                     Label4(133).Refresh
                     
                     'Para no tener que reclcular ventimientos y demas, esto lo hace autmaticamente
                     Set CF = New CFactura
                     If CF.LeerDatos("FRT", NumCod, txtCodigo(160).Text) Then
                        Set miRsAux = New ADODB.Recordset
                        Cad = "Select nomclien,domclien,codpobla,pobclien,proclien,nifclien  "
                        Cad = Cad & " from scafac WHERE numfactu =" & txtCodigo(71).Text & " AND fecfactu=" & DBSet(txtCodigo(72).Text, "F") & " AND codtipom =" & DBSet(TipoM, "T")
                        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Not miRsAux.EOF Then
                            CF.NombreClien = DBLet(miRsAux!NomClien, "T")
                            CF.DomicilioClien = DBLet(miRsAux!domclien, "T")
                            CF.CPostal = DBLet(miRsAux!codpobla, "T")
                            CF.Poblacion = DBLet(miRsAux!pobclien, "T")
                            CF.Provincia = DBLet(miRsAux!proclien, "T")
                            CF.NIF = DBLet(miRsAux!nifClien, "T")
                        End If
                        miRsAux.Close
                        Set miRsAux = Nothing
                        If CF.CuentaPrev = "" Then CF.CuentaPrev = DevuelveDesdeBD(conAri, "codmacta", "sbanpr", "codbanpr", txtCodigo(159).Text)
                        Cad = ""
                        CF.InsertarEnTesoreria "", Cad, True
                        If Cad <> "" Then MsgBox Cad, vbExclamation
                     End If
                     Espera 0.2
                     Set CF = Nothing
                     
                     cadParam = ""
                     numParam = 0
                     Cad = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoM, "T")
                     With frmImprimir
                            'Nuevo. Febrero 2010
                            .outClaveNombreArchiv = CStr(Cad & NumCod)
                            .outCodigoCliProv = Cli
                            .outTipoDocumento = 2
                            
                            .SeleccionaRPTCodigo = pRptvMultiInforme
                            .FormulaSeleccion = cadFormula
                            .OtrosParametros = cadParam
                            .NumeroParametros = 3
                            .NombrePDF = pPdfRpt
                            .SoloImprimir = False
                            .EnvioEMail = False
                            .NumeroCopias = vParamAplic.NumCop_FraRectifica
                            .Opcion = 53
                            .NombreRPT = nomRPT
                            .Titulo = "Rectificativa"
                            .Show vbModal
                    End With
                    NumCod = ""
                    
                    
                    
                    
                Else
                    conn.RollbackTrans
                    Cad = ""
                End If
                Screen.MousePointer = vbDefault
                If Cad <> "" Then
                    Cad = "OK_FRA"
                    RaiseEvent DatoSeleccionado(Cad)
                    Unload Me
                End If
            End If
        End If
    Else
        'Lo que hacia antes
        Cad = TipoM & "|"
        If txtCodigo(71).Text <> "" Then
            Cad = Cad & txtCodigo(71).Text & "|"
        Else
            Cad = Cad & "-1|"   'k no de error el select
        End If
        Cad = Cad & txtCodigo(72).Text & "|"
        Cad = Cad & QuitarCaracterEnter(txtCodigo(87).Text) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub cmdAceptarGenPed_Click()
'Solicitar datos para Generar Pedido a partir de una Oferta
Dim Cad As String
    
    If txtCodigo(24).Text = "" Or txtCodigo(25).Text = "" Or txtCodigo(26).Text = "" Or txtNombre(4).Text = "" Or txtNombre(24).Text = "" Then
    
        MsgBox "Todos los campos son obligatorios", vbExclamation
        Exit Sub
    End If
    
    
    Cad = txtCodigo(24).Text & "|"
    Cad = Cad & txtCodigo(25).Text & "|"
    Cad = Cad & txtCodigo(26).Text & "|"
    Cad = Cad & txtNombre(4).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub cmdAceptarHco_Click()
'pedir datos para Pasar de Albaranes a historico
Dim Cad As String
    Cad = ""
    'comprobar que todos los camos tienen valor
    If txtCodigo(50).Text = "" Or txtCodigo(51).Text = "" Or txtCodigo(52).Text = "" Then
        Cad = "(I)"
    Else
        If txtNombre(51).Text = "" Or txtNombre(52).Text = "" Then Cad = "(II)"
    End If

    If Cad <> "" Then
        Cad = "Debe rellenar todos los campos para pasar al hist�rico. " & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    'datos a devolver
    Cad = txtCodigo(50).Text & "|"
    Cad = Cad & txtCodigo(51).Text & "|"
    Cad = Cad & txtCodigo(52).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub cmdAceptarOfer_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim jj As Integer


    If txtCodigo(1).Text = "" Then 'And (txtCodigo(33).Text = "" Or txtCodigo(34).Text = "") Then
        MsgBox "Debe seleccionar una Oferta para Imprimir.", vbInformation
        PonerFoco txtCodigo(1)
        Exit Sub
    End If
    
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    NumeroDeCopias = 1
    If (OpcionListado = 31) Then
        If Me.OptPapelBlanco.Value Then
            'las ofertas NOrmales, las que van a cliente
            indRPT = 5 '31: Informe de Ofertas
            NumeroDeCopias = vParamAplic.NumCop_Oferta
            
            If Me.chkVarios(12).Value = 1 Then indRPT = 95
            
        Else
            indRPT = 54    'octubre 2011
            
            If Me.chkVarios(12).visible Then
                If Me.chkVarios(12).Value = 1 Then MsgBox "No existe el formato oferta interno con IVA", vbExclamation
            End If
        
            
            
        End If
    
    
    
        
    
    
    
    
    ElseIf OpcionListado = 35 Then
        'En HISTORICO NO hay ofertas internas... de momento
        devuelve = ""
        If OptPapelMembrete.Value Then devuelve = "No existe el formato oferta interno para el historico"
        
        If Me.chkVarios(12).visible Then
            If Me.chkVarios(12).Value = 1 Then devuelve = devuelve & vbCrLf & "No existe el formato oferta interno con IVA"
        End If
        If devuelve <> "" Then
            MsgBox devuelve, vbExclamation
            Exit Sub
        End If
        indRPT = 6 '35: Historico Informe de Ofertas
    End If
    conSubRPT = True
    
    
    
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then Exit Sub


    'ANTES ocubre 2011
    'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
''''''''    If Me.OptPapelBlanco.Value = True Then 'Blanco o con Membrete
''''''''        devuelve = "True"
''''''''    Else
''''''''        devuelve = "False"
''''''''    End If
''''''''    cadParam = cadParam & "pPapelB=" & devuelve & "|"
''''''''    numParam = numParam + 1
                
    
    
            
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then Exit Sub
                
    
    'Se pasa como parametro la carta a imprimir
    If Me.txtCodigo(2).Text <> "" Then
        cadParam = cadParam & "pCodCarta=" & CInt(Me.txtCodigo(2).Text) & "|"
    Else
        cadParam = cadParam & "pCodCarta=" & CInt(0) & "|"
    End If
    numParam = numParam + 1
    
    'A�adir el codigo de usuario como parametro para link con tabla Temporal en el Report
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de OFERTA
    '--------------------------------------------
    CadenaParaEnvioMail = ""
    If txtCodigo(1).Text <> "" Then
       
        
        'Si Imprimir Otras Ofertas del Cliente
        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
            campo = "{" & NomTabla & ".fecofert}"
            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
            If devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & devuelve & ")"
                cadSelect = "(" & cadSelect & " OR " & CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F") & ")"
            Else
                cadFormula = devuelve
                cadSelect = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
            End If
''''            devuelve = "{" & NomTabla & ".aceptado}=0"
''''            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
''''            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
''''
        Else
            devuelve = "{" & NomTabla & ".numofert}=" & Val(txtCodigo(1).Text)
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            cadSelect = cadFormula
            If OpcionListado = 35 Then 'solo imprime la Oferta Seleccionada (si Historico filtrar x fecha)
                devuelve = "{" & NomTabla & ".fecofert}=Date(" & Year(FecEntre) & ", " & Month(FecEntre) & ", " & Day(FecEntre) & ")"
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                devuelve = NomTabla & ".fecofert= '" & Format(FecEntre, FormatoFecha) & "'"
                AnyadirAFormula cadSelect, devuelve
            End If
        End If
        'Filtrar solo las ofertas del cliente que las solicita
        If OpcionListado = 35 Then 'Historico
            devuelve = DevuelveDesdeBDNew(conAri, NomTabla, "codclien", "numofert", txtCodigo(1).Text, "N", , "fecofert", FecEntre, "F")
        Else
            devuelve = DevuelveDesdeBDNew(conAri, NomTabla, "codclien", "numofert", txtCodigo(1).Text, "N")
        End If
        codClien = devuelve
'        If devuelve <> "" Then
'            campo = "{" & NomTabla & ".codclien}=" & devuelve
'            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
'            If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
'        End If
        
        'Para la posibnilidad de enviar por email
        CadenaParaEnvioMail = "1|" & devuelve & "|" & txtCodigo(1).Text & "|"
        

    End If
   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y a�adimos a los parametros el tipo de IVA
    'que se aplica a ese cliente (0=NORMAL)
    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", codClien, "N")
    If devuelve <> "" Then
        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
        numParam = numParam + 1
    End If


    'Agosto 2011
    'Separador
    cadParam = cadParam & "Separador=""" & vParamAplic.ArtSeparador & """|"
    numParam = numParam + 1


    'Cuando este cargada la tabla temporal a�adir un parametro con la concatenacion de
    'Todas las ofertas que se van a imprimir
    devuelve = ""
    If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then devuelve = "ok"
    'FALTA### Daba error
    'PonerParamCadOferta2 cadParam, numParam, cadSelect, cadFormula, devuelve <> "", CLng(txtCodigo(1).Text)
     PonerParamCadOferta2
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    'If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    cadFormula = CadenaDesdeOtroForm




    CadenaParaEnvioMail = "1|" & codClien & "|"
    indCodigo = InStr(1, CadenaDesdeOtroForm, "IN")
    If indCodigo = 0 Then
        CadenaDesdeOtroForm = txtCodigo(1).Text
    Else
        If InStr(1, CadenaDesdeOtroForm, ",") > 0 Then
            CadenaDesdeOtroForm = "RTAS"  'le a�ade el OFE alli
        Else
            CadenaDesdeOtroForm = txtCodigo(1).Text
        End If
    End If
    CadenaParaEnvioMail = CadenaParaEnvioMail & CadenaDesdeOtroForm & "|"
    
    
    If InstalacionEsEulerTaxco Then
        'Si hay elementos "reimpresion seleccionados
        'Cuando pulse imprimir imprimira tb los docuemntos asociados
        'y si dice exoprtar concatenara tb los documentos asociados
        conn.Execute "Delete from tmpImpresionAuxliar WHERE codusu = " & vUsu.Codigo
        
        
        If ListView2.ListItems.Count > 0 Then
            For jj = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(jj).Checked Then
                    devuelve = EulerPathCompletoArchivoOfertas(Val(txtCodigo(1).Text), ListView2.ListItems(jj).SubItems(1))
                    
                    If devuelve <> "" Then
                        devuelve = "(" & vUsu.Codigo & "," & jj & "," & DBSet(devuelve, "T") & ")"
                        devuelve = "INSERT INTO tmpImpresionAuxliar(codusu,orden,fichero) VALUES " & devuelve
                        conn.Execute devuelve
                    End If
                End If
            Next jj
        End If
    End If
    
    LlamarImprimir True, False, CadenaParaEnvioMail
End Sub
Private Sub PonerParamCadOferta2()
Dim C As String
Dim L As Boolean
    Set miRsAux = New ADODB.Recordset
    If Me.txtCodigo(3).Text = "" And txtCodigo(4).Text = "" Then
        
        L = False
    Else
        
        C = "numofert <> " & txtCodigo(1).Text
        C = C & " AND codclien = " & codClien
        If txtCodigo(3).Text <> "" Then C = C & " AND fecofert >='" & Format(txtCodigo(3).Text, FormatoFecha) & "'"
        If txtCodigo(4).Text <> "" Then C = C & " AND fecofert <='" & Format(txtCodigo(4).Text, FormatoFecha) & "'"
        L = True
    End If
    
    CadenaDesdeOtroForm = ""
    If L Then
        C = "Select * from " & NomTabla & " where " & C
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            frmListado2.Opcion = 21
            frmListado2.Show vbModal
        Else
            miRsAux.Close
        End If
        Set miRsAux = Nothing
    End If
    
    CadenaDesdeOtroForm = "{" & NomTabla & ".numofert} IN [" & txtCodigo(1).Text & CadenaDesdeOtroForm & "]"
End Sub

Private Sub cmdAceptarPedCom_Click()
'55: Informe Pedido de Compras (a Proveedor)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim CodPed As String
Dim campo1 As String, campo2 As String, campo3 As String
    
    
    
    
    If txtCodigo(73).Text = "" Then 'N� del Pedido
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        PonerFoco txtCodigo(73)
        Exit Sub
    Else
        NumCod = txtCodigo(73).Text
    End If
    
    If (OpcionListado = 239) And txtCodigo(76).Text = "" Then
        MsgBox "Debe seleccionar un Pedido y Fecha para Imprimir.", vbInformation
        PonerFoco txtCodigo(76)
        Exit Sub
    End If
    
    
    InicializarVbles
    conSubRPT = True
    CadenaParaEnvioMail = ""
    '===================================================
    '============ PARAMETROS ===========================
    Select Case OpcionListado
        Case 38
            Titulo = "Pedido de Ventas"
            indRPT = 7 '7: Pedidos de Clientes
            NumeroDeCopias = vParamAplic.NumCop_Pedido
            If Me.chkVarios(10).Value = 1 Then
                indRPT = 84 'Interno
                Titulo = Titulo & " (INTERNO)"
                NumeroDeCopias = 1
            End If
            
            
            
        Case 239
            indRPT = 8 '8: Pedidos de Clientes (Historico)
            Titulo = "Hist. Pedido de Venta"
        Case 55, 407  'impresion directa(HERBELCA)
            indRPT = 14 '14: Pedidos a Proveedores
            Titulo = "Pedidos de Compras"
        Case 56
            indRPT = 15
            Titulo = "Hist. Pedidos de Compras"
    End Select
    
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then Exit Sub
     
    If OpcionListado = 38 Or OpcionListado = 239 Then
        campo1 = "numpedcl"
        campo2 = "fecpedcl"
        campo3 = "codclien"
    Else
        campo1 = "numpedpr"
        campo2 = "fecpedpr"
        campo3 = "codprove"
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de PEDIDO
    '--------------------------------------------
    If NumCod <> "" Then
        devuelve = "{" & NomTabla & "." & campo1 & "}=" & Val(NumCod)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        If OpcionListado = 239 Then 'historico ( hay fecha)
            devuelve = "{" & NomTabla & "." & campo2 & "}= Date(" & Year(txtCodigo(76).Text) & "," & Month(txtCodigo(76).Text) & "," & Day(txtCodigo(76).Text) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            devuelve = NomTabla & "." & campo2 & "='" & Format(txtCodigo(76).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
        
        'Seleccionar otros PEdidos entre esas FEchas
        If Not (txtCodigo(74).Text = "" And txtCodigo(75).Text = "") Then
            campo = "{" & NomTabla & "." & campo2 & "}"
            devuelve = CadenaDesdeHasta(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            If devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & devuelve & ")"
                cadSelect = "((" & cadSelect & ") OR " & CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F") & ")"
            Else
                cadFormula = devuelve
                cadSelect = CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            End If
        
            'Filtrar solo los Pedidos del CLIENTE/PROVEEDOR que las solicita
            If codClien <> "" Then
                campo = "{" & NomTabla & "." & campo3 & "}=" & codClien
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
     
        End If
        
        'FALTA## para hco pedidos
        If OpcionListado = 38 Or OpcionListado = 239 Then
            CadenaParaEnvioMail = "3|" & codClien & "|" & txtCodigo(73).Text & "|"
        Else
            'Proveedores
            CadenaParaEnvioMail = "51|" & codClien & "|" & txtCodigo(73).Text & "|"
        End If
        
        
        
    Else
'        'Comprobar si se imprimen varios Pedidos
'        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
'         'Cadena para seleccion Desde y Hasta FECHA
'         '--------------------------------------------
'            campo = "{" & NomTabla & ".fecpedcl}"
'            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'            If devuelve = "Error" Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, devuelve) Then
'                Exit Sub
'            Else
'                devuelve = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'                If devuelve = "Error" Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'            End If
'        End If
    End If
    
    If OpcionListado = 38 Or OpcionListado = 239 Then
        'A�adimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", codClien, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        
        'PORTES
        cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
        numParam = numParam + 1
    End If


    'Listado proveedores
    If chkVarios(0).visible = True Then
        cadParam = cadParam & "valorado= " & Abs(Me.chkVarios(0).Value) & "|"
        numParam = numParam + 1
    End If
    'comprobar que hay datos para mostrar en el Informe
     If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    LlamarImprimir True, False, CadenaParaEnvioMail
    Unload Me
End Sub


' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
Private Sub cmdAceptarPedConfirma_Click()
'Confirmacion entrega del pedido
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim NomPDF As String 'Nombre del fichero .pdf
Dim campo As String
Dim RS As ADODB.Recordset

    If txtCodigo(116).Text = "" Then
        MsgBox "Debe seleccionar una carta para Imprimir la Confirmaci�n entrega del Pedido.", vbInformation
        PonerFoco txtCodigo(116)
        Exit Sub
    End If
    
    
    PrepararCarpetasEnvioMail True
    
    InicializarVbles
    
    'Se pasa como parametro la carta a imprimir
    If Me.txtCodigo(116).Text <> "" Then
        cadParam = cadParam & "|pCodCarta=" & CInt(Me.txtCodigo(116).Text) & "|"
    Else
        cadParam = cadParam & "|pCodCarta=" & CInt(0) & "|"
    End If
    numParam = numParam + 1
    
    
    indRPT = 40 'A�ade los parametros de la tabla scrystal para el informe
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, vImprimedirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    NomTabla = "scaped"
    
    'Cadena para seleccion Clientes de Pedido
    '--------------------------------------------
    If txtCodigo(114).Text <> "" Then
        campo = "{scaped.numpedcl}=" & txtCodigo(114).Text
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        cadSelect = cadFormula
    End If
       
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
       
       
    If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
       
'    LlamarImprimir

     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = True
        .Opcion = 238
        .Titulo = "Confirmaci�n entrega Pedido"
        .NombreRPT = nomDocu
        .NombrePDF = pPdfRpt
        .ConSubInforme = True
        .Show vbModal
    End With
    
    
    'FALTA###
'    Exit Sub
    If Dir(App.Path & "\docum.pdf", vbArchive) = "" Then
        MsgBox "No se encuentra el archivo", vbExclamation
        Exit Sub
    End If
    NomPDF = App.Path & "\Temp\PEV-" & Format(NumCod, "0000000") & ".pdf"
    FileCopy App.Path & "\docum.pdf", NomPDF
    
    'Obtener los ficheros que hay en el directorio de documentos
'    MiRuta = "" & App.Path & "" & "\PDF-Docum\"


    '-- obtener los datos para envio e-mail
    campo = "SELECT numpedcl,fecpedcl,codclien,nomclien,mailconfir"
    campo = campo & " FROM " & NomTabla & " WHERE numpedcl=" & NumCod
    Set RS = New ADODB.Recordset
    RS.Open campo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    campo = ""
    If Not RS.EOF Then
        If DBLet(RS!mailconfir, "T") <> "" Then campo = RS!NomClien & "|" & RS!mailconfir & "|"
    End If
    
    RS.Close
    Set RS = Nothing

    If campo = "" Then MsgBox "No hay direcci�n e-mail en el pedido para enviar confirmaci�n de entrega.", vbExclamation
    
    If Dir(NomPDF, vbArchive) <> "" And campo <> "" Then
    
        '- a�adir el subject del e-mail
        campo = campo & "Confirmaci�n entrega pedido " & vEmpresa.nomempre & "|"
        '- a�adir el cuerpo del mensaje
        campo = campo & "Le confirmamos que su pedido adjunto N� " & NumCod & " de fecha " & FecEntre & " le ser� entregado en la semana "
        campo = campo & DevuelveDesdeBDNew(conAri, NomTabla, "sementre", "numpedcl", NumCod, "N") & ".|"
        
        'El adjunto, para que no se llame docum.pdf
        campo = campo & NomPDF & "|"
        
        frmEMail.DatosEnvio = campo
        frmEMail.Opcion = 0 'Envio documento
        frmEMail.Show vbModal
    
        If frmEMail.DatosEnvio = "OK" Then
            campo = "UPDATE " & NomTabla & " SET envconfir=1"
            campo = campo & " WHERE numpedcl=" & NumCod
            conn.Execute campo
        End If
        frmEMail.DatosEnvio = ""
        
    End If
    
    'If Dir(NomPDF, vbArchive) <> "" Then Kill NomPDF
End Sub
' ----



Private Sub cmdAceptarPte_Click()
'LIstado Material Pendiente de recibir
Dim Codigo As String
Dim Cad As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Pasar el ORDEN del informe como parametro
    If OpcionListado = 307 Then
        If Me.OptOrdenArt Then
            Cad = "{slippr.codartic}"
            Codigo = 0
        Else
            Cad = "{scappr.numpedpr}"
            Codigo = 1
        End If
        cadParam = cadParam & "pOrden=" & Cad & "|"
        numParam = numParam + 1
        
        Cad = "pTotaliza=" & Codigo
        cadParam = cadParam & Cad & "|"
        numParam = numParam + 1
    
    
    End If
    
    If OpcionListado = 308 Then
        Cad = "detalla=" & Abs(Me.chkVarios(1).Value)
        cadParam = cadParam & Cad & "|"
        numParam = numParam + 1
        
        Cad = ""
        If Me.cboPdteFacturarProv.ListIndex > 0 Then Cad = Me.cboPdteFacturarProv.Text
        Cad = "pPago=""" & Cad
        cadParam = cadParam & Cad & """|"
        numParam = numParam + 1
        
        
    End If
    
    '===================================================
    '================= FORMULA =========================
    'ser� la cadena WHERE para el Informe
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(65).Text <> "" Or txtCodigo(66).Text <> "" Then
        Codigo = "{scappr.codprove}"
        If OpcionListado = 308 Then Codigo = "{scaalp.codprove}"
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 65, 66, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(69).Text <> "" Or txtCodigo(70).Text <> "" Then
        Codigo = "{scappr.fecpedpr}"
        If OpcionListado = 308 Then Codigo = "{scaalp.fechaalb}"
        Cad = "pDHFecha=""Fecha Ped.: "
        If OpcionListado = 308 Then Cad = "pDHFecha=""Fecha Alb.: "
        If Not PonerDesdeHasta(Codigo, "F", 69, 70, Cad) Then Exit Sub
    End If
    
    If OpcionListado = 307 Then '307: List. Materia pendiente de recibir
        'Cadena para seleccion D/H ARTICULO
        '--------------------------------------------
        If txtCodigo(67).Text <> "" Or txtCodigo(68).Text <> "" Then
            Codigo = "{slippr.codartic}"
            Cad = "pDHArticulo=""Art�culo: "
            If Not PonerDesdeHasta(Codigo, "T", 67, 68, Cad) Then Exit Sub
        End If
    End If
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If OpcionListado = 307 Then
        Cad = "scappr INNER JOIN slippr ON scappr.numpedpr=slippr.numpedpr "
        Titulo = "Material Pendiente de recibir"
        nomRPT = "rComPteRecibir.rpt"
        If Me.chkVarios(11).Value = 1 Then nomRPT = "rComPteRecibirArt.rpt"
        
    Else
    
        If cadSelect <> "" Then
            Codigo = Replace(cadSelect, "{", "")
            Codigo = Replace(Codigo, "}", "")
            Codigo = " AND " & Codigo
        End If
        Cad = "(SELECT @rownum:=0) r, scaalp INNER JOIN slialp ON scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Cad = Cad & Codigo
        conn.Execute "DELETE from tmpinformes where codusu =" & vUsu.Codigo
        
        Codigo = "insert into tmpinformes(codusu,codigo1,campo1,nombre1,fecha1,importe1) select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,scaalp.codprove,scaalp.numalbar, scaalp.fechaalb,sum(slialp.importel) FROM "
        Codigo = Codigo & Cad & "  group by scaalp.codprove,scaalp.numalbar, scaalp.fechaalb"
        If Me.cboPdteFacturarProv.ListIndex > 0 Then Codigo = Codigo & " Having Sum(slialp.ImporteL) " & IIf(Me.cboPdteFacturarProv.ListIndex = 1, ">", "<") & " 0"
        conn.Execute Codigo
        
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
        Titulo = "Pendiente de Factura"
        
        If Me.chkVarios(11).Value = 1 Then
            nomRPT = "rComPteFacturaArticulo.rpt"
        Else
            nomRPT = "rComPteFactura.rpt"
        End If
    End If
    
    If Not HayRegParaInforme(Cad, cadSelect) Then Exit Sub

    'Mostrar el Informe
    conSubRPT = False
    LlamarImprimir False, False
End Sub


Private Sub cmdAceptarReimpFac_Click()
'Reimprimir Facturas ya contabilizadas
Dim TipoM As String * 3
'Dim TipoMh As String * 3
Dim Codigo As String
Dim B As Boolean
Dim TipoFactura As Byte
Dim EsDeTelCabFactura As Boolean




    If txtCodigo(119).Text = "" Then txtCodigo(119).Text = "1"
    
    If Val(txtCodigo(119).Text) > 10 Then
        MsgBox "Numero de copias excesivo", vbExclamation
        Exit Sub
    End If
    
    
    'Agosto 2016
    'Si marca todas, es de UN solo cliente y no van tickets ni telefonia, ni FAG
    If Mid(Me.cboTipomov(1).Text, 1, 3) = "***" Then
        'Ha seleccionado la opcion de TODAS
        'Tienen que ser el mismo cliente
        Codigo = ""
        If txtCodigo(120).Text = "" Then
            Codigo = "NO"
        Else
            If txtCodigo(120).Text <> txtCodigo(121).Text Then Codigo = "NO"
        End If
        
        If Codigo <> "" Then
            Codigo = ""
            Codigo = DevuelveDesdeBD(conAri, "count(*)", "scafac", "codtipom", "FTI", "T")
            If Val(Codigo) = 0 Then
                Codigo = ""
            Else
                Codigo = "    .-Tickets"
            End If
            If vParamAplic.TieneTelefonia2 > 0 Then Codigo = Codigo & vbCrLf & "    .-Telefonia"
            If vParamAplic.AguasPotables Then Codigo = Codigo & vbCrLf & "    .-Agua"
            If Codigo <> "" Then Codigo = vbCrLf & "No se imprimir�n las facturas de: " & vbCrLf & Codigo
                
            
            
            Codigo = "Debe indicar un �nico cliente para la reimpresi�n de todas sus facturas." & vbCrLf & Codigo
            MsgBox Codigo, vbExclamation
            Exit Sub
        Else
            'Indicaremos que no saldran los tickets
            Codigo = "1900-01-01"
            If Me.txtCodigo(85).Text <> "" Then Codigo = Format(txtCodigo(85).Text, FormatoFecha)
            Codigo = "scafac.fecfactu >= '" & Codigo & "' AND codtipom = 'FTI' and codclien"
            
            
            Codigo = DevuelveDesdeBD(conAri, "count(*)", "scafac", Codigo, txtCodigo(120).Text)
            If Val(Codigo) = 0 Then
                Codigo = ""
            Else
                MsgBox "No se imprimir�n los tickets del cliente.", vbExclamation
            End If
        End If
    End If
    
    
    
    If Val(txtCodigo(119).Text) <= 0 Then txtCodigo(119).Text = "1"
    
    

    InicializarVbles
    
    '===================================================
    '================= FORMULA =========================
    'ser� la cadena WHERE para el Informe
    
    'Desde/Hasta tipo movimiento
    '---------------------------------------------
    TipoM = Mid(Me.cboTipomov(1).List(Me.cboTipomov(1).ListIndex), 1, 3)
    If TipoM <> "" Then
    
        If TipoM = "***" Then
            '*TODAS menos FTI, FAG , FAT
    
            Codigo = " NOT ({scafac.codtipom} IN ['FTI','FAG','FAT','FAZ'])"
            If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
            cadSelect = cadFormula
    
        Else
            Codigo = "({scafac.codtipom}='" & TipoM & "') "
            If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
            cadSelect = cadFormula
        End If
    End If

    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtCodigo(120).Text <> "" Or txtCodigo(121).Text <> "" Then
        Codigo = "{scafac.codclien}"
        If Not PonerDesdeHasta(Codigo, "N", 120, 121, "") Then Exit Sub
    End If
    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtCodigo(149).Text <> "" Or txtCodigo(150).Text <> "" Then
        Codigo = "{scafac.codagent}"
        If Not PonerDesdeHasta(Codigo, "N", 149, 150, "") Then Exit Sub
    End If
    
    
    
    
    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtCodigo(83).Text <> "" Or txtCodigo(84).Text <> "" Then
        Codigo = "{scafac.numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 83, 84, "") Then Exit Sub
    End If
    
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(85).Text <> "" Or txtCodigo(86).Text <> "" Then
        Codigo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "") Then Exit Sub
    End If
    
    
    'Mayo 2016
    'Marca de que al cliente se le envian las facturas por email
    If chk_duplicado2(3).Value = 1 Then
        
        
        
        
        Codigo = Replace(cadSelect, "{", "")
        Codigo = Replace(Codigo, "}", "")
        
        Set miRsAux = New ADODB.Recordset
        Codigo = "Select distinct codclien FROM scafac WHERE " & Codigo & " AND codclien IN "
        Codigo = Codigo & " (select codclien from sclien where EnvFraEmail=0 ORDER BY codclien)"
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Codigo = ""
        NumeroDeCopias = 0
        While Not miRsAux.EOF
            Codigo = Codigo & ", " & miRsAux!codClien
            NumeroDeCopias = NumeroDeCopias + 1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        If NumeroDeCopias > 1000 Then
            MsgBox "Error numero clientes distintos. 1000. [Error Crystal REPORT]", vbExclamation
            Exit Sub
        End If
        
        If NumeroDeCopias = 0 Then
            MsgBox "Ningun dato con estos criterios", vbExclamation
            Exit Sub
        End If
        Codigo = Mid(Codigo, 2)
        
        NumeroDeCopias = Val(txtCodigo(119).Text)
        
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & "  {scafac.codclien} IN [" & Codigo & "]"
        
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & "  {scafac.codclien} IN (" & Codigo & ")"
        
        
    End If
    
    
   
    
    If CBool(Me.chk_duplicado2(0).Value) Then
        cadParam = "pDuplicado=1|"
    Else
        cadParam = "pDuplicado=0|"
    End If
    
    EsDeTelCabFactura = False
    If TipoM = "FAT" Then
        If vParamAplic.TieneTelefonia2 = 1 Then EsDeTelCabFactura = True
    End If
    
    
    EsDeTelCabFactura = False    'que podremos quitar
    
    
    'Factura telefonia
    'If TipoM = "FAT" Then
    If EsDeTelCabFactura Then
    
        'Las facturas de telefonia
        Codigo = Replace(cadSelect, "{", "")
        Codigo = Replace(Codigo, "}", "")
        
        Set miRsAux = New ADODB.Recordset
        Codigo = "Select codtipom,year(fecfactu) ano,numfactu FROM scafac WHERE " & Codigo & " ORDER BY year(fecfactu),numfactu"
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Codigo = ""
        cadFormula = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoM, "T")
        cadFormula = "{tel_cab_factura.Serie} = '" & cadFormula & "'"
        NumRegElim = 0
        NomTablaLin = "" 'ara el a�o
        While Not miRsAux.EOF
            If NumRegElim <> miRsAux!Ano Then
                NumRegElim = miRsAux!Ano
                NomTablaLin = NomTablaLin & ", " & NumRegElim
            End If
            Codigo = Codigo & ", " & miRsAux!Numfactu
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        
        
        
        If Codigo = "" Then
            MsgBox "Ningun valor devuelto", vbExclamation
            Exit Sub
        Else
             Codigo = Mid(Codigo, 2)
             cadFormula = cadFormula & " AND ({tel_cab_factura.NumFact}) IN [" & Codigo & "]"
        End If
        
        
        NomTablaLin = Trim(Mid(NomTablaLin, 2))
        cadFormula = cadFormula & " AND {tel_cab_factura.ano} "
        If Len(NomTablaLin) = 4 Then
            'SOLO un a�o
            cadFormula = cadFormula & " = " & NomTablaLin
        Else
            cadFormula = cadFormula & " IN [" & NomTablaLin & "]"
        End If
        NomTablaLin = ""
        
        
        
    
    End If

    
    
    TipoFactura = 0
    Codigo = Mid(cboTipomov(1).Text, 1, 3)
    If Codigo <> "" Then
        If Codigo = "FTI" Then
            TipoFactura = 1                        'Facturas ticket
        Else
            If Codigo = "FAZ" Then TipoFactura = 2 'FAacturas B
            If Codigo = "FAT" Then
                If vParamAplic.TieneTelefonia2 > 0 Then
                    TipoFactura = 3 'FAacturas telefonia
                
                
                
                    'SOLO para los que no son ALZIRA
                    'Ya que los que no son alzira leen de scafac
                    If vParamAplic.TieneTelefonia2 > 1 Then
                        'Febrero 2014
                        'TELEFONIA   FAT
                        'En la ficha del telefono hay un campo(Factura) que utilizaremos para
                        'saber si se le imprime
                        'Ese campo se lleva a la scafac a numpedcl
                        '#   0.- Se imprime
                        '#   1.- Va por email
                        If chk_duplicado2(2).Value = 1 Then
                            Codigo = "({scafac1.numpedcl}=0) or isnull({scafac1.numpedcl}) "
                            If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
                       
                        End If
                    End If
                End If
            End If
            
            
            'En taxco
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                'Facturas alvic TipoFactura
                Codigo = Mid(cboTipomov(1).Text, 1, 3)
                Codigo = "|" & Codigo & "|"
                If InStr(1, "|FA1|FA2|FA3|FAB|FAD|", Codigo) > 0 Then TipoFactura = 6
            End If
            
            
            
        End If
    End If
    
    
    ImprimirFacturas cadFormula, cadParam, cadSelect, TipoFactura, CByte(txtCodigo(119).Text), Me.chk_duplicado2(1).Value, False, True
    
End Sub

Private Sub cmdAceptarTrasHco_Click()
Dim devuelve As String
Dim Cad As String
'IMPRIME INFORME y DESPUES PREGUNTA SI TRASPASAR AL HISTORICO

    InicializarVbles
    
    If Me.cboMotTras.ListIndex < 0 Then
        MsgBox "Seleccione motivo de traspaso", vbExclamation
        Exit Sub
    End If
    
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '===================================================
    '================= FORMULA =========================
    'ser� la cadena WHERE para el Informe y la SQL del Traspaso a Hco
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(43).Text <> "" Or txtCodigo(44).Text <> "" Then
        Codigo = "{scapre.codclien}"
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(Codigo, "N", 43, 44, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Or txtCodigo(46).Text <> "" Then
        Codigo = "{scapre.codagent}"
        Cad = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(Codigo, "N", 45, 46, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(22).Text <> "" Or txtCodigo(23).Text <> "" Then
        Codigo = "{scapre.fecofert}"
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 22, 23, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta N� OFERTA
    '---------------------------------------------
    If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
        Codigo = "{scapre.numofert}"
        Cad = "pDHOferta=""N� Oferta: "
        If Not PonerDesdeHasta(Codigo, "N", 20, 21, Cad) Then Exit Sub
    End If
    
    'Seleccionar para estos criterios solo las Ofertas que no esten Aceptadas
    '------------------------------------------------------------------------
    devuelve = " {scapre.aceptado} = 0 "
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub 'Para Crystal
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub 'Para MySQL
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    'ANTES
    'If Not HayRegParaInforme("scapre", cadSelect) Then Exit Sub
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    devuelve = "Select count(*) FROM scapre WHERE " & cadSelect
    NumRegElim = NumRegistros(devuelve, conAri)
    
    If NumRegElim = 0 Then
        MsgBox "No existen ofertas a traspasar", vbExclamation
        Exit Sub
    End If
    
    devuelve = vbCrLf & vbCrLf & "Va a traspasar a hist�rico  " & NumRegElim & "  ofertas"
    devuelve = devuelve & vbCrLf & vbCrLf & "�Continuar?"

    'Mostrar el Informe
    LlamarImprimir False, False
    
    'Preguntar si Traspasamos los Datos seleccionados al Hist�rico
    'If MsgBox("�Desea pasar estas Ofertas al Hist�rico?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
    If MsgBox(devuelve, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        If TraspasoOfertaAHco(cadSelect) Then
            MsgBox "Traspaso de Ofertas a Hist�rico realizado correctamente. ", vbInformation
            Unload Me
        End If
    End If
End Sub


Private Sub cmdAcetarConfirm_Click()
'Confirmacion de Pedidos
Dim devuelve As String, campo As String





    If txtCodigo(81).Text = "" Then
        MsgBox "Debe seleccionar una carta para Imprimir la Confirmaci�n de Pedidos.", vbInformation
        PonerFoco txtCodigo(81)
        Exit Sub
    End If
    

    If Me.chkConfirmPed(1).Value = 1 And Me.chkConfirmPed(0).Value = 0 Then
        MsgBox "La opcion de adjuntar pedidos solo es v�lidad para el envio de email.", vbExclamation
        Me.chkConfirmPed(1).Value = 0
        Exit Sub
    End If
    
    
    InicializarVbles
    
    
        
    If Not PonerParamRPT2(37, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
        Exit Sub
    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
    
    'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
    If Me.OptPapelBlanco3.Value = True Then
        campo = "True"
    Else
        campo = "False"
    End If
    cadParam = cadParam & "pPapelB=" & campo & "|"
    numParam = numParam + 1
                    
    'Si se impremen Saldos o no
    If Me.chkImpSaldo.Value = 1 Then
        campo = "True"
    Else
        campo = "False"
    End If
    cadParam = cadParam & "pImpSaldo=" & campo & "|"
    numParam = numParam + 1
    
                    
    'Se pasa como parametro la carta a imprimir
    If Me.txtCodigo(81).Text <> "" Then
        cadParam = cadParam & "pCodCarta=" & CInt(Me.txtCodigo(81).Text) & "|"
    Else
        cadParam = cadParam & "pCodCarta=" & CInt(0) & "|"
    End If
    numParam = numParam + 1
    
    'A�adir la fecha de la carta como parametro del informe
    cadParam = cadParam & "pFecha=""" & txtCodigo(82).Text & """|"
    numParam = numParam + 1
    
    'A�adir la poblacion de la empresa como parametro del informe
    cadParam = cadParam & "pPoblacion=""" & vParam.Poblacion & """|"
    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir Vien desde arriba
    'nomRPT = "rFacPedConfirm.rpt"
    conSubRPT = True
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Fechas de Pedido
    '--------------------------------------------
    If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
        campo = "{" & NomTabla & ".fecpedcl}"
        If Not PonerDesdeHasta(campo, "F", 77, 78, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Clientes de Pedido
    '--------------------------------------------
    If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
        campo = "{" & NomTabla & ".codclien}"
        If Not PonerDesdeHasta(campo, "N", 79, 80, "") Then Exit Sub
    End If
       
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub


    If Me.chkConfirmPed(0).Value = 0 Then
        LlamarImprimir True, False
    Else
        'Generaremos todos los pdf necesarios
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        frmPpal.visible = False
        GenerarConfirmacionPedidos
        Label4(97).Caption = ""
        Set miRsAux = Nothing
        
        frmEMail.Opcion = 6
        frmEMail.Show vbModal
        
        codClien = ""
        FecEntre = ""
        NumCod = ""
        Screen.MousePointer = vbDefault
        Me.visible = False
        frmPpal.visible = True
        Me.visible = True
        Unload Me
    End If
End Sub


Private Sub cmdAcetarRecorda_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim bytPrecio As Byte 'Precio valoracion seleccionado
   
    'Comprobar que hay carta si vamos a imprimir un Recordatorio de Oferta
    If (OpcionListado = 32 And txtCodigo(13).Text = "") Then
        MsgBox "Debe seleccionar una carta para Imprimir el Recordatorio.", vbInformation
        PonerFoco txtCodigo(13)
        Exit Sub
    End If
    
    InicializarVbles
    cadPDFrpt = ""
    
    '===================================================
    '============ PARAMETROS ===========================
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    If OpcionListado = 32 Then
        indRPT = 53 'A�ade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
            Exit Sub
        End If
    
        'Si tipo de Papel es blanco imprimir Datos Empresa en cabecera del Informe
        If Me.OptPapelBlancoR.Value = True Then 'Blanco o con Membrete
            devuelve = "True"
        Else
            devuelve = "False"
        End If
        cadParam = cadParam & "pPapelB=" & devuelve & "|"
        numParam = numParam + 1
                    
        'Se pasa como parametro la carta a imprimir
        If Me.txtCodigo(13).Text <> "" Then
            cadParam = cadParam & "pCodCarta=" & CInt(Me.txtCodigo(13).Text) & "|"
        Else
            cadParam = cadParam & "pCodCarta=" & CInt(0) & "|"
        End If
        numParam = numParam + 1
        
        'A�adir las 2 lineas como parametros del informe
        If Me.txtCodigo(14).Text <> "" Then 'Linea A
            cadParam = cadParam & "pLineaA=""" & Me.txtCodigo(14).Text & """|"
            numParam = numParam + 1
        End If
        If Me.txtCodigo(15).Text <> "" Then 'Linea B
            cadParam = cadParam & "pLineaB=""" & Me.txtCodigo(15).Text & """|"
            numParam = numParam + 1
        End If
    
        'A�adir la poblacion de la empresa como parametro del informe
        cadParam = cadParam & "pPoblacion=""" & vParam.Poblacion & """|"
        numParam = numParam + 1
        'Nombre fichero .rpt a Imprimir
        'nomRPT = "rFacOfeRecorda.rpt"
        nomRPT = nomDocu
        
    Else
        
        indRPT = 33 'A�ade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
            Exit Sub
        End If

        'nomRPT = "rFacOfeValoracion.rpt"
        nomRPT = nomDocu
        
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta CLIENTE
    '--------------------------------------------
    Codigo = "{scapre.codclien}"
    devuelve = CadenaDesdeHasta(txtCodigo(9).Text, txtCodigo(10).Text, Codigo, "N", "Cliente")
    If devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    'Cadena para seleccion Desde y Hasta N� OFERTA
    '--------------------------------------------
    Codigo = "{scapre.numofert}"
    devuelve = CadenaDesdeHasta(txtCodigo(5).Text, txtCodigo(6).Text, Codigo, "N", "N� Oferta")
    If devuelve = "Error" Then
        Exit Sub
    End If
    If Not AnyadirAFormula(cadFormula, devuelve) Then
        Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    Codigo = "{scapre.fecofert}"
    devuelve = CadenaDesdeHasta(txtCodigo(7).Text, txtCodigo(8).Text, Codigo, "F", "Fecha")
    If devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    'Cadena para seleccion Desde y Hasta AGENTE
    '--------------------------------------------
    Codigo = "{scapre.codagent}"
    devuelve = CadenaDesdeHasta(txtCodigo(11).Text, txtCodigo(12).Text, Codigo, "N", "Agente")
    If devuelve = "Error" Then Exit Sub
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    If OpcionListado = 32 Then
        'Cadena para seleccion de Ofertas no Aceptadas
        Codigo = "{scapre.aceptado}=0"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    End If
    
    
    'Noviembre 2013
    'Envio Outlook del recordatoriot
    devuelve = ""
    If OpcionListado = 32 Then
        'Si cliente desde y hasta son el mismo
        If Val(txtCodigo(9).Text) > 0 Then
            If txtCodigo(9).Text = txtCodigo(10).Text Then
                devuelve = Trim(Mid(txtNombre(9).Text, 1, 8)) & Format(txtCodigo(9).Text, "0000")
                devuelve = "6|" & txtCodigo(9).Text & "|" & devuelve & "|" '1|2|2013323|
            End If
        End If
    End If
    
    
    LlamarImprimir True, False, devuelve
End Sub


Private Sub cmdBajar_Click()
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)

    If Index = 18 Then
        If Me.ProgressBar1.visible Then
            If MsgBox("�Seguro que desea cancelar el proceso?", vbQuestion + vbYesNoCancel) = vbYes Then pImprimeDirecto = True
            
            Exit Sub
        End If
    End If
    Unload Me
End Sub


Private Sub cmdCliPot_Click()
Dim devuelve As String


    'Comprobaciones
    If OpcionListado = 401 Then
        devuelve = ""
        
        If txtCodigo(138).Text = "" Or txtNombre(138).Text = "" Then
            devuelve = devuelve & vbCrLf & "   -Seleccione carta"
            NumRegElim = 138
        End If
        If txtCodigo(135).Text = "" Then
            devuelve = devuelve & vbCrLf & "   -Campo ""A la atenci�n de"""
            NumRegElim = 135
        End If

            
        If devuelve <> "" Then
            devuelve = "Faltan campos: " & vbCrLf & devuelve
            MsgBox devuelve, vbExclamation
            PonerFoco txtCodigo(NumRegElim)
            Exit Sub
        End If

    End If
    
    
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
       '--------------------------------------------
     If txtCodigo(131).Text <> "" Or txtCodigo(132).Text <> "" Then
        Codigo = "{sclipot.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(Codigo, "N", 131, 132, devuelve) Then Exit Sub
    End If
    
    
    If txtCodigo(136).Text <> "" Or txtCodigo(137).Text <> "" Then
        Codigo = "{sclipot.codactiv}"
        'Parametro Desde/Hasta Actividad
        devuelve = "pDHActividad=""Actividad: "
        If Not PonerDesdeHasta(Codigo, "N", 136, 137, devuelve) Then Exit Sub
    End If
                    
    'Cadena para seleccion D/H COD. POSTAL
    '--------------------------------------------
     If txtCodigo(133).Text <> "" Or txtCodigo(134).Text <> "" Then
        Codigo = "{sclipot.codpobla}"
        'Parametro Desde/Hasta cod. Postal
        devuelve = "pDHcpostal=""CPostal: "
        If Not PonerDesdeHasta(Codigo, "T", 133, 134, devuelve) Then Exit Sub
    End If
    
    
    If OpcionListado <= 401 Then   'eitquetas y cartas
        'Parametro a la Atencion de
        cadParam = cadParam & "pAtencion=""Att. " & txtCodigo(135).Text & """|"
        numParam = numParam + 1
     End If
     
     
    'seleccionamos todos los clientes por defecto,
    cadSelect = QuitarCaracterACadena(cadFormula, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")

    'No controlo el error. Si no estaq explotara
    Titulo = " clientes potenciales"
    Select Case OpcionListado
    Case 400
        NumRegElim = 59 'etiqueta
        Titulo = "Etiquetas" & Titulo
    Case 401
        NumRegElim = 58 'carta
        Titulo = "Cartas" & Titulo
    Case 402
        NumRegElim = 60 'carta
        Titulo = "CRM " & Titulo
    End Select
    nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", CStr(NumRegElim))  'Carta a clientes potenciales
    

    
    cadParam = cadParam & "pCodCarta= " & txtCodigo(138).Text & "|"
    numParam = numParam + 1
    
    'Preseleccion
    indCodigo = 1
    If OpcionListado = 402 Then
        'CRM. Ha puesto D/H mismo cliente
        If txtCodigo(131).Text = txtCodigo(132).Text And txtCodigo(131).Text <> "" Then
            indCodigo = 0
            cadFormula = "{sclipot.codclien} = " & txtCodigo(131).Text
        End If
    End If
    If indCodigo = 1 Then
        Set frmMen = New frmMensajes
        frmMen.cadWhere = cadSelect
        frmMen.OpcionMensaje = 22 'listado clientes  potenciales con los desde hasta
        frmMen.Show vbModal
        Set frmMen = Nothing
        If cadSelect = "" Then Exit Sub
    End If
    
    
    cadFormula = Replace(cadFormula, "{sclien.", "{sclipot.")
        
    
    conSubRPT = True
    
    LlamarImprimir False, False
    
    
End Sub

Private Sub cmdComprobarCCC_NIF_Secciones_Click()


'    If chkVarios(8).Value = 1 Then
'        MsgBox "Falta comprobar", vbExclamation
'        Exit Sub
'    End If



    InicializarVbles
    
    'comprobar que se ha introducido codlien
    '---------------------------------------------------------
    If Trim(txtCodigo(147).Text) <> "" Or Trim(txtCodigo(148).Text) <> "" Then
        'Para Crystal Report
        cadParam = "{sclien.codclien}"
        nomRPT = "pdh1= ""Cliente: " 'Parametro Desde/Hasta Fecha
        If Not PonerDesdeHasta(cadParam, "N", 147, 148, nomRPT) Then Exit Sub
    End If
    
   CadenaDesdeOtroForm = nomRPT & """|"
   Screen.MousePointer = vbHourglass
   Label9(52).Caption = "Inicio proceso"
   Label9(52).Refresh
   If cadSelect <> "" Then
        cadSelect = cadSelect & " AND "
        cadSelect = QuitarCaracterACadena(cadSelect, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
   End If
   cadSelect = cadSelect & "codbanco>0 and clivario=0"
   
   
   If ComprobarDatosProcesoCCC(cadSelect, Label9(52), Me.chkVarios(8).Value = 1) Then
        Label9(52).Caption = "Mostrar datos"
        Label9(52).Refresh
        Screen.MousePointer = vbHourglass
        frmVarios3.Opcion = 1
        frmVarios3.Show vbModal
    Else
        MsgBox "Ningun dato", vbExclamation
   End If
   Label9(52).Caption = ""
   CadenaDesdeOtroForm = ""
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEnvioMail_Click()

    
    cmdEnvioMail.Enabled = False
    cmdEnvioMail.Refresh
    
    HacerCmdEnvioMail_Click
    
    cmdEnvioMail.Enabled = True
    
    Screen.MousePointer = vbDefault
    lblInd.Caption = ""
End Sub

Private Sub HacerCmdEnvioMail_Click()
Dim RS As ADODB.Recordset
Dim VanLosFTI As Boolean
Dim B As Boolean
Dim ClienteVario As Long
Dim SoloFacturaTelefonia As Boolean
Dim Aux As String


    'ENVIO MAIL
    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    
    'FACTURAE
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Copiar a   vParamAplic.PathFacturaE los archivos con el nombre como debe ser
    
    SoloFacturaTelefonia = False
    If OpcionListado = 315 Then
        If Text1(0).Text = "" Then
            MsgBox "Ponga el asunto", vbExclamation
            Exit Sub
        End If
    Else
        Codigo = ""
        If vParamAplic.PathFacturaE = "" Then
            Codigo = "Falta configurar parametros"
        Else
        
            'Enero 2016
            'Si es un path en RED , con \\, tiene que acabar en \ si no da error
            cadFormula = vParamAplic.PathFacturaE
            If Mid(cadFormula, 1, 2) = "\\" Then
                If Right(cadFormula, 1) <> "\" Then cadFormula = cadFormula & "\"
            End If
            
            'If Dir(vParamAplic.PathFacturaE, vbDirectory) = "" Then Codigo = "No existe carpeta"
            If Dir(cadFormula, vbDirectory) = "" Then Codigo = "No existe carpeta"
        End If
        If Codigo <> "" Then
            MsgBox Codigo, vbExclamation
            Exit Sub
        End If
    End If
        
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    
    'AHora pongo los tipo de facturas
    cadFormula = ""
    cadSelect = ""  'ME dira si estan todas o no
    VanLosFTI = False
    
    NomTabla = ""  'pARA SABER SI SOLO VAN LOS DE LA MARCA de telefonos (sclientfno)
    
    For indCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
        If Me.ListTipoMov(1000).Selected(indCodigo) Then
            'Esta checkeado
            cadFormula = cadFormula & " OR scafac.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) & "'"
            If Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) = "FTI" Then VanLosFTI = True
            NomTabla = NomTabla & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3))
        Else
            cadSelect = "NO"
        End If
    Next indCodigo
    
    
    If cadFormula = "" Then
        MsgBox "Seleccione algun tipo de factura", vbExclamation
        Exit Sub
    Else
        cadFormula = Mid(cadFormula, 4)
    End If
    If cadSelect = "" Then
        'Significa que estan todos. No tiene sentido poner que codtipo='fr or codtipo='FT  ESTAN TODAS
        'cadFormula = " scafac.codtipom <> 'FTI'"  'antes
        cadFormula = " scafac.codtipom <> ''"
        'FEBRERO 2013.
        'Van todos
        
    End If
    
    'SOLO VAN LAS FAT
    If NomTabla = "FAT" Then SoloFacturaTelefonia = True
        
    
    'En nomtabla tendre
    NomTabla = "(" & cadFormula & ")"

    InicializarVbles
    cadFormula = ""
    cadSelect = ""
    
    
    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
        Codigo = "scafac.fecfactu"
        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
    End If
    
    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
        Codigo = "scafac.numfactu"
        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
    End If
        
        
        
        
        
    'Para las de telefonia
    Aux = cadSelect
    
    If txtCodigo(110).Text <> "" Or txtCodigo(111).Text <> "" Then
        Codigo = "scafac.codclien"
        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
    End If
        
        
    'Junio 2011
    'facturaE. Si no esta marcado el chk de ya trasapasada
    
    If OpcionListado = 316 Then
        If Me.chkMail(1).Value = 0 Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & " (scafac.EnFacturaE = 0 )"
        End If
        
        
        
    End If
        
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & NomTabla
    cadSelect = " WHERE " & cadSelect
    
    Set RS = New ADODB.Recordset
    DoEvents
    
    
        
    'Ahora insertare en la tabla temporal tminformes las facturas que voy a generar pdf.
    'Abril 2020 AHora ponemos en numlotes el coddirec.
    'Luego buscaremos el si tiene email y actualiaremos el nomartic que llevar� la direccion email
    'Finalmante , antes de lanzar el proceso, l volveremos a poner a ''
    Codigo = "insert into tmpnlotes (codusu,numalbar,codprove,codartic,numlinea,fechaalb,codalmac,cantidad,numlotes) "
    Codigo = Codigo & " values "
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass
    lblInd.Caption = "Devolver registros"
    lblInd.Refresh
    
    
    
    If VanLosFTI Then
        NomTabla = DevuelveDesdeBD(conAri, "codclien", "spatpvg", "1", "1")
        If NomTabla = "" Then NomTabla = "0"
        ClienteVario = Val(NomTabla)
    End If
    
    
    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select codtipom,numfactu,codclien,fecfactu,totalfac,coddirec from scafac  " & cadSelect
    
    If SoloFacturaTelefonia Then
        'Imprimir solo
        
        
        'If MsgBox("�Enviar s�lo a tel�fonos con la marca de enviar por email?", vbQuestion + vbYesNo) = vbYes Then
         If Me.chkMail(2).Value = 1 Then
                If Aux <> "" Then
                    Aux = Aux & " AND "
                    Aux = Replace(Aux, "scafac", "scafac1")
                End If
                Aux = Aux & "codtipom='FAT' AND numpedcl=1"
                
                
               NomTabla = NomTabla & " AND (codtipom,numfactu,fecfactu) IN (select codtipom,numfactu,fecfactu FROM"
               NomTabla = NomTabla & " scafac1 WHERE " & Aux & ")"
                    
              
        End If
    End If
    
    
    
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY codtipom"
    NomTablaLin = ""
    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not RS.EOF
    
        lblInd.Caption = RS!Numfactu
        lblInd.Refresh
        
        B = True
        If VanLosFTI Then
            If RS!codtipom = "FTI" Then
                If RS!codClien = ClienteVario Then B = False
            End If
        End If
        If B Then
            NomTabla = ", (" & vUsu.Codigo & ",'" & RS!codtipom & "'," & RS!codClien & "," & RS!Numfactu & "," & CStr(RS!Numfactu Mod 32000) & ",'" & Format(RS!FecFactu, FormatoFecha)
            
            'El tipo de informe lo guardare en el ultimo campo
            'El report es el = 12
            NomTabla = NomTabla & "',12," & TransformaComasPuntos(CStr(DBLet(RS!TotalFac, "N")))
            
            'Abril 2020. Grabaamos en numlotes el coddirec, para ver si tiene direcion email .
            'Si no pondremos la del cliente
            NomTabla = NomTabla & "," & DBSet(RS!CodDirec, "T", "S") & ")"
            
            NomTablaLin = NomTablaLin & NomTabla
            
            
                        
            NumRegElim = NumRegElim + 1
                
            If Len(NomTablaLin) > 2000 Then
                NomTablaLin = Mid(NomTablaLin, 2) 'quiito la primera coma
                conn.Execute Codigo & NomTablaLin
                NomTablaLin = ""
            End If
                
                
           If (NumRegElim Mod 50) = 0 Then DoEvents
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If Len(NomTablaLin) > 0 Then
        NomTablaLin = Mid(NomTablaLin, 2) 'quiito la primera coma
        conn.Execute Codigo & NomTablaLin
        NomTablaLin = ""
    End If
    
    B = False
    If chkMail(3).Value = 1 Then B = True
    'La marca de envio Factura email es para Envio FACTURAE si se le envia o no
    If OpcionListado = 316 And vParamAplic.NumeroInstalacion = vbTaxco Then B = True
    
    'If chkMail(3).Value = 1 Then
    If B Then
        'Vamos a quitar aquellos clientes que no tienen la marca
        lblInd.Caption = "Clientes envio email/facturae"
        lblInd.Refresh
        Espera 0.5
        
        NomTabla = "DELETE FROM tmpnlotes"
        NomTabla = NomTabla & " WHERE  codusu=" & vUsu.Codigo & " and not codprove in (select codclien from sclien where EnvFraEmail =1)"
        conn.Execute NomTabla
        
        
        NomTabla = DevuelveDesdeBD(conAri, "count(*)", "tmpnlotes", "codusu", CStr(vUsu.Codigo))
        NumRegElim = Val(NomTabla)
        
        lblInd.Caption = ""
    End If
    

    If NumRegElim = 0 Then
        If OpcionListado = 316 Then
            NomTabla = "Ning�na factura para traspasar a FacturaE"
        Else
            NomTabla = "Ningun dato a enviar por mail"
        End If
        
        MsgBox NomTabla, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------------------------------
    '
    'Ahora cojemos las facturas que son FVA pero tienen numero terminal. COn el desde /hasta seleccionado
    'MIRAMOS en la tabla scafac1
    lblInd.Caption = "Leyendo fav "
    lblInd.Refresh
    'Compruebo si tiene codclien
    NomTabla = "select scafac1.* from scafac1 ,scafac where scafac1.codtipom=scafac.codtipom and scafac1.numfactu=scafac.numfactu and scafac1.fecfactu =scafac.fecfactu"
    'NomTabla = "Select codtipom,numfactu,fecfactu from scafac1   " & cadSelect
    'El cad select LLEVA el where.  Se lo quito
    cadSelect = Mid(cadSelect, 7)
    NomTabla = NomTabla & " AND " & cadSelect & "  AND numtermi>=0  "
    
    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        lblInd.Caption = DBLet(RS!codtipom, "T") & ": " & RS!Numfactu
        lblInd.Refresh
        NomTabla = "numalbar = '" & RS!codtipom & "' AND fechaalb = '" & Format(RS!FecFactu, FormatoFecha) & "' AND numlinea = " & CStr(RS!Numfactu Mod 32000)
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        NomTabla = "UPDATE tmpnlotes SET codalmac = 18 WHERE codusu = " & vUsu.Codigo & " AND " & NomTabla
        conn.Execute NomTabla
    
    
        RS.MoveNext
    Wend
    RS.Close
    
    
    lblInd.Caption = "Generando documentos"
    lblInd.Refresh
    
    
    
    
    'AHora las fras  FAT tienen otro report
    If vParamAplic.TieneTelefonia2 > 0 Then
        'ALZIRA
        NomTabla = "UPDATE tmpnlotes SET codalmac = 63 WHERE codusu = " & vUsu.Codigo & " AND numalbar= 'FAT'"
        conn.Execute NomTabla
    End If
    'Los tikets=66
    NomTabla = "UPDATE tmpnlotes SET codalmac = 66 WHERE codusu = " & vUsu.Codigo & " AND numalbar= 'FTI'"
    conn.Execute NomTabla
    
    'Proyectos
    NomTabla = "UPDATE tmpnlotes SET codalmac = 89 WHERE codusu = " & vUsu.Codigo & " AND numalbar= 'FPY'"
    conn.Execute NomTabla
    
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'Facturas alvic
        NomTabla = "UPDATE tmpnlotes SET codalmac = 93 WHERE codusu = " & vUsu.Codigo & " AND numalbar in ('FA1','FA2','FAB','FAD')"
        conn.Execute NomTabla
    End If
    
    
    
    'Numero de registros
    
    NomTabla = NumRegElim
    
    If OpcionListado = 315 Then
    
            'Quito los que
    
    
    
            'AHora ya tengo todos los datos de las facturas que voy  a imprimir
            'Entonces copruebo si para los clientes si tienen puesto el campo mail o no
            If optEnvioMail(0).Value Then
                'Selecciona mail comercial
                cadSelect = "2"  'de maiclie2
            Else
                cadSelect = "1"  'de maiclie1
            End If
            
            'Actualizamos en nomartic la direccion de envio
            
            cadSelect = "UPDATE tmpnlotes,sclien set nomartic=maiclie" & cadSelect
            cadSelect = cadSelect & " WHERE codusu = " & vUsu.Codigo & " and codclien=codprove"
            conn.Execute cadSelect
            
            'Junio 2021
            'A�ado ; al final si "tienen algo" de email
            lblInd.Caption = "Multidirecciones"
            lblInd.Refresh
    
            cadSelect = "UPDATE tmpnlotes SET  nomartic=concat(nomartic,';')"
            cadSelect = cadSelect & " WHERE codusu = " & vUsu.Codigo & " and coalesce(nomartic,'')<>''"
            conn.Execute cadSelect
            'Junio 2021
            'Vamos a ver si algun cliente tiene puesto el envio a mas de una direccion
            cadSelect = "SELECT codclien, GROUP_CONCAT( maidirec separator ';') from scliendp where incluirenviofacturacion =1 and maidirec<>'' and codclien "
            cadSelect = cadSelect & " IN (select distinct codprove from tmpnlotes where codusu=" & vUsu.Codigo & ")  GROUP BY codclien"
            RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not RS.EOF
                cadSelect = RS.Fields(1) & ";"
                cadSelect = "UPDATE tmpnlotes SET nomartic =concat(nomartic," & DBSet(cadSelect, "T") & ") where codusu=" & vUsu.Codigo & " AND codprove= " & RS!codClien
                conn.Execute cadSelect
                NumRegElim = NumRegElim + 1
                RS.MoveNext
            Wend
            RS.Close
            
            If NumRegElim > 0 Then
                'Ha updateado emails
                'Vere SI son correctos, ya que como mucho aceptamos 255 caracteres para la direccion email
                cadSelect = "select * from tmpnlotes where codusu=" & vUsu.Codigo & " AND nomartic<>'' and right(nomartic,1)<>';'"
                RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                cadSelect = ""
                NumRegElim = 0
                While Not RS.EOF
                    cadSelect = "Cliente: " & RS!Codprove & vbCrLf & RS!NomArtic & vbCrLf & vbCrLf
                    RS.MoveNext
                Wend
                RS.Close
                If cadSelect <> "" Then
                    cadSelect = "eMail incorrecto: " & vbCrLf & cadSelect
                    MsgBox cadSelect, vbExclamation
                    Exit Sub
                End If
            End If
            lblInd.Caption = "Envio"
            lblInd.Refresh

            
            
            
            'Sigo como estaba
            
            
            'Si hay algun departamento con direccion email
            cadSelect = "UPDATE tmpnlotes,sdirec set nomartic=maidirec"
            cadSelect = cadSelect & " WHERE codusu = " & vUsu.Codigo & " and codclien=codprove"
            cadSelect = cadSelect & " and codclien=codprove and coddirec=numlotes +0  and maidirec <>''"
            conn.Execute cadSelect
            
            
            cadSelect = "Select * from tmpnlotes "
            cadSelect = cadSelect & " WHERE codusu = " & vUsu.Codigo & " and coalesce(nomartic,'')=''"
                        
            RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                RS.MoveNext
            Wend
            RS.Close
            
            If NumRegElim > 0 Then
                If MsgBox("Tiene cliente sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                    
                'Si no salimos borramos
                
                cadSelect = "DELETE from tmpnlotes where codusu =" & vUsu.Codigo & " and coalesce(nomartic,'') =''"
                conn.Execute cadSelect
                
                
                cadSelect = "Select count(*) from tmpnlotes where codusu =" & vUsu.Codigo
                RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                NumRegElim = 0
                If Not RS.EOF Then
                    If Not IsNull(RS.Fields(0)) Then NumRegElim = DBLet(RS.Fields(0), "N")
                    
                End If
                RS.Close
                
                If NumRegElim = 0 Then
                    'NO hay datos para enviar
                    
                    Screen.MousePointer = vbDefault
                    MsgBox "No hay datos para enviar por mail", vbExclamation
                    Exit Sub
                Else
                    cadSelect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "�Continuar?"
                    If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
                End If
                If NumRegElim = 0 Then
                    Set RS = Nothing
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                NomTabla = NumRegElim
            
            Else
                cadSelect = "Hay " & NomTabla & " facturas para enviar por mail." & vbCrLf & "�Continuar?"
                If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
      
      
            cadSelect = "UPDATE tmpnlotes set numlotes=null WHERE codusu = " & vUsu.Codigo
            conn.Execute cadSelect
      
    Else
        cadSelect = "Hay " & NumRegElim & " facturas para integrar con facturaE." & vbCrLf & "�Continuar?"
        If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

    End If
    
    
    
    
    PonerTamnyosMail2 True
    
    If vUsu.Skin >= 0 Then
        frmPpalN.visible = False
    Else
        frmPpalOld.visible = False
    End If
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label4(22).Caption = "Preparando datos"
    Label4(22).Refresh
    Me.ProgressBar1.Value = 0
    SecuenciaEnProgressBar1 = 0
    Me.ProgressBar1.Max = 1000
    MaxNumerRegistroSEnProgressBar1 = CLng(NomTabla)

    
    
    
    NumRegElim = 0
    If GeneracionEnvioMail(RS) Then NumRegElim = 1
        
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        If OpcionListado = 315 Then
    
          
            cadSelect = "Select nomclien,nomartic as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
    
    
            Label14(22).Caption = "Comienzo envio email"
            Me.Refresh
            DoEvents
            Espera 0.5
            frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(chkMail(0).Value) & "|" & cadSelect & "|"
            frmEMail.Opcion = 4 'Multienvio de facturacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBar1.visible = False
        
        Else
            'Copiar a parametros
            '
            MsgBox "Proceso finalizado", vbExclamation
        
        End If
        Me.Refresh
        DoEvents
        Espera 0.7
        PrepararCarpetasEnvioMail
        Me.ProgressBar1.visible = True
        
        
    End If
    
    
    
    
    'Es para evitar la cantidad de pantallas abriendose y cerrandose
    Me.visible = False
    PonerTamnyosMail2 False
    Espera 1
    Unload Me
    
    If vUsu.Skin >= 0 Then
        frmPpalN.Show
    Else
        frmPpalOld.Show
    End If

    Screen.MousePointer = vbDefault
End Sub
        
        
        
Private Function GeneracionEnvioMail(ByRef RS As ADODB.Recordset) As Boolean
Dim EsdesdeTelCabFact As Boolean
Dim FuerzaRPT As String
Dim C_Aux As String


    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    'Si es la de facturaE voy a upadear numlotes con el numserie
    If OpcionListado = 316 Then
        Label14(22).Caption = "Preparando datos facturae"
        Label14(22).Refresh
        cadSelect = "Select numalbar from tmpnlotes where codusu = " & vUsu.Codigo & " GROUP BY 1"
        RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            cadSelect = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", RS!Numalbar, "T")
            If cadSelect = "" Then
                codClien = codClien & "        " & RS!Numalbar
            Else
                cadSelect = "UPDATE tmpnlotes set numlotes= '" & cadSelect & "' WHERE codusu = " & vUsu.Codigo & " AND numalbar=" & DBSet(RS!Numalbar, "T")
                conn.Execute cadSelect
            End If
            RS.MoveNext
        Wend
        RS.Close
    
    End If
        
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codClien = ""
    SecuenciaEnProgressBar1 = 0
    pImprimeDirecto = False 'para cancelar el proceso
    While Not RS.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & RS!Numalbar & " " & RS!codArtic
        Label14(22).Refresh
        
        If codClien <> RS!codAlmac Then   'If CodClien <> RS!codTipoM Then
            'OTRO TIPO DE DOCUMENTO
            
            '''''If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
            If Not PonerParamRPT2(RS!codAlmac, cadParam, numParam, NumCod, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
                Exit Function
            End If
            codClien = RS!codAlmac
        End If
        FuerzaRPT = ""
        EsdesdeTelCabFact = False
        If RS!Numalbar = "FAT" Then
            If vParamAplic.TieneTelefonia2 = 1 Then EsdesdeTelCabFact = True
        End If
        'If Rs!NumAlbar = "FAT" Then
        If EsdesdeTelCabFact Then
            
            'Factura de telefonia. Lleva otro SELECT     serie
            cadFormula = "{tel_cab_factura.Serie} ='" & RS!numlotes & "' and {tel_cab_factura.Ano} =" & Year(RS!FechaAlb) & " and {tel_cab_factura.NumFact} =" & RS!codArtic
        Else
            'RESTO de facturas
            cadFormula = "({scafac.codtipom}='" & RS!Numalbar & "') "
            cadFormula = cadFormula & " AND ({scafac.numfactu}=" & RS!codArtic & ") "
            cadFormula = cadFormula & " AND ({scafac.fecfactu}= Date(" & Year(RS!FechaAlb) & "," & Month(RS!FechaAlb) & "," & Day(RS!FechaAlb) & "))"
            
            C_Aux = "scafac.codtipom='" & RS!Numalbar & "' "
            C_Aux = C_Aux & " AND scafac.numfactu=" & RS!codArtic & " "
            C_Aux = C_Aux & " AND scafac.fecfactu= " & DBSet(RS!FechaAlb, "F") & " AND 1 "
                        
            FuerzaRPT = DevuelveDesdeBD(conAri, "nomrpt", "scafac", C_Aux, "1")
            If FuerzaRPT = cadPDFrpt Then FuerzaRPT = ""
            If FuerzaRPT <> "" Then
                C_Aux = App.Path & "\informes\" & FuerzaRPT
                If Dir(C_Aux, vbArchive) = "" Then
                    C_Aux = "No existe RPT con el que se imprimi�: " & C_Aux
                    C_Aux = C_Aux & vbCrLf & "�Continuar?"
                    If MsgBox(C_Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Err.Raise 513, , "Cancelado"
                    FuerzaRPT = ""
                Else
                    C_Aux = CStr(cadPDFrpt)
                    cadPDFrpt = FuerzaRPT
                End If
            End If
        End If

          
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = NumCod
            .NombrePDF = cadPDFrpt
            .Opcion = 53
            .Titulo = ""
        End With
        If FuerzaRPT <> "" Then cadPDFrpt = C_Aux
        Load frmImprimir
        Unload frmImprimir
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        SecuenciaEnProgressBar1 = SecuenciaEnProgressBar1 + 1
        PonerPb2
        If (SecuenciaEnProgressBar1 Mod 5) = 0 Then
            Me.Refresh
            DoEvents
            Espera 0.2
            If (SecuenciaEnProgressBar1 Mod 50) = 0 Then
                
                Me.lblInd.Caption = "Exportando ( " & SecuenciaEnProgressBar1 & ")"
                Me.lblInd.Refresh
                Set frmImprimir = Nothing
                Espera 1
                Set frmImprimir = Nothing
                Screen.MousePointer = vbHourglass
            End If
        End If
       
        DoEvents
        
        If pImprimeDirecto Then
            
            pImprimeDirecto = False
            RS.Close
            Set RS = Nothing
            Exit Function
            
        End If
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        If OpcionListado = 315 Then
            FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!Numalbar & Format(RS!codArtic, "0000000") & ".pdf"
        Else
            'Se tiene que llamar base_numserie_numFactura_yyyymmdd.pdf
            
            cadFormula = vEmpresa.BDAriges & "_" & RS!numlotes & "_" & RS!codArtic & "_" & Format(RS!FechaAlb, "yyyymmdd") & ".pdf"
            cadFormula = vParamAplic.PathFacturaE & "\" & cadFormula
            
            Label14(22).Caption = cadFormula
            Label14(22).Refresh
        
            FileCopy App.Path & "\docum.pdf", cadFormula
            
            
            'Ha copiado, luego yo la pongo como en facturaE
            cadFormula = "UPDATE scafac set EnFacturaE=1, nomrpt=coalesce(nomrpt," & DBSet(NumCod, "T") & ") WHERE codtipom='" & RS!Numalbar & "' AND numfactu=" & RS!codArtic
            cadFormula = cadFormula & " AND fecfactu='" & Format(RS!FechaAlb, FormatoFecha) & "'"
            ejecutar cadFormula, False
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub cmdImpresionCRM_Click()
     InicializarVbles
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(139).Text <> "" Or txtCodigo(140).Text <> "" Then
        nomRPT = "{sclien.codclien}"
        'Parametro Desde/Hasta Cliente
        Titulo = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(nomRPT, "N", 139, 140, Titulo) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtCodigo(141).Text <> "" Or txtCodigo(142).Text <> "" Then
        nomRPT = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        Titulo = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(nomRPT, "N", 141, 142, Titulo) Then Exit Sub
    End If
    
    
    
    
    
    '--------------------------------------------
     If txtCodigo(53).Text <> "" Or txtCodigo(54).Text <> "" Then
        nomRPT = "{sclien.codactiv}"
        'Parametro Desde/Hasta Actividad
        Titulo = "pDHActividad=""Actividad: "
        If Not PonerDesdeHasta(nomRPT, "N", 53, 54, Titulo) Then Exit Sub
    End If
                    
    'Cadena para seleccion SITUACION
    '--------------------------------------------
    If txtCodigo(57).Text <> "" Then
        nomRPT = "{sclien.codsitua}=" & txtCodigo(57).Text
        If Not AnyadirAFormula(cadFormula, nomRPT) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, nomRPT) Then Exit Sub
    End If
    
    'Los checks
    nomRPT = ""
    If Me.chkVarios(4).Value Then nomRPT = nomRPT & " AND {sclien.limcredi} > 0"
    If Me.chkVarios(5).Value Then nomRPT = nomRPT & " AND {sclien.credipriv} = 1"
    If Me.chkVarios(6).Value Then nomRPT = nomRPT & " AND {sclien.codaseg} <>''"
    
    If Me.chkVarios(7).Value Then nomRPT = nomRPT & AnadirClientesCobrosPendientes
    
    
    If nomRPT <> "" Then
        nomRPT = Mid(nomRPT, 5) 'quitamos el primer and
        If Not AnyadirAFormula(cadFormula, nomRPT) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, nomRPT) Then Exit Sub
    End If
    
    cadSelect = Replace(cadSelect, "{", "(")
    cadSelect = Replace(cadSelect, "}", ")")
    If cadSelect = "" Then cadSelect = "1=1"
    Titulo = "Select count(*) from sclien WHERE " & cadSelect
    NumRegElim = NumRegistros(Titulo, conAri)
    Titulo = ""
    nomRPT = ""
    If NumRegElim = 0 Then
        MsgBox "Ning�n dato a mostrar", vbExclamation
        Exit Sub
    End If
    
    
     Set frmMen = New frmMensajes
    frmMen.cadWhere = cadSelect
    frmMen.OpcionMensaje = 8 'Etiquetas clientes
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    
    'El report
    nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "46")
    If nomRPT = "" Then
        MsgBox "Falta configurar en informes(46)", vbExclamation
        Exit Sub
    End If
    'Borrar datos
    ejecutar "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo, False
    
    Titulo = "Select count(*) from sclien WHERE " & cadSelect
    NumRegElim = NumRegistros(Titulo, conAri)
  
    'Vamos p'alla
    Screen.MousePointer = vbHourglass
    Label4(123).Caption = "Incio proceso"
    Label4(124).Caption = ""
    pbCRM.Max = NumRegElim
    pbCRM.Value = 0
    
    'If pbCRM.Max > 5 Then frmPpal.Hide
    
    Me.FrameCRM.Enabled = False
    Me.FrameCRMProgess.visible = True
    Me.Refresh
    Espera 0.5
    indCodigo = 0 'Indicara si cancela el preoceso de impresion
    HacerImpresionCRM
    
    
    
    Me.FrameCRM.Enabled = True
    Me.FrameCRMProgess.visible = False
   ' If pbCRM.Max > 5 Then frmPpal.Show
    Screen.MousePointer = vbDefault
    If indCodigo = 0 Then Unload Me  'Ha ido bien
End Sub

Private Sub cmdListadoOfertasHco_Click()
Dim campo As String
 Dim Cad As String
 
    
    InicializarVbles
    
    Codigo = ""
    NumeroDeCopias = 0
    NomTabla = ""
    For NumRegElim = 1 To Me.lwMotBaja.ListItems.Count
        If lwMotBaja.ListItems(NumRegElim).Checked Then
            NumeroDeCopias = NumeroDeCopias + 1
            Codigo = Codigo & ", " & Mid(lwMotBaja.ListItems(NumRegElim).Key, 2)
            NomTabla = NomTabla & "      " & Mid(lwMotBaja.ListItems(NumRegElim).Key, 2) & "- " & lwMotBaja.ListItems(NumRegElim).Text
        End If
    Next
    
    If NumeroDeCopias = 0 Then
        MsgBox "Seleccione algun motivo", vbExclamation
        Exit Sub
    End If
    
    
    
    
    CadenaDesdeOtroForm = ""

    If txtCodigo(164).Text <> "" Or txtCodigo(165).Text <> "" Then
        campo = "{schpre.fecofert}"
        Cad = "Fecha: "
        If Not PonerDesdeHasta(campo, "F", 164, 165, Cad) Then Exit Sub
        CadenaDesdeOtroForm = Trim(CadenaDesdeOtroForm & Space(10) & Cad)
    End If
    
    
    
    If txtCodigo(162).Text <> "" Or txtCodigo(163).Text <> "" Then
        campo = "{schpre.codclien}"
        Cad = "Cliente: "
        If Not PonerDesdeHasta(campo, "N", 162, 163, Cad) Then Exit Sub
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad
    End If
    If Combo1.ListIndex > 0 Then
        campo = " {schpre.aceptado} = " & IIf(Combo1.ListIndex = 1, "1", "0")
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        campo = " schpre.aceptado = " & IIf(Combo1.ListIndex = 1, "1", "0")
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "    Aceptada: " & Combo1.Text
    
    End If
    
    
    If NumeroDeCopias = lwMotBaja.ListItems.Count Then
        'Ha codigo todos los motivos.
        
    Else
        Codigo = Mid(Codigo, 2)
        
        
        campo = " {schpre.motivoTraspaso} IN [ " & Codigo & "]"
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        campo = " schpre.motivoTraspaso IN (" & Codigo & ")"
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        
        If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & """ + chr(13) + """
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Motivo:" & NomTabla
        
        
        
    End If
    
    NumeroDeCopias = 1
    NomTabla = ""
    
    
    
    
    campo = "schpre INNER JOIN slhpre ON schpre.numofert=slhpre.numofert and schpre.fecofert=slhpre.fecofert "
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    
    CadenaDesdeOtroForm = "|pDH=""" & CadenaDesdeOtroForm & """|"
    cadParam = cadParam & CadenaDesdeOtroForm
    numParam = numParam + 2
    nomRPT = "rListaOfertasHco.rpt"
    CadenaDesdeOtroForm = ""
    
    Titulo = "Listado hist�rico de ofertas"
    LlamarImprimir True, False
    
End Sub

Private Sub cmdPararCRM_Click()
    'Paramos el proceso de impresionde CRM
    If Not FrameCRMProgess.visible Then Exit Sub
    If MsgBox("�Desea parar el proceso?", vbQuestion + vbYesNo) = vbYes Then indCodigo = 1 'Cancela el preoceso de impresion de crm
    
End Sub

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub









Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 31, 35 '31: Informe Ofertas
                        '35: Informe Historico Ofertas
                PonerFoco txtCodigo(1)
            Case 32, 33 '32: Recordatorio de Oferta
                        '33: Informe Valoracion de Oferta
                PonerFoco txtCodigo(5)
                
                If vParamAplic.NumeroInstalacion = 2 Then optPrecioUC.Value = True
                
            Case 34, 92 '34: Informe Ofertas Efectuadas
                        '92: Informe Gastos t�cnicos
                PonerFoco txtCodigo(16)
            Case 36 '36: Traspaso Ofertas a Historico
                If txtCodigo(43).Text <> "" Then
                    PonerFocoCbo cboMotTras
                Else
                    PonerFoco txtCodigo(43)
                End If
            Case 37 '37: Generar Pedido de OFerta
                PonerFoco txtCodigo(24)
            Case 40 '40: Carta Confirmacion de Pedido
                PonerFoco txtCodigo(77)
            Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '48: Informe de Altas de Nuevos Clientes
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
                PonerFoco txtCodigo(27)
            Case 47 '47: Informe de Clientes
                PonerFoco txtCodigo(33)
            Case 38, 239, 55, 56 '55: Informe de Pedido de Compras (proveedor)
            
                If OpcionListado = 38 Then
                    If vParamAplic.NumeroInstalacion = vbFenollar Then
                        
                        
                        'En fenollar el peso del albaran
                        Codigo = "sliped.codartic=sartic.codartic and numpedcl"
                        Codigo = DevuelveDesdeBD(conAri, " sum(cantidad*(coalesce(pesoarti,0)))", "sliped,sartic", Codigo, NumCod)
                                    
                        chkVarios(0).visible = True
                        chkVarios(0).Value = False
                        Label12(9).visible = True
                        Label12(8).visible = True
                        Label12(9).Caption = Codigo & " Kg"
                    End If
                End If
                
            
                PonerFoco txtCodigo(73)
            Case 57 '57: Pasar Pedido a Albaran de Compras(Proveedores)
                PonerFoco txtCodigo(47)
            Case 80, 81 '80: Pasar albaranes al historico (ventas clientes)
                            '81: Pasar pedidos al historico (ventas clientes)
                PonerFoco txtCodigo(50)
            
            Case 225 'Datos para Factura Rectificativa
                PonerFoco txtCodigo(71)
            Case 226 'Datos para Reimprimir Facturas
                PonerFocoCbo Me.cboTipomov(1)
                
            Case 230 'Listado Ventas por Familia
                chkDatosAlbaranes_Click 0
                PonerFoco txtCodigo(96)
                
            ' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
            Case 238 'Confirmacion entrega Pedido
                PonerFoco txtCodigo(116)
            ' ----
                
            Case 240 'Inf. Cierre caja TPV
                PonerFoco txtCodigo(88)
                
            Case 305, 306 '305: Listado Etiquetas proveedor
                          '306: Listado Cartas a proveedores
                PonerFoco txtCodigo(58)
            Case 307, 308 '307: List. Pendiente de Recibir (COMPRAS)
                          '308: List. Pendiente de Facturar (COMPRAS)
                PonerFoco txtCodigo(65)
                
            Case 310, 311, 312 'Listado Compras por Proveedor/Familia/Articulo
                                '312: Listado albaranes por proveedor
                PonerFoco txtCodigo(90)
            Case 315, 316
              
                PonerFoco txtCodigo(110)
            Case 400, 401, 402
              
                PonerFoco txtCodigo(131)
            Case 406
                PonerFoco txtCodigo(139)
            Case 407
                'Imprimir el PEDIDO
                cmdAceptarPedCom_Click
                Unload Me
                
            Case 409
                'Oferas hco
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim devuelve As String
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    indCodigo = 0
    NomTabla = ""
    CargaIconosAyuda
    CargaIconosAyuda2
    
    'Ocultar todos los Frames de Formulario
    Me.FrameOfertas.visible = False
    Me.FrameRecordatorio.visible = False
    Me.FrameEfectuadas.visible = False
    Me.FrameTraspasoHco.visible = False
    Me.FrameGenPedido.visible = False
    Me.FrameClienInactivos.visible = False
    Me.FrameClientes2.visible = False
    Me.FrameGenAlbCom.visible = False
    Me.FramePasarHco.visible = False
    Me.FrameEtiqProv.visible = False
    Me.FramePteRecibir.visible = False
    Me.FrameFacRectif.visible = False
    Me.FrameFacReimprimir.visible = False
    Me.FramePedidos.visible = False
    ' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
    Me.FramePedConfirma.visible = False
    ' ----
    Me.FrameConfirmPed.visible = False
    Me.FrameCierreCaja.visible = False
    Me.FrameCompras.visible = False
    Me.FrameEstVentasFam.visible = False
    FrameClientesPotenciales.visible = False
    FrameEnvioFacMail.visible = False
    FrameCRM.visible = False
    
    FrameComprobarCtaBancoSecciones.visible = False
    FrameOfertasHcoInforme.visible = False
    
    CommitConexion
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 31, 35 '31: Informe de Ofertas
                    '35: Informe Historico de Ofertas
                    
            W = 5475
            If InstalacionEsEulerTaxco Then W = 10275
            H = 5655
            cmdCancel(0).Left = W - 1635
            Me.cmdAceptarOfer.Left = cmdCancel(0).Left - 1080
            PonerFrameVisible Me.FrameOfertas, True, H, W
            'Situo el cancelar
            
            Me.OptPapelBlanco.Value = True
            indFrame = 0
            If NumCod <> "" Then txtCodigo(1).Text = NumCod
            If OpcionListado = 35 Then Me.Label5.Caption = "Informe de Ofertas (Hist�rico)"
            
            
            If InstalacionEsEulerTaxco Then
                cargaDocumentos
                chkVarios(12).visible = False  'IVA incluido
            Else
                chkVarios(12).visible = OpcionListado = 31
            End If
            
        Case 32, 33 '32: Recordatorio de Ofertas
                    '33:Informe Valoraci�n de Ofertas
            PonerFrameRecordaVisible True, H, W
            indFrame = 1
            If codClien <> "" Then
                txtCodigo(9).Text = codClien
                txtCodigo(10).Text = codClien
                devuelve = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", codClien, "N")
                txtNombre(9).Text = devuelve
                txtNombre(10).Text = devuelve
            End If
            If NumCod <> "" Then
                txtCodigo(5).Text = NumCod
                txtCodigo(6).Text = NumCod
            End If
            
        Case 34, 92 '34: Informe Ofertas Efectuadas
                    '92: Informe Gastos T�cnicos
            W = FrameEfectuadas.Width
            H = FrameEfectuadas.Height
            PonerFrameVisible Me.FrameEfectuadas, True, H, W
            
            If OpcionListado = 92 Then
                Label1.Caption = "Gastos T�cnicos"
                Label4(4).Caption = "T�cnico"
                
                Label4(145).visible = False
                Label4(144).visible = False
                imgFecha(40).visible = False
                txtCodigo(166).visible = False
                Label4(146).visible = False
                imgFecha(41).visible = False
                txtCodigo(167).visible = False
            
            
            
            Else
                If InstalacionEsEulerTaxco Then
                    Label4(145).Caption = "Fecha reclamaci�n"
                Else
                    Label4(145).Caption = "Fecha entrega"
                End If
            End If
            Me.chkPendientes.visible = (OpcionListado = 34)
            chkBusHcoOfer.visible = (OpcionListado = 34)
            indFrame = 2
            
        Case 36 '36: Traspaso a Historico (IMPRIME LISTADO Y PREGUNTA SI TRASPASO A HCO)
            W = FrameTraspasoHco.Width
            H = FrameTraspasoHco.Height
            PonerFrameVisible Me.FrameTraspasoHco, True, H, W
            indFrame = 3
            Me.Caption = "Ofertas"
             CargarCombo_Tabla Me.cboMotTras, "smotitrasofer", "codmotiv", "descmotivo", , False, "descmotivo"
            If CadenaDesdeOtroForm <> "" Then
                txtCodigo(43).Text = RecuperaValor(CadenaDesdeOtroForm, 1) 'clien
                txtCodigo(44).Text = txtCodigo(43).Text
                txtNombre(43).Text = RecuperaValor(CadenaDesdeOtroForm, 2) 'nombre
                txtNombre(44).Text = txtNombre(43).Text
                txtCodigo(22).Text = RecuperaValor(CadenaDesdeOtroForm, 3)  'fech
                txtCodigo(23).Text = txtCodigo(22).Text
                txtCodigo(20).Text = RecuperaValor(CadenaDesdeOtroForm, 4)  'fech
                txtCodigo(21).Text = txtCodigo(20).Text
            End If
            
            
        Case 37 '37: Pedir datos para pasar Oferta a Pedido (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenPedido, True, H, W
            indFrame = 4
            Me.Caption = "Generar Pedido"
            txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(26).Text = Format(FecEntre, "dd/mm/yyyy")
            txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
        
            
            Me.txtCodigo(24).Text = PonerTrabajadorConectado(Codigo)
            Me.txtNombre(24).Text = Codigo
            
            If vParamAplic.NumeroInstalacion = vbHerbelca Then
                imgBuscarOfer(8).Enabled = vUsu.Nivel = 0
                BloquearTxt txtCodigo(24), vUsu.Nivel > 0
            End If
        
         Case 40 '40: Cartas Confirmacion de Pedidos
            W = 7035
            H = 6255
            PonerFrameVisible Me.FrameConfirmPed, True, H, W
            Me.OptPapelBlanco3.Value = True
            indFrame = 13 'solo para el boton cancelar
            txtCodigo(82).Text = Format(Now, "dd/mm/yy")
            NomTabla = "scaped"
            NomTablaLin = "sliped"
            FrameTipoPapel3.visible = False
        
        Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
            PonerFrameClienInacVisible True, H, W
            indFrame = 5
            chkEnviaCorreo.visible = OpcionListado = 90
            chkEtiqDpto.visible = OpcionListado = 90
            If OpcionListado = 90 Then
                CargarComboTipoMov 2
                'FrameImpClien.visible = False  'Tb hay que ponerlo para etiquetas de clientes
            End If
            
        Case 47 '47: Informe de Clientes
            W = FrameClientes2.Width + 320
            H = FrameClientes2.Height + 320
            PonerFrameVisible Me.FrameClientes2, True, H, W
            CargarListViewOrden
            indFrame = 6
            'Viloumen de ventas
            FrameVolumen.visible = False
            Me.chkVolumen.Value = 0
            'fijo el a�o actual
            txtCodigo(122).Text = "01/01/" & Year(Now)
            txtCodigo(123).Text = Format(Now, "dd/mm/yyyy")
            Label4(51).Caption = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociaci�n", "Ruta")
        Case 38, 239, 55, 56, 407
                '38: Pedidos Venta
                '239: Hco Pedidos venta
                '55: Informe de Pedido de Compras (Proveedor)
                '56: Informe de Hist. Pedido de Compras (Proveedor)
                '407: Pedidos proveedor SIN vistaprevia
            PonerFramePedVisible H, W
            indFrame = 12
            If NumCod <> "" Then txtCodigo(73).Text = NumCod
            
            
            
        Case 57 '57: Pedir datos para pasar de Pedido a Albaran proveedor (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenAlbCom, True, H, W
            indFrame = 7
            Me.Caption = "Generar Albaran Compras"
            'Poner el trabajador conectado
            Me.txtCodigo(47).Text = PonerTrabajadorConectado(devuelve)
            Me.txtNombre(47).Text = devuelve
            Me.txtCodigo(49).Text = Format(Now, "dd/mm/yyyy")
            Me.txtCodigo(168).Text = txtCodigo(49).Text
        
            'chkImprAlbProv(0).Value = 0
            chkImprAlbProv(0).Value = 0
            If vParamAplic.NumeroInstalacion = 2 Then chkImprAlbProv(0).Value = 1 'Para herbleca esta marcado
        
            txtCodigo(168).ToolTipText = "Fecha entrada mercancia almacen"
            Me.Label4(147).ToolTipText = txtCodigo(168).ToolTipText
            
        Case 80, 81 '80: pasar albaranes al historico (ventas)
                        '81: pasar pedidos al historico (ventas)
            H = 4575
            W = 6920
            PonerFrameVisible Me.FramePasarHco, True, H, W
            indFrame = 8
            Me.Caption = "Eliminar"
            Select Case OpcionListado
                Case 80, 82: Me.Label3(4).Caption = "Pasar Albaran al hist�rico"
                Case 81: Me.Label3(4).Caption = "Pasar Pedido al hist�rico"
            End Select
            Me.txtCodigo(50).Text = Format(Now, "dd/mm/yyyy")
            Me.txtCodigo(51).Text = PonerTrabajadorConectado(devuelve)
            Me.txtNombre(51).Text = devuelve
            
        Case 225 'Factura rectificativa
            H = 4420
            W = 5740
            PonerFrameVisible Me.FrameFacRectif, True, H, W
            indFrame = 11
            Me.Caption = "Facturas rectificativas"
            CargarComboTipoMov (0)
'            Me.cboTipomov(0).ListIndex = 2
            Line2.BorderStyle = 0 'la linea
            Label4(133).Caption = ""
        Case 226 'Reimprimir Factura
            H = FrameFacReimprimir.Height
            W = 6555
            PonerFrameVisible Me.FrameFacReimprimir, True, H, W
            indFrame = 14
            CargarComboTipoMov (1)
            
     
            'PARA LAS FAT
            chk_duplicado2(2).Value = 1
            chk_duplicado2(2).visible = False
            
            cadFormula = DevuelveDesdeBDNew(conAri, "scryst", "nomcryst", "codcryst", "18", "N")
            Me.chkFormatoTPV.Value = 0
            If cadFormula = "" Then
                'NO SE HA ENCONTRADOR
                Me.chkFormatoTPV.Enabled = False
                cadFormula = "Formato NO encontrado"
            End If
            Me.chkFormatoTPV.Caption = cadFormula
            Me.txtCodigo(119).Text = vParamAplic.NumCopiasFacturacion
'            CargarComboTipoMov (2)
            
        Case 230, 231 '230: Estadistica ventas por familia
                      '231: Detalle facturacion clientes
            indFrame = 17
            
            If OpcionListado = 231 Then
                H = 7365
               ' H = 7000 '4325
                'FrameDetalleFacturacion.Left = 390
                'FrameDetalleFacturacion.Top = 3000
                FrameDetalleFacturacion.BorderStyle = 0
                Me.cmdAceptarEstVentas.Top = 6800
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarEstVentas.Top
                Me.Label9(31).Caption = "Detalle Facturaci�n Clientes"
                CargaTipoMov
                txtCodigo(98).Text = Format(vEmpresa.FechaIni, "dd/mm/yyyy")
                Label9(56).Caption = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociaci�n", "Ruta")
            Else
                H = 7845
            End If
            W = 7035
            Me.Frame12.visible = (OpcionListado = 230)
            Frame10.visible = (OpcionListado = 230)
            
            
            chkDatosAlbaranes(8).visible = (OpcionListado = 231)
            chkDatosAlbaranes(10).visible = (OpcionListado <> 231)
            FrameDetalleFacturacion.visible = OpcionListado = 231
            
            PonerFrameVisible Me.FrameEstVentasFam, True, H, W
            
            
            chkDatosAlbaranes(3).visible = (OpcionListado = 230)
            chkDatosAlbaranes(10).Value = 0
            
            
            
            
            chkDatosAlbaranes(4).Value = 0
            If vParamAplic.AlmacenB = 99 Then chkDatosAlbaranes(4).Value = 1
            
        ' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
        Case 238 'Confirmacion entrega pedido
            W = 6315
            H = 4095
            PonerFrameVisible Me.FramePedConfirma, True, H, W
            indFrame = 19
            Me.Caption = "Confirmaci�n entrega Pedido"
            If NumCod <> "" Then txtCodigo(114).Text = NumCod
            txtCodigo(115).Text = Format(FecEntre, "dd/mm/yyyy")
            BloquearTxt txtCodigo(114), True
            BloquearTxt txtCodigo(115), True
            
'            NomTabla = "scaped"
'            NomTablaLin = "sliped"
        ' ----
        
        Case 240 'Inf. cierre caja TPV
            H = 3800
            W = 6300
            PonerFrameVisible Me.FrameCierreCaja, True, H, W
            indFrame = 15
'            CargarComboTipoPago
'            Combo1.ListIndex = 0
            'Mostrar la fecha de hoy
            txtCodigo(88).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(89).Text = Format(Now, "dd/mm/yyyy")
            
        
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = FrameEtiqProv.Height + 320
            W = FrameEtiqProv.Width + 240
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (OpcionListado = 306)
            If (OpcionListado = 306) Then Me.Label9(1).Caption = "Cartas a Proveedores"
            
        Case 307, 308 '307: List. Material Pendiente de recibir (COMPRAS)
                      '308: List. Albaranes ptes de facturar (COMPRAS)
            indFrame = 10
            chkVarios(11).visible = True
            If OpcionListado = 307 Then
                Me.Label9(19).Caption = "Material pendiente de recibir"
                chkVarios(11).Top = 4840
                H = 5200
            Else
                
             chkVarios(11).Top = 3840
                Me.Label9(19).Caption = "Albaranes pendiente de factura"
                H = 4200
                Me.cmdAceptarPte.Top = 3500
                Me.cmdCancel(10).Top = Me.cmdAceptarPte.Top
                Me.cboPdteFacturarProv.ListIndex = 0
            End If
            W = 7035
            PonerFrameVisible Me.FramePteRecibir, True, H, W
            Me.Frame6.visible = (OpcionListado = 307)
            Me.Frame7.visible = (OpcionListado = 307)
            
            chkVarios(1).visible = OpcionListado = 308
            Label4(129).visible = OpcionListado = 308
            Me.cboPdteFacturarProv.visible = OpcionListado = 308
            
        Case 310, 311, 312 '310: Listado COMPRAS por proveedor
                            'compras familia /articulo
                            '312: Listado albaranes por proveedor
            indFrame = 16
            H = 5635
            chkDatosAlbaranes(7).Top = 3960
            If OpcionListado = 310 Or OpcionListado = 312 Then
                H = 4325
                Me.cmdAceptarCompras.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarCompras.Top
                If OpcionListado = 312 Then
                    Me.Label9(21).Caption = "Albaranes por Proveedor"
                    chkDatosAlbaranes(7).Top = cmdAceptarCompras.Top - 320
                Else
                    'opcion 310
                    Me.Label9(21).Caption = "Compras por Proveedor"
                End If
                Me.Label4(87).Caption = "Fecha albaran"
            End If
            W = 7035
            Label9(38).Caption = ""
            PonerFrameVisible Me.FrameCompras, True, H, W
            Me.Frame8.visible = (OpcionListado = 311)
            Me.Frame9.visible = (OpcionListado = 311)
            chkVarios(9).visible = (OpcionListado = 311)  'ordenado por nomprove
            FrameMinImporte.visible = (OpcionListado = 311)
            FrameMinImporte.BorderStyle = 0
            chkDatosAlbaranes(1).visible = (OpcionListado = 311)
            chkDatosAlbaranes(11).visible = (OpcionListado = 311)
            chkDatosAlbaranes(7).visible = OpcionListado <> 310
            If OpcionListado <> 310 Then
                'El chk es visible
                'Para el listado de albaranes NO lo marco
                If OpcionListado = 312 Then chkDatosAlbaranes(7).Value = 0
            End If
        Case 315, 316
            
            FrameEnvioMail2.ZOrder 0     'para traer al ppio el frame
            
            indFrame = 18
            If OpcionListado = 316 Then Me.FrameEnvioFacMail.Width = 5735 '5535
            lblInd.Caption = ""
            Label14(22).Caption = ""
            H = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, H, W
            CargarComboTipoMov 1000
            'Mayo 2016. Solo era para las traspasadas
            chkMail(1).visible = OpcionListado = 316 'Solo para facturae
            chkMail(3).visible = OpcionListado = 315
            If OpcionListado = 316 Then
                cmdEnvioMail.Left = 3240
                cmdCancel(indFrame).Left = 4320
                Label14(16).Caption = "Facturacion E"
                cmdEnvioMail.TabIndex = 474
            Else
                Label14(16).Caption = "Envio facturas por mail"
                
            End If
        Case 400, 401, 402
            
            
            If OpcionListado = 402 Then
                Label8(1).Caption = "Listado CRM potenciales"
                
                Me.txtCodigo(131).Text = RecuperaValor(NumCod, 1)
                Me.txtNombre(131).Text = RecuperaValor(NumCod, 2)
                Me.txtCodigo(132).Text = Me.txtCodigo(131).Text
                Me.txtNombre(132).Text = Me.txtNombre(131).Text
            Else
                If OpcionListado = 400 Then
                    cadParam = "etiquetas"
                Else
                    cadParam = "cartas"
                End If
                Label8(1).Caption = "Clientes potenciales (" & cadParam & ")"
                cadParam = ""
            End If
                
            FrameCartaPot.BorderStyle = 0
            FrameCartaPot.visible = OpcionListado = 401
            
            H = FrameClientesPotenciales.Height
            W = FrameClientesPotenciales.Width
            PonerFrameVisible FrameClientesPotenciales, True, H, W
            indFrame = 20
            
    Case 406
            indFrame = OpcionListado
            H = FrameCRM.Height
            W = FrameCRM.Width
            PonerFrameVisible FrameCRM, True, H, W
            
            'Hay un listado que se pondra siempre por encima del CRM
            'que sera el progress
            FrameCRMProgess.Top = 600
            FrameCRMProgess.Left = 240
    Case 408
            
            indFrame = OpcionListado
            H = FrameComprobarCtaBancoSecciones.Height
            W = FrameComprobarCtaBancoSecciones.Width
            PonerFrameVisible FrameComprobarCtaBancoSecciones, True, H, W
            Label9(52).Caption = ""
            
            

            chkVarios(8).visible = vParamAplic.ComprobarBancoRestoAplicaciones
            chkVarios(8).Value = Abs(vParamAplic.ComprobarBancoRestoAplicaciones)
            
            
           ' chkVarios(8).Value = 0
            'chkVarios(8).visible = False
            
    Case 409
            H = FrameOfertasHcoInforme.Height
            W = FrameOfertasHcoInforme.Width
            PonerFrameVisible FrameOfertasHcoInforme, True, H, W
            CargaMotivosBaja
            indFrame = 21
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    'Poner la tabla de Ofertas o la del Historico de Ofertas
    If NomTabla = "" Then
        If OpcionListado = 35 Then
            NomTabla = "schpre" 'Historico
        Else
            NomTabla = "scapre"
        End If
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If OpcionListado = 406 Then
        'Esta imprimiendo los CRM
        If FrameCRMProgess.visible Then Cancel = 1
    
    Else
        If OpcionListado = 315 Or OpcionListado = 316 Then
            If Me.ProgressBar1.visible Then Cancel = 1
        End If
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cadFormula = CadenaDevuelta
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de cod Postal
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If OpcionListado = 305 Or OpcionListado = 306 Then 'Proveedores
            cadFormula = "{sprove.codprove} IN [" & CadenaSeleccion & "]"
            cadSelect = "sprove.codprove IN (" & CadenaSeleccion & ")"
        Else 'clientes
            cadFormula = "{sclien.codclien} IN [" & CadenaSeleccion & "]"
            cadSelect = "sclien.codclien IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub

Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Actividades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArtic_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProve_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoRuta_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Rutas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSitua_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoZona_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Zonas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim Ayuda As String

    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Codigo = vbCrLf & Space(10)
    Select Case Index
    Case 0
        Ayuda = ""
        Ayuda = Ayuda & vbCrLf & " --> VOLUMEN "
        Ayuda = Ayuda & Codigo & "Agrupado por agente muestra ademas de los datos basicos"
        Ayuda = Ayuda & Codigo & "el volumen de ventas entre las fechas seleccionadas "
        Ayuda = Ayuda & Codigo & "y el credito que tenga. Si marca agrupado no 'salta'"
        Ayuda = Ayuda & Codigo & "por zona,ruta."
        Ayuda = Ayuda & Codigo & "  -Telefonos / mail / Forma de pago:  Muestra en listado, o telefonos o email o la forma de pago(en el agrupado) "
        Ayuda = Ayuda & Codigo & "  -Formato exportacion: facilita exportaci�n excel"
        
        Ayuda = Ayuda & vbCrLf & vbCrLf & " --> Poblacion / actividad "
        'Ordenado por codpobla,activadad. Solo rompe por codpostal
        Ayuda = Ayuda & Codigo & "Agrupado por codigo postal, muestra los datos basicos y la actividad"
        
        
        Ayuda = Ayuda & vbCrLf & vbCrLf & " --> Normal "
        Ayuda = Ayuda & Codigo & "Ordenado segun la seleccion mostrara los datos basicos:"
        Ayuda = Ayuda & Codigo & "ruta,zona,agente,codigo,nombre,domicilio,nif,telefono"
        
    Case 1
        Ayuda = ""
        Ayuda = Ayuda & vbCrLf & " Si marca la opcion de """ & chkFraRectCompleta.Caption & """ se generar�"
        Ayuda = Ayuda & " una factura FRT igual que la indicada pero con importes en negativo." & vbCrLf & vbCrLf
        Ayuda = Ayuda & " Si no lo marca crear� el albaran para seleccionar posteriormente las lineas a rectificar(pudiendo seleccionar todas)"
    
    
    End Select
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(46, "=") & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 39, 40, 45, 60, 77 'Cod. Carta
            Select Case Index
                Case 0: indCodigo = 2
                Case 1: indCodigo = 13
                Case 39: indCodigo = 63
                Case 40: indCodigo = 64
                Case 45: indCodigo = 81
                Case 60: indCodigo = 116
                Case 77: indCodigo = 138
            End Select
            
            Set frmMtoCartasOfe = New frmFacCartasOferta
            frmMtoCartasOfe.DatosADevolverBusqueda = "0|1|"
            frmMtoCartasOfe.Show vbModal
            Set frmMtoCartasOfe = Nothing
            
        Case 2, 3, 9, 10, 23, 24, 46, 47, 52, 53, 56, 57, 63, 64, 78, 79, 85, 86, 93, 94, 99, 100, 147, 148 'Cod. CLIENTE
            Select Case Index
                Case 2, 3: indCodigo = 7 + Index
                Case 9, 10: indCodigo = 18 + Index
                Case 23, 24: indCodigo = Index + 20
                Case 46, 47: indCodigo = Index + 33
                Case 52, 53: indCodigo = Index + 44
                Case 56, 57: indCodigo = Index + 54
                Case 63, 64: indCodigo = Index + 57
                Case 78, 79: indCodigo = Index + 61
                Case 85, 86, 93, 94: indCodigo = Index + 62
                Case 99, 100: indCodigo = Index + 63
            End Select
            
            
            Set frmMtoCliente = New frmBasico2
            AyudaClientes frmMtoCliente, txtCodigo(indCodigo).Text
            Set frmMtoCliente = Nothing
    
            
            
        Case 4, 5, 6, 7, 11, 12, 19, 20, 25, 26, 80, 81, 87, 88, 89, 90 'Cod. AGENTE
            Select Case Index
                Case 4, 5: indCodigo = 7 + Index
                Case 5: indCodigo = 12
                Case 6, 7: indCodigo = 12 + Index
                Case 11, 12: indCodigo = 18 + Index
                Case 19, 20, 25, 26: indCodigo = 20 + Index
                Case 80, 81: indCodigo = Index + 61
                Case 87, 88, 89, 90: indCodigo = Index + 62
            End Select
            If OpcionListado <> 92 Then
'                Set frmMtoAgente = New frmFacAgentesCom
'                frmMtoAgente.DatosADevolverBusqueda = "0|1|"
'                frmMtoAgente.Show vbModal
                Set frmMtoAgente = New frmBasico2
                AyudaAgentesComerciales frmMtoAgente, txtCodigo(indCodigo), , True
                Set frmMtoAgente = Nothing
                    
            ElseIf Index = 6 Or Index = 7 Then 'Gastos financieros (trabajador)
'                Set frmMtoTraba = New frmAdmTrabajadores
'                frmMtoTraba.DatosADevolverBusqueda = "0|1|"
'                frmMtoTraba.Show vbModal
                Set frmMtoTraba = New frmBasico2
                AyudaTrabajadores frmMtoTraba, Text1(indCodigo)
                Set frmMtoTraba = Nothing
            End If
            
        Case 8, 27, 28, 61, 62 'cod. TRABAJADOR
            indCodigo = 24
            If Index = 27 Then
                indCodigo = 47
            ElseIf Index = 28 Then indCodigo = 51
            ElseIf Index > 28 Then indCodigo = (117 + 61) - Index
            End If
'            Set frmMtoTraba = New frmAdmTrabajadores
'            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
'            frmMtoTraba.Show vbModal
            Set frmMtoTraba = New frmBasico2
            AyudaTrabajadores frmMtoTraba, Text1(indCodigo)
            Set frmMtoTraba = Nothing
            
        Case 13, 14, 30, 31, 67, 68, 75, 76, 82, 83 'cod. ACTIVIDAD
            indCodigo = 20 + Index
            If Index = 30 Or Index = 31 Then indCodigo = Index + 23
            'If Index = 67 Or Index = 68 Then indCodigo = Index + 60
            If Index >= 67 Then indCodigo = Index + 61
            
            
            Set frmMtoActiv = New frmFacActividades
            frmMtoActiv.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoActiv.Show vbModal
            Set frmMtoActiv = Nothing
            
        Case 15, 16, 91, 92 'cod. ZONA
            If Index < 91 Then
                indCodigo = 20 + Index
            Else
                indCodigo = 62 + Index
           
            End If
            Set frmMtoZona = New frmFacZonas
            frmMtoZona.DatosADevolverBusqueda = "0|1|"
            frmMtoZona.Show vbModal
            Set frmMtoZona = Nothing
            
         Case 17, 18, 95, 96 'cod. RUTA
            If Index < 95 Then
                indCodigo = 20 + Index
            Else
                indCodigo = 62 + Index
            End If
            Set frmMtoRuta = New frmFacRutas
            frmMtoRuta.DatosADevolverBusqueda = "0|1|"
            frmMtoRuta.Show vbModal
            Set frmMtoRuta = Nothing
            
        Case 21, 22, 34, 84 'cod. SITUACION
            indCodigo = 20 + Index
            If Index = 34 Then indCodigo = Index + 23
            If Index = 84 Then indCodigo = Index + 61
            Set frmMtoSitua = New frmFacSituaciones
            frmMtoSitua.DatosADevolverBusqueda = "0|1|"
            frmMtoSitua.Show vbModal
            Set frmMtoSitua = Nothing
            
        Case 29 'INCIDENCIAS
            indCodigo = 52
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            txtCodigo(indCodigo).Text = ""
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 32, 33, 37, 38, 69, 70, 73, 74 'Cod POSTAL
            If Index < 69 Then
                indCodigo = Index + 23
            Else
                indCodigo = Index + 60
            End If
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0|1|"
            txtCodigo(indCodigo).Text = ""
            frmCP.Show vbModal
            Set frmCP = Nothing
            
        Case 35, 36, 41, 42, 48, 49, 65, 66 'cod. PROVEEDOR
            Select Case Index
                Case 35, 36: indCodigo = Index + 23
                Case 41, 42: indCodigo = Index + 24
                Case 48, 49: indCodigo = Index + 42
                Case 65, 66: indCodigo = Index + 59
            End Select
'            If Index = 35 Or Index = 36 Then indCodigo = Index + 23
'            If Index = 41 Or Index = 42 Then indCodigo = Index + 24
'            If Index = 48 Or Index = 49 Then indCodigo = Index + 42
            Set frmMtoProve = New frmBasico2
'            frmMtoProve.DatosADevolverBusqueda = "0|1|"
'            frmMtoProve.Show vbModal
            AyudaProveedores frmMtoProve, txtCodigo(indCodigo)
            Set frmMtoProve = Nothing
            
        Case 43, 44, 58, 59 'cod. ARTICULO
            If Index <= 44 Then
                indCodigo = Index + 24
            Else
                indCodigo = Index + 54  'En listado de vetnas x familia articulo
            End If
            Set frmMtoArtic = New frmBasico2
            'frmMtoArtic.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
'            frmMtoArtic.DatosADevolverBusqueda = "0|1|"
'            frmMtoArtic.Show vbModal
            AyudaArticulos frmMtoArtic, txtCodigo(indCodigo)
            Set frmMtoArtic = Nothing
            
        Case 50, 51, 54, 55 'Cod. FAMILIA articulo
            Select Case Index
                Case 50, 51: indCodigo = Index + 44
                Case 54, 55: indCodigo = Index + 46
            End Select
'            Set frmMtoFamilia = New frmAlmFamiliaArticulo
'            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
'            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = New frmBasico2
            AyudaFamilias frmMtoFamilia, txtCodigo(indCodigo)
            Set frmMtoFamilia = Nothing
        Case 71, 72
            'Clientes potenciales
            AbrirBuscaGrid Index
            
        Case 97, 98
            'Bancos propios   - Formas pago
            AbrirBuscaGrid Index

            
        End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer
    If Index < 2 Then
        For i = 1 To lwCargos.ListItems.Count
            lwCargos.ListItems(i).Checked = Index = 1
        Next
    ElseIf Index < 4 Then
        'Fra electronica.  Esun listbox: empieza en cero
        For i = 0 To ListTipoMov(1000).ListCount - 1
            ListTipoMov(1000).Selected(i) = Index = 3
        Next
        chkMail(2).visible = False
    ElseIf Index < 6 Then
        For i = 1 To Me.lwFact.ListItems.Count
            Me.lwFact.ListItems(i).Checked = Index = 5
        Next
    ElseIf Index < 8 Then
        For i = 1 To Me.lwTippa.ListItems.Count
            Me.lwTippa.ListItems(i).Checked = Index = 6
        Next
    ElseIf Index < 10 Then
        For i = 1 To Me.lwMotBaja.ListItems.Count
            Me.lwMotBaja.ListItems(i).Checked = Index = 8
        Next
    
    End If
End Sub

Private Sub imgClearCmbTipomov_Click()
    cboTipomov(2).ListIndex = -1
End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            indCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            indCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            indCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            indCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            indCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            indCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            indCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            indCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            indCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            indCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            indCodigo = 49
        Case 14
            indCodigo = 50
        Case 15, 16
            indCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            indCodigo = 72
        Case 18, 19 'Ped. Compras
            indCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            indCodigo = Index + 57
        Case 22: indCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            indCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            indCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            indCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            indCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            indCodigo = Index + 73
        Case 33, 34
            indCodigo = Index + 75
            
        Case 35, 36
            indCodigo = Index + 87
        Case 37
            indCodigo = 160
        Case 38, 39, 40, 41
            indCodigo = 126 + Index
        Case 42
            indCodigo = 168
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub












Private Sub ListTipoMov_ItemCheck(Index As Integer, Item As Integer)

    '
    If Index = 1000 Then
        If PrimeraVez Then Exit Sub
        Titulo = ""
        For NumRegElim = 0 To ListTipoMov(1000).ListCount - 1
            If ListTipoMov(1000).Selected(NumRegElim) Then Titulo = Titulo & Mid(ListTipoMov(1000).List(NumRegElim), 1, 3)
        Next
        If vParamAplic.TieneTelefonia2 > 2 Then Me.chkMail(2).visible = Titulo = "FAT"
    End If
    
    
End Sub

Private Sub ListTipoMov_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1000 And OpcionListado = 316 Then
        'PonerFocoBtn cmdEnvioMail
        KEYpress KeyAscii
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub OptCompras_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub optDetalleFacturacion_Click(Index As Integer)
    If optDetalleFacturacion(1).Value Then chkDatosAlbaranes(8).Value = 1
End Sub

Private Sub optDetalleFacturacion_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    
End Sub

Private Sub optEnvioMail_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optForpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub OptPorFamilia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 1 Then KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 147: KEYBusqueda KeyAscii, 85 'cliente desde
            Case 148: KEYBusqueda KeyAscii, 86 'cliente hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 3, 4, 7, 8, 16, 17, 22, 23, 25, 26, 31, 32, 49, 50, 69, 70, 72, 74, 75, 77, 78, 82, 85, 86, 88, 89, 92, 93, 98, 99, 104, 105, 108, 109, 122, 123, 160, 164, 165, 166, 167, 168
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
            'Fecha entrega para Pedido. Poner la semana
            If Index = 26 Then
                'Comprobar que fecha entrega es posterior a la del pedido
                If Not EsFechaIgualPosterior(txtCodigo(25).Text, txtCodigo(26).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                Else
                    If Not IsDate(txtCodigo(26).Text) Then
                        txtNombre(4).Text = ""
                    Else
                        txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
                    End If
                End If
            End If
            
        Case 1, 5, 6, 20, 21, 71, 83, 84, 119 'N� de OFERTA/FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            Else
                If Index = 1 Then txtCodigo(Index).Text = ""
            End If
            If Index = 1 Then
                If InstalacionEsEulerTaxco Then cargaDocumentos
            End If
        Case 2, 13, 63, 64, 81, 116, 138 'CARTA de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 9, 10, 27, 28, 43, 44, 79, 80, 96, 97, 110, 111, 120, 121, 139, 140, 147, 148, 155, 156, 162, 163 'Cod. CLIENTE
            EsNomCod = True
            tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"

        Case 11, 12, 18, 19, 29, 30, 39, 40, 45, 46, 80, 81, 141, 142, 149, 150, 151, 152 'Cod. AGENTE
            EsNomCod = True
            Formato = "0000"
            If OpcionListado = 92 Then 'Gastos tecnicos
                If Index = 18 Or Index = 19 Then
                    'cod agente / cod. trabajador
                    tabla = "straba"
                    codCampo = "codtraba"
                    NomCampo = "nomtraba"
                    Titulo = "Trabajador"
                End If
            Else
                tabla = "sagent"
                codCampo = "codagent"
                NomCampo = "nomagent"
                Titulo = "Agente"
            End If
        
        Case 24, 47, 51, 117, 118 'Cod. TRABAJADOR
            EsNomCod = True
            tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            Formato = "0000"
            Titulo = "Trabajador"
            
        Case 33, 34, 53, 54, 127, 128, 136, 137, 143, 144 'Cod ACTIVIDAD
            EsNomCod = True
            tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            Formato = "000"
            Titulo = "Actividad de Cliente"
            
        Case 35, 36, 153, 154 'cod ZONA
            EsNomCod = True
            tabla = "szonas"
            codCampo = "codzonas"
            NomCampo = "nomzonas"
            Formato = "000"
            
        Case 37, 38, 157, 158, 162, 163 'cod RUTA
            EsNomCod = True
            tabla = "srutas"
            codCampo = "codrutas"
            NomCampo = "nomrutas"
            Formato = "000"
            Titulo = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Ruta de Asistencia")
                        
        Case 41, 42, 57, 145 'cod SITUACION
            EsNomCod = True
            tabla = "ssitua"
            codCampo = "codsitua"
            NomCampo = "nomsitua"
            Formato = "00"
            Titulo = "Situaci�n Especial"
            
        Case 52 'cod. Incidencias
            EsNomCod = True
            tabla = "sincid"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            Titulo = "Incidencias"
            
        Case 55, 56, 60, 61, 129, 130, 133, 134 'cod POSTAL
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scpostal", "provincia", "cpostal", "CPostal")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 58, 59, 65, 66, 90, 91, 124, 125 'Cod. PROVEEDOR
            EsNomCod = True
            tabla = "sprove"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
            
        Case 67, 68, 112, 113 'cod. ARTICULO
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Art�culo"
            
        Case 73  'N� de Pedido de Compras
            If txtCodigo(Index).Text = "" Then Exit Sub
            If OpcionListado = 55 Or OpcionListado = 56 Or OpcionListado = 407 Then
                NomCampo = "numpedpr"
                Titulo = "Proveedor"
            Else
                NomCampo = "numpedcl"
                Titulo = "Cliente"
            End If
            NomCampo = DevuelveDesdeBDNew(conAri, NomTabla, NomCampo, NomCampo, txtCodigo(Index).Text, "N")
            If NomCampo = "" Then
                MsgBox "No existe el N� de Pedido de " & Titulo & ": " & txtCodigo(Index).Text, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 94, 95, 100, 101 'cod. FAMILIA articulos
            EsNomCod = True
            tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        Case 126
            If Me.txtCodigo(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtCodigo(Index), 1) Then txtCodigo(Index).Text = ""
            End If
        Case 131, 132 'cliente potencial
            EsNomCod = True
            tabla = "sclipot"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Cli. potenciales"
        Case 159
            'Bancos propios
            EsNomCod = True
            tabla = "sbanpr"
            codCampo = "codbanpr"
            NomCampo = "nombanpr"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Bancos"
        Case 161
            'Bancos propios
            EsNomCod = True
            tabla = "sforpa"
            codCampo = "codforpa"
            NomCampo = "nomforpa"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Forma pago"
        
    End Select
    
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = "" 'Puesto el 25 de enero
            End If

            
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
'            If tabla = "sincid" Then
'                If txtNombre(Index).Text = "" Then txtCodigo(Index).Text = ""
'            End If
            
        End If
    End If
    
    If Index = 71 Or Index = 72 Then DatosFraRectificativa
    
    
End Sub




Private Sub PonerFrameRecordaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Ofertas Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Ofertas
Dim B As Boolean

    H = 7100
    W = 7100
    'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
    PonerFrameVisible Me.FrameRecordatorio, visible, H, W

    If visible = True Then
        '====================================
        Me.OptPapelBlancoR.Value = True

        B = (OpcionListado = 32) '32: Informe Recordatorio
                                 '33: Informe Valoracion Ofertas
        'Carta
        Me.Label4(24).visible = B
        Me.imgBuscarOfer(1).visible = B
        txtCodigo(13).visible = B
        txtNombre(13).visible = B
        'Lineas
        Me.Label4(0).visible = B
        txtCodigo(14).visible = B
        txtCodigo(15).visible = B
        'Pedir Tipo Papel (blanco o con membrete)
        Me.FrameTipoPapel2.visible = B

        'Frame Valorar coste con
        Me.FrameValorar.visible = Not B
        If Not B Then
            Me.FrameValorar.Top = 4520
            Me.FrameValorar.Left = 600
            Me.FrameRecordatorio.Width = 6800
            W = Me.FrameRecordatorio.Width
        End If

        'Poner el Titulo del Frame
        If B Then
            Me.Label7.Caption = "Recordatorio de Ofertas"
        Else
            Me.Label7.Caption = "Valoraci�n de Ofertas"
        End If
    End If
End Sub

   
Private Sub PonerFrameClienInacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Clientes Inactivos Visible y Ajustado al Formulario, y visualiza los controles
'necesarios
Dim B As Boolean

    If OpcionListado = 90 Or OpcionListado = 91 Then
        H = 6980
        Me.cmdAceptarClienInac.Top = 5980
        Me.cmdCancel(5).Top = 5980
    Else
        H = 4460
        Me.cmdAceptarClienInac.Top = 3800
        Me.cmdCancel(5).Top = 3800
    End If
    Me.frameCliexFacturas.visible = OpcionListado = 90
    FrameTipoPago.visible = OpcionListado = 46
    
    If OpcionListado = 90 Or OpcionListado = 91 Then
        W = 11000
    Else
        If OpcionListado = 46 Then
            CargaTipoPagoInactivos
            FrameTipoPago.BorderStyle = 0
            W = 11000
        Else
            W = 6800
        End If
    End If
    
    PonerFrameVisible Me.FrameClienInactivos, visible, H, W

    If visible = True Then
        B = (OpcionListado = 48)
        'Mostrar D/H Fecha
        Label4(43).visible = B
        Label4(44).visible = B
        Me.imgFecha(12).visible = B
        Me.txtCodigo(32).visible = B
        
        If B Then
            Me.Label4(36).Caption = "Fecha Alta"
            Me.Label8(0).Caption = "Altas Nuevos Clientes"
        ElseIf OpcionListado = 90 Or OpcionListado = 91 Then
            Me.Frame1.visible = True
            Me.txtCodigo(31).visible = False
            Me.FrameImpClien.visible = True
            Me.OptCliTodos.Value = True
            If OpcionListado = 90 Then
                Me.Label8(0).Caption = "Etiquetas de Clientes"
                Me.FrameImpClien.Top = 5740
                Me.FrameImpClien.Left = 600
            Else
                Me.Label8(0).Caption = "Cartas a Clientes"
                Me.FrameImpClien.Left = 6800
                Me.FrameImpClien.Top = 4500
            End If
        End If
        Me.Frame4.visible = (OpcionListado = 91)
    End If
End Sub


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function TraspasoOfertaAHco(cadWhere As String) As Boolean
'Realiza el traspaso de las ofertas seleccionadas por cadWhere
'Inserta en la tabla de Historico de ofertas (schpre, slhpre)
'Borra de las tablas de Ofertas (scapre, slipre)
Dim SQL As String
Dim Donde As String
Dim bol As Boolean

'Aqui empieza transaccion
    conn.BeginTrans
    On Error GoTo ETraspasoHco
    bol = ActualizarElTraspaso(Donde, cadWhere, "OFE", CStr(cboMotTras.ItemData(cboMotTras.ListIndex)))

    If bol Then
        '------------------------------
        'A�ado LOG
        Set LOG = New cLOG
        
        SQL = "Mot: " & Me.cboMotTras.Text & "    Traspaso a hco ofertas.  " & cadWhere
        LOG.Insertar 11, vUsu, SQL
        Set LOG = Nothing
        SQL = ""
    End If




ETraspasoHco:
        If Err.Number <> 0 Then
            SQL = "Traspaso Ofertas a Hist�rico." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            TraspasoOfertaAHco = True
        Else
            conn.RollbackTrans
            TraspasoOfertaAHco = False
        End If
End Function


Private Function ObtenerTotalOferPeriodo(cadWhere As String, TotImpA As String, TotImpNA As String, TotOfeA As String, TotOfeNA As String) As Boolean
'para INFORME DE OFERTAS EFECTUADAS
'TotImpA: suma del Importe bruto de todas las Ofertas Aceptadas del periodo seleccionado
'TotImpNA: suma del Importe bruto de todas las Ofertas NO Aceptadas del periodo
'TotOfeA: n� total de ofertas Aceptadas en el periodo
'TotOfeNA: n� total de Ofertas NO Aceptadas en el periodo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpBrutoLin As Currency
Dim ImpBrutoTotA As Currency
Dim ImpBrutoTotNA As Currency
Dim TotalOfeA As Integer
Dim TotalOfeNA As Integer
On Error GoTo ETotalPeriodo

    SQL = "SELECT scapre.numofert, scapre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal, (sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
    SQL = SQL & " FROM scapre INNER join slipre ON scapre.numofert=slipre.numofert "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " GROUP by scapre.numofert "
    
    If Me.chkBusHcoOfer.Value = 1 Then
        
        SQL = SQL & " Union "
        SQL = SQL & " SELECT schpre.numofert, schpre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal,(sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
        SQL = SQL & " FROM schpre iNNER join slhpre ON schpre.numofert=slhpre.numofert "
        If cadWhere <> "" Then
    '        cadWHERE = SustituirCadenas(cadWHERE, "scapre", "schpre")
            cadWhere = Replace(cadWhere, "scapre", "schpre")
            SQL = SQL & " WHERE " & cadWhere
        End If
        SQL = SQL & " GROUP by schpre.numofert "
    End If

    ImpBrutoTotA = 0
    ImpBrutoTotNA = 0
    TotalOfeA = 0
    TotalOfeNA = 0
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        ImpBrutoLin = RS!ImpTotal - RS!impdtopp - RS!impdtogn
        If RS!aceptado = 1 Then 'OFERTA ACEPTADA
            TotalOfeA = TotalOfeA + 1
            ImpBrutoTotA = ImpBrutoTotA + ImpBrutoLin
        Else 'OFERTA NO ACEPTADA
            TotalOfeNA = TotalOfeNA + 1
            ImpBrutoTotNA = ImpBrutoTotNA + ImpBrutoLin
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    TotImpA = Format(ImpBrutoTotA, "0.00")
    TotImpNA = Format(ImpBrutoTotNA, "0.00")
    TotOfeA = TotalOfeA
    TotOfeNA = TotalOfeNA
    ObtenerTotalOferPeriodo = True
    
ETotalPeriodo:
    If Err.Number <> 0 Then ObtenerTotalOferPeriodo = False
End Function


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameClientes
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Actividad, Zona, Ruta, Agente, Situaci�n
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Campo", 1500

    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Actividad"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Zona"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = IIf(vParamAplic.NumeroInstalacion <> vbHerbelca, "Ruta", "Asociacion")
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Agente"
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    NumeroDeCopias = 1
    numParam = 0
    pRptvMultiInforme = 0
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

'Llevara empipados los siguientes datos, para el envio del mail
'DatosEnvioMail:
'       outTipoDocumento
'       outCodigoCliProv
'       outClaveNombreArchiv
Private Sub LlamarImprimir(PonerNombrePDF As Boolean, EnviaPorEmail As Boolean, Optional DatosEnvioMail As String, Optional MostrarTreeEnPrint As Boolean)
     With frmImprimir
        
        If EnviaPorEmail Then
            If Dir(App.Path & "\docum.pdf") <> "" Then Kill App.Path & "\docum.pdf"
        End If
        
        .outTipoDocumento = 0
        If DatosEnvioMail <> "" Then
            .outTipoDocumento = RecuperaValor(DatosEnvioMail, 1)
            .outCodigoCliProv = RecuperaValor(DatosEnvioMail, 2)
            .outClaveNombreArchiv = RecuperaValor(DatosEnvioMail, 3)
        End If
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .Opcion = OpcionListado
        .SoloImprimir = False
        .MostrarTreeDesdeFuera = MostrarTreeEnPrint
        .EnvioEMail = EnviaPorEmail
        
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .NumeroCopias = NumeroDeCopias
        .SeleccionaRPTCodigo = pRptvMultiInforme
        If PonerNombrePDF Then .NombrePDF = cadPDFrpt
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
    If DatosEnvioMail <> "" Then DatosEnvioMail = ""
    pRptvMultiInforme = 0
End Sub




Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
'Pone por que campos se van a AGrupar los datos en el Informe de Crystal Report
'El informe tiene definido 4 formulas a las cuales ahora le asignamos un campo
'de la tabla segun el orden seleccionado para el agrupamiento
Dim campo As String
Dim NomCampo As String

    campo = "pGroup" & numGrupo & "="
    NomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Actividad"
            cadParam = cadParam & campo & "{sclien.codactiv}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""ACTIVIDAD:  "" & " & " totext({sclien.codactiv},""000"") & " & """  """ & " & {sactiv.nomactiv}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codactiv},""000"") & " & """ """ & " & {sactiv.nomactiv}" & "|"
                cadParam = cadParam & NomCampo & "{sactiv.nomactiv}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Actividad""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
            
        Case "Zona"
            cadParam = cadParam & campo & "{sclien.codzonas}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""ZONA:  "" & " & " totext({sclien.codzonas},""000"") & " & """  """ & " & {szonas.nomzonas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codzonas},""000"") & " & """ """ & " & {szonas.nomzonas}" & "|"
                cadParam = cadParam & NomCampo & "{szonas.nomzonas}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Zona""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
            
        Case "Ruta", "Asociacion"
            cadParam = cadParam & campo & "{sclien.codrutas}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " """ & UCase(cadgrupo) & ":  "" & " & " totext({sclien.codrutas},""000"") & " & """  """ & " & {srutas.nomrutas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codrutas},""000"") & " & """ """ & " & {srutas.nomrutas}" & "|"
                cadParam = cadParam & NomCampo & "{srutas.nomrutas}" & "|"
              '  cadParam = cadParam & "pTitulo" & numGrupo & "=""Ruta""" & "|"
                 cadParam = cadParam & "pTitulo" & numGrupo & "=""" & IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Ruta") & """" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
'            PonerGrupo = numGrupo
        Case "Agente"
            cadParam = cadParam & campo & "{sclien.codagent}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""AGENTE:  "" & " & " totext({sclien.codagent},""000000"") & " & """  """ & " & {sagent.nomagent}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codagent},""000000"") & " & """ """ & " & {sagent.nomagent}" & "|"
                cadParam = cadParam & NomCampo & "{sagent.nomagent}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Agente""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
'        Case "Situacion"
    End Select
End Function


Private Function ListaClientesMante(cadWhere As String) As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim SQL As String, Cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista

    Cad = ""
    SQL = "SELECT sclien.codclien "
    SQL = SQL & " FROM sclien INNER JOIN scaman ON sclien.codclien=scaman.codclien "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    ListaClientesMante = Cad
ELista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Clientes con mantenimientos", Err.Description
End Function




Private Function ListaClientesDesdeHastaFactura2() As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim SQL As String, Cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista2

    'Monto el cad
    Cad = ""
    If Me.cboTipomov(2).ListIndex >= 0 Then
        'Tipo mov=
        Cad = " AND codtipom = '" & Mid(Me.cboTipomov(2).List(Me.cboTipomov(2).ListIndex), 1, 3) & "'"
    End If
    If txtCodigo(102).Text <> "" Then Cad = Cad & " AND numfactu >= " & txtCodigo(102).Text
    If txtCodigo(103).Text <> "" Then Cad = Cad & " AND numfactu <= " & txtCodigo(103).Text
    If txtCodigo(104).Text <> "" Then Cad = Cad & " AND fecfactu >= '" & Format(txtCodigo(104).Text, FormatoFecha) & "'"
    If txtCodigo(105).Text <> "" Then Cad = Cad & " AND fecfactu <= '" & Format(txtCodigo(105).Text, FormatoFecha) & "'"
    If Len(Cad) > 0 Then Cad = Mid(Cad, 5) 'QUITO EL PRIMER AND
    
    
    
    'Febrero 2010
    'Si no pongo ningun dato para el desde / hasta factura, no me busca en facturados
    If Cad = "" Then
        ListaClientesDesdeHastaFactura2 = ""
        Exit Function
    End If
    
    
    'A�ado un par de desde/hastas, para acotar. Aunque realmente estan en el SELECT principal
    'si lo pong aqui, acotamos mas
    If txtCodigo(27).Text <> "" Then Cad = Cad & " AND codclien >= " & txtCodigo(27).Text
    If txtCodigo(28).Text <> "" Then Cad = Cad & " AND codclien <= " & txtCodigo(28).Text
    
    
    
    
    SQL = "SELECT DISTINCT(scafac.codclien) "
    SQL = SQL & " FROM scafac "
    If Cad <> "" Then SQL = SQL & " WHERE " & Cad


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1)
    Else
        'NO hay resultados
        Cad = "-1"
    End If
    
    ListaClientesDesdeHastaFactura2 = Cad
ELista2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Procedimiento: ListaClientesDesdeHastaFactura", Err.Description
End Function



Private Sub EnviarEMailMulti(cadWhere As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad1 As String, Cad2 As String, Lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "sprove" Then
        'seleccionamos todos los proveedores a los que queremos enviar e-mail
        SQL = "SELECT codprove,nomprove,maiprov1,maiprov2 "
    ElseIf cadTabla = "sclien" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        SQL = "SELECT codclien,nomclien,maiclie1,maiclie2 "
    End If
    SQL = SQL & "FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    SQL = "CREATE TEMPORARY TABLE tmpMail ( "
    SQL = SQL & "codusu SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    SQL = SQL & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute SQL
    
    cont = 0
    Lista = ""
    
    While Not RS.EOF
    'para cada cliente/proveedor enviamos un e-mail
        cad1 = DBLet(RS.Fields(2), "T") 'e-mail administracion
        Cad2 = DBLet(RS.Fields(3), "T") 'e-mail compras
        
        If cad1 = "" And Cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              Lista = Lista & Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & vbCrLf
        ElseIf cad1 <> "" And Cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If cadTabla = "sprove" Then
                If Me.OptMailCom(0).Value = True Then cad1 = Cad2
            Else
                If Me.OptMailCom(1).Value = True Then cad1 = Cad2
            End If
        Else 'alguno de los 2 tiene valor
            If Cad2 <> "" Then cad1 = Cad2  'e-mail para compras
        End If
        
        If cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            With frmImprimir
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                If cadTabla = "sprove" Then
                    SQL = "{sprove.codprove}=" & RS.Fields(0)
                    .Opcion = 306
                Else
                    SQL = "{sclien.codclien}=" & RS.Fields(0)
                    .Opcion = 91
                End If
                .FormulaSeleccion = SQL
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = cadTit
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(RS.Fields(0), "N") & "," & DBSet(RS.Fields(1), "T") & "," & DBSet(cad1, "T") & ")"
                    conn.Execute SQL
            
                    Me.Refresh
                    Espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    SQL = RS.Fields(0) & ".pdf"
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
                End If
            End With
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
      
    If cont > 0 Then
        Espera 0.4
        If cadTabla = "sprove" Then
            SQL = "Carta: " & txtNombre(63).Text & "|"
             SQL = SQL & "Att : " & txtCodigo(62).Text & "|"
        Else
            SQL = "Carta: " & txtNombre(64).Text & "|"
            SQL = SQL & "Att : " & txtCodigo(0).Text & "|"
        End If
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = SQL
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
        
        'Borrar la carpeta con temporales
        Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If Lista <> "" Then
        If cadTabla = "sprove" Then
            Lista = "Proveedores sin e-mail:" & vbCrLf & vbCrLf & Lista
        Else
            Lista = "Clientes sin e-mail:" & vbCrLf & vbCrLf & Lista
        End If
        MsgBox Lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub




Private Sub CargarComboTipoMov(Indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Byte

    On Error GoTo ECargaCombo

'    SQL = "select codtipom, nomtipom from stipom where tipodocu=2 " 'Documentos de Facturas
    '3 abril 2007.
    'Mostraba todas las facturas (movimientos que empizan por F, excepto las rectificativas
    'AHora tiene que mostrarlas todas
    'SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__') and (codtipom<>'FRT')"
    SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__')"  ' and (codtipom<>'FRT')"
    If Indice = 1 Then SQL = SQL & " ORDER BY codtipom"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    numParam = 120
    If Indice < 1000 Then
            'Son combos normales
         cboTipomov(Indice).Clear
        
         While Not RS.EOF
             cboTipomov(Indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
             cboTipomov(Indice).ItemData(cboTipomov(Indice).NewIndex) = i
             If Indice = 1 Then
                If RS.Fields(0) = "FAV" Then numParam = i
             End If
             i = i + 1
             RS.MoveNext
         Wend
        
        
        
        'Para la reimpresion de facturas, a�adiremos al final la opcion *** TODAS
        ' que ser�n todas menos FAT, FTI  FAG que son impresiones distintas
        If Indice = 1 And OpcionListado = 226 Then
            cboTipomov(Indice).AddItem "*** Todas *** "
            cboTipomov(Indice).ItemData(cboTipomov(Indice).NewIndex) = i
            i = i + 1
        End If
        
        
    Else
        
        ListTipoMov(Indice).Clear
        
        
        'LOS TIKCETS NO LOS ENVIO POR MAIL
        
        'Febrero 2013. Si que se pueden poner los tikets NOMINATIVOS
        While Not RS.EOF
'            If RS!codtipom <> "FTI" Then
'
'                ListTipoMov(indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
'                'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
'                ListTipoMov(indice).Selected((ListTipoMov(indice).NewIndex)) = True
'            End If
            
            
                ListTipoMov(Indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
                'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
                ListTipoMov(Indice).Selected((ListTipoMov(Indice).NewIndex)) = True
            


            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'Pongo el dos para todos menos para la de etiquetas cliente
    If Indice < 1000 Then
        If Indice <> 2 Then
            If Indice = 1 Then
                'Julio 2021
                'Ponemos FAV como predeterminado
                If numParam > 100 Then numParam = 2
            Else
                numParam = 2
            End If
            Me.cboTipomov(Indice).ListIndex = numParam
        End If
    End If
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub PonerFramePedVisible(H As Integer, W As Integer)
'Frame de Pedidos de Venta y Compra
    W = 6075
    H = 4455
    
    chkVarios(10).Value = 0
    chkVarios(10).visible = False
    PonerFrameVisible Me.FramePedidos, True, H, W
    Select Case OpcionListado
        Case 38 'PEdidos venta
            Me.Label12(0).Caption = "Informe Pedidos ventas"
            NomTabla = "scaped"
            NomTablaLin = "sliped"
            Me.Label12(3).Caption = "Imprimir otros Pedidos del Cliente:"
            'Estaba asi 09/junio/21 Lo pongo a true vUsu.Nivel < 1
            chkVarios(10).visible = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, True, vUsu.Nivel < 1) 'en herbelc siempre visible
        Case 239 'Historico de Pedidos Venta
            Me.Label12(0).Caption = "Informe Hist. Pedidos ventas"
            NomTabla = "schped" 'Cabecera  Hco de Pedidos de clientes
            NomTablaLin = "slhped"
            If FecEntre <> "" Then txtCodigo(76).Text = FecEntre
        Case 55, 407 'Cabecera de Pedidos de Compras (a proveedores)
            Me.Label12(0).Caption = "Informe Pedidos compras"
            NomTabla = "scappr"
            NomTablaLin = "slippr"
        Case 56 'Historico de Pedidos Compras
            Me.Label12(0).Caption = "Informe Hist. Pedidos compras"
            NomTabla = "schppr" 'Cabecera  Hco de Pedidos de Compras (a proveedores)
            NomTablaLin = "slhppr"
            If FecEntre <> "" Then txtCodigo(76).Text = FecEntre
    End Select
    
    Me.chkVarios(0).visible = OpcionListado = 55
    'Ver Fecha Pedido (En Hist.)
    Label12(2).visible = (OpcionListado = 239) Or OpcionListado = 56
    txtCodigo(76).visible = (OpcionListado = 239) Or OpcionListado = 56
End Sub






 
Private Sub PonerTamnyosMail2(peque As Boolean)

    
        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.visible = peque
        Me.cmdEnvioMail.Enabled = Not peque
    
    'If peque Then
    '    Me.Height = Me.FrameEnvioMail.Height
    '    Me.Width = Me.FrameEnvioMail.Width
   '
   ' Else
   '     Me.Height = Me.FrameEnvioFacMail.Height
   '     Me.Width = Me.FrameEnvioFacMail.Width
   '
   ' End If
   '
   ' Me.Height = Me.Height + 420
   ' Me.Width = Me.Width + 120
   ' Me.FrameEnvioMail.visible = peque
    DoEvents
    Me.Refresh
End Sub


Private Sub GenerarConfirmacionPedidos()
Dim C As String
Dim ListaCli As Collection

    




    
    
    'Vaciamos la carpeta.
    Label4(97).Caption = "Vaciar datos anteriores"
    Label4(97).Refresh
    
    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    
    C = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute C
    
    
'    'Crearemos los pdf de confirmacino de envio
'
'    C = "Select codclien FR        "
'      While Not miRsAux.EOF
'            If IsNull(miRsAux!el_mail) Then
'                devuelve = devuelve & "    - " & miRsAux!nomclien & vbCrLf
'            Else
'                'INSERTAMOS
'                NumRegElim = NumRegElim + 1
'                Codigo = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
'                Codigo = Codigo & vUsu.Codigo & ",1,'" & Format(txtCodigo(109).Text, FormatoFecha) & "'," & miRsAux!codClien & ","
'                Codigo = Codigo & NumRegElim & ",'" & miRsAux!nummante & "')"
'                conn.Execute Codigo
'            End If
'            miRsAux.MoveNext
'        Wend
'        miRsAux.Close
'
'
'
    NumCod = cadFormula 'WHERE original
    codClien = "" 'Clientes SIn email
    FecEntre = ""
    C = "Select codclien from scaped "
    If cadSelect <> "" Then
        C = C & " WHERE " & cadSelect

    End If
    C = C & " GROUP BY codclien"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ListaCli = New Collection
    While Not miRsAux.EOF
        Label4(97).Caption = "Carta cliente: " & miRsAux!codClien
        Label4(97).Refresh
        Codigo = "maiclie2"
        C = DevuelveDesdeBD(conAri, "maiclie1", "sclien", "codclien", CStr(miRsAux.Fields(0)), "N", Codigo)
        
        
        C = Trim(C)
        If C = "" Then C = Codigo
        If C = "" Then
            'NO eiste eail
            codClien = codClien & ", " & miRsAux!codClien
        
        Else
        
            'La carta de confirmacion en formato pdf
            If NumCod <> "" Then
                cadFormula = " AND "
            Else
                cadFormula = ""
            End If
            cadFormula = NumCod & cadFormula & " ({scaped.codclien} = " & miRsAux.Fields(0) & ")"
        
            LlamarImprimir True, True
            If Dir(App.Path & "\docum.pdf") = "" Then
                'HA HABIDO UN ERROR
                MsgBox "No se encuentra pdf para cliente: " & miRsAux!codClien, vbExclamation
            Else
                FileCopy App.Path & "\docum.pdf", App.Path & "\temp\CL" & Format(miRsAux.Fields(0), "00000") & ".pdf"
                
                'Como no me cabe voy a utlizar numalbar+nomartic+numlotes para la direccion email
                cadFormula = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic,nomartic,numlotes) values ("
                cadFormula = cadFormula & vUsu.Codigo & ",'" & Mid(C, 1, 10) & "','1972-04-12'," & miRsAux!codClien & ",1"
                cadFormula = cadFormula & ",'CL" & Format(miRsAux.Fields(0), "00000") & ".pdf" & "','" & Mid(C, 11, 40) & "',"
                If Len(C) > 50 Then
                    C = "'" & Mid(C, 51) & "'"
                Else
                    C = "NULL"
                End If
                cadFormula = cadFormula & C & ")"
                If ejecutar(cadFormula, False) Then ListaCli.Add CStr(miRsAux!codClien)
            
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    cadFormula = ""
    'Aqui deberiamos hacer un select
    If codClien <> "" Then codClien = "No existe email para clientes: " & Mid(codClien, 2)
    If ListaCli.Count = 0 Then cadFormula = "No se ha generado ningun dato" & vbCrLf & codClien
        
    If cadFormula <> "" Then
        MsgBox cadFormula, vbExclamation
        Exit Sub
    Else
        If codClien <> "" Then MsgBox codClien, vbExclamation
    End If
        
        
    If chkConfirmPed(1).Value = 0 Then Exit Sub 'NO adjuntamos los pedidos
        
    'Vamos a meter los pedidos adjuntos a las cartas anteriores
    'Memorizo el cadselect
    C = cadSelect
    InicializarVbles
    cadSelect = C
    
    
    '38. Pedidos
    If Not PonerParamRPT2(7, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
        Exit Sub
    End If

    
    For NumRegElim = 1 To ListaCli.Count
    
            Label4(97).Caption = "Pedidos para : " & CStr(ListaCli.Item(NumRegElim)) & "   (" & NumRegElim & " de " & ListaCli.Count & ")"
            Label4(97).Refresh
    
    
            'El resto de pedidos
            C = "Select numpedcl from scaped WHERE"
            
            If cadSelect <> "" Then
                cadFormula = " AND "
            Else
                cadFormula = ""
            End If
            cadFormula = " " & cadSelect & cadFormula & " scaped.codclien = " & CStr(ListaCli.Item(NumRegElim))
            C = C & cadFormula & " ORDER BY numpedcl"
            
        
            miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            C = ""
            While Not miRsAux.EOF
                C = C & " OR ({scaped.numpedcl} = " & miRsAux!NumPedcl & ")"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
           
            cadFormula = Mid(C, 4)
            
            LlamarImprimir True, True
    
            If Dir(App.Path & "\docum.pdf") = "" Then
                'HA HABIDO UN ERROR
                MsgBox "No se encuentra pdf para pedidos: " & CStr(ListaCli.Item(NumRegElim)), vbExclamation
            Else
                FileCopy App.Path & "\docum.pdf", App.Path & "\temp\PED" & Format(CStr(ListaCli.Item(NumRegElim)), "00000") & ".pdf"
            
            
                cadFormula = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
                cadFormula = cadFormula & vUsu.Codigo & ",2,'1974-04-12'," & ListaCli.Item(NumRegElim) & ",1"
                cadFormula = cadFormula & ",'PED" & Format(CStr(ListaCli.Item(NumRegElim)), "00000") & ".pdf" & "')"
                ejecutar cadFormula, False
            
            End If
    
    Next NumRegElim
    
    
    
End Sub





Private Sub PreparaDatosLineasCompras()
Dim Aux As String
Dim Cad As String
Dim R2 As ADODB.Recordset
Dim Col As Collection
Dim FinBusq As Boolean
Dim FinPpal As Boolean


    On Error GoTo EPreparaDatosLineasCompras
    Screen.MousePointer = vbHourglass
    Label9(38).Caption = "Rappel fras."
    Label9(38).Refresh
    conn.Execute "Delete from tmpcommand where codusu = " & vUsu.Codigo

    'Habra que mirar para cada
    'Las facturas SIEMPRE las pone
    Aux = "SELECT " & vUsu.Codigo & ",`scafpc`.`codprove`, `scafpc`.`nomprove`, `slifpc`.`cantidad`, `slifpc`.`importel`, `sartic`.`codfamia`, `sfamia`.`nomfamia`, `scafpc`.`fecrecep`, fechaalb,sartic.nomartic,sartic.codartic"
    Aux = Aux & ",0,0" 'Despues vere el Rappel
    Aux = Aux & " FROM   (`scafpc` `scafpc` INNER JOIN `scafpa` scafpa ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`=`scafpa`.`fecfactu`))"
    Aux = Aux & " AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
    Aux = Aux & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
    Aux = Aux & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
    Aux = Aux & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
    Aux = Aux & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"
    
    Aux = "insert into `tmpcommand` (`codusu`,`codprove`,`nomprove`,`cantidad`,`importel`,`codfamia`,`nomfamia`,`fecrecep`,`fechaalb`,`nomartic`,`codartic`,`rap1`,`rap2`) " & Aux
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "{", "")
        Codigo = Replace(Codigo, "}", "")
        Aux = Aux & " WHERE " & Codigo
    End If
    conn.Execute Aux
    
    'Si tiene puesto la marca de albranes
    If Me.chkDatosAlbaranes(1).Value = 1 Then
        Label9(38).Caption = "Rappel alb."
        Label9(38).Refresh
        Aux = "SELECT " & vUsu.Codigo & ", scaalp.`codprove`, `nomprove`, `cantidad`, `importel`, `sartic`.`codfamia`, `sfamia`.`nomfamia`,"
        Aux = Aux & " scaalp.`fechaalb` fecrecep, scaalp.`fechaalb`,sartic.nomartic,sartic.codartic"
        Aux = Aux & ",0,0" 'Despues vere el Rappel
        Aux = Aux & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
        Aux = Aux & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
        Aux = Aux & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
        Aux = "insert into `tmpcommand` (`codusu`,`codprove`,`nomprove`,`cantidad`,`importel`,`codfamia`,`nomfamia`,`fecrecep`,`fechaalb`,`nomartic`,`codartic`,`rap1`,`rap2`) " & Aux
        If cadSelect <> "" Then
            Codigo = Replace(cadSelect, "{", "")
            Codigo = Replace(Codigo, "}", "")
            Codigo = Replace(Codigo, "scafpc", "scaalp")
            Codigo = Replace(Codigo, "scafpa", "scaalp")
            Codigo = Replace(Codigo, "fecfactu", "fechaalb")
            Aux = Aux & " WHERE " & Codigo
        End If
        conn.Execute Aux
    
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Importe minimo
    
    If txtCodigo(126).Text <> "" Then
        Label9(38).Caption = "Ajuste importe minimo"
        Label9(38).Refresh
        'Me quito de en miedo las familias que no superen esto
        Aux = "select codfamia,codprove,sum(importel) s1 from tmpcommand where tmpcommand.codusu=" & vUsu.Codigo & "  group by codfamia,codprove having s1<" & TransformaComasPuntos(CStr(ImporteFormateado(txtCodigo(126).Text)))
        Aux = Aux & "  ORDER BY codprove,codfamia"
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Label9(38).Caption = miRsAux!Codprove & " / " & miRsAux!Codfamia
            Label9(38).Refresh
            Aux = "DELETE FROM tmpcommand where codusu =" & vUsu.Codigo & " AND codfamia =" & miRsAux!Codfamia & " AND codprove =" & miRsAux!Codprove
            conn.Execute Aux
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    
    
    Aux = "select tmpcommand.codprove from tmpcommand where tmpcommand.codusu=" & vUsu.Codigo & " group by 1"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    indCodigo = 0
    Set Col = New Collection
    While Not miRsAux.EOF
        indCodigo = indCodigo + 1
        codClien = codClien & ", " & miRsAux!Codprove
        'estoy agrupando por proveedor para luego ir a famila dto en sdtomp(asi iremos mas rapido)
        miRsAux.MoveNext
        If indCodigo > 20 Then
            
            Col.Add codClien
            Label9(38).Caption = Col.Count & " ...."
            Label9(38).Refresh
            codClien = ""
            indCodigo = 1
        End If
    Wend
    miRsAux.Close
    NumRegElim = 1
    If indCodigo < 1 Then
        'NIO habia ninguno
        MsgBox "No existen datos", vbExclamation
        NumRegElim = 0
        Set Col = Nothing
        Set miRsAux = Nothing
        Exit Sub
    Else
        If codClien <> "" Then Col.Add codClien
    End If
    
    
    '### Aqui deberia ir lo de dtosincargo  DAVID###
    'Para cada N proveedores voy buscando su dtopm
    DoEvents
    Me.Refresh
    Set R2 = New ADODB.Recordset
    For indCodigo = 1 To Col.Count
        Label9(38).Caption = indCodigo & " de " & Col.Count
        Label9(38).Refresh
    
    
        codClien = Col.Item(indCodigo)
        codClien = Mid(codClien, 2)  'quito la primera coma
        
        
        Aux = "Select * from sdtomp where codprove in (" & codClien & ")"
        Aux = Aux & " AND ( rap1 >0 or rap2 >0)"
        Aux = Aux & " order by codprove,codfamia,codmarca desc"
        R2.Open Aux, conn, adOpenKeyset, adLockOptimistic, adCmdText
        If Not R2.EOF Then
            'Hay alguno por lo menos
                    'FEBRERO.... NO puede cruzar por CODMARCA
                    'Aux = "select tmpcommand.codprove,tmpcommand.codfamia,codmarca from tmpcommand,sartic where"
                    'Aux = Aux & " tmpcommand.codartic=sartic.codartic and tmpcommand.codusu=" & vUsu.Codigo
                    Aux = "select tmpcommand.codprove,tmpcommand.codfamia from tmpcommand where"
                    Aux = Aux & " tmpcommand.codusu=" & vUsu.Codigo
                    Aux = Aux & " AND tmpcommand.codprove IN (" & codClien & ")"
                    Aux = Aux & " group by 1,2 "  'Aux = Aux & " group by 1,2,codmarca "
                    Aux = Aux & " order by 1,2 "  ',codmarca
                
                    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    codClien = ""
                    FinPpal = False
                    While Not miRsAux.EOF
                        'El de los dtos
                        FinBusq = False
                        'R2.MoveFirst
                        codClien = miRsAux!Codprove
                        
                        R2.Find "codprove = " & codClien, , adSearchForward, 1
                        If R2.EOF Then
                               'para un determinado mirsaux!codprove NO tiene nada en R2
                               'voy a mover hasta encontrar otro proveedor
                               FinPpal = False
                               While Not FinPpal
                                    If CStr(miRsAux!Codprove) = codClien Then
                                        miRsAux.MoveNext
                                        If miRsAux.EOF Then FinPpal = True
                                    Else
                                        FinPpal = True
                                    End If
                                Wend
                        Else
                            While Not FinBusq
                                If R2!Codprove = miRsAux!Codprove Then
                                
                                        'MISMO PROVEEDOR
                                        If R2!Codfamia > miRsAux!Codfamia Then
                                            FinBusq = True
                                            
                                        Else
                                            If R2!Codfamia = miRsAux!Codfamia Then
                                                'Si SON LA marca es NULL tb se aplica
     
                                                    Aux = TransformaComasPuntos(DBLet(R2!Rap1, "N"))
                                                    Cad = TransformaComasPuntos(DBLet(R2!Rap2, "N"))
                                                    If Aux = "0" And Cad = "0" Then
                                                        'YA ESTA EL CERO
                                                    Else
                                                        Aux = "UPDATE tmpcommand,sartic set rap1=" & Aux
                                                        Aux = Aux & ", rap2 = " & Cad
                                                        Aux = Aux & " WHERE tmpcommand.codartic = sartic.codartic"
                                                        Aux = Aux & " AND tmpcommand.codprove = " & R2!Codprove
                                                        Aux = Aux & " AND tmpcommand.codfamia = " & R2!Codfamia
                                                        Aux = Aux & " AND tmpcommand.codusu = " & vUsu.Codigo
                                                        'HERBELCA. No llevan codmarca
                                                        'If Not IsNull(R2!codmarca) Then Aux = Aux & " AND codmarca = " & R2!codmarca
                                                        conn.Execute Aux
                                                    End If
                                                    FinBusq = True
                                               
                                            End If
                                        End If
                      
                                Else
                                    '<> codprove
                                    FinBusq = True
                                End If
                                If Not FinBusq Then
                                    R2.MoveNext
                                    If R2.EOF Then FinBusq = True
                                    
                                End If
                            Wend
                                                
                            miRsAux.MoveNext
                        End If
                        
                    Wend
                    miRsAux.Close
        End If
        R2.Close
    Next indCodigo
    
    
    Label9(38).Caption = "Datos con rappel"
    Label9(38).Refresh
    Aux = "DELETE FROM tmpcommand WHERE rap1=0 and rap2 =0 and codusu = " & vUsu.Codigo
    conn.Execute Aux
    
EPreparaDatosLineasCompras:
    If Err.Number <> 0 Then MuestraError Err.Number, "PreparaDatosLineasCompras"
    Set miRsAux = Nothing
    Set R2 = Nothing
    Set Col = Nothing
    indCodigo = 0
    codClien = ""
    Label9(38).Caption = ""
    Screen.MousePointer = vbDefault
End Sub






'------
'Cuando pide el compras por articulo familia COMPARATIVO
Private Sub ponerLineasComprasComparatativo()
    Label9(38).Caption = "Preparando datos"
    Label9(38).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    
    DatosLineasComprasComparatativo False
    DatosLineasComprasComparatativo True 'perido antiguo
    
    If Me.txtCodigo(126).Text <> "" Then
        'hay que eliminar importes....
        Label9(38).Caption = "Ajuste importe minimo"
        Label9(38).Refresh
        Set miRsAux = New ADODB.Recordset
        
        Codigo = "select codigo1,campo1,sum(importe2) s1,sum(importe4) s2 from tmpinformes where codusu= " & vUsu.Codigo
        Codigo = Codigo & " group by codigo1,campo1 having s1 < " & TransformaComasPuntos(CStr(ImporteFormateado(txtCodigo(126).Text)))
        Codigo = Codigo & " and s2< " & TransformaComasPuntos(CStr(ImporteFormateado(txtCodigo(126).Text)))
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Codigo = "DELETE FROM tmpinformes WHERE codusu=" & vUsu.Codigo & " AND codigo1=" & miRsAux!Codigo1 & " AND campo1=" & miRsAux!campo1
            miRsAux.MoveNext
            conn.Execute Codigo
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
   
End Sub


Private Sub DatosLineasComprasComparatativo(Comparativo As Boolean)
Dim Aux As String
Dim Cad As String


    On Error GoTo EPreparaDatosLineasCompras
    Screen.MousePointer = vbHourglass
    Label9(38).Caption = "Compartivo facturas"
    Label9(38).Refresh
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Traspaso los datos de tmpcommand aqui
    Aux = "SELECT " & vUsu.Codigo & ",`scafpc`.`codprove`, `sartic`.`codfamia`,`scafpc`.`nomprove`, `sfamia`.`nomfamia`,"
    Aux = Aux & "  fechaalb,sum(`slifpc`.`cantidad`), sum(`slifpc`.`importel`),0,0 "
    Aux = Aux & " FROM   (`scafpc` `scafpc` INNER JOIN `scafpa` scafpa ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`=`scafpa`.`fecfactu`))"
    Aux = Aux & " AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
    Aux = Aux & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
    Aux = Aux & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
    Aux = Aux & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
    Aux = Aux & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"
    
    
    If cadSelect <> "" Then
            Codigo = Replace(cadSelect, "{", "")
            Codigo = Replace(Codigo, "}", "")
            If Comparativo Then
                'replace de fecha
                Cad = "'" & Year(txtCodigo(92).Text) & "-"
                Codigo = Replace(Codigo, Cad, "'" & CStr(CInt(Year(txtCodigo(92).Text)) - 1) & "-")
            End If
            Aux = Aux & " WHERE " & Codigo
    End If
    
    Aux = Aux & " GROUP BY 1,2,3"
    If Comparativo Then
        Aux = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,fecha1,importe3,importe4,importe1,importe2) " & Aux
    Else
        Aux = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,fecha1,importe1,importe2,importe3,importe4) " & Aux
    End If
    conn.Execute Aux
    
    
    'Si tiene puesto la marca de albranes
    If Me.chkDatosAlbaranes(1).Value = 1 Then
        Label9(38).Caption = "Compartivo albaranes"
        Label9(38).Refresh
        Aux = "SELECT " & vUsu.Codigo & ", scaalp.`codprove`,`sartic`.`codfamia`, `nomprove`,   `sfamia`.`nomfamia`,scaalp.`fechaalb`,sum(`cantidad`), sum(`importel`),0,0 "
        Aux = Aux & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
        Aux = Aux & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
        Aux = Aux & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
        If cadSelect <> "" Then
            Codigo = Replace(cadSelect, "{", "")
            Codigo = Replace(Codigo, "}", "")
            Codigo = Replace(Codigo, "scafpc", "scaalp")
            Codigo = Replace(Codigo, "scafpa", "scaalp")
            Codigo = Replace(Codigo, "fecfactu", "fechaalb")
            If Comparativo Then
                'replace de fecha
                Cad = "'" & Year(txtCodigo(92).Text) & "-"
                Codigo = Replace(Codigo, Cad, "'" & CStr(Year(txtCodigo(92).Text) - 1) & "-")
            End If
            Aux = Aux & " WHERE " & Codigo
        End If
        Aux = Aux & " GROUP BY 1,2,3"
        If Comparativo Then
            Aux = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,fecha1,importe3,importe4,importe1,importe2) " & Aux
        Else
            Aux = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,fecha1,importe1,importe2,importe3,importe4) " & Aux
        End If
        conn.Execute Aux
    
    End If
    

   

    
   
    
EPreparaDatosLineasCompras:
    If Err.Number <> 0 Then MuestraError Err.Number, "PreparaDatosLineasCompras"

    indCodigo = 0
    codClien = ""
    Label9(38).Caption = ""
    Screen.MousePointer = vbDefault
End Sub






'Informe de clientes por agente con volumen de ventas
'----------------------------------------------------------

Private Function CalculaVolumenVtas_() As Boolean

On Error GoTo ECalculaVolumenVtas_
    CalculaVolumenVtas_ = False
    
    Codigo = "DELETE FROM tmpstockfec WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    'codigo = "Select " & vUsu.codigo & ",'DAVID' ,codclien,sum(totalfac)"  VA sin IVA
    Codigo = "Select " & vUsu.Codigo & ",'DAVID' ,codclien,sum(baseimp1 + if(baseimp2 is null, 0,baseimp2) + if(baseimp3 is null, 0,baseimp3))"
    Codigo = Codigo & " from scafac where codtipom<>'FAZ'"
    If Me.txtCodigo(122).Text <> "" Then Codigo = Codigo & " AND fecfactu>=" & DBSet(txtCodigo(122).Text, "F")
    If Me.txtCodigo(123).Text <> "" Then Codigo = Codigo & " AND fecfactu<=" & DBSet(txtCodigo(123).Text, "F")
    If cadSelect <> "" Then Codigo = Codigo & " AND codclien IN (Select codclien from sclien WHERE " & cadSelect & ")"
    Codigo = Codigo & " GROUP BY 1,2,3"
    
    Codigo = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock)  " & Codigo
    conn.Execute Codigo
    
    
    'Metere los que no hubieran facturado
    Codigo = "Select " & vUsu.Codigo & ",'DAVID',codclien,0 from sclien WHERE 1=1 "
    If cadSelect <> "" Then Codigo = Codigo & " AND " & cadSelect
    Codigo = "INSERT IGNORE INTO tmpstockfec(codusu,codartic,codalmac,stock) " & Codigo
    conn.Execute Codigo
    
    
    CalculaVolumenVtas_ = True
ECalculaVolumenVtas_:
    If Err.Number <> 0 Then MuestraError Err.Number, "Calculando volumen ventas"
End Function





'Insertamos en temporal para las estadisiticas
' CadSelectPeriodoAnterior--> Qioere compartativo
Private Function InsertarTmpEstdisticasVtas(CadSelectPeriodoAnterior As String) As Boolean
Dim C As String
    On Error GoTo eInsertarTmpEstdisticasVtas

    'Con albaranes
    Codigo = cadSelect
    Codigo = QuitarCaracterACadena(Codigo, "{")
    Codigo = QuitarCaracterACadena(Codigo, "}")
    
    
    InsertarTmpEstdisticasVtas = False
    'Lo facturado
    
    C = "insert into tmpcommandest(codusu,codclien,codfamia,nomclien,nomfamia,cantidad,importel,fechaalb,codprove,nomprove,codartic,nomartic)"
    C = C & " SELECT " & vUsu.Codigo & ",scafac.codclien,sartic.codfamia,scafac.nomclien,nomfamia,cantidad,importel,scafac.fecfactu,sartic.codprove,'',slifac.codartic,slifac.nomartic FROM"
   
    C = C & " ((((`scafac1` `scafac1` INNER JOIN `scafac` `scafac` ON"
    C = C & " ((`scafac1`.`codtipom`=`scafac`.`codtipom`) AND (`scafac1`.`numfactu`=`scafac`.`numfactu`))"
    C = C & " AND (`scafac1`.`fecfactu`=`scafac`.`fecfactu`)) INNER JOIN `slifac` `slifac` ON"
    C = C & " ((((`scafac1`.`codtipom`=`slifac`.`codtipom`) AND (`scafac1`.`numfactu`=`slifac`.`numfactu`))"
    C = C & " AND (`scafac1`.`fecfactu`=`slifac`.`fecfactu`)) AND (`scafac1`.`numalbar`=`slifac`.`numalbar`))"
    C = C & " AND (`scafac1`.`codtipoa`=`slifac`.`codtipoa`)) INNER JOIN `sartic` `sartic`"
    C = C & " ON `slifac`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
    
    C = C & ")  INNER JOIN `sclien` ON `sclien`.`codclien`=`scafac`.`codclien`"
    
    If Codigo <> "" Then C = C & " WHERE " & Codigo

    conn.Execute C


    C = "insert into tmpcommandest(codusu,codclien,codfamia,nomclien,nomfamia,cantidad,importel,fechaalb,codprove,nomprove,codartic,nomartic)"
    C = C & " SELECT " & vUsu.Codigo & ",scaalb.codclien,sartic.codfamia,scaalb.nomclien,nomfamia,cantidad,importel,scaalb.fechaalb,sartic.codprove,'',slialb.codartic,slialb.nomartic FROM"
   
    C = C & "  (((`slialb` INNER JOIN scaalb ON `slialb`.`codtipom`=`scaalb`.`codtipom` AND"
    C = C & " `slialb`.`numalbar`=`scaalb`.`numalbar`)"
    C = C & " INNER JOIN `sartic` `sartic` ON `slialb`.`codartic`=`sartic`.`codartic`)"
    C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
    C = C & " INNER JOIN `sclien` ON `scaalb`.`codclien`=`sclien`.`codclien`"
    If Codigo <> "" Then
        Codigo = Replace(Codigo, "scafac1", "scaalb")
        Codigo = Replace(Codigo, "scafac", "scaalb")
        Codigo = Replace(Codigo, "slifac", "slialb")
        
        C = C & " WHERE " & Codigo
    End If
    
    
    If Me.chkDatosAlbaranes(0).Value = 1 Then conn.Execute C      'con albaranes




    'A�os anteriores
    'Lo facturado
    If CadSelectPeriodoAnterior <> "" Then
            Codigo = CadSelectPeriodoAnterior
            Codigo = QuitarCaracterACadena(Codigo, "{")
            Codigo = QuitarCaracterACadena(Codigo, "}")
            
            
            
            C = "insert into tmpcommandest(codusu,codclien,codfamia,nomclien,nomfamia,cantidadAnt,importlAnt,fechaalb,codprove,nomprove,codartic,nomartic)"
            C = C & " SELECT " & vUsu.Codigo & ",scafac.codclien,sartic.codfamia,scafac.nomclien,nomfamia,cantidad,importel,scafac.fecfactu,sartic.codprove,'',slifac.codartic,slifac.nomartic FROM"
           
            C = C & " ((((`scafac1` `scafac1` INNER JOIN `scafac` `scafac` ON"
            C = C & " ((`scafac1`.`codtipom`=`scafac`.`codtipom`) AND (`scafac1`.`numfactu`=`scafac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`scafac`.`fecfactu`)) INNER JOIN `slifac` `slifac` ON"
            C = C & " ((((`scafac1`.`codtipom`=`slifac`.`codtipom`) AND (`scafac1`.`numfactu`=`slifac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`slifac`.`fecfactu`)) AND (`scafac1`.`numalbar`=`slifac`.`numalbar`))"
            C = C & " AND (`scafac1`.`codtipoa`=`slifac`.`codtipoa`)) INNER JOIN `sartic` `sartic`"
            C = C & " ON `slifac`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            
            C = C & ")  INNER JOIN `sclien` ON `sclien`.`codclien`=`scafac`.`codclien`"
            
            If Codigo <> "" Then C = C & " WHERE " & Codigo
        
            conn.Execute C
        
        
            C = "insert into tmpcommandest(codusu,codclien,codfamia,nomclien,nomfamia,cantidad,importel,fechaalb,codprove,nomprove,codartic,nomartic)"
            C = C & " SELECT " & vUsu.Codigo & ",scaalb.codclien,sartic.codfamia,scaalb.nomclien,nomfamia,cantidad,importel,scaalb.fechaalb,sartic.codprove,'',slialb.codartic,slialb.nomartic FROM"
           
            C = C & "  (((`slialb` INNER JOIN scaalb ON `slialb`.`codtipom`=`scaalb`.`codtipom` AND"
            C = C & " `slialb`.`numalbar`=`scaalb`.`numalbar`)"
            C = C & " INNER JOIN `sartic` `sartic` ON `slialb`.`codartic`=`sartic`.`codartic`)"
            C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
            C = C & " INNER JOIN `sclien` ON `scaalb`.`codclien`=`sclien`.`codclien`"
            If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafac1", "scaalb")
                Codigo = Replace(Codigo, "scafac", "scaalb")
                Codigo = Replace(Codigo, "slifac", "slialb")
                
                C = C & " WHERE " & Codigo
            End If
        
            If Me.chkDatosAlbaranes(0).Value = 1 Then conn.Execute C      'con albaranes

    End If

    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    C = "Select distinct codprove from tmpcommandest where codusu = " & vUsu.Codigo
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        C = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", CStr(miRsAux!Codprove))
        
        C = "UPDATE tmpcommandest set nomprove=" & DBSet(C, "T") & " WHERE codusu =" & vUsu.Codigo & " AND codprove=" & miRsAux!Codprove
        conn.Execute C
        miRsAux.MoveNext
    Wend
    miRsAux.Close
















    InsertarTmpEstdisticasVtas = True
eInsertarTmpEstdisticasVtas:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    
End Function



Private Sub CargaCargos()
Dim IT As ListItem
    Set miRsAux = New ADODB.Recordset
    lwCargos.ListItems.Clear
    miRsAux.Open "Select * from scargoscli order by cargo", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'El prinmero vacio
    While Not miRsAux.EOF
        Set IT = lwCargos.ListItems.Add()
        IT.Text = miRsAux!cargo
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub



Private Function CargaDatosEstadComprasCOMMAND() As Boolean
Dim vSQL As String

    On Error GoTo eCargaDatosEstadComprasCOMMAND
    CargaDatosEstadComprasCOMMAND = False
    Label9(38).Caption = "Prepara datos"
    Label9(38).Refresh
    vSQL = "DELETE FROM tmpcommand WHERE codusu = " & vUsu.Codigo
    conn.Execute vSQL
    
    cadSelect = Replace(cadSelect, "{", "")
    cadSelect = Replace(cadSelect, "}", "")
    
    Label9(38).Caption = "Facturas"
    Label9(38).Refresh
    vSQL = "insert into `tmpcommand` (`codusu`,`codprove`,`nomprove`,`cantidad`,`importel`,`codfamia`,`nomfamia`,`fecrecep`,`fechaalb`,`nomartic`,`codartic`) "
    vSQL = vSQL & " SELECT " & vUsu.Codigo & ",`scafpc`.`codprove`, `scafpc`.`nomprove`, `slifpc`.`cantidad`, `slifpc`.`importel`,"
    vSQL = vSQL & " `sartic`.`codfamia`, `sfamia`.`nomfamia`, `scafpc`.`fecrecep`, fechaalb,sartic.nomartic,"
    vSQL = vSQL & " sartic.codartic FROM   (`scafpc` `scafpc` INNER JOIN `scafpa` `scafpa`"
    vSQL = vSQL & " ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`="
    vSQL = vSQL & " `scafpa`.`fecfactu`)) AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
    vSQL = vSQL & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
    vSQL = vSQL & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
    vSQL = vSQL & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
    vSQL = vSQL & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"

    'EL where
    If cadSelect <> "" Then vSQL = vSQL & " WHERE " & cadSelect
    conn.Execute vSQL
    
    
    
    'Los albaranes
    Label9(38).Caption = "Albaranes"
    Label9(38).Refresh
    vSQL = "insert into `tmpcommand` (`codusu`,`codprove`,`nomprove`,`cantidad`,`importel`,`codfamia`,`nomfamia`,`fecrecep`,`fechaalb`,`nomartic`,`codartic`) "
    vSQL = vSQL & " SELECT " & vUsu.Codigo & ",scaalp.`codprove`, `nomprove`, `cantidad`, `importel`,"
    vSQL = vSQL & " `sartic`.`codfamia`, `sfamia`.`nomfamia`,"
    vSQL = vSQL & " scaalp.`fechaalb` fecrecep, scaalp.`fechaalb`,sartic.nomartic,sartic.codartic"
    vSQL = vSQL & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON "
    vSQL = vSQL & " ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) "
    vSQL = vSQL & " AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
    vSQL = vSQL & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
    vSQL = vSQL & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
    'el where
    cadSelect = Replace(cadSelect, "scafpa", "scaalp")
    cadSelect = Replace(cadSelect, "scafpc", "scaalp")
    cadSelect = Replace(cadSelect, "slifpa", "slialp")
    cadSelect = Replace(cadSelect, "fecfactu", "fechaalb")
    If cadSelect <> "" Then vSQL = vSQL & " WHERE " & cadSelect
    conn.Execute vSQL
    
    
    
    CargaDatosEstadComprasCOMMAND = True
    
    
    
eCargaDatosEstadComprasCOMMAND:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Label9(38).Caption = ""
End Function

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub

Private Sub CargaIconosAyuda2()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    
    imgBuscarOfer(56).Picture = imgBuscar(0).Picture
    imgBuscarOfer(57).Picture = imgBuscar(0).Picture
    
    imgBuscarOfer(85).Picture = imgBuscar(0).Picture
    imgBuscarOfer(86).Picture = imgBuscar(0).Picture
    
    Err.Clear
End Sub

Private Sub AbrirBuscaGrid(OP As Integer)
Dim indT As Integer
    Set frmB = New frmBuscaGrid
    cadFormula = "" 'Aqui metera el valor
    Select Case OP
    Case 71, 72
        indT = OP + 60
        frmB.vCampos = "Codigo|sclipot|codclien|T||20�Descripci�n|sclipot|nomclien|T||70�"
        frmB.vTabla = "sclipot"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Clientes potenciales"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    
    Case 97
        'Banocs propios
        indT = 159
        frmB.vCampos = "Codigo|sbanpr|codbanpr|T||20�Descripci�n|sbanpr|nombanpr|T||70�"
        frmB.vTabla = "sbanpr"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "BANCOS"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        
    Case 98
        indT = 161
        frmB.vCampos = "Codigo|sforpa|codforpa|T||20�Descripci�n|sforma|nomforpa|T||70�"
        frmB.vTabla = "sforpa"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Forma Pago"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    End Select
    frmB.Show vbModal
    Set frmB = Nothing
    
    
    If cadFormula <> "" Then
        
        txtCodigo(indT).Text = Format(RecuperaValor(cadFormula, 1), "0000")
        txtNombre(indT).Text = RecuperaValor(cadFormula, 2)
    
    End If
End Sub


Private Sub HacerImpresionCRM()
Dim SQL As String

    On Error GoTo eHacerImpresionCRM
    indCodigo = 0 'Indicara si cancela el preoceso de impresion
   

    
    
    Set miRsAux = New ADODB.Recordset
     
    Codigo = "Select codclien,nomclien from sclien WHERE " & cadSelect & " ORDER BY codclien"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        pbCRM.Value = pbCRM.Value + 1
        Label4(123).Caption = pbCRM.Value & " de " & pbCRM.Max
        Label4(123).Refresh
        Label4(124).Caption = miRsAux!NomClien
        Label4(124).Refresh
        
    
        'Hacemos impresion directa
    
        
        cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"

    
        GenerarDatosInformes
    
        
        'Si habian metido algun dato...
        SQL = "insert into `tmpcrmclien` (`codusu`,`codclien`,`saldopdte`,saldototal,`nomactiv`,`nomforpa`) values ("
        SQL = SQL & vUsu.Codigo & "," & miRsAux!codClien & ","
        
        'Saldo pdte (a fecha NOW
        Codigo = "Imp"
        ComprobarCobrosCliente CStr(miRsAux!codClien), Now, Codigo
        If Codigo = "" Or Codigo = "Imp" Then Codigo = "0"
        SQL = SQL & DBSet(Codigo, "N") & ","
        'saldo totoal A fecha 31/12/2222"
        Codigo = "Imp"
        ComprobarCobrosCliente CStr(miRsAux!codClien), CDate("31/12/2222"), Codigo
        If Codigo = "" Or Codigo = "Imp" Then Codigo = "0"
        SQL = SQL & DBSet(Codigo, "N") & ","
        
        
        
        Codigo = DevuelveDesdeBD(conAri, "nomactiv", "sclien,sactiv", "sclien.codactiv=sactiv.codactiv and codclien", CStr(miRsAux!codClien))
        SQL = SQL & DBSet(Codigo, "T") & ","
        Codigo = DevuelveDesdeBD(conAri, "nomforpa", "sclien,sforpa", "sclien.codforpa=sforpa.codforpa and codclien", CStr(miRsAux!codClien))
        SQL = SQL & DBSet(Codigo, "T") & ")"
        conn.Execute SQL
        
        
        
        'Vamos a fijar la cadena de parametros de impresion
        cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|pVisVolVenta=1|pDesdeAnyo=2009|pVisCobrPdte=1|"
        cadParam = cadParam & "pVisReclamas=1|pDesdeReclamas=Date(2007, 1, 7)|pVisMtos=1|"
        cadParam = cadParam & "pDesdeOferta=Date(2010, 1, 1)|pDesdepedido=Date(2009, 1, 1)|"
        cadParam = cadParam & "pDesdeAlbaran=Date(2010, 1, 1)|pVisAccionesComer=1|pDesdeAccComer=Date(2010, 3, 1)|"
        cadParam = cadParam & "pVisLlamadas=0|pDesdeLlamada=Date(2010, 3, 1)|pVisEmails=0|pDesdeEmail=Date(2010, 1, 1)|"
        cadParam = cadParam & "pVisFreq=0|pVisAlbSat=0|pVisAvisos=0|pVisReparas=0|"
        NumRegElim = 20
        With frmImprimir
            .FormulaSeleccion = "{tmpcrmclien.codusu} = " & vUsu.Codigo
  
            .OtrosParametros = cadParam
            .NumeroParametros = NumRegElim
    
            .SoloImprimir = True
            .EnvioEMail = False
            .Opcion = 2018
            .Titulo = "CRM" & miRsAux!NomClien
            .NombreRPT = nomRPT
            .NombrePDF = ""
            'If PonerNombrePDF Then .NombrePDF = pPdfRpt
            .ConSubInforme = True
            .Show vbModal
        End With
        
        Me.Refresh
        Espera 0.1
        DoEvents
        
        For NumRegElim = 1 To 10
            Screen.MousePointer = vbHourglass
            DoEvents
            Espera 0.1
        Next
        
        
        
        If indCodigo = 1 Then
            'Han cancelado
            While Not miRsAux.EOF
                miRsAux.MoveNext
            Wend
        Else
            miRsAux.MoveNext
        End If
        
    Wend
    miRsAux.Close
    
    
eHacerImpresionCRM:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub



Private Function GenerarDatosInformes() As Boolean
Dim vCRM As cCRM
Dim Impor1 As Currency
Dim Base As Currency
Dim Cad As String
Dim Aux As String
Dim F As Date
Dim RS As ADODB.Recordset
Dim SQL As String
Dim J As Integer


    On Error GoTo eGenerarDatosInformes
    GenerarDatosInformes = False
    Set vCRM = New cCRM
    Set RS = New ADODB.Recordset
    vCRM.BorrarTemporales
    vCRM.codClien = miRsAux!codClien
    vCRM.Codmacta = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", vCRM.codClien)
    conn.Execute "commit"  'de mysql
    Espera 0.3
    
    
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'A d m i n i s t r a c i o n

    Titulo = "Volumen fact."
    
    'Volumen facturacion
    SQL = "select year(fecfactu) anyo,sum(totalfac) totalfac from scafac "
    'SEPTIEMBE 2011. Quito FRT del select
    'SQL = SQL & " where codclien=" & cstr(mirsaux!codclien) & " and codtipom <>'FAZ' and codtipom<>'FRT' "
    SQL = SQL & " where codclien=" & CStr(miRsAux!codClien) & " and codtipom <>'FAZ'"
    SQL = SQL & " AND fecfactu>='" & Format(F, FormatoFecha) & "'"
    'Aqui va lo de ultimos a�os
    SQL = SQL & " group by 1 order by 1,2"
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    
    While Not RS.EOF
        Cad = ""
    
        NumRegElim = NumRegElim + 1
        Impor1 = DBLet(RS!TotalFac, "N")
        
        SQL = "insert into `tmpcrmtesor` (`codusu`,`codigo`,`importe`,`anyotxt`,`variacion`)"
        SQL = SQL & " values (" & vUsu.Codigo & "," & NumRegElim & "," & TransformaComasPuntos(CStr(Impor1)) & ",'"
        
        If Val(RS!Anyo) = Year(Now) Then
            'Valor actual.
            SQL = SQL & "actual',"
            'Cambio la base para comprar con el mismo periodo del actual
            
            'Cad = "codtipom <>'FAZ' and codtipom<>'FRT' and "
            Cad = "codtipom <>'FAZ' and "
            Cad = Cad & " fecfactu>='" & Year(Now) - 1 & "-01-01' and "
            Cad = Cad & " fecfactu<='" & Year(Now) - 1 & "-" & Format(Now, "mm-dd") & "' AND codclien "
            Cad = DevuelveDesdeBD(conAri, "sum(totalfac)", "scafac", Cad, CStr(miRsAux!codClien))
            If Cad = "" Then Cad = "0"
            Base = CCur(Cad)
            If NumRegElim > 1 And Base <> 0 Then
                Impor1 = CStr(((100 * Impor1) / Base) - 100)
                Cad = Format(Impor1, FormatoPorcen) & "% sobre misma fecha a�o anterior"
            Else
                Cad = ""
            End If
        Else
            'Otro a�o cualquiera
             SQL = SQL & RS!Anyo & "',"
            If NumRegElim > 1 And Base <> 0 Then
                Impor1 = CStr(((100 * Impor1) / Base) - 100)
                Cad = Format(Impor1, FormatoPorcen) & "%"
            End If
             
        End If
        Base = DBLet(RS!TotalFac, "N")
        SQL = SQL & "'" & Cad & "')"
      

        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    
    
    
    Titulo = "Cobros pendientes"
    'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
    If vParamAplic.ContabilidadNueva Then
        
        SQL = "SELECT cobros.*,numfactu Codfaccl,fecfactu fecfaccl ,nomforpa FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        SQL = SQL & " WHERE cobros.codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & "  AND recedocu=0 ORDER BY fecvenci desc"
    
    Else
        SQL = "SELECT scobro.*,nomforpa FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & "  AND recedocu=0 ORDER BY fecvenci desc"
    End If
    NumRegElim = 0
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Base = 0
    Impor1 = 0
    
    While Not RS.EOF
          'trozo copiado d ela funcion de ver cobros pdtes
      If DBLet(RS!Devuelto, "N") = 1 Then
            'SALE SEGURO (si no esta girado otra vez �no?
            'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
            Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
            
        Else
            'Si esta recibido NO lo saco
            If Val(RS!recedocu) = 1 Then
                Impor1 = 0
            Else
                'NO esta recibido. Si tiene diferencia
                Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
        
            End If
      End If
      If Impor1 <> 0 Then
            NumRegElim = NumRegElim + 1
            SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
            SQL = SQL & "`importe`,`observa`,forpa) values ( "
            SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",0,'"
            SQL = SQL & RS!numSerie & Format(RS!Codfaccl, "000000")
            If RS!FecVenci < Now Then SQL = SQL & " *"
            SQL = SQL & "','" & Format(RS!fecfaccl, FormatoFecha)
            SQL = SQL & "','" & Format(RS!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
            'Antes la observa era NULL, ahora llevare el Depto
            If IsNull(RS!Departamento) Then
                Aux = "NULL"
            Else
                Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(RS!Departamento))
                If Aux = "" Then Aux = RS!Departamento
                Aux = "'" & DevNombreSQL(Aux) & "'"
                
            End If
            SQL = SQL & Aux
            'Mayo 2010
            'Con forma de pago
            SQL = SQL & ",'" & Format(RS!codforpa, "000") & " - " & DevNombreSQL(RS!nomforpa) & "')"
            conn.Execute SQL
      End If
      RS.MoveNext

        
    
    Wend
    RS.Close
        
        
    'Marzo 2011
    'Tambien sacare el riesgo. Habra que configurar el rpt de cada uno
    '----------------------------------------------------------------
    Titulo = "Riesgo tesoreria"
    'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
    If vParamAplic.ContabilidadNueva Then
        
        SQL = "SELECT cobros.*,numfactu Codfaccl,fecfactu fecfaccl ,nomforpa FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        SQL = SQL & " WHERE cobros.codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND (formapago.tipforpa between 2 and 5) "
        SQL = SQL & " AND impcobro>0 ORDER BY fecvenci desc"
    
    Else
        SQL = "SELECT scobro.*,nomforpa FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND (sforpa.tipforpa between 2 and 5) "
        SQL = SQL & " AND impcobro>0 ORDER BY fecvenci desc"
    End If
    J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Base = 0
    Impor1 = 0
    
    While Not RS.EOF
    'trozo copiado d ela funcion de ver cobros pdtes
      
            'NO esta recibido. Si tiene diferencia
            Impor1 = RS!impcobro
            NumRegElim = NumRegElim + 1
            SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
            SQL = SQL & "`importe`,`observa`,forpa) values ( "
            SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",2,'"    '2.  El 2 es RIESGO para el rpt
            SQL = SQL & RS!numSerie & Format(RS!Codfaccl, "000000")
            If RS!FecVenci < Now Then SQL = SQL & " *"
            SQL = SQL & "','" & Format(RS!fecfaccl, FormatoFecha)
            SQL = SQL & "','" & Format(RS!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
            'Antes la observa era NULL, ahora llevare el Depto
            If IsNull(RS!Departamento) Then
                Aux = "NULL"
            Else
                Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(RS!Departamento))
                If Aux = "" Then Aux = RS!Departamento
                Aux = "'" & DevNombreSQL(Aux) & "'"
                
            End If
            SQL = SQL & Aux
            'Mayo 2010
            'Con forma de pago
            SQL = SQL & ",'" & Format(RS!codforpa, "000") & " - " & DevNombreSQL(RS!nomforpa) & "')"
            conn.Execute SQL
            RS.MoveNext

        
    
    Wend
    RS.Close
        
 
    Titulo = "Hco reclamas"
    If vParamAplic.ContabilidadNueva Then
    
        SQL = "SELECT reclama.codigo,numserie,numfactu codfaccl,fecfactu fecfaccl,fecreclama,impvenci,codmacta,observaciones "
        SQL = SQL & " FROM reclama INNER JOIN reclama_facturas  ON reclama.codigo =reclama_facturas.codigo"
        SQL = SQL & " WHERE codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' "
    
    Else
        SQL = "SELECT codigo,numserie,codfaccl,fecfaccl,fecreclama,impvenci,codmacta,observaciones from shcocob "
        SQL = SQL & " WHERE codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' "
        'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
    End If
    
    J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
    
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ("
        SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",1,'"
        SQL = SQL & DBLet(RS!numSerie, "T") & Format(DBLet(RS!Codfaccl, "N"), "000000") & "','"
        SQL = SQL & Format(RS!fecfaccl, FormatoFecha) & "','" & Format(RS!fecreclama, FormatoFecha) & "',"
        SQL = SQL & TransformaComasPuntos(RS!ImpVenci) & ",'"
        Cad = DBLetMemo(RS!Observaciones)
        Cad = Replace(Cad, vbCrLf, " ")
        SQL = SQL & DevNombreSQL(Cad) & "')"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    
   
 
    SQL = "Select count(*) from scaman where codclien = " & CStr(miRsAux!codClien)
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not RS.EOF Then NumRegElim = DBLet(RS.Fields(0), "N")
    RS.Close

    
    
    






eGenerarDatosInformes:
    If Err.Number <> 0 Then MuestraError Err.Number, Titulo
    Set vCRM = Nothing
    Set RS = Nothing
End Function



Private Function AnadirClientesCobrosPendientes() As String
Dim SQ As String
    AnadirClientesCobrosPendientes = ""
    On Error GoTo eAnadirClientesCobrosPendientes
    SQ = "select distinct(codmacta) from scobro where recedocu=0 and codrem is null "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQ, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQ = ""
    While Not miRsAux.EOF
        SQ = SQ & ", '" & miRsAux!Codmacta & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    If SQ <> "" Then
        SQ = Mid(SQ, 2)
        AnadirClientesCobrosPendientes = " AND codmacta IN (" & SQ & ")"
    End If
eAnadirClientesCobrosPendientes:
    If Err.Number <> 0 Then MuestraError Err.Number, "Leyendo clientes pendiente cobros"
    Set miRsAux = Nothing
End Function




Private Sub CargaTipoMov()
Dim IT As ListItem

    lwFact.ListItems.Clear
    
    
    Codigo = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%'"
    'Para cualquiera menos root
    If (vUsu.Codigo Mod 1000) > 0 Then
        Codigo = Codigo & " AND codtipom"
        If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
            Codigo = Codigo & " = "
        Else
            Codigo = Codigo & "<>"
        End If
        Codigo = Codigo & "'FAZ'"
    End If
        
    'FTG y FLQ NO SALEN
    Codigo = Codigo & " AND codtipom <> 'FTG' AND codtipom <> 'FLQ'"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = Me.lwFact.ListItems.Add()
        IT.Text = miRsAux!codtipom
        IT.SubItems(1) = miRsAux!nomtipom
        IT.Checked = True
        

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub cargaDocumentos()
Dim i As Integer
    Me.ListView2.ListItems.Clear
    
    If Trim(txtCodigo(1).Text) = "" Then Exit Sub
    If Val(txtCodigo(1).Text) = 0 Then Exit Sub
    
'    Set miRsAux = New ADODB.Recordset
'    codigo = "sliprePdfs"
    'If EsHistorico Then txtAnterior = "slhprePdfs"
'
'    codigo = "Select * from " & codigo & " WHERE numofert =" & Val(txtCodigo(1).Text) & " ORDER BY numlinea"
'    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    i = 0
'    While Not miRsAux.EOF
'        i = i + 1
'        Me.ListView2.ListItems.Add , "C" & miRsAux!numlinea, miRsAux!ficheroDesc
'        Me.ListView2.ListItems(i).SubItems(1) = miRsAux!ficheronombre
'        ListView2.ListItems(i).Checked = False
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
    If davidCodtipom = "" Then Exit Sub
    
    'codigo = Dir(davidCodtipom & "*.pdf", vbDirectory)   ' Recupera la primera entrada.
    Codigo = Dir(davidCodtipom & "*.*", vbDirectory)   ' Los ponemos todos por si acaso ENVIA por email. Imprimiendo solo PDFs
    Do While Codigo <> ""   ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If Codigo <> "." And Codigo <> ".." Then
           ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
           If (GetAttr(davidCodtipom & Codigo) And vbDirectory) <> vbDirectory Then
                'Debug.Print MiNombre   ' Muestra la entrada
                
                 i = i + 1
                 Me.ListView2.ListItems.Add , "C" & i, Codigo
                 
                 Me.ListView2.ListItems(i).SubItems(1) = Codigo
                 
           End If   ' solamente si representa un directorio.
        End If
        Codigo = Dir   ' Obtiene siguiente entrada.
    Loop

    
    
End Sub

Private Sub CargaTipoPagoInactivos()
    Set miRsAux = New ADODB.Recordset
    lwTippa.ListItems.Clear
    miRsAux.Open "Select tipforpa,destippa from stippa order by 1", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        lwTippa.ListItems.Add , "C" & miRsAux!tipforpa, miRsAux!destippa
        lwTippa.ListItems(NumRegElim).Checked = True
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'---------------------------------------------------------------------------------------
' Factura rectificativa completa
Private Function GeneraFacturaRectificativaCompleta() As Boolean
Dim NF As CTiposMov
Dim NA As CTiposMov
Dim J As Integer
Dim Tra As Long
Dim Cad As String
Dim Cliente As Long


    On Error GoTo eGeneraFacturaRectificativaCompleta
    GeneraFacturaRectificativaCompleta = False
    
    Set NF = New CTiposMov
    Set NA = New CTiposMov
    If Not NF.Leer("FRT") Then Err.Raise 513, , "Error obteniendo contador FRT"
    If Not NA.Leer("ART") Then Err.Raise 513, , "Error obteniendo contador ART"
    
    If NF.ConseguirContador(NF.TipoMovimiento) = 0 Then Err.Raise 513, , "Error obteniendo numero FRT"
    If NA.ConseguirContador(NA.TipoMovimiento) = 0 Then Err.Raise 513, , "Error obteniendo numero FRT"
    NF.IncrementarContador NF.TipoMovimiento
    Codigo = PonerTrabajadorConectado(Cad)
    If Codigo = "" Then Err.Raise 513, , "Error obteniendo trabajador conectado"
    Tra = Val(Codigo)
    
    
    
    Codigo = " codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F") & " AND 1"
    Codigo = DevuelveDesdeBD(conAri, "codclien", "scafac", Codigo, "1")
    If Codigo = "" Then Err.Raise 513, , "Error obteniendo cliente factura"
    Cliente = Val(Codigo)
    
    'Objetivo 1.
    'INSERTAREMOS en scafac,scafac1,svenci,slifac,slifaclotes
    '
    Label4(133).Caption = "Insertar factura"
    Label4(133).Refresh
    Codigo = "INSERT INTO scafac ( codtipom ,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien"
    Codigo = Codigo & ",coddirec,codagent,codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba"
    Codigo = Codigo & ",brutofac,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3,"
    Codigo = Codigo & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,"
    Codigo = Codigo & "imporiv1,imporiv2,imporiv3,totalfac,intconta , porciva1re, porciva2re, porciva3re, "
    Codigo = Codigo & "imporiv1re, imporiv2re, imporiv3re, aridoc, Aportacion, pesoalba, portes, EnFacturaE, Iban) "
    
    
    Codigo = Codigo & " SELECT " & DBSet(NF.TipoMovimiento, "T") & "," & NF.Contador & "," & DBSet(txtCodigo(160).Text, "F")
    Codigo = Codigo & ",codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien"
    Codigo = Codigo & ",coddirec,codagent," & txtCodigo(161).Text & " codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba"
    Codigo = Codigo & ",-brutofac,impdtopp,impdtogr,-baseimp1,-baseimp2,-baseimp3,"
    Codigo = Codigo & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,"
    Codigo = Codigo & "-imporiv1,-imporiv2,-imporiv3,-totalfac,0 intconta , porciva1re, porciva2re, porciva3re, "
    Codigo = Codigo & "-imporiv1re, -imporiv2re, -imporiv3re, aridoc, -Aportacion, pesoalba, portes, EnFacturaE, Iban "
    Codigo = Codigo & " FROM scafac where codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
    conn.Execute Codigo
    
    
    
    
    'Iremos por ALBARAN (albar , linea, linealotes
    Set miRsAux = New ADODB.Recordset
    Codigo = "Select numalbar,codtipoa from scafac1 where  codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F") & " ORDER BY numalbar"
    
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 1
    Cad = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    Cad = "RECTIFICA A FACTURA: " & Cad & ", " & Val(txtCodigo(71).Text) & ", " & txtCodigo(72).Text

    While Not miRsAux.EOF
        Label4(133).Caption = "Albaran/linea factura"
        Label4(133).Refresh
        'SCAFAC1
        Codigo = "INSERT INTO scafac1(codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,codenvio,codtraba,codtrab1,codtrab2,"
        Codigo = Codigo & " observa1,observa2,coddiren,tipAlbaran,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre"
        Codigo = Codigo & ",TipoCarnet,fechaAux) "
        
        Codigo = Codigo & " SELECT " & DBSet(NF.TipoMovimiento, "T") & "," & NF.Contador & "," & DBSet(txtCodigo(160).Text, "F") & ","
        Codigo = Codigo & "'ART'," & NA.Contador + J & "," & DBSet(txtCodigo(160).Text, "F") & ","
        Codigo = Codigo & "codenvio, " & Tra & " as codtraba,codtrab1,codtrab2,"
        Codigo = Codigo & DBSet(Cad, "T") & " as observa1," & DBSet("MOTIVO: " & txtCodigo(87).Text, "T") & " as observa2,coddiren,tipAlbaran,"
        Codigo = Codigo & "ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,tipocarnet ," & DBSet(txtCodigo(160).Text, "F")
        
        Codigo = Codigo & " FROM scafac1 where  codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
        Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
        Codigo = Codigo & " AND numalbar =" & miRsAux!Numalbar & " AND codtipoa =" & DBSet(miRsAux!Codtipoa, "T")
        conn.Execute Codigo
        
        If vParamAplic.Ariagro <> "" Then
            Codigo = "INSERT INTO slifaccampos( codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codcampo,fechahora,nomvarie,nompartida)"
            Codigo = Codigo & " SELECT " & DBSet(NF.TipoMovimiento, "T") & "," & NF.Contador & "," & DBSet(txtCodigo(160).Text, "F") & ","
            Codigo = Codigo & "'ART'," & NA.Contador + J & ",numlinea,codcampo,fechahora,nomvarie,nompartida"
            Codigo = Codigo & " FROM slifaccampos WHERE codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
            Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
            Codigo = Codigo & " AND numalbar =" & miRsAux!Numalbar & " AND codtipoa =" & DBSet(miRsAux!Codtipoa, "T")
             conn.Execute Codigo
        End If
        
        'SLIFAC
        Codigo = "INSERT INTO slifac(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci"
        Codigo = Codigo & ",cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,precioiv,preciomp,preciost,preciouc,"
        Codigo = Codigo & " codproveX,numlote,codccost,codtipor,codcapit,precoste,codtraba,pvpInferior,comisionagente)"
        
        Codigo = Codigo & " SELECT " & DBSet(NF.TipoMovimiento, "T") & "," & NF.Contador & "," & DBSet(txtCodigo(160).Text, "F") & ","
        Codigo = Codigo & "'ART'," & NA.Contador + J & ", numlinea ,codalmac,codartic,nomartic,ampliaci, -cantidad, numbultos,"
        Codigo = Codigo & "precioar,dtoline1,dtoline2,-importel,origpre,precioiv,preciomp,preciost,preciouc,"
        Codigo = Codigo & " codproveX,numlote,codccost,codtipor,codcapit,precoste,codtraba,pvpInferior,comisionagente FROM slifac "
        Codigo = Codigo & " where  codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
        Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
        Codigo = Codigo & " AND numalbar =" & miRsAux!Numalbar & " AND codtipoa =" & DBSet(miRsAux!Codtipoa, "T")
        conn.Execute Codigo
        
        
        'smoval
        Label4(133).Caption = "Insertar smoval"
        Label4(133).Refresh
        
        
        Codigo = "INSERT INTO smoval( codArtic , codAlmac, FechaMov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, LetraSer, document, numlinea)"
        Codigo = Codigo & " SELECT codartic,codalmac," & DBSet(txtCodigo(160).Text, "F") & ","
        Codigo = Codigo & DBSet(txtCodigo(160).Text & " " & Format(Now, "hh:nn:ss"), "FH") & ",0,'ART',-cantidad,-importel," & Cliente
        Codigo = Codigo & ", null, '" & Format(NA.Contador + J, "0000000") & "', numlinea"
        Codigo = Codigo & " FROM slifac where  codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
        Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
        Codigo = Codigo & " AND numalbar =" & miRsAux!Numalbar & " AND codtipoa =" & DBSet(miRsAux!Codtipoa, "T")
        conn.Execute Codigo
        
       
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Codigo = " INSERT INTO slifaclotes (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic) "
            
            Codigo = Codigo & " SELECT " & DBSet(NF.TipoMovimiento, "T") & "," & NF.Contador & "," & DBSet(txtCodigo(160).Text, "F") & ","
            Codigo = Codigo & "'ART'," & NA.Contador + J & ",numlinea,sublinea,-cantidad,numlote,fecentra,codartic FROM slifaclotes "
            Codigo = Codigo & " where  codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
            Codigo = Codigo & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
            Codigo = Codigo & " AND numalbar =" & miRsAux!Numalbar & " AND codtipoa =" & DBSet(miRsAux!Codtipoa, "T")
            
            conn.Execute Codigo
        
        
        
        
        End If
        
        
                
        
        
        
        J = J + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Label4(133).Caption = "Sotcks-Lotes"
    Label4(133).Refresh
    
    'Ajustaremos los lotes, si procede
    If vParamAplic.ManipuladorFitosanitarios2 Then
        Cad = "Select numlote,codartic,fecentra,sum(cantidad) cantidad from slifaclotes "
        Cad = Cad & " where codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
        Cad = Cad & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
        Cad = Cad & " GROUP BY 1,2,3"
        
        
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = "UPDATE slotes set vendida=vendida "
            If miRsAux!cantidad < 0 Then
                Cad = Cad & " + "
            Else
                Cad = Cad & " - "
            End If
            Cad = Cad & DBSet(Abs(miRsAux!cantidad), "N") & " WHERE  codartic=" & DBSet(miRsAux!codArtic, "T")
            Cad = Cad & "and fecentra=" & DBSet(miRsAux!fecentra, "F") & " and numlotes=" & DBSet(miRsAux!numLote, "T")
            conn.Execute Cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    'Ajustaremos canstock en SALMAC, si procede
    Cad = "select  codalmac,slifac.codartic,sum(cantidad) cantidad"
    Cad = Cad & " from slifac left join sartic on slifac.codartic=sartic.codartic  where artvario=0  "
    Cad = Cad & " and codtipom=" & DBSet(Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    Cad = Cad & " AND numfactu =" & txtCodigo(71).Text & " AND fecfactu =" & DBSet(txtCodigo(72).Text, "F")
    Cad = Cad & " group by 1,2"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "UPDATE salmac set canstock = canstock  "
        If miRsAux!cantidad < 0 Then
            Cad = Cad & " - "
        Else
            Cad = Cad & " + "
        End If
        Cad = Cad & DBSet(Abs(miRsAux!cantidad), "N") & " WHERE  codartic=" & DBSet(miRsAux!codArtic, "T")
        Cad = Cad & "and codalmac=" & miRsAux!codAlmac
        conn.Execute Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    cadFormula = "{scafac.codtipom}='FRT' AND {scafac.numfactu}=" & NF.Contador & " AND {scafac.fecfactu}= Date(" & Format(CDate(txtCodigo(160).Text), "yyyy,mm,dd") & ")"
    NumCod = Format(NF.Contador, "0000000")
    
    'Actualizmos contadores
    
    NA.Contador = NA.Contador + (J - 2)  'al incremetar
    NA.IncrementarContador NA.TipoMovimiento
    
    
    
    'Para la impresion posterior  indRPT = 12
    PonerParamRPT2 12, cadParam, numParam, nomRPT, False, pPdfRpt, pRptvMultiInforme
    
    
    
    GeneraFacturaRectificativaCompleta = True
    
    
eGeneraFacturaRectificativaCompleta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set NF = Nothing
    Set NA = Nothing
    Set miRsAux = Nothing
    Label4(133).Caption = ""
    
End Function

Private Function FechaLiquidacionIVA() As String
    
    FechaLiquidacionIVA = ""
    cadFormula = "concat(anofactu,'|',perfactu,'|')"
    
    nomRPT = DevuelveDesdeBD(conConta, "periodos", "Parametros", "1", "1", "N", cadFormula)
    If nomRPT <> "" Then
        indCodigo = 1
        If nomRPT = "0" Then indCodigo = 3
        Titulo = RecuperaValor(cadFormula, 2)
        'MEs
        If Titulo <> "" Then
            Titulo = CStr(CByte(Titulo) * indCodigo)
            cadFormula = RecuperaValor(cadFormula, 1)
            indCodigo = DiasMes(CByte(Titulo), CInt(cadFormula))
            Titulo = indCodigo & "/" & Format(Titulo, "00") & "/" & cadFormula
            If CDate(Titulo) > CDate(txtCodigo(160).Text) Then FechaLiquidacionIVA = "El periodo de facturacion del IVA ya esta cerrado."
                
            
        End If
    End If
    
End Function


Private Sub DatosFraRectificativa()
    If Me.chkFraRectCompleta.Value = 0 Then Exit Sub
    
    'SI estan los datos de la factura
    If Me.cboTipomov(0).ListIndex < 0 Then Exit Sub
    If Me.txtCodigo(71).Text = "" Then Exit Sub
    If Me.txtCodigo(72).Text = "" Then Exit Sub
    
    'SI ya ha puesto datos, no los toco
    If Me.txtCodigo(161).Text <> "" Then Exit Sub
    
    
    
    
    Codigo = "numfactu =" & txtCodigo(71).Text & " AND fecfactu=" & DBSet(txtCodigo(72).Text, "F") & " AND codtipom "
    Codigo = DevuelveDesdeBD(conAri, "codforpa", "scafac", Codigo, Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3), "T")
    If Codigo <> "" Then
        txtCodigo(161).Text = Codigo
        txtNombre(161).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtCodigo(161).Text)
    End If
End Sub


Private Sub CargaMotivosBaja()
    Codigo = "select * from  smotitrasofer order by 2"
    Set miRsAux = New ADODB.Recordset
    lwMotBaja.ListItems.Clear
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroDeCopias = 0
    While Not miRsAux.EOF
        NumeroDeCopias = NumeroDeCopias + 1
        lwMotBaja.ListItems.Add , "C" & Format(miRsAux!codmotiv, "000"), DBLet(miRsAux!descmotivo, "T")
        lwMotBaja.ListItems(NumeroDeCopias).Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    NumeroDeCopias = 0
End Sub

Private Sub PonerPb2()
Dim jj As Integer
    jj = Round((SecuenciaEnProgressBar1 * 1000) / MaxNumerRegistroSEnProgressBar1, 0)
    If jj > 1000 Then jj = 1000
    Me.ProgressBar1.Value = jj
End Sub
