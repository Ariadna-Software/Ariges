VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComprasVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociar albaranes compras / ventas"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   45
      Top             =   8040
      Width           =   3735
      Begin VB.Label lblInd 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4920
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Selección de datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   7935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8055
      Begin VB.Frame FrameOrdenProv 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2400
         TabIndex        =   70
         Top             =   6120
         Width           =   5295
         Begin VB.OptionButton optOrdenProv 
            Caption         =   "Nº Albaran  / Fecha "
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   72
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optOrdenProv 
            Caption         =   "Fecha / Nº Albaran"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   71
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Mostrar todos"
         Height          =   195
         Left            =   3960
         TabIndex        =   58
         Top             =   7080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Artículo  / proveedor"
         Height          =   195
         Index           =   1
         Left            =   5160
         TabIndex        =   43
         Top             =   5880
         Width           =   1815
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Proveedor / Artículo"
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   42
         Top             =   5880
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdBus 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   11
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton cmdBus 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   10
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   5880
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   5880
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   4440
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   4080
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenacion proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   69
         Top             =   6360
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenacion ppal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   68
         Top             =   5880
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Index           =   3
         Left            =   3000
         TabIndex        =   67
         Top             =   120
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Index           =   2
         X1              =   7680
         X2              =   1440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00000080&
         Height          =   7695
         Left            =   120
         Top             =   0
         Width           =   7815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmComprasVentas.frx":0000
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "frmComprasVentas.frx":0102
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":0204
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":0306
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":0408
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":050A
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Index           =   13
         Left            =   360
         TabIndex        =   41
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
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
         Index           =   12
         Left            =   360
         TabIndex        =   40
         Top             =   4080
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   5520
         Picture         =   "frmComprasVentas.frx":060C
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   39
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   38
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":0697
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Albaranes clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   3720
         Width           =   1755
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Index           =   1
         X1              =   7680
         X2              =   1680
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label2 
         Caption         =   "Albaranes proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   2235
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Index           =   0
         X1              =   7680
         X2              =   1440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Index           =   9
         Left            =   360
         TabIndex        =   37
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
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
         Index           =   8
         Left            =   360
         TabIndex        =   36
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   5520
         Picture         =   "frmComprasVentas.frx":0722
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   35
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   34
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   33
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   32
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   31
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   30
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   915
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmComprasVentas.frx":07AD
         Top             =   3120
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   8175
      Begin VB.Frame Frame44 
         Height          =   1575
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   7815
         Begin VB.Frame FrameDes 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   3600
            TabIndex        =   61
            Top             =   120
            Visible         =   0   'False
            Width           =   3975
            Begin VB.CommandButton cmdDes 
               Height          =   375
               Index           =   3
               Left            =   3600
               Picture         =   "frmComprasVentas.frx":0838
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Ultimo"
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmdDes 
               Height          =   375
               Index           =   2
               Left            =   3120
               Picture         =   "frmComprasVentas.frx":0DC2
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Siguiente"
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmdDes 
               Height          =   375
               Index           =   1
               Left            =   2640
               Picture         =   "frmComprasVentas.frx":134C
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Anterior"
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmdDes 
               Height          =   375
               Index           =   0
               Left            =   2160
               Picture         =   "frmComprasVentas.frx":18D6
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Primero"
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Label5"
               Height          =   255
               Left            =   360
               TabIndex        =   66
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            TabIndex        =   59
            Text            =   "Text4"
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton cmdCancelar 
            Height          =   375
            Left            =   6360
            Picture         =   "frmComprasVentas.frx":1E60
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Cancelar"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdGuardar 
            Height          =   375
            Left            =   5880
            Picture         =   "frmComprasVentas.frx":2862
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Aceptar"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   4320
            TabIndex        =   55
            Text            =   "Text3"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2880
            TabIndex        =   53
            Text            =   "Text3"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   1440
            TabIndex        =   49
            Text            =   "Text3"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Text            =   "Text3"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Datos albaran compra"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label4 
            Caption         =   "Disponible"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   54
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "VENTAS"
            Height          =   255
            Index           =   3
            Left            =   4800
            TabIndex        =   52
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Asociadas"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   51
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "En albaran"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   600
         Width           =   5055
      End
      Begin MSComctlLib.ListView lwV 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8281
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cliente"
            Object.Width           =   4146
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mov"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Albaran"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lin"
            Object.Width           =   838
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cant."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Asig."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Dispo."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Aportacion"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Articulo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar asignaciones"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar albaran venta asignado"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6240
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
End
Attribute VB_Name = "frmComprasVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                
Private Const vbVerde = &H8000&       'Indicara que se puede asignar
Private Const vbMorado = &H800080     'Indicara que es el mismo albaran


Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmPro As frmComProveedores
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents FrmB As frmBuscaGrid
Attribute FrmB.VB_VarHelpID = -1


Private DatosModificados As Boolean
Private IndiceTxt As Integer
Dim Impo As Currency





Private Sub cmdBus_Click(Index As Integer)
Dim b As Boolean
    Screen.MousePointer = vbHourglass
    lblInd.Caption = "Cargando datos prov / art"
    lblInd.Refresh
    If Index = 0 Then
        b = Cargar
    Else
        b = True
    End If
    If b Then
        Frame2.visible = False
        Frame1.Enabled = True
    End If
    PonerLbl
    
    
    NumRegElim = 1
    b = False
    If Not Adodc1.Recordset Is Nothing Then
        If Not Adodc1.Recordset.EOF Then
            b = True
            'CargaAlbaranesCompras
            If Adodc1.Recordset.RecordCount > 1 Then NumRegElim = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    DesplazamientoVisible Me.Toolbar1, 17, b, CByte(NumRegElim)
    
    

    
    Screen.MousePointer = vbDefault
End Sub


Private Function Cargar() As Boolean
Dim SQL As String

    On Error GoTo ECargar
        Cargar = False
        SQL = ""
        If txtCodigo(4).Text <> "" Then SQL = SQL & " AND codartic >= '" & DevNombreSQL(txtCodigo(4).Text) & "'"
        If txtCodigo(5).Text <> "" Then SQL = SQL & " AND codartic <= '" & DevNombreSQL(txtCodigo(5).Text) & "'"
        
        If txtCodigo(0).Text <> "" Then SQL = SQL & " AND codprove >= " & txtCodigo(0).Text
        If txtCodigo(1).Text <> "" Then SQL = SQL & " AND codprove <= " & txtCodigo(1).Text
        
        
        If txtFecha(0).Text <> "" Then SQL = SQL & " AND fechaalb >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then SQL = SQL & " AND fechaalb <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        If SQL <> "" Then
            SQL = Mid(SQL, 5) 'QUITO EL PRIMER AND
            SQL = " WHERE " & SQL
        End If
        
        SQL = "Select codprove,codartic from slialp  " & SQL
        
        'Group y order
        SQL = SQL & " GROUP BY codprove,codartic ORDER BY "
        If optOrden(0).Value Then
            SQL = SQL & "codprove,codartic "
        Else
            SQL = SQL & "codartic,codprove "
        End If
        
        Adodc1.ConnectionString = conn
        Adodc1.RecordSource = SQL
        Adodc1.Refresh
        If Not Adodc1.Recordset.EOF Then
            Cargar = True
            PonerCampos
        Else
            MsgBox "Ningun dato con estos valores", vbExclamation
        End If
        
        Exit Function
        
ECargar:
    MuestraError Err.Number

End Function

Private Sub cmdCancelar_Click()

    Screen.MousePointer = vbHourglass
   
    'No, cargamos solamente el cargalist, y los datos
    PonerCamposCompras
    PonerFocoBtn Me.cmdCancelar
    
    Screen.MousePointer = vbDefault
    DatosModificados = False
   
End Sub

Private Sub cmdDes_Click(Index As Integer)
    If DatosModificados Then
        If MsgBox("Ha modificado los datos. Si continua perdara los cambios. " & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        DatosModificados = False
    End If
    Select Case Index
    Case 0
        Adodc2.Recordset.MoveFirst
    Case 1
        Adodc2.Recordset.MovePrevious
        If Adodc2.Recordset.BOF Then Adodc2.Recordset.MoveFirst
    Case 2
        Adodc2.Recordset.MoveNext
        If Adodc2.Recordset.EOF Then Adodc2.Recordset.MoveLast
    Case 3
        Adodc2.Recordset.MoveLast
    End Select
    PonerCamposCompras
    
    PonerLbl
    
End Sub

Private Sub cmdGuardar_Click()
Dim SQL As String
Dim C2 As String

    If Not DatosModificados Then Exit Sub

    'Guardar
    'Insertaremos una linea por cada registro en la tabla
    If Adodc2.Recordset.EOF Then Exit Sub
    
    
    
    Screen.MousePointer = vbHourglass
    
    '`numalbarc`,`fechaalbc`,`codprovec`,`numlineac` & ","
    C2 = "insert into `slcomven` (`numalbarc`,`fechaalbc`,`codprovec`,`numlineac`,`codartic`,`codtipom`,`numalbarv`,`numlineav`,`fechaalbv`,`cantidad`) VALUES ("
    C2 = C2 & "'" & Adodc2.Recordset!NumAlbar & "','" & Format(Adodc2.Recordset!FechaAlb, FormatoFecha) & "'," & Adodc1.Recordset!codProve & ","
    C2 = C2 & Adodc2.Recordset!numlinea & ",'" & DevNombreSQL(Adodc1.Recordset!codArtic) & "',"
       


    For NumRegElim = 1 To lwV.ListItems.Count
        'Si esta en negrita es que esta seelccionado
        If lwV.ListItems(NumRegElim).ListSubItems(7).Bold Then
            
                With lwV.ListItems(NumRegElim)
                    SQL = "'" & .SubItems(1) & "'," & .SubItems(2) & "," & .SubItems(4) & ",'" & Format(.SubItems(3), FormatoFecha) & "',"
                End With

                'CAntidad
                Impo = CCur(lwV.ListItems(NumRegElim).ListSubItems(8))
                SQL = C2 & SQL & TransformaComasPuntos(CStr(Impo)) & ")"
                conn.Execute SQL
                


        End If
     Next
     
     Me.lblInd.Caption = "Consulta DB"
     Me.Refresh
     DatosModificados = False
     
     SQL = Adodc2.Recordset!NumAlbar & "|" & Format(Adodc2.Recordset!FechaAlb, "dd/mm/yyyy") & "|" & Adodc2.Recordset!numlinea & "|"
     CargaAlbaranesCompras
     DoEvents
     Me.lblInd.Caption = "Situando...."
     lblInd.Refresh
     SituarADO2 SQL
     PonerCamposCompras
     PonerLbl
     Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    NumRegElim = 17 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
'
'        .Buttons(6).Image = 3   'Insertar Nuevo
'        .Buttons(7).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(8).Image = 10 'Stocks Almacenes
'        .Buttons(11).Image = 11 'Conjuntos
'        .Buttons(12).Image = 36 'Instalaciones
        .Buttons(12).Image = 14
        
        .Buttons(15).Image = 15  'Salir
        .Buttons(NumRegElim).Image = 6  'Primero
        .Buttons(NumRegElim + 1).Image = 7 'Anterior
        .Buttons(NumRegElim + 2).Image = 8 'Siguiente
        .Buttons(NumRegElim + 3).Image = 9 'Último
        
        
    End With
    
    limpiar Me
    
    CargaList2 True
    Frame1.Enabled = False
    Frame2.visible = True
End Sub
Private Sub CargaAlbaranesCompras()
Dim SQL As String
    lblInd.Caption = "Albaranes compras"
    lblInd.Refresh
    DatosModificados = False
    'Cargamos compras
    SQL = "SELECT s.numalbar,s.fechaalb,codprove,numlinea,s.cantidad cantidad ,"
    SQL = SQL & "if(sum(v.cantidad) is null,0,sum(v.cantidad)) asig"  'Para que pinte un cero si es null
    SQL = SQL & " from slialp s left join slcomven v on "
    SQL = SQL & " s.fechaalb=v.fechaalbc and s.numalbar=v.numalbarc and s.numlinea=v.numlineac"
    SQL = SQL & " AND s.codprove=v.codprovec AND s.codartic = v.codartic"
    
    'EL WHERE
    'aqui aqui aqui. No me enalza bien
    SQL = SQL & " WHERE s.codprove=" & Adodc1.Recordset!codProve
    SQL = SQL & " AND s.codartic='" & Adodc1.Recordset!codArtic & "'"
    
    'AHora meto los desde hasta fecha albaran
    If txtFecha(0).Text <> "" Then SQL = SQL & " AND s.fechaalb >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then SQL = SQL & " AND s.fechaalb <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"

    'El group
    SQL = SQL & " group by s.numalbar,s.fechaalb,codprove,numlinea"
    
    
    'Si muestro todos o no
    If Me.chkTodos.Value = 0 Then SQL = SQL & " Having Cantidad - asig > 0"
    
    
    
    If Me.optOrdenProv(0).Value Then
        SQL = SQL & " ORDER BY numalbar,fechaalb"
    Else
        SQL = SQL & " ORDER BY fechaalb,numalbar"
    End If
    
    Adodc2.RecordSource = SQL
    Adodc2.ConnectionString = conn
    Adodc2.LockType = adLockPessimistic
    Adodc2.CursorType = adOpenKeyset
    Adodc2.Refresh
    If Adodc2.Recordset.EOF Then
        'Ummmm no tenei a por que ser esto
        CargaList2 True
        Text3(1).Text = "":  Text3(3).Text = "": Text3(0).Text = ""
        Text4.Text = ""
        Me.FrameDes.visible = False
        MsgBox "Esta completamente asignado", vbExclamation
        Exit Sub
    End If
    Label5.Caption = Adodc2.Recordset.AbsolutePosition & " de " & Adodc2.Recordset.RecordCount
    Me.FrameDes.visible = Adodc2.Recordset.RecordCount > 1
    PonerCamposCompras
    PonerFocoBtn Me.cmdCancelar
End Sub

Private Sub PonerCamposCompras()
    Label5.Caption = Adodc2.Recordset.AbsolutePosition & " de " & Adodc2.Recordset.RecordCount
    
    If Adodc2.Recordset.EOF Then Exit Sub
    
    Text4.Text = Adodc2.Recordset.Fields(0) & "   " & Adodc2.Recordset.Fields(1) & "    " & Adodc2.Recordset.Fields(2)
    Text3(0).Text = Adodc2.Recordset.Fields(4)
    
    If IsNull(Adodc2.Recordset!asig) Then
        Text3(1).Text = "0"
        Impo = 0
        
    Else
        Text3(1).Text = Adodc2.Recordset!asig
        Impo = Adodc2.Recordset!asig
        
    End If
    Text3(3).Text = Adodc2.Recordset!Cantidad - Impo
    Text3(4).Text = ""
    CargaList2 False
End Sub

Private Sub CargaList2(limpiar As Boolean)
Dim SQL As String

Dim b As Boolean
Dim Codigo As String

On Error GoTo EcargaList
    
    lwV.ListItems.Clear
    If limpiar Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset
    
        
    'Trozo comun para todos
    'El WHERE de los desde / hasta
    Codigo = ""
    If txtFecha(2).Text <> "" Then Codigo = Codigo & " AND scaalb.fechaalb >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then Codigo = Codigo & " AND scaalb.fechaalb <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    If txtCodigo(2).Text <> "" Then Codigo = Codigo & " AND scaalb.codclien >= " & txtCodigo(2).Text
    If txtCodigo(3).Text <> "" Then Codigo = Codigo & " AND scaalb.codclien <= " & txtCodigo(3).Text
    If Codigo <> "" Then Codigo = Mid(Codigo, 5) & " AND "
    
    
    
    'Las ventas. Este SQL es mas jodido
    lblInd.Caption = "Albaranes ventas"
    lblInd.Refresh
    
    SQL = "SELECT sclien.codclien, sclien.nomclien, scaalb.fechaalb, slialb.codtipom, slialb.numalbar,"
    SQL = SQL & " slialb.numlinea, slialb.cantidad, slcomven.cantidad asig,numlineac,fechaalbc,numalbarc"
    SQL = SQL & " FROM ((scaalb INNER JOIN slialb ON (scaalb.numalbar = slialb.numalbar) AND (scaalb.codtipom ="
    SQL = SQL & " slialb.codtipom)) LEFT JOIN slcomven ON (slialb.numlinea = slcomven.numlineav) AND (slialb.numalbar"
    SQL = SQL & " = slcomven.numalbarv) AND (slialb.codtipom = slcomven.codtipom) and slialb.codartic=slcomven.codartic)"
    SQL = SQL & " INNER JOIN sclien ON scaalb.codclien = sclien.codclien WHERE "
    
    
    'EL WHERE desde / hasta
    SQL = SQL & Codigo
    'Prove articulo
    SQL = SQL & " codprovex=" & Adodc1.Recordset!codProve 'Este SEGURO
    SQL = SQL & " AND slialb.codartic='" & DevNombreSQL(Adodc1.Recordset!codArtic) & "'"
    
    'Primero cargaremos los que ya tiene asignados
    SQL = SQL & " AND numalbarc= '" & DevNombreSQL(Adodc2.Recordset!NumAlbar) & "'"
    SQL = SQL & " AND fechaalbc= '" & Format(Adodc2.Recordset!FechaAlb, FormatoFecha) & "'"
    SQL = SQL & " AND numlineac= " & Adodc2.Recordset!numlinea
    
    SQL = SQL & " order by codclien,fechaalb,codtipom,numalbar,numlinea"
    

    b = False
    lwV.Tag = "" 'AQUI llevare la marca de si tengo albaranes c
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        MeterEnListView 0
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Ahora leemos los albaranes asignados a otros, y los que faltan por asignar
    lblInd.Caption = "Albarabes s/asignar"
    lblInd.Refresh
    DoEvents
    
    
    SQL = "SELECT sclien.codclien, sclien.nomclien, scaalb.fechaalb, slialb.codtipom, slialb.numalbar, "
    SQL = SQL & " slialb.numlinea, slialb.cantidad, sum(slcomven.cantidad) asig,numlineac,fechaalbc,"
    SQL = SQL & " numalbarc  FROM ((scaalb INNER JOIN slialb ON (scaalb.numalbar = slialb.numalbar) AND  "
    SQL = SQL & " (scaalb.codtipom = slialb.codtipom))  LEFT JOIN slcomven ON (slialb.numlinea = slcomven.numlineav) AND"
    SQL = SQL & " (slialb.numalbar = slcomven.numalbarv)  AND (slialb.codtipom = slcomven.codtipom) and slialb.codartic=slcomven.codartic)"
    SQL = SQL & " INNER JOIN sclien ON scaalb.codclien = sclien.codclien  WHERE "
    
    'EL WHERE desde / hasta
    SQL = SQL & Codigo
    'Prove articulo
    SQL = SQL & " codprovex=" & Adodc1.Recordset!codProve 'Este SEGURO
    SQL = SQL & " AND slialb.codartic='" & DevNombreSQL(Adodc1.Recordset!codArtic) & "'"
    SQL = SQL & " group by 1,2,3,4,5,6"
    SQL = SQL & " order by codclien,fechaalb,codtipom,numalbar,numlinea"
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        MeterEnListView 1
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    
    
    
    
    'Y ahora veo si alguno de las asignado ha sido YA facturado
    lblInd.Caption = "Facturas"
    lblInd.Refresh
    SQL = "SELECT 0, 'FACTURADO', scafac1.fechaalb, slifac.codtipom, slifac.numalbar, slifac.numlinea, slifac.cantidad,"
    SQL = SQL & " slcomven.cantidad asig,numlineac,fechaalbc,numalbarc FROM"
    SQL = SQL & " ((scafac1 INNER JOIN slifac ON (scafac1.numalbar = slifac.numalbar) AND (scafac1.codtipoa = slifac.codtipoa))"
    SQL = SQL & " LEFT JOIN slcomven ON (slifac.numlinea = slcomven.numlineav) AND"
    SQL = SQL & " (slifac.numalbar = slcomven.numalbarv) AND (slifac.codtipoa = slcomven.codtipom) and slifac.codartic=slcomven.codartic)"
    SQL = SQL & " WHERE codprovex=" & Adodc1.Recordset!codProve 'Este SEGURO
    SQL = SQL & " AND slifac.codartic='" & DevNombreSQL(Adodc1.Recordset!codArtic) & "'"
    SQL = SQL & " AND not (slcomven.cantidad is null) "
    
    'Primero cargaremos los que ya tiene asignados
    SQL = SQL & " AND numalbarc= '" & DevNombreSQL(Adodc2.Recordset!NumAlbar) & "'"
    SQL = SQL & " AND fechaalbc= '" & Format(Adodc2.Recordset!FechaAlb, FormatoFecha) & "'"
    SQL = SQL & " AND numlineac= " & Adodc2.Recordset!numlinea
    
    SQL = SQL & " order by fechaalb,codtipom,numalbar,numlinea"
    
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        lwV.Tag = "Si" 'AQUI llevare la marca de si tiene facturas
        While Not miRsAux.EOF
            MeterEnListView 2
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
    
    
    
    
    lwV.ColumnHeaders(5).Width = 500
    Set lwV.SelectedItem = Nothing
    Set miRsAux = Nothing
    Exit Sub
EcargaList:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub

'0: Ya asignados
'1: pendientes asignar
'2: FACTURADOS
Private Sub MeterEnListView(opcion As Byte)
Dim It As ListItem
Dim C As String
Dim Impo2 As Currency

    On Error GoTo EMeterEnListView


    Impo = DBLet(miRsAux!asig, "N")
    Impo2 = miRsAux!Cantidad - Impo
    
    'Añadimos otros albaranes, pero no tienen cantidad para asignar
    If opcion = 1 And Impo2 = 0 Then Exit Sub
    


    
   'trozo comun
    Set It = lwV.ListItems.Add()
   
    'pAra evitar duplicados
    If opcion <> 2 Then
        It.Key = FijarCampoClaveLw
        It.Text = miRsAux!nomclien
    Else
        It.Text = "FACTURADO"
    End If
   
    It.SubItems(1) = miRsAux!codtipom
    It.SubItems(2) = miRsAux!NumAlbar
    It.SubItems(3) = miRsAux!FechaAlb
    It.SubItems(4) = miRsAux!numlinea
    It.SubItems(5) = miRsAux!Cantidad
    It.SubItems(6) = Impo
    It.SubItems(7) = Impo2
    
    
    If opcion = 0 Then
        It.ForeColor = vbMorado
    ElseIf opcion = 2 Then
        It.ForeColor = vbVerde
    End If
            

        
    
    It.SubItems(8) = 0
    
    Exit Sub
EMeterEnListView:
    NumRegElim = Err.Number
    C = Err.Description
    If opcion = 1 Then
        'Error 35602:  Clave duplicada
        '
        If Err.Number = 35602 Then
            C = ""
            lwV.ListItems.Remove It.Index
        End If
    End If
        
    If C <> "" Then MsgBox "Error: " & NumRegElim & "  " & C, vbExclamation
    
    Err.Clear
End Sub


Private Function FijarCampoClaveLw() As String
    FijarCampoClaveLw = miRsAux!codtipom & Format(miRsAux!NumAlbar, "00000000") & _
            Format(miRsAux!FechaAlb, "ddmmyy") & Format(miRsAux!numlinea, "0000")
End Function


Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(IndiceTxt).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndiceTxt).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(NumRegElim).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    IndiceTxt = Index
    Select Case Index
    Case 0, 1
        'proveedor
        Set frmPro = New frmComProveedores
        frmPro.DatosADevolverBusqueda = "0|1|"
        frmPro.Show vbModal
        Set frmPro = Nothing
        
    Case 2, 3
        'cliente
        Set frmCli = New frmFacClientes
        frmCli.DatosADevolverBusqueda = "0|1|"
        frmCli.Show vbModal
        Set frmCli = Nothing
        
    Case 4, 5
        'articulo
        
        Set frmArt = New frmAlmArticulos
        frmArt.DatosADevolverBusqueda = "@1@"
        frmArt.Show vbModal
        Set frmArt = Nothing
        
        
    End Select
    
    If CadenaDesdeOtroForm <> "" Then
        txtCodigo(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        txtNombre(Index).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        
        CadenaDesdeOtroForm = ""
    End If
End Sub




Private Sub imgFecha_Click(Index As Integer)
    NumRegElim = Index
    Set frmC = New frmCal
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
End Sub




Private Sub HacerToolbar(Indice As Integer)
    Select Case Indice
    Case 1
        'Frame4.visible = False
        Frame2.visible = True
        Frame1.Enabled = False
        PonerFoco txtCodigo(4)
        
    Case 7, 12
        EliminarAsignaciones (Indice = 12)  'Partiendo del albaran COMPRA
        
        
    Case 15
        Unload Me
        
    Case 17 To 20
        Screen.MousePointer = vbHourglass
        HacerDesplaz Indice
        Screen.MousePointer = vbDefault
    End Select
End Sub



Private Sub HabilitarAlbaranesVentas(IndiceAsignado As Integer)
Dim Cad As String
Dim SePuede As Boolean
Dim i As Integer
        'ESTE ES EL QUE tiene asignado. Por si quiere asignar mas
        If IndiceAsignado > 0 Then
            lwV.ListItems(IndiceAsignado).ForeColor = vbMorado
            MsgBox "Para reasignar la cantidad tendra que eliminar la asignacion primero", vbExclamation
        Else
            'El albaran de compra no esta asignado a ninguno de estas mismo albaran venta
            'Con lo cual, la ultima linea la habilitare
            IndiceTxt = IndiceTxt - 1

    
            'Si no, de aqui hasta el final veo. Aquel que tenga
            NumRegElim = IndiceTxt - 1
            Impo = 0
            While NumRegElim > 0
                    
                SePuede = False
                If lwV.ListItems(NumRegElim).Text = lwV.ListItems(IndiceTxt).Text Then
                    For i = 2 To 4
                        If lwV.ListItems(NumRegElim).SubItems(i) <> lwV.ListItems(NumRegElim).SubItems(i) Then Exit For
                    Next i
                    If i > 4 Then SePuede = True
                End If
                If SePuede Then
                    Impo = Impo + CCur(lwV.ListItems(NumRegElim).SubItems(6))  'Asignadas
                    NumRegElim = NumRegElim - 1
                Else
                    NumRegElim = 0
                End If
                    
            Wend
            'Sumo las asignadas en la linea seleccionada
            Impo = Impo + CCur(lwV.ListItems(IndiceTxt).SubItems(6))  'Asignadas
            'las disponibles son la cantiad - asignadas
            lwV.ListItems(IndiceTxt).SubItems(6) = Impo
            Impo = CCur(lwV.ListItems(IndiceTxt).SubItems(5)) - Impo
            lwV.ListItems(IndiceTxt).SubItems(7) = Impo
            lwV.ListItems(IndiceTxt).ForeColor = vbVerde
        End If
End Sub



Private Sub lwV_DblClick()
Dim Impo2 As Currency

    If Frame2.visible Then Exit Sub  'Esta seleccionando desde hasta
    
    Impo = 1
    If Adodc2.Recordset Is Nothing Then
        Impo = 0
    Else
        If Adodc2.Recordset.EOF Then Impo = 0
    End If
    
    If lwV.SelectedItem Is Nothing Then Impo = 0
    If Impo = 0 Then Exit Sub
    
    
    If lwV.SelectedItem.ForeColor = vbMorado Then
        MsgBox "Este albarán ya lo tiene asignado.", vbExclamation
        Exit Sub
    End If
    
    If lwV.SelectedItem.ForeColor = vbVerde Then
        MsgBox "Pertenece a una factura.", vbExclamation
        Exit Sub
    End If
    
    If lwV.SelectedItem.ListSubItems(7).Bold Then
        MsgBox "Lo acabas de asignar", vbExclamation
        Exit Sub
    End If
    
    
    Impo = CCur(lwV.SelectedItem.SubItems(7))
   
    'Si tiene cantidad para asignar
    If Impo = 0 Then
        MsgBox "No tienen unidades para asignar la compra / venta", vbExclamation
        Exit Sub
    End If
    
    
   
        Impo2 = CCur(Text3(3).Text)
        If Impo2 = 0 Then
            MsgBox "No hay mas cantidad para asignar", vbExclamation
            Exit Sub
        End If
    
    If Impo2 < Impo Then
        If MsgBox("No dispone de suficientes unidades.Desea continuar asignando las disponibles?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Impo = Impo - Impo2
    Else
        Impo2 = Impo
        Impo = CCur(lwV.SelectedItem.SubItems(7)) - Impo
    End If
    
    'ASig
    If Not DatosModificados Then Text3(4).Text = "0"
    DatosModificados = True
    Text3(3).Text = CCur(Text3(3).Text) - Impo2
    Text3(4).Text = CCur(Text3(4).Text) + Impo2
    
    'Asignamos importe
    lwV.SelectedItem.ListSubItems(7).ForeColor = vbRed
    lwV.SelectedItem.ListSubItems(7).Bold = True
    lwV.SelectedItem.ListSubItems(8) = Impo2
    Impo2 = CCur(lwV.SelectedItem.SubItems(6)) + Impo2
    lwV.SelectedItem.SubItems(6) = Impo2
    lwV.SelectedItem.ListSubItems(7) = Impo  '// ASIGNAMOS todo el inmporte
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Frame2.visible Then
        Exit Sub
    Else
       
            If DatosModificados Then
                MsgBox "Ha modificado datos", vbExclamation
                Exit Sub
            End If

    End If
    HacerToolbar Button.Index
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, False
    
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim C As String
    Screen.MousePointer = vbHourglass
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    If Trim(txtCodigo(Index).Text) = "" Then
        C = ""
    Else
        If Index < 4 Then
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "Campo debe ser numérico: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = "0"
                C = ""
            End If
        End If
        
    
        Select Case Index
        Case 4, 5
            'articulos
            C = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtCodigo(Index).Text, "T")
            
        Case 0, 1
            'proveedores
            C = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtCodigo(Index).Text, "N")
        Case 2, 3
            'clientes
            C = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", txtCodigo(Index).Text, "N")
        End Select
        
        If C = "" Then C = " ----- NO EXISTE "
    End If
    txtNombre(Index).Text = C
    Screen.MousePointer = vbDefault
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
 KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
End Sub

Private Sub PonerCampos()

    If Adodc1.Recordset.EOF Then
        Text1.Text = ""
        Text1.Tag = ""
        Text2.Text = ""
        Text2.Tag = ""
        
        CargaList2 True
    Else
        'SI datos
        If Text1.Tag <> CStr(Adodc1.Recordset!codProve) Then
            Text1.Text = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(Adodc1.Recordset!codProve), "N")
            Text1.Tag = CStr(Adodc1.Recordset!codProve)
        End If
        If Text2.Tag <> CStr(Adodc1.Recordset!codArtic) Then
            Text2.Text = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", CStr(Adodc1.Recordset!codArtic), "T")
            Text2.Tag = CStr(Adodc1.Recordset!codArtic)
        End If
        CargaAlbaranesCompras
    End If

    PonerLbl
End Sub

Private Sub PonerLbl()
    On Error GoTo EP
    lblInd.Caption = ""
    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    lblInd.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    lblInd.Refresh
    Exit Sub
EP:
    Err.Clear
    
End Sub


Private Sub HacerDesplaz(opcion As Integer)
Dim b As Boolean
    If opcion = 17 Or opcion = 20 Then
        b = True
        If opcion = 17 Then
            Adodc1.Recordset.MoveFirst
        Else
            Adodc1.Recordset.MoveLast
        End If
    Else
        b = False
        If opcion = 18 Then
            If Not Adodc1.Recordset.BOF Then
                Adodc1.Recordset.MovePrevious
                b = True
                If Adodc1.Recordset.BOF Then
                    Adodc1.Recordset.MoveNext
                    b = False
                End If
            End If
        Else
            If Not Adodc1.Recordset.EOF Then
                Adodc1.Recordset.MoveNext
                b = True
                If Adodc1.Recordset.EOF Then
                    Adodc1.Recordset.MovePrevious
                    b = False
                End If
                
            End If
        End If
    End If
    If b Then PonerCampos
End Sub



Private Sub EliminarAsignaciones(AlbaranVenta As Boolean)
Dim Cad As String
    On Error GoTo EElim
    
    
    'If lwC.SelectedItem Is Nothing Then Exit Sub
    If AlbaranVenta Then
        If lwV.SelectedItem Is Nothing Then Exit Sub
        
        If lwV.SelectedItem.ForeColor = vbVerde Then
            MsgBox "Esta facturado", vbExclamation
            Exit Sub
        End If
            
            
        
        Cad = "Va a eliminar las asociaciones del albaran : " & vbCrLf
    
    Else
        'Partiendo de albaran compra,borrare todas las asignaciones
        Cad = ""
        If Adodc2.Recordset Is Nothing Then
            Cad = "N"
        Else
            If Adodc2.Recordset.EOF Then
                Cad = "N"
            Else
                If DBLet(Adodc2.Recordset!asig, "N") = 0 Then
                    MsgBox "No tiene datos asignados ", vbExclamation
                    Cad = "N"
                End If
            End If
        End If
        
        If Cad = "" Then
            If lwV.Tag <> "" Then
                Cad = "Algunos de los datos asociados ya estan facturados" & vbCrLf & "¿continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then Cad = ""
            End If
        End If
        If Cad <> "" Then Exit Sub
            
            
        Cad = "Va a eliminar las asociaciones compras / ventas para: " & vbCrLf
    End If

    Cad = Cad & vbCrLf & "Proveedor: " & Text1.Text
    Cad = Cad & vbCrLf & "Artículo: " & Text2.Text
    Cad = Cad & vbCrLf & "Albarán: " & Adodc2.Recordset!NumAlbar & " / " & Adodc2.Recordset!numlinea
    Cad = Cad & vbCrLf & "Fecha: " & Adodc2.Recordset!FechaAlb
    Cad = Cad & vbCrLf & "Cantidad: " & Adodc2.Recordset!Cantidad
    If AlbaranVenta Then
        Cad = Cad & vbCrLf & String(60, "-")
        Cad = Cad & vbCrLf & "Cliente: " & lwV.SelectedItem.Text
        
        Cad = Cad & vbCrLf & "Albaran: " & lwV.SelectedItem.SubItems(1) & " " & lwV.SelectedItem.SubItems(2)
        Cad = Cad & " " & lwV.SelectedItem.SubItems(3) & " " & lwV.SelectedItem.SubItems(4)
        
        Cad = Cad & vbCrLf & "Asignado: " & lwV.SelectedItem.SubItems(6)
        
    End If
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass

    'Hacemos un delete de la tabla
    Cad = "DELETE FROM slcomven WHERE codprovec=" & Adodc1.Recordset!codProve
    Cad = Cad & " AND codartic = '" & DevNombreSQL(CStr(Adodc1.Recordset!codArtic)) & "'"
    Cad = Cad & " AND numalbarc = '" & DevNombreSQL(Adodc2.Recordset!NumAlbar) & "'"
    Cad = Cad & " AND fechaalbc = '" & Format(Adodc2.Recordset!FechaAlb, FormatoFecha) & "'"
    Cad = Cad & " AND numlineac = " & Adodc2.Recordset!numlinea

    If AlbaranVenta Then
        'SOLO EL ALBARAB ESE
        Cad = Cad & " AND codtipom = '" & lwV.SelectedItem.SubItems(1) & "'"
        Cad = Cad & " AND numalbarv = " & lwV.SelectedItem.SubItems(2)
        Cad = Cad & " AND fechaalbv = '" & Format(lwV.SelectedItem.SubItems(3), FormatoFecha) & "'"
        Cad = Cad & " AND numlineav = " & lwV.SelectedItem.SubItems(4)
    End If
    conn.Execute Cad
    
    'Cargamos los datos OTRA VEZ
    Cad = Adodc2.Recordset!NumAlbar & "|" & Format(Adodc2.Recordset!FechaAlb, "dd/mm/yyyy") & "|" & Adodc2.Recordset!numlinea & "|"
    CargaAlbaranesCompras
    DoEvents
    Me.lblInd.Caption = "Situando...."
    lblInd.Refresh
    SituarADO2 Cad
    PonerCamposCompras
    PonerLbl
EElim:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub


Private Sub SituarADO2(Referencia As String)
Dim Esta As Boolean
Dim C As String
Dim Fin As Boolean

    
    Fin = False
    Esta = False
    While Not Fin
        If Adodc2.Recordset.EOF Then
            Fin = True
        Else
            C = RecuperaValor(Referencia, 1)
            If C = Adodc2.Recordset!NumAlbar Then
                C = RecuperaValor(Referencia, 2)
                If C = Format(Adodc2.Recordset!FechaAlb) Then
                    C = RecuperaValor(Referencia, 3)
                    If C = CStr(Adodc2.Recordset!numlinea) Then
                        Esta = True
                        Fin = True
                        
                    End If
                End If
            End If
            If Not Esta Then Adodc2.Recordset.MoveNext
        End If
    Wend
    
    If Not Esta Then
        If Adodc2.Recordset.RecordCount > 0 Then Adodc2.Recordset.MoveFirst
    End If
End Sub
