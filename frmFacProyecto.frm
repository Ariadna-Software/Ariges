VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacProyecto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proyectos"
   ClientHeight    =   11355
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   18135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacProyecto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11355
   ScaleWidth      =   18135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5400
      TabIndex        =   72
      Top             =   0
      Width           =   2895
      Begin VB.ComboBox cboFiltro 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFacProyecto.frx":000C
         Left            =   120
         List            =   "frmFacProyecto.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   240
         Width           =   2535
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6780
      Left            =   120
      TabIndex        =   52
      Top             =   3840
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   11959
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Albaranes"
      TabPicture(0)   =   "frmFacProyecto.frx":0044
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFramePp(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lwLineaAlbaran"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameCampos2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Impresion"
      TabPicture(1)   =   "frmFacProyecto.frx":0060
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblFramePp(4)"
      Tab(1).Control(1)=   "lwEulerLineas"
      Tab(1).Control(2)=   "FrameToolAux(5)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tareas / costes"
      TabPicture(2)   =   "frmFacProyecto.frx":007C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblFramePp(3)"
      Tab(2).Control(1)=   "lblFramePp(5)"
      Tab(2).Control(2)=   "Label1(64)"
      Tab(2).Control(3)=   "Label1(63)"
      Tab(2).Control(4)=   "ListView1"
      Tab(2).Control(5)=   "ListView2"
      Tab(2).ControlCount=   6
      Begin VB.Frame FrameToolAux 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   -74880
         TabIndex        =   57
         Top             =   960
         Width           =   2205
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Intercalar"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Lotes"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Ordenar lineas"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameCampos2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   8895
         Begin MSComctlLib.ListView lwAlb 
            Height          =   2535
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4471
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Albaran"
               Object.Width           =   2734
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2734
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Referenca"
               Object.Width           =   5998
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Bases"
               Object.Width           =   2645
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "orderficha"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblAjuste 
            Alignment       =   1  'Right Justify
            Caption         =   "Label2"
            Height          =   255
            Left            =   5520
            TabIndex        =   74
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label lblFramePp 
            Caption         =   "Vinculados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2745
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   9240
         TabIndex        =   53
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView lwAlb 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4471
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Albaran"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   3440
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Referenca"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "orderficha"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblFramePp 
            Caption         =   "Pendientes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   5265
         End
      End
      Begin MSComctlLib.ListView lwEulerLineas 
         Height          =   5535
         Left            =   -72480
         TabIndex        =   59
         Top             =   720
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9763
         SortKey         =   5
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Articulo"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   11642
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2363
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Descuento"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "descripcionReal"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   -72840
         TabIndex        =   61
         Top             =   3480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         SortKey         =   8
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5503
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Precio"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "codartic"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lwLineaAlbaran 
         Height          =   2535
         Left            =   2520
         TabIndex        =   63
         Top             =   3960
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   4471
         SortKey         =   4
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Albaran"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   9878
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   2363
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dto"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "descripcionReal"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   -72840
         TabIndex        =   65
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Albaran"
            Object.Width           =   2998
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Trab."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5503
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tarea"
            Object.Width           =   1429
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2195
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Tiempo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Horas"
            Object.Width           =   2364
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   -74640
         TabIndex        =   68
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   -74880
         TabIndex        =   67
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblFramePp 
         Caption         =   "Fichadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Index           =   5
         Left            =   -74760
         TabIndex        =   66
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label lblFramePp 
         Caption         =   "Lineas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   4080
         Width           =   5265
      End
      Begin VB.Label lblFramePp 
         Caption         =   "Lin. imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   4
         Left            =   -74880
         TabIndex        =   62
         Top             =   480
         Width           =   5265
      End
      Begin VB.Label lblFramePp 
         Caption         =   "Costes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   3
         Left            =   -74760
         TabIndex        =   60
         Top             =   3480
         Width           =   5265
      End
   End
   Begin VB.Frame FrameCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   17820
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
         Left            =   11400
         MaxLength       =   15
         TabIndex        =   71
         Tag             =   "numfactu|N|S|||sproyecto|numfactu||N|"
         Text            =   "numfactu"
         Top             =   1800
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
         Height          =   1800
         Index           =   2
         Left            =   13200
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Tag             =   "Domicilio|T|S|||sproyecto|observa||N|"
         Text            =   "frmFacProyecto.frx":0098
         Top             =   360
         Width           =   4395
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
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   38
         Tag             =   "Domicilio|T|N|||sproyecto|domclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   840
         Width           =   4755
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
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1320
         Width           =   4365
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
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   36
         Tag             =   "Forma de Pago|N|N|0|999|sproyecto|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   1320
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
         Index           =   17
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   840
         Width           =   4365
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
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "Cod. Agente|N|N|0|9999|sproyecto|codagent|0000|N|"
         Text            =   "Text1"
         Top             =   840
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
         Index           =   6
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   33
         Tag             =   "NIF Cliente|T|N|||sproyecto|nifclien||N|"
         Text            =   "123456789"
         Top             =   360
         Width           =   2055
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
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   32
         Tag             =   "teléfono Cliente|T|S|||sproyecto|telclien||N|"
         Text            =   "12345678911234567899"
         Top             =   360
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
         Index           =   10
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   31
         Tag             =   "Población|T|N|||sproyecto|pobclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   1320
         Width           =   3765
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   30
         Tag             =   "CPostal|T|N|||sproyecto|codpobla||N|"
         Text            =   "Text15"
         Top             =   1320
         Width           =   855
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "Provincia|T|N|||sproyecto|proclien||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text22"
         Top             =   1800
         Width           =   2445
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
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Direccion/Dpto.|N|S|0|999|sproyecto|coddirec|000|N|"
         Text            =   "Text1"
         Top             =   360
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
         Index           =   12
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   27
         Tag             =   "Direccion/Dpto.|T|S|||sproyecto|nomdirec||N|"
         Text            =   "Text2"
         Top             =   360
         Width           =   4365
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
         Left            =   7680
         MaxLength       =   60
         TabIndex        =   26
         Tag             =   "Referencia Cliente|T|S|||sproyecto|referenc||N|"
         Text            =   "Text1 Text1 Text1 Te"
         Top             =   1800
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "Obser."
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
         Left            =   13200
         TabIndex        =   49
         Top             =   120
         Width           =   840
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
         Height          =   240
         Index           =   34
         Left            =   6240
         TabIndex        =   47
         Top             =   870
         Width           =   945
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
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   870
         Width           =   840
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   7440
         ToolTipText     =   "Buscar forma de pago"
         Top             =   1350
         Width           =   240
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
         Height          =   240
         Index           =   15
         Left            =   6240
         TabIndex        =   45
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   7440
         ToolTipText     =   "Buscar agente"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         ToolTipText     =   "Buscar cliente varios"
         Top             =   420
         Visible         =   0   'False
         Width           =   240
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
         Height          =   240
         Index           =   20
         Left            =   120
         TabIndex        =   44
         Top             =   420
         Width           =   555
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
         Height          =   240
         Index           =   19
         Left            =   3450
         TabIndex        =   43
         Top             =   420
         Width           =   855
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
         Height          =   240
         Index           =   16
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   930
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
         Height          =   240
         Index           =   17
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   7440
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Direc./Dpto"
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
         Left            =   6240
         TabIndex        =   40
         Top             =   420
         Width           =   1125
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   6
         Left            =   1080
         ToolTipText     =   "Buscar población"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ref. Cliente"
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
         Left            =   6240
         TabIndex        =   39
         Top             =   1860
         Width           =   1380
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   22
      Top             =   0
      Width           =   1455
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Asignar albaranes"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar factura"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Marcar para facturar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Imprimir portes"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Duplicar albaran"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8420
      TabIndex        =   20
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   21
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
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   19
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
      Left            =   16320
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
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
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   10800
      Width           =   3615
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   240
         Top             =   0
         Width           =   3300
      End
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   3075
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
      Left            =   16440
      TabIndex        =   6
      Top             =   10800
      Width           =   1335
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
      Left            =   14880
      TabIndex        =   5
      Top             =   10800
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   0
      Top             =   10800
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
      Left            =   16410
      TabIndex        =   7
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   17775
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
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   50
         Tag             =   "Fecha cierre|F|S|||sproyecto|fecfinal|dd/mm/yyyy|N|"
         Top             =   300
         Width           =   1425
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
         Left            =   1440
         TabIndex        =   24
         Tag             =   "Tipo |T|N|||sproyecto|codtipom||S|"
         Text            =   "Text3"
         Top             =   300
         Width           =   615
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
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Cod. Cliente|N|N|0|999999|sproyecto|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   300
         Width           =   885
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
         Left            =   6240
         MaxLength       =   60
         TabIndex        =   3
         Tag             =   "Nombre Cliente|T|N|||sproyecto|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   300
         Width           =   5625
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
         Left            =   12120
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Responsable|N|N|0|9999|sproyecto|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   300
         Width           =   760
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
         Left            =   13080
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   300
         Width           =   4440
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha inicio|F|N|||sproyecto|fecproyec|dd/mm/yyyy|N|"
         Top             =   300
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   120
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Codigo|N|S|0||sproyecto|numproyec|0000000|S|"
         Text            =   "Text1 7"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4680
         Picture         =   "frmFacProyecto.frx":00BC
         ToolTipText     =   "Buscar fecha"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cierre"
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
         Left            =   3720
         TabIndex        =   51
         Top             =   60
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6240
         ToolTipText     =   "Buscar cliente"
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Index           =   0
         Left            =   5280
         TabIndex        =   16
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable"
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
         Left            =   12120
         TabIndex        =   14
         Top             =   60
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   13440
         ToolTipText     =   "Buscar trabajador"
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   60
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3120
         Picture         =   "frmFacProyecto.frx":0147
         ToolTipText     =   "Buscar fecha"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   11
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
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
Attribute VB_Name = "frmFacProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public Datos_A_Ver As String    'Tendra el nº proyecto que quiere ver

                                
Private Const TipoProyecto = "ALY"
                                
                                
'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmBasico2 'Form M7to Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 '%=%=frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Modificar albaranes

'   6.-   lineas costes descripcion

'-------------------------------------------------------------------------






Dim PrimeraVez As Boolean

Dim EsCabecera As Byte   '0 cabecera   1-direc    2 direnv
'Para saber en MandaBusquedaPrevia si busca en la tabla sproyecto o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Dim txtAnterior As String
Dim PulsadoMas2 As Boolean

'Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
'Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

'Dim PorCaja As Boolean
''Para Saber si se ha salido con precio caja y hay que calcular el importe de la
''linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad
'
'Dim Precio As String 'Precio de la linea de Articulo
'
'Dim cadList As String 'cadena para pasar al historico
'Dim motivo As String 'cadena para el motivo si es factura Rectificativa
'
'

'



Dim SQL As String


'Para buscar por los chks
Private BuscaChekc As String



Private AlbaranesDelProyecto As String
Private PedidosVinculadosEnAlbaranes As String
Private CostesCargados As Boolean







Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                '
                
                 If ModificaDesdeFormulario(Me, 1) Then
                      TerminaBloquear
                      
                      PosicionarData

                
                  End If
'
                
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialb'
                        
             'If ModificarAlbaranesVinculados Then
            If False Then
                PonerModo 2
                CostesCargados = False
                AlbaranesDelProyecto = ""
                PonerCamposAlbaranes
                
            End If
            
        Case 6
            'Matriculas
                         
             
          
            
            
            
            
            
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub








Private Sub cmdCancelar_Click()
Dim EraNuevaLinea As Boolean
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
        
           
                PonerModo 2
            
            
            
            
            
            
            
                'cmdRegresar_Click
            
            
            
            
        Case 6
       
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String


    LimpiarCampos 'Vacía los TextBox

    
    PonerModo 3
    
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

  
    
    
    
    Text1(15).Text = TipoProyecto
    PonerFoco Text1(1)
End Sub



Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        
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
Dim cadB As String
    
    cadB = ""
    AnyadeFiltro cadB

'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = 0
        
        MandaBusquedaPrevia cadB
    Else
        LimpiarCampos
        If cadB = "" Then cadB = "true"
        CadenaConsulta = "Select * from " & NombreTabla
        
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB
        
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index - 1
    Screen.MousePointer = vbHourglass
    PonerCampos
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean


    



    'Por si acaso esta bloqueado
    SoloEnEfectivoAlbaranes = False
    EsClienteBloqueado2 Text1(4).Text, 0, True, SoloEnEfectivoAlbaranes
    
    
    PonerModo 4

    PonerFoco Text1(1)
   
    'Si es Cliente de Varios no se pueden modificar sus datos
    
    
    
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
    cmdCancelar.Cancel = True
End Sub




Private Sub BotonesCampos(Nuevo As Boolean)
    
    
End Sub



Private Sub Form_Activate()
    
    If PrimeraVez Then
        'Si no tiene albaranes
         PrimeraVez = False
        
        If Me.Datos_A_Ver <> "" Then
            Modo = 1
            Text1(15).Text = Mid(Me.Datos_A_Ver, 1, 3)
            Text1(0).Text = Mid(Me.Datos_A_Ver, 4)
            cmdAceptar_Click
        
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()


    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon


    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(1).Picture
    Next kCampo
    




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
       
    'Lineas
    With Me.ToolbarAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        '3 4 5
        
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 32
        .Buttons(6).Image = 39
        .Buttons(7).Image = 38
        
    End With

    
    If vParamAplic.Ariagro <> "" Then
        With Me.ToolbarAux(2)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            
            .Buttons(1).Image = 3
            .Buttons(3).Image = 5
        End With
    
    End If
    
   
    
   
  
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        '11(30   21    20   16
        
        
        '                                                                                   Indice  antiguo
        .Buttons(1).Image = 30 'Nº Serie si lineas con articulos de control Nº serie  Ant 11
        If vParamAplic.NumeroInstalacion = vbFenollar Then .Buttons(1).ToolTipText = "Reestablecer al pedido"
        .Buttons(2).Image = 21 'GEnerar factura ant 12
        .Buttons(3).Image = 20  'Marcar a facturar 13
        
        
        
        
        .Buttons(5).Image = 11 'duplicar albaran
        

        'MAYO 2015  Herbelca. ALbran ruta Castellon
        .Buttons(6).Image = 16 'Imprimir Pedido ant 16
        
        
        If vParamAplic.TipoPortes <> 1 Then
            If vParamAplic.PathFirmasAlbaran <> "" Then
                .Buttons(4).ToolTipText = "Imprimir albaran firmado"
                .Buttons(4).Style = tbrDefault
                .Buttons(4).Image = 54  '54
            Else
                .Buttons(4).Style = tbrSeparator
                .Buttons(4).visible = False
                .Buttons(4).ToolTipText = ""
            End If
        Else
            .Buttons(4).Style = tbrDefault
            .Buttons(4).ToolTipText = "Imprimir portes"
        End If

        
        
        
        
    End With


      
      
    cboFiltro.ListIndex = 1
    LimpiarCampos   'Limpia los campos TextBox
    
        
    Me.Label1(1).Caption = DevuelveTextoDepto(True)

    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
 
    Text1(18).Left = 20000 'lo quito de la vista. NUNCA estar enabled
    
    NombreTabla = "sproyecto"
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where false"


    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
   
     If Datos_A_Ver = "" Then
         PonerModo 0
     Else
        
    End If

    PrimeraVez = True
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    lblAjuste.Caption = ""
    Me.lwAlb(0).ListItems.Clear
    AlbaranesDelProyecto = ""
    Me.lwAlb(1).ListItems.Clear
    lwEulerLineas.ListItems.Clear
    ListView2.ListItems.Clear
    ListView1.ListItems.Clear
    lwLineaAlbaran.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo = 5 Then
        Cancel = 1
        Exit Sub
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Datos_A_Ver = ""
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub





Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If EsCabecera = 0 Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(15), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            
            
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 2), "0000000")
            
        Else
            If EsCabecera = 3 Then
                
                
            ElseIf EsCabecera = 1 Then
                'Llama desde Prismatico Direcciones/Departamentos
                Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                
            End If
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
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
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







Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = Val(Me.imgBuscar(3).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub




Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte


    If Modo = 0 Then Exit Sub
    If Modo = 2 And Index <> 14 Then Exit Sub
    
    TerminaBloquear
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            
            Set frmC = New frmBasico2
            AyudaClientes frmC, Text1(4).Text
            Set frmC = Nothing
         
            Indice = 5
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 4
                txtAnterior = Text1(4).Text
            End If
        Case 1 'NIF para cliente de Varios
            Indice = 6
            Set frmCV = New frmBasico2
            AyudaClientesV frmCV, Text1(Indice)
            Set frmCV = Nothing

            
        Case 2 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = 1
                   'ANTES
                '01/DICIEMBRE/2010   DAVID
                'MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
                LanzaBusquedaDpto True, CInt(Indice)
                
             End If
             
        Case 3 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            If Index = 7 Then
                Indice = 27
            ElseIf Index = 8 Then
                Indice = 28
            Else
                Indice = Index
            End If
            Me.imgBuscar(3).Tag = Indice
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            Indice = 14
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
            PonerFoco Text1(Indice)
            
            
        Case 5 'Agente
            Indice = 17
            PonerFoco Text1(Indice)
            Set frmA = New frmBasico2
            AyudaAgentesComerciales frmA, Text1(Indice), , True
            Set frmA = Nothing
            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            
            
        
            
            

        
        
    End Select
    
    

    
    PonerFoco Text1(Indice)
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




Private Sub ListView2_DblClick()
    If Modo <> 2 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    If ListView2.SelectedItem.Text <> "ALC" Then
        If ListView2.SelectedItem.Text <> "FAC" Then
            If ListView2.SelectedItem.Text <> "PED" Then Exit Sub
        End If
    End If
    
    
    If ListView2.SelectedItem.Text = "ALC" Then
    
      'IT.Tag = "numalbar =" & DBSet(Rs!NUmAlbar, "T") & " AND  fechaalb =" & DBSet(Rs!FechaAlb, "F") & " AND codprove =" & Rs!Codprove
       With frmComEntAlbaranSA
            .hcoCodMovim = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            .hcoFechaMovim = RecuperaValor(ListView2.SelectedItem.Tag, 2)
            .hcoCodProve = RecuperaValor(ListView2.SelectedItem.Tag, 3)
            .EsHistorico = False
            .Show vbModal
        End With
    
    ElseIf ListView2.SelectedItem.Text = "PED" Then
        'PEDIDOS
            frmComEntPedidosSa.MostrarDatos = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            frmComEntPedidosSa.EsHistorico = False
            frmComEntPedidosSa.Show vbModal
    
    Else
    
        
        'IT.Tag = "numfactu =" & DBSet(Rs!Numfactu, "T") & " AND  fecfactu=" & DBSet(Rs!FecFactu, "F") & " AND codprove =" & Rs!Codprove
         With frmComHcoFacturSA
            .hcoCodMovim = RecuperaValor(ListView2.SelectedItem.Tag, 1)
            .hcoFechaMovim = RecuperaValor(ListView2.SelectedItem.Tag, 2)
            .hcoCodProve = RecuperaValor(ListView2.SelectedItem.Tag, 3)
            .Show vbModal
        End With
    End If
    
    
End Sub

Private Sub lwAlb_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim columna As Integer
    
    
    columna = ColumnHeader.Index - 1
    If columna = 1 Then columna = 4
    If columna <> lwAlb(Index).SortKey Then
        lwAlb(Index).SortKey = columna
        lwAlb(Index).SortOrder = lvwAscending
    Else
        If lwAlb(Index).SortOrder = lvwAscending Then
            lwAlb(Index).SortOrder = lvwDescending
        Else
            lwAlb(Index).SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub lwAlb_DblClick(Index As Integer)
Dim Knodo As Integer
    If Modo <> 2 Then Exit Sub
    If Me.Datos_A_Ver <> "" Then Exit Sub
    If lwAlb(Index).ListItems.Count = 0 Then Exit Sub
    If lwAlb(Index).SelectedItem Is Nothing Then Exit Sub
    
    DblClickTreeview lwAlb(Index).SelectedItem.Text
    If Index = 0 Then
        Knodo = lwAlb(Index).SelectedItem.Index
        PonerCamposAlbaranes
        
        'Pongo menor porque si es igual cuando carga el listvie ya lo deja seleccioando el ultimo
        If Knodo < lwAlb(Index).ListItems.Count Then lwAlb(Index).SelectedItem = lwAlb(Index).ListItems(Knodo)
    End If
End Sub

Private Sub lwLineaAlbaran_DblClick()
    If Modo <> 2 Then Exit Sub
    If lwLineaAlbaran.ListItems.Count = 0 Then Exit Sub
    If lwLineaAlbaran.SelectedItem Is Nothing Then Exit Sub
    
    DblClickTreeview lwLineaAlbaran.SelectedItem.Text
    
End Sub

Private Sub mnBuscar_Click()
    
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
        'BotonEliminarLinea
    Else
        'Eliminar Albaran
        BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
    'Imprimir Albaran
    BotonImprimir_ 45, False '45: Informe de Albaranes
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
       '  BotonModificarLinea
    Else   'Modificar albaran
        
            If Modo <> 2 Then Exit Sub
            If Text1(16).Text <> "" Then
                    MsgBox "Proyecto cerrado", vbExclamation
                    Exit Sub
            End If
    
    
    
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
      '   BotonAnyadirLinea False
    Else 'Añadir Cabecera
         BotonAnyadir
    End If
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)

 
    If SSTab1.Tab = 1 Then LineasImpresionEul
    If SSTab1.Tab = 2 Then CargaCostesEuler Modo < 2
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    txtAnterior = Text1(Index).Text
    kCampo = Index
    'If Index = 9 Then HaCambiadoCP = False 'CPostal
   
    If Not (Index = 15 And Modo = 1) Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 38 Then KEYdown KeyCode
    
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
            Case 27, 28, 29
                Ind = Index - 20
            
            Case 43
                Ind = 13
            End Select
            If Ind >= 0 Then
                'PulsadoMas2 = True
                PulsarTeclaMas True, Ind
            End If
        End If
    End If
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
Dim devuelve As String
Dim campo As String
Dim ImpDto As Currency
        
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
    If txtAnterior = Text1(Index).Text Then
        
        

        Exit Sub
    End If
          
    
          
          
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 41 'Fecha Albaran,fecenvio
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                If Text2(Index).Text = "" And Modo >= 3 Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
              
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                   
                    Text1(5).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "N")
                Else 'If Modo = 3 Then 'Modo Insertar
                    'si es ART-Albaran de factura Rectificativa ya he cargado los
                    'datos de la factura
                     
                    
                        campo = "nomclien"
                        devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", campo)
                        If campo <> Text1(5).Text Then PonerDatosCliente Text1(Index).Text
                    
                    If Text1(Index).Text = "" Then
                        PonerFoco Text1(Index)
                    Else
                        If Text1(5).Locked Then
       
                            PonerFoco Text1(13)

                        Else
                            PonerFoco Text1(5)
                        End If
                    End If
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
            If Text1(6).Locked Then Exit Sub
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
            If (Modo = 4) Then
                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
                    
        Case 9 'Cod. Postal
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
             'If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
             '    Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
             '    Text1(Index + 2).Text = devuelve
            'End If
            'VieneDeBuscar = False
            
        Case 12 'Cod. Direc
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionada tiene esa direccion
            If PonerDptoEnCliente Then
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then ComprobarRefObligatoria
            
        Case 14 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then
                    
                    CalcularDatosFactura
                End If
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
            
       
    End Select
End Sub


Private Sub AnyadeFiltro(ByRef CadenaSQL As String)

    If cboFiltro.ListIndex > 0 Then
        If CadenaSQL <> "" Then CadenaSQL = CadenaSQL & " AND "
        CadenaSQL = CadenaSQL & IIf(cboFiltro.ListIndex = 2, " NOT ", "") & " fecfinal IS null"
    End If
End Sub

Private Sub HacerBusqueda()
Dim cadB As String



    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    AnyadeFiltro cadB
    
    
    If chkVistaPrevia = 1 Then
        EsCabecera = 0
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera = 0 Then
        Cad = Cad & ParaGrid(Text1(15), 8, "Tipo")
        Cad = Cad & ParaGrid(Text1(0), 12, "Codigo")
        Cad = Cad & ParaGrid(Text1(1), 15, "F. inicio")
        Cad = Cad & ParaGrid(Text1(4), 12, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 53, "Nombre Cliente")
        
        tabla = NombreTabla
        
        Titulo = "Proyectos"
        devuelve = "0|1|"
    
    Else
        If EsCabecera = 1 Then
                'DIRECION DEPARTAMENTO
                If vParamAplic.HayDeparNuevo = 1 Then
                    Titulo = "Dptos Cliente: "
                    Desc = "Dpto."
                ElseIf vParamAplic.HayDeparNuevo = 0 Then
                    Titulo = "Direc. Cliente: "
                    Desc = "Direc."
                Else
                    Titulo = "Obra Cliente: "
                    Desc = "Obra"
                End If
                Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
                Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|18·"
                Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
                tabla = "sdirec"
                devuelve = "0|1|"
                
        ElseIf EsCabecera = 2 Then
            'DIRENVIO
            '--------------------
            Titulo = "Dirección de envio cliente: "
            Desc = " envio"
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Codigo" & Desc & "|sdirenvio|coddiren|N|000|18·"
            Cad = Cad & "Descripción" & Desc & "|sdirenvio|nomdiren|T||65·"
            tabla = "sdirenvio"
            devuelve = "0|1|"
        
        Else
            Stop
        
        End If
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            frmB.vselElem = 2
            frmB.vDescendente = True
        Else
            frmB.vselElem = 1
            frmB.vDescendente = False
        End If
        
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        If EsCabecera > 0 Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing

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
            PonerFoco Text1(kCampo)
            Text1(0).BackColor = vbYellow
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



Private Sub PonerCampos()
Dim B As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
     'Si es un Albaran de Ticket visualizamos unos datos y sino otros
    B = (Data1.Recordset!EsTicket = 1)
    Me.Toolbar1.Buttons(11).Enabled = (Not B)
    
      

    PonerCamposForma Me, Data1
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
  
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
     
   
    
    Text2(16).Text = ""

    PonerCamposAlbaranes
    
    CalcularDatosFactura
    If SSTab1.Tab = 1 Then LineasImpresionEul
    CostesCargados = False
    If SSTab1.Tab = 2 Then CargaCostesEuler False
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo

    BuscaChekc = ""
    

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If Datos_A_Ver <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
        PonLblIndicador lblIndicador, Data1
    End If
    
    DespalzamientoVisible NumReg > 1
        
    
    If B Then 'modo=2
        If Me.FrameCampos2.visible Then
            'Tiene campos visibles
            If Not Data1.Recordset.EOF Then B = True
        Else
            B = False
        End If
    End If
    If vParamAplic.Ariagro <> "" Then
        ToolbarAux(2).Buttons(1).Enabled = B
        ToolbarAux(2).Buttons(3).Enabled = B
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(15), B
    BloquearTxt Text1(16), B  'La fecha de cierre solo es valido en buusqueda
    BloquearTxt Text1(18), True  'SIEMPRE bloqueado
    
    
    
    
  
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    Me.imgFecha(0).Enabled = B
  
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    Me.imgBuscar(1).visible = False

              
    'Modo Linea de Albaranes
    '- poner visible ampliacion linea
    BloquearTxt Text2(16), True
    '- poner visible nombre proveedor linea
    BloquearTxt Text2(9), True
      
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
       
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    
    
    'Para remarcar el cliente
    '&H00C0FFFF&
    If Modo = 2 Then Text1(5).BackColor = &HC0FFFF
    
    
   
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub




Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolbar Button.Index
End Sub

Private Sub HacerToolbar(Indice As Integer)
    Select Case Indice
        
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click 'Modificar
        Case 3: mnEliminar_Click  'Borrar
            
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos
        
            
        Case 8:
                mnImprimir_Click 'Imprimir Albaran
        
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



Private Function ModificarLinea2() As Boolean
ModificarLinea2 = True


End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next
    
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas "
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub






Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Modo <> 2 Then Exit Sub
    If Text1(16).Text <> "" Then
            MsgBox "Proyecto cerrado", vbExclamation
            Exit Sub
    End If



    If Button.Index = 1 Then
        'Habria que bloquear
        
        CadenaDesdeOtroForm = ""
        frmListado5.OpcionListado = 43
        frmListado5.OtrosDatos = Text1(4).Text & "|" & Text1(15).Text & "|" & Text1(0).Text & "|"
        frmListado5.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            Screen.MousePointer = vbHourglass
            PonerCampos
            Screen.MousePointer = vbDefault
        End If
        
        
    ElseIf Button.Index = 2 Then

            CadenaDesdeOtroForm = ""
            CadenaDesdeOtroForm = PonerTrabajadorConectado("")
            If CadenaDesdeOtroForm = "" Then
                CadenaDesdeOtroForm = "Error trabajador conectado"
            Else
                If lwAlb(0).ListItems.Count = 0 Then
                    CadenaDesdeOtroForm = "Ningún albarán vinculado"
                Else
                    CadenaDesdeOtroForm = "" 'OK
                End If
            End If
            If CadenaDesdeOtroForm <> "" Then
                MsgBox CadenaDesdeOtroForm, vbExclamation
                Exit Sub
            End If
            'Si todos los albaranes tienen lineas
            CadenaDesdeOtroForm = "(Select scaalb.codtipom FROM scaalb  left join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar =slialb.numalbar"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  WHERE codartic is null  and (scaalb.codtipom,scaalb.numalbar)  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  ) as AAA"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", CadenaDesdeOtroForm, "1", "1")
            If Val(CadenaDesdeOtroForm) > 0 Then
                MsgBox "Hay albaranes sin lineas", vbExclamation
                Exit Sub
            End If
        
            'El campo OBRA es el mismo
            CadenaDesdeOtroForm = "(Select coalesce(coddirec,0) FROM scaalb WHERE (scaalb.codtipom,scaalb.numalbar)"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  group by 1) as AAA"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", CadenaDesdeOtroForm, "1", "1")
            If Val(CadenaDesdeOtroForm) > 1 Then
                MsgBox "Distintas obras/departamento  en los albaranes", vbExclamation
                Exit Sub
            End If


            'Que no han cambiado el cliente
            CadenaDesdeOtroForm = "(Select codclien FROM scaalb WHERE (scaalb.codtipom,scaalb.numalbar)"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  group by 1) as AAA"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", CadenaDesdeOtroForm, "1", "1")
            If Val(CadenaDesdeOtroForm) > 1 Then
                MsgBox "Distintos obras/departamento  en los albaranes", vbExclamation
                Exit Sub
            End If


            'La dtoppago y dtogneral deb ser el mismo
            CadenaDesdeOtroForm = "(Select dtoppago,dtognral FROM scaalb WHERE (scaalb.codtipom,scaalb.numalbar)"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  group by 1,2) as AAA"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", CadenaDesdeOtroForm, "1", "1")
            If Val(CadenaDesdeOtroForm) > 1 Then
                MsgBox "Distintos descuentos(general y pronto pago) en los albaranes", vbExclamation
                Exit Sub
            End If
            
            CadenaDesdeOtroForm = "(Select codforpa FROM scaalb WHERE (scaalb.codtipom,scaalb.numalbar)"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "  group by 1) as AAA"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", CadenaDesdeOtroForm, "1", "1")
            If Val(CadenaDesdeOtroForm) > 1 Then
                MsgBox "Distintas formas de pago  en los albaranes", vbExclamation
                Exit Sub
            End If
            
            CadenaDesdeOtroForm = "(scaalb.codtipom,scaalb.numalbar) IN (Select codtipoa,numalbar from  sproyectolin  WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  AND 1"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "codforpa", "scaalb", CadenaDesdeOtroForm, "1")
            If CadenaDesdeOtroForm <> Val(Text1(14).Text) Then
                'Distinta forma de pago
                CadenaDesdeOtroForm = "Albaranes: " & CadenaDesdeOtroForm & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Proyecto: " & Text1(14).Text & " (" & Text2(14).Text & ")"
                CadenaDesdeOtroForm = "Formas de pago: " & vbCrLf & CadenaDesdeOtroForm & vbCrLf & "¿Continuar con la que tiene el albaran? " & vbCrLf & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Si: Forma de pago albaran" & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "NO: Forma de pago PROYECTO (" & Text2(14).Text & ")" & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Cancelar: detener proceso sin hacer cambios"
                CadenaDesdeOtroForm = MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel)
                If CByte(CadenaDesdeOtroForm) <> vbCancel Then
                
                    If CByte(CadenaDesdeOtroForm) = vbYes Then
                        CadenaDesdeOtroForm = ""
                    Else
                        CadenaDesdeOtroForm = "UPDATE scaalb set codforpa=" & Text1(14).Text & " WHERE  (scaalb.codtipom,scaalb.numalbar) IN "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & " (Select codtipoa,numalbar from  sproyectolin WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  "
                        If ejecutar(CadenaDesdeOtroForm, False) Then CadenaDesdeOtroForm = ""
                    End If
                    
                End If
                If CadenaDesdeOtroForm <> "" Then Exit Sub
            End If
            
            'Si tiene lineas  manuales
            If Not CompruebaTotales(True) Then Exit Sub

            
            
            'Veamos que tiene todos los albaranes marcados para facturar
            CadenaDesdeOtroForm = "(scaalb.codtipom,scaalb.numalbar) IN (Select codtipoa,numalbar from  sproyectolin WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  AND factursn"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "scaalb", CadenaDesdeOtroForm, "0")
            If Val(CadenaDesdeOtroForm) > 0 Then
                CadenaDesdeOtroForm = "Albaranes sin marca de facturar. Total:  " & CadenaDesdeOtroForm & vbCrLf
                If vUsu.Nivel > 1 Then
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                Else
                    If MsgBox(CadenaDesdeOtroForm & "¿Marcarlos y continuar generando la factura?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                        CadenaDesdeOtroForm = "UPDATE scaalb set factursn =1 WHERE  (scaalb.codtipom,scaalb.numalbar) IN "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & " (Select codtipoa,numalbar from  sproyectolin WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  AND factursn=0"
                        If ejecutar(CadenaDesdeOtroForm, False) Then CadenaDesdeOtroForm = ""
                    End If
                End If
            Else
                CadenaDesdeOtroForm = "" 'Todos marcados
            End If
                
                
                
                
            If CadenaDesdeOtroForm = "" Then
            
                'En notasportes grabo el numero de proyecto. Por si lo necestio mas adelante
                CadenaDesdeOtroForm = "UPDATE scaalb set notasportes='ALY" & Text1(0).Text & "' WHERE  (scaalb.codtipom,scaalb.numalbar) IN "
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " (Select codtipoa,numalbar from  sproyectolin WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " )  AND factursn=0"
                ejecutar CadenaDesdeOtroForm, False
            
                CadenaDesdeOtroForm = ""
            
                'Facturacion de Albaran de Mostrador
                frmListadoPed.codClien = "ALY"  'utilizamos esta vble para pasarle el tipo de movimiento
                frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº proyecto
                AbrirListadoPed (222)
                    
                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "fecfinal", "sproyecto", "numproyec=" & Text1(0).Text & " AND codtipom", "ALY", "T")
                If CadenaDesdeOtroForm <> "" Then
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    PosicionarDataTrasEliminar
                End If
            End If
            CadenaDesdeOtroForm = ""
    End If
End Sub

Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    If Modo <> 2 Then Exit Sub

    If Data1.Recordset.EOF Then Exit Sub

    If Index = 0 Then HacerToolBarLineas Button.Index - 1
    
    
    
    BotonesToolBarAux
    
End Sub


Private Sub HacerToolBarLineas(Index As Integer)  'index: boton

Dim Cad As String
   
    
    If Modo <> 2 Then Exit Sub
    If Text1(16).Text <> "" Then
            MsgBox "Proyecto cerrado", vbExclamation
            Exit Sub
    End If
    
    
    If Index > 0 Then
        If lwEulerLineas.ListItems.Count = 0 Then
            MsgBox "Ningun dato", vbExclamation
            Exit Sub
        End If
        If Index < 3 Then
            'Modificar eliminar.
            'el seleccionado
            If Me.lwEulerLineas.SelectedItem Is Nothing Then
                MsgBox "Seleccione una linea", vbExclamation
                Exit Sub
            End If
        End If
    Else
       ' If Me.lwEulerLineas.ListItems.Count = 0 Then Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    
    If Index < 2 Then
        'nuevo modificar
        If Index = 1 Then
            
            CadenaDesdeOtroForm = Mid(lwEulerLineas.SelectedItem.Key, 2, 4)
        Else
            CadenaDesdeOtroForm = ""  '"" = nuevo   id= linea
        End If
        frmListado5.OtrosDatos = Data1.Recordset!codtipom & "|" & Data1.Recordset!numproyec & "|"
        frmListado5.OpcionListado = 42
        frmListado5.Show vbModal
        
    
    Else
        If Index = 2 Then
            'Eliminar
            Cad = "Va a eliminar linea impresion" & vbCrLf & "Articulo : " & Me.lwEulerLineas.SelectedItem.Text & vbCrLf
            Cad = Cad & "Descripcion : " & Me.lwEulerLineas.SelectedItem.SubItems(1) & vbCrLf
            Cad = Cad & "Importe : " & Me.lwEulerLineas.SelectedItem.SubItems(5) & vbCrLf
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                Cad = " WHERE codtipom='" & Data1.Recordset!codtipom & "' AND numproyec = " & Data1.Recordset!numproyec
                 Cad = "DELETE FROM sproyectolin2 " & Cad & " AND numlinea= " & Mid(Me.lwEulerLineas.SelectedItem.Key, 2, 4)
                If ejecutar(Cad, False) Then CadenaDesdeOtroForm = "OK"
            End If

        Else
            'imprimir
            If lwEulerLineas.Tag <> "" Then
                MsgBox lwEulerLineas.Tag, vbExclamation
            Else
                'BotonImprimir 90, False '90: Informe de Albaranes lineas especiales
            End If
        End If
    End If
    
    If CadenaDesdeOtroForm <> "" Then PonerLineasImpresionEULER






 
    
    
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento Button.Index
End Sub







Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then

             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos

             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " " & NombreTabla & ".codtipom= '" & Text1(15).Text & "' and " & NombreTabla & ".numproyec= " & Val(Text1(0).Text)
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean, QueGRid As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    
    If QueGRid = 1 Then
       SQL = "SELECT codtipom, numproyec, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
   
           SQL = SQL & " WHERE false"
   
       SQL = SQL & " Order by codtipom, numproyec, ordenlin, numlinea"
    
    Else
        'Matriculas en portes
        SQL = "SELECT codtipom,numproyec,matricula,descr FROM "
        SQL = SQL & Trim(NombreTabla) & "_portes as " & NombreTabla & " WHERE "
         
        If enlaza Then
           SQL = SQL & ObtenerWhereCP(False)
           
       Else
           SQL = SQL & " false "
       End If
       
        
    End If
    MontaSQLCarga = SQL
End Function





Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        
        B = (Modo = 2) And Me.Datos_A_Ver = ""
        'Insertar
        Toolbar1.Buttons(1).Enabled = (B Or Modo = 0)
        Me.mnNuevo.Enabled = (B Or Modo = 0)
        'Modificar
        If B Then
            If Me.Data1.Recordset.EOF Then B = False
        End If
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
            
            
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
            
        Toolbar1.Buttons(8).Enabled = (Modo = 2)
        Me.mnImprimir.Enabled = (Modo = 2)
            
            
            
            
        B = (Modo = 2) And Me.Datos_A_Ver = ""
        
        'Nº Series
        Toolbar2.Buttons(1).Enabled = B
        
        'Generar Factura
        Toolbar2.Buttons(2).Enabled = B
        Toolbar2.Buttons(3).Enabled = B
        If Toolbar2.Buttons(4).Style = tbrDefault Then Toolbar2.Buttons(4).Enabled = B
        
        
        
        
        BotonesToolBarAux
        
End Sub






Private Sub LimpiarDatosCliente()
Dim i As Byte

    For i = 4 To 14
        Text1(i).Text = ""
    Next i
    Text2(12).Text = ""
    Text2(14).Text = ""
    Text2(17).Text = ""
    Text1(17).Text = "" 'agente
    
    
     
    
  
    
End Sub
    





Private Sub BotonImprimir_(OpcionListado As Byte, EsInformePortes As Boolean)
    
    Dim devuelve
    
    CompruebaTotales False    'Avisa (si hay lineas, que todas suman lo que toca
    
    
    
    
    frmImprimir.NombreRPT = "rProyectos.rpt"
    frmImprimir.NombrePDF = frmImprimir.NombreRPT

        
        
    
        
    
        With frmImprimir
            'Febrero 2010
                .outTipoDocumento = 0
            
            
            .FormulaSeleccion = "{sproyecto.codtipom}='" & Text1(15).Text & "' AND ({sproyecto.numproyec}=" & Text1(0).Text & ")"
            .OtrosParametros = "|pCodigoISO=""""|pCodigoRev=""""|pCodUsu=2000|vPortes=""""|PuntoVerde=""""|Albarcon=0|pTipoIVA=0|"

            .NumeroParametros = 7
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 12 'OpcionListado
            
            .Titulo = "PROYECTO"
                
            .ConSubInforme = True
            .Show vbModal
        End With
    
    
    
End Sub








Private Sub PosicionarDataTrasEliminar()
Dim HayDatos As Boolean
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    HayDatos = SituarDataTrasEliminar(Data1, NumRegElim)
    If HayDatos Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            Data1.Recordset.MoveLast
            If Data1.Recordset.EOF Then HayDatos = False
        End If
    End If
    If HayDatos Then
        PonerCampos
    Else
        LimpiarCampos

        PonerModo 0
    End If
End Sub


Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim B As Boolean
    
    
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
            SoloEnEfectivoAlbaranes = False
            If vCliente.ClienteBloqueado(0, SoloEnEfectivoAlbaranes) Then
                
                    LimpiarDatosCliente
                    Set vCliente = Nothing
                    Exit Sub
                
            End If
            
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)v
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        

        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
                
            End If
            
            If Modo = 3 Or Modo = 4 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                'Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
               ' Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")

                
                    
                
                
               
                
          
            End If


            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
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
Dim B As Boolean
Dim RN As ADODB.Recordset
Dim Aux As String

    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    B = vCliente.LeerDatosCliVario(nifClien)
    If B Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
            
    'Si tiene manipulador de fitosnaitarios
    If B Then
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Set RN = New ADODB.Recordset
            Aux = "Select ManipuladorNumCarnet , fcaducidad "
            Aux = Aux & ",ManipuladortipoCarnet from sclvar WHERE nifclien = " & DBSet(nifClien, "T")
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Aux = "|||||"
            If Not RN.EOF Then
                Aux = DBLet(RN!ManipuladorNumCarnet, "T") & "|"
                Aux = Aux & vCliente.Nombre & "|"
                If Not IsNull(RN!fcaducidad) Then Aux = Aux & Format(RN!fcaducidad, "dd/mm/yyyy")
                Aux = Aux & "|"
                'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
                If Val(DBLet(RN!ManipuladortipoCarnet, "N")) > 0 Then
                    Aux = Aux & IIf(RN!ManipuladortipoCarnet = 2, "Cualificado", "Básico") & "|"
                    Aux = Aux & RN!ManipuladortipoCarnet & "|"
                Else
                    Aux = Aux & "||"
                End If
            End If
            RN.Close
            Set RN = Nothing
            Me.Text1(45).Text = RecuperaValor(Aux, 1)
            Me.Text1(46).Text = RecuperaValor(Aux, 2)
            Me.Text1(47).Text = RecuperaValor(Aux, 3)
            Text2(0).Text = RecuperaValor(Aux, 4)
            'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            Me.Text1(48).Text = RecuperaValor(Aux, 5)
        End If
    End If
            
'    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente

    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
         If Not Comprobar_NIF(NIF) Then
            If MsgBox("El NIF es incorrecto. ¿Continuar de igual modo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        
        End If
       
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


Private Function ActualizarFecMovCliente() As Boolean
Dim vCliente As CCliente
Dim B As Boolean

    On Error GoTo EActFecha

    ActualizarFecMovCliente = False
    Set vCliente = New CCliente
    vCliente.Codigo = Text1(4).Text
    B = vCliente.ActualizaUltFecMovim(Text1(1).Text)
    Set vCliente = Nothing
    
EActFecha:
    If Err.Number <> 0 Then B = False
    ActualizarFecMovCliente = B
End Function


Private Sub CalcularDatosFactura()
Dim i As Integer
End Sub



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




Private Sub UpdateaNomDirec()
Dim N As Integer
Dim Ol As Integer
Dim C As String

    N = -1
    If Not IsNull(Data1.Recordset!CodDirec) Then N = Data1.Recordset!CodDirec
    
    Ol = -1
    If Text1(12).Text <> "" Then Ol = CInt(Text1(12).Text)
    
    If N <> Ol Then
        If Ol < 0 Then
            C = "NULL"
        Else
            C = DBSet(Text2(12).Text, "T")
        End If
        C = "UPDATE sproyecto set nomdirec=" & C
        C = C & " WHERE codtipom = '" & Text1(15).Text & "' AND numproyec=" & Text1(0).Text
        ejecutar C, False
    End If
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera = 0
        imgBuscar_Click Index
        
    Else
        'Lineas
        
       
        
        
    End If
        
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
        If EsCabecera = 1 Then 'Llama desde VerTodos del Form
            Text1(12).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
     
        End If
End Sub

Private Sub LanzaBusquedaDpto(Departamento As Boolean, Indice As Integer)

    Set frmDptoEnvio = New frmFacCliEnvDpto
    frmDptoEnvio.DireccionesEnvio = Not Departamento
    If Text1(Indice).Text <> "" Then
        frmDptoEnvio.VerDatoDpto = CInt(Text1(Indice).Text)
    Else
        frmDptoEnvio.VerDatoDpto = -1
    End If
    frmDptoEnvio.codClien = CLng(Text1(4).Text)
    frmDptoEnvio.NomClien = Text1(5).Text
    frmDptoEnvio.Show vbModal
    Set frmDptoEnvio = Nothing
End Sub




Private Sub BotonesToolBarAux()
Dim B As Boolean


    B = Modo = 2 And Me.Datos_A_Ver = ""
        
    
    ToolbarAux(0).Buttons(1).Enabled = B
    If B Then
       If Me.lwEulerLineas.ListItems.Count = 0 Then B = False
      
    End If

    
    ToolbarAux(0).Buttons(2).Enabled = B
    ToolbarAux(0).Buttons(3).Enabled = B
    
    ToolbarAux(0).Buttons(5).Enabled = B
    ToolbarAux(0).Buttons(6).Enabled = B
    ToolbarAux(0).Buttons(7).Enabled = B
    
    
    

End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Function DevWHERE() As String
    DevWHERE = "codtipom='" & TipoProyecto & "' AND numproyec = " & Text1(0).Text
End Function


Private Sub PonerCamposAlbaranes()
Dim Valora As String
Dim IT As ListItem
Dim Facturado As Boolean
Dim ImporteAlbaran As Currency
Dim Seleccionado As Integer
Dim ALbPpal As String
    
    
    
    lblAjuste.Caption = ""
    lblAjuste.Tag = 0
    Me.lwAlb(0).ListItems.Clear
    Me.lwAlb(1).ListItems.Clear
    lwLineaAlbaran.ListItems.Clear
    AlbaranesDelProyecto = ""
    PedidosVinculadosEnAlbaranes = ""
    Set miRsAux = New ADODB.Recordset
    DoEvents
    
    
    'Albaran PPAL
    SQL = "select codtipoa,numalbar from sproyectolin where ppal=1 and " & DevWHERE
    ALbPpal = ""
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then ALbPpal = miRsAux!Codtipoa & Format(miRsAux!Numalbar, "000000")
    miRsAux.Close
    
    
    'Si esta cerrado. Vamos a ir a buscar los albaranes a la factura
    Facturado = False
    If Text1(16).Text <> "" Then Facturado = True
    
    If Facturado Then
        SQL = " select slifac.codtipoa codtipom ,slifac.numalbar,fechaalb,referenc,codartic,nomartic,cantidad,importel,numlinea,numpedcl"
        SQL = SQL & " from scafac1 inner join slifac on"
        SQL = SQL & " scafac1.codtipom=slifac.codtipom and scafac1.numfactu=slifac.numfactu and"
        SQL = SQL & " scafac1.fecfactu=slifac.fecfactu and scafac1.numalbar=slifac.numalbar and"
        SQL = SQL & " scafac1.Codtipoa = slifac.Codtipoa        WHERE   (scafac1.codtipoa,scafac1.numalbar) IN  ("
        SQL = SQL & "          (select codtipoa,numalbar from sproyectolin where " & DevWHERE & ")"
        SQL = SQL & "  ) ORDER BY scafac1.codtipoa,scafac1.numalbar,numlinea"


    
    Else
        'De albaranes
        SQL = "select scaalb.codtipom,scaalb.numalbar,fechaalb,referenc,codartic,nomartic,cantidad,importel,numlinea,numpedcl,codartic "
        SQL = SQL & " from scaalb left join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
        SQL = SQL & " WHERE codClien = " & Text1(4).Text & " AND  (scaalb.codtipom,scaalb.numalbar) IN"
        SQL = SQL & "          (select codtipoa,numalbar from sproyectolin where " & DevWHERE & ")"
        SQL = SQL & " ORDER BY slialb.codtipom,slialb.numalbar,numlinea"
    
    End If
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Valora = ""
    While Not miRsAux.EOF
        SQL = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        If Valora <> SQL Then
            'Insertamos el padre
            
            Valora = CStr(SQL)
            Set IT = Me.lwAlb(0).ListItems.Add()
            IT.Text = SQL
            IT.SubItems(1) = miRsAux!FechaAlb
            IT.SubItems(2) = DBLet(miRsAux!referenc, "T") & " "
            IT.SubItems(3) = " "
            IT.SubItems(4) = Format(miRsAux!FechaAlb, "yymmdd") & SQL
            IT.ToolTipText = DBLet(miRsAux!referenc, "T")
            AlbaranesDelProyecto = AlbaranesDelProyecto & ", ('" & miRsAux!codtipom & "'," & miRsAux!Numalbar & ")"
            If DBLet(miRsAux!NumPedcl, "N") > 0 Then PedidosVinculadosEnAlbaranes = PedidosVinculadosEnAlbaranes & ", " & miRsAux!NumPedcl
            
            
            If ALbPpal = SQL Then
                
                IT.Bold = True
                For NumRegElim = 1 To IT.ListSubItems.Count
                    IT.ListSubItems(NumRegElim).Bold = True
                Next
                IT.ToolTipText = "Albaran principal"
                
            End If
            
            
            If DBLet(miRsAux!codArtic, "T") = "" Then
                IT.ForeColor = vbRed: IT.ListSubItems(1).ForeColor = vbRed
                IT.ToolTipText = "Sin lineas"
            End If
            
        End If
        
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    If AlbaranesDelProyecto <> "" Then AlbaranesDelProyecto = Mid(AlbaranesDelProyecto, 2)
    If PedidosVinculadosEnAlbaranes <> "" Then PedidosVinculadosEnAlbaranes = Mid(PedidosVinculadosEnAlbaranes, 2)
    
    
    'Pendientes
    SQL = "select slialb.codtipom,slialb.numalbar,fechaalb,referenc,codartic,nomartic,cantidad,importel,numlinea "
    SQL = SQL & " from scaalb inner join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    SQL = SQL & " WHERE codClien = " & Text1(4).Text & " AND  not (scaalb.codtipom,scaalb.numalbar) in"
    'SQL = SQL & " (select codtipoa,numalbar,numlinea from sproyectolin WHERE " & DevWHERE & ")"
    SQL = SQL & " (select codtipoa,numalbar from sproyectolin )"
    
    SQL = SQL & " ORDER BY slialb.codtipom,slialb.numalbar,numlinea"
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Valora = ""
    While Not miRsAux.EOF
        SQL = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        If Valora <> SQL Then
            'Insertamos el padre
            Valora = CStr(SQL)
            
            Set IT = Me.lwAlb(1).ListItems.Add()
            IT.Text = SQL
            IT.SubItems(1) = miRsAux!FechaAlb
            IT.SubItems(2) = DBLet(miRsAux!referenc, "T") & " "
            IT.SubItems(3) = Format(miRsAux!FechaAlb, "yymmdd") & SQL
            IT.ToolTipText = DBLet(miRsAux!referenc, "T")
            
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    
    
    'Lineas albaranes
    If AlbaranesDelProyecto <> "" Then
        'Facturado
        If Text1(16).Text <> "" Then
            SQL = "select  codtipoa codtipom,numalbar,codalmac,codartic,nomartic,cantidad,precioar,dtoline1,dtoline2,importel"
            SQL = SQL & " from slifac WHERE codtipom='FPY' AND numfactu=" & Data1.Recordset!Numfactu
            SQL = SQL & "  AND fecfactu=" & DBSet(Data1.Recordset!fecfinal, "F")
            SQL = SQL & " ORDER BY codtipoa,numalbar,codartic"
        
        Else
            SQL = "select  codtipom,numalbar,codalmac,codartic,nomartic,cantidad,precioar,dtoline1,dtoline2,importel"
            SQL = SQL & " from slialb WHERE (codtipom,numalbar) IN (" & AlbaranesDelProyecto & ")"
            SQL = SQL & " ORDER BY codtipom,numalbar,codartic"
        End If
        
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Valora = ""
        While Not miRsAux.EOF
            
            SQL = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
            If Valora <> SQL Then
                If Valora <> "" Then
                    For NumRegElim = 1 To Me.lwAlb(0).ListItems.Count
                        If lwAlb(0).ListItems(NumRegElim).Text = Valora Then
                            lwAlb(0).ListItems(NumRegElim).SubItems(3) = Right(String(10, " ") & Format(ImporteAlbaran, FormatoImporte), 10)
                            Exit For
                        End If
                    Next
                End If
                Valora = SQL
                ImporteAlbaran = 0
            End If
            Set IT = Me.lwLineaAlbaran.ListItems.Add
            IT.Text = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
            IT.SubItems(1) = DBLet(miRsAux!codArtic, "T")
            IT.SubItems(2) = DBLet(miRsAux!NomArtic, "T")
            IT.SubItems(3) = Format(miRsAux!cantidad, FormatoCantidad)
            IT.SubItems(4) = Format(miRsAux!precioar, FormatoPrecio)
            IT.SubItems(5) = Format(miRsAux!dtoline1 + miRsAux!dtoline2, FormatoCantidad)
            IT.SubItems(6) = Format(miRsAux!ImporteL, FormatoCantidad)
            ImporteAlbaran = ImporteAlbaran + miRsAux!ImporteL
            
            'TotalCostes = TotalCostes + RS!ImporteL
             
            miRsAux.MoveNext
        Wend
        miRsAux.Close

        If Valora <> "" Then
            For NumRegElim = 1 To Me.lwAlb(0).ListItems.Count
                If lwAlb(0).ListItems(NumRegElim).Text = Valora Then
                    lwAlb(0).ListItems(NumRegElim).SubItems(3) = Right(String(10, " ") & Format(ImporteAlbaran, FormatoImporte), 10)
                    Exit For
                End If
            Next
            
            
            If Not Facturado Then
                ImporteAlbaran = 0
                Valora = ""
                For NumRegElim = 1 To Me.lwAlb(0).ListItems.Count
                    If lwAlb(0).ListItems(NumRegElim).Bold Then
                        Valora = "HAY PPAL"
                    Else
                        ImporteAlbaran = ImporteAlbaran + ImporteFormateado(lwAlb(0).ListItems(NumRegElim).SubItems(3))
                    End If
                Next
                If Valora <> "" Then
                    lblAjuste.Caption = "Ajuste: " & Format(-ImporteAlbaran, FormatoImporte)
                    lblAjuste.Tag = -ImporteAlbaran
                End If
            End If
            
        End If


    End If
    
    
    Set miRsAux = Nothing
End Sub




Private Sub LineasImpresionEul()
    
    
    Screen.MousePointer = vbHourglass
    
    PonerLineasImpresionEULER
    
    BotonesToolBarAux

    
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerLineasImpresionEULER()
Dim SQL As String
Dim N As Integer
Dim vImpo As Currency

    Me.lwEulerLineas.ListItems.Clear
    lwEulerLineas.Tag = ""
        
    If Modo <> 2 Then Exit Sub
    
  
    Set miRsAux = New ADODB.Recordset
    SQL = "Select codtipom,numproyec,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel FROM  sproyectolin2"
    SQL = SQL & " WHERE codtipom= " & DBSet(Text1(15).Text, "T") & " AND numproyec =" & Text1(0).Text & " order by numlinea"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vImpo = 0
    N = 0
    While Not miRsAux.EOF
        N = N + 1
        lwEulerLineas.ListItems.Add , "k" & Format(miRsAux!numlinea, "000"), miRsAux!Articulo ' & miRsAux!Numalbar, miRsAux!Articulo
        lwEulerLineas.ListItems(N).SubItems(1) = Replace(miRsAux!descrarticulo, vbCrLf, " ")
        lwEulerLineas.ListItems(N).SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
        lwEulerLineas.ListItems(N).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        lwEulerLineas.ListItems(N).SubItems(4) = Format(miRsAux!dtoline1, FormatoCantidad)
        lwEulerLineas.ListItems(N).SubItems(5) = Format(miRsAux!ImporteL, FormatoCantidad)
        lwEulerLineas.ListItems(N).ToolTipText = miRsAux!descrarticulo
        vImpo = vImpo + miRsAux!ImporteL
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If N > 0 Then
        'Tiene lineas y NO suma el burto
'        If ImporteFormateado(Text3(56).Text) <> vImpo Then
'            SQL = "B.imponible albaran: " & Text3(56).Text & vbCrLf
'            SQL = SQL & "Suma lineas:           " & Format(vImpo, FormatoImporte)
'            vImpo = ImporteFormateado(Text3(56).Text) - vImpo
'            SQL = SQL & vbCrLf & "Diferencia:              " & Format(vImpo, FormatoImporte)
'            lwEulerLineas.Tag = "Importes lineas impresion: " & vbCrLf & vbCrLf & SQL
'            MsgBox SQL, vbExclamation
'        End If
    End If
        
      Set miRsAux = Nothing
 
End Sub





Private Sub DblClickTreeview(CADENA As String)
    If Modo <> 2 Then Exit Sub
    If Me.Datos_A_Ver <> "" Then Exit Sub
    
    'Si esta facturado
    If Text1(16).Text <> "" Then
        
        'Facturado = True
        With frmFacHcoFacturas2
            .DesdeFichaCliente = True
            .hcoCodMovim = Format(Data1.Recordset!Numfactu, "0000000")
            .hcoCodTipoM = "FPY"
            .hcoFechaMov = Data1.Recordset!fecfinal
            .Show vbModal
        End With
    Else
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.hcoCodTipoM = Mid(CADENA, 1, 3)
        frmFacEntAlbSAIL.hcoCodMovim = Mid(CADENA, 4, 7)
        frmFacEntAlbSAIL.Show vbModal
        
        PonerCamposAlbaranes
        
        
        
    End If
End Sub



Private Function DatosOk() As Boolean
  Dim B As Boolean
  
    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    
    
    
    If Modo = 4 Then
        SQL = ""
        'No esta cerrado
        If Text1(16).Text <> "" Then SQL = "Esta cerrado"
        'Puede eliminar
        If Val(Data1.Recordset!codClien) <> Val(Text1(4).Text) Then
            If Me.lwAlb(0).ListItems.Count > 0 Then SQL = "Tiene albaranes vinculados"
        End If
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            B = False
        End If
    End If
    

    DatosOk = B



    

End Function

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(TipoProyecto) Then
        Text1(0).Text = vTipoMov.ConseguirContador(TipoProyecto)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        SQL = CadenaInsertarDesdeForm(Me)
        
        If SQL <> "" Then
            If ejecutar(SQL, False) Then
                vTipoMov.IncrementarContador vTipoMov.TipoMovimiento
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE codtipom=" & DBSet(Text1(15).Text, "T") & " AND numproyec =" & Text1(0).Text
                PonerCadenaBusqueda
                PonerModo 2
               
            End If
        End If
    End If
    Set vTipoMov = Nothing
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    SQL = ""
    'No esta cerrado
    If Text1(16).Text <> "" Then SQL = "Esta cerrado"
    'Puede eliminar
    If Me.lwAlb(0).ListItems.Count > 0 Then SQL = "Tiene albaranes vinculados"
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If





    '### a mano
    SQL = "¿Seguro que desea eliminar el proyecto?" & vbCrLf
    SQL = SQL & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), "00000")
    SQL = SQL & vbCrLf & "Cliente: " & Data1.Recordset!codClien & " " & Data1.Recordset!NomClien
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = " codtipom= " & DBSet(Text1(15).Text, "T") & " AND numproyec =" & Text1(0).Text
        
        
        conn.Execute "Delete from sproyectolin2 where " & SQL
        conn.Execute "Delete from sproyectolin where " & SQL
        conn.Execute "Delete from sproyecto where " & SQL

        CancelaADODC Me.Data1
        
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            
            PonerModo 0
        End If
        
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar proyecto", Err.Description
End Sub






Private Sub CargaCostesEuler(limpiar As Boolean)
Dim oldC As Byte
Dim C1 As String
Dim RS As ADODB.Recordset
Dim N As Integer
Dim H As Currency
Dim TotalCostes As Currency
Dim CostesHoras As Currency
Dim IT As ListItem
Dim Aux1 As Currency


    On Error GoTo eCargaCostesEuler
    
    If CostesCargados Then Exit Sub
    ListView1.ListItems.Clear
    Me.ListView2.ListItems.Clear
    
    For N = 66 To 71
      '  Label1(N).Caption = ""
    Next
        
    If limpiar Then Exit Sub
    If Text1(0).Text = "" Then Exit Sub
    
    If Me.SSTab1.Tab <> 2 Then Exit Sub
    If AlbaranesDelProyecto = "" Then Exit Sub
    
    CostesCargados = True
    
    
    oldC = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    lblIndicador.Tag = lblIndicador.Caption
    
    
    
    
    
    C1 = "select sreloj.codtraba,nomtraba,fecha,sreloj.codtipor,nomtipor,horainicio,horafin,calculadas,codtipom,numalbar from sreloj left join stipor on sreloj.codtipor=stipor.codtipor"
    C1 = C1 & " left join straba on straba.codtraba=sreloj.codtraba"
    C1 = C1 & " WHERE  (codtipom,numalbar) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " ORDER BY fecha,horainicio"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalCostes = 0 'reutilizo pq bajo se pone a cero
    
    
    N = 0
    While Not miRsAux.EOF
        N = N + 1
        ListView1.ListItems.Add , , miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        ListView1.ListItems(N).SubItems(1) = Format(miRsAux!CodTraba, "0000")
        ListView1.ListItems(N).SubItems(2) = DBLet(miRsAux!NomTraba, "T")
        ListView1.ListItems(N).SubItems(3) = DBLet(miRsAux!codtipor, "T")
        ListView1.ListItems(N).SubItems(4) = DBLet(miRsAux!NomTipor, "T")
        ListView1.ListItems(N).SubItems(5) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        
        If Not IsNull(miRsAux!calculadas) Then
            TotalCostes = TotalCostes + miRsAux!calculadas
            ListView1.ListItems(N).SubItems(7) = Format(miRsAux!calculadas, FormatoCantidad)
            C1 = Format(Int(miRsAux!calculadas), "00") & ":"
            
            
            CostesHoras = Int((miRsAux!calculadas - Int(miRsAux!calculadas)) * 100)
            CostesHoras = Round(CostesHoras * 0.6, 2)
            C1 = C1 & Format(CostesHoras, "00")
            ListView1.ListItems(N).SubItems(6) = C1
            
            
            
        Else
            ListView1.ListItems(N).SubItems(6) = " "
            ListView1.ListItems(N).SubItems(7) = " "
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Label1(63).visible = False
    Label1(63).Caption = Format(TotalCostes, FormatoCantidad)
    If TotalCostes = 0 Then
        C1 = ""
    Else
        C1 = Format(Int(TotalCostes), "00") & ":"
        CostesHoras = Int((TotalCostes - Int(TotalCostes)) * 100)
        CostesHoras = Round(CostesHoras * 0.6, 2)
        C1 = C1 & Format(CostesHoras, "00")
    End If
    Label1(64).visible = False
    Label1(64).Caption = C1

    
    
    
    
    
    
    
    lblIndicador.Caption = "Costes alb."
    lblIndicador.Refresh
    N = 0
    TotalCostes = 0
    CostesHoras = 0
   
    
    'Si tiene horas, las aplicamos aqui
    H = 0
    If Label1(63).Caption <> "" And Label1(63).Caption <> "" Then

        C1 = ImporteFormateado(Label1(63).Caption)
        H = CCur(C1)
        ListView2.ListItems.Add , , "HOR"
        ListView2.ListItems(1).SubItems(1) = "Horas trabajadas"

        For N = 2 To 4
            ListView2.ListItems(1).SubItems(N) = " "
        Next
        ListView2.ListItems(1).SubItems(5) = Format(H, FormatoImporte)
        ListView2.ListItems(1).SubItems(6) = Format(vParamAplic.PrecioHoraCosteEUL, FormatoPrecio)
        H = H * vParamAplic.PrecioHoraCosteEUL
        TotalCostes = TotalCostes + H
        CostesHoras = H
        ListView2.ListItems(1).SubItems(7) = Format(H, FormatoImporte)
        ListView2.ListItems(1).SubItems(8) = " "  'ordenacion
        N = 1
    End If
    
    'En albaranes
    C1 = "select scaalp.numalbar,scaalp.fechaalb,nomprove,codartic,nomartic,cantidad,precioar,importel,scaalp.Codprove from scaalp,slialp  where"
    C1 = C1 & " scaalp.NumAlbar = slialp.NumAlbar And scaalp.FechaAlb = slialp.FechaAlb And scaalp.Codprove = slialp.Codprove"
    C1 = C1 & " and (codtipomv,numalbarV) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " ORDER BY Fechaalb"
    
    Set RS = New ADODB.Recordset
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "ALC"
        IT.SubItems(1) = DBLet(RS!nomprove, "T")
        IT.SubItems(2) = DBLet(RS!Numalbar, "T")
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
               
        If RS!cantidad = 0 Then
            Aux1 = 0
        Else
            Aux1 = RS!ImporteL / RS!cantidad
        End If
        IT.SubItems(6) = Format(Aux1, FormatoPrecio)
        Aux1 = Aux1 - RS!precioar
        If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
        IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numalbar  'ordenacion
        IT.SubItems(9) = RS!codArtic
        
        IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
        TotalCostes = TotalCostes + RS!ImporteL
         
        RS.MoveNext
    Wend
    RS.Close


    'ALbaranes vinculados al pedido
    If PedidosVinculadosEnAlbaranes <> "" Then
        C1 = "select scaalp.numalbar,scaalp.fechaalb,nomprove,codartic,nomartic,cantidad,precioar,importel,scaalp.Codprove from scaalp,slialp  where"
        C1 = C1 & " scaalp.NumAlbar = slialp.NumAlbar And scaalp.FechaAlb = slialp.FechaAlb And scaalp.Codprove = slialp.Codprove"
        C1 = C1 & " and codclien =" & Text1(4).Text
        C1 = C1 & " and numpedV IN (" & PedidosVinculadosEnAlbaranes & ")"
        C1 = C1 & " ORDER BY Fechaalb"

        Set RS = New ADODB.Recordset
        RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            N = N + 1

            Set IT = ListView2.ListItems.Add
            IT.Text = "ALC"
            IT.SubItems(1) = DBLet(RS!nomprove, "T")
            IT.SubItems(2) = DBLet(RS!Numalbar, "T")
            IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
            IT.SubItems(4) = DBLet(RS!NomArtic, "T")
            IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)

            If RS!cantidad = 0 Then
                Aux1 = 0
            Else
                Aux1 = RS!ImporteL / RS!cantidad
            End If
            IT.SubItems(6) = Format(Aux1, FormatoPrecio)
            Aux1 = Aux1 - RS!precioar
            If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
            IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
            IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numalbar  'ordenacion
            IT.SubItems(9) = RS!codArtic

            IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
            TotalCostes = TotalCostes + RS!ImporteL

            RS.MoveNext
        Wend
        RS.Close
    End If


    'FACTURAS PROVEEDOR
    lblIndicador.Caption = "Costes fact."
    lblIndicador.Refresh
    C1 = "select scafpc.numfactu,scafpc.fecfactu,nomprove,codartic,nomartic,cantidad,precioar,importel,scafpc.Codprove,slifpc.numalbar,scafpa.fechaalb from"
    C1 = C1 & " scafpc,scafpa,slifpc  where "
    C1 = C1 & " scafpc.codprove = scafpa.codprove And scafpc.numfactu = scafpa.numfactu And scafpc.fecfactu = scafpa.fecfactu "
    C1 = C1 & " AND scafpc.codprove = slifpc.codprove And scafpc.numfactu = slifpc.numfactu And scafpc.fecfactu = slifpc.fecfactu "
    C1 = C1 & " and scafpa.numalbar = slifpc.numalbar"
    C1 = C1 & " and (codtipomv,numalbarV) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " ORDER BY fecfactu"
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "FAC"
        IT.SubItems(1) = DBLet(RS!nomprove, "T")
        IT.SubItems(2) = DBLet(RS!Numfactu, "T")
        IT.SubItems(3) = Format(RS!FecFactu, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        
        
        If RS!cantidad = 0 Then
            Aux1 = 0
        Else
            Aux1 = RS!ImporteL / RS!cantidad
        End If
        IT.SubItems(6) = Format(Aux1, FormatoPrecio)
        Aux1 = Aux1 - RS!precioar
        If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
        
        
        
        IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
        IT.SubItems(8) = Format(RS!FecFactu, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numfactu  'ordenacion
        TotalCostes = TotalCostes + RS!ImporteL
        IT.SubItems(9) = RS!codArtic
        
        
        IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
        RS.MoveNext
    Wend
    RS.Close

    If PedidosVinculadosEnAlbaranes <> "" Then
        'FACTURAS PROVEEDOR
        lblIndicador.Caption = "Costes fact."
        lblIndicador.Refresh
        C1 = "select scafpc.numfactu,scafpc.fecfactu,nomprove,codartic,nomartic,cantidad,precioar,importel,scafpc.Codprove,slifpc.numalbar,scafpa.fechaalb from"
        C1 = C1 & " scafpc,scafpa,slifpc  where "
        C1 = C1 & " scafpc.codprove = scafpa.codprove And scafpc.numfactu = scafpa.numfactu And scafpc.fecfactu = scafpa.fecfactu "
        C1 = C1 & " AND scafpc.codprove = slifpc.codprove And scafpc.numfactu = slifpc.numfactu And scafpc.fecfactu = slifpc.fecfactu "
        C1 = C1 & " and scafpa.numalbar = slifpc.numalbar"
        C1 = C1 & " and codclien=" & Text1(4).Text
        C1 = C1 & " and numpedV IN (" & PedidosVinculadosEnAlbaranes & ")"
        C1 = C1 & " ORDER BY fecfactu"
        RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            N = N + 1
            
            Set IT = ListView2.ListItems.Add
            IT.Text = "FAC"
            IT.SubItems(1) = DBLet(RS!nomprove, "T")
            IT.SubItems(2) = DBLet(RS!Numfactu, "T")
            IT.SubItems(3) = Format(RS!FecFactu, "dd/mm/yyyy")
            IT.SubItems(4) = DBLet(RS!NomArtic, "T")
            IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
            
            
            If RS!cantidad = 0 Then
                Aux1 = 0
            Else
                Aux1 = RS!ImporteL / RS!cantidad
            End If
            IT.SubItems(6) = Format(Aux1, FormatoPrecio)
            Aux1 = Aux1 - RS!precioar
            If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
            
            
            
            IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
            IT.SubItems(8) = Format(RS!FecFactu, "yymmdd") & Format(RS!Codprove, "00000") & RS!Numfactu  'ordenacion
            TotalCostes = TotalCostes + RS!ImporteL
            IT.SubItems(9) = RS!codArtic
            
            
            IT.Tag = RS!Numalbar & "|" & RS!FechaAlb & "|" & RS!Codprove & "|"
            RS.MoveNext
        Wend
        RS.Close
    End If


    lblIndicador.Caption = "Adicionales"
    lblIndicador.Refresh
    C1 = "select fechamov ,codartic,numlinea ,nomartic ,cantidad ,precioar,round(cantidad *precioar,2) implin FROM slialb_eu "
    C1 = C1 & " WHERE "
    C1 = C1 & " (codtipom,numalbar) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " ORDER BY fechamov"
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "MAT"
        IT.SubItems(1) = " "
        IT.SubItems(2) = "L " & RS!numlinea
        IT.SubItems(3) = Format(RS!FechaMov, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        IT.SubItems(6) = Format(RS!precioar, FormatoPrecio)
        IT.SubItems(7) = Format(RS!implin, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaMov, "yymmdd") & "   " & Format(RS!numlinea, "00") 'ordenacion
        IT.SubItems(9) = RS!codArtic
        TotalCostes = TotalCostes + RS!implin
                 
        RS.MoveNext
    Wend
    RS.Close






    'En este albarane.   NO haria falta linkar con sartic
    C1 = "select scaalb.numalbar,scaalb.fechaalb,nomclien,slialb.codartic,slialb.nomartic,cantidad,preciouc,precoste ,scaalb.codtipom"
    C1 = C1 & " From scaalb, slialb, sartic"
    C1 = C1 & " Where scaalb.NumAlbar = slialb.NumAlbar And scaalb.codtipom = slialb.codtipom And slialb.codArtic = sartic.codArtic"
    C1 = C1 & " and (scaalb.codtipom,scaalb.numalbar) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " and slialb.precoste<>0"
    C1 = C1 & " ORDER BY Fechaalb"
    

    
    
    Set RS = New ADODB.Recordset
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = RS!codtipom
        IT.SubItems(1) = " "
        IT.SubItems(2) = DBLet(RS!Numalbar, "T")
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        
        Aux1 = DBLet(RS!precoste, "N")   ' DBLet(Rs!precioUC, "N")
        Aux1 = Aux1 * DBLet(RS!cantidad, "N")
        Aux1 = Round(Aux1, 2)
        'IT.SubItems(6) = " " & Format(DBLet(Rs!precioUC, "N"), FormatoPrecio)
        IT.SubItems(6) = " " & Format(DBLet(RS!precoste, "N"), FormatoPrecio)
    
        IT.SubItems(7) = Format(Aux1, FormatoImporte)
        IT.SubItems(8) = Format(RS!FechaAlb, "yymmdd") & RS!codtipom & RS!Numalbar  'ordenacion
        TotalCostes = TotalCostes + Aux1
         
        RS.MoveNext
    Wend
    RS.Close




    'SEPT 2018
    lblIndicador.Caption = "Pedidos proveedor."
    lblIndicador.Refresh
    C1 = " select scappr.numpedpr,fecpedpr,nomprove,codartic,nomartic,cantidad,precioar,importel,scappr.Codprove"
    C1 = C1 & " From scappr, slippr where  scappr.numpedpr = slippr.numpedpr "
    C1 = C1 & " and (codtipomv,numalbarV) IN (" & AlbaranesDelProyecto & ")"
    C1 = C1 & " ORDER BY fecpedpr"
    RS.Open C1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        N = N + 1
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "PED"
        IT.SubItems(1) = DBLet(RS!nomprove, "T")
        IT.SubItems(2) = DBLet(RS!numpedpr, "T")
        IT.SubItems(3) = Format(RS!fecpedpr, "dd/mm/yyyy")
        IT.SubItems(4) = DBLet(RS!NomArtic, "T")
        IT.SubItems(5) = Format(RS!cantidad, FormatoImporte)
        
        
        If RS!cantidad = 0 Then
            Aux1 = 0
        Else
            Aux1 = RS!ImporteL / RS!cantidad
        End If
        IT.SubItems(6) = Format(Aux1, FormatoPrecio)
        Aux1 = Aux1 - RS!precioar
        If Abs(Aux1) > 0.05 Then IT.ListSubItems(6).ForeColor = vbRed  'Lleva descuentos
        
        
        
        IT.SubItems(7) = Format(RS!ImporteL, FormatoImporte)
        IT.SubItems(8) = Format(RS!fecpedpr, "yymmdd") & Format(RS!Codprove, "00000") & Format(RS!numpedpr, "000000") 'ordenacion
        TotalCostes = TotalCostes + RS!ImporteL
        IT.SubItems(9) = RS!codArtic
        
        
        IT.Tag = RS!numpedpr & "|" & RS!fecpedpr & "|" & RS!Codprove & "|"
        RS.MoveNext
    Wend
    RS.Close




        
    If ListView2.ListItems.Count > 0 Then
    
'        Label1(67).Caption = "Total costes"
'        Label1(66).Caption = Format(TotalCostes, FormatoImporte)
'        Label1(68).Caption = "Costes horas"
'        Label1(69).Caption = Format(CostesHoras, FormatoImporte)
'        CostesHoras = TotalCostes - CostesHoras
'        Label1(70).Caption = "Costes materiales"
'        Label1(71).Caption = Format(CostesHoras, FormatoImporte)
        
    End If
    
eCargaCostesEuler:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    lblIndicador.Caption = lblIndicador.Tag
    Screen.MousePointer = oldC
End Sub




Private Function CompruebaTotales(DesdeFacturar As Boolean) As Boolean
Dim Totales As Currency
Dim Cad As String

     
    CompruebaTotales = True
    Cad = "count(*)"
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "sum(importel)", "sproyectolin2", "codtipom='ALY' AND numproyec", Text1(0).Text, "N", Cad)
    If Val(Cad) > 0 Then
        'Tiene lineas
        Totales = CCur(CadenaDesdeOtroForm)
        
       
        
        CadenaDesdeOtroForm = ""
        'Vemos el total albaranes
        Cad = "(slialb.codtipom,slialb.numalbar) IN (Select codtipoa,numalbar from  sproyectolin WHERE codtipom='ALY' AND numproyec=" & Text1(0).Text & " ) AND 1 "
        Cad = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", Cad, "1")
        If Cad = "" Then Cad = "0"
           
        If Me.lblAjuste.Caption <> "" Then
            'Habrá que ajustar
            Cad = CCur(Cad) + Me.lblAjuste.Tag
        
        End If
           
           
           
        If CCur(Cad) <> Totales Then
        
           SQL = Format(CCur(Cad) - Totales, FormatoImporte)
        
           Cad = vbCrLf & vbCrLf & "Albaranes vinculados:" & Cad & vbCrLf
           Cad = Cad & "Lineas proyecto   : " & Totales & vbCrLf
           If Me.lblAjuste.Caption <> "" Then Cad = Cad & "Importe " & Me.lblAjuste.Caption & vbCrLf
           Cad = "Distintos sumatorios  albaranes // lineas de impresion." & vbCrLf & Cad
           Cad = Cad & vbCrLf & vbCrLf & "Diferencia: " & SQL
           SQL = ""
           If DesdeFacturar Then
                Cad = Cad & vbCrLf & "¿Continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then CompruebaTotales = False
            Else
                MsgBox Cad, vbExclamation
            End If
        End If
        
        
    Else
        'No ha puesto ninugan linea de impresion
        If DesdeFacturar Then
            Cad = "No ha añadido ninguna linea de impresion." & vbCrLf & vbCrLf & "¿Contiuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then CompruebaTotales = False
        End If
    End If






End Function




























































