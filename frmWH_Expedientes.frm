VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmWH_Expedientes 
   Caption         =   "WHOSE"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   Icon            =   "frmWH_Expedientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   14625
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   11160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   15
      Text            =   "002341"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   8280
      TabIndex        =   11
      Top             =   4440
      Width           =   6135
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   5295
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   5775
         _cx             =   5080
         _cy             =   5080
      End
      Begin VB.CommandButton cmdVerPDF 
         Height          =   495
         Left            =   120
         Picture         =   "frmWH_Expedientes.frx":3482
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTituloDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   25
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.TextBox txtDescObra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   9840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2520
      Width           =   4575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":46F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":55CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":58E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":5C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":6054
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_Expedientes.frx":64A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "654649836"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "961398959"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "José Luis Fernández Rodriguéz"
      Top             =   360
      Width           =   5415
   End
   Begin MSComctlLib.ListView lwPRI 
      Height          =   1095
      Left            =   2160
      TabIndex        =   9
      Tag             =   "Propiedad intelectual"
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "F. Presentacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contestacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acep."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lwObras 
      Height          =   1455
      Left            =   1080
      TabIndex        =   12
      Tag             =   "Expediente"
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Expediente"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Año"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   5892
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Alta"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo presentador"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView lwSGD 
      Height          =   1695
      Left            =   2160
      TabIndex        =   13
      Tag             =   "SGD"
      Top             =   5640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SGD"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Empresa"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F. Presen."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contesta."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Acep."
         Object.Width           =   1058
      EndProperty
   End
   Begin MSComctlLib.ListView lwTitularidad 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Relacion"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lwActuaciones 
      Height          =   2055
      Left            =   2160
      TabIndex        =   20
      Tag             =   "Actuaciones"
      Top             =   7440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "F. Presen."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Horas"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descripcion"
         Object.Width           =   4304
      EndProperty
   End
   Begin MSComctlLib.ListView lwPropCom 
      Height          =   1095
      Left            =   4680
      TabIndex        =   22
      Tag             =   "Propuesta comer."
      ToolTipText     =   "Prop. comercial"
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "F. Presentacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contestacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acep."
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ListView lwContrato 
      Height          =   1095
      Left            =   9480
      TabIndex        =   23
      Tag             =   "Contrato"
      ToolTipText     =   "Contrato"
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "F. Presentacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Rechazo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Aceptado"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lwTrabajadores 
      Height          =   1215
      Left            =   2160
      TabIndex        =   29
      Tag             =   "Actuaciones"
      Top             =   9600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Trabajador"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   17
      Left            =   10920
      Picture         =   "frmWH_Expedientes.frx":8C58
      ToolTipText     =   "Eliminar contrato"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   15
      Left            =   10200
      Picture         =   "frmWH_Expedientes.frx":965A
      ToolTipText     =   "Nuevo contrato"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   16
      Left            =   10560
      Picture         =   "frmWH_Expedientes.frx":A05C
      ToolTipText     =   "Modificar datos"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2160
      Picture         =   "frmWH_Expedientes.frx":AA5E
      ToolTipText     =   "Ver cliente"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgWeb 
      Height          =   495
      Left            =   13200
      Picture         =   "frmWH_Expedientes.frx":B460
      Stretch         =   -1  'True
      Tag             =   "-1"
      ToolTipText     =   "Abrir web"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   12
      Left            =   120
      Picture         =   "frmWH_Expedientes.frx":CDE2
      ToolTipText     =   "Nueva actuación"
      Top             =   7680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador actuación"
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   28
      Top             =   9600
      Width           =   1515
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   4
      Left            =   840
      Picture         =   "frmWH_Expedientes.frx":D7E4
      ToolTipText     =   "Aceptar SGD"
      Top             =   5880
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   3
      Left            =   840
      Picture         =   "frmWH_Expedientes.frx":E1E6
      ToolTipText     =   "Presentacion P.I. aceptado"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "frmWH_Expedientes.frx":EBE8
      ToolTipText     =   "Eliminar"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   14
      Left            =   840
      Picture         =   "frmWH_Expedientes.frx":1543A
      ToolTipText     =   "Eliminar accion"
      Top             =   7680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Actuaciones"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   7440
      Width           =   885
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   13
      Left            =   480
      Picture         =   "frmWH_Expedientes.frx":15E3C
      ToolTipText     =   "Modificar"
      Top             =   7680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   10
      Left            =   1080
      Picture         =   "frmWH_Expedientes.frx":1683E
      ToolTipText     =   "Nueva persona relacionada"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   11
      Left            =   1800
      Picture         =   "frmWH_Expedientes.frx":17240
      ToolTipText     =   "Eliminar"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   8
      Left            =   120
      Picture         =   "frmWH_Expedientes.frx":17C42
      ToolTipText     =   "Nueva propuesta"
      Top             =   5880
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   9
      Left            =   480
      Picture         =   "frmWH_Expedientes.frx":18644
      ToolTipText     =   "Rechazar"
      Top             =   5880
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   6
      Left            =   120
      Picture         =   "frmWH_Expedientes.frx":19046
      ToolTipText     =   "Nueva propuesta"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   7
      Left            =   480
      Picture         =   "frmWH_Expedientes.frx":19A48
      ToolTipText     =   "Rechazar"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmWH_Expedientes.frx":1A44A
      ToolTipText     =   "Modificar"
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmWH_Expedientes.frx":1AE4C
      ToolTipText     =   "Nueva"
      Top             =   2880
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   4
      Index           =   1
      X1              =   120
      X2              =   14520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Obras"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Titularidad"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   14520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Alta gestoras de derechos "
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Alta Registro P.I."
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contrato"
      Height          =   195
      Index           =   5
      Left            =   9480
      TabIndex        =   8
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Propuesta comercial"
      Height          =   195
      Index           =   4
      Left            =   4680
      TabIndex        =   7
      Top             =   840
      Width           =   1680
   End
   Begin VB.Image imgMail 
      Height          =   480
      Left            =   13920
      Picture         =   "frmWH_Expedientes.frx":1B84E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Télefono"
      Height          =   195
      Index           =   3
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Télefono"
      Height          =   195
      Index           =   2
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnContextual 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnPop1 
         Caption         =   "Enviar por email"
         Index           =   0
      End
      Begin VB.Menu mnPop1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPop1 
         Caption         =   "Guardar como ..."
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmWH_Expedientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cliente As Long
Public expediente As Long
Public Ano As Integer

Dim IT As ListItem
Dim Cad As String


Dim PrimVez As Boolean

Dim QUienHaLlamadoAlContextual As Byte
Dim AsuntoContextual As String

'**************************************************
'Esto estara en un .bas
Private Sub CursorParaIconos(ByRef img As Image)
    img.MousePointer = vbCustom
    'Img.MouseIcon = LoadPicture(App.Path & "\hand5.ico")
End Sub





'Private Sub cboEGDA_Click()
'    CargaListviewExpWHOSE Me.lwSGD, Ano, expediente, CByte(Me.cboEGDA.ItemData(Me.cboEGDA.ListIndex))
'
'End Sub



Private Sub cmdVerPDF_Click()
    LanzaVisorMimeDocumento Me.hWnd, AcroPDF1.Tag
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        Set miRsAux = New ADODB.Recordset
        CargaExpedientes
        CargaDatosCliente
        Set miRsAux = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Load()
    PrimVez = True
    Set lwPropCom.SmallIcons = Me.ImageList1
    Set lwContrato.SmallIcons = Me.ImageList1
    Set lwPRI.SmallIcons = Me.ImageList1
    Set lwObras.SmallIcons = Me.ImageList1
    Set lwSGD.SmallIcons = Me.ImageList1
    Set lwActuaciones.SmallIcons = Me.ImageList1
    
    
    
    LimpiarDatos

    

    
    
    
    Dim I
    For I = 0 To Me.Image3.Count - 1
        'El 5 NO existe
        If I <> 5 Then CursorParaIconos Image3(I)
    Next
    CursorParaIconos imgWeb
    CursorParaIconos imgMail
    
    'Solo admin podrá...
    Me.Image3(17).visible = False  'eliminar un contrato
    
    
    
End Sub



Private Sub LimpiarDatos()
    limpiar Me
'    lwPropCom.ListItems.Clear
'    lwContrato.ListItems.Clear
    lblTituloDoc.Caption = ""
    cmdVerPDF.visible = False
    lwPRI.ListItems.Clear
    lwObras.ListItems.Clear
    lwSGD.ListItems.Clear
    lwTitularidad.ListItems.Clear
    lwActuaciones.ListItems.Clear
    lwTrabajadores.ListItems.Clear
End Sub









Private Sub Image1_Click()
    frmFacClientes.VerCliente = Cliente
    frmFacClientes.DatosADevolverBusqueda = ""
    frmFacClientes.Show vbModal
    
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select nomclien,telclie1,telclie2 from sclien where codclien =" & Cliente
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Me.Text1(1).Text = miRsAux!NomClien
    Me.Text1(2).Text = DBLet(miRsAux!telclie1, "T")
    Me.Text1(3).Text = DBLet(miRsAux!telclie2, "T")
    miRsAux.Close
End Sub


Private Sub Image3_Click(Index As Integer)

    If Index > 1 And Index < 15 Then
        'NEcesita obra
        If expediente = 0 Then Exit Sub
    Else
        'NO neesita obra
    End If

    Select Case Index
    Case 0, 1
        If Index = 0 Then
            frmWH_Varios.opcion = 4
            frmWH_Varios.ExtraData2 = Cliente
        Else
            If Me.lwObras.SelectedItem Is Nothing Then Exit Sub
            frmWH_Varios.opcion = 5
            frmWH_Varios.ExtraData2 = Cliente & "|" & lwObras.SelectedItem.Text & "|" & lwObras.SelectedItem.SubItems(1) & "|"
        End If
        
        frmWH_Varios.Show vbModal
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
        If CadenaDesdeOtroForm <> "" Then
            'HA METIDO UN NUEVA OBRA
            If Index = 0 Then
                expediente = RecuperaValor(CadenaDesdeOtroForm, 1)
                Ano = RecuperaValor(CadenaDesdeOtroForm, 2)
                Volver_A_Cargar_Datos = True  'para que refresque los datos en el LW de seleccion
            End If
            CargaExpedientes
            
        End If
   
        
    Case 6
        
            
        'NUEVA. O esta rechazada la anterior
        Cad = ""
        If lwPRI.ListItems.Count > 0 Then
            If Trim(lwPRI.ListItems(1).SubItems(1)) = "" Then
                Cad = "NO hay contestacion a la anterior"
            Else
                If lwPRI.ListItems(1).SubItems(2) = "SI" Then Cad = "YA esta aceptada"
            End If
        End If
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
    
        CadenaDesdeOtroForm = "01/01/2000"
        If lwPRI.ListItems.Count > 0 Then CadenaDesdeOtroForm = Me.lwPRI.ListItems(1).SubItems(1)
        frmWH_Varios.ExtraData2 = Format(Cliente, "000000") & "|" & expediente & "|" & Ano & "|" & CadenaDesdeOtroForm & "|" '& "|1|"
        CadenaDesdeOtroForm = ""
        frmWH_Varios.opcion = 6
        frmWH_Varios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then CargaListviewExpPRI Me.lwPRI, Ano, expediente
 
    Case 3, 4, 7, 9
        'RECHAZOs 7-9
        'Aceptado 3-4
        
        CadenaDesdeOtroForm = ""
        Cad = ""
        If Index = 7 Or Index = 3 Then
            If Me.lwPRI.SelectedItem Is Nothing Then Exit Sub
            If Trim(lwPRI.SelectedItem.SubItems(1)) <> "" Then Cad = "*"
            CadenaDesdeOtroForm = lwPRI.SelectedItem.Text & "|Prop. intelectual|"
        Else
            If Me.lwSGD.SelectedItem Is Nothing Then Exit Sub
            If Trim(lwSGD.SelectedItem.SubItems(3)) <> "" Then Cad = "*"
            CadenaDesdeOtroForm = lwSGD.SelectedItem.SubItems(2) & "|SGD|"
        End If
 
        If Cad <> "" Then
            Cad = "Ya esta contestado"
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
        
        frmWH_Varios.ExtraData2 = CStr(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = ""
        If Index >= 7 Then
            frmWH_Varios.opcion = 13
        Else
            frmWH_Varios.opcion = 14
        End If
        frmWH_Varios.Show vbModal
        
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 7 Or Index = 3 Then
                'PRopiedad intelectual
                'whoobrascli pifcontesta aceptado  where expediente  anoexp  idPI
                Cad = "UPDATE  whoobrasclipi SET fcontesta =" & DBSet(CadenaDesdeOtroForm, "F") & ", aceptado=" & Abs(Index = 3)
                Cad = Cad & " WHERE expediente=" & expediente & " AND anoexp=" & Ano
                Cad = Cad & " AND idPI =" & Mid(Me.lwPRI.SelectedItem.Key, 2)
                If Ejecutar(Cad, False) Then
                    lwPRI.SelectedItem.SubItems(1) = CadenaDesdeOtroForm
                    If Index = 3 Then lwPRI.SelectedItem.SubItems(2) = "SI"
                End If
            Else
                'SGD
                'whoobrasclisgd  fcontesta  aceptado    expediente  anoexp  SGD  IdPres
                Cad = "UPDATE whoobrasclisgd SET fcontesta =" & DBSet(CadenaDesdeOtroForm, "F") & ", aceptado=" & Abs(Index = 4)
                Cad = Cad & " WHERE expediente=" & expediente & " AND anoexp=" & Ano
                
                
                Cad = Cad & " AND SGD =" & lwSGD.SelectedItem.Text
                
                Cad = Cad & " AND IdPres =" & Mid(Me.lwSGD.SelectedItem.Key, 4)
                If Ejecutar(Cad, False) Then
                    lwSGD.SelectedItem.SubItems(3) = CadenaDesdeOtroForm
                    If Index = 4 Then lwSGD.SelectedItem.SubItems(4) = "SI"
                End If
            End If
              
        End If
 
    Case 8

        
        
            
        
        'NUEVA. O esta rechazada la anterior
        Cad = ""
        If lwSGD.ListItems.Count > 0 Then
            If Trim(lwSGD.ListItems(1).SubItems(1)) = "" Then
                Cad = "NO hay contestacion a la anterior"
            Else
                If lwSGD.ListItems(1).SubItems(2) = "SI" Then Cad = "YA esta aceptada"
            End If
        End If
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
        
        
        frmWH_Varios.ExtraData2 = Format(Cliente, "000000") & "|" & expediente & "|" & Ano & "|"
        
        
        CadenaDesdeOtroForm = ""
        frmWH_Varios.opcion = 7
        frmWH_Varios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then CargaListviewExpSGD Me.lwSGD, Ano, expediente
 
    Case 12, 13
        'ACcciones    ACTUACIONES
        If Index = 13 Then
            If Me.lwActuaciones.SelectedItem Is Nothing Then Exit Sub
        End If
        frmWH_Varios.ExtraData2 = Format(Cliente, "000000") & "|" & expediente & "|" & Ano & "|"
        If Index = 13 Then frmWH_Varios.ExtraData2 = frmWH_Varios.ExtraData2 & Mid(Me.lwActuaciones.SelectedItem.Key, 2) & "|"
        CadenaDesdeOtroForm = ""
        frmWH_Varios.opcion = Index - 3 '9 y 10---> 12-3
        frmWH_Varios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then CargaActuaciones
            
                
    Case 10, 11, 2
        
        If Index <> 10 Then
            If Me.lwTitularidad.SelectedItem Is Nothing Then Exit Sub
        End If
        
        If Index = 11 Then
            'ELIMINAR
            Cad = "Va a eliminar el elemento: " & vbCrLf & Me.lwTitularidad.SelectedItem.Text & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            Cad = " DELETE FROM whotitularidadcli  WHERE codclien =" & Cliente
            
            Cad = Cad & " AND idTitularidad =" & Mid(lwTitularidad.SelectedItem.Key, 2)
            If Ejecutar(Cad, False) Then
                CadenaDesdeOtroForm = "OK"
            Else
                CadenaDesdeOtroForm = "" 'MAL. QUe no refresque
            End If
        
        Else
            If Index = 2 Then
                Cad = lwTitularidad.SelectedItem.Text & "|" & lwTitularidad.SelectedItem.SubItems(1) & "|" & Mid(lwTitularidad.SelectedItem.Key, 2) & "|"
                frmWH_Varios.opcion = 12
            Else
                Cad = ""
                frmWH_Varios.opcion = 11
            End If
            frmWH_Varios.ExtraData2 = Format(Cliente, "000000") & "|" & Cad
            frmWH_Varios.Show vbModal
        End If
        
        If CadenaDesdeOtroForm <> "" Then CargaTitularidad
        
    Case 14
        MsgBox "No tiene permisos", vbExclamation
 
 
    Case 15, 16, 17
        'Julio 2104
        'Pueden añadir mas contratos sin necesidad de haber rechazado ninguno de ellos(ni aceptado)
        'Poder modificar las fechas
        If Index <> 15 Then
            If lwContrato.SelectedItem Is Nothing Then Exit Sub
        End If
        
        If Index = 17 Then
            'Eliminar
                
                
        Else
            ' Nuevo  modificar
            CadenaDesdeOtroForm = Cliente & "|"
            frmWH_Varios.ExtraData2 = CStr(CadenaDesdeOtroForm)
            
            If Index = 15 Then
                
                frmWH_Varios.opcion = 15
            Else
                frmWH_Varios.opcion = 16
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Mid(lwContrato.SelectedItem.Key, 2) & "|"
                If Trim(lwContrato.SelectedItem.SubItems(1)) <> "" Then
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0|" & lwContrato.SelectedItem.SubItems(1) & "|"
                Else
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1|" & lwContrato.SelectedItem.SubItems(2) & "|"
                End If
                
                frmWH_Varios.ExtraData2 = CadenaDesdeOtroForm & Me.lwContrato.SelectedItem.Text & "|"  'y la fecha presentacion
            End If
            frmWH_Varios.Show vbModal
            
            If CadenaDesdeOtroForm <> "" Then CargaListviewWHOSE lwContrato, False, False, Cliente, True
            
            CadenaDesdeOtroForm = ""
        End If
                
        
    End Select
End Sub





Private Sub ImgMail_Click()
    Screen.MousePointer = vbHourglass
    Cad = DevuelveDesdeBD(conAri, "maiclie1", "sclien", "codclien", Me.Text1(0).Text)
    If Cad = "" Then
        MsgBox "No tiene direccion de email administración", vbExclamation
    Else
        If LanzaMailGnral(Cad) Then Espera 1
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
    Screen.MousePointer = vbHourglass
    Cad = DevuelveDesdeBD(conAri, "wwwclien", "sclien", "codclien", Me.Text1(0).Text)
    If Cad = "" Then
        MsgBox "No tiene direccion web asignada", vbExclamation
    Else
        Cad = "http://" & Cad
        'If LanzaHomeGnral(Cad) Then Espera 2
        LanzaVisorMimeDocumento Me.hWnd, Cad
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub lwActuaciones_DblClick()
    If lwActuaciones.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwActuaciones.SelectedItem, 5, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad
End Sub

Private Sub lwActuaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    'Trabajadores
    CargaTrabajadoresActuacion
    'Ver documento
    PonerPDF lwActuaciones, 5
    Screen.MousePointer = vbDefault
End Sub

Private Sub lwActuaciones_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    QUienHaLlamadoAlContextual = 5 'actuaciones
    AbrirContextual Button
End Sub

Private Sub lwContrato_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF lwContrato, 1
End Sub

Private Sub lwContrato_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    QUienHaLlamadoAlContextual = 0 'contratos
    AbrirContextual Button
End Sub

Private Sub lwObras_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SeleccionarObra
End Sub

Private Sub lwPRI_DblClick()
 
    If lwPRI.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwPRI.SelectedItem, 3, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad
 
End Sub









'OBRAS
Private Sub CargaExpedientes()
    lwObras.ListItems.Clear
    Cad = "SELECT expediente,anoexp,codclien,nombre,fecaltobra,extension,desRelacion FROM whoobrascli,whorelacioncliobra  "
    Cad = Cad & " WHERE whoobrascli.tipoPresentador=whorelacioncliobra.idRelacion "
    Cad = Cad & "  AND codclien =" & Cliente
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = Me.lwObras.ListItems.Add()
        IT.Text = Format(miRsAux!expediente, "000000")
        IT.SubItems(1) = miRsAux!anoexp
        IT.SubItems(2) = miRsAux!nombre
        IT.SubItems(3) = miRsAux!fecaltobra
        
        IT.SubItems(4) = miRsAux!desRelacion
        
        IT.SmallIcon = DevuelveIconoWHOSE(miRsAux!extension)
        
        IT.Tag = IT.Text & IT.SubItems(1) & "." & miRsAux!extension
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If expediente > 0 Then
        For NumRegElim = 1 To Me.lwObras.ListItems.Count
            If Val(Me.lwObras.ListItems(NumRegElim).Text) = expediente Then
                If Val(Me.lwObras.ListItems(NumRegElim).SubItems(1)) = Ano Then
                    lwObras.ListItems(NumRegElim).Selected = True
                    Set lwObras.SelectedItem = Me.lwObras.ListItems(NumRegElim)
                    SeleccionarObra
                End If
            End If
        Next
    End If
    
End Sub

Private Sub CargaDatosCliente()
    Cad = "select whoexpedientecli.*,nomclien,telclie1,telclie2 from whoexpedientecli inner join sclien on"
    Cad = Cad & " whoexpedientecli.codclien=sclien.codclien WHERE whoexpedientecli.codclien =" & Cliente
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Me.Text1(0).Text = Format(Cliente, "000000")
    Me.Text1(1).Text = miRsAux!NomClien
    Me.Text1(2).Text = DBLet(miRsAux!telclie1, "T")
    Me.Text1(3).Text = DBLet(miRsAux!telclie2, "T")
    
    CargaListviewWHOSE lwPropCom, False, True, Cliente, True
    CargaListviewWHOSE lwContrato, False, False, Cliente, True
    
    'Tienen fecha de aceptacion de las dos cosas
    If Me.lwPropCom.ListItems.Count > 0 Then Me.lwPropCom.ListItems(1).SubItems(2) = miRsAux!fecAceptPropComer
    'If Me.lwContrato.ListItems.Count > 0 Then Me.lwContrato.ListItems(1).SubItems(2) = miRsAux!fecAceptContrato
    
    
    miRsAux.Close
            
    CargaTitularidad
            
End Sub

Private Sub CargaTitularidad()
    lwTitularidad.ListItems.Clear
    If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
    'whotitularidadcli (   codclien,idTitularidad ,nombre ,relacion )
    Cad = " select * from whotitularidadcli WHERE codclien=" & Cliente
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = Me.lwTitularidad.ListItems.Add(, "K" & Format(miRsAux!idTitularidad, "00000"))
        
        IT.Text = miRsAux!nombre
        IT.SubItems(1) = miRsAux!relacion
      
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub

Private Sub SeleccionarObra()
Dim cu As Byte

    cu = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    expediente = CLng(Me.lwObras.SelectedItem.Text)
    Ano = lwObras.SelectedItem.SubItems(1)
    
    Cad = "Expediente = " & expediente & " AND anoexp =" & Ano & " AND codclien"
    txtDescObra.Text = DevuelveDesdeBD(conAri, "Descripcion", "whoobrascli", Cad, CStr(Cliente))
         
    'Cargamos mas cosas
     
    CargaListviewExpPRI Me.lwPRI, Ano, expediente
    CargaListviewExpSGD Me.lwSGD, Ano, expediente
    CargaActuaciones
    
    
    
    
    Me.lblTituloDoc.Caption = ""
    Me.lblTituloDoc.Tag = ""
    Me.lblTituloDoc.Refresh
    cmdVerPDF.visible = False
    ComponenteAdobeVisible False
    
    Screen.MousePointer = cu
    
End Sub


Private Sub ComponenteAdobeVisible(visible As Boolean)
    Me.AcroPDF1.visible = visible
End Sub

Private Sub CargaActuaciones()
    lwActuaciones.ListItems.Clear
    Cad = "Select f_preact,importe,horas,observa,extension,idactua from whoobrascliactua WHERE expediente =  " & expediente & " AND anoexp = " & Ano & " ORDER BY   IdActua desc"
    If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = Me.lwActuaciones.ListItems.Add(, "K" & Format(miRsAux!idActua, "00000"))
        
        IT.Text = Format(miRsAux!f_preact, "dd/mm/yyyy")
        IT.SubItems(1) = " "
        If Not IsNull(miRsAux!Importe) Then IT.SubItems(1) = Format(miRsAux!Importe, FormatoImporte)
        IT.SubItems(2) = " "
        If Not IsNull(miRsAux!Horas) Then IT.SubItems(2) = Format(miRsAux!Horas, FormatoImporte)
        Cad = DBLet(miRsAux!observa, "T") & " "
        Cad = Replace(Cad, vbCrLf, " ")
        IT.SubItems(3) = Cad
        
        IT.SmallIcon = DevuelveIconoWHOSE(miRsAux!extension)
        If miRsAux!extension <> "" Then
            Cad = Format(expediente, "000000") & Ano & Format(miRsAux!idActua, "000") & "." & miRsAux!extension
        Else
            Cad = ""
        End If
        IT.Tag = Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    CargaTrabajadoresActuacion
    
    
End Sub


Private Sub CargaTrabajadoresActuacion()
    lwTrabajadores.ListItems.Clear
    
    If Me.lwActuaciones.ListItems.Count = 0 Then Exit Sub
    If Me.lwActuaciones.SelectedItem Is Nothing Then Exit Sub
    
    Cad = "Select whoobrascliactuatrab.codtraba,nomtraba from whoobrascliactuatrab,straba WHERE whoobrascliactuatrab.codtraba=straba.codtraba AND"
    Cad = Cad & " expediente =  " & expediente & " AND anoexp = " & Ano & " AND idactua = " & Mid(Me.lwActuaciones.SelectedItem.Key, 2)
    If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = Format(miRsAux!CodTraba, "00000")
        Set IT = Me.lwTrabajadores.ListItems.Add(, "K" & Cad)
        IT.Text = Cad
        IT.SubItems(1) = miRsAux!NomTraba
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
End Sub




Private Sub lwContrato_DblClick()
    If lwContrato.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwContrato.SelectedItem, 1, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad
End Sub

Private Sub lwObras_DblClick()
    If lwObras.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwObras.SelectedItem, 2, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad
End Sub

Private Sub lwPRI_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF lwPRI, 3
End Sub

Private Sub lwPRI_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    QUienHaLlamadoAlContextual = 3 'PRI
    AbrirContextual Button
End Sub

Private Sub lwPropCom_DblClick()
    If lwPropCom.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwPropCom.SelectedItem, 0, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad
End Sub



Private Sub lwPropCom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF lwPropCom, 0
End Sub

Private Sub lwPropCom_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    QUienHaLlamadoAlContextual = 1  'proupesta comercial
    AbrirContextual Button
End Sub

Private Sub lwSGD_DblClick()
    If lwSGD.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(lwSGD.SelectedItem, 4, Cliente)
    If Cad <> "" Then LanzaVisorMimeDocumento Me.hWnd, Cad

End Sub




Private Sub lwSGD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF lwSGD, 4
End Sub

Private Sub lwSGD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    QUienHaLlamadoAlContextual = 4 'SGD
    AbrirContextual Button
End Sub

Private Sub lwTitularidad_DblClick()
    
    Image3_Click 2
End Sub



Private Sub PonerPDF(ByRef LISTV As ListView, Tipo As Byte)
    
    If LISTV.SelectedItem Is Nothing Then Exit Sub
        
    If LCase(Right(LISTV.SelectedItem.Tag, 3)) <> "pdf" Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEMClientes(LISTV.SelectedItem, Tipo, Cliente)
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Me.lblTituloDoc.Caption = LISTV.Tag & " " & LISTV.SelectedItem.Tag
        Me.lblTituloDoc.Refresh
        
        cmdVerPDF.visible = False
        If Not CargaArchivo Then
            Me.lblTituloDoc.Caption = "ERROR " & Me.lblTituloDoc.Caption
        Else
            cmdVerPDF.visible = True
        End If
        
        Screen.MousePointer = vbDefault
    End If
    
    
    
    
    
    
End Sub

Private Function CargaArchivo() As Boolean
    
    On Error GoTo eCargaArchivo
    CargaArchivo = False
    
    
    AcroPDF1.LoadFile (Cad)
    AcroPDF1.Tag = Cad
    AcroPDF1.visible = True
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function



Private Sub AbrirContextual(Buton As Integer)

    If Buton = 2 Then
        PopupMenu Me.mnContextual
    End If
End Sub

Private Sub mnPop1_Click(Index As Integer)
Dim I As Integer

    On Error GoTo eEma

    'Solo hay enviar por email
    If Not ItemCorrectoParaEnvioEmail Then Exit Sub
    
    
    'YA se cual es el archivo. Lo copio en \tmp
    If Index = 0 Then
    
            I = Len(vParamAplic.PathDocsWHOSE)
            CadenaDesdeOtroForm = Mid(Cad, I + 2)
            CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, "\", "_")
            
            
            FileCopy Cad, App.Path & "\temp\" & CadenaDesdeOtroForm
            
            
            'LLega aqui, ha ido bien
            CadenaDesdeOtroForm = App.Path & "\temp\" & CadenaDesdeOtroForm
            'Nombre para|email para|Asunto|Mensaje|
            Cad = DevuelveDesdeBD(conAri, "maiclie1", "sclien", "codclien", Me.Text1(0).Text)
            Cad = Me.Text1(1).Text & "|" & Cad & "|" & AsuntoContextual & "|" & "|" & CadenaDesdeOtroForm & "|"
            
            frmEMail.DatosEnvio = Cad
            frmEMail.opcion = 0
            frmEMail.Show vbModal
    
    Else
    
            I = InStrRev(Cad, ".")
            CadenaDesdeOtroForm = Mid(Cad, I + 1)
     
            cd1.Filter = "Archivo " & CadenaDesdeOtroForm & "|*." & CadenaDesdeOtroForm
            cd1.FileName = ""
            'cd1.InitDir = "c:\"
            cd1.CancelError = False
            cd1.ShowSave
            If cd1.FileName = "" Then Exit Sub
            FileCopy Cad, cd1.FileName
    
    
    
    End If
    
    
    Exit Sub
eEma:
    MuestraError Err.Number, Err.Description
End Sub



Private Function ItemCorrectoParaEnvioEmail() As Boolean
        
    ItemCorrectoParaEnvioEmail = False
    AsuntoContextual = ""
    
    Set IT = Nothing
    Select Case QUienHaLlamadoAlContextual
    Case 5
        If Me.lwActuaciones.ListItems.Count > 0 Then
            If Not Me.lwActuaciones.SelectedItem Is Nothing Then
                Set IT = lwActuaciones.SelectedItem
                AsuntoContextual = "Actuacion: "
            End If
        End If
        
    Case 0
           
        If Me.lwContrato.ListItems.Count > 0 Then
            If Not Me.lwContrato.SelectedItem Is Nothing Then
                Set IT = lwContrato.SelectedItem
                AsuntoContextual = "Contrato: "
            End If
        End If
        AsuntoContextual = "Contrato: "
    Case 4
        If Me.lwSGD.ListItems.Count > 0 Then
            If Not Me.lwSGD.SelectedItem Is Nothing Then
                Set IT = lwSGD.SelectedItem
                AsuntoContextual = "Alta SGD: "
            End If
        End If
    Case 3
         
        If Me.lwPRI.ListItems.Count > 0 Then
            If Not Me.lwPRI.SelectedItem Is Nothing Then
                Set IT = lwPRI.SelectedItem
                AsuntoContextual = "Presentacion P.I.: "
            End If
        End If
        
    Case 1
         
        If Me.lwPropCom.ListItems.Count > 0 Then
            If Not Me.lwPropCom.SelectedItem Is Nothing Then
                Set IT = lwPropCom.SelectedItem
                AsuntoContextual = "Propuesta comercial: "
            End If
        End If
    End Select
    
    If IT Is Nothing Then
        Cad = ""
    Else
        Cad = DevuelveNombreArhivoITEMClientes(IT, QUienHaLlamadoAlContextual, Val(Me.Text1(0).Text))
    End If
    If Cad = "" Then Exit Function
    
    AsuntoContextual = AsuntoContextual & " Exp " & expediente & " / " & Ano & " - -" & IT.Text
    
    ItemCorrectoParaEnvioEmail = True
    
End Function




