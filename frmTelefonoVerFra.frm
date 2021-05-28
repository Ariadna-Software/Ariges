VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelefonoVerFra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver datos prefactura"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   31
      Top             =   45
      Width           =   1875
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   135
         TabIndex        =   32
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Suma datos seleccionados"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exportar csv"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
            EndProperty
         EndProperty
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Teléfonos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   8595
      TabIndex        =   29
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   6675
      TabIndex        =   28
      Top             =   840
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ocultar cero"
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
      Left            =   13005
      TabIndex        =   27
      Top             =   840
      Width           =   1545
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
      Index           =   9
      Left            =   7170
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
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
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
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
      Height          =   615
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "frmTelefonoVerFra.frx":0000
      Top             =   2925
      Width           =   6510
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
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5985
      Top             =   675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   1
      Left            =   195
      Picture         =   "frmTelefonoVerFra.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Suma datos seleccionados"
      Top             =   270
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   0
      Left            =   795
      Picture         =   "frmTelefonoVerFra.frx":0A08
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exportar csv"
      Top             =   270
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   2505
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
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
      Left            =   12015
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1260
      TabIndex        =   10
      Top             =   135
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
      Index           =   4
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2355
      Left            =   6675
      TabIndex        =   7
      Top             =   1200
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   4154
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Concepto"
         Object.Width           =   9852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   3422
      EndProperty
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
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4095
      Width           =   2070
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
      Height          =   375
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   6540
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   1980
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6045
      Left            =   120
      TabIndex        =   13
      Top             =   4860
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   10663
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Destino"
         Object.Width           =   3023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha-Hora"
         Object.Width           =   3052
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Codigo"
         Object.Width           =   1833
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descripcion"
         Object.Width           =   8006
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Duracion"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unidad"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   3199
      EndProperty
   End
   Begin MSComctlLib.ListView lwTelef 
      Height          =   2355
      Left            =   6675
      TabIndex        =   30
      Top             =   1200
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   4154
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
         Text            =   "Telefono"
         Object.Width           =   8617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   960
      Left            =   120
      Top             =   3630
      Width           =   14385
   End
   Begin VB.Label Label1 
      Caption         =   "Vta plazos"
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
      Left            =   7185
      TabIndex        =   26
      Top             =   3765
      Width           =   1155
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   11
      Left            =   4905
      TabIndex        =   24
      Top             =   3765
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vta plazos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Exento"
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
      Left            =   2580
      TabIndex        =   20
      Top             =   3765
      Width           =   1200
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
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2370
      Width           =   6510
   End
   Begin VB.Label Label1 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL FACTURA"
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
      Left            =   12015
      TabIndex        =   12
      Top             =   3780
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "IVA"
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
      Left            =   9495
      TabIndex        =   9
      Top             =   3765
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Imponible"
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
      TabIndex        =   6
      Top             =   3765
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2505
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tfno / agrupación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2040
   End
End
Attribute VB_Name = "frmTelefonoVerFra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Where2 As String   'fichero|where|
Public TieneAlbaranes As Boolean
Public esAgrupacion As Boolean



Private UnaVez As Boolean
Dim cad As String



Private Sub Check1_Click()
    If UnaVez Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargarDatos True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub cmdExport_Click(Index As Integer)
Dim SumaCantidad As Currency
Dim SumaImporte As Currency
Dim Seleccionados As Integer
    'Sumatorio selecccion
    If Index = 0 Then
        'Exportar
        ExportaCSV
    Else
        Seleccionados = 0
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Selected Then Seleccionados = Seleccionados + 1
        Next
        If Seleccionados = 0 Then
            MsgBox "Seleccione algún elemento", vbExclamation
            Exit Sub
        End If
        
        cad = "|"
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Selected Then
                If InStr(1, cad, "|" & Me.ListView2.ListItems(NumRegElim).SubItems(2) & "|") = 0 Then cad = cad & ListView2.ListItems(NumRegElim).SubItems(2) & "|"
            End If
        Next
        'Ya tengo los distintos conceptos
        'AAhora busco
        cad = Mid(cad, 2) 'quito el primer pipe
        CadenaDesdeOtroForm = ""
        Do
            davidNumalbar = InStr(1, cad, "|")
            If davidNumalbar = 0 Then
                cad = ""
            Else
                pPdfRpt = Mid(cad, 1, davidNumalbar - 1)
                cad = Mid(cad, davidNumalbar + 1)
                pImprimeDirecto = True
                SumaCantidad = 0
                SumaImporte = 0
                davidNumalbar = 0
                For NumRegElim = 1 To ListView2.ListItems.Count
                    If ListView2.ListItems(NumRegElim).Selected Then
                        If Me.ListView2.ListItems(NumRegElim).SubItems(2) = pPdfRpt Then
                            If pImprimeDirecto Then
                                CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & Me.ListView2.ListItems(NumRegElim).SubItems(3)
                                pImprimeDirecto = False
                            End If
                            davidNumalbar = davidNumalbar + 1
                            SumaCantidad = SumaCantidad + CCur(Me.ListView2.ListItems(NumRegElim).SubItems(4))
                            If Me.ListView2.ListItems(NumRegElim).SubItems(6) <> "" Then SumaImporte = SumaImporte + CCur(Me.ListView2.ListItems(NumRegElim).SubItems(6))
                        End If
                    End If
                Next
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & Format(davidNumalbar, "000") & "      Cantidad:    " & Format(SumaCantidad, FormatoCantidad)
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "         Importe:     " & Format(SumaImporte, FormatoPrecio) & vbCrLf
            End If
        Loop Until cad = ""
        pPdfRpt = "": pImprimeDirecto = False: davidNumalbar = 0
        CadenaDesdeOtroForm = "Resumen datos seleccionados (" & Seleccionados & ") " & vbCrLf & CadenaDesdeOtroForm
        MsgBox CadenaDesdeOtroForm, vbInformation
        CadenaDesdeOtroForm = ""
    End If
End Sub



Private Sub Form_Activate()
Dim RS As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    If UnaVez Then
        Check1.Value = CheckValueLeer(Me.Name)
        UnaVez = False
        Me.Tag = Where2
        
        CargarDatos False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    UnaVez = True
    Me.Icon = frmPpal.Icon
    ListView2.Tag = 0
    
    '++
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 32 ' suma datos seleccionados
        .Buttons(3).Image = 30 ' Exportar csv
        .Buttons(5).Image = 15 ' Salir
    End With
    
    
    Label1(12).visible = vParamAplic.TelefoniaVtaPlazos
    Text1(9).visible = vParamAplic.TelefoniaVtaPlazos
    Label1(10).visible = vParamAplic.TelefoniaVtaPlazos
    Text1(7).visible = vParamAplic.TelefoniaVtaPlazos
    
    
    Me.lwTelef.visible = esAgrupacion
    Option1(1).visible = esAgrupacion
    limpiar Me
    Text1(1).Text = "          L e y e n d o    B.D."
    
End Sub

Private Sub CargarDatos(SoloConsumos As Boolean)
Dim Importe1 As Currency
Dim IVA As Currency
Dim Aux As Currency
Dim IT As ListItem

    

    Set miRsAux = New ADODB.Recordset
    
    
    
    'Cabecera
    Where2 = Me.Tag
    cad = "Select * from tel_cab_factura WHERE  " & RecuperaValor(Where2, 2)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not SoloConsumos Then
        'NO PUESER EOF
        Text1(0).Text = miRsAux!Telefono
             cad = ""
             If Not IsNull(miRsAux!apellido1) Then cad = miRsAux!apellido1
             If Not IsNull(miRsAux!apellido2) Then cad = Trim(cad & " " & miRsAux!apellido2)
             If Not IsNull(miRsAux!Nombre) Then
                 If cad <> "" Then cad = cad & ","
                 cad = Trim(cad & " " & Trim(miRsAux!Nombre))
             End If
        Text1(1).Tag = cad
        
        Text1(2).Text = miRsAux!Fecha
        Text1(3).Text = Format(miRsAux!BaseImponible, "#,##0.00") ' Cuota Total
        Text1(4).Text = Format(miRsAux!Cuota, "#,##0.00") ' Cuota Total
        Text1(5).Text = Format(miRsAux!total, "#,##0.00") ' Cuota Total
        Text1(6).Text = ""
        If DBLet(miRsAux!base_exenta, "N") <> 0 Then Text1(6).Text = Format(miRsAux!base_exenta, "#,##0.00") ' Cuota Total
        
    End If
    
    If vParamAplic.TieneTelefonia2 = 3 Then
        cad = FormatoPrecio
    Else
        cad = FormatoImporte
    End If
    
    ListView1.ListItems.Clear
    CargaLwTelefonia Me.ListView1, miRsAux!Serie, miRsAux!Ano, miRsAux!NumFact, cad, Me.Check1.Value
    
    
    lwTelef.ListItems.Clear
    If esAgrupacion Then
        CargaLwTelefoniaAgrupadoTfono lwTelef, miRsAux!Serie, miRsAux!Ano, miRsAux!NumFact, cad, Me.Check1.Value
            
    Else
        lwTelef.Tag = "'" & Text1(0).Text & "'"
    End If
    
    Importe1 = 0
    IVA = 0
    
    miRsAux.Close
    
'    If Not SoloConsumos Then
'        Text1(1).Text = Text1(1).Tag
'        Set miRsAux = Nothing
'        Exit Sub
'    End If
'
'
    
    
    
    
    
    
    
    
    
    Label1(8).Caption = ""
    Label1(8).ToolTipText = ""
    If TieneAlbaranes Then
    
        cad = "Select nomclien,scaalb.numalbar,slialb.codartic,slialb.nomartic,importel,codigiva,factursn from scaalb left join slialb on scaalb.codtipom=slialb.codtipom"
        cad = cad & " AND scaalb.numalbar = slialb.numalbar left join sartic on slialb.codartic = sartic.codartic  "
        
        cad = cad & " WHERE scaalb.codtipom='ALT' AND referenc in (" & lwTelef.Tag & ") "
        
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If miRsAux.EOF Then
            MsgBox "Albaranes NO encontrados", vbExclamation
        Else
            If miRsAux!NomClien <> Text1(1).Tag Then
                '
                Text1(1).ForeColor = vbRed
                Label1(8).Caption = miRsAux!NomClien
                Label1(8).ToolTipText = "No coincide nombre ALBARAN con el del telefomo"
            Else
                Text1(1).ForeColor = vbBlack
            End If
            While Not miRsAux.EOF
                    
                cad = "** " & Format(miRsAux!Numalbar, "0000") & " - " & miRsAux!NomArtic & "(" & miRsAux!codArtic & ")"
                If miRsAux!factursn = 0 Then cad = cad & "  ""NO"""
                
                ListView1.ListItems.Add , , cad
                If miRsAux!factursn = 0 Then
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "0"
                    ListView1.ListItems(ListView1.ListItems.Count).ForeColor = vbRed
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems(1).ForeColor = vbRed
                    ListView1.ListItems(ListView1.ListItems.Count).ToolTipText = "No tiene marca de facturar"
                Else
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Format(miRsAux!ImporteL, "#,##0.00")
                    ListView1.ListItems(ListView1.ListItems.Count).ForeColor = vbBlue
                    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems(1).ForeColor = vbBlue
                    
                    cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", DBLet(miRsAux!Codigiva, "N"))
                    If cad = "" Then cad = "0"
                    Aux = (miRsAux!ImporteL * CCur(cad)) / 100
                    IVA = IVA + Round2(Aux, 2)
                    Importe1 = Importe1 + miRsAux!ImporteL
                End If
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
        
        
        Text1(8).Text = Format(Importe1, FormatoImporte)  'Solo la base
        
    End If
    
    
    DoEvents
    Where2 = RecuperaValor(Where2, 1)
    
    DetalleLlamada 1
    
    
    
    'Si lleva venta plazos
    If vParamAplic.TelefoniaVtaPlazos Then
        cad = "select ArtPlazos,PlazosMeses,ImportePlazo,nomartic,codigiva,idtelefono from sclientfno left join sartic on artplazos=codartic "
        cad = cad & " where PlazosMeses>0 AND idtelefono IN (" & lwTelef.Tag & ")"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Text1(7).Text = ""
        Importe1 = 0
        NumRegElim = 0
        While Not miRsAux.EOF
            
            If esAgrupacion Then
                If Text1(7).Text <> "" Then Text1(7).Text = Text1(7).Text & vbCrLf
                cad = " (" & miRsAux!idtelefono & "     " & Format(miRsAux!ImportePlazo, FormatoImporte) & "€ )"
                NumRegElim = NumRegElim + 1
            Else
                cad = ""
            End If
            cad = miRsAux!artplazos & "   " & DBLet(miRsAux!NomArtic, "T") & cad
            Text1(7).Text = Text1(7).Text & cad
            cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", DBLet(miRsAux!Codigiva, "N"))
            If cad = "" Then cad = "0"
            Aux = Round((miRsAux!ImportePlazo * CCur(cad)) / 100, 2)
            Importe1 = Importe1 + miRsAux!ImportePlazo
            IVA = IVA + Aux
            
            miRsAux.MoveNext
    
        Wend
        miRsAux.Close
        If NumRegElim > 0 Then
            Text1(9).Text = Format(Importe1, FormatoImporte)
            Label1(10).Caption = "Vta plazos " & IIf(NumRegElim > 2, CStr("(" & NumRegElim & ")"), "")
        End If
    End If
    
    
    If IVA <> 0 Then
        Aux = ImporteFormateado(Text1(4).Text)
        Aux = Aux + IVA
        Text1(4).Text = Format(Aux, FormatoImporte)
    End If
    
    If IVA <> 0 Or Text1(2).Text <> "" Or Text1(3).Text <> "" Then
        Importe1 = ImporteFormateado(Text1(3).Text)
        Aux = ImporteFormateado(Text1(8).Text)
        Importe1 = Importe1 + Aux
    
        Aux = ImporteFormateado(Text1(9).Text)
        Importe1 = Importe1 + Aux
    
        'NUEVO IVA
        Aux = ImporteFormateado(Text1(4).Text)
        Importe1 = Importe1 + Aux
    
        Text1(5).Text = Format(Importe1, FormatoImporte)
    End If
    Text1(1).Text = Text1(1).Tag
    
    Set miRsAux = Nothing
End Sub



Private Sub DetalleLlamada(Orden As Byte)
Dim I As Integer
    'Detalle llamada
    If ListView2.Tag = Orden Then Exit Sub
    If Orden = 127 Then Orden = 1
    Me.ListView2.ListItems.Clear
    ListView2.Tag = Orden
    'Cad = "select Numero_llamado,Fecha,Hora_inicio,Unidad_de_medida,Codigo_de_trafico,"
    cad = "SELECT Numero_llamado,Fecha,Codigo_de_trafico, Tipo_de_trafico,Cantidad_medida_originada,"
    cad = cad & " Unidad_de_medida,Importe,Hora_inicio from telefono.detalle_de_llamadas WHERE "
    cad = cad & "  fichero='" & Where2 & "' and   Numero_de_telefono "
    If esAgrupacion Then
        'Los he cargado en los datos agrupados
        cad = cad & " IN ("
        If Me.lwTelef.ListItems.Count = 0 Then
            cad = cad & "'no'"
        Else
            NumRegElim = 0
            For I = 1 To Me.lwTelef.ListItems.Count
                If lwTelef.ListItems(I).Checked Then
                    NumRegElim = NumRegElim + 1
                    If NumRegElim > 1 Then cad = cad & ", "
                    cad = cad & DBSet(lwTelef.ListItems(I).Text, "T")
                End If
            Next
            If NumRegElim = 0 Then cad = cad & "'NO'"
        End If
        cad = cad & ")"
    Else
        cad = cad & " = " & DBSet(Text1(0).Text, "T")
    End If
    cad = cad & " and fecha<>'0000'"
    'Cad = Cad & " ORDER BY fecha,Hora_inicio"
    cad = cad & " ORDER BY " & Orden + 1
    If Orden <> 1 Then
        cad = cad & ",2,8"
    Else
        cad = cad & ",8"
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
            'Cad = "** " & Format(miRsAux!NumAlbar, "0000") & " - " & miRsAux!NomArtic & "(" & miRsAux!codArtic & ")"
            ListView2.ListItems.Add , , DBLet(miRsAux!Numero_llamado, "T") & " "
            With ListView2.ListItems(ListView2.ListItems.Count)
                    
                    .SubItems(1) = Mid(miRsAux!Fecha, 3, 2) & "/" & Mid(miRsAux!Fecha, 1, 2) & "  " & Mid(miRsAux!Hora_inicio, 1, 2) & ":" & Mid(miRsAux!Hora_inicio, 3, 2)
                    .SubItems(2) = miRsAux!Codigo_de_trafico
                    .SubItems(3) = miRsAux!Tipo_de_trafico
                    .SubItems(4) = Format(miRsAux!Cantidad_medida_originada, "#,##0.00")
                    .SubItems(5) = miRsAux!Unidad_de_medida
                    .SubItems(6) = Format(miRsAux!Importe, "#,##0.0000")
                
                
            End With
            miRsAux.MoveNext
        Wend

    miRsAux.Close
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.Check1.Value
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    DetalleLlamada ColumnHeader.Index - 1
End Sub



Private Sub ExportaCSV()
Dim NF As Integer

    On Error GoTo eExp
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        
        'Cabecera
        cad = ""
        For NumRegElim = 1 To ListView2.ColumnHeaders.Count
            cad = cad & ";""" & ListView2.ColumnHeaders(NumRegElim).Text & """"
        Next NumRegElim
        Print #NF, Mid(cad, 2)
    
    
        'Lineas
        For NumRegElim = 1 To ListView2.ListItems.Count
            cad = """" & Trim(ListView2.ListItems(NumRegElim)) & """"
            For davidNumalbar = 1 To ListView2.ColumnHeaders.Count - 1
                cad = cad & ";""" & Trim(ListView2.ListItems(NumRegElim).SubItems(davidNumalbar)) & """"
            Next davidNumalbar
            Print #NF, cad
            
            
        Next NumRegElim

        
        Close #NF



    
        cd1.Filter = "Archivo csv|*.csv"
        'Ofertare un nombre
        cd1.FileName = ""
        'cd1.InitDir = "c:\"
        cd1.CancelError = False
        cd1.ShowSave
        If cd1.FileName = "" Then Exit Sub
        
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("Ya existe el fichero: " & cd1.FileName & vbCrLf & vbCrLf & "¿Reemplazar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        FileCopy App.Path & "\docum.csv", cd1.FileName


    Exit Sub
eExp:
    MuestraError Err.Number, Err.Description
    
End Sub

Private Sub lwTelef_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    DetalleLlamada 127
End Sub

Private Sub Option1_Click(Index As Integer)
    If UnaVez Then Exit Sub
    
    Me.Check1.visible = Option1(0).Value
    
    If Index = 0 Then
        ListView1.Left = Me.lwTelef.Left
    Else
        ListView1.Left = 12000
    End If
    
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            'suma datos seleccionados
            cmdExport_Click (1)
        Case 3:
            'exportar csv
            cmdExport_Click (0)
        Case 5:
            'salir
            Command1_Click
    End Select

End Sub
