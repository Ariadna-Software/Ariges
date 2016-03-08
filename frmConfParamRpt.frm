VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfParamRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10755
   Icon            =   "frmConfParamRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   300
      TabIndex        =   26
      Top             =   4920
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3836
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Linea"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   3422
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   6885
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "Fichero envio mail|T|N|||scryst|pdfrpt|||"
      Text            =   "Text1"
      Top             =   1680
      Width           =   3510
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Impresion directa"
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Tag             =   "Impr  directa|N|N|||scryst|imprimedirecto|||"
      Top             =   720
      Width           =   1935
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   195
      Left            =   6600
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   300
      MaxLength       =   140
      TabIndex        =   8
      Tag             =   "Linea pie 2|T|S|||scryst|lineapi2|||"
      Text            =   "Text1"
      Top             =   2910
      Width           =   10245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   300
      MaxLength       =   140
      TabIndex        =   11
      Tag             =   "Linea pie 5|T|S|||scryst|lineapi5|||"
      Text            =   "Text1"
      Top             =   4080
      Width           =   10245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   300
      MaxLength       =   140
      TabIndex        =   10
      Tag             =   "Linea pie 4|T|S|||scryst|lineapi4|||"
      Text            =   "Text1"
      Top             =   3690
      Width           =   10245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   300
      MaxLength       =   140
      TabIndex        =   9
      Tag             =   "Linea pie 3|T|S|||scryst|lineapi3|||"
      Text            =   "Text1"
      Top             =   3300
      Width           =   10245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9435
      TabIndex        =   13
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   360
      TabIndex        =   21
      Top             =   7200
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   2280
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   300
      MaxLength       =   140
      TabIndex        =   7
      Tag             =   "Linea pie 1|T|S|||scryst|lineapi1|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   10245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   9900
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "Revisión ISO|N|S|0|99|scryst|codigrev|00||"
      Text            =   "Te"
      Top             =   1110
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Código ISO|T|S|||scryst|codigiso|||"
      Text            =   "Text1"
      Top             =   1110
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1180
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Fichero rpt|T|N|||scryst|documrpt|||"
      Text            =   "Text1"
      Top             =   1740
      Width           =   3750
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1180
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Descripción|T|N|||scryst|nomcryst|||"
      Text            =   "Text1"
      Top             =   1170
      Width           =   3765
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9435
      TabIndex        =   14
      Top             =   7320
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1180
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código Documento|N|N|||scryst|codcryst||S|"
      Text            =   "Text"
      Top             =   720
      Width           =   765
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insertar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ultimo"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9240
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Image Image3 
      Height          =   240
      Index           =   1
      Left            =   2280
      Picture         =   "frmConfParamRpt.frx":000C
      ToolTipText     =   "Modificar"
      Top             =   4560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   2
      Left            =   2640
      Picture         =   "frmConfParamRpt.frx":0A0E
      ToolTipText     =   "Eliminar "
      Top             =   4560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   0
      Left            =   1920
      Picture         =   "frmConfParamRpt.frx":1410
      ToolTipText     =   "Nueva "
      Top             =   4560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informes disponibles"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   27
      Top             =   4560
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fichero envio mail"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   25
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Líneas para el  Pie del Informe"
      Height          =   255
      Index           =   7
      Left            =   300
      TabIndex        =   23
      Top             =   2235
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Revisión ISO"
      Height          =   255
      Index           =   6
      Left            =   8895
      TabIndex        =   20
      Top             =   1155
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Código ISO"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   19
      Top             =   1155
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero rpt"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   18
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   17
      Top             =   1170
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Dim HaDevueltoDatos  As Boolean
Dim CadB As String

Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar


Private Sub Check1_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim vParamRpt As CParamRpt 'Clase Parametros para Reports
Dim Cad As String, Indicador As String
Dim actualiza As Boolean

    If DatosOk Then
        'Modifica datos en la Tabla: scryst
'        I = ModificaDesdeFormulario(Me)
        'Actualizar campos de la clase
            Set vParamRpt = New CParamRpt
            vParamRpt.codigo = Text1(0).Text
            vParamRpt.Descripcion = Text1(1).Text
            vParamRpt.Documento = Text1(2).Text
            vParamRpt.PDFrpt = Text1(10).Text
            vParamRpt.CodigoISO = Text1(3).Text
            If Trim(Text1(4).Text) <> "" Then
                vParamRpt.CodigoRevision = CInt(Text1(4).Text)
            Else
                vParamRpt.CodigoRevision = -1
            End If
            vParamRpt.LineaPie1 = Text1(5).Text
            vParamRpt.LineaPie2 = Text1(6).Text
            vParamRpt.LineaPie3 = Text1(7).Text
            vParamRpt.LineaPie4 = Text1(8).Text
            vParamRpt.LineaPie5 = Text1(9).Text
            vParamRpt.ImprimeDirecto = Check1.Value
        If Modo = 3 Then 'INSERTAR
            actualiza = vParamRpt.Insertar
        ElseIf Modo = 4 Then 'MODIFICAR
            actualiza = vParamRpt.Modificar(Text1(0).Text)
            TerminaBloquear
        End If
        Set vParamRpt = Nothing
        If actualiza = 0 Then 'Inserta o Modifica
            Cad = "codcryst=" & Text1(0).Text
            If SituarData(Data1, Cad, Indicador) Then
                PonerModo 2
                Me.lblIndicador.Caption = Indicador
            End If
        End If
        PonerFocoBtn Me.cmdSalir
    End If
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 3 'Insertar
            LimpiarCampos
            PonerModo 0
            'PonerFoco Text1(0)
        Case 4 'Modificar
            TerminaBloquear
            If Data1.Recordset.EOF Then
                PonerModo 0
                LimpiarCampos
            Else
                PonerCampos
                PonerModo 2
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            End If
    End Select
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'PonerCadenaBusqueda
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        btnPrimero = 11
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(4).Image = 3   'Anyadir
        .Buttons(5).Image = 4   'Modificar
        .Buttons(8).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    NombreTabla = "scryst"
    Ordenacion = " ORDER BY codcryst"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    PonerFoco Text1(0)
    limpiar Me
    Me.chkVistaPrevia.Value = 1
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(2).Enabled = False
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Image3_Click(Index As Integer)

   

    If Modo <> 0 Then Exit Sub
    
    If Index > 0 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If
    
    
    If Index = 2 Then
        'Eliminar
        CadenaDesdeOtroForm = "Va a eliminar el Informe : " & vbCrLf & "Descripcion: " & ListView1.SelectedItem.SubItems(1)
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "Informe(RPT) : " & ListView1.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes Then
              CadenaConsulta = "DELETE FROM scryst2 WHERE codcryst = " & Data1.Recordset!codcryst
              CadenaConsulta = CadenaConsulta & " AND linea =" & ListView1.SelectedItem.Text
              If ejecutar(CadenaConsulta, False) Then ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    Else
        'NUEVO - MODIFICAR
        If Index = 0 Then
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(linea)", "scryst2", "codcryst", CStr(Data1.Recordset!codcryst))
            If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
            CadenaDesdeOtroForm = CStr(Val(CadenaDesdeOtroForm) + 1)
            CadenaConsulta = "1|" & CadenaDesdeOtroForm & "|||"
        Else
            CadenaConsulta = "0|" & ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(2) & "|"
        End If
        
        
        CadenaDesdeOtroForm = ""
        frmListado5.OpcionListado = 6
        frmListado5.OtrosDatos = CadenaConsulta
         frmListado5.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 0 Then
                CadenaConsulta = DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T") & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ")"
                CadenaConsulta = Data1.Recordset!codcryst & "," & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & CadenaConsulta
                CadenaConsulta = "INSERT INTO scryst2(codcryst,linea,descriprp,nomcryst) VALUES (" & CadenaConsulta
                ejecutar CadenaConsulta, False
                ListView1.ListItems.Clear
                CargaSubRpt
            Else
                CadenaConsulta = "UPDATE scryst2 SET descriprp = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T")
                CadenaConsulta = CadenaConsulta & ",nomcryst = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
                CadenaConsulta = CadenaConsulta & " WHERE codcryst = " & Data1.Recordset!codcryst
                CadenaConsulta = CadenaConsulta & " AND linea =" & RecuperaValor(CadenaDesdeOtroForm, 1)
                ejecutar CadenaConsulta, False
                
                ListView1.SelectedItem.SubItems(1) = RecuperaValor(CadenaDesdeOtroForm, 2)
                ListView1.SelectedItem.SubItems(2) = RecuperaValor(CadenaDesdeOtroForm, 3)
            End If
        
        
        End If
    End If
    
End Sub

Private Sub mnModificar_Click()
    If Text1(0).Text = "" Then Exit Sub
     If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
     Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    If Index = 4 Then 'cod. ISO
        If Text1(Index).Text = "" Then Exit Sub
        If Not PonerFormatoEntero(Text1(Index)) Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            'BUSCAR
            
        
        Case 2
            'Ver todos
            BotonVerTodos
        Case 4  'Anyadir
            mnNuevo_Click
        Case 5  'Modificar
            mnModificar_Click
        Case 8 'Salir
           mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Desplazamiento Registros
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub



Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    'Bloquear el código que es clave primaria
    BloquearTxt Text1(0), True
    'Si no es root o administradar no Mofificar la descripcion del documento
    If (vUsu.Nivel <> 0 And vUsu.Nivel <> 1) Then BloquearTxt Text1(1), True
    
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    DatosOk = False
    B = CompForm(Me, 1)
    DatosOk = B
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdSalir.visible = B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    
    Me.ListView1.ListItems.Clear
    
    If Val(Data1.Recordset!LlevaMulitInformes) = 1 Then
        Image3(0).visible = True
        CargaSubRpt
    Else
        Image3(0).visible = False
    End If
    Image3(1).visible = Image3(0).visible
    Image3(2).visible = Image3(0).visible

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.ListView1.ListItems.Clear
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
   
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    B = (Kmodo = 2) Or (Kmodo = 0)
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    
    '------------------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    PonerBotonCabecera Not B
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
   
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearChecks Me, Modo
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    B = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not B 'Insertar
    Me.mnNuevo.Enabled = Not B
    Me.Toolbar1.Buttons(2).Enabled = (Not B) 'Modificar
    Me.mnModificar.Enabled = Not B
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String

    'Llamamos a al form
    Screen.MousePointer = vbHourglass
    Cad = ParaGrid(Text1(0), 10, "Código")
    Cad = Cad & ParaGrid(Text1(1), 30, "Descripción")
    Cad = Cad & ParaGrid(Text1(2), 60, "Fichero")
        
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "scryst"
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Tipos documento"
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
        frmB.vCargaFrame = False

        frmB.Show vbModal
        Set frmB = Nothing
        
        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(0)
        End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaSubRpt()
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    CadenaConsulta = "Select * from scryst2 where codcryst =" & Data1.Recordset!codcryst & " ORDER BY linea"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = miRsAux!linea
        IT.SubItems(1) = miRsAux!descriprp
        IT.SubItems(2) = miRsAux!nomcryst
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Set miRsAux = Nothing
    
    
End Sub


