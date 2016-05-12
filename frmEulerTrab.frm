VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEulerTrab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partes trabajo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   15420
   ClipControls    =   0   'False
   Icon            =   "frmEulerTrab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux3 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txtAux3 
      BackColor       =   &H80000018&
      Height          =   1515
      Index           =   4
      Left            =   11760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "frmEulerTrab.frx":000C
      Top             =   3960
      Width           =   3285
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmEulerTrab.frx":0012
      Left            =   11760
      List            =   "frmEulerTrab.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   8640
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Horas|N|N|0|24|sreloj|horas|#0.00||"
      Text            =   "horas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   11760
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Salida|H|S|||sreloj|Horafin|hh:mm:ss||"
      Text            =   "salida"
      Top             =   2040
      Width           =   1200
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11760
      TabIndex        =   4
      Tag             =   "Entrada|H|N|||sreloj|HoraInicio|hh:mm:ss||"
      Text            =   "entrada"
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   20
      ToolTipText     =   "Buscar artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux3 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Tag             =   "Cod. trabajo|T|N|||sreloj|codtipor|||"
      Text            =   "numal"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   16
      ToolTipText     =   "Buscar artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux3 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||sreloj|fecha|dd/mm/yyyy||"
      Text            =   "fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Height          =   315
      Index           =   6
      Left            =   13800
      MaxLength       =   13
      TabIndex        =   10
      Tag             =   "Tipo|N|N|||sreloj|numalbar|0000||"
      Text            =   "tipo"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12720
      TabIndex        =   7
      Top             =   5565
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13875
      TabIndex        =   8
      Top             =   5565
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13875
      TabIndex        =   9
      Top             =   5565
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   360
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Técnico|N|N|0|9999|sreloj|codtraba|0000||"
      Text            =   "tecn"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   9240
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEulerTrab.frx":006A
      Height          =   5025
      Left            =   120
      TabIndex        =   11
      Top             =   495
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8864
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgRef 
      Height          =   240
      Left            =   13080
      Picture         =   "frmEulerTrab.frx":007E
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "H. Fin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   11760
      TabIndex        =   24
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   11760
      TabIndex        =   23
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "H.Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   11760
      TabIndex        =   22
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmEulerTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
'Private WithEvents frmF As frmCal 'Calendario de Fechas
Private WithEvents frmT As frmAdmTrabajadores  'Form Mantenimiento Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Private NombreTabla As String
Private Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarModificar Then
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                 If InsertarModificar Then
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion Data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
            End If
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'cod. tecnico
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
            CadenaConsulta = ""
            frmT.Show vbModal
            Set frmT = Nothing
            If CadenaConsulta <> "" Then
                txtAux(0).Text = RecuperaValor(CadenaConsulta, 1)
                txtAux3(0).Text = RecuperaValor(CadenaConsulta, 2)
            End If
        Case 1
            'Orden trabajo
            Screen.MousePointer = vbHourglass
            CadenaConsulta = "Tipo|stipor|codtipor|T||13·Nombre|stipor|nomtipor|N||60·"
            Set frmB = New frmBuscaGrid
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Tipo trabajos"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri 'Conexion a BD Ariges
            frmB.vCampos = CadenaConsulta
            frmB.vTabla = "stipor"
            If cboTipo.ListIndex < 0 Or cboTipo.ListIndex > 2 Then
                CadenaConsulta = ""
            Else
                'CadenaConsulta = "codtipor like '" & RecuperaValor("ALR|ALE|ALO|", cboTipo.ListIndex + 1) & "'"
                CadenaConsulta = "codtipor like '" & RecuperaValor("R|E|O|", cboTipo.ListIndex + 1) & "%'"
            End If
            frmB.vSQL = CadenaConsulta
            CadenaConsulta = ""
            frmB.Show vbModal
            Set frmB = Nothing
            If CadenaConsulta <> "" Then
                txtAux(1).Text = RecuperaValor(CadenaConsulta, 1)
                txtAux3(1).Text = RecuperaValor(CadenaConsulta, 2)
            End If
            
        Case 100 'albaran
    
                    
            CadenaConsulta = "Tipo|scaalb|codtipom|T||8·Numero|scaalb|numalbar|N|00000|25·Cliente|scaalb|nomclien|T||60·"
            
            
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = CadenaConsulta
            frmB.vTabla = "scaalb"
            If cboTipo.ListIndex < 0 Then
                CadenaConsulta = " IN ('ALR','ALO','ALE')"
            Else
                CadenaConsulta = " = '" & RecuperaValor("ALR|ALE|ALO|", cboTipo.ListIndex + 1) & "'"
            End If
            CadenaConsulta = "codtipom " & CadenaConsulta
            frmB.vSQL = CadenaConsulta
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Albaranes " & cboTipo.Text
            frmB.vselElem = 2
            frmB.vConexionGrid = conAri 'Conexion a BD Ariges
            CadenaConsulta = ""
            frmB.Show vbModal
            Set frmB = Nothing
            If CadenaConsulta <> "" Then
                'ALR|ALE|ALO|
                NumRegElim = 0
                CadenaDesdeOtroForm = RecuperaValor(CadenaConsulta, 1)
                If CadenaDesdeOtroForm = "ALE" Then
                    NumRegElim = 1
                ElseIf CadenaDesdeOtroForm = "ALO" Then
                    NumRegElim = 2
                End If
                cboTipo.ListIndex = NumRegElim
                txtAux(6).Text = RecuperaValor(CadenaConsulta, 2)
                txtAux_LostFocus 6
                CadenaConsulta = ""
            End If
        
    End Select
    
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
                LimpiarCampos
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
    End Select
    Exit Sub
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data1.Recordset.EOF Then
        txtAux(0).Text = DBLet(Data1.Recordset!CodTraba, "T")
       ' txtAux2(1).Text = PonerNombreDeCod(txtAux(1), conAri, "straba", "nomtraba", "codtraba", "N")

        txtAux(1).Text = DBLet(Data1.Recordset!codtipor, "T")
       ' Me.txtAux2(2).Text = PonerNombreDeCod(txtAux(2), conAri, "sclien", "nomclien", "codclien", "N")
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
    
    If Modo <> 0 Then
        If Not Data1.Recordset.EOF Then
            'Caption = data4.Recordset!Id
            PonerDatosForaGrid False
        Else
           ' Caption = "EOF"
             PonerDatosForaGrid True
        End If
    End If
    
    Exit Sub
    
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "sreloj left join straba on sreloj.codtraba=straba.codtraba"
    Ordenacion = " ORDER BY sreloj.codtraba,fecha,horainicio "
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        'se le llama desde otro form
        BotonBuscar
    End If
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim tots As String
    
    On Error GoTo ECarga
    
    tots = MontaSQLCarga(enlaza)

    
    
    CargaGridGnral DataGrid1, Me.Data1, tots, False
    'S|txtAux(0)|T|Código|700|
    tots = "S|txtAux(0)|T|Trab|750|;S|cmdAux(0)|B||0|;S|txtAux3(0)|T|Nombre|2800|;S|txtAux(1)|T|Tipo|700|;S|cmdAux(1)|B||0|;"
    tots = tots & "S|txtAux3(1)|T|Trabajo|3600|;S|txtAux(2)|T|Fecha|1150|;S|txtAux(3)|T|Trab|850|;"
    
    'en formato hora
    tots = tots & "S|txtAux3(2)|T|Horas|850|;"
    
    'tots = tots & "S|txtAux(4)|T|Entrada|1100|;S|txtAux(5)|T|Salida|1100|;"
    tots = tots & "N|||||;N|||||;"
    tots = tots & "N|||||;N|||||;N|||||;"
    
    arregla tots, DataGrid1, Me
    
    
    'dtos alineados a la dcha
    
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight

    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0 Or Modo = 2) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   Exit Sub
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim b As Boolean

    On Error Resume Next
    
    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To 3
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
        If jj < 2 Then
            txtAux3(jj).Height = DataGrid1.RowHeight
            txtAux3(jj).Top = alto
            txtAux3(jj).visible = b
            
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = b
            Me.cmdAux(jj).Enabled = b
            
            
            
        End If
    Next jj
    
    For jj = 4 To 6
        BloquearTxt txtAux(jj), Not b
    Next jj
    BloquearCmb Me.cboTipo, Not b
    If Err.Number Then Err.Clear
End Sub




Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaConsulta = CadenaDevuelta
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Trabajadores
    'txtAux(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod traba
    CadenaConsulta = CadenaSeleccion
    'txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom traba
End Sub

Private Sub imgRef_Click()
    If Modo = 0 Or Modo = 2 Then Exit Sub
    If cboTipo.ListIndex < 0 Or cboTipo.ListIndex > 2 Then Exit Sub
    
    cmdAux_Click 100
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
Dim SQL As String

    If Me.Data1.Recordset.EOF Then Exit Sub
    
    SQL = " ID=" & Data1.Recordset!Id
    If BloqueaRegistro("sreloj", SQL) Then BotonModificar
    
'     If BLOQUEADesdeFormulario(Me) Then BotonModificar
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Aux As String
Dim K As Integer
    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 10 'Imprimir
        
            'AHORA
            frmListado2.Opcion = 46
            frmListado2.Show vbModal
            Exit Sub
'            'ANTES
'            If Data1.Recordset.EOF Then Exit Sub
'            CadenaConsulta = Data1.Recordset.Source
'            NumRegElim = InStr(1, CadenaConsulta, " WHERE ")
'            If NumRegElim > 0 Then
'                CadenaConsulta = Mid(CadenaConsulta, NumRegElim + 7)
'
'                NumRegElim = InStr(1, CadenaConsulta, " ORDER BY ")
'                If NumRegElim > 0 Then CadenaConsulta = Mid(CadenaConsulta, 1, NumRegElim - 1)
'            Else
'                CadenaConsulta = ""
'            End If
'
'            With frmImprimir
'
'                .OtrosParametros = "|elSQL=""" & CadenaConsulta & """|pEmpresa=""" & vEmpresa.nomempre & """|"
'                .NumeroParametros = 2
'
'                'CORCHETEAMOS
'                 NumRegElim = 1
'                 Do
'                       K = InStr(NumRegElim, CadenaConsulta, "eulertrabajos.")
'                       If K > 0 Then
'                            CadenaConsulta = Mid(CadenaConsulta, 1, K - 1) & "{" & Mid(CadenaConsulta, K) 'añado corchete
'                            K = InStr(K, CadenaConsulta, " ")
'                            'No puede ser 0
'                            CadenaConsulta = Mid(CadenaConsulta, 1, K - 1) & "}" & Mid(CadenaConsulta, K)
'                            NumRegElim = NumRegElim + K
'                        Else
'                            NumRegElim = 0
'                        End If
'                 Loop Until NumRegElim = 0
'
'
'                'Cambiaremos la FECHA y las horas a formato CRYSTAL
'                 NumRegElim = 1
'                 Do
'                       K = InStr(NumRegElim, CadenaConsulta, ".fecha")
'                       If K > 0 Then
'                            K = InStr(K, CadenaConsulta, "'") 'primera cilla de la fecha
'                            Aux = Mid(CadenaConsulta, K + 1, 10)
'                            Aux = Replace(Aux, "-", ",")
'                            Aux = "Date(" & Aux & ")"
'
'                            CadenaConsulta = Mid(CadenaConsulta, 1, K - 1) & Aux & Mid(CadenaConsulta, K + 12) 'cambo
'
'
'                            NumRegElim = NumRegElim + K
'                        Else
'                            NumRegElim = 0
'                        End If
'                 Loop Until NumRegElim = 0
'
'
'                .FormulaSeleccion = CadenaConsulta
'
'                .Titulo = "Partes trabajo"
'                .SoloImprimir = False
'                .EnvioEMail = False
'                .opcion = 2000 '2000 generico
'                .NombrePDF = ""
'                ' PongoNombrePDF Then .NombrePDF = cadPDFrpt
'                .NombreRpt = "eulParteTrabajo.rpt"
'                .ConSubInforme = False
'                .Show vbModal
'            End With
'
        Case 11  'Salir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim I As Byte
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
                      
    'Modo Buscar
    If Kmodo = 1 Then PonerFoco txtAux(0)
                      
    'Bloquear los campos de clave primaria al modificar
    'For I = 0 To 2
        BloquearTxt txtAux(7), True
    'Next I
                      
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    'PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


'Private Sub PonerLongCampos()
''Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
''para los campos que permitan introducir criterios más largos del tamaño del campo
'      PonerLongCamposGnral Me, Modo, 3
'End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
    
    On Error Resume Next

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    b = ((Modo >= 3))
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo opciones del menú.", Err.Description

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    cboTipo.ListIndex = -1
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    On Error GoTo ErrSQL
    
    

    SQL = "SELECT sreloj.codtraba,nomtraba ,stipor.codtipor,nomtipor,fecha,calculadas"
    
    'EN formato horas
    SQL = SQL & ",concat(floor(calculadas) ,':',right(concat('0',floor(round(100*(calculadas - floor(calculadas) ) * 0.6,2))),2)) horas"
    
    
    SQL = SQL & ",HoraInicio,HoraFin,"
    'SQL = SQL & " if(sreloj.codtipom='ALE','Exter',if(codtipom='ALR','Repar',if(codtipom='ALO','Orden','Prod'))) Txt,"
    SQL = SQL & " codtipom,numalbar,ID FROM "
    SQL = SQL & NombreTabla
    SQL = SQL & " left join stipor on stipor.codtipor=sreloj.codtipor"
        
        
        
        
        
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        Else
            If Modo = 3 Then
                If Data1.Recordset.RecordCount < 1 Then SQL = SQL & " WHERE fecha=" & DBSet(txtAux(2).Text, "F")
            End If
        End If
    Else
        SQL = SQL & " WHERE sreloj.codtraba = -1"
    End If
    
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
    Exit Function
    
ErrSQL:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cadena SQL", Err.Description
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    LimpiarCampos
    
    If Modo <> 1 Then
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 10)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid Me.DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    txtAux(2).Text = Format(Now, "dd/mm/yyyy")
    
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    On Error Resume Next
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc

    'poner valores grabados
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    txtAux3(0).Text = DBLet(DataGrid1.Columns(1).Value, "F")
    
    txtAux(1).Text = DBLet(DataGrid1.Columns(2).Value, "N")
    txtAux3(1).Text = DBLet(DataGrid1.Columns(3).Value, "F")
        
    txtAux(2).Text = DBLet(Data1.Recordset!Fecha, "F")
    
    txtAux(3).Text = Format(Data1.Recordset!calculadas, FormatoCantidad)
    

    DataGrid1.Enabled = False
    PonerFoco txtAux(0)
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Botón modificar.", Err.Description
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
        
    On Error GoTo FinEliminar
        
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
        
    SQL = "¿Seguro que desea eliminar la tarea?" & vbCrLf
    SQL = SQL & vbCrLf & "Fecha: " & Format(Data1.Recordset.Fields(6).Value, "dd/mm/yyyy")
    SQL = SQL & vbCrLf & "Técnico: " & Format(Data1.Recordset.Fields(0).Value, "0000") & " - " & Data1.Recordset.Fields(1).Value
    SQL = SQL & vbCrLf & "Referencia: " & DBLet(Data1.Recordset!codtipom, "T") & " - " & DBLet(Data1.Recordset.Fields(4).Value, "T")
            
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = "Delete from sreloj where ID=" & Data1.Recordset!Id
        
        conn.Execute SQL
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
        SituarDataPosicion Me.Data1, NumRegElim, SQL
    End If
    Exit Function
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Gastos Técnicos.", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String


    If cboTipo.ListIndex = -1 Then
        MsgBox "Seleccione tipo trabajo", vbExclamation
        PonerFocoCbo cboTipo
        Exit Function
    End If
    
    If Me.txtAux(3).Text <> "" Then
        If ImporteFormateado(Me.txtAux(3).Text) <> 0 And Me.txtAux(5).Text = "" Then
            MsgBox "Lleva horas trabajadas sin especficiar hora finalizacion", vbExclamation
            Exit Function
        End If
    End If
    If Me.txtAux(3).Text = "" Then txtAux(3).Text = "0"
    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    
    
    
    
    'Comprobar que existe un Albaran de venta para ese técnico(realizado por del alb)
    'para ese cliente y en esa fecha. Si no avisar
    If Me.cboTipo.ListIndex < 3 And Me.cboTipo.ListIndex >= 0 Then
        
        SQL = RecuperaValor("ALR|ALE|ALO|", cboTipo.ListIndex + 1)
        SQL = " numalbar=" & DBSet(txtAux(6).Text, "N") & " AND codtipom=" & DBSet(SQL, "T")
        SQL = "SELECT count(*) FROM scaalb WHERE " & SQL
        
        If Not (RegistrosAListar(SQL) > 0) Then
            SQL = "No existe el Albaran de fecha referenciado. ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then b = False
        End If
    End If
    DatosOk = b
End Function


Private Sub HacerBusqueda()
Dim cadB As String

    On Error Resume Next
    
   
    cadB = ObtenerBusqueda(Me, False)
     
    If cboTipo.ListIndex >= 0 Then
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & " codtipom = '" & RecuperaValor("ALR|ALE|ALO|PRO|", cboTipo.ListIndex + 1) & "'"
    End If
    
    
    If chkVistaPrevia = 1 Then
'        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
'        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        PonerCadenaBusqueda
        PonerFocoGrid Me.DataGrid1
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerCadenaBusqueda()
Dim Cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
      

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        Cad = "No hay ningún registro en la tabla "
         If EsBusqueda Then Cad = Cad & " para ese criterio de Búsqueda."
        MsgBox Cad, vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        CargaGrid True
'        DataGrid1.Refresh
'        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'cod tecnico
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux3(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "straba", "nomtraba", "codtraba")
                If txtAux3(Index).Text = "" Then PonerFoco txtAux(Index)
            Else
                txtAux3(Index).Text = ""
            End If
        
        
        Case 1
            If txtAux(Index).Text = "" Then
                txtAux3(Index).Text = ""
            Else
                txtAux(Index).Text = UCase(txtAux(Index).Text)
                txtAux3(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "stipor", "nomtipor", "codtipor", , "T")
                If Modo <> 1 And txtAux3(Index).Text = "" Then
                   ' MsgBox "No existe el tipo de trabajo", vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
        Case 6 'Cod. tipor
            If Modo = 1 Then Exit Sub
            If cboTipo.ListIndex < 0 Then
                txtAux(Index).Text = ""
                PonerFocoCbo cboTipo
                Exit Sub
            End If
            
            If txtAux(Index) = "" Then
                txtAux3(4).Text = ""
            Else
                If PonerFormatoEntero(txtAux(Index)) Then
                
                    If cboTipo.ListIndex = 3 Then
                        txtAux3(4).Text = "Produccion"
                    Else
                        CadenaConsulta = "codtipom = '" & RecuperaValor("ALR|ALE|ALO|", cboTipo.ListIndex + 1) & "' AND numalbar "
                        txtAux3(4).Text = PonerNombreDeCod(txtAux(Index), conAri, "scaalb", "nomclien", CadenaConsulta)
                        If txtAux3(4).Text = "" Then
                            txtAux(Index).Text = ""
                            PonerFoco txtAux(Index)
                        End If
                    End If
                Else
                    txtAux(Index).Text = ""
                   txtAux3(1).Text = ""
                End If
            End If
        Case 2 'kms
            PonerFormatoFecha txtAux(Index)
'            PonerFormatoEntero txtAux(Index)
           
        Case 3  'Sum horas
            PonerFormatoDecimal txtAux(Index), 4
            
        Case 4, 5 'horas
            txtAux(Index) = Replace(txtAux(Index), ".", ":")
            PonerFormatoHora txtAux(Index)
            DiferenciaHoras
            
    End Select
    
    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub

Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data1.Recordset Is Nothing) Then
            If Not Data1.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        txtAux(4).Text = ""
        txtAux(5).Text = ""
        txtAux(6).Text = ""
        txtAux3(4).Text = ""

        
    Else
        'EL
        txtAux(4).Text = PonerCampoAux(1)
        txtAux(5).Text = PonerCampoAux(2)
        txtAux(6).Text = PonerCampoAux(3)
        txtAux3(4).Text = ""
        PonerCampoAux 4
        
        
    End If
End Sub


Private Function PonerCampoAux(Cual As Integer) As String
Dim Cad As String
Dim C As String
    
    Cad = ""
    C = RecuperaValor("HoraInicio|HoraFin|numalbar|codtipom|", Cual)
    
    
    If IsNull(Data1.Recordset.Fields(C)) Then
        Cad = ""
    Else
        If Cual < 3 Then
            Cad = Format(Data1.Recordset.Fields(C), "hh:mm:ss")
        Else
            Cad = Data1.Recordset.Fields(C)
        End If
    End If
    PonerCampoAux = Cad
    
    
    If Cual = 4 Then
        If Cad = "" Then
            Me.cboTipo.ListIndex = -1
        Else
            'Ubicamos el combo
            'Vamos a intentar localizar el ALBARAN, FACTURA
                            
            If Mid(Cad, 1, 1) = "A" Then
                'ALbaran
                
                cboTipo.ListIndex = IIf(Cad = "ALE", 1, IIf(Cad = "ALO", 2, 0))
                
                
                
                C = "codtipom='" & Cad & "' AND numalbar"
                C = DevuelveDesdeBD(conAri, "concat(codclien,'|',nomclien,'|')", "scaalb", C, txtAux(6).Text)
                If C <> "" Then
                    C = RecuperaValor(C, 1) & " - " & RecuperaValor(C, 2)
                    C = "Albaran: " & vbCrLf & C
                Else
                    'SELECT concat(codclien,'|',nomclien,'|') from scafac,scafac1 where
                    C = "scafac.codtipom=scafac1.codtipom AND scafac.numfactu=Scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu"
                    C = "codtipoa='" & Cad & "' AND numalbar"
                    C = DevuelveDesdeBD(conAri, "concat(codclien,'|',nomclien,'|',scafac.numfactu,'|',scafac.fecfactu,'|',scafac.codtipom,'|')", "scafac,scafac1", C, txtAux(6).Text)
                
                    If C <> "" Then
                        C = RecuperaValor(C, 5) & RecuperaValor(C, 3) & " de " & RecuperaValor(C, 4) & vbCrLf & RecuperaValor(C, 1) & " - " & RecuperaValor(C, 2)
                        C = "Factura: " & C
                    Else
                        C = "NO encontrado"
                    End If
                End If
            Else
                'Prodccion
                cboTipo.ListIndex = 3
                C = "Produccion"
            End If
            
            Me.txtAux3(4).Text = C
        End If
    End If
End Function


Private Function InsertarModificar() As Boolean
Dim Cad As String
    
    
    If Modo <> 4 Then
        'sreloj(ID,Fecha,codtraba,HoraInicio,HoraFin,Calculadas,codtipom,numalbar,codtipor)
        
        Cad = SugerirCodigoSiguienteStr("sreloj", "id")
        
        If Data1.Recordset.EOF Then CadenaConsulta = Cad
        
        Cad = Cad & "," & DBSet(txtAux(2).Text, "F") & "," & DBSet(txtAux(0).Text, "N") & ","
        Cad = Cad & DBSet(txtAux(4).Text, "H") & ","
        If txtAux(5).Text = "" Then
            Cad = Cad & "NULL"
        Else
            Cad = Cad & DBSet(txtAux(5).Text, "H")
        End If
        Cad = Cad & "," & DBSet(txtAux(3).Text, "N") & ","
        
        Cad = Cad & "'" & RecuperaValor("ALR|ALE|ALO|PROD|", cboTipo.ListIndex + 1) & "',"
        Cad = Cad & DBSet(txtAux(6).Text, "N") & "," & DBSet(txtAux(1).Text, "T") & ")"
        
        Cad = "insert into sreloj(ID,Fecha,codtraba,HoraInicio,HoraFin,Calculadas,codtipom,numalbar,codtipor) VALUES (" & Cad
    Else
        Cad = "UPDATE sreloj SET fecha=" & DBSet(txtAux(2).Text, "F")
        Cad = Cad & ",codtraba=" & DBSet(txtAux(0).Text, "N") & ",codtipor=" & DBSet(txtAux(1).Text, "T")
        Cad = Cad & ",HoraInicio=" & DBSet(txtAux(4).Text, "H") & ",HoraFin="
        If txtAux(5).Text = "" Then
            Cad = Cad & "NULL"
        Else
            Cad = Cad & DBSet(txtAux(5).Text, "H")
        End If
        
        Cad = Cad & ",Calculadas=" & DBSet(txtAux(3).Text, "N") & ",numalbar=" & DBSet(txtAux(6).Text, "N")
        Cad = Cad & ",codtipom='" & RecuperaValor("ALR|ALE|ALO|PROD|", cboTipo.ListIndex + 1)
        Cad = Cad & "' WHERE ID =" & Data1.Recordset!Id
    End If
    
    InsertarModificar = ejecutar(Cad, False)
    If Modo = 3 And Data1.Recordset.EOF Then
        
    End If
End Function

Private Sub DiferenciaHoras()
Dim Minutos As Integer
    If Me.txtAux(4).Text = "" Or txtAux(5).Text = "" Then Exit Sub
    
    Minutos = DateDiff("n", CDate(txtAux(4).Text), CDate(txtAux(5).Text))
    If Minutos < 0 Then
        MsgBox "Diferencia de horas negativa", vbExclamation
        PonerFoco txtAux(5)
    Else
          
        kCampo = Minutos \ 60
        Minutos = Minutos - (kCampo * 60)
        txtAux(3).Text = kCampo & ","
        kCampo = 100 * (Round((Minutos / 60), 2))
        If kCampo = 100 Then kCampo = 99
        txtAux(3).Text = txtAux(3).Text & Format(kCampo, "00")
            
        
        
    End If
    
End Sub
