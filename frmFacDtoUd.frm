VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacDtoUd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento tasas"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10965
   Icon            =   "frmFacDtoUd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacDtoUd.frx":000C
      Height          =   7440
      Left            =   135
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   945
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   13123
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   15
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   16
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
      Height          =   195
      Left            =   8775
      TabIndex        =   14
      Top             =   315
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   13
      Top             =   5400
      Width           =   225
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "Descripcion"
      Top             =   5400
      Width           =   1755
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   3960
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "Descuento|N|N|0|99|sdesca|dtolinea|0,00||"
      Text            =   "Descripcion"
      Top             =   5040
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3000
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Hasta|N|N|||sdesca|hastacan|0||"
      Text            =   "Descripcion"
      Top             =   5040
      Width           =   1035
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
      Left            =   9735
      TabIndex        =   5
      Top             =   8595
      Width           =   1065
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
      Left            =   8535
      TabIndex        =   4
      Top             =   8595
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Envase-granel|T|N|||sdesca|envagran|||"
      Text            =   "Codi"
      Top             =   5040
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Desde|N|S|||sdesca|desdecan|0||"
      Text            =   "Descripcion"
      Top             =   5040
      Width           =   1395
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
      Left            =   9720
      TabIndex        =   8
      Top             =   8595
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   135
      TabIndex        =   6
      Top             =   8550
      Width           =   1755
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   240
         TabIndex        =   7
         Top             =   195
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   480
      MaxLength       =   16
      TabIndex        =   11
      Tag             =   "id|N|N|||sdesca|id|||"
      Text            =   "Codi"
      Top             =   7080
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "El txt ID esta bajo"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
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
Attribute VB_Name = "frmFacDtoUd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)


Private WithEvents frmA As frmBasico2
Attribute frmA.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos


Dim Aux As String

Dim Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
        
    txtAux(0).visible = Not b
    txtAux2(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(2).visible = Not b
    txtAux(3).visible = Not b
    Me.Command1.visible = Not b
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If
    
    'Si estamos insertando o busqueda
   ' BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                            'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2)
    Toolbar1.Buttons(5).Enabled = b 'Buscar
    Me.mnBuscar.Enabled = b
    Toolbar1.Buttons(6).Enabled = b 'Todos
    Me.mnVerTodos.Enabled = b
    
    b = (Modo = 2) And Not DeConsulta
    'A�adir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = False 'b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
      
    anc = ObtenerAlto(DataGrid1, 10)
    
    'Obtenemos la siguiente numero de factura
    LimpiarCampos
    
    txtAux(4).Text = Format(SugerirCodigoSiguienteStr("sdesca", "id"), "000000")

    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "envagran= 'DABIZ'"
    LimpiarCampos
    LLamaLineas DataGrid1.Top + 240, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ning�n registro en la tabla descuentos", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
'        adodc1.Recordset.MoveFirst
'        PonerCampos
         DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()
Dim anc As Single
Dim i As Integer

    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(0).Text = DataGrid1.Columns(1).Text
    For i = 1 To 4
        txtAux(i).Text = DataGrid1.Columns(i + 1).Text
    Next i

    LLamaLineas anc, 4
   PonerFoco txtAux(0)
   Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Byte
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    txtAux2(0).Top = alto
    Command1.Top = alto
    For i = 0 To 3
        txtAux(i).Top = alto
    Next   '
    
    'Fijamos el ancho
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux2(0).Left = txtAux(0).Left + txtAux(0).Width + 30 '+ 90
    Me.Command1.Left = txtAux2(0).Left - Me.Command1.Width '-120
    txtAux(1).Left = txtAux2(0).Left + txtAux2(0).Width + 96 '75
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 80 '70
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 55 '45
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo Error2

    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    

    
    '### a mano
    SQL = "�Seguro que desea eliminar el descuento seleccionado? " & vbCrLf
    
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        'Hay que eliminar
        
       
        SQL = "Delete from sdesca where id = " & Adodc1.Recordset!ID
        conn.Execute SQL
        
        SQL = CStr(InStr(1, Adodc1.RecordSource, " WHERE "))
        If Val(SQL) > 0 Then
            SQL = Mid(Adodc1.RecordSource, Val(SQL) + 7)
            
            SQL = Mid(SQL, 1, InStr(1, SQL, " ORDER BY "))
            
        Else
            SQL = ""
        End If
        CancelaADODC Me.Adodc1
        CargaGrid SQL
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Tipo Unidad", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Long
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If

        Case 4  'MODIFICAR
            'If DatosOk And BLOQUEADesdeFormulario(Me) Then
            If DatosOk Then
                If Modificar() Then
                   TerminaBloquear
                   i = Adodc1.Recordset!ID
                   PonerModo 2
                   CancelaADODC Me.Adodc1
                   CargaGrid
                   Adodc1.Recordset.Find (" id =" & i)
                End If
                DataGrid1.SetFocus
            End If
        Case 1  'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
    Case 3 'Insertar
        Me.lblIndicador.Caption = ""
        DataGrid1.AllowAddNew = False
        'CargaGrid
        If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    Case 4 'Modificar
        TerminaBloquear
        Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    Case 1 'Busqueda
        CargaGrid
    End Select
    
    PonerModo 2
    DataGrid1.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String



        If Adodc1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
    
        cad = Adodc1.Recordset.Fields(0) & "|"
        cad = cad & Adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me

End Sub


Private Sub Command1_Click()
    If Modo = 2 Or Modo = 0 Then Exit Sub
        
    Aux = ""
    Set frmA = New frmBasico2
'    frmA.DatosADevolverBusqueda = "0|1|"
'    frmA.Show vbModal
    AyudaArticulos frmA, txtAux(0)
    Set frmA = Nothing
    If Aux <> "" Then
        txtAux(0).Text = RecuperaValor(Aux, 1)
        txtAux2(0).Text = RecuperaValor(Aux, 2)
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    
    'ICONOS de La toolbar.
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
    
    Screen.MousePointer = vbDefault
     
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "Select envagran,nomartic,desdecan,hastacan,dtolinea,id from sdesca inner join sartic on envagran=codartic"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Aux = CadenaSeleccion
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
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


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: BotonAnyadir
        Case 2: BotonModificar
        Case 3: BotonEliminar
        Case 5: BotonBuscar
        Case 6: BotonVerTodos
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim i As Byte
Dim b As Boolean
    
    b = DataGrid1.Enabled
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY envagran,desdecan,hastacan"
    
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False
    
    DataGrid1.RowHeight = 350
    
    i = 0 'Cod. Tipo Unidad
    DataGrid1.Columns(i).Caption = "Art�culo" 'RecuperaValor(txtAux(i).Tag, 1)
    DataGrid1.Columns(i).Width = 1600
            
    i = 1
    DataGrid1.Columns(i).Caption = "Descripci�n"
    DataGrid1.Columns(i).Width = 4150
        
    For i = 1 To 3
        DataGrid1.Columns(i + 1).Caption = RecuperaValor(txtAux(i).Tag, 1)
        If i = 3 Then
            DataGrid1.Columns(i + 1).Width = 1400 '900
            DataGrid1.Columns(i + 1).NumberFormat = "0.00"
        Else
            DataGrid1.Columns(i + 1).Width = 1450 '1200
        End If
        DataGrid1.Columns(i + 1).Alignment = dbgRight
    Next i
            
    DataGrid1.Columns(i + 1).visible = False
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux2(0).Width = DataGrid1.Columns(1).Width - 60
        For i = 2 To 4
            txtAux(i - 1).Width = DataGrid1.Columns(i).Width - 60
        Next i
        
        CadAncho = True
    End If
   
   'No permitir cambiar tama�o de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(2).Enabled Then
        Toolbar1.Buttons(2).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(3).Enabled = Not Adodc1.Recordset.EOF
        mnModificar.Enabled = Not Adodc1.Recordset.EOF
        mnEliminar.Enabled = Not Adodc1.Recordset.EOF
   End If
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   PonerOpcionesMenu
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim C As String
    'If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).BackColor = vbYellow Then txtAux(Index).BackColor = vbWhite
    If txtAux(Index).Text = "" Then Exit Sub
    
    If Modo < 3 Then Exit Sub
    
    Select Case Index
    Case 1, 2
        If Not PonerFormatoEntero(txtAux(Index)) Then PonerFoco txtAux(Index)
            
    Case 0
        'CERO. La longitud debe ser 4
'        txtAux(0).Text = UCase(txtAux(0).Text)
'        If Len(txtAux(Index).Text) <> 4 Then
'            MsgBox "Longitud debe ser 4", vbExclamation
'            Exit Sub
'        End If
        Aux = ""
        If txtAux(Index).Text <> "" Then
            C = "codartic"
            Aux = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux(Index).Text, "T", C)
            If Aux = "" Then
                MsgBox "No existe el articulo", vbExclamation
                PonerFoco txtAux(Index)
            Else
                txtAux(Index).Text = C
            End If
        End If
        txtAux2(Index).Text = Aux
    Case 3
        If Not PonerFormatoDecimal(txtAux(3), 1) Then PonerFoco txtAux(3)
        
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False

    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de tipo unidad en la tabla
    
'    If Len(txtAux(0).Text) <> 4 Then
'        MsgBox "Longitud debe ser 4", vbExclamation
'        Exit Function
'    End If
    
    '    If ExisteCP(txtAux(0)) Then b = False
    
    
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function Modificar() As Boolean
Dim C As String
Dim i As Byte
Dim J As Integer
    On Error GoTo EModificar
    Modificar = False
    
    C = ""
    For i = 0 To 3
        J = IIf(i = 0, 0, i + 1)
        C = C & ", " & RecuperaValor(txtAux(i).Tag, 7) & " = "
        If i = 0 Then
            C = C & "'" & txtAux(0) & "'"
        Else
            C = C & TransformaComasPuntos(txtAux(i).Text)
        End If
    Next
    C = Mid(C, 2) 'quito la 1� coma
    C = "UPDATE sdesca set " & C
    C = C & " WHERE id = " & Adodc1.Recordset!ID
    conn.Execute C
    Modificar = True
    Exit Function
EModificar:
    MuestraError Err.Number, Err.Description
End Function
