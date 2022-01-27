VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmArticu2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos (Busqueda)"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   Icon            =   "frmAlmArticu2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2265
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
               Object.ToolTipText     =   "Busqueda avanzada"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Lotes navidad"
            EndProperty
         EndProperty
      End
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
      Index           =   6
      Left            =   10320
      MaxLength       =   18
      TabIndex        =   6
      Tag             =   "refprove|T|N|||sartic|referprov|||"
      Text            =   "Dato2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
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
      Index           =   5
      Left            =   9360
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Stock|N|N|||salmac|canstock|#,###,###,##0.0000|N|"
      Text            =   "Dato2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   8760
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Stock|N|N|||sartic|preciove|#,##0.0000|N|"
      Text            =   "Dato2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
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
      Index           =   3
      Left            =   7560
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
      Text            =   "Dato2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
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
      Index           =   2
      Left            =   6120
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Cod. Asociacionl|T|N|||sartic|codtelem||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   1395
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
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código|T1|N|||sartic|codartic||S|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   2235
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
      Index           =   1
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||sartic|nomartic||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   3195
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmArticu2.frx":000C
      Height          =   5190
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9155
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      Left            =   12780
      TabIndex        =   8
      Top             =   6135
      Width           =   1320
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
      Left            =   11280
      TabIndex        =   7
      Top             =   6120
      Width           =   1200
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12840
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   3435
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
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3120
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Busqueda avanzad"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Articulos agrupados"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnLotes 
         Caption         =   "Lotes (articulos agrupados)"
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnBusqAvan 
         Caption         =   "&Busqueda avanza"
         HelpContextID   =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnOrdenacion 
      Caption         =   "Ordenacion"
      Begin VB.Menu mnOrdenadoPor 
         Caption         =   "Codigo"
         Index           =   0
      End
      Begin VB.Menu mnOrdenadoPor 
         Caption         =   "Nombre"
         Index           =   1
      End
      Begin VB.Menu mnExc 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnExc 
         Caption         =   "Excluir bloqueados - caducados"
         Checked         =   -1  'True
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmAlmArticu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DesdeTPV As Boolean
    'Si es desde TPV mostrara el precio con IVA (el % iva lo ha cargado en una tmp)
    'Dese cualquier otro sitio mostrara referprov

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
''''''Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos


Public Event DatoSeleccionado(CadenaSeleccion As String)

' ---- [06/11/2009] [LAURA] : añadir la cantidad de stock
'      añadimos el parametro almacen para mostrar el stock del almacen con el q trabajamos en TPV
Public parAlmacen As String
' ----


Private WithEvents frmA As frmAlmArticulosGr
Attribute frmA.VB_VarHelpID = -1

Private CadenaConsulta As String

Dim FormatoCod As String 'formato del campo de codigo
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


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(vModo As Byte)
Dim B As Boolean
Dim J As Integer

'    J = 0
'    If Modo <> 1 And vModo = 1 Then
'        J = 1
'    Else
'        If vModo = 1 And Modo <> 1 Then J = 1
'    End If
'    If J > 0 Then
'        For J = 0 To Me.DataGrid1.Columns.Count - 1
'            Me.DataGrid1.Columns(J).DividerStyle = IIf(vModo = 1, dbgNoDividers, dbgDarkGrayLine)
'        Next J
'    End If
    Modo = vModo
    B = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    Me.txtAux(0).visible = Not B
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    txtAux(3).visible = Not B
    txtAux(4).visible = Not B
    If Me.DesdeTPV Then
        txtAux(5).visible = Not B
    Else
        txtAux(6).visible = Not B
    End If
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
        
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
     
    'Si estamos en insertar o modificar
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'El PVP IVA NO SE PUEDE BUSCAR
    BloquearTxt txtAux(5), True
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                            'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean

    B = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ber Todos
    Toolbar1.Buttons(2).Enabled = B
    Me.mnVerTodos.Enabled = B
    

End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
'Dim anc As Single
'
'    'Situamos el grid al final
'    AnyadirLinea DataGrid1, adodc1
'
'    'Obtenemos la siguiente numero de factura
'    txtAux(0).Text = SugerirCodigoSiguienteStr("sactiv", "codactiv")
'    txtAux(0).Text = Format(txtAux(0).Text, FormatoCod)
'    txtAux(1).Text = ""
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 3
'
'    'Ponemos el foco
'    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid " false"  'para vaciar los datos del Grid
    limpiar Me
    LLamaLineas ObtenerAlto(DataGrid1, 45), 1
    If vParamAplic.SituaEnCodigoArticulo Then
        PonerFoco txtAux(0)
    Else
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub BotonVerTodos()
On Error Resume Next

    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla artic.", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
'Dim anc As Single
'Dim i As Integer
'
'    If adodc1.Recordset.EOF Then Exit Sub
'    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
'        i = DataGrid1.Bookmark - DataGrid1.FirstRow
'        DataGrid1.Scroll 0, i
'        DataGrid1.Refresh
'    End If
'
'    'Llamamos al form
'    txtAux(0).Text = DataGrid1.Columns(0).Text
'    txtAux(1).Text = DataGrid1.Columns(1).Text
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 4
'
'    'Como es modificar
''    PonerFoco txtAux(1)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(4).Top = alto
    txtAux(5).Top = alto
    txtAux(6).Top = alto
'    txtAux(0).Left = DataGrid1.Left + 340
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
'Dim SQL As String
'
'    On Error GoTo Error2
'
'    'Ciertas comprobaciones
'    If adodc1.Recordset.EOF Then Exit Sub
'
'    '### a mano
'    SQL = "¿Seguro que desea eliminar la Actividad?" & vbCrLf
'    SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
'    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
'    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
'        SQL = "Delete from sactiv where codactiv=" & adodc1.Recordset!codactiv
'        Conn.Execute SQL
'        CancelaADODC Me.adodc1
'        CargaGrid ""
'        CancelaADODC Me.adodc1
'        SituarDataPosicion Me.adodc1, NumRegElim, SQL
'    End If
'
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Actividad Cliente", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim cadB As String

    On Error Resume Next

    Select Case Modo
        Case 1 'HacerBusqueda
        
'            'Modificacion para que no tenga que poner *
'            For i = 1 To 2
'                If txtAux(i).Text <> "" Then
'                    'No lo ha puesto el. Se lo pongo YO
'                    If InStr(1, txtAux(i).Text, "*") = 0 Then txtAux(i).Text = "*" & txtAux(i).Text & "*"
'                End If
'            Next i
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
            
                'Febrero 2018
                '  Los articulos que esten BLOQUEADOS no salen en esta lista
                cadB = cadB & " AND sartic.codstatu<>2"
            
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
        
'        Case 3  'Hacemos insertar
'            If DatosOk Then
'                If InsertarDesdeForm(Me) Then
'                    CargaGrid
'                    BotonAnyadir
'                End If
'            End If
'
'        Case 4 'Modificar
'             If DatosOk And BLOQUEADesdeFormulario(Me) Then
'                 If ModificaDesdeFormulario(Me, 3) Then
'                      TerminaBloquear
'                      i = adodc1.Recordset.Fields(0)
'                      PonerModo 2
'                      CancelaADODC Me.adodc1
'                      CargaGrid
'                      adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
'                  End If
'                  DataGrid1.SetFocus
'            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1 'busqueda
            CargaGrid

'        Case 3 'Insertar
'            DataGrid1.AllowAddNew = False
'            'CargaGrid
'            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'
'        Case 4 'Modificar
'            'CargaGrid
'            TerminaBloquear
''            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
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
    If Modo = 1 Then
        If vParamAplic.SituaEnCodigoArticulo Then
            PonerFoco txtAux(0)
        Else
            PonerFoco txtAux(1)
        End If
    End If
End Sub


Private Sub Form_Load()
    ' ICONITOS DE LA BARRA
    Me.Icon = frmPpal.Icon
    
    If vParamAplic.HaciendoFrmulariosGrandes Then
    
      Toolbar1.visible = False
        Me.mnOpciones.visible = False
    
         With Me.Toolbar2
            .HotImageList = frmPpal.imgListComun_OM2
            .DisabledImageList = frmPpal.imgListComun_BN2
            .ImageList = frmPpal.ImgListComun2
         '   .Buttons(1).Image = 3    '3
         '   .Buttons(2).Image = 4    '4
            
            .Buttons(5).Image = 1
            .Buttons(6).Image = 2
            .Buttons(8).Image = 26
            
            .Buttons(10).visible = False
            mnLotes.visible = False
            If DesdeTPV Then
                If vParamAplic.NumeroInstalacion = 1 Then
                    .Buttons(10).Image = 21
                    .Buttons(10).visible = True
                    .Buttons(10).ToolTipText = "Lotes navidad"
                    mnLotes.visible = True
                End If
            End If
            
            
            
        End With
        DataGrid1.Top = 720
    Else
        Me.FrameBotonGnral.visible = False
        With Me.Toolbar1
            .ImageList = frmPpal.imgListComun
            .Buttons(1).Image = 1    'Botón Busqueda
            .Buttons(2).Image = 2    'Botón Recuperar Todos
            .Buttons(5).Image = 19    'Botón Añadir Nuevo Registro
    '        .Buttons(6).Image = 4    'Botón Modificar Registro
    '        .Buttons(7).Image = 5    'Botón Borrar Registro
            .Buttons(10).visible = False
            mnLotes.visible = False
            If DesdeTPV Then
                If vParamAplic.NumeroInstalacion = 1 Then
                    .Buttons(10).Image = 53  'Botón articulos agrupdados en TPV
                    .Buttons(10).visible = True
                    .Buttons(10).ToolTipText = "Lotes navidad"
                    mnLotes.visible = True
                End If
            End If
            .Buttons(17).Image = 15  'Botón Salir
        End With
        
        DataGrid1.Top = 540
    End If
    
    
    FormatoCod = CheckValueLeer(Me.Name)
    If FormatoCod <> "1" Then FormatoCod = "0"
    Me.mnOrdenadoPor(CInt(FormatoCod)).Checked = True
        
    'If vUsu.Nivel2 = 2 Then
    If False Then
        If vParamAplic.HaciendoFrmulariosGrandes Then
            Toolbar2.Buttons(8).Enabled = False
        Else
            Toolbar1.Buttons(5).Enabled = False
        End If
        mnBusqAvan.Enabled = False
    End If
    FormatoCod = FormatoCampo(txtAux(0))
    
    
    'SIEMPRE VIENEN EN MODO BUSQUEDA
    If DatosADevolverBusqueda = "" Then DatosADevolverBusqueda = "0"
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")

    
    'Novimebre 2010
    If parAlmacen = "" Then
        If vUsu.AlmacenPorDefecto2 <> "" Then
            parAlmacen = vUsu.AlmacenPorDefecto2
        Else
            parAlmacen = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
        End If
    End If
    
    'AHORA. Marzo 2010
    CadenaConsulta = "Select sartic.codartic,nomartic,codtelem,salmac.canstock,"
    CadenaConsulta = CadenaConsulta & "preciove,"
    If Me.DesdeTPV Then
        CadenaConsulta = CadenaConsulta & "if(isnull(porcen1),preciove,preciove*(1+(porcen1/100)))"
    Else
        If InstalacionEsEulerTaxco Then
            CadenaConsulta = CadenaConsulta & "if(ctrstock=0,'','Si') "
        Else
            CadenaConsulta = CadenaConsulta & "referprov"
        End If
    End If
    'CadenaConsulta = CadenaConsulta & " FROM  (sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic AND codalmac = " & parAlmacen & " ) "
    'If Me.DesdeTPV Then CadenaConsulta = CadenaConsulta & " LEFT OUTER JOIN tmpinformes ON sartic.codigiva=tmpinformes.codigo1 AND codusu = " & vUsu.codigo
    
    
    CadenaConsulta = CadenaConsulta & " FROM  sartic ,salmac "
    If Me.DesdeTPV Then CadenaConsulta = CadenaConsulta & ", tmpinformes  "
    CadenaConsulta = CadenaConsulta & " WHERE "
    If Me.DesdeTPV Then CadenaConsulta = CadenaConsulta & " codusu = " & vUsu.Codigo & " AND "
    CadenaConsulta = CadenaConsulta & " codalmac = " & parAlmacen & " AND sartic.codartic=salmac.codartic "
    If Me.DesdeTPV Then CadenaConsulta = CadenaConsulta & " AND sartic.codigiva=tmpinformes.codigo1 "
    
    
    
    BotonBuscar
'    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Modo = 0
    If Me.mnOrdenadoPor(1).Checked Then Modo = 1
    CheckValueGuardar Me.Name, Modo
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

'Private Sub Form_Unload(Cancel As Integer)
''    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
'End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub




Private Sub mnBusqAvan_Click()
Dim C As String
    C = CadenaConsulta
    CadenaConsulta = ""
    Set frmA = New frmAlmArticulosGr
    frmA.DatosADevolverBusqueda = "@1@" 'Poner en modo busqueda
    frmA.Show vbModal
    Set frmA = Nothing
    If CadenaConsulta <> "" Then
        
        
        RaiseEvent DatoSeleccionado(CadenaConsulta)
        Unload Me
        
    End If
    CadenaConsulta = C
End Sub


Private Sub mnExc_Click(Index As Integer)
      If Index = 1 Then mnExc(1).Checked = Not mnExc(1).Checked
End Sub

Private Sub mnLotes_Click()
        HacerToolbar 10
End Sub

Private Sub mnOrdenadoPor_Click(Index As Integer)
        mnOrdenadoPor(0).Checked = Index = 0
        mnOrdenadoPor(1).Checked = Index = 1
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolbar Button.Index
End Sub
Private Sub HacerToolbar(Indice As Integer)
    Select Case Indice
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5:
                mnBusqAvan_Click
                
        Case 10
            'De momento solo alzira y en TPV
            If DesdeTPV And vParamAplic.NumeroInstalacion = 1 Then
                CadenaDesdeOtroForm = ""
                frmMensajes.OpcionMensaje = 25
                frmMensajes.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    CadenaDesdeOtroForm = "##LOTESAGRUPADOS##" & CadenaDesdeOtroForm
                    RaiseEvent DatoSeleccionado(CadenaDesdeOtroForm)
                    Unload Me
                End If
            
            End If
        Case 17 'Salir
            
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional Sql As String)
Dim B As Boolean
Dim tots As String
Dim cadSel As String
Dim rc As Byte

    rc = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    B = DataGrid1.Enabled
    
    ' ---- [06/11/2009] [LAURA] : añadir la cantidad de stock
    cadSel = Sql
    
    
    
    Sql = CadenaConsulta
    
    If mnExc(1).Checked Then
        Sql = Sql & " AND sartic.codstatu<2 " 'EXLUCIMOS
    End If
    
'    If Trim(cadSel) <> "" Then Sql = Sql & " WHERE " & cadSel
    If Trim(cadSel) <> "" Then Sql = Sql & " AND " & cadSel
    If mnOrdenadoPor(1).Checked Then
        Sql = Sql & " ORDER BY sartic.nomartic"
    Else
        Sql = Sql & " ORDER BY sartic.codartic"
    End If

    CargaGridGnral DataGrid1, Me.Adodc1, Sql, False
    
    '### a mano
    tots = "S|txtAux(0)|T|Codigo|2100|;S|txtAux(1)|T|Descripcion|4640|;S|txtAux(2)|T|Cod. Asoc.|1500|;"
    ' ---- [06/11/2009] [LAURA] : añadir la cantidad de stock
    tots = tots & "S|txtAux(3)|T|Stock|1500|;"
    tots = tots & "S|txtAux(4)|T|PVP|1300|;"
    If Me.DesdeTPV Then
        tots = tots & "S|txtAux(5)|T|PVP IVA|1500|;"
    Else
        If InstalacionEsEulerTaxco Then
            tots = tots & "S|txtAux(6)|T|Ctr. stock|1000|;"
        Else
            tots = tots & "S|txtAux(6)|T|Ref. prov|2400|;"
        End If
    End If
    ' ----
    arregla tots, DataGrid1, Me, 330
    
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   Screen.MousePointer = rc
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Indice As Integer
    
    If Button.Index = 10 Then
        Indice = 10 'TPV
    Else
        If Button.Index = 8 Then
            Indice = 5
        ElseIf Button.Index = 6 Then
            Indice = 2
        Else
            Indice = 1
        End If
    End If
    HacerToolbar Indice
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 And KeyAscii = 13 Then
        If Me.txtAux(Index).Text <> "" Then
            PonerFocoBtn Me.cmdAceptar
            KeyAscii = 0
        End If
    End If
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    'If Index = 0 Then PonerFormatoEntero txtAux(Index)
End Sub


Private Function DatosOk() As Boolean
'Dim b As Boolean
'
'    b = CompForm(Me, 3)
'    If Not b Then Exit Function
'
'    If Modo = 3 Then 'Insertar
'        If ExisteCP(txtAux(0)) Then b = False
'    End If
'    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

