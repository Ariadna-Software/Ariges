VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmNumLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Números de Lote"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14625
   ClipControls    =   0   'False
   Icon            =   "frmAlmNumLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   12430
      TabIndex        =   19
      Top             =   135
      Width           =   1815
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmAlmNumLote.frx":000C
         Left            =   165
         List            =   "frmAlmNumLote.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   210
         Width           =   1530
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   17
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   18
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   15
      Top             =   135
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   16
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Comprobación"
            EndProperty
         EndProperty
      End
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
      Height          =   315
      Index           =   4
      Left            =   9570
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Vendida|N|N|||slotes|vendida|#,###,###,##0.00|N|"
      Text            =   "Vendida"
      Top             =   4335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   14
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
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
      Height          =   315
      Index           =   3
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Cantidad|N|N|||slotes|canentra|#,###,###,##0.00|N|"
      Text            =   "cantidad"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   315
      Index           =   2
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha entrada|F|N|||slotes|fecentra||S|"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   315
      Index           =   1
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Num. Lotes|T|N|||slotes|numlotes||S|"
      Text            =   "numlote"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   285
      TabIndex        =   12
      Top             =   9255
      Width           =   2535
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
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
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
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3645
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
      Height          =   315
      Index           =   0
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Artículo|T|N|||slotes|codartic||S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   12210
      TabIndex        =   5
      Top             =   9375
      Width           =   1065
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
      Left            =   13365
      TabIndex        =   6
      Top             =   9375
      Width           =   1065
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
      Left            =   13365
      TabIndex        =   7
      Top             =   9360
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
      Top             =   5280
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
      Bindings        =   "frmAlmNumLote.frx":0050
      Height          =   8145
      Left            =   225
      TabIndex        =   8
      Top             =   990
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   14367
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      TabIndex        =   9
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "Sin filtro"
         Index           =   0
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Disponible"
         Index           =   2
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimo mes"
         Index           =   3
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimos 6 meses"
         Index           =   4
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimo año"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmAlmNumLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean

Dim KK As Integer




Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    
    On Error GoTo Error1
    
    
    'FALTA#####
    ' LOG. si cambian cantidad vendida
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    EsBusqueda = True
                    CadenaBusqueda = " WHERE  slotes.codartic=" & DBSet(txtAux(0).Text, "T") & " AND numlotes=" & DBSet(txtAux(1).Text, "T")
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     'LOG
                     If vParamAplic.ManipuladorFitosanitarios2 Then
                        
                        Indicador = "Articulo: " & data1.Recordset!codArtic & " -- " & data1.Recordset!NomArtic
                        Indicador = Indicador & vbCrLf & "Anterior: " & data1.Recordset!vendida
                        Indicador = Indicador & "     ACtual: " & Me.txtAux(4).Text
                        
                        Set LOG = New cLOG
                        LOG.Insertar 35, vUsu, Indicador
                        Set LOG = Nothing
                     End If
                 
                     TerminaBloquear
                     NumReg = data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
             End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)

    Select Case Index
        Case 0 'Cod Articulo
            Set frmA = New frmBasico2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Poner Modo Busqueda
'            frmA.DesdeTPV = False
'            frmA.Show vbModal
            AyudaArticulos frmA, txtAux(0)
            Set frmA = Nothing
            PonerFoco txtAux(0)
        
        Case 1 'fecha entrada
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(2).Text <> "" Then frmF.Fecha = CDate(txtAux(2).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
            PonerFoco txtAux(2)
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
            If Not data1.Recordset.EOF Then
                data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = data1.Recordset.AbsolutePosition
            If Not data1.Recordset.EOF Then data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = data1.Recordset.Fields(0) & "|"
    cad = cad & data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not data1.Recordset.EOF Then
        lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

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


    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 20  'comprobacion
    End With

    CargaFiltros

    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "slotes" 'Tabla numeros de lotes
    Ordenacion = " ORDER BY codartic,fecentra,numlotes "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numlotes = -1" 'No recupera datos
    data1.ConnectionString = conn
    data1.RecordSource = CadenaConsulta
    data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid False
    
    KK = ByteValueLeer(Me.Name)
    PosicionarCombo Me.cboFiltro, KK
'    mnFiltro.Tag = KK
'    For kCampo = 0 To Me.mnFiltro1.Count - 1
'        If KK = kCampo Then
'            mnFiltro1(kCampo).Checked = True
'            Exit For
'        End If
'    Next
    
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.data1, SQL, False
    
    tots = "S|txtAux(0)|T|Artículo|2100|;S|cmdAux(0)|B||0|;S|txtAux2(0)|T|Nombre Artículo|4600|;S|txtAux(1)|T|Nº Lote|2000|;S|txtAux(2)|T|Fecha Entrada|1600|;S|cmdAux(1)|B||0|;S|txtAux(3)|T|Cantidad|1700|;S|txtAux(4)|T|Vendida|1700|;"
    arregla tots, DataGrid1, Me, 350


'    'dtos alineados a la dcha
    DataGrid1.Columns(3).Alignment = dbgCenter
'    DataGrid1.Columns(6).Alignment = dbgCenter

    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
    txtAux2(0).Height = Me.DataGrid1.RowHeight
    txtAux2(0).Top = alto
    txtAux2(0).visible = b
    
    
    For jj = 0 To Me.cmdAux.Count - 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = b
    Next jj
End Sub



Private Sub Form_Unload(Cancel As Integer)
    KK = 0
'    For kCampo = 0 To Me.mnFiltro1.Count - 1
'        If mnFiltro1(kCampo).Checked Then
'            KK = kCampo
'            Exit For
'        End If
'    Next
'    If mnFiltro.Tag <> KK Then ByteValueGuardar Me.Name, CByte(KK)
    KK = Me.cboFiltro.ListIndex
    ByteValueGuardar Me.Name, CByte(KK)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtAux(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnFiltro1_Click(Index As Integer)
    For KK = 0 To Me.mnFiltro1.Count - 1
        If Index <> 1 Then
            If KK = Index Then
                Me.mnFiltro1(KK).Checked = True
            Else
                Me.mnFiltro1(KK).Checked = False
            End If
        End If
    Next
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
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 1 'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            mnEliminar_Click
            
            If vUsu.Login <> "root" Then Exit Sub
            HacerComprobacion
        
        Case 8 'Imprimir
'            BotonImprimir
            
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
    
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
                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)
    Me.cmdAux(0).Enabled = (Modo <> 4)
                      
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    b = ((Modo >= 3))
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b

    '%=%= No hay impresion
    Toolbar1.Buttons(8).Enabled = False
    
    
    Toolbar5.Buttons(1).Enabled = (vUsu.Login = "root")
    FrameBotonGnral2.visible = (vUsu.Login = "root")
    FrameBotonGnral2.Enabled = (vUsu.Login = "root")


End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData data1, Index
    PonerCampos
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
    
    SQL = "SELECT " & NombreTabla & ".codartic, " & " Articulos.nomartic, numlotes, fecentra, canentra, vendida "
    SQL = SQL & " FROM " & NombreTabla & " LEFT OUTER JOIN sartic AS Articulos ON " & NombreTabla & ".codartic ="
    SQL = SQL & " Articulos.codartic"
    
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
        '%=%= añadida esta condicion
        If ValorFiltro <> "" Then
            If EsBusqueda And CadenaBusqueda <> "" Then
                SQL = SQL & " and " & ValorFiltro
            Else
                SQL = SQL & " where " & ValorFiltro
            End If
        End If
    Else
        SQL = SQL & " WHERE numlotes = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAltoNew(Me.DataGrid1, 10)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub

Private Function ValorFiltro() As String
    ValorFiltro = ""
    If Me.cboFiltro.ListIndex = 1 Then
        ValorFiltro = " canentra - vendida >0 "
'    ElseIf Me.mnFiltro1(3).Checked Then
'        MsgBox "Falt"
    End If
End Function

Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = ValorFiltro
    If CadenaConsulta <> "" Then CadenaConsulta = " WHERE " & CadenaConsulta
    
    
    CadenaConsulta = "Select * from " & NombreTabla & CadenaConsulta
    CadenaConsulta = CadenaConsulta & Ordenacion
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, data1
    
    'fecha de entrada
    txtAux(2).Text = Format(Now, "dd/mm/yyyy")
    
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
    'codartic
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")
    txtAux2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
    'numlote
    txtAux(1).Text = DBLet(Me.DataGrid1.Columns(2).Value, "T")
    'fecha entrada
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "F")
    FormateaCampo txtAux(2)
    
    'cantidad
    txtAux(3).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    FormateaCampo txtAux(3)
    
    'cantidad
    txtAux(4).Text = DBLet(DataGrid1.Columns(5).Value, "N")
    FormateaCampo txtAux(4)
    
    

    DataGrid1.Enabled = False
    PonerFoco txtAux(1)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar
    
    'Ciertas comprobaciones
    If data1.Recordset.EOF Then Exit Function
    
    If Not PuedeRealizarLaAccion Then Exit Function
    
    SQL = "¿Seguro que desea eliminar el Nº de lote?" & vbCrLf
    SQL = SQL & vbCrLf & "Nº Lote: " & data1.Recordset.Fields(2).Value
    SQL = SQL & vbCrLf & "Artículo: " & data1.Recordset.Fields(1).Value
            
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.data1.Recordset.AbsolutePosition
        SQL = "Delete from " & NombreTabla & " WHERE codartic=" & DBSet(data1.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(data1.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(data1.Recordset!fecentra, "F")
        
        conn.Execute SQL
        CancelaADODC Me.data1
        CargaGrid True
        CancelaADODC Me.data1
        SituarDataPosicion Me.data1, NumRegElim, SQL
    End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Nº de lote", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cart As CArticulo

    On Error GoTo ErrDatosOK

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    
    'comprobar que el articulo insertado tiene control de numero de serie
    If Modo = 3 Then
        Set cart = New CArticulo
        If cart.LeerDatos(txtAux(0).Text) Then
            If Not cart.TieneNumLote Then
                b = False
                MsgBox "El artículo no tiene control de nº de lote.", vbInformation
            End If
        End If
        Set cart = Nothing
    End If
    
    
    If b Then
        If ImporteFormateado(txtAux(4).Text) > ImporteFormateado(txtAux(3).Text) Then
            MsgBox "Cantidad vendida no puede ser mayor a cantidad entrada", vbExclamation
            PonerFoco txtAux(4)
            b = False
        End If
    End If
    
    If Modo = 4 Then
        If b Then b = PuedeRealizarLaAccion
    End If
    
    DatosOk = b
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
'    cad = cad & ParaGrid(txtAux(0), 20, "Código")
'    cad = cad & ParaGrid(txtAux(1), 80, "Descripción")
'    Tabla = NombreTabla
'    Titulo = "Tipos de Contrato"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco txtAux(kCampo)
''            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim Aux As String
    
    cadB = ObtenerBusqueda(Me, False)
    Aux = ValorFiltro
    If Aux <> "" Then Aux = " AND " & Aux
    
    
'    If chkVistaPrevia = 1 Then
'       MandaBusquedaPrevia cadB
'    Else
    If cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    data1.RecordSource = CadenaConsulta
    data1.Refresh
    If data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla de LOTES para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, data1
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Comprobacion
            If vUsu.Login <> "root" Then Exit Sub
            HacerComprobacion
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If Index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'codigo articulo
            Case 2: KEYBusqueda KeyAscii, 1 'codigo hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod. Articulo
            txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), conAri, "sartic", "nomartic")
            If txtAux2(0).Text = "" And txtAux(0).Text <> "" Then PonerFoco txtAux(0)
            
        Case 2 'fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 3, 4 'cantidad entrada
            PonerFormatoDecimal txtAux(Index), 1
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function PuedeRealizarLaAccion() As Boolean
Dim SQL As String

    PuedeRealizarLaAccion = True
    If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Function
            
            
    SQL = "OK"
    If Modo = 4 Then
        SQL = ""
        'Si cambia cantidad vendida o nº lote entonces comprobaremos
        If data1.Recordset!numlotes <> Me.txtAux(1).Text Then
            SQL = "OK"
        ElseIf data1.Recordset!vendida <> ImporteFormateado(Me.txtAux(4).Text) Then
            'SQL = "OK"
        End If
    Else
        SQL = "OK"
    End If
        
    PuedeRealizarLaAccion = False
    If SQL <> "" Then
        SQL = "numlote =" & DBSet(data1.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(data1.Recordset!fecentra, "F")
        SQL = SQL & " AND codartic = " & DBSet(data1.Recordset!codArtic, "T") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", SQL, "1")
        If Val(SQL) > 0 Then
            MsgBox "El lote ya ha sido vendido", vbExclamation
            Exit Function
        End If
        
        SQL = "numlote =" & DBSet(data1.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(data1.Recordset!fecentra, "F")
        SQL = SQL & " AND codartic = " & DBSet(data1.Recordset!codArtic, "T") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", SQL, "1")
        If Val(SQL) > 0 Then
            MsgBox "El lote esta siendo vendido", vbExclamation
            Exit Function
        End If
    End If
    PuedeRealizarLaAccion = True
    
End Function




Private Sub HacerComprobacion()
Dim cad As String
Dim cantidad As Currency
Dim ColArtic As Collection
Dim i As Integer
Dim EnElLote As Currency
Dim NF As Integer
Dim UpdateaSlotes As Boolean


    conn.Execute "DELETE FROM tmpstockfec WHERE codusu =" & vUsu.Codigo
    
    lblIndicador.Caption = "Ajuste devoluciones"
    lblIndicador.Refresh
    cad = "UPDATE slotes set vendida=canentra where canentra < 0"
    conn.Execute cad
    Espera 0.5
    
    
    
    lblIndicador.Caption = "Leyendo lotes"
    lblIndicador.Refresh
    cad = "insert into tmpstockfec(codusu,codartic,codalmac,stock)"
    cad = cad & " select " & vUsu.Codigo & ",codartic,1,sum(canentra-vendida) from slotes group by codartic"
    conn.Execute cad
     
     
    
      
     
    lblIndicador.Caption = "Leyendo salmac"
    lblIndicador.Refresh
    Espera 0.5
    
    Set miRsAux = New ADODB.Recordset
    cad = "Select codartic, sum(canstock) canti from salmac where codartic in (select codartic from tmpstockfec where codusu =" & vUsu.Codigo & ") GROUP BY codartic"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColArtic = New Collection
    NF = FreeFile
    Open App.Path & "\ajslot.txt" For Output As #NF
    While Not miRsAux.EOF
        
        cad = DevuelveDesdeBD(conAri, "stock", "tmpstockfec", "codusu = " & vUsu.Codigo & " AND codartic", miRsAux!codArtic, "T")
        cantidad = CCur(cad)
        
        If miRsAux!canti = cantidad Then
            cad = "DELETE FROM tmpstockfec WHERE codusu =" & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute cad
        Else
        
            cantidad = miRsAux!canti
        
            cad = miRsAux!codArtic & "|" & cantidad & "|"
            ColArtic.Add cad
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    For i = 1 To ColArtic.Count
        lblIndicador.Caption = i & " de " & ColArtic.Count
        lblIndicador.Refresh
        
        
        cad = RecuperaValor(ColArtic.Item(i), 2)
        cantidad = CCur(cad)   'canstock
        
        
        cad = RecuperaValor(ColArtic.Item(i), 1)
        Debug.Print cad
        
        
        If cantidad < 0 Then
            'STOCK NEGATIVO
            
            'Cantidad negativa
            Print #NF, RecuperaValor(ColArtic.Item(i), 1) & "::" & RecuperaValor(ColArtic.Item(i), 2) & "::" & cantidad & "      NEGATIVO"
            
        Else
            cad = "Select * from slotes WHERE  codartic=" & DBSet(cad, "T") & " AND canentra >0 order by fecentra desc"
            miRsAux.Open cad, conn, adOpenKeyset, adLockReadOnly, adCmdText
            cad = ""
            While Not miRsAux.EOF
               
               
                'If miRsAux!codArtic = "2020903032012" Then Sto p
               
                EnElLote = 0
                UpdateaSlotes = False
                If cantidad < 0 Then
                    'Significa que hay mas en lote s que en stock. Nos fiamos del stock
                    
                    
                    
                    
                Else
                    
                        If miRsAux!canentra > cantidad Then
                            'Perfecto. TOOOOdos va a este lote
                            EnElLote = miRsAux!canentra - cantidad
                            cantidad = 0
                            UpdateaSlotes = True
                        Else
                            
                            EnElLote = 0
                            cantidad = cantidad - miRsAux!canentra
                            UpdateaSlotes = True
                        End If
                                 
                End If
                
                
                If UpdateaSlotes Then
                    cad = "UPDATE slotes set vendida=" & DBSet(EnElLote, "N") & " WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
                    cad = cad & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!fecentra, "F")
                    conn.Execute cad
                End If
                miRsAux.MoveNext
                If cantidad = 0 Then
                    'HA IDO TODO BIEN
                    While Not miRsAux.EOF
                        cad = "UPDATE slotes set vendida=canentra WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
                        cad = cad & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!fecentra, "F")
                        conn.Execute cad
                        miRsAux.MoveNext
                    Wend
                End If
            Wend
            miRsAux.Close
            
            If cantidad <> 0 Then
                Print #NF, RecuperaValor(ColArtic.Item(i), 1) & "::" & RecuperaValor(ColArtic.Item(i), 2) & "::" & cantidad
            Else
                If cad = "" Then Print #NF, RecuperaValor(ColArtic.Item(i), 1) & "::" & RecuperaValor(ColArtic.Item(i), 2) & " 0STOCK"
            End If
        
        End If
    Next
    Close #NF
    
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
    Set ColArtic = Nothing
End Sub


Private Sub CargaFiltros()
Dim Aux As String
    
    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Disponible"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
'    cboFiltro.AddItem "Último mes"
'    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
'    cboFiltro.AddItem "Últimos 6 meses"
'    cboFiltro.ItemData(cboFiltro.NewIndex) = 3
'    cboFiltro.AddItem "Último año"
'    cboFiltro.ItemData(cboFiltro.NewIndex) = 4

End Sub
    

