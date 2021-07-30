VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComProveAnticipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de albaranes"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16965
   ClipControls    =   0   'False
   Icon            =   "frmComProveAnticipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   18
      Top             =   90
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
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmComProveAnticipo.frx":000C
      Left            =   12930
      List            =   "frmComProveAnticipo.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "desconado|N|N|||sproveanticipo|descontado|||"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
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
      Index           =   5
      Left            =   10920
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Importe|N|N|0||sproveanticipo|importe|#,##0.00||"
      Text            =   "mporte"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
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
      Index           =   1
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   2
      Left            =   8760
      TabIndex        =   16
      ToolTipText     =   "Buscar zona"
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
      Index           =   4
      Left            =   7800
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Forpa|N|N|||sproveanticipo|codforpa|000||"
      Text            =   "zona"
      Top             =   4320
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   1
      Left            =   7080
      TabIndex        =   15
      ToolTipText     =   "Buscar envio"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
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
      Left            =   6600
      MaxLength       =   20
      TabIndex        =   2
      Tag             =   "Documento|T|N|||sproveanticipo|numdocum|||"
      Text            =   "documento"
      Top             =   3960
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
      Index           =   3
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha|F|N|||sproveanticipo|fechaant|dd/mm/yyyy||"
      Text            =   "fecha"
      Top             =   3720
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Id|N|N|0||sproveanticipo|codprove|0000||"
      Text            =   "codprove"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   9615
      Width           =   3495
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
         TabIndex        =   14
         Top             =   180
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   0
      Left            =   3360
      TabIndex        =   12
      ToolTipText     =   "Buscar fecha"
      Top             =   3720
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3165
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
      Tag             =   "Id|N|N|0||sproveanticipo|idanticipo|0000|S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
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
      Left            =   14520
      TabIndex        =   7
      Top             =   9690
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
      Left            =   15720
      TabIndex        =   8
      Top             =   9690
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
      Bindings        =   "frmComProveAnticipo.frx":0022
      Height          =   8520
      Left            =   150
      TabIndex        =   9
      Top             =   930
      Width           =   16620
      _ExtentX        =   29316
      _ExtentY        =   15028
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Anticipo proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6720
      TabIndex        =   20
      Top             =   120
      Width           =   10095
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
      TabIndex        =   10
      Top             =   7905
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
End
Attribute VB_Name = "frmComProveAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1


Private WithEvents frmFP As frmBasico2
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2
Attribute frmProv.VB_VarHelpID = -1





Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As String



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
                If InsertarDesdeForm(Me) Then
                    Screen.MousePointer = vbHourglass
                    'Primero  eseeperamos un momentito
                    Espera 0.25
                    InsertarAnticipoEnContabilidad CLng(txtAux(0).Text)
                    
                    CargaGrid True
                    
                    BotonAnyadir
                    
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                 
                    
                     '         DatosVto:   codmactaprov|numdcoum|fecdocum|
                     Indicador = DevuelveDesdeBD(conAri, "codmacta", "sprove", "codprove", Data1.Recordset!Codprove)
                     Indicador = Indicador & "|" & Data1.Recordset!numdocum & "|" & Data1.Recordset!fechaant & "|"
                     BorrarAnticipoEnContabilidad Indicador
                
                     Indicador = ""
                   
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 30
                     If SituarDataPosicion(Data1, NumReg, Indicador) Then
                        InsertarAnticipoEnContabilidad CLng(txtAux(0).Text)
                     Else
                        MsgBox "No se ha creado en contabilidad", vbExclamation
                     End If
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
    HaDevueltoDatos = ""
    Select Case Index
       
        
        Case 2
            'MandaBusquedaPrevia2 Index = 1
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, txtAux(4).Text
            Set frmFP = Nothing
            PonerFoco txtAux(4)
        
        
        Case 0
            Set frmProv = New frmBasico2
            AyudaProveedores frmProv, txtAux(1)
            Set frmProv = Nothing
            
            
        Case 1
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index + 2).Text <> "" Then
                If IsDate(txtAux(Index + 2).Text) Then frmF.Fecha = CDate(txtAux(Index + 2).Text)
            End If
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
            If HaDevueltoDatos <> "" Then
                txtAux(Index + 2).Text = HaDevueltoDatos
                PonerFoco txtAux(Index + 2)
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
            PonerFocoGrid DataGrid1
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub






Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub







Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()

    If Me.Caption = "" Then
        Me.Caption = "Anticipo proveedor"
        Data1.ConnectionString = conn
        CargaGrid True
        If Data1.Recordset.EOF Then
            PonerModo 0
            
        Else
            PonerModo 2
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Caption = ""
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

    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    Ordenacion = " ORDER BY idanticipo "
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    

        
    
    tots = "S|txtAux(0)|T|ID|950|;S|txtAux(1)|T|Cod.|1000|;S|cmdAux(0)|B||0|;"
    tots = tots & "S|txtAux2(0)|T|Proveedor|4080|;S|txtAux(2)|T|Documento|2900|;S|txtAux(3)|T|Fecha|1400|;"
    tots = tots & "S|cmdAux(1)|B||0|;"
    tots = tots & "S|txtAux(4)|T|F.P.|1000|;S|cmdAux(2)|B||0|;"
    tots = tots & "S|txtAux2(1)|T|Forma de pago|2250|;S|txtAux(5)|T|Importe|1400|;S|Combo1|C|Desc.|650|;"
    
    arregla tots, DataGrid1, Me, 350

    
    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim B As Boolean

    DeseleccionaGrid Me.DataGrid1
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        If jj < 2 Then
            txtAux2(jj).Height = Me.DataGrid1.RowHeight
            txtAux2(jj).Top = alto
            txtAux2(jj).visible = B
        End If
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = B
    Next jj

    Me.Combo1.visible = B
    Me.Combo1.Top = alto
    
    
    For jj = 0 To Me.cmdAux.Count - 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = B
    Next jj
End Sub




Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    HaDevueltoDatos = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    
    txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
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
        Case 1 'Nuevo
           BotonAnyadir
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            BotonEliminar
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 8 'Imprimir
          
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    
                      
    If Kmodo = 1 Then PonerFoco txtAux(0)
    
                                 
    BloquearTxt txtAux(0), (Modo <> 1)
    
    
    BloquearCmb Me.Combo1, Modo <> 1
    
    'Me.cmdAux(0).Enabled = (Modo <> 4)
                   
    '-----------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B

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
Dim B As Boolean

    
    'Insertar
    If vParamAplic.SerieAnticipoProveedor = "" Then
        B = False
        Toolbar1.Buttons(1).Enabled = B
        Me.mnNuevo.Enabled = B
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
        Toolbar1.Buttons(5).Enabled = B
        Me.mnBuscar.Enabled = B
        Toolbar1.Buttons(6).Enabled = B
        Me.mnVerTodos.Enabled = B
        
                
        
    Else
        Toolbar1.Buttons(1).Enabled = Modo = 2 Or Modo = 0
        Me.mnNuevo.Enabled = False
        
        
        
        'modificar eliminar
        B = (Modo = 2)
        If B Then B = Not Data1.Recordset.EOF
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
        
        
        
        B = (Modo >= 3 Or Modo = 1)
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'VerTodos
        Toolbar1.Buttons(6).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
        
        
        Toolbar1.Buttons(8).Enabled = False
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
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
    
    SQL = "select idanticipo,sproveanticipo.codprove,nomprove,numdocum,fechaant,sproveanticipo.codforpa,sforpa.nomforpa,"
    SQL = SQL & " importe,if(descontado=1,""SI"","""") descontado from sproveanticipo left join sprove on sproveanticipo.codprove=sprove.codprove"
    SQL = SQL & " left join sforpa on sproveanticipo.codforpa=sforpa.codforpa "
    SQL = SQL & " WHERE  true "
    SQL = SQL & ""
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND  false "
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
        anc = ObtenerAlto(Me.DataGrid1, 30)
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
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function PuedeElimModif() As Boolean

    PuedeElimModif = False
    If Modo <> 2 Then Exit Function
    If Data1.Recordset.EOF Then Exit Function
    'Lo principal para poder modificar.
    'NO puede estar descontado
    
    
    If Data1.Recordset!descontado <> "" Then
        MsgBox "Efecto ya descontado", vbExclamation
        Exit Function
    End If
    
            
    If Not EstadoAnticipoEnContabilidad(CLng(Data1.Recordset!idAnticipo)) Then Exit Function
        
    PuedeElimModif = True
        

End Function


Private Sub BotonEliminar()
Dim C As String
On Error GoTo Error2
    
    If Not PuedeElimModif() Then Exit Sub
    
    If MsgBox("¿Eliminar anticipo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    

    'Hay que eliminar
    NumRegElim = Data1.Recordset.AbsolutePosition
    
    
    
    '         DatosVto:   codmactaprov|numdcoum|fecdocum|
    C = DevuelveDesdeBD(conAri, "codmacta", "sprove", "codprove", Data1.Recordset!Codprove)
    C = C & "|" & Data1.Recordset!numdocum & "|" & Data1.Recordset!fechaant & "|"
    BorrarAnticipoEnContabilidad C
    
    
    
    
    conn.Execute "Delete from sproveanticipo where idanticipo=" & Data1.Recordset!idAnticipo
    CancelaADODC Me.Data1
    CargaGrid True
    CancelaADODC Me.Data1
    SituarDataPosicion Me.Data1, NumRegElim, ""

    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar anticipo", Err.Description
    
    
    
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single


    If Not PuedeElimModif Then Exit Sub

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    DeseleccionaGrid DataGrid1
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
 
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")
    txtAux(1).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    txtAux2(0).Text = DBLet(DataGrid1.Columns(2).Value, "T")
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "T")
    txtAux(3).Text = DBLet(Me.DataGrid1.Columns(4).Value, "T")
    txtAux(4).Text = DBLet(Me.DataGrid1.Columns(5).Text, "T")
    txtAux2(1).Text = DBLet(DataGrid1.Columns(6).Value, "T")
    
    
    
    txtAux(5).Text = DBLet(DataGrid1.Columns(7).Text, "F")
    
    
    
    If UCase(DBLet(DataGrid1.Columns(8).Value, "T")) = "SI" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    
    DataGrid1.Enabled = False
    PonerFoco txtAux(3)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    

    txtAux(0).Text = SugerirCodigoSiguienteStr("sproveanticipo", "idanticipo")
    FormateaCampo txtAux(0)


    NumRegElim = Val(ObtenerAlto(Me.DataGrid1, 30))
    LLamaLineas CSng(NumRegElim)
    Combo1.ListIndex = 1
    PonerFoco txtAux(1)
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean


    On Error GoTo ErrDatosOK
    
    
    DatosOk = False
    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
    'OBBLIGATORIO codmacta
    If DevuelveDesdeBD(conAri, "codmacta", "sprove", "codprove", txtAux(1).Text) = "" Then
        MsgBox "El proveedor no tiene cuenta contable asignada para contabilidad", vbExclamation
        B = False
   
    End If
    
    'NO tiene otro codigo igual
    
    If B And Modo = 3 Then
        
        '   numdocum   codprove fechaant
        If DevuelveDesdeBD(conAri, "idanticipo", "sproveanticipo", "year(fechaant)=" & Year(CDate(txtAux(3).Text)) & " AND numdocum=" & DBSet(txtAux(2).Text, "T") & " AND codprove", txtAux(1).Text) <> "" Then
            MsgBox "Ya existe el documento para el proveedor y  año ", vbExclamation
            B = False
        End If
    End If
    
    
    DatosOk = B
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function




Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
        
    If cadB <> "" Then CadenaBusqueda = " AND " & cadB
    CadenaConsulta = MontaSQLCarga(True)
    
    PonerCadenaBusqueda
 
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla para ese criterio de Búsqueda." & vbCrLf & CadenaConsulta, vbInformation
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

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
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
'   KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1: KEYBusqueda KeyAscii, 0 'prove
            Case 3: KEYBusqueda KeyAscii, 1 'fec
            Case 4: KEYBusqueda KeyAscii, 2 'forpa
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(ByRef KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1, 4
            Cad = ""
            If txtAux(Index).Text <> "" Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                Else
                    If Index = 4 Then
                        Cad = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtAux(Index).Text)
                    Else
                        Cad = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text)
                    End If
                    If Cad = "" Then MsgBox "No existe el valor en la BD: " & txtAux(Index).Text, vbExclamation
                End If
                If Cad = "" And txtAux(Index).Text <> "" Then
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
                      
            End If
            txtAux2(IIf(Index = 1, 0, 1)).Text = Cad
        Case 3 'fecha
            PonerFormatoFecha txtAux(Index)
            If txtAux(Index).Text <> "" Then
                Cad = ""
                If CDate(txtAux(Index)) < vEmpresa.FechaIni Then Cad = "Anterior inicio ejercicio"
                If Cad <> "" Then
                    MsgBox Cad, vbExclamation
                    txtAux(Index).Text = ""
                     PonerFoco txtAux(Index)
                End If
            End If
            
        Case 5
            If txtAux(Index).Text <> "" Then
            If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
            End If
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub

