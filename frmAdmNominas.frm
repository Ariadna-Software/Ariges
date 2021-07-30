VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmNominas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nominas y Gastos"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14190
   ClipControls    =   0   'False
   Icon            =   "frmAdmNominas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   20
      Top             =   90
      Width           =   2100
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   21
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Mes"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Norma 34"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recibos"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
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
      Index           =   6
      Left            =   7440
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "Liquidacion|N|S|0||||#,###,###,##0.00|N|"
      Text            =   "liq"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
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
      Left            =   6720
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Importe anticipo|N|S|0||snomin|impanti|#,###,###,##0.00|N|"
      Text            =   "Anticipo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CmbAux 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmAdmNominas.frx":000C
      Left            =   8280
      List            =   "frmAdmNominas.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Gastos pasados a Norma 34|N|N|||snomin|n34gast||N|"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CmbAux 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmAdmNominas.frx":0010
      Left            =   7440
      List            =   "frmAdmNominas.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Nomina pasada a Norma 34|N|N|||snomin|n34nomi||N|"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
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
      Index           =   1
      Left            =   960
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Mes nómina|N|N|1|99|snomin|mesnomi|00|S|"
      Text            =   "me"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
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
      Left            =   6480
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Importe gastos|N|N|0||snomin|impgasto|#,###,###,##0.00|N|"
      Text            =   "Imp gastos"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
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
      Left            =   5520
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Importe nómina|N|N|0||snomin|impnomi|#,###,###,##0.00|N|"
      Text            =   "Imp nomina"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   7035
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
         TabIndex        =   17
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Buscar trabajador"
      Top             =   3600
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
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3045
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
      Index           =   0
      Left            =   360
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Año nómina|N|N|0|9999|snomin|anynomi|0000|S|"
      Text            =   "any"
      Top             =   3600
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
      Left            =   11625
      TabIndex        =   9
      Top             =   7155
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
      Left            =   12915
      TabIndex        =   10
      Top             =   7155
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
      Left            =   12915
      TabIndex        =   11
      Top             =   7155
      Visible         =   0   'False
      Width           =   1065
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
      Index           =   2
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Trabajador|N|N|0|9999|snomin|codtraba|0000|S|"
      Text            =   "trab"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3000
      Top             =   5640
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
      Bindings        =   "frmAdmNominas.frx":0014
      Height          =   5985
      Left            =   240
      TabIndex        =   12
      Top             =   900
      Width           =   13775
      _ExtentX        =   24289
      _ExtentY        =   10557
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
      TabIndex        =   13
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
End
Attribute VB_Name = "frmAdmNominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
'Private WithEvents frmF As frmCal 'Calendario de Fechas
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mantenimiento Trabajadores
Attribute frmT.VB_VarHelpID = -1
'Private WithEvents frmC As frmFacClientes   'Form Mantenimiento Clientes


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






Private Sub CmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
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
                If InsertarDesdeForm(Me) Then
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 3) Then
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
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 1 'cod. tecnico
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, txtAux(2)
            Set frmT = Nothing
            PonerFoco txtAux(2)
            
'        Case 2 'cod. cliente
'            Set frmC = New frmFacClientes
'            frmC.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
'            frmC.Show vbModal
'            Set frmC = Nothing
    End Select
    
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            limpiar Me
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
    End Select
    Exit Sub
    
ECancelar:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0 Or Modo = 2) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

'    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(10).Image = 10 'Generar mes aut.
'        .Buttons(11).Image = 24 'Generar Norma 34
'        .Buttons(12).Image = 16 'Imprimir
'        .Buttons(13).Image = 40
'        .Buttons(15).Image = 15 'Salir
'    End With

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
        .Buttons(1).Image = 37 'Generar mes aut.
        .Buttons(2).Image = 36 'Generar Norma 34
        .Buttons(3).Image = 45 'Recibos
    End With



    
    limpiar Me   'Limpia los campos TextBox
    DataGrid1.ClearFields 'limpiar el Grid

    
    NombreTabla = "snomin" 'Tabla Nominas y Gastos
    Ordenacion = " ORDER BY anynomi desc,mesnomi desc,codtraba "
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        'se le llama desde otro form
        BotonBuscar
    End If
    
    CargarCombo_SiNo Me.CmbAux(0)
    CargarCombo_SiNo Me.CmbAux(1)
    
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim tots As String
    
    On Error GoTo ECarga
    
    tots = MontaSQLCarga(enlaza)
    
    CargaGridGnral DataGrid1, Me.Data1, tots, False
    
    tots = "S|txtAux(0)|T|Año|650|;S|txtAux(1)|T|Mes|550|;S|txtAux(2)|T|Codigo|900|;S|cmdAux(1)|B||0|;S|txtAux2(2)|T|Trabajador|3850|;"
    tots = tots & "S|txtAux(3)|T|Imp. Nómina|1550|;S|txtAux(4)|T|Imp. Gastos|1400|;"
    
    'Marzo2010
    'Campo calculado pendiente
    tots = tots & "S|txtAux(5)|T|Anticipo|1100|;S|txtAux(6)|T|Liquidación|1440|;"
    tots = tots & "N||||0|;S|CmbAux(0)|C|N34 N.|880|;N||||0|;S|CmbAux(1)|C|N34 G.|880|;"
    
    
    
    arregla tots, DataGrid1, Me, 350

    DataGrid1.ScrollBars = dbgAutomatic
   
   Exit Sub
   
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim B As Boolean

    On Error Resume Next
    
    DeseleccionaGrid Me.DataGrid1
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = B
        txtAux(jj).Text = ""
    Next jj
    
    txtAux2(2).Height = DataGrid1.RowHeight
    txtAux2(2).Top = alto
    txtAux2(2).visible = B
    BloquearTxt txtAux(6), True
    
    For jj = 0 To 1
        Me.CmbAux(jj).Height = DataGrid1.RowHeight
        Me.CmbAux(jj).Top = alto
        Me.CmbAux(jj).visible = B
        BloquearCmb Me.CmbAux(jj), (Modo <> 1)
    Next jj
    
    'boton de busqueda
    For jj = 1 To 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = B
        Me.cmdAux(jj).Enabled = B And (Modo <> 4)
    Next jj
    
    If Err.Number Then Err.Clear
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Trabajadores
    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod traba
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom traba
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
    
    SQL = " anynomi=" & DBSet(Me.Data1.Recordset!anynomi, "N") & " AND mesnomi=" & Data1.Recordset!mesnomi
    SQL = SQL & " AND codtraba=" & Data1.Recordset!CodTraba
    
    If BloqueaRegistro(NombreTabla, SQL) Then BotonModificar
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
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click 'Eliminar
        Case 5: mnBuscar_Click 'Busqueda
        Case 6: mnVerTodos_Click 'Ver Todos
        Case 8: AbrirListadoNomi (1) 'Imprimir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim i As Byte
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
                      
    'Modo Buscar
    If Kmodo = 1 Then PonerFoco txtAux(0)
                      
    'Bloquear los campos de clave primaria al modificar
    For i = 0 To 2
        BloquearTxt txtAux(i), (Modo = 4)
    Next i
                      
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
    
    On Error Resume Next

    B = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'VerTodos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    B = (Modo = 2 Or Modo = 0)
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'generar mes
    Toolbar5.Buttons(1).Enabled = B
    Toolbar5.Buttons(2).Enabled = B
    Toolbar1.Buttons(8).Enabled = Modo = 2 Or Modo = 0
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo opciones del menú.", Err.Description

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
    
    SQL = "SELECT anynomi,mesnomi," & NombreTabla & ".codtraba,nomtraba,impnomi,impgasto"
    
    'Marzo 2010
    SQL = SQL & ",impanti,impgasto-impanti"
    
    SQL = SQL & ",n34nomi,if(n34nomi=1,'Si','No') as dn34nomi,n34gast,if(n34gast=1,'Si','No') as dn34gast"
    SQL = SQL & " FROM " & NombreTabla
    SQL = SQL & " INNER JOIN straba ON " & NombreTabla & ".codtraba = straba.codtraba"
        
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        ElseIf Modo = 3 Then
            If Data1.Recordset.RecordCount < 1 Then SQL = SQL & " WHERE anynomi=" & txtAux(0).Text & " and mesnomi=" & txtAux(1).Text
        End If
    Else
        SQL = SQL & " WHERE " & NombreTabla & ".codtraba = -1"
    End If
    
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
    Exit Function
    
ErrSQL:
   MuestraError Err.Number, "Cadena SQL", Err.Description
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    limpiar Me
    Me.CmbAux(0).ListIndex = -1
    Me.CmbAux(1).ListIndex = -1
    
    If Modo <> 1 Then
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 25)
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
    limpiar Me
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid Me.DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    limpiar Me 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
   
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
    'valores por defecto
    txtAux(0).Text = Year(Now)
    txtAux(1).Text = Format(Month(Now), "00")
    PosicionarCombo Me.CmbAux(0), 0
    PosicionarCombo Me.CmbAux(1), 0
    
'    txtAux(9).Text = "1"
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    On Error Resume Next
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc

    'poner valores grabados
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    txtAux(1).Text = DBLet(DataGrid1.Columns(1).Value, "N")
    txtAux(2).Text = DBLet(DataGrid1.Columns(2).Value, "N")
    txtAux2(2).Text = DBLet(DataGrid1.Columns(3).Value, "T")
    txtAux(3).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    txtAux(4).Text = DBLet(DataGrid1.Columns(5).Value, "N")
    txtAux(5).Text = DBLet(DataGrid1.Columns(6).Value, "N")
    txtAux(6).Text = DBLet(DataGrid1.Columns(7).Value, "N")
    
    PosicionarCombo Me.CmbAux(0), DBLet(DataGrid1.Columns(8).Value, "N")
    PosicionarCombo Me.CmbAux(1), DBLet(DataGrid1.Columns(10).Value, "N")
    
    
    FormateaCampo txtAux(1)
    FormateaCampo txtAux(2)
    For i = 3 To 4
        FormateaCampo txtAux(i)
    Next i

    DataGrid1.Enabled = False
    PonerFoco txtAux(3)
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Botón modificar.", Err.Description
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
        
    On Error GoTo FinEliminar
        
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
        
    SQL = "¿Seguro que desea eliminar la linea de nómina y gastos?" & vbCrLf
    SQL = SQL & vbCrLf & "Año: " & Data1.Recordset.Fields(0).Value
    SQL = SQL & vbCrLf & "Mes: " & Format(Data1.Recordset.Fields(1).Value, "00")
    SQL = SQL & vbCrLf & "Trab.: " & Format(Data1.Recordset.Fields(2).Value, "0000") & " - " & Data1.Recordset.Fields(3).Value
            
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = "Delete from " & NombreTabla & " where anynomi=" & DBSet(Data1.Recordset!anynomi, "N")
        SQL = SQL & " AND mesnomi=" & Data1.Recordset!mesnomi & " AND codtraba=" & Data1.Recordset!CodTraba
        
        conn.Execute SQL
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
'        SituarDataPosicion Me.Data1, NumRegElim, SQL
        SituarDataTrasEliminar Me.Data1, NumRegElim, True
    End If
    Exit Function
        
FinEliminar:
     Screen.MousePointer = vbDefault
     MuestraError Err.Number, "Eliminar Gastos Técnicos.", Err.Description
End Function


Private Sub BotonGenerarMes()
'generar automaticamente las nominas de un mes partiendo de otro
    
    
    'abrir el form de listado de nominas mostrando el frame para
    'dedir de que mes/año quiere duplicar importe nomina
    AbrirListadoNomi (2)
    
End Sub


Private Sub BotonGenerarNorma34()
'generar automaticamente fichero Norma 34
    
    
    'abrir el form de listado de nominas mostrando el frame para
    'dedir de que mes/año quiere duplicar importe nomina
    AbrirListadoNomi (3)
    
    If Me.Data1.Recordset.EOF = False Then CargaGrid True
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
    DatosOk = B
End Function


Private Sub HacerBusqueda()
Dim cadB As String

    On Error Resume Next

    cadB = ObtenerBusqueda(Me, False)

    If cadB <> "" Then 'Se muestran en el mismo form
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
        Cad = "No hay ningún registro en la tabla " & NombreTabla
        If EsBusqueda Then Cad = Cad & " para ese criterio de Búsqueda."
        Screen.MousePointer = vbDefault
        MsgBox Cad, vbInformation
       
        PonerModo Modo
        Exit Sub
    Else
        CargaGrid True
        PonerModo 2
        DataGrid1_RowColChange 1, 1
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


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1: BotonGenerarMes 'generar automat. nuevas nominas
        Case 2: BotonGenerarNorma34 'generar Norma 34
        Case 3: AbrirListadoNomi (4) 'recibos
                
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
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 1 'trabajador
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
Dim Im As Currency


    On Error GoTo ErrFoco
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0, 1 'Año,mes
            PonerFormatoEntero txtAux(Index)
            
        Case 2 'Cod. trabajador
             If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "straba", "nomtraba", "codtraba")
                If txtAux2(Index).Text = "" Then PonerFoco txtAux(Index)
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 3, 4, 5 'Importes
             PonerFormatoDecimal txtAux(Index), 1
             
             
             If Index >= 4 Then
                'A partir de los gastos menos los anticipos
                Im = ImporteFormateado(txtAux(4).Text)
                Im = Im - ImporteFormateado(txtAux(5).Text)
                txtAux(6).Text = Format(Im, FormatoImporte)
                
            End If
             
    End Select
    Exit Sub
    
ErrFoco:
    MuestraError Err.Number, "", Err.Description
End Sub



Public Sub AbrirListadoNomi(numero As Integer)
'Abre el Form con los listados de nominas
    Screen.MousePointer = vbHourglass
    frmListadoNomi.parOpcion = numero
    frmListadoNomi.Show vbModal
    
    'Refresco el grid
    CadenaConsulta = Data1.RecordSource
    PonerCadenaBusqueda
    
    Screen.MousePointer = vbDefault
End Sub
