VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacFormasPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formas de Pago"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmFacFormasPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   240
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Texto auxiliar|T|S|||sforpa|obsIBAN|||"
      Text            =   "Text1"
      Top             =   2280
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "% Gastos Financieros|N|S|0|99.90|sforpa|porgasfi|#0.00|N|"
      Text            =   "Text1"
      Top             =   1560
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Caption         =   "Forma de Pago por Adelantado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   7695
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   1965
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   360
         Width           =   2685
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   6120
         MaxLength       =   5
         TabIndex        =   11
         Tag             =   "% Adelantado|N|S|0|99.90|sforpa|poradela|#0.00|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "Forma de Pago por adelantado|N|S|0|999|sforpa|forpapor|000|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   645
      End
      Begin VB.Image imgFPago 
         Height          =   240
         Index           =   1
         Left            =   980
         Picture         =   "frmFacFormasPago.frx":000C
         ToolTipText     =   "Buscar forma de pago"
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "% Adelantado"
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   31
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pago"
         Height          =   315
         Index           =   7
         Left            =   270
         TabIndex        =   30
         Top             =   370
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Código Forma de Pago|N|N|0|999|sforpa|codforpa|000|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2800
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Resto Vencimientos|T|S|||sforpa|restoven||N|"
      Text            =   "Text1"
      Top             =   1560
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   1580
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Primer Vencimiento|N|N|0||sforpa|primerve|0|N|"
      Text            =   "Text1"
      Top             =   1560
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   240
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "Nº Vencimientos|N|N|1|99999|sforpa|numerove||N|"
      Text            =   "Text1"
      Top             =   1560
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmFacFormasPago.frx":010E
      Left            =   4200
      List            =   "frmFacFormasPago.frx":0110
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Tipo de Pago|N|N|||sforpa|tipforpa||N|"
      Top             =   840
      Width           =   2700
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6780
      TabIndex        =   13
      Top             =   5235
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1000
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre de la Forma de Pago|T|N|||sforpa|nomforpa|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   17
      Top             =   5115
      Width           =   3135
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   210
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6780
      TabIndex        =   14
      Top             =   5235
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   5235
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   150
      Top             =   5355
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin VB.Frame Frame2 
      Caption         =   "Forma de Pago Alternativa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   7695
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Forma de Pago alternativa|N|S|0|999|sforpa|forpaalt|000|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Importe Mínimo|N|S|0||sforpa|impormin|#,###,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   1965
         MaxLength       =   30
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   360
         Width           =   2685
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pago"
         Height          =   315
         Index           =   9
         Left            =   270
         TabIndex        =   32
         Top             =   370
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Mínimo"
         Height          =   315
         Index           =   6
         Left            =   4920
         TabIndex        =   28
         Top             =   375
         Width           =   1095
      End
      Begin VB.Image imgFPago 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmFacFormasPago.frx":0112
         Tag             =   "-1"
         ToolTipText     =   "Buscar forma de pago"
         Top             =   390
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5280
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Texto auxiliar"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   34
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "% Gastos Financieros"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   33
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Resto Vtos."
      Height          =   255
      Index           =   4
      Left            =   2805
      TabIndex        =   26
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Primer Vto."
      Height          =   255
      Index           =   3
      Left            =   1575
      TabIndex        =   25
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Vencimientos"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Pago"
      Height          =   195
      Index           =   10
      Left            =   4200
      TabIndex        =   21
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1005
      TabIndex        =   20
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   19
      Top             =   600
      Width           =   495
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
Attribute VB_Name = "frmFacFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    InsertarEnTesoreria
                    PonerModo 0
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    ModificarENTesoeria
                    TerminaBloquear
                    Cad = "(codforpa=" & Text1(0).Text & ")"
                    If SituarData(Data1, Cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                        PonerFoco Text1(0)
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("sforpa", "codforpa")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        '### A mano
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
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    '### A mano
    'Bloquear importe minimo y %adelantado si las formas de pago estan vacias
    If Text1(6).Text = "" Then BloquearTxt Text1(8), True
    If Text1(7).Text = "" Then BloquearTxt Text1(9), True
    'Bloquear Restos Vencimientos si nº vencimientos=1
    If Val(Text1(2).Text) = 1 Then BloquearTxt Text1(4), True
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not PuedeModificarFPenContab Then Exit Sub
    
    '### a mano
    Cad = "¿Seguro que desea eliminar la Forma de Pago?" & vbCrLf
    Cad = Cad & vbCrLf & "Cod. Forma Pago: " & Format(Data1.Recordset.Fields(0), "000")
    Cad = Cad & vbCrLf & "Desc. Forma Pago: " & Data1.Recordset.Fields(1)
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        Screen.MousePointer = vbHourglass

        Cad = "En ariges"
        Data1.Recordset.Delete
        
        
        'Para eliminar en tesoreria
        Cad = "DELETE FROM sforpa WHERE codforpa = " & Text1(0).Text
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
        
        
        'Borro en tesoreria
        
        ConnConta.Execute Cad
        
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Forma de Pago" & vbCrLf & Cad, Err.Description
    End If
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


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 13
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 15  'Salir
        .Buttons(13).Image = 6  'Primero
        .Buttons(14).Image = 7  'Anterior
        .Buttons(15).Image = 8  'Siguiente
        .Buttons(16).Image = 9  'Último
    End With
    
    LimpiarCampos
    
    'Si hay algun combo los cargamos
    CargarComboTipoPago
    
    '## A mano
    NombreTabla = "sforpa"
    Ordenacion = " ORDER BY codforpa"
           
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '## A mano
    Data1.RecordSource = "Select * from " & NombreTabla & " where codforpa=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
End Sub


Private Sub CargarComboTipoPago()
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Contado, 1-Cheque, 2-Pagaré, 3-Transferencia, 4-Efecto
Dim RS As ADODB.Recordset
Dim SQL As String

    Combo1.Clear
        
    On Error GoTo ECargar

    SQL = "SELECT tipforpa, destippa from stippa"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Combo1.AddItem RS!destippa
        Combo1.ItemData(Combo1.NewIndex) = RS!tipforpa
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando combo tipos de pago.", Err.Description
    

End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim indice As Integer

    If CadenaDevuelta <> "" Then
        If Val(imgFPago(0).Tag) >= 0 Then 'Llama desde Prismaticos
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgFPago(0).Tag)
            Text1(indice + 6).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(indice + 6).Text = RecuperaValor(CadenaDevuelta, 2)

            If Modo = 3 Then
                 Text1(indice + 8).Locked = False
                 Text1(indice + 8).BackColor = vbWhite
            End If
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub
    
    
Private Sub imgFPago_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
      
    'En el Tag almacenamos el indice de la imagen de
    'Busqueda que va a llamar a frmBuscaGrid para busqueda
    imgFPago(0).Tag = Index
    MandaBusquedaPrevia ""
    imgFPago(0).Tag = -1
    PonerFoco Text1(Index + 6)
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
   
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod Forma de Pago
           If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
           End If
            
        Case 2 'Numero Vencimientos
            PonerFormatoEntero Text1(Index)
            If Val(Text1(Index).Text) = 1 Then
                Text1(4).Text = ""
                BloquearTxt Text1(4), True
            Else
                BloquearTxt Text1(4), False
            End If
                
        Case 3, 4 'nº vencimientos
            PonerFormatoEntero Text1(Index)
        
        Case 5, 9 '5: %Gastos Financieros, 9: %Adelantado
             'Formato tipo 4: Decimal(4,2)
             PonerFormatoDecimal Text1(Index), 4

        Case 8       '8:Importe Mínimo
            'Formato tipo 1: Decimal(12,2)
             PonerFormatoDecimal Text1(Index), 1
        
        Case 6, 7 ' 6: Forma de Pago Alternativa
                  ' 7: Forma de Pago por Adelantado
             If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa", "N")
                If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                BloquearTxt Text1(Index + 2), False
             Else
                 Text2(Index).Text = ""
                 Text1(Index + 2).Text = "" 'Importe Mínimo
                 'Modo 1: Busqueda
                 If Modo <> 1 Then BloquearTxt Text1(Index + 2), True
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim Cad As String

        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 30, "Código")
        Cad = Cad & ParaGrid(Text1(1), 70, "Denominacion")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = cadB
            
            HaDevueltoDatos = False
            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Formas de Pago"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1 'Conexión a BD: Ariges
'            If imgFPago(0).Tag = -1 Then
'                frmB.vBuscaPrevia = chkVistaPrevia
'            Else
'                frmB.vBuscaPrevia = True
'            End If
            frmB.vCargaFrame = False
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
''            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                PonerFoco Text1(kCampo)
'                PonerModo Modo
'            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
'         MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
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

    On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Me.Data1
    Text2(6).Text = PonerNombreDeCod(Text1(6), 1, "sforpa", "nomforpa", "codforpa")
    Text2(7).Text = PonerNombreDeCod(Text1(7), 1, "sforpa", "nomforpa", "codforpa")
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    
    '----------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    If B Or Modo = 1 Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    BloquearText1 Me, Modo
    
    'Formas de Pago
    For i = 0 To Text2.Count - 1
        BloquearTxt Text2(i), True
    Next i
    
    Combo1.Enabled = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    
    B = (Modo = 3) 'Insertar
    'Campos Importe Mínimo y % Adelantado
    If B Then
        For i = 8 To 9
            BloquearTxt Text1(i), True
        Next i
    End If

     chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean

    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    B = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnNuevo.Enabled = Not B
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    
    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
     
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If
     
    If Not B Then Exit Function
    
    'Comprobar que si nº vencimientos es 1, el campo resto vencimientos no tiene valor
    If Val(Text1(2).Text) = 1 Then
        If Not EsVacio(Text1(4)) Then
            MsgBox "El campo Resto Vencimientos no puede tener valor si NºVtos=1", vbInformation
            B = False
        End If
    End If
    If Not B Then Exit Function
     
    'Comprobar el campo resto vencimientos
    If Not EsVacio(Text1(2)) And Not EsVacio(Text1(3)) Then
        If Val(Text1(2).Text) > 1 And EsVacio(Text1(4)) Then
            MsgBox "El Campo Resto Vencimientos debe tener valor", vbInformation
            PonerFoco Text1(4)
            B = False
        End If
    End If
    If Not B Then Exit Function
    
    'Comprobar el importe Mínimo
    'Requerido si se selecciona una forma de pago alternativa
    If Not EsVacio(Text1(6)) And EsVacio(Text1(8)) Then
       MsgBox "El campo Importe Mínimo debe tener valor", vbInformation
       PonerFoco Text1(8)
       B = False
    End If
    'Verificar que el campo Importe Minimo no tiene valor si la forma de pago es vacio
    If EsVacio(Text1(6)) And Not EsVacio(Text1(8)) Then
        MsgBox "El campo Importe Mínimo no puede tener valor si no selecciona Forma de Pago.", vbInformation
        B = False
    End If
    If Not B Then Exit Function
    
    
    'Porcentaje Adelantado
    'Requerido si se selecciona una forma de pago por adelantado
    If Not EsVacio(Text1(7)) And EsVacio(Text1(9)) Then
        MsgBox "El campo % Adelantado debe tener valor", vbInformation
        PonerFoco Text1(9)
        B = False
    End If
    'Verificar que el campo %adelantado no tiene valor si la forma de pago es vacio
    If EsVacio(Text1(7)) And Not EsVacio(Text1(9)) Then
        MsgBox "El campo %Adelantado no puede tener valor si no selecciona Forma de Pago.", vbInformation
        B = False
    End If
    If Not B Then Exit Function
        
        
    'Marzo 2011
    '----------------------------------
    'Comprobaremos que son correctas las formas de pago adelantao y alternativa, y que no son la misma
    'que la principal
    If B Then
        'Alternativa
        If Text1(6).Text <> "" Then
            If Text1(6).Text = Text1(0).Text Then
                MsgBox "Mismo código de forma de pago y la forma de pago alternativa", vbExclamation
                B = False
            End If
        End If
        
        
        'Adelantado
        If Text1(7).Text <> "" Then
            If Text1(7).Text = Text1(0).Text Then
                MsgBox "Mismo código de forma de pago y la forma de pago adelantado", vbExclamation
                B = False
            End If
        End If
        
        'AHora veamos si de la forma de pago adelantado NO es recib
        
        
        
        'Salimos si no esta bien
        If Not B Then Exit Function
    End If
        
            
        
        
    'Comprobaciones de TESORERIA
    If Modo = 4 Then
        'Estoy modificando
        If Not PuedeModificarFPenContab Then Exit Function
    End If
    
    DatosOk = B
End Function


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 10  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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


Private Function PuedeModificarFPenContab() As Boolean
Dim Cad As String
    PuedeModificarFPenContab = False
    Set miRsAux = New ADODB.Recordset

    
    'Si modo=2 esta eliminando.
    'Vere si la forma de pago es forma de pago alternativa
    If Modo = 2 Then
        NumRegElim = 0
        Cad = DevuelveDesdeBD(conAri, "count(*)", "sforpa", "forpaalt", CStr(Val(Text1(0).Text)))
        If Cad = "" Then Cad = "0"
        If Val(Cad) > 0 Then NumRegElim = 1
        Cad = DevuelveDesdeBD(conAri, "count(*)", "sforpa", "forpapor", Text1(0).Text)
        If Cad = "" Then Cad = "0"
        If Val(Cad) > 0 Then NumRegElim = 1
        
        If NumRegElim > 0 Then
            MsgBox "La forma de pago es forma de pago alternativa o por adelantado de otra", vbExclamation
            Exit Function
        End If
        
        
        If vParamAplic.AguasPotables Then
            Cad = DevuelveDesdeBD(conAri, "Count(*)", "aguacontadores", "codforpa", Text1(0).Text)
            If Cad = "" Then Cad = "0"
            If Val(Cad) > 0 Then
                MsgBox "Tiene contadores de agua asociados a esta forma de pago", vbExclamation
                Exit Function
            End If
        End If
        
        
    End If

    NumRegElim = 0
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select count(*) from cobros where codforpa=" & Text1(0).Text
    Else
        Cad = "Select count(*) from scobro where codforpa=" & Text1(0).Text
    End If
    
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select count(*) from pagos where codforpa=" & Text1(0).Text
    Else
        Cad = "Select count(*) from spagop where codforpa=" & Text1(0).Text
    End If
    
    
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        If Modo = 4 Then
            If MsgBox("Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago. ¿Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        Else
            'NO DEJO CONTINUAR
            MsgBox "Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago", vbExclamation
            Exit Function
        End If
            
            
    End If
    'Si llega aqui puede seguir
    PuedeModificarFPenContab = True
End Function


Private Sub ModificarENTesoeria()
Dim C As String

    If vParamAplic.ContabilidadNueva Then
        C = DevuelveDesdeBD(conConta, "codforpa", "formapago", "codforpa", Text1(0).Text)
        If C = "" Then
            InsertarEnTesoreria
        Else
            C = "UPDATE formapago set nomforpa = '" & DevNombreSQL(Text1(1).Text) & "', tipforpa=" & Me.Combo1.ItemData(Combo1.ListIndex)
            C = C & ", numerove =" & Text1(2).Text & ",primerve = " & Text1(3).Text & " ,restoven =" & DBSet(Text1(4).Text, "N", "S")
            C = C & " WHERE codforpa = " & Text1(0).Text
        End If
    Else
        C = "UPDATE sforpa set nomforpa = '" & DevNombreSQL(Text1(1).Text) & "', tipforpa=" & Me.Combo1.ItemData(Combo1.ListIndex)
        C = C & " WHERE codforpa = " & Text1(0).Text
    End If
    If C <> "" Then ConnConta.Execute C
End Sub


Private Sub InsertarEnTesoreria()
Dim C As String
    On Error Resume Next
    
    If vParamAplic.ContabilidadNueva Then
        C = "INSERT INTO formapago(codforpa,nomforpa, tipforpa,numerove,primerve,restoven) VALUES ("
        C = C & Text1(0).Text & ",'" & DevNombreSQL(Text1(1).Text) & "'," & Me.Combo1.ItemData(Combo1.ListIndex) & ","
        C = C & Text1(2).Text & ",'" & Text1(3).Text & "'," & DBSet(Text1(4).Text, "N", "S") & ")"
    Else
        C = "INSERT INTO sforpa(codforpa,nomforpa, tipforpa) VALUES (" & Text1(0).Text & ",'" & DevNombreSQL(Text1(1).Text) & "'," & Me.Combo1.ItemData(Combo1.ListIndex) & ")"
    End If
    ConnConta.Execute C
    If Err.Number <> 0 Then
        MsgBox "Error insertando en tesoreria: " & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub
