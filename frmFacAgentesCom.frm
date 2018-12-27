VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacAgentesCom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agentes Comerciales"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmFacAgentesCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   3735
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
         TabIndex        =   19
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   14
      Tag             =   "Comisión PVP minimo|N|S|0|99,90|sagent|comsiopvpmin|#0.00|N|"
      Text            =   "Text1"
      Top             =   5400
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   13
      Tag             =   "Comisión General|N|S|0|99,90|sagent|comisio1s|#0.00|N|"
      Text            =   "Text1"
      Top             =   4800
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   12
      Tag             =   "Comisión General|N|S|0|99,90|sagent|comsio1n|#0.00|N|"
      Text            =   "Text1"
      Top             =   4800
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Comisión General|N|S|||sagent|coddelega||N|"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   11
      Tag             =   "Comisión General|N|S|0|99,90|sagent|comisios|#0.00|N|"
      Text            =   "Text1"
      Top             =   4200
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   10
      Tag             =   "Comisión General|N|S|0|99,90|sagent|comisioc|#0.00|N|"
      Text            =   "Text1"
      Top             =   4200
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   9
      Tag             =   "Comisión General|N|S|0|99,90|sagent|comision|#0.00|N|"
      Text            =   "Text1"
      Top             =   2880
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   180
      MaxLength       =   15
      TabIndex        =   7
      Tag             =   "Móvil|T|S|||sagent|telagent||N|"
      Text            =   "Text1"
      Top             =   2880
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   4500
      MaxLength       =   9
      TabIndex        =   2
      Tag             =   "N.I.F|T|S|||sagent|nifagent||N|"
      Text            =   "Text1"
      Top             =   810
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   4440
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Provincia|T|N|||sagent|proagent||N|"
      Text            =   "Text1"
      Top             =   2115
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   180
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código Agente Comercial|N|N|0|9999|sagent|codagent|0000|S|"
      Text            =   "Text"
      Top             =   840
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Población|T|N|||sagent|pobagent||N|"
      Text            =   "Text1"
      Top             =   2088
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   180
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "C. Postal|T|N|||sagent|codpobla||N|"
      Text            =   "Text1"
      Top             =   2088
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   180
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Domicilio|T|S|||sagent|domagent||N|"
      Text            =   "Text1"
      Top             =   1464
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1020
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre del Agente Comercial|T|N|||sagent|nomagent||N|"
      Text            =   "Text1"
      Top             =   840
      Width           =   3165
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   6240
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6000
      Top             =   600
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
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
         Left            =   5160
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Comisión PVP min"
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
      Index           =   14
      Left            =   240
      TabIndex        =   38
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Familia especial  para fontenas en su report iran separados lienas a la familia y lineas distinta"
      Height          =   1455
      Left            =   5520
      TabIndex        =   37
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   4800
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   35
      Top             =   4320
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   " Comisión por lineas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   34
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label1 
      Caption         =   "Delegación"
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   33
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Sin dto."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   32
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Con dto."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   31
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Left            =   825
      Picture         =   "frmFacAgentesCom.frx":000C
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1860
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Comisión general"
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   30
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Móvil"
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   29
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "N.I.F."
      Height          =   255
      Index           =   6
      Left            =   4500
      TabIndex        =   28
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   27
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   4
      Left            =   1140
      TabIndex        =   26
      Top             =   1875
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   25
      Top             =   1875
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   24
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1020
      TabIndex        =   21
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   20
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
Attribute VB_Name = "frmFacAgentesCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin ningun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    TratarAgenteTesoreria
                    PosicionarData
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TratarAgenteTesoreria
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        Case 1  'BUSCAR
            HacerBusqueda
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3 'Insertar
        LimpiarCampos
        PonerModo 0
        PonerOpcionesMenu
    Case 4  'Modificar
        lblIndicador.Caption = ""
        TerminaBloquear
        PonerModo 2
        PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    'Sugerir el siguiente codigo
    Text1(0).Text = Format(SugerirCodigoSiguienteStr("sagent", "codagent"), "0000")
    PonerFoco Text1(0)
    Text1_GotFocus 0
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
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
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    'Copmpruebo si esta vinculado a algun cliente
    Cad = DevuelveDesdeBD(conAri, "count(*)", "sclien", "codagent", CStr(Data1.Recordset!CodAgent))
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen clientes asociados a este agente.", vbExclamation
        Exit Sub
    End If
    Cad = DevuelveDesdeBD(conAri, "count(*)", "sclien", "visitador", CStr(Data1.Recordset!CodAgent))
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen clientes(visitador) asociados a este agente.", vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Copmpruebo si esta vinculado a algun trabajador
    Cad = DevuelveDesdeBD(conAri, "count(*)", "straba", "codagent", CStr(Data1.Recordset!CodAgent))
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existe trabajador asociado a este agente.", vbExclamation
        Exit Sub
    End If
    
    If vParamAplic.ContabilidadNueva Then
        Cad = DevuelveDesdeBD(conConta, "count(*)", "cobros", "agente", CStr(Data1.Recordset!CodAgent))
    Else
        Cad = DevuelveDesdeBD(conConta, "count(*)", "scobro", "agente", CStr(Data1.Recordset!CodAgent))
    End If
    If Cad <> "" Then
        If Val(Cad) = 0 Then Cad = ""
    End If
    
    If Cad = "" Then
    
            Cad = "¿Seguro que desea eliminar el Agente Comercial? " & vbCrLf
            Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), "0000")
            Cad = Cad & vbCrLf & "Descripción: " & Data1.Recordset.Fields(1)
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        'EXSITEN Vtos
        Cad = "Existen " & Cad & " vencimiento(s) en Arimoney para este agente."
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
            Exit Sub
        Else
            Cad = Cad & vbCrLf & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    'Borramos

        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Cad = "Delete from sagent where codagent=" & Data1.Recordset!CodAgent
        conn.Execute Cad
        
        'En tesoreria
        Cad = "DELETE FROM agentes WHERE codigo = " & Data1.Recordset!CodAgent
        ConnConta.Execute Cad
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If

    Screen.MousePointer = vbDefault
    
Error2:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Agente Comercial", Err.Description
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


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
    VieneDeBuscar = False
    
    'Dicembre 2014. Nueva comision
    If DatosADevolverBusqueda <> "" Then
        Me.Height = 2595
    Else
        Me.Height = 7485
    End If
    Me.Frame1.Top = Me.Height - 1305
    Me.cmdAceptar.Top = Me.Height - 1185
    cmdCancelar.Top = Me.cmdAceptar.Top
    cmdRegresar.Top = Me.cmdAceptar.Top
    
     
     
    'Ocultamos campos para herbelca
    Label4.Caption = "Familia especial"
    If vParamAplic.NumeroInstalacion = 2 Then Label4.Caption = "Comision ECO"
    Me.Label1(9).visible = vParamAplic.NumeroInstalacion <> 2
    Me.Label1(10).visible = vParamAplic.NumeroInstalacion <> 2
    Text1(10).visible = vParamAplic.NumeroInstalacion <> 2
    Text1(12).visible = vParamAplic.NumeroInstalacion <> 2
    
    '## A mano
    NombreTabla = "sagent"
    Ordenacion = " ORDER BY codagent"
        
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codagent=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    
    CargarCombo_Tabla Me.Combo1, "sdelega", "coddelega", "nombre", , True
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub
    
Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(indice + 2).Text = devuelve
End Sub


Private Sub imgBuscar_Click()
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    'Codigo Postal
    Set frmCP = New frmCPostal
    frmCP.DatosADevolverBusqueda = "0"
    frmCP.Show vbModal
    Set frmCP = Nothing

    PonerFoco Text1(3)
    VieneDeBuscar = True
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


'### A mano
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
Dim devuelve As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'cod Agente Com.
            If PonerFormatoEntero(Text1(Index)) Then
                'comprobar si ya existe ese codigo de Agente en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If

        Case 3 'CPostal
            If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
            ElseIf Not VieneDeBuscar Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                 Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 6 'NIF
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                ValidarNIF Text1(Index).Text
            End If
            
        Case 8, 9, 10, 11, 12, 13 'Comision General
            'Tipo 4: Decimal(4,2)
            PonerFormatoDecimal Text1(Index), 4
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then     'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 30, "Código")
        Cad = Cad & ParaGrid(Text1(1), 70, "Denominacion")
'        Cad = Cad & ParaGrid(Combo1, 20, "Tipo Pago")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano

            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Agentes Comerciales"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                PonerFocoBtn Me.cmdRegresar
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerModo Modo
                PonerFoco Text1(kCampo)
            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
         Screen.MousePointer = vbDefault
         'PonerModo 0
         Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
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
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo

    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    B = Modo = 1 Or Modo >= 3 'busqueda o inser/mod
    BloquearCmb Combo1, Not B
    
    '---------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    
    B = (Modo = 2 Or Modo = 0 Or Modo = 1)
    
    If B Then
        If DatosADevolverBusqueda <> "" Then B = False
    End If
    
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
     If B Then
        If DatosADevolverBusqueda <> "" Then B = False
    End If
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    '----------------------------------------
    B = (Modo >= 3) 'Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
'Dim cad As String

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar datos OK
    If Not B Then Exit Function
        
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


Private Sub PosicionarData()
Dim Cad As String
Dim Indicador As String

    Cad = "(codagent=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
Dim Sql As String
On Error Resume Next
    Sql = " WHERE codagent= " & Text1(0).Text
    ObtenerWhereCP = Sql
End Function



Private Sub TratarAgenteTesoreria()
Dim C As String
Dim Nuevo As Boolean

    C = DevuelveDesdeBD(conConta, "Nombre", "agentes", "Codigo", Text1(0).Text)
    Nuevo = C = ""
    If vParamAplic.ContabilidadNueva Then
        If Nuevo Then
            C = "agentes(Codigo,Nombre,domagent,codpobla,pobagent,proagent,nifagent,telagent,comision,comisioc,comisios,coddelega,comsio1n,comisio1s,comsiopvpmin) VALUES ("
            C = C & Text1(0).Text & "," & DBSet(Text1(1).Text, "T") & "," & DBSet(Text1(2).Text, "T", "S") & ","
            'codpobla,pobagent,proagent
            C = C & DBSet(Text1(3).Text, "N") & "," & DBSet(Text1(4).Text, "T", "N") & "," & DBSet(Text1(5).Text, "T", "N") & ","
            'nifagent,telagent,comision
            C = C & DBSet(Text1(6).Text, "T", "S") & "," & DBSet(Text1(7).Text, "T", "S") & "," & DBSet(Text1(8).Text, "N", "S") & ","
            'comisioc,comisios,coddelega
            C = C & DBSet(Text1(9).Text, "T", "S") & "," & DBSet(Text1(10).Text, "T", "S") & ","
            If Combo1.ListIndex < 0 Then
                C = C & "null"
            Else
                C = C & Combo1.ItemData(Combo1.ListIndex)
            End If
            'comsio1n,comisio1s,comsiopvpmin
            C = C & "," & DBSet(Text1(11).Text, "T", "S") & "," & DBSet(Text1(12).Text, "T", "S") & "," & DBSet(Text1(13).Text, "T", "S") & ")"
            C = "INSERT INTO " & C
    
    
        Else
            
            C = " Nombre = " & DBSet(Text1(1).Text, "T") & ",domagent=" & DBSet(Text1(2).Text, "T", "S")
            'codpobla,pobagent,proagent
            C = C & ",codpobla = " & DBSet(Text1(3).Text, "N") & ",pobagent=" & DBSet(Text1(4).Text, "T", "N") & ",proagent=" & DBSet(Text1(5).Text, "T", "N")
            'nifagent,telagent,comision
            C = C & ", nifagent =" & DBSet(Text1(6).Text, "T", "S") & ",telagent=" & DBSet(Text1(7).Text, "T", "S") & ",comision =" & DBSet(Text1(8).Text, "N", "S")
            'comisioc,comisios,coddelega
            C = C & ",comisioc=" & DBSet(Text1(9).Text, "T", "S") & ",comisios=" & DBSet(Text1(10).Text, "T", "S") & ",coddelega="
            If Combo1.ListIndex < 0 Then
                C = C & "null"
            Else
                C = C & Combo1.ItemData(Combo1.ListIndex)
            End If
            'comsio1n,comisio1s,comsiopvpmin
            C = C & ",comsio1n=" & DBSet(Text1(11).Text, "T", "S") & ",comisio1s=" & DBSet(Text1(12).Text, "T", "S") & ",comsiopvpmin =" & DBSet(Text1(13).Text, "T", "S")
            C = "UPDATE agentes set " & C & " WHERE codigo =" & Text1(0).Text
    
        
        End If
    Else
        '// Antigua contabilidad CONTA
        If Nuevo Then
            C = "insert into `agentes` (`Codigo`,`Nombre`) values ( "
            C = C & Text1(0).Text & "," & DBSet(Text1(1).Text, "T") & ")"
            
        Else
            C = "UPDATE agentes set nombre=" & DBSet(Text1(1).Text, "T")
            C = C & " WHERE codigo = " & Text1(0).Text
        End If
    End If
    
    On Error Resume Next
    ConnConta.Execute C
    If Err.Number <> 0 Then
        C = Err.Description
        C = "Error actualizando en contabilidad: " & vbCrLf & vbCrLf & C
        MsgBox C, vbExclamation
    End If

End Sub
