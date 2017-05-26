VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComActPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar precios venta"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmComActPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   5445
      Width           =   1035
   End
   Begin VB.Frame Frame4 
      Caption         =   "Precio Venta"
      ForeColor       =   &H00972E0B&
      Height          =   950
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   8415
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "margen"
         Top             =   435
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "PVP nuevo"
         Top             =   435
         Width           =   1250
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   18
         Text            =   "PVP actual"
         Top             =   435
         Width           =   1250
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   960
         MaxLength       =   15
         TabIndex        =   17
         Text            =   "PVP actual"
         Top             =   435
         Width           =   1250
      End
      Begin VB.Label Label1 
         Caption         =   "PVP nuevo"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Margen %"
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "PVP actual"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Precio compra"
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tarifas de Precios"
      ForeColor       =   &H00972E0B&
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   2000
      Width           =   8415
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   30
         Tag             =   "Precio caja|N|N|||slista|precioa1|###,##0.0000|N|"
         Text            =   "precio caja"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   29
         Tag             =   "Precio actual|N|N|||slista|precioac|###,##0.0000||"
         Text            =   "precioac"
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   28
         Tag             =   "Cod. Tarifa|N|N|||slista|codlista|000|S|"
         Text            =   "tarif"
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   16
         TabIndex        =   27
         Tag             =   "Cod. Artic|N|N|||slista|codartic||S|"
         Text            =   "artic"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "margen"
         Top             =   2640
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   25
         Tag             =   "Precio nuevo|N|N|||slista|tmpprecioac|###,##0.0000||"
         Text            =   "precioac new"
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "precioa1 ne "
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComActPrecios.frx":000C
         Height          =   2220
         Left            =   240
         TabIndex        =   14
         Top             =   705
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3916
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   2
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
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar precio tarifa"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "Fecha cambio"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   5400
         Picture         =   "frmComActPrecios.frx":0021
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha cambio"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   8415
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         MaxLength       =   16
         TabIndex        =   11
         Text            =   "codartic codarti"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   5280
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
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   5445
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   5445
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2880
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
      TabIndex        =   6
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmComActPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public parCodArtic As String
Public parNomArtic As String
'Public parPrecioUC As Currency

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte


Dim cart As CArticulo




Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 5 'Lineas
'            If DatosOk Then
                If ModificarLinea Then
                    TerminaBloquear
                    NumReg = Data1.Recordset.AbsolutePosition
                    PonerModo 2
                    CancelaADODC Me.Data1
                    CargaGrid True
                    LLamaLineas 10
                    SituarDataPosicion Data1, NumReg, Indicador
                    PonerFocoGrid Me.DataGrid1
                End If
'            End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub





Private Sub cmdActualizar_Click()
'Actualizar los nuevos precios del artículo
' en la tabla de articulo y en la de tarifas (slista)
Dim b As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim MenError As String
Dim RN As ADODB.Recordset

    If Not BloqueaRegistro("slista", "codartic=" & DBSet(Me.parCodArtic, "T")) Then Exit Sub
    
    On Error GoTo ErrAct
    conn.BeginTrans
    
    MenError = ""
    
    'Actualizar el PVP del artículo
    '----------------------------------
    b = cart.CambiaPrecioVenta(Text1(4).Text, Text1(5).Text, Text1(1).Text, MenError)
    
    
    'Actualizar los precios de las TARIFAS del artículo si se ha modificado
    '-----------------------------------------------------------------------
    If b Then
        SQL = MontaSQLCarga(True)
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF And b
            'si los precios son distintos
            If (RS!precioac <> RS!tmpprecioac) Then
                Set cTar = New CTarifaArt
                If cTar.LeerDatos(Me.parCodArtic, RS!codlista) Then
                    b = cTar.ActualizarPrecios(Text1(1).Text, RS!tmpprecioac, 0, MenError, True)
                End If
                                
                
                Set cTar = Nothing
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    If b And vParamAplic.ActualizaPrecioEspecial Then
        ActualizarPrecioEspecialGenerico parCodArtic, ImporteFormateado(Text1(4).Text), False
        
    End If
    
    
    If b Then
        conn.CommitTrans
        MsgBox "Precios actualizados correctamente.", vbInformation
    Else
        conn.RollbackTrans
        MsgBox "NO SE HA PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf & vbCrLf & MenError, vbExclamation
    End If
    
    TerminaBloquear
    Unload Me
    Exit Sub
    
ErrAct:
    TerminaBloquear
    MuestraError Err.Number, "Actualizar Precios." & vbCrLf & MenError, Err.Description
End Sub


Private Sub cmdCancelar_Click()

    On Error GoTo ECancelar
    
    Select Case Modo
        Case 5 'Lineas
            TerminaBloquear
'            NumRegElim = Data1.Recordset.AbsolutePosition
'            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
'            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
'            lblIndicador.Caption = Indicador
        Case Else
             Unload Me
'
'        Case 4  'Modificar
'            TerminaBloquear
'            NumRegElim = Data1.Recordset.AbsolutePosition
'            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
'            PonerModo 2
'            LLamaLineas 10
'            DataGrid1.Enabled = True
'            SituarDataPosicion Data1, NumRegElim, Indicador
'            DeseleccionaGrid Me.DataGrid1
'            lblIndicador.Caption = Indicador
    End Select

ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 0 Then Modo = 4
End Sub


Private Sub Form_Load()
'Dim cArt As CArticulo

    'Icono del formulario
'    Me.Icon = frmPpal.Icon

'    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
        .Buttons(1).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(10).Image = 16 'Imprimir
'        .Buttons(11).Image = 15 'Salir
    End With

    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    Text1(0).Text = Me.parCodArtic
    Text2(0).Text = Me.parNomArtic
    
    Set cart = New CArticulo
    If cart.LeerDatos(Me.parCodArtic) Then
        Text1(2).Text = cart.PrecioUltCom
        Text1(2).Text = Format(Text1(2).Text, FormatoPrecio)
        
        Text1(3).Text = cart.PrecioVenta 'precio venta actual
        Text1(3).Text = Format(Text1(3).Text, FormatoPrecio)
        Text1(5).Text = cart.MargenComercial
        Text1(5).Text = Format(Text1(5).Text, FormatoPorcen)
        
        Text1(4).Text = cart.AplicarMargenComercial 'obtiene el nuevo PVP
        Text1(4).Text = Format(Text1(4).Text, FormatoPrecio)
    Else
        cmdActualizar.Enabled = False
    End If
'    Set cArt = Nothing
    
    'Fecha cambio tarifas
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    NombreTabla = "slista" 'Tabla tarifas de articulo
    Ordenacion = " ORDER BY codartic,codlista "
'    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = " & DBSet(Me.parCodArtic, "T")
'    Data1.ConnectionString = Conn
'    Data1.RecordSource = CadenaConsulta
'    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then PonerModo 0


    'Antes de cargar el Grid actualizamos los nuevos precios de las tarifas
    'el los campos temporales de la tabla slista
    CalcularPreciosNuevosTarifas

    CargaGrid True
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim tots As String
    
    On Error GoTo ECarga
    
    tots = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, tots, False
    
    
    tots = "N||||0|;S|txtAux(1)|T|Tarif|540|;S|txtAux2(1)|T|Desc. Tarifa|2550|;S|txtAux(2)|T|Precio|1200|;"
    tots = tots & "N||||0|;" '"S|txtAux(3)|T|Precio Caja|1100|;"
    tots = tots & "S|txtAux2(2)|T|% PVP|900|;S|txtAux(4)|T|Precio Nuevo     |1200|;"
    '                           Nuevo. Tipo: PVP o UPC
    tots = tots & "N||||0|;"
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(5).NumberFormat = FormatoPorcen & " "
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(6).NumberFormat = FormatoPrecio & " "


    DataGrid1.ScrollBars = dbgAutomatic
    

    Me.Toolbar1.Buttons(1).Enabled = (Modo <> 5) And Data1.Recordset.RecordCount > 0
    
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
'    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar
    b = (Modo = 5)

    For jj = 4 To 4
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
'    txtaux2(1).Height = Me.DataGrid1.RowHeight
'    txtaux2(1).Top = alto
'    txtaux2(1).visible = b
    
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Not cart Is Nothing Then Set cart = Nothing
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
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

                      
    'codartic siempre bloqueado
    BloquearTxt Text1(0), True
    BloquearTxt Text1(2), True
    BloquearTxt Text1(3), True
    
    BloquearTxt Text1(1), Modo = 5
    BloquearTxt Text1(4), Modo = 5 'nuevo PVP
    BloquearTxt Text1(5), Modo = 5 'margen comercial
                      
                                 
    BloquearTxt txtAux(0), (Modo = 4)
                      
    '-----------------------------------------
'    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = True
    cmdAceptar.visible = (Modo = 5)
    Me.cmdActualizar.visible = (Modo <> 5)

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
'    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub PonerModoOpcionesMenu()
''Activas unas Opciones de Menu y Toolbar según el modo en que estemos
'Dim b As Boolean

    Me.Toolbar1.Buttons(1).Enabled = (Modo <> 5)
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
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
    
    SQL = "SELECT " & NombreTabla & ".codartic, " & NombreTabla & ".codlista,starif.nomlista, precioac, if(isnull(precioa1),0,precioa1) as precioa1,"
    SQL = SQL & "if(isnull(starif.margecom),0,starif.margecom) as margen,tmpprecioac,tmpprecioa1,  "
    'NUEVO. Trifas sobre UPC (ultimo precio compra)
    'Indicare en el grid que tipo tarifa es
    SQL = SQL & " if(opcionINC=0,""PVP"",""UPC"") as tipo"
    SQL = SQL & " FROM " & NombreTabla & " INNER JOIN starif ON " & NombreTabla & ".codlista = starif.codlista "
    
    If enlaza Then
        SQL = SQL & " WHERE " & NombreTabla & ".codartic=" & DBSet(Me.parCodArtic, "T")
        SQL = SQL & " AND not isnull(margecom) "
    Else
        SQL = SQL & " WHERE codlista = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function




Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Lineas para modificar
    PonerModo 5
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 0)
    LLamaLineas anc
    
    'codartic
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")
    
    'codtarifa
    txtAux(1).Text = DBLet(Me.DataGrid1.Columns(1).Value, "N")
    FormateaCampo txtAux(1)
    'desc. tarifa
    txtAux2(1).Text = DBLet(DataGrid1.Columns(2).Value, "T")
    
    'precio actual
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "N")
    FormateaCampo txtAux(2)
    
    'precio nuevo
    txtAux(4).Text = DBLet(DataGrid1.Columns(6).Value, "N")
    FormateaCampo txtAux(4)

    DataGrid1.Enabled = False
    PonerFoco txtAux(4)
End Sub






Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

'   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)

End Sub



Private Sub Text1_GotFocus(Index As Integer)
     ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim MargenTemporal As Currency
Dim OK As Boolean

    'If Modo = 2 Then Modo = 4
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    
    
    'solo se puede modificar el campo PVP nuevo y el margen comercial
    Select Case Index
        Case 4 'PVP nuevo
            'si el precio lo hemos modificado y es distinto del valor obtenido
            'al aplicar el margen al precio_ultima_compra, recalcular para
            'las tarifas el % sobre el PVP
            If PonerFormatoDecimal(Text1(Index), 2) Then
                Text1(Index).Text = Format(Text1(Index).Text, FormatoPrecio)
                If ImporteFormateado(Text1(Index).Text) <> cart.AplicarMargenComercial Then
                    
                    If vParamAplic.RecalculoMargen Then
                        'Significa que ajustamos el maren al precio de venta /copra
                        If Text1(2).Text <> "" Then
                            MargenTemporal = Round2(ImporteFormateado(Text1(4).Text) / ImporteFormateado(Text1(2).Text), 4)
                            MargenTemporal = (MargenTemporal - 1) * 100
                            If Val(MargenTemporal) > 999 Then
                                MsgBox "margen calculado:" & Format(MargenTemporal, FormatoPrecio) & "   superior al 999%. No se actualizara"
                                MargenTemporal = cart.MargenComercial
                            Else
                                cart.MargenComercial = MargenTemporal
                            End If
                            Text1(5).Text = Format(MargenTemporal, FormatoPorcen)
                            
                                            
                                            
                        End If
                    Else
                        OK = True
                    End If
                    
                    CalcularPreciosNuevosTarifas
                    CargaGrid True
                
                End If
            Else
                Text1(4).Text = Format(cart.PrecioVenta, FormatoPrecio)
            End If
        Case 5 'margen comercial
            If Text1(Index).Text <> "" Then
                If PonerFormatoDecimal(Text1(Index), 7) Then 'tipo 7: Decimal(5,2)
                    'recalcular los nuevos PVP del articulo y de las tarifas
                    cart.MargenComercial = CCur(Text1(Index).Text)
                    Text1(4).Text = cart.AplicarMargenComercial 'obtiene el nuevo PVP
                    Text1(4).Text = Format(Text1(4).Text, FormatoPrecio)
                    CalcularPreciosNuevosTarifas
                    CargaGrid True
                Else
                    'MsgBox "Introduzca valor para el margen comercial", vbInformation
                    Text1(Index).Text = Format(cart.MargenComercial, FormatoPorcen)
                    PonerFoco Text1(Index)
                End If
            End If
    End Select

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Modificar linea
            If Data1.Recordset.EOF Then Exit Sub
            If BloqueaRegistro("slista", "codartic=" & DBSet(Me.parCodArtic, "T") & " AND codlista=" & Me.Data1.Recordset!codlista) Then
                BotonModificar
            End If
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
                If Index = 4 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    On Error Resume Next
    
    If Not PerderFocoGnralLineas(txtAux(Index), 2) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio nuevo
            If PonerFormatoDecimal(txtAux(Index), 2) Then   'tipo 2 -> Decimal(10,4)
                PonerFocoBtn Me.cmdAceptar
            Else
                PonerFoco txtAux(Index)
            End If
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function CalcularPreciosNuevosTarifas() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt

    On Error GoTo ErrTarifa
    
    SQL = MontaSQLCarga(True)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set cTar = New CTarifaArt
        If cTar.LeerDatos(RS!codArtic, RS!codlista) Then
            cTar.AplicarMargenComercial (CCur(Text1(4).Text))
        End If
        Set cTar = Nothing

        RS.MoveNext
    Wend

    RS.Close
    Set RS = Nothing
    
    CalcularPreciosNuevosTarifas = True
    Exit Function

ErrTarifa:
    CalcularPreciosNuevosTarifas = False
    MuestraError Err.Number, "Calcular nuevos precios de las tarifas.", Err.Description
End Function



Private Function ModificarLinea() As Boolean
Dim SQL As String

    On Error GoTo ErrMod
    SQL = "UPDATE " & NombreTabla & " SET tmpprecioac=" & DBSet(txtAux(4).Text, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(Me.parCodArtic, "T") & " AND codlista=" & Me.Data1.Recordset!codlista
    conn.Execute SQL
    
    ModificarLinea = True
    Exit Function
    
ErrMod:
    ModificarLinea = False
    MuestraError Err.Number, "Modificar linea tarifa.", Err.Description
End Function
