VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmInventarioGR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Inventario Real"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13545
   ClipControls    =   0   'False
   Icon            =   "frmAlmInventarioGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   18
      Top             =   90
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   19
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
      Left            =   11700
      TabIndex        =   17
      Top             =   225
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar fichero externo"
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
      Left            =   7905
      TabIndex        =   15
      Top             =   8820
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   135
      TabIndex        =   10
      Top             =   795
      Width           =   13275
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   7440
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Familia|N|N|0|9999|sinven|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Left            =   8250
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Cod. Almacen|N|N|0|999|sinven|codalmac|000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   240
         Width           =   3585
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1215
         Picture         =   "frmAlmInventarioGR.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacén"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6105
         TabIndex        =   14
         Top             =   285
         Width           =   990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   7155
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   270
         Width           =   870
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "existencia"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   8655
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   180
         Width           =   2355
      End
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
      Left            =   11100
      TabIndex        =   2
      Top             =   8820
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
      Left            =   12375
      TabIndex        =   3
      Top             =   8820
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
      Left            =   12375
      TabIndex        =   4
      Top             =   8820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmInventarioGR.frx":0A0E
      Height          =   6960
      Left            =   120
      TabIndex        =   6
      Top             =   1635
      Width           =   13295
      _ExtentX        =   23442
      _ExtentY        =   12277
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8880
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
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
      Left            =   180
      TabIndex        =   7
      Top             =   8775
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmAlmInventarioGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
'Segun el Parametro se trabajara con Familia o con Proveedor (frmFA o frmP)
Private WithEvents frmFA As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmFA.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmFI As frmInvPistolas
Attribute frmFI.VB_VarHelpID = -1

Private Modo As Byte

Dim kCampo As Integer

Dim CadenaConsulta As String

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim cad As String

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    '-- RAFA: QTD Pistolas
    cmdImportar.visible = False
    lblInfInv.Caption = ""
    Select Case Modo
        Case 1 'Buscar registros a inventariar
            If Text1(0).Text <> "" And Text1(1).Text <> "" Then
                CargaGrid True
                'Poner Modo Modificar columna Existencia Real
                If Not Data1.Recordset.EOF Then
                    PonerModo 4
                    CargaTxtAux True, True
                    cmdImportar.visible = True
                Else 'No existen registros en la tabla sinven para ese criterio de búsqueda
                    cad = "No se esta realizando inventario para esos criterios de búsqueda."
                    MsgBox cad, vbInformation
                    PonerFoco Text1(0)
                End If
            Else
                cad = "Criterio de Búsqueda incompleto." & vbCrLf
                cad = cad & "Debe introducir ambos criterios de búsqueda: cod. almacen, "
                If vParamAplic.InventarioxProv Then
                    cad = cad & "Cod. Proveedor"
                Else
                    cad = cad & "Cod. Familia"
                End If
                MsgBox cad, vbExclamation
                PonerFoco Text1(0)
            End If
            
        Case 4 'Modificar Existencia Real (Introducir Valores Reales)
            CargaTxtAux False, False
            PonerModo 2
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar
    
     
    cmdImportar.visible = False
    lblInfInv.Caption = ""
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid False
        Case 4  ' 4: Modificar
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid True
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdImportar_Click()
    Set frmFI = New frmInvPistolas
    frmFI.Show vbModal
    CargaGrid True
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(4).Image = 4 'Modificar
'        .Buttons(5).Image = 15 'Salir
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
    '++
    imgBuscar(1).Picture = imgBuscar(0).Picture
    
    'cargar el TAG en funcion del valor del parametro haydepar de la tabla spara1
    If vParamAplic.InventarioxProv Then
        Label4.Caption = "Proveedor"
        Text1(1).Tag = "Cod. Proveedor|N|N|||sinven|codprove|000000|N|"
    Else
        Label4.Caption = "Familia"
        Text1(1).Tag = "Cod. Familia|N|N|||sinven|codfamia|0000|N|"
    End If
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True

    PonerModo 0
    CargaGrid (Modo = 2)
            
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub CargaGrid(enlaza As Boolean)
Dim i As Byte
Dim SQL As String
On Error GoTo ECarga

    gridCargado = False
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    PrimeraVez = False
        
    'Cod. Articulo
    DataGrid1.Columns(0).Caption = "Artículo"
    DataGrid1.Columns(0).Width = 2150
    DataGrid1.Columns(1).Caption = "Descripción"
    DataGrid1.Columns(1).Width = 4800
       
    'Fecha Inventario
    DataGrid1.Columns(2).Caption = "Fecha"
    DataGrid1.Columns(2).Width = 1350
    DataGrid1.Columns(2).Alignment = dbgCenter
    
    'Hora Inventario
    DataGrid1.Columns(3).Caption = "Hora"
    DataGrid1.Columns(3).Width = 1100
    DataGrid1.Columns(3).NumberFormat = "hh:mm:ss"
    DataGrid1.Columns(3).Alignment = dbgCenter
    
    'Existencia Real
    DataGrid1.Columns(4).Caption = "Existencia"
    DataGrid1.Columns(4).Width = 1650
    DataGrid1.Columns(4).Alignment = dbgCenter
    
    'Diferencia
    DataGrid1.Columns(5).Caption = "Diferencia"
    DataGrid1.Columns(5).Width = 1650
    DataGrid1.Columns(5).Alignment = dbgCenter
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.RowHeight = 350
    
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = Data1.Recordset!existenc
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(4).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(4).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
        txtAux.SelStart = 0
        txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFA_DatoSeleccionado(CadenaSeleccion As String)
'Familia Articulos
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFI_Seleccionado(CADENA As String)
    Dim resultado As Boolean
    If CADENA <> "" Then
        resultado = ProcesarFicheroInventario(CADENA)
        If Not resultado Then
            MsgBox "Se ha producido errores durante el proceso de captura, consulte el fichero LOG para más detalles", vbCritical
        Else
            MsgBox "Proceso finalizado", vbInformation
        End If
    End If
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Proveedor
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0 'Codigo Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
        Case 1 'Codigo Familia / Cod. Proveedor
            If vParamAplic.InventarioxProv Then
                'Realizar inventario por Proveedor
'                Set frmP = New frmComProveedores
'                frmP.DatosADevolverBusqueda = "0"
'                frmP.Show vbModal
                Set frmP = New frmBasico2
                AyudaProveedores frmP, Text1(Index)
                Set frmP = Nothing
            Else 'Cod. Familia
'                Set frmFA = New frmAlmFamiliaArticulo
'                frmFA.DatosADevolverBusqueda = "0"
'                frmFA.Show vbModal
                Set frmFA = New frmBasico2
                AyudaFamilias frmFA, Text1(Index)
                Set frmFA = Nothing
            End If
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'almacen
            Case 1: KEYBusqueda KeyAscii, 1 'familia
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim campo As String
Dim tabla As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    If Text1(Index).Text = "" Then
        Text2(Index).Text = ""
    Else
        Select Case Index
            Case 0 'Codigo Almacen
                campo = "nomalmac"
                tabla = "salmpr"
            Case 1 'Codigo Familia/ Cod. Proveedor
                If vParamAplic.InventarioxProv Then
                'Realizar inventario por Proveedor
                    campo = "nomprove"
                    tabla = "sprove"
                Else
                    campo = "nomfamia"
                    tabla = "sfamia"
                End If
        End Select
        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, tabla, campo)
        If Text1(Index).Text <> "" And Text2(Index).Text = "" Then PonerFoco Text1(Index)
     End If
End Sub


Private Sub txtAux_GotFocus()
    ConseguirFocoLin txtAux
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
'                If DataGrid1.Row > 0 Then
'                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
''                elseif
'                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        ModificarExistencia
        PasarSigReg
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        PonerFormatoDecimal txtAux, 1
    End With
'    PonerFocoBtn Me.cmdAceptar
'    cmdAceptar_Click
'    If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
''    If Me.Data1.Recordset.EOF Then
'
'        If DataGrid1.Row <= 12 And Data1.Recordset.AbsolutePosition <> Data1.Recordset.RecordCount Then DataGrid1.Row = DataGrid1.Row + 1
''        CargaTxtAux True, True
'    Else
'        CargaTxtAux False, False
'        PonerModo 2
'    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Modificar
            BotonModificar
        Case 5 'Busqueda
            BotonBuscar
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    'b = (Kmodo = 2)
   
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
'    BloquearText1 Me, Modo
    b = (Modo <> 1)
    BloquearTxt Text1(0), b
    BloquearTxt Text1(1), b
    
    b = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera b
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
        PonerFoco Text1(0)
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select
           
    b = Modo <> 0 And Modo <> 2 And Modo <> 4
   
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i

    b = (Modo = 1)
    Toolbar1.Buttons(2).Enabled = Not b And (Not (Modo = 0 Or Modo = 4))
    Toolbar1.Buttons(5).Enabled = Not b

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
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

    SQL = "SELECT sinven.codartic, sartic.nomartic, "
    SQL = SQL & " sinven.fechainv, sinven.horainve, sinven.existenc, sinven.existenc - salmac.canstock as diferencia "
    SQL = SQL & " FROM (sinven INNER JOIN sartic on sinven.codartic=sartic.codartic) INNER JOIN salmac ON sinven.codalmac=salmac.codalmac and sinven.codartic=salmac.codartic"
    SQL = SQL & " INNER JOIN salmpr ON sinven.codalmac=salmpr.codalmac"
    
    If enlaza Then
        SQL = SQL & " WHERE sinven.codalmac = " & Text1(0).Text & " AND "
        If vParamAplic.InventarioxProv Then
            SQL = SQL & "sinven.codprove=" & Text1(1).Text
        Else
            SQL = SQL & "sinven.codfamia=" & Text1(1).Text
        End If
    Else
        SQL = SQL & " WHERE sinven.codalmac = -1"
    End If

    If vParamAplic.InventarioxProv Then
        '-- Modificado por LAURA 29/08/2007  --> DAVID 13Junio11  Vuelvo a poner codartic
        SQL = SQL & " ORDER BY sinven.codprove, sinven.codfamia, codartic"
        'Sql = Sql & " ORDER BY sinven.codprove, sinven.codfamia, nomartic"
    Else
        '-- Modificado por LAURA 29/08/2007
        'SQL = SQL & " ORDER BY sinven.codfamia, codartic"
        If vParamAplic.SituaEnCodigoArticulo Then
            'Por codigo de articulopac
            SQL = SQL & " ORDER BY sinven.codfamia, sinven.codartic"
        Else
            If vParamAplic.InventarioCodigoArticulo Then
                SQL = SQL & " ORDER BY sinven.codfamia, codartic"
            Else
                SQL = SQL & " ORDER BY sinven.codfamia, nomartic"
            End If
        End If
    End If
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        CargaTxtAux False, False
        Text1(0).BackColor = vbLightBlue
    Else
        'Ya estamos en Modo de Busqueda
'        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux.Text = Trim(txtAux.Text)

    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
        If PonerFormatoDecimal(txtAux, 1) Then
            DatosOk = True
        Else
            DatosOk = False
        End If
        'DatosOk = True
    Else
        DatosOk = False
    End If
End Function


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarExistencia(canti As String) As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim ADonde As String

    On Error GoTo EActualizar

    conn.BeginTrans
    'Actualizar la Tabla: sinven con la cantidad introducida
    '-------------------------------------------------------
    ADonde = "Modificando datos de Inventario (Tabla: sinven)."
    SQL = "UPDATE sinven Set existenc = " & DBSet(canti, "N")
    SQL = SQL & " WHERE codartic =" & DBSet(Data1.Recordset!codArtic, "T") & " AND "
    SQL = SQL & " codalmac =" & Val(Text1(0).Text)
    conn.Execute SQL
    
    
    'Actualizar la Tabla: salmac el campo stockinv con la cantidad introducida
    '-------------------------------------------------------------------------
    ADonde = "Modificando datos de Articulos en Almacen. Tabla: salmac."
    SQL = "UPDATE salmac Set stockinv = " & DBSet(canti, "N")
    SQL = SQL & " WHERE codartic =" & DBSet(Data1.Recordset!codArtic, "T") & " AND "
    SQL = SQL & " codalmac =" & Val(Text1(0).Text)
    conn.Execute SQL
    
    ActualizarExistencia = True
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         SQL = "Actualizando Diferencias de Inventario." & vbCrLf & "--------------------------------------------" & vbCrLf
         SQL = SQL & ADonde
         MuestraError Err.Number, SQL, Err.Description
         conn.RollbackTrans
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
        conn.CommitTrans
    End If
End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOk Then
        If ActualizarExistencia(txtAux.Text) Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid True
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function

'----------------------- RAFA: Fichero de inventario QTD
Private Function ActualizarExistencia2(canti As String, codArtic As String) As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim ADonde As String

    On Error GoTo EActualizar

    conn.BeginTrans
    'Actualizar la Tabla: sinven con la cantidad introducida
    '-------------------------------------------------------
    ADonde = "Modificando datos de Inventario (Tabla: sinven)."
    SQL = "UPDATE sinven Set existenc = " & DBSet(canti, "N")
    SQL = SQL & " WHERE codartic =" & DBSet(codArtic, "T") & " AND "
    SQL = SQL & " codalmac =" & Val(Text1(0).Text)
    conn.Execute SQL
    
    
    'Actualizar la Tabla: salmac el campo stockinv con la cantidad introducida
    '-------------------------------------------------------------------------
    ADonde = "Modificando datos de Articulos en Almacen. Tabla: salmac."
    SQL = "UPDATE salmac Set stockinv = " & DBSet(canti, "N")
    SQL = SQL & " WHERE codartic =" & DBSet(codArtic, "T") & " AND "
    SQL = SQL & " codalmac =" & Val(Text1(0).Text)
    conn.Execute SQL
    
    ActualizarExistencia2 = True
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         SQL = "Actualizando Diferencias de Inventario." & vbCrLf & "--------------------------------------------" & vbCrLf
         SQL = SQL & ADonde
         MuestraError Err.Number, SQL, Err.Description
         conn.RollbackTrans
         ActualizarExistencia2 = False
    Else
        ActualizarExistencia2 = True
        conn.CommitTrans
    End If
End Function

Private Function ProcesarFicheroInventario(Fichero As String) As Boolean
    '----- ProcesarFicheroInventario:
    '   Procesa el fichero pasado en el parámetro fichero y carga los datos correspondientes.
    '   si hay errores los graba en el fichero invddmmaahhmm.log
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim NfLeer As Integer ' controlador de fichero
    Dim NfLog As Integer ' controlador del fichero log
    Dim LineaLeida As String ' la linea leida del fichero
    Dim LineaLog As String ' la linea a grabar en el log
    Dim Codigo As String
    Dim cantidad As String
    Dim i As Integer
    Dim pos As Integer
    On Error GoTo err_ProcesarFichero
    ProcesarFicheroInventario = True ' por defecto consideramos que todo va a ir bien
    NfLeer = FreeFile()
    Open Fichero For Input As #NfLeer '-- el fichero que vamos a leer
    NfLog = FreeFile()
    Open App.Path & "\INV" & Format(Now, "ddMMyyhhmmss") & ".log" For Output As #NfLog '-- fichero log
    While Not EOF(NfLeer)
        Line Input #NfLeer, LineaLeida
        i = i + 1
        lblInfInv.Caption = "Registro " & CStr(i)
        lblInfInv.Refresh
        DoEvents
        pos = InStr(1, LineaLeida, ",")
        If pos Then
            Codigo = Mid(LineaLeida, 1, pos - 1)
            cantidad = Right(LineaLeida, Len(LineaLeida) - pos)
            '-- Ahora hay que buscar el artículo asociado
            '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN y quitar de la cabecera
'            Sql = "select * from sartic where codigoea = '" & Codigo & "'"
            SQL = "select * from sarti3 where codigoea = '" & Codigo & "'"
            '----
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly
            If Not RS.EOF Then
                '-- Si que lo ha encontrado vamos a actualizarlo
                If Not ActualizarExistencia2(cantidad, RS!codArtic) Then
                    LineaLog = "Linea:" & CStr(i) & "|Error actualizando existencias " & _
                        " EAN: " & Codigo & _
                        " ARTICULO: " & RS!codArtic & _
                        " CANTIDAD: " & cantidad
                    Print #NfLog, LineaLog
                    ProcesarFicheroInventario = False
                End If
            Else
                LineaLog = "Linea:" & CStr(i) & "|No se ha encontrado artículo asociado " & _
                    " EAN: " & Codigo & _
                    " CANTIDAD: " & cantidad
                Print #NfLog, LineaLog
                ProcesarFicheroInventario = False
            End If
        Else
            LineaLog = "Linea:" & CStr(i) & "|No se reconoce el formato de la línea"
            Print #NfLog, LineaLog
            ProcesarFicheroInventario = False
        End If
    Wend
    Close NfLeer
    Close NfLog
    Exit Function
err_ProcesarFichero:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical
    ProcesarFicheroInventario = False
    Close NfLeer
    Close NfLog
    
End Function

