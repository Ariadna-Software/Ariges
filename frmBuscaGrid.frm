VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscaGrid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmBuscaGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBusqueda 
      Caption         =   "Búsqueda"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   5520
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   435
         Index           =   1
         Left            =   6720
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option2"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         Caption         =   "Label5"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1080
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   5100
      Width           =   1455
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBuscaGrid.frx":1CFA
      Height          =   3075
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5424
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         AllowFocus      =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Leyendo datos servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "Búsqueda"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TITULO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cargando datos ...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   3675
   End
End
Attribute VB_Name = "frmBuscaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Selecionado(CadenaDevuelta As String)

'Variables publicas para montar datos
Public vTabla As String
Public vCampos As String 'columnas en la tabla.Empipados
Public vselElem As Integer
Public vTitulo As String
Public vSQL As String
Public vBusqueda As String
Public vCargaFrame As Boolean 'Si se cargara el frame de busqueda o no

'Dentro de campos vendra cada grupo separado por ·
'Y cada grupo sera Desc|Tabla|Tipo|Porcentaje de ancho
Public vDevuelve As String 'Empipados los campos que devuelve

'Cadena de conexión con la BD a la que hay que buscar
Public vConexionGrid As Integer

'Variables privadas
Dim PrimeraVez As Boolean
Dim SQL As String

'Las redimensionaremos
Dim TotalArray As Integer
Dim Cabeceras() As String 'Descripcion de mensajes
Dim CabTablas() As String 'Nombres de las tablas
Dim CabColumnas() As String 'Nombres de las columnas a mostrar
Dim CabAncho() As Single 'Ancho de la columna
Dim TipoCampo() As String 'Tipo campo a mostrar
Dim FormatoCampo() As String 'Formato del campo


Private Busca As Boolean
Private DbClick As Boolean



Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim columna As String
Dim J As Byte
On Error Resume Next

    If Busca Then
        Busca = False
        Exit Sub
    End If
    
    DbClick = True
    If Adodc1.Recordset.BOF Then
        If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
    End If
    
    If Adodc1.Recordset.RecordCount > 0 Then
        columna = CabColumnas(vselElem)
        
         '---- Añade: Laura 28/04/2005
        J = InStr(1, columna, " as ") 'si columna tiene if o case renombramos ( as nomcolum )
        If J > 0 Then
            columna = Mid(columna, J + 4)
            columna = Trim(columna)
        End If
        
        '---- Modifica: LAura 2005 ------------------------
        '---- se añade el formato del campo
        If FormatoCampo(vselElem) <> "" Then
            Text1.Text = Format(Adodc1.Recordset.Fields(CabColumnas(vselElem)), FormatoCampo(vselElem))
        Else
            Text1.Text = DBLet(Adodc1.Recordset.Fields(CabColumnas(vselElem)))
        End If
    End If
End Sub


Private Sub cmdBuscar_Click()
Dim cadB As String

    Screen.MousePointer = vbHourglass
    Me.Refresh
    vBusqueda = ObtenerBusqueda(Me, False)
'    cadB = ObtenerBusqueda(Me)
'    If cadB <> "" Then
'        If vSQL <> "" Then vSQL = vSQL & " AND "
'        vSQL = vSQL & cadB
'    End If
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdRegresar_Click()
Dim vDes As String
Dim i, J As Integer
Dim V As String

If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub

    i = 0
    vDes = ""
    Do
    J = i + 1
    i = InStr(J, vDevuelve, "|")
    If i > 0 Then
        V = Mid(vDevuelve, J, i - J)
        If V <> "" Then
            If IsNumeric(V) Then
                If Val(V) <= TotalArray Then vDes = vDes & Adodc1.Recordset(CabColumnas(Val(V))) & "|"
            End If
'            If IsDate(V) Then
'                If Val(V) <= TotalArray Then vDes = vDes & adodc1.Recordset(CabColumnas(Val(V))) & "|"
'            End If
        End If
    End If
Loop Until i = 0
RaiseEvent Selecionado(vDes)
Unload Me
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    cmdRegresar_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String

If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
If vselElem = ColIndex Then Exit Sub
cad = "¿Desea reordenar por el concepto " & DataGrid1.Columns(ColIndex).Caption & "?"
If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
If ColIndex <= TotalArray Then
    Me.Refresh
    Screen.MousePointer = vbHourglass
    vselElem = ColIndex
    CargaGrid
    Screen.MousePointer = vbDefault
    Else
    MsgBox "Error cargando tabla. Imposible ordenacion", vbCritical
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    PonerFoco Text1
End Sub

Private Sub Form_Activate()
Dim OK As Boolean
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        Screen.MousePointer = vbHourglass
        OK = ObtenerTamanyosArray
        If OK Then OK = SeparaCampos
        If Not OK Then
            'Error en SQL
            'Salimos
            Unload Me
            Exit Sub
        End If
    '    If vBuscaPrevia And vCargaFrame = False Then
        If vCargaFrame = False Then
            CargaGrid
        Else
            CargaFrame
        End If
        Label4.visible = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Label4.visible = True
    Me.DataGrid1.Enabled = False
    PrimeraVez = True
    Label1.Caption = vTitulo
    DbClick = True
    Adodc1.password = vUsu.Passwd
    
    'If Not vBuscaPrevia And vCargaFrame Then
    If vCargaFrame Then
        'Poner Visible el Frame y no Visible el Grid
        ConfiguraForm (1)
    End If
End Sub


Private Function SeparaCampos() As Boolean
Dim cad As String
Dim Grupo As String
Dim i As Integer
Dim J As Integer
Dim C As Integer 'Contrador dentro del array

    SeparaCampos = False
    i = 0
    C = 0
    Do
        J = i + 1
        i = InStr(J, vCampos, "·")
        If i > 0 Then
            Grupo = Mid(vCampos, J, i - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until i = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim i As Integer
Dim J As Integer
Dim cad As String

    J = 0
    cad = ""
    
    'Cabeceras
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    Cabeceras(Contador) = cad
    
    'TAblas BD
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    CabTablas(Contador) = cad
    
    'Columnas Tablas
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    CabColumnas(Contador) = cad
    
    'El tipo
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    TipoCampo(Contador) = cad
    
    'El formato
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    FormatoCampo(Contador) = cad
    
    'Por ultimo
    'ANCHO
    If Grupo = "" Then Grupo = 0
    CabAncho(Contador) = Grupo
End Sub


Private Function ObtenerTamanyosArray() As Boolean
Dim i As Integer
Dim J As Integer
Dim Grupo As String

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        i = J + 1
        J = InStr(i, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    If TotalArray < 0 Then Exit Function
    'Las redimensionaremos
    ReDim Cabeceras(TotalArray)
    ReDim CabTablas(TotalArray)
    ReDim CabColumnas(TotalArray)
    ReDim CabAncho(TotalArray)
    ReDim TipoCampo(TotalArray)
    ReDim FormatoCampo(TotalArray)
    ObtenerTamanyosArray = True
End Function


Private Sub CargaGrid()
Dim cad As String, Orden As String
Dim i As Integer
Dim anc As Single
On Error GoTo ECargaGrid

    'On Error GoTo ECargaGRid '##QUITAR
    'Generamos SQL
    cad = ""
    For i = 0 To TotalArray
        If cad <> "" Then cad = cad & ", "
        If (InStr(CabColumnas(i), "if") > 0) Or (InStr(CabColumnas(i), "case") > 0) Then
            cad = cad & CabColumnas(i)
        Else
            'Si no he indicado la tabla, NO la pongo, ni pongo el punto(.)
            If CabTablas(i) <> "" Then cad = cad & CabTablas(i) & "."
            cad = cad & CabColumnas(i)
        End If
    Next i
    cad = "SELECT " & cad & " FROM " & vTabla
    If vSQL <> "" Then
        cad = cad & " WHERE " & vSQL
        If vBusqueda <> "" Then cad = cad & " AND " & vBusqueda
    ElseIf vBusqueda <> "" Then cad = cad & " WHERE " & vBusqueda
    End If
   
   
   
   'TRUCO DEL ALMENDRUCO
   'SOLO para facturas. Si se necesita en algun sitio mas, habra que programarlo
   'Como sabemos que es fecturas
       
    If vTitulo = "Facturas" Then
        If Mid(vCampos, 1, 29) = "Tipo Fac.|scafac|codtipom|T||" Then
            If Mid(vTabla, 1, 28) = "scafac INNER JOIN scafac1 ON" Then
                'SOLO entonces, ponemos group by 1,2,3
                cad = cad & " GROUP BY 1,2,3"
            End If
        End If
    End If
   
       
       
       
   
   
    '---- Modifica: Laura 08/06/2005  ----------------------
    'antes:
    ' cad = cad & " ORDER BY " & CabColumnas(vselElem)
    Orden = CabColumnas(vselElem)
    i = InStr(1, Orden, " as ")
    If i > 0 Then Orden = Mid(Orden, i + 4)
    cad = cad & " ORDER BY " & Orden
    '--------------------------------------------------------


    Select Case vConexionGrid
        Case 1  'Conexión a BDatos: Ariges
                Adodc1.ConnectionString = conn
        Case 2  'Conexión a BDatos: Conta
                Adodc1.ConnectionString = ConnConta
    End Select

    Adodc1.RecordSource = cad
    Adodc1.Refresh

    'If adodc1.Recordset.RecordCount > 0 Then
    If (vCargaFrame = False) Or (vCargaFrame = True And Adodc1.Recordset.RecordCount > 0) Then
        DataGrid1.AllowRowSizing = False
        DataGrid1.visible = True
        'Cargamos el grid
        anc = DataGrid1.Width - 640
        
        For i = 0 To TotalArray
            DataGrid1.Columns(i).Caption = Cabeceras(i)
            If FormatoCampo(i) <> "" Then
                DataGrid1.Columns(i).NumberFormat = FormatoCampo(i)
                If InStr(1, FormatoCampo(i), ".") Then DataGrid1.Columns(i).Alignment = dbgRight
            End If
            If CabAncho(i) = 0 Then
                DataGrid1.Columns(i).visible = False
            Else
                DataGrid1.Columns(i).Width = anc * (CabAncho(i) / 100)
            End If
        Next i
    
         'Habilitamos el text1 para que escriban
        DataGrid1.Enabled = True
        Text1.Enabled = True
        Text1.visible = True
    
        If Not Adodc1.Recordset.EOF Then
            'Le ponemos el 1er registro
            cad = CabColumnas(vselElem)
            
             '---- Añade: LAura 08/06/2005
            'Si hay if/case en nombre columna cogemos el renombrado: if(colum=x,,) as colum
            i = InStr(1, cad, " as ")
            If i > 0 Then
                cad = Mid(cad, i + 4)
                cad = Trim(cad)
            End If
            
            '---- Modifica: Laura 2005 --------------
            '---- se añade el formato del campo
            If FormatoCampo(vselElem) <> "" Then
                Text1.Text = Format(Adodc1.Recordset(cad), FormatoCampo(vselElem))
            Else
               ' Text1.Text = DBLet(Adodc1.Recordset(cad))
            End If
            PonerFoco Text1
        Else
            PonerFocoBtn cmdSalir(0)
'            cmdSalir(0).SetFocus
        End If
        ConfiguraForm (0)
    Else
    '    txtBusqueda.SetFocus
    '    If vCargaFrame Then
            frameBusqueda.visible = True
            ConfiguraForm (1)
    '    Else
    '        Unload Me
    '    End If
        MsgBox "No hay ningún registro en la tabla " & vTabla
    End If
    Exit Sub
ECargaGrid:
    MuestraError Err.Number, "Carga grid." & vbCrLf & Err.Description
End Sub


Private Sub CargaFrame()
Dim cad As String
Dim i As Integer
    
    frameBusqueda.visible = True
    For i = 0 To TotalArray
        Option1(i).Caption = Cabeceras(i)
    Next i
    i = 0
    Option1(i).Value = True
    lblBusqueda.Caption = Cabeceras(i)
    txtBusqueda.Text = ""
    txtBusqueda.SetFocus
    
    CargaTagTxt (vselElem)
'    cad = Cabeceras(vselElem) & "|" & TipoCampo(vselElem) & "|N|||"
'    cad = cad & vTabla & "|" & CabTablas(vselElem) & "|||"
'    txtBusqueda.Tag = cad
    
End Sub


Private Sub ConfiguraForm(ByVal tamanyo As Integer)
'Ajustar tamaño del form
    Select Case tamanyo
    Case 0  'Tamaño normal
        DataGrid1.visible = True
        Me.Height = 6225
        Me.DataGrid1.Height = 3075
        If Me.frameBusqueda.visible Then Me.Top = 2600
'        Me.Top = 1400
        Me.cmdSalir(0).visible = True
        Me.cmdRegresar.visible = True
        Me.Label2.visible = True
        frameBusqueda.visible = False
    Case 1 'Tamaño pequeño sin Grid (Solo Frame)
        DataGrid1.visible = False
        Me.Height = 3700
        Me.DataGrid1.Height = 100
'        Me.Top = 2700
        Me.cmdSalir(0).visible = False
        Me.cmdRegresar.visible = False
        Me.Label2.visible = False
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DataGrid1.Enabled = False
End Sub


Private Sub Option1_Click(Index As Integer)
    lblBusqueda.Caption = Cabeceras(Index)
    vselElem = Index
    CargaTagTxt (vselElem)
End Sub

Private Sub Text1_Change()
Dim SQLDBGRID As String
Dim pTexto As String

    If Trim(Text1) = "" Then
        If Not Me.Adodc1.Recordset.EOF Then
            If Not Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
        End If
        Exit Sub
    End If
    If DbClick Then
        DbClick = False
        Exit Sub
    End If
    Busca = True
    SQLDBGRID = CabColumnas(vselElem)
    Select Case TipoCampo(vselElem)
        Case "N"
            If Not IsNumeric(Text1.Text) Then
                If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
                Exit Sub
            End If
            If Len(Trim(Text1)) > Len(FormatoCampo(vselElem)) Then
                SQLDBGRID = SQLDBGRID & " >= " & Val(Mid(Trim(Text1), 1, Len(FormatoCampo(vselElem))))
            Else
                SQLDBGRID = SQLDBGRID & " >= " & Val(Trim(Text1))
            End If
        Case "T"
            pTexto = Trim(Text1.Text)
            If Len(pTexto) = 1 Then
                If pTexto = "*" Then
                    pTexto = ""
                    Exit Sub
                End If
            End If
            
'            SQLDBGRID = SQLDBGRID & " >= '" & Trim(text1) & "'"
            SQLDBGRID = SQLDBGRID & " like '" & pTexto & "*'"
        Case "F"
            Exit Sub
    End Select
    Screen.MousePointer = vbHourglass
    
    Adodc1.Recordset.Find SQLDBGRID, , adSearchForward, 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdRegresar_Click 'ENTER
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdBuscar_Click 'ENTER
End Sub


Private Sub CargaTagTxt(ByVal vselElem As Integer)
Dim cad As String
Dim EsNulo As String

    If vselElem = 0 Then
        EsNulo = "N"
    Else
        EsNulo = "S"
    End If
    cad = Cabeceras(vselElem) & "|" & TipoCampo(vselElem) & "|" & EsNulo & "|||"
    cad = cad & vTabla & "|" & CabColumnas(vselElem) & "|" & FormatoCampo(vselElem) & "||"
    txtBusqueda.Tag = cad
End Sub
