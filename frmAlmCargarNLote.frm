VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmCargarNLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Introducir Nº de lotes"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   ClipControls    =   0   'False
   Icon            =   "frmAlmCargarNLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7875
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCargarNLote.frx":000C
      Height          =   3885
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   0
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
End
Attribute VB_Name = "frmAlmCargarNLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CargarNumLotes()
Public parSelSQL As String

Public Desde2 As String
    '"".- Proveedores
    '
    'codtipom|codtipoa|

Dim NombreTabla As String
Dim Ordenacion As String

Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.




Private Sub cmdAceptar_Click()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim DentroTRANS As Boolean

    On Error GoTo ErrAceptar
    DentroTRANS = False
    
    'comprobar que todos los num serie tienen valor
    If Desde2 = "" Then
        SQL = "SELECT COUNT(*) FROM " & NombreTabla
        SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND " & Me.parSelSQL & " AND (isnull(numlotes) or trim(numlotes)="""" )"
        
        If RegistrosAListar(SQL) > 0 Then
            MsgBox "Hay algún articulo para el que no se ha introducido Nº de Lote.", vbInformation
            Exit Sub
        End If
    End If
    
    'Actualizar la tabla de albaranes de proveedor con el nº de lote
    SQL = MontaSQLCarga(True)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not RS.EOF
    
        If Desde2 = "" Then
            '++++++++++++++++++++++++++++++++++++++++
            '
            '   PROVEEDORES
            '
            '++++++++++++++++++++++++++++++++++++++++
    
    
    
                SQL = "UPDATE slialp " & " SET numlotes=" & DBSet(RS!numlotes, "T")
                SQL = SQL & " WHERE " & " numalbar=" & DBSet(RS!NumAlbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F")
                SQL = SQL & " AND codprove=" & RS!codProve & " AND numlinea=" & RS!numlinea
        
                conn.Execute SQL
                
                'Insertamos en la tabla slotes
                SQL = "SELECT COUNT(*) FROM slotes WHERE "
                SQL = SQL & " codartic=" & DBSet(RS!codArtic, "T") & " AND numlotes=" & DBSet(RS!numlotes, "T") & " AND fecentra=" & DBSet(RS!FechaAlb, "F")
                If RegistrosAListar(SQL) > 0 Then
                    'si ya existe la linea aumentamos la cantidad entrada
                    SQL = "UPDATE slotes SET canentra=canentra + " & DBSet(RS!Cantidad, "N")
                    SQL = SQL & " WHERE " & " codartic=" & DBSet(RS!codArtic, "T") & " AND numlotes=" & DBSet(RS!numlotes, "T") & " AND fecentra=" & DBSet(RS!FechaAlb, "F")
                    conn.Execute SQL
                Else
                    SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
                    SQL = SQL & DBSet(RS!codArtic, "T") & ", " & DBSet(RS!numlotes, "T") & ", "
                    'fecha entrada, cantidad entrada y cantidad asignada
                    SQL = SQL & DBSet(RS!FechaAlb, "F") & "," & DBSet(RS!Cantidad, "N") & ",0)"
                    conn.Execute SQL
                End If
        
        
                
        Else
        
                '***************************
                ' LOTES en facturas clientes
                '***************************
                SQL = "UPDATE slifac " & " SET numlote=" & DBSet(RS!numlotes, "T")
                SQL = SQL & " " & Me.Desde2
                SQL = SQL & " AND numlinea=" & RS!numlinea
                conn.Execute SQL
            
            
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    Unload Me
    Exit Sub
    
ErrAceptar:
    If Not RS Is Nothing Then Set RS = Nothing
    MuestraError Err.Number, "No se ha actualizado correctamente los nº de lote en la tabla slialp.", Err.Description
End Sub

Private Sub cmdAceptar_LostFocus()
    PonerModo 4
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (Not Data1.Recordset.BOF) And (Not Data1.Recordset.EOF) Then
       If gridCargado And Modo = 4 Then BotonModificar
    Else
        Data1.Recordset.MoveLast
    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtAux
    End If
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

'    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 4 'Modificar
'        .Buttons(2).Image = 21 'Cargar Nº Series
''        .Buttons(4).Image = 15 'Salir
'    End With
    
    PulsadoSalir = False
    PrimeraVez = True
    DataGrid1.ClearFields
    
    
    If Me.Desde2 = "" Then
        Label1.Caption = "Proveedores"
    Else
        Label1.Caption = "Fra. cliente"
    End If
    NombreTabla = "tmpnlotes"
    If Me.Desde2 = "" Then
        Ordenacion = " ORDER BY codusu, numalbar,fechaalb,codartic"
    Else
        Ordenacion = " ORDER BY codusu, numlinea"
    End If
    
'    PonerModo 4
    CargaGrid True
    BotonModificar
    
    Screen.MousePointer = vbDefault
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

     SQL = "SELECT * FROM " & NombreTabla
     SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND " & Me.parSelSQL
     SQL = SQL & Ordenacion

     MontaSQLCarga = SQL
End Function


Private Sub CargaGrid(enlaza As Boolean)
Dim I As Byte
Dim SQL As String
    
    On Error GoTo ECarga

    gridCargado = False
    
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    PrimeraVez = False
        
    DataGrid1.Columns(0).visible = False 'codusu
    DataGrid1.Columns(1).visible = False 'numalbar
    DataGrid1.Columns(2).visible = False 'fecha albar
    DataGrid1.Columns(3).visible = False 'codprove
    
    DataGrid1.Columns(4).Caption = "NºLin."
    DataGrid1.Columns(4).Width = 600
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = "#" & " "
    
    DataGrid1.Columns(5).Caption = "Cod.Art."
    DataGrid1.Columns(5).Width = 1700
    
    DataGrid1.Columns(6).visible = False 'codalmac
    
    DataGrid1.Columns(7).Caption = "Desc. Articulo"
    DataGrid1.Columns(7).Width = 3200
       
    DataGrid1.Columns(8).Caption = "Cantid."
    DataGrid1.Columns(8).Width = 800
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(8).NumberFormat = FormatoCantidad & " "
       
    DataGrid1.Columns(9).Caption = "Nº Lote"
    DataGrid1.Columns(9).Width = 1500
        
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
        DataGrid1.Columns(I).Locked = True
    Next I
    
    gridCargado = True
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
       
    Modo = Kmodo
    
    'MODIFICAR
    B = (Modo = 4)
    Me.cmdAceptar.visible = B
    Me.cmdCancelar.visible = B
    
'    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub BotonModificar()
    PonerModo 4
'    CargaGrid True
    CargaTxtAux True, True
    PonerFoco txtAux
End Sub



Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            txtAux.Text = DBLet(Data1.Recordset!numlotes, "T")
            txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(9).Left + DataGrid1.Left + 10 'Nº Lotes
        txtAux.Width = DataGrid1.Columns(9).Width - 10
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux.visible = visible
    
    PonerFoco txtAux
End Sub





Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    txtAux_LostFocus
                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
                End If
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                    txtAux_LostFocus
                    DataGrid1.Row = DataGrid1.Row + 1
'                    CargaTxtAux True, True
                Else
                    Modo = 2
                    PonerFocoBtn Me.cmdAceptar
                End If
    End Select
End Sub


Private Sub txtAux_LostFocus()

    txtAux.Text = Trim(txtAux.Text)
'    If Screen.ActiveControl.Name = "cmdAceptar" Then Exit Sub
    
    
    
        If Desde2 = "" Then
            MsgBox "El Nº de lote debe tener valor", vbInformation
            PonerFoco txtAux
        End If
    
    
        GuardarLinea
   
End Sub


Private Sub GuardarLinea()

'    If Data1.Recordset.EOF = True Then PrimeraLin = True
     If ActualizarLinea Then

        NumRegElim = Data1.Recordset.AbsolutePosition
        
        CargaGrid True
        gridCargado = False
        If SituarDataPosicion(Data1, NumRegElim, "") Then
'            Data1.Recordset.MoveNext
        End If
        gridCargado = True
    End If
End Sub




Private Function ActualizarLinea() As Boolean
Dim SQL As String

'    If Not DatosOkLinea Then Exit Function
    
    On Error GoTo ErrActLinea
    
'    Conn.BeginTrans

  '  If Trim(txtAux.Text) <> "" Then
        SQL = "UPDATE " & NombreTabla & " SET numlotes=" & DBSet(txtAux.Text, "T")
        SQL = SQL & " WHERE codusu=" & Data1.Recordset!CodUsu & " AND numalbar=" & DBSet(Data1.Recordset!NumAlbar, "T") & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!codProve & " AND numlinea=" & Data1.Recordset!numlinea

        conn.Execute SQL
   ' End If
    
     
    ActualizarLinea = True
    Exit Function

ErrActLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando linea.", Err.Description
    End If
'    If b Then
'        Conn.CommitTrans
'    Else
'        Conn.RollbackTrans
'    End If
    ActualizarLinea = False
End Function

