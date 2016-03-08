VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmCambPromo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio precio"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frmAlmCambPromo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCambPromo.frx":000C
      Height          =   4965
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8758
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Left            =   480
      TabIndex        =   5
      Top             =   5400
      Width           =   5775
   End
End
Attribute VB_Name = "frmAlmCambPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Dim NombreTabla As String
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

    If MsgBox("Actualizar datos promocion ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo ErrAceptar

    SQL = "Select * FROM tmpinformes  WHERE codusu = " & vUsu.codigo
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        If DBLet(RS!importeb3, "N") <> 0 Or DBLet(RS!importeb4, "N") <> 0 Then
            SQL = "UPDATE spromo set fechain1=" & DBSet(RS!fecha1, "F") & " , fechafi1 = " & DBSet(RS!fecha2, "F")
            SQL = SQL & ", precionu = " & DBSet(RS!importeb3, "N", "S") & ",  precion1 = " & DBSet(RS!importeb4, "N", "S")
            SQL = SQL & " WHERE codlista = " & RS!campo1 & " AND codartic = " & DBSet(RS!nombre1, "T")
            conn.Execute SQL
        End If
    
        RS.MoveNext
    Wend
    RS.Close
    
    cmdAceptar.visible = False
    Unload Me
    
ErrAceptar:
    If Err.Number <> 0 Then MuestraError Err.Number, "No se ha actualizado correctamente los precios", Err.Description
    If Not RS Is Nothing Then Set RS = Nothing
End Sub

Private Sub cmdAceptar_LostFocus()
    PonerModo 4
End Sub

Private Sub cmdCancelar_Click()
    If cmdAceptar.visible Then
        If MsgBox("Salir sin actualizar los cambio?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        cmdAceptar.visible = False
    End If
    Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (Not Data1.Recordset.BOF) And (Not Data1.Recordset.EOF) Then
       If gridCargado And Modo = 4 Then
            BotonModificar
        Else
            'txtAux(0).visible = False
           ' txtAux(1).visible = False
        End If
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
    
    cmdAceptar.visible = False
    
    Label1.Caption = "Cambio precio promociones"
    'NombreTabla = "slista"
    'If Me.Desde2 = "" Then
    '    Ordenacion = " ORDER BY codusu, numalbar,fechaalb,codartic"
    'Else
    '    Ordenacion = " ORDER BY codusu, numlinea"
    'End If
    Ordenacion = " ORDER BY nombre1"
    
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

     SQL = "SELECT codusu,codigo1,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4 FROM tmpinformes "
     SQL = SQL & " WHERE codusu = " & vUsu.codigo
     If Not enlaza Then SQL = SQL & " AND codigo1=-1"
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
        
    
    DataGrid1.Columns(0).visible = False
    DataGrid1.Columns(1).visible = False
    
    DataGrid1.Columns(2).Caption = "Cod.Art."
    DataGrid1.Columns(2).Width = 1700
    
    DataGrid1.Columns(3).Caption = "Desc. Articulo"
    DataGrid1.Columns(3).Width = 3200
       
    DataGrid1.Columns(4).Caption = "Precio"
    DataGrid1.Columns(4).Width = 1200
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = FormatoPrecio & " "
       
    
    DataGrid1.Columns(5).Caption = "Pre. caja"
    DataGrid1.Columns(5).Width = 1200
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(5).NumberFormat = FormatoPrecio & " "
        
    DataGrid1.Columns(6).Caption = "Nuevo"
    DataGrid1.Columns(6).Width = 1200
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(6).NumberFormat = FormatoPrecio & " "
            
    DataGrid1.Columns(7).Caption = "Nuevo caja"
    DataGrid1.Columns(7).Width = 1200
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(7).NumberFormat = FormatoPrecio & " "
        
    gridCargado = True
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
       
    Modo = Kmodo
    
    'MODIFICAR
    B = (Modo = 4)
   ' Me.cmdAceptar2.visible = B
    Me.cmdCancelar.visible = B
    
'    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub BotonModificar()
    PonerModo 4
'    CargaGrid True
    CargaTxtAux True, True
    PonerFoco txtAux(0)
    txtAux_GotFocus 0
    'DoEvents
End Sub



Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(0).Top = 290
        txtAux(1).Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            If DBLet(Data1.Recordset!importeb3, "N") = "0" Then
                txtAux(0).Text = " "
                txtAux(1).Text = " "
            Else
                txtAux(0).Text = DBLet(Data1.Recordset!importeb3, "N")
                txtAux(1).Text = DBLet(Data1.Recordset!importeb4, "N")
                If txtAux(0).Text = "0" Then txtAux(0).Text = ""
                If txtAux(1).Text = "0" Then txtAux(1).Text = ""
                
            End If
            txtAux(0).Locked = False
            txtAux(1).Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posición Top
        '-------------------------------
        txtAux(0).Top = alto
        txtAux(0).Height = DataGrid1.RowHeight
        txtAux(1).Top = alto
        txtAux(1).Height = DataGrid1.RowHeight
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux(0).Left = DataGrid1.Columns(6).Left + DataGrid1.Left + 10 'Nº Lotes
        txtAux(0).Width = DataGrid1.Columns(6).Width - 10
        txtAux(1).Left = DataGrid1.Columns(7).Left + DataGrid1.Left + 10 'Nº Lotes
        txtAux(1).Width = DataGrid1.Columns(7).Width - 10
        
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux(0).visible = visible
    txtAux(1).visible = visible
    PonerFoco txtAux(0)
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If cmdAceptar.visible Then Cancel = 1
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    On Error Resume Next
'    If txtAux.Text = "" Then Exit Sub
'    If ImporteFormateado(txtAux.Text) = 0 Then txtAux.Text = ""
'    If Trim(txtAux(Index).Text) = "" Then Exit Sub
'    txtAux(Index).SelStart = 0
'    txtAux(Index).SelLength = Len(txtAux(Index).Text)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim J As Long
Dim I As Integer
     Select Case KeyCode
        Case 33, 34
            'page down
            If KeyCode = 33 Then
                I = -DataGrid1.VisibleRows
            Else
                I = DataGrid1.VisibleRows
            End If
            J = Data1.Recordset.AbsolutePosition + I
            If J < 1 Then
                Data1.Recordset.MoveFirst
            Else
                If J > Data1.Recordset.RecordCount Then
                    Data1.Recordset.MoveLast
                Else
                    Data1.Recordset.Move I
                End If
            End If
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    txtAux_LostFocus 1
                    If DataGrid1.Row = 0 Then
                        BotonModificar
                    Else
                        DataGrid1.Row = DataGrid1.Row - 1
                    End If
                    CargaTxtAux True, True
                Else
                    If Not Data1.Recordset.BOF Then Data1.Recordset.MovePrevious
                End If
        Case 39
            If Index = 0 Then PonerFoco txtAux(1)
            
        Case 40 'Desplazamiento Flecha Hacia Abajo
            
                If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                    txtAux_LostFocus 1
                    If Data1.Recordset.AbsolutePosition = Data1.Recordset.RecordCount Then
                         DataGrid1.Row = DataGrid1.Row - 1
                    Else
                        DataGrid1.Row = DataGrid1.Row + 1
                    End If
                    CargaTxtAux True, True
                Else
                    Modo = 2
                   ' PonerFocoBtn Me.cmdAceptar2
                End If
    End Select
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    'Si el
    If KeyAscii = 13 Then
        If Index = 0 Then
            'pongo el foco en el 1
            PonerFoco txtAux(1)
        Else
            'guarda la linea
            TxtAux_KeyDown 0, 40, 0
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(Index).Text = ""
'    If Screen.ActiveControl.Name = "cmdAceptar" Then Exit Sub
    
    
    
'        If Desde2 = "" Then
'            MsgBox "El Nº de lote debe tener valor", vbInformation
'            PonerFoco txtAux
'        End If
    
    
    If Index = 1 Then GuardarLinea
   
End Sub


Private Sub GuardarLinea()

'    If Data1.Recordset.EOF = True Then PrimeraLin = True
     If ActualizarLinea2 Then

        NumRegElim = Data1.Recordset.AbsolutePosition
        
        CargaGrid True
        gridCargado = False
        If SituarDataPosicion(Data1, NumRegElim, "") Then
'            Data1.Recordset.MoveNext
        End If
        gridCargado = True
    End If
End Sub




Private Function ActualizarLinea2() As Boolean
Dim SQL As String
Dim I As Currency
Dim I2 As Currency
'    If Not DatosOkLinea Then Exit Function
    
    On Error GoTo ErrActLinea
    
'    Conn.BeginTrans
    
    If Trim(txtAux(0).Text) <> "" Then
        I = ImporteFormateado(txtAux(0).Text)
    Else
        I = 0
    End If
    If Trim(txtAux(1).Text) <> "" Then
        I2 = ImporteFormateado(txtAux(1).Text)
    Else
        I2 = 0
    End If
    
    
        
    SQL = "UPDATE tmpinformes SET importeb3=" & TransformaComasPuntos(CStr(I)) & " , importeb4=" & TransformaComasPuntos(CStr(I2))
    SQL = SQL & " WHERE codusu=" & vUsu.codigo & " AND codigo1=" & DBSet(Data1.Recordset!Codigo1, "N")

    conn.Execute SQL
    
    If Not cmdAceptar.visible Then cmdAceptar.visible = True



    ActualizarLinea2 = True
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
    ActualizarLinea2 = False
End Function

