VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmCambPrec 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   ClipControls    =   0   'False
   Icon            =   "frmAlmCambPrec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8040
      TabIndex        =   4
      Top             =   8400
      Visible         =   0   'False
      Width           =   1290
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
      Height          =   390
      Left            =   9720
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCambPrec.frx":000C
      Height          =   7965
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14049
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   8280
      Width           =   5415
   End
End
Attribute VB_Name = "frmAlmCambPrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public parSelSQL As String
Public vFecha As Date
Public Ventas As Boolean

Dim NombreTabla As String
Dim Ordenacion As String

Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.




Private Sub cmdAceptar_Click()
Dim Sql As String
Dim RS As ADODB.Recordset

    If MsgBox("Fecha: " & vFecha & "      ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo ErrAceptar


    
    Sql = MontaSQLCarga(True)
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    While Not RS.EOF

       If Ventas Then

            If Not IsNull(RS!tmpprecioac) Then
                If RS!tmpprecioac > 0 Then
                    Sql = "UPDATE slista SET precionu=" & DBSet(RS!tmpprecioac, "N")
                    Sql = Sql & ", fechanue = " & DBSet(vFecha, "F")
                    Sql = Sql & " WHERE " & " codartic =" & DBSet(RS!codArtic, "T")
                    Sql = Sql & " AND codlista=" & RS!codlista
                    conn.Execute Sql
                    
                    
                    'Nov 2017
                    If vParamAplic.ActualizaPrecioEspecial Then
                        'Como descuento vamos a poner el dto que tiene ahora
                        Sql = "UPDATE sprees SET precionu=" & DBSet(RS!tmpprecioac, "N")
                        Sql = Sql & ", fechanue = " & DBSet(vFecha, "F")
                        Sql = Sql & ", dtoespe1 = dtoespec" 'Como descuento vamos a poner el dto que tiene ahora
                        Sql = Sql & " WHERE " & " codartic =" & DBSet(RS!codArtic, "T")
                        conn.Execute Sql
                    End If
                End If
            End If
        
        Else
            
            If DBLet(RS!precioar, "N") > 0 Then
                Sql = "UPDATE slispr SET precionu=" & DBSet(RS!precioar, "N")
                Sql = Sql & ", fechanue = " & DBSet(vFecha, "F")
                Sql = Sql & " WHERE " & " codartic =" & DBSet(RS!codArtic, "T")
                Sql = Sql & " AND codprove=" & RS!NumOfert  'Numofert en la temporal es el proveedor
                conn.Execute Sql
            End If
    
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Me.cmdAceptar.visible = False
    Unload Me
    Exit Sub
    
ErrAceptar:
    If Not RS Is Nothing Then Set RS = Nothing
    MuestraError Err.Number, "No se ha actualizado correctamente los precios", Err.Description
End Sub

Private Sub cmdAceptar_LostFocus()
    PonerModo 4
End Sub

Private Sub cmdCancelar_Click()
    If cmdAceptar.visible Then
        If MsgBox("Salir sin actualizar los cambios?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        cmdAceptar.visible = False
    End If
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


    
    PulsadoSalir = False
    PrimeraVez = True
    DataGrid1.ClearFields
    
    
    
    
    cmdAceptar.visible = False
    
    Label1.Caption = "Cambio precio"
    Me.Caption = Label1.Caption
    NombreTabla = "Ventas"
    
    If Ventas Then
        Label1.Caption = Label1.Caption & "(Ventas)"
        NombreTabla = "slista"
        
    Else
        Label1.Caption = Label1.Caption & "(Compras)"
        NombreTabla = "tmpslipreu"
        
    End If
    Ordenacion = " ORDER BY codartic"
    
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
Dim Sql As String



     If Ventas Then
        'Sql = "SELECT slista.codlista,slista.codartic,nomartic,fechanue,precionu,tmpprecioac FROM"
        Sql = "SELECT slista.codlista,slista.codartic,nomartic,fechanue,precioac,tmpprecioac,"
        Sql = Sql & " round(if(precioac>0, if(tmpprecioac>0, (tmpprecioac-precioac)/precioac*100,0),0),2) Incremento  "
        Sql = Sql & " FROM slista,sartic WHERE slista.codartic=sartic.codartic "
        Sql = Sql & " AND " & Me.parSelSQL
        Sql = Sql & Ordenacion
     Else
        Sql = "Select codusu,codartic,nomartic,numofert,ampliaci, precioar"
        'Sql = Sql & ", round(if(numofert>0, if(precioar>0, (precioar-numofert)/numofert*100,0),0),2) Incremento  "
        Sql = Sql & "  FROM tmpslipreu WHERE " & Me.parSelSQL
        Sql = Sql & Ordenacion
     End If
     MontaSQLCarga = Sql
End Function


Private Sub CargaGrid(enlaza As Boolean)
Dim i As Byte
Dim Sql As String
    
    On Error GoTo ECarga

    gridCargado = False
    
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, Sql, PrimeraVez
    PrimeraVez = False
        

    DataGrid1.Columns(0).visible = False
    
    DataGrid1.Columns(1).Caption = "Cod.Art."
    DataGrid1.Columns(1).Width = 1920
    
    DataGrid1.Columns(2).Caption = "Desc. Articulo"
    DataGrid1.Columns(2).Width = 4500
       
       
    DataGrid1.Columns(3).Caption = "Fecha"
    DataGrid1.Columns(3).Width = 0
    DataGrid1.Columns(3).NumberFormat = "dd/mm/yyyy"
          
    DataGrid1.Columns(4).Caption = "Precio"
    DataGrid1.Columns(4).Width = 1500
    DataGrid1.Columns(4).Alignment = dbgRight
    If Ventas Then DataGrid1.Columns(4).NumberFormat = FormatoPrecio & " "
        
        
        
    
    DataGrid1.Columns(5).Caption = "Nuevo"
    DataGrid1.Columns(5).Width = 1500
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(5).NumberFormat = FormatoPrecio & " "
            
    If Ventas Then
        'Porcentaje incremento
        DataGrid1.Columns(6).Caption = "%Inc."
        DataGrid1.Columns(6).Width = 840
        DataGrid1.Columns(6).Alignment = dbgRight
        DataGrid1.Columns(6).NumberFormat = FormatoPorcen
    Else
        'DataGrid1.Columns(6).visible = False
    End If
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
        
            If Ventas Then
            
                If DBLet(Data1.Recordset!tmpprecioac, "N") = "0" Then
                    txtAux.Text = ""
                Else
                    txtAux.Text = DBLet(Data1.Recordset!tmpprecioac, "N")
                End If
            Else
                If DBLet(Data1.Recordset!precioar, "N") = "0" Then
                    txtAux.Text = ""
                Else
                    txtAux.Text = DBLet(Data1.Recordset!precioar, "N")
                End If
            End If
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
        txtAux.Left = DataGrid1.Columns(5).Left + DataGrid1.Left + 10 'Nº Lotes
        txtAux.Width = DataGrid1.Columns(5).Width - 10
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux.visible = visible
    
    PonerFoco txtAux
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If cmdAceptar.visible Then Cancel = 1
End Sub

Private Sub txtAux_GotFocus()
    On Error Resume Next
'    If txtAux.Text = "" Then Exit Sub
'    If ImporteFormateado(txtAux.Text) = 0 Then txtAux.Text = ""
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
Dim J As Long
Dim i As Integer
     Select Case KeyCode
        Case 33, 34
            'page down
            If KeyCode = 33 Then
                i = -DataGrid1.VisibleRows
            Else
                i = DataGrid1.VisibleRows
            End If
            J = Data1.Recordset.AbsolutePosition + i
            If J < 1 Then
                Data1.Recordset.MoveFirst
            Else
                If J > Data1.Recordset.RecordCount Then
                    Data1.Recordset.MoveLast
                Else
                    Data1.Recordset.Move i
                End If
            End If
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    txtAux_LostFocus
                    If DataGrid1.Row = 0 Then
                        BotonModificar
                    Else
                        DataGrid1.Row = DataGrid1.Row - 1
                    End If
'CargaTxtAux True, True
                Else
                    If Not Data1.Recordset.BOF Then Data1.Recordset.MovePrevious
                End If
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
            
                If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                    txtAux_LostFocus
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


Private Sub txtAux_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtAux_KeyDown 40, 0
    End If
End Sub

Private Sub txtAux_LostFocus()

    txtAux.Text = Trim(txtAux.Text)
    If Not PonerFormatoDecimal(txtAux, 2) Then txtAux.Text = ""
'    If Screen.ActiveControl.Name = "cmdAceptar" Then Exit Sub
    
    
    
'        If Desde2 = "" Then
'            MsgBox "El Nº de lote debe tener valor", vbInformation
'            PonerFoco txtAux
'        End If
    
    
        GuardarLinea
   
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
Dim Sql As String
Dim i As Currency
'    If Not DatosOkLinea Then Exit Function
    
    On Error GoTo ErrActLinea
    
'    Conn.BeginTrans

    If Trim(txtAux.Text) <> "" Then
        i = ImporteFormateado(txtAux.Text)
        
        
        If Ventas Then
        
            Sql = "UPDATE " & NombreTabla & " SET tmpprecioac=" & TransformaComasPuntos(CStr(i))
            Sql = Sql & " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND codlista=" & DBSet(Data1.Recordset!codlista, "N")
        Else
            'Compras
            Sql = "UPDATE " & NombreTabla & " SET precioar=" & TransformaComasPuntos(CStr(i))
            Sql = Sql & " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND codusu=" & vUsu.Codigo
            
        End If
        
        conn.Execute Sql
            
        If Not cmdAceptar.visible Then cmdAceptar.visible = True
   
    End If
     
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

