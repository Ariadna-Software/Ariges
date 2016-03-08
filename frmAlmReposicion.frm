VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmReposicion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada reposicion de almacén"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13080
   ClipControls    =   0   'False
   Icon            =   "frmAlmReposicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "existencia"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5595
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11715
      TabIndex        =   1
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11715
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmReposicion.frx":000C
      Height          =   5355
      Left            =   120
      TabIndex        =   4
      Top             =   225
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   9446
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
      TabIndex        =   8
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
      Left            =   240
      TabIndex        =   5
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmAlmReposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean




Private Sub cmdAceptar_Click()
Dim cad As String
Dim miSQL As String
Dim cT As CTiposMov
Dim LineaInicioComponentes As Integer
Dim Cantidad As Currency

    On Error GoTo Error1
    
    
    davidCodtipom = PonerTrabajadorConectado(cad)
    If davidCodtipom = "" Then
        MsgBox "Error trabajador conectado(1)", vbExclamation
        Exit Sub
    End If
    
    cad = DevuelveDesdeBD(conAri, "count(*)", "straspaso", "propuesta>0 and 1", 1)
    If Val(cad) = 0 Then
        MsgBox "No se puede generar ningun dato. Cantidad =0", vbExclamation
        Exit Sub
    End If
        
        
    
    'Almacen que no sea el ppal
    'Si acepta meteremos en entrada en trasapaso de almacen
    If MsgBox("¿Desea insertar en el traspaso de almacenes?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
 
    
    
    Screen.MousePointer = vbHourglass
    
    
    'COMPONENTES
    Set miRsAux = New ADODB.Recordset
    cad = "select straspaso.codartic,sum(propuesta) cuantos from straspaso,sartic where straspaso.codartic=sartic.codartic and conjunto=1 group by 1 having sum(propuesta) > 0"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    While Not miRsAux.EOF
        'Siempre desde el almacen ppal
        miSQL = miSQL & miRsAux!codArtic & "@@" & miRsAux!Cuantos & "··"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Para cada ariculo con componentes, veremos cuales de esos compoente
    
    If miSQL <> "" Then
    
        'Sobre temporal
        'tmpsliped(codusu,numpedcl,numlinea,codartic,cantidad)
        conn.Execute "Delete from tmpsliped where codusu = " & vUsu.codigo
        Espera 0.25
    
        
        LineaInicioComponentes = 0
    
        While miSQL <> ""
            NumRegElim = InStr(1, miSQL, "··")
            If NumRegElim = 0 Then
                miSQL = ""
            Else
                cad = Mid(miSQL, 1, NumRegElim - 1)
                miSQL = Mid(miSQL, NumRegElim + 2)
                    
                NumRegElim = InStr(1, cad, "@@")
                If NumRegElim > 0 Then
                    If Len(Mid(cad, NumRegElim + 2)) > 0 Then
                        Cantidad = Mid(cad, NumRegElim + 2)
                        cad = Mid(cad, 1, NumRegElim - 1)
                        pPdfRpt = ""
                        cad = "Select sarti1.*," & TransformaComasPuntos(CStr(Cantidad)) & " CantOrigen from sarti1 where codartic=" & DBSet(cad, "T")
                        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not miRsAux.EOF
                            ''tmpsliped(codusu,numpedcl,numlinea,codartic,cantidad)
                            LineaInicioComponentes = LineaInicioComponentes + 1
                            Cantidad = miRsAux!Cantidad * miRsAux!CantOrigen
                            pPdfRpt = pPdfRpt & ", (" & vUsu.codigo & ",1," & LineaInicioComponentes & ","
                            pPdfRpt = pPdfRpt & DBSet(miRsAux!codarti1, "T") & "," & DBSet(Cantidad, "N") & ")"
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                        If pPdfRpt <> "" Then
                            pPdfRpt = Mid(pPdfRpt, 2)
                            cad = "INSERT INTO tmpsliped(codusu,numpedcl,numlinea,codartic,cantidad) VALUES " & pPdfRpt
                            conn.Execute cad
                        End If
                    End If
                End If
            End If
        Wend
        
        
        'Ya hemos insertado TODOS los componentes con sus cantidades en tmpsliped
        'AHora vamos a meterlos en straspaso
        If LineaInicioComponentes > 0 Then
            Espera 0.1
            
            'select codartic,sum(cantidad) from tmpsliped group by 1
            cad = "Select max(secuencia) from straspaso"
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUE SER EOF
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
            
            'straspaso(codusu,secuencia,codprove,codartic,propuesta)
            '@rownum:=@rownum+1 AS rownum      (SELECT @rownum:=0) r
            cad = vUsu.codigo & ",@rownum:=@rownum+1 AS rownum ,99999999,codartic,sum(cantidad)  "
            cad = cad & " from tmpsliped,(SELECT @rownum:=" & NumRegElim & ") r WHERE codusu = " & vUsu.codigo & " group by  codartic"
            cad = "INSERT INTO straspaso(codusu,secuencia,codprove,codartic,propuesta) SELECT " & cad
            conn.Execute cad
        End If
    End If
    Set miRsAux = Nothing
    
    
    
    
    
    
    
    
    
    
 
    Set cT = New CTiposMov
    If cT.Leer("TRA") Then
        cT.ConseguirContador cT.TipoMovimiento
        cT.IncrementarContador cT.TipoMovimiento
        'Cabecera
        ' scatra(codtrasp,fechatra,almaorig,almadest,codtraba,situacio,observa1)
        'En tmpinfomres llevamos, fecha1,fecha2,almacen(porcen2)
        
        
        
        cad = "Reposicion,generado por " & vUsu.Nombre & " el " & Format(Now, "dd/mm/yyyy hh:mm") & vbCrLf
        cad = cad & "Fechas ventas: " & Format(Data1.Recordset!fecha1, "dd/mm/yyyy")
        If Not IsNull(Data1.Recordset!fecha2) Then cad = cad & " - " & Format(Data1.Recordset!fecha2, "dd/mm/yyyy") & vbCrLf
        
        
        miSQL = cT.Contador & "," & DBSet(Now, "F") & ",1," & Val(CStr(Data1.Recordset!almdestino)) & "," & davidCodtipom
        miSQL = miSQL & ",0," & DBSet(cad, "T") & ")"
        
'

        
        
        miSQL = "INSERT INTO scatra(codtrasp,fechatra,almaorig,almadest,codtraba,situacio,observa1) VALUES (" & miSQL
        
        
        If Ejecutar(miSQL, False) Then
        
            'codtrasp , numlinea, codArtic, Cantidad, observa2    minimo,stock,reserva,propuest
            'Mayo 2014. NO quieren la linea con estos datos
            'miSQL = "concat(right(concat('00000',codprove),6),'  ST:',stock,'  Min:',minimo,'   Ped:',reserva)"
            
            miSQL = "if(codprove=99999999,'COMPONENTES',null)"
            
            miSQL = "Select " & cT.Contador & ",@rownum:=@rownum+1 AS rownum,codArtic,propuesta ," & miSQL
            
            
            miSQL = "INSERT INTO slitra(codtrasp , numlinea, codArtic, Cantidad, observa2) " & miSQL
            miSQL = miSQL & " from straspaso ,(SELECT @rownum:=0) r  where"
            'miSql = miSql & " CodUsu = " & vUsu.codigo & "  And importe1 >0"
            miSQL = miSQL & "  propuesta >0"
            miSQL = miSQL & " ORDER BY codprove,codartic"  'codprove,nomartic
            If Not Ejecutar(miSQL, False) Then
        
                miSQL = "DELETE FROM scatra where codtrasp = " & cT.Contador
                Ejecutar miSQL, False
                cT.DevolverContador cT.TipoMovimiento, cT.Contador
                miSQL = ""
            Else
                miSQL = "Traspaso generado correctamente: " & vbCrLf & "Nº: " & cT.Contador & vbCrLf & "Fecha: " & Format(Now, "dd/mm/yyyy")
                MsgBox miSQL, vbExclamation
                
                'Borramos la tabla
                conn.Execute "DELETE FROM straspaso"
                
            End If
        
        
        End If
        
    Else
        miSQL = ""
    End If
    Set cT = Nothing
        


        
    If miSQL <> "" Then
         Me.Tag = 0
         Unload Me
    End If


   
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Set miRsAux = Nothing
End Sub

Private Sub cmdCancelar_Click()
   
    
    'If MsgBox("Los datos serán guardados." & vbCrLf & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    Me.Tag = 0
    Unload Me
  
   
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
       'txtAux.SelStart = Len(Me.txtAux.Text)
       txtAux.SelStart = 0
       txtAux.SelLength = Len(Me.txtAux.Text)
       txtAux.Refresh
       
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
         PrimeraVez = False
         BotonModificar
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
 
    
    Me.Tag = 1 'NO se puede cerrar mas que de boton
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True

    PonerModo 4
    CargaGrid
    PrimeraVez = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()
Dim i As Byte
Dim SQL As String
On Error GoTo ECarga

    gridCargado = False
    'codusu,secuencia,codprove,codartic,articulo,proveedor"
    ',minimo,stock,reserva,propuesta,almdestino,fecha1,fecha2"
    SQL = "select secuencia,proveedor,codartic,articulo,minimo,stock,reserva,propuesta,fecha1,fecha2,almdestino"
    SQL = SQL & " from straspaso "   '"WHERE codusu = " & vUsu.codigo
    SQL = SQL & " ORDER BY codprove,codartic"
    
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    PrimeraVez = False
        
    'Cod. Articulo
    DataGrid1.Columns(0).Caption = "CodInte"
    DataGrid1.Columns(0).Width = 0
    
    DataGrid1.Columns(1).Caption = "Proveedor"
    DataGrid1.Columns(1).Width = 3000
    
    DataGrid1.Columns(2).Caption = "Codartic"
    DataGrid1.Columns(2).Width = 1500
    
        
    DataGrid1.Columns(3).Caption = "Descripcion"
    DataGrid1.Columns(3).Width = 3950
        
        
    'Existencia Real
    For i = 4 To 7
        DataGrid1.Columns(i).Caption = RecuperaValor("Mínimo|Stock|Reser|Propuesta|", i - 3)
        DataGrid1.Columns(i).Width = 900
        DataGrid1.Columns(i).Alignment = dbgRight
        DataGrid1.Columns(i).NumberFormat = FormatoCantidad
    Next
    
    'fecha1,fecha2,porcen2
    For i = 8 To 10
        DataGrid1.Columns(i).Caption = ""
        DataGrid1.Columns(i).Width = 0
    Next
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    
    DataGrid1.ScrollBars = dbgAutomatic
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
                txtAux.Text = Data1.Recordset!propuesta
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
        txtAux.Left = DataGrid1.Columns(7).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(7).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
      '  txtAux.SelStart = 0
       ' txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub






Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = 1 Then Cancel = 1 'o aceptar o cancelar
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




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    'b = (Kmodo = 2)
   
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
'    BloquearText1 Me, Modo
    B = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera B
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
      
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
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


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    'PonerOpcionesMenuGeneral Me
End Sub




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
        
        If ActualizarExistencia() Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim Cantidad As Currency


    On Error GoTo EActualizar

    Cantidad = ImporteFormateado(txtAux.Text)
    
    If Cantidad < 0 Then Err.Raise 513, , "No se permiten negativos"
        
    If Cantidad <> Data1.Recordset!propuesta Then
    
        'Actualizar la Tabla: sinven con la cantidad introducida
        '-------------------------------------------------------
        
        SQL = "UPDATE straspaso Set propuesta = " & DBSet(Cantidad, "N")
        SQL = SQL & " WHERE secuencia =" & DBSet(CStr(Data1.Recordset!secuencia), "T")
        'SQL = SQL & " AND codusu =" & vUsu.codigo
        conn.Execute SQL
        
    End If
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function

