VERSION 5.00
Begin VB.Form frmUtDeclara 
   Caption         =   "Declarar ROPO"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSubvencionados 
      Caption         =   "Lotes subvencionados"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Facturas de tratamientos"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdComenzar 
      Caption         =   "Declaración"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   95
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   2880
      Picture         =   "frmUtDeclara.frx":0000
      Top             =   330
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmUtDeclara.frx":008B
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta "
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Desde "
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   510
   End
   Begin VB.Label lblInf 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
End
Attribute VB_Name = "frmUtDeclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim db As BaseDatos
Dim Rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim SQL As String
Dim cantidad As Double
Dim resto As Double

Private Sub cmdComenzar_Click()

    Screen.MousePointer = vbHourglass
    lblInf.Caption = "Incio proceso"
    lblInf.Refresh
    'RealizarProceso
    SQL = "A" 'antiguo
    If txtFecha(0).Text <> "" Then
        If CDate(txtFecha(0).Text) >= CDate("01/01/2015") Then SQL = ""
    End If
    If SQL = "A" Then
        'Antiguo. ES ek de Rafa
        NuevoProceso
    Else
        'Nuevo. Desde slifaclotes
        ProcesoDesdeSlifac
    End If
    
    lblInf.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '-- Abrimos la base de datos para trabajar con ella
    Set db = New BaseDatos
'    db.abrir "vAriges", "root", "aritel"
    db.asignar conn
    
    db.Tipo = "MYSQL"
    '-- Por defecto desde y hasta fecha de hoy
    txtFecha(0).Text = Format(Date, "dd/mm/yyyy")
    txtFecha(1).Text = Format(Date, "dd/mm/yyyy")
    
    
    Check2.Value = 0
    Check2.visible = vParamAplic.NumeroInstalacion = 1
    
    ChkSubvencionados.Value = 0
    ChkSubvencionados.visible = vParamAplic.LotesGeneralitat
    If vParamAplic.LotesGeneralitat Then ChkSubvencionados.Value = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set db = Nothing
End Sub








'****************************************************************************************
' Diciembre 2014
Private Sub NuevoProceso()

Dim nomDocum As String
Dim L As Long
Dim Col As Collection
Dim BuscarEnSlifacCampos As Boolean

Dim cadFecha As String

Dim CantidadEnLote As Currency
Dim ArticuloATratar As String
Dim NumeroDeLote As String
Dim CantidadQuedaEnLote As Currency
Dim UtilizadaEnLote As Currency
Dim HaMovidoLinFactura As Boolean
Dim fin As Boolean

    
    SQL = ""
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        SQL = "Debe indicar las fechas"
    Else
        
        '-- comprobamos que las fechas de paso son as correctas
        If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then SQL = "Fecha inicio mayor que fecha fin"
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbInformation
        Exit Sub
    End If
    
    
    lblInf.Caption = "Preparando datos"
    lblInf.Refresh
    
    '-- Eliminamos posibles declaraciones anteriores
    SQL = "delete from declaralom"
    db.ejecutar SQL
    
    '-- Antes de empezar y como vamos a hacer uso de canasign, lo limpiamos
    SQL = "update slotes set canasign = 0"
    db.ejecutar SQL
    
    
    BuscarEnSlifacCampos = False
    If vParamAplic.Ariagro <> "" Then
        SQL = " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slifaccampos", SQL, "1")
        If SQL <> "" Then
            If Val(SQL) > 0 Then BuscarEnSlifacCampos = True
        End If
    End If
    
    
    '-- Ahora vamos a por el gran mogollón
    lblInf.Caption = "Obtener lineas facturas"
    lblInf.Refresh
    SQL = "select a.codtipom, a.numfactu, a.fecfactu, a.codartic, a.nomartic, a.cantidad " & _
            ",b.nomclien, b.nifclien,b.domclien direccion,concat(codpobla,' ',pobclien) poblacion " & _
            ",d.descateg" & _
            " from slifac as a, scafac as b, sartic as c, scateg as d" & _
            " where a.codartic in" & _
            " (select codartic from sartic" & _
            " where codcateg in (select codcateg from scateg where ctrlotes = 1))" & _
            " and a.cantidad <> 0 " & _
            " and b.codtipom = a.codtipom" & _
            " and b.numfactu = a.numfactu" & _
            " and b.fecfactu = a.fecfactu" & _
            " and c.codartic = a.codartic" & _
            " and d.codcateg = c.codcateg" & _
            " and a.fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & _
            " order by codartic,a.fecfactu desc "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
    
    If Not Rs.EOF Then
    
        'Ahora vamos a contar los que hay
        L = 0
        While Not Rs.EOF
            Rs.MoveNext
            L = L + 1
        Wend
        Rs.MoveFirst
        
        
        lblInf.Tag = L
        L = 1
        ArticuloATratar = ""
        lblInf.Caption = ""
        fin = False
        HaMovidoLinFactura = False
        
        While Not fin
            
            
            lblInf.Caption = "Registro      " & L & " de " & lblInf.Tag & " "
            lblInf.Refresh
            If (L Mod 100) = 0 Then DoEvents
            
            
                
            
            If Rs!codArtic <> ArticuloATratar Then
                'OK. NUEVO ARTICULO
                If ArticuloATratar <> "" Then
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                        SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                        SQL = SQL & " and fecentra = " & db.Fecha(rs2!fecentra)
                        db.ejecutar SQL
                    End If
                
                    rs2.Close
                End If
                ArticuloATratar = Rs!codArtic
                
                
                
                HaMovidoLinFactura = True
                SQL = "select a.codartic, a.numlotes, a.fecentra, a.canentra, a.canasign, b.numserie from slotes as a, sartic as b" & _
                    " where a.codartic = " & db.texto(Rs!codArtic) & _
                    " and (a.canentra - a.canasign > 0)" & _
                    " and a.fecentra <= " & db.Fecha(Rs!FecFactu) & _
                    " and b.codartic = a.codartic" & _
                    " order by a.fecentra desc"
            
                Set rs2 = db.cursor(SQL)
                
                UtilizadaEnLote = 0
                CantidadQuedaEnLote = 0
                If Not rs2.EOF Then
                    NumeroDeLote = rs2!numlotes
                    CantidadQuedaEnLote = rs2!canentra
                End If
                
            End If
            
            If HaMovidoLinFactura Then
                cantidad = Rs!cantidad
                resto = cantidad
                HaMovidoLinFactura = False
            End If
            
            If rs2.EOF Then
                'NO HAY MAS LOTES
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                SQL = SQL & " values("
                SQL = SQL & db.Fecha(Rs!FecFactu) & "," ' FechaVenta
                SQL = SQL & db.texto(Rs!NomArtic) & "," ' NombreComercial
                SQL = SQL & db.texto(" ") & "," ' Registro
                SQL = SQL & db.texto(Rs!descateg) & "," ' Categoria
                SQL = SQL & db.texto(" ") & "," ' Lote
                SQL = SQL & db.numero(resto) & "," ' Cantidad
                SQL = SQL & db.texto(Rs!NomClien) & "," ' NombreSocio
                SQL = SQL & db.texto(Rs!nifClien) & "," ' NIF
                SQL = SQL & db.texto(Rs!codtipom & Format(Rs!Numfactu, "0000000")) & "," ' NumFactura
                'octubre 2011 EsVenta,Direccion,Poblacion
                            SQL = SQL & "1,"   ' es vebta
                            SQL = SQL & db.texto(Rs!Direccion) & "," ' direccion cliente
                            SQL = SQL & db.texto(Rs!Poblacion) & ")" ' poblacion
                
                
                db.ejecutar SQL
                HaMovidoLinFactura = True
                
            
            Else
                If rs2!fecentra > Rs!FecFactu Then
                    'Cantidad utilizada
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                        SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                        SQL = SQL & " and fecentra = " & db.Fecha(rs2!fecentra)
                        db.ejecutar SQL
                        UtilizadaEnLote = 0
                        
                    End If
                    rs2.MoveNext
                
                    If Not rs2.EOF Then
                        NumeroDeLote = rs2!numlotes
                        CantidadQuedaEnLote = rs2!canentra
                        UtilizadaEnLote = 0
                    End If
                Else
                    'OK. Articulo y lote. Vamos asignado
                    
                    
                    If resto <= CantidadQuedaEnLote Then
                        'En el lote queda la cantidad
                            'Para guardar
                            UtilizadaEnLote = UtilizadaEnLote + resto
                            CantidadQuedaEnLote = CantidadQuedaEnLote - resto
                            
                            SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                            SQL = SQL & " values("
                            SQL = SQL & db.Fecha(Rs!FecFactu) & "," ' FechaVenta
                            SQL = SQL & db.texto(Rs!NomArtic) & "," ' NombreComercial
                            SQL = SQL & db.texto(rs2!numSerie) & "," ' Registro
                            SQL = SQL & db.texto(Rs!descateg) & "," ' Categoria
                            SQL = SQL & db.texto(rs2!numlotes) & "," ' Lote
                            SQL = SQL & TransformaComasPuntos(db.numero(resto)) & "," ' Cantidad
                            SQL = SQL & db.texto(Rs!NomClien) & "," ' NombreSocio
                            SQL = SQL & db.texto(Rs!nifClien) & "," ' NIF
                            SQL = SQL & db.texto(Rs!codtipom & Format(Rs!Numfactu, "0000000")) & "," ' NumFactura
                            
                            'octubre 2011 EsVenta,Direccion,Poblacion
                            SQL = SQL & "1,"   ' es vebta
                            SQL = SQL & db.texto(Rs!Direccion) & "," ' direccion cliente
                            SQL = SQL & db.texto(Rs!Poblacion) & ")" ' poblacion
                            db.ejecutar SQL
                            
                            HaMovidoLinFactura = True
                        Else
                            
                            'Quedaba un poco en el lote
                            If CantidadQuedaEnLote > 0 Then
                                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                                SQL = SQL & " values("
                                SQL = SQL & db.Fecha(Rs!FecFactu) & "," ' FechaVenta
                                SQL = SQL & db.texto(Rs!NomArtic) & "," ' NombreComercial
                                SQL = SQL & db.texto(rs2!numSerie) & "," ' Registro
                                SQL = SQL & db.texto(Rs!descateg) & "," ' Categoria
                                SQL = SQL & db.texto(rs2!numlotes) & "," ' Lote
                                SQL = SQL & TransformaComasPuntos(db.numero(CantidadQuedaEnLote)) & ","  ' Cantidad
                                SQL = SQL & db.texto(Rs!NomClien) & "," ' NombreSocio
                                SQL = SQL & db.texto(Rs!nifClien) & "," ' NIF
                                SQL = SQL & db.texto(Rs!codtipom & Format(Rs!Numfactu, "0000000")) & ","
                                'octubre 2011 EsVenta,Direccion,Poblacion
                                SQL = SQL & "1,"   ' es vebta
                                SQL = SQL & db.texto(Rs!Direccion) & "," ' direccion cliente
                                SQL = SQL & db.texto(Rs!Poblacion) & ")" ' poblacion
                                
                                db.ejecutar SQL
                                resto = resto - CantidadQuedaEnLote  'nos queda "resto por asignar
                                UtilizadaEnLote = UtilizadaEnLote + CantidadQuedaEnLote
                            End If
                            
                            SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                            SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                            SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                            SQL = SQL & " and fecentra = " & db.Fecha(rs2!fecentra)
                            db.ejecutar SQL
                                
                            
                            HaMovidoLinFactura = False
                            rs2.MoveNext
                            UtilizadaEnLote = 0
                            CantidadQuedaEnLote = 0
                            If Not rs2.EOF Then
                                NumeroDeLote = rs2!numlotes
                                CantidadQuedaEnLote = rs2!canentra
                                Else
                                    'tsop
                            End If
                            
                            
                        End If
                    End If
                    

            End If
            
            
            If HaMovidoLinFactura Then
                Rs.MoveNext
                L = L + 1
            End If
            If Rs.EOF Then fin = True
            
            
        Wend
        '-- Por último actualizamos las compras
'        sql = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra)"
'        sql = sql & " select a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, '---', '---','----', a.canentra"
'        sql = sql & " from slotes as a, sartic as b, scateg as c"
'        sql = sql & " where b.codartic = a.codartic"
'        sql = sql & " and c.codcateg = b.codcateg"
        lblInf.Caption = "Proveedores"
        lblInf.Refresh
        DoEvents
        SQL = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,Poblacion)"
        SQL = SQL & "select distinct a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, e.nomprove, e.nifprove, d.document, a.canentra" & _
                " ,0, domprove,trim(concat(codpobla,' ',pobprove)) " & _
                " from slotes as a, sartic as b, scateg as c, smoval as d, sprove as e" & _
                " where b.codartic = a.codartic" & _
                " and c.codcateg = b.codcateg" & _
                " and d.codartic = a.codartic" & _
                " and d.fechamov = a.fecentra" & _
                " and d.tipomovi = 1" & _
                " and d.detamovi = 'ALC'" & _
                " and a.fecentra between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & _
                " and e.codprove = d.codigope"
        
        
        db.ejecutar SQL
        
        Rs.Close
        
        'Si temenos enlace con ariagro, podemos intentar sacar los tratamientos
        
        
        'If vParamAplic.Ariagro <> "" Then
        If BuscarEnSlifacCampos Then
            lblInf.Caption = "Enlace ariagro"
            lblInf.Refresh
            DoEvents
            Set Col = New Collection
            
            'Junio 2014
            cadFecha = " FechaVenta between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            L = 0
            SQL = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then L = DBLet(Rs.Fields(0), "N")
            Rs.Close
            lblInf.Tag = L
            
            SQL = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            SQL = ""
            While Not Rs.EOF
                L = L + 1
                lblInf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblInf.Tag
                lblInf.Refresh
                
                SQL = SQL & ", (" & DBSet(Rs!FechaVenta, "F") & "," & DBSet(Rs.Fields(1), "T") & "," & Rs.Fields(2) & ")"
                Rs.MoveNext
                
                
                If L > 29 Then
                    Col.Add SQL
                    SQL = ""
                    DoEvents
                    L = 0
                End If
            Wend
            Rs.Close
            
            If L > 0 Then Col.Add SQL
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count
                lblInf.Refresh
                If (L Mod 5) = 0 Then DoEvents
                SQL = "(" & Mid(Col.item(L), 2) & ")"
                SQL = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & SQL
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    
                    'FAV0079016
                    SQL = " AND numfactura = '" & Rs!codtipom & Format(Rs!Numfactu, "0000000") & "'"
                    SQL = " WHERE esventa=1 and fechaventa= " & DBSet(Rs!FecFactu, "F") & SQL
                    
                    SQL = "UPDATE declaraLOM SET cultivo=" & Rs!codCampo & SQL
                    conn.Execute SQL
                    Rs.MoveNext
                Wend
                Rs.Close
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblInf.Caption = "Obtener variedad"
            lblInf.Refresh
            DoEvents
            SQL = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                lblInf.Caption = "Campo " & Rs!cultivo
                lblInf.Refresh
                SQL = "select rcampos.codcampo,  variedades.nomvarie"
                SQL = SQL & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
                SQL = SQL & " WHERE codcampo =" & Rs!cultivo
                
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = "N/D"
                Else
                    SQL = rs2!nomvarie
                End If
                rs2.Close
                SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T") & " WHERE cultivo =" & DBSet(Rs!cultivo, "T")
                conn.Execute SQL
                
                Rs.MoveNext
            Wend
            Rs.Close
            
            
        
        End If 'de ariagro
        
        'Abril 2015
        'ALZIRA
        If vParamAplic.NumeroInstalacion = 1 Then
            'Para aquellas facturas de servicio (que son tratamientos), si no esta indicado el cultivo, ni  la variedad
            'entonces UPDATEAMOS con los datos de la observacion
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            SQL = "select fechaventa, NombreComercial,Registro,Categoria,Lote,NIF,NumFactura"
            SQL = SQL & " from declaralom where esventa=1 and numfactura like 'FAS%' and cultivo is null and tratamiento is null"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                lblInf.Caption = "Fra: " & Rs!NumFactura
                lblInf.Refresh
                SQL = "select * from scafac1 where codtipom='FAS' "
                SQL = SQL & " and fecfactu=" & DBSet(Rs!FechaVenta, "F") & " and numfactu=" & Mid(Rs!NumFactura, 4)
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = ""
                Else
                    SQL = Trim(DBLet(rs2!observa1, "T"))
                End If
                rs2.Close
                
                If SQL <> "" Then
                    NumRegElim = 1
                    
                    'Vamos a quitar todos los espacios en blanco "duplicados"
                    Do
                        NumRegElim = InStr(NumRegElim, SQL, " ")
                        If NumRegElim > 0 Then
                            Do
                                L = InStr(NumRegElim + 1, SQL, " ")
                                If L = NumRegElim + 1 Then
                                    SQL = Mid(SQL, 1, L - 1) & Mid(SQL, L + 1)
                                Else
                                    L = 0
                                End If
                            Loop Until L = 0
                            NumRegElim = NumRegElim + 1
                        End If
                    Loop Until NumRegElim = 0
                            
                    
                
                
                    L = Len(SQL)
                    If L > 45 Then
                        cadFecha = Mid(SQL, 46)
                        SQL = Mid(SQL, 1, 45)
                    Else
                        cadFecha = ""
                    End If
                    SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T")
                    SQL = SQL & ",tratamiento= " & DBSet(cadFecha, "T", "S")
                    SQL = SQL & " where fechaventa=" & DBSet(Rs!FechaVenta, "F") & " and numfactura='" & Rs!NumFactura
                    SQL = SQL & "' and lote=" & DBSet(Rs!lote, "T") & " and nif=" & DBSet(Rs!NIF, "T")
                    SQL = SQL & " and registro=" & DBSet(Rs!registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute SQL
                
                End If
                
                Rs.MoveNext
            Wend
            Rs.Close
            


        
        
        End If
        
        DoEvents
        '-- Llamar al informe
        Dim Desde As Date
        Dim Hasta As Date
        Desde = CDate(txtFecha(0).Text)
        Hasta = CDate(txtFecha(1).Text)
        frmVisReport.OtrosParametros = "|FecDesde=Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")|" & _
                                       "FecHasta=Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")|"
        frmVisReport.NumeroParametros = 2
'        frmVisReport.Informe = App.Path & "\Informes\" & "declaracion_lom.rpt"
        
        'Añade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT2(31, "", 0, nomDocum, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
            Exit Sub
        End If
        frmVisReport.Informe = App.Path & "\Informes\" & nomDocum
        
        frmVisReport.FormulaSeleccion = "{declaralom.FechaVenta} in " & _
                                            "Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")" & _
                                            " to" & _
                                            " Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")"
        frmVisReport.Show vbModal
        '--
        lblInf.Caption = "Proceso terminado."
        lblInf.Refresh
        DoEvents
    Else
        MsgBox "NO existen datos entre las fechas", vbExclamation
        Rs.Close
    End If
    Set Rs = Nothing
    Set rs2 = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)

   
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   SQL = ""
   frmC.Show vbModal
   Set frmC = Nothing
    If SQL <> "" Then txtFecha(Index).Text = SQL

End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub




'Ultimisimo proceso.
' Carnet de Manipulador instaudado
' los lotes van en la tabla slifaclotes
' Objetivo:
'   1.)  Ventas         a) todas las slifac  deben tener lotes
'                       b) De la scafac cogeremos el manipulador. Si no existe pondremos el cliente
'
'   2.-)  Compras       a) Todas las compras slipc deben tener slotes y viceversa(como mujeres y hombres)
'
'   Insertaremos todos estos datos(del periodo) en una tabla tmp que despues mostraremos en el report

' OCTUBRE 2016
'   Metemos lotes subvenciondaos.
'   Entraran en la BD desde las tablas  slotesgeneralitat  slotesgeneralitatmov



'****************************************************************************************
' Noviembre 2015
Private Sub ProcesoDesdeSlifac()
Dim BuscarEnSlifacCampos As Boolean
Dim cadFecha As String
Dim Raux As ADODB.Recordset
Dim Errores As String
Dim Aux As String
Dim L As Long
Dim Col As Collection
Dim CadLote As String
Dim fin As Boolean
Dim LotesCorrectos As Boolean
Dim MoverRsPpal As Boolean
Dim NF As Integer
Dim SQL_Servicios As String

    SQL = ""
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        SQL = "Debe indicar las fechas"
    Else
        
        '-- comprobamos que las fechas de paso son as correctas
        If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then SQL = "Fecha inicio mayor que fecha fin"
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbInformation
        Exit Sub
    End If
    
    
    lblInf.Caption = "Preparando datos"
    lblInf.Refresh
    
    '-- Eliminamos posibles declaraciones anteriores
    SQL = "delete from declaralom"
    db.ejecutar SQL
    
    '-- No vamos a hacer uso de canasign, lo limpiamos de todas formas
    SQL = "update slotes set canasign = 0"
    db.ejecutar SQL
    
    
    BuscarEnSlifacCampos = False
    If vParamAplic.Ariagro <> "" Then
        SQL = " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slifaccampos", SQL, "1")
        If SQL <> "" Then
            If Val(SQL) > 0 Then BuscarEnSlifacCampos = True
        End If
    End If
    
    Set Rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Errores = ""
    
    '1.- Comprobamos que todos los articulos vendidos en el periodo, que deberian tener lote
    SQL = "select codcateg from scateg where ctrlotes = 1"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not Rs.EOF
        Aux = Aux & ", " & DBSet(Rs!codCateg, "T")
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Aux = "" Then
        MsgBox "Categorias sin control de lotes", vbExclamation
        Exit Sub
    End If
        
    Aux = "(" & Mid(Aux, 2) & ")"
    
    SQL = "select slifac.codartic,slifac.nomartic from  slifac,sartic where slifac.codartic=sartic.codartic "
    SQL = SQL & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " AND codcateg in " & Aux & "  and coalesce(numserie,'')=''"
    
    If vParamAplic.NumeroInstalacion = 1 Then
        'ALZIRA
        SQL_Servicios = " = "
        If Me.Check2.Value = 0 Then SQL_Servicios = " <> "
        SQL_Servicios = SQL_Servicios & "'FAS'"
        
        SQL = SQL & " AND slifac.codtipom" & SQL_Servicios
        
    End If
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    Aux = Space(20)
    While Not Rs.EOF
        SQL = SQL & Mid(Rs!codArtic & Aux, 1, 20) & Rs!NomArtic & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    If SQL <> "" Then
        Aux = "Errores en articulos. No esta indicado el numero de registro" & vbCrLf & String(40, "=") & vbCrLf & SQL
        Errores = Errores & Aux
    End If
    
    
    'Vemos que todos los articulos vendidos en el periodo que deberian tener lote, tienen lote
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute SQL
    
    
    
    lblInf.Caption = "Comprobando lotes"
    lblInf.Refresh
    
    SQL = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifac,sartic WHERE slifac.codartic=sartic.codartic"
    SQL = SQL & " and numserie<>'' and numlote<>'' AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    
    If SQL_Servicios <> "" Then SQL = SQL & " AND slifac.codtipom" & SQL_Servicios
    
    SQL = SQL & " group by 1,2,3 order by 1,2,3"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    SQL = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifaclotes"
    SQL = SQL & " WHERE fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    If SQL_Servicios <> "" Then SQL = SQL & " AND slifaclotes.codtipom" & SQL_Servicios
    
    SQL = SQL & "  group by 1,2,3 order by 1,2,3"
    rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not Rs.EOF
        
        
        
        fin = False
        LotesCorrectos = False
        MoverRsPpal = True
        If rs2.EOF Then
            
        Else
            If Rs!codtipom = rs2!codtipom Then
                If Val(Rs!Numfactu) = Val(rs2!Numfactu) Then
                    If Rs!FecFactu = rs2!FecFactu Then
                        If Rs!canti <= rs2!canti Then LotesCorrectos = True
                    End If
                End If
            End If
            
            
            If LotesCorrectos Then
                rs2.MoveNext
            Else
                
                If Rs!codtipom <> rs2!codtipom Then
                    MsgBox "Avise soporte tecnico. Err: codtipom", vbExclamation
                Else
                    If Val(Rs!Numfactu) > Val(rs2!Numfactu) Then
                        
                        rs2.MoveNext
                        MoverRsPpal = False
                    Else
                        SQL = Rs!codtipom & "|" & Rs!Numfactu & "|" & Rs!FecFactu & "|"
                        Col.Add SQL
                    End If
                End If
                
                
            End If
        End If
        
        
        If MoverRsPpal Then Rs.MoveNext
    Wend
    rs2.Close
    Rs.Close
              
    
    
    
    SQL = ""
              
    For L = 1 To Col.Count
        lblInf.Caption = "Lotes FRA" & Col.item(L)
        lblInf.Refresh
        Debug.Print Col.item(L)
        Aux = ", ('" & RecuperaValor(Col.item(L), 1) & "'," & RecuperaValor(Col.item(L), 2) & "," & DBSet(RecuperaValor(Col.item(L), 3), "F") & ")"
        SQL = SQL & Aux & vbCrLf
    Next
    If SQL <> "" Then
        If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
        Aux = "Errores en lotes. Facturas no coinciden lotes (factura/asignados) " & vbCrLf & String(40, "=") & vbCrLf & SQL
        Errores = Errores & Aux
    End If
   
    If Errores <> "" Then
        NF = FreeFile
        Open App.Path & "\ErrROPO.txt" For Output As #NF
        Print #NF, Errores
        Close #NF
        
    End If
    '-- Ahora vamos a por el gran mogollón
    lblInf.Caption = "Datos manipulador "
    lblInf.Refresh
    SQL = "select codtipom,numfactu,fecfactu,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet from"
    SQL = SQL & " scafac1 where fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " AND ManipuladorNumCarnet <> ''"
    If SQL_Servicios <> "" Then SQL = SQL & " AND scafac1.codtipom" & SQL_Servicios
    SQL = SQL & " ORDER BY codtipom,numfactu,fecfactu ,manipuladornumcarnet desc"
    'rs2.Open SQL, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    
    lblInf.Caption = "Obtener registros "
    lblInf.Refresh
    'SQL = "select a.codtipom, a.numfactu, a.fecfactu, a.codartic, c.nomartic, a.cantidad " & _
            ",b.nomclien, b.nifclien,b.domclien direccion,concat(codpobla,' ',pobclien) poblacion " & _
            ",d.descateg , numlote,numserie " & _
            " from slifaclotes as a, scafac as b, sartic as c, scateg as d" & _
            " where c.numserie<>'' AND a.codartic in" & _
            " (select codartic from sartic" & _
            " where codcateg in (select codcateg from scateg where ctrlotes = 1))" & _
            " and a.cantidad <> 0 " & _
            " and b.codtipom = a.codtipom" & _
            " and b.numfactu = a.numfactu" & _
            " and b.fecfactu = a.fecfactu" & _
            " and c.codartic = a.codartic" & _
            " and d.codcateg = c.codcateg" & _
            " and a.fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & _
            " order by codartic,a.fecfactu desc "
    
    SQL = " select a.codtipom, a.numfactu, h.fechaalb, a.codartic, c.nomartic, a.cantidad ,b.nomclien, b.nifclien,"
    SQL = SQL & " b.domclien direccion,concat(codpobla,' ',pobclien) poblacion ,d.descateg , numlote,numserie"
    SQL = SQL & " ,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet"
    SQL = SQL & " From slifaclotes as a, scafac as b,scafac1 h,sartic as c, scateg as d where c.numserie<>''"
    SQL = SQL & " and ctrlotes = 1 and a.cantidad <> 0 and c.codartic = a.codartic and d.codcateg = c.codcateg"
    SQL = SQL & " and h.fechaalb between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " and b.codtipom = h.codtipom and b.numfactu = h.numfactu and b.fecfactu = h.fecfactu"
    SQL = SQL & " and b.codtipom = a.codtipom and b.numfactu = a.numfactu and b.fecfactu = a.fecfactu"
    SQL = SQL & " and a.codtipoa = h.codtipoa and a.numalbar = h.numalbar"
    
     If SQL_Servicios <> "" Then SQL = SQL & " AND a.codtipom" & SQL_Servicios
    
    SQL = SQL & " order by codartic,h.fechaalb desc"
    
    Rs.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
    While Not Rs.EOF
    
    
                lblInf.Caption = "Ventas: " & Rs!codtipom & " " & Rs!Numfactu & " " & Rs!codArtic
                lblInf.Refresh
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet, NifMani)"
                SQL = SQL & " values("
                SQL = SQL & db.Fecha(Rs!FechaAlb) & "," ' FechaVenta
                SQL = SQL & db.texto(Rs!NomArtic) & "," ' NombreComercial
                SQL = SQL & db.texto(Rs!numSerie) & "," ' Registro
                SQL = SQL & db.texto(Rs!descateg) & "," ' Categoria
                SQL = SQL & db.texto(Rs!numLote) & "," ' Lote
                SQL = SQL & db.numero(Rs!cantidad) & "," ' Cantidad
    
                'ENERO 2016
                LotesCorrectos = DBLet(Rs!ManipuladorNombre, "T") <> ""
'
'                If LotesCorrectos Then
'                    'ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet
'                    SQL = SQL & db.texto(RS!ManipuladorNombre) & "," ' NombreSocio
'                Else
'                    SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
'                End If
                SQL = SQL & db.texto(Rs!NomClien) & "," ' NombreSocio
                
                SQL = SQL & db.texto(Rs!nifClien) & "," ' NIF
                SQL = SQL & db.texto(Rs!codtipom & Format(Rs!Numfactu, "0000000")) & "," ' NumFactura
                SQL = SQL & "1,"   ' es vebta
                SQL = SQL & db.texto(Rs!Direccion) & "," ' direccion cliente
                SQL = SQL & db.texto(Rs!Poblacion) & "," ' poblacion
                
                
 
                'Llevamos tanto el nombre del cliente como el de manipulador
                'NomCarnetMani, NumCarnet, NifMani
                If LotesCorrectos Then
                    'Datos carnet manipulador
                    SQL = SQL & db.texto(Rs!ManipuladorNombre) & "," ' NombreSocio
                    SQL = SQL & db.texto(Rs!ManipuladorNumCarnet) & ","
                    SQL = SQL & "NULL)"  ' poblacion
                
                Else
                    SQL = SQL & "NULL,NULL,NULL)"  ' poblacion
                End If
                db.ejecutar SQL
        
        
            Rs.MoveNext
        Wend
        Rs.Close
        

        lblInf.Caption = "Proveedores"
        lblInf.Refresh
        DoEvents
        SQL = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet, NifMani)"
        SQL = SQL & "select distinct a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, e.nomprove, e.nifprove, d.document, a.canentra" & _
                " ,0, domprove,trim(concat(codpobla,' ',pobprove)) " & _
                " ,NULL,NULL,NULL from slotes as a, sartic as b, scateg as c, smoval as d, sprove as e" & _
                " where b.codartic = a.codartic" & _
                " and c.codcateg = b.codcateg" & _
                " and d.codartic = a.codartic" & _
                " and d.fechamov = a.fecentra" & _
                " and d.tipomovi = 1" & _
                " and d.detamovi = 'ALC'" & _
                " and a.fecentra between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & _
                " and e.codprove = d.codigope"
        
        'Las compras, para cuando son SERVICIOS(Alzira) no van
        If SQL_Servicios <> "" Then
            'Es decir, para servicio digo que a=-1 y me devuelve EOF
            If Me.Check2.Value = 1 Then SQL = SQL & " AND a.codartic='-1A1-'"
        End If
        
        db.ejecutar SQL
        
        
        
        'If vParamAplic.Ariagro <> "" Then
        If BuscarEnSlifacCampos Then
        
            
        
            lblInf.Caption = "Enlace ariagro"
            lblInf.Refresh
            DoEvents
            Set Col = New Collection
            
            'Junio 2014
            cadFecha = " FechaVenta between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            L = 0
            SQL = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then L = DBLet(Rs.Fields(0), "N")
            Rs.Close
            lblInf.Tag = L
            
            SQL = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            SQL = ""
            While Not Rs.EOF
                L = L + 1
                lblInf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblInf.Tag
                lblInf.Refresh
                
                
                'Graba FECHA ALBARAN, y luiego no encuentra por fecha factura.
                'Buscamos , de momento, por serie+factura
                'SQL = SQL & ", (" & DBSet(Rs!fechaventa, "F") & "," & DBSet(Rs.Fields(1), "T") & "," & Rs.Fields(2) & ")"
                SQL = SQL & ", (" & DBSet(Rs.Fields(1), "T") & "," & Rs.Fields(2) & ")"
                Rs.MoveNext
                
                
                If L > 29 Then
                    Col.Add SQL
                    SQL = ""
                    DoEvents
                    L = 0
                End If
            Wend
            Rs.Close
            
            If L > 0 Then Col.Add SQL
            
            
            'Abro los tratamientos
            lblInf.Caption = "Leyendo tratamientos BD..."
            lblInf.Refresh
            
            SQL = "select codtrata,nomtrata from advtrata"
            rs2.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
            
            
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count & " Cultivo"
                lblInf.Refresh
                DoEvents
                SQL = "(" & Mid(Col.item(L), 2) & ")"
                'SQL = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & SQL
                SQL = "Select * from slifaccampos where (codtipom,numfactu) IN " & SQL
                SQL = SQL & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    
                    'FAV0079016
                    SQL = " AND numfactura = '" & Rs!codtipom & Format(Rs!Numfactu, "0000000") & "'"
                    'SQL = " WHERE esventa=1 and fechaventa= " & DBSet(Rs!FecFactu, "F") & SQL
                    SQL = " WHERE esventa=1 " & SQL
                    SQL = "UPDATE declaraLOM SET cultivo=" & Rs!codCampo & SQL
                    conn.Execute SQL
                    Rs.MoveNext
                Wend
                Rs.Close
                
                
                'Vamos a ver los tratamientos
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count & " Tratamiento"
                lblInf.Refresh

                
                SQL = "Select codtipom,numfactu,GROUP_CONCAT( substring(referenc,7) separator ' , ') from scafac1 where "
                SQL = SQL & " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                SQL = SQL & " AND referenc like 'PARTE%' "
                SQL = SQL & " AND (codtipom,numfactu) IN (" & Mid(Col.item(L), 2) & ") group by 1,2"
                
                
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    SQL = DBLet(Rs.Fields(2), "T")
                    lblInf.Caption = "parte : " & SQL
                    lblInf.Refresh
                    
                    If SQL <> "" Then
                    
                        NF = InStrRev(SQL, " , ")
                        If NF > 0 Then SQL = Mid(SQL, NF + 3)
                        
                        SQL = DevuelveDesdeBD(conAri, "codtrata", "advpartes", "numparte", SQL)
                        If SQL <> "" Then
                            rs2.Find "codtrata = " & SQL, , adSearchForward, 1
                            If Not rs2.EOF Then
                                'FAV0079016
                                SQL = " AND numfactura = '" & Rs!codtipom & Format(Rs!Numfactu, "0000000") & "'"
                                'SQL = " WHERE esventa=1 and fechaventa= " & DBSet(Rs!FecFactu, "F") & SQL
                                SQL = " WHERE esventa=1 " & SQL
                                SQL = "UPDATE declaraLOM SET tratamiento=" & DBSet(rs2!nomtrata, "T") & SQL
                                conn.Execute SQL
                            End If
                        End If
                    End If
                    Rs.MoveNext
                Wend
                Rs.Close
               
                  
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblInf.Caption = "Obtener variedad"
            lblInf.Refresh
            DoEvents
            SQL = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                lblInf.Caption = "Campo " & Rs!cultivo
                lblInf.Refresh
                SQL = "select rcampos.codcampo,  variedades.nomvarie"
                SQL = SQL & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
                SQL = SQL & " WHERE codcampo =" & Rs!cultivo
                
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = "N/D"
                Else
                    SQL = rs2!nomvarie
                End If
                rs2.Close
                SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T") & " WHERE cultivo =" & DBSet(Rs!cultivo, "T")
                conn.Execute SQL
                
                Rs.MoveNext
            Wend
            Rs.Close
            
                        
                        
                        
                        
                        
                        
                        
        
        End If 'de ariagro
        
        'Abril 2015
        'ALZIRA
        If vParamAplic.NumeroInstalacion = 1 Then
            'Para aquellas facturas de servicio (que son tratamientos), si no esta indicado el cultivo, ni  la variedad
            'entonces UPDATEAMOS con los datos de la observacion
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            SQL = "select fechaventa, NombreComercial,Registro,Categoria,Lote,NIF,NumFactura"
            SQL = SQL & " from declaralom where esventa=1 and numfactura like 'FAS%' and cultivo is null and tratamiento is null"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                lblInf.Caption = "Fra: " & Rs!NumFactura
                lblInf.Refresh
                SQL = "select * from scafac1 where codtipom='FAS' "
                'SQL = SQL & " and fecfactu=" & DBSet(Rs!fechaventa, "F") & " and numfactu=" & Mid(Rs!NumFactura, 4)
                SQL = SQL & " and numfactu=" & Mid(Rs!NumFactura, 4)
              
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = ""
                Else
                    SQL = Trim(DBLet(rs2!observa1, "T"))
                End If
                rs2.Close
                
                If SQL <> "" Then
                    NumRegElim = 1
                    
                    'Vamos a quitar todos los espacios en blanco "duplicados"
                    Do
                        NumRegElim = InStr(NumRegElim, SQL, " ")
                        If NumRegElim > 0 Then
                            Do
                                L = InStr(NumRegElim + 1, SQL, " ")
                                If L = NumRegElim + 1 Then
                                    SQL = Mid(SQL, 1, L - 1) & Mid(SQL, L + 1)
                                Else
                                    L = 0
                                End If
                            Loop Until L = 0
                            NumRegElim = NumRegElim + 1
                        End If
                    Loop Until NumRegElim = 0
                            
                    
                
                
                    L = Len(SQL)
                    If L > 45 Then
                        cadFecha = Mid(SQL, 46)
                        SQL = Mid(SQL, 1, 45)
                    Else
                        cadFecha = ""
                    End If
                    SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T")
                    SQL = SQL & ",tratamiento= " & DBSet(cadFecha, "T", "S")
                    SQL = SQL & " where fechaventa=" & DBSet(Rs!FechaVenta, "F") & " and numfactura='" & Rs!NumFactura
                    SQL = SQL & "' and lote=" & DBSet(Rs!lote, "T") & " and nif=" & DBSet(Rs!NIF, "T")
                    SQL = SQL & " and registro=" & DBSet(Rs!registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute SQL
                
                End If
                
                Rs.MoveNext
            Wend
            Rs.Close
            


        
        
        End If
        
        
        
        
        'OCTUBRE 2016
        ' Lotes fitosnatiarios SUBVENCIONDOS
        If Me.ChkSubvencionados.Value = 1 Then
            DoEvents
            
            
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio,"
                SQL = SQL & " NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet)"
                SQL = SQL & " SELECT slotesgeneralitatmov.fechamov,nomartic,slotesgeneralitat.numserie,'LO',numlote,"
                SQL = SQL & " slotesgeneralitatmov.cantidad,nomclien,nifclien,concat(""ID"" ,"
                SQL = SQL & " right(concat(""000000"",id),6) ,right(concat(""0000"",idMov),4)),1,domclien,"
                SQL = SQL & " concat(codpobla,' ',pobclien) poblacion,slotesgeneralitatmov.ManipuladorNombre , "
                SQL = SQL & " slotesgeneralitatmov.ManipuladorNumCarnet"
                SQL = SQL & " from slotesgeneralitat,slotesgeneralitatmov,sclien,sartic where slotesgeneralitat.Id = "
                SQL = SQL & " slotesgeneralitatmov.idlote and slotesgeneralitatmov.codclien=sclien.codclien"
                SQL = SQL & " and slotesgeneralitat.codartic=sartic.codartic"
                SQL = SQL & " AND slotesgeneralitatmov.fechamov between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                SQL = SQL & " ORDER BY slotesgeneralitatmov.fechamov,id"
                lblInf.Caption = "Lotes subvencionados"
                lblInf.Refresh
                
                
                
                db.ejecutar SQL
        
        
            
                DoEvents
                Me.Refresh
                
                lblInf.Caption = "Proveedores lotes subv."
                lblInf.Refresh
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, cancompra, NombreSocio,"
                SQL = SQL & " NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet)"
                SQL = SQL & " SELECT slotesgeneralitat.fecha,nomartic,slotesgeneralitat.numserie,'LO',numlote,"
                SQL = SQL & " slotesgeneralitat.cantidad,nomprove,nifprove,concat(""COD "" ,right(concat(""000000"",id),6)),0,domprove,"
                SQL = SQL & " concat(codpobla,' ',pobprove) poblacion,null , null"
                SQL = SQL & " From slotesgeneralitat, sartic, sprove Where slotesgeneralitat.Codprove = sprove.Codprove"
                SQL = SQL & " and slotesgeneralitat.codartic=sartic.codartic"
                SQL = SQL & " AND slotesgeneralitat.fecha  between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                SQL = SQL & " ORDER BY slotesgeneralitat.fecha,id"
                db.ejecutar SQL

        
        
        
        
        
        
        
        
        
        End If
        
        
        DoEvents
        '-- Llamar al informe
        Dim Desde As Date
        Dim Hasta As Date
        Desde = CDate(txtFecha(0).Text)
        Hasta = CDate(txtFecha(1).Text)
        frmVisReport.OtrosParametros = "|FecDesde=Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")|" & _
                                       "FecHasta=Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")|"
        frmVisReport.NumeroParametros = 2
'        frmVisReport.Informe = App.Path & "\Informes\" & "declaracion_lom.rpt"
        
        'Añade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT2(31, "", 0, Aux, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
            Exit Sub
        End If
        
        
        
        If Me.Check2.Value = 1 Then Aux = Replace(Aux, ".rpt", "S.rpt")
        frmVisReport.Informe = App.Path & "\Informes\" & Aux
        
        frmVisReport.FormulaSeleccion = "{declaralom.FechaVenta} in " & _
                                            "Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")" & _
                                            " to" & _
                                            " Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")"
        frmVisReport.Show vbModal
        '--
        lblInf.Caption = "Proceso terminado."
        lblInf.Refresh
        DoEvents
  
    Set Rs = Nothing
    Set rs2 = Nothing
End Sub




