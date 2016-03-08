VERSION 5.00
Begin VB.Form frmUtDeclara 
   Caption         =   "Declarar ROPO"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Verificar datos"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
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
      Caption         =   "Salir"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdComenzar 
      Caption         =   "Declaración"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
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
      Top             =   1440
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
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim SQL As String
Dim Cantidad As Double
Dim resto As Double

Private Sub cmdComenzar_Click()

    Screen.MousePointer = vbHourglass
    lblinf.Caption = "Incio proceso"
    lblinf.Refresh
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
    
    lblinf.Caption = ""
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
    
    
    lblinf.Caption = "Preparando datos"
    lblinf.Refresh
    
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
    lblinf.Caption = "Obtener lineas facturas"
    lblinf.Refresh
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
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
    
    If Not RS.EOF Then
    
        'Ahora vamos a contar los que hay
        L = 0
        While Not RS.EOF
            RS.MoveNext
            L = L + 1
        Wend
        RS.MoveFirst
        
        
        lblinf.Tag = L
        L = 1
        ArticuloATratar = ""
        lblinf.Caption = ""
        fin = False
        HaMovidoLinFactura = False
        
        While Not fin
            
            
            lblinf.Caption = "Registro      " & L & " de " & lblinf.Tag & " "
            lblinf.Refresh
            If (L Mod 100) = 0 Then DoEvents
            
            
                
            
            If RS!codArtic <> ArticuloATratar Then
                'OK. NUEVO ARTICULO
                If ArticuloATratar <> "" Then
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                        SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                        SQL = SQL & " and fecentra = " & db.Fecha(rs2!FecEntra)
                        db.ejecutar SQL
                    End If
                
                    rs2.Close
                End If
                ArticuloATratar = RS!codArtic
                
                
                'If Rs!codArtic = "010009" Then Stop
                
                HaMovidoLinFactura = True
                SQL = "select a.codartic, a.numlotes, a.fecentra, a.canentra, a.canasign, b.numserie from slotes as a, sartic as b" & _
                    " where a.codartic = " & db.texto(RS!codArtic) & _
                    " and (a.canentra - a.canasign > 0)" & _
                    " and a.fecentra <= " & db.Fecha(RS!FecFactu) & _
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
                Cantidad = RS!Cantidad
                resto = Cantidad
                HaMovidoLinFactura = False
            End If
            
            If rs2.EOF Then
                'NO HAY MAS LOTES
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                SQL = SQL & " values("
                SQL = SQL & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                SQL = SQL & db.texto(RS!NomArtic) & "," ' NombreComercial
                SQL = SQL & db.texto(" ") & "," ' Registro
                SQL = SQL & db.texto(RS!descateg) & "," ' Categoria
                SQL = SQL & db.texto(" ") & "," ' Lote
                SQL = SQL & db.numero(resto) & "," ' Cantidad
                SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
                SQL = SQL & db.texto(RS!nifClien) & "," ' NIF
                SQL = SQL & db.texto(RS!codtipom & Format(RS!NumFactu, "0000000")) & "," ' NumFactura
                'octubre 2011 EsVenta,Direccion,Poblacion
                            SQL = SQL & "1,"   ' es vebta
                            SQL = SQL & db.texto(RS!Direccion) & "," ' direccion cliente
                            SQL = SQL & db.texto(RS!Poblacion) & ")" ' poblacion
                
                
                db.ejecutar SQL
                HaMovidoLinFactura = True
                
            
            Else
                If rs2!FecEntra > RS!FecFactu Then
                    'Cantidad utilizada
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                        SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                        SQL = SQL & " and fecentra = " & db.Fecha(rs2!FecEntra)
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
                            SQL = SQL & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                            SQL = SQL & db.texto(RS!NomArtic) & "," ' NombreComercial
                            SQL = SQL & db.texto(rs2!numSerie) & "," ' Registro
                            SQL = SQL & db.texto(RS!descateg) & "," ' Categoria
                            SQL = SQL & db.texto(rs2!numlotes) & "," ' Lote
                            SQL = SQL & TransformaComasPuntos(db.numero(resto)) & "," ' Cantidad
                            SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
                            SQL = SQL & db.texto(RS!nifClien) & "," ' NIF
                            SQL = SQL & db.texto(RS!codtipom & Format(RS!NumFactu, "0000000")) & "," ' NumFactura
                            
                            'octubre 2011 EsVenta,Direccion,Poblacion
                            SQL = SQL & "1,"   ' es vebta
                            SQL = SQL & db.texto(RS!Direccion) & "," ' direccion cliente
                            SQL = SQL & db.texto(RS!Poblacion) & ")" ' poblacion
                            db.ejecutar SQL
                            
                            HaMovidoLinFactura = True
                        Else
                            
                            'Quedaba un poco en el lote
                            If CantidadQuedaEnLote > 0 Then
                                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                                SQL = SQL & " values("
                                SQL = SQL & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                                SQL = SQL & db.texto(RS!NomArtic) & "," ' NombreComercial
                                SQL = SQL & db.texto(rs2!numSerie) & "," ' Registro
                                SQL = SQL & db.texto(RS!descateg) & "," ' Categoria
                                SQL = SQL & db.texto(rs2!numlotes) & "," ' Lote
                                SQL = SQL & TransformaComasPuntos(db.numero(CantidadQuedaEnLote)) & ","  ' Cantidad
                                SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
                                SQL = SQL & db.texto(RS!nifClien) & "," ' NIF
                                SQL = SQL & db.texto(RS!codtipom & Format(RS!NumFactu, "0000000")) & ","
                                'octubre 2011 EsVenta,Direccion,Poblacion
                                SQL = SQL & "1,"   ' es vebta
                                SQL = SQL & db.texto(RS!Direccion) & "," ' direccion cliente
                                SQL = SQL & db.texto(RS!Poblacion) & ")" ' poblacion
                                
                                db.ejecutar SQL
                                resto = resto - CantidadQuedaEnLote  'nos queda "resto por asignar
                                UtilizadaEnLote = UtilizadaEnLote + CantidadQuedaEnLote
                            End If
                            
                            SQL = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                            SQL = SQL & " where codartic = " & db.texto(rs2!codArtic)
                            SQL = SQL & " and numlotes = " & db.texto(rs2!numlotes)
                            SQL = SQL & " and fecentra = " & db.Fecha(rs2!FecEntra)
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
                RS.MoveNext
                L = L + 1
            End If
            If RS.EOF Then fin = True
            
            
        Wend
        '-- Por último actualizamos las compras
'        sql = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra)"
'        sql = sql & " select a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, '---', '---','----', a.canentra"
'        sql = sql & " from slotes as a, sartic as b, scateg as c"
'        sql = sql & " where b.codartic = a.codartic"
'        sql = sql & " and c.codcateg = b.codcateg"
        lblinf.Caption = "Proveedores"
        lblinf.Refresh
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
        
        RS.Close
        
        'Si temenos enlace con ariagro, podemos intentar sacar los tratamientos
        
        
        'If vParamAplic.Ariagro <> "" Then
        If BuscarEnSlifacCampos Then
            lblinf.Caption = "Enlace ariagro"
            lblinf.Refresh
            DoEvents
            Set Col = New Collection
            
            'Junio 2014
            cadFecha = " FechaVenta between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            L = 0
            SQL = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then L = DBLet(RS.Fields(0), "N")
            RS.Close
            lblinf.Tag = L
            
            SQL = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            SQL = ""
            While Not RS.EOF
                L = L + 1
                lblinf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblinf.Tag
                lblinf.Refresh
                
                SQL = SQL & ", (" & DBSet(RS!fechaventa, "F") & "," & DBSet(RS.Fields(1), "T") & "," & RS.Fields(2) & ")"
                RS.MoveNext
                
                
                If L > 29 Then
                    Col.Add SQL
                    SQL = ""
                    DoEvents
                    L = 0
                End If
            Wend
            RS.Close
            
            If L > 0 Then Col.Add SQL
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblinf.Caption = "Ariagro " & L & " de " & Col.Count
                lblinf.Refresh
                If (L Mod 5) = 0 Then DoEvents
                SQL = "(" & Mid(Col.item(L), 2) & ")"
                SQL = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & SQL
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    
                    'FAV0079016
                    SQL = " AND numfactura = '" & RS!codtipom & Format(RS!NumFactu, "0000000") & "'"
                    SQL = " WHERE esventa=1 and fechaventa= " & DBSet(RS!FecFactu, "F") & SQL
                    
                    SQL = "UPDATE declaraLOM SET cultivo=" & RS!codCampo & SQL
                    conn.Execute SQL
                    RS.MoveNext
                Wend
                RS.Close
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblinf.Caption = "Obtener variedad"
            lblinf.Refresh
            DoEvents
            SQL = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblinf.Caption = "Campo " & RS!cultivo
                lblinf.Refresh
                SQL = "select rcampos.codcampo,  variedades.nomvarie"
                SQL = SQL & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
                SQL = SQL & " WHERE codcampo =" & RS!cultivo
                
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = "N/D"
                Else
                    SQL = rs2!nomvarie
                End If
                rs2.Close
                SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T") & " WHERE cultivo =" & DBSet(RS!cultivo, "T")
                conn.Execute SQL
                
                RS.MoveNext
            Wend
            RS.Close
            
            
        
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
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblinf.Caption = "Fra: " & RS!NumFactura
                lblinf.Refresh
                SQL = "select * from scafac1 where codtipom='FAS' "
                SQL = SQL & " and fecfactu=" & DBSet(RS!fechaventa, "F") & " and numfactu=" & Mid(RS!NumFactura, 4)
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
                    SQL = SQL & " where fechaventa=" & DBSet(RS!fechaventa, "F") & " and numfactura='" & RS!NumFactura
                    SQL = SQL & "' and lote=" & DBSet(RS!lote, "T") & " and nif=" & DBSet(RS!NIF, "T")
                    SQL = SQL & " and registro=" & DBSet(RS!registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute SQL
                
                End If
                
                RS.MoveNext
            Wend
            RS.Close
            


        
        
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
        lblinf.Caption = "Proceso terminado."
        lblinf.Refresh
        DoEvents
    Else
        MsgBox "NO existen datos entre las fechas", vbExclamation
        RS.Close
    End If
    Set RS = Nothing
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
    
    
    lblinf.Caption = "Preparando datos"
    lblinf.Refresh
    
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
    
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Errores = ""
    
    '1.- Comprobamos que todos los articulos vendidos en el periodo, que deberian tener lote
    SQL = "select codcateg from scateg where ctrlotes = 1"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not RS.EOF
        Aux = Aux & ", " & DBSet(RS!codCateg, "T")
        RS.MoveNext
    Wend
    RS.Close
    
    If Aux = "" Then Stop
    Aux = "(" & Mid(Aux, 2) & ")"
    
    SQL = "select slifac.codartic,slifac.nomartic from  slifac,sartic where slifac.codartic=sartic.codartic "
    SQL = SQL & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " AND codcateg in " & Aux & "  and coalesce(numserie,'')=''"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    Aux = Space(20)
    While Not RS.EOF
        SQL = SQL & Mid(RS!codArtic & Aux, 1, 20) & RS!NomArtic & vbCrLf
        RS.MoveNext
    Wend
    RS.Close
    If SQL <> "" Then
        Aux = "Errores en articulos. No esta indicado el numero de registro" & vbCrLf & String(40, "=") & vbCrLf & SQL
        Errores = Errores & Aux
    End If
    
    
    'Vemos que todos los articulos vendidos en el periodo que deberian tener lote, tienen lote
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute SQL
    
    
    
    lblinf.Caption = "Comprobando lotes"
    lblinf.Refresh
    
    SQL = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifac,sartic WHERE slifac.codartic=sartic.codartic"
    SQL = SQL & " and numserie<>'' and numlote<>'' AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " group by 1,2,3 order by 1,2,3"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    SQL = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifaclotes"
    SQL = SQL & " WHERE fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & "  group by 1,2,3 order by 1,2,3"
    rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RS.EOF
        
        
        'If RS!NumFactu = 67650 Then Stop
        
        fin = False
        LotesCorrectos = False
        MoverRsPpal = True
        If rs2.EOF Then
            
        Else
            If RS!codtipom = rs2!codtipom Then
                If Val(RS!NumFactu) = Val(rs2!NumFactu) Then
                    If RS!FecFactu = rs2!FecFactu Then
                        If RS!canti <= rs2!canti Then LotesCorrectos = True
                    End If
                End If
            End If
            
            
            If LotesCorrectos Then
                rs2.MoveNext
            Else
                
                If RS!codtipom <> rs2!codtipom Then
                    MsgBox "Avise soporte tecnico. Err: codtipom", vbExclamation
                Else
                    If Val(RS!NumFactu) > Val(rs2!NumFactu) Then
                        
                        rs2.MoveNext
                        MoverRsPpal = False
                    Else
                        SQL = RS!codtipom & "|" & RS!NumFactu & "|" & RS!FecFactu & "|"
                        Col.Add SQL
                    End If
                End If
                
                
            End If
        End If
        
        
        If MoverRsPpal Then RS.MoveNext
    Wend
    rs2.Close
    RS.Close
              
    
    
    
    SQL = ""
              
    For L = 1 To Col.Count
        lblinf.Caption = "Lotes FRA" & Col.item(L)
        lblinf.Refresh
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
    lblinf.Caption = "Datos manipulador "
    lblinf.Refresh
    SQL = "select codtipom,numfactu,fecfactu,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet from"
    SQL = SQL & " scafac1 where fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    SQL = SQL & " AND ManipuladorNumCarnet <> ''"
    SQL = SQL & " ORDER BY codtipom,numfactu,fecfactu ,manipuladornumcarnet desc"
    'rs2.Open SQL, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    
    lblinf.Caption = "Obtener registros "
    lblinf.Refresh
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
    SQL = SQL & " order by codartic,h.fechaalb desc"
    
    RS.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
    While Not RS.EOF
                'If Rs!NumFactu = 70679 Then Stop
    
    
                lblinf.Caption = "Ventas: " & RS!codtipom & " " & RS!NumFactu & " " & RS!codArtic
                lblinf.Refresh
                SQL = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet, NifMani)"
                SQL = SQL & " values("
                SQL = SQL & db.Fecha(RS!FechaAlb) & "," ' FechaVenta
                SQL = SQL & db.texto(RS!NomArtic) & "," ' NombreComercial
                SQL = SQL & db.texto(RS!numSerie) & "," ' Registro
                SQL = SQL & db.texto(RS!descateg) & "," ' Categoria
                SQL = SQL & db.texto(RS!numLote) & "," ' Lote
                SQL = SQL & db.numero(RS!Cantidad) & "," ' Cantidad
    
                'ENERO 2016
                LotesCorrectos = DBLet(RS!ManipuladorNombre, "T") <> ""
'
'                If LotesCorrectos Then
'                    'ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet
'                    SQL = SQL & db.texto(RS!ManipuladorNombre) & "," ' NombreSocio
'                Else
'                    SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
'                End If
                SQL = SQL & db.texto(RS!Nomclien) & "," ' NombreSocio
                
                SQL = SQL & db.texto(RS!nifClien) & "," ' NIF
                SQL = SQL & db.texto(RS!codtipom & Format(RS!NumFactu, "0000000")) & "," ' NumFactura
                SQL = SQL & "1,"   ' es vebta
                SQL = SQL & db.texto(RS!Direccion) & "," ' direccion cliente
                SQL = SQL & db.texto(RS!Poblacion) & "," ' poblacion
                
                
 
                'Llevamos tanto el nombre del cliente como el de manipulador
                'NomCarnetMani, NumCarnet, NifMani
                If LotesCorrectos Then
                    'Datos carnet manipulador
                    SQL = SQL & db.texto(RS!ManipuladorNombre) & "," ' NombreSocio
                    SQL = SQL & db.texto(RS!ManipuladorNumCarnet) & ","
                    SQL = SQL & "NULL)"  ' poblacion
                
                Else
                    SQL = SQL & "NULL,NULL,NULL)"  ' poblacion
                End If
                db.ejecutar SQL
        
        
            RS.MoveNext
        Wend
        RS.Close
        

        lblinf.Caption = "Proveedores"
        lblinf.Refresh
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
        
        
        db.ejecutar SQL
        
        
        
        'If vParamAplic.Ariagro <> "" Then
        If BuscarEnSlifacCampos Then
        
            
        
            lblinf.Caption = "Enlace ariagro"
            lblinf.Refresh
            DoEvents
            Set Col = New Collection
            
            'Junio 2014
            cadFecha = " FechaVenta between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            L = 0
            SQL = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then L = DBLet(RS.Fields(0), "N")
            RS.Close
            lblinf.Tag = L
            
            SQL = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            SQL = ""
            While Not RS.EOF
                L = L + 1
                lblinf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblinf.Tag
                lblinf.Refresh
                
                SQL = SQL & ", (" & DBSet(RS!fechaventa, "F") & "," & DBSet(RS.Fields(1), "T") & "," & RS.Fields(2) & ")"
                RS.MoveNext
                
                
                If L > 29 Then
                    Col.Add SQL
                    SQL = ""
                    DoEvents
                    L = 0
                End If
            Wend
            RS.Close
            
            If L > 0 Then Col.Add SQL
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblinf.Caption = "Ariagro " & L & " de " & Col.Count
                lblinf.Refresh
                If (L Mod 5) = 0 Then DoEvents
                SQL = "(" & Mid(Col.item(L), 2) & ")"
                SQL = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & SQL
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    
                    'FAV0079016
                    SQL = " AND numfactura = '" & RS!codtipom & Format(RS!NumFactu, "0000000") & "'"
                    SQL = " WHERE esventa=1 and fechaventa= " & DBSet(RS!FecFactu, "F") & SQL
                    
                    SQL = "UPDATE declaraLOM SET cultivo=" & RS!codCampo & SQL
                    conn.Execute SQL
                    RS.MoveNext
                Wend
                RS.Close
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblinf.Caption = "Obtener variedad"
            lblinf.Refresh
            DoEvents
            SQL = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblinf.Caption = "Campo " & RS!cultivo
                lblinf.Refresh
                SQL = "select rcampos.codcampo,  variedades.nomvarie"
                SQL = SQL & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
                SQL = SQL & " WHERE codcampo =" & RS!cultivo
                
                rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    SQL = "N/D"
                Else
                    SQL = rs2!nomvarie
                End If
                rs2.Close
                SQL = "UPDATE declaralom set cultivo=" & DBSet(SQL, "T") & " WHERE cultivo =" & DBSet(RS!cultivo, "T")
                conn.Execute SQL
                
                RS.MoveNext
            Wend
            RS.Close
            
            
        
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
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblinf.Caption = "Fra: " & RS!NumFactura
                lblinf.Refresh
                SQL = "select * from scafac1 where codtipom='FAS' "
                SQL = SQL & " and fecfactu=" & DBSet(RS!fechaventa, "F") & " and numfactu=" & Mid(RS!NumFactura, 4)
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
                    SQL = SQL & " where fechaventa=" & DBSet(RS!fechaventa, "F") & " and numfactura='" & RS!NumFactura
                    SQL = SQL & "' and lote=" & DBSet(RS!lote, "T") & " and nif=" & DBSet(RS!NIF, "T")
                    SQL = SQL & " and registro=" & DBSet(RS!registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute SQL
                
                End If
                
                RS.MoveNext
            Wend
            RS.Close
            


        
        
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
        lblinf.Caption = "Proceso terminado."
        lblinf.Refresh
        DoEvents
  
    Set RS = Nothing
    Set rs2 = Nothing
End Sub




