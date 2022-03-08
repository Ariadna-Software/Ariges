VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtDeclara 
   Caption         =   "Declarar ROPO"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFechasPlazo 
      Caption         =   "Fechas en plazo"
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
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton cmdCarnetsCaducados 
      Height          =   495
      Left            =   720
      Picture         =   "frmUtDeclara.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Comprobar carnets caducados"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CheckBox chkSoloMostrarErrores 
      Caption         =   "Sólo mostrar errores  // INE"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton cmdCodpobla 
      Height          =   495
      Left            =   120
      Picture         =   "frmUtDeclara.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Poblaciones"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CheckBox chkROPO 
      Caption         =   "R.O.P.O."
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
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ChkSubvencionados 
      Caption         =   "Lotes subvencionados"
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
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox chkTratamientos 
      Caption         =   "Facturas de tratamientos"
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
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtFecha 
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
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
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
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdComenzar 
      Caption         =   "Declaración"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblAriagro 
      AutoSize        =   -1  'True
      Caption         =   "ARIAGRO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   3840
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fechas comunicacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   95
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   2340
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   3240
      Picture         =   "frmUtDeclara.frx":0F8C
      Top             =   480
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   960
      Picture         =   "frmUtDeclara.frx":1017
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta "
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
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde "
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
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   675
   End
   Begin VB.Label lblInf 
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
      Left            =   240
      TabIndex        =   8
      Top             =   3480
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
Dim Sql As String
Dim cantidad As Double
Dim resto As Double


Dim RETO As Boolean
Dim NF As Integer
Dim DesdeAriago As Boolean


Private Sub chkFechasPlazo_KeyPress(KeyAscii As Integer)
 KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub chkROPO_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub



Private Sub chkSoloMostrarErrores_KeyPress(KeyAscii As Integer)
KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub ChkSubvencionados_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub chkTratamientos_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmdCarnetsCaducados_Click()
Dim Varios As Boolean

    Sql = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    frmC.Show vbModal
    Set frmC = Nothing
    If Sql <> "" Then
        Varios = False
        If MsgBox("Añadir clientes de varios?", vbQuestion + vbYesNo) = vbYes Then Varios = True

    
        Screen.MousePointer = vbHourglass
        lblInf.Caption = "Listado x fecha caducidad carnet"
        lblInf.Refresh
        
        
        If HacerListadoCarnets(Varios) Then
                    
            Sql = "|pEmpresa=""" & vEmpresa.nomempre & """|Valores=""Fecha " & Sql & "     Varios: " & IIf(Varios, "Si", "No") & """|"
            
        
            frmVisReport.OtrosParametros = Sql
            frmVisReport.NumeroParametros = 2
    
            frmVisReport.SoloImprimir = False
            frmVisReport.Informe = App.Path & "\Informes\rcarnetsmanipuladoCadu.rpt"
            frmVisReport.CambiaODBC = False
            frmVisReport.FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
            frmVisReport.Show vbModal
        
        
        
        End If
        lblInf.Caption = ""
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdCodpobla_Click()
    frmCPostal.Show vbModal
End Sub

Private Sub cmdComenzar_Click()


 
        


    Screen.MousePointer = vbHourglass
    lblInf.Caption = "Incio proceso"
    lblInf.Refresh
    'RealizarProceso
    
    If RETO Then
        Sql = ""
    Else
        Sql = "A" 'antiguo
        If txtFecha(0).Text <> "" Then
            If CDate(txtFecha(0).Text) >= CDate("01/01/2015") Then Sql = ""
        End If
    End If
    
    
    'Noviembre 2021
    ' Todo es comunicacion ROPO
    
        
    
    If Sql = "A" Then
        'Antiguo. ES ek de Rafa
        NuevoProceso_
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

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '-- Abrimos la base de datos para trabajar con ella
    Set db = New BaseDatos
'    db.abrir "vAriges", "root", "aritel"
    db.asignar conn
    
    db.Tipo = "MYSQL"
    
    
    RETO = True
    Me.Caption = "R.E.T.0."
    '-- Por defecto desde y hasta fecha de hoy
    ObtenerFechas
  
    chkTratamientos.Value = 0
    chkTratamientos.visible = vParamAplic.LlevaADV    'vParamAplic.NumeroInstalacion = vbAlzira
    DesdeAriago = False
    If Not vParamAplic.LlevaADV Then
        If TratamientosDesdeAriagro Then chkTratamientos.visible = True: DesdeAriago = True
    End If
    
    
    
    ChkSubvencionados.Value = 0
    ChkSubvencionados.visible = vParamAplic.LotesGeneralitat
    'If vParamAplic.LotesGeneralitat Then ChkSubvencionados.Value = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set db = Nothing
End Sub



Private Sub ObtenerFechas()
        
        
    Sql = DevuelveDesdeBD(conAri, "fechafin", "declaralom_reto", "1", " 1 ORDER  BY id DESC")
    If Sql = "" Then Sql = Format(Now - 2, "dd/mm/yyyy")
    Sql = DateAdd("d", 1, CDate(Sql))
    
    
    txtFecha(0).Text = Format(Sql, "dd/mm/yyyy")
    txtFecha(1).Text = Format(Now - 1, "dd/mm/yyyy")
    
    
    
End Sub







'****************************************************************************************
' Diciembre 2014.
' Nov 2021. Dberiamos borrar
Private Sub NuevoProceso_()

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

    
    Sql = ""
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        Sql = "Debe indicar las fechas"
    Else
        
        '-- comprobamos que las fechas de paso son as correctas
        If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then Sql = "Fecha inicio mayor que fecha fin"
    End If
    
    If Sql <> "" Then
        MsgBox Sql, vbInformation
        Exit Sub
    End If
    
    
    lblInf.Caption = "Preparando datos"
    lblInf.Refresh
    
    '-- Eliminamos posibles declaraciones anteriores
    Sql = "delete from declaralom"
    db.ejecutar Sql
    
    '-- Antes de empezar y como vamos a hacer uso de canasign, lo limpiamos
    Sql = "update slotes set canasign = 0"
    db.ejecutar Sql
    
    
    BuscarEnSlifacCampos = False
    If vParamAplic.Ariagro <> "" Then
        Sql = " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & " AND 1"
        Sql = DevuelveDesdeBD(conAri, "count(*)", "slifaccampos", Sql, "1")
        If Sql <> "" Then
            If Val(Sql) > 0 Then BuscarEnSlifacCampos = True
        End If
    End If
    
    
    '-- Ahora vamos a por el gran mogollón
    lblInf.Caption = "Obtener lineas facturas"
    lblInf.Refresh
    Sql = "select a.codtipom, a.numfactu, a.fecfactu, a.codartic, a.nomartic, a.cantidad "
     Sql = Sql & ",b.nomclien, b.nifclien,b.domclien direccion,concat(codpobla,' ',pobclien) poblacion ,d.descateg"
    Sql = Sql & " from slifac as a, scafac as b, sartic as c, scateg as d"
    
    Sql = Sql & " where a.codartic in"
    Sql = Sql & " (select codartic from sartic"
    Sql = Sql & " where codcateg in (select codcateg from scateg where ctrlotes = 1))"
    Sql = Sql & " and a.cantidad <> 0 "
    Sql = Sql & " and b.codtipom = a.codtipom"
    Sql = Sql & " and b.numfactu = a.numfactu"
    Sql = Sql & " and b.fecfactu = a.fecfactu"
    Sql = Sql & " and c.codartic = a.codartic"
    Sql = Sql & " and d.codcateg = c.codcateg"
    Sql = Sql & " and a.fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " order by codartic,a.fecfactu desc "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    DoEvents
    
    If Not RS.EOF Then
    
        'Ahora vamos a contar los que hay
        L = 0
        While Not RS.EOF
            RS.MoveNext
            L = L + 1
        Wend
        RS.MoveFirst
        
        
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
            
            
                
            
            If RS!codArtic <> ArticuloATratar Then
                'OK. NUEVO ARTICULO
                If ArticuloATratar <> "" Then
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        Sql = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        Sql = Sql & " where codartic = " & db.texto(rs2!codArtic)
                        Sql = Sql & " and numlotes = " & db.texto(rs2!numlotes)
                        Sql = Sql & " and fecentra = " & db.Fecha(rs2!fecentra)
                        db.ejecutar Sql
                    End If
                
                    rs2.Close
                End If
                ArticuloATratar = RS!codArtic
                
                
                
                HaMovidoLinFactura = True
                Sql = "select a.codartic, a.numlotes, a.fecentra, a.canentra, a.canasign, b.numserie from slotes as a, sartic as b" & _
                    " where a.codartic = " & db.texto(RS!codArtic) & _
                    " and (a.canentra - a.canasign > 0)" & _
                    " and a.fecentra <= " & db.Fecha(RS!FecFactu) & _
                    " and b.codartic = a.codartic" & _
                    " order by a.fecentra desc"
            
                Set rs2 = db.cursor(Sql)
                
                UtilizadaEnLote = 0
                CantidadQuedaEnLote = 0
                If Not rs2.EOF Then
                    NumeroDeLote = rs2!numlotes
                    CantidadQuedaEnLote = rs2!canentra
                End If
                
            End If
            
            If HaMovidoLinFactura Then
                cantidad = RS!cantidad
                resto = cantidad
                HaMovidoLinFactura = False
            End If
            
            If rs2.EOF Then
                'NO HAY MAS LOTES
                Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                Sql = Sql & " values("
                Sql = Sql & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                Sql = Sql & db.texto(RS!NomArtic) & "," ' NombreComercial
                Sql = Sql & db.texto(" ") & "," ' Registro
                Sql = Sql & db.texto(RS!descateg) & "," ' Categoria
                Sql = Sql & db.texto(" ") & "," ' Lote
                Sql = Sql & db.numero(resto) & "," ' Cantidad
                Sql = Sql & db.texto(RS!NomClien) & "," ' NombreSocio
                Sql = Sql & db.texto(RS!nifClien) & "," ' NIF
                Sql = Sql & db.texto(RS!codtipom & Format(RS!Numfactu, "0000000")) & "," ' NumFactura
                'octubre 2011 EsVenta,Direccion,Poblacion
                            Sql = Sql & "1,"   ' es venta
                            Sql = Sql & db.texto(RS!Direccion) & "," ' direccion cliente
                            Sql = Sql & db.texto(RS!Poblacion) & ")" ' poblacion
                
                
                db.ejecutar Sql
                HaMovidoLinFactura = True
                
            
            Else
                If rs2!fecentra > RS!FecFactu Then
                    'Cantidad utilizada
                    If UtilizadaEnLote > 0 Then
                        'UPDATE ENNUmero de lote en canasign
                        Sql = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                        Sql = Sql & " where codartic = " & db.texto(rs2!codArtic)
                        Sql = Sql & " and numlotes = " & db.texto(rs2!numlotes)
                        Sql = Sql & " and fecentra = " & db.Fecha(rs2!fecentra)
                        db.ejecutar Sql
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
                            
                            Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                            Sql = Sql & " values("
                            Sql = Sql & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                            Sql = Sql & db.texto(RS!NomArtic) & "," ' NombreComercial
                            Sql = Sql & db.texto(rs2!numSerie) & "," ' Registro
                            Sql = Sql & db.texto(RS!descateg) & "," ' Categoria
                            Sql = Sql & db.texto(rs2!numlotes) & "," ' Lote
                            Sql = Sql & TransformaComasPuntos(db.numero(resto)) & "," ' Cantidad
                            Sql = Sql & db.texto(RS!NomClien) & "," ' NombreSocio
                            Sql = Sql & db.texto(RS!nifClien) & "," ' NIF
                            Sql = Sql & db.texto(RS!codtipom & Format(RS!Numfactu, "0000000")) & "," ' NumFactura
                            
                            'octubre 2011 EsVenta,Direccion,Poblacion
                            Sql = Sql & "1,"   ' es venta
                            Sql = Sql & db.texto(RS!Direccion) & "," ' direccion cliente
                            Sql = Sql & db.texto(RS!Poblacion) & ")" ' poblacion
                            db.ejecutar Sql
                            
                            HaMovidoLinFactura = True
                        Else
                            
                            'Quedaba un poco en el lote
                            If CantidadQuedaEnLote > 0 Then
                                Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion)"
                                Sql = Sql & " values("
                                Sql = Sql & db.Fecha(RS!FecFactu) & "," ' FechaVenta
                                Sql = Sql & db.texto(RS!NomArtic) & "," ' NombreComercial
                                Sql = Sql & db.texto(rs2!numSerie) & "," ' Registro
                                Sql = Sql & db.texto(RS!descateg) & "," ' Categoria
                                Sql = Sql & db.texto(rs2!numlotes) & "," ' Lote
                                Sql = Sql & TransformaComasPuntos(db.numero(CantidadQuedaEnLote)) & ","  ' Cantidad
                                Sql = Sql & db.texto(RS!NomClien) & "," ' NombreSocio
                                Sql = Sql & db.texto(RS!nifClien) & "," ' NIF
                                Sql = Sql & db.texto(RS!codtipom & Format(RS!Numfactu, "0000000")) & ","
                                'octubre 2011 EsVenta,Direccion,Poblacion
                                Sql = Sql & "1,"   ' es venta
                                Sql = Sql & db.texto(RS!Direccion) & "," ' direccion cliente
                                Sql = Sql & db.texto(RS!Poblacion) & ")" ' poblacion
                                
                                db.ejecutar Sql
                                resto = resto - CantidadQuedaEnLote  'nos queda "resto por asignar
                                UtilizadaEnLote = UtilizadaEnLote + CantidadQuedaEnLote
                            End If
                            
                            Sql = "update slotes set canasign = " & TransformaComasPuntos(CStr(UtilizadaEnLote))
                            Sql = Sql & " where codartic = " & db.texto(rs2!codArtic)
                            Sql = Sql & " and numlotes = " & db.texto(rs2!numlotes)
                            Sql = Sql & " and fecentra = " & db.Fecha(rs2!fecentra)
                            db.ejecutar Sql
                                
                            
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
        lblInf.Caption = "Proveedores"
        lblInf.Refresh
        DoEvents
        Sql = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,Poblacion)"
        Sql = Sql & "select distinct a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, e.nomprove, e.nifprove, d.document, a.canentra" & _
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
        
        
        db.ejecutar Sql
        
        RS.Close
        
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
            Sql = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then L = DBLet(RS.Fields(0), "N")
            RS.Close
            lblInf.Tag = L
            
            Sql = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            Sql = ""
            While Not RS.EOF
                L = L + 1
                lblInf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblInf.Tag
                lblInf.Refresh
                
                Sql = Sql & ", (" & DBSet(RS!fechaventa, "F") & "," & DBSet(RS.Fields(1), "T") & "," & RS.Fields(2) & ")"
                RS.MoveNext
                
                
                If L > 29 Then
                    Col.Add Sql
                    Sql = ""
                    DoEvents
                    L = 0
                End If
            Wend
            RS.Close
            
            If L > 0 Then Col.Add Sql
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count
                lblInf.Refresh
                If (L Mod 5) = 0 Then DoEvents
                Sql = "(" & Mid(Col.Item(L), 2) & ")"
                Sql = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & Sql
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    
                    'FAV0079016
                    Sql = " AND numfactura = '" & RS!codtipom & Format(RS!Numfactu, "0000000") & "'"
                    Sql = " WHERE esventa=1 and fechaventa= " & DBSet(RS!FecFactu, "F") & Sql
                    
                    Sql = "UPDATE declaraLOM SET cultivo=" & RS!codCampo & Sql
                    conn.Execute Sql
                    RS.MoveNext
                Wend
                RS.Close
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblInf.Caption = "Obtener variedad"
            lblInf.Refresh
            DoEvents
            Sql = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblInf.Caption = "Campo " & RS!cultivo
                lblInf.Refresh
                Sql = "select rcampos.codcampo,  variedades.nomvarie"
                Sql = Sql & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                Sql = Replace(Sql, "@#", vParamAplic.Ariagro & ".")
                Sql = Sql & " WHERE codcampo =" & RS!cultivo
                
                rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    Sql = "N/D"
                Else
                    Sql = rs2!nomvarie
                End If
                rs2.Close
                Sql = "UPDATE declaralom set cultivo=" & DBSet(Sql, "T") & " WHERE cultivo =" & DBSet(RS!cultivo, "T")
                conn.Execute Sql
                
                RS.MoveNext
            Wend
            RS.Close
            
            
        
        End If 'de ariagro
        
        'Abril 2015
        'ALZIRA
        If vParamAplic.NumeroInstalacion = vbAlzira Then
            'Para aquellas facturas de servicio (que son tratamientos), si no esta indicado el cultivo, ni  la variedad
            'entonces UPDATEAMOS con los datos de la observacion
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            Sql = "select fechaventa, NombreComercial,Registro,Categoria,Lote,NIF,NumFactura"
            Sql = Sql & " from declaralom where esventa=1 and numfactura like 'FAS%' and cultivo is null and tratamiento is null"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblInf.Caption = "Fra: " & RS!NumFactura
                lblInf.Refresh
                Sql = "select * from scafac1 where codtipom='FAS' "
                Sql = Sql & " and fecfactu=" & DBSet(RS!fechaventa, "F") & " and numfactu=" & Mid(RS!NumFactura, 4)
                rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    Sql = ""
                Else
                    Sql = Trim(DBLet(rs2!observa1, "T"))
                End If
                rs2.Close
                
                If Sql <> "" Then
                    NumRegElim = 1
                    
                    'Vamos a quitar todos los espacios en blanco "duplicados"
                    Do
                        NumRegElim = InStr(NumRegElim, Sql, " ")
                        If NumRegElim > 0 Then
                            Do
                                L = InStr(NumRegElim + 1, Sql, " ")
                                If L = NumRegElim + 1 Then
                                    Sql = Mid(Sql, 1, L - 1) & Mid(Sql, L + 1)
                                Else
                                    L = 0
                                End If
                            Loop Until L = 0
                            NumRegElim = NumRegElim + 1
                        End If
                    Loop Until NumRegElim = 0
                            
                    
                
                
                    L = Len(Sql)
                    If L > 45 Then
                        cadFecha = Mid(Sql, 46)
                        Sql = Mid(Sql, 1, 45)
                    Else
                        cadFecha = ""
                    End If
                    Sql = "UPDATE declaralom set cultivo=" & DBSet(Sql, "T")
                    Sql = Sql & ",tratamiento= " & DBSet(cadFecha, "T", "S")
                    Sql = Sql & " where fechaventa=" & DBSet(RS!fechaventa, "F") & " and numfactura='" & RS!NumFactura
                    Sql = Sql & "' and lote=" & DBSet(RS!Lote, "T") & " and nif=" & DBSet(RS!NIF, "T")
                    Sql = Sql & " and registro=" & DBSet(RS!Registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute Sql
                
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
        frmVisReport.CambiaODBC = False
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
        RS.Close
    End If
    Set RS = Nothing
    Set rs2 = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Sql = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)

   
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   Sql = ""
   frmC.Show vbModal
   Set frmC = Nothing
    If Sql <> "" Then txtFecha(Index).Text = Sql

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



'****************************************************************************************
' Noviembre 2015
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

'NOVIEMBRE 2021
'  RETO.    Registro de trnasaccionciones


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
Dim Sql_Servicios As String
Dim GrabarRegistro As Boolean
Dim LetraFAS As String  'para ver que rectificativas son de FAS que entonces NO se graban en venta
Dim CadenaAux As String
Dim cantidad As Currency
Dim Capacidad As Currency
Dim Volumen As Currency

Dim rsTipUd As ADODB.Recordset
Dim VtaPorUnidades As Boolean

Dim ParaVerDatosINE As String

Dim Llevatratamientos As Byte  ' 0. NO      1.- Desde ariges(Alzira Cata)     2 Desde ariagro.
Dim CADENA As String
Dim F As Date
Dim F2 As Date

    On Error GoTo eProcesoDesdeSlifac

    ParaVerDatosINE = ""

    Sql = ""
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        Sql = "Debe indicar las fechas"
    Else
        
        '-- comprobamos que las fechas de paso son as correctas
        If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then Sql = "Fecha inicio mayor que fecha fin"
        
    End If
    
    If Sql <> "" Then
        MsgBox Sql, vbInformation
        Exit Sub
    End If
    
    
    If Me.chkROPO.Value = 0 Then
        If CDate(txtFecha(0).Text) < CDate("01/11/2021") Then
            If MsgBox("Fecha anterior a inicio presentacion RETO" & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    lblInf.Caption = "Preparando datos   " & IIf(chkSoloMostrarErrores.Value = 1, "CHECK", "")
    lblInf.Refresh
    
    '-- Eliminamos posibles declaraciones anteriores
    Sql = "delete from declaralom"
    db.ejecutar Sql
    
    '-- No vamos a hacer uso de canasign, lo limpiamos de todas formas
    Sql = "update slotes set canasign = 0"
    db.ejecutar Sql
    
    
    BuscarEnSlifacCampos = False
    If vParamAplic.Ariagro <> "" Then
        Sql = " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & " AND 1"
        Sql = DevuelveDesdeBD(conAri, "count(*)", "slifaccampos", Sql, "1")
        If Sql <> "" Then
            If Val(Sql) > 0 Then BuscarEnSlifacCampos = True
        End If
    End If
    
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Errores = ""
    
    '1.- Comprobamos que todos los articulos vendidos en el periodo, que deberian tener lote
    Sql = "select codcateg from scateg where ctrlotes = 1"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not RS.EOF
        Aux = Aux & ", " & DBSet(RS!codCateg, "T")
        RS.MoveNext
    Wend
    RS.Close
    
    If Aux = "" Then
        MsgBox "Categorias sin control de lotes", vbExclamation
        Exit Sub
    End If
        
    Aux = "(" & Mid(Aux, 2) & ")"  'NO TOCAR AUX hasta el final de las comprobaciones
    
    Sql = "select distinct slifac.codartic,slifac.nomartic from  slifac,sartic where slifac.codartic=sartic.codartic "
    Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " AND codcateg in " & Aux & "  and coalesce(numserie,'')=''"
    
        
    
    If chkTratamientos.visible Then
        
        
        
        
        
        
        
        
        Sql_Servicios = "  "
        If Me.chkTratamientos.Value = 0 Then Sql_Servicios = " NOT "
        Sql_Servicios = Sql_Servicios & " slifac.codtipom IN ('FAS','FAI')"
        Sql = Sql & " AND " & Sql_Servicios
        
        '
                
                
                
                
                
                
      
        If DesdeAriago And chkTratamientos.Value = 1 Then
        
        
            If chkSoloMostrarErrores.Value Then
                'Es ver errores y datos INE, con lo cual mostraré todos los articulos de ariagro que tienen puesto producto y estan enlazadas o no
                CADENA = "select lin.codartic artagro, ges.codartic,art.nomartic nomagro,ges.nomartic, min(numfactu) minfac,count(*) cuantos"
                CADENA = CADENA & " , codcateg, if(codcateg in " & Aux & ",'Si','NO') lotes,coalesce(ges.numserie,'') nser"
                CADENA = CADENA & " from ariagro.advfacturas_lineas lin inner join ariagro.advartic art on art.codartic=lin.codartic"
                CADENA = CADENA & " left join sartic ges on ges.codartic=art.codartic  WHERE tipoprod=0 and art.numserie<>'' "
                CADENA = CADENA & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                CADENA = CADENA & " group by  lin.codartic order by codcateg,ges.codartic,lin.codartic,ges.numserie"
        
                RS.Open CADENA, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                CADENA = ""
                While Not RS.EOF
                    '        Agro    Art ariges
                    CADENA = CADENA & Mid(RS!artagro & Space(20), 1, 20) & Mid(RS!nomagro & Space(35), 1, 35)
                    If IsNull(RS!codArtic) Then
                        'No esta vinculado
                        CADENA = CADENA & Space(95) & vbCrLf
                        
                    Else
                        CADENA = CADENA & Mid(RS!codArtic & Space(20), 1, 20) & Mid(RS!NomArtic & Space(35), 1, 35) & "  "
                        CADENA = CADENA & Mid(RS!minfac & Space(10), 1, 10) & Mid(RS!Cuantos & Space(5), 1, 5)
                        CADENA = CADENA & Mid(RS!codCateg & Space(5), 1, 5) & "  " & Mid(RS!lotes & Space(5), 1, 5)
                        CADENA = CADENA & Mid(RS!nser & Space(15), 1, 15) & vbCrLf
                    End If
                                        
                    RS.MoveNext
                Wend
                RS.Close
                If CADENA <> "" Then
                    Errores = Errores & vbCrLf & "Vinculacion Ariges - Ariagro " & vbCrLf
                    Errores = Errores & "                  Articulo Ariagro                          Articulo Ariges       "
                    Errores = Errores & Space(25) & "Fra(min)  Cuantos  Categ  Lotes   Serie" & vbCrLf & String(150, "=") & vbCrLf
                    Errores = Errores & CADENA
                End If
        
            End If
     
     
     
            'Comprobacion numeros de serie
      
            Sql = "select lin.codartic, ges.nomartic"
            Sql = Sql & " from ariagro.advfacturas_lineas lin inner join ariagro.advartic art on art.codartic=lin.codartic"
            Sql = Sql & " inner join sartic ges on ges.codartic=art.codartic  WHERE tipoprod=0 and art.numserie<>''"
            Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            Sql = Sql & " AND codcateg in " & Aux & "  and coalesce(ges.numserie,'')=''"
            Sql = Sql & " group by  lin.codartic "
             
     
        End If
    End If
    
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not RS.EOF
        Sql = Sql & Mid(RS!codArtic & Space(20), 1, 20) & RS!NomArtic & vbCrLf
        RS.MoveNext
    Wend
    RS.Close
    If Sql <> "" Then
        Sql = "Errores en articulos. No está indicado el número de registro" & vbCrLf & String(40, "=") & vbCrLf & Sql
        Errores = Errores & Sql
    End If
    
    
    Sql_Servicios = " "
    If Me.chkTratamientos.Value = 0 Then Sql_Servicios = " NOT "
    Sql_Servicios = " AND " & Sql_Servicios & "  slifac.codtipom IN ('FAS','FAI')"

    
    'Noviembre 2021
    Sql = ""
    If chkROPO.Value = 0 Then    'solo RETO
         Sql = "select distinct slifac.codartic,slifac.nomartic from  slifac,sartic where slifac.codartic=sartic.codartic "
         Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
         Sql = Sql & " AND codcateg in " & Aux & "  and numserie<>'' and (unidadescompra =0 or  coalesce(unicajas2,0)=0)"
         If Sql_Servicios <> "" Then Sql = Sql & Sql_Servicios
         
         
         If DesdeAriago And chkTratamientos.Value = 1 Then
                
                 'los articulos vienen
                Sql = "select lin.codartic, ges.nomartic"
                Sql = Sql & " from ariagro.advfacturas_lineas lin inner join ariagro.advartic art on art.codartic=lin.codartic"
                Sql = Sql & " inner join sartic ges on ges.codartic=art.codartic  WHERE tipoprod=0 and art.numserie<>''"
                Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                Sql = Sql & " AND codcateg in " & Aux & "  and ges.numserie<>'' and (unidadescompra =0 or  coalesce(unicajas2,0)=0)"
                Sql = Sql & " group by  lin.codartic "
         End If
         
         
         RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         Sql = ""
        
         While Not RS.EOF
             Sql = Sql & Mid(RS!codArtic & "                 ", 1, 20) & RS!NomArtic & vbCrLf
             RS.MoveNext
         Wend
         RS.Close
    End If
    If Sql <> "" Then
        Sql = "Errores en articulos. No esta indicada la capacidad / unidad " & vbCrLf & String(40, "=") & vbCrLf & Sql
        Errores = Errores & Sql
    End If
    
    
    
    'Comprobar INE de poblacion
    lblInf.Caption = "INE / poblacion"
    lblInf.Refresh
    
    Sql = "    select codtipom,numfactu,fecfactu from  slifac,sartic where slifac.codartic=sartic.codartic"
    Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " AND codcateg in " & Aux & "  and numserie<>''   "
    If Sql_Servicios <> "" Then
        Sql = Sql & Sql_Servicios
        'En las facturas internas, el FAI no debemos comprobar el INE. Grabara el del CLIENTE
        If Me.chkTratamientos.Value = 1 Then Sql = Sql & " AND slifac.codtipom<>'FAI' "
    End If
    If chkROPO.Value = 1 Then Sql = Sql & " AND FALSE"
    
    Sql = Sql & "  group by 1,2,3 "
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    davidNumalbar = 20
    NumRegElim = 0
    Sql = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        Sql = Sql & ", (" & DBSet(RS!codtipom, "T") & "," & RS!Numfactu & "," & DBSet(RS!FecFactu, "F") & ")"
        RS.MoveNext
        If RS.EOF Then NumRegElim = davidNumalbar + 1
        
        If NumRegElim > davidNumalbar Then
            Col.Add "(" & Mid(Sql, 2) & ")"
            NumRegElim = 0
        End If
    Wend
    RS.Close
    
    
    davidCodtipom = ""
    For NumRegElim = 1 To Col.Count
        lblInf.Caption = "INE / poblacion. (II) " & NumRegElim & " de " & Col.Count
        lblInf.Refresh
    
        Sql = "select scafac.codtipom,scafac.numfactu, scafac.codclien ,scafac.codpobla  from scafac left join scpostal on scafac.codpobla=scpostal.cpostal "
        Sql = Sql & " WHERE fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F") & " AND ine is null"
        Sql = Sql & " AND (codtipom,numfactu,fecfactu) in " & Col.Item(NumRegElim)
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            Sql = "#" & RS!codpobla & "#"
            If InStr(1, davidCodtipom, Sql) = 0 Then davidCodtipom = davidCodtipom & Sql & "  " & RS!codtipom & RS!Numfactu & " cliente:" & RS!codClien
           
            RS.MoveNext
        Wend
        RS.Close
    Next
    
    If davidCodtipom <> "" Then
        Sql = "Errores codigos INE - poblacion " & vbCrLf & String(40, "=") & vbCrLf & davidCodtipom
        Errores = Errores & Sql
    End If
    
    Set Col = Nothing
    
    
    If chkROPO.Value = 0 Then    'solo RETO
         If DesdeAriago And chkTratamientos.Value = 1 Then
                
            lblInf.Caption = "Comprobando lotes ariagro "
            lblInf.Refresh
    
            'los articulos vienen
            Sql = "select lin.codartic,lin.ampliaci, ges.nomartic "
            Sql = Sql & " from ariagro.advfacturas_lineas lin inner join ariagro.advartic art on art.codartic=lin.codartic"
            Sql = Sql & " inner join sartic ges on ges.codartic=art.codartic  WHERE tipoprod=0 and art.numserie<>''"
            Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            Sql = Sql & " AND codcateg in " & Aux & "  and ges.numserie<>'' "
            Sql = Sql & " group by lin.codartic,lin.ampliaci order by 1,2"
            
            Set Col = New Collection
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
            
            While Not RS.EOF
                If DBLet(RS!Ampliaci, "T") = "" Then
                    Sql = Sql & Mid(RS!codArtic & "                 ", 1, 20) & RS!NomArtic & vbCrLf
                Else
                    Col.Add RS!codArtic & "|" & RS!Ampliaci & "|"
                End If
                RS.MoveNext
            Wend
            RS.Close
            If Sql <> "" Then
                If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
                Sql = "Errores  lotes sin asignar Ariagro" & vbCrLf & String(40, "=") & vbCrLf & Sql
                Errores = Errores & Sql
            End If
            
            
            CADENA = ""
            For NumRegElim = 1 To Col.Count
                lblInf.Caption = "Lotes " & NumRegElim & " de " & Col.Count
                lblInf.Refresh
            
                Sql = "codartic= " & DBSet(RecuperaValor(Col.Item(NumRegElim), 1), "T") & " and numlotes = " & DBSet(RecuperaValor(Col.Item(NumRegElim), 2), "T") & " AND 1"
                Sql = DevuelveDesdeBD(conAri, "numlotes", "slotes", Sql, "1")
                If Sql = "" Then CADENA = CADENA & Mid(RecuperaValor(Col.Item(NumRegElim), 1) & Space(20), 1, 20) & RecuperaValor(Col.Item(NumRegElim), 2) & vbCrLf
                    
            Next
            If CADENA <> "" Then
                If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
                CADENA = "Errores  lotes NO existe Ariges" & vbCrLf & "Articulo           Lote" & vbCrLf & String(40, "=") & vbCrLf & CADENA
                Errores = Errores & CADENA
            End If
         End If
    End If
    
    
    
    
    
    'Vemos que todos los articulos vendidos en el periodo que deberian tener lote, tienen lote
    Sql = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    
    
    
    
    Sql = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifac,sartic WHERE slifac.codartic=sartic.codartic"
    Sql = Sql & " and numserie<>'' and numlote<>'' AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    If Sql_Servicios <> "" Then Sql = Sql & Sql_Servicios
    Sql = Sql & " group by 1,2,3 order by 1,2,3"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    Sql = "select codtipom, numfactu, fecfactu,sum(cantidad) as canti from slifaclotes"
    Sql = Sql & " WHERE fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    If Sql_Servicios <> "" Then Sql = Sql & Replace(Sql_Servicios, "slifac", "slifaclotes")
    Sql = Sql & "  group by 1,2,3 order by 1,2,3"
    rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RS.EOF
        
        
        
        fin = False
        LotesCorrectos = False
        MoverRsPpal = True
        If rs2.EOF Then
            
        Else
            If RS!codtipom = rs2!codtipom Then
                If Val(RS!Numfactu) = Val(rs2!Numfactu) Then
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
                    If Val(RS!Numfactu) > Val(rs2!Numfactu) Then
                        
                        rs2.MoveNext
                        MoverRsPpal = False
                    Else
                        Sql = RS!codtipom & "|" & RS!Numfactu & "|" & RS!FecFactu & "|"
                        Col.Add Sql
                    End If
                End If
                
                
            End If
        End If
        
        
        If MoverRsPpal Then RS.MoveNext
    Wend
    rs2.Close
    RS.Close
              
    
    
    
    Sql = ""
              
    For L = 1 To Col.Count
        lblInf.Caption = "Lotes FRA" & Col.Item(L)
        lblInf.Refresh
        Debug.Print Col.Item(L)
        Aux = ", ('" & RecuperaValor(Col.Item(L), 1) & "'," & RecuperaValor(Col.Item(L), 2) & "," & DBSet(RecuperaValor(Col.Item(L), 3), "F") & ")"
        Sql = Sql & Aux & vbCrLf
    Next
    If Sql <> "" Then
        If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
        Aux = "Errores en lotes. Facturas no coinciden lotes (factura/asignados) " & vbCrLf & String(40, "=") & vbCrLf & Sql
        Errores = Errores & Aux
    End If
   
   
   
    lblInf.Caption = "Comprobando lotes"
    lblInf.Refresh
   
   
   
   
    
    LetraFAS = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAS", "T")
    If LetraFAS = "" Then LetraFAS = "X@@"
    
    '-- Ahora vamos a por el gran mogollón
    lblInf.Caption = "Datos manipulador "
    lblInf.Refresh
    Sql = "select codtipom,numfactu,fecfactu,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet from"
    'SQL = SQL & " scafac1 where fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " scafac1 where fechalb between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " AND ManipuladorNumCarnet <> ''"
    If Sql_Servicios <> "" Then Sql = Sql & Replace(Sql_Servicios, "slifac", "scafac1")
    Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu ,manipuladornumcarnet desc"
    'rs2.Open SQL, conn, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    
    lblInf.Caption = "Obtener registros "
    lblInf.Refresh
    
    
    Set rsTipUd = New ADODB.Recordset
    rsTipUd.Open "select * from stipudcompra ", conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Sql = " select a.codtipom, a.numfactu, h.fechaalb, a.codartic, c.nomartic, a.cantidad ,b.nomclien, b.nifclien,"
    Sql = Sql & " b.domclien direccion,concat(codpobla,' ',pobclien) poblacion ,d.descateg , numlote,numserie"
    Sql = Sql & " ,ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet, b.codclien"
    'Noviembre
    If Me.chkROPO.Value = 1 Then
        Sql = Sql & " ,1 unicajas2, 1 unidadescompra , codpobla"
    Else
        Sql = Sql & " ,unicajas2,unidadescompra , codpobla"
    End If
    Sql = Sql & " , a.fecfactu "
    Sql = Sql & " From slifaclotes as a, scafac as b,scafac1 h,sartic as c, scateg as d  "
    Sql = Sql & " WHERE c.numserie<>''  AND ctrlotes = 1 and a.cantidad <> 0 and c.codartic = a.codartic and d.codcateg = c.codcateg"
    Sql = Sql & " AND h.fechaalb between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
    Sql = Sql & " AND b.codtipom = h.codtipom AND b.numfactu = h.numfactu and b.fecfactu = h.fecfactu"
    Sql = Sql & " AND b.codtipom = a.codtipom AND b.numfactu = a.numfactu and b.fecfactu = a.fecfactu"
    Sql = Sql & " AND a.codtipoa = h.codtipoa AND a.numalbar = h.numalbar"
    
    If Sql_Servicios <> "" Then Sql = Sql & Replace(Sql_Servicios, "slifac", "a")
    Sql = Sql & " order by codartic,h.fechaalb desc"
    
    If chkROPO.Value = 0 Then    'solo RETO
        If DesdeAriago And chkTratamientos.Value = 1 Then
            Sql = "select lin.codtipom,lin.numfactu,lin.fecfactu fechaalb,lin.codartic,art.nomartic,lin.cantidad," & DBSet(vEmpresa.nomempre, "T") & " as NomClien,"
            Sql = Sql & DBSet(vParam.CifEmpresa, "T") & " nifclien," & DBSet(vParam.DomicilioEmpresa, "T") & " direccion,"
            Sql = Sql & DBSet(vParam.CPostal & " " & vParam.Poblacion, "T") & " poblacion," & DBSet(vParam.DomicilioEmpresa, "T") & " as direccion"
            Sql = Sql & " , descateg , ampliaci numLote, ges.numserie"
            Sql = Sql & " , '' ManipuladorNumCarnet, '' ManipuladorFecCaducidad, '' ManipuladorNombre,1 TipoCarnet, 1 codclien ,unicajas2,unidadescompra"
            Sql = Sql & " , '" & vParam.CPostal & "' codpobla , lin.fecfactu"
            Sql = Sql & " from ariagro.advfacturas_lineas lin inner join ariagro.advartic art on art.codartic=lin.codartic"
            Sql = Sql & " inner join sartic ges on ges.codartic=art.codartic"
            Sql = Sql & " inner join scateg cate on ges.codcateg = cate.codcateg"
            Sql = Sql & " WHERE tipoprod=0 and art.numserie<>''"
            Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            Sql = Sql & " AND ges.codcateg in " & Aux & "   and ges.numserie<>''  order by lin.codartic,lin.fecfactu"

        End If
    End If
    
    RS.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    
    DoEvents
    
    davidNumalbar = 1000
    Sql = ""
    Capacidad = 1   'Para el ROPO da igual
    Volumen = 1   'Para el ROPO da igual
    CADENA = ""
    While Not RS.EOF
    
                'En declara.email pondremos codclien
                lblInf.Caption = "Ventas: " & RS!codtipom & " " & RS!Numfactu & " " & RS!codArtic
                lblInf.Refresh
                
                GrabarRegistro = True
                If RS!codtipom = "FRT" Then
                    CadenaAux = "codtipom =" & DBSet(RS!codtipom, "T") & " AND fecfactu=" & DBSet(RS!FecFactu, "F") & " AND numfactu"
                    CadenaAux = DevuelveDesdeBD(conAri, "observa1", "scafac1", CadenaAux, RS!Numfactu)
                    If InStr(1, CadenaAux, LetraFAS) > 0 Then GrabarRegistro = False 'es una rectificativa de servicios
                End If
                
                If GrabarRegistro Then
                    Sql = Sql & ",  ("
                    Sql = Sql & db.Fecha(RS!FechaAlb) & "," ' FechaVenta
                    Sql = Sql & db.texto(RS!NomArtic) & "," ' NombreComercial
                    Sql = Sql & db.texto(RS!numSerie) & "," ' Registro
                    Sql = Sql & db.texto(RS!descateg) & "," ' Categoria
                    Sql = Sql & db.texto(RS!numLote) & "," ' Lote
                    
                    
                    If Me.chkROPO.Value = 1 Then
                        cantidad = RS!cantidad
                        
                    Else
                    
                        'RETO
                        
                        rsTipUd.Find "tipcompra =" & DBLet(RS!unidadesCompra, "N"), , adSearchForward, 1
                        
                        'Capacidad
                        Capacidad = DBLet(RS!unicajas2, "N")
                        If Capacidad = 0 Then Capacidad = 1
                        
                        
                        VtaPorUnidades = True
                        If Me.chkTratamientos.Value = 0 Then
                            If Not rsTipUd.EOF Then VtaPorUnidades = rsTipUd!vtaindicacantidad = 0
                        Else
                            'En tratamientos, la cantidad que "Echan" es la real. Si pone 1.5 significa que han hechado 1.5
                            Capacidad = 1
                        End If
                        If rsTipUd.EOF Then Debug.Assert False
                        
                        
                        If VtaPorUnidades Then
                            cantidad = RS!cantidad
                            If chkTratamientos.Value = 0 Then
                                If cantidad > 0 And Int(cantidad) <> cantidad Then
                                    'NO permite VENTAS decimales
                                    If Me.chkSoloMostrarErrores.Value Then CADENA = CADENA & Mid(RS!codtipom & RS!Numfactu & Space(20), 1, 20) & "  " & RS!cantidad & vbCrLf
                                    cantidad = Int(cantidad)
                                    If cantidad = 0 Then cantidad = 1
                        
                                    
                                End If
                            End If
                            Volumen = cantidad * Capacidad
                            If chkTratamientos.Value = 1 Then cantidad = 1
                            
                        Else
                            'Si pone 5, esta vendiendo 5 Litros. Si el envase es de 5Lts, esta vendiedndo 1 Unidad
                            If cantidad > 0 And Int(cantidad) <> cantidad Then Debug.Assert False
                            Volumen = RS!cantidad
                            If Me.chkTratamientos.Value = 1 Then
                                cantidad = Round2(Volumen \ Capacidad, 2)
                            Else
                                cantidad = Round2(Volumen \ Capacidad, 0)
                            End If
                            If cantidad = 0 Then cantidad = 1
                        End If
                    End If
                    
                    Sql = Sql & db.numero(cantidad) & "," ' Cantidad
        
                    'ENERO 2016
                    LotesCorrectos = DBLet(RS!ManipuladorNombre, "T") <> ""
    
                    Sql = Sql & db.texto(RS!NomClien) & "," ' NombreSocio
                    
                    Sql = Sql & db.texto(RS!nifClien) & "," ' NIF
                    Sql = Sql & db.texto(RS!codtipom & Format(RS!Numfactu, "0000000")) & "," ' NumFactura
                    Sql = Sql & "1,"   ' es vebta
                    Sql = Sql & db.texto(RS!Direccion) & "," ' direccion cliente
                    Sql = Sql & db.texto(RS!Poblacion) & "," ' poblacion
                    
                    
     
                    'Llevamos tanto el nombre del cliente como el de manipulador
                    'NomCarnetMani, NumCarnet, NifMani
                    If LotesCorrectos Then
                        'Datos carnet manipulador
                        Sql = Sql & db.texto(RS!ManipuladorNombre) & "," ' NombreSocio
                        Sql = Sql & db.texto(RS!ManipuladorNumCarnet) & ","
                        Sql = Sql & "NULL"
                    
                    Else
                        Sql = Sql & "NULL,NULL,NULL"
                    End If
                    
                    
                    'Noviembre
                    'email . Tendremos codclien para actualizar datos email tfno  pais
                    Sql = Sql & "," & RS!codClien
                    
                    'Noviembre 2021
                    'capacidad,tipoud,volumen
                    
                    Sql = Sql & "," & Int(Capacidad) & "," & DBLet(RS!unidadesCompra, "N")
                    'volumen
                    Sql = Sql & "," & DBSet(Volumen, "N") & ","
                    
                    'Codpobla
                    If DBLet(RS!codpobla, "T") = "" Then
                        Sql = Sql & "- 0 )"
                    Else
                        Sql = Sql & "-" & DBLet(RS!codpobla, "N") & ")"
                    End If
                End If
                
                RS.MoveNext
                If RS.EOF Then
                    NumRegElim = davidNumalbar + 1
                Else
                    NumRegElim = Len(Sql)
                End If
                
                
                If NumRegElim > davidNumalbar Then
                    Sql = Mid(Sql, 2) 'quitamos la primera coma
                    Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet, NifMani,email,capacidad,tipoud,volumen,ine) VALUES " & Sql
                    db.ejecutar Sql
                    Sql = ""
        
                End If
        Wend
        RS.Close
        
        
        If chkSoloMostrarErrores.Value = 1 And CADENA <> "" Then
            Errores = Errores & vbCrLf & vbCrLf & vbCrLf & "Ventas decimales " & vbCrLf & String(40, "=") & vbCrLf
            Errores = Errores & "      Factura       Cantidad" & vbCrLf
            Errores = Errores & CADENA
        End If
        CADENA = ""
        
        
        'Movimientros traspasos almacen
        'Se llevan de suminiostros a compras con un slhtra
        If chkROPO.Value = 0 And chkTratamientos.Value = 1 Then
            'Veremos si tienen almacen por defecto de advparametros <> ppal
             Sql = DevuelveDesdeBD(conAri, "codalmac", "advparametros ", "1", "1")
             If Sql = "" Then Sql = "1"
             If Val(Sql) <> 1 Then  'almacen separado
                    
                    CADENA = "select slhtra.*,nomartic,numserie, ctrlotes,unidadesCompra, descateg,unicajas2"
                    CADENA = CADENA & " FROM slhtra INNER JOIN sartic on slhtra.codartic=sartic.codartic "
                    CADENA = CADENA & " INNER JOIN  scateg ON sartic.codcateg = scateg.codcateg "
                    CADENA = CADENA & " AND codtrasp in (Select codtrasp from schtra where almadest=" & Sql
                    CADENA = CADENA & " AND fechatra between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                    Sql = CADENA & ")"
                    CADENA = ""
                    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Sql = ""
                    While Not RS.EOF
                       Sql = Sql & ",  ("
                       Sql = Sql & db.Fecha(RS!FechaMov) & "," ' FechaVenta
                       Sql = Sql & db.texto(RS!NomArtic) & "," ' NombreComercial
                       Sql = Sql & db.texto(RS!numSerie) & "," ' Registro
                       Sql = Sql & db.texto(RS!descateg) & "," ' Categoria
                       Sql = Sql & db.texto(RS!observa2) & "," ' Lote
                       
                       
                           
                       rsTipUd.Find "tipcompra =" & DBLet(RS!unidadesCompra, "N"), , adSearchForward, 1
                           
                       'Capacidad
                       Capacidad = DBLet(RS!unicajas2, "N")
                       If Capacidad = 0 Then Capacidad = 1
                       
                           
                       VtaPorUnidades = True
                       If Not rsTipUd.EOF Then VtaPorUnidades = rsTipUd!vtaindicacantidad = 0
                       If rsTipUd.EOF Then Debug.Assert False
                       
                       
                       If VtaPorUnidades Then
                           cantidad = RS!cantidad
                           Volumen = cantidad * Capacidad
                           
                           
                       Else
                           'Si pone 5, esta vendiendo 5 Litros. Si el envase es de 5Lts, esta vendiedndo 1 Unidad
                           
                           Volumen = RS!cantidad
                           
                           cantidad = Round2(Volumen \ Capacidad, 0)
                           If cantidad = 0 Then cantidad = 1
                       End If
                       Sql = Sql & db.numero(cantidad) & "," ' Cantidad

                       LotesCorrectos = False 'DBLet(RS!ManipuladorNombre, "T") <> ""
                       Sql = Sql & db.texto(vEmpresa.nomempre) & "," ' NombreSocio
                       Sql = Sql & db.texto(vParam.CifEmpresa) & "," ' NIF
                       Sql = Sql & db.texto("TRA" & Format(RS!codtrasp, "0000000")) & "," ' NumFactura
                       Sql = Sql & "3,"   ' es venta  --> Traspaso almacenes
                       Sql = Sql & db.texto(vParam.DomicilioEmpresa) & "," ' direccion cliente
                       Sql = Sql & db.texto(vParam.Poblacion) & "," ' poblacion
                       
                       
        
                       'Llevamos tanto el nombre del cliente como el de manipulador
                       'NomCarnetMani, NumCarnet, NifMani
                       Sql = Sql & "NULL,NULL,NULL"
                       
                       'Noviembre
                       'email . Tendremos codclien para actualizar datos email tfno  pais
                       Sql = Sql & ",0" '& RS!codClien
                       
                       'Noviembre 2021
                       'capacidad,tipoud,volumen
                       Sql = Sql & "," & Int(Capacidad) & "," & DBLet(RS!unidadesCompra, "N")
                       'volumen
                       Sql = Sql & "," & DBSet(Volumen, "N") & ","
                       
                       'Codpobla
                       If DBLet(vParam.CPostal, "T") = "" Then
                           Sql = Sql & "- 0 )"
                       Else
                           Sql = Sql & "-" & DBLet(vParam.CPostal, "N") & ")"
                       End If
                
                       RS.MoveNext
                    Wend
                    RS.Close
                    
                    If Sql <> "" Then
                        Sql = Mid(Sql, 2) 'quitamos la primera coma
                        Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet, NifMani,email,capacidad,tipoud,volumen,ine) VALUES " & Sql
                        db.ejecutar Sql
                        Sql = ""
        
                    End If
                    
            End If   'Tiene almacen tratamientos
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        'AHora vamos a actualizar datos de cliente desde declaralom
        lblInf.Caption = "Datos cliente"
        lblInf.Refresh
        
        If Me.chkROPO.Value = 0 Then
            
             Sql = "select distinct email , esventa from declaralom  where esventa>=1 and email <>''"
             RS.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
             
             Set miRsAux = New ADODB.Recordset
             While Not RS.EOF
                 lblInf.Caption = "Datos cliente"
                 lblInf.Refresh
                 
                 
                 If RS!EsVenta = 3 Then
                    'Traspaso de almacenes
                    'Los datos son de parametros
                    Sql = "select telempre telclie1,faxempre telclie2, maiempre maiclie1,null maiclie2 FROM sparam WHERE true"
                 Else
                    'Ventas
                    Sql = "Select telclie1,telclie2,maiclie1,maiclie2 FROM sclien where codclien = " & RS!email
                 
                 End If
                 miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                 
                 If Not miRsAux.EOF Then
                     Aux = DBLet(miRsAux!telclie1, "T")
                     If Aux = "" Then Aux = DBLet(miRsAux!telclie2, "T")
                     If Aux = "" Then Aux = "N/D"
                     Sql = "UPDATE declaralom SET telefono = " & DBSet(Aux, "T")
                     Aux = DBLet(miRsAux!maiclie1, "T")
                     If Aux = "" Then Aux = DBLet(miRsAux!maiclie2, "T")
                     If Aux = "" Then Aux = vParam.MailEmpresa
                     Sql = Sql & ", email =" & DBSet(Aux, "T")
                     Sql = Sql & " WHERE email =" & RS!email
                     conn.Execute Sql
                 Else
                     Err.Raise 513, , "Datos cliente. No encontrado. " & RS!email
                 End If
                 miRsAux.Close
            
                 RS.MoveNext
                 If Sql <> "" Then conn.Execute Sql
                 
             Wend
             RS.Close
             
                         
            'Autorizados
            lblInf.Caption = "Autorizados " & IIf(chkSoloMostrarErrores.Value = 1, "CHECK", "")
            lblInf.Refresh
            CADENA = ""
            Sql = "select numcarnet,nif,nombresocio from declaralom where esventa=1 and numcarnet <>'' group by 1,2"
            RS.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
                
            Set miRsAux = New ADODB.Recordset
            While Not RS.EOF
                lblInf.Caption = "Leyendo " & DBLet(RS!numcarnet, "T")
                lblInf.Refresh
                
                Sql = "select cif,nombre from sclienmani where numcarnet=" & DBSet(RS!numcarnet, "T")
                miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                If Not miRsAux.EOF Then
                    ' Aux = DBLet(miRsAux!Nombre, "T")
                    'Sql = DBLet(RS!nombresocio)
                     
                    Aux = DBLet(miRsAux!CIF, "T")
                    Sql = DBLet(RS!NIF)
                    If Aux <> Sql Then
                         
                         
                         
                       CADENA = CADENA & Mid(RS!nombresocio & Space(40), 1, 40) & " " & Mid(RS!numcarnet & Space(15), 1, 15) & " " & Sql & " > " & Aux & vbCrLf
                         
                           
                       Sql = DBSet(Aux, "T")
                       
                       
                       
                       'Sql = Sql & " WHERE email =" & RS!email
                       'conn.Execute Sql
                   Else
                       Sql = DBSet(RS!NIF, "T")
                       
                   End If
                   
                Else
                   Sql = DBSet(RS!NIF, "T")
                End If
                miRsAux.Close
                Sql = "UPDATE declaralom SET nifmani=" & Sql & " WHERE numcarnet=" & DBSet(RS!numcarnet, "T") & " AND nif=" & DBSet(RS!NIF, "T")
                conn.Execute Sql
                
                
                RS.MoveNext
                
                
            Wend
            RS.Close
            
            
            
            
            'Autorizados VARIOS
            lblInf.Caption = "Autorizados varios " & IIf(chkSoloMostrarErrores.Value = 1, "CHECK", "")
            lblInf.Refresh
            CADENA = ""
            Sql = "select numcarnet,nif,nombresocio from declaralom where esventa=1 and numcarnet <>'' group by 1,2"
            RS.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
                
            Set miRsAux = New ADODB.Recordset
            While Not RS.EOF
                lblInf.Caption = "Leyendo " & DBLet(RS!numcarnet, "T")
                lblInf.Refresh
                
                Sql = "select nifclien cif, nomclien nombre from sclvar where ManipuladorNumCarnet=" & DBSet(RS!numcarnet, "T")
                miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                If Not miRsAux.EOF Then
                     
                    Aux = DBLet(miRsAux!CIF, "T")
                    Sql = DBLet(RS!NIF)
                    If Aux <> Sql Then
                         
                       CADENA = CADENA & Mid(RS!nombresocio & Space(40), 1, 40) & " " & Mid(RS!numcarnet & Space(15), 1, 15) & " " & Sql & " > " & Aux & "    Var." & vbCrLf
                       Sql = DBSet(Aux, "T")
                       
                   Else
                       Sql = DBSet(RS!NIF, "T")
                   End If
                   
                Else
                   Sql = DBSet(RS!NIF, "T")
                End If
                miRsAux.Close
                Sql = "UPDATE declaralom SET nifmani=" & Sql & " WHERE numcarnet=" & DBSet(RS!numcarnet, "T") & " AND nif=" & DBSet(RS!NIF, "T")
                conn.Execute Sql
                
                
                RS.MoveNext
                
                
            Wend
            RS.Close
            Set miRsAux = Nothing
            
            
            If chkSoloMostrarErrores.Value = 1 And CADENA <> "" Then
               Errores = Errores & vbCrLf & vbCrLf & vbCrLf & "AUTORIZADOS " & vbCrLf & String(40, "=") & vbCrLf
               Errores = Errores & "      Socio                                Carnet          NIF vta      NIF mani" & vbCrLf
               Errores = Errores & CADENA
               CADENA = ""
            End If
            
             
      
        End If
        
        
        lblInf.Caption = "Proveedores " & IIf(chkSoloMostrarErrores.Value = 1, "CHECK", "")
        lblInf.Refresh
        DoEvents
        Sql = ""
        'SQL = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,"
        'SQL = SQL & " Poblacion,NomCarnetMani, NumCarnet, NifMani,telefono,email,capacidad,tipoud,volumen , ine)"
        Sql = Sql & " SELECT distinct a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, e.nomprove, e.nifprove, d.document, a.canentra"
        Sql = Sql & " ,0, domprove,trim(concat(codpobla,' ',pobprove)) lapoblacio ,e.nomprove ,"
        
        
        Sql = Sql & IIf(chkROPO.Value = 1, " '' ", " e.referencia")
        
        Sql = Sql & " referencia, e.nifprove , if(telprov1 is null,coalesce(telprov2,'N/D')  ,telprov1) telefono"
        Sql = Sql & ", if(maiprov1 is null,coalesce(maiprov2,'')  ,maiprov1 ) email"
        If Me.chkROPO.Value = 1 Then
            Sql = Sql & ",1 capacidad, 1 unidadescompra, 1 volumen "
        Else
            'reto
            'SQL = SQL & ",coalesce(unicajas2,1) capacidad, unidadescompra, coalesce(unicajas2,1) *  a.canentra  volumen "
            Sql = Sql & ",unicajas2 capacidad, unidadescompra, 1 volumen "
        End If
        Sql = Sql & " , -codpobla cpPro"
        
        
        Sql = Sql & " FROM slotes as a, sartic as b, scateg as c, smoval as d, sprove as e"
        Sql = Sql & " where b.codartic = a.codartic"
        Sql = Sql & " and c.codcateg = b.codcateg"
        Sql = Sql & " and d.codartic = a.codartic"
        Sql = Sql & " and d.fechamov = a.fecentra"
        Sql = Sql & " and d.tipomovi = 1"
        Sql = Sql & " and d.detamovi = 'ALC'"
        Sql = Sql & " and a.fecentra between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
        Sql = Sql & " and e.codprove = d.codigope"
            
            
        'Las compras, para cuando son SERVICIOS(Alzira) no van
        If Sql_Servicios <> "" Then
            'Es decir, para servicio digo que FALSE y me devuelve EOF
            If Me.chkTratamientos.Value = 1 Then Sql = Sql & " AND false "
        End If
        
        
         Set miRsAux = New ADODB.Recordset
         RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         NumRegElim = 0
         Sql = ""
         Capacidad = 1
         Volumen = 1
         CADENA = ""
         While Not RS.EOF
            lblInf.Caption = "Proveedores. " & RS!nomprove
            lblInf.Refresh
            NumRegElim = NumRegElim + 1
            
            'FechaVenta,NombreComercial,Registro
            Sql = Sql & ", (" & DBSet(RS!fecentra, "F") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!numSerie, "T")
            ',Categoria,Lote,NombreSocio,NIF
            Sql = Sql & ", " & DBSet(RS!descateg, "T") & "," & DBSet(RS!numlotes, "T") & "," & DBSet(RS!nomprove, "T") & "," & DBSet(RS!nifProve, "T")
            ',NumFactura,CanCompra,EsVenta,Direccion,"
            Sql = Sql & ", " & DBSet(RS!document, "T") & "," & DBSet(RS!canentra, "T") & ",0," & DBSet(RS!domprove, "T")
            'Poblacion,NomCarnetMani, NumCarnet,
            Sql = Sql & ", " & DBSet(RS!lapoblacio, "T") & "," & DBSet(RS!nomprove, "T") & "," & DBSet(RS!Referencia, "T")
            'NifMani,telefono,email,
            Sql = Sql & ", " & DBSet(RS!nifProve, "F") & "," & DBSet(RS!Telefono, "T") & "," & DBSet(RS!email, "T") & ","
            '          Cantidad,  capacidad,tipoud,volumen , ine)"
            
             If Me.chkROPO.Value = 1 Then
                cantidad = RS!canentra
            Else
            
                'If RS!numSerie = "25449" Then Debug.Assert False
                
                'RETO
                rsTipUd.Find "tipcompra =" & DBLet(RS!unidadesCompra, "N"), , adSearchForward, 1
                VtaPorUnidades = True
                If Not rsTipUd.EOF Then
                    VtaPorUnidades = rsTipUd!vtaindicacantidad = 0
                Else
                    If chkSoloMostrarErrores.Value = 1 Then CADENA = CADENA & RS!NomArtic & "     " & DBLet(RS!Capacidad) & "  Tipo UD" & vbCrLf
                End If
                
                'Capacidad
                Capacidad = DBLet(RS!Capacidad, "N")
                If Capacidad = 0 Then Capacidad = 1
                If Capacidad > 1 And VtaPorUnidades Then
                    If chkSoloMostrarErrores.Value = 1 Then CADENA = CADENA & RS!NomArtic & "     " & Capacidad & "  Vtas x UD" & vbCrLf
                End If
                If VtaPorUnidades Then
                    cantidad = RS!canentra
                    Volumen = cantidad * Capacidad
                Else
                    'Si pone 5, esta vendiendo 5 Litros. Si el envase es de 5Lts, esta vendiedndo 1 Unidad
                    
                    Volumen = RS!canentra
                    cantidad = Round2(Volumen \ Capacidad, 0)
                    If cantidad = 0 Then cantidad = 1
                End If
            End If
            
            
            
            ',Cantidad,capacidad,tipoud,volumen , ine
            Sql = Sql & DBSet(cantidad, "N") & "," & DBSet(Capacidad, "N") & ","
            Sql = Sql & DBLet(RS!unidadesCompra, "N") & "," & DBSet(Volumen, "N") & ","
            Sql = Sql & DBLet(RS!cpPro, "N") & ")"
            
            RS.MoveNext
            If RS.EOF Then NumRegElim = 101
            If NumRegElim > 30 Then
                    
                Aux = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,"
                Aux = Aux & " Poblacion,NomCarnetMani, NumCarnet, NifMani,telefono,email,Cantidad,capacidad,tipoud,volumen , ine) VALUES  " & Mid(Sql, 2)
                db.ejecutar Aux
                Sql = ""
                NumRegElim = 0
            End If
        Wend
        RS.Close
        
        If chkSoloMostrarErrores.Value = 1 And CADENA <> "" Then
           Errores = Errores & vbCrLf & vbCrLf & vbCrLf & "COMPRAS " & vbCrLf & String(40, "=") & vbCrLf
           Errores = Errores & "      " & vbCrLf
           Errores = Errores & CADENA
           CADENA = ""
        End If
        
        
        
        
        
        
        
        lblInf.Caption = "Nº ROPO proveedores"
        lblInf.Refresh
        Espera 0.25
        If Me.chkROPO.Value = 0 Then
            Sql = "Select distinct NIF,NombreSocio from declaralom where esventa=0 and coalesce(NumCarnet,'')=''"
            RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            Sql = ""
            While Not RS.EOF
                Sql = Sql & RS!NIF & "  " & RS!nombresocio & vbCrLf
                RS.MoveNext
            Wend
            RS.Close
            If Sql <> "" Then
                Sql = vbCrLf & vbCrLf & "Errores NºRopo proveedor" & vbCrLf & String(40, "=") & vbCrLf & Sql & vbCrLf
                Errores = Errores & Sql
                
                Sql = "UPDATE declaralom  SET numcarnet = NIF where esventa=0 and coalesce(NumCarnet,'')=''"
                conn.Execute Sql
            End If
        End If
        
        
        
        'If vParamAplic.Ariagro <> "" Then
        If BuscarEnSlifacCampos Then
        
            
        
            lblInf.Caption = "Enlace ariagro"
            lblInf.Refresh
            DoEvents
            Set Col = New Collection
            
            'Junio 2014
            cadFecha = " FechaVenta between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
            L = 0
            Sql = "Select count(*) from declaralom where esventa=1 AND " & cadFecha
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then L = DBLet(RS.Fields(0), "N")
            RS.Close
            lblInf.Tag = L
            
            Sql = "select FechaVenta,substring(numfactura,1,3),substring(numfactura,4) from declaralom where esventa=1 AND " & cadFecha
            
            If Me.chkROPO.Value = 0 Then Sql = Sql & " AND false"
            
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            L = 0
            Sql = ""
            While Not RS.EOF
                L = L + 1
                lblInf.Caption = "Col   " & Col.Count + 1 & "   Reg  " & L & " de " & lblInf.Tag
                lblInf.Refresh
                
                
                'Graba FECHA ALBARAN, y luiego no encuentra por fecha factura.
                'Buscamos , de momento, por serie+factura
                'SQL = SQL & ", (" & DBSet(Rs!fechaventa, "F") & "," & DBSet(Rs.Fields(1), "T") & "," & Rs.Fields(2) & ")"
                Sql = Sql & ", (" & DBSet(RS.Fields(1), "T") & "," & RS.Fields(2) & ")"
                RS.MoveNext
                
                
                If L > 29 Then
                    Col.Add Sql
                    Sql = ""
                    DoEvents
                    L = 0
                End If
            Wend
            RS.Close
            
            If L > 0 Then Col.Add Sql
            
            
            'Abro los tratamientos
            lblInf.Caption = "Leyendo tratamientos BD..."
            lblInf.Refresh
            
            Sql = "select codtrata,nomtrata from advtrata"
            rs2.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
            
            
            
            'Para cada subgrupo buscarenmos en slifaccampos
            For L = 1 To Col.Count
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count & " Cultivo"
                lblInf.Refresh
                DoEvents
                Sql = "(" & Mid(Col.Item(L), 2) & ")"
                'SQL = "Select * from slifaccampos where (fecfactu,codtipom,numfactu) IN " & SQL
                Sql = "Select * from slifaccampos where (codtipom,numfactu) IN " & Sql
                Sql = Sql & " AND fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                
                
                
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    
                    'FAV0079016
                    Sql = " AND numfactura = '" & RS!codtipom & Format(RS!Numfactu, "0000000") & "'"
                    'SQL = " WHERE esventa=1 and fechaventa= " & DBSet(Rs!FecFactu, "F") & SQL
                    Sql = " WHERE esventa=1 " & Sql
                    Sql = "UPDATE declaraLOM SET cultivo=" & RS!codCampo & Sql
                    conn.Execute Sql
                    RS.MoveNext
                Wend
                RS.Close
                
                
                'Vamos a ver los tratamientos
                lblInf.Caption = "Ariagro " & L & " de " & Col.Count & " Tratamiento"
                lblInf.Refresh

                
                Sql = "Select codtipom,numfactu,GROUP_CONCAT( substring(referenc,7) separator ' , ') from scafac1 where "
                Sql = Sql & " fecfactu between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                Sql = Sql & " AND referenc like 'PARTE%' "
                Sql = Sql & " AND (codtipom,numfactu) IN (" & Mid(Col.Item(L), 2) & ") group by 1,2"
                
                
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    Sql = DBLet(RS.Fields(2), "T")
                    lblInf.Caption = "parte : " & Sql
                    lblInf.Refresh
                    
                    If Sql <> "" Then
                    
                        NF = InStrRev(Sql, " , ")
                        If NF > 0 Then Sql = Mid(Sql, NF + 3)
                        
                        Sql = DevuelveDesdeBD(conAri, "codtrata", "advpartes", "numparte", Sql)
                        If Sql <> "" Then
                            rs2.Find "codtrata = " & Sql, , adSearchForward, 1
                            If Not rs2.EOF Then
                                'FAV0079016
                                Sql = " AND numfactura = '" & RS!codtipom & Format(RS!Numfactu, "0000000") & "'"
                                'SQL = " WHERE esventa=1 and fechaventa= " & DBSet(Rs!FecFactu, "F") & SQL
                                Sql = " WHERE esventa=1 " & Sql
                                Sql = "UPDATE declaraLOM SET tratamiento=" & DBSet(rs2!nomtrata, "T") & Sql
                                conn.Execute Sql
                            End If
                        End If
                    End If
                    RS.MoveNext
                Wend
                RS.Close
               
                  
                
            Next
                
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            lblInf.Caption = "Obtener variedad"
            lblInf.Refresh
            DoEvents
            Sql = "Select cultivo from declaralom where cultivo <>'' GROUP BY 1"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblInf.Caption = "Campo " & RS!cultivo
                lblInf.Refresh
                Sql = "select rcampos.codcampo,  variedades.nomvarie"
                Sql = Sql & " from @#rcampos inner join @#variedades on rcampos.codvarie = variedades.codvarie"
                Sql = Replace(Sql, "@#", vParamAplic.Ariagro & ".")
                Sql = Sql & " WHERE codcampo =" & RS!cultivo
                
                rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    Sql = "N/D"
                Else
                    Sql = rs2!nomvarie
                End If
                rs2.Close
                Sql = "UPDATE declaralom set cultivo=" & DBSet(Sql, "T") & " WHERE cultivo =" & DBSet(RS!cultivo, "T")
                conn.Execute Sql
                
                RS.MoveNext
            Wend
            RS.Close
            
                        
                        
                        
                        
                        
                        
                        
        
        End If 'de ariagro
        
        'Abril 2015
        'If vParamAplic.NumeroInstalacion = vbAlzira Then
        If vParamAplic.LlevaADV Then
            'Para aquellas facturas de servicio (que son tratamientos), si no esta indicado el cultivo, ni  la variedad
            'entonces UPDATEAMOS con los datos de la observacion
            Set rs2 = Nothing
            Set rs2 = New ADODB.Recordset
            Sql = "select fechaventa, NombreComercial,Registro,Categoria,Lote,NIF,NumFactura"
            Sql = Sql & " from declaralom where esventa=1 and numfactura like 'FAS%' and cultivo is null and tratamiento is null"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                lblInf.Caption = "Fra: " & RS!NumFactura
                lblInf.Refresh
                Sql = "select * from scafac1 where codtipom='FAS' "
                'SQL = SQL & " and fecfactu=" & DBSet(Rs!fechaventa, "F") & " and numfactu=" & Mid(Rs!NumFactura, 4)
                Sql = Sql & " and numfactu=" & Mid(RS!NumFactura, 4)
              
                rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If rs2.EOF Then
                    Sql = ""
                Else
                    Sql = Trim(DBLet(rs2!observa1, "T"))
                    If Me.chkROPO.Value = 0 Then Sql = ""
                End If
                rs2.Close
                
                If Sql <> "" Then
                    NumRegElim = 1
                    
                    'Vamos a quitar todos los espacios en blanco "duplicados"
                    Do
                        NumRegElim = InStr(NumRegElim, Sql, " ")
                        If NumRegElim > 0 Then
                            Do
                                L = InStr(NumRegElim + 1, Sql, " ")
                                If L = NumRegElim + 1 Then
                                    Sql = Mid(Sql, 1, L - 1) & Mid(Sql, L + 1)
                                Else
                                    L = 0
                                End If
                            Loop Until L = 0
                            NumRegElim = NumRegElim + 1
                        End If
                    Loop Until NumRegElim = 0
                            
                    
                
                
                    L = Len(Sql)
                    If L > 45 Then
                        cadFecha = Mid(Sql, 46)
                        Sql = Mid(Sql, 1, 45)
                    Else
                        cadFecha = ""
                    End If
                    Sql = "UPDATE declaralom set cultivo=" & DBSet(Sql, "T")
                    Sql = Sql & ",tratamiento= " & DBSet(cadFecha, "T", "S")
                    Sql = Sql & " where fechaventa=" & DBSet(RS!fechaventa, "F") & " and numfactura='" & RS!NumFactura
                    Sql = Sql & "' and lote=" & DBSet(RS!Lote, "T") & " and nif=" & DBSet(RS!NIF, "T")
                    Sql = Sql & " and registro=" & DBSet(RS!Registro, "T") & " and cultivo is null and tratamiento is null"
                    
                    conn.Execute Sql
                
                End If
                
                RS.MoveNext
            Wend
            RS.Close
            


        
        
        End If
        
        
        
        
        'OCTUBRE 2016
        ' Lotes fitosnatiarios SUBVENCIONDOS
        If Me.ChkSubvencionados.Value = 1 Then
            DoEvents
            
            
                Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio,"
                Sql = Sql & " NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet)"
                Sql = Sql & " SELECT slotesgeneralitatmov.fechamov,nomartic,slotesgeneralitat.numserie,'LO',numlote,"
                Sql = Sql & " slotesgeneralitatmov.cantidad,nomclien,nifclien,concat(""ID"" ,"
                Sql = Sql & " right(concat(""000000"",id),6) ,right(concat(""0000"",idMov),4)),1,domclien,"
                Sql = Sql & " concat(codpobla,' ',pobclien) poblacion,slotesgeneralitatmov.ManipuladorNombre , "
                Sql = Sql & " slotesgeneralitatmov.ManipuladorNumCarnet"
                Sql = Sql & " from slotesgeneralitat,slotesgeneralitatmov,sclien,sartic where slotesgeneralitat.Id = "
                Sql = Sql & " slotesgeneralitatmov.idlote and slotesgeneralitatmov.codclien=sclien.codclien"
                Sql = Sql & " and slotesgeneralitat.codartic=sartic.codartic"
                Sql = Sql & " AND slotesgeneralitatmov.fechamov between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                Sql = Sql & " ORDER BY slotesgeneralitatmov.fechamov,id"
                lblInf.Caption = "Lotes subvencionados"
                lblInf.Refresh
                
                
                
                db.ejecutar Sql
        
        
            
                DoEvents
                Me.Refresh
                
                lblInf.Caption = "Proveedores lotes subv."
                lblInf.Refresh
                Sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, cancompra, NombreSocio,"
                Sql = Sql & " NIF, NumFactura,EsVenta,Direccion,Poblacion,NomCarnetMani, NumCarnet)"
                Sql = Sql & " SELECT slotesgeneralitat.fecha,nomartic,slotesgeneralitat.numserie,'LO',numlote,"
                Sql = Sql & " slotesgeneralitat.cantidad,nomprove,nifprove,concat(""COD "" ,right(concat(""000000"",id),6)),0,domprove,"
                Sql = Sql & " concat(codpobla,' ',pobprove) poblacion,null , null"
                Sql = Sql & " From slotesgeneralitat, sartic, sprove Where slotesgeneralitat.Codprove = sprove.Codprove"
                Sql = Sql & " and slotesgeneralitat.codartic=sartic.codartic"
                Sql = Sql & " AND slotesgeneralitat.fecha  between " & DBSet(txtFecha(0).Text, "F") & " AND " & DBSet(txtFecha(1).Text, "F")
                Sql = Sql & " ORDER BY slotesgeneralitat.fecha,id"
                db.ejecutar Sql

         End If
        
            
        'Obtenecion INE desde codpobla
        If Me.chkROPO.Value = 0 Then
            lblInf.Caption = "Obtener codigo INE"
            lblInf.Refresh
            Sql = "    select distinct ine from  declaralom WHERE ine<0"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set Col = New Collection
            davidNumalbar = 10
            NumRegElim = 0
            Sql = ""
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                Sql = Sql & ", " & Abs(RS!ine)
                RS.MoveNext
                If RS.EOF Then NumRegElim = davidNumalbar + 1
                
                If NumRegElim > davidNumalbar Then
                    Col.Add "(" & Mid(Sql, 2) & ")"
                    NumRegElim = 0
                End If
            Wend
            RS.Close
        
        
            
            For NumRegElim = 1 To Col.Count
                lblInf.Caption = "Ajuste INE / poblacion. (II) " & NumRegElim & " de " & Col.Count
                lblInf.Refresh
            
                Sql = "select * from  scpostal where cpostal  in " & Col.Item(NumRegElim) & " AND ine >0"
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    Sql = "UPDATE declaralom set ine=" & RS!ine & " WHERE ine = -" & RS!CPostal
                    conn.Execute Sql
                    RS.MoveNext
                Wend
                RS.Close
            Next
                
            '##Si dio mensaes de error de INE algun cpostal estara en negativo.
            ' Lo paso a positivo
            Sql = "UPDATE declaralom set ine=-ine  WHERE ine < 0"
            conn.Execute Sql
            
        End If
            
            
            
            
        lblInf.Caption = "Comprobar carnets facturas venta"
        lblInf.Refresh

        Sql = "select NumFactura ,FechaVenta ,NombreSocio from declaralom where esventa>0 and (numcarnet is null  or  NomCarnetMani is null)"
        If Me.chkTratamientos.Value = 1 Then
            'No hace falta obtener carnet de manipulador en los tratamientos, ya que es la COOPERATIVA quien lo realiza
            Sql = Sql & " AND FALSE"
        End If
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set Col = New Collection
        davidNumalbar = 10
        NumRegElim = 0
        Sql = ""
        While Not RS.EOF
            Sql = Sql & "- " & DBLet(RS!NumFactura, "T") & " " & DBLet(RS!fechaventa, "T") & "   " & DBLet(RS!nombresocio, "T") & vbCrLf
            RS.MoveNext
        Wend
        RS.Close
        If Sql <> "" Then
            If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
            Aux = "Ventas sin identificar manpulador" & vbCrLf & String(40, "=") & vbCrLf & Sql
            Errores = Errores & Aux
        End If
        
        
        
        
        
        'Ajustes finales
        '   SIN rectificativas
        '   Ajustes fecha funcion presentacion
        
        'marzo 2022. Facturas venta rectificativas NO se declaran
        lblInf.Caption = "Rectificativas"
        lblInf.Refresh
        Set miRsAux = New ADODB.Recordset
        CADENA = ""
        
        If chkSoloMostrarErrores.Value = 1 Then
            Sql = "select * from declaralom  where esventa=1 and cantidad<=0  ORDER BY numfactura"
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                CADENA = CADENA & Mid(miRsAux!nombresocio & Space(40), 1, 40) & " " & Mid(miRsAux!NumFactura & Space(15), 1, 15) & " " & miRsAux!cantidad
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        conn.Execute "DELETE from declaralom  where esventa=1 and cantidad<=0"
        
        If CADENA <> "" Then
           Errores = Errores & vbCrLf & vbCrLf & vbCrLf & "VTAS RECTIFICTIVAS " & vbCrLf & String(40, "=") & vbCrLf
           Errores = Errores & "      " & vbCrLf
           Errores = Errores & CADENA
           CADENA = ""
        End If
        
        
        
        If Me.chkTratamientos.Value = 0 And chkFechasPlazo.Value = 1 Then
            lblInf.Caption = "Plazos fechas VENTA"
            lblInf.Refresh
            F = DateAdd("m", -1, Now()) ' -un mes
            
            Sql = "select distinct FechaVenta from declaralom  where esventa= 1 and FechaVenta <" & DBSet(F, "F")
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            NumRegElim = 1
            F2 = F
            While Not miRsAux.EOF
                Sql = ""
                Do
                    NumRegElim = NumRegElim + 1
                    F2 = DateAdd("d", NumRegElim, F)
                    If F2 >= Now Then
                        F2 = F
                        NumRegElim = 1
                    End If
                    If Weekday(F2, vbMonday) <> 7 Then Sql = "S"
                Loop Until Sql <> ""
                If chkSoloMostrarErrores.Value = 1 Then CADENA = CADENA & miRsAux!fechaventa & "      " & Format(F2, "dd/mm/yyyy") & vbCrLf
                
                Sql = "UPDATE declaralom set fechaventa= " & DBSet(F2, "F") & " WHERE esventa=1 "
                Sql = Sql & "  AND fechaventa=" & DBSet(miRsAux!fechaventa, "F")
                conn.Execute Sql
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If chkSoloMostrarErrores.Value = 1 Then
               Errores = Errores & vbCrLf & vbCrLf & vbCrLf & "Ajuste fechas " & vbCrLf & String(40, "=") & vbCrLf
               Errores = Errores & "  Fecha       Ajuste" & vbCrLf
               Errores = Errores & CADENA
            End If
            NumRegElim = 0
        End If
        
        
        
        lblInf.Caption = "Comprobar carnets ROPO"
        lblInf.Refresh
        
        Sql = "select numcarnet ,NomCarnetMani,GROUP_CONCAT( numfactuRA separator ' ')  from declaralom where numcarnet <>'' "
        If Me.chkTratamientos.Value = 1 Then
            'No hace falta obtener carnet de manipulador enlos tratamientos, ya que es la COOPERATIVA quein lo realiza
            Sql = Sql & " AND FALSE"
        End If
        'ROPOR TAMPCOCO
        If Me.chkROPO.Value = 1 Then Sql = Sql & " AND FALSE"
        Sql = Sql & " GROUP BY numcarnet having length(numcarnet)<11"
        
        
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set Col = New Collection
        davidNumalbar = 10
        NumRegElim = 0
        Sql = ""
        While Not RS.EOF
            Sql = Sql & "- " & DBLet(RS!numcarnet, "T") & " " & DBLet(RS!NomCarnetMani, "T") & " -->" & RS.Fields(2) & vbCrLf
            RS.MoveNext
        Wend
        RS.Close
        If Sql <> "" Then
            If Errores <> "" Then Errores = Errores & vbCrLf & vbCrLf & vbCrLf
            Aux = "Carnet ROPO incorrecto:" & vbCrLf & String(40, "=") & vbCrLf & Sql
            Errores = Errores & Aux
        End If
        
        
       NumRegElim = 0
        ParaVerDatosINE = ""
        If Me.chkSoloMostrarErrores.Value = 1 Then
            Sql = "select trim(substring(poblacion,1,6)) codp ,ine,trim(substring(poblacion,7)) nombre from declaralom group by 1 ORDER by 1"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
            While Not RS.EOF
                Sql = Sql & "   " & DBLet(RS!codp, "T") & "  " & Mid(DBLet(RS!Nombre, "T") & Space(35), 1, 35) & "     " & DBLet(RS!ine, "T") & vbCrLf
                RS.MoveNext
            Wend
            RS.Close
            If Sql <> "" Then
                Aux = vbCrLf & vbCrLf & Sql
                ParaVerDatosINE = "Codigos INE a presentar: " & vbCrLf & String(40, "=") & vbCrLf & Aux
            End If
            NumRegElim = 1
        Else
            If Errores <> "" Then NumRegElim = 1
        End If
        
        
   
        If NumRegElim = 1 Then
                NF = FreeFile
                Sql = App.Path & "\ErrROPO.txt"
                Open Sql For Output As #NF
                Print #NF, Errores
                If ParaVerDatosINE <> "" Then Print #NF, ParaVerDatosINE
                Close #NF
                
                LanzaVisorMimeDocumento Me.hwnd, Sql
                Espera 0.5
                
                If Me.chkSoloMostrarErrores.Value Then Exit Sub
                If Errores <> "" Then
                    If MsgBox("Existen errores . ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                End If
        End If
        
        rsTipUd.Close
        Set rsTipUd = Nothing
        
        
        
        
        
        DoEvents
        
        
        'Antes de llamar al informe GENERO el fichero
        lblInf.Caption = "Generar fichero"
        lblInf.Refresh
        If CrearFicheroReto Then
        
        
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
            If Not PonerParamRPT2(IIf(chkROPO.Value = 1, 31, 98), "", 0, Aux, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
                Exit Sub
            End If
            
            
            
            If Me.chkTratamientos.Value = 1 Then Aux = Replace(Aux, ".rpt", "S.rpt")
            frmVisReport.SoloImprimir = False
            frmVisReport.Informe = App.Path & "\Informes\" & Aux
            frmVisReport.CambiaODBC = False
            frmVisReport.FormulaSeleccion = "{declaralom.FechaVenta} in " & _
                                                "Date(" & Format(Desde, "yyyy") & _
                                                "," & Format(Desde, "mm") & _
                                                "," & Format(Desde, "dd") & ")" & _
                                                " to" & _
                                                " Date(" & Format(Hasta, "yyyy") & _
                                                "," & Format(Hasta, "mm") & _
                                                "," & Format(Hasta, "dd") & ")"
            frmVisReport.Show vbModal
            
        
            'Vemos de copiarlo donde digan
            'Lanzo el
            lblInf.Caption = "Guardando"
            lblInf.Refresh
            Sql = GuardarFicheroPedirNombre
            
            If Sql <> "" Then
                CopiarYGrabarEnBBDD Errores, Sql
            Else
                MsgBox "Proceso cancelado", vbExclamation
            End If
        
        End If
        '--
        lblInf.Caption = "Proceso terminado."
        lblInf.Refresh
        DoEvents
  
  
eProcesoDesdeSlifac:
    If Err.Number <> 0 Then MuestraError Err.Number, Sql, Err.Description
        
    Set RS = Nothing
    Set rs2 = Nothing
    Set miRsAux = Nothing
    Set rsTipUd = Nothing
    davidCodtipom = ""
    davidNumalbar = 0
    
        
End Sub






Private Function CrearFicheroReto() As Boolean
Dim Clipro As String
Dim CPos As String
Dim PintaCarnet As Boolean
Dim EsVenta As Boolean
Dim Cuantos As Currency

On Error GoTo eCrearFicheroReto

    CrearFicheroReto = False
    
    Set miRsAux = New ADODB.Recordset
    
    'VAMOS a obtener los datos de presentador ROPO   Hay uno para venta y otro para suministro
    If Me.chkROPO.Value = 0 Then
        Sql = "SELECT RetoVta ,RetoSumi FROM spara1"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO puede ser EOF
        CadenaDesdeOtroForm = DBLet(miRsAux!retovta, "T")
        If CadenaDesdeOtroForm = "" Then Err.Raise 513, , "Datos codigo ROPO vacio"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & DBLet(miRsAux!RetoSumi, "T") & "|" 'aplicador fitos
        miRsAux.Close
    Else
        CadenaDesdeOtroForm = "||"
    End If
    
    NF = -1
    Sql = App.Path & "\reto.txt"    'Si se cambia aqui , mirar en CopiarYGrabarEnBBDD
    If Dir(Sql, vbArchive) <> "" Then Kill Sql

    'Abrimos fichero
    NF = FreeFile
    Open Sql For Output As #NF
   
        
    If Me.chkROPO.Value = 1 Then
        Sql = "FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra,EsVenta,Direccion,Poblacion,Cultivo,Tratamiento,NomCarnetMani,NumCarnet,NifMani"
        Print #NF, Replace(Sql, ",", ";") & ";"
        Sql = "select " & Sql & " from declaralom order by fechaventa,esventa "
    
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
            'ROPO
            Sql = ""
            For NumRegElim = 0 To miRsAux.Fields.Count - 1
                Sql = Sql & """" & DBLet(miRsAux.Fields(NumRegElim).Value, "T") & """;"
            Next
            Print #NF, Sql
            miRsAux.MoveNext
        Wend
            
    Else
        
        Clipro = ""
        Print #NF, "RT0001"
        Sql = "select declaralom.*,if(cantidad<0,1,0) esnegativo from declaralom order by esventa asc,nif,numcarnet,fechaventa,numfactura,esnegativo"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
                
            If miRsAux!esnegativo = 1 And miRsAux!EsVenta = 0 Then Debug.Assert False
        
        
            'RETO
            'Para cada "cliente /proveedor " y persona autirzada imprimira resgistro transaccion (Tipo1)
            Sql = miRsAux!EsVenta & miRsAux!NIF & "@" & miRsAux!numcarnet & "@" & Format(miRsAux!fechaventa, "ddmmyy") & "@" & miRsAux!esnegativo
            
            
            If Sql <> Clipro Then
            
            
            
                Clipro = CStr(Sql)
                
                'Imprimimos el registro
                Sql = "1;" & Format(miRsAux!fechaventa, "dd/mm/yyyy") & ";" & vParam.CifEmpresa & ";"
                
                
                    
                If Me.chkTratamientos.Value = 1 Then
                    'Aplicacon tratamientos
                    Sql = Sql & RecuperaValor(CadenaDesdeOtroForm, 2) & ";"
                    
                    If miRsAux!EsVenta = 3 Then
                        'Traspaso entre almacenes de fisoantiarios
                        Sql = Sql & "7;"
                    Else
                        'Aplicador tratamiento
                        Sql = Sql & "8;"
                    End If
                Else
                    Sql = Sql & RecuperaValor(CadenaDesdeOtroForm, 1) & ";"
                    
                    EsVenta = miRsAux!EsVenta = 1
                    'Si es venta negativo se trata como una compra
                    'Si es compra neativo se trata como una venta
                    If miRsAux!esnegativo = 1 Then EsVenta = Not EsVenta
                    Sql = Sql & IIf(EsVenta, 2, 1) & ";"  'vta:2   compra 1
                End If
            
                
                
                If Me.chkTratamientos.Value = 1 Then
                    'Es trtamiento
                    Sql = Sql & miRsAux!NIF & ";;"   '
                    'SQL = SQL & vParam.CifEmpresa & ";" & RecuperaValor(CadenaDesdeOtroForm, 2) & ";"
                Else
                    'El destino cuando es una COMPRA se supone que es el proveedor
                    If miRsAux!EsVenta = 1 Then
                        
                        If DBLet(miRsAux!nifmani, "T") <> "" Then
                            Sql = Sql & miRsAux!nifmani
                        Else
                            Sql = Sql & miRsAux!NIF
                        End If
                        Sql = Sql & ";" & miRsAux!numcarnet & ";"
                    Else
                        'lo que habia
                        Sql = Sql & miRsAux!NIF & ";" & miRsAux!numcarnet & ";"
                    End If
                End If
                Sql = Sql & miRsAux!nombresocio & ";" & DBLet(miRsAux!email, "T") & ";"
                Sql = Sql & DBLet(miRsAux!Telefono, "T") & ";;" 'fax
                
                CPos = Trim(Mid(miRsAux!Poblacion, 1, 5))
                
                If CPos = "" Then CPos = "46000"
                If Len(CPos) < 5 Then CPos = Left(CPos & "00000", 5)
                If Not IsNumeric(CPos) Then CPos = "46000"
                    
                Sql = Sql & Trim(DBLet(miRsAux!Direccion, "T")) & ";" & CPos & ";ES;"
                
                'Si ine tiene valor se queda ine, si no, cpostal
                If DBLet(miRsAux!ine, "T") <> "" Then CPos = Mid(miRsAux!ine & "00000", 1, 5)
                    
                Sql = Sql & Mid(CPos, 1, 2) & ";" & Mid(CPos, 3, 3) & ";"
                 
                'nif nombre apee1 apell2
                PintaCarnet = False
                If Me.chkTratamientos.Value = 0 Then If miRsAux!EsVenta = 1 Then PintaCarnet = True
                
                If PintaCarnet Then
                    'Venta
                    SeparaNombreManipulador DBLet(miRsAux!NomCarnetMani, "T")
                    CPos = miRsAux!nombresocio
                Else
                    Sql = Sql & ";;;;"
                    CPos = ""
                End If
                
                'Empresa explotadora
                Sql = Sql & CPos & ";"
                Print #NF, Sql
            End If
            
        
            
            'Lineas articulos
            '  Registro NombreComercial Lote capacidad tipoud Cantidad volumen N" "N"
            Sql = "2;" & miRsAux!Registro & ";" & miRsAux!NombreComercial & ";" & miRsAux!Lote & ";"
            
            'En los tratamientos NO pintaremos la capacidad
            If Me.chkTratamientos.Value = 0 Then Sql = Sql & miRsAux!Capacidad
            
            
            Sql = Sql & ";" & miRsAux!tipoud & ";"
            
            
             If Me.chkTratamientos.Value = 0 Then
                'SIEMPRE pondra CANTIDAD
                'Sql = Sql & IIf(miRsAux!esventa = 0, miRsAux!CanCompra, miRsAux!cantidad)
                Cuantos = Abs(IIf(miRsAux!EsVenta = 0, miRsAux!cantidad, miRsAux!cantidad))
                Sql = Sql & Cuantos
            End If
            'Sql = Sql & ";" & miRsAux!Volumen & ";N;N;"
            Cuantos = Abs(miRsAux!Volumen)
            Sql = Sql & ";" & Cuantos & ";N;N;"
            If Me.chkTratamientos.Value = 1 Then
                If miRsAux!EsVenta <> 3 Then Sql = Sql & "plagas;"
            End If
            Print #NF, Sql
        
        
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    



    Sql = DevuelveDesdeBD(conAri, "count(*)", "declaralom", "1", "1")
    If Val(Sql) > 0 Then
        CrearFicheroReto = True
    Else
        MsgBox "Ningun dato generado", vbExclamation
    End If

  



    
    
eCrearFicheroReto:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    If NF > -1 Then Close #NF
    Set miRsAux = Nothing
End Function

'A partir de nombre extraera (lo mejor posible) nombre apeel1 apell2 y lo mete en SQL separado por ;
Private Sub SeparaNombreManipulador(Nombre As String)
Dim i As Integer
Dim J As Integer
Dim Apell As String

    
    'Si tiene "," 4l nombre sera Apell1  Apell2  , nombre
    i = InStrRev(Nombre, ",")
    If i > 0 Then
        'Perfecto esta la coma
        Apell = Mid(Nombre, 1, i - 1)
        Nombre = Trim(Mid(Nombre, i + 1))
    Else
        'NO hay coma
        'Veo el preimer espacio en blanco
        i = InStr(1, Nombre, " ")
        If i = 0 Then
            If Nombre = "" Then Nombre = "N/D"
            'Esto debe ser un fallo grnade. No hay espacios en blanco. Pongo nombre y sin apellidos y salgo
            Sql = Sql & Nombre & ";;;"
            Exit Sub
        End If
            
        If i > 3 Then
            J = i
        Else
            J = InStr(i + 1, Nombre, " ")
            If J = 0 Then J = i 'por el motivo que sea Solo hay un espacio en blanco
        End If
        
        Apell = Mid(Nombre, J + 1)
        Nombre = Mid(Nombre, 1, J - 1)
    End If
    
    'Grabo en SQL el nombre
    Sql = Sql & Nombre & ";"
    
    'En apell tengo que divir apell 2
    i = InStr(1, Apell, " ")
    If i = 0 Then
        'No hay ningun espacio en blanco. Solo un apellido
        Nombre = ""
    Else
        'Vale hay un espacio en blanco
        If i > 3 Then
            J = i
        Else
            J = InStr(i + 1, Apell, " ")
            If J = 0 Then J = i 'por el motivo que sea Solo hay un espacio en blanco
        End If
    
        Nombre = Mid(Apell, J + 1)
        Apell = Mid(Apell, 1, J - 1)
    End If
    
    'Grabo en SQL el apell1 y 2
    Sql = Sql & Apell & ";" & Nombre & ";"
    
        
End Sub



Private Function GuardarFicheroPedirNombre() As String
On Error GoTo eDialog
    GuardarFicheroPedirNombre = ""
    cd1.CancelError = True
    cd1.DefaultExt = ".csv" 'extension por defecto
    cd1.Filter = "CSV |*.csv|" 'extensiones a mostrar
    cd1.FilterIndex = 1
    cd1.FileName = ""
    
    Me.cd1.ShowSave
    
eDialog:
    If Err.Number <> 0 Then Err.Clear
    If cd1.FileName <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("Fichero " & cd1.FileName & " ya existe. ¿Sobreescribir?", vbQuestion + vbYesNo) = vbYes Then
                Kill cd1.FileName
                If Err.Number <> 0 Then
                    MuestraError Err.Number
                    Exit Function
                End If
            Else
                Exit Function
            End If
            
        End If
        GuardarFicheroPedirNombre = cd1.FileName
    End If
    
End Function

Private Sub CopiarYGrabarEnBBDD(vErrores As String, NombreFich As String)

    On Error Resume Next

    cantidad = 0
    resto = 0
    NF = 0




    FileCopy App.Path & "\reto.txt", NombreFich
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, , "Copiando en destino: " & NombreFich
        Exit Sub
    End If
    
    'Perfecto. Guardamos en BD
    If Me.chkROPO.Value = 0 Then
        Set miRsAux = New ADODB.Recordset
        Sql = "select sum(if(esventa=0,1,0)) compras,sum(if(esventa=1,1,0)) ventas,sum(if(esventa=2,1,0)) tramto , 0 traspasos from declaralom"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "Se ha producido un error leyendo totales", vbExclamation
        Else
            ' declaralom_reto fechahora usuario pc lineasvtas lineascompras lineastraspasos  lineastramto fechainicio fechafin errores
            Sql = "(" & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vUsu.PC, "T") & ","
            Sql = Sql & DBLet(miRsAux!Ventas, "N") & "," & DBLet(miRsAux!Compras, "N") & ","
            Sql = Sql & DBLet(miRsAux!traspasos, "N") & "," & DBLet(miRsAux!tramto, "N") & ","
            Sql = Sql & DBSet(txtFecha(0).Text, "F") & "," & DBSet(txtFecha(1).Text, "F") & ","
            Sql = Sql & DBSet(vErrores, "T", "S") & ")"
            Sql = "INSERT INTO declaralom_reto(fechahora ,usuario ,pc ,lineasvtas ,lineascompras ,lineastraspasos  ,lineastramto ,fechainicio ,fechafin ,errores) VALUES " & Sql
            ejecutar Sql, False
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    MsgBox "Proceso finalizado", vbInformation
End Sub


Private Function TratamientosDesdeAriagro() As Boolean
        TratamientosDesdeAriagro = False
        'Comprobaremos en parametro apliacdor fitosantiarios
        Sql = DevuelveDesdeBD(conAri, "RetoSumi", "spara1", "1", "1")
        
        ' Tiene indicado el numero de plicador y no tienen partes.
        'De momento, ese caso es QUATRETONDA que tiene sus tablas en ariagro
        Sql = IIf(Len(Sql) > 8, "1", "0")
        If Val(Sql) = 1 Then
            Sql = DevuelveDesdeBD(conAri, "count(*)", "ariagro.advfamia", "1", "1")
            If Val(Sql) >= 1 Then TratamientosDesdeAriagro = True: lblAriagro.visible = True
        End If
End Function









Private Function HacerListadoCarnets(Varios As Boolean) As Boolean
Dim F1 As Date
Dim CADENA As String


    On Error GoTo eCmdCarnetsCaducados
    
    
    
    
    HacerListadoCarnets = False
    
    
    Set RS = New ADODB.Recordset
    F1 = CDate(Sql)
    
    conn.Execute "Delete from tmpinformes where codusu =" & vUsu.Codigo
    
    CadenaDesdeOtroForm = "INSERT INTO tmpinformes (CodUsu , Codigo1, campo1, campo2, nombre1, nombre2, nombre3, fecha1, obser) VALUES "
    cantidad = 0
    'CodUsu , Codigo1, campo1, campo2, nombre1, nombre2, nombre3, fecha1, obser
    Sql = "select codclien,nomclien,maiclie1,maiclie2,nifclien,ManipuladortipoCarnet,ManipuladorNumCarnet,ManipuladorFecCaducidad from sclien where manipuladornumcarnet<>'' "
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        CADENA = ""
        
        Do
            MontaCadenaCarnetCaducado False, F1
            CADENA = CADENA & Sql
            cantidad = cantidad + 1
            
            RS.MoveNext
            
            If RS.EOF Then cantidad = 1000
            
            If cantidad > 900 Then
                CADENA = Mid(CADENA, 2)
                Sql = CadenaDesdeOtroForm & CADENA
                conn.Execute Sql
                cantidad = 0
                CADENA = ""
            End If
        Loop Until RS.EOF
    End If
    RS.Close
    
    Sql = "select sclienmani.codclien,nomclien,maiclie1,maiclie2,nombre,cif nifclien,tipocarnet ManipuladortipoCarnet,numcarnet ManipuladorNumCarnet,fcaducidad ManipuladorFecCaducidad "
    Sql = Sql & " from sclien inner join sclienmani on sclien.codclien=sclienmani.codclien where numcarnet<>'' "
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        CADENA = ""
        
        Do
            MontaCadenaCarnetCaducado True, F1
            CADENA = CADENA & Sql
            cantidad = cantidad + 1
            
            RS.MoveNext
            
            If RS.EOF Then cantidad = 1000
            
            If cantidad > 900 Then
                CADENA = Mid(CADENA, 2)
                Sql = CadenaDesdeOtroForm & CADENA
                conn.Execute Sql
                cantidad = 0
                CADENA = ""
            End If
        Loop Until RS.EOF
    End If
    RS.Close

    
    'CodUsu , Codigo1, campo1, campo2, nombre1, nombre2, nombre3, fecha1, obser
    Sql = DevuelveDesdeBD(conAri, "codclien", "spatpvg", "1", "1")
    If Sql = "" Then
        Sql = DevuelveDesdeBD(conAri, "codclien", "sclien", "clivario", "1 ORDER BY codclien DESC")
    End If
    Sql = "select " & Sql & " codclien,nomclien,'' maiclie1,'' maiclie2,nifclien,ManipuladortipoCarnet,"
    Sql = Sql & " ManipuladorNumCarnet,fcaducidad ManipuladorFecCaducidad from sclvar where manipuladornumcarnet<>''"
    If Not Varios Then Sql = Sql & " AND FALSE"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cantidad = 0
    If Not RS.EOF Then
        CADENA = ""
        
        Do
            MontaCadenaCarnetCaducado False, F1
            CADENA = CADENA & Sql
            cantidad = cantidad + 1
            
            RS.MoveNext
            
            If RS.EOF Then cantidad = 1000
            
            If cantidad > 900 Then
                CADENA = Mid(CADENA, 2)
                Sql = CadenaDesdeOtroForm & CADENA
                conn.Execute Sql
                cantidad = 0
                CADENA = ""
            End If
        Loop Until RS.EOF
    End If
    RS.Close
    

    
    
    
    
    
    
    
    Sql = F1 'vuelvo a dejar la fecha aqui para el rpt
    HacerListadoCarnets = True
eCmdCarnetsCaducados:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set RS = Nothing
End Function

Private Sub MontaCadenaCarnetCaducado(Autorizado As Boolean, ByRef FCad As Date)
    'CadenaDesdeOtroForm = "INSERT INTO tmpinformes (CodUsu , Codigo1, campo1, campo2, nombre1, nombre2, nombre3, fecha1, obser)"
    Sql = ""
    NF = 0 'NO caduca en breve
    If IsNull(RS!ManipuladorFecCaducidad) Then
        NF = 3
    Else
        If RS!ManipuladorFecCaducidad <= FCad Then
            NF = 2
        Else
            NumRegElim = DateDiff("d", FCad, RS!ManipuladorFecCaducidad)
            If NumRegElim <= 31 Then NF = 1
            
           ' If NF = 1 Then Debug.Assert False
        End If
    End If


    Sql = Sql & ", (" & vUsu.Codigo & "," & RS!codClien & "," & NF & "," & Abs(Autorizado) & ","
    If Autorizado Then
        Sql = Sql & DBSet(RS!Nombre, "T") & "," & DBSet(RS!nifClien, "T") & "," & DBSet(RS!ManipuladorNumCarnet, "T") & ","
    Else
        Sql = Sql & DBSet(RS!NomClien, "T") & "," & DBSet(RS!nifClien, "T") & "," & DBSet(RS!ManipuladorNumCarnet, "T") & ","
    End If
    Sql = Sql & DBSet(RS!ManipuladorFecCaducidad, "F")
    
    If Not IsNull(RS!maiclie1) Then
        davidCodtipom = DBSet(RS!maiclie1, "T")
    ElseIf Not IsNull(RS!maiclie2) Then
        davidCodtipom = DBSet(RS!maiclie2, "T")
    Else
        davidCodtipom = "NULL"
    End If
    Sql = Sql & "," & davidCodtipom & ")"
    
    
    
    
End Sub
