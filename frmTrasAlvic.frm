VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasAlvic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasAlvic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   4665
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkSeparadoTabulador 
         Caption         =   "Campos separado por tabuladores"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1545
         Left            =   240
         TabIndex        =   5
         Top             =   690
         Width           =   5955
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2730
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   495
            Width           =   1200
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2730
            MaxLength       =   10
            TabIndex        =   1
            Top             =   960
            Width           =   1170
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2430
            Picture         =   "frmTrasAlvic.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   1500
            TabIndex        =   7
            Top             =   540
            Width           =   1425
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ID Turno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   1500
            TabIndex        =   6
            Top             =   960
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   3
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   2
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   375
         Left            =   210
         TabIndex        =   8
         Top             =   2760
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasAlvic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTE (Alvic)
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1




'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim cad As String
Dim cadTabla As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Dim ArtFamGenerica As String

Dim IvaNormal As Currency
Dim IvaReducido As Currency
Dim IvaSuperReducido As Currency

Dim Col As Collection   'Segun el caso sera, Factura , albaran o tiocket


Dim ColFrasAgrupadas As Collection

Dim sparamalvic As ADODB.Recordset

Dim FechaFichero As Date
Dim IdTurno As Long  'Si importa un turno, puede coger DOS dias. Seran seguidos ye la fecha sera del inicio

Dim UltimoTurnoLeido2 As Long
Dim TipoFicheroNormal As Boolean
Dim Vec() As String
Dim Turno3 As Boolean



Dim EsAlvic2 As Boolean 'Llegado el caso, habra que parametrizar





Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub cmdAceptar_Click()
    cmdAceptar.Enabled = False
    HacercmdAceptar_Click
    cmdAceptar.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub HacercmdAceptar_Click()

Dim SQL As String
Dim I As Byte
Dim cadWhere As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String

On Error GoTo eError

'    GenerarFacturasScafac
    If Not DatosOk Then Exit Sub
    
    
    
    
    CommonDialog1.DefaultExt = ".TXT"
    
    CADENA = Format(CDate(txtCodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    b = False
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        InicializarTabla
        cadSelect = "INSERT INTO tmpgasolimport(codusu,codigo,NumAlbaran,NumFactura,fechahora,IdVendedor,"
        cadSelect = cadSelect & "Cliente,NombreCliente,NifCliente,Matricula,CodigoProducto"
        cadSelect = cadSelect & ",surtidor,manguera,Precio,cantidad,descuento,importel,idtipopago,tipoIVa,importeConIva,ccoste,turno,ClivarioAlvic,doc_original,doc_relacionado ) VALUES "
        cadFormula = ""
        
        '#aqui aqui aqui

          If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo ' vSesion.Codigo
                
                SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo ' vSesion.Codigo
                SQL = SQL & " and importeb1 is null "
                
                If TotalRegistros(SQL) <> 0 Then

                    MsgBox "Hay errores en el Traspaso de Postes. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de Poste"
                    cadNombreRPT = "rErroresTrasPoste3.rpt"
                    LlamarImprimir
                    Exit Sub
                Else
                    SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo ' vSesion.Codigo
                    SQL = SQL & " and importeb1 = 0 "
                    
                    If TotalRegistros(SQL) <> 0 Then
                        MsgBox "Hay errores en el Traspaso de Postes. Revise.", vbExclamation
                        cadTitulo = "Errores de Traspaso de Poste"
                        cadNombreRPT = "rErroresTrasPoste3.rpt"
                        LlamarImprimir
                    End If
                    
                    
                    
                    
                    'AJustamos IMportes para el cobro
                    lblProgres(0).Caption = "Ajuste  cobros"
                    lblProgres(0).Refresh
    
                    CadenaDesdeOtroForm = ""
                    frmListado5.OpcionListado = 30
                    frmListado5.OtrosDatos = cadParam  'lleva el total de la integracion
                    frmListado5.Show vbModal
                    If CadenaDesdeOtroForm = "" Then Err.Raise 513, , "Proceso cancelado"
                    
                    conn.BeginTrans
                    lblProgres(0).Caption = "Generando datos"
                    lblProgres(0).Refresh
                    Set ColFrasAgrupadas = New Collection
                    
                    
                    b = GenerarFacturasAlbaranes()
                    If Not b Then
                        conn.RollbackTrans
                    Else
                        
                        conn.CommitTrans
                        Screen.MousePointer = vbHourglass
                        lblProgres(0).Caption = "Asiento cobros"
                        lblProgres(0).Refresh
                        Pb1.Value = 0
                        
                        GeneraAsientoCobros
                        
                        lblProgres(0).Caption = "Creando facturas"
                        lblProgres(0).Refresh
                        GenerarFacturasScafac
                        
                        
                        If txtCodigo(1).Text <> "" Then
                            cadFormula = "UPDATE sparamalvic set ultimoturno =  " & txtCodigo(1).Text
                            ejecutar cadFormula, False
                            UltimoTurnoLeido2 = IdTurno
                             txtCodigo(1).Text = IncremetaUnTurno()
                        End If
                    End If

                End If
          End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Then
        b = False
        If Err.Number <> 32755 Then
            cadFormula = Err.Description
            If cadFormula <> "Proceso cancelado" Then cadFormula = "No se ha podido realizar el proceso. LLame a Ariadna." & vbCrLf & cadFormula
            MsgBox cadFormula, vbExclamation
        End If
    Else
        If Not b Then
            MsgBox "No se ha podido realizar el proceso. LLame a Ariadna." & vbCrLf, vbExclamation
        Else
            MsgBox "Proceso realizado correctamente.", vbInformation
        End If
    End If
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
    If b Then
        BorrarArchivo Me.CommonDialog1.FileName
        
        'cmdCancel_Click
    End If
    
End Sub


Private Function IncremetaUnTurno() As String
Dim FechaTurnoUlt As String
    If EsAlvic2 Then
        IncremetaUnTurno = UltimoTurnoLeido2 + 1
    Else
        'stop
        
        If UltimoTurnoLeido2 > 100 Then
            
            'Las dos ultimas cifras son Parte X  Cierre W
            ' yymmddXW
            FechaTurnoUlt = Right(UltimoTurnoLeido2, 1)
            If FechaTurnoUlt = 3 Then
                'Ultimo turno dia. Es dia siguiente
                FechaTurnoUlt = Mid(CStr(UltimoTurnoLeido2), 1, 6)
                FechaTurnoUlt = Mid(FechaTurnoUlt, 5, 2) & "/" & Mid(FechaTurnoUlt, 3, 2) & "/20" & Mid(FechaTurnoUlt, 1, 2)
                
                FechaTurnoUlt = DateAdd("d", 1, CDate(FechaTurnoUlt))
                IncremetaUnTurno = Format(FechaTurnoUlt, "yymmdd") & "11"
            Else
                IncremetaUnTurno = UltimoTurnoLeido2 + 1
            End If
        Else
            MsgBox "Ultimo turno guardado. " & UltimoTurnoLeido2, vbExclamation
        
        End If
        
    End If
End Function

    
Private Sub BorrarArchivo(Archivo As String)
    On Error Resume Next
    
    
    'Kill Archivo
    If Err.Number <> 0 Then MuestraError Err.Number, , Archivo
        
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub






Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
        
        
        
        
        'Vamos a ver los ivas, desde la conta
        Set sparamalvic = New ADODB.Recordset
        cadSelect = ""
        
        
        cadParam = "select * from sarticalvic where not codartic in (select codartic from sartic)"
        sparamalvic.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cadParam = ""
        While Not sparamalvic.EOF
            cadParam = cadParam & "  -  " & sparamalvic!artculoAlvic & "(" & sparamalvic!codArtic & ")" & vbCrLf
            sparamalvic.MoveNext
        Wend
        sparamalvic.Close
        
        If cadParam <> "" Then
            cadParam = "Articulos traspaso pendientes crear" & vbCrLf & cadParam
            cadSelect = cadSelect & cadParam
        End If
        '                    21%         10%         4%             0%
        cadParam = "Select ivExento,IvaNormal1,ivaReducido1,ivaSuperRed1,IvaNormal2,ivaReducido2,ivaSuperRed2,IvaNormal3,ivaReducido3,ivaSuperRed3"
        cadParam = cadParam & ",forpa ,Clivario  ,FraDirectaD,FraDirectaT,FraDirectaA,AlbTipoD,AlbTipoT,AlbTipoA,FacturaVariosD,FacturaVariosT"
        cadParam = cadParam & ",FacturaVariosA,letraGasoleo,letraTienda,letraVarios,ultimoturno,Serie1Gasol, Serie2Tienda ,Serie3Ticket"
        cadParam = cadParam & " from sparamalvic "
        sparamalvic.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If sparamalvic.EOF Then
            cadSelect = cadSelect & "Falta configurar parametros traspaso ALVIC"
        Else
        
            For indCodigo = 0 To 9
                cadFormula = "artvario"
                cadParam = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", sparamalvic.Fields(indCodigo), "T", cadFormula)
                If cadParam = "" Then
                    cadSelect = cadSelect & indCodigo & ": " & indCodigo & " sin configurar" & vbCrLf
                Else
                    If cadFormula = "0" Then
                        cadSelect = cadSelect & cadNombreRPT & " no es de varios" & vbCrLf
                    Else
                        cadFormula = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", cadParam, "N")
                        If cadFormula = "" Then
                            'NO existe t
                            cadSelect = cadSelect & cadParam & " codigiva NO existe" & vbCrLf
                        Else
                            If indCodigo <= 3 Then
                                'OK. Todo bien. Veamos porcentaje
                                If indCodigo = 2 Then
                                    'cadNombreRPT = miRsAux
                                    IvaReducido = CCur(cadFormula)
                                ElseIf indCodigo = 3 Then
                                    'cadNombreRPT = vParamAplic.GasolArticuloSuperReducido
                                    IvaSuperReducido = CCur(cadFormula)
                                ElseIf indCodigo = 0 Then
                                    'cadNombreRPT = vParamAplic.GasolArticuloExento
                                    
                                Else
                                    'indCodigo = 1
                                    IvaNormal = CCur(cadFormula)
                                End If
                            End If
                        End If
                    End If
                End If
            
            
            Next
            
            
            'La forpa de pago que van a pasar todas los datos debe existir
            cadFormula = "Forma de pago parametros"
            If Not IsNull(sparamalvic!ForPa) Then
                cadParam = DevuelveDesdeBD(conAri, "codforpa", "sforpa", "codforpa", CStr(sparamalvic!ForPa), "T")
                If cadParam <> "" Then cadFormula = ""
            End If
            If cadFormula <> "" Then cadSelect = cadSelect & cadFormula & vbCrLf
            
            
            UltimoTurnoLeido2 = DBLet(sparamalvic!ultimoturno, "N")
            If UltimoTurnoLeido2 > 0 Then txtCodigo(1).Text = IncremetaUnTurno
                
            
            
        End If
        If cadSelect <> "" Then
            MsgBox cadSelect, vbExclamation
            cmdAceptar.Enabled = False
        End If
        cadSelect = ""
        cadNombreRPT = ""
        cadFormula = ""
        cadParam = ""
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me


    EsAlvic2 = True    '######

    
    txtCodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")
     
    FrameCobrosVisible True, H, W
    Pb1.visible = False
        
   
    
    Me.cmdCancel.Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
    Set sparamalvic = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si se ha eliminado un turno, el check ha de estar desmarcado. " & vbCrLf & vbCrLf & _
                      "El motivo es porque si se ha traspasado el fichero de compras, " & vbCrLf & _
                      "los albaranes no se eliminan cuando se borra un turno." & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.Fecha = txtCodigo(Index).Text
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30


    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'If KeyAscii = teclaBuscar Then
    If Chr(KeyAscii) = "+" Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0 'fecha
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCodigo(Index).Text = Trim(txtCodigo(Index))
    Select Case Index
        Case 0 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
                    
        Case 1
            If txtCodigo(Index).Text <> "" Then
                If Not PonerFormatoEntero(txtCodigo(Index)) Then txtCodigo(Index).Text = ""
            End If
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta   FALTA###
            'cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

 

Private Function DatosOk() As Boolean
Dim b As Boolean

   b = True

   If txtCodigo(0).Text = "" And b Then
        MsgBox "El campo fecha debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    End If
    
 
     If txtCodigo(1).Text <> "" Then
        If CLng(txtCodigo(1).Text) <= UltimoTurnoLeido2 Then
            MsgBox "Turno menor que el ultimo traspasado", vbExclamation
            
            If vUsu.Nivel = 0 Then
                If MsgBox("SEGURO QUE DESEA CONTINUAR?", vbQuestion + vbYesNoCancel) <> vbYes Then b = False
            Else
                b = False
            End If
            
        End If
    End If
 
    'Algunas comprobaciones.
    'En scaalb NO puede quedar nada de la serie ALD y referenc<>'turno: '
    Codigo = "referenc like 'Turno:%' AND codtipom in ('ALD','ALB','ALW') AND 1"
    Codigo = DevuelveDesdeBD(conAri, "count(*)", "scaalb", Codigo, "1")
    If Val(Codigo) > 0 Then
        
        MsgBox "ERROR GRAVE. Datos sin traspasar del turno anterior", vbCritical
        
        If b = True Then
            b = False
            'Enero 2021
            If vUsu.Nivel = 0 Then
                Codigo = "Si se han borrado de albaranes los datos que faltan por procesar "
                Codigo = Codigo & vbCrLf & "y vienen en este fichero, NO habra problema." & vbCrLf & vbCrLf
                Codigo = Codigo & vbCrLf & "1.- Comprobar que no existe en albaranes NINGUNO de los que lleva el fichero"
                Codigo = Codigo & vbCrLf & "2.- Si se importa correctamente, hay que buscar el apunte del cobro que se genera "
                Codigo = Codigo & " y borrarlo. Ya  hizo la generación en su momento"
                Codigo = Codigo & vbCrLf & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(Codigo, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then b = True
            End If
        End If
    End If
    DatosOk = b
End Function


'
'Private Function RecuperaFichero() As Boolean
'Dim NF As Integer
'
'    RecuperaFichero = False
'    NF = FreeFile
'    Open App.Path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
'    Line Input #NF, cad
'    Close #NF
'    If cad <> "" Then RecuperaFichero = True
'
'End Function


Private Function GenerarFacturasAlbaranes() As Boolean
Dim cad As String
Dim I As Integer
Dim fin As Boolean
Dim AlbaranFactura As String
Dim SerieEnAriges As String
Dim AlbaranesFacturaAgrupada As String
Dim LEtra As String
Dim J As Integer
Dim Col As Collection
    On Error GoTo EprocesarFichero

    GenerarFacturasAlbaranes = False

    lblProgres(0).Caption = "Generando datos"
    lblProgres(0).Refresh
    
    cad = DevuelveDesdeBD(conAri, "count(*)", "tmpgasolimport", "codusu", CStr(vUsu.Codigo))
    If Val(cad) = 0 Then
        MsgBox "ERROR.  Ningun dato a traspasar", vbExclamation
        Exit Function
    End If
        
    I = CInt(Val(cad))
    
    Pb1.visible = True
    Me.Pb1.Value = 0
    Me.Pb1.Max = I
    Me.Refresh
    
    Set miRsAux = New ADODB.Recordset
    Set Col = New Collection
    
    'El proceso se divide en 3 trozos
    ' Albaranes que el cliente quiere factura en el momento. Vienen ya con su numero de factura Y albaran
    ' Albaranes con clientes que no este en el rango (100.000 al 100.011) pasaría a facturarse a final de mes con la serie F.     FAD ALD
    ' Albaranes con cliente en el rango (100.000 al 100.011) pasarían a ser facturas simplificadas.  Debemos exceptuar las facturas que sean 0  FTI  ATI
    
    
    '--------------------------------------------------------------------------------------------------------------------
    ' Facturas al momento
    '--------------------------------------------------------------------------------------------------------------------
    lblProgres(1).Caption = "fras al momento"
    lblProgres(1).Refresh
    cad = "Select numfactura from tmpgasolimport where codusu = " & vUsu.Codigo & " AND numfactura<>'' GROUP  BY numfactura "
   
    
    
    
    AlbaranFactura = ""
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!NumFactura <> AlbaranFactura Then
            If AlbaranFactura <> "" Then Col.Add AlbaranFactura
            AlbaranFactura = miRsAux!NumFactura
            'Lineas = ""
        End If
                
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If AlbaranFactura <> "" Then Col.Add AlbaranFactura
    
    
    Pb1.Value = 0
    Pb1.Max = IIf(Col.Count = 0, 1, Col.Count)
    DoEvents
    For I = 1 To Col.Count
        AlbaranFactura = ""
        GeneraFacturaMomento Col.Item(I), AlbaranFactura
        
        ColFrasAgrupadas.Add "0" & Col.Item(I) & ":" & AlbaranFactura
        Pb1.Value = I
    Next
    
    
    
    '--------------------------------------------------------------------------------------------------------------------
    ' Facturas al momento
    '--------------------------------------------------------------------------------------------------------------------
    lblProgres(1).Caption = "Albaranes fin mes"
    lblProgres(1).Refresh
    LEtra = ""
    Pb1.Value = 0
    Set Col = Nothing
    Set Col = New Collection


    For I = 1 To 2
        If I = 1 Then
            cad = "'" & sparamalvic!letraGasoleo & "%'"
        Else
            cad = "'" & sparamalvic!letraTienda & "%'"
        End If
        cad = "Select distinct numalbaran from tmpgasolimport where codusu = " & vUsu.Codigo & " AND numalbaran like " & cad    'GASOLINERA
        cad = cad & " AND numfactura is null"
        cad = cad & " AND not " & CadenaClientesVarios
        cad = cad & " GROUP  BY numalbaran  "
        
        
        AlbaranFactura = ""
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If miRsAux!NumAlbaran <> AlbaranFactura Then
                If AlbaranFactura <> "" Then Col.Add AlbaranFactura
                
                AlbaranFactura = miRsAux!NumAlbaran
                'Lineas = ""
            End If
                    
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If AlbaranFactura <> "" Then Col.Add AlbaranFactura
    
    
    Next
    
    Pb1.Value = 0
    Pb1.Max = IIf(Col.Count = 0, 1, Col.Count)
    DoEvents
    For I = 1 To Col.Count
        
        GeneraAlbaranesFinMes Col.Item(I)
        Pb1.Value = I
    Next
    
    
    
    
    '--------------------------------------------------------------------------------------------------------------------
    '   Ventas VARIOS.  DOS pasadas.  Serie D - T
    '--------------------------------------------------------------------------------------------------------------------
    For J = 1 To 2
            lblProgres(1).Caption = "Simplif Serie " & IIf(J = 1, "COMBUSTIBLE", "TIENDA")
            lblProgres(1).Refresh
            Pb1.Value = 0
            Set Col = Nothing
            Set Col = New Collection
            DoEvents
            
            LEtra = sparamalvic!letraGasoleo
            If J = 2 Then LEtra = sparamalvic!letraTienda
            
            cad = "Select distinct numalbaran from tmpgasolimport where codusu = " & vUsu.Codigo & " AND numalbaran like '" & LEtra & "%'"
            cad = cad & " AND numfactura is null"
            cad = cad & " AND " & CadenaClientesVarios
            cad = cad & " GROUP  BY numalbaran  "
            
            
            
  
            AlbaranFactura = ""
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux!NumAlbaran <> AlbaranFactura Then
                    If AlbaranFactura <> "" Then Col.Add AlbaranFactura
                    AlbaranFactura = miRsAux!NumAlbaran
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If AlbaranFactura <> "" Then Col.Add AlbaranFactura

            
            If Col.Count > 0 Then
                AlbaranesFacturaAgrupada = ""
                SerieEnAriges = DevuelveCodtipom(LEtra, True, False)
                Pb1.Max = Col.Count
                DoEvents
                
                For I = 1 To Col.Count
                    AlbaranesFacturaAgrupada = AlbaranesFacturaAgrupada & ", " & Mid(Col.Item(I), 2)   'quito la letra
                    GeneraAlbaranesTiendaAlvic2 SerieEnAriges, Col.Item(I)
                    Pb1.Value = I
                Next
            
                'Para generar despues la factura agrupada con todos los albaranes/facturas
                AlbaranesFacturaAgrupada = SerieEnAriges & "@" & Mid(AlbaranesFacturaAgrupada, 2)
                SerieEnAriges = DevuelveCodtipom(LEtra, False, True)  'SERIE SCAFAC
                ColFrasAgrupadas.Add "1" & SerieEnAriges & AlbaranesFacturaAgrupada
            
            End If
            
    Next
    
    
    '--------------------------------------------------------------------------------------------------------------------
    lblProgres(1).Caption = "Simplif Serie A"
    lblProgres(1).Refresh
    Pb1.Value = 0
    Set Col = Nothing
    Set Col = New Collection
    LEtra = sparamalvic!letraVarios
    cad = "Select distinct numalbaran from tmpgasolimport where codusu = " & vUsu.Codigo & " AND numalbaran like '" & LEtra & "%'  AND numfactura is null"
    cad = cad & " GROUP  BY numalbaran  "
    'VAN TODAS JUNTAS, varios y no varios
    
    AlbaranFactura = ""
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!NumAlbaran <> AlbaranFactura Then
            If AlbaranFactura <> "" Then Col.Add AlbaranFactura
            AlbaranFactura = miRsAux!NumAlbaran
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If AlbaranFactura <> "" Then Col.Add AlbaranFactura
    
    If Col.Count > 0 Then
        AlbaranesFacturaAgrupada = ""
        SerieEnAriges = DevuelveCodtipom(LEtra, True, False)
        Pb1.Max = Col.Count
        DoEvents
        For I = 1 To Col.Count
            GeneraAlbaranesTiendaAlvic2 SerieEnAriges, Col.Item(I)
            Pb1.Value = I
        Next
    
        'AlbaranesFacturaAgrupada = "codtipom ='" & SerieEnAriges & "' AND numalbar in (" & Mid(AlbaranesFacturaAgrupada, 2) & ")"
        AlbaranesFacturaAgrupada = SerieEnAriges & "@"   'VAN TODOS
        
        SerieEnAriges = DevuelveCodtipom(LEtra, False, True)  'SERIE SCAFAC
        ColFrasAgrupadas.Add "1" & SerieEnAriges & AlbaranesFacturaAgrupada
        
        
    End If
    
    
    
    
    cad = "select count(*) from tmpgasolimport where codusu =" & vUsu.Codigo & " and  traspasado=0"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If Not miRsAux Then
        If miRsAux.Fields(0) > 0 Then cad = miRsAux.Fields(0)
    End If
    miRsAux.Close
    If cad <> "" Then
        cad = cad & vbCrLf & vbCrLf & "Avise a Ariadna"
        MsgBox "Datos pendientes de traspasar: " & cad, vbExclamation
    Else
        GenerarFacturasAlbaranes = False
    End If
    
    
    '--------------------------------------------------------------------------------------------------------------------
    lblProgres(1).Caption = "Generando datos fras"
    lblProgres(1).Refresh
    Pb1.Value = 0
    
    'insert INTO tmpslipreu(codusu,codalmac,codartic,nomartic,ampliaci,numlinea,numofert)
    Set Col = New Collection
    For I = 1 To ColFrasAgrupadas.Count
        cad = CStr(ColFrasAgrupadas.Item(I)) & ""
        Col.Add cad
        
    Next
    Set ColFrasAgrupadas = Nothing
    cad = ""
    Codigo = ""
    LEtra = ""
    AlbaranesFacturaAgrupada = ""
    For I = 1 To Col.Count
        cadTitulo = CStr(Col.Item(I)) & " "
        Debug.Print cadTitulo
        cad = CStr(cadTitulo)
        If Mid(cad, 1, 1) = "0" Then
            'Factura directa
            J = InStr(1, cad, ":")
            If J = 0 Then Err.Raise 513, , "Leyendo datos facturas fichero: " & cadTitulo
            LEtra = Mid(cad, 3, J - 3)
            cad = Mid(cad, J + 1)
            J = 0
        Else
            'Facturas varias
            LEtra = ""
            cad = Mid(cad, 2)
            J = 1
        End If




        'tmpslipreu(codusu,codalmac,codartic
        Codigo = ", (" & vUsu.Codigo & "," & J & "," & DBSet(LEtra, "T", "N") & ",'"
        'nomartic,ampliaci,
        LEtra = Mid(cad, 1, 3)
        Codigo = Codigo & LEtra & "','"
        LEtra = Mid(cad, 4, 3)
        Codigo = Codigo & LEtra & "'"
        
        
        
        cad = Trim(Mid(cad, 8))
'        'para cada albaran lo meteremos aquin
        'numlinea,numofert
        vContad = 0
        LEtra = ""
        If cad = "" Then
'            'Es la facturacion tickets, YA QUE son tooooodoso los albaranes
            LEtra = Codigo & ",0,0)   "
        Else
            While cad <> ""
                
                vContad = vContad + 1
                J = InStr(1, cad, ",")
                If J = 0 Then
                    AlbaranesFacturaAgrupada = Trim(cad)
                    cad = ""
                    
                Else
                    AlbaranesFacturaAgrupada = Mid(cad, 1, J - 1)
                    cad = Trim(Mid(cad, J + 1))
                  
                End If
            
                If Not IsNumeric(AlbaranesFacturaAgrupada) Then Err.Raise 513, , "Albaran no numerico: " & AlbaranesFacturaAgrupada & " " & cadTitulo
                LEtra = LEtra & Codigo & "," & vContad & "," & AlbaranesFacturaAgrupada & ") "
               
                
            Wend
        End If
        LEtra = Mid(LEtra, 2)
        cad = "INSERT INTO tmpslipreu(codusu,codalmac,codartic,nomartic,ampliaci,numlinea,numofert) VALUES " & LEtra
        conn.Execute cad
        
    Next

    
    
    
    
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    GenerarFacturasAlbaranes = True
    
EprocesarFichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        GenerarFacturasAlbaranes = False
    End If
    Set Col = Nothing
    Set miRsAux = Nothing
End Function
'
'Private Sub P1()
'Dim cad As String
'Dim I As Integer
'Dim fin As Boolean
'Dim AlbaranFactura As String
'Dim SerieEnAriges As String
'Dim AlbaranesFacturaAgrupada As String
'Dim LEtra As String
'Dim J As Integer
'Dim Col As Collection
'
'
'    Set Col = New Collection
'
'    Col.Add "0E0022017:FA1ALD@ 0512214"
'    Col.Add "1FAXALD@ 512226, 512232, 512264, 512269, 512270, 512271, 512280, 512282, 512284, 512292, 512295"
'    Col.Add "1FAYALB@ 71977, 71978"
'    Col.Add "1FAWALW@"
'
'
'
'
'cad = ""
'    Codigo = ""
'    LEtra = ""
'    AlbaranesFacturaAgrupada = ""
'    For I = 1 To Col.Count
'        cadTitulo = CStr(Col.Item(I)) & " "
'        Debug.Print cadTitulo
'        cad = CStr(cadTitulo)
'        If Mid(cad, 1, 1) = "0" Then
'            'Factura directa
'            J = InStr(1, cad, ":")
'            If J = 0 Then Err.Raise 513, , "Leyendo datos facturas fichero: " & cadTitulo
'            LEtra = Mid(cad, 3, J - 3)
'            cad = Mid(cad, J + 1)
'            J = 0
'        Else
'            'Facturas varias
'            LEtra = ""
'            cad = Mid(cad, 2)
'            J = 1
'        End If
'
'
'
'
'        'tmpslipreu(codusu,codalmac,codartic
'        Codigo = ", (" & vUsu.Codigo & "," & J & "," & DBSet(LEtra, "T", "N") & ",'"
'        'nomartic,ampliaci,
'        LEtra = Mid(cad, 1, 3)
'        Codigo = Codigo & LEtra & "','"
'        LEtra = Mid(cad, 4, 3)
'        Codigo = Codigo & LEtra & "'"
'
'
'
'        cad = Trim(Mid(cad, 8))
''        'para cada albaran lo meteremos aquin
'        'numlinea,numofert
'        vContad = 0
'        LEtra = ""
'        If cad = "" Then
''            'Es la facturacion tickets, YA QUE son tooooodoso los albaranes
'            LEtra = Codigo & ",0,0)   "
'        Else
'            While cad <> ""
'
'                vContad = vContad + 1
'                J = InStr(1, cad, ",")
'                If J = 0 Then
'                    AlbaranesFacturaAgrupada = Trim(cad)
'                    cad = ""
'
'                Else
'                    AlbaranesFacturaAgrupada = Mid(cad, 1, J - 1)
'                    cad = Trim(Mid(cad, J + 1))
'
'                End If
'
'                If Not IsNumeric(AlbaranesFacturaAgrupada) Then Err.Raise 513, , "Albaran no numerico: " & AlbaranesFacturaAgrupada & " " & cadTitulo
'                LEtra = LEtra & Codigo & "," & vContad & "," & AlbaranesFacturaAgrupada & ") "
'
'
'            Wend
'        End If
'        LEtra = Mid(LEtra, 2)
'        cad = "INSERT INTO tmpslipreu(codusu,codalmac,codartic,nomartic,ampliaci,numlinea,numofert) VALUES " & LEtra
'        conn.Execute cad
'
'    Next
'
'
'
'
'
'
'End Sub
                
                
                
Private Sub GeneraFacturaMomento(Factura As String, ByRef NumeroFactura As String)
Dim Codtipoa As String
Dim LetraAlb As String
    
    
    
    Codigo = "select * from tmpgasolimport left join sclien on cliente=codclien left join sforpa on idtipopago=sforpa.codforpa "
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " and numfactura= " & DBSet(Factura, "T")
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Codigo = ""
    
    If miRsAux.EOF Then Err.Raise 513, , "Sin albaranes para la factura: " & Factura
    LetraAlb = Mid(miRsAux!NumAlbaran, 1, 1)
    Do
        
        If miRsAux!NumAlbaran <> Codigo Then
            If Codigo <> "" Then
                miRsAux.MovePrevious
                Codtipoa = DevuelveCodtipom(miRsAux!NumAlbaran, True, False)
                CrearAlbaran Codtipoa, False    '
                miRsAux.MoveNext
            Else
                
            End If
            Codigo = miRsAux!NumAlbaran
        End If
        NumeroFactura = NumeroFactura & ", " & Mid(miRsAux!NumAlbaran, 2)
        
        If LetraAlb <> Mid(miRsAux!NumAlbaran, 1, 1) Then
            NumeroFactura = "Letras distintas para una misma factura. " & Factura & vbCrLf & NumeroFactura
            Err.Raise 513, , NumeroFactura
        End If
        
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
        
    miRsAux.MovePrevious
    Codtipoa = DevuelveCodtipom(miRsAux!NumAlbaran, True, False)
    CrearAlbaran Codtipoa, False    'las facturas al momento SON FMO  - >letra de serie la de stipom
    miRsAux.Close
    
    
    'Crearemos las lineas de albaran
    Codigo = "select * from tmpgasolimport  WHERE  codusu=" & vUsu.Codigo & " and numfactura= " & DBSet(Factura, "T") & " ORDEr BY codigo"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CrearLineasAlbaran Codtipoa
        
    
    
    NumeroFactura = Mid(NumeroFactura, 2)
    NumeroFactura = Codtipoa & "@" & NumeroFactura
    Codtipoa = DevuelveCodtipom(LetraAlb, False, False)
    NumeroFactura = Codtipoa & NumeroFactura
        
    




End Sub
                
                
                
                
                
Private Sub GeneraAlbaranesFinMes(Numalbar As String)
Dim Codtipoa As String

    'Crearemos el/los albaranes
    Codigo = "select * from tmpgasolimport left join sclien on cliente=codclien left join sforpa on idtipopago=sforpa.codforpa "
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " and numfactura is null and numalbaran=  " & DBSet(Numalbar, "T")
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Codigo = ""
    If miRsAux.EOF Then Err.Raise 513, , "No existe albaran: " & Numalbar
    Do
        If miRsAux!NumAlbaran <> Codigo Then
            If Codigo <> "" Then
                miRsAux.MovePrevious
                Codtipoa = DevuelveCodtipom(miRsAux!NumAlbaran, True, False)
                CrearAlbaran Codtipoa, False
                miRsAux.MoveNext
            End If
            Codigo = miRsAux!NumAlbaran
        End If
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
        
    miRsAux.MovePrevious
    Codtipoa = DevuelveCodtipom(miRsAux!NumAlbaran, True, False)
    CrearAlbaran Codtipoa, False
    miRsAux.Close
    
    
    'Crearemos las lineas de albaran
    Codigo = "select * from tmpgasolimport WHERE  codusu=" & vUsu.Codigo & " and numfactura is null and numalbaran=  " & DBSet(Numalbar, "T")
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CrearLineasAlbaran Codtipoa
        
    
    
    
End Sub
    
    
' Facturas albaranes socios que no son de gasolina
' una factura por socio
Private Sub GeneraAlbaranesTiendaAlvic2(SerieAriges As String, Factura As String)

    'Crearemos el/los albaranes
    Codigo = "select * from tmpgasolimport left join sclien on cliente=codclien left join sforpa on idtipopago=sforpa.codforpa "
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " and numalbaran= " & DBSet(Factura, "T")
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Codigo = ""
    If miRsAux.EOF Then Err.Raise 513, , "Sin albaranes para la factura: " & Factura
    Do
        If miRsAux!NumAlbaran <> Codigo Then
            If Codigo <> "" Then
                miRsAux.MovePrevious
                CrearAlbaran SerieAriges, True     'Aunque luego hagamos una UNICA factura con todas los albaranes(facturas) que insertemos
                miRsAux.MoveNext
            End If
            Codigo = miRsAux!NumAlbaran
        End If
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
        
    miRsAux.MovePrevious
    CrearAlbaran SerieAriges, True
    miRsAux.Close
    
    
    'Crearemos las lineas de albaran
    Codigo = "select * from tmpgasolimport  WHERE  codusu=" & vUsu.Codigo & " and numalbaran= " & DBSet(Factura, "T") & " ORDEr BY codigo"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CrearLineasAlbaran SerieAriges
        
    

End Sub
                
                
'Tickets.  Se corresponden con los albaranes A00??? del fichero de traspaso
'   Van todos juntos a contado
Private Sub GeneraAlbaranesTicketsAlvic(Factura As String)
Dim codtipom As String
    'Crearemos el/los albaranes
    codtipom = "AL" & sparamalvic!FraTickets  'sparamalvic FraCombustible
    Codigo = "select * from tmpgasolimport left join sclien on cliente=codclien left join sforpa on idtipopago=sforpa.codforpa "
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " and numalbaran= " & DBSet(Factura, "T")
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Codigo = ""
    If miRsAux.EOF Then Err.Raise 513, , "Sin albaranes para la factura: " & Factura
    Do
        If miRsAux!NumAlbaran <> Codigo Then
            If Codigo <> "" Then
                miRsAux.MovePrevious
                CrearAlbaran codtipom, True     'Aunque luego hagamos una UNICA factura con todas los albaranes(facturas) que insertemos
                miRsAux.MoveNext
            End If
            Codigo = miRsAux!NumAlbaran
        End If
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
        
    miRsAux.MovePrevious
    CrearAlbaran codtipom, True
    miRsAux.Close
    
    
    'Crearemos las lineas de albaran
    Codigo = "select * from tmpgasolimport  WHERE  codusu=" & vUsu.Codigo & " and numalbaran= " & DBSet(Factura, "T") & " ORDEr BY codigo"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CrearLineasAlbaran codtipom
        
    

End Sub
                
Private Sub CrearFacturaTiendaTickets()

End Sub
                
                
                
                
                
                
                
' Tickets y resto vetnas de ALVIC.
' Factura UNICA
Private Sub GeneraFacturaTickets(Factura As String)

    'Crearemos el/los albaranes
    Codigo = "select * from tmpgasolimport left join sclien on cliente=codclien left join sforpa on idtipopago=sforpa.codforpa "
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " and numfactura= " & DBSet(Factura, "T")
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Codigo = ""
    If miRsAux.EOF Then Err.Raise 513, , "Sin albaranes para la factura: " & Factura
    Do
        If miRsAux!NumAlbaran <> Codigo Then
            If Codigo <> "" Then
                miRsAux.MovePrevious
                CrearAlbaran "ATI", True     'las facturas al momento SON FMO  - >letra de serie la de stipom
                miRsAux.MoveNext
            End If
            Codigo = miRsAux!NumAlbaran
        End If
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
        
    miRsAux.MovePrevious
    CrearAlbaran "ATI", True     'las facturas al momento SON FMO  - >letra de serie la de stipom
    miRsAux.Close
    
    
    'Crearemos las lineas de albaran
    Codigo = "select * from tmpgasolimport  WHERE  codusu=" & vUsu.Codigo & " and numfactura= " & DBSet(Factura, "T") & " ORDEr BY codigo"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CrearLineasAlbaran "ATI"
        
    
    
    
    
    
        
    'Lo pasamos a factura




End Sub
 
                
                
                
                
                
                
                
                
                
                
'Estara abierto mirsaux con los datos desde ALVIC, cruzados con la sclien
Private Sub CrearAlbaran(codtipom As String, PonerReferencia As Boolean)
Dim vSQL As String
Dim Clivario As Boolean
Dim RVario As ADODB.Recordset
Dim DesdeClivar As Boolean

    Clivario = False
    DesdeClivar = False
    If miRsAux!codClien >= 100000 And miRsAux!codClien <= 100011 Then
        If miRsAux!codClien <> 100007 Then Clivario = True
    End If
        
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien"
    vSQL = vSQL & ",nifclien,telclien,coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,"
    vSQL = vSQL & "dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,"
    vSQL = vSQL & "numpedcl,fecpedcl,fecentre,sementre,coddiren,tipAlbaran,codzonas,fecenvio,codinter,codnatura,chofer) VALUES ("
    vSQL = vSQL & "'" & codtipom & "'," & Mid(miRsAux!NumAlbaran, 2) & ", " & DBSet(miRsAux!FechaHora, "F") & ",1,"
    
    If Clivario Then
        vSQL = vSQL & sparamalvic!Clivario
        If DBLet(miRsAux!ClivarioAlvic, "T") <> "" Then DesdeClivar = True
    Else
        vSQL = vSQL & miRsAux!codClien
    End If
    
    
    If DesdeClivar Then
        'Clientes VARIOS. Factura identificada
        cadParam = "Select * from sclvar where nifclien=" & DBSet(miRsAux!NifCliente, "T")
        Set RVario = New ADODB.Recordset
        RVario.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RVario.EOF Then
            
            vSQL = vSQL & "," & DBSet(miRsAux!NombreCliente, "T") & ",'N/D','46'"
            vSQL = vSQL & ",' ',' ',"
            'Lo insertaremos
            cadParam = "INSERT IGNORE INTO sclvar(nifClien , NomClien, domclien, codpobla, pobclien, proclien, telclien, observa)"
            cadParam = cadParam & " VALUES ('" & miRsAux!NifCliente & "'," & DBSet(miRsAux!NombreCliente, "T")
            cadParam = cadParam & " ,'N/D','46','N/D','N/D',''," & DBSet(miRsAux!ClivarioAlvic, "T") & ")"
            conn.Execute cadParam
            cadParam = ""
        Else
            'Tiene los datos
            vSQL = vSQL & "," & DBSet(RVario!NomClien, "T") & "," & DBSet(RVario!domclien, "T", "N") & "," & DBSet(RVario!codpobla, "N")
            vSQL = vSQL & "," & DBSet(RVario!pobclien, "T", "N") & "," & DBSet(RVario!proclien, "T", "N") & ","
        
        End If
        RVario.Close
        Set RVario = Nothing
    Else
        vSQL = vSQL & "," & DBSet(miRsAux!NomClien, "T") & "," & DBSet(miRsAux!domclien, "T", "N") & "," & DBSet(miRsAux!codpobla, "N")
        vSQL = vSQL & "," & DBSet(miRsAux!pobclien, "T", "N") & "," & DBSet(miRsAux!proclien, "T", "N") & ","
    End If
    If Clivario Then
        If DesdeClivar Then
            vSQL = vSQL & DBSet(miRsAux!NifCliente, "T")
        Else
            vSQL = vSQL & DBSet(vParam.CifEmpresa, "T")
        End If
    Else
        If DBLet(miRsAux!nifClien, "T") = "" Then Err.Raise 513, , "Nif vacio: " & miRsAux!codClien
        vSQL = vSQL & DBSet(miRsAux!nifClien, "T")
    End If
    vSQL = vSQL & "," & DBSet(miRsAux!telclie1, "T") & ",NULL,NULL"   'coddirec nomidirec
    'referencia
    vSQL = vSQL & "," & IIf(PonerReferencia, DBSet("Turno:" & DBLet(miRsAux!turno, "T"), "T"), "NULL")
    vSQL = vSQL & "," & miRsAux!IdVendedor & "," & miRsAux!IdVendedor & "," & miRsAux!IdVendedor
    
    vSQL = vSQL & "," & DBSet(miRsAux!CodAgent, "T") & ","
    
    'La forma de pago será la de parametros. Pero para tener constancia de la original,guardo en observaciones 3
    'PEEERO en credito SI que dejo a credito
    cadParam = Format(miRsAux!idtipopago, "000") & " " & miRsAux!nomforpa
    
    If miRsAux!idtipopago = 2 Then
        vSQL = vSQL & "2"
    Else
        vSQL = vSQL & sparamalvic!ForPa
    End If
    
    
    
    vSQL = vSQL & "," & DBSet(miRsAux!CodEnvio, "N") & ",0,0," & DBSet(miRsAux!TipoFact, "N")
    '           observa 1   2   y 3
    vSQL = vSQL & "," & DBSet(miRsAux!NumAlbaran, "T") & "," & DBSet(miRsAux!NumFactura, "T") & "," & DBSet(cadParam, "T") & ","
    cadParam = ""
    '               observa 4 y 5
    If DBLet(miRsAux!Matricula, "T") = "" Then
        vSQL = vSQL & "NULL"
    Else
        vSQL = vSQL & DBSet(miRsAux!Matricula, "T")
    End If
    If PonerReferencia Then
        vSQL = vSQL & "," & DBSet(DBLet(miRsAux!NomClien, "T") & " " & DBLet(miRsAux!nifClien, "T"), "T")
    Else
        vSQL = vSQL & ",null"
    End If
    vSQL = vSQL & ",null,null,null,null"
    vSQL = vSQL & ",null,null,null,0," & DBSet(miRsAux!codzonas, "T")
    
    'fecenvio 'codinter,codnatura,chofer
    vSQL = vSQL & ",null,null,null,null)"
    
    'Insertar Cabecera
    conn.Execute vSQL, , adCmdText
    
End Sub

             
             
'Estara abierto mirsaux con los datos desde ALVIC, cruzados con la sclien
Private Sub CrearLineasAlbaran(codtipom As String)
Dim vSQL As String
Dim traspasdado As String
    
    
    
    
    
    
    'slialbcodtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codproveX,codccost,precoste,dtoCantidad
    Codigo = ""     'slialb
    cadSelect = "" 'smoval
    vContad = 0
    NumRegElim = CLng(Mid(miRsAux!NumAlbaran, 2))
    traspasdado = ""
    While Not miRsAux.EOF
        vContad = vContad + 1
        
        traspasdado = traspasdado & ", " & miRsAux!Codigo
      
       ' If miRsAux!TipoIVA > 1 Then S top
      
        cadTitulo = DevuelveArticuloAlbaran(Mid(miRsAux!NumAlbaran, 1, 1), miRsAux!TipoIVA)
        
        '           codtipom,numalbar,numlinea,codalmac,codartic      'codalnmac:2
        cad = ", ('" & codtipom & "'," & NumRegElim & "," & vContad & ",11," & DBSet(cadTitulo, "T") & ","
        ',nomartic,,ampliaci,
        cad = cad & DBSet(miRsAux!CodigoProducto, "T") & ","
        cadFormula = ""
        If DBLet(miRsAux!Matricula, "T") <> "" Then cadFormula = cadFormula & "   Matr. " & miRsAux!Matricula
        If DBLet(miRsAux!surtidor, "N") <> "" Then cadFormula = cadFormula & "   Surtidor " & miRsAux!surtidor & "-" & miRsAux!manguera
        If cadFormula <> "" Then cadFormula = miRsAux!FechaHora & cadFormula
        cad = cad & DBSet(cadFormula, "T") & ","
        
        'cantidad , NumBultos, precioar, dtoline1, dtoline2, ImporteL, origpre, codproveX, CodCCost ,precoste
        cad = cad & DBSet(miRsAux!cantidad, "N") & ",1," & DBSet(miRsAux!Precio, "N") & ","
        cad = cad & "0"     'DBSet(miRsAux!descuento, "N")
        cad = cad & ",0," & DBSet(miRsAux!ImporteL, "N") & ",'T',1,"
        
        
        
        cad = cad & DBSet(miRsAux!ccoste, "T") & "," & DBSet(miRsAux!importeConIva, "N") & "," & DBSet(miRsAux!descuento, "N") & ")"
        Codigo = Codigo & cad
        
        'If miRsAux!descuento > 0 Then MsgBox """"""
        'smoval
        'codartic ,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,document,numlinea
        cad = ", (" & DBSet(cadTitulo, "T") & ",11," & DBSet(miRsAux!FechaHora, "F") & "," & DBSet(miRsAux!FechaHora, "FH") & ",0,'" & codtipom & "',"
        cad = cad & DBSet(miRsAux!cantidad, "N") & "," & DBSet(miRsAux!ImporteL, "N") & "," & miRsAux!Cliente & "," & Format(NumRegElim, "0000000") & "," & vContad & ")"
        cadSelect = cadSelect & cad
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Codigo = Mid(Codigo, 2)
    cadSelect = Mid(cadSelect, 2)
    
    
    cad = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,"
    cad = cad & "importel,origpre,codproveX,codccost,precoste,dtoCantidad) VALUES " & Codigo
    
    'Insertar Cabecera
    conn.Execute cad
    
    'smoval
    cad = "INSERT INTO smoval (codartic ,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,document,numlinea) "
    cad = cad & " VALUES " & cadSelect
    conn.Execute cad
    
    
    cad = " codigo in (" & Mid(traspasdado, 2) & " )"
    cad = "UPDATE tmpgasolimport SET traspasado=1 WHERE codusu = " & vUsu.Codigo & " AND  " & cad
    conn.Execute cad
    
    
End Sub
             
             
Private Function DevuelveArticuloAlbaran(SerieAlbaran As String, TipoIVA As Byte) As String
Dim C As String
Dim C2 As String

    On Error Resume Next
    
    'Si es articulo "controlado"
    If Not IsNull(miRsAux!codArtic) Then
        DevuelveArticuloAlbaran = miRsAux!codArtic
        Exit Function
    End If
    
    C = ""
    If SerieAlbaran = sparamalvic!serie1gasol Then
        'OK
        C = "1"
    ElseIf SerieAlbaran = sparamalvic!Serie2Tienda Then
        C = "2"
    ElseIf SerieAlbaran = sparamalvic!Serie3Ticket Then C = "3"
    End If
    
    If TipoIVA = 1 Then
        C = "IvaNormal" & C
    ElseIf TipoIVA = 2 Then
        C = "ivaReducido" & C
    ElseIf TipoIVA = 3 Then
        C = "ivaSuperRed" & C
    ElseIf TipoIVA = 4 Then
        C = "ivexento"
    End If
    
    DevuelveArticuloAlbaran = sparamalvic.Fields(C)
    If Err.Number <> 0 Then
        'ERROR
        Err.Raise 513, , "ArticuloxAlbaran.  No se encuentra " & SerieAlbaran & " IVA: " & TipoIVA
    
    End If
End Function
                
                
                
                
Public Function DevuelveCodtipom(ByVal LEtra As String, ParaAlbaran As Boolean, EsFacturaVarios As Boolean) As String
    
    LEtra = Mid(LEtra, 1, 1)
    'ALBARAN
    'FraDirectaD FraDirectaT FraDirectaA AlbTipoD AlbTipoT AlbTipoA letraGasoleo letraTienda letraVarios
    If ParaAlbaran Then
        If LEtra = sparamalvic!letraGasoleo Then
            DevuelveCodtipom = sparamalvic!AlbTipoD
        ElseIf LEtra = sparamalvic!letraTienda Then
            DevuelveCodtipom = sparamalvic!AlbTipot
        ElseIf LEtra = sparamalvic!letraVarios Then
            DevuelveCodtipom = sparamalvic!AlbTipoa
        Else
            Err.Raise 513, , "Letra albaranes ALVIC no contemplada"
        End If
    Else
        
        If EsFacturaVarios Then
            'FacturaVariosD FacturaVariosY FacturaVariosA
            If LEtra = sparamalvic!letraGasoleo Then
                DevuelveCodtipom = sparamalvic!FacturaVariosD
            ElseIf LEtra = sparamalvic!letraTienda Then
                DevuelveCodtipom = sparamalvic!FacturaVariosT
            ElseIf LEtra = sparamalvic!letraVarios Then
                DevuelveCodtipom = sparamalvic!FacturaVariosA
            Else
                Err.Raise 513, , "Letra facturas ALVIC no contemplada"
            End If
        
        
        Else
            
            If LEtra = sparamalvic!letraGasoleo Then
                DevuelveCodtipom = sparamalvic!FraDirectaD
            ElseIf LEtra = sparamalvic!letraTienda Then
                DevuelveCodtipom = sparamalvic!FraDirectaT
            ElseIf LEtra = sparamalvic!letraVarios Then
                DevuelveCodtipom = sparamalvic!FraDirectaA
            Else
                Err.Raise 513, , "Letra facturas ALVIC no contemplada"
            End If
        End If
    End If
    
End Function

                
                
                
                
                
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim Longitud As Long
Dim SQL As String
Dim Sql1 As String
Dim b As Boolean
Dim CodCCost As String
Dim Impor1 As Currency
Dim Tot As Currency
Dim jj As Byte
Dim R2 As ADODB.Recordset

    On Error GoTo eProcesarFichero2
    
    IdTurno = 0
    If txtCodigo(1).Text <> "" Then IdTurno = CLng(txtCodigo(1).Text)
    
    FechaFichero = CDate("01/01/2000")
    Turno3 = False
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF
    
    
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    Longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = Longitud
    Me.Refresh
    Me.Pb1.Value = 0
    
    
    TipoFicheroNormal = chkSeparadoTabulador.Value = 0
    
    
    
  
        
    
    
    NumRegElim = 0
    vContad = 0
    CodCCost = DevuelveDesdeBD(conConta, "codccost", "ccoste", "1", "1 ORDER BY nomccost DESC", "N")
    Do
        
        
        If EOF(NF) Then
            b = False
    
        Else
            I = I + 1
            vContad = vContad + 1   'Para hacer coincidir la linea, con el registro
            
            
            Line Input #NF, cad
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & I
        
            If Not TipoFicheroNormal And I = 1 Then
                cad = ""
                b = True
            End If
              
            
            If cad <> "" Then b = ComprobarRegistroLineaFichero(cad, CodCCost)
            If Not b Then
                I = 0
                cadFormula = "" 'Ha habido error
            Else
                If Len(cadFormula) > 2000 Then
                    cadFormula = Mid(cadFormula, 2)
                    cadFormula = cadSelect & cadFormula
                    
                    conn.Execute cadFormula
                    cadFormula = ""
                End If
            End If
        End If
    Loop Until Not b
    Close #NF
    
    If cadFormula <> "" Then
        cadFormula = Mid(cadFormula, 2)
        cadFormula = cadSelect & cadFormula
        conn.Execute cadFormula
    End If
    
    
    
        
    
    
    
    
    
    
    '---------------------------------------------------
    'Unas cuantas comprobaciones
    If I > 0 Then
        lblProgres(0).Caption = "Comprobaciones BD"
        lblProgres(0).Refresh
        lblProgres(1).Caption = "Leyendo"
        lblProgres(1).Refresh
        
        'El fichero lo HA procesado OK
        Set miRsAux = New ADODB.Recordset
        SQL = "Select distinct idtipopago FROM tmpgasolimport  WHERE codusu = " & vUsu.Codigo 'IdVendedor
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = miRsAux.Fields(0)                                             'quito la f
            Sql1 = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "idForpaT", Mid(SQL, 2), "N")
            If Sql1 = "" Then
                cad = "No existe la forma de pago Alvic"
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'0'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1," & DBSet(miRsAux.Fields(0), "T") & "," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
                
                conn.Execute SQL

            Else
                Sql1 = "UPDATE tmpgasolimport SET idtipopago =" & Sql1 & " WHERE codusu = " & vUsu.Codigo & " AND idtipopago = '" & miRsAux.Fields(0) & "'"
                conn.Execute Sql1
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        SQL = "Select distinct IdVendedor FROM tmpgasolimport  WHERE codusu = " & vUsu.Codigo
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = miRsAux.Fields(0)                                             'quito la f
            Sql1 = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", Mid(SQL, 2), "N")
            If Sql1 = "" Then
                cad = "No existe el trabajador Alvic"
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'0'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1," & DBSet(miRsAux.Fields(0), "T") & "," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
                
                conn.Execute SQL

            Else
                Sql1 = "UPDATE tmpgasolimport SET IdVendedor =" & Sql1 & " WHERE codusu = " & vUsu.Codigo & " AND IdVendedor = '" & miRsAux.Fields(0) & "'"
                conn.Execute Sql1
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
            
        lblProgres(1).Caption = "clientes"
        lblProgres(1).Refresh
            
        'Veremos que todos los clientes que viene, el nif es el mismo que en tabla clientes
        SQL = "select  codigo,nifcliente , nifclien FROM  tmpgasolimport  left join sclien on cliente=codclien"
        SQL = SQL & " Where CodUsu = " & vUsu.Codigo & " And codClien >= 0 and nifcliente <> nifclien"
        SQL = SQL & " AND NOT " & CadenaClientesVarios & " ORDER BY codigo"
        
        
        
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'select * from sclien where nifclien ='Y1184807E'
            
            
            cad = "NIF distinto " & DBLet(miRsAux!NifCliente, "T") & " // " & DBLet(miRsAux!nifClien, "T")
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & ",'" & miRsAux!Codigo & "'," & DBSet(Me.txtCodigo(0).Text, "F")
            SQL = SQL & ",23,59,-1," & DBSet(miRsAux.Fields(0), "T") & "," & _
                    DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"

            conn.Execute SQL

            miRsAux.MoveNext
        Wend
        miRsAux.Close
            
            
            
            
        lblProgres(1).Caption = "cuadre importes"
        lblProgres(1).Refresh
        'Por si queremos coger datos de aqui
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            SQL = "Select distinct substring(numalbaran,1,1) FROM tmpgasolimport  WHERE codusu = " & vUsu.Codigo
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            'letraGasoleo  letraTienda   letraVarios
            While Not miRsAux.EOF
                If miRsAux.Fields(0) <> sparamalvic!letraGasoleo Then
                    If miRsAux.Fields(0) <> sparamalvic!letraTienda Then
                        If miRsAux.Fields(0) <> sparamalvic!letraVarios Then SQL = miRsAux.Fields(0) & "  "
                    End If
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If SQL <> "" Then
                cad = "Series" & Trim(SQL)
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'0'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1,'Albaran'," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
                
                conn.Execute SQL
            End If
            
        End If
            
            
        lblProgres(1).Caption = "Articulos ALVIC tratados"
        lblProgres(1).Refresh
        'Por si queremos coger datos de aqui
        SQL = "select * from sarticalvic where artculoAlvic in "
        SQL = SQL & " (Select codigoproducto  FROM  tmpgasolimport WHERE codusu = " & vUsu.Codigo & ")"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
            'Comprobamos que esta con el tipo de iva del articulo
            SQL = "tipoiva <> " & miRsAux!IVA & " AND codartic = " & DBSet(miRsAux!artculoAlvic, "T") & " AND codusu"
            SQL = DevuelveDesdeBD(conAri, "count(*)", "tmpgasolimport", SQL, CStr(vUsu.Codigo))
            If Val(SQL) > 0 Then
                cad = "Articulo tratado codigo iva distinto. " & miRsAux!artculoAlvic
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'0'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1,'IVA'," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
                conn.Execute SQL
            Else
                SQL = "UPDATE tmpgasolimport set codartic = " & DBSet(miRsAux!codArtic, "T") & " WHERE "
                SQL = SQL & " codusu =" & vUsu.Codigo & " AND codigoproducto = " & DBSet(miRsAux!artculoAlvic, "T")
                conn.Execute SQL
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
            
            
        'Clientes que voy a cREAR autmoaticamente
        'Seran de tres tipos
        '               1.-  Mayores que 128000. DIRECTAMENTE los creamos
        '               2.-  Entre 60000 y 65000
        '               3-4  Entre 0 y 60000 y entre 65000 y 128000
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            For jj = 1 To 4
                DoEvents
                lblProgres(1).Caption = "Clientes nuevos: " & jj
                lblProgres(1).Refresh
    
                If jj = 1 Then
                    SQL = " AND cliente >128000 "
                ElseIf jj = 2 Then
                    SQL = " AND cliente between 60001 AND 62998 "
                ElseIf jj = 3 Then
                    SQL = " AND  cliente < 60001 "
                Else
                    SQL = " AND cliente between  62998 and 128000"
                End If
                SQL = "Select Cliente ,NombreCliente ,NifCliente  FROM  tmpgasolimport WHERE codusu = " & vUsu.Codigo & SQL & "  and not cliente in"
                SQL = SQL & " (select codclien from sclien ) GROUP BY cliente ORDER BY cliente"
                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                While Not miRsAux.EOF
                    If jj = 1 Then
                        'Son nuevos . De momento los creo automaticamente
                        'Los creo en contablididad
                        SQL = "insert into ariconta" & vParamAplic.NumeroConta & ".cuentas(codmacta,nommacta,razosoci,apudirec,nifdatos) VALUES ("
                        SQL = SQL & "'43" & Right("00000000" & miRsAux!Cliente, 8) & "'," & DBSet(miRsAux!NombreCliente, "T")
                        SQL = SQL & "," & DBSet(miRsAux!NombreCliente, "T") & ",'S' ," & DBSet(miRsAux!NifCliente, "T") & ")"
                        ejecutar SQL, False
                        
                        
                        
                        
                        'Y en la gestion
                        SQL = "insert into sclien(codclien,nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,fechaalt,codactiv,codenvio,codzonas,codrutas,codagent"
                        SQL = SQL & " ,codforpa,codmacta,maiclie1,visitador,codtarif,tipocredito) VALUES ("
                        SQL = SQL & miRsAux!Cliente & "," & DBSet(miRsAux!NombreCliente, "T") & "," & DBSet(miRsAux!NombreCliente, "T") & ", 'N/D' ,46000 ,"
                        SQL = SQL & "'Valencia' ,'VALENCIA' ," & DBSet(miRsAux!NifCliente, "T", "N") & " ," & DBSet(Now, "F") & " ,"
                        SQL = SQL & vParamAplic.PorDefecto_Activ & " , " & vParamAplic.PorDefecto_Envio & " ," & vParamAplic.PorDefecto_Zona & " ,"
                        SQL = SQL & vParamAplic.PorDefecto_Ruta & " ," & vParamAplic.PorDefecto_Agente & " , 1 , "
                        SQL = SQL & "'43" & Right("00000000" & miRsAux!Cliente, 8) & "',null ,1 ,1,9)"
                        ejecutar SQL, False
        
        
                     Else
                        
                        
                        If jj = 2 Then
                            'Vehiculos de socios PARTICULARs, es decir , las facturas no se las pueden desgrabar
                            SQL = "insert into sclien(codclien,nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,fechaalt,codactiv,codenvio,codzonas,codrutas,codagent"
                            SQL = SQL & " ,codforpa,codmacta,maiclie1,visitador,codtarif,tipocredito) SELECT "
                            SQL = SQL & miRsAux!Cliente & " codclien,nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,fechaalt,codactiv,codenvio,codzonas,codrutas,codagent,codforpa,"
                            SQL = SQL & "'43" & Right("00000000" & miRsAux!Cliente, 8) & "' codmacta"
                            SQL = SQL & " ,maiclie1,visitador,codtarif,tipocredito "
                            SQL = SQL & " FROM sclien where codclien =" & miRsAux!Cliente - 60000 & " AND nifclien = " & DBSet(miRsAux!NifCliente, "T")
                            
                        Else
                            
                            SQL = "INSERT  INTO sclien(codclien,nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,fechaalt,codactiv,codenvio,codzonas,codrutas,codagent"
                            SQL = SQL & " ,codforpa,codmacta,maiclie1,visitador,codtarif)"
                            SQL = SQL & " SELECT codigo,NOMBRE,NOMBRE,DIRECCION dirdatos,CP codposta,POBLACION despobla,PROVINCIA desprovi,NIF nifclien,'2020-01-01' fechaalt,"
                            SQL = SQL & " IF(tipocli='TAXISTA SOCIO',3,1)  codactiv, 1 codenvio,1 codzonas,1 codrutas,1 codagent , IF(COALESCE(cred,'')<>'CREDITO',1,2) codforpa,"
                            SQL = SQL & " CONCAT('43',RIGHT(CONCAT('00000000',codigo),8)) ,NULL maiclie1,1 visitador,1 codtarif"
                            SQL = SQL & " FROM wrk_clientes_almacen WHERE"
                            SQL = SQL & " codigo = " & miRsAux!Cliente & " AND nif=" & DBSet(miRsAux!NifCliente, "T")
                            
                            
                        End If
                        ejecutar SQL, False
                        
                        
                        'En la  conta  ******
                        If jj = 2 Then
                            'Vehiculos de socios PARTICULARs, es decir , las facturas no se las pueden desgrabar
                            SQL = " SELECT '43" & Right("00000000" & miRsAux!Cliente, 8) & "' codmacta,nomclien,nomcomer,'S' apudirec,nifclien nifdatos,domclien dirdatos,codpobla codposta,pobclien despobla,proclien desprovi "
                            SQL = SQL & " FROM sclien where codclien =" & miRsAux!Cliente - 60000 & " AND nifclien = " & DBSet(miRsAux!NifCliente, "T")
                        
                        
                            SQL = "INSERT IGNORE INTO ariconta1.cuentas(codmacta,nommacta,razosoci,apudirec, nifdatos,dirdatos,codposta,despobla,desprovi)" & SQL
                        Else
                            SQL = "INSERT IGNORE INTO ariconta1.cuentas(codmacta,nommacta,razosoci,apudirec,nifdatos,dirdatos,codposta,despobla,desprovi)"
                            SQL = SQL & " SELECT CONCAT('43',RIGHT(CONCAT('00000000',codigo),8)),NOMBRE,NOMBRE, 'S' ,NIF nifdatos,DIRECCION dirdatos,"
                            SQL = SQL & " cp codposta,POBLACION despobla,PROVINCIA desprovi"
                            SQL = SQL & "  FROM wrk_clientes_almacen  WHERE "
                            SQL = SQL & " codigo = " & miRsAux!Cliente & " AND nif=" & DBSet(miRsAux!NifCliente, "T")

                        
                        End If
                        ejecutar SQL, False
                    
                                    
                                    
                                    
                                    
                                    
                    
                    End If
                    miRsAux.MoveNext
                    
                Wend
                miRsAux.Close
            
            Next jj
        End If
        If SQL <> "" Then Espera 1

            
        lblProgres(1).Caption = "Resto"
        lblProgres(1).Refresh

        SQL = "Select cliente,min(codigo) linea,min(nifcliente) nif  FROM  tmpgasolimport WHERE codusu = " & vUsu.Codigo & "  and not cliente in"
        SQL = SQL & " (select codclien from sclien ) GROUP BY cliente ORDER BY cliente"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
'            'select * from sclien where nifclien ='Y1184807E'
'            cad = DBLet(miRsAux!NIF, "T")
'            If cad <> "" Then
'                cad = DevuelveDesdeBD(conAri, "codclien", "sclien", "nifclien", cad, "T")
'            End If
'            If cad = "" Then
            
                cad = "No existe el cliente Ariges"
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'" & miRsAux!linea & "'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1," & DBSet(miRsAux.Fields(0), "T") & "," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
            
            
           ' Else
           '
           '
           '     SQL = "UPDATE tmpgasolimport SET cliente = " & cad & " WHERE codusu =" & vUsu.Codigo & " AND cliente =" & miRsAux!Cliente
           '
           ' End If
            conn.Execute SQL

    
            miRsAux.MoveNext
        Wend
        miRsAux.Close
            
            
        'JUNIO 2020
        'Si el cliente es FACE, entonces tenemos que si tiene dtoporcantidad, se pone a cero y el precio ar es importel/cantidad
        lblProgres(1).Caption = "FACE"
        lblProgres(1).Refresh
        Espera 0.25
        SQL = "select  tmpgasolimport.* FROM  tmpgasolimport  left join sclien on cliente=codclien "
        SQL = SQL & " WHERE codusu = " & vUsu.Codigo & "  And organogestor <> ''"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = "UPDATE tmpgasolimport SET descuento=0,precio=round((importel/cantidad),4)"
            SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND codigo = " & miRsAux!Codigo
            conn.Execute SQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        Espera 0.25
        SQL = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
        If Val(SQL) = 0 Then
            'NO HAY ERRORES. que valide formas de pago
      
            
            'Resumen formas de pago
            If False Then
                    SQL = "select idtipopago,numalbaran,importeconiva from tmpgasolimport   WHERE codusu = " & vUsu.Codigo
                    SQL = SQL & " group by 1,2 ORDER BY  1,2"
                    
                    
                    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    SQL = ""
                    Sql1 = ""
                    Tot = 0 'EN IVA redudido va el total
                    'tmpscapla(codusu,codplant,cantidad)
                    While Not miRsAux.EOF
                        If miRsAux!idtipopago <> SQL Then
                            If SQL <> "" Then
                                Sql1 = Sql1 & ", (" & vUsu.Codigo & "," & SQL & "," & DBSet(Impor1, "N") & ")"
                                Tot = Tot + Impor1
                            End If
                            SQL = miRsAux!idtipopago
                            Impor1 = 0
                        End If
                        Impor1 = Impor1 + miRsAux!importeConIva
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
            
            Else
                    SQL = "select idtipopago,sum(importeconiva) importeconiva from tmpgasolimport   WHERE codusu = " & vUsu.Codigo
                    SQL = SQL & " group by 1 ORDER BY  1"
                    
                    
                    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    SQL = ""
                    Sql1 = ""
                    Tot = 0 'EN IVA redudido va el total
                    'tmpscapla(codusu,codplant,cantidad)
                    While Not miRsAux.EOF
                        If miRsAux!idtipopago <> SQL Then
                            If SQL <> "" Then
                                Sql1 = Sql1 & ", (" & vUsu.Codigo & "," & SQL & "," & DBSet(Impor1, "N") & ")"
                                Tot = Tot + Impor1
                            End If
                            SQL = miRsAux!idtipopago
                            Impor1 = 0
                        End If
                        Impor1 = Impor1 + miRsAux!importeConIva
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
            End If
            If SQL <> "" Then
                Sql1 = Sql1 & ", (" & vUsu.Codigo & "," & SQL & "," & DBSet(Impor1, "N") & ")"
                Tot = Tot + Impor1   'EN IVA redudido va el total
            End If
            
            cadParam = Tot
            If Sql1 <> "" Then
                Sql1 = Mid(Sql1, 2)
                Sql1 = "INSERT INTO tmpscapla(codusu,codplant,cantidad) VALUES " & Sql1
                conn.Execute Sql1
            End If
        End If
        
            
            
        If Not EsAlvic2 Then
            lblProgres(1).Caption = "Comporbacion anulaciones"
            lblProgres(1).Refresh
            
            SQL = "select * from tmpgasolimport where codusu=" & vUsu.Codigo & "  and doc_relacionado<>''"
            SQL = SQL & " and not doc_relacionado  in (select doc_original from tmpgasolimport where codusu=" & vUsu.Codigo & ")"
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
            
                
                cad = "Error doc-vinculado. " & miRsAux!doc_relacionado
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & ",'" & miRsAux!Codigo & "'," & DBSet(Me.txtCodigo(0).Text, "F")
                SQL = SQL & ",23,59,-1,'Doc. vinc'," & _
                        DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(cad, "T") & ")"
                conn.Execute SQL
        
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
            
            
        Set miRsAux = Nothing
            
    End If
    
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = (I > 0)
    Exit Function

eProcesarFichero2:
    Sql1 = Err.Description
    MsgBox Sql1, vbExclamation
    Err.Clear
    conn.Errors.Clear
    ProcesarFichero2 = False
    Set R2 = Nothing
End Function

Private Function ComprobarRegistroLineaFichero(cad As String, ccoste As String) As Boolean
Dim SQL As String

Dim Base As String
Dim NombreBase As String
Dim turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim FechaHora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim surtidor As String
Dim manguera As String
Dim PrecioLitro As String
Dim cantidad As String
Dim Importe As String
Dim descuento As String
Dim idtipopago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_PrecioSinIVA As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim hora As String

Dim Mens As String
Dim Kilometros As String
Dim b As Boolean
Dim codsoc As String

Dim IvaArticulo As String
Dim NombreArticulo As String
Dim NomArtic As String
Dim CodIVA As String
Dim Porciva As Currency
Dim importeConIva  As Currency
Dim TIpoDeIva_D As Byte  '0. No establecido  1. Normal   2 REducido  3 Supe reducido    4 exento

Dim idClienteVarioAlvic As String
Dim Aux3 As Currency

' GESVEN.  Cuando un ticket lo pasan a factura, existe un ticket, una anulacion y una factura identificada
Dim DocumentoOriginal As String
Dim DocumentoRelacionado As String


Dim CampoNumeroFacturaAvalon As String 'le quito la primera letra y formateamos a 6 digitos

    On Error GoTo eComprobarRegistroAlz

    ComprobarRegistroLineaFichero = True

    DocumentoRelacionado = ""
    DocumentoOriginal = ""
    If EsAlvic2 Then
        'ALVIC

        If TipoFicheroNormal Then
        
        
            Base = Mid(cad, 1, 10)
            NombreBase = Mid(cad, 11, 50)
            turno = Trim(Mid(cad, 61, 10))
        
            NumAlbaran = Trim(Mid(cad, 71, 20))
            NumFactura = Trim(Mid(cad, 91, 20))
            IdVendedor = Trim(Mid(cad, 121, 10))
            NombreVendedor = Mid(cad, 131, 50)
            FechaHora = Trim(Mid(cad, 181, 14))
            Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
            hora = Mid(FechaHora, 9, 6)
            NombreCliente = Trim(Mid(cad, 215, 70))
            tarjeta = Trim(Mid(cad, 195, 20))
            Matricula = Trim(Mid(cad, 370, 20))
            IdProducto = Trim(Mid(cad, 493, 20))
            surtidor = Trim(Mid(cad, 538, 10))
            manguera = Trim(Mid(cad, 548, 10))
        
            PrecioLitro = Trim(Mid(cad, 568, 18))
            cantidad = Trim(Mid(cad, 650, 18))
            Importe = Trim(Mid(cad, 668, 18))
            descuento = Trim(Mid(cad, 586, 18))
            idtipopago = Trim(Mid(cad, 784, 10))
            DescrTipoPago = Trim(Mid(cad, 794, 25))
            CodigoTipoPago = Trim(Mid(cad, 1, 10))
            NifCliente = Trim(Mid(cad, 834, 9))
            
            IvaArticulo = Trim(Mid(cad, 609, 5))
            NombreArticulo = Trim(Mid(cad, 513, 25))
            Kilometros = Trim(Mid(cad, 415, 18))
            
            
        Else
        
        
            Vec = Split(cad, Chr(9))
           
            Base = Vec(0)
            NombreBase = Vec(1)
            turno = Vec(2)
        
            NumAlbaran = Trim(Vec(3))
            NumFactura = Trim(Vec(4))
            IdVendedor = Trim(Vec(6))
            NombreVendedor = Vec(7)
            FechaHora = Trim(Vec(8))
            Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
            hora = Mid(FechaHora, 9, 6)
            NombreCliente = Trim(Vec(10))
            tarjeta = Trim(Vec(9))
            Matricula = Trim(Vec(15))
            IdProducto = Trim(Vec(20))
            surtidor = Trim(Vec(22))
            manguera = Trim(Vec(23))
                
            PrecioLitro = Trim(Vec(25))
            
            
            cantidad = Trim(Vec(31))
            Importe = Trim(Vec(32))
            descuento = Trim(Vec(26))
           
            
            
            
            idtipopago = Trim(Vec(38))
            DescrTipoPago = Trim(Vec(39))
            CodigoTipoPago = Trim(Vec(40))
            NifCliente = Trim(Vec(41))
            
            'If idtipopago <> CodigoTipoPago Then Stop
            
            
            IvaArticulo = Trim(Vec(28))
            NombreArticulo = Trim(Vec(21))
            Kilometros = Trim(Vec(17))
            
            
            
            
        End If
        DocumentoOriginal = NumAlbaran
    Else
        'AVALON    AGosto 2020
            'Vec = Split(cad, ";")
            Vec = Split(cad, Chr(9))
            
            'Debug.Print UBound(Vec)
            If False Then
                    For NumRegElim = 0 To UBound(Vec) - 1
                        Debug.Print Vec(NumRegElim)
                        
                    Next
            End If
                        
            Base = Vec(0)
            NombreBase = Vec(0)
            '-----
            'turno
            SQL = Vec(2)
            
            turno = Val(SQL)   ' seeraá yymmddPC donde año mes dia campo P y C
                       
                        
            'Cliente
            tarjeta = Trim(Vec(13))
            If tarjeta = "UNKNOWN" Then tarjeta = sparamalvic!Clivario

            NifCliente = Trim(Vec(15))
                        
        
            
            
            'Factura / ALbaran
            DocumentoOriginal = Trim(Vec(6))
            
            SQL = LCase(Trim(Vec(3)))
            CampoNumeroFacturaAvalon = Right("000000" & Mid(Vec(5), 2), 6)
            If SQL = "factura simplificada" Then
                
                
                
                'TICKET
                
                
                'En alvic
                'D0031350
                If Vec(9) = "Venta" Then
                    DocumentoRelacionado = ""
                    'Ticket normal
                    SQL = "D"
                            

                Else
                    'DocumentoRelacionado
                    SQL = Mid(Vec(8), 1, 1)
                    SQL = SQL & Right("000000" & Mid(Vec(8), 2), 6)
                    DocumentoRelacionado = SQL
                
                
                    'ANULACION. Sera un albaran pero en negativo. Por si coincide
                    SQL = "D1"
                    
                End If
                SQL = SQL & CampoNumeroFacturaAvalon
                'Se hace una factura por cada ticket
                NumAlbaran = SQL
                NumFactura = ""
            Else
            
                Stop
                        
                'AVALON
                
                If Mid(LCase(SQL), 1, 5) = "albar" Then 'cuidado con formato fichero. Quito el acento pq pude ser que venga como caracter especial
                    'ALBARAN
                    CampoNumeroFacturaAvalon = Right("000000" & Mid(Vec(5), 2), 6)
                    SQL = "D" & Right(Vec(7), 6)
                    NumFactura = ""
                Else
                    
                    If SQL = "factura" Then
                        
     
                        NumAlbaran = Mid(DocumentoRelacionado, 1, 1) & Right("000000" & Mid(DocumentoRelacionado, 2), 6)
                        
                        
                        NumFactura = Mid(Vec(7), 1, 1) & Right("000000" & Mid(Vec(7), 2), 6)
                        
                    Else
                        Err.Raise 513, , "Tipo documento incorrecto: " & SQL
                    End If
                End If
                
               
            End If
           
            
            'Vendedor
            ' 001001_5651 BENITEZ
            IdVendedor = Trim(Vec(10))
            NombreVendedor = Vec(11)
  
                FechaHora = Trim(Vec(12))
                
            '    FechaHora = "20210213164926"
                
                
                 Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
                hora = Mid(FechaHora, 9, 8)
         
            NombreCliente = Trim(Vec(14))
           ' If NombreCliente <> "UNKNOWN" Then Stop
            
            
            Matricula = Trim(Vec(16))
            IdProducto = Trim(Vec(17))
            surtidor = ""
            manguera = ""
                
            PrecioLitro = Trim(Vec(17))   ''120500   =1.20500
            If PrecioLitro <> "" And PrecioLitro <> "0,00" Then PrecioLitro = CCur(PrecioLitro) * 100000
                
            
            cantidad = Trim(Vec(24)) * 100 'bajo lo trata divido /100
            Importe = Trim(Vec(25)) * 100
            descuento = Vec(21)
            If descuento <> "" And descuento <> "0" Then descuento = CCur(descuento) * 100000
                
 
            DescrTipoPago = Trim(Vec(28))
            idtipopago = Trim(Vec(27))
            idtipopago = 0
            
            IvaArticulo = CCur(Trim(Vec(23))) * "100"  'Son 4 decimales
            NombreArticulo = Trim(Vec(18))
            Kilometros = 0 '
                
        
    End If
        
        
        
        
    'Esto ers comun
    'CUIDADO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If vParamAplic.NumeroInstalacion = vbTaxco Then IdVendedor = Val(IdVendedor) + 500
        
        
    
    If Trim(Importe) = "" Then Importe = 0
    If CCur(Importe) = 0 Then Exit Function
    
    
    
    idClienteVarioAlvic = ""
    If Mid(tarjeta, 1, 2) = "1Z" Then
        'Cliente vario
        
        
        idClienteVarioAlvic = tarjeta
        tarjeta = sparamalvic!Clivario
    End If
    
    If FechaFichero < CDate("01/01/01") Then
        'Es la primera linea procesada
        FechaFichero = CDate(txtCodigo(0).Text)
    
        'Es la primera linea. La fecha debe coincidir con la del fichero
        If CDate(Fecha) <> FechaFichero Then
            Mens = "Fechas: " & Fecha & "  // " & FechaFichero
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                  "," & DBSet(Mid(hora, 4, 2), "N") & "," & DBSet(tarjeta, "N") & "," & DBSet(NifCliente, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
        End If
    End If
    
    If IdTurno > 0 Then
        If Val(turno) <> IdTurno Then
            Mens = "Err.turno:Fichero " & turno & "//" & IdTurno
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                  "," & DBSet(Mid(hora, 4, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(NifCliente, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    importeConIva = c_Importe
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 4)
    
    
    'If vContad = 285 Then Stop
    
    
    'Comprobamos que el IVA esta en alguno de los articulos de parametros
    
   
    
    Mens = ""
    TIpoDeIva_D = 0
    Porciva = Round2(CInt(ComprobarCero(IvaArticulo)) / 100, 0)
    If Porciva = IvaNormal Then
        TIpoDeIva_D = 1
    Else
        If Porciva = IvaReducido Then
            TIpoDeIva_D = 2
        Else
            If Porciva = IvaSuperReducido Then
                TIpoDeIva_D = 3
            Else
                If Porciva = 0 Then
                    TIpoDeIva_D = 4
                Else
                    Mens = "Porcentaje de iva no tratado: " & Porciva
                End If
            End If
        End If
    End If
    If Mens <> "" Then
        'Metemos en errores
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
        Porciva = IvaNormal 'para que no de error
    End If
    
    
   ' If IvaArticulo <> "2100" Then Stop
    If Trim(descuento) <> "" Then
        If CCur(descuento) <> 0 Then
            c_Descuento = Round2(CCur(descuento) / 100000, 5)
            If c_Descuento > 100 Then Err.Raise 513, , "Error  descuento: " & c_Descuento
            
             
        Else
            c_Descuento = 0
        End If
       
     
    End If
    
    
    c_PrecioSinIVA = 1 + (Porciva / 100)   'factor IVA
    If Not EsAlvic2 Then
        'EL importe IVA nos lo indican en el fichero
        c_PrecioSinIVA = Trim(Vec(25))
        c_Importe1 = importeConIva - c_PrecioSinIVA
    Else
        'Lo que habia
        c_Importe1 = (importeConIva / c_PrecioSinIVA)
    End If
    c_Importe2 = c_Importe1 / c_Cantidad
    If c_Descuento > 0 Then
        'EL DESCUENTO ES POR CANTIDAD !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!    Febrero 2020
        Aux3 = c_Descuento / (1 + (Porciva / 100))
        c_Importe2 = c_Importe2 + Aux3
        
    End If
    c_PrecioSinIVA = c_Importe2
    
  

    If Trim(NumFactura) <> "" Then
    
        If idClienteVarioAlvic <> "" Then
            'Es una factura A cliente varios identificado. Lo meteremos sclvar
            CodigoCliente = tarjeta
        
        Else
            'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
            'Y SI NO EXISTE ERROR
            Tarje = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "nifclien", NifCliente, "T")
            If Tarje = "" Then
                    
                   Mens = "No existe NIF en clientes"
                   SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                          "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(tarjeta, "N") & "," & DBSet(NifCliente, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                    
                   conn.Execute SQL
            End If
            CodigoCliente = Tarje
        End If
    Else
        'UN ALBARAN
        CodigoCliente = tarjeta
    End If
        
        
        
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        
        b = True  'ok por defecto
        If IdTurno > 0 Then
            'Esta traspasadno un turno. La fecha puede ser de la seleccionada, o un dia mas
            If CDate(Fecha) <> FechaFichero Then FechaFichero = DateAdd("d", 1, CDate(txtCodigo(0).Text)): Turno3 = True
            
        
        End If
            
        'If CDate(Fecha) <> CDate(txtcodigo(0).Text) Then b = False
        If CDate(Fecha) <> FechaFichero Then
            'No es la misma fecha.
            b = False
        End If

        If Not b Then
            Mens = "Fecha no es del traspaso" ' o no es del turno"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Val(Mid(hora, 3, 2)), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comporbamos el IVA esta en los tratados
    'IvaArticulo
    

    
    'Comprobamos que la forma de pago existe
    idtipopago = Trim(idtipopago)
    If Trim(idtipopago) = "" Then idtipopago = "VACIO"
        
    If Not IsNumeric(idtipopago) Then
        Mens = "Forma de pago incorrecta "
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(idtipopago, "T") & "," & _
                DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    
    Else
        idtipopago = "F" & idtipopago
    End If
    
    
    'Comprobamos que el codigo de trabajador existe
    'COMPROBAMOS QUE ES NUMERICO
    
    If Not IsNumeric(IdVendedor) Then
        
        Mens = "Codigo trabajador incorreto"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(vContad, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    
    Else
        IdVendedor = "T" & IdVendedor
    End If

    '------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------
    'INSERTAMOS EN TMP
    'cadSelect = "INSERT INTO tmpgasolimport(codusu,codigo,NumAlbaran,NumFactura,fechahora,IdVendedor,Cliente,NombreCliente,NifCliente,Matricula,CodigoProducto,surtidor,manguera,Precio,cantidad,descuento,idtipopago,importeConIva,doc_relacionado)"
    hora = Mid(hora, 1, 2) & ":" & Mid(hora, 3, 2) & ":" & Mid(hora, 5, 2)
    Mens = Format(Fecha, FormatoFecha) & " " & hora
    
    
    
    
   
    
    'codusu,codigo,NumAlbaran,NumFactura,fechahora,IdVendedor
    SQL = ", (" & vUsu.Codigo & "," & vContad & "," & DBSet(Trim(NumAlbaran), "T", "N") & "," & DBSet(Trim(NumFactura), "T", "S") & ",'" & Mens & "','" & IdVendedor
    
    'Cliente,NombreCliente,NifCliente,Matricula
    SQL = SQL & "'," & DBSet(CodigoCliente, "N", "N") & "," & DBSet(NombreCliente, "T", "N") & "," & DBSet(NifCliente, "T", "N") & "," & DBSet(Matricula, "T", "N")
    
    'CodigoProducto,surtidor,manguera
    SQL = SQL & "," & DBSet(NombreArticulo, "T", "N") & "," & DBSet(surtidor, "T", "N") & "," & DBSet(manguera, "T", "N")
    
    ',Precio,cantidad,descuento,importel,idtipopago ,ccoste)"
    SQL = SQL & "," & DBSet(c_PrecioSinIVA, "N", "N") & "," & DBSet(c_Cantidad, "N", "N") & "," & DBSet(c_Descuento, "N", "N")
    SQL = SQL & "," & DBSet(c_Importe1, "N", "N") & "," & DBSet(idtipopago, "T", "N") & "," & TIpoDeIva_D & "," & DBSet(importeConIva, "N")
    SQL = SQL & "," & DBSet(ccoste, "T", "N") & ",'" & turno & "','" & idClienteVarioAlvic & "',"
    
    'DocumentoOriginal DocumentoRelacionado
    SQL = SQL & DBSet(DocumentoOriginal, "T", "N") & "," & DBSet(DocumentoRelacionado, "T") & ")"
    
    'insertamos
    cadFormula = cadFormula & SQL
    
  
eComprobarRegistroAlz:
    If Err.Number <> 0 Then
        ComprobarRegistroLineaFichero = False
        Err.Raise 513, , Err.Description
    End If
End Function
            
            
     
            
            

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()

    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute "delete from tmpGasolImport where codusu = " & vUsu.Codigo
    conn.Execute "delete from tmpscapla where codusu = " & vUsu.Codigo
    conn.Execute "delete from tmpslipreu where codusu = " & vUsu.Codigo
    conn.Execute "delete from tmpimpresionauxliar where codusu = " & vUsu.Codigo
    
    
End Sub




 


Public Function CadenaClientesVarios() As String
    CadenaClientesVarios = "(cliente IN ( 100000,100001,100002,100003,100004,100005,100006,100008,100009,100010,100011) )"
End Function



Private Sub GeneraAsientoCobros()
Dim Mc As Contadores
Dim FechaAsi As Date
Dim SQL As String
Dim Importe As Currency

    On Error Resume Next

    'idforpaparametr
    cadTitulo = DevuelveDesdeBD(conAri, "Ctacierre", "sforpa", "codforpa", sparamalvic!ForPa)

Set Mc = New Contadores
    
    FechaAsi = CDate(txtCodigo(0).Text)
    Mc.ConseguirContador "0", FechaAsi <= vEmpresa.FechaFin, False
    cad = "Cierre caja ALVIC "
    If IdTurno > 0 Then cad = cad & "   turno: " & Format(IdTurno, "00000")
    
    'Cabecera del hco de apuntes
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari"
    SQL = SQL & ",feccreacion,usucreacion,desdeaplicacion"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "1" & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & "," & DBSet(cad, "T", "S")
    SQL = SQL & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARIGES'"
    ConnConta.Execute SQL & ")"
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    
        
    cad = ", (" & 1 & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Mc.Contador & ","

    Set miRsAux = New ADODB.Recordset
    Codigo = "select tmpgasolimport.*,codmacta from tmpgasolimport left join sclien on cliente=codclien"
    Codigo = Codigo & " WHERE  codusu=" & vUsu.Codigo & " AND tmpgasolimport.idtipopago<>2"   'CREDITO NO ENTRA
    Codigo = Codigo & "  ORDER BY sclien.codmacta,nombrecliente"
    
    
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Importe = 0
    Codigo = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Importe = Importe + miRsAux!importeConIva
        
        
        ' linliapu, codmacta, numdocum, "
        SQL = cad & NumRegElim & "," & DBSet(miRsAux!Codmacta, "T") & "," & DBSet(miRsAux!NumAlbaran, "T")
        'codconce,ampconce,
        SQL = SQL & ",1," & DBSet(miRsAux!NombreCliente, "T") & ","
        ' timporteD, timporteH,
        SQL = SQL & "NULL," & DBSet(miRsAux!importeConIva, "N")
        'codccost, ctacontr, idcontab, punteada
        SQL = SQL & ",NULL," & DBSet(cadTitulo, "T") & ",'contab',0)"
        Codigo = Codigo & SQL
        
        miRsAux.MoveNext
        If miRsAux.EOF Then
            indCodigo = 10001
        Else
            indCodigo = Len(Codigo)
        End If
        If indCodigo > 10000 Then
            Codigo = Mid(Codigo, 2)
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
            SQL = SQL & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada) VALUES "
            SQL = SQL & Codigo
            Codigo = ""
            ConnConta.Execute SQL
            If Err.Number <> 0 Then
                MsgBox "Creando asiento: " & Err.Description, vbExclamation
                Err.Clear
            End If
        End If
    Wend
    miRsAux.Close
    
    
    'Cerramos el importe
    If False Then
        'Esto YA NO LO HACEMOS. Borar en un futuro
        NumRegElim = NumRegElim + 1
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
        SQL = SQL & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada) VALUES "
        SQL = SQL & Mid(cad, 2) & NumRegElim & "," & DBSet(cadTitulo, "T") & "," & DBSet("cierre turno", "T")
        'codconce,ampconce,
        SQL = SQL & ",1," & DBSet("Cierre " & txtCodigo(0).Text, "T") & ","
        ' timporteD, timporteH,
        SQL = SQL & DBSet(Importe, "N") & ",NULL"
        'codccost, ctacontr, idcontab, punteada
        SQL = SQL & ",NULL,NULL,'contab',0)"
        ConnConta.Execute SQL
    
    
        'Distribuimos los importes entre las forpas de pago del fichero segun lo que viene en fich
        NumRegElim = NumRegElim + 1
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
        SQL = SQL & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada) VALUES "
        SQL = SQL & Mid(cad, 2) & NumRegElim & "," & DBSet(cadTitulo, "T") & "," & DBSet("cierre turno", "T")
        'codconce,ampconce,
        SQL = SQL & ",1," & DBSet("Cierre " & txtCodigo(0).Text, "T") & ","
        ' timporteD, timporteH,
        SQL = SQL & "NULL," & DBSet(Importe, "N")
        'codccost, ctacontr, idcontab, punteada
        SQL = SQL & ",NULL,NULL,'contab',0)"
        ConnConta.Execute SQL
     
   End If
         
         
         
    'GEneramos pod forma de pago del traspaso
    Codigo = "select * from tmpscapla   ,sforpa where codforpa=codplant and codusu=" & vUsu.Codigo
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Codigo = ""

    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Importe = Importe - miRsAux!cantidad
        
        
        ' linliapu, codmacta, numdocum, "
        SQL = ""
        If IdTurno > 0 Then
            SQL = "Cierre " & Format(IdTurno, "00000")
        Else
            SQL = "Cierre turno"
        End If
        SQL = cad & NumRegElim & "," & DBSet(miRsAux!Ctacierre, "T") & "," & DBSet(SQL, "T")
        
        
        
        'codconce,ampconce,
        SQL = SQL & ",1," & DBSet("Cierre turno " & txtCodigo(1).Text, "T") & ","
        ' timporteD, timporteH,
        SQL = SQL & DBSet(miRsAux!cantidad, "N") & ",NULL"
        'codccost, ctacontr, idcontab, punteada
        SQL = SQL & ",NULL," & DBSet(cadTitulo, "T") & ",'contab',0)"
        Codigo = Codigo & SQL
        
        miRsAux.MoveNext
        If miRsAux.EOF Then
            indCodigo = 1
        Else
            indCodigo = 0
        End If
        If indCodigo > 0 Then
            Codigo = Mid(Codigo, 2)
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
            SQL = SQL & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada) VALUES "
            SQL = SQL & Codigo
            Codigo = ""
            ConnConta.Execute SQL
        End If
    Wend
    miRsAux.Close
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub



Private Sub GenerarFacturasScafac()
Dim RT As ADODB.Recordset
Dim T1 As Single


    Set RT = New ADODB.Recordset
    Codigo = "select count(*) from tmpslipreu where codusu =" & vUsu.Codigo & " ORDER BY nomartic,codartic"
    RT.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = "0"
    If Not RT.EOF Then Codigo = RT.Fields(0)
    RT.Close
    
    If Val(Codigo) = 0 Then
        MsgBox "ninguna factura a generar", vbExclamation
        Exit Sub
    End If
    Pb1.Value = 0
    Pb1.Max = CInt(Codigo)
    
    
'    Dim TipCod As String
'Dim cad As String
'Dim cadTabla As String
Dim Fecha As Date
        
    Codigo = "select * from tmpslipreu where codusu =" & vUsu.Codigo & " ORDER BY nomartic,codartic"
    RT.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    
    While Not RT.EOF
        cad = RT!NomArtic & RT!codArtic
        lblProgres(0).Caption = cad
        lblProgres(0).Refresh
        Pb1.Value = Pb1.Value + 1
        Screen.MousePointer = vbHourglass
        
        
        If cad <> Codigo Then
            If Codigo <> "" Then GeneraLaFactura Fecha
            'Busco la fecha
            cadTabla = RT!NomArtic 'SERIE de la factura
            If cadTabla = sparamalvic!FraDirectaD Or cadTabla = sparamalvic!FacturaVariosD Then
                'FraDirectaD FacturaVariosD
                cadTabla = sparamalvic!letraGasoleo
            ElseIf cadTabla = sparamalvic!FraDirectaT Or cadTabla = sparamalvic!FacturaVariosT Then
                'FraDirectaT FacturaVariosT
                cadTabla = sparamalvic!letraTienda
            Else
                'FacturaVariosA FacturaVariosA
                cadTabla = sparamalvic!letraVarios
             End If
             
             If RT!NumOfert = 0 Then
                'VARIOS
                cadTabla = "'" & cadTabla & "%' AND numfactura is null"
                cadTabla = "numalbaran like " & cadTabla & " AND codusu "
                cadTabla = DevuelveDesdeBD(conAri, "fechahora", "tmpgasolimport", cadTabla, CStr(vUsu.Codigo))
            Else
                 cadTabla = cadTabla & Format(RT!NumOfert, "0000000")
                cadTabla = "numalbaran = '" & cadTabla & "' AND codusu "
                cadTabla = DevuelveDesdeBD(conAri, "fechahora", "tmpgasolimport", cadTabla, CStr(vUsu.Codigo))
            End If
            If cadTabla = "" Then cadTabla = Me.txtCodigo(0).Text
             Fecha = Format(cadTabla, "dd/mm/yyyy")
                
            Codigo = RT!NomArtic & RT!codArtic
            cadTabla = ""
            cadParam = RT!Ampliaci
            cadTitulo = RT!NomArtic
            
            
            If (Pb1.Value Mod 12) = 0 Then
                Espera 0.5
                DoEvents
            End If
        End If
        If RT!NumOfert > 0 Then cadTabla = cadTabla & ", " & RT!NumOfert
         
        RT.MoveNext
    Wend
    RT.Close
    
    If Codigo <> "" Then GeneraLaFactura Fecha
     
    Set RT = Nothing
End Sub


Private Sub GeneraLaFactura(FE As Date)
Dim Aux As String
Dim C As String
Dim Resumen As Boolean
    
    Aux = Mid(Codigo, 1, 3)
    ' FAW   FAX  FAY
    Resumen = False
    
    If Aux = "FAW" Or Aux = "FAX" Or Aux = "FAY" Then Resumen = True
    Aux = Mid(Codigo, 4)
    
    If Not Turno3 Then Resumen = False
    
    If Resumen Then FE = DateAdd("d", 1, CDate(txtCodigo(0).Text))
    
    
    
    cadTabla = Mid(cadTabla, 2)
    If cadTabla <> "" Then cadTabla = " AND numalbar in (" & cadTabla & ")"
    cadTabla = "codtipom = '" & cadParam & "'" & cadTabla
    TraspasoFacturasGasol cadTitulo, cadTabla, Format(FE, "dd/mm/yyyy"), "", Nothing, Me.lblProgres(1), False, cadTitulo, Aux, 0, False
    DoEvents
    Espera 0.1
    
End Sub
