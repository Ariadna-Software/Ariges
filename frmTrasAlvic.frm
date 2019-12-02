VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasAlvic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasAlvic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
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
            Width           =   1080
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2730
            MaxLength       =   1
            TabIndex        =   1
            Top             =   870
            Width           =   330
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
            Caption         =   "Nº Turno"
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
            Top             =   900
            Width           =   645
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
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
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
Dim Cad As String
Dim cadTabla As String

Dim vContad As Long

Dim primeravez As Boolean

Dim ArtFamGenerica As String

Dim IvaNormal As Currency
Dim IvaReducido As Currency
Dim IvaSuperReducido As Currency



Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, 0, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim SQL As String
Dim I As Byte
Dim cadWhere As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String

On Error GoTo eError


    If Not DatosOk Then Exit Sub
    
    
    Me.CommonDialog1.DefaultExt = "TXT"
    CADENA = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        InicializarTabla

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
                    
                    conn.BeginTrans
                    b = ProcesarFichero(Me.CommonDialog1.FileName)
                    
            

                End If
          End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not b Then
        If Err.Number = 32755 Then Exit Sub
        
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        'FALTA###
        'BorrarArchivo Me.CommonDialog1.FileName
        'BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totaliza")
        'BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "compras")
        
        'If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
        ' solo en el caso de alzira se graba en la srecau
        If False Then
         '   BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "caja")
          '  BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totales")
        End If
        cmdCancel_Click
    End If
    
End Sub

    




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Cmdleer_Click()

End Sub




Private Sub Form_activate()
    If primeravez Then
        primeravez = False
        PonerFoco txtcodigo(0)
        
        
        
        
        'Vamos a ver los ivas, desde la conta
        cadSelect = ""
        For indCodigo = 1 To 4
 
            cadFormula = "artvario"
            If indCodigo = 2 Then
                cadNombreRPT = vParamAplic.GasolArticuloReducido
            ElseIf indCodigo = 3 Then
                cadNombreRPT = vParamAplic.GasolArticuloSuperReducido
            ElseIf indCodigo = 4 Then
                cadNombreRPT = vParamAplic.GasolArticuloExento
            Else
                'indCodigo = 1
                cadNombreRPT = vParamAplic.GasolArticuloNormal
            End If
            cadParam = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", cadNombreRPT, "T", cadFormula)
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
                        'OK. Todo bien. Veamos porcentaje
                        If indCodigo = 2 Then
                            'cadNombreRPT = vParamAplic.GasolArticuloReducido
                            IvaReducido = CCur(cadParam)
                        ElseIf indCodigo = 3 Then
                            'cadNombreRPT = vParamAplic.GasolArticuloSuperReducido
                            IvaSuperReducido = CCur(cadParam)
                        ElseIf indCodigo = 4 Then
                            'cadNombreRPT = vParamAplic.GasolArticuloExento
                            
                        Else
                            'indCodigo = 1
                            IvaNormal = CCur(cadParam)
                        End If
                    End If
                        
                End If
            End If
            
        Next
        
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

    primeravez = True
    limpiar Me

    
    txtcodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")
     
    FrameCobrosVisible True, H, W
    Pb1.visible = False
        
    
    
    Me.cmdCancel.Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgAyuda_Click(index As Integer)
Dim vCadena As String
    Select Case index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si se ha eliminado un turno, el check ha de estar desmarcado. " & vbCrLf & vbCrLf & _
                      "El motivo es porque si se ha traspasado el fichero de compras, " & vbCrLf & _
                      "los albaranes no se eliminan cuando se borra un turno." & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgFec_Click(index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(index).Left
    dalt = imgFec(index).Top

    Set obj = imgFec(index).Container

    While imgFec(index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(index).Parent.Left + 30
    frmC.Top = dalt + imgFec(index).Parent.Top + imgFec(index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(index).Text <> "" Then frmC.Fecha = txtcodigo(index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub txtCodigo_GotFocus(index As Integer)
    ConseguirFoco txtcodigo(index), 3
End Sub

Private Sub txtCodigo_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(index As Integer, KeyAscii As Integer)
    
    'If KeyAscii = teclaBuscar Then
    If Chr(KeyAscii) = "+" Then
        Select Case index
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

Private Sub txtCodigo_LostFocus(index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(index).Text = Trim(txtcodigo(index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case index
        Case 0 'FECHAS
            If txtcodigo(index).Text <> "" Then PonerFormatoFecha txtcodigo(index)
            
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
Dim SQL As String
   b = True

   If txtcodigo(0).Text = "" And b Then
        MsgBox "El campo fecha debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(0)
    End If
    
    If txtcodigo(1).Text = "" And b Then
        MsgBox "El número de Turno debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(1)
    End If
 
 
    If b Then
 
        'If vParamAplic.Cooperativa = 5 Then
        If False Then
            SQL = "SELECT count(*) FROM scaalb WHERE fecalbar = " & DBSet(txtcodigo(0).Text, "F")
            
            If txtcodigo(1).Text <> "" Then SQL = SQL & " AND codturno = " & DBSet(txtcodigo(1).Text, "N")
            
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                b = False
                PonerFoco txtcodigo(1)
            End If
        Else
            ' faltaba comprobar que en el regaixo que no llevan turnos no se haya hecho ya el traspaso
            'If vParamAplic.Cooperativa = 2 Then
            If False Then
                SQL = "SELECT count(*) FROM srecau WHERE fechatur = " & DBSet(txtcodigo(0).Text, "F")
                If TotalRegistros(SQL) <> 0 Then
                    MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(1)
                End If
            Else
                SQL = "SELECT count(*) FROM srecau WHERE fechatur = " & DBSet(txtcodigo(0).Text, "F") & _
                      " AND codturno = " & DBSet(txtcodigo(1).Text, "N")
                If TotalRegistros(SQL) <> 0 Then
                    MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(1)
                End If
            End If
        End If
    
    End If
 
    DatosOk = b
End Function



Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.Path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, Cad
    Close #NF
    If Cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim I As Integer
Dim Longitud As Long
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    Longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = Longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    While Not EOF(NF)
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        
        b = InsertarLineaAlz(Cad)
        
        
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        b = InsertarLineaAlz(Cad)
      
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim I As Integer
Dim Longitud As Long
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    Longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = Longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT
    NumRegElim = 0
    Do
        
        
        If EOF(NF) Then
            b = False
    
        Else
            I = I + 1
            Line Input #NF, Cad
            Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
            lblProgres(1).Caption = "Linea " & I
        
        
            b = ComprobarRegistroAlz(Cad)
            If Not b Then I = 0
        End If
    Loop Until Not b
    Close #NF
    
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = (I > 0)
    Exit Function

eProcesarFichero2:
    
    ProcesarFichero2 = False
End Function
                
Private Function InsertarCabecera(Cad As String) As Boolean
Dim Numfactu As String
Dim TipDocu As String
Dim FechaCa As String
Dim turno As String
Dim hora As String
Dim ForPa As String
Dim Tarje As String
Dim Tarje1 As String
Dim Matric As String
Dim NomCli As String
Dim NifCli As String
Dim Ticket As String
Dim CtaConta As String ' cuenta contable de clientes contado
Dim codsoc As String
Dim SQL As String

    On Error GoTo eInsertarCabecera

    InsertarCabecera = False

    Numfactu = 0
    TipDocu = Mid(Cad, 10, 1)
    FechaCa = Mid(Cad, 11, 2) & Mid(Cad, 13, 2) & "20" & Mid(Cad, 15, 2)
    turno = Mid(Cad, 17, 1)
    hora = Mid(Cad, 18, 2) & ":" & Mid(Cad, 21, 2) & ":00"
    ForPa = Mid(Cad, 49, 2)
    Tarje = Mid(Cad, 53, 7)
    Tarje1 = Mid(Cad, 60, 5)
    Matric = Mid(Cad, 65, 10)
    NomCli = Mid(Cad, 91, 25)
    NifCli = Mid(Cad, 116, 9)
            
    '06/03/2007 añadida estas 2 lineas que faltaba
    If CInt(ForPa) <> 2 And Trim(Tarje) <> Trim(Tarje1) Then Tarje = Tarje1
    If Tarje = "" Then Tarje = "0"
    
    Select Case TipDocu
        Case "O"
            Ticket = Mid(Cad, 2, 8)
        Case "T"
            Ticket = Mid(Cad, 23, 8)
        Case "A"
            Ticket = Mid(Cad, 31, 8)
        Case "F"
            Ticket = Mid(Cad, 2, 8)
            Numfactu = Mid(Cad, 39, 8)
        
            'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
            'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
            Tarje = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCli, "T")
            If Tarje = "" Then
                Tarje = 900000
                Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
                
                CtaConta = ""
                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "0", "N")
                
                SQL = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                      "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                      "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                      "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                      DBSet(Tarje, "N") & ",0," & DBSet(NomCli, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                      "'VALENCIA'," & DBSet(NifCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                      DBSet(txtcodigo(0).Text, "F") & "," & _
                      ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                      "0,0,0,0,0," & ValorNulo & "," & ValorNulo & ")"
                      
                conn.Execute SQL
                      
                SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                      "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(NomCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                      ValorNulo & "," & ValorNulo & ",0)"
                
                conn.Execute SQL
            End If
    End Select
   

    'MIRAMOS SI EXISTE LA TARJETA
    codsoc = ""
    codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarje, "T")
    If Tarje = "       " Then Tarje = "0000000"
    If codsoc = "" Then
    
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Ticket, "N") & ",'" & Mid(FechaCa, 5, 4) & Mid(FechaCa, 3, 2) & Mid(FechaCa, 1, 2) & "'," & DBSet(Format(hora, "hh"), "N") & _
              "," & DBSet(Format(hora, "mm"), "N") & "," & DBSet(Tarje, "N") & ",'Nro. Tarjeta no existe') "
              
        conn.Execute SQL
        
        
    Else
        SQL = "update scaalb set codsocio = " & DBSet(codsoc, "N") & ", numtarje = " & DBSet(Tarje, "N") & ", numalbar = " & _
               DBSet(Ticket, "T") & ", horalbar = " & DBSet(txtcodigo(0).Text & " " & hora, "FH") & ", matricul = " & DBSet(Matric, "T") & _
               ", codforpa = " & DBSet(ForPa, "N") & ", numfactu = " & DBSet(Numfactu, "N") & _
               " where fecalbar = " & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N") & _
               " and numalbar = " & DBSet(vContad, "T")
               
        conn.Execute SQL
    End If
    
    vContad = vContad + 1

    InsertarCabecera = True
    
eInsertarCabecera:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Cabecera " & Err.Description, vbExclamation
    End If

End Function
            
Private Function ComprobarRegistro(Cad As String) As Boolean
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
Dim PrecioSinDto As String
Dim cantidad As String
Dim Importe As String
Dim idtipopago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency

Dim Fecha As String
Dim hora As String

Dim Mens As String
Dim Kilometros As String


Dim codsoc As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    Base = Mid(Cad, 1, 10)
    NombreBase = Mid(Cad, 11, 50)
    turno = Mid(Cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(turno) > 9 Then turno = "9"
    
    NumAlbaran = Mid(Cad, 72, 19)
    NumFactura = Mid(Cad, 94, 17) 'antes 91,20
    IdVendedor = Mid(Cad, 121, 10)
    NombreVendedor = Mid(Cad, 131, 50)
    FechaHora = Mid(Cad, 181, 14)
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 6)
    CodigoCliente = Mid(Cad, 195, 20)
    NombreCliente = Mid(Cad, 215, 70)
    tarjeta = Mid(Cad, 290, 20)
    Matricula = Mid(Cad, 370, 20)
    IdProducto = Mid(Cad, 493, 20)
    surtidor = Mid(Cad, 538, 10)
    manguera = Mid(Cad, 548, 10)
    
    
    '[Monica]24/08/2015: el precio es sin el descuento en la linea 864, antes ponia 568
    PrecioLitro = Mid(Cad, 864, 18)
    
    cantidad = Mid(Cad, 650, 18)
    Importe = Mid(Cad, 668, 18)
    idtipopago = Mid(Cad, 784, 10)
    DescrTipoPago = Mid(Cad, 794, 25)
    CodigoTipoPago = Mid(Cad, 1, 10)
    NifCliente = Mid(Cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(Cad, 415, 18)
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(NumFactura) <> "" Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            SQL = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES ("
            'Sql = Sql & DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', "
            'Sql = Sql & "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtcodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            conn.Execute SQL
                  
            SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            conn.Execute SQL
        End If
    End If
    
    'MIRAMOS SI EXISTE LA TARJETA
    '[Monica]17/06/2013: añadida la condicion de que la tarjeta no venga con asteriscos: instr(1, Tarjeta, "*") = 0
    If Mid(tarjeta, 1, 4) <> "****" And Trim(tarjeta) <> "" And InStr(1, tarjeta, "*") = 0 Then
        '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
        If Len(Trim(tarjeta)) = 16 Then
            tarjeta = Mid(tarjeta, 9, 16)
        End If
        '++
        codsoc = ""
        codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", tarjeta, "T")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                  "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
            
        End If
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        If CDate(Fecha) <> CDate(txtcodigo(0).Text) Or CByte(turno) <> CByte(txtcodigo(1).Text) Then
            Mens = "Fecha incorrecta" ' o no es del turno"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "codartic", "codartic", IdProducto, "N")
    If SQL = "" Then
        Mens = "No existe el artículo"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    
    'Comprobamos que el socio existe
    If CodigoCliente <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
        If SQL = "" Then
            Mens = "No existe el cliente"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que la forma de pago existe
    If idtipopago <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "forpaalvic", idtipopago, "N")
        If SQL = "" Then
            Mens = "No existe la forma de pago Alvic"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(idtipopago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If SQL = "" Then
        Mens = "No existe el trabajador"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function

Private Function ComprobarRegistroAlz(Cad As String) As Boolean
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
Dim c_Precio2 As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim hora As String

Dim Mens As String
Dim Kilometros As String

Dim codsoc As String

Dim IvaArticulo As String
Dim NombreArticulo As String
Dim NomArtic As String
Dim CodIVA As String
Dim Porciva As Currency

    On Error GoTo eComprobarRegistroAlz

    ComprobarRegistroAlz = True

    Base = Mid(Cad, 1, 10)
    NombreBase = Mid(Cad, 11, 50)
    turno = Mid(Cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(turno) > 9 Then turno = "9"
    
    NumAlbaran = Mid(Cad, 71, 20)
    NumFactura = Mid(Cad, 94, 17) 'antes 91,20
    IdVendedor = Mid(Cad, 121, 10)
    NombreVendedor = Mid(Cad, 131, 50)
    FechaHora = Mid(Cad, 181, 14)
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 6)
    NombreCliente = Mid(Cad, 215, 70)
    tarjeta = Mid(Cad, 195, 20)
    Matricula = Mid(Cad, 370, 20)
    IdProducto = Mid(Cad, 493, 20)
    surtidor = Mid(Cad, 538, 10)
    manguera = Mid(Cad, 548, 10)

    PrecioLitro = Mid(Cad, 568, 18)
    

    cantidad = Mid(Cad, 650, 18)
    Importe = Mid(Cad, 668, 18)
    
    
    
    
    descuento = Mid(Cad, 586, 18)
    idtipopago = Mid(Cad, 784, 10)
    DescrTipoPago = Mid(Cad, 794, 25)
    CodigoTipoPago = Mid(Cad, 1, 10)
    NifCliente = Mid(Cad, 834, 9)
    
    
    IvaArticulo = Mid(Cad, 609, 5)
    NombreArticulo = Mid(Cad, 513, 25)
    
    Kilometros = Mid(Cad, 415, 18)
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(descuento) <> "" Then
        If CCur(descuento) <> 0 Then
            c_Descuento = Round2(CCur(descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
        Else
            c_Descuento = 0
        End If
    End If
    


    If Trim(NumFactura) <> "" Then
    
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE ERROR
        Tarje = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "nifclien", NifCliente, "T")
        If Tarje = "" Then
               Mens = "No existe NIF en clientes"
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                      "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(NifCliente, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                
                conn.Execute SQL
            
            
        
        
        End If
    End If
        
        
        
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        If CDate(Fecha) <> CDate(txtcodigo(0).Text) Or CByte(turno) <> CByte(txtcodigo(1).Text) Then
            Mens = "Fecha incorrecta" ' o no es del turno"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comporbamos el IVA esta en los tratados
    'IvaArticulo
    
    'Comprobamos que el IVA esta en alguno de los articulos de parametros
    Mens = ""
    
    Porciva = Round2(CInt(ComprobarCero(IvaArticulo)) / 100, 0)
    If Porciva <> IvaNormal Then
        If Porciva <> IvaReducido Then
            If Porciva <> IvaSuperReducido Then Mens = "Porcentaje de iva no tratado: " & Porciva
        End If
    End If
    If Mens <> "" Then
        'Metemos en errores
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If

    
    'Comprobamos que la forma de pago existe
    If idtipopago <> "" Then
    
        
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "idForpaT", idtipopago, "N")
        If SQL = "" Then
            Mens = "No existe la forma de pago Alvic"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(idtipopago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If SQL = "" Then
        Mens = "No existe el trabajador"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    'Comprobamos si hay descuento que el codigo de articulo de dto existe
    
eComprobarRegistroAlz:
    If Err.Number <> 0 Then
        ComprobarRegistroAlz = False
    End If
End Function
            
            
Private Function ComprobarRegistroRib(Cad As String) As Boolean
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
Dim c_Precio2 As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim hora As String

Dim Mens As String
Dim Kilometros As String

Dim codsoc As String
Dim NuevaCuenta As String


    On Error GoTo eComprobarRegistroRib

    ComprobarRegistroRib = True

    Base = Mid(Cad, 1, 10)
    NombreBase = Mid(Cad, 11, 50)
    turno = Mid(Cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(turno) > 9 Then turno = "9"
    
    NumAlbaran = Mid(Cad, 71, 20)
    NumFactura = Mid(Cad, 92, 7) 'antes 91,20
    IdVendedor = Mid(Cad, 121, 10)
    NombreVendedor = Mid(Cad, 131, 50)
    FechaHora = Mid(Cad, 181, 14)
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 6)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(Cad, 215, 70)
    tarjeta = Mid(Cad, 195, 20)
    Matricula = Mid(Cad, 370, 20)
    IdProducto = Mid(Cad, 493, 20)
    surtidor = Mid(Cad, 538, 10)
    manguera = Mid(Cad, 548, 10)
    PrecioLitro = Mid(Cad, 568, 18)
    cantidad = Mid(Cad, 650, 18)
    Importe = Mid(Cad, 668, 18)
    descuento = Mid(Cad, 586, 18)
    idtipopago = Mid(Cad, 784, 10)
    DescrTipoPago = Mid(Cad, 794, 25)
    CodigoTipoPago = Mid(Cad, 1, 10)
    NifCliente = Mid(Cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(Cad, 415, 18)
    
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(descuento) <> "" Then
        If CCur(descuento) <> 0 Then
            c_Descuento = Round2(CCur(descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
        Else
            c_Descuento = 0
        End If
    End If
    

    If Trim(NumFactura) <> "" And InStr(1, tarjeta, "Z") <> 0 Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            '[Monica]02/01/2019: ahora los clientes de paso tienen que estar entre 8001 y 9998 antes entre 900000 y 999998
            'Tarje = 8000 '900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 8001 and codsocio <= 9998")
            If TotalRegistros("select codsocio from ssocio where codsocio >= 8001 and codsocio <= 9998") = 0 Then Tarje = "8001"
            If Tarje = "1" Then Tarje = ""
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            '[Monica]03/01/2019: en caso de que no se puedan crear mas clientes de paso damos un error
            If Tarje = "" Then
                MsgBox "No podemos crear socio de paso. Llame a Ariadna.", vbExclamation
                ComprobarRegistroRib = False
                Exit Function
            End If
            
            
            NuevaCuenta = "43." & Tarje
            
            'Rellenamos si procede
            NuevaCuenta = RellenaCodigoCuenta(NuevaCuenta)
            
            
            SQL = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES ("
                  
                  'DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtcodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "1,0,0,0,0," & DBSet(NuevaCuenta, "T") & "," & ValorNulo & ")" ' antes vParamAplic.CtaContable
                  
            conn.Execute SQL
                  
            SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            conn.Execute SQL
            
            '[Monica]03/01/2018: introducimos tambien la cuenta en la contabilidad
            SQL = "insert ignore into cuentas (codmacta,nommacta,apudirec,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) values ("
            SQL = SQL & DBSet(NuevaCuenta, "T") & "," & DBSet(NombreCliente, "T") & ",'S'," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA',"
            SQL = SQL & "46,'VALENCIA','VALENCIA'," & DBSet(NifCliente, "T") & ")"
            
            ConnConta.Execute SQL
            
        Else
            '[Monica]07/02/2011: caso de que sea un socio que quiere la factura (me viene en fichero nro de factura y Z)
            ' añadida esta parte del else que no estaba
            '[Monica]03/01/2019: ahora entre 8001 y 9998
            If CLng(Tarje) >= 8001 And CLng(Tarje) <= 9998 Then ' 900000 Then
                ' miro si existe tarjeta sino la creo
                SQL = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                If TotalRegistros(SQL) = 0 Then
                    SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                          "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                          ValorNulo & "," & ValorNulo & ",0)"
                    
                    conn.Execute SQL
                End If
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                Stop 'FALTATipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
                    SQL = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 1"
                    If TotalRegistros(SQL) = 0 Then
            
                        SQL = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                        If TotalRegistros(SQL) = 0 Then
                            SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                                  ValorNulo & "," & ValorNulo & ",0)"

                            conn.Execute SQL
                        End If

                    End If
                Else
                    SQL = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 0"
                    If TotalRegistros(SQL) = 0 Then
                        Mens = "Nro. Tarjeta no existe"
                        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                              "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                        
                        conn.Execute SQL
                    End If
                End If
            End If '07/02/2011: hasta aqui la parte añadida
        
        End If
    Else
        'MIRAMOS SI EXISTE LA TARJETA
        ' en alzira lo pongo dentro
        codsoc = ""
        '++monica:050508 el numero de tarjeta puede venir a blanco--> dar error
        If Trim(tarjeta) <> "" Then codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", tarjeta, "N")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                  "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
        Else
            ' comprobamos que el socio existe
            ' no haria falta pq hay clave referencial a ssocio
            SQL = ""
            SQL = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "codsocio", codsoc, "N")
            If SQL = "" Then
                Mens = "No existe el cliente"
                SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                      "importe4, importe5, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
                SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "T") & "," & _
                        DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                
                conn.Execute SQL
            End If
        End If
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        '[Monica]09/01/2013: en Ribarroja meten todos los turnos del dia a diferencia de Alzira
        If CDate(Fecha) <> CDate(txtcodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que el articulo existe en sartic
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "codartic", "codartic", IdProducto, "N")
    If SQL = "" Then
        Mens = "No existe el artículo"
        Dim IdProducto1 As Currency
        IdProducto1 = CCur(IdProducto)
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto1, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    'Comprobamos que la forma de pago existe
    If idtipopago <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "forpaalvic", idtipopago, "N")
        If SQL = "" Then
            Mens = "No existe la forma de pago Alvic"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(idtipopago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If SQL = "" Then
        Mens = "No existe el trabajador"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    'Comprobamos si hay descuento que el codigo de articulo de dto existe
    If c_Descuento <> 0 Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "artdto", "codartic", IdProducto, "N")
        If SQL = "" Then
            Mens = "No tiene artículo de descuento"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
        End If
    End If
    
eComprobarRegistroRib:
    If Err.Number <> 0 Then
        ComprobarRegistroRib = False
    End If
End Function
            
Private Function ComprobarRegistroReg(ByRef RS As Recordset) As Boolean
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
Dim PrecioSinDto As String
Dim cantidad As String
Dim Importe As String
Dim idtipopago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency

Dim Fecha As String
Dim hora As String

Dim Mens As String
Dim Kilometros As String


Dim codsoc As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistroReg = True

    turno = DBLet(RS!turno, "N")
    
    NumAlbaran = DBLet(RS!Albaran, "N")
    NumFactura = DBLet(RS!Factura, "T")
    If NumFactura <> "" Then
'        NumFactura = Mid(NumFactura, 5, Len(NumFactura) - 4)
        If Mid(NumFactura, 1, 3) = "FAV" Then
            NumFactura = "9" & Mid(NumFactura, Len(NumFactura) - 5, 6)
        Else
            NumFactura = Mid(NumFactura, Len(NumFactura) - 6, 7)
        End If
    End If
    FechaHora = DBLet(RS!Fecha, "T")
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 6)
    CodigoCliente = DBLet(RS!Cliente, "T")
    NombreCliente = DBLet(RS!NomClien, "T")
    tarjeta = DBLet(RS!tarjeta, "N")
    Matricula = DBLet(RS!Matricula, "T")
    IdProducto = DBLet(RS!producto, "N")
    surtidor = DBLet(RS!surtidor, "N")
    manguera = DBLet(RS!manguera, "N")
    
    
    PrecioLitro = DBLet(RS!Precio, "N")
    
    cantidad = DBLet(RS!cantidad, "N")
    Importe = DBLet(RS!Importe, "N")
    idtipopago = DBLet(RS!idtipopago, "N")
    DescrTipoPago = DBLet(RS!desctipopago, "T")
    CodigoTipoPago = DBLet(RS!idtipopago, "N")
    NifCliente = DBLet(RS!NIF, "T")
    
    Kilometros = DBLet(RS!km, "N")
    
    ' en caso de que el codigo de cliente y el nombre no me vengan cojo el asociado a la forma de pago
    If CodigoCliente = "" And NombreCliente = "" Then
        CodigoCliente = DevuelveDesdeBDNew(conAri, "sforpa", "codsocio", "forpaalvic", idtipopago, "N")
        NombreCliente = DevuelveDesdeBDNew(conAri, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
        tarjeta = CodigoCliente
        If tarjeta = "0" Then tarjeta = CodigoCliente
    End If
    '++
    If Mid(CodigoCliente, 1, 2) = "1Z" Then
        CodigoCliente = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If tarjeta = "0" Then tarjeta = CodigoCliente
    
    End If
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = cantidad
    c_Importe = Importe
    c_Precio = PrecioLitro
    
    
    
    If Trim(NumFactura) <> "" Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        
        If NifCliente = "" Then
            NifCliente = DevuelveDesdeBDNew(conAri, "ssocio", "nifsocio", "codsocio", CodigoCliente, "N")
        End If
        
        Tarje = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            
            SQL = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES ("
                  
                  'DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtcodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            conn.Execute SQL
                  
            SQL = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            conn.Execute SQL
            
            tarjeta = Tarje
            
        End If
    End If
    
    'MIRAMOS SI EXISTE LA TARJETA
    If Trim(tarjeta) <> "0" Then
        codsoc = ""
        codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", tarjeta, "N")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                  "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
        End If
    Else
        'COGEMOS LA PRIMERA TARJETA DEPENDIENDO DEL TIPO DE ARTICULO
        Dim tipogaso As String
        tipogaso = DevuelveDesdeBD("tipogaso", "sartic", "codartic", IdProducto, "N")
        Select Case tipogaso
            Case "3" ' bonificado
                Tarje = DevuelveDesdeBDNew(conAri, "starje", "numtarje", "tiptarje", "1", "N", , "codsocio", CodigoCliente, "N")
                If Tarje = "" Then
                    Mens = "Nro.Tarjeta Bonif.no existe"
                    SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                          "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                          
                    conn.Execute SQL
                End If
            Case "0", "1", "2", "4"
                Stop 'Tarje = DevuelveValor("select numtarje from starje where tiptarje <> 1 and codsocio =" & DBSet(CodigoCliente, "N"))
                
                If Tarje = "0" Then
                    Mens = "Nro.Tarjeta no existe"
                    SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N") & _
                          "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                          
                    conn.Execute SQL
                End If
        End Select
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        If CDate(Fecha) <> CDate(txtcodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "codartic", "codartic", IdProducto, "N")
    If SQL = "" Then
        Mens = "No existe el artículo"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
        SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    
    'Comprobamos que el socio existe
    If CodigoCliente <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
        If SQL = "" Then
            Mens = "No existe el cliente"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    'Comprobamos que la forma de pago existe
    If idtipopago <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "forpaalvic", idtipopago, "N")
        
        
        If SQL = "" Then
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            SQL = "select codforpa from ssocio where codsocio = " & DBSet(CodigoCliente, "N") & " and codcoope in (1,2) "
            If TotalRegistros(SQL) <> 0 Then Exit Function
            
            
            Mens = "No existe la forma de pago Alvic"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(hora, 1, 2), "N")
            SQL = SQL & "," & DBSet(Mid(hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(idtipopago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistroReg = False
    End If
End Function
            
            
            
            
            
            
            
Private Function InsertarLineaAlz(Cad As String) As Boolean
Dim numlin As String
Dim codpro As String
Dim Articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim SQL As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim FechaHora As String
Dim Fecha As String
Dim hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim tarjeta As String
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

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency
Dim IdProductoDes As String

Dim Tarje As String


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim ForPa As String
Dim Kilometros As String


    On Error GoTo eInsertarLineaAlz

    InsertarLineaAlz = True
    

    Base = Mid(Cad, 1, 10)
    NombreBase = Mid(Cad, 11, 50)
    turno = Mid(Cad, 982, 10) 'txtcodigo(1).Text 'el turno que yo le diga, antes: Mid(cad, 61, 10)
    If CByte(turno) > 9 Then turno = "9"
    NumAlbaran = Mid(Cad, 71, 20)
    NumFactura = Mid(Cad, 94, 17)
    IdVendedor = Mid(Cad, 121, 10)
    NombreVendedor = Mid(Cad, 131, 50)
    FechaHora = Mid(Cad, 181, 14)
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 2) & ":" & Mid(FechaHora, 11, 2) & ":" & Mid(FechaHora, 13, 2)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(Cad, 215, 70)
'    Tarjeta = Mid(cad, 290, 20)
    tarjeta = Mid(Cad, 195, 20)
    Matricula = Mid(Cad, 370, 20)
    IdProducto = Mid(Cad, 493, 20)
    surtidor = Mid(Cad, 538, 10)
    manguera = Mid(Cad, 548, 10)
    PrecioLitro = Mid(Cad, 568, 18)
    cantidad = Mid(Cad, 650, 18)
    Importe = Mid(Cad, 668, 18)
    descuento = Mid(Cad, 586, 18)
    idtipopago = Mid(Cad, 784, 10)
    DescrTipoPago = Mid(Cad, 794, 25)
    CodigoTipoPago = Mid(Cad, 1, 10)
    NifCliente = Mid(Cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(Cad, 415, 18)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)

    If Trim(descuento) <> "" Then
        If CCur(descuento) <> 0 Then
            c_Descuento = Round2(CCur(descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
            IdProductoDes = DevuelveDesdeBDNew(conAri, "sartic", "artdto", "codartic", IdProducto, "N")
        Else
            c_Descuento = 0
        End If
    End If

    Stop   'David:   Actualizando preventa ? estamos locos?
    SQL = "update sartic set preventa = " & DBSet(c_Precio, "N") & _
          ", canstock = canstock - " & DBSet(c_Cantidad, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    conn.Execute SQL
    
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    ForPa = ""
    ForPa = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "forpaalvic", idtipopago, "N")
    

    
    '[Monica]30/11/2011 añadida segunda condicion
    If Trim(NumFactura) <> "" And InStr(1, tarjeta, "Z") <> 0 Then
        codsoc = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Mid(tarjeta, 1, 4) = "****" Or Trim(tarjeta) = "" Then
            tarjeta = codsoc
            
        Else '[Monica]07/02/2011 buscamos la tarjeta que corresponda para meter pq me viene Z
            If codsoc >= 900000 Then
                Stop 'tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N"))
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                
                'DAVID
                ' Lo he comentado para compliar
'                TipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
'                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
'                    tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 1")
'
'                    If tarjeta = "0" Then
'                        codsoc = DevuelveValor("select codsocio from ssocio where codsocio >= 900000 and nifsocio = " & DBSet(NifCliente, "T"))
'                        tarjeta = DevuelveValor("select numtarje from starje where codsocio = " & DBSet(codsoc, "N") & " and tiptarje = 1")
'                    Else
'                        tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 1")
'                    End If
'                Else
'                    tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 0")
'                End If
                'FIN: lo he comentado para compliar

            End If
            
        End If
        'fechahora--> txtcodigo(0).Text & " " & Time
        
        SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        SQL = SQL & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
        
        '[monica]24/06/2013: añadimos los kilometros
        SQL = SQL & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
   
        conn.Execute SQL
        
        If c_Descuento <> 0 Then
            SQL = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                  " where codartic = " & DBSet(IdProductoDes, "N")
            conn.Execute SQL
            
            Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
           
            SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                   DBSet(c_Importe2, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = numlinea + 1
            SQL = SQL & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
            
            '[monica]24/06/2013: añadimos los kilometros
            SQL = SQL & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
        
            conn.Execute SQL
        End If
        
    Else
        '[Monica]30/11/2010
        If Trim(NumFactura) <> "" Then
            codsoc = DevuelveDesdeBDNew(conAri, "starje", "codsocio", "numtarje", tarjeta, "N")
            If Mid(tarjeta, 1, 4) = "****" Or Trim(tarjeta) = "" Then
                tarjeta = codsoc
            End If
            'fechahora--> txtcodigo(0).Text & " " & Time
            
            SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
            SQL = SQL & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
            
            '[monica]24/06/2013: añadimos los kilometros
            SQL = SQL & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
            
            
            conn.Execute SQL
            
            If c_Descuento <> 0 Then
                SQL = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                conn.Execute SQL
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
               
                SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            
                numlinea = numlinea + 1
                SQL = SQL & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ")"
            
                conn.Execute SQL
            End If
        
        Else
            CodigoCliente = DevuelveDesdeBDNew(conAri, "starje", "codsocio", "numtarje", tarjeta, "N")
            SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            SQL = SQL & "0,0,"
            
            '[monica]24/06/2013: añadimos los kilometros
            SQL = SQL & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
            
            
            conn.Execute SQL
            
            If c_Descuento <> 0 Then
                SQL = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                conn.Execute SQL
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
                
                SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
                SQL = SQL & "0,0)"
            
                conn.Execute SQL
            End If
        End If
    End If
 
    
    
eInsertarLineaAlz:
    If Err.Number <> 0 Then
        InsertarLineaAlz = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            

Private Function InsertarSalida(Cad As String) As Boolean
Dim tipMov As String
Dim Importe As Currency
Dim SQL As String
Dim I  As Integer

    On Error GoTo eInsertarSalida
    
    
    InsertarSalida = False
    tipMov = Mid(Cad, 2, 6)
    I = InStr(Mid(Cad, 8, 10), "-")
    If I = 0 Then
        Importe = Format(CCur(TransformaPuntosComas(Mid(Cad, 8, 10))), "######0.00")
    Else
        Importe = Format(CCur(Replace(TransformaPuntosComas(Mid(Cad, 8, 10)), "-", "") * (-1)), "######0.00")
    End If
    
    If tipMov = "MOVIMI" And CCur(Importe) <> 0 Then
        SQL = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
              "99, " & DBSet(Importe, "N") & ",0)"
              
        conn.Execute SQL
    End If
    InsertarSalida = True
eInsertarSalida:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Salida en " & Err.Description, vbExclamation
    End If
End Function

Private Sub InsertarLineaTurno(Cad As String)
Dim codpro As String
Dim cantidad As String
Dim Precio As String
Dim Importe As String
Dim SQL As String
Dim numlin As Long
Dim cWhere As String


    codpro = Mid(Cad, 35, 2)
    cantidad = Mid(Cad, 54, 6) & "," & Mid(Cad, 60, 2)
    Precio = Mid(Cad, 42, 2) & "," & Mid(Cad, 44, 2)
    Importe = Mid(Cad, 47, 5) & "," & Mid(Cad, 52, 2)
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "sturno", "codturno", "fechatur", txtcodigo(0).Text, "F", , "codturno", txtcodigo(1).Text, "N", "codartic", codpro, "N")
    If SQL = "" Then
    
        cWhere = "fechatur=" & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
        numlin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
        'insertamos
        SQL = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
              " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & DBSet(numlin, "N") & ",2,1,1," & _
              DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ",0,0,0)"
              
        conn.Execute SQL
    Else
        'actualizamos
        SQL = "UPDATE sturno SET importel = importel + " & DBSet(Importe, "N") & ", litrosve = litrosve +  " & DBSet(cantidad, "N") & " WHERE fechatur = " & _
              DBSet(txtcodigo(0).Text, "F") & " AND codturno = " & DBSet(txtcodigo(1).Text, "N") & " AND codartic = " & _
              DBSet(codpro, "N")
              
        conn.Execute SQL
    End If
End Sub


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
Dim SQL As String
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute SQL
End Sub






Private Function CrearTMP() As Boolean
' temporales de lineas para insertar posteriormente en scaalp y slialp
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    'tabla temporal con la que cargaremos: scaalp
    SQL = "CREATE TEMPORARY TABLE tmpscaalp ( " '
    SQL = SQL & "`numalbar` varchar(10) NOT NULL default '', "
    SQL = SQL & "`fechaalb` date NOT NULL default '0000-00-00', "
    SQL = SQL & "`codprove` int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "`nomprove` varchar(40) NOT NULL, "
    SQL = SQL & "`domprove` varchar(35) NOT NULL, "
    SQL = SQL & "`codpobla` varchar(6) NOT NULL default '46',"
    SQL = SQL & "`pobprove` varchar(30) NOT NULL default 'A',"
    SQL = SQL & "`proprove` varchar(30) NOT NULL default 'A',"
    SQL = SQL & "`nifprove` varchar(15) NOT NULL default 'A',"
    SQL = SQL & "`telprove` varchar(15) default NULL,"
    SQL = SQL & "`codforpa` smallint(2) NOT NULL default '0',"
    SQL = SQL & "`dtoppago` decimal(4,2) NOT NULL default '0.00',"
    SQL = SQL & "`dtognral` decimal(4,2) NOT NULL default '0.00',"
    SQL = SQL & "`fecturno` date NOT NULL default '0000-00-00', "
    SQL = SQL & "`codturno` tinyint(1) NOT NULL) "
    
    conn.Execute SQL
    
    'tabla temporal con la que cargaremos: slialp
    SQL = "CREATE TEMPORARY TABLE tmpslialp ( " 'TEMPORARY
    SQL = SQL & "`numalbar` varchar(10) NOT NULL default '',"
    SQL = SQL & "`fechaalb` date NOT NULL default '0000-00-00',"
    SQL = SQL & "`codprove` int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "`numlinea` smallint(5) unsigned NOT NULL default '0',"
    SQL = SQL & "`codartic` int(6) NOT NULL,"
    SQL = SQL & "`codalmac` smallint(3) unsigned NOT NULL default '0',"
    SQL = SQL & "`nomartic` varchar(40) NOT NULL default '',"
    SQL = SQL & "`ampliaci` varchar(60) default NULL, "
    SQL = SQL & "`cantidad` decimal(12,2) default NULL,"
    SQL = SQL & "`precioar` decimal(10,5) NOT NULL default '0.00000',"
    SQL = SQL & "`dtoline1` decimal(4,2) NOT NULL default '0.00',"
    SQL = SQL & "`dtoline2` decimal(4,2) NOT NULL default '0.00',"
    SQL = SQL & "`importel` decimal(12,2) NOT NULL default '0.00',"
    SQL = SQL & "`fechahora` datetime)"
    
    conn.Execute SQL
     
    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpscaalp;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpslialp;"
        conn.Execute SQL
    End If
End Function


Private Sub BorrarTMP()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpslialp;"
    conn.Execute " DROP TABLE IF EXISTS tmpscaalp;"
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function PasarTemporales() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset

On Error GoTo ePasar

    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    ' insertamos en tmpinformes: los albaranes que ya estaban en la scaalp CAMPO1 = 1
    SQL = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1) "
    SQL = SQL & " select " & vUsu.Codigo & ", numalbar, fechaalb, codprove, 1 from tmpscaalp "
    SQL = SQL & " where (numalbar, fechaalb, codprove) in (select numalbar,fechaalb,codprove from scaalp) "

    conn.Execute SQL


    conn.Execute " INSERT INTO scaalp (numalbar,fechaalb,codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codforpa,dtoppago,dtognral,fecturno,codturno) SELECT * FROM tmpscaalp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vUsu.Codigo & ") ; "
    conn.Execute " INSERT INTO slialp (numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel) SELECT numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel FROM tmpslialp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vUsu.Codigo & ") ; "
    
    'aqui es donde tenemos que actualizar la cantidad en stock, la fecha y ultimo precio de compra del articulo
    SQL = "SELECT * FROM tmpslialp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "update sartic set ultpreci = " & DBSet(RS!precioar, "N") & _
              ", ultfecha = " & DBSet(txtcodigo(0).Text, "F") & _
              " where codartic = " & DBSet(RS!codArtic, "N") & _
              " and ultfecha < " & DBSet(txtcodigo(0).Text, "F")
        conn.Execute SQL
'        ' solo si tiene control de stock
'        If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(RS!codArtic, "N")) = 1 Then
            SQL = "update sartic set canstock = canstock + " & DBSet(RS!cantidad, "N") & _
                  " where codartic = " & DBSet(RS!codArtic, "N")
            conn.Execute SQL
'        End If
        ' falta insertar en la smoval
        SQL = "insert into smoval (codartic,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,letraser,document,numlinea) values ("
        SQL = SQL & DBSet(RS!codArtic, "N") & ",1,"
        SQL = SQL & DBSet(RS!FechaAlb, "F") & ","
        SQL = SQL & DBSet(RS!FechaHora, "FH") & ","
        SQL = SQL & "'S','ALC'," & DBSet(RS!cantidad, "N") & ","
        SQL = SQL & DBSet(RS!ImporteL, "N") & ","
        SQL = SQL & DBSet(RS!Codprove, "N") & ","
        SQL = SQL & ValorNulo & ","
        SQL = SQL & DBSet(RS!Numalbar, "T") & ","
        SQL = SQL & DBSet(RS!numlinea, "N") & ")"
        
        conn.Execute SQL
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ' actualizamos la fecha de ultimo movimiento del proveedor
    SQL = "SELECT * FROM tmpscaalp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "update proveedor set fechamov = " & DBSet(txtcodigo(0).Text, "F") & _
              " where codprove = " & DBSet(RS!Codprove, "N") & _
              " and fechamov < " & DBSet(txtcodigo(0).Text, "F")
        conn.Execute SQL
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ' insertamos en tmpinformes: los proveedores que estan introducidos automaticamente CAMPO1 = 2
    SQL = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    SQL = SQL & " select " & vUsu.Codigo & ", '' ," & ValorNulo & ", codprove, 2, nomprove from proveedor where domprove = 'AUTOMATICO'"
    
    conn.Execute SQL
    
    ' insertamos en tmpinformes: los articulos que estan introducidos automaticamente CAMPO1 = 3
    SQL = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    SQL = SQL & " select " & vUsu.Codigo & ", '', " & ValorNulo & ", codartic, 3, nomartic from sartic where artnuevo = 1 "
        
    conn.Execute SQL
    
    ' insertamos en tmpinformes: las familias que se han generado automaticamente CAMPO1 = 4
    SQL = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    SQL = SQL & " select " & vUsu.Codigo & ", '', " & ValorNulo & ", codfamia, 4, nomfamia from sfamia where nomfamia = 'AUTOMATICO'"
        
    conn.Execute SQL
    
    
    PasarTemporales = True
    Exit Function
ePasar:
    PasarTemporales = False
End Function



Private Function ComprobarFechaAlbaran(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim I As Integer
Dim Longitud As Long
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eComprobarFechaAlbaran
    
    ComprobarFechaAlbaran = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    I = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    Longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = Longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO COMPRAS

    b = True

    While Not EOF(NF) And b
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        b = ComprobarFecha(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        b = ComprobarFecha(Cad)
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ComprobarFechaAlbaran = b
    Exit Function

eComprobarFechaAlbaran:
    ComprobarFechaAlbaran = False
End Function




Private Function ComprobarFecha(Cad As String) As Boolean
Dim SQL As String

Dim Albaran As String
Dim FechaHora As String

Dim Fecha As String
Dim hora As String

Dim Mens As String


Dim codsoc As String

    On Error GoTo eComprobarFecha

    ComprobarFecha = True

    Albaran = Mid(Cad, 92, 15)
    FechaHora = Mid(Cad, 122, 14)
    
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 2) & ":" & Mid(FechaHora, 11, 2) & ":" & Mid(FechaHora, 13, 2)

    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
    Else
        If CDate(Fecha) <> CDate(txtcodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
eComprobarFecha:
    If Err.Number <> 0 Then
        ComprobarFecha = False
    End If
End Function



Private Function InsertarLineaReg(ByRef RS As ADODB.Recordset) As Boolean
Dim numlin As String
Dim codpro As String
Dim Articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim SQL As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim FechaHora As String
Dim Fecha As String
Dim hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim tarjeta As String
Dim CodigoProducto As String
Dim surtidor As String
Dim manguera As String
Dim PrecioLitro As String
Dim descuento As String
Dim PorcDescuento As Currency
Dim cantidad As String
Dim Importe As String
Dim idtipopago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency
Dim c_Descuento As Currency
Dim c_Vale As Currency
Dim c_Devolucion As Currency
Dim Tarje As String

Dim SqlVale As String
Dim RsVale As ADODB.Recordset


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim ForPa As String

Dim Kilometros As String
Dim NomArtic As String

    On Error GoTo EInsertarLinea

    InsertarLineaReg = True
    
    turno = DBLet(RS!turno, "N")
    
    NumAlbaran = DBLet(RS!Albaran, "N")
    NumFactura = DBLet(RS!Factura, "T")
'    If NumFactura <> "" Then
'        NumFactura = Mid(NumFactura, 5, Len(NumFactura) - 4)
'    End If
    If NumFactura <> "" Then
        If Mid(NumFactura, 1, 3) = "FAV" Then
            NumFactura = "9" & Mid(NumFactura, Len(NumFactura) - 5, 6)
        Else
            NumFactura = Mid(NumFactura, Len(NumFactura) - 6, 7)
        End If
    End If

    
    FechaHora = DBLet(RS!Fecha, "T")
    Fecha = Mid(FechaHora, 7, 2) & "/" & Mid(FechaHora, 5, 2) & "/" & Mid(FechaHora, 1, 4)
    hora = Mid(FechaHora, 9, 2) & ":" & Mid(FechaHora, 11, 2) & ":" & Mid(FechaHora, 13, 2)
    CodigoCliente = DBLet(RS!Cliente, "T")
    NombreCliente = DBLet(RS!NomClien, "T")
    
    tarjeta = DBLet(RS!tarjeta, "N")
    Matricula = DBLet(RS!Matricula, "T")
    IdProducto = DBLet(RS!producto, "N")
    surtidor = DBLet(RS!surtidor, "N")
    manguera = DBLet(RS!manguera, "N")
    
    PrecioLitro = DBLet(RS!Precio, "N")
    cantidad = DBLet(RS!cantidad, "N")
    Importe = DBLet(RS!Importe, "N")
    idtipopago = DBLet(RS!idtipopago, "N")
    DescrTipoPago = DBLet(RS!desctipopago, "T")
    CodigoTipoPago = DBLet(RS!idtipopago, "N")
    NifCliente = DBLet(RS!NIF, "T")
    
    ' en caso de que el codigo de cliente y el nombre no me vengan cojo el asociado a la forma de pago
    If CodigoCliente = "" And NombreCliente = "" Then
        CodigoCliente = DevuelveDesdeBDNew(conAri, "sforpa", "codsocio", "forpaalvic", idtipopago, "N")
        NombreCliente = DevuelveDesdeBDNew(conAri, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
        tarjeta = CodigoCliente
    End If
    
    Kilometros = DBLet(RS!km, "N")
    PorcDescuento = DBLet(RS!descuentoporc, "N")
    descuento = Round(PrecioLitro * PorcDescuento / 100, 3)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
'    If NifCliente = "20763891C" Then
'        Stop
'    End If
    
    c_Cantidad = cantidad 'Round2(CCur(cantidad) / 100, 2)
    c_Importe = Importe 'Round2(CCur(Importe) / 100, 2)
    
    
    '[Monica]03/01/2017: antes estaba preciolitro - descuento
'    c_Precio = PrecioLitro - Descuento 'Round2(CCur(PrecioLitro) / 100000, 5)
    c_Precio = PrecioLitro - descuento
    If c_Cantidad <> 0 Then
        c_Precio = Round2(c_Importe / c_Cantidad, 3)
    End If
    '[Monica]03/01/2017: el descuento ahora lo calculo
'    c_Descuento = Descuento 'Round2(CCur(Descuento) / 100000, 5)
    If c_Cantidad <> 0 Then
        c_Descuento = PrecioLitro - c_Precio
    Else
        c_Descuento = descuento
    End If
    
    
    c_Vale = 0
    
    SqlVale = "select * from tmptraspaso where codusu = " & DBSet(vUsu.Codigo, "N") & " and albaran = " & DBSet(NumAlbaran, "N")
    SqlVale = SqlVale & " and idtipopago in (select forpaalvic from sforpa where tipovale = 1) "
    Set RsVale = New ADODB.Recordset
    RsVale.Open SqlVale, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RsVale.EOF Then
        c_Vale = DBLet(RsVale!Importe, "N")
    End If
    Set RsVale = Nothing
    
    c_Importe = c_Importe + c_Vale
    
    ' lo mismo con la devolucion de billetes
    c_Devolucion = 0
    
    SqlVale = "select * from tmptraspaso where codusu = " & DBSet(vUsu.Codigo, "N") & " and albaran = " & DBSet(NumAlbaran, "N")
    SqlVale = SqlVale & " and idtipopago in (select forpaalvic from sforpa where tipovale = 2) "
    Set RsVale = New ADODB.Recordset
    RsVale.Open SqlVale, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RsVale.EOF Then
        c_Devolucion = DBLet(RsVale!Importe, "N")
    End If
    Set RsVale = Nothing
    
    c_Importe = c_Importe + c_Devolucion
    
    
    'VRS:4.0.1(0) actualizamos el precio de articulo
    SQL = "update sartic set preventa = " & DBSet(PrecioLitro, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    conn.Execute SQL
    
    If DevuelveDesdeBD(conAri, "ctrstock", " sartic", "codartic", IdProducto, "T") = 1 Then
        SQL = "update sartic set " & _
              "  canstock = canstock - " & DBSet(c_Cantidad, "N") & _
              " where codartic = " & DBSet(IdProducto, "N")
        conn.Execute SQL
    End If
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    ForPa = ""
    ForPa = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "forpaalvic", idtipopago, "N")
    
    
    If Trim(NumFactura) <> "" Then
        codsoc = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        
        '[Monica]04/01/2015: en el caso de venga una factura sin nif, cogemos el de la forma de pago
        If codsoc = "" Then
            CodigoCliente = DevuelveDesdeBDNew(conAri, "sforpa", "codsocio", "forpaalvic", idtipopago, "N")
            NombreCliente = DevuelveDesdeBDNew(conAri, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
            tarjeta = CodigoCliente
            If tarjeta = "0" Then tarjeta = CodigoCliente
            codsoc = CodigoCliente
        Else
            '[Monica]17/06/2013: miramos si la tarjeta viene con algun asterisco
            If Mid(tarjeta, 1, 4) = "****" Or Trim(tarjeta) = "0" Or InStr(1, tarjeta, "*") <> 0 Then
                tarjeta = codsoc
            Else '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
                If Len(Trim(tarjeta)) = 16 Then
                    tarjeta = Mid(tarjeta, 9, 16)
                End If
                '++
            End If
            'fechahora--> txtcodigo(0).Text & " " & Time
        End If
        
        
        SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        SQL = SQL & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
    Else
        If InStr(1, CodigoCliente, "1Z") <> 0 Then
            
            codsoc = DevuelveDesdeBDNew(conAri, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
            
            If tarjeta = "0" Then
                Tarje = DevuelveDesdeBDNew(conAri, "starje", "numtarje", "numtarje", tarjeta, "T")
                If Tarje = "" Then tarjeta = codsoc
            End If
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            SQL = "select codforpa from ssocio where codsocio = " & DBSet(codsoc, "N") & " and codcoope in (1,2) "
            If TotalRegistros(SQL) <> 0 Then
                ForPa = DevuelveValor(SQL)
            End If
            
            
            
            SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            SQL = SQL & "0,0,"
        Else
            If tarjeta = "0" Then
                'COGEMOS LA PRIMERA TARJETA DEPENDIENDO DEL TIPO DE ARTICULO
                Dim tipogaso As String
                tipogaso = DevuelveDesdeBD("tipogaso", "sartic", "codartic", IdProducto, "N")
                Select Case tipogaso
                    Case "3" ' bonificado
                        tarjeta = DevuelveDesdeBDNew(conAri, "starje", "numtarje", "tiptarje", "1", "N", , "codsocio", CodigoCliente, "N")
                    Case "0", "1", "2", "4"
                        tarjeta = DevuelveValor("select numtarje from starje where tiptarje <> 1 and codsocio = " & DBSet(CodigoCliente, "N"))
                End Select
            End If
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            SQL = "select codforpa from ssocio where codsocio = " & DBSet(CodigoCliente, "N") & " and codcoope in (1,2) "
            If TotalRegistros(SQL) <> 0 Then
                ForPa = DevuelveValor(SQL)
            End If
            
            
            
            SQL = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & hora, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(ForPa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            SQL = SQL & "0,0,"
            
        End If
    End If
    
    '[monica]24/06/2013: añadimos los kilometros
    SQL = SQL & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & "," '& ")"
 
 
    '[Monica]24/08/2015: añadimos el descuento
    SQL = SQL & DBSet(c_Descuento, "N") & "," & DBSet(c_Vale, "N") & ")"
 
    conn.Execute SQL
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaReg = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function

Private Function InsertarLineaTurnoReg(ByRef RS As ADODB.Recordset) As Boolean
Dim NF As Long
Dim I As Long
Dim Longitud As Long


Dim codpro As String
Dim cantidad As String
Dim Precio As String
Dim Importe As String
Dim SQL As String
Dim numlin As Long
Dim cWhere As String

Dim surtidor As String
Dim manguera As String
Dim Inicial As String
Dim Final As String
Dim vInicial As Currency
Dim vFinal As Currency

    On Error GoTo eInsertarLineaTurnoNew

    InsertarLineaTurnoReg = True

            
    codpro = DBLet(RS!producto, "N")
    cantidad = DBLet(RS!cantidad, "N")
    Precio = DBLet(RS!Precio, "N")
    Importe = DBLet(RS!Importe, "N")
    surtidor = DBLet(RS!surtidor, "N")
    manguera = DBLet(RS!manguera, "N")
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(conAri, "sturno", "codturno", "fechatur", txtcodigo(0).Text, "F", , "codturno", txtcodigo(1).Text, "N", "codartic", codpro, "N")
    If SQL = "" Then
    
        cWhere = "fechatur=" & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
        numlin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
        'insertamos
        ' antes surtidor y manguera: 1,1,
        SQL = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
              " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & DBSet(numlin, "N") & ",2," & DBSet(surtidor, "N") & "," & DBSet(manguera, "N") & "," & _
              DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ",0,0,0)"
              
        conn.Execute SQL
    Else
        'actualizamos
        SQL = "UPDATE sturno SET importel = importel + " & DBSet(Importe, "N") & ", litrosve = litrosve +  " & DBSet(cantidad, "N") & " WHERE fechatur = " & _
              DBSet(txtcodigo(0).Text, "F") & " AND codturno = " & DBSet(txtcodigo(1).Text, "N") & " AND codartic = " & _
              DBSet(codpro, "N")
              
        conn.Execute SQL
    End If
            
eInsertarLineaTurnoNew:
    If Err.Number <> 0 Then
        InsertarLineaTurnoReg = False
        MsgBox "Error en Insertar Turno en " & Err.Description, vbExclamation
    End If
End Function

Private Function InsertarRecaudacionReg() As Boolean
Dim ForPa As String
Dim Importe As String
Dim SQL As String
Dim vImporte As String
Dim vForpaVale As String
Dim idtipopago As String
Dim Existe As String

    On Error GoTo eInsertarRecaudacion

    InsertarRecaudacionReg = True
    
    SQL = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) "
    SQL = SQL & " select " & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & ", codforpa, sum(importel-coalesce(importevale,0)), 0 "
    SQL = SQL & " from scaalb where fecalbar = " & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
    SQL = SQL & " group by 1,2,3 "
    SQL = SQL & " order by 1,2,3 "
    
    conn.Execute SQL

    SQL = "select sum(coalesce(importevale,0)) from scaalb where fecalbar = " & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
    vImporte = DevuelveValor(SQL)
    vForpaVale = DevuelveValor("select codforpa from sforpa where tipovale = 1")
    If vImporte <> 0 Then
        SQL = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
        SQL = SQL & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & DBSet(vForpaVale, "N") & "," & DBSet(vImporte, "N") & ",0) "
    
        conn.Execute SQL
    End If


eInsertarRecaudacion:
    If Err.Number <> 0 Then
        InsertarRecaudacionReg = False
        MsgBox "Error en Insertar Recaudacion en " & Err.Description, vbExclamation
    End If
    
End Function




Public Function EsArticuloCombustible(Articulo As String) As Boolean
Dim Famia As String
Dim tipoF As String

    EsArticuloCombustible = False
    Famia = ""
    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", Articulo, "N")
    If Famia = "" Then Exit Function
    tipoF = ""
    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
    If tipoF = "" Then Exit Function
    If tipoF = "1" Then EsArticuloCombustible = True

End Function

