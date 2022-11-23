VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión listados"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAlbaran1A1 
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   6375
      Begin VB.OptionButton optAlbValorado 
         Caption         =   "Cantidad"
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
         Index           =   2
         Left            =   4680
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optAlbValorado 
         Caption         =   "Cantidad y precio"
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
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optAlbValorado 
         Caption         =   "Valorado"
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
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame FrameSelecRPT 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionar informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   180
         Width           =   2895
      End
   End
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3060
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   3060
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
    'Equivale a OpcionListado en frmListado
    'SI ES MAYOR QUE 2000 es ke viene de frmListado2
    
Public FormulaSeleccion As String 'Formula de Seleccion para Crystal Report

Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer
'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes

Public SoloImprimir As Boolean
Public EnvioEMail As Boolean
Public PulsaAceptar As Boolean   'Si no es solo imprimirm y no es envioemail, si este esta activo simula el pulsar boton


Public NombreRPT As String 'Nombre del fichero de crystal Report .Rpt
Public NombrePDF As String 'Para cunado envie por email. Al unload se pone a ""


Public Titulo As String 'Titulo informe a mostrar en el text1

Public NombreSubRptConta As String 'Nombre del subreport si va conectado a la BDatos Contabilidad

Public ConSubInforme As Boolean 'Para saber si hay subinformes y hay que enlazar las
                                 'tablas a la BD correspondiente
Public MostrarTreeDesdeFuera As Boolean
        'Para indicar si muestra el tree o no





'Febrero 2010.
'Vamos a enviar mail abriendo el outlook
'Con lo cual, pasaremos ciertos valores
Public outCodigoCliProv As Long
Public outTipoDocumento As Byte
        '0 UNDEFINNED. Si es cero NO va por este trozo de programa
        '1.- Oferta cliente
        '2.- Pedido cliente
        '
        '
        'a partir del 50 van proveedores

Public outClaveNombreArchiv As String  'Llevara el codigo oferta, pedido alb.....  SIN el .pdf, solo el nombre
Public NumeroCopias As Byte

'ENERO 2015
' EULER. Puede seleccionar que RPT seleccionara
Public SeleccionaRPTCodigo As Integer




Private MostrarTree As Boolean

Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean


Private EstabaMarcado As Boolean


Private ImpresoraPorDefectoAnterior As String


'Private ReestableceSoloImprimir As Boolean
Private Sub chkEMAIL_Click()
    If chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEmail.Value = 0
End Sub


Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    CommonDialog1.ShowPrinter
    PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
 
    If Me.chkSoloImprimir.Value = 1 And Me.chkEmail.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
        Exit Sub
    End If
    
    
    
    
    Imprime
    
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Combo1_Click()
Dim i As Integer
Dim C  As String
    If PrimeraVez Then Exit Sub
    'En nomrpt pondra el valor entrecorchetado
    C = Combo1.Text
    i = InStr(1, C, "[")
    C = Mid(C, i + 1)
    i = InStr(1, C, "]")
    C = Mid(C, 1, i - 1)
    Me.NombreRPT = C
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        
        
        
        If SoloImprimir Then
           
            Imprime
            
            Unload Me
            
        ElseIf Me.EnvioEMail Then
            
            Unload Me
            
           
        
        Else
            
           
                
            If PulsaAceptar Then Imprime
                

        End If
        If Not EnvioEMail Then Espera 0.1
        CommitConexion
    End If
    Screen.MousePointer = vbDefault
End Sub

'Si no he ajustado el NombrePDF y no le he puesto valor entonces,
'cogera el mismo que tiene en NombreRPT
Private Sub Form_Load()
Dim Cad As String
Dim jj As Integer

    PrimeraVez = True
    Lanzado = False
    CargaICO
    Cad = Dir(App.Path & "\impre.dat", vbArchive)
    HaPulsadoElBotonDeImprimir = False
    
    FrameAlbaran1A1.visible = False
        
    
    
    'ReestableceSoloImprimir = False
    If Cad = "" Then
        chkSoloImprimir.Value = 0
    Else
        chkSoloImprimir.Value = 1
        'ReestableceSoloImprimir = True
    End If
    EstabaMarcado = chkSoloImprimir.Value = 1
    cmdImprimir.Enabled = True
    Me.FrameSelecRPT.visible = False
    If SoloImprimir Then
        chkSoloImprimir.Value = 0
        Me.Frame2.Enabled = False
        chkSoloImprimir.visible = False
    Else
        Frame2.Enabled = True
        chkSoloImprimir.visible = True
         If SeleccionaRPTCodigo > 0 Then CargaComboRPTS
    End If
    PonerNombreImpresora
    MostrarTree = False

'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)

    MIPATH = App.Path & "\Informes\"
'    ConSubInforme = False


    If Opcion >= 2000 Then
        'LISTADOS QUE VIENE de frmlistado2 o de listado 3
        If Opcion = 3000 Then
            'Desde frmlistado 3, TODO viene por parametros
            MostrarTree = MostrarTreeDesdeFuera
        Else
            Select Case Opcion
            Case 2001 'Confirmacion de Pedido
                Text1.Text = "Reparaciones efectuadas"
                ConSubInforme = False
                MostrarTree = True
                NombreRPT = "rRepEfectuadas.rpt"
                 Titulo = ""
            Case 2002
                
                Text1.Text = IIf(Titulo = "", "Listado reparacion x tecnico", Titulo)
                 Titulo = ""
            
            Case 2003
                'Esta libre. Lo utlizo para la impresion del justificante del pago de la regarga
                Text1.Text = "Justificante recarga móviles"
                
            Case 2004
                Text1.Text = "Listado Recarga móviles"
                ConSubInforme = False
                MostrarTree = True
                NombreRPT = "rRecargaMov.rpt"
            Case 2006
                Text1.Text = "Listados ventas por proveedor"
                MostrarTree = True
                ConSubInforme = False
                'El nombre lo dejo que venga del form listado2
                 Titulo = ""
            Case 2009
                Text1.Text = "Facturas proveedor"
                 Titulo = ""
                
            Case 2010
                Text1.Text = "Albarán proveedor"
                Titulo = ""
            Case 2014
                Text1.Text = "Listado tickets agrupados"
          
            Case 2015
                Text1.Text = "Informe traza."
                
            Case 2016
                Text1.Text = "Ventas por agente"
                If NombreRPT = "" Then NombreRPT = "ragentes.rpt"
                MostrarTree = True
                 Titulo = ""
            Case 2017
                Text1.Text = "Trabajadores"
                 Titulo = ""
            Case 2018
                Text1.Text = "CRM"
            Case 2023
                Text1.Text = "Situacion de albaranes"
            Case 2027
                Text1.Text = "Albaranes trasporte"
                Titulo = ""
                        
            Case 2030
                Text1.Text = "Reparaciones garantia proveedor"
                MostrarTree = True
                 Titulo = ""
            Case 2032
                Text1.Text = "Propuesta pedido"
                MostrarTree = True
                 Titulo = ""
            Case 2033
                Text1.Text = "Informe descuentos proveedor"
                MostrarTree = True
                NombreRPT = "rCompraDto.rpt"
                 Titulo = ""
        
            Case 2035
                Text1.Text = "Informe descuentos actividad"
                MostrarTree = False
                NombreRPT = "rFacActivDtos.rpt"
                 Titulo = ""
            Case 2036
                Text1.Text = "Vtas agente/marca"
                MostrarTree = True
                Titulo = ""
                
            Case 2037, 2040
                
                Text1.Text = "Beneficio por agente"
                If Opcion = 2040 Then Text1.Text = "Beneficio por proveedor"
                MostrarTree = True
                Titulo = ""
            Case 2041
                
                Text1.Text = "Beneficio por cliente"
                MostrarTree = True
                Titulo = ""
            Case 2042
                Text1.Text = "Listado control albaranes"
            Case 2043
                Text1.Text = "Listado control albaranes facturados"
            Case 2046
                Text1.Text = "Listado control productividad"
                    
            Case 2048
                Text1.Text = "Beneficio Marca, Agente, Proveedor"
                MostrarTree = True
            Case 2049
                Text1.Text = "Ventas marca-familia"
                MostrarTree = True
            Case 2050
                Text1.Text = "Compras marca-familia"
                MostrarTree = True
            
            Case 2051
                'Text1.Text = "Compras marca-familia"
                NombreRPT = "rFacPedxClienIVA.rpt"
            Case 2052
                'pEdidos por articulo agr
                NombreRPT = "rFacPedxArticAgr.rpt"
            Case 2053
                Text1.Text = "Costes cliente-factura"
            Case 2054
                Text1.Text = "Ficha mantenimiento"
                
            Case 2055
                Text1.Text = "Etiquetas QR envasado"
            End Select
        End If
    Else
        'Normal. Los de antes
                If Opcion <= 40 Then
                    Select Case Opcion
                    
                    
                    '---------------- Algunos listados basicos
                    Case 5
                        'Tipos de contrato de mantenimiento
                        Text1.Text = "Tipo contrato mantenimiento"
                        
                    Case 18 'Informe Stocks Maximos o Minimos
                        Text1.Text = "Stocks Máximos-Mínimos"
                    Case 30
                        MostrarTree = True
                    Case 31 'Listado de Ofertas
                        Text1.Text = "Listado de Ofertas"
                        ConSubInforme = True
                    Case 32 'Listado Recordatorio de Ofertas
                        Text1.Text = "Recordatorio de Ofertas"
                        ConSubInforme = True
                    Case 33 'Listado Valoracion de Ofertas
                        Text1.Text = "Valoracion de Ofertas"
                
                    Case 35 'Listado Historico de Ofertas
                        Text1.Text = "Histórico de Ofertas"
                        ConSubInforme = True
                    Case 36 'Listado Ofertas Pendientes y Traspaso a Historico
                        Text1.Text = "Ofertas Pendientes"
                        NombreRPT = "rFacOfePtes.rpt"
                
                    Case 39 'Orden de Instalacion
                        Text1.Text = "Orden de Instalación"
                        ConSubInforme = True
                    Case 40 'Confirmacion de Pedido
                        Text1.Text = "Confirmación de Pedido"
                        ConSubInforme = True
                        
                    
                    Case Else
                        Text1.Text = "Opcion incorrecta"
                        Me.cmdImprimir.Enabled = False
                    End Select
                ElseIf Opcion < 100 Then
                    Select Case Opcion
                    Case 41 'Informe de Pedidos por Articulo
                        Text1.Text = "Pedidos por Articulo"
                        NombreRPT = "rFacPedxArtic.rpt"
                    Case 42 'Informe de Disponibilidad de Stocks
                        Text1.Text = "Disponibilidad de Stocks"
                        'NombreRPT = "rFacPedDispStocks.rpt"
                        ConSubInforme = True
                        MostrarTree = True
                    Case 44 'Informe de Pedidos por Cliente
                        Text1.Text = "Pedidos por Cliente"
                        NombreRPT = "rFacPedxClien.rpt"
                        MostrarTree = True
                    '45: Informe de Albaranes
                    Case 45
                    
                        'Viene un parametro ... |Albarcon=0|   '0 o 1
                        FrameAlbaran1A1.Tag = ""
                        If vParamAplic.NumeroInstalacion = vbHerbelca Then
                            jj = InStr(1, OtrosParametros, "|Albarcon=")
                        Else
                            jj = 0  'Solo para herbelca
                        End If
                        If jj > 0 Then
                            
                            
                            
                            
                            
                            Cad = Mid(OtrosParametros, jj + 10, 2)   'Buscamos ->  0| o 1|
                            ' 0 todo    1  cantidad y precio    2- cantidad
                            If Cad = "0|" Then
                                Me.optAlbValorado(0).Value = True
                            Else
                                If Cad = "1|" Then
                                    Me.optAlbValorado(1).Value = True
                                Else
                                    If Cad = "2|" Then
                                        Me.optAlbValorado(2).Value = True
                                    Else
                                        Cad = ""
                                    End If
                                End If
                            End If
                            
                            If Cad <> "" Then
                                FrameAlbaran1A1.Tag = "|Albarcon=" & Cad
                                FrameAlbaran1A1.Left = 240
                                FrameAlbaran1A1.visible = True
                            End If
                            
                        End If
                        
                    
                    
                    
                    Case 46
                        'TRAMPA.
                        'No estaba utlizado.
                        'Lo aprovecho para llamar a un report parecidor al 44
                        Text1.Text = "Pedidos por Cliente SIN VALORAR"
                        NombreRPT = "rFacPedxClienSIN.rpt"
                        MostrarTree = True
                        
                    Case 47 'Listado de Clientes
                        Text1.Text = "Listado de Cliente"
                        'Feb 2011 Ahora se lo manda el formlistado
                        MostrarTree = Not (NombreRPT = "rFacClientes.rpt")
                        
                    Case 48 'Informe Altas Nuevos Clientes
                        Text1.Text = "Altas Nuevos Clientes"
                    Case 49 'Informe de Albaranes por Articulo
                        Text1.Text = "" ' dejamos la cadena vacía para que use Titulo [SERVICIOS]
                        NombreRPT = "rFacAlbxArtic.rpt"
                    Case 53 'Factura cliente
                        Text1.Text = "Factura Cliente"
                        ConSubInforme = True
                    Case 54 'Listado Descuentos Familia/Marca
                        Text1.Text = "Listado Descuentos Familia/Marca"
                        'NombreRPT = "rFacDtosFM.rpt" Se lo indico en frmlistado
                    Case 58 'Listado Proveedor
                        Text1.Text = "Listado Proveedores"
                        ConSubInforme = False
                         NombreRPT = "rComProve.rpt"
                    Case 60 'Informe Equipos con Nº Serie
                        Text1.Text = "Equipos con Nº Serie"
                        ConSubInforme = True
                    Case 61 'Informe Motivos Pend. Rep.
                        NombreRPT = "rRepMotivosPend.rpt"
                        Text1.Text = "Motivos Pend. Rep."
                        
                    ' ---- [11/11/2009] [LAURA] : lo paso como parametro al llamar al form
'                    Case 62 'Listado Resguardo Reparacion
'                        Text1.Text = "Resguardo Reparación"
                    ' ----
                    
                    Case 63 'FACTURAs del TPV
                        Text1.Text = "Facturas formato TPV"
                        ConSubInforme = True
                    
                    Case 65 'Informe Motivos Baja equipos
                        NombreRPT = "rRepMotivosBaja.rpt"
                        Text1.Text = "Motivos Baja equipos"
                    
                    Case 78
                        Titulo = "Carta renovación mantenimientos"
                    
                    
                    Case 79
                        Titulo = "Etiquetas de mantenimiento"
                        'NombreRPT = "rManClienEtiq.rpt"
                        
                    Case 230
                        
                    ' ---- [06/11/2009] [LAURA] : corregir informe de frecuencias
                    '      estos valores se pasan ya al llamar al form desde mto frecuencias
'                    Case 96
                        'FRECUENCIAS
'                        Titulo = "Frecuencias"
'                        ConSubInforme = False
'                        NombreRPT = "rFrequ.rpt"
                    ' ----

                    Case Else
                        If Titulo = "" And NombreRPT = "" Then
                            Text1.Text = "Opcion incorrecta"
                            Me.cmdImprimir.Enabled = False
                        End If
                    End Select
                Else
                    If Opcion = 230 Then MostrarTree = MostrarTreeDesdeFuera
                End If
End If
    If Titulo <> "" Then
        Text1.Text = Titulo
        Me.cmdImprimir.Enabled = True
    End If
    
    If NombrePDF = "" Then NombrePDF = NombreRPT
    
    If EnvioEMail Then
        Imprime
        Screen.MousePointer = vbDefault
    Else
        'Enero 2021
        'Si es factura cliente(53) y solo imprimir tambien lo lanzo aqui
        If Opcion = 53 And SoloImprimir Then
            Screen.MousePointer = vbHourglass
            Imprime   'para que no haga la hola
        End If
    End If
    
    
End Sub


Private Function Imprime() As Boolean
Dim LanzaAbrirOutlook As Boolean
Dim ImpresionFacturas As Boolean
Dim OtrosParam2 As String
Dim NumParam2 As Integer
Dim HaPulsadoImprimir As Boolean
Dim J As Integer
Dim EulerT As String
Dim Aux As String
Dim EsPorEmail As Boolean

    Screen.MousePointer = vbHourglass
    OtrosParam2 = OtrosParametros
    NumParam2 = NumeroParametros
    HaPulsadoImprimir = False
    If Opcion = 53 And Me.chkEmail.Value = 0 Then
        'Estamos en
        '   -reimpresion de facturas
        '   -facturacion
        'Con lo cual, si manda mas de una copia haremos luego el bucle
        ImpresionFacturas = True
        If NumeroCopias > 1 Then
            OtrosParam2 = OtrosParametros & "KTexto=1|"
            NumParam2 = NumeroParametros + 1
        End If

    Else
        ImpresionFacturas = False
    
    End If
    
    
    
    If Opcion = 45 And CStr(Me.FrameAlbaran1A1.Tag) <> "" Then
        Aux = ""
        
        If Me.optAlbValorado(0).Value Then Aux = "0"
        If Me.optAlbValorado(1).Value Then Aux = "1"
        If Me.optAlbValorado(2).Value Then Aux = "2"
         
        If Aux <> "" Then
            Aux = "|Albarcon=" & Aux & "|"
            If FrameAlbaran1A1.Tag <> Aux Then OtrosParam2 = Replace(OtrosParam2, FrameAlbaran1A1.Tag, Aux)
            
        End If
    End If
    If False Then
        
    End If
    
    CadenaDesdeOtroForm = ""
    With frmVisReport
            
        .CambiaODBC = False
        .OcultarElMensajeDeError = False
            
        '.ForzarNombreImpresora
        'ETIQUETAS TAXCo
        If Opcion = 513 And vParamAplic.NumeroInstalacion = vbTaxco Then
            .ForzarNombreImpresora = "GODEX500"
            .SoloImprimir = True
        End If
         If Opcion = 2055 And vParamAplic.NumeroInstalacion = vbFontenas Then
            .ForzarNombreImpresora = "Godex G500-1"
            .SoloImprimir = True
        End If
        If Opcion = 93 Then
            If vParamTPV.NomImpresora <> "" Then .ForzarNombreImpresora = vParamTPV.NomImpresora
        End If
        EsPorEmail = False
        If EnvioEMail Then
            EsPorEmail = True
        Else
            If Me.chkEmail.Value = 1 Then EsPorEmail = True
        End If
        If EsPorEmail Then
            'EMAIL
            
            'En EULER, el rp
            If SeleccionaRPTCodigo > 0 Then
                If InstalacionEsEulerTaxco Then NombrePDF = NombreRPT
            End If
            .Informe = MIPATH & NombrePDF
        Else
            'IMPRIMIR
            .Informe = MIPATH & NombreRPT
        End If
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParam2
        .NumeroParametros = NumParam2
        .ConSubInforme = ConSubInforme
        'Si es impresion de facturas el proceso de numero de copias es distinto
        If ImpresionFacturas Then
            .NumCopias = 1 'hay un bucle ahi abajo
        Else
            .NumCopias = NumeroCopias
        End If
        .Opcion = Opcion
        .ExportarPDF = EsPorEmail
        .MostrarTree = MostrarTree
        
        If EsPorEmail Then
            Load frmVisReport
            Unload frmVisReport
            
            'De momento ciertos casos
            '
            'if outTipoDocumento<>"" the
            If Opcion = 238 Then .EstaImpreso = True   'confirmacion de entrega
        Else
             .Show vbModal
        End If
        
        
        
        HaPulsadoImprimir = .EstaImpreso
        HaPulsadoElBotonDeImprimir = HaPulsadoImprimir
    End With
    
    
    
    
    If ImpresionFacturas Then
        If Me.chkSoloImprimir.Value = 0 Then HaPulsadoImprimir = True
        
        If HaPulsadoImprimir And NumeroCopias > 1 Then
            Text1.Text = "Enviando copias"
            Text1.Refresh
            Espera 0.5
            For J = 2 To NumeroCopias
                Me.Refresh
                DoEvents
                Text1.Text = "Copia:" & J
                Text1.Refresh
                OtrosParam2 = OtrosParametros & "KTexto=" & J & "|"
                NumParam2 = NumeroParametros + 1
                Espera 0.5
                With frmVisReport
                    .Informe = MIPATH & NombreRPT
    
                 
                    .FormulaSeleccion = Me.FormulaSeleccion
                    .SoloImprimir = True
                    .OtrosParametros = OtrosParam2
                    .NumeroParametros = NumParam2
                    .ConSubInforme = ConSubInforme
                    .Opcion = Opcion
                    .ExportarPDF = False
                    .MostrarTree = MostrarTree
                    .Show vbModal
                    
                End With
            Next
        End If
            
            
        If HaPulsadoImprimir Then
            'Vamos a guardar con que RPT han impreso la factura
            CadenaRPTFactura True, NombreRPT, FormulaSeleccion
        End If
        
    End If
    
    
    
    
    
    
    If Me.chkEmail.Value = 1 Then
        If CadenaDesdeOtroForm <> "" Then 'se exporto el informe OK (.pdf)
            
            If Me.EnvioEMail Then  'se llamo desde envio masivo
'                frmEMail.Show vbModal
                
            Else 'informe normal, pero que se selecciono enviar e-mail
            
                'Febrero 2010
                ' Nuevo
                LanzaAbrirOutlook = False
                If vParamAplic.ExeEnvioMail <> "" Then
                
                    '28 MAYO 2021
                    'Todos se envian por outlook (si tiene el parametro)
                    ' Si no se personaliza, se envia como generico
                
                    'LO que habia
                    'If Me.outTipoDocumento = 0 Then
                    '    'MsgBox "Tipo de documento sin definir en el envio.", vbExclamation
                    'Else
                    '    LanzaAbrirOutlook = True
                    'End If
                    LanzaAbrirOutlook = True
                End If
            
                If LanzaAbrirOutlook Then
                
                
                    
                    If InstalacionEsEulerTaxco Then
                        If davidCodtipom <> "" Then
                            If Dir(davidCodtipom, vbDirectory) <> "" Then LanzaVisorMimeDocumento Me.hwnd, davidCodtipom
                            
                        End If
                     End If
                
                    '
                    LanzaProgramaAbrirOutlook
                Else
                    'El que habia
                    frmEMail.Opcion = 0
                    frmEMail.Show vbModal
                End If
            End If
            CadenaDesdeOtroForm = ""
        End If
    End If
    
    Dim HazUnload As Boolean
    HazUnload = False
    'If Not EnvioEMail Then Unload Me
    If Not EnvioEMail Then
        If Opcion = 53 And SoloImprimir Then
            HazUnload = False 'para que no haga la hola
        Else
            HazUnload = True
        End If
    End If
    If HazUnload Then Unload Me
    
End Function


Private Sub Form_Unload(Cancel As Integer)

    NumeroCopias = 0
    outTipoDocumento = 0 'Para restear esta variable
    davidNumalbar = 0 'Log impresion albaranes  tb la reestablezco
    PulsaAceptar = False
    SeleccionaRPTCodigo = 0
    NombreSubRptConta = ""
    NombrePDF = ""
    MostrarTreeDesdeFuera = False
    Titulo = ""

    If EnvioEMail Then Exit Sub
    

    If Me.chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    'Dejo la marca como estaba
    If SoloImprimir Then
        If EstabaMarcado Then chkSoloImprimir.Value = 1
    End If
    
    OperacionesArchivoDefecto
    
    
    

End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

    crear = (Me.chkSoloImprimir.Value = 1)
    'crear = crear And ReestableceSoloImprimir
    If Not crear Then
        Kill App.Path & "\impre.dat"
        Else
            FileCopy App.Path & "\Vacio.dat", App.Path & "\impre.dat"
    End If
ErrOperacionesArchivoDefecto:
        If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub Text1_DblClick()
    Frame2.Tag = Val(Frame2.Tag) + 1
    If Val(Frame2.Tag) > 2 Then
        Frame2.Enabled = True
        chkSoloImprimir.visible = True
    End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub


Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub LanzaProgramaAbrirOutlook()
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String
Dim i As Integer
Dim UpdateSQL As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    UpdateSQL = ""
    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Oferta
        Aux = "OFE" & Me.outClaveNombreArchiv & ".pdf"
    Case 2
        'Fra
         Aux = Me.outClaveNombreArchiv & ".pdf"
    Case 3
         Aux = "PED" & Me.outClaveNombreArchiv & ".pdf"
    Case 4
         Aux = Me.outClaveNombreArchiv & ".pdf"
    Case 5
        Aux = "FPROF" & Me.outClaveNombreArchiv & ".pdf"
    
    Case 6
        'RECORDATORIO OFERTA
        Aux = "REC_OFE_" & Me.outClaveNombreArchiv & ".pdf"
    
    Case 7
        'ALbaran facturado
        Aux = "" & Me.outClaveNombreArchiv & ".pdf"
        
        
    Case 8
        'Confirmacion entraga pedido
        Aux = "Entrega_Ped" & Me.outClaveNombreArchiv & ".pdf"
        
    Case 51
        Aux = "PEDP" & Me.outClaveNombreArchiv & ".pdf"
        
        UpdateSQL = "UPDATE scappr set emailenviado =1 WHERE numpedpr = " & outClaveNombreArchiv
    Case Else
        'GENERICO.  Sin especificar.
        'Con lo cual pondremos Documento
        Aux = Text1.Text
        If Aux <> "" Then
            For i = 1 To Len(Aux)
                Aux = Replace(Aux, Mid("\/:*""?<>|", i, 1), " ")
            Next
            Aux = Aux & ".pdf"
        End If
        If Aux = "" Then Aux = "Documento.pdf"
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
    Case 1
        'ofertas
        If outClaveNombreArchiv = "RTAS" Then
            Aux = "Ofertas"
        Else
            Aux = "Oferta nº" & outClaveNombreArchiv
        End If
    Case 2
        Aux = "Factura nº" & outClaveNombreArchiv
    Case 3
        Aux = "Pedido cliente nº" & outClaveNombreArchiv
    Case 4
        Aux = "Albarán nº" & outClaveNombreArchiv
    Case 5
        Aux = "Factura proforma desde Oferta: " & outClaveNombreArchiv
        
    Case 6
        Aux = "Recordatorio de oferta."
    Case 7
        Aux = "Albaran facturado."
    '--------------------------------------------------
    Case 8
        Aux = "Confirmacion entrega pedido nº: " & outClaveNombreArchiv
    Case 51
        Aux = "Pedido proveedor nº: " & outClaveNombreArchiv
    
    
    Case Else
        'Todos los demas
        Aux = "Informe.  " & Text1.Text
    End Select
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.Path & "\" & vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    If UpdateSQL <> "" Then ejecutar UpdateSQL, True
    
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub


Private Function FijaDireccionEmail() As String
Dim campoemail As String
Dim otromail As String
Dim Maiclie1_Primero As Byte   '0 NO , primero maicli2 ,   1: si       2 NO busques

    FijaDireccionEmail = ""
    
    
    If outTipoDocumento = 0 Then
        '28 Mayo 21
        'GENERICO. NO especifo direccion email
        campoemail = ""
    
    
    Else
        'LO que habia
        If outTipoDocumento < 50 Then
            
            
            Maiclie1_Primero = 1
            
            If outTipoDocumento = 8 Then
                'Confirmacion PEDIDO
                campoemail = DevuelveDesdeBD(conAri, "mailconfir", "scaped", "numpedcl", outClaveNombreArchiv)
                otromail = ""
                If campoemail <> "" Then Maiclie1_Primero = 2
            End If
        
            
            If Maiclie1_Primero < 2 Then  'si es dos es que ya "han fijado " el mail
            
                If outTipoDocumento = 1 Or outTipoDocumento = 2 Or outTipoDocumento = 3 Or outTipoDocumento = 7 Or outTipoDocumento = 8 Then
                    campoemail = "maiclie1"
                    otromail = "maiclie2"
                Else
                    campoemail = "maiclie2"
                    otromail = "maiclie1"
                End If
                campoemail = DevuelveDesdeBD(conAri, campoemail, "sclien", "codclien", Me.outCodigoCliProv, "N", otromail)
                If campoemail = "" Then campoemail = otromail
            End If
            
        Else
            'Para provedores
            If outTipoDocumento = 52 Or outTipoDocumento = 53 Then
                campoemail = "maiprov1"
                otromail = "maiprov2"
            Else
                'outTipoDocumento = 51  LO paso aqui bajo. Ped prov
                campoemail = "maiprov2"
                otromail = "maiprov1"
            End If
            campoemail = DevuelveDesdeBD(conAri, campoemail, "sprove", "codprove", Me.outCodigoCliProv, "N", otromail)
            If campoemail = "" Then campoemail = otromail
            
        End If
    End If
    FijaDireccionEmail = campoemail
End Function


Private Sub CargaComboRPTS()
Dim RN As ADODB.Recordset
Dim C As String

    'Primero METEMOS el rpt por defecto, el dee la scryst
    C = "Informe por  defecto [" & Me.NombreRPT & "]"
    Combo1.AddItem C
    Combo1.ListIndex = 0
    
    C = "Select * from scryst2 where codcryst =" & SeleccionaRPTCodigo & " ORDER BY linea"
    Set RN = New ADODB.Recordset
    RN.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        If Not FrameSelecRPT.visible Then FrameSelecRPT.visible = True
        C = RN!descriprp & " [" & RN!nomcryst & "]"
        Combo1.AddItem C
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub




Private Sub ForzarImpresoraPorDefecto(sNombreImpresora As String)

    On Error GoTo eForzarImpresoraPorDefecto
    Dim nom As String
    Dim bEncontrada As Boolean
    Dim i As Integer
    'Selecciona la impresora para imprimir, si no puede seleccionarla devuelve false
    
    bEncontrada = False
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = sNombreImpresora Then
            bEncontrada = True
            Exit For
        End If
    Next i
    
    If bEncontrada Then
        
        Set Printer = Printers(i)
        
    Else
        
        nom = ""
        i = InStrRev(sNombreImpresora, "\")
        If i > 0 Then
            nom = Mid(sNombreImpresora, i + 1)
            For i = 0 To Printers.Count - 1
                If Printers(i).DeviceName = sNombreImpresora Then
                    bEncontrada = True
                    Exit For
                End If
            Next i
            If bEncontrada Then Set Printer = Printers(i)
            
        
        End If
        
    End If

        
    
    
    Exit Sub
eForzarImpresoraPorDefecto:
    Err.Clear
    ImpresoraPorDefectoAnterior = ""
End Sub
