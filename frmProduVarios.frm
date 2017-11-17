VERSION 5.00
Begin VB.Form frmProduVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi form para muchas cosas de produccion"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrCierreOrdenProduccion 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdCierreOrdProd 
         Caption         =   "Cerrar orden"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Cierre orden de producción"
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
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2280
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   480
         Picture         =   "frmProduVarios.frx":0000
         Top             =   720
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmProduVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0  .-Cierrer de una orden de produccion
    '1 .-  "             "        envasado
Public Intercambio As String

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim cad As String  'multi proposito
Dim i As Integer

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreOrdProd_Click()
    If txtFecha(0).Text = "" Then Exit Sub
    cad = "¿Seguro que desea cerrar la orden de "
    If Opcion = 0 Then
        cad = cad & "producción"
    Else
        cad = cad & "envasado"
    End If
    cad = cad & RecuperaValor(Intercambio, 1) & " - " & RecuperaValor(Intercambio, 2)
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If CerrarOrdenProduccion(True) Then
        If CerrarOrdenProduccion(False) Then Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
 Dim i As Integer
    Me.Icon = frmPpal.Icon
    FrCierreOrdenProduccion.visible = False
    limpiar Me
    i = Opcion
    Select Case Opcion
    Case 0, 1
        PonerFrameVisible FrCierreOrdenProduccion
        Me.Caption = "Cierre orden producción"
        lbFec(0).Caption = "Cod:   " & RecuperaValor(Intercambio, 1) & "   " & RecuperaValor(Intercambio, 2) & "   "
        If Opcion = 0 Then
            lbFec(0).Caption = lbFec(0).Caption & "PROD"
        Else
            lbFec(0).Caption = lbFec(0).Caption & "Envasado"
            i = 0
        End If
        
    End Select
    
    cmdCancelar(i).Cancel = True
End Sub



Private Sub PonerFrameVisible(ByRef Fr As Frame)

    Fr.visible = True
    Fr.Top = 30
    Fr.Left = 30
    Me.Width = Fr.Width + 180
    Me.Height = Fr.Height + 520
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    txtFecha(i).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'El index tiene que ser el mismo que el del txtfecha al que acompaña
    Set frmC = New frmCal
    frmC.Fecha = Now
    i = Index
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    
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

Private Function CerrarOrdenProduccion(SoloComprobar As Boolean) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim tabla As String
    
    If Opcion = 0 Then
        tabla = "sliordpr"
    Else
        tabla = "slienvpr"
    End If

    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    CerrarOrdenProduccion = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New CStock
    'Veamos las sub lineas  si tienen stock. Antes comprobabamos cantidad x sarti1.cntidad
    'Cad = "select codarti1,codalmac,sarti1.cantidad multiplicador,sum(sliordpr.cantidad) cantilinea from sliordpr,sarti1 where "
    'Cad = Cad & " sliordpr.codartic=sarti1.codartic and  codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2,3"
    'AHora hay una tabla para los componentes
'    Cad = "select codarti2,sliordpr.codalmac,sliordpr2.cantidad cantilinea from sliordpr,sliordpr2 where"
'    Cad = Cad & " sliordpr.codartic=sliordpr2.codartic and sliordpr.codalmac=sliordpr2.codalmac and"
'    Cad = Cad & " sliordpr.codigo=1 group by 1,2"
'
    cad = "select * from " & tabla & "2 where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    If Not SoloComprobar Then conn.BeginTrans
        
    While Not miRsAux.EOF
        b = False
        If InicializarCStock(vCStock, "S") Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    b = vCStock.MoverStock(False, False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock(False, True)
                End If
            Else
                b = True
            End If
            
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = "select codartic codarti2,codalmac,sum(" & tabla & ".cantidad) cantidad from " & tabla & " where "
    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    While Not miRsAux.EOF
        b = False
        If InicializarCStock(vCStock, "E") Then   'Las lineas son de netrada
        
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    b = vCStock.MoverStock(False, False, True)
                Else
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        conn.CommitTrans
        If Opcion = 0 Then
            cad = "sordprod"
        Else
            cad = "senvprod"
        End If
        cad = "UPDATE " & cad & "  set fecproduccion = " & DBSet(txtFecha(0).Text, "F")
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        conn.Execute cad
    End If
    
    CerrarOrdenProduccion = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'No le paso el recodset pq es mirsaux que es comun
Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String) As Boolean
Dim CantidadNecesaria As Currency
    On Error Resume Next

    vCStock.tipoMov = TipoM
    If Opcion = 0 Then
        vCStock.DetaMov = "PRO"
    Else
        vCStock.DetaMov = "PRE"
    End If
    vCStock.Trabajador = PonerTrabajadorConectado(cad)
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.FechaMov = txtFecha(0).Text '
    
   
    vCStock.codArtic = miRsAux!codarti2
    vCStock.codAlmac = CInt(miRsAux!codAlmac)
    CantidadNecesaria = miRsAux!cantidad
    vCStock.cantidad = CSng(CantidadNecesaria)
    vCStock.Importe = 0
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function



