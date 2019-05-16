VERSION 5.00
Begin VB.Form frmProduVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi form para muchas cosas de produccion"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrCierreOrdenProduccion 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkCierreParcial 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre parcial"
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCierreOrdProd 
         Caption         =   "Cerrar orden"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha cierre"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Cierre orden de producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2640
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmProduVarios.frx":0000
         Top             =   960
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
Dim I As Integer

Private Sub chkCierreParcial_Click()
    Me.txtcantidad.visible = chkCierreParcial.Value = 1
    Me.Label2.visible = chkCierreParcial.Value = 1
    
End Sub

Private Sub cmdCancelar_Click(index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreOrdProd_Click()
    If txtFecha(0).Text = "" Then Exit Sub
    
    If Me.txtcantidad.visible Then
        If Me.txtcantidad.Text = "" Then
            MsgBox "Cierre parcial. Indique cantidad", vbExclamation
            Exit Sub
        End If
    End If
    
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
 Dim I As Integer
    Me.Icon = frmPpal.Icon
    FrCierreOrdenProduccion.visible = False
    limpiar Me
    I = Opcion
    Select Case Opcion
    Case 0, 1
        PonerFrameVisible FrCierreOrdenProduccion
        Me.Caption = "Cierre orden producción"
        lbFec(0).Caption = "Cod:   " & RecuperaValor(Intercambio, 1) & "   " & RecuperaValor(Intercambio, 2) & "   "
        chkCierreParcial.visible = False
        If Opcion = 0 Then
            lbFec(0).Caption = lbFec(0).Caption & "PROD"
            'If vParamAplic.NumeroInstalacion = vbFenollar Then chkCierreParcial.visible = True
        Else
            lbFec(0).Caption = lbFec(0).Caption & "Envasado"
            I = 0
        End If
        
        
                
        
        
    End Select
    
    cmdCancelar(I).Cancel = True
End Sub



Private Sub PonerFrameVisible(ByRef Fr As Frame)

    Fr.visible = True
    Fr.Top = 30
    Fr.Left = 30
    Me.Width = Fr.Width + 180
    Me.Height = Fr.Height + 520
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    txtFecha(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(index As Integer)
    'El index tiene que ser el mismo que el del txtfecha al que acompaña
    Set frmC = New frmCal
    frmC.Fecha = Now
    I = index
    If txtFecha(index).Text <> "" Then frmC.Fecha = CDate(txtFecha(index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub

Private Sub txtcantidad_GotFocus()
    ConseguirFoco txtcantidad, 3
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtcantidad_LostFocus()
   txtcantidad.Text = Trim(txtcantidad.Text)
    If txtcantidad.Text <> "" Then
        If Not PonerFormatoDecimal(txtcantidad, 2) Then
        
            MsgBox "Cantidad incorrecta: " & txtcantidad.Text, vbExclamation
            txtcantidad.Text = ""
            PonerFoco txtcantidad
        End If
    End If
 
End Sub

Private Sub txtFecha_GotFocus(index As Integer)
    ConseguirFoco txtFecha(index), 3
End Sub

Private Sub txtFecha_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(index As Integer)
Dim T As String
    txtFecha(index).Text = Trim(txtFecha(index).Text)
    If txtFecha(index).Text <> "" Then
        T = txtFecha(index).Text
        If EsFechaOK(T) Then
            txtFecha(index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(index).Text, vbExclamation
            txtFecha(index).Text = ""
            PonerFoco txtFecha(index)
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
    If cad = "" Then Err.Raise 513, , "Imposible asignar trabajador conectado"
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.FechaMov = txtFecha(0).Text '
    
   
    vCStock.codArtic = miRsAux!codarti2
    vCStock.codAlmac = CInt(miRsAux!codAlmac)
    CantidadNecesaria = miRsAux!cantidad
    vCStock.cantidad = CSng(CantidadNecesaria)
    vCStock.Importe = 0
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock" & vbCrLf & Err.Description, vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function



