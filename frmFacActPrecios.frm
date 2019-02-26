VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacActPrecios2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Precios"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmFacActPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCabel 
      Caption         =   "CABEL"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Frame Frameprov 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtDescProv 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtProv 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgProve 
         Height          =   240
         Left            =   1320
         Picture         =   "frmFacActPrecios.frx":000C
         Top             =   360
         Width           =   240
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   3135
      Begin VB.CheckBox chkPreuEsp 
         Caption         =   "Precios especiales"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkPreuAct 
         Caption         =   "Precios actuales"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblProgreso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   6105
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblProgreso 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblTitulo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de cambio"
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
      Index           =   21
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   1410
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1920
      Picture         =   "frmFacActPrecios.frx":010E
      Top             =   720
      Width           =   240
   End
End
Attribute VB_Name = "frmFacActPrecios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Proveedor As Boolean


Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmProv As frmComProveedores
Attribute frmProv.VB_VarHelpID = -1

Private menErrProceso As String 'mensaje final del proceso actualizacion de precios

Private Sub chkCabel_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub chkPreuAct_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub AceptarClientes()
'actualizar los nuevos precios actuales  y/o especiales
Dim cadSel As String
Dim SQL As String
Dim totRegPA As Long 'total registros a cambiar de precios actuales
Dim totRegPE As Long 'total registros a cambia de precios especiales


    '--- COMPROBACIONES DE DATOS
    '-----------------------------
    
    '- comprobar q se ha seleccionado fecha de cambio
    If txtCodigo(0).Text = "" Then
        MsgBox "El campo fecha de cambio debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    '- comprobar que es una fecha valida
    PonerFormatoFecha txtCodigo(0)
    If txtCodigo(0).Text = "" Then
        Exit Sub
    End If
    
    '- comprobar q se ha seleccionado al menos un check
    If Me.chkPreuAct.Value <> 1 And Me.chkPreuEsp <> 1 Then
        MsgBox "Debe seleccionar al menos un precio para actualizar.", vbExclamation
        Exit Sub
    End If
    
    
    
    '--- COMPROBAR Q HAY REGISTROS A PROCESAR
    '------------------------------------------
    
    '- obtener la cadena de seleccion de registros de tarifas de precio q se van
    '    a actualizar: los q cumplan q slista.fechanue <= valor_introducido
    cadSel = "fechanue"
    cadSel = CadenaDesdeHastaBD("", txtCodigo(0).Text, cadSel, "F")
    cadSel = " slista.codartic =sartic.codartic AND " & cadSel
    If Me.txtProv.Text <> "" Then cadSel = cadSel & " AND sartic.codprove = " & txtProv.Text
    
    '- comprabar q existen registros para ese criterio de seleccion
    totRegPA = 0
    totRegPE = 0
    If Me.chkPreuAct.Value = 1 Then
        'si marcado actualizar PRECIOS ACTUALES
        SQL = "SELECT COUNT(*) FROM slista,sartic WHERE " & cadSel
        totRegPA = TotalRegistros(SQL)
        
        If Not (totRegPA > 0) Then
            If Me.chkPreuEsp.Value = 1 Then
                'comprobar si se actualizar precios especiales
                SQL = "SELECT COUNT(*) FROM sprees,sartic WHERE " & Replace(cadSel, "slista", "sprees")
                totRegPE = TotalRegistros(SQL)
                If Not (totRegPE > 0) Then
                    MsgBox "No hay tarifas de precios ni precios especiales a actualizar para esa fecha.", vbExclamation
                    Exit Sub
                End If
            ElseIf Me.chkPreuEsp.Value <> 1 Then
                'no hay registros a procesar y fin
                MsgBox "No hay tarifas de precios a actualizar para esa fecha.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    If Me.chkPreuEsp.Value = 1 Then
        'comprobar si se actualizar PRECIOS ESPECIALES
        SQL = "SELECT COUNT(*) FROM sprees,sartic WHERE " & Replace(cadSel, "slista", "sprees")
        totRegPE = TotalRegistros(SQL)
        
        If Not ((totRegPE) > 0) And totRegPA = 0 Then
            MsgBox "No hay precios especiales a actualizar para esa fecha.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    '--- ACTUALIZAR LOS PRECIOS
    '---------------------------------
    menErrProceso = ""
    
    '-- Bloquear para que nadie mas pueda actualizar precios
    DesBloqueoManual ("ACTPRE")
    If Not BloqueoManual("ACTPRE", "1") Then
        MsgBox "Hay otro usuario actualizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'PRECIOS ACTUALES
    If Me.chkPreuAct.Value = 1 And totRegPA > 0 Then
        '-- bloquear los registros a actualizar
        If Not BloqueaRegistro("slista,sartic", " not isnull(fechanue) and " & cadSel) Then
            MsgBox "No se ha podido actualizar precios actuales.", vbExclamation
        Else
            '-- proceso actualizar precios actuales
            Screen.MousePointer = vbHourglass
            ProcesoActualizarPrecios_Actuales cadSel, totRegPA
            Screen.MousePointer = vbDefault
        End If
        TerminaBloquear
    End If
    
    
    'PRECIOS ESPECIALES
    If Me.chkPreuEsp.Value = 1 And totRegPE > 0 Then
        '-- bloquear los registros a actualizar
        If Not BloqueaRegistro("sprees,sartic", " not isnull(fechanue) and " & Replace(cadSel, "slista", "sprees")) Then
            MsgBox "No se ha podido actualizar precios especiales.", vbExclamation
        Else
            '-- proceso actualizar precios especiales
            Screen.MousePointer = vbHourglass
            ProcesoActualizarPrecios_Especiales Replace(cadSel, "slista", "sprees"), totRegPE
            Screen.MousePointer = vbDefault
        End If
        TerminaBloquear
    End If
    
    DesBloqueoManual ("ACTPRE")
    
    If menErrProceso <> "" Then MsgBox menErrProceso, vbInformation
    cmdCancel_Click



End Sub



Private Sub chkPreuEsp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    If MsgBox("Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If Me.Proveedor Then
        AceptarProveedores
    Else
        AceptarClientes
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Me.ProgressBar1.visible = False
    Me.lblProgreso(0).visible = False
    Me.lblProgreso(1).visible = False
    Me.Height = 4100
    Frame1.visible = Not Me.Proveedor
    chkCabel.visible = Me.Proveedor And vParamAplic.NumeroInstalacion = 2
   ' Frameprov.visible = Me.Proveedor
    txtProv.Text = ""
    Me.txtDescProv.Text = ""
    If Me.Proveedor Then
        Me.lblTitulo(3).Caption = "Actualizar precios proveedor"
        Caption = "PROVEEDORR"

    Else
        Caption = "TARIFAS "
        Me.lblTitulo(3).Caption = "Actualizar Tarifas de Precios"
    End If
    
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtProv.Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescProv.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgFecha(0).Tag = Index
    Set frmF = New frmCal
    frmF.Fecha = Now
    
    PonerFormatoFecha txtCodigo(Index)
    If txtCodigo(Index).Text <> "" Then frmF.Fecha = CDate(txtCodigo(Index).Text)
   
    Screen.MousePointer = vbDefault
    frmF.Show vbModal
    Set frmF = Nothing
    PonerFoco txtCodigo(Index)
End Sub





Private Sub imgProve_Click()
    Set frmProv = New frmComProveedores
    frmProv.DatosADevolverBusqueda = "0"
    frmProv.Show vbModal
    Set frmProv = Nothing
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, 2, Cerrar
    If Cerrar Then Unload Me
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Index = 0 Then 'fecha de cambio
        If txtCodigo(Index).Text <> "" Then
           PonerFormatoFecha txtCodigo(Index)
        End If
    End If
End Sub




Private Sub ProcesoActualizarPrecios_Actuales(cadWhere As String, totReg As Long)
'Actualizar los precios Actuales de las Tarifas
'(IN) cadWHERE: cadena seleccion de tarifas a actualizar
'Para cada tarifa a actualizar:
'   - insertar. en historico (slist1) linea con slista.fechanue y con el slista.precioac
'   - actualizar slista con slista.precioac=slista.precionu
'   - si slista.codlista es la tarifa de los parametros de la aplicacion: actualizar PVP del articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim hayErr As Boolean

    On Error GoTo ErrActPreu
   
    '-- iniciar la barra de progreso
    Me.Height = 4600
    Me.lblProgreso(0).Caption = "Actualizando precios actuales."
    Me.lblProgreso(0).visible = True
    Me.lblProgreso(1).visible = True
    CargarProgresNew Me.ProgressBar1, 100
    i = 0
    Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
    Me.ProgressBar1.visible = True
    
    
    '-- seleccionar todos los registros actuales a procesar
    SQL = "SELECT * FROM slista,sartic WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa a cambiar
    hayErr = False
    While Not RS.EOF
        '-- actualizar tarifas precios y PVP si corresponde
        If Not ActualizarTarifa(DBLet(RS!codArtic, "T"), DBLet(RS!codlista, "N")) Then
            hayErr = True
        End If
        
        '-- actualizar la progress bar
'        IncrementarProgresNew Me.ProgressBar1, 1
        i = i + 1
        Me.ProgressBar1.Value = CInt((i * 100) / totReg)
        Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
        Me.lblProgreso(0).Caption = "Actualizando precios actuales.     (" & i & " de " & totReg & ")"
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    
    Screen.MousePointer = vbDefault
    If Not hayErr Then
        Me.lblProgreso(0).Caption = "Proceso finalizado correctamente.     (" & i & " de " & totReg & ")"
'        MsgBox "Proceso actualización precios actuales finalizado correctamente.", vbInformation
        menErrProceso = "Proceso actualización precios actuales finalizado correctamente." & vbCrLf
    Else
        Me.lblProgreso(0).Caption = "Proceso finalizado con errores.     (" & i & " de " & totReg & ")"
        'MsgBox "Algunos precios actuales no se actualizaron correctamente.", vbExclamation
        menErrProceso = "Algunos precios actuales no se actualizaron correctamente." & vbCrLf
    End If
    Espera 0.2
    
    Exit Sub
    
ErrActPreu:
    MuestraError Err.Number, "Actualizar precios actuales.", Err.Description
End Sub




Private Function ActualizarTarifa(codArt As String, codLis As Integer) As Boolean
Dim cadErr As String
Dim cTar As CTarifaArt
Dim b As Boolean
Dim margen As Currency
Dim newPrecio As Currency

    On Error GoTo ErrAct
    conn.BeginTrans
    
    
    Set cTar = New CTarifaArt
    b = cTar.LeerDatos(codArt, codLis)
    
    If b Then
        'actualizar la tarifa precios
        b = cTar.ActualizarPrecios(cTar.FechaCambio, cTar.PrecioNuevo, cTar.PrecioCajaNuevo, cadErr, True)
        
        'si tarifa es la de parametros actualizar PVP del articulo
        If b And codLis = vParamAplic.CodTarifa Then
            b = BloquearArticulo(codArt)
            If b Then
                margen = Round2(cTar.MargenComercial / 100, 4)
                newPrecio = Round2((cTar.PrecioNuevo / (margen + 1)), 4)
                b = ActualizarPVPArticulo(codArt, newPrecio)
            End If
        End If
    End If
    Set cTar = Nothing
    
    
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    
    ActualizarTarifa = b
    If Not b And cadErr <> "" Then MsgBox cadErr, vbExclamation
    Exit Function
    
ErrAct:
    conn.RollbackTrans
    MuestraError Err.Number, "Actualizar precio actual tarifa.", Err.Description
End Function



Private Function BloquearArticulo(codArt As String) As Boolean
Dim cadWhere As String

    cadWhere = "codartic=" & DBSet(codArt, "T")
    BloquearArticulo = BloqueaRegistro("sartic", cadWhere)
End Function




Private Function ActualizarPVPArticulo(codArt As String, newPreu As Currency) As Boolean
Dim SQL As String
    
    On Error GoTo ErrActPVP
    ActualizarPVPArticulo = False
    
    SQL = "UPDATE sartic SET preciove=" & DBSet(newPreu, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(codArt, "T")
    conn.Execute SQL
    
    ActualizarPVPArticulo = True
    Exit Function
    
ErrActPVP:
    ActualizarPVPArticulo = False
    MuestraError Err.Number, "Actualizar precio PVP del articulo.", Err.Description
End Function







Private Sub ProcesoActualizarPrecios_Especiales(cadWhere As String, totReg As Long)
'Actualizar los precios especiales de las Tarifas
'(IN) cadWHERE: cadena seleccion de precios a actualizar
'Para cada precio especial a actualizar:
'   - insertar. en historico (spree1) linea con sprees.fechanue y con el sprees.precioac
'   - actualizar sprees con sprees.precioac=sprees.precionu
'   - poner a nulos los valores nuevos
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim hayErr As Boolean

    On Error GoTo ErrActPreu
    
    '-- iniciar la barra de progreso
    Me.Height = 4600
    Me.lblProgreso(0).Caption = "Actualizando precios especiales."
    Me.lblProgreso(0).visible = True
    Me.lblProgreso(1).visible = True
    CargarProgresNew Me.ProgressBar1, 100 'CInt(totReg)
    i = 0
    Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
    Me.ProgressBar1.visible = True
    
    
    '-- seleccionar todos los registros actuales a procesar
    SQL = "SELECT * FROM sprees,sartic WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada precio especial a cambiar
    hayErr = False
    While Not RS.EOF
        '-- actualizar precios especiales
        If Not ActualizarPrecioEspec(RS!codClien, RS!codArtic) Then
            'procesar errores!!!!!!!!!!!
            hayErr = True
        End If
        
        '-- actualizar la progress bar
'        IncrementarProgresNew Me.ProgressBar1, 1
        i = i + 1
        Me.ProgressBar1.Value = CInt((i * 100) / totReg)
        Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
        Me.lblProgreso(0).Caption = "Actualizando precios especiales.     (" & i & " de " & totReg & ")"
        If (i Mod 30) = 0 Then DoEvents
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    Screen.MousePointer = vbDefault
    If Not hayErr Then
        Me.lblProgreso(0).Caption = "Proceso finalizado correctamente.     (" & i & " de " & totReg & ")"
        'MsgBox "Proceso actualización precios especiales finalizado correctamente.", vbInformation
        menErrProceso = menErrProceso & "Proceso actualización precios especiales finalizado correctamente."
    Else
        Me.lblProgreso(0).Caption = "Proceso finalizado con errores.     (" & i & " de " & totReg & ")"
        'MsgBox "Algunos precios especiales no se actualizaron correctamente.", vbExclamation
         menErrProceso = menErrProceso & "Algunos precios especiales no se actualizaron correctamente."
    End If
    
    Exit Sub
    
ErrActPreu:
    MuestraError Err.Number, "Actualizar precios especiales.", Err.Description
End Sub




Private Function ActualizarPrecioEspec(codCli As Long, codArt As String) As Boolean
'actualizar precio especial
Dim SQL As String
Dim RS As ADODB.Recordset
Dim NumF As String

    On Error GoTo ErrAct
    
    conn.BeginTrans
    ActualizarPrecioEspec = False
    
    SQL = "SELECT * FROM sprees WHERE codclien=" & codCli & " AND codartic=" & DBSet(codArt, "T")
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        '-- Insertar en el historico spree1
        'numero de linea
        NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", "codartic=" & DBSet(codArt, "T") & " AND codclien=" & codCli)
    
        SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec) "
        SQL = SQL & " VALUES (" & codCli & "," & DBSet(codArt, "T") & "," & NumF & ","
        SQL = SQL & DBSet(RS!fechanue, "F") & "," & DBSet(RS!precioac, "N") & "," & DBSet(DBLet(RS!precioa1, "N"), "N") & "," & DBSet(RS!dtoespec, "N") & ")"
        conn.Execute SQL
        
        
        '-- Actualizar precios actuales con nuevo y resetear valores nuevos
        SQL = "UPDATE sprees SET precioac=" & DBSet(RS!precionu, "N")
        SQL = SQL & "," & " precioa1=" & DBSet(RS!precion1, "N")
        'Si dtosespec nuevo esta vacio(null), como esto es herbequiano, no UPDATEO el dtoespecial
        If Not IsNull(RS!dtoespe1) Then SQL = SQL & ", dtoespec=" & DBSet(RS!dtoespe1, "N")
        SQL = SQL & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo & ", precion1=" & ValorNulo
        SQL = SQL & ", dtoespe1=" & ValorNulo
        SQL = SQL & " WHERE codclien=" & codCli & " and codartic=" & DBSet(codArt, "T")
        conn.Execute SQL
    End If
    RS.Close
    Set RS = Nothing


    conn.CommitTrans
    ActualizarPrecioEspec = True
    Exit Function
    
ErrAct:
    ActualizarPrecioEspec = False
    conn.RollbackTrans
    MuestraError Err.Number, "Actualizar precio especial.", Err.Description
End Function



Private Sub txtProv_GotFocus()
    ConseguirFoco txtProv, 3
End Sub

Private Sub txtProv_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtProv_LostFocus()
    menErrProceso = ""
    txtProv.Text = Trim(txtProv.Text)
    If txtProv.Text <> "" Then
        If Not IsNumeric(txtProv.Text) Then
            MsgBox "Campo numerico", vbExclamation
        Else
            menErrProceso = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtProv.Text)
            If menErrProceso = "" Then MsgBox "no existe existe Proveedor", vbExclamation
        End If
        If menErrProceso = "" Then
            txtProv.Text = ""
            PonerFoco txtProv
        End If
    End If
    Me.txtDescProv.Text = menErrProceso
    menErrProceso = ""
    
End Sub



Private Sub AceptarProveedores()
'actualizar los nuevos precios actuales  y/o especiales
Dim cadSel As String
Dim SQL As String
Dim totRegPA As Long 'total registros a cambiar de precios actuales
Dim totRegPE As Long 'total registros a cambia de precios especiales



    If txtCodigo(0).Text = "" Then
        MsgBox "El campo fecha de cambio debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    '- comprobar que es una fecha valida
    PonerFormatoFecha txtCodigo(0)
    If txtCodigo(0).Text = "" Then
        Exit Sub
    End If
    
    SQL = ""
    cadSel = ""
    If Me.txtProv.Text = "" Or Me.txtDescProv.Text = "" Then SQL = "1"
    
    If Me.chkCabel.Value Then
        If SQL <> "" Then cadSel = "No debe indicar proveedor"
    Else
        If SQL = "" Then cadSel = "Indique proveedor"
    End If
    
    
    
    '--- COMPROBAR Q HAY REGISTROS A PROCESAR
    '------------------------------------------
    
    '- obtener la cadena de seleccion de registros de tarifas de precio q se van
    '    a actualizar: los q cumplan q slista.fechanue <= valor_introducido
    cadSel = "fechanue"
    cadSel = CadenaDesdeHastaBD("", txtCodigo(0).Text, cadSel, "F")
    If Me.chkCabel.Value Then
        'CABEL
        cadSel = cadSel & " AND codartic in (select codartic from sartic,sfamia"
        cadSel = cadSel & " WHERE sartic.codfamia=sfamia.codfamia and marcapropia=1 )"
    Else
        cadSel = cadSel & " AND codprove = " & txtProv.Text
    End If
    
    '- comprabar q existen registros para ese criterio de seleccion
    totRegPA = 0
    totRegPE = 0
    If Me.chkPreuAct.Value = 1 Then
        'si marcado actualizar PRECIOS ACTUALES
        SQL = "SELECT COUNT(*) FROM slispr WHERE " & cadSel
        totRegPA = TotalRegistros(SQL)
        
        If totRegPA = 0 Then
                'no hay registros a procesar y fin
                MsgBox "No hay precios a actualizar para ese proveedor.", vbExclamation
                Exit Sub
           
        End If
    End If
    
    
    
    '--- ACTUALIZAR LOS PRECIOS
    '---------------------------------
    menErrProceso = ""
    
    '-- Bloquear para que nadie mas pueda actualizar precios
    DesBloqueoManual ("ACTPRE")
    If Not BloqueoManual("ACTPRE", "1") Then
        MsgBox "Hay otro usuario actualizando precio.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'PRECIOS ACTUALES
            '-- proceso actualizar precios actuales
    Screen.MousePointer = vbHourglass
    ProcesoActualizarPreciosProvee cadSel, totRegPA
    Screen.MousePointer = vbDefault

    
    
    DesBloqueoManual ("ACTPRE")
    
    If menErrProceso <> "" Then MsgBox menErrProceso, vbInformation
    cmdCancel_Click
End Sub



Private Sub ProcesoActualizarPreciosProvee(cadWhere As String, totReg As Long)
'Actualizar los precios Actuales de las Tarifas
'(IN) cadWHERE: cadena seleccion de tarifas a actualizar
'Para cada tarifa a actualizar:
'   - insertar. en historico (slist1) linea con slista.fechanue y con el slista.precioac
'   - actualizar slista con slista.precioac=slista.precionu
'   - si slista.codlista es la tarifa de los parametros de la aplicacion: actualizar PVP del articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Long
Dim hayErr As Boolean

    On Error GoTo ErrActPreu
   
    '-- iniciar la barra de progreso
    Me.Height = 4600
    Me.lblProgreso(0).Caption = "Actualizando precios actuales."
    Me.lblProgreso(0).visible = True
    Me.lblProgreso(1).visible = True
    CargarProgresNew Me.ProgressBar1, 100
    i = 0
    Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
    Me.ProgressBar1.visible = True
    
    
    '-- seleccionar todos los registros actuales a procesar
    SQL = "SELECT * FROM slispr WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa a cambiar
    hayErr = False
    While Not RS.EOF
        '-- actualizar tarifas precios y PVP si corresponde
        If Not ActualizarPreciosProvee(RS) Then hayErr = True
        
        
        '-- actualizar la progress bar
'        IncrementarProgresNew Me.ProgressBar1, 1
        i = i + 1
        Me.ProgressBar1.Value = CInt((i * 100) / totReg)
        Me.lblProgreso(1).Caption = CLng((i * 100) / totReg) & " %"
        Me.lblProgreso(0).Caption = "Actualizando precios actuales.     (" & i & " de " & totReg & ")"
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    
    Screen.MousePointer = vbDefault
    If Not hayErr Then
        Me.lblProgreso(0).Caption = "Proceso finalizado correctamente.     (" & i & " de " & totReg & ")"
'        MsgBox "Proceso actualización precios actuales finalizado correctamente.", vbInformation
        menErrProceso = "Proceso actualización precios actuales finalizado correctamente." & vbCrLf
    Else
        Me.lblProgreso(0).Caption = "Proceso finalizado con errores.     (" & i & " de " & totReg & ")"
        'MsgBox "Algunos precios actuales no se actualizaron correctamente.", vbExclamation
        menErrProceso = "Algunos precios actuales no se actualizaron correctamente." & vbCrLf
    End If
    Espera 0.2
    
    Exit Sub
    
ErrActPreu:
    MuestraError Err.Number, "Actualizar precios actuales.", Err.Description
End Sub





Public Function ActualizarPreciosProvee(ByRef rsa As ADODB.Recordset) As Boolean
Dim SQL As String
Dim NumF As String
Dim vCodProve As Long

    On Error GoTo ErrAct

    ActualizarPreciosProvee = False
    
    If Me.chkCabel.Value Then
        vCodProve = rsa!Codprove
    Else
        vCodProve = txtProv.Text
    End If
    
    
    'Mover los precios actuales al histórcio slist1
    '------------------------------------------------
    SQL = "INSERT INTO slisp1(codartic,codprove,numlinea,fechacam,precioac) "
    SQL = SQL & " VALUES (" & DBSet(rsa!codArtic, "T") & "," & vCodProve

    'numero de linea
    NumF = SugerirCodigoSiguienteStr("slisp1", "numlinea", "codartic=" & DBSet(rsa!codArtic, "T") & " AND codprove=" & vCodProve)
    SQL = SQL & "," & NumF & "," & DBSet(Me.txtCodigo(0).Text, "F") & "," & DBSet(rsa!precioac, "N") & ")"




    conn.Execute SQL
    

    'Actualizar los precios actuales con valores nuevos
    'y quitar el valor de los precios nuevos y poner a nulos
    '--------------------------------------------------
    SQL = "UPDATE slispr SET precioac=" & DBSet(rsa!precionu, "N")
'    SQL = SQL & "," & " precioa1=" & DBSet(newPrecioA1, "N")
    SQL = SQL & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo
    SQL = SQL & " WHERE codartic=" & DBSet(rsa!codArtic, "T") & " AND codprove=" & vCodProve
    conn.Execute SQL
    
    
    
    ActualizarPreciosProvee = True
    Exit Function
    
ErrAct:


    MuestraError Err.Number, "Actualizar precios proveedor.", Err.Description
End Function
