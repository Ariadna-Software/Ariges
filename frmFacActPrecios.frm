VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacActPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Precios"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmFacActPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFam 
      Caption         =   "Familia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   6735
      Begin VB.TextBox txtFam 
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
         Left            =   270
         TabIndex        =   2
         Top             =   315
         Width           =   990
      End
      Begin VB.TextBox txtFam 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Left            =   1620
         TabIndex        =   17
         Top             =   315
         Width           =   4830
      End
      Begin VB.Image ImgFam 
         Height          =   240
         Left            =   1305
         Picture         =   "frmFacActPrecios.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.CheckBox chkCabel 
      Caption         =   "CABEL"
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
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Frame Frameprov 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   360
      TabIndex        =   14
      Top             =   1170
      Width           =   6735
      Begin VB.TextBox txtDescProv 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Left            =   1620
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   315
         Width           =   4830
      End
      Begin VB.TextBox txtProv 
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
         Left            =   270
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   315
         Width           =   990
      End
      Begin VB.Image imgProve 
         Height          =   240
         Left            =   1305
         Picture         =   "frmFacActPrecios.frx":0A0E
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   360
         Width           =   240
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   3600
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   3600
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   3720
      Begin VB.CheckBox chkPreuEsp 
         Caption         =   "Precios especiales"
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
         Top             =   720
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox chkPreuAct 
         Caption         =   "Precios actuales"
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
         Top             =   360
         Value           =   1  'Checked
         Width           =   3330
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label lblProgreso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6105
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblProgreso 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblTitulo 
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   10
      Top             =   120
      Width           =   6330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de cambio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   21
      Left            =   360
      TabIndex        =   8
      Top             =   765
      Width           =   1800
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   2205
      Picture         =   "frmFacActPrecios.frx":1410
      Top             =   765
      Width           =   240
   End
End
Attribute VB_Name = "frmFacActPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Proveedor As Boolean


Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 '%=%=frmComProveedores
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
Dim Sql As String
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
        Sql = "SELECT COUNT(*) FROM slista,sartic WHERE " & cadSel
        totRegPA = TotalRegistros(Sql)
        
        If Not (totRegPA > 0) Then
            If Me.chkPreuEsp.Value = 1 Then
                'comprobar si se actualizar precios especiales
                Sql = "SELECT COUNT(*) FROM sprees,sartic WHERE " & Replace(cadSel, "slista", "sprees")
                totRegPE = TotalRegistros(Sql)
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
        Sql = "SELECT COUNT(*) FROM sprees,sartic WHERE " & Replace(cadSel, "slista", "sprees")
        totRegPE = TotalRegistros(Sql)
        
        If Not ((totRegPE) > 0) And totRegPA = 0 Then
            MsgBox "No hay precios especiales a actualizar para esa fecha.", vbExclamation
            Exit Sub
        End If
    End If
    
    If MsgBox("Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    
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
    Me.Height = 5100
    Frame1.visible = Not Me.Proveedor
    Me.FrameFam.visible = Me.Proveedor
    chkCabel.visible = Me.Proveedor And vParamAplic.NumeroInstalacion = 2
   ' Frameprov.visible = Me.Proveedor
    txtProv.Text = ""
    Me.txtDescProv.Text = ""
    If Me.Proveedor Then
        Me.lblTitulo(3).Caption = "Actualizar precios proveedor"
        Caption = "PROVEEDOR"
    Else
        Caption = "TARIFAS "
        Me.lblTitulo(3).Caption = "Actualizar Tarifas de Precios"
        Me.Frame1.Top = FrameFam.Top
    End If
    
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)

    menErrProceso = CadenaSeleccion

End Sub

Private Sub ImgFam_Click()
    LanzaAyuda False
    If menErrProceso <> "" Then
        txtFam(0).Text = RecuperaValor(menErrProceso, 1)
        txtFam(1).Text = RecuperaValor(menErrProceso, 2)
    End If
    menErrProceso = ""
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

    LanzaAyuda True
    If menErrProceso <> "" Then
        txtProv.Text = RecuperaValor(menErrProceso, 1)
        Me.txtDescProv.Text = RecuperaValor(menErrProceso, 2)
    End If
    menErrProceso = ""
End Sub

Private Sub LanzaAyuda(prov As Boolean)
    menErrProceso = ""
    Set frmProv = New frmBasico2
    If prov Then
        AyudaProveedores frmProv, txtProv
    Else
        AyudaFamilias frmProv, txtFam(0)
    End If
    Set frmProv = Nothing
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        KEYFecha KeyAscii, 0 'fechar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
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
Dim Sql As String
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
    Sql = "SELECT * FROM slista,sartic WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim B As Boolean
Dim margen As Currency
Dim newPrecio As Currency

    On Error GoTo ErrAct
    conn.BeginTrans
    
    
    Set cTar = New CTarifaArt
    B = cTar.LeerDatos(codArt, codLis)
    
    If B Then
        'actualizar la tarifa precios
        B = cTar.ActualizarPrecios(cTar.FechaCambio, cTar.PrecioNuevo, cTar.PrecioCajaNuevo, cadErr, True)
        
        'si tarifa es la de parametros actualizar PVP del articulo
        If B And codLis = vParamAplic.CodTarifa Then
            B = BloquearArticulo(codArt)
            If B Then
                margen = Round2(cTar.MargenComercial / 100, 4)
                newPrecio = Round2((cTar.PrecioNuevo / (margen + 1)), 4)
                B = ActualizarPVPArticulo(codArt, newPrecio)
            End If
        End If
    End If
    Set cTar = Nothing
    
    
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    
    ActualizarTarifa = B
    If Not B And cadErr <> "" Then MsgBox cadErr, vbExclamation
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
Dim Sql As String
    
    On Error GoTo ErrActPVP
    ActualizarPVPArticulo = False
    
    Sql = "UPDATE sartic SET preciove=" & DBSet(newPreu, "N")
    Sql = Sql & " WHERE codartic=" & DBSet(codArt, "T")
    conn.Execute Sql
    
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
Dim Sql As String
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
    Sql = "SELECT * FROM sprees,sartic WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim Sql As String
Dim RS As ADODB.Recordset
Dim NumF As String

    On Error GoTo ErrAct
    
    conn.BeginTrans
    ActualizarPrecioEspec = False
    
    Sql = "SELECT * FROM sprees WHERE codclien=" & codCli & " AND codartic=" & DBSet(codArt, "T")
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        '-- Insertar en el historico spree1
        'numero de linea
        NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", "codartic=" & DBSet(codArt, "T") & " AND codclien=" & codCli)
    
        Sql = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec) "
        Sql = Sql & " VALUES (" & codCli & "," & DBSet(codArt, "T") & "," & NumF & ","
        Sql = Sql & DBSet(RS!fechanue, "F") & "," & DBSet(RS!precioac, "N") & "," & DBSet(DBLet(RS!precioa1, "N"), "N") & "," & DBSet(RS!dtoespec, "N") & ")"
        conn.Execute Sql
        
        
        '-- Actualizar precios actuales con nuevo y resetear valores nuevos
        Sql = "UPDATE sprees SET precioac=" & DBSet(RS!precionu, "N")
        Sql = Sql & "," & " precioa1=" & DBSet(RS!precion1, "N")
        'Si dtosespec nuevo esta vacio(null), como esto es herbequiano, no UPDATEO el dtoespecial
        If Not IsNull(RS!dtoespe1) Then Sql = Sql & ", dtoespec=" & DBSet(RS!dtoespe1, "N")
        Sql = Sql & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo & ", precion1=" & ValorNulo
        Sql = Sql & ", dtoespe1=" & ValorNulo
        Sql = Sql & " WHERE codclien=" & codCli & " and codartic=" & DBSet(codArt, "T")
        conn.Execute Sql
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



Private Sub txtFam_GotFocus(Index As Integer)
    ConseguirFoco txtFam(0), 2
End Sub

Private Sub txtFam_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYBusqueda KeyAscii, 1 'fam
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtFam_LostFocus(Index As Integer)
    If Index = 1 Then Exit Sub
    
    menErrProceso = ""
    txtFam(Index).Text = Trim(txtFam(Index).Text)
    If txtFam(Index).Text <> "" Then
        If Not IsNumeric(txtFam(Index).Text) Then
            MsgBox "Campo numerico", vbExclamation
        Else
            menErrProceso = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFam(Index).Text)
            If menErrProceso = "" Then MsgBox "no existe existe familia", vbExclamation
        End If
        If menErrProceso = "" Then
            txtFam(Index).Text = ""
            PonerFoco txtFam(Index)
        End If
    End If
    Me.txtFam(1).Text = menErrProceso
    menErrProceso = ""
   
End Sub

Private Sub txtProv_GotFocus()
    ConseguirFoco txtProv, 3
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    If Indice = 0 Then
        imgProve_Click
    Else
        ImgFam_Click
    End If
End Sub

Private Sub txtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYBusqueda KeyAscii, 0 'proveedor
    Else
        KEYpress KeyAscii
    End If
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
Dim Sql As String
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
    
    Sql = ""
    cadSel = ""
    If Me.txtProv.Text <> "" Or Me.txtDescProv.Text <> "" Then Sql = "1"
    
    If Me.chkCabel.Value Then
        If Sql <> "" Then cadSel = "No debe indicar proveedor"
        
        'Si indica familia, debe ser cabel
        If Me.txtFam(0).Text <> "" Then
            Sql = DevuelveDesdeBD(conAri, "marcapropia", "sfamia", "codfamia", txtFam(0).Text)
            If Sql = "" Then
                cadSel = "Error leyendo familia"
            Else
                If Sql = "0" Then cadSel = "No es familia CABEL"
            End If
        End If
    Else
        If Sql = "" Then cadSel = "Indique proveedor"
        
        
    End If
    If cadSel <> "" Then
        MsgBox cadSel, vbExclamation
        Exit Sub
    End If
    
        
    
    '--- COMPROBAR Q HAY REGISTROS A PROCESAR
    '------------------------------------------
    
    '- obtener la cadena de seleccion de registros de tarifas de precio q se van
    '    a actualizar: los q cumplan q slista.fechanue <= valor_introducido
    cadSel = "fechanue"
    cadSel = CadenaDesdeHastaBD("", txtCodigo(0).Text, cadSel, "F")
    
    If Not chkCabel.Value = 1 Then cadSel = cadSel & " AND codprove = " & txtProv.Text
    
    cadSel = cadSel & " AND codartic in (select codartic from sartic,sfamia"
    cadSel = cadSel & " WHERE sartic.codfamia=sfamia.codfamia "
    If Me.txtFam(0).Text <> "" Then cadSel = cadSel & " AND sfamia.codfamia =" & txtFam(0).Text
    'CABEL
    If Me.chkCabel.Value Then cadSel = cadSel & " AND marcapropia=1 "
    cadSel = cadSel & " )"
        
        
    
    
    
    
    
    
    
    
    '- comprabar q existen registros para ese criterio de seleccion
    totRegPA = 0
    totRegPE = 0
    If Me.chkPreuAct.Value = 1 Then
        'si marcado actualizar PRECIOS ACTUALES
        Sql = "SELECT COUNT(*) FROM slispr WHERE " & cadSel
        totRegPA = TotalRegistros(Sql)
        
        If totRegPA = 0 Then
                'no hay registros a procesar y fin
                MsgBox "No hay precios a actualizar para con estos valores.", vbExclamation
                Exit Sub
           
        End If
    End If
    
    
    If MsgBox("Continuar con el proceso?  (" & totRegPA & ")", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    
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
    
End Sub



Private Sub ProcesoActualizarPreciosProvee(cadWhere As String, totReg As Long)
'Actualizar los precios Actuales de las Tarifas
'(IN) cadWHERE: cadena seleccion de tarifas a actualizar
'Para cada tarifa a actualizar:
'   - insertar. en historico (slist1) linea con slista.fechanue y con el slista.precioac
'   - actualizar slista con slista.precioac=slista.precionu
'   - si slista.codlista es la tarifa de los parametros de la aplicacion: actualizar PVP del articulo
Dim Sql As String
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
    Sql = "SELECT * FROM slispr WHERE " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim Sql As String
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
    Sql = "INSERT INTO slisp1(codartic,codprove,numlinea,fechacam,precioac) "
    Sql = Sql & " VALUES (" & DBSet(rsa!codArtic, "T") & "," & vCodProve

    'numero de linea
    NumF = SugerirCodigoSiguienteStr("slisp1", "numlinea", "codartic=" & DBSet(rsa!codArtic, "T") & " AND codprove=" & vCodProve)
    Sql = Sql & "," & NumF & "," & DBSet(Me.txtCodigo(0).Text, "F") & "," & DBSet(rsa!precioac, "N") & ")"




    conn.Execute Sql
    

    'Actualizar los precios actuales con valores nuevos
    'y quitar el valor de los precios nuevos y poner a nulos
    '--------------------------------------------------
    Sql = "UPDATE slispr SET precioac=" & DBSet(rsa!precionu, "N")
'    SQL = SQL & "," & " precioa1=" & DBSet(newPrecioA1, "N")
    Sql = Sql & ", " & "precionu=" & ValorNulo & ", fechanue=" & ValorNulo
    Sql = Sql & " WHERE codartic=" & DBSet(rsa!codArtic, "T") & " AND codprove=" & vCodProve
    conn.Execute Sql
    
    
    
    ActualizarPreciosProvee = True
    Exit Function
    
ErrAct:


    MuestraError Err.Number, "Actualizar precios proveedor.", Err.Description
End Function
