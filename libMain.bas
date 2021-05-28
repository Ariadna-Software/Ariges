Attribute VB_Name = "libMain"

Public Sub Main()
Dim T1 As Single
Dim LanzaWhose As Boolean


       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       
       frmIdentifica.Show vbModal
               

               
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set conn = Nothing
            End
       End If
       
       
    
      
       If vUsu.Skin >= 0 Then
           
            
            

            
            'NUEVO MENU
            Load frmPpalN
            
            DoEvent2
            frmPpalN.Show vbModal
            
            
            
        Else
       
           
           
           
           
           
           
           'No deberiamos utiliza gotos, pero me ahorro tanta faena y no se hacerlo de otra forma
AQUI:
           CadenaDesdeOtroForm = ""
           frmLogin.Show vbModal
           
            If CadenaDesdeOtroForm = "" Then
                'No ha seleccionado ninguna empresa
                Set conn = Nothing
                End
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
    
        
            'Mayo 2011
            'Borraremos de todos los arigiess que hayan la tabla zbloqueos where codusu
            If Not App.PrevInstance Then BorrarEnZbloqueos
        
    
    '        LeerEmpresa 'Carga los Datos de la empresa
            'Carga los Datos Básicos de la empresa
            LeerDatosEmpresa
            
            'Cerramos la conexion con BD: Usuarios
            conn.Close
    
            'Abre la conexión a BDatos:Ariges
            If AbrirConexion() = False Then
                MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
                End
            Else
                'Carga Parametros Generales y Contables de la empresa
                LeerParametros
            End If
                    
            'Abrir conexión a la BDatos de Contabilidad para acceder a
            'Tablas: Cuentas, Tipos IVA
            If AbrirConexionConta(False) = False Then
                MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
                End
            End If
            
            
            'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
            LeerNivelesEmpresa
            
    '        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
    '        GestionaPC
            
            'Otras acciones
            OtrasAcciones
             
            LanzaWhose = False
            If LCase(vUsu.Login) <> "root" Then
                If vParamAplic.QueEmpresaEs = 1 Then LanzaWhose = True
            End If
            
            If LanzaWhose Then
                'Leere este parametro
                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "PathDocs", "sparamwhose", "1", "1")
                If CadenaDesdeOtroForm = "" Then
                    MsgBox "ERROR obteniendo PATH documentos", vbCritical
                    End
                End If
                
                vParamAplic.PathDocsWHOSE = CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
                
                frmPPalWhose.Show vbModal
            Else
    '            frmPpal.Show
                
                If vParamAplic.QueEmpresaEs = 2 Then
                    
                    frmPpalGessocial.Show vbModal, frmPpalOld
                    If CadenaDesdeOtroForm = "" Then
                        GoTo AQUI
                    Else
                        Exit Sub
                    End If
                End If
                frmPpalOld.Show
            End If

     End If

End Sub




Private Sub BorrarEnZbloqueos()
Dim Cad As String
Dim RS As ADODB.Recordset
    On Error GoTo EBorrarEnZbloqueos
    
    Cad = "Select ariges from empresasariges"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Cad = "DELETE FROM " & RS.Fields(0) & ".zbloqueos where codusu = " & vUsu.Codigo
        ejecutar Cad, True
        RS.MoveNext
    Wend
    RS.Close
    
EBorrarEnZbloqueos:
   Err.Clear
   Set RS = Nothing
End Sub

Public Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    
    
    
    'Por si paraemtrizamos la ampliacion
    FormatoPrecio = "###,##0.0000"  'Decimal(10,4)
    FormatoPrecio2 = "###,##0." & String(PrecioDecimales, "0") 'Decimal(10,4)
    
    '=%=%
    teclaBuscar = 43
    
    'Por si acasomcambaimos la aplicacion los numeros de decimales
    'ANTES
    'FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    'FormatoCantidad2 = "###,##0.00"   'Decimal(8,2)
    'Ahora
    FormatoCantidad = "##,###,##0." & String(NumeroDeDecimales, "0")
    FormatoCantidad2 = "###,##0." & String(NumeroDeDecimales, "0")
    
    FormatoDescuento = "#0.00" 'Decima(4,2)
    FormatoKms = "#,##0.00##" 'Decimal(8,4)
    FormatoPorcen = "##0.00" 'Decima(5,2)
    
    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    conn.Execute "Delete from zbloqueos " & CadenaDesdeOtroForm
    
    
    
    
    'If vParamAplic.NumeroInstalacion <> vbFenollar Then HaMostradoCanal2_El_B = True    borrar
    
    
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
End Sub



Public Sub AbrirGeolocalizacion(ByVal Coordendadas As String)

    Coordendadas = "https://www.google.com/maps/?q=" & Coordendadas
    LanzaVisorMimeDocumento frmPpal.hwnd, Coordendadas
    
End Sub




'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

'           SII.  4. No seguir, pero no dar msgbox
'                 5. NO seguir, pero dar msgbox con MensajeFechaOkConta
Public Function EsFechaOKConta(Fecha As Date, DejarContinuarSII As Boolean) As Byte
Dim F2 As Date
    
    
    If vEmpresa.FechaIni > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresa.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If
    If EsFechaOKConta = 0 Then
        'Si tiene SII
        If vParamAplic.ContabilidadNueva Then
            If vParamAplic.SII_Tiene Then
                
                If Fecha < UltimaFechaCorrectaSII(vParamAplic.Sii_Dias, Now) Then
                    MensajeFechaOkConta = "Fecha fuera de periodo de comunicación SII."
                    'LLEVA SII y han trascurrido los dias
                    If vUsu.Nivel = 0 Then
                        If MsgBox(MensajeFechaOkConta & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            EsFechaOKConta = 4
                        End If
                        
                    Else
                        'NO tienen nivel
                        EsFechaOKConta = 5
                    End If
                End If
            End If
        End If
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If
End Function


'Para proveedores la necesito
Public Function EsFechaOKConta_SinSII(Fecha As Date, DejarContinuarSII As Boolean) As Byte
Dim F2 As Date
    
    
    If vEmpresa.FechaIni > Fecha Then
        EsFechaOKConta_SinSII = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresa.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta_SinSII = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta_SinSII = 0
        End If
    End If
    If EsFechaOKConta_SinSII = 0 Then
       
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If
End Function











Public Sub AbrirListado(numero As Integer)
    Screen.MousePointer = vbHourglass
    
    frmListado.OpcionListado = numero
    frmListado.Show vbModal
    
    Screen.MousePointer = vbDefault
End Sub
Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.OpcionListado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Sub AbrirListadoPed(numero As Integer)
'Abre el Form con los listados de Pedidos
    Screen.MousePointer = vbHourglass
    frmListadoPed.OpcionListado = numero
    frmListadoPed.Show vbModal
    Screen.MousePointer = vbDefault
    
    
End Sub

Public Sub MostrarAvisosPantalla(LosErrores As String)
    frmMensajes.vCampos = LosErrores
    frmMensajes.OpcionMensaje = 13
    frmMensajes.Show vbModal
End Sub



 









