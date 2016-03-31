VERSION 5.00
Begin VB.Form frmPPalWhose 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WHOSE "
   ClientHeight    =   8445
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12810
   Icon            =   "frmPPalWhose.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   2160
      Picture         =   "frmPPalWhose.frx":0A02
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4935
   End
   Begin VB.Menu mnClientes 
      Caption         =   "Clientes"
      Begin VB.Menu mnClientes1 
         Caption         =   "Mantenimiento clientes"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "Facturas "
         Index           =   1
         Shortcut        =   ^F
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "Facturas rectificativas"
         Index           =   2
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "Histórico de facturas"
         Index           =   3
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "Contabilizacion facturas"
         Index           =   4
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnClientes1 
         Caption         =   "Clientes potenciales"
         Index           =   6
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnExped 
      Caption         =   "Expedientes"
      Begin VB.Menu mnVerExpedientes 
         Caption         =   "Ver expedientes"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnInformes 
      Caption         =   "Informes"
      Begin VB.Menu mnInformes1 
         Caption         =   "Ventas por cliente"
         Index           =   0
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Ventas por meses"
         Index           =   1
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Ventas por familia"
         Index           =   2
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Ventas por artículo"
         Index           =   3
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Detalle facturacion"
         Index           =   4
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Listado de clientes"
         Index           =   6
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Listado clientes nuevos"
         Index           =   7
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Etiquetas "
         Index           =   8
      End
      Begin VB.Menu mnInformes1 
         Caption         =   "Cartas a clientes"
         Index           =   9
      End
   End
   Begin VB.Menu mnCRMmenu 
      Caption         =   "CRM"
      Begin VB.Menu mnCRM 
         Caption         =   "Mantenimiento acciones comerciales"
         Index           =   0
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Tipos acciones comerciales"
         Index           =   1
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Generar acciones comerciales"
         Index           =   2
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion masiva"
         Index           =   3
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion resumen CRM"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmPPalWhose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    frmFacClientes.Show vbModal
    
End Sub

Private Sub Command2_Click()
    frmFacEntAlbaranes2.Show vbModal
End Sub

Private Sub cmdAccion_Click()

End Sub

Private Sub Form_Activate()



    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta
    'Finalizamos
    End
End Sub

Private Sub mnClientes1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'mto
        frmFacClientes.Show vbModal
        
        
    Case 1


        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALV"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal

    Case 2
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ART"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal

    Case 3
        'Factura
        frmFacHcoFacturas2.hcoCodMovim = ""
        frmFacHcoFacturas2.Show vbModal
    
    Case 4
        AbrirListado 223
        
    Case 6
        frmFacClienPot.Show vbModal
    End Select
    
End Sub


Private Sub mnCRM_Click(Index As Integer)
    
        Select Case Index
        Case 0
            frmCRMMto.DesdeElCliente = 0 'No clien
            frmCRMMto.TipoPredefinido = 0   'Ninguno
            frmCRMMto.Show vbModal
            
        Case 1
            frmCRMtipos.Show vbModal
        
        Case 2
            frmCRMVarios.Opcion = 0
            frmCRMVarios.Show vbModal
            
        Case 3
            frmListadoOfer.OpcionListado = 406
            frmListadoOfer.Show vbModal
        Case 4
            frmCRMVarios.Opcion = 1
            frmCRMVarios.Show vbModal
            
        End Select
        
End Sub


Private Sub mnInformes1_Click(Index As Integer)
    Select Case Index
    Case 0
        AbrirListadoPed (227)
        BorrarTempInformes
        
    Case 1
        AbrirListadoPed (229)
    Case 2
        AbrirListadoOfer (230)
    Case 3
        
        frmListado3.Opcion = 18
        frmListado3.Show vbModal
    Case 4
        AbrirListadoOfer (231)
        
        
        
    'CLIENTES
    Case 6
        AbrirListadoOfer (47) '47: Informes Clientes
        
    Case 7
         'Informe de Altas de Nuevos Clientes
        AbrirListadoOfer (48) '48: Informes Altas Clientes

        
    Case 8
         'Etiquetas de clientes
        AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
        
    Case 9
        AbrirListadoOfer (91) '91: Informe Cartas a Clientes
         
    Case Else
        MsgBox "Error en opcion", vbExclamation
    
    End Select
    
    
    
End Sub

Private Sub mnVerExpedientes_Click()
    frmWH_SelExp.Show vbModal
End Sub
