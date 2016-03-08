VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensaje2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ForeColor       =   &H00000000&
   Icon            =   "frmMensaje2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame FrameEnviar 
      Height          =   5895
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8775
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   4260
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   2445
         Index           =   2
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmMensaje2.frx":000C
         Top             =   3240
         Width           =   5415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   2520
         Width           =   7455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "e-mail"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "frmMensaje2.frx":0012
         Left            =   240
         List            =   "frmMensaje2.frx":0014
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1230
         Left            =   3600
         TabIndex        =   27
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1320
         Picture         =   "frmMensaje2.frx":0016
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   960
         Picture         =   "frmMensaje2.frx":6868
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   2
         Left            =   6120
         Picture         =   "frmMensaje2.frx":D0BA
         Top             =   240
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   1
         Left            =   5760
         Picture         =   "frmMensaje2.frx":1390C
         Top             =   240
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   5400
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Destinatarios externos"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "MENSAJE"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   26
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Adjuntos"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   21
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1920
         Picture         =   "frmMensaje2.frx":1A15E
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Destinatarios ARIDOC"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FrameRecibido 
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8775
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmMensaje2.frx":209B0
         Top             =   600
         Width           =   7575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1440
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   3405
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmMensaje2.frx":209B6
         Top             =   2400
         Width           =   8535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "E-mail"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "PARA"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMensaje2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enviado As Boolean
Public Opcion As Byte
Public vM As Cmailc
Public Carpetas As String   ' La primera sera la carpeta ppal, a partir de ahi, las subcarpetas
Public TodasCarpetas As String
Public ImagenAEnviar As String

    '0  - NUEVO
Dim I As Integer
Dim PrimeraVez  As Boolean
Dim It As ListItem


Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub cmdCerrar_Click()
    Set vM = Nothing
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Command1_Click()
Dim CadenaPara As String
Dim Externo As Collection
Dim ListaFicheros As String

    'Si tiene adjuntos, ira por e-mail
    If ListView1.ListItems.Count > 0 Then
        If Check2.Value = 0 Then
            MsgBox "Para enviar adjuntos debe indicar la opcion de e-mail", vbExclamation
            Exit Sub
        End If
    End If

    CadenaPara = ""
    If List1.ListCount > 0 Then CadenaPara = "S"
    If ListView2.ListItems.Count > 0 Then
        CadenaPara = CadenaPara & "S"
        Check2.Value = 1 'Pongo a TRUE el envio por mail
    End If
    
    If Len(CadenaPara) = 0 Then
        MsgBox "Selecione algun destinatario", vbExclamation
        Exit Sub
    End If
    
    Text2(1).Text = Trim(Text2(1).Text)
    If Text2(1).Text = "" Then
        MsgBox "Asunto no puede estar vacio", vbExclamation
        Exit Sub
    End If
    
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione el tipo de mensaje", vbExclamation
        Exit Sub
    End If
    
    
    'Si tiene enviar e-mail deberiamos comprobar que todos tienen
    'direccion e-mail
    '### FALTA
'    BorrarTemporal1
    Set listacod = Nothing
    Set listacod = New Collection
    
    
    CadenaPara = ""
    For I = 0 To List1.ListCount - 1
        If CadenaPara <> "" Then CadenaPara = CadenaPara & ";"
        CadenaPara = CadenaPara & List1.List(I)
        '### FALTA
'        InsertaTemporal List1.ItemData(I)
        listacod.Add List1.ItemData(I)
    Next I
    
    Set Externo = Nothing
    Set Externo = New Collection
    For I = 1 To ListView2.ListItems.Count
        
        If CadenaPara <> "" Then CadenaPara = CadenaPara & ";"
        CadenaPara = CadenaPara & ListView2.ListItems(I).Text
        Externo.Add CStr(ListView2.ListItems(I).Text & "|" & ListView2.ListItems(I).Tag & "|")
    Next I
    
    If Len(CadenaPara) > 255 Then CadenaPara = Mid(CadenaPara, 1, 251) & " ..."
        
    
    'Tomo prestado esta variable
    Set listaimpresion = Nothing
    Set listaimpresion = New Collection
    
    Set miRsAux = New ADODB.Recordset
    DatosCopiados = "Select nombre,usuarios.codusu from tmpFich,usuarios Where imagen = usuarios.codusu"
    DatosCopiados = DatosCopiados & " AND tmpFich.codusu =" & vUsu.Codigo
    DatosCopiados = DatosCopiados & " AND codequipo= " & vUsu.PC & " AND (maildir ='' or (maildir is null))"
    miRsAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DatosCopiados = ""
    While Not miRsAux.EOF
        DatosCopiados = DatosCopiados & miRsAux!Nombre & vbCrLf
       
        listaimpresion.Add CStr(miRsAux!CodUsu)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Check2.Value Then
        If DatosCopiados <> "" Then
            DatosCopiados = "Los siguientes usuarios no tienen direccion e-mail:" & vbCrLf & vbCrLf & DatosCopiados
            If ListView1.ListItems.Count = 0 Then
                DatosCopiados = DatosCopiados & vbCrLf & vbCrLf & "¿Desea continuar?"
                If MsgBox(DatosCopiados, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                
            Else
                'Ha marcado adjuntos. No debe dejar enviar a los que no tengan e-mail
                DatosCopiados = DatosCopiados & vbCrLf & vbCrLf & "Debe desmarcarlos."
                MsgBox DatosCopiados, vbExclamation
                Exit Sub
            End If
        End If
        
        
    End If
    
    
    
    'Si tiene adjuntos Traemos los ficheros
    ListaFicheros = ""
    If ListView1.ListItems.Count > 0 Then
        If Not TraerLosDatosAdjuntos(ListaFicheros) Then Exit Sub
    End If
    
    'Llegados aqui, creamos el mensaje.
    
        Set vM = Nothing
        Set vM = New Cmailc
        
    vM.asunto = Text2(1).Text
    vM.Fecha = CDate(Text2(0).Text)
    vM.Origen = vUsu.Codigo
    vM.Texto = Text2(2).Text
    vM.email = Abs(Check2.Value)
    vM.Tipo = Combo1.ItemData(Combo1.ListIndex)
    If vM.GenerarMensaje(listacod, CadenaPara, Externo) = 1 Then
        'Borramos
        Conn.Execute "Delete from maill where codmail =" & vM.codmail
        Conn.Execute "Delete from mailc where codmail =" & vM.codmail
        Conn.Execute "Delete from maile where codmail =" & vM.codmail
        Conn.Execute "Delete from maildestexth where codmail =" & vM.codmail
        
    Else
        'Ha ido todo bien
        If Check2.Value = 1 Then
            'HAY QUE ENVIAR POR MAIL, excepto los que no tienen mail
            DatosCopiados = "UPDATE mailc SET email=0 where"
            DatosCopiados = DatosCopiados & " origen = " & vUsu.Codigo & " and codmail = " & vM.codmail
            DatosCopiados = DatosCopiados & " AND destino = "
            If Not listaimpresion Is Nothing Then
                For I = 1 To listaimpresion.Count
                    Conn.Execute DatosCopiados & listaimpresion(I)
                Next I
                'PARA Abriremos la pantalla de envio de mail
                I = List1.ListCount - listaimpresion.Count
                
            Else
                I = List1.ListCount
            End If
            
            If ListView2.ListItems.Count > 0 Then I = 1
            
            
            If I > 0 Then
                'HAY que enviar mensajes
'                frmEMail.ListaDeFicheros = ListaFicheros
'                frmEMail.IdMail = vM.codmail
'                frmEMail.Show vbModal
            End If
            'PONERMO A NOTINH ALGUNOS VALORES
            Set listaimpresion = Nothing
        End If
    End If
    Set listacod = Nothing
    Set vM = Nothing
    DatosModificados = True
    Unload Me
        
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 0 Then
            If List1.ListCount = 0 Then
                List1.SetFocus
            Else
                Text2(2).SetFocus
            End If
            
            If Me.ImagenAEnviar <> "" Then
                DatosCopiados = Me.ImagenAEnviar
                InsertaAdjunto
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
'    Set ListView1.SmallIcons = frmPpal.ImageList2
     'Icono del formulario principal
    Me.Icon = frmPpal.imgListComun.ListImages(20).Picture

    'cargar los iconos
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(44).Picture

    PrimeraVez = True
    limpiar Me
    Me.FrameEnviar.visible = False
    Me.FrameRecibido.visible = False
    
    On Error GoTo ECargar
    
    If Opcion = 0 Then
        Me.FrameEnviar.visible = True
        H = FrameEnviar.Height
        W = FrameEnviar.Width
        
        If vM Is Nothing Then
            Text2(0).Text = Format(Now, "dd/mm/yyyy")
'            Check2.Enabled = (vUsu.e_server <> "")
                 
        Else
            'Es un reenvio o respuesta
            Text2(0).Text = Format(vM.Fecha, "dd/mm/yyyy")
            Text2(1).Text = vM.asunto
            Text2(2).Text = vM.Texto
            'Añadimos el usario
            List1.Clear
            
            If vM.email Then
                Check2.Value = 1
            Else
                Check2.Value = 0
            End If
            
            PonCampos
        End If
        Set vM = Nothing


        'Text2(0).BackColor = CLng("&H80000018")
        Command1.visible = True
    Else
        Command1.visible = False
        FrameRecibido.visible = True
        H = FrameRecibido.Height
        W = FrameRecibido.Width
        
        'Ahora ponemos los campos del mensaje ande corresponda
        Text1(0).Text = ""
        Text1(4).Text = ""
        Text1(1).Text = Format(vM.Fecha, "dd/mm/yyyy")
        Text1(2).Text = vM.Texto
        Text1(3).Text = vM.asunto
        Text3.Text = ""
        PonCampos
                
        
    End If
    
    Me.Width = W + 120
    Me.Height = H + 920
    

   
    Combo1.Clear
    For H = 0 To TotalTipos
        If ArrayTipoMen(H).Descripcion <> "" Then
            Combo1.AddItem ArrayTipoMen(H).Descripcion
            Combo1.ItemData(Combo1.NewIndex) = H
        End If
    Next H
    Combo1.ListIndex = 0

ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub


Private Sub PonCampos()
    On Error GoTo EPonCa
    If Opcion = 1 Then
        If vM.Recibido Then
            Label1(0).Caption = "DE"
            DatosCopiados = DevuelveDesdeBD("nombre", "usuarios", "codusu", CStr(vM.Origen), "N")
        Else
            Label1(0).Caption = "PARA"
            DatosCopiados = vM.Textopara
        End If
        Text1(0).Text = DatosCopiados
        DatosCopiados = ArrayTipoMen(vM.Tipo).Descripcion
        Text1(4).Text = DatosCopiados
        
        'Hacemos el SELECT
        If Not vM.Recibido Then
            Set miRsAux = New ADODB.Recordset
            DatosCopiados = "Select * from maildestext"
            If vM.EnHco Then DatosCopiados = DatosCopiados & "h"
            DatosCopiados = DatosCopiados & " WHERE codmail =" & vM.codmail
            miRsAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                Text3.Text = Text3.Text & miRsAux!Nombre & "   ( " & miRsAux!mail & ")"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
    Else
        
        If vM.Destino >= 0 Then
                DatosCopiados = DevuelveDesdeBD("nombre", "usuarios", "codusu", CStr(vM.Destino), "N")
                If DatosCopiados <> "" Then
                    List1.AddItem DatosCopiados
                    List1.ItemData(List1.NewIndex) = vM.Destino
                End If
        End If
            
        For I = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(I) = vM.Tipo Then
                'Es este
                Combo1.ListIndex = I
                Exit For
            End If
        Next I
    End If
    Exit Sub
EPonCa:
    MuestraError Err.Number, "Poniendo campos(2)"
End Sub

Private Sub Image1_Click()


    'Borramos temporal
    '### FALTA
'    BorrarTemporal1
    
    'Insertamos
    DatosCopiados = ""
    For I = 0 To List1.ListCount - 1
        DatosCopiados = DatosCopiados & List1.ItemData(I) & "|"
    Next

    '### FALTA
'    frmListaUserPcs2.Opcion = 0
'    frmListaUserPcs2.Show vbModal
    
    If DatosCopiados <> "" Then
        List1.Clear
'        DatosCopiados = "Select nombre,codusu from tmpFich,usaurios where codusu =" & vUsu.codusu
'        DatosCopiados = DatosCopiados & " AND codpc= " & vUsu.PC
'        DatosCopiados = DatosCopiados & " AND tmpfich.imagen = usuarios.codusu"
'        DatosCopiados = DatosCopiados & " ORDER BY nomusu"
'
        DatosCopiados = "Select nombre,usuarios.codusu from tmpFich,usuarios Where imagen = usuarios.codusu"
        DatosCopiados = DatosCopiados & " AND tmpFich.codusu =" & vUsu.Codigo
        DatosCopiados = DatosCopiados & " AND codequipo= " & vUsu.PC & "  ORDER BY nombre"
        
        
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            List1.AddItem miRsAux!Nombre
            List1.ItemData(List1.NewIndex) = miRsAux!CodUsu
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Sub Image2_Click()

'    'Para adjuntar archivos
'    On Error GoTo E1
'    DatosCopiados = ""
'    frmBusca2.DesdeEmail = True
'    frmBusca2.Carpetas = Carpetas
'    frmBusca2.TodasCarpetas = TodasCarpetas
'    frmBusca2.Show vbModal
'    If DatosCopiados <> "" Then InsertaAdjunto
'
'
'E1:
'    If Err.Number <> 0 Then MuestraError Err.Number
'    Set miRSAux = Nothing
End Sub


Private Sub InsertaAdjunto()
Dim C As String
    On Error GoTo EI
        Set miRsAux = New ADODB.Recordset
        C = "Select campo1,codigo,timagen.codext,codcarpeta,exten from timagen,extension where "
        C = C & " timagen.codext=extension.codext AND "
        C = C & " codigo =" & Mid(RecuperaValor(DatosCopiados, 2), 2)
        miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            Set It = ListView1.ListItems.Add(, "C" & miRsAux!Codigo)
            It.Text = miRsAux!campo1
            It.Tag = miRsAux!codcarpeta & "|" & miRsAux!Exten & "|"
            'ICONO
            It.SmallIcon = miRsAux!codext + 1
        End If
        miRsAux.Close
EI:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub


Private Sub Image3_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("¿Quitar de adjuntos: " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    ListView1.ListItems.Remove ListView1.SelectedItem.Key
End Sub

Private Sub ImgMail_Click(Index As Integer)

'    If Index > 0 Then
'        If ListView2.SelectedItem Is Nothing Then Exit Sub
'    End If
'
'    If Index = 2 Then
'        ListView2.ListItems.Remove ListView2.SelectedItem.Index
'
'    Else
'        If Index = 1 Then
'            DatosCopiados = ListView2.SelectedItem.Text & "|" & ListView2.SelectedItem.Tag & "|"
'        Else
'            DatosCopiados = ""
'        End If
'        frmPregunta.Opcion = 21
'        frmPregunta.Show vbModal
'        If DatosCopiados <> "" Then
'            If Index = 0 Then
'                Set It = ListView2.ListItems.Add()
'                It.Text = RecuperaValor(DatosCopiados, 1)
'                It.Tag = RecuperaValor(DatosCopiados, 2)
'                It.ToolTipText = It.Tag
'            Else
'                ListView2.SelectedItem.Text = RecuperaValor(DatosCopiados, 1)
'                ListView2.SelectedItem.Tag = RecuperaValor(DatosCopiados, 2)
'                ListView2.SelectedItem.ToolTipText = ListView2.SelectedItem.Tag
'            End If
'            Check2.Value = 1
'        End If
'    End If
'
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Image1_Click
    Else
        KEYpressGnral KeyAscii, 2, False
    End If
End Sub



Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Then KEYpressGnral KeyAscii, 2, False
End Sub




Private Function TraerLosDatosAdjuntos(ByRef Lista As String) As Boolean
'Dim I As Integer
'Dim C As Ccarpetas
'Dim Cad As String
'Dim F, FS
'On Error GoTo ETraerLosDatosAdjuntos
'
'    TraerLosDatosAdjuntos = False
'    If Dir(App.Path & "\mail", vbDirectory) = "" Then MkDir (App.Path & "\mail")
'
'    If Dir(App.Path & "\mail\*.*", vbArchive) <> "" Then Kill App.Path & "\mail\*.*"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'
'    Set C = New Ccarpetas
'
'    For I = 1 To ListView1.ListItems.Count
'        If C.Leer(RecuperaValor(ListView1.ListItems(I).Tag, 1), ModoTrabajo <> vbNorm) = 1 Then
'            'msgbox
'            GoTo ETraerLosDatosAdjuntos
'        Else
'            DevuelveNombreFichero ListView1.ListItems(I).Text, RecuperaValor(ListView1.ListItems(I).Tag, 2), Cad, True
'
'            If Not TraerFicheroFisico(C, Cad, Val(Mid(ListView1.ListItems(I).Key, 2))) Then
'                Cad = "Fichero: " & Cad & vbCrLf & "Codigo imagen: " & Val(Mid(ListView1.ListItems(I).Key, 2))
'
'                MsgBox "Error trayendo datos. " & vbCrLf & Cad, vbExclamation
'                GoTo ETraerLosDatosAdjuntos
'            End If
'
'            If Cad <> "" Then
'                If Dir(Cad, vbArchive) = "" Then
'                    MsgBox "Fichero no encontrado en carpeta \mail", vbExclamation
'                    GoTo ETraerLosDatosAdjuntos
'                Else
'                    Set F = FS.GetFile(Cad)
'                    Lista = Lista & F.shortpath & "|"
'                End If
'            End If
'
'
'
'
'
'
'        End If
'    Next I
'    'Ahora copiamos los archivos recibidos
'    TraerLosDatosAdjuntos = True
'
'
'
'ETraerLosDatosAdjuntos:
'    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
'    Set C = Nothing
'    Set FS = Nothing
'    Set F = Nothing
End Function
