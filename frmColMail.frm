VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColMail 
   Caption         =   "Mensajes"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   Icon            =   "frmColMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameAcciones 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   2175
      Begin VB.Label Label1 
         Caption         =   "Comprobando datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.CommandButton cmdComun 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   9000
      Picture         =   "frmColMail.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   150
      Width           =   495
   End
   Begin VB.CommandButton cmdComun 
      Height          =   495
      Index           =   0
      Left            =   8400
      Picture         =   "frmColMail.frx":1A84
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Tipos mensaje"
      Top             =   150
      Width           =   495
   End
   Begin VB.Frame FrameRecibidos 
      Height          =   700
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   6
         Left            =   600
         Picture         =   "frmColMail.frx":2B06
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Leer"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   5
         Left            =   3840
         Picture         =   "frmColMail.frx":2FB3
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Pasar HCO"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   4
         Left            =   3240
         Picture         =   "frmColMail.frx":345A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Eliminar"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   3
         Left            =   2520
         Picture         =   "frmColMail.frx":3952
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar como leido"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   2
         Left            =   1800
         Picture         =   "frmColMail.frx":3E5E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Renviar"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   1
         Left            =   1320
         Picture         =   "frmColMail.frx":4358
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Responder"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdRec 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmColMail.frx":4848
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo"
         Top             =   150
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "De"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Asunto"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "@"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmColMail.frx":4D5F
      Left            =   120
      List            =   "frmColMail.frx":4D6F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmColMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean



'Para los adjuntos en los nuevos correos
Public Carpetas1 As String   ' La primera sera la carpeta ppal, a partir de ahi, las subcarpetas
Public TodasCarpetas1 As String

'variables para zona comun
Dim SQL As String
Dim It As ListItem
Dim I As Long


Private Sub cmdComun_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
    Else
        frmTiposMensajes.Show vbModal
    End If
End Sub

Private Sub cmdRec_Click(Index As Integer)
Dim VMe As Cmailc
Dim Dest As Integer
Dim VN As Cmailc
Dim Recibido As Boolean
Dim HCO As Boolean

    HCO = Combo1.ListIndex > 1  'En HCO

    Select Case Index
    Case 0
            'NUEVO
            DatosModificados = False

            Set frmMensaje2.vM = Nothing
            frmMensaje2.Carpetas = Carpetas1
            frmMensaje2.TodasCarpetas = TodasCarpetas1
            frmMensaje2.ImagenAEnviar = ""
            frmMensaje2.Opcion = 0
            frmMensaje2.Show vbModal
            'Si han aceptado
            
            If DatosModificados = True Then
                'Refrescar
                If Combo1.ListIndex = 1 Then
                    'SOLO SI ESTAMOS EN LA DE ENVIADOS
                    Me.Refresh
                    CargaMensajes False
                End If
                
            End If
    Case 1, 2
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            I = Val(ListView1.SelectedItem.Tag)
            
            Set VMe = New Cmailc
            If VMe.Leer(I, -1, HCO) Then
                Set VMe = Nothing
                Exit Sub
            End If
            
            Set VN = New Cmailc
            VN.Origen = vUsu.Codigo
            
            If Index = 1 Then
                VN.asunto = "RE: " & VMe.asunto
                VN.Texto = "Respuesta"
                VN.Destino = VMe.Origen
            Else
                VN.asunto = "RV: " & VMe.asunto
                VN.Texto = "Reenvio"
                VN.Destino = -1
            End If
            SQL = vbCrLf & vbCrLf & "-----------------------------------------------" & vbCrLf
            SQL = SQL & VN.Texto & "  del mensaje enviado por :" & ListView1.SelectedItem.Text
            SQL = SQL & vbCrLf & "--------------------------------------------" & vbCrLf & VMe.Texto
            
            VN.Texto = SQL
            VN.Fecha = Now
            VN.leido = 0
            Set VMe = Nothing
            frmMensaje2.Carpetas = Carpetas1
            frmMensaje2.TodasCarpetas = TodasCarpetas1
            frmMensaje2.ImagenAEnviar = ""
            frmMensaje2.Opcion = 0
            Set frmMensaje2.vM = VN
            frmMensaje2.Show vbModal
            Set VN = Nothing
    Case 3, 6
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            
            I = Val(ListView1.SelectedItem.Tag)
            
            Set VMe = New Cmailc
            If Combo1.ListIndex = 0 Or Combo1.ListIndex = 2 Then
                Recibido = True
            Else
                Recibido = False
            End If
            If VMe.Leer(I, Recibido, HCO) Then
                Set VMe = Nothing
                Exit Sub
            End If
                
            If Index = 6 Then
                Set frmMensaje2.vM = VMe
                frmMensaje2.Opcion = 1
                frmMensaje2.Carpetas = Carpetas1
                frmMensaje2.ImagenAEnviar = ""
                frmMensaje2.TodasCarpetas = TodasCarpetas1
                frmMensaje2.Show vbModal
            End If
            'FALTA poner leido
            If ListView1.SelectedItem.Bold Then
                VMe.MarcarComoLeido
                ListView1.SelectedItem.Bold = False
                For I = 1 To 3
                    ListView1.SelectedItem.ListSubItems(I).Bold = False
                Next I
            End If
            Set VMe = Nothing
            ListView1.SetFocus
            Me.Refresh
    Case 4, 5
            'BORRAR o PASAR A HCO
            'Tambien decir que si es paso a HCO,  y ya estamos en HCO, no hacemos nada
            If Index = 5 And HCO Then Exit Sub
            
            SQL = ""
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Selected Then SQL = SQL & "1"
            Next I
            If SQL = "" Then Exit Sub
            I = Len(SQL)
        
        
            
        
            If Index = 4 Then
                SQL = "eliminar "
            Else
                SQL = "traspasar a historico"
            End If
            SQL = "Seguro que desea " & SQL
            If I = 1 Then
                SQL = SQL & " el archivo seleccionado"
            Else
                SQL = SQL & " los (" & I & ")archivos seleccionados"
            End If
            If MsgBox(SQL & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
            
            If Combo1.ListIndex = 0 Or Combo1.ListIndex = 2 Then
                Recibido = True
            Else
                Recibido = False
            End If

        
        
            Set VMe = New Cmailc
            For I = ListView1.ListItems.Count To 1 Step -1
                If ListView1.ListItems(I).Selected Then
                    If VMe.Leer(Val(ListView1.ListItems(I).Tag), Recibido, HCO) = 0 Then
                        If Index = 4 Then
                            '------------
                            'B O R R A R
                            '------------
                            VMe.Eliminar
                            
                        Else
                            '-----------
                            ' A HCO
                            '------------
                            VMe.PasarAHistorico
                        End If
                        ListView1.ListItems.Remove I
                    End If
                End If
            Next I
    End Select
End Sub




Private Sub Combo1_Click()
    If Combo1.Tag = -1 Then Exit Sub
    If Combo1.Tag = Combo1.ListIndex Then Exit Sub
    
    If Combo1.ListIndex = 1 Or Combo1.ListIndex = 3 Then
        I = 1
        ListView1.ColumnHeaders(1).Text = "PARA"
    Else
        I = 0
        ListView1.ColumnHeaders(1).Text = "DE"
    End If
    Me.cmdRec(1).Enabled = I = 0
    'Me.cmdRec(3).Enabled = Combo1.ListIndex = 0
    Me.cmdRec(5).Enabled = Combo1.ListIndex < 2
    
    CargaMensajes I = 0
    Combo1.Tag = Combo1.ListIndex
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
'        If vUsu.preferencias.mailPasarHCO > 0 And vUsu.preferencias.mailPasarHCO < 100 Then
'            'Realizar paso HCO
'            RealizandoTraspaso
'            vUsu.preferencias.mailPasarHCO = vUsu.preferencias.mailPasarHCO + 100
'        End If
        
        
        
        Set Me.ListView1.SmallIcons = frmPpal.ImageListMAIL
        CargaMensajes True
        Combo1.Tag = 0 'Para poner ya el valor
        Me.FrameAcciones.visible = False
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Combo1.Tag = -1
    Combo1.ListIndex = 0
    PrimeraVez = True
    FrameRecibidos.visible = True
    FrameAcciones.visible = True
    
End Sub

Private Sub Form_Resize()
Dim X As Integer
   If WindowState = 1 Then Exit Sub ' ha pulsado minimizar
    
    If Me.Width < 5850 Then Me.Width = 5850
    If Me.Height < 4100 Then Me.Height = 4100
    
    ListView1.Width = Me.Width - 320
    ListView1.Height = Me.Height - ListView1.Top - 800
    
    
    cmdComun(1).Left = Me.Width - 150 - cmdComun(1).Width
    cmdComun(0).Left = cmdComun(1).Left - cmdComun(0).Width - 60
    
    Me.FrameAcciones.Width = Me.Width - 320
    
    
    Me.FrameRecibidos.Width = cmdComun(0).Left - Me.FrameRecibidos.Left - 220
    
    
    
    
    'Me.FrameEnviados.Width = FrameRecibidos.Width
    
    X = ListView1.Width - 2350
    X = CInt(X / 5)
    ListView1.ColumnHeaders(1).Width = 2 * X
    ListView1.ColumnHeaders(3).Width = 3 * X
End Sub


Private Sub CargaMensajes(Recibidos As Boolean)
    If Recibidos Then
        CargaMensajesRecibidos
    Else
        CargaMensajesEnviados
    End If
End Sub


Private Sub CargaMensajesRecibidos()
Dim EnHco As Boolean
    EnHco = (Combo1.ListIndex > 1)
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    SQL = "select nombre,mailc.*,maill.* from mailc"
    If EnHco Then SQL = SQL & "h"
    SQL = SQL & " as mailc,maill,usuarios WHERE maill.codmail=mailc.codmail AND "
    SQL = SQL & "mailc.origen = usuarios.codigo and destino = " & vUsu.Codigo
    'ORDEN
    SQL = SQL & " ORDER BY fecha,maill.codmail,nombre"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add
        It.Text = miRsAux!Nombre

            
        It.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        It.SubItems(2) = miRsAux!asunto

        
        If miRsAux!email = 1 Then
            It.SubItems(3) = "*"
        Else
            It.SubItems(3) = ""
        End If
        
        If Not EnHco Then
            If miRsAux!leido = 0 Then
                It.Bold = True
                It.ListSubItems(1).Bold = True
                It.ListSubItems(2).Bold = True
                It.ListSubItems(3).Bold = True
            End If
        End If
        It.Tag = miRsAux!codmail
    
        
        'el color
        If ArrayTipoMen(miRsAux!Tipo).Color <> 0 Then
            It.ForeColor = ArrayTipoMen(miRsAux!Tipo).Color
            For I = 1 To 3
                It.ListSubItems(I).ForeColor = ArrayTipoMen(miRsAux!Tipo).Color
            Next I
            It.ToolTipText = ArrayTipoMen(miRsAux!Tipo).Descripcion
        End If
        
        I = ArrayTipoMen(miRsAux!Tipo).Icono
        If I > 0 Then It.SmallIcon = I
        'If miRSAux!Tipo > 0 Then It.SmallIcon = miRSAux!Tipo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub CargaMensajesEnviados()

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    SQL = "select maile.*,maill.* from maile"
    If Combo1.ListIndex > 1 Then SQL = SQL & "h"
    SQL = SQL & " as maile,maill WHERE maill.codmail=maile.codmail AND "
    SQL = SQL & " origen = " & vUsu.Codigo
    'ORDEN
    SQL = SQL & " ORDER BY fecha,maile.codmail"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add
        It.Text = miRsAux!Textopara

            
        It.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        It.SubItems(2) = miRsAux!asunto

        
        If miRsAux!email = 1 Then
            It.SubItems(3) = "*"
        Else
            It.SubItems(3) = ""
        End If
        
        
           
        It.Tag = miRsAux!codmail & "|"
        
        
        'el color
        If ArrayTipoMen(miRsAux!Tipo).Color <> 0 Then
            It.ForeColor = ArrayTipoMen(miRsAux!Tipo).Color
            For I = 1 To 3
                It.ListSubItems(I).ForeColor = ArrayTipoMen(miRsAux!Tipo).Color
            Next I
            It.ToolTipText = ArrayTipoMen(miRsAux!Tipo).Descripcion
        End If
                
        I = ArrayTipoMen(miRsAux!Tipo).Icono
        If I > 0 Then It.SmallIcon = I
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub




Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdRec_Click 6
End Sub



Private Sub RealizandoTraspaso()
'Dim F As Date
'
'
'
'    Me.Label1.Caption = "Realizando Traspaso HCO enviados"
'    Me.Refresh
'    I = vUsu.preferencias.mailPasarHCO * -1
'    F = DateAdd("m", I, Now)
'
'
'    SQL = "select maile.codmail from maile,maill"
'    SQL = SQL & " WHERE maill.codmail = maill.codmail "
'    SQL = SQL & " AND fecha <='" & Format(F, FormatoFecha) & "'"
'    SQL = SQL & " AND origen = " & vUsu.codigo
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not miRsAux.EOF Then
'
'        'Pasamos a HCO
'        SQL = "INSERT INTO maileh Select maile.* from maile,maill"
'        SQL = SQL & " WHERE maile.codmail = maill.codmail "
'        SQL = SQL & " AND fecha <='" & Format(F, FormatoFecha) & "'"
'        SQL = SQL & " AND origen = " & vUsu.codigo
'        Conn.Execute SQL
'
'
'        SQL = "DELETE from maile where origen = " & vUsu.codigo & " AND codmail = "
'        While Not miRsAux.EOF
'            Conn.Execute SQL & miRsAux!codmail
'            miRsAux.MoveNext
'        Wend
'
'    End If
'    miRsAux.Close
'
'
'    Me.Label1.Caption = "Realizando Traspaso HCO enviados"
'    Me.Refresh
'
'    SQL = "select mailc.codmail from mailc,maill"
'    SQL = SQL & " WHERE mailc.codmail = maill.codmail "
'    SQL = SQL & " AND fecha <='" & Format(F, FormatoFecha) & "'"
'    SQL = SQL & " AND destino = " & vUsu.codigo
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not miRsAux.EOF Then
'
'        'Pasamos a HCO
'        SQL = "INSERT INTO mailch Select mailc.* from mailc,maill"
'        SQL = SQL & " WHERE mailc.codmail = maill.codmail "
'        SQL = SQL & " AND fecha <='" & Format(F, FormatoFecha) & "'"
'        SQL = SQL & " AND destino = " & vUsu.codigo
'        Conn.Execute SQL
'
'
'        SQL = "DELETE from mailc where destino = " & vUsu.codigo & " AND codmail = "
'
'        While Not miRsAux.EOF
'            Conn.Execute SQL & miRsAux!codmail
'            miRsAux.MoveNext
'        Wend
'
'    End If
'    miRsAux.Close
'
        
    
End Sub
