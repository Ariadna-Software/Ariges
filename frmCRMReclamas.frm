VERSION 5.00
Begin VB.Form frmCRMReclamas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reclamaciones tesoreria(Arimoney)"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Index           =   7
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmCRMReclamas.frx":0000
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3360
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   240
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2280
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   600
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "1"
      Top             =   2280
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   240
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2280
      Width           =   345
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1200
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5520
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   5
      Left            =   1320
      Picture         =   "frmCRMReclamas.frx":0006
      ToolTipText     =   "Buscar fecha"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   3000
      Picture         =   "frmCRMReclamas.frx":0091
      ToolTipText     =   "Buscar fecha"
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Importe"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "F: reclamacion"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Vto"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   19
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha fact"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Factura"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Datos reclamación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   5505
   End
   Begin VB.Label Label3 
      Caption         =   "Datos vencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   5505
   End
   Begin VB.Label Label2 
      Caption         =   "Datos cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5505
   End
   Begin VB.Label Label1 
      Caption         =   "Serie"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta contable"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label LabelCRM 
      Caption         =   "aqui ira el nomclient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   5505
   End
End
Attribute VB_Name = "frmCRMReclamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Intercambio As String  ' codig|nomclien|codmacta|nommacta|   NUEVO. Para la ariconta lleva el numlinea (si procede
Private Codigo2  As Long
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim SQL As String
Dim I As Integer

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Text1(5).Text = "" Or Text1(6).Text = "" Then
            MsgBox "Campos obligatorios: fecha reclamacion e importe", vbExclamation
            Exit Sub
        End If
        'Acciones...
        I = 0
        If Codigo2 < 0 Then
            'NUEVO
            I = 1 'por si da error reestablacer el codigo2 a menos1
            'shcocob (codigo,numserie,codfaccl,fecfaccl,numorden,impvenci,codmacta,nommacta,carta,fecreclama,observaciones)
            SQL = DevuelveDesdeBD(conConta, "max(codigo)", "shcocob", "1", "1")
            If SQL = "" Then SQL = "0"
            Codigo2 = Val(SQL) + 1
            SQL = "INSERT INTO shcocob (codigo,numserie,codfaccl,fecfaccl,numorden,impvenci,codmacta,nommacta,carta,"
            SQL = SQL & "fecreclama,observaciones) VALUES (" & Codigo2 & ","
            SQL = SQL & DBSet(Text1(1).Text, "T", "S") & "," '
            SQL = SQL & DBSet(Text1(2).Text, "N", "S") & "," '= DBLet(miRsAux!Codfaccl, "T")
            SQL = SQL & DBSet(Text1(3).Text, "F", "S") & "," '= DBLet(miRsAux!fecfaccl, "F")
            SQL = SQL & DBSet(Text1(4).Text, "N", "S") & "," '= DBLet(miRsAux!numorden, "T")
            
            SQL = SQL & DBSet(Text1(6).Text, "N") & "," ' z'= DBLet(miRsAux!ImpVenci, "N")
            SQL = SQL & DBSet(Text1(0).Text, "T", "S") & "," ' = miRsAux!Codmacta
            SQL = SQL & DBSet(Text2(0).Text, "T", "S") & ",0," ' = miRsAux!nommacta  Y  CARTA que le pondre un 0
            SQL = SQL & DBSet(Text1(5).Text, "F") & "," ' = DBLet(miRsAux!fecreclama, "F")
            SQL = SQL & DBSet(Text1(7).Text, "T", "S") & ")" ' = DBLet(miRsAux!Observaciones, "T")
            
        Else
            'MODIFICAR
            
            SQL = DBSet(Text1(7).Text, "T")
            SQL = "UPDATE shcocob set observaciones = " & SQL & " WHERE codigo = " & Codigo2
        End If
        
        ConnConta.Execute SQL
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not PrimeraVez Then Exit Sub
    PrimeraVez = False
    
    If Codigo2 > 0 Then
        Set miRsAux = New ADODB.Recordset
        
        
        If vParamAplic.ContabilidadNueva Then
            SQL = RecuperaValor(Intercambio, 3)
            If SQL <> "" Then SQL = " AND numlinea =" & SQL
            SQL = " WHERE reclama.codigo =" & Codigo2 & SQL & " ORDER BY fecreclama desc ,reclama.codigo,numlinea  DESC"
            SQL = " FROM  reclama  left join reclama_facturas  on reclama.codigo=reclama_facturas.codigo" & SQL
            SQL = " if (impvenci is null, importes,impvenci) impvenci,observaciones,fecreclama " & SQL
            SQL = "select codmacta,nommacta,numserie,numfactu Codfaccl,fecfactu fecfaccl,numorden," & SQL
        Else
            SQL = "Select * from shcocob where codigo = " & Codigo2
        End If
        
        miRsAux.Open SQL, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            Text1(0).Text = miRsAux!Codmacta
            Text2(0).Text = miRsAux!Nommacta
            Text1(1).Text = DBLet(miRsAux!numSerie, "T")
            Text1(2).Text = DBLet(miRsAux!Codfaccl, "T")
            Text1(3).Text = DBLet(miRsAux!fecfaccl, "F")
            Text1(4).Text = DBLet(miRsAux!numorden, "T")
            Text1(5).Text = DBLet(miRsAux!fecreclama, "F")
            Text1(6).Text = DBLet(miRsAux!ImpVenci, "N")
            Text1(7).Text = DBLet(miRsAux!Observaciones, "T")
        Else
            MsgBox "Imposible abrir reclamacion cod: " & Codigo2, vbExclamation
            Codigo2 = -1
            limpiar Me
            Me.Command1(0).visible = False
        End If
        miRsAux.Close
    End If
    
    If Codigo2 < 0 Then
        For I = 1 To 7
            Text1(I).Text = ""
        Next I
        Text1(5).Text = Format(Now, "dd/mm/yyyy")
    Else
        If vParamAplic.ContabilidadNueva Then
            'DE momento, NO dejo modificar el ariconta la reclamacion
            Command1(0).visible = False
            Image1.visible = False
            Text1(7).Locked = True
        End If
    End If
    Text1(0).Locked = True
    For I = 1 To 6
        Text1(I).Locked = Codigo2 >= 0
    Next
    Me.imgFecha(3).visible = Codigo2 = -1
    Me.imgFecha(5).visible = Codigo2 = -1
    If Codigo2 >= 0 Then
        If vParamAplic.ContabilidadNueva Then
            PonerFocoBtn Command1(1)
        Else
            PonerFoco Text1(7)
        End If
    Else
        PonerFoco Text1(1)
    End If
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    Codigo2 = Val(RecuperaValor(Intercambio, 1))
    Me.LabelCRM.Caption = RecuperaValor(Intercambio, 2)
    If Codigo2 < 0 Then
        Text1(0).Text = RecuperaValor(Intercambio, 3)
        Text2(0).Text = RecuperaValor(Intercambio, 4)
    End If
    Image1.Picture = frmPpal.imgListComun.ListImages(46).Picture
End Sub

Private Sub Image1_Click()
    SQL = "Reclamaciones" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
    SQL = SQL & "Si esta modificando solo se le permite cambiar las observaciones." & vbCrLf
    SQL = SQL & "Si es nueva son obligatorios los campos: fecha reclamacion e importe " & vbCrLf
    MsgBox SQL, vbInformation
End Sub

Private Sub imgFecha_Click(Index As Integer)
    SQL = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    frmC.Show vbModal
    If SQL <> "" Then Text1(Index).Text = SQL
    SQL = ""
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Not Text1(Index).Locked And Index <> 7 Then ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    Select Case Index
    Case 2, 4
        If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
    
    Case 3, 5
        PonerFormatoFecha Text1(Index)
    
    Case 6
        If Not PonerFormatoDecimal(Text1(Index), 6) Then Text1(Index).Text = ""
    
    End Select
    
    
End Sub
