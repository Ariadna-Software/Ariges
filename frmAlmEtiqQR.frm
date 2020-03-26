VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmEtiqQR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR Code"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameFontenas 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdQR_Fontenas 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3840
         MaxLength       =   700
         TabIndex        =   0
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdQR_Fontenas 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   700
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   700
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   3810
         Left            =   6480
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   5
         Top             =   240
         Width           =   3810
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   0
         Left            =   240
         MaxLength       =   700
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Etiquetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblInd 
         Caption         =   "Indicador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   3780
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Referencia"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "LOTE"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblEan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   3255
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAlmEtiqQR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'AutorQR:  Leandro Ascierto
'Web:    leandroascierto.com
'Date:   09/09/2011
'----------------------------------



Public Opcion As Byte
    '   0.- Etiquestas envasado FONTENAS




Public Datos As String



Dim PrimVez As Boolean






Dim cQrCode As ClsQrCode


Private Function Guardar() As Boolean
    On Error GoTo eG
    
    lblInd(0).Caption = "Preparando IMG"
    lblInd(0).Refresh

    
    Guardar = False
    
    If Dir(App.Path & "\p1.jpg", vbArchive) <> "" Then Kill App.Path & "\p1.jpg"
    
    SavePicture Picture1.Picture, App.Path & "\p1.jpg"
    Guardar = True
    
    Exit Function
eG:
    MuestraError Err.Number, , Err.Description
End Function

Private Sub cmdQR_Fontenas_Click(index As Integer)
    
    If index = 0 Then
        If Not InsertarTmp Then Exit Sub
        lblInd(0).Caption = "Abriendo rpt"
        lblInd(0).Refresh
        
        
        
        With frmImprimir
            .FormulaSeleccion = "{tmpetienvas.codusu} = " & vUsu.Codigo
            .OtrosParametros = ""
            .NumeroParametros = 0
            .Titulo = "Capitulo"
            .SoloImprimir = True
            .EnvioEMail = False
            .Opcion = 2055
            .NombrePDF = ""
            
            .NombreRPT = "rEtiqQR.rpt"
            .ConSubInforme = False
            .Show vbModal
        End With
        
        
    End If
    lblInd(0).Caption = ""
    Unload Me
    
    
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        If Opcion = 0 Then PonerCamposEtiqueta
        
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
    
    Set cQrCode = New ClsQrCode
    'TxtFile.Text = App.Path & "\casquillo_gorra.jpg"
    
    
    
    lblInd(0).Caption = ""
    PrimVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cQrCode = Nothing
End Sub

Private Sub Text1_GotFocus(index As Integer)
    ConseguirFoco Text1(index), 3
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(index As Integer)
    If index = 3 Then
        If Not PonerFormatoEntero(Text1(index)) Then Text1(index).Text = "1"
    End If
End Sub




Private Sub PonerCamposEtiqueta()
    cmdQR_Fontenas(0).Enabled = False

    lblInd(0).Caption = "Obteniendo orden"
    lblInd(0).Refresh
    
    Text1(0).Text = RecuperaValor(Datos, 4)
    Text1(1).Text = RecuperaValor(Datos, 7)
    Text1(2).Text = RecuperaValor(Datos, 5)
    Text1(3).Text = RecuperaValor(Datos, 3)
    Me.lblEan.Caption = RecuperaValor(Datos, 6)
    ForzarEAN
    
    
    'Y el TAG llevara lo que quiero que pinte el QR
    Text1(0).Tag = RecuperaValor(Datos, 2) & "|" & Text1(0).Text & "|" & Text1(1).Text & "|" & RecuperaValor(Datos, 1) & "|"
    lblInd(0).Caption = "Obteniendo QR"
    lblInd(0).Refresh
    
    

    
    Picture1.Picture = cQrCode.GetPictureQrCode(Text1(0).Tag, Picture1.ScaleWidth, Picture1.ScaleHeight)
    If Picture1.Picture Is Nothing Then
        MsgBox "Error creand QR!"
    
    Else
        If Guardar Then cmdQR_Fontenas(0).Enabled = True
    End If
    
    lblInd(0).Caption = ""
    
End Sub


Private Sub ForzarEAN()
 On Error Resume Next
    lblEan.FontName = "IDAutomation.com Code39"
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function InsertarTmp() As Boolean
    
    InsertarTmp = False
    lblInd(0).Caption = "Guardando IMG"
    lblInd(0).Refresh
        'Abro parar guardar el binary
    conn.Execute "DELETE FROM tmpetienvas WHERE codusu = " & vUsu.Codigo
    Espera 0.25
    
    Datos = "INSERT INTO tmpetienvas ( codusu, secuencial, texto, lote ,cantidad ,ean ) VALUES ("
    Datos = Datos & vUsu.Codigo & "," & 1 & "," & DBSet(Text1(0).Text, "T") & "," & DBSet(Text1(1).Text, "T")
    Datos = Datos & "," & DBSet(Text1(2).Text, "T") & "," & DBSet(Me.lblEan.Caption, "T") & ")"
    If Not ejecutar(Datos, False) Then Exit Function
    
    adodc1.ConnectionString = conn
    adodc1.RecordSource = "Select * from tmpetienvas WHERE codusu =" & vUsu.Codigo
    adodc1.Refresh
'
    If adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL
        MsgBox "Error insertando imagen", vbExclamation
        Exit Function
    Else
        
        
        GuardarBinary adodc1.Recordset!img, App.Path & "\p1.jpg"
        adodc1.Recordset.Update
        
        
        
        DoEvents
        Espera 1
    End If
    
    
        
        
    'Si hay mas de una etiqueta hacemos el insert into, select * from
    lblInd(0).Caption = "Generando etiqueta"
    lblInd(0).Refresh
    
    
    For NumRegElim = 2 To Val(Me.Text1(3).Text)
        
        Datos = "INSERT INTO tmpetienvas(codusu,secuencial,texto,lote,cantidad,ean,img) "
        Datos = Datos & " select  codusu, " & NumRegElim & ",texto,lote,cantidad,ean,img from tmpetienvas WHERE "
        Datos = Datos & " codusu = " & vUsu.Codigo & " AND secuencial= 1"
        conn.Execute Datos
    Next
    
    
    
    
    
    
        
        
        
    InsertarTmp = True
    
End Function
