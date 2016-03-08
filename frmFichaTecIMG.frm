VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFichaTecIMG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Documentos. IMAGENES"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmFichaTecIMG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   7335
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   6495
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   4590
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
      Top             =   840
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   150
      MaxLength       =   40
      TabIndex        =   0
      Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   840
      Width           =   4305
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3870
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4320
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin VB.Label lblCarga2 
      Caption         =   "Label1"
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
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   5385
   End
   Begin VB.Label Label1 
      Caption         =   "Orden"
      Height          =   255
      Index           =   2
      Left            =   4620
      TabIndex        =   6
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción Fichero"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   7605
      Left            =   120
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   7260
      Left            =   300
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   6405
   End
   Begin VB.Label Label1 
      Caption         =   "Imagen"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1230
      Width           =   1455
   End
End
Attribute VB_Name = "frmFichaTecIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const CarpetaIMG = "ImgFicFT"
Public vDatos As String 'codsocio|nomsocio|y para iopcion1, nomimagen|

Public Opcion_ As Integer    ' -1:  VER
                            ' 0=insertar documento
                            ' 1= DNI
                            ' 2= Carnet manipulador
                        
                        
                            'FITO SANA
                            '201  DNI asociado
                            '202  Carnet fito

Dim InsertandoImg As Boolean
Dim PrimeraVez As Boolean


Dim IT As ListItem
'Dim Contador As Integer
Dim Fichero As String
Dim TipoDocu As Byte

Dim DirectorioInicio As String

Private Sub InsertarDesdeFichero()
Dim CADENA As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer

    Fichero = ""
    cd1.FileName = ""
    'cd1.InitDir = "c:\"   esta establecida en la funcion:
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos PDF|*.pdf|Archivos JPG|*.jpg|Archivos PNG|*.png|Archivos TIFF|*.tif"
    cd1.FilterIndex = 2
    cd1.ShowOpen
    cd1.MaxFileSize = 256
    cd1.CancelError = False
    
    If cd1.FileName = "" Then
        Exit Sub
    End If
    
    If FileLen(cd1.FileName) / 1000 > 1024 Then
        MsgBox "No se permite insertar ficheros de tamaño superior a 1 M", vbExclamation
        Exit Sub
    End If
    
    
    J = InStr(1, cd1.FileName, cd1.FileTitle)
    If J > 0 Then
        CADENA = Mid(cd1.FileName, 1, 25 - 1)
        cd1.InitDir = CADENA
    End If
    
'    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass
    InsertandoImg = True

    J = InStr(1, cd1.FileName, Chr(0))
    CADENA = cd1.FileName
    TipoDocu = 0
    If LCase(Right(cd1.FileName, 3)) = "pdf" Then TipoDocu = 1
    If LCase(Right(cd1.FileName, 3)) = "png" Then TipoDocu = 2
    If LCase(Right(cd1.FileName, 3)) = "tif" Then TipoDocu = 3
    Fichero = CADENA
        
    AcroPDF1.visible = TipoDocu = 1
    Image1.visible = TipoDocu <> 1
    
    CargarIMG (CADENA)
    InsertandoImg = False
    Screen.MousePointer = vbDefault
    
    
    If Opcion_ < 200 Then
        Text1(0).Text = Val(DevuelveDesdeBD(conAri, "max(orden)", "sfichdocs", "codclien", DBSet(RecuperaValor(vDatos, 1), "N"))) + 1
    Else
        Text1(0).Text = "0"
    End If
    
    Select Case Opcion_
    Case 1, 2
        
        If Opcion_ = 1 Then
            CADENA = "DNI"
        Else
            CADENA = "CarnetMa"
        End If
        CADENA = CADENA & "_" & Me.lblCarga2.Caption
        Text1(1).Text = CADENA
        
    Case 201, 202
        If Opcion_ = 201 Then
            CADENA = "DNI"
        Else
            CADENA = "CarnetMa"
        End If
        CADENA = CADENA & "_" & Me.lblCarga2.Caption
        Text1(1).Text = CADENA
    Case Else
        Text1(1).Text = Dir(CADENA)
    End Select
    PonerFoco Text1(1)
End Sub


Private Function CargarIMG(Archivo As String) As Boolean
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    lblCarga2.Caption = "Cargando ..."
    lblCarga2.Refresh
    CargarIMG = False
    If TipoDocu = 1 Then
    
         Me.AcroPDF1.LoadFile Archivo
    Else
        Select Case TipoDocu
        Case 1
            Me.Image1.Picture = LoadPicture(App.Path & "\pdf.dat")
        Case 2
            Me.Image1.Picture = LoadPicture(App.Path & "\png.dat")
        Case 3
            Me.Image1.Picture = LoadPicture(App.Path & "\tif.dat")
        Case Else
            Me.Image1.Picture = LoadPicture(Archivo)
        End Select
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    Else
        CargarIMG = True
    End If
    lblCarga2.Caption = lblCarga2.Tag
    Screen.MousePointer = vbDefault
End Function

Private Function InsertarImagen() As Boolean
Dim RS As ADODB.Recordset
Dim C As String
Dim L As Long
Dim Aux As String
    
    On Error GoTo eInsertarImagen
    
    InsertarImagen = False
    
  
    If Opcion_ > 200 Then
        '    codclien id
        C = "ImgDNI,DocDNI  "
        If Opcion_ = 202 Then C = "ImgManipula,DocManipula"
        C = "Select " & C & " ,codclien, id from sclienmani where codclien=" & RecuperaValor(vDatos, 1)
        C = C & " AND id =" & RecuperaValor(vDatos, 3)
    
        Aux = Dir(Fichero)
        L = InStr(1, Aux, ".")
        If L > 0 Then Aux = Mid(Aux, L)
        
        Aux = Text1(1).Text & Aux
        
    Else
        'IMAGENES sclien
        C = "Select max(codigo) from sfichdocs" '  where codsocio = " & RecuperaValor(vDatos, 1)
        Set RS = New ADODB.Recordset
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        L = 0
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then L = RS.Fields(0)
        End If
        L = L + 1
        RS.Close
        
        
    
        ' es nuevo
        C = "insert into sfichdocs (codigo, codclien, descripfich, orden, docum,TipoDoc) values"
        C = C & " (" & DBSet(L, "N") & "," & RecuperaValor(vDatos, 1) & "," & DBSet(Me.Text1(1).Text, "T") & "," & DBSet(Text1(0).Text, "N") & ","
        C = C & DBSet(Dir(Fichero), "T") & "," & Opcion_ & ")"
        conn.Execute C
        
        Espera 0.2
        
        'Abro parar guardar el binary
        C = "Select campo,codigo from sfichdocs where codigo =" & L
        
    End If
    
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = C
    Adodc1.Refresh
'
    If Adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        InsertandoImg = True
        CargarIMG Fichero 'lw1.ListItems(k).SubItems(2)
        If Opcion_ < 200 Then
            GuardarBinary Adodc1.Recordset!campo, Fichero
        Else
            GuardarBinary Adodc1.Recordset.Fields(0), Fichero
            Adodc1.Recordset.Fields(1).Value = Aux
        End If
        Adodc1.Recordset.Update
        DoEvents
        Espera 1
    End If


    CadenaDesdeOtroForm = L
    InsertarImagen = True
    Exit Function
    
eInsertarImagen:
    MuestraError Err.Number, "Insertar Imágen", Err.Description
End Function


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()

    If Text1(1).Text = "" Then
        MsgBox "Debe introducir una descripción de Fichero. Reintroduzca.", vbExclamation
        PonerFoco Text1(1)
        Exit Sub
    End If
    
    If Text1(0).Text = "" Then
        MsgBox "Debe introducir el orden de la imágen en la lista del socio. Reintroduzca.", vbExclamation
        PonerFoco Text1(0)
        Exit Sub
    End If

    If Opcion_ < 200 Then
        Text1(1).Tag = "docum = " & DBSet(Me.Text1(1).Text, "T") & " AND codclien"
        Text1(1).Tag = DevuelveDesdeBD(conAri, "docum", "sfichdocs", Text1(1).Tag, RecuperaValor(vDatos, 1))
        If Text1(1).Tag <> "" Then
            MsgBox "Ya existe una descripcion igual para el cliente", vbExclamation
            PonerFoco Text1(0)
            Exit Sub
        End If
    End If
    
    If InsertarImagen Then
        
    
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass

        
        If Opcion_ = -1 Then
            CargarIMG (RecuperaValor(vDatos, 3))
            
        Else
            InsertarDesdeFichero
        End If
        
        Me.cmdGuardar.visible = Opcion_ >= 0
        Me.Text1(0).Enabled = cmdGuardar.visible
        Me.Text1(1).Enabled = cmdGuardar.visible
        
        lblCarga2.Caption = lblCarga2.Tag
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    Me.lblCarga2.Tag = RecuperaValor(Me.vDatos, 2)
    lblCarga2.Caption = "Leyendo datos BD"
    If Opcion_ >= 0 Then
        Me.Text1(1).Text = ""
        Me.Text1(0).Text = ""
    End If
    
    FicjarInitDir True
End Sub


Private Sub FicjarInitDir(Leer As Boolean)
Dim I As Integer
    
    
    
    If Leer Then
    
        DirectorioInicio = "C:\"
        If Opcion_ <> 0 Then
            
            I = FreeFile
            Fichero = App.Path & "\FichInitD.xdf"
            If Dir(Fichero, vbArchive) <> "" Then
                Open Fichero For Input As #I
                Line Input #I, Fichero
                Close #I
                
                If Trim(Fichero) <> "" Then
                    If Dir(Fichero, vbDirectory) <> "" Then DirectorioInicio = Fichero
                End If
                
            
            End If
        End If
        cd1.InitDir = DirectorioInicio

    Else
        
        If cd1.InitDir <> DirectorioInicio Then
            I = FreeFile
            Fichero = App.Path & "\FichInitD.xdf"
         
            Open Fichero For Output As #I
            Fichero = cd1.InitDir
            Print #I, Fichero
            Close #I
    
        End If
        

    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    FicjarInitDir False
End Sub

'Private Sub Imprimir()
'        With frmImprimir
'            .FormulaSeleccion = "{rsocios.codsocio}=" & RecuperaValor(vDatos, 1)
'            .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
'            .Titulo = "Imágenes adjuntas"
'            .NumeroParametros = 1
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .NombreRPT = "rImgDocs.rpt"
'
'            .Opcion = 2015
'            .Show vbModal
'        End With
'End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        Unload Me
    End If
End Sub

