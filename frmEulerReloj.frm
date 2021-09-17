VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEulerReloj 
   Caption         =   "Reloj"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18990
   Icon            =   "frmEulerReloj.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   18990
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAlbaran 
      Caption         =   "Nuevo"
      Height          =   495
      Index           =   0
      Left            =   10440
      Picture         =   "frmEulerReloj.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Generar albaran"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprimir listado"
      Top             =   6120
      Width           =   615
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3495
      Left            =   6960
      TabIndex        =   7
      Top             =   1200
      Width           =   11775
      Begin VB.ComboBox cboVehiculo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2640
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         TabIndex        =   19
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   9015
      End
      Begin VB.ComboBox cboAlb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   11655
      End
      Begin VB.ComboBox cboTipoTrabajo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmEulerReloj.frx":0894
         Left            =   1680
         List            =   "frmEulerReloj.frx":0896
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label5 
         Caption         =   "Veh�culo/Maquinaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   21
         Top             =   2280
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Tarea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Albar�n / Orden producci�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   6840
      TabIndex        =   5
      Top             =   240
      Width           =   10455
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   8640
         Picture         =   "frmEulerReloj.frx":0898
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpiar"
         Top             =   120
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "  Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9240
         TabIndex        =   18
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5175
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tr"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Numero"
         Object.Width           =   2565
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tarea"
         Object.Width           =   8828
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Incio"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fin"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cliente/PRod"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Id Registro"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Ref."
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Matr"
         Object.Width           =   3422
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   12720
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "SALIR"
      Height          =   495
      Left            =   14160
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   56.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   6240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmEulerReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Soloreloj = False

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim Cad As String
Dim T1 As Date


Dim HayQueCerrarNodo As Integer  'Marcara el IT que tiene que actualizar como HORAFIn, 0=Inicio
Dim UltimaLecturaReloj As Date


Dim UltimaTareaLeida As Date

Dim TrabajadoresZaldibia As String

Dim NumeroTareasPendientesCerrar As Integer





Private Sub cmdAlbaran_Click(Index As Integer)
    Cad = ""
    If Combo1.ListIndex < 0 Then Cad = "Seleccione el trabajador "
        
    If Me.cboTipoTrabajo.ListIndex <= 0 Then Cad = "Seleccione el tipo de albaran"
    'If Me.cboTipoTrabajo.ListIndex = 4 Then cad = "No puede seleccionar produccion"
        
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If

    
    
    
    Cad = "  " & UCase(Me.cboTipoTrabajo.Text)
        
    Cad = "Desea crear el albar�n del tipo " & Cad & "?"
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    CrearAlbaran
End Sub

Private Sub cmdBuscarCli_Click(Index As Integer)
    
    
    If Index = 0 Then
        'Nusca grid
        
    
    Else
        'fechas
        
        
    End If
End Sub



Private Sub cmdImprimir_Click()
    'VERSION RELOJ: comentar lineas #Soloreloj
   ' frmListado2.opcion = 46
   ' frmListado2.Show vbModal
End Sub



Private Sub Combo1_Click()
Dim J As Integer
Dim EsTrabajadorZaldibia As Boolean

     If Me.Command1.Tag = 1 Then Exit Sub
    
    If Combo1.ListIndex < 0 Then Exit Sub
    
    
    J = 0
    If Val(Combo1.Tag) >= 0 Then
        '�Si no esta marcado varios
        If Me.Check1.Value = 0 Then
            If Combo1.ListIndex <> Combo1.Tag Then J = 1
        End If
    End If
    'Si ya habia seleccionado un trabjador antoeriormente, volvemos a cargar datos
    If J = 1 Then CagarMarcajes
    Combo1.Tag = Combo1.ListIndex
    
    
    'Borramos nodos no trabajador
    
    Cad = "|" & Combo1.ItemData(Combo1.ListIndex) & "|"
  
    EsTrabajadorZaldibia = InStr(1, "|" & TrabajadoresZaldibia, Cad) > 0
    
    'Si marca todos, no hacemos nada
    If Me.Check1.Value = 0 Then
    
        For J = Me.ListView2.ListItems.Count To 1 Step -1
            If InStr(1, "|" & TrabajadoresZaldibia, "|" & ListView2.ListItems(J).Text & "|") > 0 Then
                If Not EsTrabajadorZaldibia Then ListView2.ListItems.Remove J
            Else
                If EsTrabajadorZaldibia Then ListView2.ListItems.Remove J
            End If
        Next
    End If
    Cad = "|" & TrabajadoresZaldibia

    BuscarNodo True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If Me.Command1.Tag = 1 Then Exit Sub
    
    If Combo1.ListIndex < 0 Then Exit Sub

    
    BuscarNodo True
    
End Sub

Private Sub BuscarNodo(DesdeTrabajador As Boolean)
Dim K As Integer
Dim NodosFechasAnteriores As Boolean
Dim fin As Boolean
Dim NodoCierreAnterior As Integer
Dim Horas As Currency
Dim F1 As Date

    HayQueCerrarNodo = 0
    NodosFechasAnteriores = False
    fin = False
    
    Do
            For K = ListView2.ListItems.Count To 1 Step -1
                ListView2.ListItems(K).ForeColor = vbBlack
                ListView2.ListItems(K).Bold = False
                If Combo1.ListIndex >= 0 Then
                    If Val(ListView2.ListItems(K).Text) = Val(Combo1.ItemData(Combo1.ListIndex)) Then
                        
                        ListView2.ListItems(K).Bold = True
                        If Trim(ListView2.ListItems(K).SubItems(6)) = "" Then
                            HayQueCerrarNodo = K
                            ListView2.ListItems(K).ForeColor = vbBlue
                            
                            'Fechas anteriores
                            If ListView2.ListItems(K).ListSubItems(5).ForeColor = vbRed Then
                                NodosFechasAnteriores = True
                                NodoCierreAnterior = K
                            End If
                        End If
                    End If
                End If
            Next
            
            If Not NodosFechasAnteriores Then
                fin = True
            Else
                
                
                'El nodo es del dia anterior,   Vamos a cerrar con la fecha de ayer
                Cad = "Trabajador : " & ListView2.ListItems(NodoCierreAnterior).ListSubItems(1).Text & vbCrLf
                Cad = Cad & "FECHA : " & ListView2.ListItems(NodoCierreAnterior).ListSubItems(5).ToolTipText & vbCrLf
                Cad = Cad & "hora inicio : " & ListView2.ListItems(NodoCierreAnterior).ListSubItems(5).Text
                Cad = InputBox(Cad, "Cierre dia anterior", "23:45")
                If Cad = "" Then
                    'Fin = True
                    
                Else
                    Cad = Trim(Replace(Cad, ".", ":"))
                    If Not EsHoraOK(Cad) Then
                        MsgBox "Formato hora incorrecto", vbExclamation
                        Cad = ""
                    Else
                    
                    End If
                End If
                If Cad = "" Then
                    fin = True
                    Command2_Click
              
                Else
                    'UPDATEAMOS
                    
                    RedondeaLaHora Cad
                    
                    
                    Cad = ListView2.ListItems(NodoCierreAnterior).ListSubItems(5).ToolTipText & " " & Cad
                    
                   
                    F1 = CDate(ListView2.ListItems(NodoCierreAnterior).ListSubItems(5).ToolTipText & " " & ListView2.ListItems(NodoCierreAnterior).Tag)
                    
        
                    'lo paso a segundos y divido por 3600
                    Horas = DateDiff("s", CDate(Cad), F1)
                    Horas = Abs(Round2(Horas / 3600, 2))
                    
                    'Cad = "UPDATE sreloj SET HoraFin =" & DBSet(Label1(0).Caption & " " & Label1(1).Caption, "FH")
                    
                    Cad = "UPDATE sreloj SET HoraFin =" & DBSet(Cad, "FH")
                    
                    Cad = Cad & " ,calculadas=" & DBSet(Horas, "N")
                    Cad = Cad & " WHERE codtraba = " & Combo1.ItemData(Combo1.ListIndex) & " AND fecha = "
                    Cad = Cad & DBSet(F1, "F")
                    Cad = Cad & " AND HoraInicio = " & DBSet(F1, "FH")
                    
                    If ejecutar(Cad, False) Then
                        ListView2.ListItems.Remove NodoCierreAnterior
                        fin = True
                    End If
                End If
            End If
            
      Loop Until fin
    PonerFrames2
    
End Sub


Private Sub cboTipoTrabajo_Click()
Dim OrdenProduccion As Boolean
Dim Aux As String
Dim Cad2 As String
    OrdenProduccion = False
    
    
    If cboTipoTrabajo.ListIndex <= 0 Then
        cboAlb.Clear
        Combo4.Clear
        cboVehiculo.ListIndex = -1
        Exit Sub
    End If
    
    Cad = "ALR"
    If cboTipoTrabajo.ListIndex = 2 Then
        Cad = "ALE"
    ElseIf cboTipoTrabajo.ListIndex = 3 Then
        Cad = "ALO"
        
        
    ElseIf cboTipoTrabajo.ListIndex = 4 Then
        'Orden produccion
        'OrdenProduccion = True
        Cad = "ALV"
        
        
    End If
    
    If OrdenProduccion Then
    
            
            Cad = Format(DateAdd("yyyy", -1, Now), FormatoFecha)
            CargarCombo_Tabla Me.cboAlb, "sordprod", "concat(right(concat(""000000"",codigo),6),' - ',coalesce(descripcion,feccreacion))", "codigo", "feccreacion>='" & Cad & "'"
    
    Else
    
    
    
            
                Aux = ""
                If Combo1.ListIndex >= 0 Then
                    If Me.Check1.Value = 0 Then
                        If TrabajadoresZaldibia <> "" Then
                            Aux = "|" & Combo1.ItemData(Combo1.ListIndex) & "|"
                            If InStr(1, "|" & TrabajadoresZaldibia, Aux) > 0 Then
                                Aux = ""
                            Else
                                Aux = "NOT"
                            End If
                        
                            Cad2 = Mid(TrabajadoresZaldibia, 1, Len(TrabajadoresZaldibia) - 1) 'quito el ultimo pipe
                            Aux = " AND " & Aux & " codtraba IN (" & Replace(Cad2, "|", ",") & ")"
                        End If
                    End If
                End If
                
                Cad = "codtipom = '" & Cad & "' AND  (origdat is null or origdat<>2)" & Aux
                CargarCombo_Tabla Me.cboAlb, "scaalb", "concat(numalbar,' - ',nomclien,if(coalesce(referenc ,'')='','',concat('  [',referenc,']')))", "NumAlbar", Cad, , "fechaalb desc,numalbar desc"
          
    End If
            
    Cad = "R"
    If cboTipoTrabajo.ListIndex = 2 Then
        Cad = "E"
    ElseIf cboTipoTrabajo.ListIndex = 3 Then
        Cad = "O"
    ElseIf cboTipoTrabajo.ListIndex = 4 Then
        Cad = "V"
        
        
        
    End If
    Cad = "codtipor like '" & Cad & "_'"
    CargarCombo_Tabla Me.Combo4, "stipor", "concat(nomtipor,' [',codtipor,']')", "1", Cad

    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim F1 As Date
Dim linea As Integer
Dim Horas As Currency
Dim C As String

    'Por si las moscas
    
    On Error GoTo EC
    
    If Combo1.ListIndex < 0 Then Exit Sub
    If cboTipoTrabajo.ListIndex < 0 Then Exit Sub

    
    
    If cboTipoTrabajo.ListIndex < 0 Then Exit Sub
 
         
    If cboTipoTrabajo.ListIndex = 0 Then
      'Hay que cerrar nodo.
        If HayQueCerrarNodo = 0 Then
            MsgBox "Ninguna tarea iniciada para el trabajador", vbExclamation
            Exit Sub
        End If
        
       
    
    Else
        'No ha asiganado tarea
        If Combo4.ListIndex < 0 Then Exit Sub
        
    End If          'de cerrar tarea
        
           
    'Hay que cerrar nodo
    If HayQueCerrarNodo > 0 Then
        If HayQueCerrarNodo > ListView2.ListItems.Count Then
            MsgBox "Error situando datos trabajador", vbExclamation
            Exit Sub
        End If
        
        'Horas
        'Debug.Print ListView2.ListItems(HayQueCerrarNodo).ListSubItems(5).Text
        If ListView2.ListItems(HayQueCerrarNodo).ListSubItems(5).ToolTipText = "" Then
            Cad = Label1(0).Caption
        Else
            Cad = ListView2.ListItems(HayQueCerrarNodo).ListSubItems(4).ToolTipText
        End If
        Cad = Cad & " " & ListView2.ListItems(HayQueCerrarNodo).Tag
        F1 = CDate(Cad)
        
        'Mayo 2019
        'Horas = DateDiff("n", F1, CDate(Label1(0).Caption & " " & Label1(1).Caption))
        'lo paso a segundos y divido por 3600
        Horas = DateDiff("s", F1, CDate(Label1(0).Caption & " " & Label1(1).Caption))
        
        
        
        
        If Horas < 0 Then
            MsgBox "Calculos negativos!!!", vbExclamation
            Horas = 0
        End If
        
        'En horas, llevamos los minutos. Ahora lo pasamos a horas en decimal
        'MAyo 19 -- ahora son segundos
        'linea = Horas Mod 60  'los minutos que exende de la hora
        'Horas = Horas \ 60
        'Horas = Horas + Round((linea / 60), 2)
        Horas = Round2(Horas / 3600, 2)
        
        
        
        
       ' Fecha codtraba HoraInicio
        Cad = "UPDATE sreloj SET HoraFin =" & DBSet(Label1(0).Caption & " " & Label1(1).Caption, "FH")
        Cad = Cad & " ,calculadas=" & DBSet(Horas, "N")
        Cad = Cad & " WHERE codtraba = " & Combo1.ItemData(Combo1.ListIndex) & " AND fecha = " & DBSet(F1, "F")
        Cad = Cad & " AND HoraInicio = " & DBSet(F1, "FH")
                        
        conn.Execute Cad
    End If
    
    'Insertamios la nueva, si es que hay que insertarla
    If Me.cboTipoTrabajo.ListIndex > 0 Then
        C = Combo4.Text
        NumRegElim = InStr(1, C, "[")
        If NumRegElim = 0 Then Err.Raise 513, , "MAl"
            
        C = Mid(C, NumRegElim + 1)
        C = Mid(C, 1, Len(C) - 1)
        
        Cad = ",'" & C & "')"
       
        C = cboAlb.Text
        NumRegElim = InStr(1, C, " - ")
        If NumRegElim = 0 Then Err.Raise 513, , "MAl 2"
        C = Trim(Mid(C, 1, NumRegElim))
        
        Cad = "," & C & Cad 'Concatenamos
        C = "'ALR'"
        If cboTipoTrabajo.ListIndex = 2 Then
            C = "'ALE'"
        ElseIf cboTipoTrabajo.ListIndex = 3 Then
            C = "'ALO'"
        ElseIf cboTipoTrabajo.ListIndex = 4 Then
            C = "'ALV'"
        End If
        Cad = "," & C & Cad
        
        Cad = DBSet(Label1(0).Caption, "F") & "," & Combo1.ItemData(Combo1.ListIndex) & "," & DBSet(Label1(0).Caption & " " & Label1(1).Caption, "FH") & ",null,0" & Cad
        C = Trim(cboVehiculo.Text)
        Cad = DBSet(C, "T", "S") & "," & Cad
        
        C = DevuelveDesdeBD(conAri, "max(id)", "sreloj", "1", "1")
        C = str(Val(C) + 1)
        Cad = C & "," & Cad
        Cad = "INSERT INTO sreloj(ID,codflota,Fecha,codtraba,HoraInicio,HoraFin,Calculadas,codtipom,numalbar,codtipor) VALUES (" & Cad
       
        conn.Execute Cad
            
    End If
    
    
    cboAlb.Clear
    Combo4.Clear
    cboVehiculo.ListIndex = -1
    Combo1.ListIndex = -1
    Combo1.Tag = -1
    cboTipoTrabajo.ListIndex = -1
    CagarMarcajes
    
    
EC:
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub Command2_Click()
    limpiar Me
    Me.cboTipoTrabajo.ListIndex = -1
    Me.Combo1.ListIndex = -1
    Combo1.Tag = -1
  
    
    CagarMarcajes
End Sub

Private Sub Command3_Click()
    
End Sub

Private Sub Form_Activate()
    
    If Me.Command1.Tag = 1 Then
    
        Me.Command1.Tag = 0
        
        
        
        Set miRsAux = New ADODB.Recordset
        Label1(1).Caption = "Leyendo"
        Label1(1).Refresh
        
        TrabajadoresZaldibia = ""
        miRsAux.Open "SELECT * from straba where codalmac=10", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            TrabajadoresZaldibia = TrabajadoresZaldibia & miRsAux!CodTraba & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        

        
            
            
        miRsAux.Open "Select now()", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        T1 = miRsAux.Fields(0)
        miRsAux.Close
        Labels
            
            
        miRsAux.Open "Select distinct fecha from sreloj where horafin is null and fecha<" & DBSet(Label1(1).Caption, "F"), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        HayQueCerrarNodo = 0
        While Not miRsAux.EOF
            HayQueCerrarNodo = HayQueCerrarNodo + 1
            Cad = Cad & Format(miRsAux!Fecha, "dd/mm/yyyy") & "     "
            If (HayQueCerrarNodo Mod 3) = 0 Then Cad = Cad & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
         If Cad <> "" Then MsgBox "Dias por finalizar tareas: " & vbCrLf & Cad, vbExclamation
            
            
            
            
        Labels
        
        'Cargo vehiculos Cad = "codtipor like '" & Cad & "_'"
        CargarCombo_Tabla Me.cboVehiculo, "sflotas", "codflota", "codflota", "", True
        
        
        
        'Cargo listitiem
        HayQueCerrarNodo = 0
        CagarMarcajes

        
        
        
        PonerFrames2
       
        
        Timer1.Enabled = True

        
        
    End If
End Sub




Private Sub Form_Load()
    'Me.Icon = frmPpal.Icon   formulario compartido de proyecto
    
    'Cargar Trabajadores
    Cad = "fechabaj is null and codagent1 >=0 "
    CargarCombo_Tabla Me.Combo1, "straba", "codtraba", "nomtraba", Cad
    
    'Carga combo trabajos
    'ALR|ALE|ALO|
    cboTipoTrabajo.AddItem "** Salida empresa **"
    cboTipoTrabajo.AddItem "Reparaci�n"
    cboTipoTrabajo.AddItem "T. exterior"
    cboTipoTrabajo.AddItem "Orden de trabajo"
    cboTipoTrabajo.AddItem "Albaran venta"
    
    '
    'cboTipoTrabajo.AddItem "Producci�n"   'orden de produccion

    
    Me.Command1.Tag = 1
    Combo1.Tag = -1
  
    Caption = "Reloj"
    If Soloreloj Then Caption = Caption & "    ver: " & App.Major & "." & App.Minor & "." & App.Revision
    
    CargaImangenBtn
    UltimaLecturaReloj = Now
End Sub


Private Sub CargaImangenBtn()
'#Soloreloj

    If Soloreloj Then
        'Version SOLORELOJ. Comentar en NORMAL
        Me.cmdImprimir.visible = False
        cmdAlbaran(0).Picture = Nothing
        cmdAlbaran(0).Caption = "Generar"
    Else
        'Version normal
        Me.cmdImprimir.Picture = frmPpal.imgListComun.ListImages(16).Picture
        Me.cmdAlbaran(0).Picture = frmPpal.ImgListPpal.ListImages(10).Picture
        Me.cmdAlbaran(0).Caption = ""
    End If
End Sub

Private Sub Form_Resize()
Dim H As Long
    If WindowState = 1 Then Exit Sub ' ha pulsado minimizar
    
    H = Me.Width - (ListView2.Left + 600)
    If H < 0 Then H = 6375
    ListView2.Width = H
    
    
    H = Me.Width - 7500
    If H < 300 Then H = 11655
    Me.cboAlb.Width = H - 200
    Me.Frame4.Width = H
    
    
    
    ListView2.ColumnHeaders(10).Width = IIf(Me.Width > 24000, 4500, 2300)
    
    
    H = Me.Height - 1200 - ListView2.Top
    If H < 0 Then H = 6375
    ListView2.Height = H
    
    
    
    
    
    Command1.Top = Me.Height - 1140
    cmdAceptar.Top = Command1.Top
    cmdImprimir.Top = Command1.Top
    Me.cmdAlbaran(0).Top = Command1.Top
    
    
    Command1.Left = Me.Width - 1800
    cmdAceptar.Left = Command1.Left - 320 - cmdAceptar.Width
    
    
    
    If Soloreloj Then
        'SOLO RELOJ  #Soloreloj
        cmdImprimir.Left = cmdAceptar.Left
    
    Else
        'NORMAL
        cmdImprimir.Left = cmdAceptar.Left - 560 - cmdImprimir.Width
        Me.cmdAlbaran(0).Left = cmdImprimir.Left - 560 - cmdAlbaran(0).Width
    End If
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub ListView2_DblClick()
Dim i As Integer

    
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    
    For i = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(i) = Val(ListView2.SelectedItem.Text) Then
            'Este es el trabajador
            Combo1.ListIndex = i
            Exit For
        End If
    Next
    
    If ListView2.SelectedItem Is Nothing Then Exit Sub
       
    Select Case UCase(Trim(ListView2.SelectedItem.SubItems(2)))
    Case "ALE"
        i = 2
    Case "ALO"
        i = 3
    Case Else
        i = 1
    End Select
    cboTipoTrabajo.ListIndex = i
    DoEvents
    Espera 0.1
    
    For i = 0 To cboAlb.ListCount - 1
        Cad = cboAlb.List(i)
        Cad = Mid(Cad, 1, InStr(1, Cad, "-") - 1)
        
        If Val(Cad) = Val(ListView2.SelectedItem.SubItems(3)) Then
            'Este es el trabajador
            cboAlb.ListIndex = i
            Exit For
        End If
    Next
    
    
    'Matricula
    
    'El primero esta vacio
    For i = 1 To cboVehiculo.ListCount - 1
        Cad = UCase(cboVehiculo.List(i))
        
        
        If UCase(Cad) = UCase(Trim(ListView2.SelectedItem.SubItems(10))) Then
            'Este es el trabajador
            cboVehiculo.ListIndex = i
            Exit For
        End If
    Next
    
    
    
 
    ListView2.Tag = Trim(Mid(ListView2.SelectedItem.SubItems(4), 1, InStr(2, ListView2.SelectedItem.SubItems(4), " ")))
    For i = 0 To Combo4.ListCount - 1
        Cad = Trim(Combo4.List(i))
        
        Cad = Mid(Cad, InStr(1, Cad, "[") + 1)  'quitamos primer corchete
        Cad = Mid(Cad, 1, Len(Cad) - 1)  'quitamos segundo corchete
        
        If Cad = ListView2.Tag Then
            'Este es el trabajador
            Combo4.ListIndex = i
            Exit For
        End If
    Next
    ListView2.Tag = ""
    
End Sub


Private Sub Timer1_Timer()
Dim Hacer As Boolean
    
    
    T1 = DateAdd("s", 1, T1)
    
    Hacer = True
    If Me.Combo1.ListIndex >= 0 Then Hacer = False
    If Me.cboTipoTrabajo.ListIndex >= 0 Then Hacer = False
    If Me.cboAlb.ListIndex >= 0 Then Hacer = False
    If Combo4.ListIndex >= 0 Then Hacer = False
    
    
    If Hacer Then
        If DateDiff("s", UltimaLecturaReloj, T1) >= 25 Then
            
            
            
            
            
            
            Label1(1).Caption = "Leyendo"
            Label1(1).Refresh
            If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
            Cad = "HoraFin is null AND 1 "
            Cad = DevuelveDesdeBD(conAri, "count(*)", "sreloj", Cad, "1")
            If Val(Cad) <> NumeroTareasPendientesCerrar Then
                Espera 0.5
                CagarMarcajes
            Else
                miRsAux.Open "Select now()", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                T1 = miRsAux.Fields(0)
                miRsAux.Close
                UltimaLecturaReloj = T1
            End If
    

            
            
        End If
    End If
    Labels
End Sub

Private Sub Labels()
    Me.Label1(0).Caption = Format(T1, "dd/mm/yyyy")
    Me.Label1(1).Caption = Format(T1, "hh:mm:ss")
    
    
    
End Sub


Private Sub CagarMarcajes()
Dim IT As ListItem
Dim i As Integer

    limpiar Me

    Set miRsAux = New ADODB.Recordset
    Me.ListView2.ListItems.Clear
    
    
    NumeroTareasPendientesCerrar = 0

    'Cargariamos las anteriores que no esten cerradas
    Cad = "select sreloj.*,nomtraba,nomclien,nomtipor,referenc from sreloj inner join straba on sreloj.codtraba=straba.codtraba"
    Cad = Cad & " LEFT JOIN scaalb ON scaalb.codtipom = sreloj.codtipom AND sreloj.numalbar=scaalb.numalbar"
    Cad = Cad & " LEFT JOIN stipor ON sreloj.codtipor = stipor.codtipor"
    Cad = Cad & " where  fecha<" & DBSet(Label1(0).Caption, "F") & " AND horafin is null"
    Cad = Cad & " ORDER BY fecha,HoraInicio"
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!CodTraba
        IT.SubItems(1) = miRsAux!NomTraba
        
        IT.SubItems(2) = DBLet(miRsAux!codtipom, "T") & " "
        IT.SubItems(3) = DBLet(miRsAux!Numalbar, "T") & " "
        
        If IsNull(miRsAux!NomTipor) Then
            Cad = "--- ** No encotrado"
        Else
            Cad = miRsAux!NomTipor
        End If
        Cad = miRsAux!codtipor & " " & Cad
        IT.SubItems(4) = Cad
        IT.SubItems(5) = Format(miRsAux!horainicio, "hh:mm")
        IT.Tag = Format(miRsAux!horainicio, "hh:mm:ss")
        
        
        If IsNull(miRsAux!HoraFin) Then
            IT.SubItems(6) = " "
            IT.ListSubItems(5).ForeColor = vbRed
            
        Else
            IT.SubItems(6) = Format(miRsAux!HoraFin, "hh:mm")
        End If
        
        
        If IsNull(miRsAux!codtipom) Then
            'ORden de produccion
            IT.SubItems(7) = "Prod. " & Format(miRsAux!Numalbar, "000000")
            IT.SubItems(9) = " "
        Else
        
            IT.SubItems(7) = DBLet(miRsAux!NomClien, "T") & " "
            IT.SubItems(9) = DBLet(miRsAux!referenc, "T") & " "
        End If
        IT.SubItems(10) = DBLet(miRsAux!codflota, "T") & " "
        
        
        For i = 1 To IT.ListSubItems.Count
            IT.ListSubItems(i).ToolTipText = Format(miRsAux!horainicio, "dd/mm/yyyy")
        Next
        IT.ListSubItems(1).ToolTipText = miRsAux!NomTraba
        
        
        NumeroTareasPendientesCerrar = NumeroTareasPendientesCerrar + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    Cad = "select sreloj.*,nomtraba,nomclien,nomtipor,referenc from sreloj inner join straba on sreloj.codtraba=straba.codtraba"
    Cad = Cad & " LEFT JOIN scaalb ON scaalb.codtipom = sreloj.codtipom AND sreloj.numalbar=scaalb.numalbar"
    Cad = Cad & " LEFT JOIN stipor ON sreloj.codtipor = stipor.codtipor"
    Cad = Cad & " where fecha=" & DBSet(Label1(0).Caption, "F")
    Cad = Cad & " ORDER BY HoraInicio,HoraFin"
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumeroTareasPendientesCerrar = NumeroTareasPendientesCerrar + 1
        
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!CodTraba
        IT.SubItems(1) = miRsAux!NomTraba
        
        IT.SubItems(2) = DBLet(miRsAux!codtipom, "T") & " "
        IT.SubItems(3) = DBLet(miRsAux!Numalbar, "T") & " "
        
        If IsNull(miRsAux!NomTipor) Then
            Cad = "--- ** No encotrado"
        Else
            Cad = miRsAux!NomTipor
        End If
        Cad = miRsAux!codtipor & " " & Cad
        IT.SubItems(4) = Cad
        IT.SubItems(5) = Format(miRsAux!horainicio, "hh:mm")
        IT.Tag = Format(miRsAux!horainicio, "hh:mm:ss")
        
        If IsNull(miRsAux!HoraFin) Then
            IT.SubItems(6) = " "
        Else
            IT.SubItems(6) = Format(miRsAux!HoraFin, "hh:mm")
        End If
        
        
        If IsNull(miRsAux!codtipom) Then
            'ORden de produccion
            IT.SubItems(7) = "Prod. " & Format(miRsAux!Numalbar, "000000")
            IT.SubItems(9) = " "
        Else
        
            IT.SubItems(7) = DBLet(miRsAux!NomClien, "T") & " "
            IT.SubItems(9) = DBLet(miRsAux!referenc, "T") & " "
        End If
        IT.SubItems(10) = DBLet(miRsAux!codflota, "T") & " "
        IT.ListSubItems(1).ToolTipText = miRsAux!NomTraba
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Not IT Is Nothing Then IT.EnsureVisible
    
    
    
    
    miRsAux.Open "Select now()", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    T1 = miRsAux.Fields(0)
    miRsAux.Close
    UltimaLecturaReloj = T1
    
    
    
    
    
End Sub


Private Sub PonerFrames2()
Dim B As Boolean

    


    
    If Combo1.ListIndex < 0 Then
        B = False
    Else
        B = HayQueCerrarNodo = 0
    End If
    
    Frame4.visible = True
    
    
    
End Sub


Private Sub CrearAlbaran()
Dim vC As CTiposMov
Dim vCli As CCliente
Dim HaCambiadoContador As Boolean
    On Error GoTo eCrearAlbaran
    
    Set vC = New CTiposMov
    Set vCli = New CCliente
    
    
    conn.BeginTrans

    
    'VErsion normal
    Cad = DBSet(vParam.CifEmpresa, "T") & " ORDER BY codclien"
    

    
    Cad = DevuelveDesdeBD(conAri, "codclien", "sclien", "nifclien", Cad)
    If Cad = "" Then Err.Raise 513, , "Obteniendo cliente EULER"
    
    If Not vCli.LeerDatos(Cad) Then Err.Raise 513, , "Obteniendo datos cliente EULER " & Cad
    
    'Febrero 2016
    'Si es EUSKADI o VALENCIA para los albaranes de reparacion cogera el CAR o el ALR
    
    Select Case Me.cboTipoTrabajo.ListIndex
    Case 2
        Cad = "ALE"
    Case 3
        Cad = "ALO"
    Case 4
        Cad = "ALV"
    Case Else
            
            Cad = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", CStr(Combo1.ItemData(Combo1.ListIndex)))
            If Cad = "10" Then
                Cad = "CAR"
            Else
                Cad = "ALR"
            End If
    End Select
    
    vC.ConseguirContador Cad
    
    Cad = "INSERT INTO scaalb(codtipom,numalbar,fechaalb,factursn,codclien,"
    Cad = Cad & "nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    Cad = Cad & "facturkm,codtraba,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,esticket,fechaAux) VALUES ('"
    If Me.cboTipoTrabajo.ListIndex = 1 Then
        'ALR
         Cad = Cad & "ALR"
    Else
        Cad = Cad & vC.TipoMovimiento
    End If
    Cad = Cad & "'," & vC.Contador + 1 & "," & DBSet(Label1(0).Caption, "F") & ",1," & vCli.Codigo
    Cad = Cad & "," & DBSet(vCli.Nombre, "T") & "," & DBSet(vCli.Domicilio, "T") & "," & DBSet(vCli.CPostal, "T") & "," & DBSet(vCli.Poblacion, "T")
    Cad = Cad & "," & DBSet(vCli.Provincia, "T") & "," & DBSet(vCli.NIF, "T") & "," & DBSet(vCli.TfnoClien, "T", "S") & ",0," & Val(Combo1.ItemData(Combo1.ListIndex)) & "," & Val(Combo1.ItemData(Combo1.ListIndex))
    Cad = Cad & "," & DBSet(vCli.Agente, "T") & "," & DBSet(vCli.ForPago, "T") & "," & DBSet(vCli.FEnvio, "T") & ",0,0,0,0," & DBSet(Label1(0).Caption, "F") & ")"
    conn.Execute Cad
    
    'If HaCambiadoContador Then vC.Contador = vC.Contador + 1
    vC.IncrementarContador vC.TipoMovimiento
    
    
eCrearAlbaran:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
                
        Cad = "Se ha generado el albaran:   " & IIf(vC.TipoMovimiento = "CAR", "ALR", vC.TipoMovimiento) & " " & Format(vC.Contador, "000000") & vbCrLf
        Cad = Cad & "--> " & vC.NombreMovimiento
        MsgBox Cad, vbInformation
        
        
        
        
        CrearCarpetaComun vC.Contador, IIf(vC.TipoMovimiento = "CAR", "ALR", vC.TipoMovimiento)
        
        'Cargamos el combo de albaranes
        Cad = IIf(vC.TipoMovimiento = "CAR", "ALR", vC.TipoMovimiento)
        Cad = "codtipom = '" & Cad & "'"
        
        
        CargarCombo_Tabla Me.cboAlb, "scaalb", "concat(numalbar,' - ',nomclien,if(coalesce(referenc ,'')='','',concat('  [',referenc,']')))", "NumAlbar", Cad, , "fechaalb desc,numalbar desc"
        
        'Situamos el
        For NumRegElim = 0 To cboAlb.ListCount - 1
            Cad = Mid(cboAlb.List(NumRegElim), 1, InStr(1, cboAlb.List(NumRegElim), "-") - 1)
            If Val(Cad) = vC.Contador Then
                
                cboAlb.ListIndex = NumRegElim
            End If
            
        Next
        
        
    End If

    Set vC = New CTiposMov
    Set vCli = New CCliente
    
End Sub




Private Sub CrearCarpetaComun(NUmeroNuevoDeAlbaranRepacacion As Long, hcoCodTipoM As String)
Dim Referencia As String
Dim OtraCadena As String
Dim J As Byte
Dim Numalbar As Long
Dim trab As Long
Dim CadenaConsulta As String
Dim txtAnterior As String
Dim CarpetaTipoAlbaran As String


    On Error GoTo eCrearCarpetaComun

    If Not InstalacionEsEulerTaxco Then Exit Sub
    
    If EulerParam = "" Then Exit Sub
    
    
    J = 0
    
    CarpetaTipoAlbaran = ""
    If hcoCodTipoM = "ALR" Then J = 1: CarpetaTipoAlbaran = "REPARACIONES"
    If hcoCodTipoM = "ALO" Then J = 2: CarpetaTipoAlbaran = "orden de trabajo"
    If hcoCodTipoM = "ALE" Then J = 3: CarpetaTipoAlbaran = "trabajo exterior"
    If J = 0 Then Exit Sub
    



    
    If hcoCodTipoM = "ALR" Then
        'Z:(VALENCIA O ZALDIVIA)REPARACIONES\YYYY\NNNNNNN\
        trab = Combo1.ItemData(Combo1.ListIndex)
        CadenaConsulta = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", CStr(trab), "N")
        If CadenaConsulta = "" Then Err.Raise 513, , "Obteniendo centro trabajo (coste) trabajador " & trab
        
        If CadenaConsulta = "1" Then
            CadenaConsulta = "ZALDIBIA"
        Else
            CadenaConsulta = "VALENCIA"
        End If
        
        'cOMPRUEBO EL A�O
        'txtAnterior = EulerParam & "\REPARACIONES\" & CadenaConsulta & "\" & Year(CDate(Label1(0).Caption))
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & CadenaConsulta & "\" & Year(CDate(Label1(0).Caption))
    Else
        txtAnterior = EulerParam & "\" & CarpetaTipoAlbaran & "\" & Year(CDate(Label1(0).Caption))
    End If
        
        
    
    If Dir(txtAnterior, vbDirectory) = "" Then MkDir txtAnterior
    
    
    'Reemplazamos los carteres especiales de carpeta \/:*?<>| por espacios en blanco
    Referencia = ""   'Text1(13).Text
    For J = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", J, 1), " ")
    Next
    
    
    Numalbar = NUmeroNuevoDeAlbaranRepacacion
    
    Referencia = Trim(Format(Numalbar, "0000000") & " " & Referencia)
     
    
    
    CadenaConsulta = txtAnterior & "\" & Referencia
    If Dir(CadenaConsulta, vbDirectory) <> "" Then
        Err.Raise 513, , "Ya existe la carpeta: " & CadenaConsulta
    Else
        'NO existe vemos, si existe una carpeta que empieza por
        OtraCadena = txtAnterior & "\" & Format(Numalbar, "0000000") & "*"
        OtraCadena = Dir(OtraCadena, vbDirectory)
        If OtraCadena <> "" Then
            OtraCadena = txtAnterior & "\" & OtraCadena
            Name OtraCadena As CadenaConsulta
            OtraCadena = ""
        Else
            MkDir CadenaConsulta
        End If
    End If
    
    'Documentacion interna
    
    CadenaConsulta = CadenaConsulta & "\Documentacion interna"
    If Dir(CadenaConsulta, vbDirectory) <> "" Then
        Err.Raise 513, , "Ya existe la carpeta: " & CadenaConsulta
    Else
        MkDir CadenaConsulta
    End If
eCrearCarpetaComun:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description & vbCrLf & CadenaConsulta & vbCrLf & OtraCadena
End Sub

Private Sub RedondeaLaHora(HoraTexto As String)
Dim F As Date
Dim min As Integer
Dim Ya As Boolean
    
    min = Minute(CDate(HoraTexto))

    If min < 8 Then
        min = 0
    ElseIf min < 23 Then
        min = 15
    ElseIf min < 38 Then
        min = 30
    ElseIf min < 53 Then
        min = 45
    Else
        min = -1
    End If
            
    If min < 0 Then
        min = Hour(CDate(HoraTexto)) + 1
        If min >= 24 Then
            HoraTexto = "23:59:59"
        Else
            HoraTexto = Format(min, "00") & ":00:00"
        End If
    Else
        HoraTexto = Format(HoraTexto, "hh:") & Format(min, "00") & ":00"
    End If
    
End Sub


