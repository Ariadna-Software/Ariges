VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FramePreciosProveArt 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   9375
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1826
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Proveedor"
            Object.Width           =   6509
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Actual"
            Object.Width           =   2258
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "F. cambio"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio N."
            Object.Width           =   2258
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame FrPrecioPr 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9375
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6800
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblProve 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   8160
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblProve 
         AutoSize        =   -1  'True
         Caption         =   "Dto2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7200
         TabIndex        =   9
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblProve 
         AutoSize        =   -1  'True
         Caption         =   "Dto1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblProve 
         Caption         =   "PRECIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblProve 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblProve 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProve 
         Caption         =   "Alb/Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblInd 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   5535
   End
End
Attribute VB_Name = "frmListadoPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
'   0.-  Listado precios proveedor
'   1.-  precios un articulo por cada proveedor desde slispr

Public CadenaPasoDatos As String
'   0.- codartic|codprove|




Dim SQL As String
Dim Cad As String


Dim PrimeraVez As Boolean

Private Sub cmdCancelar_Click(Index As Integer)
'    For NumRegElim = 1 To ListView1.ColumnHeaders.Count
'        Debug.Print ListView1.ColumnHeaders(NumRegElim).Text & "-" & ListView1.ColumnHeaders(NumRegElim).Width
'    Next
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
    
        If Opcion = 0 Then
            Caption = "Ver precios proveedor"
            CargaPreciosProveedor
        
        ElseIf Opcion = 1 Then
            Caption = "Lista precios"
            CargaPreciosDesdeSlispr
        End If
        
        
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        lblInd.Caption = ""
    End If
    
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    FramePreciosProveArt.visible = False
    FrPrecioPr.visible = False
    lblInd.Caption = ""
    Select Case Opcion
    Case 0
        FrPrecioPr.visible = True
    Case 1
        FramePreciosProveArt.visible = True
    End Select
End Sub


Private Sub CargaPreciosProveedor()
Dim N As Node

    TreeView1.Nodes.Clear
    lblInd.Caption = "Leyendo BD"
    lblInd.Refresh
    Set TreeView1.ImageList = frmPpal.ImgListPpal
    
    
    SQL = "DELETE FROM tmpslipreu WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL
    DoEvents
    
    'Metemos los oprecios del proveedor en question
    'nuofert: proveedor    numliene 0 alb  1 fra    codartic:fecha mov         nomartic: nºfra/alb

    Cad = "insert into `tmpslipreu` (`codusu`,codalmac,`numofert`,`numlinea`,codartic,nomartic,`cantidad`,`precioar`,`dtoline1`,`dtoline2`,`importel`) "

    'Codalmac lo utilizare para que me ponga primero los datos del proveedor

    'ALBARANES
    SQL = "select " & vUsu.Codigo & ",0,codprove,0,date_format(fechaalb,'%Y/%m/%d'),numalbar,cantidad,precioar,`dtoline1`,`dtoline2`,`importel` from slialp"
    SQL = SQL & " WHERE codartic = " & DBSet(RecuperaValor(CadenaPasoDatos, 1), "T")
 
    SQL = Cad & SQL
    conn.Execute SQL
    
    'FRA
    SQL = "select " & vUsu.Codigo & ",0,codprove,1,date_format(fecfactu,'%Y/%m/%d'),numfactu,cantidad,precioar,`dtoline1`,`dtoline2`,`importel` from slifpc"
    SQL = SQL & " WHERE codartic = " & DBSet(RecuperaValor(CadenaPasoDatos, 1), "T")
    SQL = Cad & SQL
    conn.Execute SQL
    
    SQL = "UPDATE tmpslipreu SET codalmac = 1 WHERE numofert <> " & RecuperaValor(CadenaPasoDatos, 2)
    SQL = SQL & " AND codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    
    'oK
    'Ya tengo en CadenaPasoDatos
    Me.Refresh
    lblInd.Caption = "Mostrando reg."
    lblInd.Refresh
    SQL = "Select * from tmpslipreu WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " ORDER BY codalmac,numofert,codartic desc"  '1º el proveedor k copnsulto. 2 otros proveedores  3. fechamov
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
        If miRsAux!NumOfert <> NumRegElim Then
            If NumRegElim >= 0 Then
                'Para ver si el prroveedor que acabo de insertar es el que tocaba
                 If NumRegElim = Val(RecuperaValor(CadenaPasoDatos, 2)) Then
                    N.Parent.EnsureVisible
                    N.Parent.Expanded = True
                End If
            End If
        
            'Proveedor.
            NumRegElim = miRsAux!NumOfert
            SQL = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", CStr(NumRegElim))
            Set N = TreeView1.Nodes.Add(, , "C" & NumRegElim, SQL)
            N.Image = 4 'proveedor
        End If
        
        'Primero
        SQL = miRsAux!NomArtic & Space(11)
        SQL = Mid(SQL, 1, 11) & Format(miRsAux!codArtic, "dd/mm/yyyy") & "  "
        'Cantidad, precioar, dto1,dto2-->  cantidad,precioar,`dtoline1`,`dtoline2`,`importel`
        Cad = Space(9) & Format(miRsAux!cantidad, FormatoCantidad)
        SQL = SQL & Right(Cad, 9)
        
        Cad = Space(13) & Format(miRsAux!precioar, FormatoPrecio)
        SQL = SQL & Right(Cad, 13)
        
        Cad = Space(7) & Format(miRsAux!dtoline1, FormatoDescuento)
        SQL = SQL & Right(Cad, 7)
        
        Cad = Space(7) & Format(miRsAux!dtoline1, FormatoDescuento)
        SQL = SQL & Right(Cad, 7)
        
        Cad = Space(10) & Format(miRsAux!ImporteL, FormatoImporte)
        SQL = SQL & Right(Cad, 10)
        
        
        Set N = TreeView1.Nodes.Add(CStr("C" & NumRegElim), tvwChild, , SQL)
        
        
        
        If miRsAux!numlinea = 0 Then
            N.Image = 9
        Else
            N.Image = 10
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
            'Para ver si expande el unico nodo
            If NumRegElim >= 0 Then
                'Para ver si el prroveedor que acabo de insertar es el que tocaba
                 If NumRegElim = Val(RecuperaValor(CadenaPasoDatos, 2)) Then
                    
                    N.Parent.EnsureVisible
                    N.Parent.Expanded = True
                End If
            End If
    
End Sub




        


Private Sub CargaPreciosDesdeSlispr()

    TreeView1.Nodes.Clear
    lblInd.Caption = "Leyendo BD"
    lblInd.Refresh
    
    Cad = Val(RecuperaValor(CadenaPasoDatos, 2))
    
    SQL = "select slispr.codprove,nomprove,precioac,fechanue,precionu from slispr,sprove where"
    SQL = SQL & " slispr.codprove=sprove.codprove AND codartic = " & DBSet(RecuperaValor(CadenaPasoDatos, 1), "T")
    SQL = SQL & " ORDER BY 1"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        ListView1.ListItems.Add , , Format(miRsAux!CodProve, "000000")
        ListView1.ListItems(NumRegElim).SubItems(1) = miRsAux!nomprove
        
        If Val(Cad) = miRsAux!CodProve Then
            ListView1.ListItems(NumRegElim).Bold = True
            ListView1.ListItems(NumRegElim).ListSubItems(1).Bold = True
        End If
        
        
        ListView1.ListItems(NumRegElim).SubItems(2) = Format(miRsAux!precioac, FormatoPrecio)
        If IsNull(miRsAux!fechanue) Then
            ListView1.ListItems(NumRegElim).SubItems(3) = " "
            ListView1.ListItems(NumRegElim).SubItems(4) = " "
        Else
            ListView1.ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioac, FormatoFecha)
            ListView1.ListItems(NumRegElim).SubItems(4) = Format(miRsAux!precionu, FormatoPrecio)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub
