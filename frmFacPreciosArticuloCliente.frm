VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacPreciosArticuloCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vta cliente Articulo"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8281
      SortKey         =   7
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "numero"
         Object.Width           =   2432
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Refer/Obr"
         Object.Width           =   7479
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cantidad"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Importe"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Dto1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Orden"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "Leyendo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9405
   End
End
Attribute VB_Name = "frmFacPreciosArticuloCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Datos As String

Dim primeravez As Boolean

Private Sub cmdCerrar_Click()
Unload Me
End Sub



Private Sub Form_activate()
    Dim I As Integer
    Dim SQL As String
    Dim Referencia As String
    Dim IT As ListItem
    If primeravez Then
        primeravez = False
        Set miRsAux = New ADODB.Recordset
        Referencia = RecuperaValor(Datos, 3)
           
        For I = 1 To 3
        
   
            Label9.Caption = "Leyendo " & I
            Label9.Refresh
        
  
            If I = 1 Then
               SQL = " select 'Ped' Tipo,fecpedcl fec,scaped.numpedcl clave,referenc,cantidad,precioar,dtoline1 from scaped,sliped Where scaped.NumPedcl = Sliped.NumPedcl"
                SQL = SQL & " "
            ElseIf I = 2 Then
                SQL = "select 'ALB' Tipo,fechaalb fec,scaalb.numalbar clave, referenc,cantidad,precioar,dtoline1"
                SQL = SQL & " from scaalb,slialb WHERE scaalb.numalbar=slialb.numalbar and"
                SQL = SQL & " scaalb.codtipom=slialb.codtipom  "
            Else
                SQL = " select 'FAC' Tipo,scafac.fecfactu fec,scafac.numfactu clave, referenc,cantidad,precioar,dtoline1 FROM scafac,scafac1,slifac"
                SQL = SQL & "  WHERE scafac1.codtipom = scafac.codtipom And scafac1.Numfactu = scafac.Numfactu And scafac1.FecFactu = scafac.FecFactu"
                SQL = SQL & " AND scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
                SQL = SQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa"

              
            End If
            SQL = SQL & " AND codclien =" & RecuperaValor(Datos, 1)
             SQL = SQL & " AND codartic =" & DBSet(RecuperaValor(Datos, 2), "T")
            
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                Set IT = ListView1.ListItems.Add()
                'scaped.numpedcl,referenc,fecpedcl,numlinea,solicitadas,servidas
                IT.Text = miRsAux!Tipo
                IT.SubItems(1) = miRsAux!fec
            
                IT.SubItems(2) = Format(miRsAux!Clave, "00000")
                IT.SubItems(3) = DBLet(miRsAux!referenc, "T")
                IT.SubItems(4) = Format(miRsAux!cantidad, FormatoCantidad)
                IT.SubItems(5) = Format(miRsAux!precioar, FormatoPrecio)
                
                IT.SubItems(6) = Format(miRsAux!dtoline1, FormatoCantidad)
                SQL = Format(miRsAux!fec, "yymmdd") & miRsAux!Tipo
                IT.SubItems(7) = SQL
                If Referencia <> "" Then
                    If DBLet(miRsAux!referenc, "T") = Referencia Then
                        IT.ListSubItems(3).ForeColor = vbRed
                        IT.ListSubItems(3).ToolTipText = "Misma referencia"
                    End If
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
        Next
        
        Set miRsAux = Nothing
    End If
    Me.Label9.Caption = "Cliente: " & RecuperaValor(Datos, 1) & "   Art: " & RecuperaValor(Datos, 2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    primeravez = True
    Me.Icon = frmPpal.Icon
    Screen.MousePointer = vbHourglass
End Sub

