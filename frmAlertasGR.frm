VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlertasGR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alertas"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   17400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Dias desde la fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   45
      TabIndex        =   10
      Top             =   135
      Width           =   17195
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   200
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "ped. cli|N|S|0||spara1|avipedcli|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "ped.pro.|N|S|0||spara1|avipedpro|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "alb.cli.|N|S|0||spara1|avialbcli|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6200
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "alb.pro.|N|S|0||spara1|avialbpro|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   12300
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "avi.mante|N|S|0||spara1|avimanteni|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   10620
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "avi.avisos|N|S|0||spara1|aviavios|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   8625
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "avi.repa.|N|S|0||spara1|avirepara|##0||"
         Text            =   "3"
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   200
         TabIndex        =   17
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1950
         TabIndex        =   16
         Top             =   315
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Albaranes Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4140
         TabIndex        =   15
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Albaranes Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6195
         TabIndex        =   14
         Top             =   315
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mantenimientos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   12300
         TabIndex        =   13
         Top             =   315
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Reparaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   8625
         TabIndex        =   12
         Top             =   315
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Avisos "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   10620
         TabIndex        =   11
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   7
      Top             =   8910
      Width           =   1065
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlertasGR.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlertasGR.frx":6862
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7155
      Left            =   3240
      TabIndex        =   8
      Top             =   1575
      Width           =   14045
      _ExtentX        =   24765
      _ExtentY        =   12621
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
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2382
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Codigo"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Forma Pago"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   2787
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7110
      Left            =   45
      TabIndex        =   9
      Top             =   1575
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   12541
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   3
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAlertasGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String
Dim F As Date


Private Sub Command1_Click()
    'Unload Me
    Caption = "UNLOAD"
End Sub

Private Sub Form_Activate()
    PonerFoco Text1(0)
End Sub

Private Sub Form_Load()
Dim I As Long

    Me.Icon = frmPpal.Icon
    Set miRsAux = New ADODB.Recordset
    Set TreeView1.ImageList = Me.ImageList1
    Set ListView1.SmallIcons = frmPpal.ImgListPpal
    CargaTreeView
    Set TreeView1.SelectedItem = Nothing
    Screen.MousePointer = vbDefault
    
    I = 200
    
    If vParamAplic.avipedcli = 0 Then
        Text1(0).Enabled = False
'        Text1(0).visible = False
        Text1(0).Text = ""
'        Label1(0).visible = False
    Else
        Text1(0).Enabled = True
        Text1(0).visible = True
        Text1(0).Text = vParamAplic.avipedcli
'        Label1(0).visible = True
'        Text1(0).Left = I
'        Label1(0).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.avipedpro = 0 Then
        Text1(1).Enabled = False
'        Text1(1).visible = False
        Text1(1).Text = ""
'        Label1(1).visible = False
    Else
        Text1(1).Enabled = True
        Text1(1).visible = True
        Text1(1).Text = vParamAplic.avipedpro
'        Label1(1).visible = True
'        Text1(1).Left = I
'        Label1(1).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.avialbcli = 0 Then
        Text1(2).Enabled = False
'        Text1(2).visible = False
        Text1(2).Text = ""
'        Label1(2).visible = False
    Else
        Text1(2).Enabled = True
        Text1(2).visible = True
        Text1(2).Text = vParamAplic.avialbcli
'        Label1(2).visible = True
'        Text1(2).Left = I
'        Label1(2).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.avialbpro = 0 Then
        Text1(3).Enabled = False
'        Text1(3).visible = False
        Text1(3).Text = ""
'        Label1(3).visible = False
    Else
        Text1(3).Enabled = True
        Text1(3).visible = True
        Text1(3).Text = vParamAplic.avialbpro
'        Label1(3).visible = True
'        Text1(3).Left = I
'        Label1(3).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.avirepara = 0 Then
        Text1(4).Enabled = False
'        Text1(4).visible = False
        Text1(4).Text = ""
'        Label1(4).visible = False
    Else
        Text1(4).Enabled = True
        Text1(4).visible = True
        Text1(4).Text = vParamAplic.avirepara
'        Label1(4).visible = True
'        Text1(4).Left = I
'        Label1(4).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.aviavisos = 0 Then
        Text1(5).Enabled = False
'        Text1(5).visible = False
        Text1(5).Text = ""
'        Label1(5).visible = False
    Else
        Text1(5).Enabled = True
        Text1(5).visible = True
        Text1(5).Text = vParamAplic.aviavisos
'        Label1(5).visible = True
'        Text1(5).Left = I
'        Label1(5).Left = I
'        I = I + 2200
    End If
    
    If vParamAplic.avimanteni = 0 Then
        Text1(6).Enabled = False
'        Text1(6).visible = False
        Text1(6).Text = ""
'        Label1(6).visible = False
    Else
        Text1(6).Enabled = True
        Text1(6).visible = True
        Text1(6).Text = vParamAplic.avipedpro
'        Label1(6).visible = True
'        Text1(6).Left = I
'        Label1(6).Left = I
'        I = I + 1800
    End If
    
End Sub


Private Sub CargaTreeView()
Dim NO As Node
    TreeView1.Nodes.Clear
    
    'Para cada opcion de alertas vamos viendo si lo ponemos.
    Set NO = TreeView1.Nodes.Add(, , "c1", "PEDIDOS CLIENTE")
    NO.Image = 1
    NO.Tag = 3
    If vParamAplic.avipedcli = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Tag = 1   'Pondremos el icono
        NO.Image = 2
    End If
    
    
    Set NO = TreeView1.Nodes.Add(, , "c2", "PEDIDOS PROVEEDORES")
    NO.Image = 1
    NO.Tag = 4
    If vParamAplic.avipedpro = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    
    Set NO = TreeView1.Nodes.Add(, , "c3", "ALBARANES CLIENTE")
    NO.Image = 1
    NO.Tag = 7
    If vParamAplic.avialbcli = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    Set NO = TreeView1.Nodes.Add(, , "c4", "ALBARANES PROVEEDORES")
    NO.Image = 1
    NO.Tag = 10
    If vParamAplic.avipedpro = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    Set NO = TreeView1.Nodes.Add(, , "c5", "REPARACIONES")
    NO.Image = 1
    NO.Tag = 16
    If vParamAplic.avirepara = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    Set NO = TreeView1.Nodes.Add(, , "c6", "AVISOS")
    NO.Image = 1
    NO.Tag = 1
    If vParamAplic.aviavisos = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    Set NO = TreeView1.Nodes.Add(, , "c7", "MANTENIMIENTOS")
    NO.Image = 1
    NO.Tag = 12
    If vParamAplic.avimanteni = 0 Then
        NO.ForeColor = RGB(192, 192, 192)
        NO.Image = 2
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set miRsAux = Nothing
End Sub


Private Sub CargaListView(NumNod As Integer, LaImagen As Integer)
Dim IT As ListItem

    On Error GoTo ECA
    FijaCadenaSQL NumNod
    If SQL = "" Then Exit Sub
    
    'SI no cargamos. SIiiiempre sera el mismo orden para los campos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add(, "c" & NumRegElim)
        
        IT.Text = miRsAux.Fields(0)
        IT.SubItems(1) = miRsAux.Fields(1)
        IT.SubItems(2) = Format(miRsAux.Fields(2), "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux.Fields(3)
        IT.SubItems(4) = miRsAux.Fields(4)
'        IT.SubItems(4) = miRsAux.Fields(4)
        IT.SubItems(5) = Format(miRsAux.Fields(5), FormatoImporte)
        
        'IT.SmallIcon = LaImagen
        miRsAux.MoveNext
        NumRegElim = NumRegElim + 1
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Exit Sub
ECA:
    MuestraError Err.Number, SQL
    Set miRsAux = Nothing
End Sub



Private Sub FijaCadenaSQL(Opcion As Integer)
    
    SQL = ""
    Select Case Opcion
    Case 1
        ' "PEDIDOS CLIENTE")
        SQL = "select scaped.numpedcl,'',scaped.fecpedcl,scaped.codclien,scaped.nomclien,nomforpa,sum(importel) "
        SQL = SQL & " from scaped,sliped,sforpa WHERE scaped.numpedcl=sliped.numpedcl and scaped.codforpa = sforpa.codforpa "
        'WHERE del alerta
'        F = DateAdd("d", -vParamAplic.avipedcli, Now)
        F = DateAdd("d", -Val(Text1(0)), Now)
        SQL = SQL & " AND scaped.fecpedcl <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fecpedcl"
    
    Case 2
        '"PEDIDOS PROVEEDORES")
        SQL = "select scappr.numpedpr,'',scappr.fecpedpr,scappr.codprove,scappr.nomprove,nomforpa,sum(importel)"
        SQL = SQL & " from scappr,slippr,sforpa WHERE scappr.numpedpr=slippr.numpedpr and scappr.codforpa = sforpa.codforpa "
'        F = DateAdd("d", -vParamAplic.avipedpro, Now)
        F = DateAdd("d", -Val(Text1(1)), Now)
        SQL = SQL & " AND scappr.fecpedpr <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fecpedpr"
    
    Case 3
        'Set NO = TreeView1.Nodes.Add(, , "c3", "ALBARANES CLIENTE")
        SQL = "select concat(scaalb.codtipom, scaalb.numalbar),replace(ucase(stipom.nomtipom),'ALBARAN',''),scaalb.fechaalb,scaalb.codclien,scaalb.nomclien,nomforpa,sum(importel)"
        SQL = SQL & " from scaalb,slialb, stipom,sforpa  WHERE scaalb.codtipom = stipom.codtipom and scaalb.numalbar=slialb.numalbar and scaalb.codtipom=slialb.codtipom and scaalb.codforpa = sforpa.codforpa "
'        F = DateAdd("d", -vParamAplic.avialbcli, Now)
        F = DateAdd("d", -Val(Text1(2)), Now)
        SQL = SQL & " AND scaalb.fechaalb <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fechaalb"

    Case 4
        '"ALBARANES PROVEEDORES"
        SQL = "select  scaalp.numalbar,'',scaalp.fechaalb,scaalp.codprove,scaalp.nomprove,nomforpa,sum(importel)"
        SQL = SQL & " from scaalp,slialp,sforpa WHERE scaalp.numalbar=slialp.numalbar and scaalp.fechaalb=slialp.fechaalb and scaalp.codforpa=sforpa.codforpa"
        SQL = SQL & " and scaalp.codprove=slialp.codprove"
'        F = DateAdd("d", -vParamAplic.avialbpro, Now)
        F = DateAdd("d", -Val(Text1(3)), Now)
        SQL = SQL & " AND scaalp.fechaalb <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fechaalb"
    
    Case 5
        'Set NO = TreeView1.Nodes.Add(, , "c5", "REPARACIONES")
'        F = DateAdd("d", -vParamAplic.avirepara, Now)
        F = DateAdd("d", -Val(Text1(4)), Now)
        SQL = "select scarep.numrepar,'',fecrepar,scarep.codclien,scarep.nomclien,if(imppresu1 is null,'0.0',imppresu1) from"
        SQL = SQL & " scarep,sclien where scarep.codclien=sclien.codclien  AND motivore is null "
        SQL = SQL & " AND scarep.fecrepar <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fecrepar"
    Case 6
        '"AVISOS"
'        F = DateAdd("d", -vParamAplic.aviavisos, Now)
        F = DateAdd("d", -Val(Text1(5)), Now)
        SQL = "select numaviso,'',fechaavi,codclien,nomclien,'' from scaavi where situacio=0 "
        SQL = SQL & " AND scaavi.fechaavi <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fechaavi"
    
    Case 7
        'Mantenimientos
'        F = DateAdd("d", -vParamAplic.avimanteni, Now)
        F = DateAdd("d", -Val(Text1(6)), Now)

        SQL = "select scaman.nummante,'',concat(""01/"" , lpad(ulmesfac+1,2,'0'),""/" & Year(F) & """),scaman.codclien,nomclien,"" "" from scaman,sclien"
        SQL = SQL & " Where scaman.CodClien = sclien.CodClien And ("
        SQL = SQL & "(tipopago = 0 And ulmesfac < " & Month(F)
        SQL = SQL & ") Or (tipopago = 1 And ulmesfac < " & Month(F) - 3
        SQL = SQL & ") Or (tipopago = 2 And ulmesfac <" & Month(F) - 6
        'Noviembre 2013
        SQL = SQL & ") Or (tipopago = 4 And ulmesfac <" & Month(F) - 2
        
        SQL = SQL & ") Or (tipopago = 3 And ulmesfac = 0))"
    
    
    
    End Select
    
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 4
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
 '   KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean
    KEYpressGnral KeyAscii, 1, cerrar
    If cerrar Then
        Unload Me
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
    If Not PerderFocoGnral(Text1(Index), 4) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
  
    If PonerFormatoEntero(Text1(Index)) Then
        If Not TreeView1.SelectedItem Is Nothing Then
            If TreeView1.Nodes(Index + 1) = TreeView1.SelectedItem Then TreeView1_NodeClick TreeView1.Nodes(Index + 1)
        Else
            TreeView1_NodeClick TreeView1.Nodes(Index + 1)
        End If
    Else
        PonerFoco Text1(Index)
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
'    If ListView1.Tag = CStr(Node.Index) Then Exit Sub
    
    ListView1.ListItems.Clear
    ListView1.Tag = Node.Index
    If Node.Image <> 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargaListView Node.Index, CInt(Node.Tag)
    Screen.MousePointer = vbDefault
    
End Sub
