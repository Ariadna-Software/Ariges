VERSION 5.00
Begin VB.Form frmFacTPVMovCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos caja"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Index           =   1
      Left            =   1560
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4080
      Width           =   4260
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   5760
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   5760
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2760
      Width           =   3495
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
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Tag             =   "Inicial|N|N|0||||#,##0.00||"
      Text            =   "Text1"
      Top             =   3360
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   240
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   840
      Width           =   4275
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1560
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   3120
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Observacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Terminal"
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
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   825
   End
End
Attribute VB_Name = "frmFacTPVMovCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SoloVer As Boolean
Public linea As Long

Dim Cad As String

Private Sub cmdAceptar_Click()

    If linea = 0 Then ' insertar

        If Combo1.ListIndex < 0 Then Exit Sub
        If Text1(0).Text = "" Then Exit Sub
        
        NumRegElim = 0  'Salida
        '%=%= cambiado por lo de abajo
        'If Me.Option1(1).Value Then NumRegElim = 1  '1Entrada
        NumRegElim = DevuelveDesdeBDNew(conAri, "stpvtipogastos", "tipo", "id", CStr(Combo1.ItemData(Combo1.ListIndex)), "N")
        Cad = DevuelveDesdeBD(conAri, "max(idlin)", "stpventradassalidas", "numtermi", vParamTPV.NumeroDeTerminal, "F")
        Cad = CStr(Val(Cad) + 1)
        Cad = vParamTPV.NumeroDeTerminal & "," & Cad & "," & DBSet(Now, "FH") & "," & NumRegElim & ","
        Cad = Cad & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "T", "S") & "," & Combo1.ItemData(Combo1.ListIndex) & "," & Text2(1).Tag & ")"
        
        
        Cad = "INSERT INTO stpventradassalidas(numtermi,idLin,diahora,Entrada,importe,descrip,concepto,codtraba) VALUES (" & Cad
        If ejecutar(Cad, False) Then Unload Me
        
    Else ' modificar
        NumRegElim = DevuelveDesdeBDNew(conAri, "stpvtipogastos", "tipo", "id", CStr(Combo1.ItemData(Combo1.ListIndex)), "N")
        
        Cad = "UPDATE stpventradassalidas SET "
        Cad = Cad & "concepto = " & Combo1.ItemData(Combo1.ListIndex)
        Cad = Cad & ",entrada = " & NumRegElim
        Cad = Cad & ",importe = " & DBSet(Text1(0).Text, "N")
        Cad = Cad & ",descrip = " & DBSet(Text1(1).Text, "T", "S")
        Cad = Cad & " where numtermi = " & vParamTPV.NumeroDeTerminal
        Cad = Cad & " and idlin = " & DBSet(linea, "N")
        
        If ejecutar(Cad, False) Then Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Load()

    CargaCombo

    If linea = 0 Then
        limpiar Me
        
        cmdAceptar.Enabled = True
        Text2(0).Text = Format(Now, "dd/mm/yyyy")
        Text2(1).Tag = PonerTrabajadorConectado(Cad)
        If Cad = "" Then
            Cad = "Error trabajador conectado"
            cmdAceptar.Enabled = False
        End If
        Text2(1).Text = Text2(1).Tag & " - " & Cad
        
        'destermi   numtermi   spatpvt
        Cad = DevuelveDesdeBD(conAri, "destermi", "spatpvt", "numtermi", vParamTPV.NumeroDeTerminal)
        Text2(2).Text = Cad
        
        'PosicionarCombo Combo1, 1
        Combo1.ListIndex = 0
    Else
        cmdAceptar.Enabled = True
        
        cmdAceptar.visible = Not SoloVer
        
        Text2(0).Text = Format(Now, "dd/mm/yyyy")
        Text2(1).Tag = PonerTrabajadorConectado(Cad)
        If Cad = "" Then
            Cad = "Error trabajador conectado"
            cmdAceptar.Enabled = False
        End If
        Text2(1).Text = Text2(1).Tag & " - " & Cad
        
        'destermi   numtermi   spatpvt
        Cad = DevuelveDesdeBD(conAri, "destermi", "spatpvt", "numtermi", vParamTPV.NumeroDeTerminal)
        Text2(2).Text = Cad
    
        'situamos los campos para que los modifique
        SituarCampos
    
    End If
'%=%=    Option1_Click 0
    
End Sub

Private Sub SituarCampos()
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "select * from stpventradassalidas where numtermi = " & vParamTPV.NumeroDeTerminal
    Sql = Sql & " and idlin = " & DBSet(linea, "N")

    Text1(0).Text = ""
    Text1(1).Text = ""

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        PosicionarCombo Combo1, DBLet(Rs!concepto, "N")
        Text1(0).Text = Format(DBLet(Rs!Importe, "N"), "###,###,##0.00")
        Text1(1).Text = DBLet(Rs!Descrip, "T")
    End If

End Sub


Private Sub Option1_Click(Index As Integer)
    Me.Combo1.Clear
    Set miRsAux = New ADODB.Recordset
    Cad = "select * from stpvtipogastos where tipo =" & Index & " order by 1 "
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux.Fields!observa
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!ID
        If Combo1.NewIndex = 0 Then Combo1.ListIndex = 0
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    PonerFoco Text1(0)
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False

End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim mTag As cTag
    
    If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
           
                If Text1(Index).Text = "" Then Exit Sub
                Set mTag = New cTag
                If mTag.Cargar(Text1(Index)) Then
                    If mTag.Cargado Then
                        If mTag.Comprobar(Text1(Index)) Then
                            FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
                        Else
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
                Set mTag = Nothing

        
    End Select
    '---
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
Dim miRsAux As ADODB.Recordset
    
    Combo1.Clear
    
    Set miRsAux = New ADODB.Recordset
    
    ' camaras
    Sql = "Select * from stpvtipogastos order by id"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!observa
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!ID
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub
