VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmVerDtos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3960
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5760
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "data1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Maximo descuento particulares"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   2190
   End
   Begin VB.Label Label3 
      Caption         =   "Descuentos"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmAlmVerDtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mCodArtic As String

Dim PRimVez As Boolean


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PRimVez Then
        PRimVez = False
        CargaGrid

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Icon = frmPpal.Icon
    PRimVez = True
End Sub


Private Sub CargaGrid()
Dim B As Boolean
Dim SQL As String
Dim Aux As String
    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    SQL = "nomartic"
    Aux = DevuelveDesdeBD(conAri, "codfamia", "sartic", "codartic", mCodArtic, "T", SQL)
    Me.Label3.Caption = SQL
    
    SQL = "select sfamiadtos.clasifica,nombre,dtoline1,dtoline2 from sfamiadtos,sfamiatipodto where"
    SQL = SQL & " sfamiadtos.clasifica=sfamiatipodto.clasifica and codfamia="
    SQL = SQL & Aux
    SQL = SQL & " Order by sfamiadtos.clasifica"
    
    
    CargaGridGnral DataGrid1, data1, SQL, True

    
    CargaGrid2 DataGrid1, data1
    DataGrid1.ScrollBars = dbgAutomatic
        
   
    DataGrid1.Enabled = True

    SQL = DevuelveDesdeBD(conAri, "maxdtopar", "sfamia", "codfamia", Aux, "T")
    Text1.Text = SQL
    'PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

                vDataGrid.Columns(0).Caption = "Código"
                vDataGrid.Columns(0).visible = False
                
                vDataGrid.Columns(1).Caption = "Descripción"
                vDataGrid.Columns(1).Width = 3200
 
                vDataGrid.Columns(2).Caption = "Dto. 1"
                vDataGrid.Columns(2).Width = 900
                vDataGrid.Columns(2).Alignment = dbgRight
                vDataGrid.Columns(2).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(3).Caption = "Dto. 2"
                vDataGrid.Columns(3).Width = 900
                vDataGrid.Columns(3).Alignment = dbgRight
                vDataGrid.Columns(3).NumberFormat = FormatoDescuento
                
                


    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub
